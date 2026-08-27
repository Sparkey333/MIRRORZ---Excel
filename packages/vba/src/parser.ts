/**
 * VBA parser: token stream to abstract syntax tree.
 *
 * The point of this stage is not yet to run anything. It is to know exactly what
 * a macro DOES before deciding whether to run it, which is what makes the
 * compatibility report in `compat.ts` possible - and that report is the feature
 * that separates us from implementations whose defining failure is running a
 * macro until it silently does the wrong thing.
 *
 * The grammar covered is the executable subset real Excel macros use. Constructs
 * that exist but cannot be safely executed - notably `Declare ... Lib`, which
 * calls arbitrary native code - are parsed rather than rejected, so a file still
 * loads and can be inspected, and then refused at execution time. Parsing
 * something is not a promise to run it.
 */

import { type Token, Tk, tokenize, significant } from './lexer.js';

export class VbaParseError extends Error {
  constructor(
    message: string,
    readonly line: number,
  ) {
    super(`${message} (line ${line})`);
    this.name = 'VbaParseError';
  }
}

// --- expressions ---------------------------------------------------------

export type Expr =
  | { kind: 'number'; value: number; suffix?: string }
  | { kind: 'string'; value: string }
  | { kind: 'date'; value: string }
  | { kind: 'bool'; value: boolean }
  | { kind: 'nothing' }
  | { kind: 'null' }
  | { kind: 'empty' }
  | { kind: 'me' }
  | { kind: 'identifier'; name: string }
  | { kind: 'member'; target: Expr | null; name: string }
  | { kind: 'call'; target: Expr; args: Argument[] }
  | { kind: 'index'; target: Expr; args: Argument[] }
  | { kind: 'unary'; op: string; operand: Expr }
  | { kind: 'binary'; op: string; left: Expr; right: Expr }
  | { kind: 'new'; typeName: string }
  | { kind: 'paren'; inner: Expr };

export interface Argument {
  /** Named argument, as in `Range(Cell1:="A1")`. */
  name?: string;
  /** ByVal / ByRef marker at the call site. */
  modifier?: 'ByVal' | 'ByRef';
  value: Expr | null;
}

// --- statements ----------------------------------------------------------

export type Stmt =
  | { kind: 'option'; text: string }
  | { kind: 'dim'; scope: DeclScope; vars: VarDecl[] }
  | { kind: 'const'; scope: DeclScope; vars: ConstDecl[] }
  | { kind: 'redim'; preserve: boolean; vars: VarDecl[] }
  | { kind: 'assign'; target: Expr; value: Expr; isSet: boolean }
  | { kind: 'call'; target: Expr; args: Argument[] }
  | { kind: 'if'; branches: { condition: Expr; body: Stmt[] }[]; elseBody?: Stmt[] }
  | { kind: 'for'; variable: Expr; from: Expr; to: Expr; step?: Expr; body: Stmt[] }
  | { kind: 'forEach'; variable: Expr; collection: Expr; body: Stmt[] }
  | { kind: 'do'; test?: { kind: 'while' | 'until'; condition: Expr; post: boolean }; body: Stmt[] }
  | { kind: 'while'; condition: Expr; body: Stmt[] }
  | { kind: 'selectCase'; subject: Expr; cases: CaseClause[]; elseBody?: Stmt[] }
  | { kind: 'with'; subject: Expr; body: Stmt[] }
  | { kind: 'exit'; target: string }
  | { kind: 'goto'; label: string }
  | { kind: 'gosub'; label: string }
  | { kind: 'return' }
  | { kind: 'onError'; mode: 'resumeNext' | 'goto' | 'goto0'; label?: string }
  | { kind: 'resume'; target?: string }
  | { kind: 'label'; name: string }
  | { kind: 'erase'; targets: Expr[] }
  | { kind: 'stop' }
  | { kind: 'end' }
  | { kind: 'raiseEvent'; name: string; args: Argument[] }
  | { kind: 'declare'; name: string; lib: string; alias?: string; isFunction: boolean }
  | { kind: 'type'; name: string; fields: VarDecl[] }
  | { kind: 'enum'; name: string; members: { name: string; value?: Expr }[] }
  | { kind: 'attribute'; text: string }
  | { kind: 'directive'; text: string }
  | { kind: 'unknown'; text: string };

export type DeclScope = 'Dim' | 'Private' | 'Public' | 'Global' | 'Static' | 'Friend';

export interface VarDecl {
  name: string;
  typeName?: string;
  /** Array bounds, empty for a dynamic array declared as `x()`. */
  bounds?: { lower?: Expr; upper: Expr }[];
  isArray?: boolean;
  withEvents?: boolean;
  isNew?: boolean;
}

export interface ConstDecl {
  name: string;
  typeName?: string;
  value: Expr;
}

export interface CaseClause {
  /** Each label is a value, a range (`1 To 5`), or a comparison (`Is > 5`). */
  labels: CaseLabel[];
  body: Stmt[];
}

export type CaseLabel =
  | { kind: 'value'; value: Expr }
  | { kind: 'range'; from: Expr; to: Expr }
  | { kind: 'compare'; op: string; value: Expr };

export interface Parameter {
  name: string;
  optional: boolean;
  byRef: boolean;
  paramArray: boolean;
  typeName?: string;
  defaultValue?: Expr;
}

export type ProcedureKind = 'Sub' | 'Function' | 'PropertyGet' | 'PropertyLet' | 'PropertySet';

export interface Procedure {
  kind: ProcedureKind;
  name: string;
  visibility: 'Public' | 'Private' | 'Friend';
  isStatic: boolean;
  params: Parameter[];
  returnType?: string;
  body: Stmt[];
  line: number;
}

export interface Module {
  /** Statements outside any procedure: options, declarations, types. */
  declarations: Stmt[];
  procedures: Procedure[];
  attributes: string[];
  /** Problems that did not stop the parse. */
  warnings: string[];
}

/**
 * Operator precedence, lowest binding first. VBA's set is larger than most
 * languages': the logical operators run all the way out to `Imp`, and `Like`
 * and `Is` sit with the comparisons.
 */
const PRECEDENCE: Readonly<Record<string, number>> = {
  IMP: 1,
  EQV: 2,
  XOR: 3,
  OR: 4,
  AND: 5,
  '=': 7,
  '<>': 7,
  '<': 7,
  '>': 7,
  '<=': 7,
  '>=': 7,
  IS: 7,
  LIKE: 7,
  '&': 8,
  '+': 9,
  '-': 9,
  MOD: 10,
  '\\': 11,
  '*': 12,
  '/': 12,
  '^': 14,
};

/** `Not` is a prefix operator sitting between comparison and And. */
const NOT_PRECEDENCE = 6;
/** Unary minus binds tighter than any binary operator except exponentiation. */
const UNARY_PRECEDENCE = 13;

export function parseModule(source: string): Module {
  return new Parser(significant(tokenize(source))).parseModule();
}

class Parser {
  private i = 0;

  constructor(private readonly tokens: Token[]) {}

  private get current(): Token {
    return this.tokens[this.i] ?? this.tokens[this.tokens.length - 1]!;
  }

  private at(type: Tk, upper?: string): boolean {
    const t = this.current;
    return t.type === type && (upper === undefined || t.upper === upper);
  }

  private atKeyword(...words: string[]): boolean {
    const t = this.current;
    return t.type === Tk.Keyword && words.includes(t.upper);
  }

  private advance(): Token {
    const t = this.current;
    if (this.i < this.tokens.length - 1) this.i++;
    return t;
  }

  private expect(type: Tk, what: string): Token {
    if (this.current.type !== type) {
      throw new VbaParseError(`expected ${what}, found ${JSON.stringify(this.current.text)}`, this.current.line);
    }
    return this.advance();
  }

  private skipTerminators(): void {
    while (this.at(Tk.Terminator)) this.advance();
  }

  parseModule(): Module {
    const module: Module = { declarations: [], procedures: [], attributes: [], warnings: [] };

    this.skipTerminators();
    while (this.current.type !== Tk.EOF) {
      if (this.at(Tk.Attribute)) {
        module.attributes.push(this.advance().text);
        this.skipTerminators();
        continue;
      }

      const procedure = this.tryParseProcedure(module.warnings);
      if (procedure) {
        module.procedures.push(procedure);
        this.skipTerminators();
        continue;
      }

      try {
        module.declarations.push(this.parseStatement());
      } catch (err) {
        // One unparseable declaration must not lose the whole module: an
        // inspectable partial parse is far more useful than a refusal.
        module.warnings.push(err instanceof Error ? err.message : String(err));
        this.recover();
      }
      this.skipTerminators();
    }
    return module;
  }

  /** Skip to the next statement boundary after an error. */
  private recover(): void {
    while (!this.at(Tk.Terminator) && this.current.type !== Tk.EOF) this.advance();
  }

  private tryParseProcedure(warnings: string[]): Procedure | undefined {
    const start = this.i;
    let visibility: Procedure['visibility'] = 'Public';
    let isStatic = false;

    while (this.atKeyword('PUBLIC', 'PRIVATE', 'FRIEND', 'STATIC')) {
      const word = this.advance().upper;
      if (word === 'STATIC') isStatic = true;
      else visibility = (word.charAt(0) + word.slice(1).toLowerCase()) as Procedure['visibility'];
    }

    let kind: ProcedureKind;
    if (this.atKeyword('SUB')) {
      kind = 'Sub';
      this.advance();
    } else if (this.atKeyword('FUNCTION')) {
      kind = 'Function';
      this.advance();
    } else if (this.atKeyword('PROPERTY')) {
      this.advance();
      const accessor = this.advance().upper;
      kind =
        accessor === 'GET' ? 'PropertyGet' : accessor === 'SET' ? 'PropertySet' : 'PropertyLet';
    } else {
      // Not a procedure; put back whatever modifiers were consumed.
      this.i = start;
      return undefined;
    }

    const line = this.current.line;
    const name = this.expect(Tk.Identifier, 'a procedure name').text;
    const params = this.at(Tk.LParen) ? this.parseParameterList() : [];

    let returnType: string | undefined;
    if (this.atKeyword('AS')) {
      this.advance();
      returnType = this.parseTypeName();
    }
    this.skipTerminators();

    const endWord = kind.startsWith('Property') ? 'PROPERTY' : kind.toUpperCase();
    const body = this.parseBlock(() => this.atEndOf(endWord));
    if (this.atKeyword('END')) {
      this.advance();
      if (this.current.type === Tk.Keyword) this.advance();
    } else {
      warnings.push(`procedure ${name} is missing its End ${kind} (line ${line})`);
    }

    const procedure: Procedure = { kind, name, visibility, isStatic, params, body, line };
    if (returnType !== undefined) procedure.returnType = returnType;
    return procedure;
  }

  private atEndOf(word: string): boolean {
    return (
      this.current.type === Tk.Keyword &&
      this.current.upper === 'END' &&
      this.tokens[this.i + 1]?.upper === word
    );
  }

  private parseParameterList(): Parameter[] {
    this.expect(Tk.LParen, '(');
    const params: Parameter[] = [];
    while (!this.at(Tk.RParen) && this.current.type !== Tk.EOF) {
      let optional = false;
      let byRef = true; // VBA passes by reference unless told otherwise
      let paramArray = false;
      while (this.atKeyword('OPTIONAL', 'BYVAL', 'BYREF', 'PARAMARRAY')) {
        const word = this.advance().upper;
        if (word === 'OPTIONAL') optional = true;
        else if (word === 'BYVAL') byRef = false;
        else if (word === 'BYREF') byRef = true;
        else paramArray = true;
      }
      const name = this.advance().text;
      // An array parameter is written `values()`.
      if (this.at(Tk.LParen)) {
        this.advance();
        this.expect(Tk.RParen, ')');
      }
      const param: Parameter = { name, optional, byRef, paramArray };
      if (this.atKeyword('AS')) {
        this.advance();
        param.typeName = this.parseTypeName();
      }
      if (this.at(Tk.Operator, '=')) {
        this.advance();
        param.defaultValue = this.parseExpression();
      }
      params.push(param);
      if (this.at(Tk.Comma)) this.advance();
      else break;
    }
    this.expect(Tk.RParen, ')');
    return params;
  }

  private parseTypeName(): string {
    let name = this.advance().text;
    // A qualified type such as `Excel.Range`.
    while (this.at(Tk.Dot)) {
      this.advance();
      name += `.${this.advance().text}`;
    }
    return name;
  }

  private parseBlock(stop: () => boolean): Stmt[] {
    const body: Stmt[] = [];
    this.skipTerminators();
    while (this.current.type !== Tk.EOF && !stop()) {
      try {
        body.push(this.parseStatement());
      } catch (err) {
        body.push({ kind: 'unknown', text: err instanceof Error ? err.message : String(err) });
        this.recover();
      }
      this.skipTerminators();
    }
    return body;
  }

  private parseStatement(): Stmt {
    const t = this.current;

    if (t.type === Tk.Attribute) return { kind: 'attribute', text: this.advance().text };
    if (t.type === Tk.Directive) return { kind: 'directive', text: this.advance().text };
    if (t.type === Tk.Label) return { kind: 'label', name: this.advance().text };

    if (t.type === Tk.Keyword) {
      switch (t.upper) {
        case 'OPTION':
          return this.parseOption();
        case 'DIM':
        case 'PRIVATE':
        case 'PUBLIC':
        case 'GLOBAL':
        case 'STATIC':
        case 'FRIEND':
          return this.parseDeclaration();
        case 'CONST':
          return this.parseConst('Dim');
        case 'REDIM':
          return this.parseRedim();
        case 'SET':
          return this.parseAssignment(true);
        case 'LET':
          this.advance();
          return this.parseAssignment(false);
        case 'CALL': {
          this.advance();
          const expr = this.parseExpression();
          return expr.kind === 'call'
            ? { kind: 'call', target: expr.target, args: expr.args }
            : { kind: 'call', target: expr, args: [] };
        }
        case 'IF':
          return this.parseIf();
        case 'FOR':
          return this.parseFor();
        case 'DO':
          return this.parseDo();
        case 'WHILE':
          return this.parseWhile();
        case 'SELECT':
          return this.parseSelectCase();
        case 'WITH':
          return this.parseWith();
        case 'EXIT': {
          this.advance();
          return { kind: 'exit', target: this.advance().upper };
        }
        case 'GOTO': {
          this.advance();
          return { kind: 'goto', label: this.advance().text };
        }
        case 'GOSUB': {
          this.advance();
          return { kind: 'gosub', label: this.advance().text };
        }
        case 'RETURN':
          this.advance();
          return { kind: 'return' };
        case 'ON':
          return this.parseOnError();
        case 'RESUME': {
          this.advance();
          if (this.at(Tk.Terminator)) return { kind: 'resume' };
          return { kind: 'resume', target: this.advance().text };
        }
        case 'ERASE': {
          this.advance();
          const targets: Expr[] = [this.parseExpression()];
          while (this.at(Tk.Comma)) {
            this.advance();
            targets.push(this.parseExpression());
          }
          return { kind: 'erase', targets };
        }
        case 'STOP':
          this.advance();
          return { kind: 'stop' };
        case 'END':
          this.advance();
          return { kind: 'end' };
        case 'RAISEEVENT': {
          this.advance();
          const name = this.advance().text;
          const args = this.at(Tk.LParen) ? this.parseArguments(true) : [];
          return { kind: 'raiseEvent', name, args };
        }
        case 'DECLARE':
          return this.parseDeclare();
        case 'TYPE':
          return this.parseUserType();
        case 'ENUM':
          return this.parseEnum();
        case 'IMPLEMENTS': {
          this.advance();
          return { kind: 'unknown', text: `Implements ${this.advance().text}` };
        }
        default:
          break;
      }
    }

    // Anything else is an assignment or a bare procedure call.
    return this.parseAssignment(false);
  }

  private parseOption(): Stmt {
    const parts: string[] = [];
    while (!this.at(Tk.Terminator) && this.current.type !== Tk.EOF) parts.push(this.advance().text);
    return { kind: 'option', text: parts.join(' ') };
  }

  private parseDeclaration(): Stmt {
    const scope = this.advance().upper;
    const scopeName = (scope.charAt(0) + scope.slice(1).toLowerCase()) as DeclScope;

    // `Public Const`, `Private Const` and friends.
    if (this.atKeyword('CONST')) return this.parseConst(scopeName);
    // `Private Declare Function ...`
    if (this.atKeyword('DECLARE')) return this.parseDeclare();
    if (this.atKeyword('TYPE')) return this.parseUserType();
    if (this.atKeyword('ENUM')) return this.parseEnum();

    return { kind: 'dim', scope: scopeName, vars: this.parseVarList() };
  }

  private parseVarList(): VarDecl[] {
    const vars: VarDecl[] = [];
    for (;;) {
      let withEvents = false;
      if (this.atKeyword('WITHEVENTS')) {
        withEvents = true;
        this.advance();
      }
      const name = this.advance().text;
      const decl: VarDecl = { name };
      if (withEvents) decl.withEvents = true;

      if (this.at(Tk.LParen)) {
        decl.isArray = true;
        this.advance();
        const bounds: { lower?: Expr; upper: Expr }[] = [];
        while (!this.at(Tk.RParen) && this.current.type !== Tk.EOF) {
          const first = this.parseExpression();
          if (this.atKeyword('TO')) {
            this.advance();
            bounds.push({ lower: first, upper: this.parseExpression() });
          } else {
            bounds.push({ upper: first });
          }
          if (this.at(Tk.Comma)) this.advance();
          else break;
        }
        this.expect(Tk.RParen, ')');
        if (bounds.length > 0) decl.bounds = bounds;
      }

      if (this.atKeyword('AS')) {
        this.advance();
        if (this.atKeyword('NEW')) {
          decl.isNew = true;
          this.advance();
        }
        decl.typeName = this.parseTypeName();
      }
      vars.push(decl);
      if (this.at(Tk.Comma)) {
        this.advance();
        continue;
      }
      break;
    }
    return vars;
  }

  private parseConst(scope: DeclScope): Stmt {
    this.advance(); // Const
    const vars: ConstDecl[] = [];
    for (;;) {
      const name = this.advance().text;
      let typeName: string | undefined;
      if (this.atKeyword('AS')) {
        this.advance();
        typeName = this.parseTypeName();
      }
      if (this.at(Tk.Operator, '=')) this.advance();
      const value = this.parseExpression();
      vars.push(typeName === undefined ? { name, value } : { name, typeName, value });
      if (this.at(Tk.Comma)) {
        this.advance();
        continue;
      }
      break;
    }
    return { kind: 'const', scope, vars };
  }

  private parseRedim(): Stmt {
    this.advance();
    let preserve = false;
    if (this.atKeyword('PRESERVE')) {
      preserve = true;
      this.advance();
    }
    return { kind: 'redim', preserve, vars: this.parseVarList() };
  }

  private parseAssignment(isSet: boolean): Stmt {
    if (isSet) this.advance();
    const target = this.parseExpression();

    if (this.at(Tk.Operator, '=')) {
      this.advance();
      // Parse the value above precedence 0 so that a further `=` is read as
      // comparison rather than ending the expression: `x = a & b = c` assigns
      // the boolean result, it is not two assignments.
      return { kind: 'assign', target, value: this.parseExpression(1), isSet };
    }

    // A bare call with unparenthesised arguments: `MsgBox "hi", vbOK`.
    if (!this.at(Tk.Terminator) && this.current.type !== Tk.EOF) {
      const args = this.parseArguments(false);
      return { kind: 'call', target, args };
    }

    if (target.kind === 'call') return { kind: 'call', target: target.target, args: target.args };
    return { kind: 'call', target, args: [] };
  }

  private parseIf(): Stmt {
    this.advance(); // If
    const condition = this.parseExpression();
    if (this.atKeyword('THEN')) this.advance();

    // A single-line If has its statement on the same line: `If x Then y = 1`.
    if (!this.at(Tk.Terminator)) {
      const body: Stmt[] = [this.parseStatement()];
      const branches = [{ condition, body }];
      if (this.atKeyword('ELSE')) {
        this.advance();
        return { kind: 'if', branches, elseBody: [this.parseStatement()] };
      }
      return { kind: 'if', branches };
    }

    const branches: { condition: Expr; body: Stmt[] }[] = [];
    let elseBody: Stmt[] | undefined;
    let currentCondition = condition;

    for (;;) {
      const body = this.parseBlock(
        () => this.atKeyword('ELSEIF', 'ELSE') || this.atEndOf('IF'),
      );
      branches.push({ condition: currentCondition, body });

      if (this.atKeyword('ELSEIF')) {
        this.advance();
        currentCondition = this.parseExpression();
        if (this.atKeyword('THEN')) this.advance();
        continue;
      }
      if (this.atKeyword('ELSE')) {
        this.advance();
        elseBody = this.parseBlock(() => this.atEndOf('IF'));
      }
      break;
    }

    if (this.atEndOf('IF')) {
      this.advance();
      this.advance();
    }
    return elseBody === undefined ? { kind: 'if', branches } : { kind: 'if', branches, elseBody };
  }

  private parseFor(): Stmt {
    this.advance(); // For
    if (this.atKeyword('EACH')) {
      this.advance();
      const variable = this.parseExpression();
      if (this.atKeyword('IN')) this.advance();
      const collection = this.parseExpression();
      const body = this.parseBlock(() => this.atKeyword('NEXT'));
      if (this.atKeyword('NEXT')) {
        this.advance();
        if (!this.at(Tk.Terminator)) this.advance(); // optional loop variable
      }
      return { kind: 'forEach', variable, collection, body };
    }

    const variable = this.parseExpression();
    if (this.at(Tk.Operator, '=')) this.advance();
    const from = this.parseExpression();
    if (this.atKeyword('TO')) this.advance();
    const to = this.parseExpression();
    let step: Expr | undefined;
    if (this.atKeyword('STEP')) {
      this.advance();
      step = this.parseExpression();
    }
    const body = this.parseBlock(() => this.atKeyword('NEXT'));
    if (this.atKeyword('NEXT')) {
      this.advance();
      while (!this.at(Tk.Terminator) && this.current.type !== Tk.EOF) this.advance();
    }
    return step === undefined
      ? { kind: 'for', variable, from, to, body }
      : { kind: 'for', variable, from, to, step, body };
  }

  private parseDo(): Stmt {
    this.advance(); // Do
    // The test may sit at the top or the bottom, which changes whether the body
    // is guaranteed to run once.
    let test: { kind: 'while' | 'until'; condition: Expr; post: boolean } | undefined;
    if (this.atKeyword('WHILE', 'UNTIL')) {
      const which = this.advance().upper === 'WHILE' ? 'while' : 'until';
      test = { kind: which, condition: this.parseExpression(), post: false };
    }
    const body = this.parseBlock(() => this.atKeyword('LOOP'));
    if (this.atKeyword('LOOP')) {
      this.advance();
      if (this.atKeyword('WHILE', 'UNTIL')) {
        const which = this.advance().upper === 'WHILE' ? 'while' : 'until';
        test = { kind: which, condition: this.parseExpression(), post: true };
      }
    }
    return test === undefined ? { kind: 'do', body } : { kind: 'do', test, body };
  }

  private parseWhile(): Stmt {
    this.advance();
    const condition = this.parseExpression();
    const body = this.parseBlock(() => this.atKeyword('WEND'));
    if (this.atKeyword('WEND')) this.advance();
    return { kind: 'while', condition, body };
  }

  private parseSelectCase(): Stmt {
    this.advance(); // Select
    if (this.atKeyword('CASE')) this.advance();
    const subject = this.parseExpression();
    this.skipTerminators();

    const cases: CaseClause[] = [];
    let elseBody: Stmt[] | undefined;

    while (this.atKeyword('CASE')) {
      this.advance();
      if (this.atKeyword('ELSE')) {
        this.advance();
        elseBody = this.parseBlock(() => this.atKeyword('CASE') || this.atEndOf('SELECT'));
        break;
      }
      const labels: CaseLabel[] = [];
      for (;;) {
        if (this.atKeyword('IS')) {
          this.advance();
          const op = this.advance().text;
          labels.push({ kind: 'compare', op, value: this.parseExpression() });
        } else {
          const first = this.parseExpression();
          if (this.atKeyword('TO')) {
            this.advance();
            labels.push({ kind: 'range', from: first, to: this.parseExpression() });
          } else {
            labels.push({ kind: 'value', value: first });
          }
        }
        if (this.at(Tk.Comma)) {
          this.advance();
          continue;
        }
        break;
      }
      const body = this.parseBlock(() => this.atKeyword('CASE') || this.atEndOf('SELECT'));
      cases.push({ labels, body });
    }

    if (this.atEndOf('SELECT')) {
      this.advance();
      this.advance();
    }
    return elseBody === undefined
      ? { kind: 'selectCase', subject, cases }
      : { kind: 'selectCase', subject, cases, elseBody };
  }

  private parseWith(): Stmt {
    this.advance();
    const subject = this.parseExpression();
    const body = this.parseBlock(() => this.atEndOf('WITH'));
    if (this.atEndOf('WITH')) {
      this.advance();
      this.advance();
    }
    return { kind: 'with', subject, body };
  }

  private parseOnError(): Stmt {
    this.advance(); // On
    if (this.atKeyword('ERROR')) {
      this.advance();
      if (this.atKeyword('RESUME')) {
        this.advance();
        this.advance(); // Next
        return { kind: 'onError', mode: 'resumeNext' };
      }
      if (this.atKeyword('GOTO')) {
        this.advance();
        const label = this.advance().text;
        return label === '0' ? { kind: 'onError', mode: 'goto0' } : { kind: 'onError', mode: 'goto', label };
      }
    }
    // `On x GoTo a, b, c` - a computed jump, rare and not modelled.
    const parts: string[] = ['On'];
    while (!this.at(Tk.Terminator) && this.current.type !== Tk.EOF) parts.push(this.advance().text);
    return { kind: 'unknown', text: parts.join(' ') };
  }

  private parseDeclare(): Stmt {
    this.advance(); // Declare
    if (this.atKeyword('PTRSAFE') || this.current.upper === 'PTRSAFE') this.advance();
    const isFunction = this.advance().upper === 'FUNCTION';
    const name = this.advance().text;
    let lib = '';
    let alias: string | undefined;
    if (this.atKeyword('LIB')) {
      this.advance();
      lib = stripQuotes(this.advance().text);
    }
    if (this.current.upper === 'ALIAS') {
      this.advance();
      alias = stripQuotes(this.advance().text);
    }
    if (this.at(Tk.LParen)) this.parseParameterList();
    if (this.atKeyword('AS')) {
      this.advance();
      this.parseTypeName();
    }
    return alias === undefined
      ? { kind: 'declare', name, lib, isFunction }
      : { kind: 'declare', name, lib, alias, isFunction };
  }

  private parseUserType(): Stmt {
    this.advance(); // Type
    const name = this.advance().text;
    this.skipTerminators();
    const fields: VarDecl[] = [];
    while (!this.atEndOf('TYPE') && this.current.type !== Tk.EOF) {
      fields.push(...this.parseVarList());
      this.skipTerminators();
    }
    if (this.atEndOf('TYPE')) {
      this.advance();
      this.advance();
    }
    return { kind: 'type', name, fields };
  }

  private parseEnum(): Stmt {
    this.advance(); // Enum
    const name = this.advance().text;
    this.skipTerminators();
    const members: { name: string; value?: Expr }[] = [];
    while (!this.atEndOf('ENUM') && this.current.type !== Tk.EOF) {
      const memberName = this.advance().text;
      if (this.at(Tk.Operator, '=')) {
        this.advance();
        members.push({ name: memberName, value: this.parseExpression() });
      } else {
        members.push({ name: memberName });
      }
      this.skipTerminators();
    }
    if (this.atEndOf('ENUM')) {
      this.advance();
      this.advance();
    }
    return { kind: 'enum', name, members };
  }

  // --- expression parsing ------------------------------------------------

  parseExpression(minPrecedence = 0): Expr {
    let left = this.parseUnary();

    for (;;) {
      const t = this.current;
      const key = t.type === Tk.Keyword ? t.upper : t.text;
      const precedence = PRECEDENCE[key];
      if (precedence === undefined || precedence < minPrecedence) break;
      // `=` inside a statement context is assignment, not comparison; the
      // statement parser handles that before calling us at precedence 0.
      if (key === '=' && minPrecedence === 0) break;
      this.advance();
      // Exponentiation is the only right-associative operator.
      const next = key === '^' ? precedence : precedence + 1;
      left = { kind: 'binary', op: key, left, right: this.parseExpression(next) };
    }
    return left;
  }

  private parseUnary(): Expr {
    if (this.atKeyword('NOT')) {
      this.advance();
      return { kind: 'unary', op: 'Not', operand: this.parseExpression(NOT_PRECEDENCE) };
    }
    if (this.at(Tk.Operator, '-')) {
      this.advance();
      return { kind: 'unary', op: '-', operand: this.parseExpression(UNARY_PRECEDENCE) };
    }
    if (this.at(Tk.Operator, '+')) {
      this.advance();
      return this.parseExpression(UNARY_PRECEDENCE);
    }
    if (this.atKeyword('NEW')) {
      this.advance();
      return { kind: 'new', typeName: this.parseTypeName() };
    }
    return this.parsePostfix();
  }

  private parsePostfix(): Expr {
    let expr = this.parsePrimary();
    for (;;) {
      if (this.at(Tk.Dot)) {
        this.advance();
        expr = { kind: 'member', target: expr, name: this.advance().text };
        continue;
      }
      if (this.at(Tk.LParen)) {
        // VBA does not distinguish a call from an array index syntactically;
        // only the resolved target tells them apart, which is a runtime matter.
        expr = { kind: 'call', target: expr, args: this.parseArguments(true) };
        continue;
      }
      break;
    }
    return expr;
  }

  private parseArguments(parenthesised: boolean): Argument[] {
    const args: Argument[] = [];
    if (parenthesised) this.expect(Tk.LParen, '(');

    const done = () =>
      parenthesised ? this.at(Tk.RParen) : this.at(Tk.Terminator) || this.current.type === Tk.EOF;

    if (done()) {
      if (parenthesised) this.advance();
      return args;
    }

    for (;;) {
      // An omitted argument, as in `Foo(1, , 3)`.
      if (this.at(Tk.Comma)) {
        args.push({ value: null });
        this.advance();
        continue;
      }
      if (done()) break;

      const arg: Argument = { value: null };
      // A named argument uses `:=`.
      if (this.current.type === Tk.Identifier && this.tokens[this.i + 1]?.text === ':=') {
        arg.name = this.advance().text;
        this.advance();
      }
      if (this.atKeyword('BYVAL', 'BYREF')) {
        arg.modifier = this.advance().upper === 'BYVAL' ? 'ByVal' : 'ByRef';
      }
      arg.value = this.parseExpression();
      args.push(arg);

      if (this.at(Tk.Comma)) {
        this.advance();
        continue;
      }
      break;
    }

    if (parenthesised) this.expect(Tk.RParen, ')');
    return args;
  }

  private parsePrimary(): Expr {
    const t = this.current;

    switch (t.type) {
      case Tk.Number: {
        this.advance();
        return parseNumberLiteral(t.text);
      }
      case Tk.String:
        this.advance();
        return { kind: 'string', value: stripQuotes(t.text) };
      case Tk.DateLiteral:
        this.advance();
        return { kind: 'date', value: t.text.slice(1, -1) };
      case Tk.Identifier:
        this.advance();
        return { kind: 'identifier', name: t.text };
      case Tk.LParen: {
        this.advance();
        const inner = this.parseExpression();
        this.expect(Tk.RParen, ')');
        return { kind: 'paren', inner };
      }
      case Tk.Dot: {
        // A leading dot inside a With block: `.Value = 1`.
        this.advance();
        return { kind: 'member', target: null, name: this.advance().text };
      }
      case Tk.Keyword:
        switch (t.upper) {
          case 'TRUE':
            this.advance();
            return { kind: 'bool', value: true };
          case 'FALSE':
            this.advance();
            return { kind: 'bool', value: false };
          case 'NOTHING':
            this.advance();
            return { kind: 'nothing' };
          case 'NULL':
            this.advance();
            return { kind: 'null' };
          case 'ME':
            this.advance();
            return { kind: 'me' };
          default:
            // A keyword used as a name, which VBA permits in several places.
            this.advance();
            return { kind: 'identifier', name: t.text };
        }
      default:
        throw new VbaParseError(`unexpected ${JSON.stringify(t.text)}`, t.line);
    }
  }
}

function stripQuotes(text: string): string {
  if (text.startsWith('"') && text.endsWith('"') && text.length >= 2) {
    return text.slice(1, -1).replaceAll('""', '"');
  }
  return text;
}

export function parseNumberLiteral(text: string): Expr {
  const suffixChar = /[%&@!#$]$/.exec(text)?.[0];
  const body = suffixChar ? text.slice(0, -1) : text;

  let value: number;
  if (/^&H/i.test(body)) value = Number.parseInt(body.slice(2).replace(/&$/, ''), 16);
  else if (/^&O/i.test(body)) value = Number.parseInt(body.slice(2).replace(/&$/, ''), 8);
  // VBA writes double-precision exponents with D rather than E.
  else value = Number(body.replace(/[dD]/, 'e'));

  return suffixChar === undefined
    ? { kind: 'number', value }
    : { kind: 'number', value, suffix: suffixChar };
}
