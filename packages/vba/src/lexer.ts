/**
 * VBA tokenizer, with the preprocessing the language requires before a parser
 * can see a sensible token stream.
 *
 * VBA is line-oriented in a way most languages are not, and four transformations
 * have to happen before parsing or the grammar becomes unmanageable:
 *
 *   A trailing space-underscore continues a logical line onto the next physical
 *   one. Statements are otherwise terminated by end of line, so missing this
 *   turns one statement into two malformed ones.
 *
 *   A colon separates statements on one line, so `x = 1: y = 2` is two
 *   statements and `Label:` is not.
 *
 *   Comments start with an apostrophe or the `Rem` keyword and run to end of
 *   line, but an apostrophe inside a string literal is just an apostrophe.
 *
 *   Module headers carry `Attribute` lines the VBE writes and hides. They are
 *   real syntax, not comments, and must be recognised rather than tripped over.
 *
 * Everything in VBA is case-insensitive, so keyword matching works on an
 * upper-cased copy while the original spelling is preserved for display.
 */

export const enum Tk {
  Identifier,
  Keyword,
  Number,
  String,
  /** A date literal delimited by hashes: `#1/1/2024#`. */
  DateLiteral,
  Operator,
  /** Statement separator: end of line or a colon. */
  Terminator,
  LParen,
  RParen,
  Comma,
  Dot,
  /** A line label such as `ErrorHandler:`. */
  Label,
  Comment,
  /** `Attribute VB_Name = "Module1"` and friends. */
  Attribute,
  /** A conditional-compilation directive: `#If`, `#Const`, `#Else`, `#End If`. */
  Directive,
  EOF,
}

export interface Token {
  type: Tk;
  /** Source text, with the author's original casing. */
  text: string;
  /** Upper-cased text, for case-insensitive matching. */
  upper: string;
  line: number;
  column: number;
}

/**
 * VBA's reserved words. Case-insensitive, and deliberately complete for the
 * statement forms we parse - an unrecognised keyword becomes an identifier,
 * which is how a program using a name we do not know still parses.
 */
export const KEYWORDS: ReadonlySet<string> = new Set([
  'AND', 'AS', 'BOOLEAN', 'BYREF', 'BYTE', 'BYVAL', 'CALL', 'CASE', 'CONST',
  'CURRENCY', 'DATE', 'DECLARE', 'DEFBOOL', 'DEFBYTE', 'DEFCUR', 'DEFDATE',
  'DEFDBL', 'DEFINT', 'DEFLNG', 'DEFOBJ', 'DEFSNG', 'DEFSTR', 'DEFVAR', 'DIM',
  'DO', 'DOUBLE', 'EACH', 'ELSE', 'ELSEIF', 'END', 'ENUM', 'EQV', 'ERASE',
  'ERROR', 'EVENT', 'EXIT', 'FALSE', 'FOR', 'FRIEND', 'FUNCTION', 'GET',
  'GLOBAL', 'GOSUB', 'GOTO', 'IF', 'IMP', 'IMPLEMENTS', 'IN', 'INTEGER', 'IS',
  'LET', 'LIB', 'LIKE', 'LONG', 'LOOP', 'LSET', 'ME', 'MOD', 'NEW', 'NEXT',
  'NOT', 'NOTHING', 'NULL', 'OBJECT', 'ON', 'OPTION', 'OPTIONAL', 'OR',
  'PARAMARRAY', 'PRESERVE', 'PRIVATE', 'PROPERTY', 'PUBLIC', 'RAISEEVENT',
  'REDIM', 'REM', 'RESUME', 'RETURN', 'RSET', 'SELECT', 'SET', 'SINGLE',
  'STATIC', 'STEP', 'STOP', 'STRING', 'SUB', 'THEN', 'TO', 'TRUE', 'TYPE',
  'UNTIL', 'VARIANT', 'WEND', 'WHILE', 'WITH', 'WITHEVENTS', 'XOR',
]);

/** Type-declaration suffixes: `count%` is an Integer, `name$` a String. */
export const TYPE_SUFFIXES: Readonly<Record<string, string>> = {
  '%': 'Integer',
  '&': 'Long',
  '@': 'Currency',
  '!': 'Single',
  '#': 'Double',
  $: 'String',
};

const TWO_CHAR_OPERATORS = ['<=', '>=', '<>', ':='] as const;
const SINGLE_CHAR_OPERATORS = '+-*/\\^&<>=';

export class VbaLexError extends Error {
  constructor(
    message: string,
    readonly line: number,
  ) {
    super(`${message} (line ${line})`);
    this.name = 'VbaLexError';
  }
}

/**
 * Join continued lines, so the tokenizer sees logical lines.
 *
 * Returns the joined text plus a map from joined-line index back to the original
 * physical line, so errors still point at a line the user can find.
 */
export function joinContinuations(source: string): { lines: string[]; origins: number[] } {
  const physical = source.split(/\r\n|\r|\n/);
  const lines: string[] = [];
  const origins: number[] = [];

  let buffer = '';
  let start = 0;
  for (let i = 0; i < physical.length; i++) {
    const line = physical[i]!;
    if (buffer === '') start = i;
    // A continuation is a space followed by an underscore at end of line. The
    // space is required, which is what stops an identifier ending in underscore
    // from being read as a continuation.
    const continued = /(^|\s)_[ \t]*$/.test(line);
    if (continued) {
      // Collapse to exactly one space: the separator must exist so two tokens
      // do not merge across the join, but its width should not depend on how
      // the author happened to indent.
      buffer += `${line.replace(/\s*_[ \t]*$/, '')} `;
      continue;
    }
    lines.push(buffer + line);
    origins.push(start + 1);
    buffer = '';
  }
  if (buffer !== '') {
    lines.push(buffer);
    origins.push(start + 1);
  }
  return { lines, origins };
}

export function tokenize(source: string): Token[] {
  const { lines, origins } = joinContinuations(source);
  const tokens: Token[] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]!;
    const lineNo = origins[i]!;
    tokenizeLine(line, lineNo, tokens);
    // Every logical line ends a statement, unless it produced nothing at all.
    const last = tokens[tokens.length - 1];
    if (last && last.type !== Tk.Terminator) {
      tokens.push(term(lineNo, line.length));
    }
  }
  tokens.push({ type: Tk.EOF, text: '', upper: '', line: lines.length + 1, column: 0 });
  return tokens;
}

function term(line: number, column: number): Token {
  return { type: Tk.Terminator, text: '\n', upper: '\n', line, column };
}

function tokenizeLine(line: string, lineNo: number, out: Token[]): void {
  let p = 0;
  const push = (type: Tk, text: string, column: number) => {
    out.push({ type, text, upper: text.toUpperCase(), line: lineNo, column });
  };

  // A leading Attribute line is module metadata the VBE maintains.
  const trimmed = line.trimStart();
  if (/^Attribute\s/i.test(trimmed)) {
    push(Tk.Attribute, trimmed, line.length - trimmed.length);
    push(Tk.Terminator, '\n', line.length);
    return;
  }
  // Conditional compilation. We record the directive rather than evaluating it;
  // a static analysis has to see both branches anyway.
  if (/^#(If|ElseIf|Else|End\s+If|Const)\b/i.test(trimmed)) {
    push(Tk.Directive, trimmed, line.length - trimmed.length);
    push(Tk.Terminator, '\n', line.length);
    return;
  }

  while (p < line.length) {
    const ch = line[p]!;

    if (ch === ' ' || ch === '\t') {
      p++;
      continue;
    }

    // Comments run to end of line. `Rem` only counts as a comment when it is a
    // whole word, so a variable called `Remainder` is safe.
    if (ch === "'" || /^rem(\s|$)/i.test(line.slice(p))) {
      push(Tk.Comment, line.slice(p), p);
      break;
    }

    if (ch === '"') {
      const start = p;
      p++;
      while (p < line.length) {
        if (line[p] === '"') {
          // A doubled quote is an escaped quote.
          if (line[p + 1] === '"') {
            p += 2;
            continue;
          }
          p++;
          break;
        }
        p++;
      }
      push(Tk.String, line.slice(start, p), start);
      continue;
    }

    // A date literal is delimited by hashes. A bare `#` elsewhere is the Double
    // type suffix, so only a closing hash later on the line makes this a date.
    if (ch === '#') {
      const close = line.indexOf('#', p + 1);
      if (close > p) {
        push(Tk.DateLiteral, line.slice(p, close + 1), p);
        p = close + 1;
        continue;
      }
    }

    if (ch === ':' && line[p + 1] === '=') {
      p += 2;
      push(Tk.Operator, ':=', p - 2);
      continue;
    }

    if (ch === ':') {
      // A colon after a bare identifier at the start of a statement is a line
      // label; otherwise it separates statements.
      const previous = out[out.length - 1];
      const atStatementStart =
        out.length >= 1 &&
        previous?.type === Tk.Identifier &&
        (out.length === 1 || out[out.length - 2]?.type === Tk.Terminator);
      if (atStatementStart) {
        previous.type = Tk.Label;
        p++;
        push(Tk.Terminator, ':', p);
        continue;
      }
      p++;
      push(Tk.Terminator, ':', p);
      continue;
    }

    if (ch === '(') {
      p++;
      push(Tk.LParen, '(', p - 1);
      continue;
    }
    if (ch === ')') {
      p++;
      push(Tk.RParen, ')', p - 1);
      continue;
    }
    if (ch === ',') {
      p++;
      push(Tk.Comma, ',', p - 1);
      continue;
    }
    if (ch === '.') {
      // A leading dot is member access; a dot between digits belongs to a number.
      if (!/[0-9]/.test(line[p + 1] ?? '')) {
        p++;
        push(Tk.Dot, '.', p - 1);
        continue;
      }
    }

    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(line[p + 1] ?? ''))) {
      const start = p;
      // Hex and octal literals use &H and &O prefixes.
      if (ch === '0' && /[&]/.test(line[p + 1] ?? '')) {
        // handled below by the ampersand branch
      }
      while (p < line.length && /[0-9]/.test(line[p]!)) p++;
      if (line[p] === '.') {
        p++;
        while (p < line.length && /[0-9]/.test(line[p]!)) p++;
      }
      if (/[eEdD]/.test(line[p] ?? '')) {
        const save = p;
        p++;
        if (line[p] === '+' || line[p] === '-') p++;
        if (/[0-9]/.test(line[p] ?? '')) {
          while (p < line.length && /[0-9]/.test(line[p]!)) p++;
        } else {
          p = save;
        }
      }
      // A type suffix binds to the literal.
      if (line[p] !== undefined && TYPE_SUFFIXES[line[p]!] !== undefined) p++;
      push(Tk.Number, line.slice(start, p), start);
      continue;
    }

    // &H and &O introduce hex and octal literals; a bare & is concatenation.
    if (ch === '&' && /[hHoO]/.test(line[p + 1] ?? '')) {
      const start = p;
      p += 2;
      while (p < line.length && /[0-9a-fA-F]/.test(line[p]!)) p++;
      if (line[p] === '&') p++; // optional Long marker
      push(Tk.Number, line.slice(start, p), start);
      continue;
    }

    if (/[A-Za-z_]/.test(ch) || ch.charCodeAt(0) > 127) {
      const start = p;
      while (p < line.length && (/[A-Za-z0-9_]/.test(line[p]!) || line[p]!.charCodeAt(0) > 127)) p++;
      // A trailing type suffix is part of the name for declaration purposes.
      if (line[p] !== undefined && TYPE_SUFFIXES[line[p]!] !== undefined) p++;
      const text = line.slice(start, p);
      const bare = text.replace(/[%&@!#$]$/, '').toUpperCase();
      push(KEYWORDS.has(bare) ? Tk.Keyword : Tk.Identifier, text, start);
      continue;
    }

    // A bracketed identifier: [My Name] allows spaces and reserved words.
    if (ch === '[') {
      const close = line.indexOf(']', p);
      if (close > p) {
        push(Tk.Identifier, line.slice(p, close + 1), p);
        p = close + 1;
        continue;
      }
    }

    const two = line.slice(p, p + 2);
    if ((TWO_CHAR_OPERATORS as readonly string[]).includes(two)) {
      p += 2;
      push(Tk.Operator, two, p - 2);
      continue;
    }
    if (SINGLE_CHAR_OPERATORS.includes(ch)) {
      p++;
      push(Tk.Operator, ch, p - 1);
      continue;
    }

    throw new VbaLexError(`unexpected character ${JSON.stringify(ch)}`, lineNo);
  }
}

/** Strip comments and blank statements, which a parser does not need. */
export function significant(tokens: readonly Token[]): Token[] {
  const out: Token[] = [];
  for (const t of tokens) {
    if (t.type === Tk.Comment) continue;
    // Collapse runs of terminators so an empty line does not become a statement.
    if (t.type === Tk.Terminator && out[out.length - 1]?.type === Tk.Terminator) continue;
    out.push(t);
  }
  return out;
}
