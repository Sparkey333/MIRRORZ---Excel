/**
 * Pratt parser for Excel formulas.
 *
 * Pratt parsing suits this grammar because Excel's operator set is large, mostly
 * infix, and has a couple of precedence rules that a recursive-descent cascade
 * would need an extra level for each of. The binding powers live in ast.ts and
 * are verified against the fixture oracle, not assumed.
 *
 * Three things the parser has to get right that a naive implementation misses:
 *
 *   - A space between two references is the intersection operator, so
 *     whitespace tokens are consulted rather than skipped when deciding whether
 *     an infix operator follows.
 *   - A comma is the union operator inside parentheses but an argument
 *     separator inside a call. That is context, not grammar, so the parser
 *     carries a flag rather than duplicating the expression rules.
 *   - Unary minus binds tighter than `^`, which means the prefix parser must
 *     parse its operand at a binding power ABOVE the exponent operator's.
 */

import { CellError, type ErrorCode, errorFromCode, nameToCol, unquoteSheetName } from '@mirrorz/core';
import {
  type Ast,
  type BeamNode,
  type BinaryOp,
  BINDING_POWER,
  Node,
  POSTFIX_BINDING_POWER,
  PREFIX_BINDING_POWER,
  type PostfixOp,
  type RefNode,
  type UnaryOp,
} from './ast.js';
import { Lexer, Tok, type Token, isColumnShape, tokenText } from './lexer.js';

export class ParseError extends Error {
  constructor(
    message: string,
    readonly offset: number,
  ) {
    super(`${message} (at offset ${offset})`);
    this.name = 'ParseError';
  }
}

export interface ParseOptions {
  /** The address the formula lives at, so references can be stored relatively. */
  origin?: { row: number; col: number };
  /** Argument separator; `;` in many European locales. */
  argSeparator?: ',' | ';';
}

interface Context {
  sheet?: string;
  sheetEnd?: string;
  book?: string;
}

export function parseFormula(src: string, options: ParseOptions = {}): Ast {
  return new Parser(src, options).parse();
}

class Parser {
  private readonly tokens: Token[];
  private i = 0;
  private readonly origin: { row: number; col: number };
  private readonly argSeparator: ',' | ';';
  /** Depth of parentheses that are grouping rather than call arguments. */
  private groupDepth = 0;

  constructor(
    private readonly src: string,
    options: ParseOptions,
  ) {
    this.tokens = Lexer.tokenize(src);
    this.origin = options.origin ?? { row: 0, col: 0 };
    this.argSeparator = options.argSeparator ?? ',';
  }

  parse(): Ast {
    // A leading `=` is optional; formulas stored in xlsx omit it.
    if (this.peek()?.type === Tok.Equals) this.i++;
    const ast = this.expression(0);
    this.skipWhitespace();
    const rest = this.peek();
    if (rest && rest.type !== Tok.EOF) {
      throw new ParseError(`unexpected ${JSON.stringify(tokenText(this.src, rest))}`, rest.start);
    }
    return ast;
  }

  private peek(offset = 0): Token | undefined {
    return this.tokens[this.i + offset];
  }

  private skipWhitespace(): string {
    let ws = '';
    while (this.peek()?.type === Tok.Whitespace) {
      ws += tokenText(this.src, this.tokens[this.i]!);
      this.i++;
    }
    return ws;
  }

  /**
   * Core Pratt loop: parse a prefix expression, then absorb infix and postfix
   * operators whose binding power exceeds `minBp`.
   */
  private expression(minBp: number): Ast {
    const ws = this.skipWhitespace();
    let left = this.prefix();
    if (ws) left = { ...left, ws };

    for (;;) {
      // Whitespace is only consumed here once we know it is not an intersection
      // operator, so look ahead across it without committing.
      const save = this.i;
      const gap = this.skipWhitespace();
      const t = this.peek();
      if (!t || t.type === Tok.EOF) {
        this.i = save;
        break;
      }

      // Intersection: whitespace directly between two reference-valued operands.
      if (gap && this.startsOperand(t)) {
        if (BINDING_POWER[' ']! <= minBp) {
          this.i = save;
          break;
        }
        const right = this.expression(BINDING_POWER[' ']! + 1);
        left = { kind: Node.Binary, op: ' ', left, right, wsOp: gap };
        continue;
      }

      if (t.type === Tok.Operator) {
        const op = tokenText(this.src, t);
        if (op === '%' || op === '#') {
          if (POSTFIX_BINDING_POWER <= minBp) {
            this.i = save;
            break;
          }
          this.i++;
          left = { kind: Node.Postfix, op: op as PostfixOp, operand: left, wsOp: gap };
          continue;
        }
        const bp = BINDING_POWER[op];
        if (bp === undefined || bp <= minBp) {
          this.i = save;
          break;
        }
        this.i++;
        // Every binary operator here is left-associative, exponentiation
        // included - `2^3^2` is 64 in Excel, not 512.
        const right = this.expression(bp);
        left = this.makeBinary(op as BinaryOp, left, right, gap);
        continue;
      }

      // A comma is the union operator only inside grouping parentheses; inside a
      // call it separates arguments and belongs to the caller.
      if (t.type === Tok.Comma && this.argSeparator === ',' && this.groupDepth > 0) {
        const bp = BINDING_POWER[',']!;
        if (bp <= minBp) {
          this.i = save;
          break;
        }
        this.i++;
        const right = this.expression(bp);
        left = { kind: Node.Binary, op: ',', left, right, wsOp: gap };
        continue;
      }

      this.i = save;
      break;
    }
    return left;
  }

  /** Could this token begin an operand? Used to detect the intersection operator. */
  private startsOperand(t: Token): boolean {
    switch (t.type) {
      case Tok.Ref:
      case Tok.Beam:
      case Tok.Name:
      case Tok.StructRef:
      case Tok.Context:
      case Tok.Function:
      case Tok.OpenParen:
        return true;
      default:
        return false;
    }
  }

  /**
   * Combine two operands with `:`.
   *
   * When both sides are plain references this collapses into a single Range
   * node, which is the overwhelmingly common case and keeps the evaluator from
   * having to handle a binary `:` at all. `A1:INDEX(...)` stays a binary node,
   * because its right-hand side is only known at evaluation time.
   */
  private makeBinary(op: BinaryOp, left: Ast, right: Ast, wsOp: string): Ast {
    if (op === ':' && left.kind === Node.Ref && right.kind === Node.Ref) {
      return { kind: Node.Range, start: left, end: right };
    }
    return wsOp ? { kind: Node.Binary, op, left, right, wsOp } : { kind: Node.Binary, op, left, right };
  }

  private prefix(): Ast {
    const t = this.peek();
    if (!t) throw new ParseError('unexpected end of formula', this.src.length);

    switch (t.type) {
      case Tok.Number: {
        this.i++;
        return { kind: Node.Number, value: Number(tokenText(this.src, t)) };
      }

      case Tok.String: {
        this.i++;
        const raw = tokenText(this.src, t);
        // Strip the surrounding quotes and undouble any escaped quote.
        const inner = raw.slice(1, raw.endsWith('"') && raw.length > 1 ? -1 : undefined);
        return { kind: Node.Text, value: inner.replaceAll('""', '"') };
      }

      case Tok.Boolean: {
        this.i++;
        return { kind: Node.Bool, value: tokenText(this.src, t).toUpperCase() === 'TRUE' };
      }

      case Tok.ErrorLit: {
        this.i++;
        const code = tokenText(this.src, t).toUpperCase();
        const err = errorFromCode(code);
        return { kind: Node.ErrorLit, code: (err?.code ?? '#VALUE!') as ErrorCode };
      }

      case Tok.Operator: {
        const op = tokenText(this.src, t);
        if (op === '-' || op === '+' || op === '@') {
          this.i++;
          // Above `^`'s binding power, so `-2^2` parses as `(-2)^2` = 4.
          const operand = this.expression(PREFIX_BINDING_POWER);
          return { kind: Node.Unary, op: op as UnaryOp, operand };
        }
        throw new ParseError(`unexpected operator ${JSON.stringify(op)}`, t.start);
      }

      case Tok.OpenParen: {
        this.i++;
        this.groupDepth++;
        const inner = this.expression(0);
        this.groupDepth--;
        const wsClose = this.skipWhitespace();
        this.expectHere(Tok.CloseParen, ')');
        return wsClose ? { kind: Node.Paren, inner, wsClose } : { kind: Node.Paren, inner };
      }

      case Tok.OpenBrace:
        return this.arrayConstant();

      case Tok.Function:
        return this.call(undefined);

      case Tok.Context:
        return this.qualified();

      case Tok.Ref:
        this.i++;
        return this.reference(tokenText(this.src, t), {});

      case Tok.Beam:
        this.i++;
        return this.beam(tokenText(this.src, t), {});

      case Tok.StructRef: {
        this.i++;
        const text = tokenText(this.src, t);
        const bracket = text.indexOf('[');
        return {
          kind: Node.StructRef,
          table: bracket < 0 ? text : text.slice(0, bracket),
          spec: bracket < 0 ? '' : text.slice(bracket),
        };
      }

      case Tok.Name: {
        this.i++;
        return { kind: Node.Name, name: tokenText(this.src, t) };
      }

      case Tok.Comma:
      case Tok.Semicolon:
        // An omitted argument, as in `IF(A1,,0)`. The caller consumes the
        // separator itself, so we do not advance here.
        return { kind: Node.Missing };

      case Tok.CloseParen:
        return { kind: Node.Missing };

      default:
        throw new ParseError(`unexpected ${JSON.stringify(tokenText(this.src, t))}`, t.start);
    }
  }

  /** A reference qualified by a sheet or workbook prefix. */
  private qualified(): Ast {
    const t = this.tokens[this.i]!;
    this.i++;
    const ctx = parseContext(tokenText(this.src, t));

    const next = this.peek();
    if (!next) throw new ParseError('sheet prefix with no reference', t.start);

    if (next.type === Tok.Ref) {
      this.i++;
      const ref = this.reference(tokenText(this.src, next), ctx);
      return ctx.sheetEnd ? { kind: Node.ThreeD, sheetStart: ctx.sheet!, sheetEnd: ctx.sheetEnd, inner: ref } : ref;
    }
    if (next.type === Tok.Beam) {
      this.i++;
      const beam = this.beam(tokenText(this.src, next), ctx);
      return ctx.sheetEnd
        ? { kind: Node.ThreeD, sheetStart: ctx.sheet!, sheetEnd: ctx.sheetEnd, inner: beam }
        : beam;
    }
    if (next.type === Tok.Function) {
      return this.call(ctx);
    }
    if (next.type === Tok.Name) {
      this.i++;
      const node = { kind: Node.Name as const, name: tokenText(this.src, next) };
      return ctx.sheet === undefined
        ? node
        : ctx.book === undefined
          ? { ...node, sheet: ctx.sheet }
          : { ...node, sheet: ctx.sheet, book: ctx.book };
    }
    if (next.type === Tok.ErrorLit) {
      this.i++;
      return { kind: Node.ErrorLit, code: tokenText(this.src, next).toUpperCase() as ErrorCode };
    }
    throw new ParseError('sheet prefix with no reference', next.start);
  }

  private call(ctx: Context | undefined): Ast {
    const t = this.tokens[this.i]!;
    this.i++;
    const name = tokenText(this.src, t);
    this.expect(Tok.OpenParen, '(');

    const args: Ast[] = [];
    const wsArgs: string[] = [];
    // Inside a call, a comma separates arguments rather than forming a union,
    // so grouping depth resets and is restored on the way out.
    const savedGroupDepth = this.groupDepth;
    this.groupDepth = 0;

    const openWs = this.skipWhitespace();
    if (this.peek()?.type === Tok.CloseParen) {
      this.i++;
      this.groupDepth = savedGroupDepth;
      return openWs
        ? { kind: Node.Call, name, args, wsArgs: [openWs] }
        : { kind: Node.Call, name, args };
    }
    // The whitespace after `(` belongs to the first argument.
    let pendingLead = openWs;

    for (;;) {
      let arg: Ast;
      try {
        arg = this.expression(0);
      } catch (err) {
        // Running out of input inside a call is overwhelmingly a half-typed
        // formula, and naming the function is far more useful in the formula
        // bar than reporting the raw position where the tokens ended.
        if (err instanceof ParseError && this.peek() === undefined) {
          throw new ParseError(`unterminated call to ${name}`, t.start);
        }
        throw err;
      }
      args.push(pendingLead ? { ...arg, ws: pendingLead + (arg.ws ?? '') } : arg);
      pendingLead = '';
      wsArgs.push(this.skipWhitespace());
      const next = this.peek();
      if (!next) throw new ParseError(`unterminated call to ${name}`, t.start);
      if (next.type === Tok.CloseParen) {
        this.i++;
        break;
      }
      const isSeparator =
        (this.argSeparator === ',' && next.type === Tok.Comma) ||
        (this.argSeparator === ';' && next.type === Tok.Semicolon) ||
        // Accept the other separator too: files written in one locale are
        // routinely opened in another, and rejecting them helps nobody.
        next.type === Tok.Comma ||
        next.type === Tok.Semicolon;
      if (!isSeparator) {
        throw new ParseError(`expected an argument separator in ${name}`, next.start);
      }
      this.i++;
    }

    this.groupDepth = savedGroupDepth;
    void ctx;
    return wsArgs.some((w) => w) ? { kind: Node.Call, name, args, wsArgs } : { kind: Node.Call, name, args };
  }

  /** `{1,2;3,4}` - a literal array. Commas separate columns, semicolons rows. */
  private arrayConstant(): Ast {
    const open = this.tokens[this.i]!;
    this.i++;
    const rows: Ast[][] = [];
    let row: Ast[] = [];

    this.skipWhitespace();
    if (this.peek()?.type === Tok.CloseBrace) {
      this.i++;
      return { kind: Node.Array, rows: [] };
    }

    for (;;) {
      // Array elements are literals only, so parse above the separator powers.
      row.push(this.expression(BINDING_POWER[',']!));
      this.skipWhitespace();
      const next = this.peek();
      if (!next) throw new ParseError('unterminated array constant', open.start);
      if (next.type === Tok.Comma) {
        this.i++;
        continue;
      }
      if (next.type === Tok.Semicolon) {
        this.i++;
        rows.push(row);
        row = [];
        continue;
      }
      if (next.type === Tok.CloseBrace) {
        this.i++;
        rows.push(row);
        break;
      }
      throw new ParseError('unexpected token in array constant', next.start);
    }

    const width = rows[0]?.length ?? 0;
    if (rows.some((r) => r.length !== width)) {
      throw new ParseError('array constant rows have different lengths', open.start);
    }
    return { kind: Node.Array, rows };
  }

  /** Build a Ref or Range node from A1 text, relative to the formula's origin. */
  private reference(text: string, ctx: Context): Ast {
    const colon = text.indexOf(':');
    if (colon < 0) {
      return this.cellRef(text, ctx);
    }
    const start = this.cellRef(text.slice(0, colon), ctx);
    const end = this.cellRef(text.slice(colon + 1), ctx);
    return { kind: Node.Range, start, end };
  }

  private cellRef(text: string, ctx: Context): RefNode {
    let i = 0;
    const colAbs = text.charCodeAt(i) === 36;
    if (colAbs) i++;
    const colStart = i;
    while (i < text.length) {
      const c = text.charCodeAt(i) & ~0x20;
      if (c < 65 || c > 90) break;
      i++;
    }
    const col = nameToCol(text.slice(colStart, i));
    const rowAbs = text.charCodeAt(i) === 36;
    if (rowAbs) i++;
    const row = Number(text.slice(i)) - 1;

    const node: RefNode = {
      kind: Node.Ref,
      // A relative reference stores its offset from the formula; an absolute one
      // stores the address itself. Both are then position-independent.
      row: rowAbs ? row : row - this.origin.row,
      col: colAbs ? col : col - this.origin.col,
      rowAbs,
      colAbs,
    };
    if (ctx.sheet !== undefined) node.sheet = ctx.sheet;
    if (ctx.book !== undefined) node.book = ctx.book;
    return node;
  }

  private beam(text: string, ctx: Context): BeamNode {
    const colon = text.indexOf(':');
    const left = text.slice(0, colon);
    const right = text.slice(colon + 1);
    const fromAbs = left.charCodeAt(0) === 36;
    const toAbs = right.charCodeAt(0) === 36;
    const leftBare = fromAbs ? left.slice(1) : left;
    const rightBare = toAbs ? right.slice(1) : right;

    if (isColumnShape(leftBare)) {
      const from = nameToCol(leftBare);
      const to = nameToCol(rightBare);
      const node: BeamNode = {
        kind: Node.Beam,
        axis: 'col',
        from: fromAbs ? from : from - this.origin.col,
        to: toAbs ? to : to - this.origin.col,
        fromAbs,
        toAbs,
      };
      if (ctx.sheet !== undefined) node.sheet = ctx.sheet;
      if (ctx.book !== undefined) node.book = ctx.book;
      return node;
    }

    const from = Number(leftBare) - 1;
    const to = Number(rightBare) - 1;
    const node: BeamNode = {
      kind: Node.Beam,
      axis: 'row',
      from: fromAbs ? from : from - this.origin.row,
      to: toAbs ? to : to - this.origin.row,
      fromAbs,
      toAbs,
    };
    if (ctx.sheet !== undefined) node.sheet = ctx.sheet;
    if (ctx.book !== undefined) node.book = ctx.book;
    return node;
  }

  private expect(type: Tok, what: string): void {
    this.skipWhitespace();
    this.expectHere(type, what);
  }

  /** Like expect, but without consuming whitespace the caller already captured. */
  private expectHere(type: Tok, what: string): void {
    const t = this.peek();
    if (!t || t.type !== type) {
      throw new ParseError(`expected ${JSON.stringify(what)}`, t?.start ?? this.src.length);
    }
    this.i++;
  }
}

/**
 * Split a context token such as `Sheet1!`, `'My Sheet'!`, `Sheet1:Sheet3!`,
 * or `'[Book 1.xlsx]Sheet 1'!` into its parts.
 */
export function parseContext(token: string): Context {
  let text = token.endsWith('!') ? token.slice(0, -1) : token;
  text = unquoteSheetName(text);

  let book: string | undefined;
  if (text.startsWith('[')) {
    const close = text.indexOf(']');
    if (close > 0) {
      book = text.slice(1, close);
      text = text.slice(close + 1);
    }
  }

  // A 3-D reference spans a range of sheets. Sheet names may themselves contain
  // a colon only when quoted, and unquoting has already happened, so this split
  // is only safe for the unquoted form - which is exactly when Excel allows the
  // 3-D syntax.
  const colon = text.indexOf(':');
  if (colon > 0 && !token.startsWith("'")) {
    const ctx: Context = { sheet: text.slice(0, colon), sheetEnd: text.slice(colon + 1) };
    if (book !== undefined) ctx.book = book;
    return ctx;
  }
  const ctx: Context = { sheet: text };
  if (book !== undefined) ctx.book = book;
  return ctx;
}

export { CellError };
