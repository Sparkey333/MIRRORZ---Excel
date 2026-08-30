/**
 * Formula tokenizer.
 *
 * Single pass, no regex alternation, no backtracking. Tokens carry offsets into
 * the source rather than substrings, so scanning a formula allocates nothing
 * until a caller actually asks for the text.
 *
 * Two decisions here are not retrofittable later, so they are made deliberately:
 *
 * 1. Whitespace is emitted as a token, never discarded. A space between two
 *    references is Excel's intersection operator - `SUM(A1:C3 B1:B5)` is a real
 *    formula - so a lexer that swallows spaces can never support intersection.
 *    Keeping whitespace also lets us round-trip the user's own formatting
 *    inside a formula rather than reflowing it on every save.
 *
 * 2. A name immediately followed by `(` is a function call; the same characters
 *    without the paren are a defined name. That distinction is made here, at the
 *    lexer, because the parser cannot recover it without unbounded lookahead.
 */

export const enum Tok {
  Number,
  String,
  Boolean,
  ErrorLit,
  /** Identifier followed directly by `(`. */
  Function,
  /** Sheet or workbook prefix ending in `!`, e.g. `Sheet1!` or `Sheet1:Sheet3!`. */
  Context,
  /** A1-style cell or range reference. */
  Ref,
  /** Whole-column or whole-row reference: `A:A`, `1:1`. */
  Beam,
  /** Structured (table) reference: `Table1[Column]`. */
  StructRef,
  /** Defined name, or anything identifier-shaped that is not a reference. */
  Name,
  Operator,
  OpenParen,
  CloseParen,
  OpenBrace,
  CloseBrace,
  Comma,
  Semicolon,
  Whitespace,
  /** Leading `=` of a formula. */
  Equals,
  Unknown,
  EOF,
}

export interface Token {
  type: Tok;
  start: number;
  end: number;
}

/**
 * Error literals, longest-first so that scanning matches `#N/A` before `#N`.
 * `#REF!` is included here but is special downstream: it is reference-valued and
 * can appear as an operand of the range operator.
 */
const ERROR_LITERALS = [
  '#GETTING_DATA',
  '#BLOCKED!',
  '#CONNECT!',
  '#UNKNOWN!',
  '#EXTERNAL!',
  '#DIV/0!',
  '#SPILL!',
  '#VALUE!',
  '#NAME?',
  '#FIELD!',
  '#NULL!',
  '#BUSY!',
  '#CALC!',
  '#NUM!',
  '#REF!',
  '#N/A',
] as const;

/** Multi-character operators, checked before their single-character prefixes. */
const TWO_CHAR_OPERATORS = ['<=', '>=', '<>'] as const;

const SINGLE_CHAR_OPERATORS = '+-*/^&<>=:%#@!';

export class Lexer {
  private pos = 0;
  private readonly len: number;

  constructor(readonly src: string) {
    this.len = src.length;
  }

  /** Tokenize the whole formula. */
  static tokenize(src: string): Token[] {
    const lexer = new Lexer(src);
    const out: Token[] = [];
    for (;;) {
      const t = lexer.next();
      if (t.type === Tok.EOF) break;
      out.push(t);
    }
    return out;
  }

  next(): Token {
    if (this.pos >= this.len) return { type: Tok.EOF, start: this.pos, end: this.pos };
    const start = this.pos;
    const c = this.src.charCodeAt(this.pos);

    if (isSpace(c)) {
      while (this.pos < this.len && isSpace(this.src.charCodeAt(this.pos))) this.pos++;
      return { type: Tok.Whitespace, start, end: this.pos };
    }

    // A leading `=` only counts as the formula marker at position 0; anywhere
    // else it is the comparison operator.
    if (c === 61 /* = */ && start === 0) {
      this.pos++;
      return { type: Tok.Equals, start, end: this.pos };
    }

    if (c === 34 /* " */) return this.scanString(start);
    if (c === 39 /* ' */) return this.scanQuotedContext(start);
    if (c === 35 /* # */) return this.scanError(start);
    if (c === 91 /* [ */) return this.scanBracketed(start);

    if (isDigit(c) || (c === 46 /* . */ && isDigit(this.src.charCodeAt(this.pos + 1)))) {
      return this.scanNumber(start);
    }

    switch (c) {
      case 40:
        this.pos++;
        return { type: Tok.OpenParen, start, end: this.pos };
      case 41:
        this.pos++;
        return { type: Tok.CloseParen, start, end: this.pos };
      case 123:
        this.pos++;
        return { type: Tok.OpenBrace, start, end: this.pos };
      case 125:
        this.pos++;
        return { type: Tok.CloseBrace, start, end: this.pos };
      case 44:
        this.pos++;
        return { type: Tok.Comma, start, end: this.pos };
      case 59:
        this.pos++;
        return { type: Tok.Semicolon, start, end: this.pos };
      default:
        break;
    }

    if (isIdentStart(c) || c === 36 /* $ */) return this.scanWord(start);

    const two = this.src.slice(this.pos, this.pos + 2);
    if ((TWO_CHAR_OPERATORS as readonly string[]).includes(two)) {
      this.pos += 2;
      return { type: Tok.Operator, start, end: this.pos };
    }
    if (SINGLE_CHAR_OPERATORS.includes(this.src[this.pos]!)) {
      this.pos++;
      return { type: Tok.Operator, start, end: this.pos };
    }

    this.pos++;
    return { type: Tok.Unknown, start, end: this.pos };
  }

  private scanString(start: number): Token {
    this.pos++; // opening quote
    while (this.pos < this.len) {
      if (this.src.charCodeAt(this.pos) === 34) {
        // A doubled quote is an escaped quote, not the end of the string.
        if (this.src.charCodeAt(this.pos + 1) === 34) {
          this.pos += 2;
          continue;
        }
        this.pos++;
        return { type: Tok.String, start, end: this.pos };
      }
      this.pos++;
    }
    // Unterminated: report what we have rather than throwing, so an editor can
    // still highlight a half-typed formula.
    return { type: Tok.String, start, end: this.pos };
  }

  /**
   * A single-quoted sheet or workbook name, e.g. `'My Sheet'!A1` or
   * `'[Book 1.xlsx]Sheet 1'!A1`. Inside the quotes a literal apostrophe is
   * doubled.
   */
  private scanQuotedContext(start: number): Token {
    this.pos++;
    while (this.pos < this.len) {
      if (this.src.charCodeAt(this.pos) === 39) {
        if (this.src.charCodeAt(this.pos + 1) === 39) {
          this.pos += 2;
          continue;
        }
        this.pos++;
        break;
      }
      this.pos++;
    }
    // Only a following `!` makes this a sheet context; otherwise it is junk.
    if (this.src.charCodeAt(this.pos) === 33 /* ! */) {
      this.pos++;
      return { type: Tok.Context, start, end: this.pos };
    }
    return { type: Tok.Unknown, start, end: this.pos };
  }

  private scanError(start: number): Token {
    for (const lit of ERROR_LITERALS) {
      if (this.src.startsWith(lit, start)) {
        this.pos = start + lit.length;
        return { type: Tok.ErrorLit, start, end: this.pos };
      }
    }
    // A bare `#` is the postfix spill-range operator: `A1#`.
    this.pos++;
    return { type: Tok.Operator, start, end: this.pos };
  }

  /** A bracketed structured-reference tail, e.g. the `[Column]` of `Table1[Column]`. */
  private scanBracketed(start: number): Token {
    let depth = 0;
    while (this.pos < this.len) {
      const ch = this.src.charCodeAt(this.pos);
      if (ch === 91) depth++;
      else if (ch === 93) {
        depth--;
        if (depth === 0) {
          this.pos++;
          break;
        }
      }
      this.pos++;
    }
    return { type: Tok.StructRef, start, end: this.pos };
  }

  private scanNumber(start: number): Token {
    while (this.pos < this.len && isDigit(this.src.charCodeAt(this.pos))) this.pos++;
    if (this.src.charCodeAt(this.pos) === 46 /* . */) {
      this.pos++;
      while (this.pos < this.len && isDigit(this.src.charCodeAt(this.pos))) this.pos++;
    }
    const e = this.src.charCodeAt(this.pos);
    if (e === 101 || e === 69 /* e E */) {
      const save = this.pos;
      this.pos++;
      const sign = this.src.charCodeAt(this.pos);
      if (sign === 43 || sign === 45) this.pos++;
      if (isDigit(this.src.charCodeAt(this.pos))) {
        while (this.pos < this.len && isDigit(this.src.charCodeAt(this.pos))) this.pos++;
      } else {
        // `1E` with no exponent digits: the E belongs to whatever follows.
        this.pos = save;
      }
    }
    return { type: Tok.Number, start, end: this.pos };
  }

  /**
   * An identifier-shaped run, resolved into a function name, a sheet context, a
   * reference, a beam, a boolean, or a defined name.
   */
  private scanWord(start: number): Token {
    while (this.pos < this.len && isIdentPart(this.src.charCodeAt(this.pos))) this.pos++;

    // A sheet or 3-D context prefix: `Sheet1!`, `Sheet1:Sheet3!`.
    if (this.src.charCodeAt(this.pos) === 33 /* ! */) {
      this.pos++;
      return { type: Tok.Context, start, end: this.pos };
    }
    if (this.src.charCodeAt(this.pos) === 58 /* : */) {
      const save = this.pos;
      this.pos++;
      while (this.pos < this.len && isIdentPart(this.src.charCodeAt(this.pos))) this.pos++;
      if (this.src.charCodeAt(this.pos) === 33) {
        this.pos++;
        return { type: Tok.Context, start, end: this.pos };
      }
      this.pos = save;
    }

    const word = this.src.slice(start, this.pos);

    // A name directly followed by `(` is a call. No whitespace is allowed
    // between them in Excel, which is what makes this decidable here.
    if (this.src.charCodeAt(this.pos) === 40 /* ( */) {
      return { type: Tok.Function, start, end: this.pos };
    }

    // `Table1[...]` is a structured reference; absorb the bracketed part.
    if (this.src.charCodeAt(this.pos) === 91 /* [ */) {
      this.scanBracketed(this.pos);
      return { type: Tok.StructRef, start, end: this.pos };
    }

    const upper = word.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') {
      return { type: Tok.Boolean, start, end: this.pos };
    }

    // A range like `A1:B2` is scanned as one token so the parser does not have
    // to distinguish it from the range operator applied to two names.
    if (isCellShape(word) && this.src.charCodeAt(this.pos) === 58 /* : */) {
      const save = this.pos;
      this.pos++;
      const rhsStart = this.pos;
      while (this.pos < this.len && isIdentPart(this.src.charCodeAt(this.pos))) this.pos++;
      const rhs = this.src.slice(rhsStart, this.pos);
      if (isCellShape(rhs) && this.src.charCodeAt(this.pos) !== 40) {
        return { type: Tok.Ref, start, end: this.pos };
      }
      this.pos = save;
    }

    if (isCellShape(word)) return { type: Tok.Ref, start, end: this.pos };

    // Whole-column beams: `A:A`, `$A:$C`.
    if (isColumnShape(word) && this.src.charCodeAt(this.pos) === 58) {
      const save = this.pos;
      this.pos++;
      const rhsStart = this.pos;
      while (this.pos < this.len && isIdentPart(this.src.charCodeAt(this.pos))) this.pos++;
      if (isColumnShape(this.src.slice(rhsStart, this.pos))) {
        return { type: Tok.Beam, start, end: this.pos };
      }
      this.pos = save;
    }

    return { type: Tok.Name, start, end: this.pos };
  }
}

function isSpace(c: number): boolean {
  return c === 32 || c === 9 || c === 10 || c === 13;
}

function isDigit(c: number): boolean {
  return c >= 48 && c <= 57;
}

function isIdentStart(c: number): boolean {
  return (
    (c >= 65 && c <= 90) ||
    (c >= 97 && c <= 122) ||
    c === 95 /* _ */ ||
    c === 92 /* \ - legal in defined names */ ||
    c > 127
  );
}

function isIdentPart(c: number): boolean {
  return isIdentStart(c) || isDigit(c) || c === 46 /* . */ || c === 36 /* $ */ || c === 63 /* ? */;
}

/**
 * Does this text have the shape of an A1 cell address that is actually on the
 * grid?
 *
 * The range check is what stops `A1048577` and `XFE1` from lexing as references
 * - both are legal defined names, and Excel treats them as such. Getting this
 * wrong means a workbook with a name like `TAX2024` silently becomes a broken
 * reference.
 */
export function isCellShape(word: string): boolean {
  let i = 0;
  if (word.charCodeAt(i) === 36) i++;
  const colStart = i;
  while (i < word.length) {
    const c = word.charCodeAt(i) & ~0x20;
    if (c < 65 || c > 90) break;
    i++;
  }
  const colLen = i - colStart;
  if (colLen === 0 || colLen > 3) return false;
  if (word.charCodeAt(i) === 36) i++;
  const rowStart = i;
  while (i < word.length && isDigit(word.charCodeAt(i))) i++;
  if (i !== word.length || i === rowStart) return false;

  // Within the grid: columns A..XFD, rows 1..1048576.
  const colName = word.slice(colStart, colStart + colLen).toUpperCase();
  let col = 0;
  for (let k = 0; k < colName.length; k++) col = col * 26 + (colName.charCodeAt(k) - 64);
  if (col > 16_384) return false;
  const row = Number(word.slice(rowStart));
  return row >= 1 && row <= 1_048_576;
}

/** `A`, `$A`, `XFD` - the operand shape of a whole-column beam. */
export function isColumnShape(word: string): boolean {
  let i = 0;
  if (word.charCodeAt(i) === 36) i++;
  const start = i;
  while (i < word.length) {
    const c = word.charCodeAt(i) & ~0x20;
    if (c < 65 || c > 90) break;
    i++;
  }
  if (i !== word.length || i - start === 0 || i - start > 3) return false;
  const name = word.slice(start).toUpperCase();
  let col = 0;
  for (let k = 0; k < name.length; k++) col = col * 26 + (name.charCodeAt(k) - 64);
  return col <= 16_384;
}

/** Token text, sliced on demand. */
export function tokenText(src: string, t: Token): string {
  return src.slice(t.start, t.end);
}
