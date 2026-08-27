/**
 * Delimited text (CSV/TSV) import and export.
 *
 * The parsing half of this file is small and mechanical; the interesting part is
 * type inference, and the design rule there is that a guess must be *safe* and
 * *reversible*. Excel's import guesses aggressively and destroys data doing it:
 * it turns the gene names SEPT1 and MARCH1 into dates, strips the leading zeros
 * off "007" and off half the postcodes in the world, and reads "1-2" as the
 * second of January. Those conversions are lossy and silent, which is the worst
 * possible combination - the user only finds out much later, from a colleague.
 *
 * So inference here follows three rules. Only shapes with exactly one plausible
 * reading are converted (ISO dates yes, "1/2/2024" no, unless the caller states
 * the order). Anything carrying a signal that it is an identifier rather than a
 * quantity stays text: a leading zero, a leading apostrophe, more than fifteen
 * significant digits, a known false-positive shape. And a conversion that
 * changes how a value reads carries the number format that renders it back -
 * dates, times and percentages all do - so exporting the sheet reproduces the
 * text that came in. Two shapes are converted without one: a grouped "1,234"
 * and an accounting "(1,234)" become the plain number they denote, and export
 * writes 1234 and -1234. The value is right and nothing is lost that a format
 * cannot restore; the punctuation is not preserved.
 *
 * Malformed input is reported, not thrown. A ragged CSV is still worth opening -
 * refusing the whole file because row 4,132 has an extra comma would be useless
 * behaviour for a spreadsheet application, so structural complaints are
 * collected as warnings and the caller decides what to show.
 *
 * The row parser is a character-level state machine that keeps its state between
 * calls to `push`, so it already works on a chunked source; `parseDelimited` is
 * simply the single-string convenience wrapper around it.
 */

import {
  MAX_COLS,
  MAX_ROWS,
  Workbook,
  type CellStyle,
  type DateSystem,
  type Scalar,
  type Sheet,
  type StyleId,
  type StyleTable,
  daysInMonth,
  isError,
  utcMsToSerial,
} from '../../core/src/index.js';

// ---------------------------------------------------------------------------
// Shared types
// ---------------------------------------------------------------------------

/** Text encodings we can decode without a locale or ICU dependency. */
export type CsvEncoding = 'utf-8' | 'utf-16le' | 'utf-16be' | 'windows-1252';

export type CsvWarningCode =
  /** A row whose field count differs from the rest of the file. */
  | 'ragged-row'
  /** Characters between a closing quote and the next delimiter. */
  | 'text-after-quote'
  /** A double quote inside a field that did not start with one. */
  | 'quote-in-unquoted-field'
  /** End of input reached inside a quoted field. */
  | 'unterminated-quote'
  /** Two or more candidate delimiters scored equally well. */
  | 'delimiter-ambiguous'
  /** The data is larger than a sheet can hold, and was truncated. */
  | 'sheet-limit'
  /** Further warnings of some code were suppressed. */
  | 'warning-limit';

export interface CsvWarning {
  code: CsvWarningCode;
  message: string;
  /** 0-based row index in the source, where the warning is row-specific. */
  row?: number;
  /** 0-based field index, where the warning is field-specific. */
  col?: number;
}

/** Order of the day, month and year components in ambiguous numeric dates. */
export type DateOrder = 'ymd' | 'dmy' | 'mdy';

export interface CsvInferenceOptions {
  /** Import every field as text, exactly as written. Overrides the rest. */
  raw?: boolean;
  /** Default true. */
  inferNumbers?: boolean;
  /** ISO 8601 dates and date-times. Default true. */
  inferDates?: boolean;
  /**
   * Bare times of day such as "13:37". Default false: "1:30" is as often a
   * score, a ratio or a duration as it is a time, and there is no signal in the
   * text to tell them apart.
   */
  inferTimes?: boolean;
  /** TRUE/FALSE, case-insensitive. Default true. */
  inferBooleans?: boolean;
  /**
   * Enables ambiguous forms such as "1/2/2024" and states how to read them.
   * Unset means those stay text, which is the default.
   */
  dateOrder?: DateOrder;
  /** Date system the produced serial numbers belong to. Default 1900. */
  dateSystem?: DateSystem;
  /**
   * Drop a leading apostrophe and keep the rest as text, mirroring the
   * spreadsheet convention for "this is not a number". Default true, and the
   * inverse of the writer's `sanitise` option.
   */
  stripLeadingApostrophe?: boolean;
}

export interface CsvReadOptions extends CsvInferenceOptions {
  /** Single character. Auto-detected when unset. */
  delimiter?: string;
  /** Forced encoding. Unset means detect from a byte-order mark, else UTF-8. */
  encoding?: CsvEncoding;
  /** Treat the first row as column headings, imported as text. Default false. */
  headerRow?: boolean;
  /** Rows sampled when auto-detecting the delimiter. Default 20. */
  sampleRows?: number;
  /** Cap on collected warnings. Default 100. */
  maxWarnings?: number;
  /** Name given to the sheet `readCsv` creates. Default "Sheet1". */
  sheetName?: string;
}

/** The result of the text layer: fields as written, with nothing inferred. */
export interface DelimitedText {
  rows: string[][];
  delimiter: string;
  encoding: CsvEncoding;
  warnings: CsvWarning[];
}

export interface InferredValue {
  value: Scalar;
  /** Number format the value should carry, for dates, times and percentages. */
  numFmt?: string;
}

export interface CsvImport {
  workbook: Workbook;
  sheet: Sheet;
  delimiter: string;
  encoding: CsvEncoding;
  /** Present only when `headerRow` was set. */
  header?: string[];
  warnings: CsvWarning[];
  rowCount: number;
  colCount: number;
}

const QUOTE = '"';

/** Excel carries 15 significant decimal digits; beyond that a value is lossy. */
export const MAX_SIGNIFICANT_DIGITS = 15;

export const DEFAULT_DELIMITER_CANDIDATES: readonly string[] = [',', ';', '\t', '|'];

// ---------------------------------------------------------------------------
// Encoding
// ---------------------------------------------------------------------------

/**
 * windows-1252, as the WHATWG Encoding Standard defines it.
 *
 * The table is written out rather than handed to `TextDecoder`, because that
 * label - like `utf-16be` - resolves only in builds carrying full ICU data and
 * throws a RangeError everywhere else. Only the 0x80-0x9F range differs from
 * Latin-1; the five bytes Microsoft leaves undefined (0x81, 0x8D, 0x8F, 0x90,
 * 0x9D) map to the matching C1 control, which is what the Encoding Standard
 * specifies and what every browser does.
 */
const CP1252_80_9F = [
  0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039,
  0x0152, 0x008d, 0x017d, 0x008f, 0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
  0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
] as const;

function decodeWindows1252(bytes: Uint8Array): string {
  const units = new Uint16Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i]!;
    units[i] = b >= 0x80 && b <= 0x9f ? CP1252_80_9F[b - 0x80]! : b;
  }
  // Chunked, because spreading a multi-megabyte array blows the argument limit.
  const CHUNK = 0x8000;
  if (units.length <= CHUNK) return String.fromCharCode(...units);
  let out = '';
  for (let i = 0; i < units.length; i += CHUNK) {
    out += String.fromCharCode(...units.subarray(i, i + CHUNK));
  }
  return out;
}

/**
 * Decode a byte buffer, honouring a byte-order mark unless told otherwise.
 *
 * UTF-16BE is byte-swapped and handed to the little-endian decoder rather than
 * named directly, because the big-endian label is only available in Node builds
 * with full ICU and we will not have the behaviour depend on how Node was
 * compiled. windows-1252 has the same problem and is decoded from a table here.
 */
export function decodeCsvBytes(
  bytes: Uint8Array,
  forced?: CsvEncoding,
): { text: string; encoding: CsvEncoding } {
  const bom = sniffBom(bytes);
  const encoding = forced ?? bom?.encoding ?? 'utf-8';
  // A mark is stripped whenever it matches the encoding actually in use, so a
  // forced encoding does not leave U+FEFF sitting in the first field.
  const body = bom && bom.encoding === encoding ? bytes.subarray(bom.length) : bytes;

  if (encoding === 'utf-16be') {
    const swapped = new Uint8Array(body.length & ~1);
    for (let i = 0; i + 1 < body.length; i += 2) {
      swapped[i] = body[i + 1]!;
      swapped[i + 1] = body[i]!;
    }
    return { text: new TextDecoder('utf-16le').decode(swapped), encoding };
  }
  if (encoding === 'windows-1252') {
    return { text: decodeWindows1252(body), encoding };
  }
  return { text: new TextDecoder(encoding).decode(body), encoding };
}

function sniffBom(bytes: Uint8Array): { encoding: CsvEncoding; length: number } | undefined {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return { encoding: 'utf-8', length: 3 };
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return { encoding: 'utf-16le', length: 2 };
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return { encoding: 'utf-16be', length: 2 };
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Row parser
// ---------------------------------------------------------------------------

export interface CsvParserSink {
  /** Called per row. Return false to stop parsing. */
  row(fields: string[], index: number): void | boolean;
  warn?(warning: CsvWarning): void;
}

/**
 * RFC 4180 field splitter as a resumable state machine.
 *
 * Everything that has to survive a chunk boundary lives in a field of this
 * class, so `push` may be called with the input cut at any character - inside a
 * quoted field, between the two halves of a CRLF, or between a quote and the
 * character that decides whether it doubled or closed.
 *
 * It is deliberately forgiving in the two places real files misbehave. A quote
 * inside an unquoted field is literal data, and text after a closing quote is
 * appended rather than discarded; both raise a warning instead of an error,
 * because the alternative is refusing to open a file that every other tool
 * opens. At most one such warning is raised per field.
 */
export class CsvRowParser {
  private field = '';
  private row: string[] = [];
  private index = 0;
  private inQuotes = false;
  private fieldWasQuoted = false;
  /** A quote was seen inside a quoted field; the next character decides. */
  private pendingQuote = false;
  /** A row ended on CR, so a following LF belongs to that same terminator. */
  private pendingCR = false;
  /** A CR was written into a quoted field, so a following LF is its pair. */
  private quotedCR = false;
  /** This field has already been complained about; one warning per field. */
  private malformedWarned = false;
  private halted = false;

  constructor(
    private readonly delimiter: string,
    private readonly sink: CsvParserSink,
  ) {
    if (delimiter.length !== 1) {
      throw new RangeError(`delimiter must be a single character, got ${JSON.stringify(delimiter)}`);
    }
    if (delimiter === QUOTE || delimiter === '\r' || delimiter === '\n') {
      throw new RangeError(`delimiter may not be ${JSON.stringify(delimiter)}`);
    }
  }

  get stopped(): boolean {
    return this.halted;
  }

  push(chunk: string): void {
    for (let i = 0; i < chunk.length && !this.halted; i++) {
      this.consume(chunk[i]!);
    }
  }

  /** Flush whatever is buffered. A trailing line terminator adds no row. */
  end(): void {
    if (this.halted) return;
    if (this.inQuotes && !this.pendingQuote) {
      this.warn({
        code: 'unterminated-quote',
        message: `row ${this.index + 1}: input ended inside a quoted field`,
        row: this.index,
        col: this.row.length,
      });
    }
    this.inQuotes = false;
    this.pendingQuote = false;
    if (this.field !== '' || this.fieldWasQuoted || this.row.length > 0) {
      this.endField();
      this.endRow();
    }
  }

  private consume(ch: string): void {
    if (this.pendingCR) {
      this.pendingCR = false;
      if (ch === '\n') return;
    }

    if (this.inQuotes) {
      if (this.pendingQuote) {
        this.pendingQuote = false;
        if (ch === QUOTE) {
          this.field += QUOTE;
          return;
        }
        this.inQuotes = false;
        // Falls through: the character after a closing quote is handled by the
        // unquoted rules below, so `"a",b` and `"a"x` both behave sensibly.
      } else {
        if (ch === QUOTE) {
          this.pendingQuote = true;
          return;
        }
        // Line endings inside a field are normalised, so a file authored on
        // Windows yields the same string as the same file authored on Unix.
        if (ch === '\r') {
          this.quotedCR = true;
          this.field += '\n';
          return;
        }
        if (ch === '\n') {
          if (this.quotedCR) {
            this.quotedCR = false;
            return;
          }
          this.field += '\n';
          return;
        }
        this.quotedCR = false;
        this.field += ch;
        return;
      }
    }

    if (ch === this.delimiter) {
      this.endField();
      return;
    }
    if (ch === '\r') {
      this.endField();
      this.endRow();
      this.pendingCR = true;
      return;
    }
    if (ch === '\n') {
      this.endField();
      this.endRow();
      return;
    }
    if (ch === QUOTE && this.field === '' && !this.fieldWasQuoted) {
      this.inQuotes = true;
      this.fieldWasQuoted = true;
      return;
    }
    if (!this.malformedWarned) {
      if (this.fieldWasQuoted) {
        this.malformedWarned = true;
        this.warn({
          code: 'text-after-quote',
          message: `row ${this.index + 1}, field ${this.row.length + 1}: text after a closing quote`,
          row: this.index,
          col: this.row.length,
        });
      } else if (ch === QUOTE) {
        // Reached only when the field already has content, since a quote in
        // first position opens a quoted field. RFC 4180 has no reading for it,
        // so it is kept as data - but the file is malformed and says so.
        this.malformedWarned = true;
        this.warn({
          code: 'quote-in-unquoted-field',
          message: `row ${this.index + 1}, field ${this.row.length + 1}: quote inside an unquoted field`,
          row: this.index,
          col: this.row.length,
        });
      }
    }
    this.field += ch;
  }

  private endField(): void {
    this.row.push(this.field);
    this.field = '';
    this.fieldWasQuoted = false;
    this.malformedWarned = false;
    this.quotedCR = false;
  }

  private endRow(): void {
    const row = this.row;
    this.row = [];
    const keepGoing = this.sink.row(row, this.index);
    this.index++;
    if (keepGoing === false) this.halted = true;
  }

  private warn(w: CsvWarning): void {
    this.sink.warn?.(w);
  }
}

// ---------------------------------------------------------------------------
// Delimiter detection
// ---------------------------------------------------------------------------

export interface DelimiterScore {
  delimiter: string;
  /** Most common field count per row for this delimiter. */
  fields: number;
  /** Fraction of sampled rows agreeing with that count, 0..1. */
  consistency: number;
  rows: number;
}

/**
 * Characters of the input examined when guessing the delimiter.
 *
 * The row cap alone is not a bound: a file with no line terminator at all - one
 * 500 MB line, which is exactly what a hostile or truncated upload looks like -
 * never reaches the row limit, and every candidate would scan the whole string.
 * A megabyte holds far more than the twenty rows the heuristic wants.
 */
const DETECT_SAMPLE_CHARS = 1 << 20;

/**
 * Guess the delimiter by splitting the sample with each candidate and seeing
 * which one produces a consistent table.
 *
 * Counting occurrences would be simpler and wrong: a file of prose sentences
 * contains far more commas than a semicolon-delimited file contains semicolons,
 * yet only the latter forms rectangular rows. Consistency of the field count is
 * the signal that actually distinguishes a delimiter from ordinary punctuation,
 * and it is computed through the real parser so quoted commas do not count.
 *
 * Only the first `sampleRows` rows of the first megabyte are examined.
 */
export function detectDelimiter(
  text: string,
  options: { candidates?: readonly string[]; sampleRows?: number } = {},
): { delimiter: string; confident: boolean; scores: DelimiterScore[] } {
  const candidates = options.candidates ?? DEFAULT_DELIMITER_CANDIDATES;
  const sampleRows = options.sampleRows ?? 20;
  const sample = text.length > DETECT_SAMPLE_CHARS ? text.slice(0, DETECT_SAMPLE_CHARS) : text;

  const scores: DelimiterScore[] = candidates.map((delimiter) => {
    const counts: number[] = [];
    const parser = new CsvRowParser(delimiter, {
      row(fields) {
        // Blank lines say nothing about the delimiter.
        if (fields.length === 1 && fields[0] === '') return counts.length < sampleRows;
        counts.push(fields.length);
        return counts.length < sampleRows;
      },
    });
    parser.push(sample);
    parser.end();

    if (counts.length === 0) return { delimiter, fields: 0, consistency: 0, rows: 0 };
    const tally = new Map<number, number>();
    for (const n of counts) tally.set(n, (tally.get(n) ?? 0) + 1);
    let best = 0;
    let bestHits = 0;
    for (const [n, hits] of tally) {
      if (hits > bestHits || (hits === bestHits && n > best)) {
        best = n;
        bestHits = hits;
      }
    }
    return { delimiter, fields: best, consistency: bestHits / counts.length, rows: counts.length };
  });

  const ranked = [...scores].sort((a, b) => {
    const aSplits = a.fields > 1 ? 1 : 0;
    const bSplits = b.fields > 1 ? 1 : 0;
    if (aSplits !== bSplits) return bSplits - aSplits;
    if (a.consistency !== b.consistency) return b.consistency - a.consistency;
    if (a.fields !== b.fields) return b.fields - a.fields;
    return candidates.indexOf(a.delimiter) - candidates.indexOf(b.delimiter);
  });

  const winner = ranked[0];
  // Nothing split the sample at all: a single-column file. Comma is the
  // conventional answer and produces the same one-column result regardless.
  if (!winner || winner.fields <= 1) {
    return { delimiter: candidates[0] ?? ',', confident: false, scores };
  }
  const runnerUp = ranked[1];
  const tied =
    runnerUp !== undefined &&
    runnerUp.fields > 1 &&
    runnerUp.consistency === winner.consistency &&
    runnerUp.fields === winner.fields;
  return { delimiter: winner.delimiter, confident: !tied && winner.consistency >= 0.9, scores };
}

// ---------------------------------------------------------------------------
// Text layer
// ---------------------------------------------------------------------------

/** Split delimited text into fields, inferring nothing. */
export function parseDelimited(
  input: string | Uint8Array,
  options: CsvReadOptions = {},
): DelimitedText {
  const maxWarnings = options.maxWarnings ?? 100;
  const warnings: CsvWarning[] = [];
  let suppressed = 0;
  const warn = (w: CsvWarning): void => {
    if (warnings.length < maxWarnings) warnings.push(w);
    else suppressed++;
  };

  let text: string;
  let encoding: CsvEncoding;
  if (typeof input === 'string') {
    encoding = options.encoding ?? 'utf-8';
    // A string source has already been decoded, but a mark may have survived it.
    text = input.charCodeAt(0) === 0xfeff ? input.slice(1) : input;
  } else {
    const decoded = decodeCsvBytes(input, options.encoding);
    text = decoded.text;
    encoding = decoded.encoding;
  }

  let delimiter = options.delimiter;
  if (delimiter === undefined) {
    const detected = detectDelimiter(text, { sampleRows: options.sampleRows ?? 20 });
    delimiter = detected.delimiter;
    const rivals = detected.scores.filter((s) => s.fields > 1);
    // A one-column file has no rival to be ambiguous with, so say nothing.
    if (!detected.confident && rivals.length > 1) {
      warn({
        code: 'delimiter-ambiguous',
        message:
          `delimiter could not be established with confidence; chose ` +
          `${JSON.stringify(delimiter)} from ` +
          rivals.map((s) => JSON.stringify(s.delimiter)).join(', '),
      });
    }
  }

  const rows: string[][] = [];
  const parser = new CsvRowParser(delimiter, {
    row(fields) {
      rows.push(fields);
    },
    warn,
  });
  parser.push(text);
  parser.end();

  reportRaggedRows(rows, warn);
  if (suppressed > 0) {
    warnings.push({
      code: 'warning-limit',
      message: `${suppressed} further warnings suppressed`,
    });
  }
  return { rows, delimiter, encoding, warnings };
}

/**
 * Flag rows whose field count departs from the file's own norm.
 *
 * The norm is the most common count rather than the first row's, so a file with
 * a one-cell title line above a proper table reports the title as the anomaly
 * and not every row beneath it.
 */
function reportRaggedRows(rows: string[][], warn: (w: CsvWarning) => void): void {
  const tally = new Map<number, number>();
  for (const row of rows) {
    if (isBlankRow(row)) continue;
    tally.set(row.length, (tally.get(row.length) ?? 0) + 1);
  }
  if (tally.size <= 1) return;

  let expected = 0;
  let hits = 0;
  for (const [n, count] of tally) {
    if (count > hits || (count === hits && n > expected)) {
      expected = n;
      hits = count;
    }
  }
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]!;
    if (isBlankRow(row) || row.length === expected) continue;
    warn({
      code: 'ragged-row',
      message: `row ${i + 1}: ${row.length} fields, expected ${expected}`,
      row: i,
    });
  }
}

function isBlankRow(row: string[]): boolean {
  return row.length === 0 || (row.length === 1 && row[0] === '');
}

// ---------------------------------------------------------------------------
// Type inference
// ---------------------------------------------------------------------------

/**
 * Shapes that a date-guessing importer famously mangles.
 *
 * The human gene nomenclature committee renamed dozens of genes in 2020 because
 * spreadsheets kept eating them, which is a remarkable concession by biology to
 * a text importer. Our grammars would not convert these anyway, since we do not
 * read month names at all - the stop-list is here so that no later widening of
 * the date grammar can reintroduce the failure without deleting this function
 * and the reasoning attached to it.
 */
const GENE_LIKE =
  /^(?:JAN|FEB|MAR|MARC|MARCH|APR|APRIL|MAY|JUN|JUNE|JUL|JULY|AUG|SEP|SEPT|OCT|NOV|DEC)-?\d{1,2}$/i;

export function looksLikeGeneName(text: string): boolean {
  return GENE_LIKE.test(text.trim());
}

const NUMBER_RE = /^([+-]?)(?:(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d*))?|\.(\d+))(?:[eE]([+-]?\d+))?$/;
const ISO_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})$/;
const LOOSE_DATE_RE = /^(\d{1,4})([/\-.])(\d{1,2})\2(\d{1,4})$/;
const TIME_RE = /^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?$/;

/**
 * Decide what a single field means.
 *
 * The order is fixed: the escapes that force text come first, then the shapes
 * with no overlap between them. Nothing here consults a locale, and nothing
 * falls back to `Date.parse`, whose behaviour on partial input is
 * implementation-defined and has produced silently different results across
 * engines for years.
 */
export function inferValue(text: string, options: CsvInferenceOptions = {}): InferredValue {
  if (text === '') return { value: null };
  if (options.raw) return { value: text };

  if ((options.stripLeadingApostrophe ?? true) && text.startsWith("'")) {
    return { value: text.slice(1) };
  }
  if (looksLikeGeneName(text)) return { value: text };

  const trimmed = text.trim();
  if (trimmed === '') return { value: text };

  if (options.inferBooleans ?? true) {
    const upper = trimmed.toUpperCase();
    if (upper === 'TRUE') return { value: true };
    if (upper === 'FALSE') return { value: false };
  }

  if (options.inferNumbers ?? true) {
    const num = parseNumber(trimmed);
    if (num) return num;
  }

  const date = parseDateLike(trimmed, options);
  if (date) return date;

  return { value: text };
}

/**
 * Parse a numeric literal, or return undefined to leave the field as text.
 *
 * Rejection is as much the point as acceptance. A leading zero means the digits
 * are an identifier - a part number, a postcode, an hour written "08" - and
 * turning "007" into 7 is not a formatting change, it is data loss, because
 * nothing in the sheet records how many zeros there were. Sixteen or more
 * significant digits cannot survive a double either, so those stay text too;
 * an order reference would come back off by one and look plausible doing it.
 */
function parseNumber(s: string): InferredValue | undefined {
  let negate = false;
  let body = s;

  // Accounting notation: (1,234) is negative twelve hundred and thirty four.
  const parens = /^\((.*)\)$/s.exec(body);
  if (parens) {
    negate = true;
    body = parens[1]!.trim();
    if (body === '' || /^[+-]/.test(body)) return undefined;
  }

  let percent = false;
  if (body.endsWith('%')) {
    percent = true;
    body = body.slice(0, -1).trim();
  }

  const m = NUMBER_RE.exec(body);
  if (!m) return undefined;
  const sign = m[1] ?? '';
  const intPart = (m[2] ?? '').replace(/,/g, '');
  const fracPart = m[3] ?? m[4] ?? '';
  const exp = m[5];

  if (intPart.length > 1 && intPart.startsWith('0')) return undefined;
  const digits = intPart + fracPart;
  if (significantDigits(digits) > MAX_SIGNIFICANT_DIGITS) return undefined;

  const mantissa = `${intPart === '' ? '0' : intPart}.${fracPart === '' ? '0' : fracPart}`;
  let value = Number(`${sign}${mantissa}e${exp ?? '0'}`);
  if (!Number.isFinite(value)) return undefined;
  // Underflow: a non-zero literal that collapses to zero has lost everything.
  if (value === 0 && /[1-9]/.test(digits)) return undefined;

  if (percent) value /= 100;
  if (negate) value = -value;
  if (Object.is(value, -0)) value = 0;

  if (percent) {
    // As many decimals as the text showed: "50.00%" is a measurement written to
    // two places and must come back out as "50.00%", not "50%". An exponent
    // moves the point, so the count is taken after applying it.
    const places = Math.max(0, fracPart.length - Number(exp ?? 0));
    return { value, numFmt: places > 0 ? `0.${'0'.repeat(places)}%` : '0%' };
  }
  return { value };
}

/** Digits that carry information: leading and trailing zeros carry none. */
function significantDigits(digits: string): number {
  return digits.replace(/^0+/, '').replace(/0+$/, '').length;
}

function parseDateLike(s: string, options: CsvInferenceOptions): InferredValue | undefined {
  const system = options.dateSystem ?? 1900;
  const parts = s.split(/[T ]/);
  const datePart = parts[0] ?? '';
  const timePart = parts.length === 2 ? parts[1] ?? '' : undefined;
  if (parts.length > 2) return undefined;

  if (options.inferDates ?? true) {
    const date = parseDatePart(datePart, options.dateOrder);
    if (date) {
      const time = timePart === undefined ? undefined : parseTimePart(timePart);
      if (timePart !== undefined && !time) return undefined;
      const serial =
        dateToSerial(date.y, date.m, date.d, time?.h ?? 0, time?.mi ?? 0, time?.s ?? 0, system) +
        (time?.frac ?? 0);
      // Excel's calendar starts at 1 January 1900 (serial 1) or, in the 1904
      // system, at 1 January 1904 (serial 0). Serial 0 in the 1900 system is
      // the fictitious "January 0", so 1899-12-31 is not a date Excel can
      // hold - it stays text rather than becoming a number that renders wrong.
      if (serial < (system === 1904 ? 0 : 1)) return undefined;
      if (!time) return { value: serial, numFmt: 'yyyy-mm-dd' };
      return { value: serial, numFmt: time.hasSeconds ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd hh:mm' };
    }
  }

  if ((options.inferTimes ?? false) && timePart === undefined) {
    const time = parseTimePart(datePart);
    if (time) {
      const serial = (time.h * 3600 + time.mi * 60 + time.s) / 86_400 + time.frac;
      return { value: serial, numFmt: time.hasSeconds ? 'hh:mm:ss' : 'hh:mm' };
    }
  }
  return undefined;
}

/**
 * Serial number for a fully specified calendar date.
 *
 * Core's `partsToSerial` implements Excel's DATE(), which maps any year below
 * 1900 into the twentieth century - DATE(99,1,1) is 1999, by design. That rule
 * is right for a formula argument and wrong for a written-out ISO year, so this
 * goes through the epoch conversion directly and keeps 0099-01-01 in the first
 * century, where it will fall outside the serial range and stay text.
 */
function dateToSerial(
  y: number,
  m: number,
  d: number,
  h: number,
  mi: number,
  s: number,
  system: DateSystem,
): number {
  let ms = Date.UTC(y, m - 1, d, h, mi, s);
  if (y >= 0 && y <= 99) {
    // Date.UTC reads a two-digit year as 19xx; undo that for a genuine one.
    const dt = new Date(ms);
    dt.setUTCFullYear(y);
    ms = dt.getTime();
  }
  return utcMsToSerial(ms, system);
}

function parseDatePart(
  s: string,
  order: DateOrder | undefined,
): { y: number; m: number; d: number } | undefined {
  const iso = ISO_DATE_RE.exec(s);
  if (iso) {
    const y = Number(iso[1]);
    const m = Number(iso[2]);
    const d = Number(iso[3]);
    return validDate(y, m, d) ? { y, m, d } : undefined;
  }
  if (!order) return undefined;

  const loose = LOOSE_DATE_RE.exec(s);
  if (!loose) return undefined;
  const a = loose[1]!;
  const b = loose[3]!;
  const c = loose[4]!;
  const [yRaw, mRaw, dRaw] =
    order === 'ymd' ? [a, b, c] : order === 'dmy' ? [c, b, a] : [c, a, b];
  const y = expandYear(yRaw);
  if (y === undefined) return undefined;
  const m = Number(mRaw);
  const d = Number(dRaw);
  return validDate(y, m, d) ? { y, m, d } : undefined;
}

/** Two-digit years follow Excel: 00-29 are 2000s, 30-99 are 1900s. */
function expandYear(raw: string): number | undefined {
  if (raw.length === 4) return Number(raw);
  if (raw.length > 2) return undefined;
  const n = Number(raw);
  return n <= 29 ? 2000 + n : 1900 + n;
}

function validDate(y: number, m: number, d: number): boolean {
  if (m < 1 || m > 12 || d < 1) return false;
  return d <= daysInMonth(y, m);
}

function parseTimePart(
  s: string,
): { h: number; mi: number; s: number; frac: number; hasSeconds: boolean } | undefined {
  const m = TIME_RE.exec(s);
  if (!m) return undefined;
  const h = Number(m[1]);
  const mi = Number(m[2]);
  const sec = m[3] === undefined ? 0 : Number(m[3]);
  // Hours past 23 are an elapsed duration, not a time of day, and only a number
  // format can tell the two apart. Left as text rather than guessed at.
  if (h > 23 || mi > 59 || sec > 59) return undefined;
  const frac = m[4] === undefined ? 0 : Number(`0.${m[4]}`) / 86_400;
  return { h, mi, s: sec, frac, hasSeconds: m[3] !== undefined };
}

// ---------------------------------------------------------------------------
// Import into a sheet
// ---------------------------------------------------------------------------

/**
 * Write parsed rows into a sheet, interning a number format per inferred date.
 *
 * Dates become a serial number plus a format, which is what a spreadsheet date
 * is; storing the text instead would leave DATEDIF and sorting broken, and
 * storing the serial without the format would show the user 45000.
 */
export function rowsToSheet(
  rows: readonly string[][],
  sheet: Sheet,
  styles: StyleTable,
  options: CsvReadOptions = {},
  warn: (w: CsvWarning) => void = () => {},
): { rowCount: number; colCount: number; header?: string[] } {
  const formats = new Map<string, StyleId>();
  const styleFor = (numFmt: string): StyleId => {
    const existing = formats.get(numFmt);
    if (existing !== undefined) return existing;
    const id = styles.intern({ numFmt });
    formats.set(numFmt, id);
    return id;
  };

  let header: string[] | undefined;
  let colCount = 0;
  let rowCount = 0;
  let truncated = false;

  for (let r = 0; r < rows.length; r++) {
    if (r >= MAX_ROWS) {
      truncated = true;
      break;
    }
    const row = rows[r]!;
    const isHeader = r === 0 && options.headerRow === true;
    if (isHeader) header = [...row];
    rowCount = r + 1;

    for (let c = 0; c < row.length; c++) {
      if (c >= MAX_COLS) {
        truncated = true;
        break;
      }
      const raw = row[c]!;
      if (c + 1 > colCount) colCount = c + 1;
      // Headings are labels by definition, so they never go through inference.
      const inferred = isHeader ? { value: raw === '' ? null : raw } : inferValue(raw, options);
      if (inferred.value === null) continue;
      if ('numFmt' in inferred && inferred.numFmt) {
        sheet.setCell(r, c, { value: inferred.value, style: styleFor(inferred.numFmt) });
      } else {
        sheet.setValue(r, c, inferred.value);
      }
    }
  }

  if (truncated) {
    warn({
      code: 'sheet-limit',
      message: `input exceeds the ${MAX_ROWS} x ${MAX_COLS} sheet limit and was truncated`,
    });
  }
  return header ? { rowCount, colCount, header } : { rowCount, colCount };
}

/** Read delimited text into a fresh single-sheet workbook. */
export function readCsv(input: string | Uint8Array, options: CsvReadOptions = {}): CsvImport {
  const parsed = parseDelimited(input, options);
  const workbook = new Workbook();
  const sheet = workbook.addSheet(options.sheetName ?? 'Sheet1');
  const warnings = parsed.warnings;
  const filled = rowsToSheet(parsed.rows, sheet, workbook.styles, options, (w) => {
    warnings.push(w);
  });
  return {
    workbook,
    sheet,
    delimiter: parsed.delimiter,
    encoding: parsed.encoding,
    warnings,
    rowCount: filled.rowCount,
    colCount: filled.colCount,
    ...(filled.header ? { header: filled.header } : {}),
  };
}

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------

/**
 * Renders one cell. Returning undefined defers to the built-in rendering, so a
 * caller can format only the cells it has an opinion about.
 *
 * The number-format engine lives in another module and is deliberately not
 * imported here: CSV export must keep working, and keep being testable, without
 * dragging in a format interpreter, and a caller that wants Excel-faithful text
 * passes one in.
 */
export type CsvValueFormatter = (
  value: Scalar,
  numFmt: string | undefined,
  context: { row: number; col: number; style: CellStyle },
) => string | undefined;

export interface CsvWriteOptions {
  delimiter?: string;
  /** Default CRLF, as RFC 4180 specifies. */
  lineEnding?: '\r\n' | '\n' | '\r';
  /** Emit U+FEFF first. Excel needs it to read a UTF-8 file as UTF-8. */
  bom?: boolean;
  /** Quote every field, not just the ones that need it. */
  quoteAll?: boolean;
  /** End the last row with a line terminator. Default true. */
  trailingNewline?: boolean;
  /** Style table backing the sheet, needed to pass number formats along. */
  styles?: StyleTable;
  format?: CsvValueFormatter;
  /** Restrict the export to a rectangle. Defaults to the used range. */
  range?: { minRow: number; minCol: number; maxRow: number; maxCol: number };
  /**
   * Prefix text fields beginning with =, +, -, @, tab or CR with an apostrophe.
   *
   * Off by default, and that is a deliberate choice rather than an oversight. A
   * CSV is a data interchange format; the injection risk belongs to whichever
   * spreadsheet later opens the file and decides to execute a formula it found
   * in a data file. Sanitising on export corrupts the data for every honest
   * consumer - a column of ranges like "-5" or a genuine "=" comes back with a
   * stray apostrophe, checksums stop matching, and re-importing needs the
   * matching strip step. Callers exporting to a destination that will open the
   * file in a spreadsheet should switch it on; callers writing a data feed
   * should not.
   */
  sanitise?: boolean;
}

const INJECTION_PREFIX = /^[=+\-@\t\r]/;

/** True when the field cannot be written bare. */
export function needsQuoting(field: string, delimiter: string): boolean {
  if (field === '') return false;
  if (
    field.includes(delimiter) ||
    field.includes(QUOTE) ||
    field.includes('\n') ||
    field.includes('\r')
  ) {
    return true;
  }
  // Leading and trailing spaces survive only inside quotes; several importers,
  // Excel among them, strip them otherwise.
  return /^\s|\s$/.test(field);
}

function encodeField(field: string, delimiter: string, quoteAll: boolean): string {
  if (quoteAll || needsQuoting(field, delimiter)) {
    return QUOTE + field.replace(/"/g, '""') + QUOTE;
  }
  return field;
}

/** Serialise ready-made text fields. */
export function writeRows(
  rows: readonly (readonly string[])[],
  options: CsvWriteOptions = {},
): string {
  const delimiter = options.delimiter ?? ',';
  const eol = options.lineEnding ?? '\r\n';
  const quoteAll = options.quoteAll === true;
  const out: string[] = [];
  for (const row of rows) {
    out.push(row.map((f) => encodeField(f, delimiter, quoteAll)).join(delimiter));
  }
  let text = out.join(eol);
  // Keyed off the row count, not the text: a single row of empty fields is one
  // row and needs its terminator, or it comes back from the reader as no rows.
  if (out.length > 0 && (options.trailingNewline ?? true)) text += eol;
  if (options.bom) text = '\uFEFF' + text;
  return text;
}

/**
 * Default rendering for a cell value.
 *
 * Numbers are rounded to Excel's fifteen significant digits and then printed in
 * JavaScript's shortest round-tripping form, with the exponent upper-cased.
 *
 * The rounding is not cosmetic. `String(0.1 + 0.2)` is "0.30000000000000004",
 * seventeen significant digits, which the reader on the other side of this file
 * deliberately refuses to treat as a number - so a computed cell would leave as
 * a number and come back as text. Excel stores fifteen digits and writes "0.3",
 * which is both what a user expects to see and a value that survives the round
 * trip. Anything needing a cell's real number format passes a `format` callback.
 */
export function formatScalar(value: Scalar): string {
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (isError(value)) return value.code;
  if (typeof value === 'number') return formatNumber(value);
  return value;
}

function formatNumber(value: number): string {
  if (!Number.isFinite(value)) return String(value);
  const rounded = value === 0 ? 0 : Number(value.toPrecision(MAX_SIGNIFICANT_DIGITS));
  return String(rounded).replace('e', 'E');
}

/** Serialise a sheet as delimited text. */
export function writeCsv(sheet: Sheet, options: CsvWriteOptions = {}): string {
  const delimiter = options.delimiter ?? ',';
  const styles = options.styles;
  // Without an explicit range the export is anchored at A1, not at the first
  // populated cell, so an empty leading row or column stays empty rather than
  // shifting every value one place up or left on the way out.
  const used = sheet.bounds();
  const bounds =
    options.range ??
    (used ? { minRow: 0, minCol: 0, maxRow: used.maxRow, maxCol: used.maxCol } : undefined);
  if (!bounds) return options.bom ? '\uFEFF' : '';

  const rows: string[][] = [];
  for (let r = Math.max(0, bounds.minRow); r <= bounds.maxRow; r++) {
    const row: string[] = [];
    for (let c = Math.max(0, bounds.minCol); c <= bounds.maxCol; c++) {
      const cell = sheet.getCell(r, c);
      const value: Scalar = cell?.value ?? null;
      const style: CellStyle = styles ? styles.get(sheet.getStyle(r, c)) : {};
      const custom = options.format?.(value, style.numFmt, { row: r, col: c, style });
      let text = custom ?? formatScalar(value);
      // Sanitising is confined to text cells: an apostrophe in front of a
      // rendered number would change the value on the way back in.
      if (options.sanitise && typeof value === 'string' && INJECTION_PREFIX.test(text)) {
        text = "'" + text;
      }
      row.push(text);
    }
    rows.push(row);
  }
  return writeRows(rows, { ...options, delimiter });
}

/** As `writeCsv`, encoded as UTF-8 bytes. */
export function writeCsvBytes(sheet: Sheet, options: CsvWriteOptions = {}): Uint8Array {
  return new TextEncoder().encode(writeCsv(sheet, options));
}
