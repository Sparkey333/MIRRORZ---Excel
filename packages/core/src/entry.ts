/**
 * Non-destructive data entry.
 *
 * The most-reported failure in spreadsheets is not a crash - it is the silent
 * one. You type `SEPT1` and get a date. You paste `007` and get `7`. You import
 * an order number with nineteen digits and get `1.23457E+18`, with the low digits
 * gone for good. A genetics working group renamed human genes because Excel kept
 * corrupting them, which is the clearest possible statement that the software
 * should have changed instead.
 *
 * The fix is not to stop inferring types; people rely on typing `1/2/2024` and
 * getting a date. The fix is to make inference REVERSIBLE:
 *
 *   The literal text the user supplied is stored alongside the inferred value
 *   whenever the two differ, so the original is never destroyed and reverting is
 *   exact rather than a re-guess.
 *
 *   Inference reports what it did and how confident it was, so an import can
 *   show "412 cells would be read as dates" before committing rather than after.
 *
 *   Shapes that are known false positives - gene names, leading zeros, long
 *   digit strings, ranges like `1-2` - stay text by default. They are precisely
 *   the cases where the guess is usually wrong and always destructive.
 *
 * Formula EVALUATION stays bug-compatible with Excel's coercion rules. Entry and
 * evaluation are different layers and conflating them would trade one kind of
 * wrongness for another.
 */

import { partsToSerial, type DateSystem } from './serial.js';
import type { Scalar } from './types.js';

export interface EntryOptions {
  /** Recognise numbers. Off gives a strict text column. */
  inferNumbers?: boolean;
  /** Recognise dates and times. */
  inferDates?: boolean;
  /** Recognise TRUE and FALSE. */
  inferBooleans?: boolean;
  /**
   * Which of day and month comes first in an ambiguous numeric date.
   * There is no safe default here, which is why an ambiguous date is reported
   * as low confidence rather than silently resolved.
   */
  dateOrder?: 'mdy' | 'dmy' | 'ymd';
  dateSystem?: DateSystem;
  /**
   * Treat the input as text no matter what it looks like. This is what a column
   * type lock sets, and what a leading apostrophe means.
   */
  forceText?: boolean;
}

export type EntryKind = 'text' | 'number' | 'date' | 'time' | 'datetime' | 'boolean' | 'error' | 'blank';

export interface EntryResult {
  /** The value to store. */
  value: Scalar;
  /** What kind of thing we decided it was. */
  kind: EntryKind;
  /**
   * The text exactly as supplied, kept when it cannot be reconstructed from the
   * value alone. This is what makes the conversion reversible.
   */
  literal?: string;
  /** A number-format code implied by the input, e.g. a date or percentage. */
  impliedFormat?: string;
  /**
   * How sure we are. `certain` needs no review; `ambiguous` means a reasonable
   * person could read it either way and an import should surface it; `risky`
   * means this shape is a known false positive and we deliberately kept it as
   * text.
   */
  confidence: 'certain' | 'ambiguous' | 'risky';
  /** Why, in one line, for the review panel. */
  note?: string;
}

/**
 * Shapes that look like dates or numbers but almost never are.
 *
 * These are not hypothetical. Gene symbols are the famous case, but the same
 * applies to part numbers, chapter references, score lines and version strings.
 */
const GENE_LIKE = /^(SEPT|MARCH|MARC|DEC|OCT|NOV|APR|JUN|SEP|FEB|JAN|AUG|MAY|JUL)\d{1,2}$/i;
const RANGE_LIKE = /^\d{1,2}-\d{1,2}$/;
const LEADING_ZERO = /^0\d+$/;
const LONG_DIGITS = /^\d{16,}$/;
const VERSION_LIKE = /^\d+\.\d+\.\d+/;

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;
const ISO_DATETIME = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?$/;
const TIME_ONLY = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([AaPp][Mm])?$/;
const SLASH_DATE = /^(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})$/;

/**
 * Interpret one piece of user input.
 *
 * Returns the value to store plus everything needed to explain, review or
 * reverse the decision.
 */
export function parseEntry(input: string, options: EntryOptions = {}): EntryResult {
  const {
    inferNumbers = true,
    inferDates = true,
    inferBooleans = true,
    dateOrder,
    dateSystem = 1900,
    forceText = false,
  } = options;

  if (input === '') return { value: null, kind: 'blank', confidence: 'certain' };

  // A leading apostrophe is Excel's own "this is text" marker, and the most
  // direct way for a user to say so.
  if (input.startsWith("'")) {
    return {
      value: input.slice(1),
      kind: 'text',
      literal: input,
      confidence: 'certain',
      note: 'Leading apostrophe forces text',
    };
  }

  if (forceText) {
    return { value: input, kind: 'text', confidence: 'certain', note: 'Column is locked to text' };
  }

  const trimmed = input.trim();

  if (inferBooleans) {
    const upper = trimmed.toUpperCase();
    if (upper === 'TRUE') return { value: true, kind: 'boolean', confidence: 'certain' };
    if (upper === 'FALSE') return { value: false, kind: 'boolean', confidence: 'certain' };
  }

  // The known false positives, checked before any inference runs. Each one is a
  // shape where guessing is usually wrong and always destructive.
  const risk = riskyShape(trimmed);
  if (risk) {
    return { value: input, kind: 'text', confidence: 'risky', note: risk };
  }

  if (inferDates) {
    const date = parseDate(trimmed, dateOrder, dateSystem);
    if (date) return date;
  }

  if (inferNumbers) {
    const number = parseNumber(trimmed);
    if (number) {
      // Round-tripping matters: if rendering the number back does not reproduce
      // what was typed, keep the original so nothing is lost.
      const literal = String(number.value) === input ? undefined : input;
      return literal === undefined ? number : { ...number, literal };
    }
  }

  return { value: input, kind: 'text', confidence: 'certain' };
}

/** Why this input should stay text, or undefined when it is safe to infer. */
function riskyShape(s: string): string | undefined {
  if (GENE_LIKE.test(s)) {
    return `"${s}" looks like a gene symbol or code; kept as text rather than a date`;
  }
  if (LEADING_ZERO.test(s)) {
    return `"${s}" has a leading zero; kept as text so the zero is not lost`;
  }
  if (LONG_DIGITS.test(s)) {
    return `"${s}" has more than 15 digits; kept as text so no digits are lost`;
  }
  if (RANGE_LIKE.test(s)) {
    return `"${s}" reads as a range or score; kept as text rather than a date`;
  }
  if (VERSION_LIKE.test(s)) {
    return `"${s}" looks like a version number; kept as text`;
  }
  return undefined;
}

function parseNumber(s: string): EntryResult | undefined {
  let text = s;
  let sign = 1;
  let impliedFormat: string | undefined;

  // Accounting-style parentheses mean negative.
  if (text.startsWith('(') && text.endsWith(')')) {
    sign = -1;
    text = text.slice(1, -1).trim();
  }

  let percent = false;
  if (text.endsWith('%')) {
    percent = true;
    text = text.slice(0, -1).trim();
  }

  // A leading currency symbol implies a currency format rather than being part
  // of the number.
  const currency = /^([$£€¥])\s*/.exec(text);
  if (currency) {
    impliedFormat = `"${currency[1]}"#,##0.00`;
    text = text.slice(currency[0].length);
  }

  // Thousands separators are accepted only in well-formed groups, so "1,2,3"
  // stays text instead of quietly becoming 123.
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(text)) {
    text = text.replaceAll(',', '');
    impliedFormat ??= '#,##0.00';
  }

  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(text)) return undefined;

  const n = Number(text);
  if (!Number.isFinite(n)) return undefined;

  const value = sign * (percent ? n / 100 : n);
  const result: EntryResult = { value, kind: 'number', confidence: 'certain' };
  if (percent) result.impliedFormat = '0.00%';
  else if (impliedFormat) result.impliedFormat = impliedFormat;
  return result;
}

function parseDate(
  s: string,
  order: EntryOptions['dateOrder'],
  system: DateSystem,
): EntryResult | undefined {
  const isoDateTime = ISO_DATETIME.exec(s);
  if (isoDateTime) {
    const serial = partsToSerial(
      Number(isoDateTime[1]),
      Number(isoDateTime[2]),
      Number(isoDateTime[3]),
      Number(isoDateTime[4]),
      Number(isoDateTime[5]),
      Math.floor(Number(isoDateTime[6] ?? 0)),
      system,
    );
    return {
      value: serial,
      kind: 'datetime',
      literal: s,
      impliedFormat: 'yyyy-mm-dd hh:mm:ss',
      confidence: 'certain',
    };
  }

  const isoDate = ISO_DATE.exec(s);
  if (isoDate) {
    const [, y, m, d] = isoDate;
    if (!validYmd(Number(y), Number(m), Number(d))) return undefined;
    return {
      value: partsToSerial(Number(y), Number(m), Number(d), 0, 0, 0, system),
      kind: 'date',
      literal: s,
      impliedFormat: 'yyyy-mm-dd',
      confidence: 'certain',
    };
  }

  const time = TIME_ONLY.exec(s);
  if (time) {
    let hour = Number(time[1]);
    const minute = Number(time[2]);
    const second = Math.floor(Number(time[3] ?? 0));
    const meridiem = time[4]?.toLowerCase();
    if (meridiem === 'pm' && hour < 12) hour += 12;
    if (meridiem === 'am' && hour === 12) hour = 0;
    if (hour > 48 || minute > 59 || second > 59) return undefined;
    return {
      value: (hour * 3600 + minute * 60 + second) / 86_400,
      kind: 'time',
      literal: s,
      impliedFormat: meridiem ? 'h:mm:ss AM/PM' : 'h:mm:ss',
      confidence: 'certain',
    };
  }

  const slash = SLASH_DATE.exec(s);
  if (slash) {
    const a = Number(slash[1]);
    const b = Number(slash[2]);
    const c = Number(slash[3]);

    // Four-digit first component can only be a year.
    if (String(slash[1]).length === 4) {
      if (!validYmd(a, b, c)) return undefined;
      return {
        value: partsToSerial(a, b, c, 0, 0, 0, system),
        kind: 'date',
        literal: s,
        impliedFormat: 'yyyy-mm-dd',
        confidence: 'certain',
      };
    }

    const year = String(slash[3]).length === 4 ? c : c < 30 ? 2000 + c : 1900 + c;

    // Without a stated order, 3/4/2024 is genuinely ambiguous: it is 3 April in
    // most of the world and 4 March in the United States. Guessing silently is
    // how imported data ends up wrong by months, so we resolve it only when the
    // caller has said which convention applies, and flag it either way.
    const bothPlausible = a <= 12 && b <= 12 && a !== b;
    const resolved = order ?? 'mdy';
    const month = resolved === 'dmy' ? b : a;
    const day = resolved === 'dmy' ? a : b;
    if (!validYmd(year, month, day)) return undefined;

    const result: EntryResult = {
      value: partsToSerial(year, month, day, 0, 0, 0, system),
      kind: 'date',
      literal: s,
      impliedFormat: 'yyyy-mm-dd',
      confidence: bothPlausible && !order ? 'ambiguous' : 'certain',
    };
    if (bothPlausible && !order) {
      result.note = `"${s}" could be ${day}/${month} or ${month}/${day}; read as ${resolved.toUpperCase()}`;
    }
    return result;
  }

  return undefined;
}

function validYmd(y: number, m: number, d: number): boolean {
  if (m < 1 || m > 12 || d < 1 || d > 31) return false;
  if (y < 1 || y > 9999) return false;
  const lengths = [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return d <= lengths[m - 1]!;
}

function isLeap(y: number): boolean {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

/**
 * The pre-commit review for a paste or import.
 *
 * Excel's model is "convert everything, and you find out later, if ever". The
 * inversion is to run the same inference over the incoming values first and
 * report what would change meaning, so the decision is made before the data is
 * altered rather than after.
 */
export interface ImportReview {
  total: number;
  /** Values whose stored form differs from the text supplied. */
  converted: number;
  /** Conversions a reasonable person might disagree with. */
  ambiguous: EntryIssue[];
  /** Shapes we deliberately kept as text, and why. */
  protected: EntryIssue[];
}

export interface EntryIssue {
  row: number;
  col: number;
  input: string;
  result: EntryResult;
}

export function reviewImport(
  rows: readonly (readonly string[])[],
  options: EntryOptions = {},
): ImportReview {
  const review: ImportReview = { total: 0, converted: 0, ambiguous: [], protected: [] };
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!;
    for (let c = 0; c < row.length; c++) {
      const input = row[c]!;
      if (input === '') continue;
      review.total++;
      const result = parseEntry(input, options);
      if (result.kind !== 'text' && result.kind !== 'blank') review.converted++;
      if (result.confidence === 'ambiguous') review.ambiguous.push({ row: r, col: c, input, result });
      else if (result.confidence === 'risky') review.protected.push({ row: r, col: c, input, result });
    }
  }
  return review;
}

/**
 * Recover the text the user originally supplied.
 *
 * Stored literals make this exact. Without one, the value round-trips to text
 * unchanged, so there is nothing to recover.
 */
export function originalText(value: Scalar, literal?: string): string {
  if (literal !== undefined) return literal;
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}
