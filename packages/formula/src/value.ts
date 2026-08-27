/**
 * The value model, Excel's coercion rules, and its deliberate departures from
 * IEEE-754.
 *
 * Excel is not a bit-exact IEEE calculator, and treating it as one produces
 * visibly different numbers from Excel on ordinary financial models - which
 * users notice immediately and report as corruption. The rules below were
 * derived empirically from fixtures/generated/precedence.calc.xlsx rather than
 * copied from folklore, and the probe formulas that establish each one are named
 * in the comments so they can be re-checked against real Excel later.
 *
 * What the probes established:
 *
 *   Arithmetic itself is plain f64. `=1/3*3` is exactly 1, which rules out
 *   rounding to 15 digits after every operation - that would give
 *   0.999999999999999.
 *
 *   Addition and subtraction snap to zero when the result is negligible beside
 *   its operands. `=0.1+0.2-0.3` is 0, and crucially `=(0.1+0.2-0.3)*1E20` is
 *   also 0, so the snap happens at the subtraction and is stored, not applied
 *   cosmetically at the end of the formula.
 *
 *   Comparison and display work at 15 significant decimal digits. `=0.5-0.4` is
 *   0.09999999999999998 in f64 but displays as 0.1 and compares equal to 0.1.
 *
 *   Numeric literals are truncated to 15 significant digits on entry.
 *   `=1.0000000000000002` stores as 1, and `=9007199254740993` as
 *   9007199254740990.
 */

import { CellError, type Scalar, isError } from '@mirrorz/core';

/** Excel carries fifteen significant decimal digits, not seventeen. */
export const SIGNIFICANT_DIGITS = 15;

/**
 * Relative threshold below which an addition or subtraction result collapses to
 * zero. Every cancellation residue observed in the probes sat between 1.2e-16
 * and 2.8e-16 relative to the larger operand, and the next representable
 * difference a user could mean is far above 1e-15.
 */
const SNAP_RELATIVE_THRESHOLD = 1e-15;

/** Largest and smallest magnitudes Excel will hold before #NUM!. */
export const MAX_MAGNITUDE = 1.7976931348623157e308;
export const MIN_MAGNITUDE = 2.2250738585072014e-308;

/** Excel's text length ceiling for a single cell. */
export const MAX_TEXT_LENGTH = 32_767;

/**
 * Round to fifteen significant decimal digits.
 *
 * `toPrecision` does the work; the parse back through Number is what actually
 * collapses 0.09999999999999998 onto the double nearest 0.1.
 */
export function toExcelPrecision(v: number): number {
  if (!Number.isFinite(v) || v === 0) return v;
  return Number(v.toPrecision(SIGNIFICANT_DIGITS));
}

/**
 * Apply the entry-time truncation to a literal or typed number.
 * Identical to `toExcelPrecision`, named separately because the two rules are
 * independent and could diverge if Excel's behaviour is ever pinned down more
 * precisely.
 */
export function truncateLiteral(v: number): number {
  return toExcelPrecision(v);
}

/**
 * Addition and subtraction, with the cancellation snap.
 *
 * Applied at every add and subtract rather than only the last one: the probe
 * `=(0.1+0.2-0.3)*1E20` returns 0, which it could not if the residue survived
 * the inner subtraction.
 */
export function excelAdd(a: number, b: number): number {
  return snapCancellation(a + b, a, b);
}

export function excelSub(a: number, b: number): number {
  return snapCancellation(a - b, a, b);
}

function snapCancellation(result: number, a: number, b: number): number {
  if (result === 0 || !Number.isFinite(result)) return result;
  const scale = Math.max(Math.abs(a), Math.abs(b));
  if (scale > 0 && Math.abs(result) < scale * SNAP_RELATIVE_THRESHOLD) return 0;
  return result;
}

/**
 * Numeric comparison at display precision.
 *
 * `=0.5-0.4=0.1` is TRUE in Excel although the two doubles differ, because the
 * comparison happens after both sides are rounded to fifteen significant
 * digits. Comparing raw doubles here is the single most common reason a
 * spreadsheet clone disagrees with Excel on a boolean.
 */
export function excelNumbersEqual(a: number, b: number): boolean {
  if (a === b) return true;
  return toExcelPrecision(a) === toExcelPrecision(b);
}

export function excelCompareNumbers(a: number, b: number): number {
  if (excelNumbersEqual(a, b)) return 0;
  return a < b ? -1 : 1;
}

/** Clamp an out-of-range magnitude to the #NUM! error Excel reports. */
export function checkMagnitude(v: number): number | CellError {
  if (Number.isNaN(v)) return CellError.NUM;
  if (!Number.isFinite(v)) return CellError.NUM;
  if (v !== 0 && Math.abs(v) > MAX_MAGNITUDE) return CellError.NUM;
  return v;
}

/**
 * A reference value.
 *
 * References are first-class rather than being dereferenced at parse time,
 * because several functions must return one: OFFSET, INDIRECT, INDEX, CHOOSE
 * and IF can all be used as the left operand of `:`, and ROW/COLUMN/ROWS/
 * COLUMNS/AREAS/CELL need the reference itself rather than its contents.
 */
export interface RefValue {
  readonly kind: 'ref';
  readonly sheet: string;
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
}

export function makeRef(
  sheet: string,
  startRow: number,
  startCol: number,
  endRow = startRow,
  endCol = startCol,
): RefValue {
  return { kind: 'ref', sheet, startRow, startCol, endRow, endCol };
}

export function isRef(v: unknown): v is RefValue {
  return typeof v === 'object' && v !== null && (v as RefValue).kind === 'ref';
}

/** A dense rectangular block of scalars, row-major. */
export interface ArrayValue {
  readonly kind: 'array';
  readonly rows: number;
  readonly cols: number;
  readonly data: Scalar[];
}

export function makeArray(rows: number, cols: number, data: Scalar[]): ArrayValue {
  return { kind: 'array', rows, cols, data };
}

export function isArray(v: unknown): v is ArrayValue {
  return typeof v === 'object' && v !== null && (v as ArrayValue).kind === 'array';
}

export function arrayAt(a: ArrayValue, row: number, col: number): Scalar {
  return a.data[row * a.cols + col] ?? null;
}

/** Wrap a single scalar as a 1x1 array. */
export function scalarToArray(v: Scalar): ArrayValue {
  return makeArray(1, 1, [v]);
}

/** Everything an expression can evaluate to. */
export type Value = Scalar | ArrayValue | RefValue;

/**
 * Coerce a value to a number for arithmetic.
 *
 * Excel's rules: blank is 0, TRUE is 1 and FALSE is 0, and text is parsed as a
 * number if it looks like one - `="1"+1` is 2 - or yields #VALUE! if it does
 * not. Errors propagate unchanged.
 */
export function toNumber(v: Scalar): number | CellError {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return v;
  const parsed = parseNumericText(v);
  return parsed === undefined ? CellError.VALUE : parsed;
}

/**
 * Parse the text forms Excel accepts as numbers in arithmetic.
 *
 * Deliberately narrower than the CSV importer's inference: a formula coercing
 * text is an exact conversion, so a trailing unit or a stray word must fail
 * rather than be guessed at.
 */
export function parseNumericText(s: string): number | undefined {
  const t = s.trim();
  if (t === '') return 0;

  // Percentages: "50%" is 0.5.
  if (t.endsWith('%')) {
    const inner = parseNumericText(t.slice(0, -1));
    return inner === undefined ? undefined : inner / 100;
  }

  // Parenthesised negatives, as used in accounting formats.
  if (t.startsWith('(') && t.endsWith(')')) {
    const inner = parseNumericText(t.slice(1, -1));
    return inner === undefined ? undefined : -inner;
  }

  // Thousands separators are accepted only in well-formed groups, so that
  // "1,2,3" fails rather than silently becoming 123.
  const degrouped = /^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t) ? t.replaceAll(',', '') : t;

  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(degrouped)) return undefined;
  const n = Number(degrouped);
  return Number.isFinite(n) ? n : undefined;
}

/** Coerce to text, the way `&` and TEXT-family functions do. */
export function toText(v: Scalar): string | CellError {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return v;
  return formatNumberForConcat(v);
}

/**
 * How a number renders when concatenated, which is not the same as how it
 * renders in a cell: concatenation ignores the cell's number format and uses
 * the General-style shortest representation at 15 significant digits.
 */
export function formatNumberForConcat(v: number): string {
  if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
  const rounded = toExcelPrecision(v);
  if (Number.isInteger(rounded) && Math.abs(rounded) < 1e15) return String(rounded);
  const s = String(rounded);
  // JavaScript writes e+21 where Excel writes E+21.
  return s.includes('e') ? s.replace('e', 'E').replace('E+', 'E+').toUpperCase() : s;
}

/** Coerce to boolean, as IF and the logical functions do. */
export function toBoolean(v: Scalar): boolean | CellError {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isError(v)) return v;
  const upper = v.trim().toUpperCase();
  if (upper === 'TRUE') return true;
  if (upper === 'FALSE') return false;
  // Unlike arithmetic, a numeric string is NOT accepted here: `=IF("1",...)`
  // is #VALUE! in Excel.
  return CellError.VALUE;
}

/**
 * Type ranking for comparison.
 *
 * Excel orders values by type before value: number < text < FALSE < TRUE. This
 * is why `="a">1` is TRUE and `=TRUE>1` is TRUE, and it is a place LibreOffice
 * disagrees (it coerces the boolean and answers FALSE), so this rule follows
 * Excel rather than our fixture oracle. See docs/oracle-divergences.md.
 */
function typeRank(v: Scalar): number {
  if (v === null) return 0;
  if (typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return v ? 4 : 3;
  return 5;
}

/**
 * Compare two scalars the way Excel's relational operators do.
 * Returns a negative number, zero, or a positive number, or a CellError when
 * either side is an error.
 */
export function compareScalars(a: Scalar, b: Scalar): number | CellError {
  if (isError(a)) return a;
  if (isError(b)) return b;

  // A blank compares as the zero value of whatever it is compared against, so
  // `=Z99=""` and `=Z99=0` are both TRUE.
  if (a === null && b === null) return 0;
  if (a === null) return compareScalars(zeroOf(b), b);
  if (b === null) return compareScalars(a, zeroOf(a));

  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra - rb;

  if (typeof a === 'number' && typeof b === 'number') return excelCompareNumbers(a, b);
  if (typeof a === 'string' && typeof b === 'string') {
    // Text comparison is case-insensitive in Excel.
    const la = a.toUpperCase();
    const lb = b.toUpperCase();
    return la < lb ? -1 : la > lb ? 1 : 0;
  }
  // Both booleans, and typeRank already separated TRUE from FALSE.
  return 0;
}

function zeroOf(other: Scalar): Scalar {
  if (typeof other === 'string') return '';
  if (typeof other === 'boolean') return false;
  return 0;
}

/** Does this value stop evaluation and propagate? */
export function firstError(...values: Value[]): CellError | undefined {
  for (const v of values) {
    if (isError(v)) return v;
    if (isArray(v)) {
      for (const cell of v.data) if (isError(cell)) return cell;
    }
  }
  return undefined;
}

/**
 * Criteria matching for the *IF and *IFS family.
 *
 * A criterion is either a bare value ("apple", 5) meaning equality, or a
 * comparison prefix (">=5", "<>", "<10"). Text criteria support the `*` and `?`
 * wildcards, escaped with `~`.
 */
export interface Criterion {
  op: '=' | '<>' | '<' | '<=' | '>' | '>=';
  value: Scalar;
  /** Compiled wildcard matcher, when the criterion is a text pattern. */
  pattern?: RegExp;
}

export function parseCriterion(raw: Scalar): Criterion {
  if (typeof raw !== 'string') return { op: '=', value: raw };

  let op: Criterion['op'] = '=';
  let rest = raw;
  for (const candidate of ['<>', '<=', '>=', '<', '>', '='] as const) {
    if (raw.startsWith(candidate)) {
      op = candidate;
      rest = raw.slice(candidate.length);
      break;
    }
  }

  // A criterion body that parses as a number compares numerically.
  const asNumber = parseNumericText(rest);
  if (asNumber !== undefined && rest.trim() !== '') {
    return { op, value: asNumber };
  }
  const upper = rest.trim().toUpperCase();
  if (upper === 'TRUE') return { op, value: true };
  if (upper === 'FALSE') return { op, value: false };

  const criterion: Criterion = { op, value: rest };
  if ((op === '=' || op === '<>') && /[*?]/.test(rest)) {
    criterion.pattern = wildcardToRegExp(rest);
  }
  return criterion;
}

export function matchesCriterion(value: Scalar, criterion: Criterion): boolean {
  if (criterion.pattern) {
    const text = typeof value === 'string' ? value : toText(value);
    if (isError(text)) return false;
    const hit = criterion.pattern.test(text);
    return criterion.op === '<>' ? !hit : hit;
  }

  // An empty "=" criterion matches blank cells, which is how COUNTIF(range,"=")
  // counts blanks.
  if (criterion.value === '' && (criterion.op === '=' || criterion.op === '<>')) {
    const blank = value === null || value === '';
    return criterion.op === '=' ? blank : !blank;
  }

  const cmp = compareScalars(value, criterion.value);
  if (isError(cmp)) return false;
  switch (criterion.op) {
    case '=':
      return cmp === 0;
    case '<>':
      return cmp !== 0;
    case '<':
      return cmp < 0;
    case '<=':
      return cmp <= 0;
    case '>':
      return cmp > 0;
    case '>=':
      return cmp >= 0;
  }
}

/**
 * Excel wildcards: `*` is any run, `?` is one character, and `~` escapes either
 * (and itself). Everything else is literal, so a criterion containing a `.` or
 * `(` must not be read as a regular expression.
 */
export function wildcardToRegExp(pattern: string): RegExp {
  let out = '^';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]!;
    if (ch === '~') {
      const next = pattern[i + 1];
      if (next === '*' || next === '?' || next === '~') {
        out += escapeRegExp(next);
        i++;
        continue;
      }
      out += escapeRegExp('~');
      continue;
    }
    if (ch === '*') {
      out += '[\\s\\S]*';
      continue;
    }
    if (ch === '?') {
      out += '[\\s\\S]';
      continue;
    }
    out += escapeRegExp(ch);
  }
  return new RegExp(`${out}$`, 'i');
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
