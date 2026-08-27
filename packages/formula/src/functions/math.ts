/**
 * Math and trigonometry, plus the aggregates that share their machinery.
 *
 * Four decisions shape everything below.
 *
 * First, the range/argument asymmetry. `SUM(A1:A3)` ignores text and booleans
 * sitting in those cells, but `SUM(TRUE)` is 1 and `SUM("3")` is 3. Excel
 * decides by where the value came from, not by what it is, so every aggregate
 * walks its arguments through `gather`, which tags each value as direct (typed
 * into the formula) or indirect (read out of a reference or array) and applies
 * the two rules separately. Collapsing that distinction is the single most
 * commonly reported difference between a spreadsheet clone and Excel.
 *
 * Second, references are walked with `ctx.iterate`, never materialised. A
 * `SUM(A:A)` must cost the number of cells that exist, not a million; iterate
 * also skips truly blank cells for free, which is exactly the semantics COUNT,
 * AVERAGE and MIN want. Only the criteria functions and the matrix functions
 * dereference, because they need positional alignment.
 *
 * Third, rounding happens on Excel's fifteen-digit decimal view of a number,
 * not on the raw double. `ROUND(2.675,2)` is 2.68 in Excel although the stored
 * double is 2.67499999999999982, because Excel rounds what it would display.
 * `roundMagnitude` therefore snaps the input to fifteen significant digits and
 * shifts the decimal exponent through the string form rather than multiplying
 * by a power of ten, which would reintroduce the error it is trying to avoid.
 * The rounding itself is half away from zero: `ROUND(-2.5,0)` is -3, so naive
 * `Math.round` (which gives -2) is never used on a signed value.
 *
 * Fourth, transcendental results are left as full doubles. `LOG(8,2)` is
 * 2.9999999999999996 here, as it is in IEEE arithmetic; value.ts already
 * rounds to fifteen digits at comparison and display time, which is where Excel
 * does it too. Rounding at the source instead would corrupt any later
 * multiplication by a large factor.
 *
 * Two behaviours in this file are deliberately incomplete, and silently wrong
 * answers were preferred to none only where Excel's own answer coincides:
 * SUBTOTAL's 101-111 forms cannot exclude hidden rows and AGGREGATE's
 * hidden-row options cannot either, because FunctionContext exposes cell values
 * but no row visibility; likewise neither can skip nested SUBTOTAL/AGGREGATE
 * cells, which would need the formula text of the cells being summed. On a
 * sheet with nothing hidden and no nesting - the overwhelmingly common case -
 * the answers agree with Excel; otherwise they over-count. See the note on
 * SUBTOTAL below.
 */

import { CellError, type Scalar, isError } from '@mirrorz/core';
import {
  ArgKind,
  type FunctionContext,
  type FunctionSpec,
  type ParamSpec,
  p,
} from '../registry.js';
import {
  type ArrayValue,
  type Criterion,
  type RefValue,
  type Value,
  checkMagnitude,
  excelAdd,
  excelSub,
  isArray,
  isRef,
  makeArray,
  makeRef,
  matchesCriterion,
  parseCriterion,
  toExcelPrecision,
  toNumber,
  toText,
} from '../value.js';

// ---------------------------------------------------------------------------
// Argument plumbing
// ---------------------------------------------------------------------------

/** The scalar an ArgKind.Scalar parameter delivered, with omissions as blank. */
function scalarArg(v: Value | undefined): Scalar {
  if (v === undefined) return null;
  if (isArray(v)) return v.data[0] ?? null;
  if (isRef(v)) return null;
  return v;
}

/**
 * The same, for a parameter declared ArgKind.Any: the evaluator leaves a
 * reference undereferenced there, so SUMIF's criteria and AGGREGATE's k have to
 * read the cell themselves.
 */
function anyScalarArg(v: Value | undefined, ctx: FunctionContext): Scalar {
  if (isRef(v)) return ctx.getScalar(v.sheet, v.startRow, v.startCol);
  return scalarArg(v);
}

function numArg(v: Value | undefined, whenOmitted?: number): number | CellError {
  if (v === undefined && whenOmitted !== undefined) return whenOmitted;
  return toNumber(scalarArg(v));
}

/** A numeric argument Excel truncates towards zero before using. */
function intArg(v: Value | undefined, whenOmitted?: number): number | CellError {
  const n = numArg(v, whenOmitted);
  return isError(n) ? n : Math.trunc(n);
}

/** Range-check a computed result the way Excel does: overflow is #NUM!. */
function finite(v: number): number | CellError {
  return checkMagnitude(v);
}

/** One value pulled out of an argument, with its provenance. */
interface Item {
  value: Scalar;
  /** Typed directly into the formula rather than read from a range or array. */
  direct: boolean;
}

function* items(arg: Value | undefined, ctx: FunctionContext): Generator<Item> {
  if (arg === undefined) return;
  if (isRef(arg)) {
    // Blank cells never surface here, which is what COUNT and AVERAGE want.
    for (const cell of ctx.iterate(arg)) yield { value: cell.value, direct: false };
    return;
  }
  if (isArray(arg)) {
    for (const v of arg.data) yield { value: v, direct: false };
    return;
  }
  yield { value: arg, direct: true };
}

interface Gathered {
  numbers: number[];
  /** Values that are not blank, for COUNTA. */
  nonEmpty: number;
  /** A real error value met along the way. */
  error?: CellError;
  /**
   * A direct argument that is text and does not parse as a number. SUM reports
   * it as #VALUE!; COUNT ignores it, which is why the two are kept apart.
   */
  conversionError?: CellError;
}

/**
 * Collect the numbers an aggregate should see.
 *
 * Direct arguments are coerced (text that looks numeric converts, booleans
 * become 1 and 0); values reached through a reference or an array are used only
 * when they are already numbers.
 */
function gather(
  args: readonly (Value | undefined)[],
  ctx: FunctionContext,
  ignoreErrors = false,
): Gathered {
  const out: Gathered = { numbers: [], nonEmpty: 0 };
  for (const arg of args) {
    for (const item of items(arg, ctx)) {
      const v = item.value;
      if (isError(v)) {
        if (!ignoreErrors && !out.error) out.error = v;
        continue;
      }
      if (v !== null) out.nonEmpty++;
      if (item.direct) {
        const n = toNumber(v);
        if (isError(n)) {
          if (!out.conversionError) out.conversionError = n;
          continue;
        }
        out.numbers.push(n);
        continue;
      }
      if (typeof v === 'number') out.numbers.push(v);
    }
  }
  return out;
}

/** The error an ordinary numeric aggregate should report, if any. */
function gatherError(g: Gathered): CellError | undefined {
  return g.error ?? g.conversionError;
}

/** A rectangular block, for the functions that need positional alignment. */
function block(v: Value | undefined, ctx: FunctionContext): ArrayValue {
  if (v === undefined) return makeArray(0, 0, []);
  if (isRef(v)) return ctx.deref(v);
  if (isArray(v)) return v;
  return makeArray(1, 1, [v]);
}

/**
 * SUMIF's sum_range is not read as written: Excel anchors it at its top-left
 * corner and then takes the same shape as the criteria range, so
 * `SUMIF(A1:A9,">1",B1)` sums B1:B9.
 */
function alignedBlock(
  v: Value | undefined,
  rows: number,
  cols: number,
  ctx: FunctionContext,
): ArrayValue {
  if (isRef(v)) {
    const anchored: RefValue = makeRef(
      v.sheet,
      v.startRow,
      v.startCol,
      v.startRow + rows - 1,
      v.startCol + cols - 1,
    );
    return ctx.deref(anchored);
  }
  const b = block(v, ctx);
  if (b.rows === rows && b.cols === cols) return b;
  const data: Scalar[] = new Array(rows * cols).fill(null);
  for (let r = 0; r < Math.min(rows, b.rows); r++) {
    for (let c = 0; c < Math.min(cols, b.cols); c++) {
      data[r * cols + c] = b.data[r * b.cols + c] ?? null;
    }
  }
  return makeArray(rows, cols, data);
}

// ---------------------------------------------------------------------------
// Aggregate primitives
// ---------------------------------------------------------------------------

/** Addition goes through excelAdd so cancellation snaps as it does in Excel. */
function sumOf(nums: readonly number[]): number {
  let total = 0;
  for (const n of nums) total = excelAdd(total, n);
  return total;
}

function productOf(nums: readonly number[]): number {
  // An empty PRODUCT is 0, not the mathematical identity 1.
  if (nums.length === 0) return 0;
  let acc = 1;
  for (const n of nums) acc *= n;
  return acc;
}

function averageOf(nums: readonly number[]): number | CellError {
  if (nums.length === 0) return CellError.DIV0;
  return sumOf(nums) / nums.length;
}

function varianceOf(nums: readonly number[], sample: boolean): number | CellError {
  const n = nums.length;
  if (n < (sample ? 2 : 1)) return CellError.DIV0;
  const mean = sumOf(nums) / n;
  let ss = 0;
  for (const x of nums) {
    const d = excelSub(x, mean);
    ss += d * d;
  }
  return ss / (sample ? n - 1 : n);
}

function medianOf(nums: readonly number[]): number | CellError {
  if (nums.length === 0) return CellError.NUM;
  const s = [...nums].sort((a, b) => a - b);
  const mid = s.length >> 1;
  return s.length % 2 === 1 ? s[mid]! : excelAdd(s[mid - 1]!, s[mid]!) / 2;
}

function modeOf(nums: readonly number[]): number | CellError {
  let best: number | undefined;
  let bestCount = 1;
  for (let i = 0; i < nums.length; i++) {
    let count = 0;
    for (const x of nums) if (x === nums[i]) count++;
    if (count > bestCount) {
      bestCount = count;
      best = nums[i];
    }
  }
  return best === undefined ? CellError.NA : best;
}

function nthOf(nums: readonly number[], k: number, largest: boolean): number | CellError {
  const rank = Math.trunc(k);
  if (nums.length === 0 || rank < 1 || rank > nums.length) return CellError.NUM;
  const s = [...nums].sort((a, b) => a - b);
  return largest ? s[s.length - rank]! : s[rank - 1]!;
}

function percentileInc(nums: readonly number[], q: number): number | CellError {
  if (nums.length === 0 || q < 0 || q > 1) return CellError.NUM;
  const s = [...nums].sort((a, b) => a - b);
  const pos = q * (s.length - 1);
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  if (lo === hi) return s[lo]!;
  return excelAdd(s[lo]!, (pos - lo) * excelSub(s[hi]!, s[lo]!));
}

function percentileExc(nums: readonly number[], q: number): number | CellError {
  const n = nums.length;
  if (n === 0) return CellError.NUM;
  // The exclusive form cannot reach the extremes: with n values only the open
  // interval (1/(n+1), n/(n+1)) is defined.
  if (q <= 0 || q >= 1 || q < 1 / (n + 1) || q > n / (n + 1)) return CellError.NUM;
  const s = [...nums].sort((a, b) => a - b);
  const pos = q * (n + 1) - 1;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  if (lo === hi) return s[lo]!;
  return excelAdd(s[lo]!, (pos - lo) * excelSub(s[hi]!, s[lo]!));
}

function quartile(nums: readonly number[], q: number, inclusive: boolean): number | CellError {
  const k = Math.trunc(q);
  if (k < 0 || k > 4) return CellError.NUM;
  if (inclusive) return percentileInc(nums, k / 4);
  if (k === 0 || k === 4) return CellError.NUM;
  return percentileExc(nums, k / 4);
}

/**
 * The shared body of SUBTOTAL and AGGREGATE, addressed by Excel's function
 * numbers. 1-11 are the SUBTOTAL set; 12-19 exist only in AGGREGATE.
 */
function applyAggregate(
  fn: number,
  g: Gathered,
  k: number | undefined,
): number | CellError {
  const nums = g.numbers;
  switch (fn) {
    case 1:
      return averageOf(nums);
    case 2:
      return nums.length;
    case 3:
      return g.nonEmpty;
    case 4:
      return nums.length === 0 ? 0 : Math.max(...nums);
    case 5:
      return nums.length === 0 ? 0 : Math.min(...nums);
    case 6:
      return productOf(nums);
    case 7: {
      const v = varianceOf(nums, true);
      return isError(v) ? v : Math.sqrt(v);
    }
    case 8: {
      const v = varianceOf(nums, false);
      return isError(v) ? v : Math.sqrt(v);
    }
    case 9:
      return sumOf(nums);
    case 10:
      return varianceOf(nums, true);
    case 11:
      return varianceOf(nums, false);
    case 12:
      return medianOf(nums);
    case 13:
      return modeOf(nums);
    case 14:
      return k === undefined ? CellError.VALUE : nthOf(nums, k, true);
    case 15:
      return k === undefined ? CellError.VALUE : nthOf(nums, k, false);
    case 16:
      return k === undefined ? CellError.VALUE : percentileInc(nums, k);
    case 17:
      return k === undefined ? CellError.VALUE : quartile(nums, k, true);
    case 18:
      return k === undefined ? CellError.VALUE : percentileExc(nums, k);
    case 19:
      return k === undefined ? CellError.VALUE : quartile(nums, k, false);
    default:
      return CellError.VALUE;
  }
}

/** COUNT and COUNTA see a #VALUE!-producing direct text argument as ignorable. */
function errorFor(fn: number, g: Gathered): CellError | undefined {
  if (g.error) return g.error;
  return fn === 2 || fn === 3 ? undefined : g.conversionError;
}

// ---------------------------------------------------------------------------
// Rounding
// ---------------------------------------------------------------------------

/**
 * Move a number's decimal point by `by` places without multiplying.
 *
 * `2.675 * 100` is 267.49999999999997, which rounds to 267 and gives the wrong
 * answer; rewriting the exponent of the decimal string gives exactly 267.5.
 */
function shiftExponent(v: number, by: number): number {
  if (v === 0 || !Number.isFinite(v)) return v;
  const [mantissa, exponent] = v.toExponential().split('e');
  return Number(`${mantissa}e${Number(exponent) + by}`);
}

/**
 * Round at a decimal place, always applying `roundMagnitude` to the absolute
 * value so that half-way cases go away from zero in both directions.
 */
function roundAt(x: number, digits: number, roundMagnitude: (n: number) => number): number {
  if (x === 0 || !Number.isFinite(x)) return x;
  // Excel rounds the fifteen-digit number it would display, not the double.
  const value = toExcelPrecision(x);
  const sign = value < 0 ? -1 : 1;
  const d = Math.max(-330, Math.min(330, Math.trunc(digits)));
  const shifted = shiftExponent(Math.abs(value), d);
  // More digits than a double can hold leaves the value untouched, as in Excel.
  if (!Number.isFinite(shifted)) return value;
  return sign * shiftExponent(roundMagnitude(shifted), -d);
}

/** Half away from zero, on a value already known to be non-negative. */
const halfUp = (n: number): number => Math.round(n);

/**
 * Round a quotient to Excel's fifteen digits before taking its floor or
 * ceiling. `0.29/0.01` is 28.999999999999996, and flooring that raw would make
 * FLOOR(0.29,0.01) report 0.28.
 */
function quotientFor(n: number, s: number): number {
  return toExcelPrecision(n / s);
}

// ---------------------------------------------------------------------------
// Spec builders
// ---------------------------------------------------------------------------

function unary(
  name: string,
  fn: (x: number) => number | CellError,
  summary: string,
): FunctionSpec {
  return {
    name,
    params: [p.scalar('number')],
    broadcast: true,
    summary,
    impl: (args) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const r = fn(x);
      return isError(r) ? r : finite(r);
    },
  };
}

function binary(
  name: string,
  first: string,
  second: string,
  fn: (a: number, b: number) => number | CellError,
  summary: string,
): FunctionSpec {
  return {
    name,
    params: [p.scalar(first), p.scalar(second)],
    broadcast: true,
    summary,
    impl: (args) => {
      const a = numArg(args[0]);
      if (isError(a)) return a;
      const b = numArg(args[1]);
      if (isError(b)) return b;
      const r = fn(a, b);
      return isError(r) ? r : finite(r);
    },
  };
}

// ---------------------------------------------------------------------------
// Aggregates
// ---------------------------------------------------------------------------

const SUM: FunctionSpec = {
  name: 'SUM',
  params: [p.any('number1'), p.rest('number2', ArgKind.Any)],
  summary: 'Adds its arguments.',
  impl: (args, ctx) => {
    const g = gather(args, ctx);
    const err = gatherError(g);
    return err ?? sumOf(g.numbers);
  },
};

const PRODUCT: FunctionSpec = {
  name: 'PRODUCT',
  params: [p.any('number1'), p.rest('number2', ArgKind.Any)],
  summary: 'Multiplies its arguments.',
  impl: (args, ctx) => {
    const g = gather(args, ctx);
    const err = gatherError(g);
    return err ?? finite(productOf(g.numbers));
  },
};

const SUMSQ: FunctionSpec = {
  name: 'SUMSQ',
  params: [p.any('number1'), p.rest('number2', ArgKind.Any)],
  summary: 'Returns the sum of the squares of its arguments.',
  impl: (args, ctx) => {
    const g = gather(args, ctx);
    const err = gatherError(g);
    if (err) return err;
    return finite(sumOf(g.numbers.map((n) => n * n)));
  },
};

/** SUMIF and SUMIFS both reduce to "walk aligned blocks, test every row". */
function sumMatching(
  sumBlock: ArrayValue,
  tests: { block: ArrayValue; criterion: Criterion }[],
): Value {
  let total = 0;
  const cells = sumBlock.rows * sumBlock.cols;
  for (let i = 0; i < cells; i++) {
    let hit = true;
    for (const t of tests) {
      // An error in a criteria range simply fails to match; only the summed
      // value propagates one.
      if (!matchesCriterion(t.block.data[i] ?? null, t.criterion)) {
        hit = false;
        break;
      }
    }
    if (!hit) continue;
    const v = sumBlock.data[i] ?? null;
    if (isError(v)) return v;
    if (typeof v === 'number') total = excelAdd(total, v);
  }
  return total;
}

const SUMIF: FunctionSpec = {
  name: 'SUMIF',
  params: [p.any('range'), p.scalar('criteria'), p.any('sum_range', true)],
  summary: 'Adds the cells specified by a given condition.',
  impl: (args, ctx) => {
    const range = block(args[0], ctx);
    const criterion = parseCriterion(anyScalarArg(args[1], ctx));
    const sumBlock =
      args[2] === undefined ? range : alignedBlock(args[2], range.rows, range.cols, ctx);
    return sumMatching(sumBlock, [{ block: range, criterion }]);
  },
};

const SUMIFS: FunctionSpec = {
  name: 'SUMIFS',
  params: [
    p.any('sum_range'),
    p.any('criteria_range1'),
    p.scalar('criteria1'),
    p.rest('more', ArgKind.Any),
  ],
  summary: 'Adds the cells specified by a given set of conditions.',
  impl: (args, ctx) => {
    // Pairs after the sum range: an even tail means a criterion has no range.
    if (args.length < 3 || (args.length - 1) % 2 !== 0) return CellError.VALUE;
    const sumBlock = block(args[0], ctx);
    const tests: { block: ArrayValue; criterion: Criterion }[] = [];
    for (let i = 1; i < args.length; i += 2) {
      const b = block(args[i], ctx);
      // Unlike SUMIF, SUMIFS requires every range to have the sum range's shape.
      if (b.rows !== sumBlock.rows || b.cols !== sumBlock.cols) return CellError.VALUE;
      tests.push({ block: b, criterion: parseCriterion(anyScalarArg(args[i + 1], ctx)) });
    }
    return sumMatching(sumBlock, tests);
  },
};

const SUMPRODUCT: FunctionSpec = {
  name: 'SUMPRODUCT',
  params: [p.array('array1'), p.rest('array2', ArgKind.Array)],
  summary: 'Returns the sum of the products of corresponding array components.',
  impl: (args, ctx) => {
    const blocks = args.filter((a) => a !== undefined).map((a) => block(a, ctx));
    if (blocks.length === 0) return CellError.VALUE;
    const first = blocks[0]!;
    for (const b of blocks) {
      if (b.rows !== first.rows || b.cols !== first.cols) return CellError.VALUE;
    }
    let total = 0;
    const cells = first.rows * first.cols;
    for (let i = 0; i < cells; i++) {
      let product = 1;
      for (const b of blocks) {
        const v = b.data[i] ?? null;
        if (isError(v)) return v;
        // Text, blanks and booleans count as zero rather than erroring.
        product *= typeof v === 'number' ? v : 0;
      }
      total = excelAdd(total, product);
    }
    return finite(total);
  },
};

/** The three paired-array sums differ only in what they do with each pair. */
function pairwiseSum(
  name: string,
  combine: (x: number, y: number) => number,
  summary: string,
): FunctionSpec {
  return {
    name,
    params: [p.array('array_x'), p.array('array_y')],
    summary,
    impl: (args, ctx) => {
      const x = block(args[0], ctx);
      const y = block(args[1], ctx);
      const nx = x.rows * x.cols;
      const ny = y.rows * y.cols;
      // Excel compares counts, not shapes: a row and a column of equal length
      // pair up happily.
      if (nx !== ny) return CellError.NA;
      let total = 0;
      for (let i = 0; i < nx; i++) {
        const a = x.data[i] ?? null;
        const b = y.data[i] ?? null;
        if (isError(a)) return a;
        if (isError(b)) return b;
        // A pair is skipped entirely unless both halves are numbers.
        if (typeof a !== 'number' || typeof b !== 'number') continue;
        total = excelAdd(total, combine(a, b));
      }
      return finite(total);
    },
  };
}

/**
 * SUBTOTAL.
 *
 * The 101-111 forms are meant to skip manually hidden rows and every form is
 * meant to skip cells that are themselves SUBTOTAL formulas. FunctionContext
 * hands us cell values only - no row visibility and no formula text - so both
 * exclusions are accepted and then not applied. The answer is Excel's whenever
 * nothing is hidden and no SUBTOTAL is nested, and too large otherwise; that is
 * a narrower failure than rejecting the function outright, and it becomes
 * correct the moment the context grows a row-visibility query.
 */
const SUBTOTAL: FunctionSpec = {
  name: 'SUBTOTAL',
  params: [p.scalar('function_num'), p.rest('ref', ArgKind.Any)],
  summary: 'Returns a subtotal in a list or database.',
  impl: (args, ctx) => {
    const raw = intArg(args[0]);
    if (isError(raw)) return raw;
    const fn = raw > 100 ? raw - 100 : raw;
    if (fn < 1 || fn > 11 || (raw > 11 && raw < 101)) return CellError.VALUE;
    const rest = args.slice(1).filter((a) => a !== undefined);
    if (rest.length === 0) return CellError.VALUE;
    const g = gather(rest, ctx);
    const err = errorFor(fn, g);
    if (err) return err;
    const r = applyAggregate(fn, g, undefined);
    return isError(r) ? r : finite(r);
  },
};

/**
 * AGGREGATE.
 *
 * Options 2, 3, 6 and 7 ignore error values, which is the reason the trailing
 * parameter is declared error-transparent: the evaluator would otherwise
 * short-circuit on `AGGREGATE(9,6,1/0)` before we could ignore anything. The
 * hidden-row options (1, 3, 5, 7) and the always-on "ignore nested
 * SUBTOTAL/AGGREGATE" rule are accepted and not applied, for the reason given
 * on SUBTOTAL.
 */
const AGGREGATE: FunctionSpec = {
  name: 'AGGREGATE',
  params: [
    p.scalar('function_num'),
    p.scalar('options'),
    { name: 'ref', kind: ArgKind.Any, repeating: true, optional: true, errorTransparent: true },
  ],
  summary: 'Returns an aggregate in a list or database, with errors optionally ignored.',
  impl: (args, ctx) => {
    const fn = intArg(args[0]);
    if (isError(fn)) return fn;
    const options = intArg(args[1], 0);
    if (isError(options)) return options;
    if (options < 0 || options > 7) return CellError.VALUE;
    if (fn < 1 || fn > 19) return CellError.VALUE;

    const rest = args.slice(2).filter((a) => a !== undefined);
    if (rest.length === 0) return CellError.VALUE;

    const ignoreErrors = options === 2 || options === 3 || options === 6 || options === 7;

    // 14-19 take a single array plus a k argument; 1-13 take references only.
    let k: number | undefined;
    let data = rest;
    if (fn >= 14) {
      if (rest.length < 2) return CellError.VALUE;
      const n = toNumber(anyScalarArg(rest[rest.length - 1], ctx));
      if (isError(n)) return ignoreErrors ? CellError.VALUE : n;
      k = n;
      data = rest.slice(0, -1);
    }

    const g = gather(data, ctx, ignoreErrors);
    const err = errorFor(fn, g);
    if (err) return err;
    const r = applyAggregate(fn, g, k);
    return isError(r) ? r : finite(r);
  },
};

// ---------------------------------------------------------------------------
// Rounding and sign
// ---------------------------------------------------------------------------

const ROUND_FAMILY: FunctionSpec[] = [
  {
    name: 'ROUND',
    params: [p.scalar('number'), p.scalar('num_digits')],
    broadcast: true,
    summary: 'Rounds a number to a specified number of digits.',
    impl: (args) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const d = numArg(args[1]);
      if (isError(d)) return d;
      return finite(roundAt(x, d, halfUp));
    },
  },
  {
    name: 'ROUNDUP',
    params: [p.scalar('number'), p.scalar('num_digits')],
    broadcast: true,
    summary: 'Rounds a number up, away from zero.',
    impl: (args) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const d = numArg(args[1]);
      if (isError(d)) return d;
      return finite(roundAt(x, d, Math.ceil));
    },
  },
  {
    name: 'ROUNDDOWN',
    params: [p.scalar('number'), p.scalar('num_digits')],
    broadcast: true,
    summary: 'Rounds a number down, towards zero.',
    impl: (args) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const d = numArg(args[1]);
      if (isError(d)) return d;
      return finite(roundAt(x, d, Math.floor));
    },
  },
  {
    name: 'MROUND',
    params: [p.scalar('number'), p.scalar('multiple')],
    broadcast: true,
    summary: 'Returns a number rounded to the desired multiple.',
    impl: (args) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const m = numArg(args[1]);
      if (isError(m)) return m;
      if (m === 0) return 0;
      if (x !== 0 && Math.sign(x) !== Math.sign(m)) return CellError.NUM;
      const q = quotientFor(x, m);
      return finite(toExcelPrecision(Math.sign(q) * Math.round(Math.abs(q)) * m));
    },
  },
];

/**
 * CEILING and FLOOR, the pre-2013 pair.
 *
 * `ceil(n/s)*s` and `floor(n/s)*s` reproduce every documented example once the
 * signs are taken as given rather than normalised: CEILING(-2.5,2) is -2 and
 * CEILING(-2.5,-2) is -4, FLOOR(-2.5,2) is -4 and FLOOR(-2.5,-2) is -2. Only a
 * positive number with a negative significance is an error, and only FLOOR
 * treats a zero significance as a division - CEILING answers 0. That asymmetry
 * is Excel's, not a slip.
 */
function legacyRounder(name: string, up: boolean, summary: string): FunctionSpec {
  return {
    name,
    params: [p.scalar('number'), p.scalar('significance')],
    broadcast: true,
    summary,
    impl: (args) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const s = numArg(args[1]);
      if (isError(s)) return s;
      if (x > 0 && s < 0) return CellError.NUM;
      if (s === 0) return up ? 0 : CellError.DIV0;
      if (x === 0) return 0;
      const q = quotientFor(x, s);
      return finite(toExcelPrecision((up ? Math.ceil(q) : Math.floor(q)) * s));
    },
  };
}

/**
 * The .MATH and .PRECISE forms ignore the sign of significance entirely. They
 * differ from each other only in that .MATH takes a mode that flips negative
 * numbers to round away from zero instead of towards negative infinity.
 */
function modernRounder(
  name: string,
  up: boolean,
  withMode: boolean,
  summary: string,
): FunctionSpec {
  const params: ParamSpec[] = [p.scalar('number'), p.scalar('significance', true)];
  if (withMode) params.push(p.scalar('mode', true));
  return {
    name,
    params,
    broadcast: true,
    summary,
    impl: (args) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const sRaw = numArg(args[1], 1);
      if (isError(sRaw)) return sRaw;
      const mode = withMode ? numArg(args[2], 0) : 0;
      if (isError(mode)) return mode;
      // An omitted significance is 1, but an explicit blank is 0 and wins.
      const s = Math.abs(sRaw);
      if (s === 0) return 0;
      if (x === 0) return 0;
      const q = quotientFor(x, s);
      // Away-from-zero mode inverts the direction for negative numbers only.
      const awayFromZero = x < 0 && mode !== 0;
      const roundUp = awayFromZero ? !up : up;
      return finite(toExcelPrecision((roundUp ? Math.ceil(q) : Math.floor(q)) * s));
    },
  };
}

// ---------------------------------------------------------------------------
// Combinatorics
// ---------------------------------------------------------------------------

function factorial(n: number): number | CellError {
  if (n < 0) return CellError.NUM;
  let acc = 1;
  for (let i = 2; i <= n; i++) acc *= i;
  return acc;
}

function doubleFactorial(n: number): number | CellError {
  if (n < 0) return CellError.NUM;
  let acc = 1;
  for (let i = n; i > 1; i -= 2) acc *= i;
  return acc;
}

/**
 * n choose k by the multiplicative formula, which stays inside double range far
 * longer than n!/(k!(n-k)!) and is exact for every result Excel can represent.
 */
function combinations(n: number, k: number): number | CellError {
  if (n < 0 || k < 0 || k > n) return CellError.NUM;
  const m = Math.min(k, n - k);
  let acc = 1;
  for (let i = 1; i <= m; i++) acc = (acc * (n - m + i)) / i;
  return Math.round(acc);
}

function permutations(n: number, k: number): number | CellError {
  if (n < 0 || k < 0 || k > n) return CellError.NUM;
  let acc = 1;
  for (let i = 0; i < k; i++) acc *= n - i;
  return acc;
}

function gcd2(a: number, b: number): number {
  let x = a;
  let y = b;
  while (y !== 0) {
    const t = x % y;
    x = y;
    y = t;
  }
  return x;
}

/** GCD and LCM share their argument handling and their two rejection rules. */
function integerListFn(
  name: string,
  reduce: (acc: number, next: number) => number | CellError,
  seed: number,
  summary: string,
): FunctionSpec {
  return {
    name,
    params: [p.any('number1'), p.rest('number2', ArgKind.Any)],
    summary,
    impl: (args, ctx) => {
      const g = gather(args, ctx);
      const err = gatherError(g);
      if (err) return err;
      let acc = seed;
      for (const raw of g.numbers) {
        const n = Math.trunc(raw);
        // Excel refuses negatives outright and cannot hold an exact integer
        // beyond 2^53, where the answer would silently stop being a divisor.
        if (n < 0 || n >= 2 ** 53) return CellError.NUM;
        const next = reduce(acc, n);
        if (isError(next)) return next;
        acc = next;
      }
      return acc;
    },
  };
}

// ---------------------------------------------------------------------------
// Numeral systems
// ---------------------------------------------------------------------------

const ROMAN_SYMBOLS = ['I', 'V', 'X', 'L', 'C', 'D', 'M'] as const;
const ROMAN_VALUES = [1, 5, 10, 50, 100, 500, 1000] as const;

/**
 * The tokens available at a given ROMAN form.
 *
 * Classic notation subtracts only I, X or C, and only from the next two symbols
 * up: for a big symbol at index j the classic subtractor is the largest even
 * index below j. Each increment of `form` allows the subtractor to move one
 * more place down the list, which is exactly what generates Microsoft's own
 * ladder for 1999: MCMXCIX, MLMVLIV, MXMIX, MVMIV, MIM.
 */
function romanTokens(form: number): { value: number; text: string }[] {
  const tokens: { value: number; text: string }[] = [];
  for (let j = 0; j < 7; j++) {
    tokens.push({ value: ROMAN_VALUES[j]!, text: ROMAN_SYMBOLS[j]! });
    if (j === 0) continue;
    const classic = 2 * Math.floor((j - 1) / 2);
    for (let i = classic; i >= 0 && classic - i <= form; i--) {
      tokens.push({
        value: ROMAN_VALUES[j]! - ROMAN_VALUES[i]!,
        text: ROMAN_SYMBOLS[i]! + ROMAN_SYMBOLS[j]!,
      });
    }
  }
  return tokens.sort((a, b) => b.value - a.value);
}

const DIGITS = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';

// ---------------------------------------------------------------------------
// Matrices
// ---------------------------------------------------------------------------

/** A numeric matrix, or the error Excel reports for a non-numeric cell. */
function numericMatrix(b: ArrayValue): number[][] | CellError {
  if (b.rows === 0 || b.cols === 0) return CellError.VALUE;
  const m: number[][] = [];
  for (let r = 0; r < b.rows; r++) {
    const row: number[] = [];
    for (let c = 0; c < b.cols; c++) {
      const v = b.data[r * b.cols + c] ?? null;
      if (isError(v)) return v;
      // The matrix functions accept numbers only: blanks, text and booleans are
      // all #VALUE!, which is what Excel documents.
      if (typeof v !== 'number') return CellError.VALUE;
      row.push(v);
    }
    m.push(row);
  }
  return m;
}

/** LU decomposition with partial pivoting; returns the determinant. */
function determinant(m: readonly (readonly number[])[]): number {
  const n = m.length;
  const a = m.map((row) => [...row]);
  let det = 1;
  for (let i = 0; i < n; i++) {
    let pivot = i;
    for (let r = i + 1; r < n; r++) {
      if (Math.abs(a[r]![i]!) > Math.abs(a[pivot]![i]!)) pivot = r;
    }
    if (a[pivot]![i] === 0) return 0;
    if (pivot !== i) {
      const t = a[i]!;
      a[i] = a[pivot]!;
      a[pivot] = t;
      det = -det;
    }
    det *= a[i]![i]!;
    for (let r = i + 1; r < n; r++) {
      const f = a[r]![i]! / a[i]![i]!;
      for (let c = i; c < n; c++) a[r]![c] = excelSub(a[r]![c]!, f * a[i]![c]!);
    }
  }
  return det;
}

/** Gauss-Jordan inversion; null when the matrix is singular. */
function invert(m: readonly (readonly number[])[]): number[][] | null {
  const n = m.length;
  const a = m.map((row) => [...row]);
  const inv: number[][] = Array.from({ length: n }, (_, i) =>
    Array.from({ length: n }, (_, j) => (i === j ? 1 : 0)),
  );
  for (let i = 0; i < n; i++) {
    let pivot = i;
    for (let r = i + 1; r < n; r++) {
      if (Math.abs(a[r]![i]!) > Math.abs(a[pivot]![i]!)) pivot = r;
    }
    if (a[pivot]![i] === 0) return null;
    if (pivot !== i) {
      const t = a[i]!;
      a[i] = a[pivot]!;
      a[pivot] = t;
      const u = inv[i]!;
      inv[i] = inv[pivot]!;
      inv[pivot] = u;
    }
    const d = a[i]![i]!;
    for (let c = 0; c < n; c++) {
      a[i]![c] = a[i]![c]! / d;
      inv[i]![c] = inv[i]![c]! / d;
    }
    for (let r = 0; r < n; r++) {
      if (r === i) continue;
      const f = a[r]![i]!;
      if (f === 0) continue;
      for (let c = 0; c < n; c++) {
        a[r]![c] = excelSub(a[r]![c]!, f * a[i]![c]!);
        inv[r]![c] = excelSub(inv[r]![c]!, f * inv[i]![c]!);
      }
    }
  }
  return inv;
}

// ---------------------------------------------------------------------------
// The table
// ---------------------------------------------------------------------------

export const MATH_FUNCTIONS: readonly FunctionSpec[] = [
  SUM,
  SUMIF,
  SUMIFS,
  SUMPRODUCT,
  SUMSQ,
  pairwiseSum('SUMX2MY2', (x, y) => x * x - y * y, 'Sums the difference of squares of two arrays.'),
  pairwiseSum('SUMX2PY2', (x, y) => x * x + y * y, 'Sums the sum of squares of two arrays.'),
  pairwiseSum('SUMXMY2', (x, y) => (x - y) ** 2, 'Sums the squares of differences of two arrays.'),
  PRODUCT,
  SUBTOTAL,
  AGGREGATE,

  unary('ABS', (x) => Math.abs(x), 'Returns the absolute value of a number.'),
  unary('SIGN', (x) => Math.sign(x), 'Returns the sign of a number.'),
  unary('INT', (x) => Math.floor(toExcelPrecision(x)), 'Rounds a number down to the nearest integer.'),
  {
    name: 'TRUNC',
    params: [p.scalar('number'), p.scalar('num_digits', true)],
    broadcast: true,
    summary: 'Truncates a number to an integer.',
    impl: (args) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const d = numArg(args[1], 0);
      if (isError(d)) return d;
      return finite(roundAt(x, d, Math.floor));
    },
  },
  ...ROUND_FAMILY,
  legacyRounder('CEILING', true, 'Rounds a number to the nearest multiple of significance.'),
  legacyRounder('FLOOR', false, 'Rounds a number down to the nearest multiple of significance.'),
  modernRounder('CEILING.MATH', true, true, 'Rounds a number up to the nearest integer or multiple.'),
  modernRounder('FLOOR.MATH', false, true, 'Rounds a number down to the nearest integer or multiple.'),
  modernRounder('CEILING.PRECISE', true, false, 'Rounds a number up, regardless of its sign.'),
  modernRounder('FLOOR.PRECISE', false, false, 'Rounds a number down, regardless of its sign.'),
  unary(
    'EVEN',
    (x) => {
      const v = toExcelPrecision(x);
      const sign = v < 0 ? -1 : 1;
      return sign * 2 * Math.ceil(Math.abs(v) / 2);
    },
    'Rounds a number up to the nearest even integer.',
  ),
  unary(
    'ODD',
    (x) => {
      const v = toExcelPrecision(x);
      if (v === 0) return 1;
      const sign = v < 0 ? -1 : 1;
      const mag = Math.abs(v);
      return sign * (2 * Math.ceil((mag + 1) / 2) - 1);
    },
    'Rounds a number up to the nearest odd integer.',
  ),

  {
    name: 'MOD',
    params: [p.scalar('number'), p.scalar('divisor')],
    broadcast: true,
    summary: 'Returns the remainder from division.',
    impl: (args) => {
      const n = numArg(args[0]);
      if (isError(n)) return n;
      const d = numArg(args[1]);
      if (isError(d)) return d;
      if (d === 0) return CellError.DIV0;
      // n - d*INT(n/d) rather than JavaScript's %, so the result takes the sign
      // of the divisor: MOD(-7,3) is 2, not -1.
      const q = Math.floor(quotientFor(n, d));
      if (!Number.isFinite(q)) return CellError.NUM;
      return finite(excelSub(n, d * q));
    },
  },
  binary(
    'QUOTIENT',
    'numerator',
    'denominator',
    (a, b) => (b === 0 ? CellError.DIV0 : Math.trunc(a / b)),
    'Returns the integer portion of a division.',
  ),
  binary(
    'POWER',
    'number',
    'power',
    (x, y) => {
      if (x === 0 && y === 0) return CellError.NUM;
      if (x === 0 && y < 0) return CellError.DIV0;
      const r = x ** y;
      // A negative base with a fractional exponent is #NUM!, not NaN.
      return Number.isNaN(r) ? CellError.NUM : r;
    },
    'Returns the result of a number raised to a power.',
  ),
  unary('SQRT', (x) => (x < 0 ? CellError.NUM : Math.sqrt(x)), 'Returns a positive square root.'),
  unary(
    'SQRTPI',
    (x) => (x < 0 ? CellError.NUM : Math.sqrt(x * Math.PI)),
    'Returns the square root of a number times pi.',
  ),
  unary('EXP', (x) => Math.exp(x), 'Returns e raised to the power of a number.'),
  unary('LN', (x) => (x <= 0 ? CellError.NUM : Math.log(x)), 'Returns the natural logarithm.'),
  unary('LOG10', (x) => (x <= 0 ? CellError.NUM : Math.log10(x)), 'Returns the base-10 logarithm.'),
  {
    name: 'LOG',
    params: [p.scalar('number'), p.scalar('base', true)],
    broadcast: true,
    summary: 'Returns the logarithm of a number to a specified base.',
    impl: (args) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const base = numArg(args[1], 10);
      if (isError(base)) return base;
      if (x <= 0 || base < 0) return CellError.NUM;
      const denominator = Math.log(base);
      // LOG(n,1) divides by ln(1); LOG(n,0) takes ln(0) and is out of domain.
      if (base === 0) return CellError.NUM;
      if (denominator === 0) return CellError.DIV0;
      return finite(Math.log(x) / denominator);
    },
  },
  {
    name: 'PI',
    params: [],
    summary: 'Returns the value of pi.',
    impl: () => Math.PI,
  },

  unary('SIN', (x) => Math.sin(x), 'Returns the sine of an angle.'),
  unary('COS', (x) => Math.cos(x), 'Returns the cosine of an angle.'),
  unary('TAN', (x) => Math.tan(x), 'Returns the tangent of an angle.'),
  unary(
    'ASIN',
    (x) => (x < -1 || x > 1 ? CellError.NUM : Math.asin(x)),
    'Returns the arcsine of a number.',
  ),
  unary(
    'ACOS',
    (x) => (x < -1 || x > 1 ? CellError.NUM : Math.acos(x)),
    'Returns the arccosine of a number.',
  ),
  unary('ATAN', (x) => Math.atan(x), 'Returns the arctangent of a number.'),
  binary(
    'ATAN2',
    'x_num',
    'y_num',
    // Excel's argument order is (x, y); Math.atan2 takes (y, x).
    (x, y) => (x === 0 && y === 0 ? CellError.DIV0 : Math.atan2(y, x)),
    'Returns the arctangent from x- and y-coordinates.',
  ),
  unary('SINH', (x) => Math.sinh(x), 'Returns the hyperbolic sine of a number.'),
  unary('COSH', (x) => Math.cosh(x), 'Returns the hyperbolic cosine of a number.'),
  unary('TANH', (x) => Math.tanh(x), 'Returns the hyperbolic tangent of a number.'),
  unary('ASINH', (x) => Math.asinh(x), 'Returns the inverse hyperbolic sine of a number.'),
  unary(
    'ACOSH',
    (x) => (x < 1 ? CellError.NUM : Math.acosh(x)),
    'Returns the inverse hyperbolic cosine of a number.',
  ),
  unary(
    'ATANH',
    (x) => (x <= -1 || x >= 1 ? CellError.NUM : Math.atanh(x)),
    'Returns the inverse hyperbolic tangent of a number.',
  ),
  unary(
    'SEC',
    (x) => {
      const c = Math.cos(x);
      return c === 0 ? CellError.DIV0 : 1 / c;
    },
    'Returns the secant of an angle.',
  ),
  unary(
    'CSC',
    (x) => {
      const s = Math.sin(x);
      return s === 0 ? CellError.DIV0 : 1 / s;
    },
    'Returns the cosecant of an angle.',
  ),
  unary(
    'COT',
    (x) => {
      const s = Math.sin(x);
      return s === 0 ? CellError.DIV0 : Math.cos(x) / s;
    },
    'Returns the cotangent of an angle.',
  ),
  unary(
    'ACOT',
    // The principal value runs from 0 to pi, so this is not simply atan(1/x).
    (x) => Math.PI / 2 - Math.atan(x),
    'Returns the arccotangent of a number.',
  ),
  unary('SECH', (x) => 1 / Math.cosh(x), 'Returns the hyperbolic secant of an angle.'),
  unary(
    'CSCH',
    (x) => {
      const s = Math.sinh(x);
      return s === 0 ? CellError.DIV0 : 1 / s;
    },
    'Returns the hyperbolic cosecant of an angle.',
  ),
  unary(
    'COTH',
    (x) => {
      const s = Math.sinh(x);
      return s === 0 ? CellError.DIV0 : Math.cosh(x) / s;
    },
    'Returns the hyperbolic cotangent of an angle.',
  ),
  unary('DEGREES', (x) => (x * 180) / Math.PI, 'Converts radians to degrees.'),
  unary('RADIANS', (x) => (x * Math.PI) / 180, 'Converts degrees to radians.'),

  integerListFn(
    'GCD',
    (acc, n) => gcd2(acc, n),
    0,
    'Returns the greatest common divisor.',
  ),
  integerListFn(
    'LCM',
    (acc, n) => {
      if (acc === 0 || n === 0) return 0;
      const l = (acc / gcd2(acc, n)) * n;
      return l >= 2 ** 53 ? CellError.NUM : l;
    },
    1,
    'Returns the least common multiple.',
  ),
  unary('FACT', (x) => factorial(Math.trunc(x)), 'Returns the factorial of a number.'),
  unary(
    'FACTDOUBLE',
    (x) => doubleFactorial(Math.trunc(x)),
    'Returns the double factorial of a number.',
  ),
  binary(
    'COMBIN',
    'number',
    'number_chosen',
    (n, k) => combinations(Math.trunc(n), Math.trunc(k)),
    'Returns the number of combinations for a given number of objects.',
  ),
  binary(
    'COMBINA',
    'number',
    'number_chosen',
    (nRaw, kRaw) => {
      const n = Math.trunc(nRaw);
      const k = Math.trunc(kRaw);
      if (n < 0 || k < 0) return CellError.NUM;
      // Choosing from nothing is only possible when nothing is chosen.
      if (n === 0) return k === 0 ? 1 : CellError.NUM;
      return combinations(n + k - 1, k);
    },
    'Returns the number of combinations with repetitions.',
  ),
  binary(
    'PERMUT',
    'number',
    'number_chosen',
    (n, k) => permutations(Math.trunc(n), Math.trunc(k)),
    'Returns the number of permutations for a given number of objects.',
  ),
  binary(
    'PERMUTATIONA',
    'number',
    'number_chosen',
    (nRaw, kRaw) => {
      const n = Math.trunc(nRaw);
      const k = Math.trunc(kRaw);
      if (n < 0 || k < 0) return CellError.NUM;
      return n ** k;
    },
    'Returns the number of permutations with repetitions.',
  ),

  {
    name: 'RAND',
    params: [],
    volatile: true,
    summary: 'Returns a random number between 0 and 1.',
    impl: () => Math.random(),
  },
  {
    name: 'RANDBETWEEN',
    params: [p.scalar('bottom'), p.scalar('top')],
    volatile: true,
    summary: 'Returns a random number between the numbers you specify.',
    impl: (args) => {
      const lo = numArg(args[0]);
      if (isError(lo)) return lo;
      const hi = numArg(args[1]);
      if (isError(hi)) return hi;
      // Only whole numbers inside the range are candidates, so the bounds move
      // inwards rather than being truncated towards zero.
      const from = Math.ceil(lo);
      const to = Math.floor(hi);
      if (from > to) return CellError.NUM;
      return from + Math.floor(Math.random() * (to - from + 1));
    },
  },

  {
    name: 'ROMAN',
    params: [p.scalar('number'), p.scalar('form', true)],
    broadcast: true,
    summary: 'Converts an arabic numeral to roman, as text.',
    impl: (args) => {
      const n = numArg(args[0]);
      if (isError(n)) return n;
      const value = Math.trunc(n);
      if (value < 0 || value > 3999) return CellError.VALUE;
      const rawForm = scalarArg(args[1]);
      let form: number;
      if (args[1] === undefined || rawForm === null) form = 0;
      else if (typeof rawForm === 'boolean') form = rawForm ? 0 : 4;
      else {
        const f = toNumber(rawForm);
        if (isError(f)) return f;
        form = Math.trunc(f);
      }
      if (form < 0 || form > 4) return CellError.VALUE;

      let remaining = value;
      let out = '';
      for (const token of romanTokens(form)) {
        while (remaining >= token.value) {
          out += token.text;
          remaining -= token.value;
        }
        if (remaining === 0) break;
      }
      return out;
    },
  },
  {
    name: 'ARABIC',
    params: [p.scalar('text')],
    broadcast: true,
    summary: 'Converts a roman numeral to arabic, as a number.',
    impl: (args) => {
      const raw = scalarArg(args[0]);
      const text = toText(raw);
      if (isError(text)) return text;
      const trimmed = text.trim().toUpperCase();
      if (trimmed.length > 255) return CellError.VALUE;
      const negative = trimmed.startsWith('-');
      const body = negative ? trimmed.slice(1) : trimmed;
      if (body === '') return 0;

      const values: number[] = [];
      for (const ch of body) {
        const i = ROMAN_SYMBOLS.indexOf(ch as (typeof ROMAN_SYMBOLS)[number]);
        if (i < 0) return CellError.VALUE;
        values.push(ROMAN_VALUES[i]!);
      }
      // A symbol standing before a larger one is subtracted, which is all the
      // validation Excel does: it accepts strings no Roman ever wrote.
      let total = 0;
      for (let i = 0; i < values.length; i++) {
        const v = values[i]!;
        const next = values[i + 1];
        total += next !== undefined && next > v ? -v : v;
      }
      return negative ? -total : total;
    },
  },
  {
    name: 'BASE',
    params: [p.scalar('number'), p.scalar('radix'), p.scalar('min_length', true)],
    broadcast: true,
    summary: 'Converts a number into a text representation in the given radix.',
    impl: (args) => {
      const n = numArg(args[0]);
      if (isError(n)) return n;
      const radix = intArg(args[1]);
      if (isError(radix)) return radix;
      const minLength = intArg(args[2], 0);
      if (isError(minLength)) return minLength;
      const value = Math.trunc(n);
      if (value < 0 || value >= 2 ** 53) return CellError.NUM;
      if (radix < 2 || radix > 36) return CellError.NUM;
      if (minLength < 0 || minLength > 255) return CellError.NUM;
      let out = '';
      let rest = value;
      while (rest > 0) {
        out = DIGITS[rest % radix]! + out;
        rest = Math.floor(rest / radix);
      }
      if (out === '') out = '0';
      return out.padStart(minLength, '0');
    },
  },
  {
    name: 'DECIMAL',
    params: [p.scalar('text'), p.scalar('radix')],
    broadcast: true,
    summary: 'Converts a text representation in the given radix into a number.',
    impl: (args) => {
      const text = toText(scalarArg(args[0]));
      if (isError(text)) return text;
      const radix = intArg(args[1]);
      if (isError(radix)) return radix;
      if (radix < 2 || radix > 36) return CellError.NUM;
      const body = text.trim().toUpperCase();
      // No digits at all is zero, the value the accumulation below would reach.
      let total = 0;
      for (const ch of body) {
        const d = DIGITS.indexOf(ch);
        if (d < 0 || d >= radix) return CellError.NUM;
        total = total * radix + d;
      }
      return finite(total);
    },
  },

  {
    name: 'TRANSPOSE',
    params: [p.array('array')],
    summary: 'Returns the transpose of an array.',
    impl: (args, ctx) => {
      const b = block(args[0], ctx);
      const data: Scalar[] = new Array(b.rows * b.cols).fill(null);
      for (let r = 0; r < b.rows; r++) {
        for (let c = 0; c < b.cols; c++) {
          data[c * b.rows + r] = b.data[r * b.cols + c] ?? null;
        }
      }
      return makeArray(b.cols, b.rows, data);
    },
  },
  {
    name: 'MMULT',
    params: [p.array('array1'), p.array('array2')],
    summary: 'Returns the matrix product of two arrays.',
    impl: (args, ctx) => {
      const a = numericMatrix(block(args[0], ctx));
      if (isError(a)) return a;
      const b = numericMatrix(block(args[1], ctx));
      if (isError(b)) return b;
      const inner = a[0]!.length;
      if (inner !== b.length) return CellError.VALUE;
      const rows = a.length;
      const cols = b[0]!.length;
      const data: Scalar[] = new Array(rows * cols).fill(0);
      for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
          let acc = 0;
          for (let k = 0; k < inner; k++) acc = excelAdd(acc, a[i]![k]! * b[k]![j]!);
          const cell = finite(acc);
          if (isError(cell)) return cell;
          data[i * cols + j] = cell;
        }
      }
      return makeArray(rows, cols, data);
    },
  },
  {
    name: 'MDETERM',
    params: [p.array('array')],
    summary: 'Returns the matrix determinant of an array.',
    impl: (args, ctx) => {
      const m = numericMatrix(block(args[0], ctx));
      if (isError(m)) return m;
      if (m.length !== m[0]!.length) return CellError.VALUE;
      // Elimination residue is noise below Excel's fifteen digits, and Excel
      // reports the determinant of an integer matrix as an integer.
      return finite(toExcelPrecision(determinant(m)));
    },
  },
  {
    name: 'MINVERSE',
    params: [p.array('array')],
    summary: 'Returns the matrix inverse of an array.',
    impl: (args, ctx) => {
      const m = numericMatrix(block(args[0], ctx));
      if (isError(m)) return m;
      if (m.length !== m[0]!.length) return CellError.VALUE;
      const inv = invert(m);
      if (!inv) return CellError.NUM;
      const n = m.length;
      const data: Scalar[] = new Array(n * n);
      for (let r = 0; r < n; r++) {
        for (let c = 0; c < n; c++) data[r * n + c] = toExcelPrecision(inv[r]![c]!);
      }
      return makeArray(n, n, data);
    },
  },
  {
    name: 'SERIESSUM',
    params: [p.scalar('x'), p.scalar('n'), p.scalar('m'), p.array('coefficients')],
    summary: 'Returns the sum of a power series.',
    impl: (args, ctx) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const n = numArg(args[1]);
      if (isError(n)) return n;
      const m = numArg(args[2]);
      if (isError(m)) return m;
      const coefficients = block(args[3], ctx);
      let total = 0;
      let i = 0;
      for (const raw of coefficients.data) {
        if (isError(raw)) return raw;
        if (raw !== null && typeof raw !== 'number') return CellError.VALUE;
        // A blank coefficient is zero and still consumes its power, so a gap
        // does not shift every later term down an exponent.
        const term = (raw ?? 0) * x ** (n + i * m);
        if (!Number.isFinite(term)) return CellError.NUM;
        total = excelAdd(total, term);
        i++;
      }
      return finite(total);
    },
  },
];
