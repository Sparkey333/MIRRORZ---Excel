/**
 * Statistical functions: the descriptive aggregates, the regression family and
 * the probability distributions.
 *
 * Six decisions shape this file.
 *
 * First, the same range/argument asymmetry the math aggregates live with. A
 * value read out of a reference or an array counts only if it is already a
 * number, while a value typed straight into the formula is coerced, so
 * `AVERAGE(A1:A3)` ignores the text in A2 but `AVERAGE(1,"3")` is 2. `collect`
 * tags every value with that provenance and applies the two rules separately.
 *
 * Second, the A-suffixed variants are the same walk with one substitution:
 * text reached through a reference counts as 0 and booleans count as 1 and 0,
 * where the plain versions skip both. That is the whole difference between
 * AVERAGE and AVERAGEA, and it is expressed here as a mode flag rather than a
 * second copy of the traversal.
 *
 * Third, COUNT, COUNTA and COUNTBLANK are three genuinely different questions
 * and users lean on the differences. COUNT sees numbers only and quietly steps
 * over errors. COUNTA counts anything that is not an empty cell, error values
 * and formula-produced empty strings included, which is why its parameters are
 * error-transparent - an error must reach the implementation to be counted
 * rather than short-circuit the call. COUNTBLANK counts what is left: truly
 * empty cells plus cells holding "". It is therefore the one function here that
 * needs the reference's rectangle rather than only its populated cells, since
 * the blanks are exactly the cells the store does not store.
 *
 * Fourth, sample versus population is a spelling difference with a silent
 * numerical consequence, so the two denominators are chosen once, in
 * `varianceOf(nums, sample)`, and every dependent function names which it wants.
 * STDEV.S/VAR.S/STDEVA/VARA divide by n-1; STDEV.P/VAR.P/STDEVPA/VARPA divide
 * by n.
 *
 * Fifth, every distribution rests on four special functions - log-gamma, the
 * regularised incomplete gamma, the regularised incomplete beta and the error
 * function - implemented once, below, with their convergence criteria stated.
 * Getting these right once is what makes forty distributions right; getting one
 * wrong quietly biases a whole family. Measured against known reference values
 * they carry about fifteen significant digits (see the comment on each), which
 * is the whole of the precision Excel itself displays.
 *
 * Sixth, the inverse distributions are solved rather than approximated. A
 * bracket-and-bisect on the corresponding CDF converges to the last bit of the
 * bracket in a bounded number of steps and cannot diverge, which a Newton
 * iteration on a flat tail can; the cost is a few dozen CDF evaluations, which
 * is nothing beside the recalculation around it. The single exception is the
 * inverse normal, where Acklam's rational approximation plus one Halley
 * refinement is both faster and more accurate in the extreme tails.
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
  toBoolean,
  toNumber,
} from '../value.js';

// ---------------------------------------------------------------------------
// Argument plumbing
// ---------------------------------------------------------------------------

function scalarArg(v: Value | undefined): Scalar {
  if (v === undefined) return null;
  if (isArray(v)) return v.data[0] ?? null;
  if (isRef(v)) return null;
  return v;
}

/** The same, for ArgKind.Any, where a reference arrives undereferenced. */
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

function boolArg(v: Value | undefined, whenOmitted: boolean): boolean | CellError {
  if (v === undefined) return whenOmitted;
  return toBoolean(scalarArg(v));
}

/** Range-check a computed result the way Excel does: overflow is #NUM!. */
function finite(v: number): number | CellError {
  return checkMagnitude(v);
}

interface Item {
  value: Scalar;
  /** Typed directly into the formula rather than read from a range or array. */
  direct: boolean;
}

function* items(arg: Value | undefined, ctx: FunctionContext): Generator<Item> {
  if (arg === undefined) return;
  if (isRef(arg)) {
    // Blank cells never surface here, which is what every aggregate wants.
    for (const cell of ctx.iterate(arg)) yield { value: cell.value, direct: false };
    return;
  }
  if (isArray(arg)) {
    for (const v of arg.data) yield { value: v, direct: false };
    return;
  }
  yield { value: arg, direct: true };
}

/** Whether text and booleans in a range participate: the A-variant switch. */
const enum Mode {
  /** AVERAGE, MAX, STDEV.S: only values that are already numbers count. */
  Numbers,
  /** AVERAGEA, MAXA, STDEVA: text counts as 0 and booleans as 1 and 0. */
  Inclusive,
}

interface Collected {
  numbers: number[];
  /** Values that are not an empty cell, for COUNTA. */
  nonEmpty: number;
  /** An error value met inside a range or array. */
  error?: CellError;
  /** Direct text that does not parse as a number; COUNT ignores it, AVERAGE does not. */
  conversionError?: CellError;
}

function collect(
  args: readonly (Value | undefined)[],
  ctx: FunctionContext,
  mode: Mode = Mode.Numbers,
): Collected {
  const out: Collected = { numbers: [], nonEmpty: 0 };
  for (const arg of args) {
    for (const item of items(arg, ctx)) {
      const v = item.value;
      if (isError(v)) {
        // Counted as present: COUNTA counts error cells.
        out.nonEmpty++;
        if (!out.error) out.error = v;
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
      if (typeof v === 'number') {
        out.numbers.push(v);
        continue;
      }
      if (mode === Mode.Inclusive && v !== null) {
        // Text is 0, TRUE is 1, FALSE is 0. Numeric-looking text is still 0:
        // AVERAGEA over a cell holding "3" is 0, not 3.
        out.numbers.push(typeof v === 'boolean' ? (v ? 1 : 0) : 0);
      }
    }
  }
  return out;
}

/** The error an ordinary numeric aggregate reports, if any. */
function collectError(c: Collected): CellError | undefined {
  return c.error ?? c.conversionError;
}

/** The numbers, or the error that must be reported instead of them. */
function numbersOf(
  args: readonly (Value | undefined)[],
  ctx: FunctionContext,
  mode: Mode = Mode.Numbers,
): number[] | CellError {
  const c = collect(args, ctx, mode);
  return collectError(c) ?? c.numbers;
}

/** A rectangular block, for the functions that need positional alignment. */
function block(v: Value | undefined, ctx: FunctionContext): ArrayValue {
  if (v === undefined) return makeArray(0, 0, []);
  if (isRef(v)) return ctx.deref(v);
  if (isArray(v)) return v;
  return makeArray(1, 1, [v]);
}

/**
 * AVERAGEIF's average_range is not read as written: Excel anchors it at its
 * top-left corner and gives it the criteria range's shape, so
 * `AVERAGEIF(A1:A9,">1",B1)` averages B1:B9.
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

/** Every numeric value of a block, in reading order, or the first error in it. */
function numbersIn(b: ArrayValue): number[] | CellError {
  const out: number[] = [];
  for (const v of b.data) {
    if (isError(v)) return v;
    if (typeof v === 'number') out.push(v);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Descriptive primitives
// ---------------------------------------------------------------------------

/** Addition goes through excelAdd so cancellation snaps as it does in Excel. */
function sumOf(nums: readonly number[]): number {
  let total = 0;
  for (const n of nums) total = excelAdd(total, n);
  return total;
}

function meanOf(nums: readonly number[]): number {
  return sumOf(nums) / nums.length;
}

function averageOf(nums: readonly number[]): number | CellError {
  if (nums.length === 0) return CellError.DIV0;
  return meanOf(nums);
}

/**
 * Two-pass variance. The textbook one-pass form (E[x^2] - E[x]^2) cancels
 * catastrophically on data with a large mean and small spread - salaries around
 * 150,000 varying by 20,000 already lose four digits - so the mean is computed
 * first and the deviations summed second.
 */
function varianceOf(nums: readonly number[], sample: boolean): number | CellError {
  const n = nums.length;
  if (n < (sample ? 2 : 1)) return CellError.DIV0;
  const mean = meanOf(nums);
  let ss = 0;
  for (const x of nums) {
    const d = excelSub(x, mean);
    ss += d * d;
  }
  return ss / (sample ? n - 1 : n);
}

function stdevOf(nums: readonly number[], sample: boolean): number | CellError {
  const v = varianceOf(nums, sample);
  return isError(v) ? v : Math.sqrt(v);
}

function medianOf(nums: readonly number[]): number | CellError {
  if (nums.length === 0) return CellError.NUM;
  const s = [...nums].sort((a, b) => a - b);
  const mid = s.length >> 1;
  return s.length % 2 === 1 ? s[mid]! : excelAdd(s[mid - 1]!, s[mid]!) / 2;
}

/** All values tying for the highest repeat count, in first-appearance order. */
function modesOf(nums: readonly number[]): number[] {
  const counts = new Map<number, number>();
  for (const x of nums) counts.set(x, (counts.get(x) ?? 0) + 1);
  let best = 1;
  for (const c of counts.values()) if (c > best) best = c;
  if (best < 2) return [];
  const out: number[] = [];
  for (const [value, count] of counts) if (count === best) out.push(value);
  return out;
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
  // interval [1/(n+1), n/(n+1)] is defined.
  if (q <= 0 || q >= 1 || q < 1 / (n + 1) || q > n / (n + 1)) return CellError.NUM;
  const s = [...nums].sort((a, b) => a - b);
  const pos = q * (n + 1) - 1;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  if (lo === hi) return s[lo]!;
  return excelAdd(s[lo]!, (pos - lo) * excelSub(s[hi]!, s[lo]!));
}

function quartileOf(nums: readonly number[], q: number, inclusive: boolean): number | CellError {
  const k = Math.trunc(q);
  if (k < 0 || k > 4) return CellError.NUM;
  if (inclusive) return percentileInc(nums, k / 4);
  if (k === 0 || k === 4) return CellError.NUM;
  return percentileExc(nums, k / 4);
}

/**
 * Truncate a percent rank to `significance` places.
 *
 * Microsoft documents the default as "three digits (0.xxx)", which reads as
 * decimal places rather than significant digits, and truncates rather than
 * rounds - PERCENTRANK of a value two thirds of the way up a list is 0.666, not
 * 0.667. The two readings only differ for ranks below 0.1, which is the one
 * case worth knowing about here.
 */
function truncateRank(value: number, significance: number): number {
  const scale = 10 ** significance;
  return Math.floor(value * scale) / scale;
}

function percentRank(
  nums: readonly number[],
  x: number,
  significance: number,
  exclusive: boolean,
): number | CellError {
  const n = nums.length;
  if (n === 0) return CellError.NUM;
  if (significance < 1) return CellError.NUM;
  const s = [...nums].sort((a, b) => a - b);
  if (x < s[0]! || x > s[n - 1]!) return CellError.NA;

  // The first index whose value is at least x. Ties resolve to the first
  // occurrence, so a repeated value has one rank rather than several.
  let i = 0;
  while (i < n && s[i]! < x) i++;
  const denominator = exclusive ? n + 1 : n - 1;
  if (denominator <= 0) return CellError.NUM;

  if (i < n && s[i]! === x) {
    return truncateRank((i + (exclusive ? 1 : 0)) / denominator, significance);
  }
  // Strictly between s[i-1] and s[i]: interpolate linearly between their ranks.
  const lo = s[i - 1]!;
  const hi = s[i]!;
  const fraction = (x - lo) / (hi - lo);
  const loRank = (i - 1 + (exclusive ? 1 : 0)) / denominator;
  const hiRank = (i + (exclusive ? 1 : 0)) / denominator;
  return truncateRank(loRank + fraction * (hiRank - loRank), significance);
}

/**
 * RANK.EQ and RANK.AVG. `order` of 0 ranks descending, which is Excel's
 * default and the opposite of what the argument name suggests.
 */
function rankOf(
  x: number,
  nums: readonly number[],
  descending: boolean,
  average: boolean,
): number | CellError {
  let below = 0;
  let ties = 0;
  for (const v of nums) {
    if (v === x) ties++;
    else if (descending ? v > x : v < x) below++;
  }
  if (ties === 0) return CellError.NA;
  const first = below + 1;
  // The average form spreads the tied block over the positions it occupies.
  return average ? first + (ties - 1) / 2 : first;
}

/** Sum of squared deviations from the mean, the core of the variance family. */
function devsqOf(nums: readonly number[]): number {
  const mean = meanOf(nums);
  let ss = 0;
  for (const x of nums) {
    const d = excelSub(x, mean);
    ss += d * d;
  }
  return ss;
}

/** The k-th central moment, used by SKEW, SKEW.P and KURT. */
function centralMoment(nums: readonly number[], k: number): number {
  const mean = meanOf(nums);
  let acc = 0;
  for (const x of nums) acc += excelSub(x, mean) ** k;
  return acc;
}

// ---------------------------------------------------------------------------
// Special functions
//
// Everything below this line is shared by the distributions. It is written once
// and carefully, because an error here is invisible at the call site and shows
// up only as a distribution that is slightly wrong in one tail.
// ---------------------------------------------------------------------------

const LOG_SQRT_2PI = 0.9189385332046727;
const SQRT_2PI = 2.5066282746310002;
const SQRT_2 = 1.4142135623730951;

/**
 * Lanczos coefficients, g = 7, n = 9.
 *
 * This is the classical Numerical-Recipes parameterisation. Its relative error
 * on gamma itself is below 1e-15 across the right half plane, so `lnGamma`
 * carries about fifteen significant digits - checked against LN(GAMMA(x)) at
 * x = 0.5, 1.5, 10, 100 and 1000.
 */
const LANCZOS = [
  0.9999999999998099, 676.5203681218851, -1259.1392167224028, 771.3234287776531,
  -176.6150291621406, 12.507343278686905, -0.13857109526572012, 9.984369578019572e-6,
  1.5056327351493116e-7,
] as const;

/** log(gamma(x)) for real x, using reflection below 0.5. */
function lnGamma(x: number): number {
  if (x < 0.5) {
    // Euler's reflection, which moves the argument into the accurate half.
    return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - lnGamma(1 - x);
  }
  const z = x - 1;
  let series = LANCZOS[0];
  for (let i = 1; i < LANCZOS.length; i++) series += LANCZOS[i]! / (z + i);
  const t = z + 7.5;
  return LOG_SQRT_2PI + (z + 0.5) * Math.log(t) - t + Math.log(series);
}

/** Exact factorials, so GAMMA and the discrete distributions stay integral. */
const FACTORIALS: number[] = (() => {
  const out = [1];
  for (let i = 1; i <= 170; i++) out.push(out[i - 1]! * i);
  return out;
})();

/** gamma(x). Undefined at zero and the negative integers, as in Excel. */
function gammaFn(x: number): number | CellError {
  if (Number.isInteger(x)) {
    if (x <= 0) return CellError.NUM;
    if (x <= 171) return FACTORIALS[x - 1]!;
    return CellError.NUM;
  }
  if (x > 0) {
    const r = Math.exp(lnGamma(x));
    return Number.isFinite(r) ? r : CellError.NUM;
  }
  // Reflection carries the sign, which the log form has thrown away.
  const sin = Math.sin(Math.PI * x);
  const r = Math.PI / (sin * Math.exp(lnGamma(1 - x)));
  return Number.isFinite(r) ? r : CellError.NUM;
}

/** ln(n!) for the discrete distributions, exact while the factorial is. */
function lnFactorial(n: number): number {
  return n <= 170 ? Math.log(FACTORIALS[n]!) : lnGamma(n + 1);
}

/** ln(C(n,k)), the log binomial coefficient. */
function lnChoose(n: number, k: number): number {
  return lnFactorial(n) - lnFactorial(k) - lnFactorial(n - k);
}

/** C(n,k) as a number, exact for the sizes a spreadsheet actually uses. */
function chooseFn(n: number, k: number): number {
  if (k < 0 || k > n) return 0;
  if (k === 0 || k === n) return 1;
  const m = Math.min(k, n - k);
  let acc = 1;
  // The running quotient stays integral at every step, so no rounding creeps in
  // until the true result exceeds 2^53.
  for (let i = 1; i <= m; i++) acc = (acc * (n - m + i)) / i;
  // The quotient is integral in exact arithmetic; snapping it back removes the
  // last-bit dust while the true value is still representable.
  return acc < 9007199254740992 ? Math.round(acc) : acc;
}

/**
 * Regularised lower incomplete gamma P(a,x) = gamma(a,x)/gamma(a).
 *
 * Series below the turning point x < a+1, Lentz's continued fraction above it,
 * which is the split that keeps both branches convergent. Both stop when a term
 * changes the running value by less than 1e-16 relative, with a hard cap of
 * 1000 iterations that is never reached for arguments a spreadsheet produces;
 * accuracy is about fifteen significant digits.
 */
function gammaP(a: number, x: number): number {
  if (x <= 0) return 0;
  if (a <= 0) return 1;
  if (x < a + 1) {
    let term = 1 / a;
    let sum = term;
    for (let n = 1; n < 1000; n++) {
      term *= x / (a + n);
      sum += term;
      if (Math.abs(term) < Math.abs(sum) * 1e-16) break;
    }
    const r = sum * Math.exp(-x + a * Math.log(x) - lnGamma(a));
    return r > 1 ? 1 : r;
  }
  return 1 - gammaQContinued(a, x);
}

function gammaQ(a: number, x: number): number {
  if (x <= 0) return 1;
  if (a <= 0) return 0;
  if (x < a + 1) return 1 - gammaP(a, x);
  return gammaQContinued(a, x);
}

/** The Q branch: modified Lentz on the continued fraction, valid for x >= a+1. */
function gammaQContinued(a: number, x: number): number {
  const tiny = 1e-300;
  let b = x + 1 - a;
  let c = 1 / tiny;
  let d = 1 / b;
  let h = d;
  for (let i = 1; i < 1000; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < tiny) d = tiny;
    c = b + an / c;
    if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) < 1e-16) break;
  }
  return Math.exp(-x + a * Math.log(x) - lnGamma(a)) * h;
}

/**
 * Regularised incomplete beta I_x(a,b).
 *
 * The continued fraction converges quickly only on the side of the distribution
 * where x is below the mode, so the symmetry I_x(a,b) = 1 - I_{1-x}(b,a) moves
 * the evaluation there first. Lentz's algorithm again, tolerance 1e-16 relative
 * on the multiplier, capped at 500 iterations; about fifteen significant digits
 * for the argument ranges the distributions use.
 */
function betaI(a: number, b: number, x: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const front = Math.exp(
    lnGamma(a + b) - lnGamma(a) - lnGamma(b) + a * Math.log(x) + b * Math.log(1 - x),
  );
  if (x < (a + 1) / (a + b + 2)) return (front * betaContinued(a, b, x)) / a;
  // The leading factor is symmetric under (a,b,x) -> (b,a,1-x), so the same
  // `front` serves the mirrored branch.
  return 1 - (front * betaContinued(b, a, 1 - x)) / b;
}

function betaContinued(a: number, b: number, x: number): number {
  const tiny = 1e-300;
  const qab = a + b;
  const qap = a + 1;
  const qam = a - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < tiny) d = tiny;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= 500; m++) {
    const m2 = 2 * m;
    // The fraction alternates between two forms of the numerator, one for the
    // even step and one for the odd.
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < tiny) d = tiny;
    c = 1 + aa / c;
    if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < tiny) d = tiny;
    c = 1 + aa / c;
    if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) < 1e-16) break;
  }
  return h;
}

/**
 * erf and erfc through the incomplete gamma, which is exact rather than an
 * independent approximation: erfc(x) = Q(1/2, x^2) for x >= 0. Routing the
 * negative side through 2 - erfc(-x) loses no accuracy that matters, since the
 * result is then of order 1 anyway.
 */
function erfc(x: number): number {
  if (x < 0) return 2 - erfc(-x);
  return gammaQ(0.5, x * x);
}

function erf(x: number): number {
  if (x < 0) return -erf(-x);
  return x < 0.5 ? gammaP(0.5, x * x) : 1 - gammaQ(0.5, x * x);
}

/** Standard normal density. */
function normPdf(z: number): number {
  return Math.exp(-0.5 * z * z) / SQRT_2PI;
}

/**
 * Standard normal CDF, written with erfc so that the far left tail keeps its
 * relative accuracy: 1 - Phi(z) computed by subtraction would be zero long
 * before Phi(-z) underflows.
 */
function normCdf(z: number): number {
  return 0.5 * erfc(-z / SQRT_2);
}

/**
 * Inverse standard normal.
 *
 * Acklam's rational approximation (relative error below 1.15e-9) refined by one
 * Halley step against the erfc-based CDF, which takes it to about 1e-15. The
 * refinement is skipped where exp(z^2/2) would overflow, i.e. below roughly
 * p = 1e-160, where the approximation's own accuracy is all that is on offer.
 */
const ACKLAM_A = [
  -3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2,
  -3.066479806614716e1, 2.506628277459239,
] as const;
const ACKLAM_B = [
  -5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1,
  -1.328068155288572e1,
] as const;
const ACKLAM_C = [
  -7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734,
  4.374664141464968, 2.938163982698783,
] as const;
const ACKLAM_D = [
  7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416,
] as const;

function normInv(p: number): number {
  const low = 0.02425;
  let z: number;
  if (p < low) {
    const q = Math.sqrt(-2 * Math.log(p));
    z =
      (((((ACKLAM_C[0] * q + ACKLAM_C[1]) * q + ACKLAM_C[2]) * q + ACKLAM_C[3]) * q +
        ACKLAM_C[4]) *
        q +
        ACKLAM_C[5]) /
      ((((ACKLAM_D[0] * q + ACKLAM_D[1]) * q + ACKLAM_D[2]) * q + ACKLAM_D[3]) * q + 1);
  } else if (p <= 1 - low) {
    const q = p - 0.5;
    const r = q * q;
    z =
      ((((((ACKLAM_A[0] * r + ACKLAM_A[1]) * r + ACKLAM_A[2]) * r + ACKLAM_A[3]) * r +
        ACKLAM_A[4]) *
        r +
        ACKLAM_A[5]) *
        q) /
      (((((ACKLAM_B[0] * r + ACKLAM_B[1]) * r + ACKLAM_B[2]) * r + ACKLAM_B[3]) * r +
        ACKLAM_B[4]) *
        r +
        1);
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    z =
      -(
        ((((ACKLAM_C[0] * q + ACKLAM_C[1]) * q + ACKLAM_C[2]) * q + ACKLAM_C[3]) * q +
          ACKLAM_C[4]) *
          q +
        ACKLAM_C[5]
      ) /
      ((((ACKLAM_D[0] * q + ACKLAM_D[1]) * q + ACKLAM_D[2]) * q + ACKLAM_D[3]) * q + 1);
  }
  if (z * z < 1400) {
    const e = normCdf(z) - p;
    const u = e * SQRT_2PI * Math.exp((z * z) / 2);
    z -= u / (1 + (z * u) / 2);
  }
  return z;
}

/**
 * Invert a non-decreasing CDF on [lo, hi] by bracketing then bisection.
 *
 * `hi` may be Infinity, in which case the bracket doubles outwards until the
 * CDF passes the target. Sixty halvings exhaust a double's mantissa, so the
 * loop terminates on the interval width long before its 200-step cap; the
 * result is as accurate as the CDF handed in.
 */
function invertCdf(
  cdf: (x: number) => number,
  target: number,
  lo: number,
  hi: number,
): number | CellError {
  let a = lo;
  let b = hi;
  if (!Number.isFinite(b)) {
    b = Math.max(1, Math.abs(a) * 2);
    let guard = 0;
    while (cdf(b) < target) {
      b *= 2;
      if (++guard > 2000 || b > 1e300) return CellError.NUM;
    }
  }
  if (!Number.isFinite(a)) {
    a = -Math.max(1, Math.abs(b) * 2);
    let guard = 0;
    while (cdf(a) > target) {
      a *= 2;
      if (++guard > 2000 || a < -1e300) return CellError.NUM;
    }
  }
  for (let i = 0; i < 200; i++) {
    const mid = a + (b - a) / 2;
    if (mid === a || mid === b) break;
    if (cdf(mid) < target) a = mid;
    else b = mid;
    if (b - a <= Math.abs(mid) * 1e-16) break;
  }
  return a + (b - a) / 2;
}

// ---------------------------------------------------------------------------
// Paired data and regression
// ---------------------------------------------------------------------------

interface Pairs {
  xs: number[];
  ys: number[];
}

/**
 * The pairs CORREL, SLOPE, COVARIANCE and friends work on.
 *
 * Excel compares counts rather than shapes, so a row and a column of the same
 * length pair up; a pair is dropped entirely unless both halves are numbers,
 * which is what "text, logical values or empty cells are ignored" means when
 * two parallel ranges are involved.
 */
function pairsOf(xBlock: ArrayValue, yBlock: ArrayValue): Pairs | CellError {
  const nx = xBlock.rows * xBlock.cols;
  const ny = yBlock.rows * yBlock.cols;
  if (nx !== ny) return CellError.NA;
  const xs: number[] = [];
  const ys: number[] = [];
  for (let i = 0; i < nx; i++) {
    const a = xBlock.data[i] ?? null;
    const b = yBlock.data[i] ?? null;
    if (isError(a)) return a;
    if (isError(b)) return b;
    if (typeof a !== 'number' || typeof b !== 'number') continue;
    xs.push(a);
    ys.push(b);
  }
  return { xs, ys };
}

/** Sums of squares and cross-products about the means. */
interface Moments {
  n: number;
  meanX: number;
  meanY: number;
  sxx: number;
  syy: number;
  sxy: number;
}

function momentsOf(pairs: Pairs): Moments {
  const n = pairs.xs.length;
  const meanX = n === 0 ? 0 : meanOf(pairs.xs);
  const meanY = n === 0 ? 0 : meanOf(pairs.ys);
  let sxx = 0;
  let syy = 0;
  let sxy = 0;
  for (let i = 0; i < n; i++) {
    const dx = excelSub(pairs.xs[i]!, meanX);
    const dy = excelSub(pairs.ys[i]!, meanY);
    sxx += dx * dx;
    syy += dy * dy;
    sxy += dx * dy;
  }
  return { n, meanX, meanY, sxx, syy, sxy };
}

interface Regression {
  /** [intercept, m1 .. mk]; the intercept is 0 when it was forced off. */
  coef: number[];
  se: number[];
  r2: number;
  sey: number;
  f: number;
  df: number;
  ssReg: number;
  ssResid: number;
  k: number;
  useConst: boolean;
}

/**
 * Ordinary least squares through the normal equations.
 *
 * Forming X'X squares the condition number, which a QR factorisation would
 * avoid; for the one- and two-variable fits a worksheet actually contains the
 * difference is far below Excel's fifteen displayed digits, and the inverse of
 * X'X is needed anyway for the coefficient standard errors.
 */
function leastSquares(xs: number[][], y: number[], useConst: boolean): Regression | undefined {
  const n = y.length;
  const k = xs[0]?.length ?? 0;
  if (n === 0 || k === 0) return undefined;
  const c = useConst ? k + 1 : k;

  const design = (i: number, j: number): number => (useConst ? (j === 0 ? 1 : xs[i]![j - 1]!) : xs[i]![j]!);

  const a: number[][] = Array.from({ length: c }, () => new Array<number>(c).fill(0));
  const rhs = new Array<number>(c).fill(0);
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < c; j++) {
      const dj = design(i, j);
      rhs[j] = rhs[j]! + dj * y[i]!;
      for (let l = 0; l < c; l++) a[j]![l] = a[j]![l]! + dj * design(i, l);
    }
  }

  const inv = invertMatrix(a);
  if (!inv) return undefined;

  const beta = new Array<number>(c).fill(0);
  for (let j = 0; j < c; j++) {
    let acc = 0;
    for (let l = 0; l < c; l++) acc += inv[j]![l]! * rhs[l]!;
    beta[j] = acc;
  }

  const meanY = meanOf(y);
  let ssResid = 0;
  let ssReg = 0;
  for (let i = 0; i < n; i++) {
    let fit = 0;
    for (let j = 0; j < c; j++) fit += beta[j]! * design(i, j);
    const resid = excelSub(y[i]!, fit);
    ssResid += resid * resid;
    // Without an intercept the regression sum of squares is measured from zero,
    // not from the mean, which is why R^2 can look implausibly high there.
    const centred = useConst ? excelSub(fit, meanY) : fit;
    ssReg += centred * centred;
  }

  const df = n - c;
  const mse = df > 0 ? ssResid / df : 0;
  const sey = df > 0 ? Math.sqrt(mse) : 0;
  const se = new Array<number>(c).fill(0);
  for (let j = 0; j < c; j++) se[j] = Math.sqrt(Math.abs(inv[j]![j]!) * mse);

  const total = ssReg + ssResid;
  return {
    coef: useConst ? beta : [0, ...beta],
    se: useConst ? se : [0, ...se],
    r2: total === 0 ? 1 : ssReg / total,
    sey,
    f: df > 0 && ssResid > 0 ? ssReg / k / mse : 0,
    df,
    ssReg,
    ssResid,
    k,
    useConst,
  };
}

/** Gauss-Jordan inverse with partial pivoting; undefined when singular. */
function invertMatrix(m: readonly number[][]): number[][] | undefined {
  const n = m.length;
  const a = m.map((row) => [...row]);
  const inv: number[][] = Array.from({ length: n }, (_, i) =>
    Array.from({ length: n }, (_, j) => (i === j ? 1 : 0)),
  );
  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let r = col + 1; r < n; r++) {
      if (Math.abs(a[r]![col]!) > Math.abs(a[pivot]![col]!)) pivot = r;
    }
    if (Math.abs(a[pivot]![col]!) < 1e-300) return undefined;
    [a[col], a[pivot]] = [a[pivot]!, a[col]!];
    [inv[col], inv[pivot]] = [inv[pivot]!, inv[col]!];
    const d = a[col]![col]!;
    for (let j = 0; j < n; j++) {
      a[col]![j] = a[col]![j]! / d;
      inv[col]![j] = inv[col]![j]! / d;
    }
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const factor = a[r]![col]!;
      if (factor === 0) continue;
      for (let j = 0; j < n; j++) {
        a[r]![j] = a[r]![j]! - factor * a[col]![j]!;
        inv[r]![j] = inv[r]![j]! - factor * inv[col]![j]!;
      }
    }
  }
  return inv;
}

interface RegressionInput {
  y: number[];
  /** One row per observation, one column per independent variable. */
  xs: number[][];
  k: number;
  /** known_y arrived as a column, which decides the shape TREND returns. */
  yIsColumn: boolean;
}

/** Every cell of a block as a number, or the error Excel reports instead. */
function strictNumbers(b: ArrayValue): number[] | CellError {
  const out: number[] = [];
  for (const v of b.data) {
    if (isError(v)) return v;
    if (typeof v !== 'number') return CellError.VALUE;
    out.push(v);
  }
  return out;
}

/**
 * Work out which way round known_x is.
 *
 * With one independent variable the two arrays simply have to hold the same
 * number of values. With several, the variables run down the columns when
 * known_y is a column and along the rows when it is a row; anything else is
 * #REF!, which is the error Excel reports for a shape mismatch here rather than
 * the #VALUE! the rest of the family uses.
 */
function regressionInput(
  yBlock: ArrayValue,
  xBlock: ArrayValue | undefined,
): RegressionInput | CellError {
  const y = strictNumbers(yBlock);
  if (isError(y)) return y;
  const n = y.length;
  if (n === 0) return CellError.REF;
  const yIsColumn = yBlock.cols === 1;

  if (xBlock === undefined || xBlock.rows * xBlock.cols === 0) {
    return { y, xs: y.map((_, i) => [i + 1]), k: 1, yIsColumn };
  }
  const xValues = strictNumbers(xBlock);
  if (isError(xValues)) return xValues;

  if (xValues.length === n) {
    return { y, xs: xValues.map((v) => [v]), k: 1, yIsColumn };
  }
  if (yIsColumn && xBlock.rows === n) {
    const k = xBlock.cols;
    const xs = Array.from({ length: n }, (_, i) =>
      Array.from({ length: k }, (_, j) => xBlock.data[i * xBlock.cols + j] as number),
    );
    return { y, xs, k, yIsColumn };
  }
  if (!yIsColumn && xBlock.cols === n) {
    const k = xBlock.rows;
    const xs = Array.from({ length: n }, (_, i) =>
      Array.from({ length: k }, (_, j) => xBlock.data[j * xBlock.cols + i] as number),
    );
    return { y, xs, k, yIsColumn };
  }
  return CellError.REF;
}

/** The new observations TREND and GROWTH predict at, in the same orientation. */
function newObservations(
  b: ArrayValue | undefined,
  input: RegressionInput,
): { rows: number[][]; shape: { rows: number; cols: number } } | CellError {
  if (b === undefined || b.rows * b.cols === 0) {
    return {
      rows: input.xs,
      shape: input.yIsColumn
        ? { rows: input.xs.length, cols: 1 }
        : { rows: 1, cols: input.xs.length },
    };
  }
  const values = strictNumbers(b);
  if (isError(values)) return values;
  if (input.k === 1) {
    // One variable: every cell is an observation and the answer keeps the shape
    // of new_x, so a row of new x values produces a row of predictions.
    return { rows: values.map((v) => [v]), shape: { rows: b.rows, cols: b.cols } };
  }
  if (b.cols === input.k) {
    const m = b.rows;
    return {
      rows: Array.from({ length: m }, (_, i) =>
        Array.from({ length: input.k }, (_, j) => b.data[i * b.cols + j] as number),
      ),
      shape: { rows: m, cols: 1 },
    };
  }
  if (b.rows === input.k) {
    const m = b.cols;
    return {
      rows: Array.from({ length: m }, (_, i) =>
        Array.from({ length: input.k }, (_, j) => b.data[j * b.cols + i] as number),
      ),
      shape: { rows: 1, cols: m },
    };
  }
  return CellError.REF;
}

function predict(reg: Regression, row: readonly number[]): number {
  let acc = reg.coef[0]!;
  for (let j = 0; j < row.length; j++) acc = excelAdd(acc, reg.coef[j + 1]! * row[j]!);
  return acc;
}

// ---------------------------------------------------------------------------
// Distribution cores
// ---------------------------------------------------------------------------

function chiSqCdf(x: number, df: number): number {
  return x <= 0 ? 0 : gammaP(df / 2, x / 2);
}

function chiSqPdf(x: number, df: number): number {
  if (x < 0) return 0;
  if (x === 0) return df === 2 ? 0.5 : df < 2 ? Number.POSITIVE_INFINITY : 0;
  const k = df / 2;
  return Math.exp((k - 1) * Math.log(x) - x / 2 - k * Math.LN2 - lnGamma(k));
}

function studentTCdf(x: number, df: number): number {
  const z = df / (df + x * x);
  const half = betaI(df / 2, 0.5, z) / 2;
  return x <= 0 ? half : 1 - half;
}

function studentTPdf(x: number, df: number): number {
  return Math.exp(
    lnGamma((df + 1) / 2) -
      lnGamma(df / 2) -
      0.5 * Math.log(df * Math.PI) -
      ((df + 1) / 2) * Math.log(1 + (x * x) / df),
  );
}

function fCdf(x: number, df1: number, df2: number): number {
  if (x <= 0) return 0;
  return betaI(df1 / 2, df2 / 2, (df1 * x) / (df1 * x + df2));
}

function fPdf(x: number, df1: number, df2: number): number {
  if (x < 0) return 0;
  if (x === 0) return df1 < 2 ? Number.POSITIVE_INFINITY : df1 === 2 ? 1 : 0;
  const a = df1 / 2;
  const b = df2 / 2;
  return Math.exp(
    lnGamma(a + b) -
      lnGamma(a) -
      lnGamma(b) +
      a * Math.log(df1 / df2) +
      (a - 1) * Math.log(x) -
      (a + b) * Math.log(1 + (df1 * x) / df2),
  );
}

function binomPmf(k: number, n: number, prob: number): number {
  if (k < 0 || k > n) return 0;
  if (prob === 0) return k === 0 ? 1 : 0;
  if (prob === 1) return k === n ? 1 : 0;
  return Math.exp(lnChoose(n, k) + k * Math.log(prob) + (n - k) * Math.log1p(-prob));
}

/**
 * The binomial CDF as a regularised incomplete beta rather than a running sum,
 * so a thousand-trial cumulative costs the same as a ten-trial one and does not
 * accumulate a thousand rounding errors.
 */
function binomCdf(k: number, n: number, prob: number): number {
  if (k < 0) return 0;
  if (k >= n) return 1;
  return betaI(n - k, k + 1, 1 - prob);
}

function poissonPmf(x: number, mean: number): number {
  if (mean === 0) return x === 0 ? 1 : 0;
  return Math.exp(-mean + x * Math.log(mean) - lnFactorial(x));
}

function poissonCdf(x: number, mean: number): number {
  if (x < 0) return 0;
  if (mean === 0) return 1;
  return gammaQ(x + 1, mean);
}

function hypgeomPmf(k: number, n: number, successes: number, population: number): number {
  if (k < 0 || k > n || k > successes || n - k > population - successes) return 0;
  return Math.exp(
    lnChoose(successes, k) + lnChoose(population - successes, n - k) - lnChoose(population, n),
  );
}

// ---------------------------------------------------------------------------
// Spec builders
// ---------------------------------------------------------------------------

/** An aggregate over a repeating argument list, in one of the two coercion modes. */
function aggregate(
  name: string,
  mode: Mode,
  summary: string,
  reduce: (nums: number[]) => number | CellError,
): FunctionSpec {
  return {
    name,
    params: [p.any('number1'), p.rest('number2', ArgKind.Any)],
    summary,
    impl: (args, ctx) => {
      const nums = numbersOf(args, ctx, mode);
      return isError(nums) ? nums : reduce(nums);
    },
  };
}

function maxOf(nums: readonly number[]): number {
  // A loop rather than Math.max(...nums): spreading a whole column would blow
  // the argument limit long before the sheet ran out of rows.
  let best = Number.NEGATIVE_INFINITY;
  for (const n of nums) if (n > best) best = n;
  return best === Number.NEGATIVE_INFINITY ? 0 : best;
}

function minOf(nums: readonly number[]): number {
  let best = Number.POSITIVE_INFINITY;
  for (const n of nums) if (n < best) best = n;
  return best === Number.POSITIVE_INFINITY ? 0 : best;
}

/** Criteria ranges paired with their parsed criteria, for the *IFS family. */
interface Test {
  block: ArrayValue;
  criterion: Criterion;
}

function conditions(
  args: readonly (Value | undefined)[],
  from: number,
  shape: { rows: number; cols: number } | undefined,
  ctx: FunctionContext,
): Test[] | CellError {
  const supplied = args.slice(from).filter((a) => a !== undefined);
  if (supplied.length === 0 || supplied.length % 2 !== 0) return CellError.VALUE;
  const tests: Test[] = [];
  for (let i = 0; i < supplied.length; i += 2) {
    const b = block(supplied[i], ctx);
    if (shape && (b.rows !== shape.rows || b.cols !== shape.cols)) return CellError.VALUE;
    tests.push({ block: b, criterion: parseCriterion(anyScalarArg(supplied[i + 1], ctx)) });
  }
  return tests;
}

/** Indices of the cells that satisfy every condition. */
function matchingIndices(cells: number, tests: readonly Test[]): number[] {
  const out: number[] = [];
  for (let i = 0; i < cells; i++) {
    let hit = true;
    for (const t of tests) {
      // An error inside a criteria range simply fails to match rather than
      // propagating; only the aggregated range can raise one.
      if (!matchesCriterion(t.block.data[i] ?? null, t.criterion)) {
        hit = false;
        break;
      }
    }
    if (hit) out.push(i);
  }
  return out;
}

/** The numbers at the given positions of a block, propagating any error found. */
function valuesAt(b: ArrayValue, indices: readonly number[]): number[] | CellError {
  const out: number[] = [];
  for (const i of indices) {
    const v = b.data[i] ?? null;
    if (isError(v)) return v;
    if (typeof v === 'number') out.push(v);
  }
  return out;
}

// ---------------------------------------------------------------------------
// The function table
// ---------------------------------------------------------------------------

const COUNTING: FunctionSpec[] = [
  {
    name: 'COUNT',
    params: [p.any('value1'), p.rest('value2', ArgKind.Any)],
    summary: 'Counts how many numbers are in the list of arguments.',
    impl: (args, ctx) => collect(args, ctx).numbers.length,
  },
  {
    name: 'COUNTA',
    params: [
      { name: 'value1', kind: ArgKind.Any, errorTransparent: true },
      { name: 'value2', kind: ArgKind.Any, repeating: true, optional: true, errorTransparent: true },
    ],
    summary: 'Counts how many values are in the list of arguments.',
    impl: (args, ctx) => collect(args, ctx).nonEmpty,
  },
  {
    name: 'COUNTBLANK',
    params: [{ name: 'range', kind: ArgKind.Any, errorTransparent: true }],
    summary: 'Counts the number of blank cells in a range.',
    impl: (args, ctx) => {
      const arg = args[0];
      if (isRef(arg)) {
        // The blanks are exactly the cells the store does not hold, so they are
        // counted by subtraction from the rectangle's area.
        const area =
          (arg.endRow - arg.startRow + 1) * (arg.endCol - arg.startCol + 1);
        let occupied = 0;
        for (const cell of ctx.iterate(arg)) {
          // A formula that returned "" reads as blank here, which is the whole
          // reason COUNTBLANK and COUNTA disagree on such a cell.
          if (cell.value !== '') occupied++;
        }
        return area - occupied;
      }
      const b = block(arg, ctx);
      let blanks = 0;
      for (const v of b.data) if (v === null || v === '') blanks++;
      return blanks;
    },
  },
  {
    name: 'COUNTIF',
    params: [p.any('range'), p.scalar('criteria')],
    summary: 'Counts the cells in a range that meet a condition.',
    impl: (args, ctx) => {
      const range = block(args[0], ctx);
      const criterion = parseCriterion(anyScalarArg(args[1], ctx));
      return matchingIndices(range.rows * range.cols, [{ block: range, criterion }]).length;
    },
  },
  {
    name: 'COUNTIFS',
    params: [p.any('criteria_range1'), p.scalar('criteria1'), p.rest('more', ArgKind.Any)],
    summary: 'Counts the cells that meet several conditions.',
    impl: (args, ctx) => {
      const first = block(args[0], ctx);
      const tests = conditions(args, 0, { rows: first.rows, cols: first.cols }, ctx);
      if (isError(tests)) return tests;
      return matchingIndices(first.rows * first.cols, tests).length;
    },
  },
];

const AVERAGES: FunctionSpec[] = [
  aggregate('AVERAGE', Mode.Numbers, 'Returns the average of its arguments.', averageOf),
  aggregate(
    'AVERAGEA',
    Mode.Inclusive,
    'Returns the average of its arguments, counting text as zero.',
    averageOf,
  ),
  {
    name: 'AVERAGEIF',
    params: [p.any('range'), p.scalar('criteria'), p.any('average_range', true)],
    summary: 'Returns the average of the cells that meet a condition.',
    impl: (args, ctx) => {
      const range = block(args[0], ctx);
      const criterion = parseCriterion(anyScalarArg(args[1], ctx));
      const target =
        args[2] === undefined ? range : alignedBlock(args[2], range.rows, range.cols, ctx);
      const hits = matchingIndices(range.rows * range.cols, [{ block: range, criterion }]);
      const nums = valuesAt(target, hits);
      return isError(nums) ? nums : averageOf(nums);
    },
  },
  {
    name: 'AVERAGEIFS',
    params: [
      p.any('average_range'),
      p.any('criteria_range1'),
      p.scalar('criteria1'),
      p.rest('more', ArgKind.Any),
    ],
    summary: 'Returns the average of the cells that meet several conditions.',
    impl: (args, ctx) => {
      const target = block(args[0], ctx);
      const tests = conditions(args, 1, { rows: target.rows, cols: target.cols }, ctx);
      if (isError(tests)) return tests;
      const nums = valuesAt(target, matchingIndices(target.rows * target.cols, tests));
      return isError(nums) ? nums : averageOf(nums);
    },
  },
];

const EXTREMES: FunctionSpec[] = [
  aggregate('MAX', Mode.Numbers, 'Returns the largest value in a set of values.', maxOf),
  aggregate(
    'MAXA',
    Mode.Inclusive,
    'Returns the largest value, counting text as zero.',
    maxOf,
  ),
  aggregate('MIN', Mode.Numbers, 'Returns the smallest value in a set of values.', minOf),
  aggregate(
    'MINA',
    Mode.Inclusive,
    'Returns the smallest value, counting text as zero.',
    minOf,
  ),
  {
    name: 'MAXIFS',
    params: [
      p.any('max_range'),
      p.any('criteria_range1'),
      p.scalar('criteria1'),
      p.rest('more', ArgKind.Any),
    ],
    summary: 'Returns the largest value among cells meeting several conditions.',
    impl: (args, ctx) => {
      const target = block(args[0], ctx);
      const tests = conditions(args, 1, { rows: target.rows, cols: target.cols }, ctx);
      if (isError(tests)) return tests;
      const nums = valuesAt(target, matchingIndices(target.rows * target.cols, tests));
      return isError(nums) ? nums : maxOf(nums);
    },
  },
  {
    name: 'MINIFS',
    params: [
      p.any('min_range'),
      p.any('criteria_range1'),
      p.scalar('criteria1'),
      p.rest('more', ArgKind.Any),
    ],
    summary: 'Returns the smallest value among cells meeting several conditions.',
    impl: (args, ctx) => {
      const target = block(args[0], ctx);
      const tests = conditions(args, 1, { rows: target.rows, cols: target.cols }, ctx);
      if (isError(tests)) return tests;
      const nums = valuesAt(target, matchingIndices(target.rows * target.cols, tests));
      return isError(nums) ? nums : minOf(nums);
    },
  },
];

const ORDER_STATISTICS: FunctionSpec[] = [
  aggregate('MEDIAN', Mode.Numbers, 'Returns the median of the given numbers.', medianOf),
  aggregate('MODE.SNGL', Mode.Numbers, 'Returns the most common value in a data set.', (nums) => {
    const modes = modesOf(nums);
    // No value repeats, so there is no mode to report.
    return modes.length === 0 ? CellError.NA : modes[0]!;
  }),
  {
    name: 'MODE.MULT',
    params: [p.any('number1'), p.rest('number2', ArgKind.Any)],
    summary: 'Returns a vertical array of the most frequently occurring values.',
    impl: (args, ctx) => {
      const nums = numbersOf(args, ctx);
      if (isError(nums)) return nums;
      const modes = modesOf(nums);
      if (modes.length === 0) return CellError.NA;
      return makeArray(modes.length, 1, modes);
    },
  },
  {
    name: 'LARGE',
    params: [p.any('array'), p.scalar('k')],
    summary: 'Returns the k-th largest value in a data set.',
    impl: (args, ctx) => {
      const nums = numbersOf([args[0]], ctx);
      if (isError(nums)) return nums;
      const k = numArg(args[1]);
      return isError(k) ? k : nthOf(nums, k, true);
    },
  },
  {
    name: 'SMALL',
    params: [p.any('array'), p.scalar('k')],
    summary: 'Returns the k-th smallest value in a data set.',
    impl: (args, ctx) => {
      const nums = numbersOf([args[0]], ctx);
      if (isError(nums)) return nums;
      const k = numArg(args[1]);
      return isError(k) ? k : nthOf(nums, k, false);
    },
  },
  rankSpec('RANK.EQ', false),
  rankSpec('RANK.AVG', true),
  percentileSpec('PERCENTILE.INC', true),
  percentileSpec('PERCENTILE.EXC', false),
  quartileSpec('QUARTILE.INC', true),
  quartileSpec('QUARTILE.EXC', false),
  percentRankSpec('PERCENTRANK.INC', false),
  percentRankSpec('PERCENTRANK.EXC', true),
];

function rankSpec(name: string, average: boolean): FunctionSpec {
  return {
    name,
    params: [p.scalar('number'), p.any('ref'), p.scalar('order', true)],
    summary:
      average
        ? 'Returns the rank of a number in a list, averaging ties.'
        : 'Returns the rank of a number in a list, giving ties the top rank.',
    impl: (args, ctx) => {
      const x = numArg(args[0]);
      if (isError(x)) return x;
      const nums = numbersOf([args[1]], ctx);
      if (isError(nums)) return nums;
      const order = numArg(args[2], 0);
      if (isError(order)) return order;
      return rankOf(x, nums, order === 0, average);
    },
  };
}

function percentileSpec(name: string, inclusive: boolean): FunctionSpec {
  return {
    name,
    params: [p.any('array'), p.scalar('k')],
    summary: 'Returns the k-th percentile of values in a range.',
    impl: (args, ctx) => {
      const nums = numbersOf([args[0]], ctx);
      if (isError(nums)) return nums;
      const k = numArg(args[1]);
      if (isError(k)) return k;
      return inclusive ? percentileInc(nums, k) : percentileExc(nums, k);
    },
  };
}

function quartileSpec(name: string, inclusive: boolean): FunctionSpec {
  return {
    name,
    params: [p.any('array'), p.scalar('quart')],
    summary: 'Returns the quartile of a data set.',
    impl: (args, ctx) => {
      const nums = numbersOf([args[0]], ctx);
      if (isError(nums)) return nums;
      const q = numArg(args[1]);
      if (isError(q)) return q;
      return quartileOf(nums, q, inclusive);
    },
  };
}

function percentRankSpec(name: string, exclusive: boolean): FunctionSpec {
  return {
    name,
    params: [p.any('array'), p.scalar('x'), p.scalar('significance', true)],
    summary: 'Returns the percentage rank of a value in a data set.',
    futureFunction: true,
    impl: (args, ctx) => {
      const nums = numbersOf([args[0]], ctx);
      if (isError(nums)) return nums;
      const x = numArg(args[1]);
      if (isError(x)) return x;
      const significance = intArg(args[2], 3);
      if (isError(significance)) return significance;
      return percentRank(nums, x, significance, exclusive);
    },
  };
}
