/**
 * Financial functions.
 *
 * Seven decisions shape this module.
 *
 * First, the sign convention is expressed once. Excel's annuity family is one
 * balance equation - `pv*(1+r)^n + pmt*(1+r*type)*((1+r)^n-1)/r + fv = 0` -
 * rearranged four ways, so PMT, PV, FV and NPER are written as four
 * rearrangements of `annuityBalance` rather than four independently transcribed
 * formulas. Money paid out is negative and money received positive; a loan of
 * 300,000 is entered as `PMT(0.05/12, 360, -300000)` and yields a positive
 * payment. Transcribing the four separately is how implementations end up with
 * a PV that disagrees in sign with their own FV.
 *
 * Second, the end-of-period / beginning-of-period switch is a real branch, not
 * a multiplier bolted on afterwards. An annuity-due payment earns one extra
 * period of interest, and IPMT's first period carries no interest at all under
 * type 1, which no single scaling factor reproduces.
 *
 * Third, degenerate closed forms are separated by cause. An exact division by
 * zero in a closed form - nper of 0 in PMT, a zero rate with a zero payment in
 * NPER - is #DIV/0!, because that is the arithmetic that actually failed.
 * Mathematically impossible input (the logarithm of a non-positive quantity, a
 * fractional root of a negative, a solver that will not converge) is #NUM!.
 * Returning #NUM! where Excel returns #VALUE! or #DIV/0! is a real defect, so
 * the two causes are never merged into one guard.
 *
 * Fourth, the iterative functions carry Microsoft's documented budgets and fail
 * the way Excel fails. RATE and IRR are documented at 20 iterations, XIRR and
 * YIELD at 100; when the iteration does not converge inside its budget the
 * answer is #NUM!, not the last iterate. Every solver also gives up when its
 * slope collapses or an iterate leaves the domain, because a silently wrong
 * number here is worse than an error a user can see.
 *
 * Fifth, NPV and XNPV genuinely differ and are not unified. NPV discounts its
 * first argument by one full period, so a time-zero outlay passed to it is
 * discounted once - the classic user error, which we reproduce rather than
 * correct. XNPV discounts from the first date, actual/365, so its first flow is
 * undiscounted.
 *
 * Sixth, the day-count conventions live here rather than being shared with
 * datetime.ts. Basis 0 for the securities functions is NASD 30/360, which folds
 * the last day of February onto the 30th; DAYS360's US method does not, and the
 * two must not be served by one routine. The coupon schedule is anchored at
 * maturity and stepped in whole months, keeping the original day of month and
 * treating an end-of-month maturity as sticky, which is what makes
 * COUPNCD(.., 31 Aug .., 2) land on 28 February rather than on the 28th of
 * every subsequent month.
 *
 * Seventh, date arguments here accept serial numbers, not text. datetime.ts
 * owns the date-text parser and does not export it, and duplicating a second
 * parser that drifted from the first would be worse than the gap: a text date
 * handed to PRICE yields #VALUE! where Excel would parse it. Numeric text
 * ("45000") still coerces, as it does everywhere else.
 */

import {
  CellError,
  type DateSystem,
  type Scalar,
  isError,
  isLeapYear,
  partsToSerial,
  serialToParts,
} from '@mirrorz/core';
import { ArgKind, type FunctionContext, type FunctionSpec, p } from '../registry.js';
import {
  type Value,
  checkMagnitude,
  excelSub,
  isArray,
  isRef,
  toBoolean,
  toNumber,
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

function numArg(v: Value | undefined, whenOmitted?: number): number | CellError {
  if (v === undefined && whenOmitted !== undefined) return whenOmitted;
  const n = toNumber(scalarArg(v));
  if (isError(n)) return n;
  return Number.isFinite(n) ? n : CellError.NUM;
}

/** A count argument, truncated towards zero the way Excel does before using it. */
function intArg(v: Value | undefined, whenOmitted?: number): number | CellError {
  const n = numArg(v, whenOmitted);
  return isError(n) ? n : Math.trunc(n);
}

/** A flag argument, defaulting when omitted or left blank. */
function flagArg(v: Value | undefined, whenOmitted: boolean): boolean | CellError {
  const s = scalarArg(v);
  if (s === null) return whenOmitted;
  return toBoolean(s);
}

/**
 * The 0/1 annuity-due switch. Excel accepts any non-zero number as 1, but
 * rejects nothing, so this only has to separate zero from non-zero.
 */
function typeArg(v: Value | undefined): 0 | 1 | CellError {
  const n = numArg(v, 0);
  if (isError(n)) return n;
  return n === 0 ? 0 : 1;
}

/** Range-check a computed result: an overflow is #NUM! in Excel, not Infinity. */
function finite(v: number): number | CellError {
  return checkMagnitude(v);
}

/** An error hiding inside an array argument, which the evaluator does not unwrap. */
function arrayError(v: Value | undefined): CellError | undefined {
  if (isError(v)) return v;
  if (isArray(v)) {
    for (const cell of v.data) if (isError(cell)) return cell;
  }
  return undefined;
}

/**
 * Numbers from a cash-flow argument.
 *
 * Values reached through a range or an array follow Excel's aggregate rule -
 * blanks, text and logicals are skipped - while a value typed straight into the
 * formula is coerced, so `NPV(0.1, TRUE)` counts a 1.
 */
function cashFlows(v: Value | undefined): number[] | CellError {
  const err = arrayError(v);
  if (err) return err;
  if (v === undefined) return [];
  if (isArray(v)) {
    const out: number[] = [];
    for (const cell of v.data) if (typeof cell === 'number') out.push(cell);
    return out;
  }
  if (isRef(v)) return CellError.VALUE;
  const n = toNumber(v);
  return isError(n) ? n : [n];
}

/**
 * Every cell of an argument, coerced and kept in position.
 *
 * XNPV and XIRR pair values against dates by position, so a skipped blank would
 * silently shift the whole schedule. Blanks become 0, which is also what makes
 * a missing date fall before the first one and produce Excel's #NUM!.
 */
function positional(v: Value | undefined): number[] | CellError {
  const err = arrayError(v);
  if (err) return err;
  if (v === undefined) return [];
  if (isArray(v)) {
    const out: number[] = [];
    for (const cell of v.data) {
      const n = toNumber(cell);
      if (isError(n)) return n;
      out.push(n);
    }
    return out;
  }
  if (isRef(v)) return CellError.VALUE;
  const n = toNumber(v);
  return isError(n) ? n : [n];
}

// ---------------------------------------------------------------------------
// Root finding
// ---------------------------------------------------------------------------

/**
 * Relative step below which an iterate counts as settled. Excel documents its
 * own tolerances on the answer (1e-7 for RATE, 1e-6 for XIRR); converging
 * harder than that costs one extra Newton step and keeps us bit-comparable with
 * a fully converged oracle, while the iteration budget - which is what actually
 * decides #NUM! - stays exactly Excel's.
 */
const SOLVER_TOLERANCE = 1e-13;

/** A derivative this flat carries no usable information about where the root is. */
const FLAT_SLOPE = 1e-300;

/**
 * Newton-Raphson with an explicit budget.
 *
 * Returns undefined rather than a number whenever the iteration cannot be
 * trusted: a collapsed derivative, a non-finite iterate, or the budget running
 * out. Every caller turns that into #NUM!, which is what Excel reports.
 */
function newton(
  f: (x: number) => number,
  df: (x: number) => number,
  guess: number,
  budget: number,
): number | undefined {
  let x = guess;
  for (let i = 0; i < budget; i++) {
    const y = f(x);
    if (!Number.isFinite(y)) return undefined;
    if (y === 0) return x;
    const slope = df(x);
    if (!Number.isFinite(slope) || Math.abs(slope) < FLAT_SLOPE) return undefined;
    const next = x - y / slope;
    if (!Number.isFinite(next)) return undefined;
    if (Math.abs(next - x) <= SOLVER_TOLERANCE * Math.max(1, Math.abs(next))) return next;
    x = next;
  }
  return undefined;
}

/**
 * Secant iteration, for the roots whose derivative is not worth writing down.
 *
 * The second starting point is nudged rather than supplied, so callers express
 * only the guess Excel's own argument names.
 */
function secant(
  f: (x: number) => number,
  guess: number,
  budget: number,
): number | undefined {
  let x0 = guess;
  let x1 = guess === 0 ? 1e-4 : guess * 1.0001;
  let y0 = f(x0);
  if (!Number.isFinite(y0)) return undefined;
  for (let i = 0; i < budget; i++) {
    const y1 = f(x1);
    if (!Number.isFinite(y1)) return undefined;
    if (y1 === 0) return x1;
    const slope = y1 - y0;
    if (!Number.isFinite(slope) || Math.abs(slope) < FLAT_SLOPE) return undefined;
    const next = x1 - (y1 * (x1 - x0)) / slope;
    if (!Number.isFinite(next)) return undefined;
    const settled = Math.abs(next - x1) <= SOLVER_TOLERANCE * Math.max(1, Math.abs(next));
    x0 = x1;
    y0 = y1;
    x1 = next;
    if (settled) return next;
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// The annuity equation
// ---------------------------------------------------------------------------

/** (1+rate)^nper, the growth factor every annuity rearrangement shares. */
function growth(rate: number, nper: number): number {
  return Math.pow(1 + rate, nper);
}

/**
 * `(1+rate)^nper - 1`, computed so that a small rate does not cancel away.
 *
 * At a rate of 1e-15 over 120 periods the true value is 1.2e-13, and taking it
 * as `pow(1+r, n) - 1` leaves barely two correct digits - enough to turn a
 * payment of 1000 into 900. `expm1` of `log1p` keeps the whole quantity in the
 * range where doubles are dense. The switch happens only where the direct form
 * has actually lost digits, so every ordinary rate still travels the identical
 * `pow` path and reproduces the oracle bit for bit.
 */
const CANCELLATION_RATE = 1e-6;

function growthMinusOne(rate: number, nper: number): number {
  if (Math.abs(rate) < CANCELLATION_RATE) return Math.expm1(nper * Math.log1p(rate));
  return growth(rate, nper) - 1;
}

/**
 * The present-value-weighted annuity factor, `(1+r*type)*((1+r)^n-1)/r`, with
 * the removable singularity at r = 0 filled in by its limit.
 */
function annuityFactor(rate: number, nper: number, type: 0 | 1): number {
  if (rate === 0) return nper;
  return ((1 + rate * type) * growthMinusOne(rate, nper)) / rate;
}

/**
 * The residual Excel's whole annuity family is defined by. It is zero exactly
 * when the five arguments are mutually consistent, which is what RATE solves.
 */
function annuityBalance(
  rate: number,
  nper: number,
  pmt: number,
  pv: number,
  fv: number,
  type: 0 | 1,
): number {
  if (rate === 0) return pv + pmt * nper + fv;
  return pv * growth(rate, nper) + pmt * annuityFactor(rate, nper, type) + fv;
}

function fvOf(rate: number, nper: number, pmt: number, pv: number, type: 0 | 1): number {
  if (rate === 0) return -(pv + pmt * nper);
  return -(pv * growth(rate, nper) + pmt * annuityFactor(rate, nper, type));
}

function pvOf(
  rate: number,
  nper: number,
  pmt: number,
  fv: number,
  type: 0 | 1,
): number | CellError {
  if (rate === 0) return -(fv + pmt * nper);
  const f = growth(rate, nper);
  if (f === 0) return CellError.DIV0;
  return -(fv + pmt * annuityFactor(rate, nper, type)) / f;
}

function pmtOf(
  rate: number,
  nper: number,
  pv: number,
  fv: number,
  type: 0 | 1,
): number | CellError {
  if (nper === 0) return CellError.DIV0;
  if (rate === 0) return -(pv + fv) / nper;
  const factor = annuityFactor(rate, nper, type);
  if (factor === 0) return CellError.DIV0;
  return -(fv + pv * growth(rate, nper)) / factor;
}

// ---------------------------------------------------------------------------
// Calendar and day counts
// ---------------------------------------------------------------------------

/** A date argument, reduced to whole days as every securities function does. */
function dateArg(v: Value | undefined): number | CellError {
  const n = numArg(v);
  if (isError(n)) return n;
  // A negative serial is not a date Excel will represent.
  if (n < 0) return CellError.NUM;
  return Math.floor(n);
}

interface Dt {
  serial: number;
  year: number;
  month: number;
  day: number;
}

function dt(serial: number, system: DateSystem): Dt {
  const parts = serialToParts(serial, system);
  return { serial, year: parts.year, month: parts.month, day: parts.day };
}

function epochYear(system: DateSystem): number {
  return system === 1904 ? 1904 : 1900;
}

/**
 * Build a serial from calendar parts, refusing anything before the workbook's
 * epoch. `partsToSerial` folds years 0-1899 into the twentieth century, which
 * is right for DATE() and catastrophic for a coupon schedule stepped backwards
 * off the bottom of the calendar.
 */
function serialFromYmd(
  year: number,
  month: number,
  day: number,
  system: DateSystem,
): number | CellError {
  if (year < epochYear(system) || year > 9999) return CellError.NUM;
  return partsToSerial(year, month, day, 0, 0, 0, system);
}

function daysInMonthOf(year: number, month: number): number {
  if (month === 2) return isLeapYear(year) ? 29 : 28;
  return [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1] ?? 30;
}

function isLastOfFebruary(d: Dt): boolean {
  return d.month === 2 && d.day === daysInMonthOf(d.year, 2);
}

/**
 * NASD 30/360, the day count behind basis 0 of every securities function.
 *
 * The end-of-February fold is what separates this from DAYS360's US method,
 * which leaves February alone; sharing one routine between them disagrees with
 * Excel across every February.
 */
function days360Nasd(from: Dt, to: Dt): number {
  let d1 = from.day;
  let d2 = to.day;
  if (isLastOfFebruary(from) && isLastOfFebruary(to)) d2 = 30;
  if (isLastOfFebruary(from)) d1 = 30;
  if (d2 === 31 && d1 >= 30) d2 = 30;
  if (d1 === 31) d1 = 30;
  return (to.year - from.year) * 360 + (to.month - from.month) * 30 + (d2 - d1);
}

/** European 30/360, basis 4: both endpoints simply clamp to the 30th. */
function days360Euro(from: Dt, to: Dt): number {
  const d1 = Math.min(from.day, 30);
  const d2 = Math.min(to.day, 30);
  return (to.year - from.year) * 360 + (to.month - from.month) * 30 + (d2 - d1);
}

/** Days between two dates under a day-count basis. */
function dayCount(from: Dt, to: Dt, basis: number): number {
  if (basis === 0) return days360Nasd(from, to);
  if (basis === 4) return days360Euro(from, to);
  return excelSub(to.serial, from.serial);
}

/**
 * Length of the year a fraction is measured against.
 *
 * Basis 1 is the awkward one: Excel does not use 365.25 but the length of the
 * year the interval sits in, or the mean length of every calendar year the
 * interval touches once it spans more than one.
 */
function yearLength(from: Dt, to: Dt, basis: number): number {
  switch (basis) {
    case 0:
    case 2:
    case 4:
      return 360;
    case 3:
      return 365;
    default:
      return actualYearLength(from, to);
  }
}

function actualYearLength(from: Dt, to: Dt): number {
  const spansMoreThanAYear =
    to.year > from.year + 1 ||
    (to.year === from.year + 1 &&
      (to.month > from.month || (to.month === from.month && to.day > from.day)));

  if (spansMoreThanAYear) {
    let total = 0;
    for (let y = from.year; y <= to.year; y++) total += isLeapYear(y) ? 366 : 365;
    return total / (to.year - from.year + 1);
  }
  if (from.year === to.year) return isLeapYear(from.year) ? 366 : 365;
  // One calendar boundary is crossed, so the year is 366 days long only when a
  // 29 February actually falls inside the interval.
  for (const y of [from.year, to.year]) {
    if (!isLeapYear(y)) continue;
    const leapDay = Date.UTC(y, 1, 29);
    if (leapDay >= Date.UTC(from.year, from.month - 1, from.day) &&
        leapDay <= Date.UTC(to.year, to.month - 1, to.day)) {
      return 366;
    }
  }
  return 365;
}

/** The fraction of a year between two dates, as the securities functions use it. */
function yearFraction(from: Dt, to: Dt, basis: number): number {
  if (from.serial === to.serial) return 0;
  const reversed = from.serial > to.serial;
  const a = reversed ? to : from;
  const b = reversed ? from : to;
  const frac = dayCount(a, b, basis) / yearLength(a, b, basis);
  return reversed ? -frac : frac;
}

// ---------------------------------------------------------------------------
// Coupon schedules
// ---------------------------------------------------------------------------

/** Excel accepts only annual, semi-annual and quarterly coupons. */
function frequencyArg(v: Value | undefined): 1 | 2 | 4 | CellError {
  const n = intArg(v);
  if (isError(n)) return n;
  if (n === 1 || n === 2 || n === 4) return n;
  return CellError.NUM;
}

function basisArg(v: Value | undefined): number | CellError {
  const n = intArg(v, 0);
  if (isError(n)) return n;
  if (n < 0 || n > 4) return CellError.NUM;
  return n;
}

/**
 * The coupon schedule, anchored at maturity.
 *
 * Coupon dates are maturity stepped back by whole months, keeping the original
 * day of month - so a 31st that passes through February comes back out as a
 * 31st - and treating an end-of-month maturity as sticky, which is how Excel
 * puts a 31 August bond's semi-annual coupon on 28 February.
 */
interface Schedule {
  monthsPerPeriod: number;
  anchor: Dt;
  anchorIsLastDay: boolean;
  system: DateSystem;
}

function scheduleOf(maturity: Dt, frequency: number, system: DateSystem): Schedule {
  return {
    monthsPerPeriod: 12 / frequency,
    anchor: maturity,
    anchorIsLastDay: maturity.day === daysInMonthOf(maturity.year, maturity.month),
    system,
  };
}

/** The coupon date `k` whole periods before maturity. */
function couponDate(s: Schedule, k: number): Dt | CellError {
  const total = s.anchor.year * 12 + (s.anchor.month - 1) - k * s.monthsPerPeriod;
  const year = Math.floor(total / 12);
  const month = total - year * 12 + 1;
  const day = s.anchorIsLastDay
    ? daysInMonthOf(year, month)
    : Math.min(s.anchor.day, daysInMonthOf(year, month));
  const serial = serialFromYmd(year, month, day, s.system);
  if (isError(serial)) return serial;
  return { serial, year, month, day };
}

/** How far a runaway coupon search is allowed to walk before it is a bug. */
const MAX_COUPON_PERIODS = 1_000_000;

/**
 * The number of coupon periods from settlement to maturity, which is also the
 * index of the coupon date at or before settlement. COUPNUM is exactly this
 * count, because stepping back k periods moves k*12/frequency months and
 * COUPNUM is that month span scaled back by the frequency.
 */
function periodsToMaturity(s: Schedule, settle: number): number | CellError {
  const settleDt = dt(settle, s.system);
  const months = (s.anchor.year - settleDt.year) * 12 + (s.anchor.month - settleDt.month);
  let k = Math.max(0, Math.floor(months / s.monthsPerPeriod));

  for (let guard = 0; guard <= MAX_COUPON_PERIODS; guard++) {
    const here = couponDate(s, k);
    if (isError(here)) return here;
    if (here.serial > settle) {
      k++;
      continue;
    }
    if (k === 0) return CellError.NUM;
    const previous = couponDate(s, k - 1);
    if (isError(previous)) return previous;
    if (previous.serial <= settle) {
      k--;
      continue;
    }
    return k;
  }
  return CellError.NUM;
}

/** The four coupon-period measurements the bond formulas are written in terms of. */
interface CouponGeometry {
  /** Coupons remaining, N in Excel's documentation. */
  count: number;
  /** Days in the coupon period containing settlement, E. */
  periodDays: number;
  /** Days from the previous coupon date to settlement, A. */
  sinceCoupon: number;
  /** Days from settlement to the next coupon date, DSC. */
  toCoupon: number;
  previous: Dt;
  next: Dt;
}

function couponGeometry(
  settle: number,
  maturity: number,
  frequency: number,
  basis: number,
  system: DateSystem,
): CouponGeometry | CellError {
  const s = scheduleOf(dt(maturity, system), frequency, system);
  const k = periodsToMaturity(s, settle);
  if (isError(k)) return k;
  const previous = couponDate(s, k);
  if (isError(previous)) return previous;
  const next = couponDate(s, k - 1);
  if (isError(next)) return next;

  const settleDt = dt(settle, system);
  const periodDays =
    basis === 1
      ? excelSub(next.serial, previous.serial)
      : yearLength(previous, next, basis) / frequency;
  const sinceCoupon = dayCount(previous, settleDt, basis);
  // For the 30/360 bases the two halves must add up to the period exactly, so
  // the far side is taken as the remainder rather than counted independently.
  const toCoupon =
    basis === 0 || basis === 4
      ? periodDays - sinceCoupon
      : dayCount(settleDt, next, basis);

  return { count: k, periodDays, sinceCoupon, toCoupon, previous, next };
}

// ---------------------------------------------------------------------------
// Bond price and yield
// ---------------------------------------------------------------------------

/**
 * Excel's PRICE, as a plain function so YIELD can invert it.
 *
 * The single-period branch is not the general sum evaluated at N = 1: Excel
 * discounts a bond in its last coupon period simple, not compound, which is the
 * money-market convention and a visibly different number.
 */
function bondPrice(
  g: CouponGeometry,
  rate: number,
  yld: number,
  redemption: number,
  frequency: number,
): number {
  const coupon = (100 * rate) / frequency;
  const accrued = coupon * (g.sinceCoupon / g.periodDays);

  if (g.count <= 1) {
    const remaining = (g.periodDays - g.sinceCoupon) / g.periodDays;
    return (coupon + redemption) / (1 + (yld / frequency) * remaining) - accrued;
  }

  const offset = g.toCoupon / g.periodDays;
  const step = 1 + yld / frequency;
  let price = redemption / Math.pow(step, g.count - 1 + offset);
  for (let k = 1; k <= g.count; k++) {
    price += coupon / Math.pow(step, k - 1 + offset);
  }
  return price - accrued;
}

// ---------------------------------------------------------------------------
// Depreciation
// ---------------------------------------------------------------------------

/** One period of double-declining depreciation, salvage floor included. */
function ddbAt(
  cost: number,
  salvage: number,
  life: number,
  period: number,
  factor: number,
): number {
  let rate = factor / life;
  let opening: number;
  if (rate >= 1) {
    // A factor at or beyond the life writes the whole asset off at once.
    rate = 1;
    opening = period === 1 ? cost : 0;
  } else {
    opening = cost * Math.pow(1 - rate, period - 1);
  }
  const closing = cost * Math.pow(1 - rate, period);
  const amount = closing < salvage ? opening - salvage : opening - closing;
  return amount < 0 ? 0 : amount;
}

/** How long a VDB walk may run before the input is treated as a mistake. */
const MAX_DEPRECIATION_PERIODS = 1_000_000;

/**
 * Declining balance with the switch to straight line, summed to `period`.
 *
 * `remainingLife` is separate from `life` because VDB re-bases the asset after
 * the start period: the declining-balance rate keeps referring to the original
 * life while the straight-line leg spreads what is left over the life that
 * remains.
 */
function vdbSwitching(
  cost: number,
  salvage: number,
  life: number,
  remainingLife: number,
  period: number,
  factor: number,
): number {
  const wholePeriods = Math.ceil(period);
  let total = 0;
  let straightLine = 0;
  let switched = false;
  let remaining = cost - salvage;

  for (let i = 1; i <= wholePeriods; i++) {
    let amount: number;
    if (switched) {
      amount = straightLine;
    } else {
      const declining = ddbAt(cost, salvage, life, i, factor);
      straightLine = remaining / (remainingLife - (i - 1));
      if (straightLine > declining) {
        amount = straightLine;
        switched = true;
      } else {
        amount = declining;
        remaining -= declining;
      }
    }
    // The final period is pro-rated when the caller asked for a fraction of it.
    if (i === wholePeriods) amount *= period + 1 - wholePeriods;
    total += amount;
  }
  return total;
}

// ---------------------------------------------------------------------------
// Fixed-yield helpers shared by the discounted-security functions
// ---------------------------------------------------------------------------

interface Span {
  from: Dt;
  to: Dt;
  basis: number;
}

function spanOf(
  settle: number,
  maturity: number,
  basis: number,
  system: DateSystem,
): Span | CellError {
  if (settle >= maturity) return CellError.NUM;
  return { from: dt(settle, system), to: dt(maturity, system), basis };
}

function spanFraction(s: Span): number {
  return yearFraction(s.from, s.to, s.basis);
}

// ---------------------------------------------------------------------------
// Function specifications: the annuity family
// ---------------------------------------------------------------------------

const PMT: FunctionSpec = {
  name: 'PMT',
  params: [
    p.scalar('rate'),
    p.scalar('nper'),
    p.scalar('pv'),
    p.scalar('fv', true),
    p.scalar('type', true),
  ],
  broadcast: true,
  summary: 'The periodic payment for a loan or annuity at a constant rate.',
  impl: (args) => {
    const rate = numArg(args[0]);
    if (isError(rate)) return rate;
    const nper = numArg(args[1]);
    if (isError(nper)) return nper;
    const pv = numArg(args[2]);
    if (isError(pv)) return pv;
    const fv = numArg(args[3], 0);
    if (isError(fv)) return fv;
    const type = typeArg(args[4]);
    if (isError(type)) return type;
    const result = pmtOf(rate, nper, pv, fv, type);
    return isError(result) ? result : finite(result);
  },
};

const PV: FunctionSpec = {
  name: 'PV',
  params: [
    p.scalar('rate'),
    p.scalar('nper'),
    p.scalar('pmt'),
    p.scalar('fv', true),
    p.scalar('type', true),
  ],
  broadcast: true,
  summary: 'The present value of a stream of equal payments.',
  impl: (args) => {
    const rate = numArg(args[0]);
    if (isError(rate)) return rate;
    const nper = numArg(args[1]);
    if (isError(nper)) return nper;
    const pmt = numArg(args[2]);
    if (isError(pmt)) return pmt;
    const fv = numArg(args[3], 0);
    if (isError(fv)) return fv;
    const type = typeArg(args[4]);
    if (isError(type)) return type;
    const result = pvOf(rate, nper, pmt, fv, type);
    return isError(result) ? result : finite(result);
  },
};

const FV: FunctionSpec = {
  name: 'FV',
  params: [
    p.scalar('rate'),
    p.scalar('nper'),
    p.scalar('pmt'),
    p.scalar('pv', true),
    p.scalar('type', true),
  ],
  broadcast: true,
  summary: 'The future value of a stream of equal payments.',
  impl: (args) => {
    const rate = numArg(args[0]);
    if (isError(rate)) return rate;
    const nper = numArg(args[1]);
    if (isError(nper)) return nper;
    const pmt = numArg(args[2]);
    if (isError(pmt)) return pmt;
    const pv = numArg(args[3], 0);
    if (isError(pv)) return pv;
    const type = typeArg(args[4]);
    if (isError(type)) return type;
    return finite(fvOf(rate, nper, pmt, pv, type));
  },
};

const NPER: FunctionSpec = {
  name: 'NPER',
  params: [
    p.scalar('rate'),
    p.scalar('pmt'),
    p.scalar('pv'),
    p.scalar('fv', true),
    p.scalar('type', true),
  ],
  broadcast: true,
  summary: 'The number of periods needed to repay a loan or reach a future value.',
  impl: (args) => {
    const rate = numArg(args[0]);
    if (isError(rate)) return rate;
    const pmt = numArg(args[1]);
    if (isError(pmt)) return pmt;
    const pv = numArg(args[2]);
    if (isError(pv)) return pv;
    const fv = numArg(args[3], 0);
    if (isError(fv)) return fv;
    const type = typeArg(args[4]);
    if (isError(type)) return type;

    if (rate === 0) {
      if (pmt === 0) return CellError.DIV0;
      return finite(-(pv + fv) / pmt);
    }
    if (rate <= -1) return CellError.NUM;
    const due = pmt * (1 + rate * type);
    const numerator = due - fv * rate;
    const denominator = pv * rate + due;
    if (denominator === 0) return CellError.DIV0;
    const ratio = numerator / denominator;
    // The logarithm is where an inconsistent set of arguments shows up: no
    // number of periods can reconcile them, which is Excel's #NUM!.
    if (!(ratio > 0)) return CellError.NUM;
    return finite(Math.log(ratio) / Math.log(1 + rate));
  },
};

/** Excel documents RATE as giving up after twenty iterations. */
const RATE_BUDGET = 20;

const RATE: FunctionSpec = {
  name: 'RATE',
  params: [
    p.scalar('nper'),
    p.scalar('pmt'),
    p.scalar('pv'),
    p.scalar('fv', true),
    p.scalar('type', true),
    p.scalar('guess', true),
  ],
  broadcast: true,
  summary: 'The interest rate per period of an annuity, found iteratively.',
  impl: (args) => {
    const nper = numArg(args[0]);
    if (isError(nper)) return nper;
    const pmt = numArg(args[1]);
    if (isError(pmt)) return pmt;
    const pv = numArg(args[2]);
    if (isError(pv)) return pv;
    const fv = numArg(args[3], 0);
    if (isError(fv)) return fv;
    const type = typeArg(args[4]);
    if (isError(type)) return type;
    const guess = numArg(args[5], 0.1);
    if (isError(guess)) return guess;

    if (nper <= 0) return CellError.NUM;
    if (guess <= -1) return CellError.NUM;

    // The residual is the balance equation divided through by its own annuity
    // factor - that is, the payment the other four arguments imply, less the
    // payment asked for. Iterating on the raw balance instead overflows on any
    // ordinary mortgage: (1+0.1)^360 from the default guess is 8e14, and a
    // secant step taken on numbers that size lands nowhere near the root.
    const residual = (r: number): number => {
      const implied = pmtOf(r, nper, pv, fv, type);
      return isError(implied) ? Number.NaN : implied - pmt;
    };
    const root = secant(residual, guess, RATE_BUDGET);
    if (root === undefined || root <= -1) return CellError.NUM;
    return finite(root);
  },
};

const IPMT: FunctionSpec = {
  name: 'IPMT',
  params: [
    p.scalar('rate'),
    p.scalar('per'),
    p.scalar('nper'),
    p.scalar('pv'),
    p.scalar('fv', true),
    p.scalar('type', true),
  ],
  broadcast: true,
  summary: 'The interest portion of one payment of an annuity.',
  impl: (args) => {
    const parts = periodSplit(args);
    if (isError(parts)) return parts;
    return finite(parts.interest);
  },
};

const PPMT: FunctionSpec = {
  name: 'PPMT',
  params: [
    p.scalar('rate'),
    p.scalar('per'),
    p.scalar('nper'),
    p.scalar('pv'),
    p.scalar('fv', true),
    p.scalar('type', true),
  ],
  broadcast: true,
  summary: 'The principal portion of one payment of an annuity.',
  impl: (args) => {
    const parts = periodSplit(args);
    if (isError(parts)) return parts;
    return finite(parts.principal);
  },
};

/**
 * Split one payment into interest and principal.
 *
 * Interest is charged on the balance standing at the start of the period, which
 * under an annuity due is the balance after that period's own payment has
 * already been made - hence the two-period lookback and the interest-free first
 * period. IPMT and PPMT share this so they cannot drift apart; PPMT is defined
 * as the remainder, which keeps IPMT + PPMT exactly equal to PMT.
 */
function periodSplit(
  args: Value[],
): { interest: number; principal: number } | CellError {
  const rate = numArg(args[0]);
  if (isError(rate)) return rate;
  const per = numArg(args[1]);
  if (isError(per)) return per;
  const nper = numArg(args[2]);
  if (isError(nper)) return nper;
  const pv = numArg(args[3]);
  if (isError(pv)) return pv;
  const fv = numArg(args[4], 0);
  if (isError(fv)) return fv;
  const type = typeArg(args[5]);
  if (isError(type)) return type;

  if (per < 1 || per > nper) return CellError.NUM;
  const payment = pmtOf(rate, nper, pv, fv, type);
  if (isError(payment)) return payment;

  let balance: number;
  if (per === 1) {
    balance = type === 1 ? 0 : -pv;
  } else if (type === 1) {
    balance = fvOf(rate, per - 2, payment, pv, 1) - payment;
  } else {
    balance = fvOf(rate, per - 1, payment, pv, 0);
  }
  const interest = balance * rate;
  return { interest, principal: payment - interest };
}

const CUMIPMT: FunctionSpec = {
  name: 'CUMIPMT',
  params: [
    p.scalar('rate'),
    p.scalar('nper'),
    p.scalar('pv'),
    p.scalar('start_period'),
    p.scalar('end_period'),
    p.scalar('type'),
  ],
  broadcast: true,
  summary: 'The cumulative interest paid between two periods of a loan.',
  impl: (args) => {
    const c = cumulative(args);
    if (isError(c)) return c;
    return finite(c.periods * c.payment - c.principal);
  },
};

const CUMPRINC: FunctionSpec = {
  name: 'CUMPRINC',
  params: [
    p.scalar('rate'),
    p.scalar('nper'),
    p.scalar('pv'),
    p.scalar('start_period'),
    p.scalar('end_period'),
    p.scalar('type'),
  ],
  broadcast: true,
  summary: 'The cumulative principal repaid between two periods of a loan.',
  impl: (args) => {
    const c = cumulative(args);
    if (isError(c)) return c;
    return finite(c.principal);
  },
};

/**
 * The shared body of CUMIPMT and CUMPRINC.
 *
 * Cumulative principal is the change in outstanding balance across the span,
 * not a loop over PPMT: the closed form is exact, is O(1) on a thirty-year
 * schedule, and cannot be made to hang by a hostile `nper`. Cumulative interest
 * is then whatever the payments did not repay.
 */
function cumulative(
  args: Value[],
): { payment: number; principal: number; periods: number } | CellError {
  const rate = numArg(args[0]);
  if (isError(rate)) return rate;
  const nper = numArg(args[1]);
  if (isError(nper)) return nper;
  const pv = numArg(args[2]);
  if (isError(pv)) return pv;
  const start = numArg(args[3]);
  if (isError(start)) return start;
  const end = numArg(args[4]);
  if (isError(end)) return end;
  const rawType = numArg(args[5]);
  if (isError(rawType)) return rawType;

  // Excel documents every one of these as #NUM!, including a type that is
  // neither 0 nor 1 - which it accepts everywhere else.
  if (rate <= 0 || nper <= 0 || pv <= 0) return CellError.NUM;
  if (start < 1 || end < 1 || start > end) return CellError.NUM;
  if (rawType !== 0 && rawType !== 1) return CellError.NUM;
  if (end > nper) return CellError.NUM;
  const type: 0 | 1 = rawType === 0 ? 0 : 1;

  const payment = pmtOf(rate, nper, pv, 0, type);
  if (isError(payment)) return payment;
  const principal = excelSub(
    principalThrough(end, rate, payment, pv, type),
    principalThrough(start - 1, rate, payment, pv, type),
  );
  return { payment, principal, periods: end - start + 1 };
}

/**
 * Principal repaid over the first `count` periods.
 *
 * With payment in arrears this is just the movement in the outstanding
 * balance. With payment in advance it is not, and assuming so is the mistake
 * that makes an annuity-due CUMPRINC report a first-period repayment of 68
 * where the whole first payment of 998 is principal: under type 1 the balance
 * at the end of period k already carries interest charged on a payment made at
 * its start, so the principal in period k telescopes one step further back.
 */
function principalThrough(
  count: number,
  rate: number,
  payment: number,
  pv: number,
  type: 0 | 1,
): number {
  if (count <= 0) return 0;
  if (type === 0) {
    return excelSub(fvOf(rate, 0, payment, pv, 0), fvOf(rate, count, payment, pv, 0));
  }
  return (
    payment + excelSub(fvOf(rate, 0, payment, pv, 1), fvOf(rate, count - 1, payment, pv, 1))
  );
}

const ISPMT: FunctionSpec = {
  name: 'ISPMT',
  params: [p.scalar('rate'), p.scalar('per'), p.scalar('nper'), p.scalar('pv')],
  broadcast: true,
  summary: 'Interest paid during a period of a loan with equal principal repayments.',
  impl: (args) => {
    const rate = numArg(args[0]);
    if (isError(rate)) return rate;
    const per = numArg(args[1]);
    if (isError(per)) return per;
    const nper = numArg(args[2]);
    if (isError(nper)) return nper;
    const pv = numArg(args[3]);
    if (isError(pv)) return pv;
    if (nper === 0) return CellError.DIV0;
    // ISPMT numbers its periods from zero, which is why the first period's
    // interest is the full pv*rate rather than one step in from it.
    return finite(pv * rate * (per / nper - 1));
  },
};

// ---------------------------------------------------------------------------
// Function specifications: cash-flow series
// ---------------------------------------------------------------------------

const NPV: FunctionSpec = {
  name: 'NPV',
  params: [p.scalar('rate'), p.array('value1'), p.rest('value', ArgKind.Array)],
  summary: 'Net present value of a series of cash flows, discounted from period 1.',
  impl: (args) => {
    const rate = numArg(args[0]);
    if (isError(rate)) return rate;
    if (rate === -1) return CellError.DIV0;

    let total = 0;
    let period = 0;
    for (let i = 1; i < args.length; i++) {
      const flows = cashFlows(args[i]);
      if (isError(flows)) return flows;
      for (const flow of flows) {
        period++;
        total += flow / Math.pow(1 + rate, period);
      }
    }
    return finite(total);
  },
};

/** Excel documents IRR, like RATE, as giving up after twenty iterations. */
const IRR_BUDGET = 20;
/** XIRR is documented at a hundred, and needs them: its exponents are unequal. */
const XIRR_BUDGET = 100;

const IRR: FunctionSpec = {
  name: 'IRR',
  params: [p.array('values'), p.scalar('guess', true)],
  summary: 'The internal rate of return of a series of periodic cash flows.',
  impl: (args) => {
    const flows = cashFlows(args[0]);
    if (isError(flows)) return flows;
    const guess = numArg(args[1], 0.1);
    if (isError(guess)) return guess;

    if (flows.length < 2) return CellError.NUM;
    if (!flows.some((v) => v > 0) || !flows.some((v) => v < 0)) return CellError.NUM;
    if (guess <= -1) return CellError.NUM;

    const value = (r: number): number => {
      let sum = 0;
      for (let i = 0; i < flows.length; i++) sum += flows[i]! / Math.pow(1 + r, i);
      return sum;
    };
    const slope = (r: number): number => {
      let sum = 0;
      for (let i = 1; i < flows.length; i++) sum -= (i * flows[i]!) / Math.pow(1 + r, i + 1);
      return sum;
    };

    const root = newton(value, slope, guess, IRR_BUDGET);
    if (root === undefined || root <= -1) return CellError.NUM;
    return finite(root);
  },
};

const MIRR: FunctionSpec = {
  name: 'MIRR',
  params: [p.array('values'), p.scalar('finance_rate'), p.scalar('reinvest_rate')],
  summary: 'Internal rate of return with distinct borrowing and reinvestment rates.',
  impl: (args) => {
    const flows = cashFlows(args[0]);
    if (isError(flows)) return flows;
    const finance = numArg(args[1]);
    if (isError(finance)) return finance;
    const reinvest = numArg(args[2]);
    if (isError(reinvest)) return reinvest;

    const n = flows.length;
    if (n < 2) return CellError.DIV0;
    if (!flows.some((v) => v > 0) || !flows.some((v) => v < 0)) return CellError.DIV0;
    if (finance === -1 || reinvest === -1) return CellError.DIV0;

    let positive = 0;
    let negative = 0;
    for (let i = 0; i < n; i++) {
      const v = flows[i]!;
      if (v > 0) positive += v / Math.pow(1 + reinvest, i + 1);
      else if (v < 0) negative += v / Math.pow(1 + finance, i + 1);
    }
    if (negative === 0) return CellError.DIV0;
    const ratio = -(positive * Math.pow(1 + reinvest, n)) / (negative * (1 + finance));
    if (!(ratio > 0)) return CellError.NUM;
    return finite(Math.pow(ratio, 1 / (n - 1)) - 1);
  },
};

/** XNPV and XIRR discount on actual days over a fixed 365-day year. */
const XNPV_YEAR = 365;

function irregularSeries(
  valuesArg: Value | undefined,
  datesArg: Value | undefined,
): { values: number[]; offsets: number[] } | CellError {
  const values = positional(valuesArg);
  if (isError(values)) return values;
  const rawDates = positional(datesArg);
  if (isError(rawDates)) return rawDates;
  if (values.length !== rawDates.length || values.length < 2) return CellError.NUM;

  const dates = rawDates.map((d) => Math.floor(d));
  const first = dates[0]!;
  const offsets: number[] = [];
  for (const d of dates) {
    if (d < 0) return CellError.NUM;
    // Excel rejects a schedule that steps backwards past its own start date.
    if (d < first) return CellError.NUM;
    offsets.push(excelSub(d, first) / XNPV_YEAR);
  }
  return { values, offsets };
}

const XNPV: FunctionSpec = {
  name: 'XNPV',
  params: [p.scalar('rate'), p.array('values'), p.array('dates')],
  summary: 'Net present value of cash flows on arbitrary dates, actual/365.',
  impl: (args) => {
    const rate = numArg(args[0]);
    if (isError(rate)) return rate;
    const series = irregularSeries(args[1], args[2]);
    if (isError(series)) return series;
    if (rate <= -1) return CellError.NUM;

    let total = 0;
    for (let i = 0; i < series.values.length; i++) {
      total += series.values[i]! / Math.pow(1 + rate, series.offsets[i]!);
    }
    return finite(total);
  },
};

const XIRR: FunctionSpec = {
  name: 'XIRR',
  params: [p.array('values'), p.array('dates'), p.scalar('guess', true)],
  summary: 'Internal rate of return for cash flows on arbitrary dates.',
  impl: (args) => {
    const series = irregularSeries(args[0], args[1]);
    if (isError(series)) return series;
    const guess = numArg(args[2], 0.1);
    if (isError(guess)) return guess;

    const { values, offsets } = series;
    if (!values.some((v) => v > 0) || !values.some((v) => v < 0)) return CellError.NUM;
    if (guess <= -1) return CellError.NUM;

    const value = (r: number): number => {
      let sum = 0;
      for (let i = 0; i < values.length; i++) sum += values[i]! / Math.pow(1 + r, offsets[i]!);
      return sum;
    };
    const slope = (r: number): number => {
      let sum = 0;
      for (let i = 0; i < values.length; i++) {
        sum -= (offsets[i]! * values[i]!) / Math.pow(1 + r, offsets[i]! + 1);
      }
      return sum;
    };

    const root = newton(value, slope, guess, XIRR_BUDGET);
    if (root === undefined || root <= -1) return CellError.NUM;
    return finite(root);
  },
};

const FVSCHEDULE: FunctionSpec = {
  name: 'FVSCHEDULE',
  params: [p.scalar('principal'), p.array('schedule')],
  summary: 'The future value of a principal compounded at a series of rates.',
  impl: (args) => {
    const principal = numArg(args[0]);
    if (isError(principal)) return principal;
    const err = arrayError(args[1]);
    if (err) return err;

    let value = principal;
    const schedule = args[1];
    if (isArray(schedule)) {
      for (const cell of schedule.data) {
        // Excel documents an empty cell in a schedule as a zero rate.
        const rate = toNumber(cell);
        if (isError(rate)) return rate;
        value *= 1 + rate;
      }
    } else {
      const rate = toNumber(scalarArg(schedule));
      if (isError(rate)) return rate;
      value *= 1 + rate;
    }
    return finite(value);
  },
};

// ---------------------------------------------------------------------------
// Function specifications: depreciation
// ---------------------------------------------------------------------------

const SLN: FunctionSpec = {
  name: 'SLN',
  params: [p.scalar('cost'), p.scalar('salvage'), p.scalar('life')],
  broadcast: true,
  summary: 'Straight-line depreciation for one period.',
  impl: (args) => {
    const cost = numArg(args[0]);
    if (isError(cost)) return cost;
    const salvage = numArg(args[1]);
    if (isError(salvage)) return salvage;
    const life = numArg(args[2]);
    if (isError(life)) return life;
    if (life === 0) return CellError.DIV0;
    return finite(excelSub(cost, salvage) / life);
  },
};

const SYD: FunctionSpec = {
  name: 'SYD',
  params: [p.scalar('cost'), p.scalar('salvage'), p.scalar('life'), p.scalar('per')],
  broadcast: true,
  summary: 'Sum-of-years-digits depreciation for one period.',
  impl: (args) => {
    const cost = numArg(args[0]);
    if (isError(cost)) return cost;
    const salvage = numArg(args[1]);
    if (isError(salvage)) return salvage;
    const life = numArg(args[2]);
    if (isError(life)) return life;
    const per = numArg(args[3]);
    if (isError(per)) return per;
    if (salvage < 0 || life <= 0) return CellError.NUM;
    if (per <= 0 || per > life) return CellError.NUM;
    return finite((excelSub(cost, salvage) * (life - per + 1) * 2) / (life * (life + 1)));
  },
};

const DB: FunctionSpec = {
  name: 'DB',
  params: [
    p.scalar('cost'),
    p.scalar('salvage'),
    p.scalar('life'),
    p.scalar('period'),
    p.scalar('month', true),
  ],
  broadcast: true,
  summary: 'Fixed-declining-balance depreciation for one period.',
  impl: (args) => {
    const cost = numArg(args[0]);
    if (isError(cost)) return cost;
    const salvage = numArg(args[1]);
    if (isError(salvage)) return salvage;
    const life = intArg(args[2]);
    if (isError(life)) return life;
    const period = intArg(args[3]);
    if (isError(period)) return period;
    const month = intArg(args[4], 12);
    if (isError(month)) return month;

    if (cost < 0 || salvage < 0 || life <= 0 || period <= 0) return CellError.NUM;
    if (month < 1 || month > 12) return CellError.NUM;
    if (period > life + 1) return CellError.NUM;
    if (cost === 0) return 0;

    // Excel rounds the declining rate to three decimals before using it, and
    // the rounding is visible in the answer, not cosmetic.
    const rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000;
    const first = (cost * rate * month) / 12;
    if (period === 1) return finite(first);

    const base = excelSub(cost, first);
    if (period === life + 1) {
      return finite((base * Math.pow(1 - rate, life - 1) * rate * (12 - month)) / 12);
    }
    return finite(base * Math.pow(1 - rate, period - 2) * rate);
  },
};

const DDB: FunctionSpec = {
  name: 'DDB',
  params: [
    p.scalar('cost'),
    p.scalar('salvage'),
    p.scalar('life'),
    p.scalar('period'),
    p.scalar('factor', true),
  ],
  broadcast: true,
  summary: 'Double-declining-balance depreciation for one period.',
  impl: (args) => {
    const cost = numArg(args[0]);
    if (isError(cost)) return cost;
    const salvage = numArg(args[1]);
    if (isError(salvage)) return salvage;
    const life = numArg(args[2]);
    if (isError(life)) return life;
    const period = numArg(args[3]);
    if (isError(period)) return period;
    const factor = numArg(args[4], 2);
    if (isError(factor)) return factor;

    if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0) {
      return CellError.NUM;
    }
    if (period > life) return CellError.NUM;
    return finite(ddbAt(cost, salvage, life, period, factor));
  },
};

const VDB: FunctionSpec = {
  name: 'VDB',
  params: [
    p.scalar('cost'),
    p.scalar('salvage'),
    p.scalar('life'),
    p.scalar('start_period'),
    p.scalar('end_period'),
    p.scalar('factor', true),
    p.scalar('no_switch', true),
  ],
  broadcast: true,
  summary: 'Declining-balance depreciation across a span of periods.',
  impl: (args) => {
    const cost = numArg(args[0]);
    if (isError(cost)) return cost;
    const salvage = numArg(args[1]);
    if (isError(salvage)) return salvage;
    const life = numArg(args[2]);
    if (isError(life)) return life;
    const start = numArg(args[3]);
    if (isError(start)) return start;
    const end = numArg(args[4]);
    if (isError(end)) return end;
    const factor = numArg(args[5], 2);
    if (isError(factor)) return factor;
    const noSwitch = flagArg(args[6], false);
    if (isError(noSwitch)) return noSwitch;

    if (cost < 0 || salvage > cost || life <= 0 || factor <= 0) return CellError.NUM;
    if (start < 0 || end < start || end > life) return CellError.NUM;
    // The walk below is per period; a schedule this long is a mistake, not a
    // model, and must not be allowed to hang a recalculation.
    if (end > MAX_DEPRECIATION_PERIODS) return CellError.NUM;

    if (noSwitch) {
      const wholeStart = Math.floor(start);
      const wholeEnd = Math.ceil(end);
      let total = 0;
      for (let i = wholeStart + 1; i <= wholeEnd; i++) {
        let amount = ddbAt(cost, salvage, life, i, factor);
        if (i === wholeStart + 1) amount *= Math.min(end, wholeStart + 1) - start;
        else if (i === wholeEnd) amount *= end + 1 - wholeEnd;
        total += amount;
      }
      return finite(total);
    }

    let from = start;
    let to = end;
    // A fractional start already inside the straight-line half of the schedule
    // is shifted to the switch point, so the re-based second leg starts where
    // the method actually changes rather than mid-period.
    if (from !== Math.floor(from) && factor > 1 && from >= life / 2) {
      const shift = from - life / 2;
      from = life / 2;
      to -= shift;
    }
    const consumed = vdbSwitching(cost, salvage, life, life, from, factor);
    const rebased = excelSub(cost, consumed);
    return finite(vdbSwitching(rebased, salvage, life, life - from, to - from, factor));
  },
};

/**
 * French declining-balance depreciation.
 *
 * The coefficient table is statutory, not derived, and the per-period rounding
 * to whole currency units is part of the definition: dropping it produces
 * numbers that are close but never equal to the ones the tax authority expects.
 */
function amorCoefficient(usefulLife: number): number {
  if (usefulLife < 3) return 1;
  if (usefulLife < 5) return 1.5;
  if (usefulLife <= 6) return 2;
  return 2.5;
}

/** Round half away from zero, which is what Excel's own rounding does. */
function roundAway(v: number): number {
  return v < 0 ? -Math.round(-v) : Math.round(v);
}

interface AmorArgs {
  cost: number;
  purchased: Dt;
  firstPeriod: Dt;
  salvage: number;
  period: number;
  rate: number;
  basis: number;
}

function amorArgs(args: Value[], ctx: FunctionContext): AmorArgs | CellError {
  const cost = numArg(args[0]);
  if (isError(cost)) return cost;
  const purchased = dateArg(args[1]);
  if (isError(purchased)) return purchased;
  const firstPeriod = dateArg(args[2]);
  if (isError(firstPeriod)) return firstPeriod;
  const salvage = numArg(args[3]);
  if (isError(salvage)) return salvage;
  const period = intArg(args[4]);
  if (isError(period)) return period;
  const rate = numArg(args[5]);
  if (isError(rate)) return rate;
  const basis = basisArg(args[6]);
  if (isError(basis)) return basis;

  // The French schemes are not defined on actual/360, so Excel omits basis 2.
  if (basis === 2) return CellError.NUM;
  if (cost <= 0 || salvage < 0 || rate <= 0 || cost < salvage) return CellError.NUM;
  if (period < 0) return CellError.NUM;
  if (purchased > firstPeriod) return CellError.NUM;

  return {
    cost,
    purchased: dt(purchased, ctx.dateSystem),
    firstPeriod: dt(firstPeriod, ctx.dateSystem),
    salvage,
    period,
    rate,
    basis,
  };
}

const AMORDEGRC: FunctionSpec = {
  name: 'AMORDEGRC',
  params: [
    p.scalar('cost'),
    p.scalar('date_purchased'),
    p.scalar('first_period'),
    p.scalar('salvage'),
    p.scalar('period'),
    p.scalar('rate'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'French declining-balance depreciation with the statutory coefficient.',
  impl: (args, ctx) => {
    const a = amorArgs(args, ctx);
    if (isError(a)) return a;
    if (a.period > MAX_DEPRECIATION_PERIODS) return CellError.NUM;

    const rate = a.rate * amorCoefficient(1 / a.rate);
    let book = a.cost;
    let amount = roundAway(
      yearFraction(a.purchased, a.firstPeriod, a.basis) * rate * a.cost,
    );
    book -= amount;
    let remaining = excelSub(book, a.salvage);

    for (let n = 0; n < a.period; n++) {
      amount = roundAway(rate * book);
      remaining -= amount;
      if (remaining < 0) {
        // Past the point where the asset is written down to salvage the scheme
        // splits what is left over the final two periods and then stops.
        amount = a.period - n <= 1 ? roundAway(book * 0.5) : 0;
      }
      book -= amount;
    }
    return finite(amount);
  },
};

const AMORLINC: FunctionSpec = {
  name: 'AMORLINC',
  params: [
    p.scalar('cost'),
    p.scalar('date_purchased'),
    p.scalar('first_period'),
    p.scalar('salvage'),
    p.scalar('period'),
    p.scalar('rate'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'French straight-line depreciation, pro-rated over the first period.',
  impl: (args, ctx) => {
    const a = amorArgs(args, ctx);
    if (isError(a)) return a;

    const perPeriod = a.cost * a.rate;
    const depreciable = excelSub(a.cost, a.salvage);
    const firstAmount = yearFraction(a.purchased, a.firstPeriod, a.basis) * a.rate * a.cost;
    if (perPeriod === 0) return CellError.NUM;
    const fullPeriods = Math.trunc((depreciable - firstAmount) / perPeriod);

    if (a.period === 0) return finite(firstAmount);
    if (a.period <= fullPeriods) return finite(perPeriod);
    if (a.period === fullPeriods + 1) {
      return finite(depreciable - perPeriod * fullPeriods - firstAmount);
    }
    return 0;
  },
};

// ---------------------------------------------------------------------------
// Function specifications: rate conversions
// ---------------------------------------------------------------------------

const EFFECT: FunctionSpec = {
  name: 'EFFECT',
  params: [p.scalar('nominal_rate'), p.scalar('npery')],
  broadcast: true,
  summary: 'The effective annual rate for a nominal rate and compounding frequency.',
  impl: (args) => {
    const nominal = numArg(args[0]);
    if (isError(nominal)) return nominal;
    const periods = intArg(args[1]);
    if (isError(periods)) return periods;
    if (nominal <= 0 || periods < 1) return CellError.NUM;
    return finite(Math.pow(1 + nominal / periods, periods) - 1);
  },
};

const NOMINAL: FunctionSpec = {
  name: 'NOMINAL',
  params: [p.scalar('effect_rate'), p.scalar('npery')],
  broadcast: true,
  summary: 'The nominal annual rate for an effective rate and compounding frequency.',
  impl: (args) => {
    const effect = numArg(args[0]);
    if (isError(effect)) return effect;
    const periods = intArg(args[1]);
    if (isError(periods)) return periods;
    if (effect <= 0 || periods < 1) return CellError.NUM;
    return finite((Math.pow(effect + 1, 1 / periods) - 1) * periods);
  },
};

const PDURATION: FunctionSpec = {
  name: 'PDURATION',
  params: [p.scalar('rate'), p.scalar('pv'), p.scalar('fv')],
  broadcast: true,
  summary: 'The number of periods an investment needs to reach a given value.',
  impl: (args) => {
    const rate = numArg(args[0]);
    if (isError(rate)) return rate;
    const pv = numArg(args[1]);
    if (isError(pv)) return pv;
    const fv = numArg(args[2]);
    if (isError(fv)) return fv;
    if (rate <= 0 || pv <= 0 || fv <= 0) return CellError.NUM;
    return finite((Math.log(fv) - Math.log(pv)) / Math.log(1 + rate));
  },
};

const RRI: FunctionSpec = {
  name: 'RRI',
  params: [p.scalar('nper'), p.scalar('pv'), p.scalar('fv')],
  broadcast: true,
  summary: 'The equivalent interest rate implied by a growth in value.',
  impl: (args) => {
    const nper = numArg(args[0]);
    if (isError(nper)) return nper;
    const pv = numArg(args[1]);
    if (isError(pv)) return pv;
    const fv = numArg(args[2]);
    if (isError(fv)) return fv;
    if (nper <= 0 || pv === 0) return CellError.NUM;
    const ratio = fv / pv;
    // A fractional root of a negative ratio has no real value.
    if (ratio < 0) return CellError.NUM;
    return finite(Math.pow(ratio, 1 / nper) - 1);
  },
};

/**
 * The power of ten a fractional quotation is written against.
 *
 * A price quoted in sixteenths writes the numerator in the two digits after the
 * point, so 1.02 is one and two sixteenths - the scale comes from the number of
 * digits in the denominator, not from the denominator itself.
 */
function fractionScale(fraction: number): number {
  return Math.pow(10, Math.ceil(Math.log10(fraction)));
}

const DOLLARDE: FunctionSpec = {
  name: 'DOLLARDE',
  params: [p.scalar('fractional_dollar'), p.scalar('fraction')],
  broadcast: true,
  summary: 'Convert a price quoted as a fraction into a decimal price.',
  impl: (args) => {
    const quoted = numArg(args[0]);
    if (isError(quoted)) return quoted;
    const raw = numArg(args[1]);
    if (isError(raw)) return raw;
    if (raw < 0) return CellError.NUM;
    const fraction = Math.trunc(raw);
    if (fraction === 0) return CellError.DIV0;

    const whole = Math.trunc(quoted);
    const part = excelSub(quoted, whole);
    return finite(whole + (part * fractionScale(fraction)) / fraction);
  },
};

const DOLLARFR: FunctionSpec = {
  name: 'DOLLARFR',
  params: [p.scalar('decimal_dollar'), p.scalar('fraction')],
  broadcast: true,
  summary: 'Convert a decimal price into a price quoted as a fraction.',
  impl: (args) => {
    const decimal = numArg(args[0]);
    if (isError(decimal)) return decimal;
    const raw = numArg(args[1]);
    if (isError(raw)) return raw;
    if (raw < 0) return CellError.NUM;
    const fraction = Math.trunc(raw);
    if (fraction === 0) return CellError.DIV0;

    const whole = Math.trunc(decimal);
    const part = excelSub(decimal, whole);
    return finite(whole + (part * fraction) / fractionScale(fraction));
  },
};

// ---------------------------------------------------------------------------
// Function specifications: coupon dates
// ---------------------------------------------------------------------------

/** The settlement/maturity/frequency/basis preamble every COUP* function shares. */
function couponArgs(
  args: Value[],
  ctx: FunctionContext,
): CouponGeometry | CellError {
  const settle = dateArg(args[0]);
  if (isError(settle)) return settle;
  const maturity = dateArg(args[1]);
  if (isError(maturity)) return maturity;
  const frequency = frequencyArg(args[2]);
  if (isError(frequency)) return frequency;
  const basis = basisArg(args[3]);
  if (isError(basis)) return basis;
  if (settle >= maturity) return CellError.NUM;
  return couponGeometry(settle, maturity, frequency, basis, ctx.dateSystem);
}

const COUPON_PARAMS = [
  p.scalar('settlement'),
  p.scalar('maturity'),
  p.scalar('frequency'),
  p.scalar('basis', true),
];

const COUPDAYBS: FunctionSpec = {
  name: 'COUPDAYBS',
  params: COUPON_PARAMS,
  broadcast: true,
  summary: 'Days from the start of the coupon period to settlement.',
  impl: (args, ctx) => {
    const g = couponArgs(args, ctx);
    return isError(g) ? g : finite(g.sinceCoupon);
  },
};

const COUPDAYS: FunctionSpec = {
  name: 'COUPDAYS',
  params: COUPON_PARAMS,
  broadcast: true,
  summary: 'Days in the coupon period containing the settlement date.',
  impl: (args, ctx) => {
    const g = couponArgs(args, ctx);
    return isError(g) ? g : finite(g.periodDays);
  },
};

const COUPDAYSNC: FunctionSpec = {
  name: 'COUPDAYSNC',
  params: COUPON_PARAMS,
  broadcast: true,
  summary: 'Days from settlement to the next coupon date.',
  impl: (args, ctx) => {
    const g = couponArgs(args, ctx);
    return isError(g) ? g : finite(g.toCoupon);
  },
};

const COUPNCD: FunctionSpec = {
  name: 'COUPNCD',
  params: COUPON_PARAMS,
  broadcast: true,
  summary: 'The next coupon date after settlement.',
  impl: (args, ctx) => {
    const g = couponArgs(args, ctx);
    return isError(g) ? g : g.next.serial;
  },
};

const COUPPCD: FunctionSpec = {
  name: 'COUPPCD',
  params: COUPON_PARAMS,
  broadcast: true,
  summary: 'The coupon date on or before settlement.',
  impl: (args, ctx) => {
    const g = couponArgs(args, ctx);
    return isError(g) ? g : g.previous.serial;
  },
};

const COUPNUM: FunctionSpec = {
  name: 'COUPNUM',
  params: COUPON_PARAMS,
  broadcast: true,
  summary: 'The number of coupons payable between settlement and maturity.',
  impl: (args, ctx) => {
    const g = couponArgs(args, ctx);
    return isError(g) ? g : g.count;
  },
};

// ---------------------------------------------------------------------------
// Function specifications: coupon-bearing securities
// ---------------------------------------------------------------------------

interface BondArgs {
  geometry: CouponGeometry;
  frequency: number;
}

function bondArgs(
  settleArg: Value | undefined,
  maturityArg: Value | undefined,
  frequencyArg_: Value | undefined,
  basisArg_: Value | undefined,
  ctx: FunctionContext,
): BondArgs | CellError {
  const settle = dateArg(settleArg);
  if (isError(settle)) return settle;
  const maturity = dateArg(maturityArg);
  if (isError(maturity)) return maturity;
  const frequency = frequencyArg(frequencyArg_);
  if (isError(frequency)) return frequency;
  const basis = basisArg(basisArg_);
  if (isError(basis)) return basis;
  if (settle >= maturity) return CellError.NUM;
  const geometry = couponGeometry(settle, maturity, frequency, basis, ctx.dateSystem);
  if (isError(geometry)) return geometry;
  return { geometry, frequency };
}

const PRICE: FunctionSpec = {
  name: 'PRICE',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('rate'),
    p.scalar('yld'),
    p.scalar('redemption'),
    p.scalar('frequency'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The price per 100 face value of a security paying periodic interest.',
  impl: (args, ctx) => {
    const rate = numArg(args[2]);
    if (isError(rate)) return rate;
    const yld = numArg(args[3]);
    if (isError(yld)) return yld;
    const redemption = numArg(args[4]);
    if (isError(redemption)) return redemption;
    if (rate < 0 || yld < 0 || redemption <= 0) return CellError.NUM;
    const b = bondArgs(args[0], args[1], args[5], args[6], ctx);
    if (isError(b)) return b;
    return finite(bondPrice(b.geometry, rate, yld, redemption, b.frequency));
  },
};

/** YIELD inverts PRICE numerically; Excel documents a hundred attempts. */
const YIELD_BUDGET = 100;

const YIELD: FunctionSpec = {
  name: 'YIELD',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('rate'),
    p.scalar('pr'),
    p.scalar('redemption'),
    p.scalar('frequency'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The yield of a security paying periodic interest.',
  impl: (args, ctx) => {
    const rate = numArg(args[2]);
    if (isError(rate)) return rate;
    const price = numArg(args[3]);
    if (isError(price)) return price;
    const redemption = numArg(args[4]);
    if (isError(redemption)) return redemption;
    if (rate < 0 || price <= 0 || redemption <= 0) return CellError.NUM;
    const b = bondArgs(args[0], args[1], args[5], args[6], ctx);
    if (isError(b)) return b;
    const g = b.geometry;

    if (g.count <= 1) {
      // In the last coupon period the yield is a money-market rate and inverts
      // in closed form; iterating here would be slower and no more accurate.
      const remaining = g.periodDays - g.sinceCoupon;
      if (remaining === 0) return CellError.NUM;
      const carried = price / 100 + (g.sinceCoupon / g.periodDays) * (rate / b.frequency);
      if (carried === 0) return CellError.NUM;
      const received = redemption / 100 + rate / b.frequency;
      return finite(((received - carried) / carried) * b.frequency * (g.periodDays / remaining));
    }

    const root = secant(
      (y) => bondPrice(g, rate, y, redemption, b.frequency) - price,
      0.05,
      YIELD_BUDGET,
    );
    if (root === undefined) return CellError.NUM;
    return finite(root);
  },
};

/**
 * Macaulay duration, and the modified duration derived from it.
 *
 * The exponents are the same fractional coupon counts PRICE discounts by, which
 * is what makes a settlement date sitting exactly on a coupon date come out at
 * whole periods rather than an actual/actual approximation to them.
 */
function macaulay(
  g: CouponGeometry,
  coupon: number,
  yld: number,
  frequency: number,
): number | CellError {
  const step = 1 + yld / frequency;
  if (step <= 0) return CellError.NUM;
  const payment = (100 * coupon) / frequency;
  const offset = g.toCoupon / g.periodDays;

  let weighted = 0;
  let value = 0;
  for (let k = 1; k <= g.count; k++) {
    const time = k - 1 + offset;
    const cash = k === g.count ? payment + 100 : payment;
    const discounted = cash / Math.pow(step, time);
    weighted += time * discounted;
    value += discounted;
  }
  if (value === 0) return CellError.NUM;
  return weighted / value / frequency;
}

const DURATION: FunctionSpec = {
  name: 'DURATION',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('coupon'),
    p.scalar('yld'),
    p.scalar('frequency'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The Macaulay duration of a security paying periodic interest.',
  impl: (args, ctx) => {
    const coupon = numArg(args[2]);
    if (isError(coupon)) return coupon;
    const yld = numArg(args[3]);
    if (isError(yld)) return yld;
    if (coupon < 0 || yld < 0) return CellError.NUM;
    const b = bondArgs(args[0], args[1], args[4], args[5], ctx);
    if (isError(b)) return b;
    if (b.geometry.count > MAX_COUPON_PERIODS) return CellError.NUM;
    const d = macaulay(b.geometry, coupon, yld, b.frequency);
    return isError(d) ? d : finite(d);
  },
};

const MDURATION: FunctionSpec = {
  name: 'MDURATION',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('coupon'),
    p.scalar('yld'),
    p.scalar('frequency'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The modified duration of a security paying periodic interest.',
  impl: (args, ctx) => {
    const coupon = numArg(args[2]);
    if (isError(coupon)) return coupon;
    const yld = numArg(args[3]);
    if (isError(yld)) return yld;
    if (coupon < 0 || yld < 0) return CellError.NUM;
    const b = bondArgs(args[0], args[1], args[4], args[5], ctx);
    if (isError(b)) return b;
    if (b.geometry.count > MAX_COUPON_PERIODS) return CellError.NUM;
    const d = macaulay(b.geometry, coupon, yld, b.frequency);
    if (isError(d)) return d;
    return finite(d / (1 + yld / b.frequency));
  },
};

const ACCRINT: FunctionSpec = {
  name: 'ACCRINT',
  params: [
    p.scalar('issue'),
    p.scalar('first_interest'),
    p.scalar('settlement'),
    p.scalar('rate'),
    p.scalar('par'),
    p.scalar('frequency'),
    p.scalar('basis', true),
    p.scalar('calc_method', true),
  ],
  broadcast: true,
  summary: 'Accrued interest for a security that pays periodic interest.',
  impl: (args, ctx) => {
    const issue = dateArg(args[0]);
    if (isError(issue)) return issue;
    const firstInterest = dateArg(args[1]);
    if (isError(firstInterest)) return firstInterest;
    const settle = dateArg(args[2]);
    if (isError(settle)) return settle;
    const rate = numArg(args[3]);
    if (isError(rate)) return rate;
    const par = numArg(args[4], 1000);
    if (isError(par)) return par;
    const frequency = frequencyArg(args[5]);
    if (isError(frequency)) return frequency;
    const basis = basisArg(args[6]);
    if (isError(basis)) return basis;
    const fromIssue = flagArg(args[7], true);
    if (isError(fromIssue)) return fromIssue;

    if (rate <= 0 || par <= 0) return CellError.NUM;
    if (issue >= settle) return CellError.NUM;

    // calc_method only bites once settlement has passed the first coupon: FALSE
    // then accrues from that coupon rather than from issue.
    const start = fromIssue || settle <= firstInterest ? issue : firstInterest;
    const system = ctx.dateSystem;
    const schedule = scheduleOf(dt(firstInterest, system), frequency, system);

    // Walk the quasi-coupon periods the accrual spans. For the linear bases
    // this collapses to one term; for actual/actual it must not, because each
    // period has a different normal length.
    let accrued = 0;
    let k = periodsToMaturity(schedule, start);
    if (isError(k)) {
      // The accrual starts on or after the first coupon date, so the schedule
      // has to run forward from it rather than back.
      k = 0;
    }
    for (let guard = 0; guard <= MAX_COUPON_PERIODS; guard++) {
      const from = couponDate(schedule, k);
      if (isError(from)) return from;
      const to = couponDate(schedule, k - 1);
      if (isError(to)) return to;
      if (from.serial >= settle) break;
      if (to.serial > start) {
        const a = from.serial > start ? from : dt(start, system);
        const b = to.serial < settle ? to : dt(settle, system);
        const normal =
          basis === 1
            ? excelSub(to.serial, from.serial)
            : yearLength(from, to, basis) / frequency;
        if (normal > 0) accrued += dayCount(a, b, basis) / normal;
      }
      if (to.serial >= settle) break;
      k--;
      if (k < -MAX_COUPON_PERIODS) return CellError.NUM;
    }
    return finite((par * rate * accrued) / frequency);
  },
};

const ACCRINTM: FunctionSpec = {
  name: 'ACCRINTM',
  params: [
    p.scalar('issue'),
    p.scalar('settlement'),
    p.scalar('rate'),
    p.scalar('par'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'Accrued interest for a security that pays interest at maturity.',
  impl: (args, ctx) => {
    const issue = dateArg(args[0]);
    if (isError(issue)) return issue;
    const settle = dateArg(args[1]);
    if (isError(settle)) return settle;
    const rate = numArg(args[2]);
    if (isError(rate)) return rate;
    const par = numArg(args[3], 1000);
    if (isError(par)) return par;
    const basis = basisArg(args[4]);
    if (isError(basis)) return basis;

    if (rate <= 0 || par <= 0) return CellError.NUM;
    if (issue >= settle) return CellError.NUM;
    const system = ctx.dateSystem;
    return finite(par * rate * yearFraction(dt(issue, system), dt(settle, system), basis));
  },
};

const PRICEMAT: FunctionSpec = {
  name: 'PRICEMAT',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('issue'),
    p.scalar('rate'),
    p.scalar('yld'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The price per 100 face value of a security paying interest at maturity.',
  impl: (args, ctx) => {
    const m = maturityInstrument(args, ctx);
    if (isError(m)) return m;
    const yld = numArg(args[4]);
    if (isError(yld)) return yld;
    if (yld < 0) return CellError.NUM;

    const denominator = 1 + m.toMaturity * yld;
    if (denominator === 0) return CellError.DIV0;
    return finite(
      (100 + m.sinceIssueToMaturity * m.rate * 100) / denominator -
        m.accrued * m.rate * 100,
    );
  },
};

const YIELDMAT: FunctionSpec = {
  name: 'YIELDMAT',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('issue'),
    p.scalar('rate'),
    p.scalar('pr'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The annual yield of a security that pays interest at maturity.',
  impl: (args, ctx) => {
    const m = maturityInstrument(args, ctx);
    if (isError(m)) return m;
    const price = numArg(args[4]);
    if (isError(price)) return price;
    if (price <= 0) return CellError.NUM;

    const carried = price / 100 + m.accrued * m.rate;
    if (carried === 0 || m.toMaturity === 0) return CellError.DIV0;
    return finite(((1 + m.sinceIssueToMaturity * m.rate) / carried - 1) / m.toMaturity);
  },
};

/** The three year-fractions PRICEMAT and YIELDMAT are both written in terms of. */
function maturityInstrument(
  args: Value[],
  ctx: FunctionContext,
): {
  rate: number;
  /** Issue to maturity, DIM/B. */
  sinceIssueToMaturity: number;
  /** Issue to settlement, A/B. */
  accrued: number;
  /** Settlement to maturity, DSM/B. */
  toMaturity: number;
} | CellError {
  const settle = dateArg(args[0]);
  if (isError(settle)) return settle;
  const maturity = dateArg(args[1]);
  if (isError(maturity)) return maturity;
  const issue = dateArg(args[2]);
  if (isError(issue)) return issue;
  const rate = numArg(args[3]);
  if (isError(rate)) return rate;
  const basis = basisArg(args[5]);
  if (isError(basis)) return basis;

  if (rate < 0) return CellError.NUM;
  if (settle >= maturity || issue > settle) return CellError.NUM;

  const system = ctx.dateSystem;
  const i = dt(issue, system);
  const s = dt(settle, system);
  const m = dt(maturity, system);
  return {
    rate,
    sinceIssueToMaturity: yearFraction(i, m, basis),
    accrued: yearFraction(i, s, basis),
    toMaturity: yearFraction(s, m, basis),
  };
}

// ---------------------------------------------------------------------------
// Function specifications: discounted securities
// ---------------------------------------------------------------------------

const DISC: FunctionSpec = {
  name: 'DISC',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('pr'),
    p.scalar('redemption'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The discount rate of a security sold below its redemption value.',
  impl: (args, ctx) => {
    const price = numArg(args[2]);
    if (isError(price)) return price;
    const redemption = numArg(args[3]);
    if (isError(redemption)) return redemption;
    const basis = basisArg(args[4]);
    if (isError(basis)) return basis;
    const settle = dateArg(args[0]);
    if (isError(settle)) return settle;
    const maturity = dateArg(args[1]);
    if (isError(maturity)) return maturity;
    if (price <= 0 || redemption <= 0) return CellError.NUM;
    const span = spanOf(settle, maturity, basis, ctx.dateSystem);
    if (isError(span)) return span;
    const fraction = spanFraction(span);
    if (fraction === 0) return CellError.DIV0;
    return finite(excelSub(redemption, price) / redemption / fraction);
  },
};

const INTRATE: FunctionSpec = {
  name: 'INTRATE',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('investment'),
    p.scalar('redemption'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The interest rate of a fully invested security.',
  impl: (args, ctx) => {
    const investment = numArg(args[2]);
    if (isError(investment)) return investment;
    const redemption = numArg(args[3]);
    if (isError(redemption)) return redemption;
    const basis = basisArg(args[4]);
    if (isError(basis)) return basis;
    const settle = dateArg(args[0]);
    if (isError(settle)) return settle;
    const maturity = dateArg(args[1]);
    if (isError(maturity)) return maturity;
    if (investment <= 0 || redemption <= 0) return CellError.NUM;
    const span = spanOf(settle, maturity, basis, ctx.dateSystem);
    if (isError(span)) return span;
    const fraction = spanFraction(span);
    if (fraction === 0) return CellError.DIV0;
    return finite(excelSub(redemption, investment) / investment / fraction);
  },
};

const RECEIVED: FunctionSpec = {
  name: 'RECEIVED',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('investment'),
    p.scalar('discount'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The amount received at maturity for a fully invested security.',
  impl: (args, ctx) => {
    const investment = numArg(args[2]);
    if (isError(investment)) return investment;
    const discount = numArg(args[3]);
    if (isError(discount)) return discount;
    const basis = basisArg(args[4]);
    if (isError(basis)) return basis;
    const settle = dateArg(args[0]);
    if (isError(settle)) return settle;
    const maturity = dateArg(args[1]);
    if (isError(maturity)) return maturity;
    if (investment <= 0 || discount <= 0) return CellError.NUM;
    const span = spanOf(settle, maturity, basis, ctx.dateSystem);
    if (isError(span)) return span;
    const denominator = excelSub(1, discount * spanFraction(span));
    if (denominator === 0) return CellError.DIV0;
    return finite(investment / denominator);
  },
};

const PRICEDISC: FunctionSpec = {
  name: 'PRICEDISC',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('discount'),
    p.scalar('redemption'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The price per 100 face value of a discounted security.',
  impl: (args, ctx) => {
    const discount = numArg(args[2]);
    if (isError(discount)) return discount;
    const redemption = numArg(args[3]);
    if (isError(redemption)) return redemption;
    const basis = basisArg(args[4]);
    if (isError(basis)) return basis;
    const settle = dateArg(args[0]);
    if (isError(settle)) return settle;
    const maturity = dateArg(args[1]);
    if (isError(maturity)) return maturity;
    if (discount <= 0 || redemption <= 0) return CellError.NUM;
    const span = spanOf(settle, maturity, basis, ctx.dateSystem);
    if (isError(span)) return span;
    return finite(excelSub(redemption, discount * redemption * spanFraction(span)));
  },
};

const YIELDDISC: FunctionSpec = {
  name: 'YIELDDISC',
  params: [
    p.scalar('settlement'),
    p.scalar('maturity'),
    p.scalar('pr'),
    p.scalar('redemption'),
    p.scalar('basis', true),
  ],
  broadcast: true,
  summary: 'The annual yield of a discounted security.',
  impl: (args, ctx) => {
    const price = numArg(args[2]);
    if (isError(price)) return price;
    const redemption = numArg(args[3]);
    if (isError(redemption)) return redemption;
    const basis = basisArg(args[4]);
    if (isError(basis)) return basis;
    const settle = dateArg(args[0]);
    if (isError(settle)) return settle;
    const maturity = dateArg(args[1]);
    if (isError(maturity)) return maturity;
    if (price <= 0 || redemption <= 0) return CellError.NUM;
    const span = spanOf(settle, maturity, basis, ctx.dateSystem);
    if (isError(span)) return span;
    const fraction = spanFraction(span);
    if (fraction === 0) return CellError.DIV0;
    return finite(excelSub(redemption, price) / price / fraction);
  },
};

// ---------------------------------------------------------------------------
// Function specifications: Treasury bills
// ---------------------------------------------------------------------------

/**
 * A Treasury bill's term, in days.
 *
 * These three functions ignore the day-count basis entirely - the market
 * convention is actual days over a 360-day year - and Excel refuses a bill
 * maturing more than a year out, which is what makes the flat 365-day
 * comparison below correct rather than sloppy.
 */
function billDays(settleArg: Value | undefined, maturityArg: Value | undefined): number | CellError {
  const settle = dateArg(settleArg);
  if (isError(settle)) return settle;
  const maturity = dateArg(maturityArg);
  if (isError(maturity)) return maturity;
  if (settle >= maturity) return CellError.NUM;
  const days = excelSub(maturity, settle);
  if (days > 365) return CellError.NUM;
  return days;
}

const TBILLEQ: FunctionSpec = {
  name: 'TBILLEQ',
  params: [p.scalar('settlement'), p.scalar('maturity'), p.scalar('discount')],
  broadcast: true,
  summary: 'The bond-equivalent yield of a Treasury bill.',
  impl: (args) => {
    const days = billDays(args[0], args[1]);
    if (isError(days)) return days;
    const discount = numArg(args[2]);
    if (isError(discount)) return discount;
    if (discount <= 0) return CellError.NUM;
    const denominator = excelSub(360, discount * days);
    if (denominator === 0) return CellError.DIV0;
    return finite((365 * discount) / denominator);
  },
};

const TBILLPRICE: FunctionSpec = {
  name: 'TBILLPRICE',
  params: [p.scalar('settlement'), p.scalar('maturity'), p.scalar('discount')],
  broadcast: true,
  summary: 'The price per 100 face value of a Treasury bill.',
  impl: (args) => {
    const days = billDays(args[0], args[1]);
    if (isError(days)) return days;
    const discount = numArg(args[2]);
    if (isError(discount)) return discount;
    if (discount <= 0) return CellError.NUM;
    return finite(100 * excelSub(1, (discount * days) / 360));
  },
};

const TBILLYIELD: FunctionSpec = {
  name: 'TBILLYIELD',
  params: [p.scalar('settlement'), p.scalar('maturity'), p.scalar('pr')],
  broadcast: true,
  summary: 'The yield of a Treasury bill.',
  impl: (args) => {
    const days = billDays(args[0], args[1]);
    if (isError(days)) return days;
    const price = numArg(args[2]);
    if (isError(price)) return price;
    if (price <= 0) return CellError.NUM;
    return finite((excelSub(100, price) / price) * (360 / days));
  },
};

export const FINANCIAL_FUNCTIONS: readonly FunctionSpec[] = [
  PMT,
  IPMT,
  PPMT,
  PV,
  FV,
  FVSCHEDULE,
  RATE,
  NPER,
  NPV,
  XNPV,
  IRR,
  XIRR,
  MIRR,
  CUMIPMT,
  CUMPRINC,
  SLN,
  SYD,
  DB,
  DDB,
  VDB,
  EFFECT,
  NOMINAL,
  PDURATION,
  RRI,
  ISPMT,
  DOLLARDE,
  DOLLARFR,
  PRICE,
  PRICEDISC,
  PRICEMAT,
  YIELD,
  YIELDDISC,
  YIELDMAT,
  DISC,
  INTRATE,
  RECEIVED,
  DURATION,
  MDURATION,
  ACCRINT,
  ACCRINTM,
  COUPDAYBS,
  COUPDAYS,
  COUPDAYSNC,
  COUPNCD,
  COUPNUM,
  COUPPCD,
  TBILLEQ,
  TBILLPRICE,
  TBILLYIELD,
  AMORDEGRC,
  AMORLINC,
];
