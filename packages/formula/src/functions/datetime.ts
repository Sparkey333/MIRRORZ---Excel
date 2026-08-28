/**
 * Date and time.
 *
 * Six decisions shape this module.
 *
 * First, no JavaScript `Date` is constructed or read here. Every conversion
 * goes through packages/core/src/serial.ts, which works entirely in UTC and
 * reproduces the 1900 system's phantom 29 February 1900. A serial routed
 * through the host `Date` would render the same workbook differently in
 * different timezones and across a DST boundary, and would be off by one day
 * for every date before 1 March 1900.
 *
 * Second, `ctx.dateSystem` is threaded into every conversion rather than left
 * to the serial helpers' 1900 default. A 1904 workbook is 1462 days out
 * otherwise, which is silent and total corruption rather than a visible error.
 *
 * Third, NOW and TODAY read `ctx.now`. Reading the clock inside the
 * implementation would let two NOW() cells in one recalculation disagree, and
 * would make every test that touches them unrepeatable.
 *
 * Fourth, `partsToSerial` folds years 0-1899 into the twentieth century. That
 * is exactly DATE's documented rule, and exactly wrong for a year that
 * arithmetic produced: EDATE stepping back from January 1900 would silently
 * land in 3799 rather than reporting #NUM!. `serialFromYmd` rejects anything
 * below the workbook's epoch before that rewrite can happen, and every function
 * that computes a year uses it.
 *
 * Fifth, the two 30/360 conventions are genuinely different and are implemented
 * separately. Both fold the last day of February onto the 30th, but DAYS360's
 * US method folds only the start date, while YEARFRAC's NASD basis also folds
 * the end date when the start was folded. So DAYS360(28 Feb 2015, 29 Feb 2016)
 * is 359 while YEARFRAC over the same interval on basis 0 is exactly 1.
 *
 * Sixth, the working-day walkers jump whole weeks whenever no holiday falls
 * inside the jump. WORKDAY(date, 100000) is a legitimate formula and stepping
 * a day at a time would make it visibly slow.
 */

import {
  CellError,
  type DateSystem,
  type Scalar,
  daysInMonth,
  isError,
  isLeapYear,
  partsToSerial,
  serialToParts,
  weekdayFromSerial,
} from '@mirrorz/core';
import { type FunctionContext, type FunctionSpec, p } from '../registry.js';
import {
  type Value,
  excelSub,
  isArray,
  isRef,
  parseNumericText,
  toBoolean,
  toNumber,
  toText,
} from '../value.js';

// ---------------------------------------------------------------------------
// Serial plumbing
// ---------------------------------------------------------------------------

/** 9999-12-31, the last date Excel will hold, in each system. */
const MAX_SERIAL: Readonly<Record<DateSystem, number>> = { 1900: 2_958_465, 1904: 2_957_003 };

const SECONDS_PER_DAY = 86_400;

function epochYear(system: DateSystem): number {
  return system === 1904 ? 1904 : 1900;
}

/** The scalar an ArgKind.Scalar parameter delivered, with omissions as blank. */
function scalarArg(v: Value | undefined): Scalar {
  if (v === undefined) return null;
  if (isArray(v)) return v.data[0] ?? null;
  if (isRef(v)) return null;
  return v;
}

/** Reject serials outside the range Excel will represent as a date. */
function checkSerial(n: number, system: DateSystem): number | CellError {
  if (!Number.isFinite(n)) return CellError.NUM;
  if (n < 0 || n >= MAX_SERIAL[system] + 1) return CellError.NUM;
  return n;
}

/**
 * Coerce a scalar to a date serial the way every date function's first argument
 * does: numbers pass through, blanks are 0, booleans coerce, and text is parsed
 * as a date, a time, or failing both as a plain number.
 */
function serialFromScalar(s: Scalar, ctx: FunctionContext): number | CellError {
  if (isError(s)) return s;
  const system = ctx.dateSystem;
  if (typeof s === 'string') {
    const parsed = parseDateTimeText(s, ctx);
    if (parsed !== undefined) return checkSerial(parsed.dateSerial + parsed.timeFraction, system);
    const n = parseNumericText(s);
    if (n === undefined) return CellError.VALUE;
    return checkSerial(n, system);
  }
  const n = toNumber(s);
  return isError(n) ? n : checkSerial(n, system);
}

function dateArg(v: Value | undefined, ctx: FunctionContext): number | CellError {
  return serialFromScalar(scalarArg(v), ctx);
}

/** A date argument reduced to whole days, which is what Excel's INT() of it is. */
function dayArg(v: Value | undefined, ctx: FunctionContext): number | CellError {
  const s = dateArg(v, ctx);
  return isError(s) ? s : Math.floor(s);
}

/** A count argument, truncated towards zero as Excel does before using it. */
function intArg(v: Value | undefined, whenOmitted?: number): number | CellError {
  if (v === undefined && whenOmitted !== undefined) return whenOmitted;
  const n = toNumber(scalarArg(v));
  if (isError(n)) return n;
  if (!Number.isFinite(n)) return CellError.NUM;
  return Math.trunc(n);
}

interface Ymd {
  year: number;
  month: number;
  day: number;
}

/**
 * Calendar parts of a serial's date component.
 *
 * Serial 0 is special-cased: Excel calls it "January 0, 1900", so YEAR is 1900,
 * MONTH is 1 and DAY is 0, where the epoch arithmetic alone would say
 * 31 December 1899.
 */
function ymd(serial: number, system: DateSystem): Ymd {
  const days = Math.floor(serial);
  if (system === 1900 && days === 0) return { year: 1900, month: 1, day: 0 };
  const parts = serialToParts(days, system);
  return { year: parts.year, month: parts.month, day: parts.day };
}

/**
 * Days in a month on Excel's calendar rather than the real one.
 *
 * In the 1900 system February 1900 has twenty-nine days, because serial 60 is
 * the phantom leap day. EOMONTH and EDATE must agree with that, or they land a
 * day away from the date Excel shows for the same formula.
 */
function excelDaysInMonth(year: number, month: number, system: DateSystem): number {
  if (system === 1900 && year === 1900 && month === 2) return 29;
  return daysInMonth(year, month);
}

/**
 * Days elapsed since the first of a month, added in serial space.
 *
 * Going through the first of the month and then counting days is what makes the
 * phantom 29 February 1900 fall out for free: serial arithmetic already knows
 * the day exists, while `partsToSerial(1900, 2, 29, ...)` would roll it forward
 * onto 1 March. It also gives Excel's roll-over for a day outside the month,
 * DATE(2024, 1, 0) and DATE(2024, 1, 32) alike, with no extra cases.
 */
function serialOfDayInMonth(
  year: number,
  month: number,
  day: number,
  system: DateSystem,
): number {
  return partsToSerial(year, month, 1, 0, 0, 0, system) + (day - 1);
}

/**
 * Build a serial from parts that arithmetic produced.
 *
 * Unlike DATE, a computed year is taken literally: `partsToSerial` would map
 * 1899 to 3799 under its two-digit-year rule, so years below the epoch are
 * rejected first.
 */
function serialFromYmd(
  year: number,
  month: number,
  day: number,
  system: DateSystem,
): number | CellError {
  if (year < epochYear(system) || year > 9999) return CellError.NUM;
  return checkSerial(serialOfDayInMonth(year, month, day, system), system);
}

/** Whole seconds since midnight, rounded, which is the precision Excel reports. */
function secondsOfDay(serial: number): number {
  const frac = serial - Math.floor(serial);
  return Math.round(frac * SECONDS_PER_DAY) % SECONDS_PER_DAY;
}

function isLastDayOfFebruary(d: Ymd, system: DateSystem): boolean {
  return d.month === 2 && d.day === excelDaysInMonth(d.year, 2, system);
}

// ---------------------------------------------------------------------------
// Date and time text
// ---------------------------------------------------------------------------

/**
 * The date and time halves are kept apart rather than pre-added, so TIMEVALUE
 * can hand back the exact fraction it parsed. Adding 0.5732638888888889 to
 * 45352 and subtracting the day again loses four digits, which is visible
 * against Excel.
 */
interface ParsedText {
  /** Whole days, or 0 when the text carried no date. */
  dateSerial: number;
  /** Time of day as a fraction of a day. */
  timeFraction: number;
  hasDate: boolean;
}

const MONTH_NAMES = [
  'january',
  'february',
  'march',
  'april',
  'may',
  'june',
  'july',
  'august',
  'september',
  'october',
  'november',
  'december',
];

/** A clock reading anywhere in the string, with an optional meridiem. */
const CLOCK = /(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?(?:\s*([ap])\.?m\.?)?/i;
/** "3 PM", which Excel accepts as a time even without a colon. */
const BARE_MERIDIEM = /(^|\s)(\d{1,2})\s*([ap])\.?m\.?(\s|$)/i;

const ISO_DATE = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/;
const NUMERIC_DATE = /^(\d{1,2})[-/](\d{1,2})(?:[-/](\d{1,4}))?$/;
const DAY_MONTH = /^(\d{1,2})[- ]([a-z]+)(?:[- ](\d{1,4}))?$/i;
const MONTH_DAY = /^([a-z]+)[- ](\d{1,2})(?:[- ](\d{1,4}))?$/i;
const MONTH_YEAR = /^([a-z]+)[- ](\d{4})$/i;

function monthFromName(name: string): number | undefined {
  const t = name.toLowerCase();
  // Three letters is the shortest unambiguous abbreviation Excel accepts, and
  // it also stops "m" from silently meaning March.
  if (t.length < 3) return undefined;
  for (let i = 0; i < 12; i++) {
    if (MONTH_NAMES[i]!.startsWith(t)) return i + 1;
  }
  return undefined;
}

/** Excel's two-digit-year window: 00-29 is this century, 30-99 the last. */
function normaliseYear(y: number, digits: number): number {
  if (digits > 2) return y;
  return y < 30 ? y + 2000 : y + 1900;
}

function dateSerialFromParts(
  year: number,
  month: number,
  day: number,
  system: DateSystem,
): number | undefined {
  if (month < 1 || month > 12) return undefined;
  if (day < 1 || day > excelDaysInMonth(year, month, system)) return undefined;
  if (year < epochYear(system) || year > 9999) return undefined;
  return serialOfDayInMonth(year, month, day, system);
}

function parseDatePart(text: string, system: DateSystem, defaultYear: number): number | undefined {
  let m = ISO_DATE.exec(text);
  if (m) return dateSerialFromParts(Number(m[1]), Number(m[2]), Number(m[3]), system);

  m = NUMERIC_DATE.exec(text);
  if (m) {
    const raw = m[3];
    const year = raw === undefined ? defaultYear : normaliseYear(Number(raw), raw.length);
    return dateSerialFromParts(year, Number(m[1]), Number(m[2]), system);
  }

  m = DAY_MONTH.exec(text);
  if (m) {
    const month = monthFromName(m[2]!);
    if (month === undefined) return undefined;
    const raw = m[3];
    const year = raw === undefined ? defaultYear : normaliseYear(Number(raw), raw.length);
    return dateSerialFromParts(year, month, Number(m[1]), system);
  }

  m = MONTH_DAY.exec(text);
  if (m) {
    const month = monthFromName(m[1]!);
    if (month === undefined) return undefined;
    const raw = m[3];
    const year = raw === undefined ? defaultYear : normaliseYear(Number(raw), raw.length);
    return dateSerialFromParts(year, month, Number(m[2]), system);
  }

  m = MONTH_YEAR.exec(text);
  if (m) {
    const month = monthFromName(m[1]!);
    if (month === undefined) return undefined;
    return dateSerialFromParts(Number(m[2]), month, 1, system);
  }

  return undefined;
}

function clockToDayFraction(
  hourText: string,
  minuteText: string | undefined,
  secondText: string | undefined,
  meridiem: string | undefined,
): number | undefined {
  let hour = Number(hourText);
  const minute = minuteText === undefined ? 0 : Number(minuteText);
  const second = secondText === undefined ? 0 : Number(secondText);
  if (minute > 59 || second >= 60) return undefined;
  if (meridiem !== undefined) {
    if (hour < 1 || hour > 12) return undefined;
    const pm = meridiem.toLowerCase() === 'p';
    if (pm && hour < 12) hour += 12;
    if (!pm && hour === 12) hour = 0;
  }
  return (hour * 3600 + minute * 60 + second) / SECONDS_PER_DAY;
}

/**
 * Parse the date and time text forms DATEVALUE, TIMEVALUE and the coercion of a
 * text argument accept.
 *
 * Deliberately en-US ordered (month before day for an all-numeric date) because
 * the workbook model carries no locale yet; the ISO form is tried first so that
 * the unambiguous spelling is never reinterpreted.
 */
function parseDateTimeText(raw: string, ctx: FunctionContext): ParsedText | undefined {
  const system = ctx.dateSystem;
  let text = raw.trim();
  if (text === '') return undefined;

  let timeFraction = 0;
  let hasTime = false;

  const clock = CLOCK.exec(text);
  if (clock) {
    const frac = clockToDayFraction(clock[1]!, clock[2], clock[3], clock[4]);
    if (frac === undefined) return undefined;
    timeFraction = frac;
    hasTime = true;
    text = `${text.slice(0, clock.index)} ${text.slice(clock.index + clock[0].length)}`;
  } else {
    const bare = BARE_MERIDIEM.exec(text);
    if (bare) {
      const frac = clockToDayFraction(bare[2]!, undefined, undefined, bare[3]);
      if (frac === undefined) return undefined;
      timeFraction = frac;
      hasTime = true;
      text = `${text.slice(0, bare.index)} ${text.slice(bare.index + bare[0].length)}`;
    }
  }

  text = text.replaceAll(',', ' ').replace(/\s+/g, ' ').trim();
  if (text === '') {
    return hasTime ? { dateSerial: 0, timeFraction, hasDate: false } : undefined;
  }

  const defaultYear = ymd(Math.floor(ctx.now), system).year;
  const date = parseDatePart(text, system, defaultYear);
  if (date === undefined) return undefined;
  return { dateSerial: date, timeFraction, hasDate: true };
}

// ---------------------------------------------------------------------------
// Weekday and week numbering
// ---------------------------------------------------------------------------

/** WEEKDAY's return_type -> the day the week starts on, Sunday = 0. */
const WEEK_START: Readonly<Record<number, number>> = {
  1: 0,
  2: 1,
  11: 1,
  12: 2,
  13: 3,
  14: 4,
  15: 5,
  16: 6,
  17: 0,
};

function isoWeekNumber(serial: number, system: DateSystem): number {
  const day = Math.floor(serial);
  // ISO 8601 numbers the week by the year its Thursday falls in.
  const isoDow = ((weekdayFromSerial(day, system) + 6) % 7) + 1;
  const thursday = day - isoDow + 4;
  const year = ymd(thursday, system).year;
  const jan1 = partsToSerial(year, 1, 1, 0, 0, 0, system);
  return Math.floor((thursday - jan1) / 7) + 1;
}

// ---------------------------------------------------------------------------
// Day-count conventions
// ---------------------------------------------------------------------------

function days360(from: Ymd, to: Ymd, european: boolean, system: DateSystem): number {
  let d1 = from.day;
  let d2 = to.day;
  let m2 = to.month;
  let y2 = to.year;

  if (d1 === 31) {
    d1 = 30;
  } else if (!european && isLastDayOfFebruary(from, system)) {
    // The US method folds the last day of February onto the 30th, which is why
    // Microsoft's own documentation warns that 28 February to 28 March returns
    // 28 rather than the 30 a "full month" would suggest. Only the start date
    // is folded: the end of February is left alone, which is where DAYS360 and
    // YEARFRAC's NASD basis part company.
    d1 = 30;
  }

  if (d2 === 31) {
    if (european || d1 === 30) {
      d2 = 30;
    } else {
      // "the ending date becomes the 1st of the next month".
      d2 = 1;
      m2 += 1;
      if (m2 > 12) {
        m2 = 1;
        y2 += 1;
      }
    }
  }

  return (y2 - from.year) * 360 + (m2 - from.month) * 30 + (d2 - d1);
}

/**
 * NASD 30/360, the convention behind YEARFRAC basis 0.
 *
 * The difference from DAYS360's US method is the second fold: the end date
 * moves to the 30th when it too is the last day of February, so a whole number
 * of years between two ends of February comes out exact.
 */
function days360Nasd(from: Ymd, to: Ymd, system: DateSystem): number {
  let d1 = from.day;
  let d2 = to.day;
  if (isLastDayOfFebruary(from, system) && isLastDayOfFebruary(to, system)) d2 = 30;
  if (isLastDayOfFebruary(from, system)) d1 = 30;
  if (d2 === 31 && d1 >= 30) d2 = 30;
  if (d1 === 31) d1 = 30;
  return (to.year - from.year) * 360 + (to.month - from.month) * 30 + (d2 - d1);
}

/**
 * The denominator YEARFRAC basis 1 divides by.
 *
 * Excel does not use a plain 365.25: for spans of a year or less it uses the
 * length of the year the span sits in, and for longer spans the mean length of
 * every calendar year the span touches, endpoints included.
 */
function actualYearLength(from: Ymd, to: Ymd, system: DateSystem): number {
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
  const start = partsToSerial(from.year, from.month, from.day, 0, 0, 0, system);
  const end = partsToSerial(to.year, to.month, to.day, 0, 0, 0, system);
  for (const y of [from.year, to.year]) {
    if (!isLeapYear(y)) continue;
    const leapDay = partsToSerial(y, 2, 29, 0, 0, 0, system);
    if (leapDay >= start && leapDay <= end) return 366;
  }
  return 365;
}

// ---------------------------------------------------------------------------
// Working days
// ---------------------------------------------------------------------------

/** WORKDAY.INTL / NETWORKDAYS.INTL weekend codes, as masks indexed from Monday. */
const WEEKEND_MASKS: Readonly<Record<number, string>> = {
  1: '0000011',
  2: '1000001',
  3: '1100000',
  4: '0110000',
  5: '0011000',
  6: '0001100',
  7: '0000110',
  11: '0000001',
  12: '1000000',
  13: '0100000',
  14: '0010000',
  15: '0001000',
  16: '0000100',
  17: '0000010',
};

const DEFAULT_WEEKEND = WEEKEND_MASKS[1]!;

function weekendMask(v: Value | undefined): string | CellError {
  if (v === undefined) return DEFAULT_WEEKEND;
  const s = scalarArg(v);
  if (isError(s)) return s;
  if (typeof s === 'string') {
    // A malformed string is #VALUE!, while a number outside the table is #NUM!.
    return /^[01]{7}$/.test(s) ? s : CellError.VALUE;
  }
  const n = toNumber(s);
  if (isError(n)) return n;
  return WEEKEND_MASKS[Math.trunc(n)] ?? CellError.NUM;
}

/** Day-of-week index with Monday = 0, the order the weekend mask is written in. */
function mondayIndex(serial: number, system: DateSystem): number {
  return (weekdayFromSerial(serial, system) + 6) % 7;
}

function isWeekend(serial: number, mask: string, system: DateSystem): boolean {
  return mask[mondayIndex(serial, system)] === '1';
}

function workingDaysPerWeek(mask: string): number {
  let n = 0;
  for (const ch of mask) if (ch === '0') n++;
  return n;
}

function holidaySet(v: Value | undefined, ctx: FunctionContext): Set<number> | CellError {
  const out = new Set<number>();
  if (v === undefined) return out;
  const add = (s: Scalar): CellError | undefined => {
    if (s === null) return undefined;
    const serial = serialFromScalar(s, ctx);
    if (isError(serial)) return serial;
    out.add(Math.floor(serial));
    return undefined;
  };
  if (isArray(v)) {
    for (const cell of v.data) {
      const bad = add(cell);
      if (bad) return bad;
    }
    return out;
  }
  if (isRef(v)) {
    for (const cell of ctx.iterate(v)) {
      const bad = add(cell.value);
      if (bad) return bad;
    }
    return out;
  }
  const bad = add(v as Scalar);
  return bad ?? out;
}

function countWorkdays(
  from: number,
  to: number,
  mask: string,
  holidays: Set<number>,
  system: DateSystem,
): number {
  const perWeek = workingDaysPerWeek(mask);
  if (perWeek === 0) return 0;

  const total = to - from + 1;
  const fullWeeks = Math.floor(total / 7);
  let count = fullWeeks * perWeek;
  for (let d = from + fullWeeks * 7; d <= to; d++) {
    if (!isWeekend(d, mask, system)) count++;
  }
  for (const h of holidays) {
    if (h >= from && h <= to && !isWeekend(h, mask, system)) count--;
  }
  return count;
}

function advanceWorkdays(
  start: number,
  days: number,
  mask: string,
  holidays: Set<number>,
  system: DateSystem,
): number {
  if (days === 0) return start;
  const perWeek = workingDaysPerWeek(mask);
  const step = days > 0 ? 1 : -1;
  let remaining = Math.abs(days);
  let cur = start;

  const holidayInSpan = (a: number, b: number): boolean => {
    const lo = Math.min(a, b);
    const hi = Math.max(a, b);
    for (const h of holidays) if (h >= lo && h <= hi) return true;
    return false;
  };

  while (remaining > 0) {
    // A whole-week jump keeps the weekday and so passes exactly perWeek working
    // days, provided no holiday sits inside the span it skips over. At least one
    // working day is always left for the day-at-a-time walk below, because the
    // result has to *land* on a working day: jumping the last week from a
    // Saturday start would return the following Saturday.
    const weeks = Math.floor((remaining - 1) / perWeek);
    if (weeks > 0 && !holidayInSpan(cur + step, cur + step * weeks * 7)) {
      cur += step * weeks * 7;
      remaining -= weeks * perWeek;
      continue;
    }
    cur += step;
    if (!isWeekend(cur, mask, system) && !holidays.has(cur)) remaining--;
  }
  return cur;
}

// ---------------------------------------------------------------------------
// Functions
// ---------------------------------------------------------------------------

const DATE: FunctionSpec = {
  name: 'DATE',
  params: [p.scalar('year'), p.scalar('month'), p.scalar('day')],
  broadcast: true,
  summary: 'The serial number of a date given its year, month and day.',
  impl: (args, ctx) => {
    const year = intArg(args[0]);
    if (isError(year)) return year;
    const month = intArg(args[1]);
    if (isError(month)) return month;
    const day = intArg(args[2]);
    if (isError(day)) return day;
    if (year < 0 || year > 9999) return CellError.NUM;
    // partsToSerial rolls an out-of-range month over and applies DATE's own
    // two-digit-year rule, both of which Excel wants; the day is then counted in
    // serial space so that DATE(1900, 2, 29) reaches the phantom leap day
    // instead of rolling forward onto 1 March.
    const shiftedYear = year >= 0 && year <= 1899 ? year + 1900 : year;
    const monthIndex = month - 1;
    const normalisedYear = shiftedYear + Math.floor(monthIndex / 12);
    const normalisedMonth = (((monthIndex % 12) + 12) % 12) + 1;
    // Below the epoch partsToSerial would re-apply its two-digit-year shift, and
    // the date is unrepresentable anyway.
    if (normalisedYear < epochYear(ctx.dateSystem) || normalisedYear > 9999) return CellError.NUM;
    return checkSerial(
      serialOfDayInMonth(normalisedYear, normalisedMonth, day, ctx.dateSystem),
      ctx.dateSystem,
    );
  },
};

const TIME: FunctionSpec = {
  name: 'TIME',
  params: [p.scalar('hour'), p.scalar('minute'), p.scalar('second')],
  broadcast: true,
  summary: 'The fraction of a day represented by a time.',
  impl: (args) => {
    const hour = intArg(args[0]);
    if (isError(hour)) return hour;
    const minute = intArg(args[1]);
    if (isError(minute)) return minute;
    const second = intArg(args[2]);
    if (isError(second)) return second;
    if (hour < 0 || minute < 0 || second < 0) return CellError.NUM;
    if (hour > 32767 || minute > 32767 || second > 32767) return CellError.NUM;
    // TIME keeps only the time of day: 27:00 is 03:00, not tomorrow morning.
    const seconds = (hour * 3600 + minute * 60 + second) % SECONDS_PER_DAY;
    return seconds / SECONDS_PER_DAY;
  },
};

const DATEVALUE: FunctionSpec = {
  name: 'DATEVALUE',
  params: [p.scalar('date_text')],
  broadcast: true,
  summary: 'The serial number of a date written as text.',
  impl: (args, ctx) => {
    const s = scalarArg(args[0]);
    if (isError(s)) return s;
    // A number is not text, and DATEVALUE refuses it rather than passing it on.
    if (typeof s !== 'string') return CellError.VALUE;
    const parsed = parseDateTimeText(s, ctx);
    if (parsed === undefined || !parsed.hasDate) return CellError.VALUE;
    return checkSerial(parsed.dateSerial, ctx.dateSystem);
  },
};

const TIMEVALUE: FunctionSpec = {
  name: 'TIMEVALUE',
  params: [p.scalar('time_text')],
  broadcast: true,
  summary: 'The fraction of a day of a time written as text.',
  impl: (args, ctx) => {
    const s = scalarArg(args[0]);
    if (isError(s)) return s;
    if (typeof s !== 'string') return CellError.VALUE;
    const parsed = parseDateTimeText(s, ctx);
    if (parsed === undefined) return CellError.VALUE;
    // Date information in the text is ignored rather than rejected.
    return parsed.timeFraction % 1;
  },
};

const NOW: FunctionSpec = {
  name: 'NOW',
  params: [],
  volatile: true,
  summary: 'The current date and time.',
  impl: (_args, ctx) => ctx.now,
};

const TODAY: FunctionSpec = {
  name: 'TODAY',
  params: [],
  volatile: true,
  summary: "Today's date, with no time of day.",
  impl: (_args, ctx) => Math.floor(ctx.now),
};

const YEAR: FunctionSpec = {
  name: 'YEAR',
  params: [p.scalar('serial_number')],
  broadcast: true,
  summary: 'The year of a date.',
  impl: (args, ctx) => {
    const s = dayArg(args[0], ctx);
    return isError(s) ? s : ymd(s, ctx.dateSystem).year;
  },
};

const MONTH: FunctionSpec = {
  name: 'MONTH',
  params: [p.scalar('serial_number')],
  broadcast: true,
  summary: 'The month of a date, 1 to 12.',
  impl: (args, ctx) => {
    const s = dayArg(args[0], ctx);
    return isError(s) ? s : ymd(s, ctx.dateSystem).month;
  },
};

const DAY: FunctionSpec = {
  name: 'DAY',
  params: [p.scalar('serial_number')],
  broadcast: true,
  summary: 'The day of the month of a date, 1 to 31.',
  impl: (args, ctx) => {
    const s = dayArg(args[0], ctx);
    return isError(s) ? s : ymd(s, ctx.dateSystem).day;
  },
};

const HOUR: FunctionSpec = {
  name: 'HOUR',
  params: [p.scalar('serial_number')],
  broadcast: true,
  summary: 'The hour of a time, 0 to 23.',
  impl: (args, ctx) => {
    const s = dateArg(args[0], ctx);
    return isError(s) ? s : Math.floor(secondsOfDay(s) / 3600);
  },
};

const MINUTE: FunctionSpec = {
  name: 'MINUTE',
  params: [p.scalar('serial_number')],
  broadcast: true,
  summary: 'The minute of a time, 0 to 59.',
  impl: (args, ctx) => {
    const s = dateArg(args[0], ctx);
    return isError(s) ? s : Math.floor((secondsOfDay(s) % 3600) / 60);
  },
};

const SECOND: FunctionSpec = {
  name: 'SECOND',
  params: [p.scalar('serial_number')],
  broadcast: true,
  summary: 'The second of a time, 0 to 59.',
  impl: (args, ctx) => {
    const s = dateArg(args[0], ctx);
    return isError(s) ? s : secondsOfDay(s) % 60;
  },
};

const WEEKDAY: FunctionSpec = {
  name: 'WEEKDAY',
  params: [p.scalar('serial_number'), p.scalar('return_type', true)],
  broadcast: true,
  summary: 'The day of the week of a date, numbered per return_type.',
  impl: (args, ctx) => {
    const s = dayArg(args[0], ctx);
    if (isError(s)) return s;
    const type = intArg(args[1], 1);
    if (isError(type)) return type;
    const dow = weekdayFromSerial(s, ctx.dateSystem);
    // Type 3 is the only variant numbered from zero.
    if (type === 3) return (dow + 6) % 7;
    const start = WEEK_START[type];
    if (start === undefined) return CellError.NUM;
    return ((dow - start + 7) % 7) + 1;
  },
};

const WEEKNUM: FunctionSpec = {
  name: 'WEEKNUM',
  params: [p.scalar('serial_number'), p.scalar('return_type', true)],
  broadcast: true,
  summary: 'The week of the year a date falls in.',
  impl: (args, ctx) => {
    const s = dayArg(args[0], ctx);
    if (isError(s)) return s;
    const type = intArg(args[1], 1);
    if (isError(type)) return type;
    const system = ctx.dateSystem;
    if (type === 21) return isoWeekNumber(s, system);
    const start = WEEK_START[type];
    if (start === undefined) return CellError.NUM;
    // Week 1 is whichever week contains 1 January, however short it is.
    const jan1 = partsToSerial(ymd(s, system).year, 1, 1, 0, 0, 0, system);
    const offset = (weekdayFromSerial(jan1, system) - start + 7) % 7;
    return Math.floor((s - jan1 + offset) / 7) + 1;
  },
};

const ISOWEEKNUM: FunctionSpec = {
  name: 'ISOWEEKNUM',
  params: [p.scalar('date')],
  broadcast: true,
  summary: 'The ISO 8601 week number of a date.',
  impl: (args, ctx) => {
    const s = dayArg(args[0], ctx);
    return isError(s) ? s : isoWeekNumber(s, ctx.dateSystem);
  },
};

const EDATE: FunctionSpec = {
  name: 'EDATE',
  params: [p.scalar('start_date'), p.scalar('months')],
  broadcast: true,
  summary: 'The date a whole number of months before or after another date.',
  impl: (args, ctx) => {
    const start = dayArg(args[0], ctx);
    if (isError(start)) return start;
    const months = intArg(args[1]);
    if (isError(months)) return months;
    const from = ymd(start, ctx.dateSystem);
    const total = from.year * 12 + (from.month - 1) + months;
    const year = Math.floor(total / 12);
    const month = (((total % 12) + 12) % 12) + 1;
    // The day of month is clamped, so a month after 31 January is 28 or 29
    // February rather than spilling into March.
    const day = Math.min(from.day, excelDaysInMonth(year, month, ctx.dateSystem));
    return serialFromYmd(year, month, day, ctx.dateSystem);
  },
};

const EOMONTH: FunctionSpec = {
  name: 'EOMONTH',
  params: [p.scalar('start_date'), p.scalar('months')],
  broadcast: true,
  summary: 'The last day of the month a whole number of months away.',
  impl: (args, ctx) => {
    const start = dayArg(args[0], ctx);
    if (isError(start)) return start;
    const months = intArg(args[1]);
    if (isError(months)) return months;
    const from = ymd(start, ctx.dateSystem);
    const total = from.year * 12 + (from.month - 1) + months;
    const year = Math.floor(total / 12);
    const month = (((total % 12) + 12) % 12) + 1;
    return serialFromYmd(year, month, excelDaysInMonth(year, month, ctx.dateSystem), ctx.dateSystem);
  },
};

const DATEDIF: FunctionSpec = {
  name: 'DATEDIF',
  params: [p.scalar('start_date'), p.scalar('end_date'), p.scalar('unit')],
  broadcast: true,
  summary: 'The difference between two dates in years, months or days.',
  impl: (args, ctx) => {
    const start = dayArg(args[0], ctx);
    if (isError(start)) return start;
    const end = dayArg(args[1], ctx);
    if (isError(end)) return end;
    const unitText = toText(scalarArg(args[2]));
    if (isError(unitText)) return unitText;
    // Unlike every other date function, DATEDIF refuses a reversed interval.
    if (start > end) return CellError.NUM;

    const system = ctx.dateSystem;
    const a = ymd(start, system);
    const b = ymd(end, system);
    const dayShort = b.day < a.day;

    switch (unitText.trim().toUpperCase()) {
      case 'D':
        return excelSub(end, start);
      case 'Y': {
        const monthShort = b.month < a.month || (b.month === a.month && dayShort);
        return b.year - a.year - (monthShort ? 1 : 0);
      }
      case 'M':
        return (b.year - a.year) * 12 + (b.month - a.month) - (dayShort ? 1 : 0);
      case 'YM': {
        const months = (b.month - a.month) - (dayShort ? 1 : 0);
        return ((months % 12) + 12) % 12;
      }
      case 'MD': {
        // Excel subtracts the days and, when that goes negative, adds the length
        // of the month *before* the end date. The result can therefore be
        // negative - a long-standing Excel bug that files depend on.
        if (!dayShort) return b.day - a.day;
        const prevMonth = b.month === 1 ? 12 : b.month - 1;
        const prevYear = b.month === 1 ? b.year - 1 : b.year;
        return b.day - a.day + excelDaysInMonth(prevYear, prevMonth, system);
      }
      case 'YD': {
        const sameYear = b.month > a.month || (b.month === a.month && b.day >= a.day);
        const anchorYear = sameYear ? b.year : b.year - 1;
        const anchor = serialFromYmd(anchorYear, a.month, a.day, system);
        if (isError(anchor)) return anchor;
        return excelSub(end, anchor);
      }
      default:
        return CellError.NUM;
    }
  },
};

const DAYS: FunctionSpec = {
  name: 'DAYS',
  params: [p.scalar('end_date'), p.scalar('start_date')],
  broadcast: true,
  summary: 'The number of days between two dates.',
  impl: (args, ctx) => {
    const end = dayArg(args[0], ctx);
    if (isError(end)) return end;
    const start = dayArg(args[1], ctx);
    if (isError(start)) return start;
    return excelSub(end, start);
  },
};

const DAYS360: FunctionSpec = {
  name: 'DAYS360',
  params: [p.scalar('start_date'), p.scalar('end_date'), p.scalar('method', true)],
  broadcast: true,
  summary: 'The days between two dates on a 360-day year.',
  impl: (args, ctx) => {
    const start = dayArg(args[0], ctx);
    if (isError(start)) return start;
    const end = dayArg(args[1], ctx);
    if (isError(end)) return end;
    const european = args[2] === undefined ? false : toBoolean(scalarArg(args[2]));
    if (isError(european)) return european;
    const system = ctx.dateSystem;
    // A reversed interval is negated rather than rejected.
    if (start > end) return -days360(ymd(end, system), ymd(start, system), european, system);
    return days360(ymd(start, system), ymd(end, system), european, system);
  },
};

const YEARFRAC: FunctionSpec = {
  name: 'YEARFRAC',
  params: [p.scalar('start_date'), p.scalar('end_date'), p.scalar('basis', true)],
  broadcast: true,
  summary: 'The fraction of a year between two dates on a given day-count basis.',
  impl: (args, ctx) => {
    const first = dayArg(args[0], ctx);
    if (isError(first)) return first;
    const second = dayArg(args[1], ctx);
    if (isError(second)) return second;
    const basis = intArg(args[2], 0);
    if (isError(basis)) return basis;
    if (basis < 0 || basis > 4) return CellError.NUM;

    // YEARFRAC is symmetric; only the magnitude of the interval matters.
    const start = Math.min(first, second);
    const end = Math.max(first, second);
    const system = ctx.dateSystem;
    const a = ymd(start, system);
    const b = ymd(end, system);

    switch (basis) {
      case 0:
        return days360Nasd(a, b, system) / 360;
      case 1:
        return excelSub(end, start) / actualYearLength(a, b, system);
      case 2:
        return excelSub(end, start) / 360;
      case 3:
        return excelSub(end, start) / 365;
      default:
        return days360(a, b, true, system) / 360;
    }
  },
};

function workdayImpl(
  args: Value[],
  ctx: FunctionContext,
  maskArg: Value | undefined,
  holidayArg: Value | undefined,
): Value {
  const start = dayArg(args[0], ctx);
  if (isError(start)) return start;
  const days = intArg(args[1]);
  if (isError(days)) return days;
  const mask = weekendMask(maskArg);
  if (isError(mask)) return mask;
  // With no working days at all the walk could never terminate.
  if (workingDaysPerWeek(mask) === 0) return CellError.VALUE;
  const holidays = holidaySet(holidayArg, ctx);
  if (isError(holidays)) return holidays;
  return checkSerial(advanceWorkdays(start, days, mask, holidays, ctx.dateSystem), ctx.dateSystem);
}

function networkdaysImpl(
  args: Value[],
  ctx: FunctionContext,
  maskArg: Value | undefined,
  holidayArg: Value | undefined,
): Value {
  const start = dayArg(args[0], ctx);
  if (isError(start)) return start;
  const end = dayArg(args[1], ctx);
  if (isError(end)) return end;
  const mask = weekendMask(maskArg);
  if (isError(mask)) return mask;
  const holidays = holidaySet(holidayArg, ctx);
  if (isError(holidays)) return holidays;
  const system = ctx.dateSystem;
  // Both endpoints are counted, and a reversed interval counts negative.
  if (start > end) return -countWorkdays(end, start, mask, holidays, system);
  return countWorkdays(start, end, mask, holidays, system);
}

const WORKDAY: FunctionSpec = {
  name: 'WORKDAY',
  params: [p.scalar('start_date'), p.scalar('days'), p.array('holidays', true)],
  summary: 'The date a number of working days away, skipping weekends and holidays.',
  impl: (args, ctx) => workdayImpl(args, ctx, undefined, args[2]),
};

const WORKDAY_INTL: FunctionSpec = {
  name: 'WORKDAY.INTL',
  params: [
    p.scalar('start_date'),
    p.scalar('days'),
    p.scalar('weekend', true),
    p.array('holidays', true),
  ],
  summary: 'WORKDAY with a configurable weekend.',
  impl: (args, ctx) => workdayImpl(args, ctx, args[2], args[3]),
};

const NETWORKDAYS: FunctionSpec = {
  name: 'NETWORKDAYS',
  params: [p.scalar('start_date'), p.scalar('end_date'), p.array('holidays', true)],
  summary: 'Whole working days between two dates, both ends included.',
  impl: (args, ctx) => networkdaysImpl(args, ctx, undefined, args[2]),
};

const NETWORKDAYS_INTL: FunctionSpec = {
  name: 'NETWORKDAYS.INTL',
  params: [
    p.scalar('start_date'),
    p.scalar('end_date'),
    p.scalar('weekend', true),
    p.array('holidays', true),
  ],
  summary: 'NETWORKDAYS with a configurable weekend.',
  impl: (args, ctx) => networkdaysImpl(args, ctx, args[2], args[3]),
};

export const DATETIME_FUNCTIONS: readonly FunctionSpec[] = [
  DATE,
  TIME,
  DATEVALUE,
  TIMEVALUE,
  NOW,
  TODAY,
  YEAR,
  MONTH,
  DAY,
  HOUR,
  MINUTE,
  SECOND,
  WEEKDAY,
  WEEKNUM,
  ISOWEEKNUM,
  EDATE,
  EOMONTH,
  DATEDIF,
  DAYS,
  DAYS360,
  YEARFRAC,
  WORKDAY,
  WORKDAY_INTL,
  NETWORKDAYS,
  NETWORKDAYS_INTL,
];
