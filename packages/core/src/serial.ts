/**
 * Excel date/time serial numbers.
 *
 * A serial is "days since the epoch", with the fractional part being the time of
 * day. Two epochs exist: the 1900 system (Windows default) and the 1904 system
 * (originally Mac). A workbook declares which it uses in workbook.xml.
 *
 * The 1900 system carries a deliberate, permanent bug: Excel treats 1900 as a
 * leap year, so serial 60 is the non-existent "29 Feb 1900". Lotus 1-2-3 had the
 * bug, Excel copied it for file compatibility in 1985, and it can never be fixed
 * because a billion workbooks encode dates that way. We reproduce it exactly -
 * any implementation that "corrects" it is off by one day for every date before
 * 1 March 1900, and, worse, disagrees with Excel about DATEDIF and NETWORKDAYS
 * results near that boundary.
 */

export type DateSystem = 1900 | 1904;

const MS_PER_DAY = 86_400_000;

/** UTC ms for 1899-12-31, the day *before* 1900 serial 1. */
const EPOCH_1900_UTC = Date.UTC(1899, 11, 31);
/** UTC ms for 1904-01-01, which is serial 0 in the 1904 system. */
const EPOCH_1904_UTC = Date.UTC(1904, 0, 1);

/** The phantom 29 Feb 1900 in the 1900 system. */
export const PHANTOM_LEAP_SERIAL = 60;

/**
 * Convert a serial number to a UTC `Date`.
 *
 * Dates are handled in UTC throughout. Spreadsheet dates are wall-clock values
 * with no timezone, so introducing the host timezone anywhere would make the
 * same file render differently in different places - a bug users have reported
 * against many spreadsheet libraries.
 */
export function serialToDate(serial: number, system: DateSystem = 1900): Date {
  return new Date(serialToUtcMs(serial, system));
}

export function serialToUtcMs(serial: number, system: DateSystem = 1900): number {
  if (system === 1904) {
    return EPOCH_1904_UTC + serial * MS_PER_DAY;
  }
  // Serials >= 61 are shifted one day earlier to skip the phantom 29 Feb 1900.
  const adjusted = serial < PHANTOM_LEAP_SERIAL ? serial : serial - 1;
  return EPOCH_1900_UTC + adjusted * MS_PER_DAY;
}

/** Convert a UTC `Date` to a serial number. */
export function dateToSerial(date: Date, system: DateSystem = 1900): number {
  return utcMsToSerial(date.getTime(), system);
}

export function utcMsToSerial(ms: number, system: DateSystem = 1900): number {
  if (system === 1904) {
    return (ms - EPOCH_1904_UTC) / MS_PER_DAY;
  }
  const raw = (ms - EPOCH_1900_UTC) / MS_PER_DAY;
  // Anything from 1 Mar 1900 onwards gains the phantom day back.
  return raw >= PHANTOM_LEAP_SERIAL ? raw + 1 : raw;
}

export interface DateParts {
  year: number;
  month: number; // 1-12
  day: number; // 1-31
  hour: number;
  minute: number;
  second: number;
  millisecond: number;
}

/**
 * Decompose a serial into calendar parts, reproducing Excel's rendering of the
 * phantom day as 29 February 1900.
 */
export function serialToParts(serial: number, system: DateSystem = 1900): DateParts {
  const days = Math.floor(serial);
  const frac = serial - days;

  // Round the time-of-day to whole milliseconds before splitting it up, so that
  // 0.5 does not come back as 11:59:59.999 through floating-point drift.
  let msOfDay = Math.round(frac * MS_PER_DAY);
  let dayCarry = 0;
  if (msOfDay >= MS_PER_DAY) {
    msOfDay -= MS_PER_DAY;
    dayCarry = 1;
  }

  let y: number, mo: number, d: number;
  if (system === 1900 && days + dayCarry === PHANTOM_LEAP_SERIAL) {
    // The day that never existed.
    y = 1900;
    mo = 2;
    d = 29;
  } else {
    const utc = new Date(serialToUtcMs(days + dayCarry, system));
    y = utc.getUTCFullYear();
    mo = utc.getUTCMonth() + 1;
    d = utc.getUTCDate();
  }

  const hour = Math.floor(msOfDay / 3_600_000);
  const minute = Math.floor((msOfDay % 3_600_000) / 60_000);
  const second = Math.floor((msOfDay % 60_000) / 1000);
  const millisecond = msOfDay % 1000;
  return { year: y, month: mo, day: d, hour, minute, second, millisecond };
}

/**
 * Build a serial from calendar parts, matching Excel's DATE().
 *
 * Excel rolls out-of-range components over (month 13 becomes January of the next
 * year, day 0 becomes the last day of the previous month) and, in the 1900
 * system, maps a two-digit-ish year below 1900 by adding 1900 - which is why
 * `DATE(24,1,1)` is 1924, not 24 AD.
 */
export function partsToSerial(
  year: number,
  month: number,
  day: number,
  hour = 0,
  minute = 0,
  second = 0,
  system: DateSystem = 1900,
): number {
  let y = year;
  if (y >= 0 && y <= 1899) y += 1900;

  // Normalise the month first; Date.UTC already rolls days and times over.
  const monthIndex = month - 1;
  const yAdj = y + Math.floor(monthIndex / 12);
  const mAdj = ((monthIndex % 12) + 12) % 12;

  const ms = Date.UTC(yAdj, mAdj, day, hour, minute, second);
  // Date.UTC maps years 0-99 into 1900-1999; undo that for genuine early years.
  const corrected =
    yAdj >= 0 && yAdj <= 99 ? shiftCentury(ms, yAdj) : ms;
  return utcMsToSerial(corrected, system);
}

function shiftCentury(ms: number, year: number): number {
  const d = new Date(ms);
  d.setUTCFullYear(year);
  return d.getTime();
}

/** Whole days only, discarding the time component. */
export function serialDatePart(serial: number): number {
  return Math.floor(serial);
}

/** Time of day only, as a fraction of a day in [0, 1). */
export function serialTimePart(serial: number): number {
  const f = serial - Math.floor(serial);
  return f < 0 ? f + 1 : f;
}

/**
 * Day of week as an index with Sunday = 0, matching what Excel's WEEKDAY()
 * reports for `return_type` 1 (minus one).
 *
 * Excel derives this straight from the serial number, phantom leap day included,
 * so for dates before 1 March 1900 its answer disagrees with the real calendar:
 * Excel calls 1 January 1900 a Sunday, though it was a Monday. We reproduce
 * Excel's answer rather than the true one - a workbook that formats a column as
 * "dddd" must render the same text here as it does in Excel, and correctness
 * against the real calendar would be a visible mismatch, not a fix.
 */
export function weekdayFromSerial(serial: number, system: DateSystem = 1900): number {
  const days = Math.floor(serial);
  if (system === 1904) {
    // 1 Jan 1904 (serial 0) genuinely was a Friday, and the 1904 system has no
    // phantom day, so this branch agrees with the real calendar throughout.
    return ((days % 7) + 5) % 7;
  }
  return (((days - 1) % 7) + 7) % 7;
}

/** True when `year` is a leap year on the proleptic Gregorian calendar. */
export function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

export function daysInMonth(year: number, month: number): number {
  if (month === 2) return isLeapYear(year) ? 29 : 28;
  return [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1] ?? 30;
}
