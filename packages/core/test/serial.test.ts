import { describe, expect, it } from 'vitest';
import {
  PHANTOM_LEAP_SERIAL,
  daysInMonth,
  dateToSerial,
  isLeapYear,
  partsToSerial,
  serialDatePart,
  serialTimePart,
  serialToDate,
  serialToParts,
  weekdayFromSerial,
} from '../src/serial.js';

describe('1900 date system', () => {
  it('anchors serial 1 at 1900-01-01', () => {
    const p = serialToParts(1);
    expect([p.year, p.month, p.day]).toEqual([1900, 1, 1]);
  });

  it('renders serial 60 as the phantom 29 Feb 1900', () => {
    const p = serialToParts(PHANTOM_LEAP_SERIAL);
    expect([p.year, p.month, p.day]).toEqual([1900, 2, 29]);
  });

  it('renders serial 59 as 28 Feb and serial 61 as 1 Mar', () => {
    expect(serialToParts(59)).toMatchObject({ year: 1900, month: 2, day: 28 });
    expect(serialToParts(61)).toMatchObject({ year: 1900, month: 3, day: 1 });
  });

  // Anchors cross-checked against LibreOffice-recalculated fixture values.
  it.each([
    [45351, 2024, 2, 29],
    [45322, 2024, 1, 31],
    [45306, 2024, 1, 15],
    [36526, 2000, 1, 1],
    [25569, 1970, 1, 1],
    [44927, 2023, 1, 1],
  ])('serial %i is %i-%i-%i', (serial, y, m, d) => {
    expect(serialToParts(serial)).toMatchObject({ year: y, month: m, day: d });
    expect(partsToSerial(y, m, d)).toBe(serial);
  });

  it('round-trips serial -> parts -> serial across a wide span', () => {
    for (let s = 1; s < 60_000; s += 37) {
      const p = serialToParts(s);
      expect(partsToSerial(p.year, p.month, p.day)).toBe(s);
    }
  });

  it('round-trips through Date for post-1900-03-01 values', () => {
    for (let s = 61; s < 60_000; s += 101) {
      expect(dateToSerial(serialToDate(s))).toBe(s);
    }
  });
});

describe('1904 date system', () => {
  it('anchors serial 0 at 1904-01-01 and has no phantom day', () => {
    expect(serialToParts(0, 1904)).toMatchObject({ year: 1904, month: 1, day: 1 });
    expect(serialToParts(59, 1904)).toMatchObject({ year: 1904, month: 2, day: 29 });
  });

  it('is offset from the 1900 system by 1462 days', () => {
    // The two systems differ by four years, one of which (1904) is a real leap year.
    expect(partsToSerial(2024, 2, 29, 0, 0, 0, 1900) - partsToSerial(2024, 2, 29, 0, 0, 0, 1904)).toBe(1462);
  });
});

describe('time components', () => {
  it('splits noon exactly', () => {
    const p = serialToParts(45351.5);
    expect([p.hour, p.minute, p.second]).toEqual([12, 0, 0]);
  });

  it('does not drift on 13:45:30', () => {
    const t = partsToSerial(1899, 12, 31, 13, 45, 30);
    const p = serialToParts(t);
    expect([p.hour, p.minute, p.second]).toEqual([13, 45, 30]);
  });

  it.each([0, 0.25, 0.5, 0.75, 1 / 3, 2 / 3, 0.1, 0.9999])('round-trips fraction %f', (f) => {
    const p = serialToParts(1000 + f);
    const back = (p.hour * 3600 + p.minute * 60 + p.second) / 86_400 + p.millisecond / 86_400_000;
    expect(back).toBeCloseTo(f, 6);
  });

  it('separates date and time parts', () => {
    expect(serialDatePart(45351.75)).toBe(45351);
    expect(serialTimePart(45351.75)).toBeCloseTo(0.75, 10);
  });
});

describe('weekday', () => {
  // 2024-03-01 was a Friday; Excel's default return_type has Sunday = 1, so 6.
  // This value comes from the LibreOffice-recalculated fixture, not from memory.
  it('matches the oracle for 2024-03-01', () => {
    expect(weekdayFromSerial(partsToSerial(2024, 3, 1)) + 1).toBe(6);
  });

  it('advances by one per day', () => {
    const base = partsToSerial(2024, 3, 1);
    for (let i = 0; i < 14; i++) {
      expect(weekdayFromSerial(base + i)).toBe((5 + i) % 7);
    }
  });

  it("reproduces Excel's off-by-one for dates before the phantom leap day", () => {
    // 1 Jan 1900 was really a Monday, but Excel reports Sunday because the
    // phantom 29 Feb 1900 shifts every earlier serial. We match Excel.
    expect(weekdayFromSerial(1)).toBe(0);
    expect(weekdayFromSerial(61)).toBe(4); // 1 Mar 1900, a Thursday - correct again
  });

  it('agrees with the real calendar in the 1904 system', () => {
    expect(weekdayFromSerial(0, 1904)).toBe(5); // 1 Jan 1904 was a Friday
  });
});

describe('DATE() roll-over semantics', () => {
  it('rolls month 13 into the next January', () => {
    expect(partsToSerial(2024, 13, 1)).toBe(partsToSerial(2025, 1, 1));
  });

  it('rolls month 0 into the previous December', () => {
    expect(partsToSerial(2024, 0, 1)).toBe(partsToSerial(2023, 12, 1));
  });

  it('rolls day 0 into the last day of the previous month', () => {
    expect(partsToSerial(2024, 3, 0)).toBe(partsToSerial(2024, 2, 29));
  });

  it('rolls day 32 into the next month', () => {
    expect(partsToSerial(2024, 1, 32)).toBe(partsToSerial(2024, 2, 1));
  });

  it('maps years below 1900 by adding 1900', () => {
    expect(partsToSerial(24, 1, 1)).toBe(partsToSerial(1924, 1, 1));
  });
});

describe('calendar helpers', () => {
  it.each([
    [2024, true],
    [2023, false],
    [2000, true],
    [1900, false],
    [2100, false],
  ])('isLeapYear(%i) === %s', (y, want) => {
    expect(isLeapYear(y)).toBe(want);
  });

  it('knows February lengths', () => {
    expect(daysInMonth(2024, 2)).toBe(29);
    expect(daysInMonth(2023, 2)).toBe(28);
    expect(daysInMonth(2024, 1)).toBe(31);
  });
});
