/**
 * Date and time function tests.
 *
 * The oracle block asserts against the values LibreOffice actually computed in
 * fixtures/generated/formulas.calc.xlsx, read back through our own xlsx reader,
 * so sixteen of these expectations are not written from memory. Everything
 * after it covers what the fixture cannot reach: the 1900 phantom leap day, the
 * 1904 system, all eleven WEEKDAY variants, DATEDIF's undocumented units
 * including the negative-MD bug, both 30/360 conventions, all five YEARFRAC
 * bases, the INTL weekend codes, and the usual coercion edges - blanks against
 * zeros, numeric-looking text, and error propagation.
 *
 * Only DATETIME_FUNCTIONS is registered, plus three one-line stubs for TRUE,
 * FALSE and NA, which live in categories other agents are writing. A failure
 * here is therefore never another module's fault.
 */

import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { CellError, type Scalar, type Workbook, isError } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { Evaluator, type SheetStore } from '../src/evaluator.js';
import { DATETIME_FUNCTIONS } from '../src/functions/datetime.js';
import { parseFormula } from '../src/parser.js';
import { FunctionRegistry } from '../src/registry.js';
import { type Value, toExcelPrecision } from '../src/value.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

function storeFor(workbook: Workbook): SheetStore {
  const byName = new Map(workbook.sheets.map((s) => [s.name.toLowerCase(), s]));
  const sheet = (name: string) => byName.get(name.toLowerCase());
  return {
    getScalar: (name, row, col) => sheet(name)?.getValue(row, col) ?? null,
    *iterate(name, startRow, startCol, endRow, endCol) {
      const s = sheet(name);
      if (!s) return;
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const cell = s.getCell(r, c);
          if (cell !== undefined) yield { row: r, col: c, value: cell.value };
        }
      }
    },
    hasSheet: (name) => byName.has(name.toLowerCase()),
    sheetNames: () => workbook.sheets.map((s) => s.name),
    getDefinedName: () => undefined,
    usedBounds(name) {
      const b = sheet(name)?.bounds();
      return b ? { maxRow: b.maxRow, maxCol: b.maxCol } : null;
    },
  };
}

function registry(): FunctionRegistry {
  return new FunctionRegistry().registerAll(DATETIME_FUNCTIONS).registerAll([
    { name: 'TRUE', params: [], impl: () => true },
    { name: 'FALSE', params: [], impl: () => false },
    { name: 'NA', params: [], impl: () => CellError.NA },
  ]);
}

/** A fixed instant so NOW and TODAY are testable: 2024-03-01 18:00. */
const FIXED_NOW = 45352.75;

let workbook: Workbook;
let evaluator: Evaluator;
/** The same library over a 1904 workbook, to prove the epoch is honoured. */
let evaluator1904: Evaluator;
/** Oracle case name -> the formula and the value LibreOffice computed. */
let oracle: Map<string, { formula: string; value: Scalar }>;

beforeAll(() => {
  workbook = readXlsx(
    new Uint8Array(readFileSync(new URL('formulas.calc.xlsx', FIXTURES))),
  ).workbook;
  const store = storeFor(workbook);
  evaluator = new Evaluator(store, registry(), { dateSystem: 1900, now: FIXED_NOW });
  evaluator1904 = new Evaluator(store, registry(), { dateSystem: 1904, now: FIXED_NOW });

  oracle = new Map();
  const sheet = workbook.getSheet('Formulas')!;
  for (const { row, col, cell } of sheet.entries()) {
    if (col !== 2 || !cell.formula) continue;
    const name = sheet.getValue(row, 0);
    if (typeof name === 'string') oracle.set(name, { formula: cell.formula, value: cell.value });
  }
});

/** Evaluate a formula written without the leading `=`. */
function calc(formula: string): Value {
  const ast = parseFormula(formula, { origin: { row: 0, col: 0 } });
  return evaluator.evaluate({ ast, sheet: 'Formulas', row: 0, col: 0 });
}

function calc1904(formula: string): Value {
  const ast = parseFormula(formula, { origin: { row: 0, col: 0 } });
  return evaluator1904.evaluate({ ast, sheet: 'Formulas', row: 0, col: 0 });
}

function code(v: Value): string {
  return isError(v) ? v.code : `not an error: ${JSON.stringify(v)}`;
}

/** Compare at Excel's fifteen significant digits, the precision it compares at. */
function near(actual: Value, expected: number): void {
  expect(typeof actual).toBe('number');
  expect(toExcelPrecision(actual as number)).toBe(toExcelPrecision(expected));
}

// ---------------------------------------------------------------------------

describe('oracle: formulas.calc.xlsx', () => {
  const CASES = [
    'DATE', 'YEAR', 'MONTH', 'DAY', 'EOMONTH', 'EDATE', 'WEEKDAY', 'WEEKNUM',
    'DATEDIF', 'DAYS', 'NETWORKDAYS', 'WORKDAY', 'TIME', 'HOUR', 'MINUTE', 'SECOND',
  ];

  for (const name of CASES) {
    it(`reproduces ${name}`, () => {
      const c = oracle.get(name);
      expect(c, `case ${name} missing from the fixture`).toBeDefined();
      const actual = calc(c!.formula);
      if (typeof c!.value === 'number') near(actual, c!.value);
      else expect(actual).toEqual(c!.value);
    });
  }

  it('covers every function the fixture exercises', () => {
    // Guards against a rename quietly dropping a case from the list above.
    expect(CASES.every((n) => oracle.has(n))).toBe(true);
  });
});

describe('DATE', () => {
  it('builds the serial the oracle recorded', () => {
    expect(calc('DATE(2024,2,29)')).toBe(45351);
  });

  it('rolls months and days over in both directions', () => {
    expect(calc('DATE(2024,13,1)')).toBe(calc('DATE(2025,1,1)'));
    expect(calc('DATE(2024,0,1)')).toBe(calc('DATE(2023,12,1)'));
    expect(calc('DATE(2024,-1,1)')).toBe(calc('DATE(2023,11,1)'));
    expect(calc('DATE(2024,1,0)')).toBe(calc('DATE(2023,12,31)'));
    expect(calc('DATE(2024,1,32)')).toBe(calc('DATE(2024,2,1)'));
    expect(calc('DATE(2024,1,-1)')).toBe(calc('DATE(2023,12,30)'));
  });

  it('maps a year below 1900 into the twentieth century', () => {
    expect(calc('DATE(24,1,1)')).toBe(calc('DATE(1924,1,1)'));
    expect(calc('DATE(0,1,1)')).toBe(calc('DATE(1900,1,1)'));
    expect(calc('DATE(1899,12,31)')).toBe(calc('DATE(3799,12,31)'));
  });

  it('truncates fractional arguments towards zero', () => {
    expect(calc('DATE(2024.9,2.9,29.9)')).toBe(45351);
  });

  it('coerces numeric text and booleans', () => {
    expect(calc('DATE("2024","2","29")')).toBe(45351);
    expect(calc('DATE(2024,TRUE(),29)')).toBe(calc('DATE(2024,1,29)'));
  });

  it('reaches the phantom 29 February 1900', () => {
    expect(calc('DATE(1900,2,29)')).toBe(60);
    expect(calc('DATE(1900,3,1)')).toBe(61);
    expect(calc('DATE(1900,2,28)')).toBe(59);
    // Counting days from 1 January also lands on it.
    expect(calc('DATE(1900,1,60)')).toBe(60);
  });

  it('rejects years and results outside the representable range', () => {
    expect(code(calc('DATE(-1,1,1)'))).toBe('#NUM!');
    expect(code(calc('DATE(10000,1,1)'))).toBe('#NUM!');
    expect(code(calc('DATE(9999,13,1)'))).toBe('#NUM!');
    expect(code(calc('DATE(1900,1,-1)'))).toBe('#NUM!');
    expect(code(calc('DATE(1900,0,1)'))).toBe('#NUM!');
    expect(calc('DATE(9999,12,31)')).toBe(2958465);
  });

  it('propagates errors from its arguments', () => {
    expect(code(calc('DATE(1/0,1,1)'))).toBe('#DIV/0!');
    expect(code(calc('DATE(2024,NA(),1)'))).toBe('#N/A');
    expect(code(calc('DATE("abc",1,1)'))).toBe('#VALUE!');
  });
});

describe('TIME', () => {
  it('builds the fraction the oracle recorded', () => {
    near(calc('TIME(13,45,30)'), 49530 / 86400);
  });

  it('keeps only the time of day', () => {
    expect(calc('TIME(0,0,0)')).toBe(0);
    expect(calc('TIME(24,0,0)')).toBe(0);
    expect(calc('TIME(27,0,0)')).toBe(0.125);
    near(calc('TIME(12,60,0)'), 13 / 24);
    near(calc('TIME(23,59,59)'), 86399 / 86400);
    // 32767 is the documented ceiling for every component, so a whole day of
    // seconds is out of range rather than wrapping.
    expect(code(calc('TIME(0,0,86399)'))).toBe('#NUM!');
  });

  it('truncates its arguments', () => {
    expect(calc('TIME(1.9,2.9,3.9)')).toBe(calc('TIME(1,2,3)'));
  });

  it('rejects negative and out-of-range components', () => {
    expect(code(calc('TIME(-1,0,0)'))).toBe('#NUM!');
    expect(code(calc('TIME(0,-1,0)'))).toBe('#NUM!');
    expect(code(calc('TIME(0,0,-1)'))).toBe('#NUM!');
    expect(code(calc('TIME(32768,0,0)'))).toBe('#NUM!');
    expect(calc('TIME(32767,0,0)')).not.toBeInstanceOf(CellError);
  });
});

describe('DATEVALUE and TIMEVALUE', () => {
  it('reads the date spellings Excel accepts', () => {
    for (const text of [
      '2024-03-01',
      '2024/03/01',
      '3/1/2024',
      '3-1-2024',
      '1-Mar-2024',
      '1 March 2024',
      'March 1, 2024',
      'Mar 1 2024',
    ]) {
      expect(calc(`DATEVALUE("${text}")`), text).toBe(45352);
    }
  });

  it('applies the two-digit year window', () => {
    expect(calc('DATEVALUE("3/1/24")')).toBe(calc('DATE(2024,3,1)'));
    expect(calc('DATEVALUE("3/1/29")')).toBe(calc('DATE(2029,3,1)'));
    expect(calc('DATEVALUE("3/1/30")')).toBe(calc('DATE(1930,3,1)'));
    expect(calc('DATEVALUE("3/1/99")')).toBe(calc('DATE(1999,3,1)'));
  });

  it('drops the time of day', () => {
    expect(calc('DATEVALUE("2024-03-01 13:45:30")')).toBe(45352);
    expect(calc('DATEVALUE("March 1, 2024 1:45 PM")')).toBe(45352);
  });

  it('parses the phantom leap day the 1900 system believes in', () => {
    expect(calc('DATEVALUE("1900-02-29")')).toBe(60);
  });

  it('refuses text that is not a date', () => {
    expect(code(calc('DATEVALUE("13:45")'))).toBe('#VALUE!');
    expect(code(calc('DATEVALUE("not a date")'))).toBe('#VALUE!');
    expect(code(calc('DATEVALUE("2024-02-30")'))).toBe('#VALUE!');
    expect(code(calc('DATEVALUE("2023-02-29")'))).toBe('#VALUE!');
    expect(code(calc('DATEVALUE("2024-13-01")'))).toBe('#VALUE!');
    expect(code(calc('DATEVALUE("")'))).toBe('#VALUE!');
    // A serial is not text, however much it looks like a date.
    expect(code(calc('DATEVALUE(45352)'))).toBe('#VALUE!');
    expect(code(calc('DATEVALUE(TRUE())'))).toBe('#VALUE!');
  });

  it('reads clock text, with or without a meridiem', () => {
    near(calc('TIMEVALUE("13:45:30")'), 49530 / 86400);
    near(calc('TIMEVALUE("1:45:30 PM")'), 49530 / 86400);
    near(calc('TIMEVALUE("1:45:30 p.m.")'), 49530 / 86400);
    expect(calc('TIMEVALUE("12:00 AM")')).toBe(0);
    expect(calc('TIMEVALUE("12:00 PM")')).toBe(0.5);
    near(calc('TIMEVALUE("6 PM")'), 0.75);
  });

  it('ignores any date in the text rather than rejecting it', () => {
    expect(calc('TIMEVALUE("2024-03-01")')).toBe(0);
    near(calc('TIMEVALUE("2024-03-01 13:45:30")'), 49530 / 86400);
  });

  it('refuses text that is neither a date nor a time', () => {
    expect(code(calc('TIMEVALUE("abc")'))).toBe('#VALUE!');
    expect(code(calc('TIMEVALUE("12:60")'))).toBe('#VALUE!');
    expect(code(calc('TIMEVALUE("13:00 PM")'))).toBe('#VALUE!');
    expect(code(calc('TIMEVALUE(0.5)'))).toBe('#VALUE!');
  });
});

describe('NOW and TODAY', () => {
  it('read the recalculation-wide clock rather than the host clock', () => {
    expect(calc('NOW()')).toBe(FIXED_NOW);
    expect(calc('TODAY()')).toBe(45352);
    expect(calc('YEAR(TODAY())')).toBe(2024);
    expect(calc('HOUR(NOW())')).toBe(18);
  });

  it('is the only pair declared volatile', () => {
    const volatiles = DATETIME_FUNCTIONS.filter((f) => f.volatile).map((f) => f.name);
    expect(volatiles.sort()).toEqual(['NOW', 'TODAY']);
    // Nothing here depends on sheet shape.
    expect(DATETIME_FUNCTIONS.some((f) => f.structural)).toBe(false);
  });
});

describe('YEAR, MONTH and DAY', () => {
  it('decomposes a date', () => {
    expect(calc('YEAR(DATE(2024,2,29))')).toBe(2024);
    expect(calc('MONTH(DATE(2024,2,29))')).toBe(2);
    expect(calc('DAY(DATE(2024,2,29))')).toBe(29);
  });

  it('renders serial 0 as Excel does, January 0 1900', () => {
    expect(calc('YEAR(0)')).toBe(1900);
    expect(calc('MONTH(0)')).toBe(1);
    expect(calc('DAY(0)')).toBe(0);
  });

  it('renders the phantom leap day', () => {
    expect(calc('YEAR(60)')).toBe(1900);
    expect(calc('MONTH(60)')).toBe(2);
    expect(calc('DAY(60)')).toBe(29);
    expect(calc('DAY(61)')).toBe(1);
    expect(calc('MONTH(61)')).toBe(3);
    expect(calc('DAY(59)')).toBe(28);
  });

  it('ignores the time of day', () => {
    expect(calc('DAY(45352.999)')).toBe(1);
    expect(calc('MONTH(45352.999)')).toBe(3);
  });

  it('coerces blanks, booleans and text', () => {
    expect(calc('YEAR(Data!Z1)')).toBe(1900);
    expect(calc('DAY(Data!Z1)')).toBe(0);
    expect(calc('YEAR(TRUE())')).toBe(1900);
    expect(calc('DAY(TRUE())')).toBe(1);
    expect(calc('YEAR("2019-03-01")')).toBe(2019);
    expect(calc('YEAR(Data!E2)')).toBe(2019);
    expect(calc('MONTH(Data!E2)')).toBe(3);
    // Text that is a bare number is still read as a serial.
    expect(calc('YEAR("45352")')).toBe(2024);
  });

  it('rejects negative and overlarge serials', () => {
    expect(code(calc('YEAR(-1)'))).toBe('#NUM!');
    expect(code(calc('MONTH(-0.5)'))).toBe('#NUM!');
    expect(calc('YEAR(2958465)')).toBe(9999);
    expect(code(calc('YEAR(2958466)'))).toBe('#NUM!');
    expect(code(calc('YEAR("nonsense")'))).toBe('#VALUE!');
    expect(code(calc('YEAR(1/0)'))).toBe('#DIV/0!');
  });
});

describe('HOUR, MINUTE and SECOND', () => {
  it('splits the time of day', () => {
    expect(calc('HOUR(0.5)')).toBe(12);
    expect(calc('MINUTE(0.5)')).toBe(0);
    expect(calc('SECOND(0.5)')).toBe(0);
    expect(calc('HOUR(45352.75)')).toBe(18);
    expect(calc('HOUR(TIME(13,45,30))')).toBe(13);
    expect(calc('MINUTE(TIME(13,45,30))')).toBe(45);
    expect(calc('SECOND(TIME(13,45,30))')).toBe(30);
  });

  it('rounds to the nearest second, carrying past midnight', () => {
    expect(calc('SECOND(0.6/86400)')).toBe(1);
    expect(calc('SECOND(0.4/86400)')).toBe(0);
    expect(calc('HOUR(0.9999999)')).toBe(0);
    expect(calc('MINUTE(0.9999999)')).toBe(0);
    expect(calc('SECOND(0.9999999)')).toBe(0);
  });

  it('reads a whole date as midnight', () => {
    expect(calc('HOUR(DATE(2024,3,1))')).toBe(0);
    expect(calc('HOUR(Data!Z1)')).toBe(0);
  });

  it('rejects a negative serial', () => {
    expect(code(calc('HOUR(-0.5)'))).toBe('#NUM!');
  });
});

describe('WEEKDAY', () => {
  // 1 March 2024 is a Friday.
  const friday = 'DATE(2024,3,1)';

  it('implements all eleven documented return types', () => {
    const expected: [number, number][] = [
      [1, 6],
      [2, 5],
      [3, 4],
      [11, 5],
      [12, 4],
      [13, 3],
      [14, 2],
      [15, 1],
      [16, 7],
      [17, 6],
    ];
    for (const [type, want] of expected) {
      expect(calc(`WEEKDAY(${friday},${type})`), `type ${type}`).toBe(want);
    }
    expect(calc(`WEEKDAY(${friday})`)).toBe(6);
  });

  it('derives the day from the serial, phantom leap day included', () => {
    // Excel calls 1 January 1900 a Sunday although it was a Monday, because the
    // serial arithmetic carries the phantom day.
    expect(calc('WEEKDAY(1)')).toBe(1);
    expect(calc('WEEKDAY(60)')).toBe(4);
    expect(calc('WEEKDAY(61)')).toBe(5);
  });

  it('rejects an undocumented return type', () => {
    for (const type of [0, 4, 10, 18, -1]) {
      expect(code(calc(`WEEKDAY(${friday},${type})`)), `type ${type}`).toBe('#NUM!');
    }
  });
});

describe('WEEKNUM and ISOWEEKNUM', () => {
  it('numbers from the week containing 1 January', () => {
    expect(calc('WEEKNUM(DATE(2024,1,1),1)')).toBe(1);
    expect(calc('WEEKNUM(DATE(2024,1,6),1)')).toBe(1);
    expect(calc('WEEKNUM(DATE(2024,1,7),1)')).toBe(2);
    expect(calc('WEEKNUM(DATE(2024,3,1),1)')).toBe(9);
    expect(calc('WEEKNUM(DATE(2024,12,31),1)')).toBe(53);
  });

  it('honours the week-start variants', () => {
    // 7 January 2024 is a Sunday: it ends week 1 when weeks start on Monday and
    // begins week 2 when they start on Sunday.
    expect(calc('WEEKNUM(DATE(2024,1,7),2)')).toBe(1);
    expect(calc('WEEKNUM(DATE(2024,1,7),11)')).toBe(1);
    expect(calc('WEEKNUM(DATE(2024,1,7),17)')).toBe(2);
    // Type 13 starts weeks on Wednesday, so week 1 of 2024 is just Monday the
    // 1st and Tuesday the 2nd.
    expect(calc('WEEKNUM(DATE(2024,1,7),13)')).toBe(2);
    expect(calc('WEEKNUM(DATE(2024,1,2),13)')).toBe(1);
    expect(calc('WEEKNUM(DATE(2024,3,1),21)')).toBe(calc('ISOWEEKNUM(DATE(2024,3,1))'));
  });

  it('rejects an undocumented return type', () => {
    for (const type of [0, 3, 5, 10, 18, 20, 22]) {
      expect(code(calc(`WEEKNUM(DATE(2024,3,1),${type})`)), `type ${type}`).toBe('#NUM!');
    }
  });

  it('numbers ISO weeks by the year of their Thursday', () => {
    expect(calc('ISOWEEKNUM(DATE(2024,1,1))')).toBe(1);
    expect(calc('ISOWEEKNUM(DATE(2024,3,1))')).toBe(9);
    // A Sunday 1 January belongs to the last week of the previous year.
    expect(calc('ISOWEEKNUM(DATE(2023,1,1))')).toBe(52);
    expect(calc('ISOWEEKNUM(DATE(2021,1,1))')).toBe(53);
    expect(calc('ISOWEEKNUM(DATE(2020,12,31))')).toBe(53);
    // A Monday 30 December already belongs to the new year's week 1.
    expect(calc('ISOWEEKNUM(DATE(2024,12,30))')).toBe(1);
  });
});

describe('EDATE and EOMONTH', () => {
  it('clamps the day of month rather than spilling over', () => {
    expect(calc('EDATE(DATE(2024,1,31),1)')).toBe(45351);
    expect(calc('EDATE(DATE(2024,3,31),-1)')).toBe(45351);
    expect(calc('EDATE(DATE(2024,2,29),12)')).toBe(calc('DATE(2025,2,28)'));
    expect(calc('EDATE(DATE(2024,1,15),0)')).toBe(calc('DATE(2024,1,15)'));
    expect(calc('EDATE(DATE(2024,1,15),-13)')).toBe(calc('DATE(2022,12,15)'));
    expect(calc('EDATE(DATE(2024,1,15),24)')).toBe(calc('DATE(2026,1,15)'));
  });

  it('finds the end of month', () => {
    expect(calc('EOMONTH(DATE(2024,1,15),0)')).toBe(45322);
    expect(calc('EOMONTH(DATE(2024,2,15),0)')).toBe(calc('DATE(2024,2,29)'));
    expect(calc('EOMONTH(DATE(2023,2,1),0)')).toBe(calc('DATE(2023,2,28)'));
    expect(calc('EOMONTH(DATE(2024,1,31),-1)')).toBe(calc('DATE(2023,12,31)'));
    expect(calc('EOMONTH(DATE(2024,12,1),1)')).toBe(calc('DATE(2025,1,31)'));
  });

  it('uses the phantom February for 1900', () => {
    expect(calc('EOMONTH(DATE(1900,2,1),0)')).toBe(60);
    expect(calc('EOMONTH(DATE(1900,1,1),0)')).toBe(31);
  });

  it('reports #NUM! below the epoch and beyond 9999', () => {
    expect(code(calc('EDATE(DATE(1900,1,15),-1)'))).toBe('#NUM!');
    expect(code(calc('EOMONTH(DATE(1900,1,15),-1)'))).toBe('#NUM!');
    expect(code(calc('EDATE(DATE(9999,12,1),1)'))).toBe('#NUM!');
  });

  it('truncates the month count and propagates errors', () => {
    expect(calc('EDATE(DATE(2024,1,15),1.9)')).toBe(calc('DATE(2024,2,15)'));
    expect(code(calc('EDATE(DATE(2024,1,15),NA())'))).toBe('#N/A');
    expect(code(calc('EOMONTH("x",0)'))).toBe('#VALUE!');
  });
});

describe('DATEDIF', () => {
  const a = 'DATE(2020,1,1)';
  const b = 'DATE(2024,3,1)';

  it('implements the six undocumented units', () => {
    expect(calc(`DATEDIF(${a},${b},"Y")`)).toBe(4);
    expect(calc(`DATEDIF(${a},${b},"M")`)).toBe(50);
    expect(calc(`DATEDIF(${a},${b},"D")`)).toBe(1521);
    expect(calc(`DATEDIF(${a},${b},"MD")`)).toBe(0);
    expect(calc(`DATEDIF(${a},${b},"YM")`)).toBe(2);
    expect(calc(`DATEDIF(${a},${b},"YD")`)).toBe(60);
  });

  it('accepts the unit in any case', () => {
    expect(calc(`DATEDIF(${a},${b},"m")`)).toBe(50);
    expect(calc(`DATEDIF(${a},${b},"yd")`)).toBe(60);
  });

  it('counts only complete years and months', () => {
    expect(calc('DATEDIF(DATE(2020,3,1),DATE(2024,2,29),"Y")')).toBe(3);
    expect(calc('DATEDIF(DATE(2020,3,1),DATE(2024,3,1),"Y")')).toBe(4);
    expect(calc('DATEDIF(DATE(2024,1,31),DATE(2024,2,29),"M")')).toBe(0);
    expect(calc('DATEDIF(DATE(2024,1,31),DATE(2024,3,31),"M")')).toBe(2);
    expect(calc('DATEDIF(DATE(2020,11,15),DATE(2021,1,10),"YM")')).toBe(1);
  });

  it('reproduces the negative MD bug rather than correcting it', () => {
    // Excel subtracts the day numbers and, when that goes negative, adds the
    // length of the month before the end date - which here is short enough to
    // leave the answer below zero.
    expect(calc('DATEDIF(DATE(2015,1,31),DATE(2015,3,1),"MD")')).toBe(-2);
    expect(calc('DATEDIF(DATE(2024,1,31),DATE(2024,3,1),"MD")')).toBe(-1);
    expect(calc('DATEDIF(DATE(2000,1,31),DATE(2000,2,29),"MD")')).toBe(29);
    expect(calc('DATEDIF(DATE(2024,1,31),DATE(2024,3,31),"MD")')).toBe(0);
  });

  it('handles a zero-length interval', () => {
    expect(calc(`DATEDIF(${b},${b},"D")`)).toBe(0);
    expect(calc(`DATEDIF(${b},${b},"Y")`)).toBe(0);
    expect(calc(`DATEDIF(${b},${b},"MD")`)).toBe(0);
    expect(calc(`DATEDIF(${b},${b},"YD")`)).toBe(0);
  });

  it('refuses a reversed interval and an unknown unit', () => {
    expect(code(calc(`DATEDIF(${b},${a},"D")`))).toBe('#NUM!');
    expect(code(calc(`DATEDIF(${a},${b},"Q")`))).toBe('#NUM!');
    expect(code(calc(`DATEDIF(${a},${b},"")`))).toBe('#NUM!');
    expect(code(calc(`DATEDIF(${a},${b},5)`))).toBe('#NUM!');
  });
});

describe('DAYS', () => {
  it('subtracts whole days in either direction', () => {
    expect(calc('DAYS(DATE(2024,3,1),DATE(2024,1,1))')).toBe(60);
    expect(calc('DAYS(DATE(2024,1,1),DATE(2024,3,1))')).toBe(-60);
    expect(calc('DAYS(DATE(2024,1,1),DATE(2024,1,1))')).toBe(0);
  });

  it('discards the time of day and reads date text', () => {
    expect(calc('DAYS(45352.9,45292.1)')).toBe(60);
    expect(calc('DAYS("2024-03-01","2024-01-01")')).toBe(60);
    expect(calc('DAYS("45352","45292")')).toBe(60);
  });

  it('counts the phantom day, as the serials do', () => {
    expect(calc('DAYS(DATE(1900,3,1),DATE(1900,2,28))')).toBe(2);
  });

  it('propagates errors', () => {
    expect(code(calc('DAYS("x",1)'))).toBe('#VALUE!');
    expect(code(calc('DAYS(1,-1)'))).toBe('#NUM!');
  });
});

describe('DAYS360', () => {
  it('applies the US method by default', () => {
    expect(calc('DAYS360(DATE(2024,1,1),DATE(2024,12,31))')).toBe(360);
    expect(calc('DAYS360(DATE(2024,1,30),DATE(2024,3,31))')).toBe(60);
    expect(calc('DAYS360(DATE(2024,1,31),DATE(2024,3,31))')).toBe(60);
    expect(calc('DAYS360(DATE(2024,1,31),DATE(2024,2,29))')).toBe(29);
  });

  it('does not treat the end of February as the end of a month', () => {
    // The true NASD rule would give 30 here; Excel's DAYS360 gives 33, which is
    // the difference between this function and YEARFRAC basis 0.
    expect(calc('DAYS360(DATE(2015,2,28),DATE(2015,3,31))')).toBe(33);
  });

  it('applies the European method when asked', () => {
    expect(calc('DAYS360(DATE(2024,1,1),DATE(2024,12,31),TRUE())')).toBe(359);
    expect(calc('DAYS360(DATE(2024,1,31),DATE(2024,3,31),TRUE())')).toBe(60);
    expect(calc('DAYS360(DATE(2015,2,28),DATE(2015,3,31),TRUE())')).toBe(32);
  });

  it('negates a reversed interval', () => {
    expect(calc('DAYS360(DATE(2024,3,31),DATE(2024,1,1))')).toBe(-90);
    expect(calc('DAYS360(DATE(2024,1,1),DATE(2024,1,1))')).toBe(0);
  });

  it('propagates errors', () => {
    expect(code(calc('DAYS360("x",DATE(2024,1,1))'))).toBe('#VALUE!');
  });
});

describe('YEARFRAC', () => {
  it('uses NASD 30/360 for basis 0', () => {
    near(calc('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),0)'), 0.5);
    near(calc('YEARFRAC(DATE(2024,1,31),DATE(2024,2,29),0)'), 29 / 360);
    // Both ends the last day of February: the NASD rule folds both onto the
    // 30th, which is exactly where DAYS360's US method differs.
    near(calc('YEARFRAC(DATE(2024,2,29),DATE(2025,2,28),0)'), 1);
    near(calc('YEARFRAC(DATE(2015,2,28),DATE(2015,3,31),0)'), 30 / 360);
  });

  it('uses actual/actual for basis 1', () => {
    near(calc('YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),1)'), 1);
    near(calc('YEARFRAC(DATE(2024,1,1),DATE(2024,12,31),1)'), 365 / 366);
    near(calc('YEARFRAC(DATE(2023,1,1),DATE(2023,12,31),1)'), 364 / 365);
    // Over more than a year the denominator is the mean length of every
    // calendar year the interval touches, endpoints included.
    near(calc('YEARFRAC(DATE(2020,1,1),DATE(2024,1,1),1)'), 1461 / ((366 + 365 + 365 + 365 + 366) / 5));
  });

  it('uses actual/360 and actual/365 for bases 2 and 3', () => {
    near(calc('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),2)'), 182 / 360);
    near(calc('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)'), 182 / 365);
  });

  it('uses European 30/360 for basis 4', () => {
    near(calc('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),4)'), 0.5);
    near(calc('YEARFRAC(DATE(2024,1,31),DATE(2024,3,31),4)'), 60 / 360);
    near(calc('YEARFRAC(DATE(2015,2,28),DATE(2015,3,31),4)'), 32 / 360);
  });

  it('is symmetric and zero on an empty interval', () => {
    near(calc('YEARFRAC(DATE(2024,7,1),DATE(2024,1,1),0)'), 0.5);
    expect(calc('YEARFRAC(DATE(2024,1,1),DATE(2024,1,1),1)')).toBe(0);
  });

  it('rejects a basis outside 0 to 4', () => {
    expect(code(calc('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),5)'))).toBe('#NUM!');
    expect(code(calc('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),-1)'))).toBe('#NUM!');
    expect(calc('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),4.9)')).toBe(
      calc('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),4)'),
    );
  });
});

describe('WORKDAY and WORKDAY.INTL', () => {
  it('steps over weekends', () => {
    expect(calc('WORKDAY(DATE(2024,1,1),10)')).toBe(45306);
    expect(calc('WORKDAY(DATE(2024,1,1),0)')).toBe(45292);
    // Friday plus one working day is the following Monday.
    expect(calc('WORKDAY(DATE(2024,1,5),1)')).toBe(calc('DATE(2024,1,8)'));
    expect(calc('WORKDAY(DATE(2024,1,8),-1)')).toBe(calc('DATE(2024,1,5)'));
    // A start date that is itself a weekend is not counted.
    expect(calc('WORKDAY(DATE(2024,1,6),1)')).toBe(calc('DATE(2024,1,8)'));
  });

  it('jumps whole weeks without losing alignment', () => {
    // 260 working days is exactly 52 weeks from a Monday.
    expect(calc('WORKDAY(DATE(2024,1,1),260)')).toBe(calc('DATE(2024,12,30)'));
    expect(calc('WORKDAY(DATE(2024,12,30),-260)')).toBe(calc('DATE(2024,1,1)'));
  });

  it('skips holidays', () => {
    expect(calc('WORKDAY(DATE(2024,1,1),10,DATE(2024,1,8))')).toBe(calc('DATE(2024,1,16)'));
    expect(calc('WORKDAY(DATE(2024,1,1),5,{45296;45297})')).toBe(calc('DATE(2024,1,9)'));
    // A holiday that lands on a weekend changes nothing.
    expect(calc('WORKDAY(DATE(2024,1,1),10,DATE(2024,1,6))')).toBe(45306);
    // Neither does one outside the span walked.
    expect(calc('WORKDAY(DATE(2024,1,1),10,DATE(2024,6,3))')).toBe(45306);
    expect(calc('WORKDAY(DATE(2024,1,1),10,Data!Z1:Z5)')).toBe(45306);
    expect(calc('WORKDAY(DATE(2024,1,1),10,"2024-01-08")')).toBe(calc('DATE(2024,1,16)'));
  });

  it('honours the INTL weekend codes', () => {
    expect(calc('WORKDAY.INTL(DATE(2024,1,5),1)')).toBe(calc('DATE(2024,1,8)'));
    expect(calc('WORKDAY.INTL(DATE(2024,1,5),1,1)')).toBe(calc('DATE(2024,1,8)'));
    expect(calc('WORKDAY.INTL(DATE(2024,1,5),1,"0000011")')).toBe(calc('DATE(2024,1,8)'));
    // Weekend 11 is Sunday only, so Saturday becomes a working day.
    expect(calc('WORKDAY.INTL(DATE(2024,1,5),1,11)')).toBe(calc('DATE(2024,1,6)'));
    // Weekend 7 is Friday and Saturday.
    expect(calc('WORKDAY.INTL(DATE(2024,1,4),1,7)')).toBe(calc('DATE(2024,1,7)'));
    expect(calc('WORKDAY.INTL(DATE(2024,1,5),1,"0000000")')).toBe(calc('DATE(2024,1,6)'));
    expect(calc('WORKDAY.INTL(DATE(2024,1,1),10,1,DATE(2024,1,8))')).toBe(
      calc('DATE(2024,1,16)'),
    );
  });

  it('rejects a malformed or exhaustive weekend', () => {
    expect(code(calc('WORKDAY.INTL(DATE(2024,1,5),1,"1111111")'))).toBe('#VALUE!');
    expect(code(calc('WORKDAY.INTL(DATE(2024,1,5),1,"abcdefg")'))).toBe('#VALUE!');
    expect(code(calc('WORKDAY.INTL(DATE(2024,1,5),1,"00011")'))).toBe('#VALUE!');
    expect(code(calc('WORKDAY.INTL(DATE(2024,1,5),1,0)'))).toBe('#NUM!');
    expect(code(calc('WORKDAY.INTL(DATE(2024,1,5),1,8)'))).toBe('#NUM!');
    expect(code(calc('WORKDAY.INTL(DATE(2024,1,5),1,18)'))).toBe('#NUM!');
  });

  it('reports a result outside the date range as #NUM!', () => {
    expect(code(calc('WORKDAY(DATE(1900,1,1),-10)'))).toBe('#NUM!');
    expect(code(calc('WORKDAY(DATE(9999,12,31),10)'))).toBe('#NUM!');
  });

  it('propagates errors from dates and holidays', () => {
    expect(code(calc('WORKDAY("x",1)'))).toBe('#VALUE!');
    expect(code(calc('WORKDAY(DATE(2024,1,1),NA())'))).toBe('#N/A');
    expect(code(calc('WORKDAY(DATE(2024,1,1),1,{45296;"x"})'))).toBe('#VALUE!');
  });
});

describe('NETWORKDAYS and NETWORKDAYS.INTL', () => {
  it('counts both endpoints', () => {
    expect(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))')).toBe(23);
    expect(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,1))')).toBe(1);
    // A weekend-only span has no working days.
    expect(calc('NETWORKDAYS(DATE(2024,1,6),DATE(2024,1,7))')).toBe(0);
  });

  it('counts backwards as negative', () => {
    expect(calc('NETWORKDAYS(DATE(2024,1,31),DATE(2024,1,1))')).toBe(-23);
  });

  it('subtracts holidays exactly once', () => {
    expect(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),DATE(2024,1,15))')).toBe(22);
    expect(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),{45292;45293})')).toBe(21);
    expect(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),{45306;45306})')).toBe(22);
    // Outside the range, or on a weekend, a holiday changes nothing.
    expect(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),DATE(2024,2,1))')).toBe(23);
    expect(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),DATE(2024,1,6))')).toBe(23);
    // An empty range contributes nothing rather than a blank-as-zero holiday.
    expect(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),Data!Z1:Z5)')).toBe(23);
  });

  it('honours the INTL weekend codes', () => {
    expect(calc('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31))')).toBe(23);
    expect(calc('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),1)')).toBe(23);
    expect(calc('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),"0000011")')).toBe(23);
    // Sunday only: January 2024 has four Sundays.
    expect(calc('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),11)')).toBe(27);
    // Every day a weekend gives zero rather than an error.
    expect(calc('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),"1111111")')).toBe(0);
    expect(calc('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),"0000000")')).toBe(31);
    expect(calc('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),1,DATE(2024,1,15))')).toBe(22);
  });

  it('rejects a malformed weekend', () => {
    expect(code(calc('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),"11")'))).toBe('#VALUE!');
    expect(code(calc('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),0)'))).toBe('#NUM!');
  });

  it('propagates errors', () => {
    expect(code(calc('NETWORKDAYS("x",DATE(2024,1,31))'))).toBe('#VALUE!');
    expect(code(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),{45296;"x"})'))).toBe('#VALUE!');
    expect(code(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),NA())'))).toBe('#N/A');
  });
});

describe('the 1904 date system', () => {
  it('shifts every serial by the 1462-day epoch difference', () => {
    expect(calc1904('DATE(2024,1,1)')).toBe(45292 - 1462);
    expect(calc1904('YEAR(0)')).toBe(1904);
    expect(calc1904('MONTH(0)')).toBe(1);
    expect(calc1904('DAY(0)')).toBe(1);
    // 1 January 1904 genuinely was a Friday, and the 1904 system has no
    // phantom day to disagree with the calendar.
    expect(calc1904('WEEKDAY(0)')).toBe(6);
    expect(calc1904('DATEVALUE("2024-01-01")')).toBe(45292 - 1462);
  });

  it('keeps differences and intervals identical to the 1900 system', () => {
    expect(calc1904('DAYS(DATE(2024,3,1),DATE(2024,1,1))')).toBe(60);
    expect(calc1904('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))')).toBe(23);
    expect(calc1904('WORKDAY(DATE(2024,1,1),10)')).toBe(45306 - 1462);
    expect(calc1904('DATEDIF(DATE(2020,1,1),DATE(2024,3,1),"M")')).toBe(50);
    near(calc1904('YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),0)'), 0.5);
  });

  it('rejects dates before 1904 and beyond 9999', () => {
    expect(code(calc1904('DATE(1900,1,1)'))).toBe('#NUM!');
    expect(code(calc1904('DATE(1903,12,31)'))).toBe('#NUM!');
    expect(calc1904('DATE(9999,12,31)')).toBe(2957003);
    expect(code(calc1904('YEAR(2957004)'))).toBe('#NUM!');
  });
});

describe('registry metadata', () => {
  it('registers all twenty-five functions', () => {
    const r = new FunctionRegistry().registerAll(DATETIME_FUNCTIONS);
    expect(r.size).toBe(25);
    expect(r.names()).toEqual([
      'DATE', 'DATEDIF', 'DATEVALUE', 'DAY', 'DAYS', 'DAYS360', 'EDATE', 'EOMONTH',
      'HOUR', 'ISOWEEKNUM', 'MINUTE', 'MONTH', 'NETWORKDAYS', 'NETWORKDAYS.INTL',
      'NOW', 'SECOND', 'TIME', 'TIMEVALUE', 'TODAY', 'WEEKDAY', 'WEEKNUM',
      'WORKDAY', 'WORKDAY.INTL', 'YEAR', 'YEARFRAC',
    ]);
  });

  it('marks the post-2007 names as future functions', () => {
    const r = new FunctionRegistry().registerAll(DATETIME_FUNCTIONS);
    for (const name of ['DAYS', 'ISOWEEKNUM', 'WORKDAY.INTL', 'NETWORKDAYS.INTL']) {
      expect(r.get(name)?.futureFunction, name).toBe(true);
    }
    for (const name of ['DATE', 'TIME', 'WORKDAY', 'NETWORKDAYS']) {
      expect(r.get(name)?.futureFunction, name).toBe(false);
    }
    // A formula that still carries the storage prefix must resolve.
    expect(r.get('_xlfn.ISOWEEKNUM')?.name).toBe('ISOWEEKNUM');
  });

  it('reports arity problems as #VALUE!', () => {
    expect(code(calc('DATE(2024,1)'))).toBe('#VALUE!');
    expect(code(calc('NOW(1)'))).toBe('#VALUE!');
    expect(code(calc('DAYS(1)'))).toBe('#VALUE!');
  });
});
