/**
 * Math and trigonometry.
 *
 * The first block replays the LibreOffice-recalculated fixtures: for every case
 * in fixtures/generated whose formula uses only functions from this module, the
 * formula is re-parsed and re-evaluated here and compared against the value the
 * oracle actually computed. Numbers are compared at Excel's fifteen significant
 * digits, which is the precision the oracle stored and the precision Excel
 * compares at, so EXP(1) matches 2.71828182845905 without pretending the double
 * is that value.
 *
 * The registry under test holds MATH_FUNCTIONS alone, so a failure here is
 * never another category's fault.
 */

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, type Scalar, Workbook, isError } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { Evaluator } from '../src/evaluator.js';
import { MATH_FUNCTIONS } from '../src/functions/math.js';
import { parseFormula } from '../src/parser.js';
import { FunctionRegistry } from '../src/registry.js';
import { WorkbookStore } from '../src/store.js';
import { type Value, isArray, toExcelPrecision } from '../src/value.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

const registry = new FunctionRegistry().registerAll(MATH_FUNCTIONS);

function evaluatorFor(workbook: Workbook): Evaluator {
  return new Evaluator(new WorkbookStore(workbook), registry, {
    dateSystem: workbook.dateSystem,
  });
}

/** Case name -> the formula the oracle recalculated and the value it produced. */
function oracleCases(file: string, sheetName: string): {
  cases: Map<string, { formula: string; value: Scalar; row: number }>;
  evaluate: (formula: string, row: number) => Value;
} {
  const { workbook } = readXlsx(new Uint8Array(readFileSync(new URL(file, FIXTURES))));
  const sheet = workbook.getSheet(sheetName)!;
  const ev = evaluatorFor(workbook);
  const cases = new Map<string, { formula: string; value: Scalar; row: number }>();
  for (const { row, col, cell } of sheet.entries()) {
    if (col !== 2 || !cell.formula) continue;
    const name = sheet.getValue(row, 0);
    if (typeof name === 'string') cases.set(name, { formula: cell.formula, value: cell.value, row });
  }
  return {
    cases,
    evaluate: (formula, row) =>
      ev.evaluate({
        // LibreOffice stores the union operator as '~'; the file format layer
        // has not normalised it yet, so the oracle text needs it here.
        ast: parseFormula(formula.replaceAll('~', ','), { origin: { row, col: 2 } }),
        sheet: sheetName,
        row,
        col: 2,
      }),
  };
}

function sameScalar(actual: Value, expected: Scalar): void {
  if (typeof expected === 'number' && typeof actual === 'number') {
    expect(toExcelPrecision(actual)).toBe(toExcelPrecision(expected));
    return;
  }
  expect(actual).toEqual(expected);
}

const FORMULAS = oracleCases('formulas.calc.xlsx', 'Formulas');
const PRECEDENCE = oracleCases('precedence.calc.xlsx', 'Precedence');

/** Cases in formulas.calc.xlsx whose formulas this module can evaluate alone. */
const FORMULA_CASES = [
  'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUMSQ', 'PRODUCT',
  'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'MROUND', 'CEILING.MATH', 'FLOOR.MATH',
  'INT', 'TRUNC', 'ABS', 'SIGN', 'MOD', 'QUOTIENT', 'POWER', 'SQRT',
  'EXP', 'LN', 'LOG10', 'LOG', 'PI', 'SIN', 'COS', 'TAN', 'ATAN2',
  'DEGREES', 'RADIANS', 'GCD', 'LCM', 'FACT', 'COMBIN',
  'ARRAY_SUM', 'RANGE_OP', 'UNION_OP', 'ERR_REF',
];

const PRECEDENCE_CASES = [
  'round_half', 'round_half_neg', 'banker_check', 'sum_cancel', 'sum_cancel_eq',
  'sqrt_sq', 'sqrt_sq_eq', 'inner_residue', 'range_op', 'intersect_op', 'union_op',
];

describe('oracle: formulas.calc.xlsx', () => {
  for (const name of FORMULA_CASES) {
    it(`reproduces ${name}`, () => {
      const c = FORMULAS.cases.get(name);
      expect(c, `case ${name} missing from the fixture`).toBeDefined();
      sameScalar(FORMULAS.evaluate(c!.formula, c!.row), c!.value);
    });
  }
});

describe('oracle: precedence.calc.xlsx', () => {
  for (const name of PRECEDENCE_CASES) {
    it(`reproduces ${name}`, () => {
      const c = PRECEDENCE.cases.get(name);
      expect(c, `case ${name} missing from the fixture`).toBeDefined();
      sameScalar(PRECEDENCE.evaluate(c!.formula, c!.row), c!.value);
    });
  }
});

describe('oracle: the aggregate machinery behind SUBTOTAL and AGGREGATE', () => {
  // AGGREGATE reaches the same statistics the oracle computed through
  // AVERAGE/MEDIAN/STDEV.S/VAR.S/LARGE/SMALL/PERCENTILE.INC, so those cached
  // values verify code this module owns even though it does not own the names.
  const viaAggregate = (fn: number, k?: string): Value =>
    FORMULAS.evaluate(
      `AGGREGATE(${fn},6,Data!D2:D9${k === undefined ? '' : `,${k}`})`,
      1,
    );

  it('matches AVERAGE', () => sameScalar(viaAggregate(1), FORMULAS.cases.get('AVERAGE')!.value));
  it('matches COUNT', () => expect(viaAggregate(2)).toBe(8));
  it('matches MAX', () => sameScalar(viaAggregate(4), FORMULAS.cases.get('MAX')!.value));
  it('matches MIN', () => sameScalar(viaAggregate(5), FORMULAS.cases.get('MIN')!.value));
  it('matches STDEV.S', () => sameScalar(viaAggregate(7), FORMULAS.cases.get('STDEV.S')!.value));
  it('matches SUM', () => sameScalar(viaAggregate(9), FORMULAS.cases.get('SUM')!.value));
  it('matches VAR.S', () => sameScalar(viaAggregate(10), FORMULAS.cases.get('VAR.S')!.value));
  it('matches MEDIAN', () => sameScalar(viaAggregate(12), FORMULAS.cases.get('MEDIAN')!.value));
  it('matches LARGE', () => sameScalar(viaAggregate(14, '2'), FORMULAS.cases.get('LARGE')!.value));
  it('matches SMALL', () => sameScalar(viaAggregate(15, '2'), FORMULAS.cases.get('SMALL')!.value));
  it('matches PERCENTILE.INC', () =>
    sameScalar(viaAggregate(16, '0.5'), FORMULAS.cases.get('PCT_RANK')!.value));
  it('matches SUBTOTAL to AGGREGATE for the shared function numbers', () => {
    for (const fn of [1, 2, 4, 5, 6, 7, 8, 9, 10, 11]) {
      const a = FORMULAS.evaluate(`AGGREGATE(${fn},0,Data!D2:D9)`, 1);
      const b = FORMULAS.evaluate(`SUBTOTAL(${fn},Data!D2:D9)`, 1);
      const c = FORMULAS.evaluate(`SUBTOTAL(${fn + 100},Data!D2:D9)`, 1);
      expect(b).toEqual(a);
      expect(c).toEqual(a);
    }
  });
});

// ---------------------------------------------------------------------------
// A hand-built workbook for the cases the fixtures do not reach.
// ---------------------------------------------------------------------------

const book = new Workbook();
const data = book.addSheet('S');
book.addSheet('Calc');

// A1:A5 mixes the value kinds an aggregate has to sort out; A5 stays blank.
data.setValue(0, 0, 1);
data.setValue(1, 0, 2);
data.setValue(2, 0, '3');
data.setValue(3, 0, true);
// B1:B5, the parallel numeric column.
data.setValue(0, 1, 10);
data.setValue(1, 1, 20);
data.setValue(2, 1, 30);
data.setValue(3, 1, 40);
data.setValue(4, 1, 50);
// C1:C3 carries an error value.
data.setValue(0, 2, 5);
data.setValue(1, 2, CellError.DIV0);
data.setValue(2, 2, 7);
// D1:D4, a criteria column.
data.setValue(0, 3, 'apple');
data.setValue(1, 3, 'apricot');
data.setValue(2, 3, 'banana');
data.setValue(3, 3, 'apple');
// E1:E4, the matching amounts.
data.setValue(0, 4, 1);
data.setValue(1, 4, 2);
data.setValue(2, 4, 4);
data.setValue(3, 4, 8);
// A far-away cell, so a whole-column reference has a large but sparse extent.
data.setValue(9999, 0, 100);

// A column holding more values than a function call can take as arguments, so
// that MAX and MIN are exercised past the point where spreading them into
// Math.max would overflow the call stack.
const wide = book.addSheet('Wide');
for (let r = 0; r < 200_000; r++) wide.setValue(r, 0, (r * 7919) % 100_003);

const ev = evaluatorFor(book);

/** Evaluate at Calc!A1, keeping arrays and references intact. */
function calc(formula: string): Value {
  return ev.evaluate({
    ast: parseFormula(formula, { origin: { row: 0, col: 0 } }),
    sheet: 'Calc',
    row: 0,
    col: 0,
  });
}

/** Evaluate and round to Excel's display precision, for transcendental results. */
function num(formula: string): number | Value {
  const v = calc(formula);
  return typeof v === 'number' ? toExcelPrecision(v) : v;
}

function code(formula: string): string {
  const v = calc(formula);
  return isError(v) ? v.code : `not an error: ${String(v)}`;
}

function grid(formula: string): { rows: number; cols: number; data: Scalar[] } {
  const v = calc(formula);
  if (!isArray(v)) throw new Error(`expected an array, got ${String(v)}`);
  return {
    rows: v.rows,
    cols: v.cols,
    data: v.data.map((x) => (typeof x === 'number' ? toExcelPrecision(x) : x)),
  };
}

describe('SUM', () => {
  it('ignores text and booleans inside a range', () => {
    expect(calc('SUM(S!A1:A5)')).toBe(3);
  });

  it('coerces the same values when they are direct arguments', () => {
    expect(calc('SUM(TRUE)')).toBe(1);
    expect(calc('SUM("3")')).toBe(3);
    expect(calc('SUM(1,TRUE,"3")')).toBe(5);
  });

  it('reports #VALUE! for direct text that is not a number', () => {
    expect(code('SUM("abc")')).toBe('#VALUE!');
  });

  it('ignores non-numeric text sitting in a range', () => {
    expect(calc('SUM(S!D1:D4)')).toBe(0);
  });

  it('is zero over an entirely blank range', () => {
    expect(calc('SUM(S!Z1:Z9)')).toBe(0);
  });

  it('treats an omitted argument as zero', () => {
    expect(calc('SUM(5,)')).toBe(5);
  });

  it('propagates an error found inside a range', () => {
    expect(code('SUM(S!C1:C3)')).toBe('#DIV/0!');
  });

  it('sums an array constant, ignoring its booleans and text', () => {
    expect(calc('SUM({1,2,3;4,5,6})')).toBe(21);
    expect(calc('SUM({1,TRUE,"x"})')).toBe(1);
  });

  it('adds through excelAdd, so cancellation snaps as Excel does', () => {
    expect(calc('SUM(0.1,0.2)-0.3')).toBe(0);
  });

  it('reports an overflow as #NUM! rather than handing back an Infinity', () => {
    expect(code('SUM(1E308,1E308)')).toBe('#NUM!');
    expect(code('SUMIF({1,2},">0",{1E308,1E308})')).toBe('#NUM!');
    expect(code('SUMIFS({1E308,1E308},{1,2},">0")')).toBe('#NUM!');
  });

  it('walks only materialised cells of a whole-column reference', () => {
    // Column A holds five cells and one at row 10000; a dense walk would be a
    // million reads, and the answer must still be right.
    expect(calc('SUM(S!A:A)')).toBe(103);
  });

  it('handles the intersection and union reference operators', () => {
    expect(calc('SUM(S!A1:B3 S!B1:B5)')).toBe(60);
    expect(calc('SUM((S!B1:B2,S!B4:B5))')).toBe(150);
  });
});

describe('PRODUCT, SUMSQ and the paired sums', () => {
  it('multiplies direct arguments', () => {
    expect(calc('PRODUCT(2,3,4)')).toBe(24);
  });

  it('is zero, not one, over an empty range', () => {
    expect(calc('PRODUCT(S!Z1:Z9)')).toBe(0);
  });

  it('ignores text in a range but coerces it directly', () => {
    expect(calc('PRODUCT(S!A1:A5)')).toBe(2);
    expect(calc('PRODUCT("3",4)')).toBe(12);
  });

  it('squares and sums', () => {
    expect(calc('SUMSQ(3,4)')).toBe(25);
    expect(calc('SUMSQ(S!B1:B5)')).toBe(5500);
  });

  it('computes the three paired-array sums', () => {
    expect(calc('SUMX2MY2({3,4},{1,2})')).toBe(20);
    expect(calc('SUMX2PY2({3,4},{1,2})')).toBe(30);
    expect(calc('SUMXMY2({3,4},{1,2})')).toBe(8);
  });

  it('reports #N/A when the two arrays hold different counts', () => {
    expect(code('SUMX2MY2({1,2,3},{1,2})')).toBe('#N/A');
  });

  it('skips a pair unless both halves are numbers', () => {
    expect(calc('SUMX2PY2({3,"x"},{1,2})')).toBe(10);
  });
});

describe('SUMPRODUCT', () => {
  it('multiplies elementwise and sums', () => {
    expect(calc('SUMPRODUCT({1,2,3},{4,5,6})')).toBe(32);
  });

  it('sums a single array', () => {
    expect(calc('SUMPRODUCT(S!B1:B5)')).toBe(150);
  });

  it('treats text, blanks and booleans as zero', () => {
    expect(calc('SUMPRODUCT(S!A1:A5,S!B1:B5)')).toBe(50);
  });

  it('reports #VALUE! when the shapes differ', () => {
    expect(code('SUMPRODUCT({1,2,3},{1,2})')).toBe('#VALUE!');
  });

  it('propagates an error from any array', () => {
    expect(code('SUMPRODUCT(S!C1:C3,S!C1:C3)')).toBe('#DIV/0!');
  });
});

describe('SUMIF and SUMIFS', () => {
  it('sums with a comparison criterion', () => {
    expect(calc('SUMIF(S!B1:B5,">25")')).toBe(120);
  });

  it('sums a parallel range', () => {
    expect(calc('SUMIF(S!D1:D4,"apple",S!E1:E4)')).toBe(9);
  });

  it('takes the criteria range shape when sum_range is a single cell', () => {
    expect(calc('SUMIF(S!D1:D4,"apple",S!E1)')).toBe(9);
  });

  it('matches wildcards', () => {
    expect(calc('SUMIF(S!D1:D4,"ap*",S!E1:E4)')).toBe(11);
    expect(calc('SUMIF(S!D1:D4,"a?ple",S!E1:E4)')).toBe(9);
  });

  it('matches case-insensitively, as Excel does', () => {
    expect(calc('SUMIF(S!D1:D4,"APPLE",S!E1:E4)')).toBe(9);
  });

  it('lets a criteria-range error fail to match instead of propagating', () => {
    expect(calc('SUMIF(S!C1:C3,">4",S!B1:B3)')).toBe(40);
  });

  it('propagates an error from a matching row of the sum range', () => {
    expect(code('SUMIF(S!B1:B3,">5",S!C1:C3)')).toBe('#DIV/0!');
  });

  it('applies every SUMIFS condition', () => {
    expect(calc('SUMIFS(S!E1:E4,S!D1:D4,"apple",S!E1:E4,">1")')).toBe(8);
  });

  it('reports #VALUE! when a SUMIFS range is the wrong shape', () => {
    expect(code('SUMIFS(S!E1:E4,S!D1:D3,"apple")')).toBe('#VALUE!');
  });

  it('reports #VALUE! for a dangling SUMIFS criterion', () => {
    expect(code('SUMIFS(S!E1:E4,S!D1:D4,"apple",S!E1:E4)')).toBe('#VALUE!');
  });

  it('reads a criterion out of a cell reference', () => {
    expect(calc('SUMIF(S!D1:D4,S!D1,S!E1:E4)')).toBe(9);
  });
});

describe('SUBTOTAL and AGGREGATE', () => {
  it('rejects a function number outside 1-11 and 101-111', () => {
    expect(code('SUBTOTAL(12,S!B1:B5)')).toBe('#VALUE!');
    expect(code('SUBTOTAL(112,S!B1:B5)')).toBe('#VALUE!');
    expect(code('SUBTOTAL(0,S!B1:B5)')).toBe('#VALUE!');
  });

  it('computes each SUBTOTAL function', () => {
    expect(calc('SUBTOTAL(1,S!B1:B5)')).toBe(30);
    expect(calc('SUBTOTAL(2,S!A1:A5)')).toBe(2);
    expect(calc('SUBTOTAL(3,S!A1:A5)')).toBe(4);
    expect(calc('SUBTOTAL(4,S!B1:B5)')).toBe(50);
    expect(calc('SUBTOTAL(5,S!B1:B5)')).toBe(10);
    expect(calc('SUBTOTAL(6,S!B1:B3)')).toBe(6000);
    expect(calc('SUBTOTAL(9,S!B1:B5)')).toBe(150);
  });

  it('propagates errors, having no option to ignore them', () => {
    expect(code('SUBTOTAL(9,S!C1:C3)')).toBe('#DIV/0!');
  });

  it('never reports an error from COUNT or COUNTA, which do not propagate one', () => {
    // COUNT ignores error values in a reference; COUNTA counts them as values.
    // Neither has any way to return one, whatever the range holds.
    expect(calc('SUBTOTAL(2,S!C1:C3)')).toBe(2);
    expect(calc('SUBTOTAL(3,S!C1:C3)')).toBe(3);
    expect(calc('SUBTOTAL(102,S!C1:C3)')).toBe(2);
    expect(calc('AGGREGATE(2,0,S!C1:C3)')).toBe(2);
    expect(calc('AGGREGATE(3,0,S!C1:C3)')).toBe(3);
    // Ignoring error values drops them from COUNTA's tally as well.
    expect(calc('AGGREGATE(3,6,S!C1:C3)')).toBe(2);
  });

  it('takes MAX and MIN over more values than a call can hold as arguments', () => {
    expect(calc('SUBTOTAL(4,Wide!A1:A200000)')).toBe(100002);
    expect(calc('AGGREGATE(5,0,Wide!A1:A200000)')).toBe(0);
  });

  it('ignores error values only for the options that say so', () => {
    expect(calc('AGGREGATE(9,6,S!C1:C3)')).toBe(12);
    expect(calc('AGGREGATE(9,2,S!C1:C3)')).toBe(12);
    expect(calc('AGGREGATE(9,3,S!C1:C3)')).toBe(12);
    expect(calc('AGGREGATE(9,7,S!C1:C3)')).toBe(12);
    expect(code('AGGREGATE(9,0,S!C1:C3)')).toBe('#DIV/0!');
    expect(code('AGGREGATE(9,1,S!C1:C3)')).toBe('#DIV/0!');
    expect(code('AGGREGATE(9,4,S!C1:C3)')).toBe('#DIV/0!');
    expect(code('AGGREGATE(9,5,S!C1:C3)')).toBe('#DIV/0!');
  });

  it('ignores an error handed in directly, which needs an error-transparent parameter', () => {
    expect(calc('AGGREGATE(9,6,1/0,5)')).toBe(5);
  });

  it('serves the array-form functions from a k argument', () => {
    expect(calc('AGGREGATE(14,6,S!B1:B5,2)')).toBe(40);
    expect(calc('AGGREGATE(15,6,S!B1:B5,2)')).toBe(20);
    expect(calc('AGGREGATE(16,6,S!B1:B5,0.25)')).toBe(20);
    expect(calc('AGGREGATE(17,6,S!B1:B5,1)')).toBe(20);
    expect(calc('AGGREGATE(18,6,S!B1:B5,0.5)')).toBe(30);
    expect(calc('AGGREGATE(19,6,S!B1:B5,2)')).toBe(30);
  });

  it('reports #VALUE! for a missing k and #NUM! for one out of range', () => {
    expect(code('AGGREGATE(14,6,S!B1:B5)')).toBe('#VALUE!');
    expect(code('AGGREGATE(14,6,S!B1:B5,9)')).toBe('#NUM!');
    expect(code('AGGREGATE(18,6,S!B1:B5,0.01)')).toBe('#NUM!');
  });

  it('rejects an unknown function number or option', () => {
    expect(code('AGGREGATE(20,0,S!B1:B5)')).toBe('#VALUE!');
    expect(code('AGGREGATE(9,8,S!B1:B5)')).toBe('#VALUE!');
  });

  it('reports #DIV/0! for an average or deviation with nothing to work on', () => {
    expect(code('AGGREGATE(1,6,S!Z1:Z9)')).toBe('#DIV/0!');
    expect(code('AGGREGATE(7,6,S!B1)')).toBe('#DIV/0!');
  });

  it('returns #N/A from MODE.SNGL when no value repeats', () => {
    expect(code('AGGREGATE(13,6,S!B1:B5)')).toBe('#N/A');
    expect(calc('AGGREGATE(13,6,{1,2,2,3})')).toBe(2);
  });
});

describe('sign, truncation and rounding', () => {
  it('ABS and SIGN', () => {
    expect(calc('ABS(-7)')).toBe(7);
    expect(calc('ABS(0)')).toBe(0);
    expect(calc('SIGN(-7)')).toBe(-1);
    expect(calc('SIGN(0)')).toBe(0);
    expect(calc('SIGN(0.0001)')).toBe(1);
  });

  it('INT floors and TRUNC chops', () => {
    expect(calc('INT(-3.5)')).toBe(-4);
    expect(calc('INT(3.9)')).toBe(3);
    expect(calc('TRUNC(-3.5)')).toBe(-3);
    expect(calc('TRUNC(3.9)')).toBe(3);
    expect(calc('TRUNC(3.14159,2)')).toBe(3.14);
    expect(calc('TRUNC(-3.14159,2)')).toBe(-3.14);
    expect(calc('TRUNC(1234,-2)')).toBe(1200);
  });

  it('rounds half away from zero in both directions', () => {
    expect(calc('ROUND(2.5,0)')).toBe(3);
    expect(calc('ROUND(-2.5,0)')).toBe(-3);
    expect(calc('ROUND(0.5,0)')).toBe(1);
    expect(calc('ROUND(1.5,0)')).toBe(2);
    expect(calc('ROUND(-1.5,0)')).toBe(-2);
  });

  it('rounds the fifteen-digit value, not the raw double', () => {
    // 2.675 is stored as 2.67499999999999982; Excel still answers 2.68.
    expect(calc('ROUND(2.675,2)')).toBe(2.68);
    expect(calc('ROUNDUP(1.1*3,1)')).toBe(3.3);
  });

  it('rounds at negative digit positions', () => {
    expect(calc('ROUND(1234.5678,-2)')).toBe(1200);
    expect(calc('ROUND(1250,-3)')).toBe(1000);
    expect(calc('ROUND(1500,-3)')).toBe(2000);
    expect(calc('ROUND(123,-10)')).toBe(0);
  });

  it('rounds away from and towards zero on demand', () => {
    expect(calc('ROUNDUP(3.14159,2)')).toBe(3.15);
    expect(calc('ROUNDUP(-3.14159,2)')).toBe(-3.15);
    expect(calc('ROUNDDOWN(3.99,1)')).toBe(3.9);
    expect(calc('ROUNDDOWN(-3.99,1)')).toBe(-3.9);
    expect(calc('ROUNDUP(4,0)')).toBe(4);
  });

  it('leaves zero and very small magnitudes alone', () => {
    expect(calc('ROUND(0,5)')).toBe(0);
    expect(calc('ROUND(0.00001,2)')).toBe(0);
    expect(calc('ROUND(1E-300,2)')).toBe(0);
  });

  it('MROUND takes the nearest multiple, half away from zero', () => {
    expect(calc('MROUND(17,5)')).toBe(15);
    expect(calc('MROUND(7.5,5)')).toBe(10);
    expect(calc('MROUND(-17,-5)')).toBe(-15);
    expect(calc('MROUND(-7.5,-5)')).toBe(-10);
    expect(calc('MROUND(1.3,0.1)')).toBe(1.3);
    expect(calc('MROUND(5,0)')).toBe(0);
  });

  it('MROUND refuses mismatched signs', () => {
    expect(code('MROUND(17,-5)')).toBe('#NUM!');
    expect(code('MROUND(-17,5)')).toBe('#NUM!');
  });

  it('EVEN and ODD round away from zero', () => {
    expect(calc('EVEN(1.5)')).toBe(2);
    expect(calc('EVEN(3)')).toBe(4);
    expect(calc('EVEN(2)')).toBe(2);
    expect(calc('EVEN(-1.5)')).toBe(-2);
    expect(calc('EVEN(0)')).toBe(0);
    expect(calc('ODD(1.5)')).toBe(3);
    expect(calc('ODD(2)')).toBe(3);
    expect(calc('ODD(1)')).toBe(1);
    expect(calc('ODD(-2)')).toBe(-3);
    expect(calc('ODD(0)')).toBe(1);
  });
});

describe('CEILING and FLOOR', () => {
  it('follows the legacy sign rules', () => {
    expect(calc('CEILING(2.5,1)')).toBe(3);
    expect(calc('CEILING(-2.5,-2)')).toBe(-4);
    expect(calc('CEILING(-2.5,2)')).toBe(-2);
    expect(calc('CEILING(1.5,0.1)')).toBe(1.5);
    expect(calc('CEILING(0.234,0.01)')).toBe(0.24);
    expect(calc('FLOOR(3.7,2)')).toBe(2);
    expect(calc('FLOOR(-2.5,-2)')).toBe(-2);
    expect(calc('FLOOR(-2.5,2)')).toBe(-4);
    expect(calc('FLOOR(1.58,0.1)')).toBe(1.5);
    expect(calc('FLOOR(0.234,0.01)')).toBe(0.23);
  });

  it('rejects a positive number with a negative significance', () => {
    expect(code('CEILING(2.5,-2)')).toBe('#NUM!');
    expect(code('FLOOR(2.5,-2)')).toBe('#NUM!');
  });

  it('splits on a zero significance the way Excel does', () => {
    expect(calc('CEILING(5,0)')).toBe(0);
    expect(code('FLOOR(5,0)')).toBe('#DIV/0!');
    expect(calc('CEILING(0,5)')).toBe(0);
    expect(calc('FLOOR(0,5)')).toBe(0);
  });

  it('rounds the quotient at fifteen digits, so exact multiples stay put', () => {
    expect(calc('FLOOR(0.29,0.01)')).toBe(0.29);
    expect(calc('CEILING(0.29,0.01)')).toBe(0.29);
  });

  it('CEILING.MATH and FLOOR.MATH default to significance 1', () => {
    expect(calc('CEILING.MATH(4.2)')).toBe(5);
    expect(calc('CEILING.MATH(-4.2)')).toBe(-4);
    expect(calc('FLOOR.MATH(4.8)')).toBe(4);
    expect(calc('FLOOR.MATH(-4.8)')).toBe(-5);
  });

  it('CEILING.MATH and FLOOR.MATH flip direction for negatives when mode is set', () => {
    expect(calc('CEILING.MATH(-4.2,1,-1)')).toBe(-5);
    expect(calc('FLOOR.MATH(-4.2,1,-1)')).toBe(-4);
    // The mode is inert for positive numbers.
    expect(calc('CEILING.MATH(4.2,1,-1)')).toBe(5);
    expect(calc('FLOOR.MATH(4.2,1,-1)')).toBe(4);
  });

  it('ignores the sign of significance in the modern forms', () => {
    expect(calc('CEILING.MATH(-5.5,-2)')).toBe(-4);
    expect(calc('FLOOR.MATH(-5.5,-2)')).toBe(-6);
    expect(calc('CEILING.PRECISE(-4.2,-2)')).toBe(-4);
    expect(calc('FLOOR.PRECISE(-4.2,-2)')).toBe(-6);
    expect(calc('CEILING.PRECISE(4.2,2)')).toBe(6);
    expect(calc('FLOOR.PRECISE(4.2,2)')).toBe(4);
  });

  it('returns zero for a zero significance in the modern forms', () => {
    expect(calc('CEILING.MATH(5,0)')).toBe(0);
    expect(calc('FLOOR.MATH(5,0)')).toBe(0);
  });
});

describe('modular and power arithmetic', () => {
  it('MOD takes the sign of the divisor', () => {
    expect(calc('MOD(-7,3)')).toBe(2);
    expect(calc('MOD(7,-3)')).toBe(-2);
    expect(calc('MOD(7,3)')).toBe(1);
    expect(calc('MOD(-7,-3)')).toBe(-1);
    expect(calc('MOD(6,3)')).toBe(0);
  });

  it('MOD divides by zero into #DIV/0!', () => {
    expect(code('MOD(7,0)')).toBe('#DIV/0!');
  });

  it('MOD stays exact on fractional divisors', () => {
    expect(calc('MOD(1.1,0.1)')).toBe(0);
    expect(calc('MOD(5.5,2)')).toBe(1.5);
  });

  it('QUOTIENT truncates towards zero', () => {
    expect(calc('QUOTIENT(7,2)')).toBe(3);
    expect(calc('QUOTIENT(-7,2)')).toBe(-3);
    expect(code('QUOTIENT(7,0)')).toBe('#DIV/0!');
  });

  it('POWER handles the awkward bases', () => {
    expect(calc('POWER(2,10)')).toBe(1024);
    expect(calc('POWER(-8,2)')).toBe(64);
    expect(code('POWER(-8,0.5)')).toBe('#NUM!');
    expect(code('POWER(0,0)')).toBe('#NUM!');
    expect(code('POWER(0,-1)')).toBe('#DIV/0!');
    expect(calc('POWER(0,2)')).toBe(0);
    expect(code('POWER(10,400)')).toBe('#NUM!');
  });

  it('agrees with the ^ operator on every awkward base', () => {
    // Excel documents ^ as another spelling of POWER, so the two cannot differ.
    for (const [fn, op] of [
      ['POWER(0,0)', '0^0'],
      ['POWER(0,-1)', '0^-1'],
      ['POWER(-8,0.5)', '(-8)^0.5'],
      ['POWER(2,10)', '2^10'],
      ['POWER(0,2)', '0^2'],
    ] as const) {
      expect(calc(op), op).toEqual(calc(fn));
    }
    expect(code('0^0')).toBe('#NUM!');
    expect(code('0^-1')).toBe('#DIV/0!');
  });

  it('SQRT and SQRTPI reject negatives', () => {
    expect(calc('SQRT(144)')).toBe(12);
    expect(calc('SQRT(0)')).toBe(0);
    expect(code('SQRT(-1)')).toBe('#NUM!');
    expect(num('SQRTPI(4)')).toBe(toExcelPrecision(Math.sqrt(4 * Math.PI)));
    expect(code('SQRTPI(-1)')).toBe('#NUM!');
  });

  it('EXP, LN, LOG10 and LOG', () => {
    expect(num('LN(EXP(2))')).toBe(2);
    expect(num('LOG10(1000)')).toBe(3);
    expect(num('LOG(8,2)')).toBe(3);
    expect(num('LOG(100)')).toBe(2);
    expect(code('EXP(1000)')).toBe('#NUM!');
    expect(code('LN(0)')).toBe('#NUM!');
    expect(code('LN(-1)')).toBe('#NUM!');
    expect(code('LOG10(0)')).toBe('#NUM!');
    expect(code('LOG(8,1)')).toBe('#DIV/0!');
    expect(code('LOG(8,0)')).toBe('#NUM!');
    expect(code('LOG(8,-2)')).toBe('#NUM!');
  });
});

describe('trigonometry', () => {
  it('the circular functions', () => {
    expect(num('SIN(PI()/2)')).toBe(1);
    expect(num('COS(0)')).toBe(1);
    expect(num('TAN(0)')).toBe(0);
    expect(num('ASIN(1)')).toBe(toExcelPrecision(Math.PI / 2));
    expect(num('ACOS(1)')).toBe(0);
    expect(num('ATAN(1)')).toBe(toExcelPrecision(Math.PI / 4));
  });

  it('ATAN2 takes x before y, unlike the C library', () => {
    expect(num('ATAN2(1,1)')).toBe(toExcelPrecision(Math.PI / 4));
    expect(num('ATAN2(-1,0)')).toBe(toExcelPrecision(Math.PI));
    expect(num('ATAN2(0,1)')).toBe(toExcelPrecision(Math.PI / 2));
    expect(code('ATAN2(0,0)')).toBe('#DIV/0!');
  });

  it('rejects arguments outside the principal domains', () => {
    expect(code('ASIN(2)')).toBe('#NUM!');
    expect(code('ACOS(-1.5)')).toBe('#NUM!');
    expect(code('ACOSH(0.5)')).toBe('#NUM!');
    expect(code('ATANH(1)')).toBe('#NUM!');
    expect(code('ATANH(-1)')).toBe('#NUM!');
  });

  it('the hyperbolic functions', () => {
    expect(num('SINH(0)')).toBe(0);
    expect(num('COSH(0)')).toBe(1);
    expect(num('TANH(0)')).toBe(0);
    expect(num('ASINH(0)')).toBe(0);
    expect(num('ACOSH(1)')).toBe(0);
    expect(num('ATANH(0)')).toBe(0);
  });

  it('the reciprocal functions and their poles', () => {
    expect(num('SEC(0)')).toBe(1);
    expect(num('CSC(PI()/2)')).toBe(1);
    expect(num('COT(PI()/4)')).toBe(1);
    expect(num('SECH(0)')).toBe(1);
    expect(code('CSC(0)')).toBe('#DIV/0!');
    expect(code('COT(0)')).toBe('#DIV/0!');
    expect(code('CSCH(0)')).toBe('#DIV/0!');
    expect(code('COTH(0)')).toBe('#DIV/0!');
  });

  it('COTH saturates at plus and minus one instead of overflowing', () => {
    // cosh(1000)/sinh(1000) is Infinity/Infinity; the answer is 1.
    expect(num('COTH(1000)')).toBe(1);
    expect(num('COTH(-1000)')).toBe(-1);
    expect(num('COTH(20)')).toBe(1);
  });

  it('ACOT returns the principal value between 0 and pi', () => {
    expect(num('ACOT(0)')).toBe(toExcelPrecision(Math.PI / 2));
    expect(num('ACOT(1)')).toBe(toExcelPrecision(Math.PI / 4));
    // atan(1/x) would answer -pi/4 here; the principal value is 3pi/4.
    expect(num('ACOT(-1)')).toBe(toExcelPrecision((3 * Math.PI) / 4));
  });

  it('converts between degrees and radians', () => {
    expect(num('DEGREES(PI())')).toBe(180);
    expect(num('RADIANS(180)')).toBe(toExcelPrecision(Math.PI));
    expect(num('DEGREES(RADIANS(45))')).toBe(45);
  });
});

describe('integer and combinatorial functions', () => {
  it('GCD and LCM', () => {
    expect(calc('GCD(24,36)')).toBe(12);
    expect(calc('GCD(0,0)')).toBe(0);
    expect(calc('GCD(0,5)')).toBe(5);
    expect(calc('GCD(S!B1:B5)')).toBe(10);
    expect(calc('LCM(4,6)')).toBe(12);
    expect(calc('LCM(1,0)')).toBe(0);
    expect(calc('LCM(2,3,4)')).toBe(12);
  });

  it('truncates non-integers and refuses negatives', () => {
    expect(calc('GCD(5.9,3)')).toBe(1);
    expect(calc('GCD(24.9,36.9)')).toBe(12);
    expect(code('GCD(-4,2)')).toBe('#NUM!');
    expect(code('LCM(-4,2)')).toBe('#NUM!');
    // -0.5 is below zero even though Math.trunc would make it -0.
    expect(code('GCD(-0.5,4)')).toBe('#NUM!');
    expect(code('LCM(-0.5,4)')).toBe('#NUM!');
  });

  it('FACT and FACTDOUBLE', () => {
    expect(calc('FACT(6)')).toBe(720);
    expect(calc('FACT(0)')).toBe(1);
    expect(calc('FACT(6.9)')).toBe(720);
    expect(code('FACT(-1)')).toBe('#NUM!');
    expect(code('FACT(171)')).toBe('#NUM!');
    expect(calc('FACTDOUBLE(7)')).toBe(105);
    expect(calc('FACTDOUBLE(6)')).toBe(48);
    expect(calc('FACTDOUBLE(0)')).toBe(1);
    expect(code('FACTDOUBLE(-1)')).toBe('#NUM!');
  });

  it('answers #NUM! for a huge argument instead of counting up to it', () => {
    // Each of these overflows a double long before its loop would end, so the
    // answer has to arrive without running the loop to completion.
    expect(code('FACT(1E15)')).toBe('#NUM!');
    expect(code('FACTDOUBLE(1E15)')).toBe('#NUM!');
    expect(code('COMBIN(1E15,5E14)')).toBe('#NUM!');
    expect(code('PERMUT(1E15,1E15)')).toBe('#NUM!');
  });

  it('COMBIN and COMBINA', () => {
    expect(calc('COMBIN(10,3)')).toBe(120);
    expect(calc('COMBIN(52,5)')).toBe(2598960);
    expect(calc('COMBIN(0,0)')).toBe(1);
    expect(code('COMBIN(3,10)')).toBe('#NUM!');
    expect(code('COMBIN(-1,1)')).toBe('#NUM!');
    expect(calc('COMBINA(4,3)')).toBe(20);
    expect(calc('COMBINA(0,0)')).toBe(1);
    expect(code('COMBINA(0,1)')).toBe('#NUM!');
    expect(code('COMBINA(4,-1)')).toBe('#NUM!');
  });

  it('PERMUT and PERMUTATIONA', () => {
    expect(calc('PERMUT(5,2)')).toBe(20);
    expect(calc('PERMUT(3,3)')).toBe(6);
    expect(code('PERMUT(2,5)')).toBe('#NUM!');
    // Microsoft documents #NUM! when number is less than *or equal to* zero,
    // which is where PERMUT parts company with COMBIN(0,0).
    expect(code('PERMUT(0,0)')).toBe('#NUM!');
    expect(code('PERMUT(-1,0)')).toBe('#NUM!');
    expect(calc('PERMUTATIONA(3,2)')).toBe(9);
    expect(calc('PERMUTATIONA(0,0)')).toBe(1);
    expect(code('PERMUTATIONA(-1,2)')).toBe('#NUM!');
  });
});

describe('random numbers', () => {
  it('RAND stays in [0,1)', () => {
    for (let i = 0; i < 50; i++) {
      const v = calc('RAND()');
      expect(typeof v).toBe('number');
      expect(v as number).toBeGreaterThanOrEqual(0);
      expect(v as number).toBeLessThan(1);
    }
  });

  it('RANDBETWEEN covers its inclusive bounds', () => {
    const seen = new Set<number>();
    for (let i = 0; i < 200; i++) seen.add(calc('RANDBETWEEN(1,3)') as number);
    expect([...seen].sort()).toEqual([1, 2, 3]);
    expect(calc('RANDBETWEEN(4,4)')).toBe(4);
    expect(code('RANDBETWEEN(5,1)')).toBe('#NUM!');
  });

  it('marks only the two random functions volatile', () => {
    const volatiles = MATH_FUNCTIONS.filter((f) => f.volatile).map((f) => f.name);
    expect(volatiles.sort()).toEqual(['RAND', 'RANDBETWEEN']);
  });
});

describe('ROMAN and ARABIC', () => {
  it('walks Microsoft\'s ladder of forms for 1999', () => {
    expect(calc('ROMAN(1999,0)')).toBe('MCMXCIX');
    expect(calc('ROMAN(1999,1)')).toBe('MLMVLIV');
    expect(calc('ROMAN(1999,2)')).toBe('MXMIX');
    expect(calc('ROMAN(1999,3)')).toBe('MVMIV');
    expect(calc('ROMAN(1999,4)')).toBe('MIM');
  });

  it('and the same ladder for 499', () => {
    expect(calc('ROMAN(499,0)')).toBe('CDXCIX');
    expect(calc('ROMAN(499,1)')).toBe('LDVLIV');
    expect(calc('ROMAN(499,2)')).toBe('XDIX');
    expect(calc('ROMAN(499,3)')).toBe('VDIV');
    expect(calc('ROMAN(499,4)')).toBe('ID');
  });

  it('defaults to the classic form and handles the boundaries', () => {
    expect(calc('ROMAN(499)')).toBe('CDXCIX');
    expect(calc('ROMAN(0)')).toBe('');
    expect(calc('ROMAN(3999)')).toBe('MMMCMXCIX');
    expect(calc('ROMAN(4.9)')).toBe('IV');
    expect(code('ROMAN(4000)')).toBe('#VALUE!');
    expect(code('ROMAN(-1)')).toBe('#VALUE!');
    expect(code('ROMAN(5,5)')).toBe('#VALUE!');
  });

  it('reads TRUE as classic and FALSE as simplified', () => {
    expect(calc('ROMAN(1999,TRUE)')).toBe('MCMXCIX');
    expect(calc('ROMAN(1999,FALSE)')).toBe('MIM');
  });

  it('ARABIC inverts the classic forms and tolerates the concise ones', () => {
    expect(calc('ARABIC("MCMXCIX")')).toBe(1999);
    expect(calc('ARABIC("MIM")')).toBe(1999);
    expect(calc('ARABIC("ID")')).toBe(499);
    expect(calc('ARABIC("mcmxcix")')).toBe(1999);
    expect(calc('ARABIC("")')).toBe(0);
    expect(calc('ARABIC("-IV")')).toBe(-4);
    expect(code('ARABIC("Q")')).toBe('#VALUE!');
  });

  it('round-trips every value ROMAN can produce', () => {
    for (let n = 1; n <= 3999; n++) {
      for (const form of [0, 1, 2, 3, 4]) {
        const text = calc(`ROMAN(${n},${form})`);
        expect(calc(`ARABIC("${String(text)}")`), `${n} form ${form} -> ${String(text)}`).toBe(n);
      }
    }
  });
});

describe('BASE and DECIMAL', () => {
  it('converts to a radix', () => {
    expect(calc('BASE(255,16)')).toBe('FF');
    expect(calc('BASE(7,2)')).toBe('111');
    expect(calc('BASE(7,2,8)')).toBe('00000111');
    expect(calc('BASE(0,2)')).toBe('0');
    expect(calc('BASE(35,36)')).toBe('Z');
    expect(calc('BASE(7.9,2)')).toBe('111');
  });

  it('rejects an impossible radix, length or number', () => {
    expect(code('BASE(7,1)')).toBe('#NUM!');
    expect(code('BASE(7,37)')).toBe('#NUM!');
    expect(code('BASE(-1,2)')).toBe('#NUM!');
    expect(code('BASE(7,2,256)')).toBe('#NUM!');
  });

  it('converts from a radix', () => {
    expect(calc('DECIMAL("FF",16)')).toBe(255);
    expect(calc('DECIMAL("ff",16)')).toBe(255);
    expect(calc('DECIMAL("111",2)')).toBe(7);
    expect(calc('DECIMAL("Z",36)')).toBe(35);
    expect(calc('DECIMAL("",10)')).toBe(0);
    expect(code('DECIMAL("2",2)')).toBe('#NUM!');
    expect(code('DECIMAL("FF",10)')).toBe('#NUM!');
    expect(code('DECIMAL("FF",1)')).toBe('#NUM!');
  });

  it('round-trips through both directions', () => {
    for (const radix of [2, 8, 16, 36]) {
      expect(calc(`DECIMAL(BASE(123456,${radix}),${radix})`)).toBe(123456);
    }
  });
});

describe('matrix functions', () => {
  it('TRANSPOSE swaps the axes and keeps every value', () => {
    expect(grid('TRANSPOSE({1,2,3;4,5,6})')).toEqual({
      rows: 3,
      cols: 2,
      data: [1, 4, 2, 5, 3, 6],
    });
    expect(grid('TRANSPOSE(S!A1:B2)')).toEqual({ rows: 2, cols: 2, data: [1, 2, 10, 20] });
  });

  it('MMULT multiplies conformable matrices', () => {
    expect(grid('MMULT({1,2;3,4},{5,6;7,8})')).toEqual({
      rows: 2,
      cols: 2,
      data: [19, 22, 43, 50],
    });
    expect(grid('MMULT({1,2,3},{1;2;3})')).toEqual({ rows: 1, cols: 1, data: [14] });
  });

  it('MMULT reports #VALUE! for non-conformable or non-numeric input', () => {
    expect(code('MMULT({1,2;3,4},{1,2,3})')).toBe('#VALUE!');
    expect(code('MMULT({1,"x";3,4},{5,6;7,8})')).toBe('#VALUE!');
    expect(code('MMULT(S!A1:A2,S!D1:D2)')).toBe('#VALUE!');
  });

  it('MDETERM', () => {
    expect(calc('MDETERM({1,2;3,4})')).toBe(-2);
    expect(calc('MDETERM({1,2,3;4,5,6;7,8,10})')).toBe(-3);
    expect(calc('MDETERM({1,2;2,4})')).toBe(0);
    expect(calc('MDETERM({5})')).toBe(5);
    expect(code('MDETERM({1,2,3;4,5,6})')).toBe('#VALUE!');
  });

  it('MINVERSE, and #NUM! for a singular matrix', () => {
    expect(grid('MINVERSE({1,2;3,4})')).toEqual({
      rows: 2,
      cols: 2,
      data: [-2, 1, 1.5, -0.5],
    });
    expect(grid('MINVERSE({2,0;0,4})')).toEqual({ rows: 2, cols: 2, data: [0.5, 0, 0, 0.25] });
    expect(code('MINVERSE({1,2;2,4})')).toBe('#NUM!');
    expect(code('MINVERSE({1,2,3;4,5,6})')).toBe('#VALUE!');
  });

  it('MMULT of a matrix and its inverse is the identity', () => {
    expect(grid('MMULT({4,7;2,6},MINVERSE({4,7;2,6}))')).toEqual({
      rows: 2,
      cols: 2,
      data: [1, 0, 0, 1],
    });
  });
});

describe('SERIESSUM', () => {
  it('sums a power series', () => {
    expect(calc('SERIESSUM(2,0,1,{1,1,1})')).toBe(7);
    expect(calc('SERIESSUM(3,1,2,{1,1})')).toBe(30);
  });

  it('sums the alternating even powers of a cosine series', () => {
    const x = Math.PI / 4;
    const expected =
      1 - 0.5 * x ** 2 + 0.041666666666666666 * x ** 4 - 0.0013888888888888889 * x ** 6;
    const v = calc('SERIESSUM(PI()/4,0,2,{1,-0.5,0.041666666666666666,-0.0013888888888888889})');
    expect(v as number).toBeCloseTo(expected, 12);
    // Four terms is enough to be recognisably the cosine of pi/4.
    expect(v as number).toBeCloseTo(Math.cos(x), 5);
  });

  it('treats a blank coefficient as zero without shifting the powers', () => {
    expect(calc('SERIESSUM(2,0,1,S!Z1:Z3)')).toBe(0);
  });

  it('reports #VALUE! for a text coefficient', () => {
    expect(code('SERIESSUM(2,0,1,{1,"x"})')).toBe('#VALUE!');
  });
});

describe('registry metadata', () => {
  it('registers every promised function exactly once', () => {
    const names = MATH_FUNCTIONS.map((f) => f.name);
    expect(new Set(names).size).toBe(names.length);
    const expected = [
      'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUMSQ', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2',
      'PRODUCT', 'SUBTOTAL', 'AGGREGATE', 'ABS', 'SIGN', 'INT', 'TRUNC', 'ROUND', 'ROUNDUP',
      'ROUNDDOWN', 'MROUND', 'CEILING', 'CEILING.MATH', 'CEILING.PRECISE', 'FLOOR', 'FLOOR.MATH',
      'FLOOR.PRECISE', 'EVEN', 'ODD', 'MOD', 'QUOTIENT', 'POWER', 'SQRT', 'SQRTPI', 'EXP', 'LN',
      'LOG', 'LOG10', 'PI', 'SIN', 'COS', 'TAN', 'ASIN', 'ACOS', 'ATAN', 'ATAN2', 'SINH', 'COSH',
      'TANH', 'ASINH', 'ACOSH', 'ATANH', 'SEC', 'CSC', 'COT', 'ACOT', 'SECH', 'CSCH', 'COTH',
      'DEGREES', 'RADIANS', 'GCD', 'LCM', 'FACT', 'FACTDOUBLE', 'COMBIN', 'COMBINA', 'PERMUT',
      'PERMUTATIONA', 'RAND', 'RANDBETWEEN', 'ROMAN', 'ARABIC', 'BASE', 'DECIMAL', 'TRANSPOSE',
      'MMULT', 'MDETERM', 'MINVERSE', 'SERIESSUM',
    ];
    expect(names.slice().sort()).toEqual(expected.slice().sort());
  });

  it('claims no structural dependency, since none of these read sheet shape', () => {
    expect(MATH_FUNCTIONS.filter((f) => f.structural)).toEqual([]);
  });

  it('marks the post-2007 names as future functions once registered', () => {
    for (const name of ['CEILING.MATH', 'FLOOR.PRECISE', 'AGGREGATE', 'BASE', 'ACOT', 'COMBINA']) {
      expect(registry.get(name)!.futureFunction, name).toBe(true);
    }
    expect(registry.get('SUM')!.futureFunction).toBe(false);
  });

  it('resolves a stored _xlfn. prefix back to the function', () => {
    expect(registry.get('_xlfn.CEILING.MATH')!.name).toBe('CEILING.MATH');
  });

  it('declares broadcast on the elementwise scalar functions and not on the aggregates', () => {
    expect(registry.get('SIN')!.broadcast).toBe(true);
    expect(registry.get('ROUND')!.broadcast).toBe(true);
    expect(registry.get('SUM')!.broadcast).toBeUndefined();
    expect(registry.get('SUMPRODUCT')!.broadcast).toBeUndefined();
    expect(registry.get('TRANSPOSE')!.broadcast).toBeUndefined();
  });
});

describe('error propagation', () => {
  it('passes an error argument straight through the scalar functions', () => {
    for (const f of ['ABS(1/0)', 'SIN(1/0)', 'ROUND(1/0,2)', 'ROUND(1,1/0)', 'MOD(1/0,2)']) {
      expect(code(f)).toBe('#DIV/0!');
    }
  });

  it('reports #VALUE! for text that is not a number', () => {
    for (const f of ['ABS("x")', 'ROUND("x",0)', 'POWER("x",2)', 'GCD("x")']) {
      expect(code(f)).toBe('#VALUE!');
    }
  });

  it('accepts text that looks like a number', () => {
    expect(calc('ABS("-7")')).toBe(7);
    expect(calc('ROUND("2.5",0)')).toBe(3);
    expect(calc('POWER("2","10")')).toBe(1024);
  });

  it('treats a blank cell as zero', () => {
    expect(calc('ABS(S!Z1)')).toBe(0);
    expect(calc('SUM(S!Z1,5)')).toBe(5);
  });
});
