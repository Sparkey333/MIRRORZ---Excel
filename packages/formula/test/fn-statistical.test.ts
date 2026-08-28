/**
 * Statistical functions.
 *
 * Three layers of assertion, in decreasing order of authority.
 *
 * The first replays fixtures/generated/formulas.calc.xlsx: sixteen of its cases
 * are formulas this module owns outright, and each is re-parsed and re-evaluated
 * here against the value the oracle actually computed.
 *
 * The second is a table of distribution values produced by running the same
 * LibreOffice recalculation pipeline over a probe workbook - the mechanism that
 * builds fixtures/generated, pointed at the fifty-odd functions the shipped
 * fixtures happen not to exercise. The numbers are quoted at the fifteen digits
 * the oracle stored and compared with a relative tolerance of 1e-13, since the
 * two implementations agree to about 1e-14 and the fifteenth digit of a stored
 * decimal is not a meaningful assertion.
 *
 * The third is the places where LibreOffice is known to disagree with Excel, and
 * where these tests therefore assert Excel's answer against the documentation
 * rather than the oracle's. There are five, all verified deliberately rather
 * than discovered by a failing test:
 *
 *   - AVERAGE, COUNT and friends over a range holding TRUE. Excel ignores
 *     logical values inside a reference; LibreOffice counts TRUE as 1.
 *   - AVERAGE(1,"3"). Excel coerces text typed directly into the argument list
 *     and answers 2; LibreOffice answers #VALUE!.
 *   - MODE.SNGL with no repeated value. Excel documents #N/A; LibreOffice
 *     reports #VALUE!.
 *   - PERCENTRANK truncates to the requested number of places, so a rank of
 *     five ninths is 0.555 (this is the value in Microsoft's own help page);
 *     LibreOffice rounds and answers 0.556.
 *   - Almost every out-of-domain argument. Excel documents #NUM! for these and
 *     LibreOffice reports #VALUE! nearly across the board, so the error codes
 *     below follow the Microsoft function reference, not the oracle.
 *
 * A fourth, smaller layer checks the special functions against independently
 * computed high-precision values, because they are shared by forty
 * distributions and an error in one of them would otherwise be spread thinly
 * enough to look like rounding everywhere.
 *
 * The registry under test holds STATISTICAL_FUNCTIONS alone, so a failure here
 * is never another category's fault.
 */

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, type Scalar, Workbook, isError } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { Evaluator } from '../src/evaluator.js';
import { FUNCTION_ALIASES } from '../src/functions/index.js';
import { STATISTICAL_FUNCTIONS } from '../src/functions/statistical.js';
import { parseFormula } from '../src/parser.js';
import { FUTURE_FUNCTIONS, FunctionRegistry } from '../src/registry.js';
import { WorkbookStore } from '../src/store.js';
import { type Value, isArray, toExcelPrecision } from '../src/value.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

const registry = new FunctionRegistry().registerAll(STATISTICAL_FUNCTIONS);

// ---------------------------------------------------------------------------
// Layer one: the shipped oracle
// ---------------------------------------------------------------------------

const { workbook: oracleBook } = readXlsx(
  new Uint8Array(readFileSync(new URL('formulas.calc.xlsx', FIXTURES))),
);
const oracleSheet = oracleBook.getSheet('Formulas')!;
const oracleEval = new Evaluator(new WorkbookStore(oracleBook), registry, {
  dateSystem: oracleBook.dateSystem,
});
const oracleCases = new Map<string, { formula: string; value: Scalar; row: number }>();
for (const { row, col, cell } of oracleSheet.entries()) {
  if (col !== 2 || !cell.formula) continue;
  const name = oracleSheet.getValue(row, 0);
  if (typeof name === 'string') oracleCases.set(name, { formula: cell.formula, value: cell.value, row });
}

function sameScalar(actual: Value, expected: Scalar): void {
  if (typeof expected === 'number' && typeof actual === 'number') {
    expect(toExcelPrecision(actual)).toBe(toExcelPrecision(expected));
    return;
  }
  expect(actual).toEqual(expected);
}

describe('oracle: formulas.calc.xlsx', () => {
  const cases = [
    'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'COUNTBLANK', 'MEDIAN',
    'STDEV.S', 'VAR.S', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF',
    'LARGE', 'SMALL', 'RANK', 'PCT_RANK',
  ];
  for (const name of cases) {
    it(`reproduces ${name}`, () => {
      const c = oracleCases.get(name);
      expect(c, `case ${name} missing from the fixture`).toBeDefined();
      const value = oracleEval.evaluate({
        ast: parseFormula(c!.formula, { origin: { row: c!.row, col: 2 } }),
        sheet: 'Formulas',
        row: c!.row,
        col: 2,
      });
      sameScalar(value, c!.value);
    });
  }

  it('reads the salary column the same way from every aggregate', () => {
    // The oracle's own numbers cross-check each other: the mean is the sum over
    // the count, and the sample variance is the sum of squared deviations over
    // n-1. If the traversal skipped or double-counted a cell, these would part
    // company even though each function alone might look plausible.
    const average = oracleCases.get('AVERAGE')!.value as number;
    const count = oracleCases.get('COUNT')!.value as number;
    const variance = oracleCases.get('VAR.S')!.value as number;
    const stdev = oracleCases.get('STDEV.S')!.value as number;
    expect(count).toBe(8);
    expect(toExcelPrecision(average * count)).toBe(1_207_000);
    expect(toExcelPrecision(Math.sqrt(variance))).toBe(toExcelPrecision(stdev));
  });
});

// ---------------------------------------------------------------------------
// A hand-built workbook, laid out to match the probe the oracle table below was
// generated from.
// ---------------------------------------------------------------------------

const book = new Workbook();
const d = book.addSheet('D');
book.addSheet('Calc');

// A: the working sample, with 4 repeated so there is a mode.
[4, 5, 8, 7, 11, 4, 3, 9].forEach((v, i) => d.setValue(i, 0, v));
// B: a parallel numeric column.
[2, 4, 6, 8, 10, 12, 14, 16].forEach((v, i) => d.setValue(i, 1, v));
// C: a criteria column.
['a', 'b', 'a', 'c'].forEach((v, i) => d.setValue(i, 2, v));
// D: all distinct, so MODE has nothing to report.
[1, 2, 3, 4, 5, 6, 7, 8].forEach((v, i) => d.setValue(i, 3, v));
// E: the mixed column that separates AVERAGE from AVERAGEA. E4 is a real empty
// string, not a blank cell.
d.setValue(0, 4, 10);
d.setValue(1, 4, 'x');
d.setValue(2, 4, true);
d.setValue(3, 4, '');
d.setValue(4, 4, 20);
// F is left entirely blank.
// G and H: a nearly linear relationship, for the regression family.
[2.1, 4.3, 5.9, 8.4, 10.1].forEach((v, i) => d.setValue(i, 6, v));
[1, 2, 3, 4, 5].forEach((v, i) => d.setValue(i, 7, v));
// I: a second independent variable.
[3, 1, 4, 1, 5].forEach((v, i) => d.setValue(i, 8, v));
// K: an error sitting in the middle of a range.
d.setValue(0, 10, 1);
d.setValue(1, 10, CellError.DIV0);
d.setValue(2, 10, 3);
// L: a formula-produced empty string beside a number.
d.setFormula(0, 11, 'IF(TRUE,"","")', '');
d.setValue(1, 11, 5);
// M: numeric-looking text and text that is not.
d.setValue(0, 12, '7');
d.setValue(1, 12, 'abc');

const ev = new Evaluator(new WorkbookStore(book), registry, {});

function calc(formula: string): Value {
  return ev.evaluate({
    ast: parseFormula(formula, { origin: { row: 0, col: 0 } }),
    sheet: 'Calc',
    row: 0,
    col: 0,
  });
}

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

/** Compare against a reference value at a stated relative tolerance. */
function close(formula: string, expected: number, tolerance = 1e-13): void {
  const actual = calc(formula);
  if (typeof actual !== 'number') {
    throw new Error(`${formula} produced ${String(actual)}, not a number`);
  }
  const slack = Math.abs(expected) * tolerance;
  expect(Math.abs(actual - expected), `${formula} gave ${actual}, expected ${expected}`)
    .toBeLessThanOrEqual(slack);
}

// ---------------------------------------------------------------------------

describe('COUNT, COUNTA and COUNTBLANK are three different questions', () => {
  it('counts only numbers with COUNT', () => {
    expect(calc('COUNT(D!E1:E5)')).toBe(2);
    expect(calc('COUNT(D!C1:C4)')).toBe(0);
  });

  it('ignores logical values sitting in a range, as Excel does', () => {
    // E3 holds TRUE. Excel skips it; LibreOffice counts it as 1, which is why
    // this assertion follows the documentation rather than the oracle.
    expect(calc('COUNT(D!E1:E5)')).toBe(2);
    expect(calc('AVERAGE(D!E1:E5)')).toBe(15);
  });

  it('coerces the same values when they are typed directly', () => {
    expect(calc('COUNT(TRUE)')).toBe(1);
    expect(calc('COUNT("3")')).toBe(1);
    expect(calc('COUNT("abc")')).toBe(0);
    expect(calc('AVERAGE(1,"3")')).toBe(2);
  });

  it('counts everything non-empty with COUNTA, error values included', () => {
    expect(calc('COUNTA(D!E1:E5)')).toBe(5);
    expect(calc('COUNTA(D!K1:K3)')).toBe(3);
    expect(calc('COUNT(D!K1:K3)')).toBe(2);
  });

  it('counts a formula-produced empty string as present but blank', () => {
    // L1 is =IF(TRUE,"",""). COUNTA sees a value, COUNTBLANK sees a blank, and
    // both are right at the same time.
    expect(calc('COUNTA(D!L1:L2)')).toBe(2);
    expect(calc('COUNTBLANK(D!L1:L2)')).toBe(1);
    expect(calc('COUNT(D!L1:L2)')).toBe(1);
  });

  it('counts the cells a sparse store never holds', () => {
    expect(calc('COUNTBLANK(D!F1:F5)')).toBe(5);
    expect(calc('COUNTBLANK(D!C1:C8)')).toBe(4);
  });

  it('counts a directly supplied empty argument', () => {
    expect(calc('COUNT(1,)')).toBe(2);
    expect(calc('COUNTA(1,)')).toBe(2);
  });

  it('is zero over an entirely blank range', () => {
    expect(calc('COUNT(D!F1:F9)')).toBe(0);
    expect(calc('COUNTA(D!F1:F9)')).toBe(0);
  });

  it('propagates an error only where the function looks at values', () => {
    expect(code('AVERAGE(D!K1:K3)')).toBe('#DIV/0!');
    expect(code('MAX(D!K1:K3)')).toBe('#DIV/0!');
    expect(calc('COUNT(D!K1:K3)')).toBe(2);
    expect(calc('COUNTIF(D!K1:K3,">0")')).toBe(2);
  });
});

describe('the A-suffixed variants', () => {
  it('count text as zero and booleans as one where the plain form skips them', () => {
    // E holds 10, "x", TRUE, "" and 20: AVERAGE sees two numbers, AVERAGEA five
    // values.
    expect(calc('AVERAGE(D!E1:E5)')).toBe(15);
    expect(calc('AVERAGEA(D!E1:E5)')).toBe(6.2);
    expect(calc('MAXA(D!C1:C4)')).toBe(0);
    expect(calc('MINA(D!E1:E5)')).toBe(0);
    expect(calc('MAX(D!C1:C4)')).toBe(0);
  });

  it('treats numeric-looking text in a range as zero, not as its number', () => {
    // M1 holds the text "7". AVERAGE ignores it; AVERAGEA counts it as 0.
    expect(code('AVERAGE(D!M1:M2)')).toBe('#DIV/0!');
    expect(calc('AVERAGEA(D!M1:M2)')).toBe(0);
  });

  it('keeps the sample and population denominators apart', () => {
    // E reads as 10, 0, 1, 0, 20 for the A-variants: mean 6.2, sum of squared
    // deviations 308.8, so the two denominators are 4 and 5.
    expect(num('VARA(D!E1:E5)')).toBe(77.2);
    expect(num('VARPA(D!E1:E5)')).toBe(61.76);
    expect(num('STDEVA(D!E1:E5)')).toBe(8.78635305459552);
    expect(num('STDEVPA(D!E1:E5)')).toBe(7.8587530817554);
  });
});

describe('AVERAGE, MIN and MAX', () => {
  it('reports #DIV/0! for an average with nothing to average', () => {
    expect(code('AVERAGE(D!F1:F5)')).toBe('#DIV/0!');
    expect(code('AVERAGEA(D!F1:F5)')).toBe('#DIV/0!');
  });

  it('returns zero, not an error, for MIN and MAX of nothing', () => {
    expect(calc('MAX(D!F1:F5)')).toBe(0);
    expect(calc('MIN(D!F1:F5)')).toBe(0);
  });

  it('handles a single value and a repeated one', () => {
    expect(calc('MAX(D!A1:A8)')).toBe(11);
    expect(calc('MIN(D!A1:A8)')).toBe(3);
    expect(calc('AVERAGE(5)')).toBe(5);
  });

  it('averages across several arguments of mixed shape', () => {
    expect(calc('AVERAGE(D!A1:A2,10,{20,30})')).toBe(13.8);
  });
});

describe('the conditional aggregates', () => {
  it('counts and averages on one condition', () => {
    expect(calc('COUNTIF(D!C1:C4,"a")')).toBe(2);
    expect(calc('COUNTIF(D!A1:A8,">=5")')).toBe(5);
    expect(calc('AVERAGEIF(D!A1:A8,">4")')).toBe(8);
  });

  it('takes the criteria range shape when the value range is a single cell', () => {
    expect(calc('AVERAGEIF(D!C1:C4,"a",D!A1)')).toBe(6);
  });

  it('matches wildcards and ignores case', () => {
    expect(calc('COUNTIF(D!C1:C4,"a*")')).toBe(2);
    expect(calc('COUNTIF(D!C1:C4,"A")')).toBe(2);
    expect(calc('COUNTIF(D!C1:C4,"<>a")')).toBe(2);
  });

  it('applies every condition of the plural forms', () => {
    expect(calc('COUNTIFS(D!C1:C4,"a",D!A1:A4,">4")')).toBe(1);
    expect(calc('AVERAGEIFS(D!A1:A8,D!A1:A8,">4")')).toBe(8);
    expect(calc('MAXIFS(D!A1:A8,D!B1:B8,">6")')).toBe(11);
    expect(calc('MINIFS(D!A1:A8,D!B1:B8,">6")')).toBe(3);
  });

  it('reports #DIV/0! when an average has no matching cell but zero for MAXIFS', () => {
    expect(code('AVERAGEIF(D!A1:A8,">100")')).toBe('#DIV/0!');
    expect(calc('MAXIFS(D!A1:A8,D!A1:A8,">100")')).toBe(0);
    expect(calc('MINIFS(D!A1:A8,D!A1:A8,">100")')).toBe(0);
  });

  it('requires the plural forms ranges to have the aggregated range shape', () => {
    expect(code('COUNTIFS(D!A1:A8,">1",D!B1:B4,">1")')).toBe('#VALUE!');
    expect(code('AVERAGEIFS(D!A1:A8,D!B1:B4,">1")')).toBe('#VALUE!');
  });

  it('reports #VALUE! for a criterion with no range', () => {
    expect(code('COUNTIFS(D!A1:A8,">1",D!B1:B8)')).toBe('#VALUE!');
  });

  it('reads a criterion out of a cell', () => {
    expect(calc('COUNTIF(D!C1:C4,D!C1)')).toBe(2);
  });
});

describe('order statistics', () => {
  it('takes the median of an even and an odd count', () => {
    expect(calc('MEDIAN(D!A1:A8)')).toBe(6);
    expect(calc('MEDIAN(1,2,3,4)')).toBe(2.5);
    expect(calc('MEDIAN(D!A1:A7)')).toBe(5);
    expect(code('MEDIAN(D!F1:F5)')).toBe('#NUM!');
  });

  it('reports the single mode and every tied mode', () => {
    expect(calc('MODE.SNGL(D!A1:A8)')).toBe(4);
    expect(grid('MODE.MULT(1,2,2,3,3)')).toEqual({ rows: 2, cols: 1, data: [2, 3] });
    // Excel documents #N/A when nothing repeats; LibreOffice answers #VALUE!.
    expect(code('MODE.SNGL(D!D1:D8)')).toBe('#N/A');
    expect(code('MODE.MULT(1,2,3)')).toBe('#N/A');
  });

  it('picks the k-th largest and smallest', () => {
    expect(calc('LARGE(D!A1:A8,3)')).toBe(8);
    expect(calc('SMALL(D!A1:A8,3)')).toBe(4);
    expect(calc('LARGE(D!A1:A8,1)')).toBe(11);
    expect(calc('SMALL(D!A1:A8,8)')).toBe(11);
  });

  it('reports #NUM! for a rank outside the data', () => {
    expect(code('LARGE(D!A1:A8,0)')).toBe('#NUM!');
    expect(code('LARGE(D!A1:A8,9)')).toBe('#NUM!');
    expect(code('SMALL(D!F1:F5,1)')).toBe('#NUM!');
  });

  it('ranks descending by default and ascending on request', () => {
    // A holds 4,5,8,7,11,4,3,9 and 4 appears twice, so both fours take rank 6.
    expect(calc('RANK.EQ(4,D!A1:A8)')).toBe(6);
    expect(calc('RANK.EQ(4,D!A1:A8,1)')).toBe(2);
    expect(calc('RANK.AVG(4,D!A1:A8)')).toBe(6.5);
    expect(code('RANK.EQ(99,D!A1:A8)')).toBe('#N/A');
  });

  it('interpolates percentiles inclusively and exclusively', () => {
    expect(calc('PERCENTILE.INC(D!A1:A8,0.25)')).toBe(4);
    expect(calc('PERCENTILE.INC(D!A1:A8,0)')).toBe(3);
    expect(calc('PERCENTILE.INC(D!A1:A8,1)')).toBe(11);
    // The exclusive quarter point lands between the two fours, so both forms
    // report 4 here and part company at the tenth percentile below.
    expect(calc('PERCENTILE.EXC(D!A1:A8,0.25)')).toBe(4);
    expect(calc('PERCENTILE.EXC(D!A1:A8,0.2)')).toBe(3.8);
    expect(calc('QUARTILE.INC(D!A1:A8,1)')).toBe(4);
    expect(calc('QUARTILE.EXC(D!A1:A8,1)')).toBe(4);
  });

  it('refuses the percentiles the exclusive form cannot reach', () => {
    // With eight values only 1/9 to 8/9 is defined.
    expect(code('PERCENTILE.EXC(D!A1:A8,0.05)')).toBe('#NUM!');
    expect(code('PERCENTILE.INC(D!A1:A8,1.5)')).toBe('#NUM!');
    expect(code('QUARTILE.INC(D!A1:A8,5)')).toBe('#NUM!');
    expect(code('QUARTILE.EXC(D!A1:A8,0)')).toBe('#NUM!');
  });

  it('truncates a percent rank rather than rounding it', () => {
    // Five ninths is 0.5555...; Excel's own help page prints 0.555 for this
    // shape of answer, while LibreOffice rounds to 0.556.
    expect(calc('PERCENTRANK.EXC(D!A1:A8,7)')).toBe(0.555);
    expect(calc('PERCENTRANK.INC(D!A1:A8,7)')).toBe(0.571);
    expect(calc('PERCENTRANK.INC(D!A1:A8,6)')).toBe(0.5);
    expect(calc('PERCENTRANK.INC(D!A1:A8,3)')).toBe(0);
    expect(code('PERCENTRANK.INC(D!A1:A8,100)')).toBe('#N/A');
  });

  it('gives a repeated value one rank rather than several', () => {
    // 4 appears twice; both occurrences share the rank of the first.
    expect(calc('PERCENTRANK.INC({1,4,4,7},4)')).toBe(0.333);
  });
});

describe('dispersion', () => {
  it('divides by n-1 for a sample and by n for a population', () => {
    expect(num('VAR.S(D!A1:A8)')).toBe(7.98214285714286);
    expect(num('VAR.P(D!A1:A8)')).toBe(6.984375);
    expect(num('STDEV.S(D!A1:A8)')).toBe(2.82526863450944);
    expect(num('STDEV.P(D!A1:A8)')).toBe(2.64279681398325);
    expect(calc('VAR.S(1,2)')).toBe(0.5);
    expect(calc('VAR.P(1,2)')).toBe(0.25);
  });

  it('needs two values for the sample forms and one for the population forms', () => {
    expect(code('VAR.S(1)')).toBe('#DIV/0!');
    expect(code('STDEV.S(1)')).toBe('#DIV/0!');
    expect(calc('VAR.P(1)')).toBe(0);
    expect(code('VAR.P(D!F1:F5)')).toBe('#DIV/0!');
  });

  it('computes the deviation aggregates', () => {
    expect(calc('DEVSQ(D!A1:A8)')).toBe(55.875);
    expect(calc('AVEDEV(D!A1:A8)')).toBe(2.375);
    expect(calc('DEVSQ(1,2,3,4)')).toBe(5);
    expect(calc('AVEDEV(1,2,3,4)')).toBe(1);
  });

  it('reports the empty set the way each function reaches it', () => {
    expect(code('AVEDEV(D!F1:F5)')).toBe('#NUM!');
    expect(code('DEVSQ(D!F1:F5)')).toBe('#DIV/0!');
  });

  it('computes the geometric and harmonic means and rejects non-positive data', () => {
    expect(num('GEOMEAN(D!A1:A8)')).toBe(5.82779555555977);
    expect(num('HARMEAN(D!A1:A8)')).toBe(5.32194197124961);
    expect(num('GEOMEAN(1,2,3,4)')).toBe(2.21336383940064);
    expect(calc('HARMEAN(1,2,4)')).toBe(12 / 7);
    expect(code('GEOMEAN(-1,2)')).toBe('#NUM!');
    expect(code('GEOMEAN(0,2)')).toBe('#NUM!');
    expect(code('HARMEAN(0,2)')).toBe('#NUM!');
  });

  it('trims a multiple of two points from a trimmed mean', () => {
    // Ten values at 20 per cent trims two, one from each end.
    expect(calc('TRIMMEAN({1,2,3,4,5,6,7,8,9,10},0.2)')).toBe(5.5);
    expect(calc('TRIMMEAN({1,2,3,4,5,6,7,8,9,100},0.2)')).toBe(5.5);
    // Three points asked for rounds down to two.
    expect(calc('TRIMMEAN({1,2,3,4,5,6,7,8,9,10},0.3)')).toBe(5.5);
    expect(code('TRIMMEAN(D!A1:A8,1.5)')).toBe('#NUM!');
  });

  it('standardizes and rejects a non-positive spread', () => {
    expect(calc('STANDARDIZE(7,5,2)')).toBe(1);
    expect(code('STANDARDIZE(1,0,0)')).toBe('#NUM!');
    expect(code('STANDARDIZE(1,0,-1)')).toBe('#NUM!');
  });

  it('computes both skewness conventions and the excess kurtosis', () => {
    close('SKEW(D!A1:A8)', 0.453719401749366);
    close('SKEW.P(D!A1:A8)', 0.363784832373893);
    close('KURT(D!A1:A8)', -1.08413134543489);
    // A symmetric sample has no skew at all.
    expect(num('SKEW(1,2,3,4,5)')).toBe(0);
  });

  it('needs enough points, and some spread, for skew and kurtosis', () => {
    expect(code('SKEW(1,2)')).toBe('#DIV/0!');
    expect(code('SKEW.P(1,2)')).toBe('#DIV/0!');
    expect(code('SKEW(1,1,1)')).toBe('#DIV/0!');
    expect(code('KURT(1,2,3)')).toBe('#DIV/0!');
    expect(code('KURT(1,1,1,1)')).toBe('#DIV/0!');
  });
});

describe('correlation and regression', () => {
  it('computes the covariances, correlation and fit', () => {
    close('COVARIANCE.P(D!G1:G5,D!H1:H5)', 4.02);
    close('COVARIANCE.S(D!G1:G5,D!H1:H5)', 5.025);
    close('CORREL(D!G1:G5,D!H1:H5)', 0.99813645654035);
    close('PEARSON(D!G1:G5,D!H1:H5)', 0.99813645654035);
    close('RSQ(D!G1:G5,D!H1:H5)', 0.996276385874926);
    close('SLOPE(D!G1:G5,D!H1:H5)', 2.01);
    close('INTERCEPT(D!G1:G5,D!H1:H5)', 0.13);
    close('FORECAST.LINEAR(6,D!G1:G5,D!H1:H5)', 12.19);
  });

  it('is exact on a perfectly linear relationship', () => {
    expect(num('SLOPE({2,4,6},{1,2,3})')).toBe(2);
    expect(num('INTERCEPT({2,4,6},{1,2,3})')).toBe(0);
    expect(num('RSQ({2,4,6},{1,2,3})')).toBe(1);
    expect(num('CORREL({2,4,6},{1,2,3})')).toBe(1);
    expect(num('CORREL({2,4,6},{3,2,1})')).toBe(-1);
  });

  it('skips a pair unless both halves are numbers', () => {
    expect(num('CORREL({1,2,"x",3},{2,4,9,6})')).toBe(1);
  });

  it('reports #N/A for arrays of different lengths and #DIV/0! for no spread', () => {
    expect(code('CORREL({1,2},{1,2,3})')).toBe('#N/A');
    expect(code('COVARIANCE.P({1,2},{1,2,3})')).toBe('#N/A');
    expect(code('SLOPE({1,1},{2,2})')).toBe('#DIV/0!');
    expect(code('INTERCEPT({1,1},{2,2})')).toBe('#DIV/0!');
    expect(code('RSQ({1,1},{2,2})')).toBe('#DIV/0!');
    expect(code('FORECAST.LINEAR(1,{1,1},{2,2})')).toBe('#DIV/0!');
  });

  it('returns LINEST slope and intercept in Excel order', () => {
    // INDEX belongs to another category, so the array is read directly here.
    const g = grid('LINEST(D!G1:G5,D!H1:H5)');
    expect(g.rows).toBe(1);
    expect(g.cols).toBe(2);
    expect(g.data[0]).toBe(2.01);
    // The intercept carries the elimination residue of the fit, which Excel
    // shows as 0.13 and stores just as inexactly.
    expect(g.data[1] as number).toBeCloseTo(0.13, 12);
  });

  it('fills the LINEST statistics block, padding what does not apply', () => {
    const g = grid('LINEST(D!G1:G5,D!H1:H5,TRUE,TRUE)');
    expect(g.rows).toBe(5);
    expect(g.cols).toBe(2);
    // Row 1 is the coefficients, row 2 their standard errors, row 3 is
    // [r squared, standard error of y], row 4 is [F, degrees of freedom] and
    // row 5 is [regression sum of squares, residual sum of squares].
    expect(g.data[0]).toBe(2.01);
    expect(g.data[1] as number).toBeCloseTo(0.13, 12);
    expect(g.data[2]).toBe(0.0709459888459759);
    expect(g.data[4]).toBe(0.996276385874926);
    expect(g.data[5]).toBe(0.224350915606185);
    expect(g.data[6]).toBe(802.668874172185);
    expect(g.data[7]).toBe(3);
    expect(g.data[8]).toBe(40.401);
    expect(g.data[9]).toBe(0.151);
    // The two sums of squares add to the total, which is what makes r squared
    // and F consistent with the coefficients above them.
    expect(toExcelPrecision((g.data[8] as number) + (g.data[9] as number))).toBe(40.552);
  });

  it('reports #N/A for the intercept error when the intercept is forced off', () => {
    const g = grid('LINEST(D!G1:G5,D!H1:H5,FALSE,TRUE)');
    expect(g.data[1]).toBe(0);
    expect(g.data[3]).toEqual(CellError.NA);
  });

  it('fits two independent variables at once', () => {
    // Excel writes the coefficients backwards, so the first column belongs to
    // the last variable and the last column is the intercept.
    const g = grid('LINEST(D!G1:G5,D!H1:I5)');
    expect(g.rows).toBe(1);
    expect(g.cols).toBe(3);
    expect(g.data[0]).toBe(-0.105357142857144);
    expect(g.data[1]).toBe(2.05214285714286);
    expect(g.data[2]).toBe(0.298571428571421);
    // Five statistics rows and one column per coefficient, with the unused
    // cells of rows three to five padded with #N/A.
    const stats = grid('LINEST(D!G1:G5,D!H1:I5,TRUE,TRUE)');
    expect(stats.rows).toBe(5);
    expect(stats.cols).toBe(3);
    expect(stats.data[8]).toEqual(CellError.NA);
    expect(stats.data[11]).toEqual(CellError.NA);
  });

  it('predicts with TREND and GROWTH, keeping the shape of new_x', () => {
    expect(grid('TREND(D!G1:G5,D!H1:H5,{6;7})')).toEqual({
      rows: 2,
      cols: 1,
      data: [12.19, 14.2],
    });
    expect(grid('TREND(D!G1:G5,D!H1:H5,{6,7})').cols).toBe(2);
    // With known_x omitted the trend runs against 1, 2, 3 ... which is exactly
    // what column H holds, so the fitted values are the same.
    expect(grid('TREND(D!G1:G5)')).toEqual({
      rows: 5,
      cols: 1,
      data: [2.14, 4.15, 6.16, 8.17, 10.18],
    });
    expect(grid('GROWTH({1,2,4,8},{1,2,3,4},{5})')).toEqual({
      rows: 1,
      cols: 1,
      data: [16],
    });
  });

  it('reports #NUM! when GROWTH is handed a non-positive observation', () => {
    expect(code('GROWTH({1,0,4},{1,2,3},{4})')).toBe('#NUM!');
  });

  it('reports #REF! when the regression arrays cannot be lined up', () => {
    expect(code('LINEST({1,2,3},{1,2})')).toBe('#REF!');
    expect(code('TREND({1;2;3},{1,2;3,4})')).toBe('#REF!');
  });
});

describe('FREQUENCY and PROB', () => {
  it('bins values inclusively at the upper edge', () => {
    // A holds 3,4,4,5,7,8,9,11: three at or below 4, two more at or below 8.
    expect(grid('FREQUENCY(D!A1:A8,{4,8})')).toEqual({
      rows: 3,
      cols: 1,
      data: [3, 3, 2],
    });
  });

  it('returns one bucket when there are no bins', () => {
    expect(grid('FREQUENCY({1,2,3},D!F1:F3)')).toEqual({ rows: 1, cols: 1, data: [3] });
  });

  it('ignores text and blanks in the data', () => {
    expect(grid('FREQUENCY({1,"x",3},{2})')).toEqual({ rows: 2, cols: 1, data: [1, 1] });
  });

  it('sums the probabilities inside the limits', () => {
    expect(calc('PROB({1,2,3,4},{0.1,0.2,0.3,0.4},2,3)')).toBe(0.5);
    expect(calc('PROB({1,2,3,4},{0.1,0.2,0.3,0.4},2)')).toBe(0.2);
  });

  it('insists the probabilities form a distribution', () => {
    expect(code('PROB({1,2},{0.2,0.3},1,2)')).toBe('#NUM!');
    expect(code('PROB({1,2},{0.5,0.6},1,2)')).toBe('#NUM!');
    expect(code('PROB({1,2,3},{0.5,0.5},1,2)')).toBe('#N/A');
  });
});

// ---------------------------------------------------------------------------
// Layer two: distribution values from the LibreOffice recalculation pipeline.
// ---------------------------------------------------------------------------

const ORACLE_DISTRIBUTIONS: readonly (readonly [string, number])[] = [
  ['NORM.DIST(42,40,1.5,TRUE)', 0.908788780274132],
  ['NORM.DIST(42,40,1.5,FALSE)', 0.109340049783996],
  ['NORM.INV(0.908789,40,1.5)', 42.0000020095662],
  ['NORM.S.DIST(1.333333,TRUE)', 0.908788725604095],
  ['NORM.S.DIST(1.333333,FALSE)', 0.164010147569367],
  ['NORM.S.DIST(-8,TRUE)', 6.22096057427176e-16],
  ['NORM.S.INV(0.908789)', 1.33333467304411],
  ['NORM.S.INV(0.000000000001)', -7.03448382530113],
  ['LOGNORM.DIST(4,3.5,1.2,TRUE)', 0.0390835557068005],
  ['LOGNORM.DIST(4,3.5,1.2,FALSE)', 0.0176175966818192],
  ['LOGNORM.INV(0.039084,3.5,1.2)', 4.00002521868064],
  ['PHI(0.75)', 0.301137432154804],
  ['GAUSS(2)', 0.477249868051821],
  ['CONFIDENCE.NORM(0.05,2.5,50)', 0.692951912174839],
  ['CONFIDENCE.T(0.05,1,50)', 0.28419685549573],
  ['T.DIST(60,1,TRUE)', 0.994695326367377],
  ['T.DIST(-1.5,10,TRUE)', 0.0822536632227202],
  ['T.DIST(1.5,10,FALSE)', 0.127444794287092],
  ['T.DIST.2T(1.96,60)', 0.0546449297365292],
  ['T.DIST.RT(1.96,60)', 0.0273224648682646],
  ['T.INV(0.75,2)', 0.816496580927726],
  ['T.INV(0.05,10)', -1.81246112281168],
  ['T.INV.2T(0.05,10)', 2.22813885198628],
  ['F.DIST(15.2069,6,4,TRUE)', 0.990000043002763],
  ['F.DIST(15.2069,6,4,FALSE)', 0.00122379170878317],
  ['F.DIST.RT(15.2069,6,4)', 0.00999995699723731],
  ['F.INV(0.01,6,4)', 0.109309914124579],
  ['F.INV.RT(0.01,6,4)', 15.2068648611575],
  ['CHISQ.DIST(0.5,1,TRUE)', 0.520499877813047],
  ['CHISQ.DIST(0.5,1,FALSE)', 0.439391289467722],
  ['CHISQ.DIST.RT(18.307,10)', 0.0500005890913981],
  ['CHISQ.INV(0.93,1)', 3.28302028675954],
  ['CHISQ.INV.RT(0.05,10)', 18.3070380532752],
  ['BINOM.DIST(6,10,0.5,FALSE)', 0.205078125],
  ['BINOM.DIST(6,10,0.5,TRUE)', 0.828125],
  ['BINOM.DIST(60,100,0.5,TRUE)', 0.982399899891148],
  ['BINOM.INV(6,0.5,0.75)', 4],
  ['BINOM.INV(100,0.3,0.95)', 38],
  ['NEGBINOM.DIST(10,5,0.25,FALSE)', 0.0550486603751779],
  ['NEGBINOM.DIST(10,5,0.25,TRUE)', 0.313514058478177],
  ['POISSON.DIST(2,5,FALSE)', 0.0842243374885683],
  ['POISSON.DIST(2,5,TRUE)', 0.124652019483081],
  ['HYPGEOM.DIST(1,4,8,20,FALSE)', 0.363261093911249],
  ['HYPGEOM.DIST(1,4,8,20,TRUE)', 0.465428276573787],
  ['EXPON.DIST(0.2,10,TRUE)', 0.864664716763387],
  ['EXPON.DIST(0.2,10,FALSE)', 1.35335283236613],
  ['WEIBULL.DIST(105,20,100,TRUE)', 0.929581390069277],
  ['WEIBULL.DIST(105,20,100,FALSE)', 0.0355888640245043],
  ['GAMMA.DIST(10,9,2,TRUE)', 0.0680936347218484],
  ['GAMMA.DIST(10,9,2,FALSE)', 0.0326390196740794],
  ['GAMMA.INV(0.068094,9,2)', 10.0000111914372],
  ['BETA.DIST(2,8,10,TRUE,1,3)', 0.685470581054688],
  ['BETA.DIST(2,8,10,FALSE,1,3)', 1.4837646484375],
  ['BETA.INV(0.6854706,8,10,1,3)', 2.00000001276841],
  ['GAMMA(2.5)', 1.32934038817914],
  ['GAMMA(-3.75)', 0.267866128861416],
  ['GAMMALN(4.5)', 2.45373657084244],
  ['GAMMALN(0.5)', 0.5723649429247],
  ['FISHER(0.75)', 0.972955074527657],
  ['FISHERINV(0.972955)', 0.749999967394148],
  ['Z.TEST(D!A1:A8,4)', 0.00871155661664841],
  ['T.TEST(D!A1:A5,D!G1:G5,2,1)', 0.248575700108013],
  ['T.TEST(D!A1:A5,D!G1:G5,1,2)', 0.333281780323974],
  ['T.TEST(D!A1:A5,D!G1:G5,2,3)', 0.666822716987087],
  ['F.TEST(D!A1:A5,D!G1:G5)', 0.777327645175304],
  ['CHISQ.TEST({58,35;11,25},{45.35,47.65;17.65,18.35})', 0.000591553680236598],
];

describe('oracle: distributions recalculated by LibreOffice', () => {
  for (const [formula, expected] of ORACLE_DISTRIBUTIONS) {
    it(`reproduces ${formula}`, () => close(formula, expected));
  }
});

describe('the special functions the distributions rest on', () => {
  // Reference values computed to forty decimal places with an independent
  // continued-fraction implementation, then rounded to a double. These pin the
  // shared kernels directly rather than through a distribution that might
  // cancel an error away.
  // ERF and ERFC themselves are engineering functions in Excel and are asserted
  // in fn-engineering.test.ts. The erfc kernel is still shared with everything
  // below, and the normal tail exercises it here: Phi(-x) is erfc(x/sqrt 2)/2,
  // so an error in the kernel shows up in these values just as directly.

  it('computes log-gamma against its closed forms', () => {
    // gamma(1/2) is the square root of pi.
    close('GAMMALN(0.5)', Math.log(Math.sqrt(Math.PI)), 1e-15);
    close('GAMMALN(170)', 701.4372638087369, 1e-14);
    // The recurrence pins the shape of the function rather than one value of it.
    close('GAMMALN(51)-GAMMALN(50)', Math.log(50), 1e-13);
    // log(gamma(1)) is zero, which no relative tolerance can express.
    expect(calc('GAMMALN(1)')).toBeCloseTo(0, 14);
    expect(calc('GAMMALN(2)')).toBeCloseTo(0, 14);
  });

  it('reproduces the factorials exactly through GAMMA', () => {
    expect(calc('GAMMA(6)')).toBe(120);
    expect(calc('GAMMA(1)')).toBe(1);
    // 170 factorial, which Excel displays as 7.25741561530799E+306.
    expect(num('GAMMA(171)')).toBe(7.25741561530799e306);
    close('GAMMA(0.5)', Math.sqrt(Math.PI), 1e-15);
  });

  it('has no gamma at zero or the negative integers', () => {
    expect(code('GAMMA(0)')).toBe('#NUM!');
    expect(code('GAMMA(-1)')).toBe('#NUM!');
    expect(code('GAMMA(172)')).toBe('#NUM!');
    expect(code('GAMMALN(0)')).toBe('#NUM!');
    expect(code('GAMMALN(-1)')).toBe('#NUM!');
  });

  it('keeps the normal tail accurate where 1-CDF would not', () => {
    // Phi(-6.7175144212722014) to sixteen digits. Computing this as
    // 1 - Phi(6.72) leaves five correct digits; the oracle itself only has six.
    // The literal is truncated to fifteen digits on entry, which moves the
    // answer by a part in 1e14 all by itself; Z.TEST below reaches the same
    // point from exact inputs and holds to a part in 1e15.
    close('NORM.S.DIST(-6.7175144212722014,TRUE)', 9.242523860742656e-12, 1e-13);
    close('Z.TEST(D!A1:A8,4,1)', 9.242523860742656e-12, 1e-14);
  });

  it('inverts each distribution back to where it started', () => {
    close('NORM.INV(NORM.DIST(1.234,2,3,TRUE),2,3)', 1.234, 1e-12);
    close('CHISQ.INV(CHISQ.DIST(7.5,4,TRUE),4)', 7.5, 1e-12);
    close('F.INV(F.DIST(2.25,7,11,TRUE),7,11)', 2.25, 1e-12);
    close('T.INV(T.DIST(-0.75,9,TRUE),9)', -0.75, 1e-12);
    close('GAMMA.INV(GAMMA.DIST(3.5,2,1.5,TRUE),2,1.5)', 3.5, 1e-12);
    close('BETA.INV(BETA.DIST(0.4,2,3,TRUE),2,3)', 0.4, 1e-12);
    close('LOGNORM.INV(LOGNORM.DIST(2.5,0.5,0.75,TRUE),0.5,0.75)', 2.5, 1e-12);
    close('CHISQ.INV.RT(CHISQ.DIST.RT(21,10),10)', 21, 1e-12);
    close('F.INV.RT(F.DIST.RT(3,4,8),4,8)', 3, 1e-12);
    close('T.INV.2T(T.DIST.2T(1.8,12),12)', 1.8, 1e-12);
  });

  it('agrees with the identities that link the distributions', () => {
    // Chi-squared with two degrees of freedom is exponential with mean two.
    close('CHISQ.DIST(3,2,TRUE)', -Math.expm1(-1.5), 1e-14);
    // A t distribution with one degree of freedom is Cauchy.
    close('T.DIST(1,1,TRUE)', 0.75, 1e-14);
    // The square of a standard normal is chi-squared with one degree.
    close('CHISQ.DIST(4,1,TRUE)', 2 * 0.9772498680518208 - 1, 1e-13);
    // The standard normal density and PHI are the same function.
    close('PHI(1.5)', 0.12951759566589174, 1e-14);
    close('NORM.S.DIST(1.5,FALSE)', 0.12951759566589174, 1e-14);
    // GAUSS is the CDF measured from the middle.
    close('GAUSS(1.5)+0.5', 0.9331927987311419, 1e-14);
  });

  it('sums a discrete distribution to one', () => {
    close('BINOM.DIST(10,10,0.3,TRUE)', 1, 1e-15);
    close('HYPGEOM.DIST(4,4,8,20,TRUE)', 1, 1e-14);
    close(
      'BINOM.DIST(0,4,0.3,FALSE)+BINOM.DIST(1,4,0.3,FALSE)+BINOM.DIST(2,4,0.3,FALSE)+BINOM.DIST(3,4,0.3,FALSE)+BINOM.DIST(4,4,0.3,FALSE)',
      1,
      1e-15,
    );
    close('POISSON.DIST(0,2,FALSE)', Math.exp(-2), 1e-15);
  });

  it('handles the degenerate parameter values a discrete distribution allows', () => {
    expect(calc('BINOM.DIST(0,10,0,FALSE)')).toBe(1);
    expect(calc('BINOM.DIST(10,10,1,FALSE)')).toBe(1);
    expect(calc('BINOM.DIST(3,10,0,FALSE)')).toBe(0);
    expect(calc('BINOM.INV(10,0.5,0)')).toBe(0);
    expect(calc('BINOM.INV(10,0.5,1)')).toBe(10);
    expect(calc('POISSON.DIST(0,0,TRUE)')).toBe(1);
    expect(calc('EXPON.DIST(0,3,TRUE)')).toBe(0);
    expect(calc('WEIBULL.DIST(0,2,1,TRUE)')).toBe(0);
  });
});

describe('the domain errors the distributions report', () => {
  it('rejects a non-positive spread', () => {
    expect(code('NORM.DIST(1,0,0,TRUE)')).toBe('#NUM!');
    expect(code('NORM.DIST(1,0,-1,TRUE)')).toBe('#NUM!');
    expect(code('LOGNORM.DIST(1,0,0,TRUE)')).toBe('#NUM!');
    expect(code('CONFIDENCE.NORM(0.05,0,10)')).toBe('#NUM!');
  });

  it('rejects a probability outside the open unit interval', () => {
    expect(code('NORM.INV(0,0,1)')).toBe('#NUM!');
    expect(code('NORM.INV(1,0,1)')).toBe('#NUM!');
    expect(code('NORM.S.INV(1)')).toBe('#NUM!');
    expect(code('LOGNORM.INV(0,0,1)')).toBe('#NUM!');
    expect(code('T.INV(0,5)')).toBe('#NUM!');
    expect(code('T.INV.2T(0,5)')).toBe('#NUM!');
    expect(code('T.INV.2T(1.5,5)')).toBe('#NUM!');
    expect(code('F.INV(1.5,1,1)')).toBe('#NUM!');
    expect(code('CHISQ.INV(1,5)')).toBe('#NUM!');
    expect(code('BETA.INV(0,1,1)')).toBe('#NUM!');
    expect(code('BETA.INV(1.5,1,1)')).toBe('#NUM!');
    expect(code('GAMMA.INV(1.5,1,1)')).toBe('#NUM!');
  });

  it('rejects degrees of freedom below one', () => {
    expect(code('T.DIST(1,0,TRUE)')).toBe('#NUM!');
    expect(code('F.DIST(1,0,1,TRUE)')).toBe('#NUM!');
    expect(code('F.DIST(1,1,0,TRUE)')).toBe('#NUM!');
    expect(code('CHISQ.DIST(1,0,TRUE)')).toBe('#NUM!');
    // Degrees of freedom are truncated, so 1.9 is one and legal.
    expect(typeof calc('T.DIST(1,1.9,TRUE)')).toBe('number');
  });

  it('rejects an argument outside the distribution support', () => {
    expect(code('T.DIST.2T(-1,5)')).toBe('#NUM!');
    expect(code('F.DIST(-1,1,1,TRUE)')).toBe('#NUM!');
    expect(code('CHISQ.DIST(-1,1,TRUE)')).toBe('#NUM!');
    expect(code('EXPON.DIST(-1,1,TRUE)')).toBe('#NUM!');
    expect(code('WEIBULL.DIST(-1,1,1,TRUE)')).toBe('#NUM!');
    expect(code('GAMMA.DIST(-1,1,1,TRUE)')).toBe('#NUM!');
    expect(code('LOGNORM.DIST(0,0,1,TRUE)')).toBe('#NUM!');
    expect(code('BETA.DIST(5,1,1,TRUE,0,1)')).toBe('#NUM!');
    expect(code('BETA.DIST(0.5,2,3,TRUE,1,1)')).toBe('#NUM!');
  });

  it('rejects impossible counts', () => {
    expect(code('BINOM.DIST(11,10,0.5,FALSE)')).toBe('#NUM!');
    expect(code('BINOM.DIST(5,10,1.5,FALSE)')).toBe('#NUM!');
    expect(code('NEGBINOM.DIST(-1,5,0.5,FALSE)')).toBe('#NUM!');
    expect(code('NEGBINOM.DIST(5,0,0.5,FALSE)')).toBe('#NUM!');
    expect(code('POISSON.DIST(-1,5,FALSE)')).toBe('#NUM!');
    expect(code('POISSON.DIST(5,-1,FALSE)')).toBe('#NUM!');
    // More successes drawn than the sample holds, and a population smaller than
    // the sample.
    expect(code('HYPGEOM.DIST(5,4,8,20,FALSE)')).toBe('#NUM!');
    expect(code('HYPGEOM.DIST(1,4,8,3,FALSE)')).toBe('#NUM!');
  });

  it('rejects the shape and rate parameters that make no distribution', () => {
    expect(code('GAMMA.DIST(1,0,1,TRUE)')).toBe('#NUM!');
    expect(code('GAMMA.DIST(1,1,0,TRUE)')).toBe('#NUM!');
    expect(code('BETA.DIST(0.5,0,1,TRUE)')).toBe('#NUM!');
    expect(code('WEIBULL.DIST(1,0,1,TRUE)')).toBe('#NUM!');
    expect(code('EXPON.DIST(1,0,TRUE)')).toBe('#NUM!');
    expect(code('FISHER(1)')).toBe('#NUM!');
    expect(code('FISHER(-1)')).toBe('#NUM!');
    expect(code('CONFIDENCE.NORM(0,1,10)')).toBe('#NUM!');
    expect(code('CONFIDENCE.NORM(0.05,1,0)')).toBe('#NUM!');
  });

  it('has no t interval to report from a single observation', () => {
    expect(code('CONFIDENCE.T(0.05,1,1)')).toBe('#DIV/0!');
  });
});

describe('the hypothesis tests', () => {
  it('doubles the one-tailed t probability for the two-tailed form', () => {
    close(
      'T.TEST(D!A1:A5,D!G1:G5,2,2)',
      2 * (calc('T.TEST(D!A1:A5,D!G1:G5,1,2)') as number),
      1e-14,
    );
  });

  it('is symmetric in its arguments for the F test', () => {
    close(
      'F.TEST(D!A1:A5,D!G1:G5)',
      calc('F.TEST(D!G1:G5,D!A1:A5)') as number,
      1e-14,
    );
  });

  it('reports a certainty of no difference between identical samples', () => {
    expect(calc('T.TEST({1,2,3},{1,2,3},2,2)')).toBe(1);
    close('F.TEST({1,2,3},{1,2,3})', 1, 1e-14);
  });

  it('rejects the argument combinations that are not tests', () => {
    expect(code('T.TEST({1,2},{1,2,3},2,1)')).toBe('#N/A');
    expect(code('T.TEST({1,2,3},{1,2,3},3,1)')).toBe('#NUM!');
    expect(code('T.TEST({1,2,3},{1,2,3},2,4)')).toBe('#NUM!');
    expect(code('F.TEST({1},{1,2})')).toBe('#DIV/0!');
    expect(code('CHISQ.TEST({1,2},{1,2,3})')).toBe('#N/A');
    expect(code('Z.TEST(D!F1:F5,1)')).toBe('#N/A');
  });

  it('counts a goodness-of-fit table with n-1 degrees of freedom', () => {
    // A single row, so df is 2 rather than the 0 a two-way table would give.
    // Chi-squared is 25/15 + 0 + 25/25, and two degrees of freedom make the
    // right tail a plain exponential.
    close('CHISQ.TEST({10,20,30},{15,20,25})', Math.exp(-(25 / 15 + 1) / 2), 1e-13);
  });
});

// ---------------------------------------------------------------------------
// Metadata and the compatibility names
// ---------------------------------------------------------------------------

describe('the function table itself', () => {
  it('registers every function the category promises', () => {
    const expected = [
      'AVERAGE', 'AVERAGEA', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA',
      'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'MAX', 'MAXA', 'MAXIFS', 'MIN',
      'MINA', 'MINIFS', 'MEDIAN', 'MODE.SNGL', 'MODE.MULT', 'LARGE', 'SMALL',
      'RANK.EQ', 'RANK.AVG', 'PERCENTILE.INC', 'PERCENTILE.EXC',
      'QUARTILE.INC', 'QUARTILE.EXC', 'PERCENTRANK.INC', 'PERCENTRANK.EXC',
      'STDEV.S', 'STDEV.P', 'STDEVA', 'STDEVPA', 'VAR.S', 'VAR.P', 'VARA',
      'VARPA', 'COVARIANCE.P', 'COVARIANCE.S', 'CORREL', 'SLOPE', 'INTERCEPT',
      'RSQ', 'FORECAST.LINEAR', 'TREND', 'GROWTH', 'LINEST', 'DEVSQ', 'AVEDEV',
      'GEOMEAN', 'HARMEAN', 'TRIMMEAN', 'STANDARDIZE', 'SKEW', 'SKEW.P', 'KURT',
      'FREQUENCY', 'PROB', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV',
      'T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'F.DIST',
      'F.DIST.RT', 'F.INV', 'F.INV.RT', 'CHISQ.DIST', 'CHISQ.DIST.RT',
      'CHISQ.INV', 'CHISQ.INV.RT', 'BINOM.DIST', 'BINOM.INV', 'NEGBINOM.DIST',
      'POISSON.DIST', 'HYPGEOM.DIST', 'EXPON.DIST', 'WEIBULL.DIST',
      'LOGNORM.DIST', 'LOGNORM.INV', 'BETA.DIST', 'BETA.INV', 'GAMMA.DIST',
      'GAMMA.INV', 'GAMMA', 'GAMMALN', 'GAMMALN.PRECISE', 'CONFIDENCE.NORM',
      'CONFIDENCE.T', 'Z.TEST', 'T.TEST', 'F.TEST', 'CHISQ.TEST', 'FISHER',
      'FISHERINV', 'PHI', 'GAUSS', 'PEARSON',
      // The six pre-2010 names that are functions rather than aliases, because
      // their argument lists differ from their modern replacements.
      'NORMSDIST', 'LOGNORMDIST', 'TDIST', 'BETADIST', 'NEGBINOMDIST',
      'HYPGEOMDIST',
    ];
    for (const name of expected) {
      expect(registry.has(name), `${name} is not registered`).toBe(true);
    }
    expect(registry.size).toBe(expected.length);
  });

  it('declares no statistical function volatile or structural', () => {
    // None of these depend on sheet shape or on the recalculation clock, and
    // marking one volatile would drag its whole dependent closure into every
    // pass.
    for (const spec of STATISTICAL_FUNCTIONS) {
      expect(spec.volatile, `${spec.name} should not be volatile`).toBeUndefined();
      expect(spec.structural, `${spec.name} should not be structural`).toBeUndefined();
    }
  });

  it('marks the post-2007 names for the _xlfn storage prefix', () => {
    const built = STATISTICAL_FUNCTIONS.map((s) => s.name);
    for (const name of built) {
      const spec = registry.get(name)!;
      if (FUTURE_FUNCTIONS.has(name)) expect(spec.futureFunction).toBe(true);
    }
    // The names Excel 2010 and later introduced that the shared list does not
    // already carry are declared on the spec itself.
    for (const name of [
      'PERCENTRANK.INC', 'PERCENTRANK.EXC', 'T.DIST.2T', 'T.DIST.RT',
      'T.INV.2T', 'F.DIST.RT', 'F.INV.RT', 'CHISQ.DIST.RT', 'CHISQ.INV.RT',
      'GAMMALN.PRECISE', 'FORECAST.LINEAR',
    ]) {
      expect(registry.get(name)!.futureFunction, `${name} needs the prefix`).toBe(true);
    }
  });
});

describe('the pre-2010 compatibility names', () => {
  const aliased = new Map(FUNCTION_ALIASES.map(([oldName, canonical]) => [oldName, canonical]));

  const required: readonly (readonly [string, string])[] = [
    ['NORMDIST', 'NORM.DIST'], ['NORMINV', 'NORM.INV'], ['NORMSDIST', 'NORM.S.DIST'],
    ['NORMSINV', 'NORM.S.INV'], ['TDIST', 'T.DIST.2T'], ['TINV', 'T.INV.2T'],
    ['FDIST', 'F.DIST.RT'], ['FINV', 'F.INV.RT'], ['CHIDIST', 'CHISQ.DIST.RT'],
    ['CHIINV', 'CHISQ.INV.RT'], ['BINOMDIST', 'BINOM.DIST'],
    ['NEGBINOMDIST', 'NEGBINOM.DIST'], ['POISSON', 'POISSON.DIST'],
    ['HYPGEOMDIST', 'HYPGEOM.DIST'], ['EXPONDIST', 'EXPON.DIST'],
    ['WEIBULL', 'WEIBULL.DIST'], ['LOGNORMDIST', 'LOGNORM.DIST'],
    ['LOGINV', 'LOGNORM.INV'], ['BETADIST', 'BETA.DIST'], ['BETAINV', 'BETA.INV'],
    ['GAMMADIST', 'GAMMA.DIST'], ['GAMMAINV', 'GAMMA.INV'],
    ['CONFIDENCE', 'CONFIDENCE.NORM'], ['ZTEST', 'Z.TEST'], ['TTEST', 'T.TEST'],
    ['FTEST', 'F.TEST'], ['CHITEST', 'CHISQ.TEST'], ['COVAR', 'COVARIANCE.P'],
    ['STDEV', 'STDEV.S'], ['STDEVP', 'STDEV.P'], ['VAR', 'VAR.S'], ['VARP', 'VAR.P'],
    ['PERCENTILE', 'PERCENTILE.INC'], ['QUARTILE', 'QUARTILE.INC'],
    ['PERCENTRANK', 'PERCENTRANK.INC'], ['MODE', 'MODE.SNGL'], ['RANK', 'RANK.EQ'],
    ['CRITBINOM', 'BINOM.INV'],
  ];

  it('maps every old name a pre-2010 file can contain', () => {
    for (const [oldName, canonical] of required) {
      expect(aliased.get(oldName), `${oldName} has no alias`).toBe(canonical);
      expect(registry.has(canonical), `${canonical} is not registered`).toBe(true);
    }
  });

  it('evaluates through the old names once they are registered', () => {
    // The alias mechanism is the registry's, so this checks the wiring rather
    // than re-testing the functions themselves.
    const compat = new FunctionRegistry().registerAll(STATISTICAL_FUNCTIONS);
    for (const [oldName, canonical] of required) {
      // The six names with their own implementations are already present, and
      // createRegistry skips them the same way.
      if (!compat.has(oldName)) compat.alias(oldName, canonical);
    }
    const compatEval = new Evaluator(new WorkbookStore(book), compat, {});
    const run = (formula: string): Value =>
      compatEval.evaluate({
        ast: parseFormula(formula, { origin: { row: 0, col: 0 } }),
        sheet: 'Calc',
        row: 0,
        col: 0,
      });
    expect(run('STDEV(D!A1:A8)')).toEqual(calc('STDEV.S(D!A1:A8)'));
    expect(run('VARP(D!A1:A8)')).toEqual(calc('VAR.P(D!A1:A8)'));
    expect(run('RANK(4,D!A1:A8)')).toBe(6);
    expect(run('MODE(D!A1:A8)')).toBe(4);
    expect(run('NORMSDIST(1.333333)')).toEqual(calc('NORM.S.DIST(1.333333,TRUE)'));
    expect(run('CRITBINOM(6,0.5,0.75)')).toBe(4);
    // CHIDIST and TDIST are the right-tailed and two-tailed forms, not the
    // left-tailed ones the modern names default to.
    expect(run('CHIDIST(18.307,10)')).toEqual(calc('CHISQ.DIST.RT(18.307,10)'));
  });

  it('keeps the six divergent old signatures out of the alias mechanism', () => {
    // TDIST chooses its tail from a third argument, so aliasing it to
    // T.DIST.2T would silently halve or double the answer.
    expect(calc('TDIST(1.96,60,2)')).toEqual(calc('T.DIST.2T(1.96,60)'));
    expect(calc('TDIST(1.96,60,1)')).toEqual(calc('T.DIST.RT(1.96,60)'));
    expect(code('TDIST(-1,60,2)')).toBe('#NUM!');
    expect(code('TDIST(1,60,3)')).toBe('#NUM!');
    // The old distributions have no cumulative flag: one is the CDF, the other
    // two the probability mass at a point.
    expect(calc('NORMSDIST(1.333333)')).toEqual(calc('NORM.S.DIST(1.333333,TRUE)'));
    expect(calc('LOGNORMDIST(4,3.5,1.2)')).toEqual(calc('LOGNORM.DIST(4,3.5,1.2,TRUE)'));
    expect(calc('NEGBINOMDIST(10,5,0.25)')).toEqual(
      calc('NEGBINOM.DIST(10,5,0.25,FALSE)'),
    );
    expect(calc('HYPGEOMDIST(1,4,8,20)')).toEqual(calc('HYPGEOM.DIST(1,4,8,20,FALSE)'));
    // BETADIST's bounds sit one position to the left of BETA.DIST's.
    expect(calc('BETADIST(2,8,10,1,3)')).toEqual(calc('BETA.DIST(2,8,10,TRUE,1,3)'));
    for (const name of [
      'NORMSDIST', 'LOGNORMDIST', 'TDIST', 'BETADIST', 'NEGBINOMDIST', 'HYPGEOMDIST',
    ]) {
      expect(registry.get(name)!.deprecatedAliasOf, `${name} should name its successor`)
        .toBeDefined();
    }
  });

  it('keeps the old names out of the _xlfn prefix set', () => {
    const compat = new FunctionRegistry().registerAll(STATISTICAL_FUNCTIONS);
    compat.alias('STDEV', 'STDEV.S');
    expect(compat.get('STDEV')!.futureFunction).toBe(false);
    expect(compat.get('STDEV')!.deprecatedAliasOf).toBe('STDEV.S');
  });
});
