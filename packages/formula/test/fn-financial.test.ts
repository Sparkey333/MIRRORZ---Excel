/**
 * Financial functions.
 *
 * Three layers of evidence, in descending order of authority.
 *
 * The first is the LibreOffice-recalculated fixture: eight financial cases in
 * fixtures/generated/formulas.calc.xlsx carry the values the oracle actually
 * computed, and every one of them is replayed here. Six match Excel's fifteen
 * significant digits exactly. NPER and RATE do not, and are compared at a
 * relative tolerance instead, for reasons worth stating rather than papering
 * over: NPER ends in a ratio of two logarithms, where LibreOffice's libm and
 * ours may differ in the last bit, and RATE is iterative, where the oracle's
 * own iteration stopped about eighty units in the last place short of the true
 * root. Ours is the closer of the two - the RATE test below proves it by
 * feeding the answer back through PMT - so demanding bit equality with the
 * oracle would be demanding that we reproduce its early exit.
 *
 * The second is Microsoft's own published examples. Excel's documentation
 * carries a worked example with a printed result for nearly every function in
 * this category, and those are Excel's answers rather than an oracle's, so they
 * are the closest thing to ground truth available offline. They are asserted at
 * the precision Microsoft printed them.
 *
 * The third is internal consistency: identities that must hold whatever the
 * numbers are. IPMT plus PPMT is PMT for every period of every schedule; the
 * principal repaid over a whole loan is the loan; PRICE and YIELD invert each
 * other; the two halves of a coupon period add up to it on the bases that count
 * them the same way. These catch the class
 * of error a table of examples cannot - a formula that is right at the
 * documented point and wrong two periods later.
 *
 * The registry under test holds FINANCIAL_FUNCTIONS alone, so a failure here is
 * never another category's fault. Date arguments are therefore written as
 * serial numbers built with the core calendar rather than through DATE().
 */

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, type Scalar, Workbook, isError, partsToSerial } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { Evaluator } from '../src/evaluator.js';
import { FINANCIAL_FUNCTIONS } from '../src/functions/financial.js';
import { parseFormula } from '../src/parser.js';
import { FunctionRegistry } from '../src/registry.js';
import { WorkbookStore } from '../src/store.js';
import { type Value, toExcelPrecision } from '../src/value.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

const registry = new FunctionRegistry().registerAll(FINANCIAL_FUNCTIONS);

// ---------------------------------------------------------------------------
// The oracle fixture
// ---------------------------------------------------------------------------

const { workbook: oracleBook } = readXlsx(
  new Uint8Array(readFileSync(new URL('formulas.calc.xlsx', FIXTURES))),
);
const oracleSheet = oracleBook.getSheet('Formulas')!;
const oracleEval = new Evaluator(new WorkbookStore(oracleBook), registry, {
  dateSystem: oracleBook.dateSystem,
});

interface OracleCase {
  formula: string;
  value: Scalar;
  row: number;
}

const oracleCases = new Map<string, OracleCase>();
for (const { row, col, cell } of oracleSheet.entries()) {
  if (col !== 2 || !cell.formula) continue;
  const name = oracleSheet.getValue(row, 0);
  if (typeof name === 'string') oracleCases.set(name, { formula: cell.formula, value: cell.value, row });
}

function replayOracle(name: string): { actual: number; expected: number } {
  const c = oracleCases.get(name);
  expect(c, `case ${name} missing from the fixture`).toBeDefined();
  const actual = oracleEval.evaluate({
    ast: parseFormula(c!.formula, { origin: { row: c!.row, col: 2 } }),
    sheet: 'Formulas',
    row: c!.row,
    col: 2,
  });
  expect(typeof actual, `${name} did not produce a number: ${String(actual)}`).toBe('number');
  expect(typeof c!.value).toBe('number');
  return { actual: actual as number, expected: c!.value as number };
}

describe('oracle: formulas.calc.xlsx', () => {
  for (const name of ['PMT', 'FV', 'PV', 'NPV', 'IRR', 'SLN']) {
    it(`reproduces ${name} to fifteen significant digits`, () => {
      const { actual, expected } = replayOracle(name);
      expect(toExcelPrecision(actual)).toBe(toExcelPrecision(expected));
    });
  }

  it('reproduces NPER to within a last-bit difference in log()', () => {
    const { actual, expected } = replayOracle('NPER');
    expect(Math.abs(actual - expected) / Math.abs(expected)).toBeLessThan(1e-14);
  });

  it('reproduces RATE, and lands closer to the true root than the oracle did', () => {
    const { actual, expected } = replayOracle('RATE');
    expect(Math.abs(actual - expected) / Math.abs(expected)).toBeLessThan(1e-13);
    // The oracle case is RATE(360, -1500, 250000). Feeding the rate back
    // through PMT is the residual test the solver itself minimised, so the
    // implied payment must come back as -1500 to within rounding.
    const impliedOurs = calc(`PMT(${actual},360,250000)`) as number;
    const impliedOracle = calc(`PMT(${expected},360,250000)`) as number;
    expect(Math.abs(impliedOurs + 1500)).toBeLessThan(Math.abs(impliedOracle + 1500));
    expect(Math.abs(impliedOurs + 1500)).toBeLessThan(1e-9);
  });

  it('confirms the sign convention the fixture encodes', () => {
    // PMT(0.05/12, 360, -300000): the loan is money received, so it is
    // negative, and the payment that services it comes back positive.
    const { actual } = replayOracle('PMT');
    expect(actual).toBeGreaterThan(0);
    expect(calc('PMT(0.05/12,360,300000)')).toBeLessThan(0);
  });
});

// ---------------------------------------------------------------------------
// A hand-built workbook for everything the fixture does not reach
// ---------------------------------------------------------------------------

const book = new Workbook();
const data = book.addSheet('S');
book.addSheet('Calc');

// A1:A5 - a cash-flow column whose middle cell is blank, so NPV's period
// numbering can be checked against Excel's skip-the-blank rule.
data.setValue(0, 0, -100);
data.setValue(1, 0, 50);
data.setValue(3, 0, 60);
data.setValue(4, 0, 70);
// B1:B4 - the same flows with no gap.
data.setValue(0, 1, -100);
data.setValue(1, 1, 50);
data.setValue(2, 1, 60);
data.setValue(3, 1, 70);
// C1:C6 - the same four flows contaminated with the value kinds Excel ignores
// inside a range, so dropping them has to leave B's series exactly.
data.setValue(0, 2, -100);
data.setValue(1, 2, 'fifty');
data.setValue(2, 2, true);
data.setValue(3, 2, 50);
data.setValue(4, 2, 60);
data.setValue(5, 2, 70);
// D1:D3 - a column carrying an error, which must propagate.
data.setValue(0, 3, -100);
data.setValue(1, 3, CellError.NA);
data.setValue(2, 3, 70);
// E1:E3 - a compounding schedule with a blank, which is a zero rate.
data.setValue(0, 4, 0.09);
data.setValue(2, 4, 0.1);

const evaluator = new Evaluator(new WorkbookStore(book), registry, {});

function calc(formula: string): Value {
  return evaluator.evaluate({
    ast: parseFormula(formula, { origin: { row: 0, col: 0 } }),
    sheet: 'Calc',
    row: 0,
    col: 0,
  });
}

function num(formula: string): number {
  const v = calc(formula);
  if (typeof v !== 'number') throw new Error(`${formula} produced ${String(v)}`);
  return v;
}

function code(formula: string): string {
  const v = calc(formula);
  return isError(v) ? v.code : `not an error: ${String(v)}`;
}

/** A 1900-system date serial, so the tests read as dates without needing DATE(). */
function d(year: number, month: number, day: number): number {
  return partsToSerial(year, month, day);
}

/** Assert to the number of decimals Microsoft printed the example at. */
function near(formula: string, expected: number, decimals: number): void {
  expect(num(formula), formula).toBeCloseTo(expected, decimals);
}

// ---------------------------------------------------------------------------
// Microsoft's published examples
// ---------------------------------------------------------------------------

describe("Microsoft's documented examples: the annuity family", () => {
  it('PMT', () => near('PMT(0.08/12,10,10000)', -1037.03, 2));
  it('PV', () => near('PV(0.08/12,12*20,500,0,0)', -59777.15, 2));
  it('FV', () => near('FV(0.06/12,10,-200,-500,1)', 2581.4, 2));
  it('NPER at the end of the period', () =>
    near('NPER(0.12/12,-100,-1000,10000)', 60.0821229, 6));
  it('NPER at the beginning of the period', () => {
    // Microsoft prints this one to two decimals only, so the round trip
    // through FV is what pins down the digits beyond them.
    const nper = num('NPER(0.12/12,-100,-1000,10000,1)');
    expect(nper).toBeCloseTo(59.67, 2);
    near(`FV(0.12/12,${nper},-100,-1000,1)`, 10000, 6);
  });
  it('NPER with no future value', () => near('NPER(0.12/12,-100,-1000)', -9.57859404, 6));
  it('RATE', () => near('RATE(4*12,-200,8000)', 0.00770147, 8));
  it('RATE annualised', () => near('RATE(4*12,-200,8000)*12', 0.0924177, 6));
  it('IPMT for the first month', () => near('IPMT(0.1/12,1,3*12,8000)', -66.6666667, 5));
  it('IPMT for the last year', () => near('IPMT(0.1,3,3,8000)', -292.45, 2));
  it('PPMT for the first month', () => near('PPMT(0.1/12,1,2*12,2000)', -75.6231855, 5));
  it('PPMT in the tenth year', () => near('PPMT(0.08,10,10,200000)', -27598.0534, 3));
  it('CUMIPMT over the second year', () =>
    near('CUMIPMT(0.09/12,30*12,125000,13,24,0)', -11135.23, 2));
  it('CUMIPMT over the first month', () =>
    near('CUMIPMT(0.09/12,30*12,125000,1,1,0)', -937.5, 4));
  it('CUMPRINC over the second year', () =>
    near('CUMPRINC(0.09/12,30*12,125000,13,24,0)', -934.1071, 4));
  it('CUMPRINC over the first month', () =>
    near('CUMPRINC(0.09/12,30*12,125000,1,1,0)', -68.27827, 5));
  it('ISPMT', () => near('ISPMT(0.1/12,1,3*12,8000000)', -64814.8148, 4));
});

describe("Microsoft's documented examples: cash-flow series", () => {
  it('FVSCHEDULE', () => near('FVSCHEDULE(1,{0.09,0.11,0.1})', 1.33089, 5));
  it('MIRR over five years', () =>
    near('MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.12)', 0.126094, 6));
  it('MIRR over three years', () =>
    near('MIRR({-120000,39000,30000,21000},0.1,0.12)', -0.0480446, 6));
  it('MIRR at a different reinvestment rate', () =>
    near('MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.14)', 0.134759, 6));

  const flows = '{-10000,2750,4250,3250,2750}';
  const dates = `{${d(2008, 1, 1)},${d(2008, 3, 1)},${d(2008, 10, 30)},${d(2009, 2, 15)},${d(2009, 4, 1)}}`;
  it('XNPV', () => near(`XNPV(0.09,${flows},${dates})`, 2086.6476, 4));
  it('XIRR', () => near(`XIRR(${flows},${dates})`, 0.373362535, 8));
});

describe("Microsoft's documented examples: depreciation", () => {
  it('SLN', () => near('SLN(30000,7500,10)', 2250, 6));
  it('SYD in the first year', () => near('SYD(30000,7500,10,1)', 4090.91, 2));
  it('SYD in the last year', () => near('SYD(30000,7500,10,10)', 409.09, 2));
  it('DB in the first year', () => near('DB(1000000,100000,6,1,7)', 186083.33, 2));
  it('DB in the second year', () => near('DB(1000000,100000,6,2,7)', 259639.42, 2));
  it('DB in the stub year', () => near('DB(1000000,100000,6,7,7)', 15845.1, 2));
  it('DDB by month', () => near('DDB(2400,300,10*12,1,2)', 40, 6));
  it('DDB by year', () => near('DDB(2400,300,10,1,2)', 480, 6));
  it('DDB at factor 1.5', () => near('DDB(2400,300,10,2,1.5)', 306, 6));
  it('DDB in the final year, held at salvage', () =>
    near('DDB(2400,300,10,10)', 22.1225472, 6));
  it('VDB for one day', () => near('VDB(2400,300,10*365,0,1)', 1.315068493, 6));
  it('VDB for one month', () => near('VDB(2400,300,10*12,0,1)', 40, 6));
  it('VDB for one year', () => near('VDB(2400,300,10,0,1)', 480, 6));
  it('VDB across months six to eighteen', () =>
    near('VDB(2400,300,10*12,6,18)', 396.31, 2));
  it('VDB across months six to eighteen at factor 1.5', () =>
    near('VDB(2400,300,10*12,6,18,1.5)', 311.81, 2));
  it('VDB over a fractional first period', () =>
    near('VDB(2400,300,10,0,0.875,1.5)', 315, 6));
  it('AMORDEGRC', () =>
    near(`AMORDEGRC(2400,${d(2008, 8, 19)},${d(2008, 12, 31)},300,1,0.15,1)`, 776, 6));
  it('AMORLINC', () =>
    near(`AMORLINC(2400,${d(2008, 8, 19)},${d(2008, 12, 31)},300,1,0.15,1)`, 360, 6));
});

describe("Microsoft's documented examples: rate conversions", () => {
  it('EFFECT', () => near('EFFECT(0.0525,4)', 0.053543, 6));
  it('NOMINAL', () => near('NOMINAL(0.053543,4)', 0.05250032, 8));
  it('PDURATION', () => near('PDURATION(0.025,2000,2200)', 3.85986616, 8));
  it('RRI', () => near('RRI(96,10000,11000)', 0.0009933, 7));
  it('DOLLARDE in sixteenths', () => near('DOLLARDE(1.02,16)', 1.125, 6));
  it('DOLLARDE in thirty-seconds', () => near('DOLLARDE(1.1,32)', 1.3125, 6));
  it('DOLLARFR in sixteenths', () => near('DOLLARFR(1.125,16)', 1.02, 6));
  it('DOLLARFR in thirty-seconds', () => near('DOLLARFR(1.3125,32)', 1.1, 6));
});

describe("Microsoft's documented examples: securities", () => {
  it('PRICE', () =>
    near(`PRICE(${d(2008, 2, 15)},${d(2017, 11, 15)},0.0575,0.065,100,2,0)`, 94.63436162, 7));
  it('YIELD', () =>
    near(`YIELD(${d(2008, 2, 15)},${d(2016, 11, 15)},0.0575,95.04287,100,2,0)`, 0.065, 6));
  it('DURATION', () =>
    near(`DURATION(${d(2008, 1, 1)},${d(2016, 1, 1)},0.08,0.09,2,1)`, 5.993775, 6));
  it('MDURATION', () =>
    near(`MDURATION(${d(2008, 1, 1)},${d(2016, 1, 1)},0.08,0.09,2,1)`, 5.73567, 5));
  it('ACCRINT', () =>
    near(
      `ACCRINT(${d(2008, 3, 1)},${d(2008, 8, 31)},${d(2008, 5, 1)},0.1,1000,2,0)`,
      16.66666667,
      6,
    ));
  it('ACCRINT with settlement before the first coupon', () =>
    near(
      `ACCRINT(${d(2008, 3, 5)},${d(2008, 9, 1)},${d(2008, 5, 1)},0.1,1000,2,0,FALSE)`,
      15.55555556,
      6,
    ));
  it('ACCRINTM', () =>
    near(`ACCRINTM(${d(2008, 4, 1)},${d(2008, 6, 15)},0.1,1000,3)`, 20.54794521, 7));
  it('DISC', () =>
    near(`DISC(${d(2007, 1, 25)},${d(2007, 6, 15)},97.975,100,1)`, 0.052420213, 8));
  it('INTRATE', () =>
    near(`INTRATE(${d(2008, 2, 15)},${d(2008, 5, 15)},1000000,1014420,2)`, 0.05768, 6));
  it('RECEIVED', () =>
    near(`RECEIVED(${d(2008, 2, 15)},${d(2008, 5, 15)},1000000,0.0575,2)`, 1014584.654, 3));
  it('PRICEDISC', () =>
    near(`PRICEDISC(${d(2008, 2, 16)},${d(2008, 3, 1)},0.0525,100,2)`, 99.79583333, 6));
  it('PRICEMAT', () =>
    near(
      `PRICEMAT(${d(2008, 2, 15)},${d(2008, 4, 13)},${d(2007, 11, 11)},0.061,0.061,0)`,
      99.98449888,
      7,
    ));
  it('YIELDDISC', () =>
    near(`YIELDDISC(${d(2008, 2, 16)},${d(2008, 3, 1)},99.795,100,2)`, 0.05282257, 7));
  it('YIELDMAT', () =>
    near(
      `YIELDMAT(${d(2008, 3, 15)},${d(2008, 11, 3)},${d(2007, 11, 8)},0.0625,100.0123,0)`,
      0.060954, 6,
    ));
  it('TBILLEQ', () =>
    near(`TBILLEQ(${d(2008, 3, 31)},${d(2008, 6, 1)},0.0914)`, 0.09415149, 8));
  it('TBILLPRICE', () =>
    near(`TBILLPRICE(${d(2008, 3, 31)},${d(2008, 6, 1)},0.09)`, 98.45, 6));
  it('TBILLYIELD', () =>
    near(`TBILLYIELD(${d(2008, 3, 31)},${d(2008, 6, 1)},98.45)`, 0.09141696, 8));
});

describe("Microsoft's documented examples: coupon dates", () => {
  const bond = `${d(2011, 1, 25)},${d(2011, 11, 15)},2,1`;
  it('COUPDAYBS', () => expect(num(`COUPDAYBS(${bond})`)).toBe(71));
  it('COUPDAYS', () => expect(num(`COUPDAYS(${bond})`)).toBe(181));
  it('COUPDAYSNC', () => expect(num(`COUPDAYSNC(${bond})`)).toBe(110));
  it('COUPNCD', () => expect(num(`COUPNCD(${bond})`)).toBe(d(2011, 5, 15)));
  it('COUPPCD', () => expect(num(`COUPPCD(${bond})`)).toBe(d(2010, 11, 15)));
  it('COUPNUM', () =>
    expect(num(`COUPNUM(${d(2007, 1, 25)},${d(2008, 11, 15)},2,1)`)).toBe(4));
});

// ---------------------------------------------------------------------------
// Sign convention and the annuity-due branch
// ---------------------------------------------------------------------------

describe('sign convention', () => {
  it('treats money received as positive and money paid out as negative', () => {
    // Borrowing 300,000 and repaying it: the loan is an inflow, the payments
    // outflows, so exactly one of the two is negative whichever way round the
    // arguments are written.
    expect(num('PMT(0.05/12,360,-300000)')).toBeGreaterThan(0);
    expect(num('PMT(0.05/12,360,300000)')).toBeLessThan(0);
    expect(num('FV(0.05/12,360,-1000)')).toBeGreaterThan(0);
    expect(num('PV(0.05/12,360,-1000)')).toBeGreaterThan(0);
  });

  it('keeps PV, FV and PMT mutually consistent', () => {
    const pmt = num('PMT(0.07/12,240,-250000)');
    near(`PV(0.07/12,240,${-pmt})`, 250000, 6);
    near(`FV(0.07/12,240,${-pmt},250000)`, 0, 6);
    near(`NPER(0.07/12,${-pmt},250000)`, 240, 8);
  });
});

describe('the end-of-period versus beginning-of-period branch', () => {
  it('discounts an annuity due by one extra period in PMT', () => {
    near('PMT(0.1,10,1000,0,1)', num('PMT(0.1,10,1000)') / 1.1, 10);
  });

  it('grows an annuity due by one extra period in FV and PV', () => {
    near('FV(0.1,10,-100,0,1)', num('FV(0.1,10,-100)') * 1.1, 8);
    near('PV(0.1,10,-100,0,1)', num('PV(0.1,10,-100)') * 1.1, 8);
  });

  it('charges no interest in the first period of an annuity due', () => {
    expect(num('IPMT(0.1,1,10,1000,0,1)')).toBe(0);
    // With payment in arrears the first period does carry interest.
    near('IPMT(0.1,1,10,1000)', -100, 10);
  });

  it('shortens the term of an annuity due', () => {
    expect(num('NPER(0.12/12,-100,-1000,10000,1)')).toBeLessThan(
      num('NPER(0.12/12,-100,-1000,10000)'),
    );
  });

  it('is not a single scale factor: type 1 changes RATE too', () => {
    const arrears = num('RATE(48,-200,8000)');
    const due = num('RATE(48,-200,8000,0,1)');
    expect(due).toBeGreaterThan(arrears);
    near(`PMT(${due},48,8000,0,1)`, -200, 8);
  });
});

// ---------------------------------------------------------------------------
// Identities that must hold for every period, not just the documented one
// ---------------------------------------------------------------------------

describe('amortisation identities', () => {
  const schedules: Array<[number, number, number, 0 | 1]> = [
    [0.09 / 12, 360, 125000, 0],
    [0.09 / 12, 360, 125000, 1],
    [0.05, 10, -5000, 0],
    [0.05, 10, -5000, 1],
  ];

  it('splits every payment into interest plus principal', () => {
    for (const [rate, nper, pv, type] of schedules) {
      const pmt = num(`PMT(${rate},${nper},${pv},0,${type})`);
      for (const per of [1, 2, 7, nper - 1, nper]) {
        const i = num(`IPMT(${rate},${per},${nper},${pv},0,${type})`);
        const p = num(`PPMT(${rate},${per},${nper},${pv},0,${type})`);
        expect(i + p, `period ${per} of ${rate}/${nper}/${pv}/${type}`).toBeCloseTo(pmt, 8);
      }
    }
  });

  it('repays exactly the principal over the whole schedule', () => {
    for (const [rate, nper, pv, type] of schedules) {
      let total = 0;
      for (let per = 1; per <= nper; per++) {
        total += num(`PPMT(${rate},${per},${nper},${pv},0,${type})`);
      }
      expect(total).toBeCloseTo(-pv, 6);
    }
  });

  it('agrees with the closed-form cumulative functions', () => {
    // CUMPRINC and CUMIPMT are computed from the balance equation rather than
    // by summing; the sums are what that shortcut has to reproduce.
    for (const [rate, nper, pv, type] of schedules) {
      if (pv < 0) continue; // CUMIPMT and CUMPRINC reject a non-positive pv.
      for (const [from, to] of [
        [1, 1],
        [13, 24],
        [1, nper],
      ]) {
        let interest = 0;
        let principal = 0;
        for (let per = from; per <= to; per++) {
          interest += num(`IPMT(${rate},${per},${nper},${pv},0,${type})`);
          principal += num(`PPMT(${rate},${per},${nper},${pv},0,${type})`);
        }
        expect(num(`CUMIPMT(${rate},${nper},${pv},${from},${to},${type})`)).toBeCloseTo(
          interest,
          6,
        );
        expect(num(`CUMPRINC(${rate},${nper},${pv},${from},${to},${type})`)).toBeCloseTo(
          principal,
          6,
        );
      }
    }
  });

  it('writes the loan off completely across every period', () => {
    near('CUMPRINC(0.09/12,360,125000,1,360,0)', -125000, 4);
  });
});

describe('round-trip identities', () => {
  it('YIELD inverts PRICE', () => {
    for (const [frequency, basis] of [
      [1, 0],
      [2, 0],
      [2, 1],
      [4, 2],
      [2, 3],
      [2, 4],
    ]) {
      const bond = `${d(2020, 6, 10)},${d(2029, 3, 31)},0.045,%,100,${frequency},${basis}`;
      const price = num(`PRICE(${bond.replace('%', '0.052')})`);
      const yld = num(`YIELD(${bond.replace('%', String(price))})`);
      expect(yld, `frequency ${frequency} basis ${basis}`).toBeCloseTo(0.052, 8);
    }
  });

  it('YIELD inverts PRICE inside the final coupon period as well', () => {
    // The last period uses simple discounting and a closed-form inversion, so
    // it is a genuinely different code path from the iterative branch.
    const bond = `${d(2024, 1, 15)},${d(2024, 5, 1)},0.06,%,100,2,0`;
    const price = num(`PRICE(${bond.replace('%', '0.048')})`);
    expect(num(`COUPNUM(${d(2024, 1, 15)},${d(2024, 5, 1)},2,0)`)).toBe(1);
    expect(num(`YIELD(${bond.replace('%', String(price))})`)).toBeCloseTo(0.048, 9);
  });

  it('NOMINAL inverts EFFECT', () => {
    for (const periods of [1, 2, 4, 12, 365]) {
      near(`NOMINAL(EFFECT(0.0725,${periods}),${periods})`, 0.0725, 12);
    }
  });

  it('DOLLARFR inverts DOLLARDE, including for negative quotations', () => {
    for (const fraction of [2, 4, 8, 16, 32]) {
      near(`DOLLARFR(DOLLARDE(1.05,${fraction}),${fraction})`, 1.05, 10);
      near(`DOLLARDE(-1.02,16)`, -1.125, 10);
      near(`DOLLARFR(-1.125,16)`, -1.02, 10);
    }
  });

  it('RRI inverts compound growth, and PDURATION finds the term', () => {
    near('RRI(10,1000,FV(0.06,10,0,-1000))', 0.06, 10);
    near('PDURATION(0.06,1000,FV(0.06,10,0,-1000))', 10, 8);
  });

  it('XIRR is the rate that makes XNPV vanish', () => {
    const flows = '{-10000,2750,4250,3250,2750}';
    const dates = `{${d(2008, 1, 1)},${d(2008, 3, 1)},${d(2008, 10, 30)},${d(2009, 2, 15)},${d(2009, 4, 1)}}`;
    const rate = num(`XIRR(${flows},${dates})`);
    expect(Math.abs(num(`XNPV(${rate},${flows},${dates})`))).toBeLessThan(1e-6);
  });

  it('IRR is the rate that makes the flows discount to zero', () => {
    const rate = num('IRR({-100,50,60,70})');
    // NPV over the same flows discounts from period 1, so the series has to be
    // written one period later to reach the same equation IRR solved.
    expect(Math.abs(num(`NPV(${rate},-100,50,60,70)`) * (1 + rate))).toBeLessThan(1e-9);
  });
});

// ---------------------------------------------------------------------------
// NPV against XNPV: the time-zero flow
// ---------------------------------------------------------------------------

describe('NPV and XNPV disagree about time zero, deliberately', () => {
  it('discounts NPV from period one, initial outlay included', () => {
    // The classic user error is passing the outlay to NPV, which discounts it.
    // Matching Excel means reproducing that, not correcting it.
    near('NPV(0.1,-100,50,60,70)', 43.3030530701455, 10);
    // The correct-by-finance-textbook form places the outlay outside.
    near('NPV(0.1,50,60,70)-100', 47.6333583771601, 8);
  });

  it('leaves the first XNPV flow undiscounted', () => {
    // 2021 is not a leap year, so these two dates are 365 days apart: the
    // second flow discounts exactly once and the first not at all.
    const dates = `{${d(2021, 1, 1)},${d(2022, 1, 1)}}`;
    expect(d(2022, 1, 1) - d(2021, 1, 1)).toBe(365);
    near(`XNPV(0.1,{-100,110},${dates})`, 0, 8);
  });

  it('uses a 365-day year regardless of leap years', () => {
    // 2020 is a leap year: 366 actual days, so a full calendar year discounts
    // by slightly more than one period.
    const dates = `{${d(2020, 1, 1)},${d(2021, 1, 1)}}`;
    near(`XNPV(0.1,{0,100},${dates})`, 100 / Math.pow(1.1, 366 / 365), 8);
  });
});

// ---------------------------------------------------------------------------
// Blanks, text, logicals and errors inside ranges
// ---------------------------------------------------------------------------

describe('what a range contributes', () => {
  it('skips blank cells in NPV rather than counting them as a period', () => {
    // S!A1:A5 has a gap at A3; S!B1:B4 is the same four numbers with no gap.
    expect(num('NPV(0.1,S!A1:A5)')).toBeCloseTo(num('NPV(0.1,S!B1:B4)'), 12);
  });

  it('ignores text and logicals inside a range', () => {
    expect(num('NPV(0.1,S!C1:C6)')).toBeCloseTo(num('NPV(0.1,S!B1:B4)'), 12);
    expect(num('IRR(S!C1:C6)')).toBeCloseTo(num('IRR(S!B1:B4)'), 12);
  });

  it('coerces the same values when they are typed into the formula', () => {
    // Excel decides by provenance, not by type: a bare TRUE is a 1.
    near('NPV(0.1,TRUE)', 1 / 1.1, 12);
    near('NPV(0.1,"50")', 50 / 1.1, 12);
  });

  it('propagates an error out of a range', () => {
    expect(code('NPV(0.1,S!D1:D3)')).toBe('#N/A');
    expect(code('IRR(S!D1:D3)')).toBe('#N/A');
    expect(code('XNPV(0.1,S!D1:D3,S!B1:B3)')).toBe('#N/A');
  });

  it('treats a blank rate in an FVSCHEDULE as zero, not as a skipped period', () => {
    // S!E1:E3 is {0.09, blank, 0.1}: three periods, the middle one flat.
    near('FVSCHEDULE(1,S!E1:E3)', 1.09 * 1.1, 12);
  });

  it('propagates an error from a scalar argument before doing any work', () => {
    expect(code('PMT(1/0,10,100)')).toBe('#DIV/0!');
    // S!D2 holds an #N/A, which reaches FV as a scalar argument.
    expect(code('FV(0.1,S!D2,100)')).toBe('#N/A');
  });

  it('coerces numeric text in scalar arguments', () => {
    expect(num('SLN("10000","1000","5")')).toBe(1800);
    expect(code('SLN("ten thousand",1000,5)')).toBe('#VALUE!');
  });

  it('treats an omitted optional argument as blank, which is zero', () => {
    expect(num('PMT(0.1,10,1000,,1)')).toBe(num('PMT(0.1,10,1000,0,1)'));
    expect(num('FV(0.1,10,-100)')).toBe(num('FV(0.1,10,-100,0,0)'));
  });
});

// ---------------------------------------------------------------------------
// Errors: the exact code, not merely an error
// ---------------------------------------------------------------------------

describe('degenerate closed forms report the arithmetic that failed', () => {
  it('divides by zero when the term is zero', () => {
    expect(code('PMT(0.1,0,1000)')).toBe('#DIV/0!');
    expect(code('PMT(0,0,1000)')).toBe('#DIV/0!');
    expect(code('ISPMT(0.1,1,0,1000)')).toBe('#DIV/0!');
    expect(code('SLN(10000,1000,0)')).toBe('#DIV/0!');
  });

  it('divides by zero when a zero rate leaves no payment to divide by', () => {
    expect(code('NPER(0,0,1000)')).toBe('#DIV/0!');
  });

  it('reports #NUM! when no term can reconcile the arguments', () => {
    // Repaying 1,000 at 10% with a payment of 1 never terminates.
    expect(code('NPER(0.1,-1,1000)')).toBe('#NUM!');
    // Two inflows do have a solution - a negative one - and Excel returns it
    // rather than refusing, which its own NPER(0.12/12,-100,-1000) example
    // shows.
    expect(num('NPER(0.1,100,1000)')).toBeCloseTo(-7.2725409, 6);
  });

  it('handles a zero rate exactly rather than by a limit', () => {
    expect(num('PMT(0,10,-1000)')).toBe(100);
    expect(num('FV(0,10,-100)')).toBe(1000);
    expect(num('PV(0,10,-100)')).toBe(1000);
    expect(num('NPER(0,-100,1000)')).toBe(10);
  });

  it('reports #NUM! for a rate at or below -100 per cent', () => {
    expect(code('PV(-1,10,-100)')).toBe('#DIV/0!');
    expect(code('NPER(-1,-100,1000)')).toBe('#NUM!');
    expect(code('RATE(10,-100,1000,0,0,-1)')).toBe('#NUM!');
  });
});

describe('solvers fail the way Excel fails', () => {
  it('reports #NUM! when RATE cannot converge inside its budget', () => {
    // Same sign on both sides: nothing repays anything, so there is no rate.
    expect(code('RATE(10,100,1000)')).toBe('#NUM!');
    expect(code('RATE(10,-100,-1000)')).toBe('#NUM!');
    // Ten payments of 100 cannot both clear 1,000 and leave 5,000 behind.
    expect(code('RATE(10,-100,1000,5000)')).toBe('#NUM!');
    // A guess this far from any root exhausts the budget before reaching one.
    expect(code('RATE(360,-1500,250000,0,0,1000000)')).toBe('#NUM!');
    expect(code('RATE(0,-100,1000)')).toBe('#NUM!');
  });

  it('solves a zero-interest schedule at zero rather than refusing it', () => {
    // Ten payments of 100 repay 1,000 exactly: the root is r = 0, and the
    // iteration has to land on it rather than divide by it.
    expect(Math.abs(num('RATE(10,-100,1000)'))).toBeLessThan(1e-12);
    // A negative rate is a legitimate root too.
    near('RATE(10,-50,1000)', -0.10956029, 8);
    near(`PMT(${num('RATE(10,-50,1000)')},10,1000)`, -50, 8);
  });

  it('reports #NUM! when IRR has no sign change to work with', () => {
    expect(code('IRR({100,50,60})')).toBe('#NUM!');
    expect(code('IRR({-100,-50,-60})')).toBe('#NUM!');
    expect(code('IRR({-100})')).toBe('#NUM!');
  });

  it('finds the root nearest the guess when a series has more than one', () => {
    // A flow that changes sign twice has two internal rates; Excel returns the
    // one the guess is closest to, which is the defining property of Newton
    // started there rather than of a bracketing search.
    const series = '{-100,300,-200}';
    const low = num(`IRR(${series},0.05)`);
    const high = num(`IRR(${series},0.9)`);
    expect(low).toBeLessThan(high);
    for (const r of [low, high]) {
      const npv = -100 + 300 / (1 + r) - 200 / Math.pow(1 + r, 2);
      expect(Math.abs(npv)).toBeLessThan(1e-8);
    }
  });

  it('rejects a guess that is not a rate', () => {
    expect(code('IRR({-100,50,60,70},-1)')).toBe('#NUM!');
    expect(code('XIRR({-100,110},{1,366},-2)')).toBe('#NUM!');
  });

  it('reports #DIV/0! when MIRR has nothing on one side', () => {
    expect(code('MIRR({100,50,60},0.1,0.12)')).toBe('#DIV/0!');
    expect(code('MIRR({-100,-50,-60},0.1,0.12)')).toBe('#DIV/0!');
  });

  it('rejects an XNPV schedule that is not aligned or not increasing', () => {
    expect(code(`XNPV(0.1,{-100,110},{${d(2020, 1, 1)}})`)).toBe('#NUM!');
    // A date before the first one has a negative exponent Excel refuses.
    expect(code(`XNPV(0.1,{-100,110},{${d(2020, 1, 1)},${d(2019, 1, 1)}})`)).toBe('#NUM!');
    expect(code(`XIRR({-100,110},{${d(2020, 1, 1)},${d(2019, 1, 1)}})`)).toBe('#NUM!');
  });
});

describe('argument-domain errors in depreciation', () => {
  it('rejects an impossible SYD', () => {
    expect(code('SYD(30000,7500,0,1)')).toBe('#NUM!');
    expect(code('SYD(30000,-1,10,1)')).toBe('#NUM!');
    expect(code('SYD(30000,7500,10,11)')).toBe('#NUM!');
    expect(code('SYD(30000,7500,10,0)')).toBe('#NUM!');
  });

  it('rejects an impossible DDB and DB', () => {
    expect(code('DDB(2400,300,10,11)')).toBe('#NUM!');
    expect(code('DDB(2400,300,10,1,0)')).toBe('#NUM!');
    expect(code('DDB(-2400,300,10,1)')).toBe('#NUM!');
    expect(code('DB(1000000,100000,6,8,7)')).toBe('#NUM!');
    expect(code('DB(1000000,100000,6,1,13)')).toBe('#NUM!');
    expect(code('DB(1000000,100000,6,1,0)')).toBe('#NUM!');
  });

  it('rejects an impossible VDB span', () => {
    expect(code('VDB(2400,300,10,5,11)')).toBe('#NUM!');
    expect(code('VDB(2400,300,10,5,4)')).toBe('#NUM!');
    expect(code('VDB(2400,3000,10,0,1)')).toBe('#NUM!');
  });

  it('accepts a factor large enough to write the asset off at once', () => {
    expect(num('DDB(1000,0,5,1,5)')).toBe(1000);
    expect(num('DDB(1000,0,5,2,5)')).toBe(0);
  });

  it('holds depreciation at the salvage floor', () => {
    // Once the declining balance would fall below salvage, the period's charge
    // is only the distance left to it, and every later period is nothing.
    expect(num('DDB(1000,900,5,4,2)')).toBeCloseTo(0, 10);
    expect(num('DDB(1000,900,5,5,2)')).toBe(0);
  });

  it('leaves a zero-cost asset at zero rather than dividing by it', () => {
    expect(num('DB(0,0,5,1)')).toBe(0);
  });

  it('rejects the day-count bases the French schemes are not defined on', () => {
    expect(code(`AMORLINC(2400,${d(2008, 8, 19)},${d(2008, 12, 31)},300,1,0.15,2)`)).toBe(
      '#NUM!',
    );
    expect(code(`AMORDEGRC(2400,${d(2008, 8, 19)},${d(2008, 12, 31)},300,1,0.15,2)`)).toBe(
      '#NUM!',
    );
    expect(code(`AMORLINC(2400,${d(2008, 8, 19)},${d(2008, 12, 31)},2500,1,0.15,1)`)).toBe(
      '#NUM!',
    );
  });

  it('runs AMORLINC off the end of the schedule to zero', () => {
    const args = `2400,${d(2008, 8, 19)},${d(2008, 12, 31)},300,%,0.15,1`;
    // Six full periods of 360 plus a stub cannot exceed 2100 of depreciation.
    let total = 0;
    for (let period = 0; period <= 9; period++) {
      total += num(args.replace('%', String(period)).replace(/^/, 'AMORLINC(') + ')');
    }
    expect(total).toBeCloseTo(2100, 6);
    expect(num(`AMORLINC(${args.replace('%', '9')})`)).toBe(0);
  });
});

describe('argument-domain errors in the securities functions', () => {
  const bond = `${d(2011, 1, 25)},${d(2011, 11, 15)}`;

  it('rejects a frequency that is not annual, semi-annual or quarterly', () => {
    for (const frequency of [0, 3, 5, 12, -2]) {
      expect(code(`COUPNUM(${bond},${frequency},0)`), `frequency ${frequency}`).toBe('#NUM!');
    }
  });

  it('rejects a basis outside zero to four', () => {
    for (const basis of [-1, 5, 9]) {
      expect(code(`COUPDAYS(${bond},2,${basis})`), `basis ${basis}`).toBe('#NUM!');
    }
  });

  it('rejects settlement on or after maturity', () => {
    expect(code(`COUPNUM(${d(2011, 11, 15)},${d(2011, 11, 15)},2,0)`)).toBe('#NUM!');
    expect(code(`COUPNUM(${d(2012, 1, 1)},${d(2011, 11, 15)},2,0)`)).toBe('#NUM!');
    expect(code(`PRICE(${d(2012, 1, 1)},${d(2011, 11, 15)},0.05,0.05,100,2)`)).toBe('#NUM!');
    expect(code(`DISC(${d(2012, 1, 1)},${d(2011, 11, 15)},97,100)`)).toBe('#NUM!');
  });

  it('rejects a negative date', () => {
    expect(code('COUPNUM(-1,40000,2,0)')).toBe('#NUM!');
    expect(code('TBILLPRICE(-1,40000,0.09)')).toBe('#NUM!');
  });

  it('rejects the impossible price, rate and redemption values', () => {
    expect(code(`PRICE(${bond},-0.01,0.05,100,2)`)).toBe('#NUM!');
    expect(code(`PRICE(${bond},0.05,-0.01,100,2)`)).toBe('#NUM!');
    expect(code(`PRICE(${bond},0.05,0.05,0,2)`)).toBe('#NUM!');
    expect(code(`YIELD(${bond},0.05,0,100,2)`)).toBe('#NUM!');
    expect(code(`DISC(${bond},0,100)`)).toBe('#NUM!');
    expect(code(`ACCRINTM(${bond},0,1000)`)).toBe('#NUM!');
  });

  it('refuses a Treasury bill maturing more than a year out', () => {
    expect(code(`TBILLPRICE(${d(2008, 3, 31)},${d(2009, 6, 1)},0.09)`)).toBe('#NUM!');
    expect(code(`TBILLEQ(${d(2008, 3, 31)},${d(2009, 6, 1)},0.09)`)).toBe('#NUM!');
    expect(code(`TBILLYIELD(${d(2008, 3, 31)},${d(2009, 6, 1)},98)`)).toBe('#NUM!');
    expect(code(`TBILLPRICE(${d(2008, 6, 1)},${d(2008, 3, 31)},0.09)`)).toBe('#NUM!');
  });

  it('rejects an impossible rate conversion', () => {
    expect(code('EFFECT(0,4)')).toBe('#NUM!');
    expect(code('EFFECT(-0.05,4)')).toBe('#NUM!');
    expect(code('EFFECT(0.05,0)')).toBe('#NUM!');
    expect(code('NOMINAL(0,4)')).toBe('#NUM!');
    expect(code('NOMINAL(0.05,0.5)')).toBe('#NUM!');
    expect(code('PDURATION(0,2000,2200)')).toBe('#NUM!');
    expect(code('PDURATION(0.025,-2000,2200)')).toBe('#NUM!');
    expect(code('RRI(0,10000,11000)')).toBe('#NUM!');
    expect(code('RRI(10,0,11000)')).toBe('#NUM!');
    expect(code('RRI(10,1000,-11000)')).toBe('#NUM!');
  });

  it('separates a fraction of zero from a negative one', () => {
    // Microsoft documents these as two different errors, and they are.
    expect(code('DOLLARDE(1.02,0)')).toBe('#DIV/0!');
    expect(code('DOLLARDE(1.02,0.5)')).toBe('#DIV/0!');
    expect(code('DOLLARDE(1.02,-1)')).toBe('#NUM!');
    expect(code('DOLLARFR(1.125,0)')).toBe('#DIV/0!');
    expect(code('DOLLARFR(1.125,-1)')).toBe('#NUM!');
  });

  it('rejects the CUMIPMT and CUMPRINC arguments Excel documents as #NUM!', () => {
    const loan = '0.09/12,360,125000';
    expect(code(`CUMIPMT(0,360,125000,1,12,0)`)).toBe('#NUM!');
    expect(code(`CUMIPMT(${loan},0,12,0)`)).toBe('#NUM!');
    expect(code(`CUMIPMT(${loan},13,12,0)`)).toBe('#NUM!');
    expect(code(`CUMIPMT(${loan},1,361,0)`)).toBe('#NUM!');
    // A type that is neither 0 nor 1 is accepted everywhere else and rejected
    // here, which is a documented quirk rather than an oversight.
    expect(code(`CUMIPMT(${loan},1,12,2)`)).toBe('#NUM!');
    expect(code(`CUMPRINC(${loan},1,12,2)`)).toBe('#NUM!');
    expect(code('CUMPRINC(0.09/12,360,-125000,1,12,0)')).toBe('#NUM!');
  });

  it('rejects a period outside the schedule in IPMT and PPMT', () => {
    expect(code('IPMT(0.1,0,10,1000)')).toBe('#NUM!');
    expect(code('IPMT(0.1,11,10,1000)')).toBe('#NUM!');
    expect(code('PPMT(0.1,0,10,1000)')).toBe('#NUM!');
    expect(code('PPMT(0.1,11,10,1000)')).toBe('#NUM!');
  });
});

// ---------------------------------------------------------------------------
// Day counts and coupon arithmetic
// ---------------------------------------------------------------------------

describe('coupon-period arithmetic', () => {
  it('splits every coupon period into the two halves that make it up', () => {
    // Bases 2 and 3 are excluded on purpose: they measure the two halves in
    // actual days but the whole period as 360 or 365 over the frequency, so
    // in Excel the parts genuinely do not add up to the whole. The test below
    // pins that discrepancy down rather than hiding it here.
    for (const basis of [0, 1, 4]) {
      for (const frequency of [1, 2, 4]) {
        const args = `${d(2021, 3, 17)},${d(2029, 8, 31)},${frequency},${basis}`;
        const before = num(`COUPDAYBS(${args})`);
        const after = num(`COUPDAYSNC(${args})`);
        const whole = num(`COUPDAYS(${args})`);
        expect(before + after, `basis ${basis} frequency ${frequency}`).toBeCloseTo(whole, 9);
      }
    }
  });

  it('leaves the halves of a basis 2 or 3 period not adding up, as Excel does', () => {
    const args = (basis: number): string => `${d(2021, 3, 17)},${d(2029, 8, 31)},1,${basis}`;
    // 28 February 2021 to 28 February 2022 is 365 actual days, while
    // COUPDAYS reports the convention's 360.
    expect(num(`COUPDAYBS(${args(2)})`) + num(`COUPDAYSNC(${args(2)})`)).toBe(365);
    expect(num(`COUPDAYS(${args(2)})`)).toBe(360);
  });

  it('brackets settlement between the two coupon dates', () => {
    const settle = d(2021, 3, 17);
    for (const basis of [0, 1, 3]) {
      for (const frequency of [1, 2, 4]) {
        const args = `${settle},${d(2029, 8, 31)},${frequency},${basis}`;
        expect(num(`COUPPCD(${args})`)).toBeLessThanOrEqual(settle);
        expect(num(`COUPNCD(${args})`)).toBeGreaterThan(settle);
      }
    }
  });

  it('keeps an end-of-month maturity on the end of the month', () => {
    // 31 August steps back six months to 28 February, not to the 28th of every
    // subsequent period: the next one is 31 August again.
    expect(num(`COUPNCD(${d(2011, 1, 25)},${d(2011, 8, 31)},2,0)`)).toBe(d(2011, 2, 28));
    expect(num(`COUPPCD(${d(2011, 4, 1)},${d(2011, 8, 31)},2,0)`)).toBe(d(2011, 2, 28));
    expect(num(`COUPNCD(${d(2011, 4, 1)},${d(2011, 8, 31)},2,0)`)).toBe(d(2011, 8, 31));
    // A leap year moves that same coupon to the 29th.
    expect(num(`COUPNCD(${d(2012, 1, 25)},${d(2012, 8, 31)},2,0)`)).toBe(d(2012, 2, 29));
  });

  it('keeps a mid-month maturity on its own day of month', () => {
    // 30 May is not the end of May, so the 30th is preserved rather than made
    // sticky: three months back it clamps to 28 February, and stepping on it
    // is a 30th again.
    expect(num(`COUPPCD(${d(2021, 3, 15)},${d(2021, 5, 30)},4,0)`)).toBe(d(2021, 2, 28));
    expect(num(`COUPNCD(${d(2021, 3, 15)},${d(2021, 5, 30)},4,0)`)).toBe(d(2021, 5, 30));
    expect(num(`COUPPCD(${d(2020, 12, 15)},${d(2021, 5, 30)},4,0)`)).toBe(d(2020, 11, 30));
    // 30 April is the end of April, so that same schedule is sticky instead
    // and lands on 31 January rather than on the 30th.
    expect(num(`COUPPCD(${d(2021, 3, 15)},${d(2021, 4, 30)},4,0)`)).toBe(d(2021, 1, 31));
  });

  it('counts a coupon for every period from the previous coupon to maturity', () => {
    for (const frequency of [1, 2, 4]) {
      const args = `${d(2021, 3, 17)},${d(2029, 8, 31)},${frequency},0`;
      const count = num(`COUPNUM(${args})`);
      const previous = num(`COUPPCD(${args})`);
      // Stepping COUPNCD forward COUPNUM times from the previous coupon date
      // must land exactly on maturity.
      let cursor = previous;
      for (let i = 0; i < count; i++) {
        cursor = num(`COUPNCD(${cursor},${d(2029, 8, 31)},${frequency},0)`);
      }
      expect(cursor, `frequency ${frequency}`).toBe(d(2029, 8, 31));
    }
  });

  it('gives 360 days a year on the 30/360 bases and 365 or 366 on actual', () => {
    const args = (basis: number): string => `${d(2021, 3, 17)},${d(2029, 8, 31)},2,${basis}`;
    expect(num(`COUPDAYS(${args(0)})`)).toBe(180);
    expect(num(`COUPDAYS(${args(2)})`)).toBe(180);
    expect(num(`COUPDAYS(${args(4)})`)).toBe(180);
    expect(num(`COUPDAYS(${args(3)})`)).toBe(182.5);
    // Basis 1 counts the actual days of the period settlement falls in:
    // 28 February 2021 to 31 August 2021.
    expect(num(`COUPDAYS(${args(1)})`)).toBe(d(2021, 8, 31) - d(2021, 2, 28));
  });

  it('folds the end of February onto the 30th on basis 0 but not basis 4', () => {
    // NASD 30/360 treats 28 February as a 30th when it is the start date; the
    // European convention does not, so the two disagree by two days.
    const nasd = num(`ACCRINTM(${d(2021, 2, 28)},${d(2021, 6, 30)},0.1,36000,0)`);
    const euro = num(`ACCRINTM(${d(2021, 2, 28)},${d(2021, 6, 30)},0.1,36000,4)`);
    expect(nasd).toBeCloseTo(1200, 6); // 120 days of 360
    expect(euro).toBeCloseTo(1220, 6); // 122 days of 360
  });

  it('accrues linearly on the linear bases', () => {
    // Half a 30/360 year at ten per cent on 1,000 par is 50, whatever the
    // coupon frequency, because the quasi-periods tile the interval exactly.
    for (const frequency of [1, 2, 4]) {
      near(
        `ACCRINT(${d(2020, 1, 1)},${d(2021, 1, 1)},${d(2020, 7, 1)},0.1,1000,${frequency},0)`,
        50,
        8,
      );
    }
  });

  it('agrees with ACCRINTM over one whole period on a linear basis', () => {
    const issue = d(2020, 1, 15);
    const settle = d(2020, 7, 15);
    near(`ACCRINT(${issue},${settle},${settle},0.08,1000,2,0)`, 40, 8);
    near(`ACCRINTM(${issue},${settle},0.08,1000,0)`, 40, 8);
  });

  it('normalises ACCRINT by the coupon period and ACCRINTM by the year', () => {
    // On actual/actual the two disagree, and correctly so: half a coupon
    // period is not half a year when the year has 366 days in it.
    const issue = d(2020, 1, 15);
    const settle = d(2020, 7, 15);
    near(`ACCRINT(${issue},${settle},${settle},0.08,1000,2,1)`, 40, 8);
    near(`ACCRINTM(${issue},${settle},0.08,1000,1)`, (1000 * 0.08 * 182) / 366, 8);
  });
});

// ---------------------------------------------------------------------------
// Magnitudes and boundaries
// ---------------------------------------------------------------------------

describe('boundaries and magnitudes', () => {
  it('overflows to #NUM! rather than to infinity', () => {
    expect(code('FV(9,10000,-1)')).toBe('#NUM!');
    expect(code('FVSCHEDULE(1E+300,{1000,1000,1000})')).toBe('#NUM!');
  });

  it('survives a rate small enough to be indistinguishable from zero', () => {
    near('PMT(1E-15,120,-120000)', 1000, 6);
    near('FV(1E-300,10,-100)', 1000, 6);
  });

  it('survives a very long term', () => {
    // A thousand-year monthly annuity: the growth factor is large but finite.
    expect(num('PV(0.05/12,12000,-1000)')).toBeCloseTo(240000, 0);
  });

  it('handles a single-period annuity', () => {
    expect(num('PMT(0.1,1,-1000)')).toBeCloseTo(1100, 8);
    expect(num('RATE(1,-1100,1000)')).toBeCloseTo(0.1, 10);
    expect(num('IPMT(0.1,1,1,-1000)')).toBeCloseTo(100, 8);
    expect(num('PPMT(0.1,1,1,-1000)')).toBeCloseTo(1000, 8);
  });

  it('handles a negative rate', () => {
    // Deflation is a legitimate input, not an error.
    expect(num('PMT(-0.01,10,-1000)')).toBeGreaterThan(0);
    expect(num('FV(-0.01,10,-100)')).toBeGreaterThan(0);
    near('PV(-0.02,5,0,-1000)', 1000 / Math.pow(0.98, 5), 6);
  });

  it('gives zero depreciation for a zero-cost or fully depreciated asset', () => {
    expect(num('SLN(1000,1000,5)')).toBe(0);
    expect(num('SYD(1000,1000,5,1)')).toBe(0);
    expect(num('DDB(1000,1000,5,1)')).toBe(0);
    expect(num('VDB(1000,1000,5,0,5)')).toBe(0);
  });

  it('sums a whole VDB schedule to the depreciable amount', () => {
    near('VDB(2400,300,10,0,10)', 2100, 6);
    // The salvage floor stops the no-switch form at the same total here.
    near('VDB(2400,300,10,0,10,2,TRUE)', 2100, 6);
  });

  it('splits a VDB span into the same total as two adjacent spans', () => {
    const whole = num('VDB(2400,300,10,0,7)');
    const first = num('VDB(2400,300,10,0,3)');
    const second = num('VDB(2400,300,10,3,7)');
    expect(first + second).toBeCloseTo(whole, 8);
  });

  it('honours no_switch by staying on the declining balance', () => {
    // With no salvage floor to stop it, declining balance keeps shrinking
    // while the switching form has moved to the larger straight-line charge.
    near('VDB(2400,0,10,7,10,2,TRUE)', 245.61844224, 6);
    near('VDB(2400,0,10,7,10)', 471.8592, 6);
    // The whole schedule shows the same thing: declining balance alone never
    // finishes writing the asset off, while the switch does.
    near('VDB(2400,0,10,0,10)', 2400, 6);
    expect(num('VDB(2400,0,10,0,10,2,TRUE)')).toBeLessThan(2400);
  });

  it('prices a par bond at par', () => {
    const bond = `${d(2020, 1, 1)},${d(2030, 1, 1)},0.05,0.05,100,2,0`;
    near(`PRICE(${bond})`, 100, 8);
  });

  it('gives a premium bond a duration below its term and above zero', () => {
    const dur = num(`DURATION(${d(2020, 1, 1)},${d(2030, 1, 1)},0.05,0.03,2,0)`);
    expect(dur).toBeGreaterThan(0);
    expect(dur).toBeLessThan(10);
    // Modified duration is Macaulay duration discounted one period.
    near(
      `MDURATION(${d(2020, 1, 1)},${d(2030, 1, 1)},0.05,0.03,2,0)`,
      dur / (1 + 0.03 / 2),
      10,
    );
  });

  it('gives a zero-coupon bond a duration equal to its term', () => {
    // With no coupons every unit of value sits at maturity.
    near(`DURATION(${d(2020, 1, 1)},${d(2030, 1, 1)},0,0.04,1,0)`, 10, 8);
  });
});

// ---------------------------------------------------------------------------
// Metadata
// ---------------------------------------------------------------------------

describe('metadata', () => {
  it('marks nothing volatile or structural', () => {
    for (const spec of FINANCIAL_FUNCTIONS) {
      expect(spec.volatile, spec.name).toBeUndefined();
      expect(spec.structural, spec.name).toBeUndefined();
    }
  });

  it('broadcasts the scalar functions and not the series functions', () => {
    const byName = new Map(FINANCIAL_FUNCTIONS.map((s) => [s.name, s]));
    for (const name of ['PMT', 'FV', 'PRICE', 'COUPNUM', 'DDB']) {
      expect(byName.get(name)?.broadcast, name).toBe(true);
    }
    // These consume ranges, so an array argument is data rather than a request
    // to map over it.
    for (const name of ['NPV', 'IRR', 'XIRR', 'XNPV', 'MIRR', 'FVSCHEDULE']) {
      expect(byName.get(name)?.broadcast, name).toBeUndefined();
    }
  });

  it('registers every function exactly once', () => {
    const names = FINANCIAL_FUNCTIONS.map((s) => s.name);
    expect(new Set(names).size).toBe(names.length);
    expect(names).toHaveLength(51);
  });
});
