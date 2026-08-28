/**
 * Engineering and database functions.
 *
 * The oracle does not cover this category. Neither formulas.calc.xlsx nor
 * precedence.calc.xlsx contains a single case for a base conversion, a bitwise
 * operation, a complex number, CONVERT, ERF, BESSEL or a D function - the first
 * block below asserts exactly that, so if the fixture generator ever grows them
 * this file fails and stops claiming the gap exists. Every expected value here
 * is therefore specified from Microsoft's published documentation for the
 * function, and the documented worked examples are quoted verbatim wherever
 * Microsoft supplies one.
 *
 * Two exceptions, where something better than a documented example is
 * available. The D functions are checked against the oracle indirectly: the
 * fixture's SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, SUM, MIN, MAX, COUNT,
 * STDEV.S and VAR.S cases all describe the same eight employee records this
 * category can select with a criteria range, so DSUM and its siblings are
 * required to reproduce the numbers LibreOffice actually computed for them. The
 * Bessel functions are checked against their recurrence relations and against
 * the Wronskians J(n+1)Y(n) - J(n)Y(n+1) = 2/(pi x) and
 * I(n+1)K(n) + I(n)K(n+1) = 1/x, which pin all four families together to the
 * last few bits and would catch any error a table of remembered values could.
 *
 * The registry under test holds ENGINEERING_FUNCTIONS alone, so a failure here
 * is never another category's fault.
 */

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, type Scalar, Workbook, isError } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { Evaluator } from '../src/evaluator.js';
import { ENGINEERING_FUNCTIONS } from '../src/functions/engineering.js';
import { parseFormula } from '../src/parser.js';
import { FunctionRegistry, storageName } from '../src/registry.js';
import { WorkbookStore } from '../src/store.js';
import { toExcelPrecision } from '../src/value.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const registry = new FunctionRegistry().registerAll(ENGINEERING_FUNCTIONS);

// ---------------------------------------------------------------------------
// The fixture, and the gap in it
// ---------------------------------------------------------------------------

const { workbook: oracleBook } = readXlsx(
  new Uint8Array(readFileSync(new URL('formulas.calc.xlsx', FIXTURES))),
);
const oracleSheet = oracleBook.getSheet('Formulas')!;
const ORACLE = new Map<string, Scalar>();
for (const { row, col, cell } of oracleSheet.entries()) {
  if (col !== 2 || !cell.formula) continue;
  const name = oracleSheet.getValue(row, 0);
  if (typeof name === 'string') ORACLE.set(name, cell.value);
}

/** Every formula in both oracle workbooks, on every sheet and in every cell. */
const ORACLE_FORMULAS: string[] = [];
for (const file of ['formulas.calc.xlsx', 'precedence.calc.xlsx']) {
  const { workbook } = readXlsx(new Uint8Array(readFileSync(new URL(file, FIXTURES))));
  for (const sheet of workbook.sheets) {
    for (const { cell } of sheet.entries()) if (cell.formula) ORACLE_FORMULAS.push(cell.formula);
  }
}

function oracleNumber(name: string): number {
  const value = ORACLE.get(name);
  expect(typeof value, `case ${name} missing from the fixture`).toBe('number');
  return value as number;
}

describe('the oracle', () => {
  it('has no case for any function in this category, on any sheet of either file', () => {
    const names = new Set(ENGINEERING_FUNCTIONS.map((f) => f.name));
    const used = ORACLE_FORMULAS.flatMap((f) => [...f.matchAll(/([A-Z][A-Z0-9._]*)\s*\(/gi)])
      .map((m) => m[1]!.toUpperCase())
      .filter((n) => names.has(n));
    expect(used).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// The D functions, against the fixture's own aggregates
// ---------------------------------------------------------------------------

/**
 * Criteria ranges written into the fixture's Data sheet, to the right of its
 * five used columns. The sheet holds eight employees with n, name, dept, salary
 * and hired columns, which is exactly the shape the D functions want.
 */
const oracleData = oracleBook.getSheet('Data')!;
for (const [row, col, value] of [
  [0, 6, 'dept'], [1, 6, 'Eng'],
  [0, 7, 'dept'], [0, 8, 'salary'], [1, 7, 'Eng'], [1, 8, '>170000'],
  [0, 9, 'dept'], [1, 9, 'Science'],
  [0, 10, 'dept'], [0, 11, 'salary'], [1, 10, 'Ops'], [1, 11, '<120000'],
  // A criteria range that is nothing but a header selects every record.
  [0, 12, 'salary'],
] as const) {
  oracleData.setValue(row, col, value);
}

const oracleEval = new Evaluator(new WorkbookStore(oracleBook), registry, {
  dateSystem: oracleBook.dateSystem,
});

function onFixture(formula: string): Scalar {
  return oracleEval.evaluateScalar({
    ast: parseFormula(formula, { origin: { row: 0, col: 2 } }),
    sheet: 'Formulas',
    row: 0,
    col: 2,
  });
}

/** Compare at the fifteen digits the oracle stored and Excel compares at. */
function sameNumber(actual: Scalar, expected: number): void {
  expect(typeof actual, `expected a number, got ${JSON.stringify(actual)}`).toBe('number');
  expect(toExcelPrecision(actual as number)).toBe(toExcelPrecision(expected));
}

describe('oracle: the D functions reproduce the fixture aggregates', () => {
  const db = 'Data!A1:E9';

  it('DSUM matches SUMIF over the same records', () => {
    sameNumber(onFixture(`DSUM(${db},"salary",Data!G1:G2)`), oracleNumber('SUMIF'));
  });

  it('DAVERAGE matches AVERAGEIF over the same records', () => {
    sameNumber(onFixture(`DAVERAGE(${db},"salary",Data!G1:G2)`), oracleNumber('AVERAGEIF'));
  });

  it('DSUM with two criteria columns matches SUMIFS', () => {
    sameNumber(onFixture(`DSUM(${db},"salary",Data!H1:I2)`), oracleNumber('SUMIFS'));
  });

  it('DCOUNT matches COUNTIF', () => {
    sameNumber(onFixture(`DCOUNT(${db},"salary",Data!J1:J2)`), oracleNumber('COUNTIF'));
  });

  it('DCOUNT with two criteria columns matches COUNTIFS', () => {
    sameNumber(onFixture(`DCOUNT(${db},"salary",Data!K1:L2)`), oracleNumber('COUNTIFS'));
  });

  it('a header-only criteria range selects every record', () => {
    sameNumber(onFixture(`DSUM(${db},"salary",Data!M1:M1)`), oracleNumber('SUM'));
    sameNumber(onFixture(`DAVERAGE(${db},"salary",Data!M1:M1)`), oracleNumber('AVERAGE'));
    sameNumber(onFixture(`DMIN(${db},"salary",Data!M1:M1)`), oracleNumber('MIN'));
    sameNumber(onFixture(`DMAX(${db},"salary",Data!M1:M1)`), oracleNumber('MAX'));
    sameNumber(onFixture(`DCOUNT(${db},"salary",Data!M1:M1)`), oracleNumber('COUNT'));
  });

  it('DSTDEV and DVAR match the sample statistics', () => {
    sameNumber(onFixture(`DSTDEV(${db},"salary",Data!M1:M1)`), oracleNumber('STDEV.S'));
    sameNumber(onFixture(`DVAR(${db},"salary",Data!M1:M1)`), oracleNumber('VAR.S'));
  });

  it('DGET returns the one record a name selects', () => {
    // Grace's salary, which the fixture's INDEX_MATCH case also picks out.
    oracleData.setValue(0, 13, 'name');
    oracleData.setValue(1, 13, 'Grace');
    sameNumber(onFixture(`DGET(${db},"salary",Data!N1:N2)`), oracleNumber('INDEX_MATCH'));
  });
});

// ---------------------------------------------------------------------------
// A synthetic workbook for everything the fixture cannot reach
// ---------------------------------------------------------------------------

/**
 * Sheet S carries the same eight employees in A1:E9, criteria blocks in G:S,
 * and a scratch row at 20 holding a blank, an error, numeric text and a
 * boolean.
 */
function makeBook(): Workbook {
  const wb = new Workbook();
  const s = wb.addSheet('S');
  const rows: (string | number)[][] = [
    ['n', 'name', 'dept', 'salary', 'hired'],
    [1, 'Ada', 'Eng', 165000, 43525],
    [2, 'Grace', 'Eng', 172000, 43296],
    [3, 'Katherine', 'Science', 158000, 43850],
    [4, 'Dorothy', 'Science', 149000, 44502],
    [5, 'Margaret', 'Eng', 181000, 42885],
    [6, 'Annie', 'Ops', 121000, 44816],
    [7, 'Henrietta', 'Ops', 118000, 44985],
    [8, 'Mary', 'Science', 143000, 44020],
  ];
  rows.forEach((row, r) => row.forEach((value, c) => s.setValue(r, c, value)));

  for (const [row, col, value] of [
    [0, 6, 'dept'], [1, 6, 'Eng'],                                     // G1:G2
    [0, 7, 'dept'], [0, 8, 'salary'], [1, 7, 'Eng'], [1, 8, '>170000'], // H1:I3
    [2, 7, 'Ops'],
    [0, 9, 'name'], [1, 9, 'Annie'],                                    // J1:J2
    [0, 10, 'name'], [1, 10, 'Zeno'],                                   // K1:K2
    [0, 11, 'dept'],                                                    // L1:L2, blank criterion
    [0, 12, 'nosuchfield'], [1, 12, 'x'],                               // M1:M2
    [0, 13, 'name'], [1, 13, 'A'],                                      // N1:N2, prefix match
    [0, 14, 'name'], [1, 14, '?ary'],                                   // O1:O2, wildcard
    [0, 15, 'dept'], [1, 15, 'Ops'],                                    // P1:P2
  ] as const) {
    s.setValue(row, col, value);
  }

  s.setValue(19, 1, '0011');
  s.setValue(19, 2, 255);
  s.setValue(19, 3, true);

  // Two small databases well below the first: A26:B28 carries an error in its
  // value column, F26:G29 a mix of text, a blank and a number. D26 and I26 are
  // header-only criteria ranges over them, which select every record.
  s.setValue(25, 0, 'k');
  s.setValue(25, 1, 'v');
  s.setValue(26, 0, 1);
  s.setValue(26, 1, CellError.DIV0);
  s.setValue(27, 0, 2);
  s.setValue(27, 1, 5);
  s.setValue(25, 3, 'k');

  s.setValue(25, 5, 'k');
  s.setValue(25, 6, 'v');
  s.setValue(26, 5, 1);
  s.setValue(26, 6, 'text');
  s.setValue(27, 5, 2);
  s.setValue(28, 5, 3);
  s.setValue(28, 6, 7);
  s.setValue(25, 8, 'k');
  return wb;
}

const book = makeBook();
const ev = new Evaluator(new WorkbookStore(book), registry, { dateSystem: book.dateSystem });

/** Evaluate at U30, well away from the data, and collapse to what a cell shows. */
function run(formula: string): Scalar {
  return ev.evaluateScalar({
    ast: parseFormula(formula, { origin: { row: 29, col: 20 } }),
    sheet: 'S',
    row: 29,
    col: 20,
  });
}

function num(formula: string): number {
  const value = run(formula);
  expect(typeof value, `${formula} gave ${JSON.stringify(value)}`).toBe('number');
  return value as number;
}

/** Relative agreement, for the transcendental results. */
function near(formula: string, expected: number, tolerance = 1e-14): void {
  const actual = num(formula);
  const scale = Math.max(Math.abs(expected), Number.MIN_VALUE);
  expect(Math.abs(actual - expected) / scale, `${formula} gave ${actual}`).toBeLessThan(tolerance);
}

function expectAll(cases: readonly (readonly [string, Scalar])[]): void {
  for (const [formula, expected] of cases) {
    expect(run(formula), formula).toEqual(expected);
  }
}

// ---------------------------------------------------------------------------
// Base conversion
// ---------------------------------------------------------------------------

describe('base conversion', () => {
  it('reproduces the documented examples', () => {
    expectAll([
      ['BIN2DEC(1100100)', 100],
      ['BIN2DEC(1111111111)', -1],
      ['BIN2HEX(11111011,4)', '00FB'],
      ['BIN2HEX(1110)', 'E'],
      ['BIN2OCT(1001,3)', '011'],
      ['BIN2OCT(1100100)', '144'],
      ['DEC2BIN(9,4)', '1001'],
      ['DEC2BIN(-100)', '1110011100'],
      ['DEC2HEX(100,4)', '0064'],
      ['DEC2HEX(-54)', 'FFFFFFFFCA'],
      ['DEC2OCT(58)', '72'],
      ['DEC2OCT(58,3)', '072'],
      ['DEC2OCT(-100)', '7777777634'],
      ['HEX2BIN("F",8)', '00001111'],
      ['HEX2BIN("B7")', '10110111'],
      ['HEX2DEC("A5")', 165],
      ['HEX2DEC("FFFFFFFF5B")', -165],
      ['HEX2DEC("3DA408B9")', 1034160313],
      ['HEX2OCT("F",3)', '017'],
      ['HEX2OCT("3B4E")', '35516'],
      ['OCT2BIN(3,3)', '011'],
      ['OCT2DEC(54)', 44],
      ['OCT2DEC(7777777533)', -165],
      ['OCT2HEX(100,4)', '0040'],
    ]);
  });

  it('encodes negatives as ten-digit two-s complement', () => {
    expectAll([
      ['DEC2BIN(-1)', '1111111111'],
      ['DEC2OCT(-1)', '7777777777'],
      ['DEC2HEX(-1)', 'FFFFFFFFFF'],
      ['BIN2OCT(1111111111)', '7777777777'],
      ['BIN2HEX(1111111111)', 'FFFFFFFFFF'],
      ['OCT2BIN(7777777000)', '1000000000'],
      ['HEX2BIN("FFFFFFFFFF")', '1111111111'],
      ['OCT2HEX(7777777533)', 'FFFFFFFF5B'],
    ]);
  });

  it('holds the width boundaries and refuses anything past them', () => {
    expectAll([
      ['DEC2BIN(511)', '111111111'],
      ['DEC2BIN(-512)', '1000000000'],
      ['DEC2BIN(512)', CellError.NUM],
      ['DEC2BIN(-513)', CellError.NUM],
      ['DEC2OCT(536870911)', '3777777777'],
      ['DEC2OCT(-536870912)', '4000000000'],
      ['DEC2OCT(536870912)', CellError.NUM],
      ['DEC2OCT(-536870913)', CellError.NUM],
      ['DEC2HEX(549755813887)', '7FFFFFFFFF'],
      ['DEC2HEX(-549755813888)', '8000000000'],
      ['DEC2HEX(549755813888)', CellError.NUM],
      ['DEC2HEX(-549755813889)', CellError.NUM],
      ['HEX2DEC("7FFFFFFFFF")', 549755813887],
      ['HEX2DEC("8000000000")', -549755813888],
      // A value that fits in the source width but not the target one.
      ['HEX2BIN("200")', CellError.NUM],
      ['HEX2BIN("FFFFFFFDFF")', CellError.NUM],
      ['OCT2BIN(2000)', CellError.NUM],
      ['HEX2OCT("20000000")', CellError.NUM],
    ]);
  });

  it('rejects digits the radix does not have, and numerals that are too long', () => {
    expectAll([
      ['BIN2DEC(2)', CellError.NUM],
      ['BIN2DEC("1a")', CellError.NUM],
      ['OCT2DEC(8)', CellError.NUM],
      ['HEX2DEC("G")', CellError.NUM],
      ['BIN2DEC(11111111111)', CellError.NUM],
      ['OCT2DEC(11111111111)', CellError.NUM],
      ['HEX2DEC("11111111111")', CellError.NUM],
      ['BIN2DEC("")', CellError.NUM],
      ['BIN2DEC(1.5)', CellError.NUM],
    ]);
  });

  it('pads to places and refuses a places that cannot hold the result', () => {
    expectAll([
      ['DEC2BIN(0)', '0'],
      ['DEC2BIN(0,3)', '000'],
      ['DEC2HEX(255,2)', 'FF'],
      ['DEC2HEX(255,1)', CellError.NUM],
      ['DEC2BIN(1,10)', '0000000001'],
      ['DEC2BIN(1,11)', CellError.NUM],
      ['DEC2BIN(9,0)', CellError.NUM],
      ['DEC2BIN(9,-1)', CellError.NUM],
      ['DEC2BIN(9,"x")', CellError.VALUE],
      // Places is truncated, not rounded.
      ['DEC2BIN(1,3.9)', '001'],
      // A negative result fills the width, and places is ignored outright.
      ['DEC2BIN(-1,2)', '1111111111'],
      ['DEC2HEX(-54,6)', 'FFFFFFFFCA'],
      ['OCT2HEX(7777777533,6)', 'FFFFFFFF5B'],
      // Ignored, but only after it has been read: a places that is not a number
      // is #VALUE! whatever the sign of the result, because the conversion
      // happens before the rule that discards it.
      ['DEC2BIN(-1,"x")', CellError.VALUE],
      ['DEC2HEX(-54,"x")', CellError.VALUE],
    ]);
  });

  it('reads hexadecimal case-insensitively but rejects logical values', () => {
    expectAll([
      ['HEX2DEC("ff")', 255],
      ['HEX2DEC("fF")', 255],
      ['DEC2BIN(TRUE)', CellError.VALUE],
      ['BIN2DEC(TRUE)', CellError.VALUE],
      ['DEC2BIN(D20)', CellError.VALUE],
    ]);
  });

  it('takes its input from cells, numerals included', () => {
    expectAll([
      // B20 holds the text 0011, C20 the number 255 - which HEX2DEC reads as a
      // hexadecimal numeral, not as the decimal it looks like.
      ['BIN2DEC(B20)', 3],
      ['HEX2DEC(C20)', 597],
      ['DEC2BIN("9")', '1001'],
      ['DEC2BIN(9.9)', '1001'],
      ['DEC2BIN(-9.9)', '1111110111'],
    ]);
  });

  it('propagates an error argument unchanged', () => {
    expectAll([
      ['DEC2BIN(1/0)', CellError.DIV0],
      ['BIN2DEC(1/0)', CellError.DIV0],
      ['DEC2BIN(1,1/0)', CellError.DIV0],
    ]);
  });
});

// ---------------------------------------------------------------------------
// Bitwise
// ---------------------------------------------------------------------------

describe('bitwise', () => {
  it('reproduces the documented examples', () => {
    expectAll([
      ['BITAND(13,25)', 9],
      ['BITAND(1,5)', 1],
      ['BITOR(23,10)', 31],
      ['BITXOR(5,3)', 6],
      ['BITLSHIFT(4,2)', 16],
      ['BITRSHIFT(13,2)', 3],
    ]);
  });

  it('works above the 32-bit boundary a JavaScript operator would truncate at', () => {
    expectAll([
      ['BITAND(281474976710655,281474976710655)', 281474976710655],
      ['BITXOR(281474976710655,1)', 281474976710654],
      ['BITOR(4294967296,1)', 4294967297],
      ['BITAND(4294967296,4294967296)', 4294967296],
      ['BITLSHIFT(1,47)', 140737488355328],
      ['BITRSHIFT(281474976710655,47)', 1],
    ]);
  });

  it('refuses negatives, fractions and anything past 2^48-1', () => {
    expectAll([
      ['BITAND(-1,1)', CellError.NUM],
      ['BITOR(1,-1)', CellError.NUM],
      ['BITAND(1.5,1)', CellError.NUM],
      ['BITXOR(281474976710656,1)', CellError.NUM],
      ['BITLSHIFT(281474976710655,1)', CellError.NUM],
      ['BITLSHIFT(1,48)', CellError.NUM],
      ['BITLSHIFT(1,54)', CellError.NUM],
      ['BITRSHIFT(1,-54)', CellError.NUM],
      ['BITLSHIFT(1,1.5)', CellError.NUM],
      ['BITAND(TRUE,1)', CellError.VALUE],
      ['BITAND("x",1)', CellError.VALUE],
    ]);
  });

  it('shifts the other way for a negative amount, and not at all for zero', () => {
    expectAll([
      ['BITRSHIFT(13,-2)', 52],
      ['BITLSHIFT(13,-2)', 3],
      ['BITLSHIFT(13,0)', 13],
      ['BITRSHIFT(0,10)', 0],
      // A right shift discards the bits that fall off the end.
      ['BITRSHIFT(1,1)', 0],
    ]);
  });
});

// ---------------------------------------------------------------------------
// Complex numbers
// ---------------------------------------------------------------------------

describe('complex numbers', () => {
  it('builds and formats the documented shapes', () => {
    expectAll([
      ['COMPLEX(3,4)', '3+4i'],
      ['COMPLEX(3,4,"j")', '3+4j'],
      ['COMPLEX(0,1)', 'i'],
      ['COMPLEX(0,-1)', '-i'],
      ['COMPLEX(1,-1)', '1-i'],
      ['COMPLEX(3,0)', '3'],
      ['COMPLEX(0,0)', '0'],
      ['COMPLEX(-3,-4)', '-3-4i'],
      ['COMPLEX(1.5,2.25)', '1.5+2.25i'],
      // The suffix is case-sensitive, and there are only two of them.
      ['COMPLEX(3,4,"I")', CellError.VALUE],
      ['COMPLEX(3,4,"k")', CellError.VALUE],
      ['COMPLEX(TRUE,1)', CellError.VALUE],
    ]);
  });

  it('picks a complex number apart', () => {
    expectAll([
      ['IMREAL("6-9i")', 6],
      ['IMAGINARY("6-9i")', -9],
      ['IMAGINARY("0-j")', -1],
      ['IMAGINARY("i")', 1],
      ['IMREAL("i")', 0],
      ['IMREAL(4)', 4],
      ['IMAGINARY(4)', 0],
      ['IMABS("5+12i")', 13],
      ['IMCONJUGATE("3+4i")', '3-4i'],
      ['IMREAL("1.5e2+2.5e-3i")', 150],
      ['IMAGINARY("1.5e2+2.5e-3i")', 0.0025],
      ['IMREAL("+3-4i")', 3],
    ]);
    near('IMARGUMENT("3+4i")', 0.927295218001612);
  });

  it('refuses text that is not a complex number, and logical values', () => {
    expectAll([
      ['IMREAL("abc")', CellError.NUM],
      ['IMREAL("")', CellError.NUM],
      // Excel accepts only a lower-case suffix.
      ['IMABS("3+4I")', CellError.NUM],
      ['IMABS("3+4k")', CellError.NUM],
      ['IMREAL("3i4")', CellError.NUM],
      ['IMREAL("3++4i")', CellError.NUM],
      ['IMREAL(TRUE)', CellError.VALUE],
      ['IMREAL(1/0)', CellError.DIV0],
      // Zero has no argument.
      ['IMARGUMENT("0")', CellError.DIV0],
      ['IMLN("0")', CellError.NUM],
      ['IMDIV("1","0")', CellError.NUM],
    ]);
  });

  it('carries the suffix through, and refuses to mix i with j', () => {
    expectAll([
      ['IMSUM("3+4j","5-3j")', '8+j'],
      ['IMSUB("13+4j","5+3j")', '8+j'],
      ['IMPRODUCT("3+4j","5-3j")', '27+11j'],
      ['IMCONJUGATE("3+4j")', '3-4j'],
      // A real operand has no suffix to disagree with.
      ['IMSUM(5,"1+i")', '6+i'],
      ['IMSUM(5,"1+j")', '6+j'],
      ['IMSUM("3+4i","5-3j")', CellError.VALUE],
      ['IMDIV("3+4i","5-3j")', CellError.VALUE],
      ['IMPRODUCT("i","j")', CellError.VALUE],
    ]);
  });

  it('reproduces the documented arithmetic', () => {
    expectAll([
      ['IMSUM("3+4i","5-3i")', '8+i'],
      ['IMSUB("13+4i","5+3i")', '8+i'],
      ['IMPRODUCT("3+4i","5-3i")', '27+11i'],
      ['IMDIV("-238+240i","10+24i")', '5+12i'],
      ['IMPRODUCT("i","i")', '-1'],
      ['IMPRODUCT(1,2,3)', '6'],
      ['IMSUM("1+i","1+i","1+i")', '3+3i'],
    ]);
  });

  it('reproduces the documented transcendental results', () => {
    expectAll([
      ['IMSQRT("1+i")', '1.09868411346781+0.455089860562227i'],
      ['IMEXP("1+i")', '1.46869393991589+2.28735528717884i'],
      ['IMLN("3+4i")', '1.6094379124341+0.927295218001612i'],
      ['IMLOG2("3+4i")', '2.32192809488736+1.33780421245098i'],
      ['IMLOG10("3+4i")', '0.698970004336019+0.402719196273373i'],
      ['IMSIN("4+3i")', '-7.61923172032141-6.548120040911i'],
      ['IMCOS("4+3i")', '-6.58066304055116+7.58155274274654i'],
      ['IMTAN("4+3i")', '0.00490825806749606+1.00070953606723i'],
      ['IMSINH("4+3i")', '-27.0168132580039+3.85373803791938i'],
      ['IMCOSH("4+3i")', '-27.0349456030742+3.85115333481178i'],
      ['IMSEC("4+3i")', '-0.0652940278579471-0.0752249603027732i'],
      ['IMCSC("4+3i")', '-0.0754898329158637+0.0648774713706355i'],
      ['IMCOT("4+3i")', '0.00490118239430447-0.999266927805902i'],
    ]);
  });

  it('raises a complex number to a power, including the degenerate ones', () => {
    expectAll([
      ['IMPOWER("2+3i",3)', '-46+9.00000000000001i'],
      ['IMPOWER("0",0)', '1'],
      ['IMPOWER("0",2)', '0'],
      ['IMPOWER("0",-1)', CellError.NUM],
      ['IMPOWER("2",10)', '1024'],
      ['IMPOWER("2+3i","x")', CellError.VALUE],
    ]);
    // The square root of -1 is i, up to the rounding of cos(pi/2).
    near('IMAGINARY(IMSQRT("-1"))', 1);
    expect(Math.abs(num('IMREAL(IMSQRT("-1"))'))).toBeLessThan(1e-15);
    // IMPOWER at a half power is the square root.
    near('IMREAL(IMPOWER("3+4i",0.5))', 2);
    near('IMAGINARY(IMPOWER("3+4i",0.5))', 1);
  });

  it('satisfies the identities that tie the family together', () => {
    // exp(ln z) = z, sin^2 + cos^2 = 1, and sec is the reciprocal of cos.
    near('IMREAL(IMEXP(IMLN("3+4i")))', 3);
    near('IMAGINARY(IMEXP(IMLN("3+4i")))', 4);
    near(
      'IMREAL(IMSUM(IMPRODUCT(IMSIN("2+1i"),IMSIN("2+1i")),IMPRODUCT(IMCOS("2+1i"),IMCOS("2+1i"))))',
      1,
    );
    near('IMREAL(IMPRODUCT("4+3i",IMDIV(1,"4+3i")))', 1);
    near('IMREAL(IMPRODUCT(IMSEC("2+1i"),IMCOS("2+1i")))', 1);
    near('IMREAL(IMPRODUCT(IMCSC("2+1i"),IMSIN("2+1i")))', 1);
    near('IMREAL(IMPRODUCT(IMCOT("2+1i"),IMTAN("2+1i")))', 1);
    // cosh^2 - sinh^2 = 1.
    near(
      'IMREAL(IMSUB(IMPRODUCT(IMCOSH("1+2i"),IMCOSH("1+2i")),IMPRODUCT(IMSINH("1+2i"),IMSINH("1+2i"))))',
      1,
    );
  });

  it('reads its operands from ranges as well as from arguments', () => {
    // A1:A3 of the employee table holds n, 1 and 2, so the numeric rows sum to
    // 3 once the header text fails to parse.
    expect(run('IMSUM(A2:A3)')).toBe('3');
    expect(run('IMSUM(A1:A3)')).toEqual(CellError.NUM);
  });
});

// ---------------------------------------------------------------------------
// The error function
// ---------------------------------------------------------------------------

describe('ERF and ERFC', () => {
  it('matches the values of the integral', () => {
    expect(run('ERF(0)')).toBe(0);
    expect(run('ERFC(0)')).toBe(1);
    near('ERF(0.5)', 0.5204998778130465);
    near('ERF(1)', 0.8427007929497149);
    near('ERF(2)', 0.9953222650189527);
    near('ERF(3)', 0.9999779095030014);
    near('ERFC(1)', 0.15729920705028513);
    near('ERFC(2)', 0.004677734981047266);
    near('ERFC(3)', 2.2090496998585441e-5);
    near('ERFC(5)', 1.5374597944280348e-12);
  });

  it('integrates between two limits, and knows its own symmetry', () => {
    // The two-limit form is erf(upper) - erf(lower).
    near('ERF(0.5,1)', 0.8427007929497149 - 0.5204998778130465);
    expect(run('ERF(1,1)')).toBe(0);
    near('ERF(-1)', -0.8427007929497149);
    near('ERFC(-1)', 1.8427007929497148);
    near('ERF(-1,1)', 2 * 0.8427007929497149);
    // ERF.PRECISE is the one-limit form under a modern name.
    near('ERF.PRECISE(1)', 0.8427007929497149);
    near('ERFC.PRECISE(1)', 0.15729920705028513);
    near('ERF.PRECISE(-0.5)', -0.5204998778130465);
  });

  it('stays complementary and saturates rather than overflowing', () => {
    for (const x of [0.1, 0.75, 1.5, 2.5, 4, 6]) {
      near(`ERF(${x})+ERFC(${x})`, 1);
      near(`ERF(-${x})+ERFC(-${x})`, 1);
    }
    expect(run('ERF(30)')).toBe(1);
    expect(run('ERF(-30)')).toBe(-1);
    expect(run('ERFC(30)')).toBe(0);
    expect(run('ERFC(-30)')).toBe(2);
    // The far tail stays a positive denormal rather than collapsing to zero.
    expect(num('ERFC(26)')).toBeGreaterThan(0);
    expect(num('ERFC(26)')).toBeLessThan(1e-290);
  });

  it('rejects what is not a number', () => {
    expectAll([
      ['ERF("x")', CellError.VALUE],
      ['ERF(TRUE)', CellError.VALUE],
      ['ERFC(1/0)', CellError.DIV0],
    ]);
  });
});

// ---------------------------------------------------------------------------
// Bessel functions
// ---------------------------------------------------------------------------

describe('BESSEL', () => {
  it('matches the tabulated values', () => {
    near('BESSELJ(1,0)', 0.7651976865579666);
    near('BESSELJ(1,1)', 0.4400505857449335);
    near('BESSELJ(2.5,0)', -0.048383776468197996, 1e-13);
    near('BESSELJ(5,2)', 0.046565116277752215, 1e-13);
    near('BESSELI(1,0)', 1.2660658777520084);
    near('BESSELI(1,1)', 0.5651591039924851);
    near('BESSELI(2.5,0)', 3.2898391440501231);
    near('BESSELY(1,0)', 0.0882569642156769, 1e-13);
    near('BESSELY(1,1)', -0.7812128213002887);
    near('BESSELK(1,0)', 0.4210244382407083);
    near('BESSELK(1,1)', 0.6019072301972346);
  });

  it('satisfies the three-term recurrences', () => {
    for (const [x, n] of [
      [1, 1], [2.5, 2], [7.3, 3], [15, 5], [0.4, 2], [30, 4],
    ] as const) {
      const scale = (2 * n) / x;
      near(`BESSELJ(${x},${n - 1})+BESSELJ(${x},${n + 1})`, scale * num(`BESSELJ(${x},${n})`), 1e-12);
      near(`BESSELI(${x},${n - 1})-BESSELI(${x},${n + 1})`, scale * num(`BESSELI(${x},${n})`), 1e-12);
      near(`BESSELY(${x},${n - 1})+BESSELY(${x},${n + 1})`, scale * num(`BESSELY(${x},${n})`), 1e-12);
      near(`BESSELK(${x},${n + 1})-BESSELK(${x},${n - 1})`, scale * num(`BESSELK(${x},${n})`), 1e-12);
    }
  });

  it('satisfies the Wronskians that tie J to Y and I to K', () => {
    for (const [x, n] of [
      [0.4, 0], [1, 1], [2.5, 2], [7.3, 3], [15, 5], [30, 4],
    ] as const) {
      const jy =
        num(`BESSELJ(${x},${n + 1})`) * num(`BESSELY(${x},${n})`) -
        num(`BESSELJ(${x},${n})`) * num(`BESSELY(${x},${n + 1})`);
      expect(Math.abs(jy - 2 / (Math.PI * x)) / (2 / (Math.PI * x))).toBeLessThan(1e-12);
      const ik =
        num(`BESSELI(${x},${n + 1})`) * num(`BESSELK(${x},${n})`) +
        num(`BESSELI(${x},${n})`) * num(`BESSELK(${x},${n + 1})`);
      expect(Math.abs(ik - 1 / x) * x).toBeLessThan(1e-12);
    }
  });

  it('keeps its accuracy where the order dwarfs the argument', () => {
    // J and I both approach (x/2)^n / n! there, which is 2.6912e-20 for n = 10
    // and x = 0.1 - far below anything a cancelling quadrature could resolve.
    near('BESSELJ(0.1,10)', 2.6905328954342e-20, 1e-12);
    near('BESSELI(0.1,10)', 2.6917561429221e-20, 1e-12);
    expect(num('BESSELY(0.1,10)')).toBeLessThan(-1e18);
    expect(num('BESSELK(0.1,10)')).toBeGreaterThan(1e18);
  });

  it('reflects across the origin for J and I, and refuses it for Y and K', () => {
    // J and I are even in x for even orders and odd for odd ones.
    near('BESSELJ(-1,0)', num('BESSELJ(1,0)'));
    near('BESSELJ(-1,1)', -num('BESSELJ(1,1)'));
    near('BESSELI(-1,1)', -num('BESSELI(1,1)'));
    expect(run('BESSELJ(0,0)')).toBe(1);
    expect(run('BESSELI(0,0)')).toBe(1);
    expect(run('BESSELJ(0,3)')).toBe(0);
    expect(run('BESSELI(0,3)')).toBe(0);
    expectAll([
      ['BESSELY(0,0)', CellError.NUM],
      ['BESSELY(-1,0)', CellError.NUM],
      ['BESSELK(0,0)', CellError.NUM],
      ['BESSELK(-1,0)', CellError.NUM],
    ]);
  });

  it('truncates the order and refuses a negative one', () => {
    expect(num('BESSELJ(1,0.9)')).toBe(num('BESSELJ(1,0)'));
    expectAll([
      ['BESSELJ(1,-1)', CellError.NUM],
      ['BESSELI(1,-1)', CellError.NUM],
      ['BESSELY(1,-1)', CellError.NUM],
      ['BESSELK(1,-1)', CellError.NUM],
      ['BESSELJ("x",0)', CellError.VALUE],
      ['BESSELJ(1,TRUE)', CellError.VALUE],
      ['BESSELJ(1/0,0)', CellError.DIV0],
    ]);
  });

  it('stays inside the amplitude that bounds J and Y for a large argument', () => {
    // |J_n(x)| and |Y_n(x)| can never exceed sqrt(2/(pi x)) by more than the
    // first correction term. A quadrature that has run out of nodes does not
    // merely lose digits, it returns a number no Bessel function can take, and
    // that is what this catches.
    for (const [x, n] of [
      [3000, 0], [20000, 0], [100000, 0], [1e6, 1], [1e7, 0], [1e9, 0],
    ] as const) {
      const bound = Math.sqrt(2 / (Math.PI * x)) * 1.01;
      expect(Math.abs(num(`BESSELJ(${x},${n})`)), `J(${x},${n})`).toBeLessThan(bound);
      expect(Math.abs(num(`BESSELY(${x},${n})`)), `Y(${x},${n})`).toBeLessThan(bound);
    }
  });

  it('satisfies J^2 + Y^2 = 2/(pi x) far out, where only the phase is left', () => {
    // The identity holds to O(1/x^2) - exactly (4n^2-1)/(8x^2) - so it pins
    // both functions absolutely, without a table and without the recurrences.
    for (const [x, n] of [
      [3000, 0], [20000, 0], [100000, 0], [1e6, 1], [1e7, 0], [1e9, 0], [1e12, 0],
    ] as const) {
      const j = num(`BESSELJ(${x},${n})`);
      const y = num(`BESSELY(${x},${n})`);
      const expected = 2 / (Math.PI * x);
      const correction = Math.abs(4 * n * n - 1) / (8 * x * x) + 1e-12;
      expect(Math.abs((j * j + y * y) / expected - 1), `x=${x} n=${n}`).toBeLessThan(correction);
    }
  });

  it('holds the recurrence and the Wronskian out where the argument is large', () => {
    for (const [x, n] of [[3000, 1], [100000, 2], [1e6, 1]] as const) {
      const scale = (2 * n) / x;
      const amplitude = Math.sqrt(2 / (Math.PI * x));
      const recurrence =
        num(`BESSELJ(${x},${n - 1})`) + num(`BESSELJ(${x},${n + 1})`) -
        scale * num(`BESSELJ(${x},${n})`);
      expect(Math.abs(recurrence) / amplitude, `J recurrence at ${x}`).toBeLessThan(1e-10);
      const wronskian =
        num(`BESSELJ(${x},${n + 1})`) * num(`BESSELY(${x},${n})`) -
        num(`BESSELJ(${x},${n})`) * num(`BESSELY(${x},${n + 1})`);
      expect(Math.abs(wronskian / (2 / (Math.PI * x)) - 1), `Wronskian at ${x}`).toBeLessThan(1e-12);
    }
  });

  it('costs a bounded amount of work however large the argument is', () => {
    // The quadratures need work proportional to x; without a ceiling on it
    // BESSELY(1E9,0) alone runs for the better part of an hour.
    const started = Date.now();
    for (const formula of ['BESSELY(1E9,0)', 'BESSELJ(1E9,0)', 'BESSELY(1E12,3)', 'BESSELK(1E9,0)']) {
      expect(typeof run(formula), formula).toBe('number');
    }
    expect(Date.now() - started).toBeLessThan(2000);
  });

  it('reflects a large negative argument rather than integrating one', () => {
    near('BESSELJ(-100000,0)', num('BESSELJ(100000,0)'));
    near('BESSELJ(-100000,1)', -num('BESSELJ(100000,1)'));
  });

  it('reports an overflow as #NUM! rather than as infinity', () => {
    // I0 grows like e^x/sqrt(2 pi x), so it is finite at 700 and not at 800.
    expect(num('BESSELI(700,0)')).toBeGreaterThan(1e302);
    expect(run('BESSELI(800,0)')).toEqual(CellError.NUM);
  });
});

// ---------------------------------------------------------------------------
// CONVERT
// ---------------------------------------------------------------------------

/**
 * Excel compares and displays at fifteen significant digits, so a conversion
 * that lands one unit in the last place away from a round number is the same
 * number to a user. This is the comparison value.ts already applies to `=`.
 */
function fifteen(formula: string, expected: number): void {
  expect(toExcelPrecision(num(formula)), formula).toBe(toExcelPrecision(expected));
}

describe('CONVERT', () => {
  it('reproduces the documented examples', () => {
    fifteen('CONVERT(1,"lbm","kg")', 0.45359237);
    fifteen('CONVERT(68,"F","C")', 20);
    fifteen('CONVERT(CONVERT(100,"ft","m"),"ft","m")', 9.290304);
    expect(run('CONVERT(2.5,"ft","sec")')).toEqual(CellError.NA);
  });

  it('converts within each unit group', () => {
    fifteen('CONVERT(1,"km","m")', 1000);
    fifteen('CONVERT(1,"m","mi")', 1 / 1609.344);
    fifteen('CONVERT(1,"Nmi","m")', 1852);
    fifteen('CONVERT(1,"Pica","in")', 1 / 72);
    fifteen('CONVERT(1,"pica","in")', 1 / 6);
    fifteen('CONVERT(1,"stone","lbm")', 14);
    fifteen('CONVERT(1,"uk_ton","lbm")', 2240);
    fifteen('CONVERT(1,"grain","g")', 0.06479891);
    fifteen('CONVERT(1,"hr","sec")', 3600);
    fifteen('CONVERT(1,"day","hr")', 24);
    fifteen('CONVERT(1,"yr","day")', 365.25);
    fifteen('CONVERT(1,"atm","Pa")', 101325);
    fifteen('CONVERT(1,"N","dyn")', 100000);
    fifteen('CONVERT(1,"lbf","N")', 4.4482216152605);
    fifteen('CONVERT(1,"J","e")', 1e7);
    fifteen('CONVERT(1,"cal","J")', 4.1868);
    fifteen('CONVERT(1,"c","J")', 4.184);
    fifteen('CONVERT(1,"Wh","J")', 3600);
    fifteen('CONVERT(1,"HP","W")', 745.69987158227);
    fifteen('CONVERT(1,"PS","W")', 735.49875);
    fifteen('CONVERT(1,"T","ga")', 10000);
    fifteen('CONVERT(1,"gal","l")', 3.785411784);
    fifteen('CONVERT(1,"m3","l")', 1000);
    fifteen('CONVERT(1,"GRT","ft3")', 100);
    fifteen('CONVERT(1,"ha","m2")', 10000);
    fifteen('CONVERT(1,"uk_acre","m2")', 4046.8564224);
    fifteen('CONVERT(1,"byte","bit")', 8);
    fifteen('CONVERT(1,"mph","m/s")', 0.44704);
    fifteen('CONVERT(1,"kn","m/s")', 1852 / 3600);
  });

  it('handles temperature, which needs an offset and not just a ratio', () => {
    fifteen('CONVERT(100,"C","F")', 212);
    fifteen('CONVERT(-40,"C","F")', -40);
    fifteen('CONVERT(32,"F","C")', 0);
    fifteen('CONVERT(0,"C","K")', 273.15);
    fifteen('CONVERT(273.15,"K","C")', 0);
    fifteen('CONVERT(0,"C","Rank")', 491.67);
    fifteen('CONVERT(100,"C","Reau")', 80);
    fifteen('CONVERT(100,"cel","fah")', 212);
    // The scale is what a prefix multiplies, so a millidegree is a thousandth
    // of one.
    fifteen('CONVERT(1000,"mC","C")', 1);
  });

  it('applies metric prefixes, raised to the unit-s dimension', () => {
    fifteen('CONVERT(1,"g","kg")', 0.001);
    fifteen('CONVERT(1,"ug","g")', 1e-6);
    fifteen('CONVERT(1,"m2","cm2")', 10000);
    fifteen('CONVERT(1,"km2","m2")', 1e6);
    fifteen('CONVERT(1,"m3","cm3")', 1e6);
    fifteen('CONVERT(1,"ml","l")', 0.001);
    fifteen('CONVERT(1,"kbyte","byte")', 1000);
    fifteen('CONVERT(1,"kibyte","byte")', 1024);
    fifteen('CONVERT(1,"Mibyte","byte")', 1048576);
    // The exact spelling wins over any prefix reading of it: kn is a knot, not
    // a kilonewton, and h is horsepower, not hecto-anything.
    fifteen('CONVERT(1,"kn","m/h")', 1852);
    fifteen('CONVERT(1,"h","W")', 745.69987158227);
    // A binary prefix belongs to the information units alone.
    expect(run('CONVERT(1,"kim","m")')).toEqual(CellError.NA);
  });

  it('reports an unknown or incompatible unit as #N/A', () => {
    expectAll([
      ['CONVERT(1,"m","g")', CellError.NA],
      ['CONVERT(1,"xyz","m")', CellError.NA],
      ['CONVERT(1,"m","xyz")', CellError.NA],
      ['CONVERT(1,"C","m")', CellError.NA],
      // Unit names are case-sensitive.
      ['CONVERT(1,"KM","m")', CellError.NA],
      ['CONVERT(1,"g","G")', CellError.NA],
      // A prefix on its own is not a unit.
      ['CONVERT(1,"da","m")', CellError.NA],
    ]);
  });

  it('rejects a non-numeric value, and propagates an error', () => {
    expectAll([
      ['CONVERT("x","m","ft")', CellError.VALUE],
      ['CONVERT(TRUE,"m","ft")', CellError.VALUE],
      ['CONVERT(1/0,"m","ft")', CellError.DIV0],
    ]);
    expect(run('CONVERT(0,"m","ft")')).toBe(0);
    fifteen('CONVERT(-3,"m","ft")', -3 / 0.3048);
  });
});

// ---------------------------------------------------------------------------
// DELTA and GESTEP
// ---------------------------------------------------------------------------

describe('DELTA and GESTEP', () => {
  it('compares against a second argument that defaults to zero', () => {
    expectAll([
      ['DELTA(5,4)', 0],
      ['DELTA(5,5)', 1],
      ['DELTA(0.5,0)', 0],
      ['DELTA(0)', 1],
      ['DELTA(-1,-1)', 1],
      ['GESTEP(5,4)', 1],
      ['GESTEP(5,5)', 1],
      ['GESTEP(-4,-5)', 1],
      ['GESTEP(-1)', 0],
      ['GESTEP(0)', 1],
      ['GESTEP(0.00001)', 1],
    ]);
  });

  it('rejects logical values and propagates errors', () => {
    expectAll([
      ['DELTA(TRUE,1)', CellError.VALUE],
      ['GESTEP(TRUE)', CellError.VALUE],
      ['DELTA("x",1)', CellError.VALUE],
      ['DELTA(1/0,1)', CellError.DIV0],
      // Numeric text is still a number.
      ['DELTA("5",5)', 1],
    ]);
  });
});

// ---------------------------------------------------------------------------
// The D functions
// ---------------------------------------------------------------------------

describe('the D functions', () => {
  const db = 'A1:E9';

  it('names the field by header, by index, and without regard to case', () => {
    expectAll([
      [`DSUM(${db},"salary",G1:G2)`, 518000],
      [`DSUM(${db},4,G1:G2)`, 518000],
      [`DSUM(${db},"SALARY",G1:G2)`, 518000],
      [`DSUM(${db},4.9,G1:G2)`, 518000],
      // A field that is not there, or an index outside the database.
      [`DSUM(${db},"nosuchfield",G1:G2)`, CellError.VALUE],
      [`DSUM(${db},9,G1:G2)`, CellError.VALUE],
      [`DSUM(${db},0,G1:G2)`, CellError.VALUE],
    ]);
  });

  it('ORs the criteria rows and ANDs the cells within one row', () => {
    // H1:I3 is (dept = Eng AND salary > 170000) OR (dept = Ops).
    expectAll([
      [`DSUM(${db},"salary",H1:I2)`, 353000],
      [`DSUM(${db},"salary",H1:I3)`, 592000],
      [`DCOUNT(${db},"salary",H1:I3)`, 4],
      [`DSUM(${db},"salary",P1:P2)`, 239000],
    ]);
  });

  it('treats a blank criterion, and a header-only range, as no constraint', () => {
    expectAll([
      [`DSUM(${db},"salary",L1:L2)`, 1207000],
      [`DSUM(${db},"salary",L1:L1)`, 1207000],
      [`DCOUNT(${db},"salary",L1:L2)`, 8],
      [`DCOUNTA(${db},"name",L1:L2)`, 8],
    ]);
  });

  it('matches bare text by prefix, and honours wildcards and comparisons', () => {
    // N1:N2 is name = "A", which selects Ada and Annie; O1:O2 is name = "?ary",
    // which selects Mary alone.
    expectAll([
      [`DSUM(${db},"salary",N1:N2)`, 286000],
      [`DSUM(${db},"salary",O1:O2)`, 143000],
      [`DCOUNT(${db},"salary",N1:N2)`, 2],
    ]);
  });

  it('counts records when no field is named', () => {
    expectAll([
      [`DCOUNT(${db},,G1:G2)`, 3],
      [`DCOUNTA(${db},,G1:G2)`, 3],
      [`DCOUNT(${db},,L1:L2)`, 8],
      // A range in the field position can only be the criteria, so the
      // two-argument spelling means the same thing.
      [`DCOUNT(${db},G1:G2)`, 3],
      [`DCOUNTA(${db},G1:G2)`, 3],
    ]);
  });

  it('aggregates the numeric cells of the field and ignores the rest', () => {
    expectAll([
      [`DMAX(${db},"salary",G1:G2)`, 181000],
      [`DMIN(${db},"salary",G1:G2)`, 165000],
      [`DPRODUCT(${db},"n",G1:G2)`, 10],
      [`DGET(${db},"salary",J1:J2)`, 121000],
      [`DGET(${db},"name",J1:J2)`, 'Annie'],
      // A text field has no numbers in it to sum.
      [`DSUM(${db},"name",G1:G2)`, 0],
      [`DCOUNT(${db},"name",G1:G2)`, 0],
      [`DCOUNTA(${db},"name",G1:G2)`, 3],
    ]);
    fifteen(`DAVERAGE(${db},"salary",G1:G2)`, 518000 / 3);
    fifteen(`DSTDEV(${db},"salary",P1:P2)`, 2121.32034355964);
    fifteen(`DSTDEVP(${db},"salary",P1:P2)`, 1500);
    fifteen(`DVAR(${db},"salary",P1:P2)`, 4500000);
    fifteen(`DVARP(${db},"salary",P1:P2)`, 2250000);
  });

  it('separates the two ways DGET can fail', () => {
    // Three Eng records is too many; no Zeno is too few.
    expectAll([
      [`DGET(${db},"salary",G1:G2)`, CellError.NUM],
      [`DGET(${db},"salary",K1:K2)`, CellError.VALUE],
    ]);
  });

  it('answers an empty selection the way the matching aggregate does', () => {
    expectAll([
      [`DSUM(${db},"salary",K1:K2)`, 0],
      [`DMAX(${db},"salary",K1:K2)`, 0],
      [`DMIN(${db},"salary",K1:K2)`, 0],
      [`DCOUNT(${db},"salary",K1:K2)`, 0],
      [`DPRODUCT(${db},"salary",K1:K2)`, 0],
      [`DAVERAGE(${db},"salary",K1:K2)`, CellError.DIV0],
      [`DVARP(${db},"salary",K1:K2)`, CellError.DIV0],
      // The sample statistics need two records, not one.
      [`DSTDEV(${db},"salary",J1:J2)`, CellError.DIV0],
      [`DVAR(${db},"salary",J1:J2)`, CellError.DIV0],
      [`DSTDEVP(${db},"salary",J1:J2)`, 0],
    ]);
  });

  it('reports a criteria header that names no field', () => {
    expect(run(`DSUM(${db},"salary",M1:M2)`)).toEqual(CellError.VALUE);
  });

  it('propagates an error out of the field it is aggregating', () => {
    expectAll([
      ['DSUM(A26:B28,"v",D26:D26)', CellError.DIV0],
      ['DMAX(A26:B28,"v",D26:D26)', CellError.DIV0],
      ['DAVERAGE(A26:B28,"v",D26:D26)', CellError.DIV0],
      // COUNT and COUNTA see through an error in a range, and so do their
      // database counterparts.
      ['DCOUNT(A26:B28,"v",D26:D26)', 1],
      ['DCOUNTA(A26:B28,"v",D26:D26)', 2],
      // A field with no error in it is unaffected by one elsewhere.
      ['DSUM(A26:B28,"k",D26:D26)', 3],
    ]);
  });

  it('separates blanks from zeros and text from numbers', () => {
    // F26:G29 holds the text "text", a blank, and the number 7 in column v.
    expectAll([
      ['DSUM(F26:G29,"v",I26:I26)', 7],
      ['DCOUNT(F26:G29,"v",I26:I26)', 1],
      ['DCOUNTA(F26:G29,"v",I26:I26)', 2],
      ['DCOUNT(F26:G29,,I26:I26)', 3],
      ['DAVERAGE(F26:G29,"v",I26:I26)', 7],
    ]);
  });

  it('propagates an error argument, and refuses a database that is not a range', () => {
    expectAll([
      [`DSUM(${db},"salary",1/0)`, CellError.DIV0],
      [`DSUM(1/0,"salary",G1:G2)`, CellError.DIV0],
      [`DSUM(${db},1/0,G1:G2)`, CellError.DIV0],
      [`DGET(${db},,G1:G2)`, CellError.VALUE],
    ]);
  });

  it('takes the field name from a cell', () => {
    // I1 holds the text "salary", which is how a field is usually parameterised.
    expect(run(`DSUM(${db},I1,G1:G2)`)).toBe(518000);
  });
});

// ---------------------------------------------------------------------------
// Metadata
// ---------------------------------------------------------------------------

describe('metadata', () => {
  it('declares nothing volatile or structural', () => {
    for (const spec of ENGINEERING_FUNCTIONS) {
      expect(spec.volatile, spec.name).toBeUndefined();
      expect(spec.structural, spec.name).toBeUndefined();
      expect(spec.summary, spec.name).toBeTruthy();
    }
  });

  it('flags the post-2007 names that xlsx stores prefixed', () => {
    for (const name of [
      'BITAND', 'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT',
      'IMTAN', 'IMCOT', 'IMSINH', 'IMCOSH', 'IMSEC', 'IMCSC',
      'ERF.PRECISE', 'ERFC.PRECISE',
    ]) {
      expect(registry.get(name)?.futureFunction, name).toBe(true);
    }
    // The Analysis ToolPak names, which have been built in since 2007 and are
    // stored unprefixed.
    for (const name of ['DEC2BIN', 'COMPLEX', 'CONVERT', 'ERF', 'BESSELJ', 'DSUM']) {
      expect(registry.get(name)?.futureFunction, name).toBe(false);
    }
    // The flag on the spec is not enough on its own: the xlsx writer prefixes
    // from the name list, so a 2013 function missing from it round-trips into a
    // file Excel shows as #NAME?.
    for (const name of ['IMCOT', 'IMSEC', 'BITAND', 'ERF.PRECISE']) {
      expect(storageName(name), name).toBe(`_xlfn.${name}`);
    }
    expect(storageName('BESSELJ')).toBe('BESSELJ');
  });

  it('broadcasts the scalar functions and not the aggregating ones', () => {
    for (const name of ['DEC2BIN', 'BITAND', 'IMABS', 'IMSUB', 'CONVERT', 'ERF', 'BESSELJ']) {
      expect(registry.get(name)?.broadcast, name).toBe(true);
    }
    for (const name of ['IMSUM', 'IMPRODUCT', 'DSUM', 'DGET', 'DCOUNT']) {
      expect(registry.get(name)?.broadcast, name).toBeUndefined();
    }
  });
});
