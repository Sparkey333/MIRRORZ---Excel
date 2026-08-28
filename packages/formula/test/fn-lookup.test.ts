/**
 * Lookup and reference.
 *
 * The first block replays the LibreOffice-recalculated fixtures: every case in
 * formulas.calc.xlsx whose formula this module can evaluate on its own is
 * re-parsed, re-evaluated against the same workbook, and compared with the value
 * the oracle actually computed.
 *
 * XLOOKUP and XMATCH are the exception. LibreOffice 24.2 does not implement
 * them, so the fixture holds #NAME? for both cases and there is nothing to
 * assert against; their expected values are hand-specified from Microsoft's
 * documentation and noted as such at the point of use. See
 * docs/oracle-divergences.md.
 *
 * Everything else runs against a small synthetic workbook, because the fixture
 * has no sorted numeric column, no descending column, no duplicates and no
 * blanks - which is exactly where a lookup goes wrong. Formulas are evaluated at
 * K21, well away from the data, so nothing accidentally depends on implicit
 * intersection.
 *
 * The registry under test holds LOOKUP_FUNCTIONS alone, so a failure here is
 * never another category's fault.
 */

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, type Scalar, Workbook, isError } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { Evaluator } from '../src/evaluator.js';
import { LOOKUP_FUNCTIONS } from '../src/functions/lookup.js';
import { parseFormula } from '../src/parser.js';
import { FunctionRegistry } from '../src/registry.js';
import { WorkbookStore } from '../src/store.js';
import { type ArrayValue, type Value, isArray, isRef, makeRef } from '../src/value.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

const registry = new FunctionRegistry().registerAll(LOOKUP_FUNCTIONS);

// ---------------------------------------------------------------------------
// The oracle
// ---------------------------------------------------------------------------

const { workbook: oracleBook } = readXlsx(
  new Uint8Array(readFileSync(new URL('formulas.calc.xlsx', FIXTURES))),
);
const oracleSheet = oracleBook.getSheet('Formulas')!;
const oracleEval = new Evaluator(new WorkbookStore(oracleBook), registry, {
  dateSystem: oracleBook.dateSystem,
});

const ORACLE = new Map<string, { formula: string; value: Scalar; row: number }>();
for (const { row, col, cell } of oracleSheet.entries()) {
  if (col !== 2 || !cell.formula) continue;
  const name = oracleSheet.getValue(row, 0);
  if (typeof name === 'string') ORACLE.set(name, { formula: cell.formula, value: cell.value, row });
}

/**
 * LibreOffice writes the logical constants as calls, `FALSE()` rather than
 * `FALSE`. Those names belong to the logical module, and this registry holds
 * only ours, so they are folded back to literals before parsing.
 */
function normalise(formula: string): string {
  return formula.replace(/\bTRUE\(\)/gi, 'TRUE').replace(/\bFALSE\(\)/gi, 'FALSE');
}

/** Evaluate a fixture formula in the cell the oracle evaluated it in. */
function oracle(name: string): { actual: Scalar; expected: Scalar } {
  const c = ORACLE.get(name);
  expect(c, `case ${name} missing from the fixture`).toBeDefined();
  const actual = oracleEval.evaluateScalar({
    ast: parseFormula(normalise(c!.formula), { origin: { row: c!.row, col: 2 } }),
    sheet: 'Formulas',
    row: c!.row,
    col: 2,
  });
  return { actual, expected: c!.value };
}

/** Evaluate a formula of our own against the fixture workbook. */
function onFixture(formula: string): Scalar {
  return oracleEval.evaluateScalar({
    ast: parseFormula(formula, { origin: { row: 0, col: 2 } }),
    sheet: 'Formulas',
    row: 0,
    col: 2,
  });
}

describe('oracle: formulas.calc.xlsx', () => {
  for (const name of [
    'VLOOKUP',
    'HLOOKUP',
    'INDEX_MATCH',
    'CHOOSE',
    'OFFSET',
    'INDIRECT',
    'ROW',
    'COLUMN',
    'ROWS',
    'COLUMNS',
  ]) {
    it(`reproduces ${name}`, () => {
      const { actual, expected } = oracle(name);
      expect(actual).toEqual(expected);
    });
  }

  it('leaves XLOOKUP and XMATCH unverified by the oracle', () => {
    // Documenting the gap rather than working around it: if LibreOffice ever
    // learns these, this test fails and the two cases below become assertable.
    expect(ORACLE.get('XLOOKUP')!.value).toEqual(CellError.NAME);
    expect(ORACLE.get('XMATCH')!.value).toEqual(CellError.NAME);
  });

  it('computes the XLOOKUP and XMATCH cases from the Microsoft definition', () => {
    // Annie is row 7 of Data, salary 121000; Mary is the 8th name in B2:B9.
    expect(onFixture('XLOOKUP("Annie",Data!B2:B9,Data!D2:D9,"none")')).toBe(121000);
    expect(onFixture('XLOOKUP("Nobody",Data!B2:B9,Data!D2:D9,"none")')).toBe('none');
    expect(onFixture('XLOOKUP("Nobody",Data!B2:B9,Data!D2:D9)')).toEqual(CellError.NA);
    expect(onFixture('XMATCH("Mary",Data!B2:B9,0)')).toBe(8);
  });
});

// ---------------------------------------------------------------------------
// A synthetic workbook for the cases the fixture does not cover
// ---------------------------------------------------------------------------

/**
 * Sheet S:
 *
 *        A     B    C         D     E    F         G    H    I
 *   1     1     a    apple     30    10   #DIV/0!   1    2    3
 *   2     2     b    banana    20    10             x    y    z
 *   3     4     c    cherry    10    20
 *   4     8     d               30
 *   5    16     e               30
 *
 * A is ascending, D is descending, E has runs of duplicates, C is sorted text,
 * G1:I2 is a horizontal table for HLOOKUP, and F1 holds an error.
 */
function makeBook(): Workbook {
  const wb = new Workbook();
  const s = wb.addSheet('S');
  const numbers = [1, 2, 4, 8, 16];
  const letters = ['a', 'b', 'c', 'd', 'e'];
  for (let i = 0; i < 5; i++) {
    s.setValue(i, 0, numbers[i]!);
    s.setValue(i, 1, letters[i]!);
  }
  s.setValue(0, 2, 'apple');
  s.setValue(1, 2, 'banana');
  s.setValue(2, 2, 'cherry');
  s.setValue(0, 3, 30);
  s.setValue(1, 3, 20);
  s.setValue(2, 3, 10);
  // Runs of duplicates, for the search-direction tests.
  s.setValue(0, 4, 10);
  s.setValue(1, 4, 10);
  s.setValue(2, 4, 20);
  s.setValue(3, 4, 30);
  s.setValue(4, 4, 30);
  s.setValue(0, 5, CellError.DIV0);
  // A horizontal table for HLOOKUP: G1:I1 keys over G2:I2 values.
  s.setValue(0, 6, 1);
  s.setValue(0, 7, 2);
  s.setValue(0, 8, 3);
  s.setValue(1, 6, 'x');
  s.setValue(1, 7, 'y');
  s.setValue(1, 8, 'z');
  wb.addSheet('Other');
  return wb;
}

const book = makeBook();
const ev = new Evaluator(new WorkbookStore(book), registry, { dateSystem: book.dateSystem });

/** Evaluate at K21 unless told otherwise, and collapse to what a cell shows. */
function run(formula: string, row = 20, col = 10): Scalar {
  return ev.evaluateScalar({
    ast: parseFormula(formula, { origin: { row, col } }),
    sheet: 'S',
    row,
    col,
  });
}

/** The same, without collapsing, for the functions that return shapes. */
function shape(formula: string, row = 20, col = 10): Value {
  return ev.evaluate({
    ast: parseFormula(formula, { origin: { row, col } }),
    sheet: 'S',
    row,
    col,
  });
}

function cellsOf(v: Value): Scalar[] {
  expect(isArray(v), `expected an array, got ${JSON.stringify(v)}`).toBe(true);
  return (v as ArrayValue).data;
}

// ---------------------------------------------------------------------------

describe('VLOOKUP', () => {
  it('defaults to approximate match, taking the largest entry at or below', () => {
    expect(run('VLOOKUP(5,A1:B5,2)')).toBe('c');
    expect(run('VLOOKUP(4,A1:B5,2)')).toBe('c');
    expect(run('VLOOKUP(100,A1:B5,2)')).toBe('e');
  });

  it('is #N/A when the target is below every entry', () => {
    expect(run('VLOOKUP(0,A1:B5,2)')).toEqual(CellError.NA);
  });

  it('does exact match only when told to', () => {
    expect(run('VLOOKUP(2,A1:B5,2,FALSE)')).toBe('b');
    expect(run('VLOOKUP(5,A1:B5,2,FALSE)')).toEqual(CellError.NA);
    // An explicitly blank fourth argument is FALSE, not the omitted default.
    expect(run('VLOOKUP(5,A1:B5,2,)')).toEqual(CellError.NA);
  });

  it('never coerces text to number or back', () => {
    expect(run('VLOOKUP("2",A1:B5,2,FALSE)')).toEqual(CellError.NA);
    expect(run('VLOOKUP(2,C1:C3,1,FALSE)')).toEqual(CellError.NA);
  });

  it('matches wildcards in exact mode', () => {
    expect(run('VLOOKUP("ban*",C1:C3,1,FALSE)')).toBe('banana');
    expect(run('VLOOKUP("?anana",C1:C3,1,FALSE)')).toBe('banana');
    expect(run('VLOOKUP("*rr*",C1:C3,1,FALSE)')).toBe('cherry');
  });

  it('reports a bad column index the way Excel does', () => {
    expect(run('VLOOKUP(2,A1:B5,3,FALSE)')).toEqual(CellError.REF);
    expect(run('VLOOKUP(2,A1:B5,0,FALSE)')).toEqual(CellError.VALUE);
    expect(run('VLOOKUP(2,A1:B5,-1,FALSE)')).toEqual(CellError.VALUE);
  });

  it('treats a blank lookup value as zero', () => {
    expect(run('VLOOKUP(Z50,A1:B5,2,FALSE)')).toEqual(CellError.NA);
    expect(run('VLOOKUP(Z50,A1:B5,2)')).toEqual(CellError.NA);
  });

  it('skips error cells rather than propagating them', () => {
    expect(run('VLOOKUP("apple",C1:C3,1,FALSE)')).toBe('apple');
    expect(run('VLOOKUP(1,F1:F1,1,FALSE)')).toEqual(CellError.NA);
  });

  it('works over an array constant', () => {
    expect(run('VLOOKUP(2,{1,"one";2,"two"},2,FALSE)')).toBe('two');
  });
});

describe('HLOOKUP', () => {
  it('searches the top row and returns from the numbered row', () => {
    expect(run('HLOOKUP(2,G1:I2,2,FALSE)')).toBe('y');
    expect(run('HLOOKUP(2,G1:I2,2)')).toBe('y');
  });

  it('takes the largest key at or below in approximate mode', () => {
    expect(run('HLOOKUP(2.5,G1:I2,2)')).toBe('y');
    expect(run('HLOOKUP(0,G1:I2,2)')).toEqual(CellError.NA);
  });

  it('reports a bad row index', () => {
    expect(run('HLOOKUP(2,G1:I2,3,FALSE)')).toEqual(CellError.REF);
    expect(run('HLOOKUP(2,G1:I2,0,FALSE)')).toEqual(CellError.VALUE);
  });
});

describe('LOOKUP', () => {
  it('uses the vector form when a result vector is given', () => {
    expect(run('LOOKUP(3,A1:A5,B1:B5)')).toBe('b');
    expect(run('LOOKUP(16,A1:A5,B1:B5)')).toBe('e');
    expect(run('LOOKUP(0,A1:A5,B1:B5)')).toEqual(CellError.NA);
  });

  it('reads the last column of the array in the array form', () => {
    expect(run('LOOKUP(3,A1:B5)')).toBe('b');
    // Wider than tall, so the first row is searched and the last row returned.
    expect(run('LOOKUP(2,G1:I2)')).toBe('y');
  });

  it('searches the first column of a square array, as Microsoft documents', () => {
    // "If array covers an area that is wider than it is tall (more columns than
    // rows), LOOKUP searches for lookup_value in the first row. If array is
    // square or is taller than it is wide, LOOKUP searches in the first
    // column." A1:B2 is square, so 2 is found in A2 and B2 comes back. Reading
    // a square array by rows instead would find the 1 in A1 and answer 2.
    expect(run('LOOKUP(2,A1:B2)')).toBe('b');
    expect(run('LOOKUP(1,A1:B2)')).toBe('a');
  });

  it('has no exact mode, so a miss still returns the nearest below', () => {
    expect(run('LOOKUP("bb",C1:C3)')).toBe('banana');
  });
});

describe('MATCH', () => {
  it('defaults to match type 1', () => {
    expect(run('MATCH(3,A1:A5)')).toBe(2);
    expect(run('MATCH(16,A1:A5)')).toBe(5);
    expect(run('MATCH(0,A1:A5)')).toEqual(CellError.NA);
  });

  it('does exact match for type 0, in any order', () => {
    expect(run('MATCH(10,D1:D3,0)')).toBe(3);
    expect(run('MATCH(15,D1:D3,0)')).toEqual(CellError.NA);
    expect(run('MATCH("cherry",C1:C3,0)')).toBe(3);
    // Text comparison is case-insensitive, as everywhere else in Excel.
    expect(run('MATCH("CHERRY",C1:C3,0)')).toBe(3);
  });

  it('takes the smallest entry at or above for type -1 on descending data', () => {
    expect(run('MATCH(15,D1:D3,-1)')).toBe(2);
    expect(run('MATCH(30,D1:D3,-1)')).toBe(1);
    expect(run('MATCH(31,D1:D3,-1)')).toEqual(CellError.NA);
  });

  it('resolves a run of duplicates to the last of the run', () => {
    expect(run('MATCH(10,E1:E5,1)')).toBe(2);
    expect(run('MATCH(30,E1:E5,1)')).toBe(5);
  });

  it('supports wildcards only in exact mode', () => {
    expect(run('MATCH("ch*",C1:C3,0)')).toBe(3);
    // In approximate mode `*` is an ordinary character, and "ch*" sorts below
    // "cherry", so the answer is the entry before it rather than a match.
    expect(run('MATCH("ch*",C1:C3,1)')).toBe(2);
  });

  it('needs a vector, not a block', () => {
    expect(run('MATCH(2,A1:B5,0)')).toEqual(CellError.NA);
  });

  it('propagates an error in the lookup value', () => {
    expect(run('MATCH(F1,A1:A5,0)')).toEqual(CellError.DIV0);
  });
});

describe('INDEX', () => {
  it('reads a single cell', () => {
    expect(run('INDEX(A1:B5,2,2)')).toBe('b');
    expect(run('INDEX(A1:A5,3)')).toBe(4);
  });

  it('returns a reference, so it can be a range operand', () => {
    expect(shape('INDEX(A1:A5,3)')).toEqual(makeRef('S', 2, 0, 2, 0));
    expect(shape('A1:INDEX(A1:A5,3)')).toEqual(makeRef('S', 0, 0, 2, 0));
    expect(shape('INDEX(A1:B5,0,1)')).toEqual(makeRef('S', 0, 0, 4, 0));
    expect(shape('INDEX(A1:B5,2,0)')).toEqual(makeRef('S', 1, 0, 1, 1));
  });

  it('treats the single argument of a one-row source as the column', () => {
    expect(run('INDEX(G1:I1,3)')).toBe(3);
  });

  it('separates an out-of-range index from a negative one', () => {
    expect(run('INDEX(A1:A5,6)')).toEqual(CellError.REF);
    expect(run('INDEX(A1:A5,-1)')).toEqual(CellError.VALUE);
    expect(run('INDEX(A1:B5,1,3)')).toEqual(CellError.REF);
  });

  it('works on an array constant, slicing where an index is zero', () => {
    expect(run('INDEX({1,2;3,4},2,1)')).toBe(3);
    expect(cellsOf(shape('INDEX({1,2;3,4},0,2)'))).toEqual([2, 4]);
    expect(cellsOf(shape('INDEX({1,2;3,4},1,0)'))).toEqual([1, 2]);
  });

  it('reports #REF! for an area past the first, which unions cannot express', () => {
    expect(run('INDEX(A1:A5,1,1,2)')).toEqual(CellError.REF);
    expect(run('INDEX(A1:A5,1,1,1)')).toBe(1);
  });

  it('is not volatile', () => {
    expect(LOOKUP_FUNCTIONS.find((f) => f.name === 'INDEX')!.volatile).toBeUndefined();
  });
});

describe('CHOOSE', () => {
  it('picks by one-based index and truncates', () => {
    expect(run('CHOOSE(2,"a","b","c")')).toBe('b');
    expect(run('CHOOSE(2.9,"a","b","c")')).toBe('b');
  });

  it('rejects an index outside the value list', () => {
    expect(run('CHOOSE(0,"a","b")')).toEqual(CellError.VALUE);
    expect(run('CHOOSE(3,"a","b")')).toEqual(CellError.VALUE);
  });

  it('evaluates only the chosen argument', () => {
    expect(run('CHOOSE(1,"a",1/0)')).toBe('a');
    expect(run('CHOOSE(2,"a",1/0)')).toEqual(CellError.DIV0);
  });

  it('passes a reference through unchanged', () => {
    expect(shape('CHOOSE(2,A1:A2,A1:A5)')).toEqual(makeRef('S', 0, 0, 4, 0));
    expect(shape('A1:CHOOSE(1,A3)')).toEqual(makeRef('S', 0, 0, 2, 0));
  });
});

describe('OFFSET', () => {
  it('shifts and resizes a reference', () => {
    expect(shape('OFFSET(A1,1,1)')).toEqual(makeRef('S', 1, 1, 1, 1));
    expect(shape('OFFSET(A1,0,0,3,2)')).toEqual(makeRef('S', 0, 0, 2, 1));
    expect(run('OFFSET(A1,2,1)')).toBe('c');
  });

  it('keeps the anchor size when height and width are omitted or blank', () => {
    expect(shape('OFFSET(A1:B2,1,0)')).toEqual(makeRef('S', 1, 0, 2, 1));
    expect(shape('OFFSET(A1:B2,1,0,,)')).toEqual(makeRef('S', 1, 0, 2, 1));
  });

  it('extends backwards for a negative height or width', () => {
    expect(shape('OFFSET(A3,0,0,-3,1)')).toEqual(makeRef('S', 0, 0, 2, 0));
    expect(shape('OFFSET(C1,0,0,1,-3)')).toEqual(makeRef('S', 0, 0, 0, 2));
  });

  it('is #REF! off the edge of the sheet and #REF! for an empty rectangle', () => {
    expect(run('OFFSET(A1,-1,0)')).toEqual(CellError.REF);
    expect(run('OFFSET(A1,0,-1)')).toEqual(CellError.REF);
    expect(run('OFFSET(A1,0,0,0,1)')).toEqual(CellError.REF);
    expect(run('OFFSET(A1,0,0,1,0)')).toEqual(CellError.REF);
  });

  it('needs a reference', () => {
    expect(run('OFFSET(5,1,1)')).toEqual(CellError.VALUE);
    expect(run('OFFSET({1,2},1,1)')).toEqual(CellError.VALUE);
  });

  it('is volatile', () => {
    expect(LOOKUP_FUNCTIONS.find((f) => f.name === 'OFFSET')!.volatile).toBe(true);
  });
});

describe('INDIRECT', () => {
  it('resolves A1 text, absolute or relative, qualified or not', () => {
    expect(run('INDIRECT("A2")')).toBe(2);
    expect(run('INDIRECT("$A$2")')).toBe(2);
    expect(run('INDIRECT("S!A2")')).toBe(2);
    expect(shape('INDIRECT("A1:A3")')).toEqual(makeRef('S', 0, 0, 2, 0));
  });

  it('resolves R1C1 text when the second argument is FALSE', () => {
    expect(run('INDIRECT("R2C1",FALSE)')).toBe(2);
    expect(run('INDIRECT("R3C1:R3C1",FALSE)')).toBe(4);
    // Relative brackets are offsets from the calling cell.
    expect(run('INDIRECT("R[1]C[-9]",FALSE)', 0, 9)).toBe(2);
    expect(run('INDIRECT("RC",FALSE)', 0, 0)).toBe(1);
  });

  it('is unaffected by the calling cell in A1 mode', () => {
    expect(run('INDIRECT("A2")', 0, 0)).toBe(2);
    expect(run('INDIRECT("A2")', 40, 4)).toBe(2);
  });

  it('is #REF! for anything that is not a reference', () => {
    expect(run('INDIRECT("")')).toEqual(CellError.REF);
    expect(run('INDIRECT("not a ref")')).toEqual(CellError.REF);
    expect(run('INDIRECT("A1+1")')).toEqual(CellError.REF);
    expect(run('INDIRECT("Nosuch!A1")')).toEqual(CellError.REF);
    expect(run('INDIRECT("A2",FALSE)')).toEqual(CellError.REF);
    expect(run('INDIRECT("R0C1",FALSE)')).toEqual(CellError.REF);
    expect(run('INDIRECT("A1048577")')).toEqual(CellError.REF);
  });

  it('reads a whole column, clipped exactly as the same text typed directly is', () => {
    // The engine clips a whole-column reference to the used range rather than
    // materialising 1,048,576 rows; INDIRECT has to agree, or A:A and
    // INDIRECT("A:A") are two different ranges spelled the same way.
    const ref = shape('INDIRECT("A:A")');
    expect(isRef(ref)).toBe(true);
    expect(ref).toEqual(shape('A:A'));
    expect(ref).toEqual(makeRef('S', 0, 0, 4, 0));
    expect(run('ROWS(INDIRECT("A:A"))')).toBe(run('ROWS(A:A)'));
  });

  it('is volatile', () => {
    expect(LOOKUP_FUNCTIONS.find((f) => f.name === 'INDIRECT')!.volatile).toBe(true);
  });
});

describe('ADDRESS', () => {
  it('reproduces the documented examples', () => {
    expect(run('ADDRESS(2,3)')).toBe('$C$2');
    expect(run('ADDRESS(2,3,2)')).toBe('C$2');
    expect(run('ADDRESS(2,3,2,FALSE)')).toBe('R2C[3]');
    expect(run('ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")')).toBe("'[Book1]Sheet1'!R2C3");
    expect(run('ADDRESS(2,3,1,TRUE,"EXCEL SHEET")')).toBe("'EXCEL SHEET'!$C$2");
  });

  it('covers all four absolute forms', () => {
    expect(run('ADDRESS(2,3,3)')).toBe('$C2');
    expect(run('ADDRESS(2,3,4)')).toBe('C2');
    expect(run('ADDRESS(2,3,4,FALSE)')).toBe('R[2]C[3]');
    expect(run('ADDRESS(2,3,3,FALSE)')).toBe('R[2]C3');
  });

  it('leaves a plain sheet name unquoted', () => {
    expect(run('ADDRESS(1,1,1,TRUE,"Sheet1")')).toBe('Sheet1!$A$1');
  });

  it('reaches the last cell and rejects anything past it', () => {
    expect(run('ADDRESS(1048576,16384)')).toBe('$XFD$1048576');
    expect(run('ADDRESS(0,1)')).toEqual(CellError.VALUE);
    expect(run('ADDRESS(1,0)')).toEqual(CellError.VALUE);
    expect(run('ADDRESS(1048577,1)')).toEqual(CellError.VALUE);
    expect(run('ADDRESS(1,1,5)')).toEqual(CellError.VALUE);
  });
});

describe('ROW, COLUMN, ROWS, COLUMNS', () => {
  it('reports the calling cell when the argument is omitted', () => {
    expect(run('ROW()', 20, 10)).toBe(21);
    expect(run('COLUMN()', 20, 10)).toBe(11);
  });

  it('collapses a single cell and spreads a block', () => {
    expect(run('ROW(B7)')).toBe(7);
    expect(run('COLUMN(D1)')).toBe(4);
    expect(cellsOf(shape('ROW(A2:A4)'))).toEqual([2, 3, 4]);
    expect(cellsOf(shape('COLUMN(B1:D1)'))).toEqual([2, 3, 4]);
  });

  it('shapes the spread along the right axis', () => {
    const rows = shape('ROW(A2:B4)') as ArrayValue;
    expect([rows.rows, rows.cols]).toEqual([3, 1]);
    const cols = shape('COLUMN(A2:B4)') as ArrayValue;
    expect([cols.rows, cols.cols]).toEqual([1, 2]);
  });

  it('needs a reference, not a value', () => {
    expect(run('ROW("x")')).toEqual(CellError.VALUE);
    expect(run('COLUMN(1)')).toEqual(CellError.VALUE);
    expect(run('ROW({1,2})')).toEqual(CellError.VALUE);
  });

  it('counts rows and columns of both references and arrays', () => {
    expect(run('ROWS(A1:B5)')).toBe(5);
    expect(run('COLUMNS(A1:B5)')).toBe(2);
    expect(run('ROWS(A1)')).toBe(1);
    expect(run('ROWS({1,2;3,4;5,6})')).toBe(3);
    expect(run('COLUMNS({1,2;3,4;5,6})')).toBe(2);
    expect(run('ROWS(5)')).toEqual(CellError.VALUE);
  });

  it('counts a whole column against the used range, not the sheet limit', () => {
    // A deliberate engine-wide divergence, not a lookup decision: Excel answers
    // 1048576 here, and the evaluator clips every whole-column reference to the
    // used range so that summing one does not materialise a million cells. The
    // number below is therefore the engine's, and the point of asserting it is
    // that INDIRECT above gives the same answer.
    expect(run('ROWS(A:A)')).toBe(5);
  });

  it('is not volatile, despite the folklore', () => {
    for (const name of ['ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'AREAS']) {
      expect(LOOKUP_FUNCTIONS.find((f) => f.name === name)!.volatile).toBeUndefined();
    }
  });
});

describe('AREAS', () => {
  it('counts the areas of a union', () => {
    expect(run('AREAS(A1:B2)')).toBe(1);
    expect(run('AREAS(A1)')).toBe(1);
    expect(run('AREAS((A1:B2,D1))')).toBe(2);
    expect(run('AREAS((A1,B1,C1))')).toBe(3);
  });

  it('needs a reference', () => {
    expect(run('AREAS(5)')).toEqual(CellError.VALUE);
    expect(run('AREAS(1/0)')).toEqual(CellError.DIV0);
    // A reference to an error cell is still one area; the error is inside it.
    expect(run('AREAS(F1)')).toBe(1);
  });
});

describe('XLOOKUP', () => {
  // LibreOffice does not implement XLOOKUP, so these expectations come from the
  // Microsoft documentation rather than from the oracle.
  it('is exact by default and returns the fallback on a miss', () => {
    expect(run('XLOOKUP(4,A1:A5,B1:B5)')).toBe('c');
    expect(run('XLOOKUP(5,A1:A5,B1:B5)')).toEqual(CellError.NA);
    expect(run('XLOOKUP(5,A1:A5,B1:B5,"none")')).toBe('none');
    // A fallback of zero or empty text is still a fallback, not an omission.
    expect(run('XLOOKUP(5,A1:A5,B1:B5,0)')).toBe(0);
  });

  it('falls back to the next smaller or next larger item', () => {
    expect(run('XLOOKUP(5,A1:A5,B1:B5,"none",-1)')).toBe('c');
    expect(run('XLOOKUP(5,A1:A5,B1:B5,"none",1)')).toBe('d');
    expect(run('XLOOKUP(0,A1:A5,B1:B5,"none",-1)')).toBe('none');
    expect(run('XLOOKUP(99,A1:A5,B1:B5,"none",1)')).toBe('none');
  });

  it('matches wildcards only in mode 2', () => {
    expect(run('XLOOKUP("ch*",C1:C3,C1:C3,"none",2)')).toBe('cherry');
    expect(run('XLOOKUP("ch*",C1:C3,C1:C3,"none",0)')).toBe('none');
  });

  it('searches last-to-first when asked', () => {
    expect(run('XLOOKUP(10,E1:E5,B1:B5)')).toBe('a');
    expect(run('XLOOKUP(10,E1:E5,B1:B5,"none",0,-1)')).toBe('b');
    expect(run('XLOOKUP(30,E1:E5,B1:B5,"none",0,-1)')).toBe('e');
  });

  it('binary-searches sorted data in modes 2 and -2', () => {
    expect(run('XLOOKUP(8,A1:A5,B1:B5,"none",0,2)')).toBe('d');
    expect(run('XLOOKUP(5,A1:A5,B1:B5,"none",-1,2)')).toBe('c');
    expect(run('XLOOKUP(5,A1:A5,B1:B5,"none",1,2)')).toBe('d');
    expect(run('XLOOKUP(20,D1:D3,C1:C3,"none",0,-2)')).toBe('banana');
    expect(run('XLOOKUP(15,D1:D3,C1:C3,"none",-1,-2)')).toBe('cherry');
  });

  it('returns a whole row or column when the return array is wide', () => {
    expect(cellsOf(shape('XLOOKUP(2,A1:A5,A1:B5)'))).toEqual([2, 'b']);
    expect(cellsOf(shape('XLOOKUP(2,G1:I1,G1:I2)'))).toEqual([2, 'y']);
  });

  it('rejects mismatched shapes and unusable modes', () => {
    expect(run('XLOOKUP(2,A1:A5,B1:B4)')).toEqual(CellError.VALUE);
    expect(run('XLOOKUP(2,A1:B5,B1:B5)')).toEqual(CellError.VALUE);
    expect(run('XLOOKUP(2,A1:A5,B1:B5,"none",7)')).toEqual(CellError.VALUE);
    expect(run('XLOOKUP(2,A1:A5,B1:B5,"none",0,3)')).toEqual(CellError.VALUE);
  });
});

describe('XMATCH', () => {
  it('is exact by default, unlike MATCH', () => {
    expect(run('XMATCH(4,A1:A5)')).toBe(3);
    expect(run('XMATCH(5,A1:A5)')).toEqual(CellError.NA);
    expect(run('MATCH(5,A1:A5)')).toBe(3);
  });

  it('honours the match and search modes', () => {
    expect(run('XMATCH(5,A1:A5,-1)')).toBe(3);
    expect(run('XMATCH(5,A1:A5,1)')).toBe(4);
    expect(run('XMATCH("ban*",C1:C3,2)')).toBe(2);
    expect(run('XMATCH(10,E1:E5,0,1)')).toBe(1);
    expect(run('XMATCH(10,E1:E5,0,-1)')).toBe(2);
    expect(run('XMATCH(8,A1:A5,0,2)')).toBe(4);
  });

  it('needs a vector', () => {
    expect(run('XMATCH(2,A1:B5,0)')).toEqual(CellError.VALUE);
  });
});

describe('FORMULATEXT and HYPERLINK', () => {
  it('reports #N/A from FORMULATEXT, which cannot reach a formula from here', () => {
    expect(run('FORMULATEXT(A1)')).toEqual(CellError.NA);
    expect(run('FORMULATEXT("A1")')).toEqual(CellError.VALUE);
  });

  it('shows the friendly name, or the location when there is none', () => {
    expect(run('HYPERLINK("http://example.com","click")')).toBe('click');
    expect(run('HYPERLINK("http://example.com")')).toBe('http://example.com');
  });

  it('returns the friendly name as a value, not as text', () => {
    // Microsoft calls friendly_name "the jump text or numeric value that is
    // displayed in the cell", and documents HYPERLINK("...",12345) as
    // displaying the number. Stringifying here would make the result of a
    // HYPERLINK cell text for every formula that reads it.
    expect(run('HYPERLINK("http://example.com",42)')).toBe(42);
    expect(run('HYPERLINK("http://example.com",TRUE)')).toBe(true);
    // A blank friendly name displays as zero, as a blank reference does anywhere
    // a value is wanted.
    expect(run('HYPERLINK("http://example.com",Z50)')).toBe(0);
  });
});

describe('the module as a whole', () => {
  it('registers eighteen functions with no duplicates', () => {
    expect(registry.size).toBe(18);
    expect(LOOKUP_FUNCTIONS.length).toBe(18);
  });

  it('marks the post-2007 names as future functions, for xlsx round-tripping', () => {
    for (const name of ['XLOOKUP', 'XMATCH', 'FORMULATEXT']) {
      expect(registry.get(name)!.futureFunction).toBe(true);
    }
    expect(registry.get('VLOOKUP')!.futureFunction).toBe(false);
  });

  it('propagates an error argument without calling the implementation', () => {
    expect(run('VLOOKUP(1/0,A1:B5,2,FALSE)')).toEqual(CellError.DIV0);
    expect(run('INDEX(A1:A5,1/0)')).toEqual(CellError.DIV0);
    expect(run('ROWS(1/0)')).toEqual(CellError.DIV0);
  });

  it('reports an arity problem as #VALUE!', () => {
    const bad = run('CHOOSE(1)');
    expect(isError(bad)).toBe(true);
    expect((bad as CellError).code).toBe('#VALUE!');
  });
});
