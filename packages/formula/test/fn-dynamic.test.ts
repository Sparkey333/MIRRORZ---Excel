/**
 * Dynamic arrays and lambdas.
 *
 * THE ORACLE DOES NOT REACH THIS CATEGORY. LibreOffice 24.2 implements none of
 * FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY, LET, LAMBDA, BYROW, BYCOL,
 * MAP, REDUCE, SCAN, MAKEARRAY, ISOMITTED, VSTACK, HSTACK, TOROW, TOCOL,
 * WRAPROWS, WRAPCOLS, TAKE, DROP, CHOOSEROWS, CHOOSECOLS, EXPAND, ANCHORARRAY or
 * SINGLE, so fixtures/generated/formulas.calc.xlsx holds no case for any of them
 * - which the first test asserts rather than assumes. Every expected value below
 * is therefore specified from Microsoft's published documentation for the
 * function, and the places where the documentation is silent (blank ordering in
 * SORT, an unused LET binding) are called out at the point of use as our own
 * decision rather than a checked fact. See docs/oracle-divergences.md.
 *
 * The oracle is still used, twice over. The fixture workbook supplies the data,
 * and the second block cross-checks this module against numbers LibreOffice
 * really did compute: the sum of a FILTER against the oracle's SUMIF, the ends
 * of a SORT against its MIN and MAX, the second row of a sorted column against
 * its LARGE and SMALL, and a REDUCE over a column against its SUM. Those are
 * real oracle assertions - they just reach it through a function it does know.
 *
 * The registry under test holds DYNAMIC_ARRAY_FUNCTIONS plus one deliberately
 * minimal SUM, defined here rather than imported, because a lambda body needs
 * something to do and a failure in this file must never be math.ts's fault.
 *
 * Formulas are evaluated far from the data (row 31, column M) so that nothing
 * accidentally implicitly intersects.
 */

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, type Scalar, Workbook, isError } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { Evaluator } from '../src/evaluator.js';
import { DYNAMIC_ARRAY_FUNCTIONS } from '../src/functions/dynamic.js';
import { parseFormula } from '../src/parser.js';
import { ArgKind, FunctionRegistry, type FunctionSpec, p, storageName } from '../src/registry.js';
import { WorkbookStore } from '../src/store.js';
import { type Value, excelAdd, isArray, isRef } from '../src/value.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

/** A stand-in for SUM, so lambda bodies have something to compute. */
const TEST_SUM: FunctionSpec = {
  name: 'SUM',
  params: [p.rest('number', ArgKind.Array)],
  impl: (args) => {
    let total = 0;
    for (const a of args) {
      if (a === undefined) continue;
      const cells = isArray(a) ? a.data : [a as Scalar];
      for (const cell of cells) {
        if (isError(cell)) return cell;
        if (typeof cell === 'number') total = excelAdd(total, cell);
      }
    }
    return total;
  },
};

const registry = new FunctionRegistry().registerAll(DYNAMIC_ARRAY_FUNCTIONS).register(TEST_SUM);

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

const ORACLE = new Map<string, { formula: string; value: Scalar }>();
for (const { row, col, cell } of oracleSheet.entries()) {
  if (col !== 2 || !cell.formula) continue;
  const name = oracleSheet.getValue(row, 0);
  if (typeof name === 'string') ORACLE.set(name, { formula: cell.formula, value: cell.value });
}

/** Evaluate one of our formulas against the fixture workbook, well below it. */
function onFixture(formula: string): Value {
  const row = 200;
  const col = 6;
  return oracleEval.evaluate({
    ast: parseFormula(formula, { origin: { row, col } }),
    sheet: 'Formulas',
    row,
    col,
  });
}

describe('oracle: formulas.calc.xlsx', () => {
  it('has no case for any function in this category', () => {
    // If LibreOffice ever learns these, this fails and the cases become
    // assertable directly instead of through the cross-checks below.
    for (const spec of DYNAMIC_ARRAY_FUNCTIONS) {
      expect(ORACLE.has(spec.name), `oracle unexpectedly covers ${spec.name}`).toBe(false);
    }
  });

  it('agrees with the oracle SUMIF and SUMIFS through FILTER', () => {
    expect(ORACLE.get('SUMIF')!.value).toBe(518000);
    expect(onFixture('SUM(FILTER(Data!D2:D9,Data!C2:C9="Eng"))')).toBe(518000);

    expect(ORACLE.get('SUMIFS')!.value).toBe(353000);
    expect(
      onFixture('SUM(FILTER(Data!D2:D9,(Data!C2:C9="Eng")*(Data!D2:D9>170000)))'),
    ).toBe(353000);
  });

  it('agrees with the oracle MIN, MAX, LARGE and SMALL through SORT', () => {
    expect(onFixture('TAKE(SORT(Data!D2:D9),1)')).toEqual(makeRowsValue([[118000]]));
    expect(ORACLE.get('MIN')!.value).toBe(118000);
    expect(onFixture('TAKE(SORT(Data!D2:D9,,-1),1)')).toEqual(makeRowsValue([[181000]]));
    expect(ORACLE.get('MAX')!.value).toBe(181000);

    // The second largest and second smallest, which the oracle computed with
    // LARGE and SMALL.
    expect(onFixture('TAKE(DROP(SORT(Data!D2:D9,,-1),1),1)')).toEqual(makeRowsValue([[172000]]));
    expect(ORACLE.get('LARGE')!.value).toBe(172000);
    expect(onFixture('TAKE(DROP(SORT(Data!D2:D9),1),1)')).toEqual(makeRowsValue([[121000]]));
    expect(ORACLE.get('SMALL')!.value).toBe(121000);
  });

  it('agrees with the oracle SUM through REDUCE and MAP', () => {
    expect(ORACLE.get('SUM')!.value).toBe(1207000);
    expect(onFixture('REDUCE(0,Data!D2:D9,LAMBDA(a,v,a+v))')).toBe(1207000);
    expect(onFixture('SUM(MAP(Data!D2:D9,LAMBDA(v,v*2)))')).toBe(2414000);
    expect(onFixture('SUM(BYROW(Data!A2:D3,LAMBDA(r,SUM(r))))')).toBe(337003);
  });

  it('orders names by salary with SORTBY, and lists departments with UNIQUE', () => {
    // Margaret earns the maximum the oracle found, Henrietta the minimum.
    expect(onFixture('TAKE(SORTBY(Data!B2:B9,Data!D2:D9,-1),1)')).toEqual(
      makeRowsValue([['Margaret']]),
    );
    expect(onFixture('TAKE(SORTBY(Data!B2:B9,Data!D2:D9,-1),-1)')).toEqual(
      makeRowsValue([['Henrietta']]),
    );
    // Three departments, in first-appearance order; the oracle counted three
    // Science rows out of the eight.
    expect(rowsOf(onFixture('UNIQUE(Data!C2:C9)'))).toEqual([['Eng'], ['Science'], ['Ops']]);
    expect(ORACLE.get('COUNTIF')!.value).toBe(3);
    expect(onFixture('SUM(MAP(FILTER(Data!D2:D9,Data!C2:C9="Science"),LAMBDA(v,1)))')).toBe(3);
  });
});

// ---------------------------------------------------------------------------
// A synthetic workbook for everything the fixture cannot express
// ---------------------------------------------------------------------------

/**
 * Sheet S:
 *
 *        A       B     C       D    E       F         G
 *   1    3       b     1       10   TRUE    #DIV/0!   5
 *   2    1       a     text    20   FALSE             (blank)
 *   3    2       B     FALSE   30   TRUE    (blank)   7
 *   4    1       c     TRUE    40   FALSE
 *   5  (blank)   a             50   TRUE
 *
 * A has duplicates and a trailing blank, B differs only by case, C holds one of
 * each type in Excel's ranking order, E is a boolean mask, F is an error beside
 * a blank, and G is a column with a hole in it.
 */
function makeBook(): Workbook {
  const wb = new Workbook();
  const s = wb.addSheet('S');
  for (const [row, value] of [3, 1, 2, 1].entries()) s.setValue(row, 0, value);
  for (const [row, value] of ['b', 'a', 'B', 'c', 'a'].entries()) s.setValue(row, 1, value);
  s.setValue(0, 2, 1);
  s.setValue(1, 2, 'text');
  s.setValue(2, 2, false);
  s.setValue(3, 2, true);
  for (const [row, value] of [10, 20, 30, 40, 50].entries()) s.setValue(row, 3, value);
  for (const [row, value] of [true, false, true, false, true].entries()) s.setValue(row, 4, value);
  s.setValue(0, 5, CellError.DIV0);
  s.setValue(0, 6, 5);
  s.setValue(2, 6, 7);
  return wb;
}

const book = makeBook();
const ev = new Evaluator(new WorkbookStore(book), registry, { dateSystem: book.dateSystem });

/** Evaluate at M31, keeping whatever shape the result has. */
function shape(formula: string, row = 30, col = 12): Value {
  return ev.evaluate({
    ast: parseFormula(formula, { origin: { row, col } }),
    sheet: 'S',
    row,
    col,
  });
}

/** Evaluate and collapse to what a single cell would show. */
function run(formula: string, row = 30, col = 12): Scalar {
  return ev.evaluateScalar({
    ast: parseFormula(formula, { origin: { row, col } }),
    sheet: 'S',
    row,
    col,
  });
}

/** An array result as nested rows, for readable expectations. */
function rows(formula: string, row = 30, col = 12): Scalar[][] | Value {
  return rowsOf(shape(formula, row, col));
}

function rowsOf(v: Value): Scalar[][] | Value {
  if (!isArray(v)) return v;
  const out: Scalar[][] = [];
  for (let r = 0; r < v.rows; r++) {
    out.push(v.data.slice(r * v.cols, r * v.cols + v.cols));
  }
  return out;
}

/** The same shape the engine builds, for comparing whole array values. */
function makeRowsValue(cells: Scalar[][]): Value {
  const flat = cells.flat();
  return { kind: 'array', rows: cells.length, cols: cells[0]!.length, data: flat };
}

/** The error code of a result, or the result itself when it is not an error. */
function code(v: unknown): unknown {
  return isError(v) ? v.code : v;
}

// ---------------------------------------------------------------------------
// FILTER
// ---------------------------------------------------------------------------

describe('FILTER', () => {
  it('keeps the rows the mask marks', () => {
    expect(rows('FILTER(D1:D5,E1:E5)')).toEqual([[10], [30], [50]]);
    expect(rows('FILTER(D1:D5,E1:E5=FALSE)')).toEqual([[20], [40]]);
  });

  it('keeps every column of a two-dimensional array', () => {
    expect(rows('FILTER(B1:D5,E1:E5)')).toEqual([
      ['b', 1, 10],
      ['B', false, 30],
      ['a', null, 50],
    ]);
  });

  it('filters columns when the mask is a row', () => {
    expect(rows('FILTER(A1:D1,{TRUE,FALSE,TRUE,FALSE})')).toEqual([[3, 1]]);
  });

  it('coerces a numeric mask and rejects a text one', () => {
    expect(rows('FILTER(D1:D5,{1;0;1;0;1})')).toEqual([[10], [30], [50]]);
    expect(code(run('FILTER(D1:D5,{"y";"n";"y";"n";"y"})'))).toBe('#VALUE!');
  });

  it('is #CALC! when nothing matches and #N/A is never the answer', () => {
    // Microsoft: FILTER returns #CALC! for an empty result unless if_empty is
    // supplied. Returning #N/A here is the classic wrong answer.
    expect(code(run('FILTER(D1:D5,D1:D5>100)'))).toBe('#CALC!');
    expect(run('FILTER(D1:D5,D1:D5>100,"none")')).toBe('none');
    expect(run('FILTER(D1:D5,D1:D5>100,0)')).toBe(0);
  });

  it('rejects a mask that does not match either dimension', () => {
    expect(code(run('FILTER(D1:D5,E1:E3)'))).toBe('#VALUE!');
    expect(code(run('FILTER(D1:D5,B1:C5)'))).toBe('#VALUE!');
  });

  it('propagates an error inside the mask', () => {
    expect(code(run('FILTER(D1:D5,F1:F5)'))).toBe('#DIV/0!');
  });
});

// ---------------------------------------------------------------------------
// SORT and SORTBY
// ---------------------------------------------------------------------------

describe('SORT', () => {
  it('sorts ascending and descending, blanks last in both', () => {
    // Excel's sort engine always places empty cells at the end; the SORT
    // documentation is silent, so this is that engine's rule, not a guess.
    expect(rows('SORT(A1:A5)')).toEqual([[1], [1], [2], [3], [null]]);
    expect(rows('SORT(A1:A5,,-1)')).toEqual([[3], [2], [1], [1], [null]]);
  });

  it('compares text case-insensitively and stably', () => {
    // b, a, B, c, a: the two a's keep their order, and b stays ahead of B.
    expect(rows('SORT(B1:B5)')).toEqual([['a'], ['a'], ['b'], ['B'], ['c']]);
  });

  it('uses Excel type ordering rather than JavaScript ordering', () => {
    // number < text < FALSE < TRUE.
    expect(rows('SORT({TRUE;"text";FALSE;1})')).toEqual([[1], ['text'], [false], [true]]);
    expect(rows('SORT({TRUE;"text";FALSE;1},1,-1)')).toEqual([[true], [false], ['text'], [1]]);
  });

  it('keeps equal keys in their original order in both directions', () => {
    expect(rows('SORT({1,"x";1,"y";0,"z"},1)')).toEqual([
      [0, 'z'],
      [1, 'x'],
      [1, 'y'],
    ]);
    expect(rows('SORT({1,"x";1,"y";0,"z"},1,-1)')).toEqual([
      [1, 'x'],
      [1, 'y'],
      [0, 'z'],
    ]);
  });

  it('sorts by a chosen column and moves whole rows', () => {
    expect(rows('SORT(A1:B4,2)')).toEqual([
      [1, 'a'],
      [3, 'b'],
      [2, 'B'],
      [1, 'c'],
    ]);
  });

  it('sorts columns when by_col is TRUE', () => {
    expect(rows('SORT(A1:D1,,,TRUE)')).toEqual([[1, 3, 10, 'b']]);
  });

  it('orders errors after values and before blanks', () => {
    expect(rows('SORT(F1:F3)')).toEqual([[CellError.DIV0], [null], [null]]);
  });

  it('rejects an out-of-range index and an order that is not 1 or -1', () => {
    expect(code(run('SORT(D1:D5,2)'))).toBe('#VALUE!');
    expect(code(run('SORT(D1:D5,0)'))).toBe('#VALUE!');
    expect(code(run('SORT(D1:D5,1,2)'))).toBe('#VALUE!');
  });

  it('accepts several keys', () => {
    expect(rows('SORT({1,"b";1,"a";0,"c"},{1,2},{1,1})')).toEqual([
      [0, 'c'],
      [1, 'a'],
      [1, 'b'],
    ]);
  });
});

describe('SORTBY', () => {
  it('sorts one array by the values of another', () => {
    expect(rows('SORTBY(B1:B5,A1:A5)')).toEqual([['a'], ['c'], ['B'], ['b'], ['a']]);
    expect(rows('SORTBY(B1:B4,A1:A4,-1)')).toEqual([['b'], ['B'], ['a'], ['c']]);
  });

  it('applies later keys only to ties', () => {
    expect(rows('SORTBY({"p";"q";"r"},{1;1;0},1,{2;1;9},1)')).toEqual([['r'], ['q'], ['p']]);
  });

  it('rejects a by_array of the wrong length', () => {
    expect(code(run('SORTBY(B1:B5,A1:A3)'))).toBe('#VALUE!');
    expect(code(run('SORTBY(B1:B5,A1:A5,2)'))).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// UNIQUE
// ---------------------------------------------------------------------------

describe('UNIQUE', () => {
  it('returns the distinct rows in first-appearance order', () => {
    expect(rows('UNIQUE(A1:A5)')).toEqual([[3], [1], [2], [null]]);
    // "b" and "B" are one value, as everywhere else in Excel.
    expect(rows('UNIQUE(B1:B5)')).toEqual([['b'], ['a'], ['c']]);
  });

  it('compares whole rows, not cells', () => {
    expect(rows('UNIQUE({1,2;1,2;1,3})')).toEqual([
      [1, 2],
      [1, 3],
    ]);
  });

  it('works across columns when by_col is TRUE', () => {
    expect(rows('UNIQUE({1,1,2},TRUE)')).toEqual([[1, 2]]);
  });

  it('returns only the values appearing exactly once when asked', () => {
    expect(rows('UNIQUE(B1:B5,FALSE,TRUE)')).toEqual([['c']]);
    expect(rows('UNIQUE(A1:A5,FALSE,TRUE)')).toEqual([[3], [2], [null]]);
  });

  it('is #CALC! when exactly_once finds nothing', () => {
    expect(code(run('UNIQUE({1;1},FALSE,TRUE)'))).toBe('#CALC!');
  });
});

// ---------------------------------------------------------------------------
// SEQUENCE and RANDARRAY
// ---------------------------------------------------------------------------

describe('SEQUENCE', () => {
  it('fills row by row', () => {
    expect(rows('SEQUENCE(3)')).toEqual([[1], [2], [3]]);
    expect(rows('SEQUENCE(2,3)')).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ]);
  });

  it('honours start and step, including a negative one', () => {
    expect(rows('SEQUENCE(2,2,10,-5)')).toEqual([
      [10, 5],
      [0, -5],
    ]);
  });

  it('accumulates at Excel precision rather than raw binary', () => {
    // 0.1 + 0.1 + 0.1 is 0.30000000000000004 in f64; Excel shows and stores 0.3.
    expect(rows('SEQUENCE(3,1,0.1,0.1)')).toEqual([[0.1], [0.2], [0.3]]);
  });

  it('truncates its dimensions and rejects impossible ones', () => {
    expect(rows('SEQUENCE(2.9)')).toEqual([[1], [2]]);
    expect(code(run('SEQUENCE(0)'))).toBe('#CALC!');
    expect(code(run('SEQUENCE(3,0)'))).toBe('#CALC!');
    expect(code(run('SEQUENCE(-1)'))).toBe('#VALUE!');
    expect(code(run('SEQUENCE("x")'))).toBe('#VALUE!');
  });

  it('refuses a result larger than the sheet instead of allocating it', () => {
    expect(code(run('SEQUENCE(1048577)'))).toBe('#NUM!');
    expect(code(run('SEQUENCE(2000,1000)'))).toBe('#NUM!');
  });

  it('is not volatile', () => {
    expect(registry.get('SEQUENCE')!.volatile).toBeFalsy();
  });
});

describe('RANDARRAY', () => {
  it('defaults to one number between 0 and 1', () => {
    const v = run('RANDARRAY()');
    expect(typeof v).toBe('number');
    expect(v as number).toBeGreaterThanOrEqual(0);
    expect(v as number).toBeLessThan(1);
  });

  it('fills the requested shape', () => {
    const v = shape('RANDARRAY(2,3)');
    expect(isArray(v) && v.rows === 2 && v.cols === 3).toBe(true);
  });

  it('returns whole numbers inside the closed range when asked', () => {
    const v = shape('RANDARRAY(1,50,1,6,TRUE)');
    expect(isArray(v)).toBe(true);
    if (!isArray(v)) return;
    for (const cell of v.data) {
      expect(Number.isInteger(cell as number)).toBe(true);
      expect(cell as number).toBeGreaterThanOrEqual(1);
      expect(cell as number).toBeLessThanOrEqual(6);
    }
  });

  it('rejects a reversed range and fractional whole-number bounds', () => {
    expect(code(run('RANDARRAY(1,1,10,1)'))).toBe('#VALUE!');
    expect(code(run('RANDARRAY(1,1,1.5,6,TRUE)'))).toBe('#VALUE!');
    expect(code(run('RANDARRAY(0,1)'))).toBe('#VALUE!');
  });

  it('is volatile, which SEQUENCE is not', () => {
    expect(registry.get('RANDARRAY')!.volatile).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// LET and LAMBDA
// ---------------------------------------------------------------------------

describe('LET', () => {
  it('binds a name and uses it', () => {
    expect(run('LET(x,5,x*2)')).toBe(10);
    expect(run('LET(x,5,y,x+1,x*y)')).toBe(30);
  });

  it('binds a range, and the binding is still a range', () => {
    expect(run('LET(r,D1:D5,SUM(r))')).toBe(150);
    expect(isRef(shape('LET(r,D1:D5,r)'))).toBe(true);
  });

  it('binds an array and a text value', () => {
    expect(rows('LET(a,{1,2},a)')).toEqual([[1, 2]]);
    expect(run('LET(s,"ab",s&s)')).toBe('abab');
  });

  it('lets an inner binding shadow an outer one', () => {
    expect(run('LET(x,1,LET(x,2,x))')).toBe(2);
    expect(run('LET(x,1,LET(y,x,LET(x,9,y)))')).toBe(1);
  });

  it('reaches into a lambda body that does not rebind the name', () => {
    expect(rows('LET(k,10,MAP(D1:D3,LAMBDA(v,v+k)))')).toEqual([[20], [30], [40]]);
    expect(rows('LET(v,99,MAP(D1:D2,LAMBDA(v,v)))')).toEqual([[10], [20]]);
  });

  it('applies a bound lambda written as a call', () => {
    expect(run('LET(f,LAMBDA(x,x*2),f(21))')).toBe(42);
    expect(run('LET(f,LAMBDA(x,x*2),LET(g,LAMBDA(y,f(y)+1),g(5)))')).toBe(11);
  });

  it('rejects a malformed binding list', () => {
    expect(code(run('LET(x,1,y,2)'))).toBe('#VALUE!');
    expect(code(run('LET(x,1,x,2,x)'))).toBe('#NAME?');
    // A1 parses as a reference, so it is not a name LET can bind.
    expect(code(run('LET(A1,5,A1)'))).toBe('#NAME?');
  });

  it('does not evaluate a binding nothing mentions', () => {
    // Our decision, not a documented one: Microsoft says a LET value is
    // calculated once, and says nothing about a value that is never used.
    expect(run('LET(x,1/0,5)')).toBe(5);
    expect(code(run('LET(x,1/0,x)'))).toBe('#DIV/0!');
  });

  it('leaves the syntax tree exactly as it found it', () => {
    // The bindings are substituted into the tree and must be substituted back
    // out, or the second recalculation of a cached AST would see the first
    // recalculation's values.
    const ast = parseFormula('LET(x,5,y,x+1,x*y)', { origin: { row: 30, col: 12 } });
    const before = JSON.stringify(ast);
    const first = ev.evaluateScalar({ ast, sheet: 'S', row: 30, col: 12 });
    expect(JSON.stringify(ast)).toBe(before);
    const second = ev.evaluateScalar({ ast, sheet: 'S', row: 30, col: 12 });
    expect(second).toBe(first);
    expect(second).toBe(30);
  });
});

describe('LAMBDA', () => {
  it('is #CALC! when it reaches a cell without being applied', () => {
    expect(code(run('LAMBDA(x,x+1)'))).toBe('#CALC!');
  });

  it('rejects a parameter that is not a name', () => {
    expect(code(run('MAP(D1:D2,LAMBDA(1,2))'))).toBe('#VALUE!');
    expect(code(run('MAP(D1:D2,LAMBDA(x,x,x))'))).toBe('#VALUE!');
  });

  it('accepts the _xlpm-prefixed form a file stores', () => {
    expect(rows('_xlfn.MAP(D1:D2,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*3))')).toEqual([[30], [60]]);
  });
});

describe('ISOMITTED', () => {
  it('reports whether a parameter was supplied', () => {
    expect(run('LET(f,LAMBDA(x,y,ISOMITTED(y)),f(1))')).toBe(true);
    expect(run('LET(f,LAMBDA(x,y,ISOMITTED(y)),f(1,2))')).toBe(false);
    // An empty argument slot is an omission, not a blank.
    expect(run('LET(f,LAMBDA(x,y,ISOMITTED(y)),f(1,))')).toBe(true);
  });

  it('is FALSE for anything that is not an omitted parameter', () => {
    expect(run('ISOMITTED(1)')).toBe(false);
    expect(run('ISOMITTED(1/0)')).toBe(false);
    expect(run('ISOMITTED(D1:D5)')).toBe(false);
  });

  it('makes an omitted parameter #VALUE! if it is used anyway', () => {
    expect(code(run('LET(f,LAMBDA(x,y,x+y),f(1))'))).toBe('#VALUE!');
  });

  it('rejects more arguments than the lambda has parameters', () => {
    expect(code(run('LET(f,LAMBDA(x,x),f(1,2))'))).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// The higher-order family
// ---------------------------------------------------------------------------

describe('MAP', () => {
  it('applies a lambda element-wise', () => {
    expect(rows('MAP(D1:D3,LAMBDA(v,v*2))')).toEqual([[20], [40], [60]]);
    expect(rows('MAP({1,2;3,4},LAMBDA(v,v+1))')).toEqual([
      [2, 3],
      [4, 5],
    ]);
  });

  it('walks several arrays in step', () => {
    expect(rows('MAP(D1:D3,D1:D3,LAMBDA(a,b,a+b))')).toEqual([[20], [40], [60]]);
  });

  it('rejects mismatched shapes, arities and non-lambdas', () => {
    expect(code(run('MAP(D1:D3,D1:D2,LAMBDA(a,b,a+b))'))).toBe('#VALUE!');
    expect(code(run('MAP(D1:D3,LAMBDA(a,b,a+b))'))).toBe('#VALUE!');
    expect(code(run('MAP(D1:D3,5)'))).toBe('#VALUE!');
  });

  it('is #CALC! where the lambda returns an array', () => {
    expect(rows('MAP(D1:D2,LAMBDA(v,SEQUENCE(2)))')).toEqual([
      [CellError.CALC],
      [CellError.CALC],
    ]);
  });

  it('passes blanks and errors through to the lambda', () => {
    expect(rows('MAP(A4:A5,LAMBDA(v,v+1))')).toEqual([[2], [1]]);
    expect(rows('MAP(F1:F1,LAMBDA(v,ISOMITTED(v)))')).toEqual([[false]]);
  });
});

describe('BYROW and BYCOL', () => {
  it('reduces each row to one value and each column to one value', () => {
    expect(rows('BYROW({1,2;3,4},LAMBDA(r,SUM(r)))')).toEqual([[3], [7]]);
    expect(rows('BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))')).toEqual([[4, 6]]);
  });

  it('hands the lambda the whole line, not a cell', () => {
    expect(rows('BYROW(A1:B2,LAMBDA(r,SUM(r)))')).toEqual([[3], [1]]);
  });

  it('needs a one-parameter lambda', () => {
    expect(code(run('BYROW({1,2},LAMBDA(a,b,a))'))).toBe('#VALUE!');
    expect(code(run('BYCOL({1,2},5)'))).toBe('#VALUE!');
  });

  it('is #CALC! where the lambda returns an array', () => {
    expect(rows('BYROW({1;2},LAMBDA(r,SEQUENCE(2)))')).toEqual([
      [CellError.CALC],
      [CellError.CALC],
    ]);
  });
});

describe('REDUCE and SCAN', () => {
  it('folds an array from an initial value', () => {
    expect(run('REDUCE(0,D1:D5,LAMBDA(a,v,a+v))')).toBe(150);
    expect(run('REDUCE(100,{1,2,3},LAMBDA(a,v,a+v))')).toBe(106);
  });

  it('treats an omitted initial value as blank', () => {
    expect(run('REDUCE(,{1,2,3},LAMBDA(a,v,a+v))')).toBe(6);
  });

  it('reduces row by row, not column by column', () => {
    expect(run('REDUCE("",{"a","b";"c","d"},LAMBDA(a,v,a&v))')).toBe('abcd');
  });

  it('keeps every intermediate value in SCAN, in the shape of the input', () => {
    expect(rows('SCAN(0,{1,2,3},LAMBDA(a,v,a+v))')).toEqual([[1, 3, 6]]);
    expect(rows('SCAN(0,{1,2;3,4},LAMBDA(a,v,a+v))')).toEqual([
      [1, 3],
      [6, 10],
    ]);
  });

  it('needs a two-parameter lambda', () => {
    expect(code(run('REDUCE(0,{1,2},LAMBDA(a,a))'))).toBe('#VALUE!');
    expect(code(run('SCAN(0,{1,2},5)'))).toBe('#VALUE!');
  });

  it('propagates an error out of the fold', () => {
    expect(code(run('REDUCE(0,{1,2},LAMBDA(a,v,a/0))'))).toBe('#DIV/0!');
    expect(code(run('REDUCE(0,F1:F1,LAMBDA(a,v,a+v))'))).toBe('#DIV/0!');
  });
});

describe('MAKEARRAY', () => {
  it('builds from one-based row and column indices', () => {
    expect(rows('MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))')).toEqual([
      [11, 12, 13],
      [21, 22, 23],
    ]);
  });

  it('rejects an empty shape and a lambda of the wrong arity', () => {
    expect(code(run('MAKEARRAY(0,1,LAMBDA(r,c,1))'))).toBe('#VALUE!');
    expect(code(run('MAKEARRAY(1,1,LAMBDA(r,1))'))).toBe('#VALUE!');
    expect(code(run('MAKEARRAY(2000,2000,LAMBDA(r,c,1))'))).toBe('#NUM!');
  });
});

// ---------------------------------------------------------------------------
// Stacking
// ---------------------------------------------------------------------------

describe('VSTACK and HSTACK', () => {
  it('stacks in the obvious direction', () => {
    expect(rows('VSTACK({1,2},{3,4})')).toEqual([
      [1, 2],
      [3, 4],
    ]);
    expect(rows('HSTACK({1;2},{3;4})')).toEqual([
      [1, 3],
      [2, 4],
    ]);
  });

  it('pads a short argument with #N/A rather than squaring it off', () => {
    expect(rows('VSTACK({1,2},{3})')).toEqual([
      [1, 2],
      [3, CellError.NA],
    ]);
    expect(rows('HSTACK({1;2},{3;4;5})')).toEqual([
      [1, 3],
      [2, 4],
      [CellError.NA, 5],
    ]);
  });

  it('accepts scalars and ranges together', () => {
    expect(rows('VSTACK(D1:D2,99)')).toEqual([[10], [20], [99]]);
  });

  it('carries an error through as an element', () => {
    expect(rows('VSTACK(1,1/0)')).toEqual([[1], [CellError.DIV0]]);
    expect(rows('VSTACK(F1:F1,2)')).toEqual([[CellError.DIV0], [2]]);
  });
});

// ---------------------------------------------------------------------------
// Reshaping
// ---------------------------------------------------------------------------

describe('TOROW and TOCOL', () => {
  it('scans by row by default and by column on request', () => {
    expect(rows('TOCOL({1,2;3,4})')).toEqual([[1], [2], [3], [4]]);
    expect(rows('TOCOL({1,2;3,4},0,TRUE)')).toEqual([[1], [3], [2], [4]]);
    expect(rows('TOROW({1,2;3,4})')).toEqual([[1, 2, 3, 4]]);
    expect(rows('TOROW({1,2;3,4},0,TRUE)')).toEqual([[1, 3, 2, 4]]);
  });

  it('ignores blanks, errors, or both', () => {
    expect(rows('TOCOL(G1:G3,1)')).toEqual([[5], [7]]);
    expect(rows('TOCOL(F1:F3,2)')).toEqual([[null], [null]]);
    expect(rows('TOCOL(G1:G3,0)')).toEqual([[5], [null], [7]]);
  });

  it('is #CALC! when everything is ignored', () => {
    expect(code(run('TOCOL(F1:F3,3)'))).toBe('#CALC!');
  });

  it('rejects an ignore flag outside 0 to 3', () => {
    expect(code(run('TOCOL({1},4)'))).toBe('#VALUE!');
    expect(code(run('TOROW({1},-1)'))).toBe('#VALUE!');
  });
});

describe('WRAPROWS and WRAPCOLS', () => {
  it('wraps a vector the documented way round', () => {
    expect(rows('WRAPROWS(SEQUENCE(1,5),3)')).toEqual([
      [1, 2, 3],
      [4, 5, CellError.NA],
    ]);
    expect(rows('WRAPCOLS(SEQUENCE(1,5),3)')).toEqual([
      [1, 4],
      [2, 5],
      [3, CellError.NA],
    ]);
  });

  it('pads with what it is told to pad with', () => {
    expect(rows('WRAPROWS(SEQUENCE(1,4),3,"-")')).toEqual([
      [1, 2, 3],
      [4, '-', '-'],
    ]);
  });

  it('wraps a column as readily as a row', () => {
    expect(rows('WRAPROWS(D1:D4,2)')).toEqual([
      [10, 20],
      [30, 40],
    ]);
  });

  it('rejects a two-dimensional input and a wrap count below one', () => {
    expect(code(run('WRAPROWS({1,2;3,4},2)'))).toBe('#VALUE!');
    expect(code(run('WRAPROWS(SEQUENCE(1,3),0)'))).toBe('#NUM!');
    expect(code(run('WRAPCOLS(SEQUENCE(1,3),-1)'))).toBe('#NUM!');
  });
});

describe('TAKE and DROP', () => {
  it('takes from the start for a positive count and the end for a negative one', () => {
    expect(rows('TAKE(SEQUENCE(3,3),2)')).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ]);
    expect(rows('TAKE(SEQUENCE(3,3),-1)')).toEqual([[7, 8, 9]]);
    expect(rows('TAKE(SEQUENCE(3,3),2,-2)')).toEqual([
      [2, 3],
      [5, 6],
    ]);
  });

  it('drops from the start for a positive count and the end for a negative one', () => {
    expect(rows('DROP(SEQUENCE(3,3),2)')).toEqual([[7, 8, 9]]);
    expect(rows('DROP(SEQUENCE(3,3),-2)')).toEqual([[1, 2, 3]]);
    expect(rows('DROP(SEQUENCE(2,3),,1)')).toEqual([
      [2, 3],
      [5, 6],
    ]);
  });

  it('clamps a count larger than the array', () => {
    expect(rows('TAKE(SEQUENCE(2,2),9)')).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  it('is #CALC! when the result would be empty', () => {
    expect(code(run('TAKE(SEQUENCE(3,3),0)'))).toBe('#CALC!');
    expect(code(run('DROP(SEQUENCE(3,3),3)'))).toBe('#CALC!');
    expect(code(run('DROP(SEQUENCE(3,3),9)'))).toBe('#CALC!');
  });
});

describe('CHOOSEROWS and CHOOSECOLS', () => {
  it('picks lines in the order asked for, counting back from the end', () => {
    expect(rows('CHOOSEROWS(SEQUENCE(4,1),1,-1)')).toEqual([[1], [4]]);
    expect(rows('CHOOSECOLS(SEQUENCE(1,3),3,1)')).toEqual([[3, 1]]);
  });

  it('accepts an array of indices and repeats', () => {
    expect(rows('CHOOSEROWS(SEQUENCE(3,1),{1,1,2})')).toEqual([[1], [1], [2]]);
  });

  it('rejects zero and out-of-range indices', () => {
    expect(code(run('CHOOSEROWS(SEQUENCE(2,1),0)'))).toBe('#VALUE!');
    expect(code(run('CHOOSEROWS(SEQUENCE(2,1),3)'))).toBe('#VALUE!');
    expect(code(run('CHOOSECOLS(SEQUENCE(1,2),-3)'))).toBe('#VALUE!');
  });
});

describe('EXPAND', () => {
  it('pads out to the requested size', () => {
    expect(rows('EXPAND({1,2},2,3)')).toEqual([
      [1, 2, CellError.NA],
      [CellError.NA, CellError.NA, CellError.NA],
    ]);
    expect(rows('EXPAND({1},2,2,0)')).toEqual([
      [1, 0],
      [0, 0],
    ]);
  });

  it('leaves a dimension alone when it is omitted', () => {
    expect(rows('EXPAND({1,2},,3,"-")')).toEqual([[1, 2, '-']]);
  });

  it('refuses to shrink', () => {
    expect(code(run('EXPAND(SEQUENCE(2,2),1)'))).toBe('#VALUE!');
    expect(code(run('EXPAND(SEQUENCE(2,2),2,1)'))).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// The spill-reference pair
// ---------------------------------------------------------------------------

describe('ANCHORARRAY and SINGLE', () => {
  it('ANCHORARRAY returns the anchor reference itself', () => {
    // The spill rectangle lives in the dependency graph, so this is the same
    // answer the `#` operator gives today.
    const v = shape('ANCHORARRAY(D1)');
    expect(isRef(v)).toBe(true);
    expect(run('ANCHORARRAY(D1)')).toBe(10);
  });

  it('SINGLE intersects a range with the formula row', () => {
    expect(run('SINGLE(D1:D5)', 2, 12)).toBe(30);
    expect(code(run('SINGLE(D1:D5)', 30, 12))).toBe('#VALUE!');
    expect(run('SINGLE(D1:D1)')).toBe(10);
  });

  it('SINGLE takes the first element of an array', () => {
    expect(run('SINGLE({7,8;9,10})')).toBe(7);
    expect(run('SINGLE(42)')).toBe(42);
  });
});

// ---------------------------------------------------------------------------
// Registry metadata
// ---------------------------------------------------------------------------

describe('metadata', () => {
  it('registers every documented function and nothing volatile but RANDARRAY', () => {
    const documented = [
      'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'LET', 'LAMBDA',
      'BYROW', 'BYCOL', 'MAP', 'REDUCE', 'SCAN', 'MAKEARRAY', 'ISOMITTED', 'VSTACK',
      'HSTACK', 'TOROW', 'TOCOL', 'WRAPROWS', 'WRAPCOLS', 'TAKE', 'DROP', 'CHOOSEROWS',
      'CHOOSECOLS', 'EXPAND', 'ANCHORARRAY', 'SINGLE',
    ];
    for (const name of documented) expect(registry.has(name), name).toBe(true);
    for (const spec of DYNAMIC_ARRAY_FUNCTIONS) {
      expect(spec.summary, `${spec.name} has no summary`).toBeTruthy();
      if (spec.name !== 'RANDARRAY') expect(spec.volatile ?? false, spec.name).toBe(false);
    }
  });

  it('stores the future-function prefixes the file format needs', () => {
    expect(storageName('FILTER')).toBe('_xlfn._xlws.FILTER');
    expect(storageName('SINGLE')).toBe('_xlfn._xlws.SINGLE');
    expect(storageName('ANCHORARRAY')).toBe('_xlfn._xlws.ANCHORARRAY');
    expect(storageName('SEQUENCE')).toBe('_xlfn.SEQUENCE');
    expect(storageName('LAMBDA')).toBe('_xlfn.LAMBDA');
    // A formula still carrying the prefix has to resolve to the same function.
    expect(registry.get('_xlfn._xlws.FILTER')!.name).toBe('FILTER');
    expect(registry.get('_xlfn.TAKE')!.name).toBe('TAKE');
  });
});
