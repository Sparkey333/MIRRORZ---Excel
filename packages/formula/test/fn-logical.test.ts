/**
 * Logical and information functions.
 *
 * The first block replays the LibreOffice-recalculated fixtures: every case in
 * formulas.calc.xlsx whose formula this module can evaluate on its own is
 * re-parsed, re-evaluated against a registry holding LOGICAL_FUNCTIONS alone,
 * and compared with the value the oracle actually computed. A failure there is
 * never another category's fault.
 *
 * The rest is a hand-built workbook, because the fixture has one case per
 * function and the interesting behaviour of this module is in the corners:
 * which branches are evaluated at all, which arguments may carry an error
 * through, and the difference between a blank cell, an empty string and a zero.
 */

import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, type Scalar, Workbook, isError } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { type Ast, Node } from '../src/ast.js';
import { Evaluator } from '../src/evaluator.js';
import { LOGICAL_FUNCTIONS } from '../src/functions/logical.js';
import { parseFormula } from '../src/parser.js';
import { FunctionRegistry } from '../src/registry.js';
import { WorkbookStore } from '../src/store.js';
import { type Value, isArray } from '../src/value.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

const registry = new FunctionRegistry().registerAll(LOGICAL_FUNCTIONS);

// ---------------------------------------------------------------------------
// The oracle
// ---------------------------------------------------------------------------

const { workbook: fixture } = readXlsx(
  new Uint8Array(readFileSync(new URL('formulas.calc.xlsx', FIXTURES))),
);
const fixtureEval = new Evaluator(new WorkbookStore(fixture), registry, {
  dateSystem: fixture.dateSystem,
});

/** Case name -> the formula the oracle recalculated and the value it produced. */
const oracle = new Map<string, { formula: string; value: Scalar; row: number }>();
for (const { row, col, cell } of fixture.getSheet('Formulas')!.entries()) {
  if (col !== 2 || !cell.formula) continue;
  const name = fixture.getSheet('Formulas')!.getValue(row, 0);
  if (typeof name === 'string') oracle.set(name, { formula: cell.formula, value: cell.value, row });
}

function replay(name: string): { actual: Value; expected: Scalar } {
  const c = oracle.get(name);
  expect(c, `case ${name} missing from the fixture`).toBeDefined();
  const actual = fixtureEval.evaluate({
    ast: parseFormula(c!.formula, { origin: { row: c!.row, col: 2 } }),
    sheet: 'Formulas',
    row: c!.row,
    col: 2,
  });
  return { actual, expected: c!.value };
}

function sameScalar(actual: Value, expected: Scalar): void {
  if (isError(expected)) {
    expect(isError(actual) ? actual.code : actual).toBe(expected.code);
    return;
  }
  expect(actual).toEqual(expected);
}

describe('oracle: formulas.calc.xlsx', () => {
  const cases = [
    'IF', 'IFS', 'AND', 'OR', 'NOT', 'XOR', 'IFERROR', 'IFNA',
    'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'ISERR', 'ISNA', 'ISLOGICAL',
    'ISEVEN', 'ISODD', 'N', 'TYPE',
    // The stored value of these three is a real error value, so they check that
    // errors survive the round trip as values rather than as text.
    'ERR_DIV0', 'ERR_VALUE', 'ERR_NAME',
  ];
  for (const name of cases) {
    it(`reproduces ${name}`, () => {
      const { actual, expected } = replay(name);
      sameScalar(actual, expected);
    });
  }
});

// ---------------------------------------------------------------------------
// A hand-built workbook for everything the fixture does not reach.
// ---------------------------------------------------------------------------

const book = new Workbook();
const s = book.addSheet('S');
book.addSheet('Calc');

// A1:A5 mixes the value kinds AND and OR have to sort out; A5 stays blank.
s.setValue(0, 0, 1);
s.setValue(1, 0, 2);
s.setValue(2, 0, '3');
s.setValue(3, 0, true);
// C1:C3 carries an error value in the middle.
s.setValue(0, 2, 5);
s.setValue(1, 2, CellError.DIV0);
s.setValue(2, 2, 7);
// D1:D4 is all text, so it holds no logical values at all.
s.setValue(0, 3, 'apple');
s.setValue(1, 3, 'apricot');
s.setValue(2, 3, 'banana');
s.setValue(3, 3, 'apple');
// F1:F3: a zero, a non-zero, and an empty string, which is text and not blank.
s.setValue(0, 5, 0);
s.setValue(1, 5, 7);
s.setValue(2, 5, '');

const ev = new Evaluator(new WorkbookStore(book), registry);

/** Evaluate at Calc!A1, keeping arrays and references intact. */
function calc(formula: string): Value {
  return ev.evaluate({
    ast: parseFormula(formula, { origin: { row: 0, col: 0 } }),
    sheet: 'Calc',
    row: 0,
    col: 0,
  });
}

function code(formula: string): string {
  const v = calc(formula);
  return isError(v) ? v.code : `not an error: ${String(v)}`;
}

function grid(formula: string): Scalar[] {
  const v = calc(formula);
  if (!isArray(v)) throw new Error(`expected an array, got ${String(v)}`);
  return v.data;
}

// ---------------------------------------------------------------------------
// Short-circuiting
// ---------------------------------------------------------------------------

describe('short-circuiting', () => {
  it('never evaluates the branch IF did not take', () => {
    expect(calc('IF(FALSE,1/0,"ok")')).toBe('ok');
    expect(calc('IF(TRUE,"ok",1/0)')).toBe('ok');
  });

  it('still propagates an error in the condition itself', () => {
    expect(code('IF(1/0,"a","b")')).toBe('#DIV/0!');
  });

  it('evaluates only the IFS pair that matched', () => {
    expect(calc('IFS(FALSE,1/0,TRUE,"ok")')).toBe('ok');
    expect(calc('IFS(TRUE,"ok",TRUE,1/0)')).toBe('ok');
  });

  it('evaluates only the SWITCH result that matched', () => {
    expect(calc('SWITCH("b","a",1/0,"b","ok")')).toBe('ok');
  });

  it('leaves the IFERROR fallback unevaluated when there is no error', () => {
    expect(calc('IFERROR(5,1/0)')).toBe(5);
    expect(calc('IFNA(5,1/0)')).toBe(5);
  });

  it('does not short-circuit AND and OR, because Excel does not', () => {
    // AND evaluates every argument; guarding a division is what IF is for.
    expect(code('AND(FALSE,1/0)')).toBe('#DIV/0!');
    expect(code('OR(TRUE,1/0)')).toBe('#DIV/0!');
  });
});

// ---------------------------------------------------------------------------
// IF, IFS, SWITCH
// ---------------------------------------------------------------------------

describe('IF', () => {
  it('coerces the condition the way Excel does', () => {
    expect(calc('IF(1,"a","b")')).toBe('a');
    expect(calc('IF(0,"a","b")')).toBe('b');
    expect(calc('IF("TRUE","a","b")')).toBe('a');
    expect(calc('IF("false","a","b")')).toBe('b');
    expect(calc('IF(S!Z1,"a","b")')).toBe('b');
  });

  it('rejects a condition that is text but not a logical word', () => {
    expect(code('IF("1","a","b")')).toBe('#VALUE!');
    expect(code('IF("x","a","b")')).toBe('#VALUE!');
  });

  it('returns FALSE when the false branch is omitted', () => {
    expect(calc('IF(FALSE,"a")')).toBe(false);
    expect(calc('IF(TRUE,"a")')).toBe('a');
  });

  it('turns a blank branch into zero', () => {
    expect(calc('IF(TRUE,)')).toBe(0);
    expect(calc('IF(TRUE,S!Z1)')).toBe(0);
    expect(calc('IF(FALSE,1,)')).toBe(0);
  });

  it('keeps an empty string distinct from a blank', () => {
    expect(calc('IF(TRUE,"")')).toBe('');
    expect(calc('IF(TRUE,S!F3)')).toBe('');
  });

  it('passes a multi-cell reference through so it can still be aggregated', () => {
    const v = calc('IF(TRUE,S!A1:A4)');
    expect(v).toMatchObject({ kind: 'ref', startRow: 0, endRow: 3 });
  });
});

describe('IFS', () => {
  it('takes the first true condition, not the best one', () => {
    expect(calc('IFS(TRUE,"first",TRUE,"second")')).toBe('first');
  });

  it('is #N/A when nothing matched', () => {
    expect(code('IFS(FALSE,1,FALSE,2)')).toBe('#N/A');
  });

  it('is #N/A when the last condition has no value', () => {
    expect(code('IFS(FALSE,1,TRUE)')).toBe('#N/A');
  });

  it('propagates an error from a condition it reached', () => {
    expect(code('IFS(FALSE,1,1/0,2)')).toBe('#DIV/0!');
  });

  it('rejects a non-logical condition', () => {
    expect(code('IFS("x",1)')).toBe('#VALUE!');
  });
});

describe('SWITCH', () => {
  it('matches by value and falls back to the trailing default', () => {
    expect(calc('SWITCH(2,1,"one",2,"two",3,"three")')).toBe('two');
    expect(calc('SWITCH(9,1,"one",2,"two","other")')).toBe('other');
  });

  it('is #N/A when nothing matched and there is no default', () => {
    expect(code('SWITCH(9,1,"one",2,"two")')).toBe('#N/A');
  });

  it('compares text case-insensitively, as = does', () => {
    expect(calc('SWITCH("A","a","hit","miss")')).toBe('hit');
  });

  it('matches a blank cell against zero, following comparison rules', () => {
    expect(calc('SWITCH(S!Z1,0,"blank is zero","no")')).toBe('blank is zero');
  });

  it('propagates an error in the expression', () => {
    expect(code('SWITCH(1/0,1,"one","other")')).toBe('#DIV/0!');
  });
});

// ---------------------------------------------------------------------------
// IFERROR and IFNA
// ---------------------------------------------------------------------------

describe('IFERROR and IFNA', () => {
  it('catches every error, and only #N/A respectively', () => {
    expect(calc('IFERROR(1/0,"caught")')).toBe('caught');
    expect(calc('IFERROR(NA(),"caught")')).toBe('caught');
    expect(calc('IFNA(NA(),"caught")')).toBe('caught');
    expect(code('IFNA(1/0,"caught")')).toBe('#DIV/0!');
  });

  it('reads a blank cell as zero rather than passing the blank on', () => {
    expect(calc('IFERROR(S!Z1,"caught")')).toBe(0);
  });

  it('keeps a non-error value untouched, empty string included', () => {
    expect(calc('IFERROR(S!F3,"caught")')).toBe('');
    expect(calc('IFERROR(FALSE,"caught")')).toBe(false);
  });

  it('repairs an array element-wise', () => {
    expect(grid('IFERROR({1,2}/{1,0},"z")')).toEqual([1, 'z']);
    expect(grid('IFNA({1,2}/{1,0},"z")')).toEqual([1, CellError.DIV0]);
  });

  it('propagates an error raised by the fallback itself', () => {
    expect(code('IFERROR(1/0,1/0)')).toBe('#DIV/0!');
  });
});

// ---------------------------------------------------------------------------
// AND, OR, XOR, NOT
// ---------------------------------------------------------------------------

describe('AND, OR and XOR', () => {
  it('ignores text and blanks inside a range', () => {
    // A1:A5 is 1, 2, "3", TRUE, blank: two non-zero numbers and a TRUE.
    expect(calc('AND(S!A1:A5)')).toBe(true);
    expect(calc('OR(S!A1:A5)')).toBe(true);
  });

  it('is #VALUE! when the arguments hold no logical value at all', () => {
    expect(code('AND(S!D1:D4)')).toBe('#VALUE!');
    expect(code('OR(S!D1:D4)')).toBe('#VALUE!');
    expect(code('AND(S!Z1:Z9)')).toBe('#VALUE!');
    expect(code('XOR(S!D1:D4)')).toBe('#VALUE!');
  });

  it('coerces a direct argument that a range would have ignored', () => {
    expect(calc('AND("TRUE")')).toBe(true);
    expect(calc('AND("false")')).toBe(false);
    expect(code('AND("x")')).toBe('#VALUE!');
    expect(code('AND("1")')).toBe('#VALUE!');
  });

  it('treats a zero as FALSE and any other number as TRUE', () => {
    expect(calc('AND(S!F1:F2)')).toBe(false);
    expect(calc('OR(S!F1:F2)')).toBe(true);
    expect(calc('AND(-1)')).toBe(true);
  });

  it('propagates an error sitting inside a range', () => {
    expect(code('AND(S!C1:C3)')).toBe('#DIV/0!');
    expect(code('XOR(S!C1:C3)')).toBe('#DIV/0!');
  });

  it('counts parity for XOR rather than stopping at the first TRUE', () => {
    expect(calc('XOR(TRUE,FALSE)')).toBe(true);
    expect(calc('XOR(TRUE,TRUE)')).toBe(false);
    expect(calc('XOR(TRUE,TRUE,TRUE)')).toBe(true);
    expect(calc('XOR(FALSE,FALSE)')).toBe(false);
  });

  it('skips an omitted argument', () => {
    expect(calc('AND(TRUE,)')).toBe(true);
  });
});

describe('NOT', () => {
  it('inverts what IF would have accepted', () => {
    expect(calc('NOT(TRUE)')).toBe(false);
    expect(calc('NOT(0)')).toBe(true);
    expect(calc('NOT(S!Z1)')).toBe(true);
    expect(calc('NOT("FALSE")')).toBe(true);
  });

  it('rejects text that is not a logical word', () => {
    expect(code('NOT("x")')).toBe('#VALUE!');
  });

  it('propagates an error', () => {
    expect(code('NOT(1/0)')).toBe('#DIV/0!');
  });
});

describe('TRUE and FALSE', () => {
  it('are functions, not just literals', () => {
    expect(calc('TRUE()')).toBe(true);
    expect(calc('FALSE()')).toBe(false);
    expect(calc('IF(TRUE(),1,2)')).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// The IS predicates
// ---------------------------------------------------------------------------

describe('ISBLANK', () => {
  it('separates an empty cell from an empty string and a zero', () => {
    expect(calc('ISBLANK(S!Z1)')).toBe(true);
    expect(calc('ISBLANK(S!F3)')).toBe(false);
    expect(calc('ISBLANK(S!F1)')).toBe(false);
    expect(calc('ISBLANK("")')).toBe(false);
  });

  it('answers FALSE for an error instead of propagating it', () => {
    expect(calc('ISBLANK(1/0)')).toBe(false);
  });
});

describe('ISERROR, ISERR and ISNA', () => {
  it('differ only over #N/A', () => {
    expect(calc('ISERROR(1/0)')).toBe(true);
    expect(calc('ISERROR(NA())')).toBe(true);
    expect(calc('ISERR(1/0)')).toBe(true);
    expect(calc('ISERR(NA())')).toBe(false);
    expect(calc('ISNA(NA())')).toBe(true);
    expect(calc('ISNA(1/0)')).toBe(false);
  });

  it('is FALSE for ordinary values', () => {
    expect(calc('ISERROR(1)')).toBe(false);
    expect(calc('ISERROR(S!Z1)')).toBe(false);
    expect(calc('ISERR("#N/A")')).toBe(false);
  });

  it('sees an error stored in a cell', () => {
    expect(calc('ISERROR(S!C2)')).toBe(true);
  });
});

describe('ISNUMBER, ISTEXT, ISNONTEXT and ISLOGICAL', () => {
  it('answers FALSE for an error, which is what the ISNUMBER(SEARCH()) idiom needs', () => {
    expect(calc('ISNUMBER(1/0)')).toBe(false);
    expect(calc('ISTEXT(NA())')).toBe(false);
    expect(calc('ISLOGICAL(1/0)')).toBe(false);
    expect(calc('ISNONTEXT(1/0)')).toBe(true);
  });

  it('does not treat numeric-looking text as a number', () => {
    expect(calc('ISNUMBER("1")')).toBe(false);
    expect(calc('ISNUMBER(S!A3)')).toBe(false);
    expect(calc('ISTEXT(S!A3)')).toBe(true);
  });

  it('treats a blank cell as neither number nor text', () => {
    expect(calc('ISNUMBER(S!Z1)')).toBe(false);
    expect(calc('ISTEXT(S!Z1)')).toBe(false);
    expect(calc('ISNONTEXT(S!Z1)')).toBe(true);
  });

  it('counts an empty string as text', () => {
    expect(calc('ISTEXT(S!F3)')).toBe(true);
    expect(calc('ISNONTEXT(S!F3)')).toBe(false);
  });

  it('keeps booleans out of the number and text buckets', () => {
    expect(calc('ISLOGICAL(TRUE)')).toBe(true);
    expect(calc('ISLOGICAL("TRUE")')).toBe(false);
    expect(calc('ISLOGICAL(1)')).toBe(false);
    expect(calc('ISNUMBER(TRUE)')).toBe(false);
    expect(calc('ISNONTEXT(TRUE)')).toBe(true);
  });
});

describe('ISREF', () => {
  it('sees the reference rather than its contents', () => {
    expect(calc('ISREF(S!A1)')).toBe(true);
    expect(calc('ISREF(S!A1:B2)')).toBe(true);
    expect(calc('ISREF(S!Z1)')).toBe(true);
  });

  it('is FALSE for anything that is not a reference', () => {
    expect(calc('ISREF("S!A1")')).toBe(false);
    expect(calc('ISREF(1)')).toBe(false);
    expect(calc('ISREF(1/0)')).toBe(false);
  });
});

describe('ISEVEN and ISODD', () => {
  it('truncates towards zero before testing parity', () => {
    expect(calc('ISEVEN(4)')).toBe(true);
    expect(calc('ISEVEN(4.9)')).toBe(true);
    expect(calc('ISODD(3.9)')).toBe(true);
    expect(calc('ISEVEN(-3)')).toBe(false);
    expect(calc('ISODD(-3)')).toBe(true);
    expect(calc('ISEVEN(0)')).toBe(true);
  });

  it('treats a blank as zero', () => {
    expect(calc('ISEVEN(S!Z1)')).toBe(true);
    expect(calc('ISODD(S!Z1)')).toBe(false);
  });

  it('accepts numeric text but nothing else', () => {
    expect(calc('ISEVEN("4")')).toBe(true);
    expect(code('ISEVEN("x")')).toBe('#VALUE!');
  });

  it('propagates errors, unlike the other IS functions', () => {
    expect(code('ISEVEN(1/0)')).toBe('#DIV/0!');
    expect(code('ISODD(NA())')).toBe('#N/A');
  });

  it('survives magnitudes past the safe-integer range', () => {
    expect(calc('ISEVEN(1E+20)')).toBe(true);
  });
});

describe('ISFORMULA', () => {
  it('rejects an argument that is not a reference', () => {
    expect(code('ISFORMULA("A1")')).toBe('#VALUE!');
  });

  it('reports the answer as unavailable rather than inventing FALSE', () => {
    // FunctionContext exposes cell values, not their formulas, and FALSE would
    // be wrong for exactly the cells this function exists to find.
    expect(code('ISFORMULA(S!A1)')).toBe('#N/A');
  });
});

// ---------------------------------------------------------------------------
// N, TYPE, ERROR.TYPE, NA
// ---------------------------------------------------------------------------

describe('N', () => {
  it('converts logicals and leaves numbers alone', () => {
    expect(calc('N(TRUE)')).toBe(1);
    expect(calc('N(FALSE)')).toBe(0);
    expect(calc('N(7.5)')).toBe(7.5);
    expect(calc('N(S!A1)')).toBe(1);
  });

  it('is zero for text, including text that looks numeric', () => {
    expect(calc('N("abc")')).toBe(0);
    expect(calc('N("3")')).toBe(0);
    expect(calc('N(S!A3)')).toBe(0);
    expect(calc('N(S!Z1)')).toBe(0);
  });

  it('passes an error through', () => {
    expect(code('N(1/0)')).toBe('#DIV/0!');
    expect(code('N(S!C2)')).toBe('#DIV/0!');
  });
});

describe('TYPE', () => {
  it('uses Excel ordinals', () => {
    expect(calc('TYPE(1)')).toBe(1);
    expect(calc('TYPE("a")')).toBe(2);
    expect(calc('TYPE(TRUE)')).toBe(4);
    expect(calc('TYPE(1/0)')).toBe(16);
    expect(calc('TYPE({1,2})')).toBe(64);
  });

  it('reports a blank cell as a number, as Excel does', () => {
    expect(calc('TYPE(S!Z1)')).toBe(1);
  });
});

describe('ERROR.TYPE', () => {
  it('numbers the errors in the documented order', () => {
    expect(calc('ERROR.TYPE(1/0)')).toBe(2);
    expect(calc('ERROR.TYPE("a"+1)')).toBe(3);
    expect(calc('ERROR.TYPE(NOTAFUNCTION())')).toBe(5);
    expect(calc('ERROR.TYPE(NA())')).toBe(7);
    expect(calc('ERROR.TYPE(S!C2)')).toBe(2);
  });

  it('is #N/A for anything that is not an error', () => {
    expect(code('ERROR.TYPE(1)')).toBe('#N/A');
    expect(code('ERROR.TYPE(S!Z1)')).toBe('#N/A');
  });
});

describe('NA', () => {
  it('produces the error value itself', () => {
    expect(code('NA()')).toBe('#N/A');
    expect(calc('ISNA(NA())')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// CELL, INFO, SHEET, SHEETS
// ---------------------------------------------------------------------------

describe('CELL', () => {
  it('reports the geometry of an explicit reference', () => {
    expect(calc('CELL("row",S!B3)')).toBe(3);
    expect(calc('CELL("col",S!B3)')).toBe(2);
    expect(calc('CELL("address",S!B3)')).toBe('S!$B$3');
  });

  it('uses the top-left corner of a range', () => {
    expect(calc('CELL("row",S!B3:D9)')).toBe(3);
    expect(calc('CELL("col",S!B3:D9)')).toBe(2);
  });

  it('falls back to the cell holding the formula', () => {
    expect(calc('CELL("address")')).toBe('$A$1');
    expect(calc('CELL("row")')).toBe(1);
  });

  it('classifies contents as blank, label or value', () => {
    expect(calc('CELL("type",S!Z1)')).toBe('b');
    expect(calc('CELL("type",S!A3)')).toBe('l');
    expect(calc('CELL("type",S!A1)')).toBe('v');
    expect(calc('CELL("contents",S!A1)')).toBe(1);
  });

  it('is case-insensitive about the info type', () => {
    expect(calc('CELL("ROW",S!B3)')).toBe(3);
  });

  it('refuses the info types that need state we do not have', () => {
    for (const type of ['width', 'format', 'color', 'filename', 'prefix', 'protect', 'nonsense']) {
      expect(code(`CELL("${type}",S!A1)`)).toBe('#VALUE!');
    }
  });

  it('rejects a second argument that is not a reference', () => {
    expect(code('CELL("row","A1")')).toBe('#VALUE!');
  });
});

describe('INFO', () => {
  it('refuses every type, because all of them describe the host application', () => {
    for (const type of ['directory', 'numfile', 'origin', 'osversion', 'recalc', 'release',
      'system', 'totmem', 'memavail', 'memused', 'nonsense']) {
      expect(code(`INFO("${type}")`)).toBe('#VALUE!');
    }
  });
});

describe('SHEET and SHEETS', () => {
  it('counts the sheets a reference spans', () => {
    expect(calc('SHEETS(S!A1:B2)')).toBe(1);
    expect(calc('SHEETS(S!A1)')).toBe(1);
  });

  it('is #REF! for a non-reference, as Excel documents', () => {
    expect(code('SHEETS("S")')).toBe('#REF!');
  });

  it('reports the workbook-wide forms as unavailable', () => {
    // Neither the tab order nor the tab count reaches a worksheet function
    // through FunctionContext.
    expect(code('SHEETS()')).toBe('#N/A');
    expect(code('SHEET()')).toBe('#N/A');
    expect(code('SHEET(S!A1)')).toBe('#N/A');
  });

  it('is #N/A for a sheet name that does not exist', () => {
    expect(code('SHEET("Nope")')).toBe('#N/A');
  });
});

// ---------------------------------------------------------------------------
// Declared metadata
// ---------------------------------------------------------------------------

/** The argument ASTs of a parsed call, for exercising a volatility predicate. */
function callArgs(formula: string): Ast[] {
  const ast = parseFormula(formula, { origin: { row: 0, col: 0 } });
  if (ast.kind !== Node.Call) throw new Error('expected a call');
  return ast.args;
}

function volatileFor(name: string, formula: string): boolean {
  const spec = registry.get(name)!;
  const flag = spec.volatile;
  return typeof flag === 'function' ? flag(callArgs(formula)) : flag === true;
}

describe('metadata', () => {
  it('marks nothing volatile except CELL and INFO', () => {
    for (const spec of LOGICAL_FUNCTIONS) {
      if (spec.name === 'CELL' || spec.name === 'INFO') continue;
      expect(spec.volatile, `${spec.name} should not be volatile`).toBeUndefined();
    }
  });

  it('makes CELL volatile only for the shapes that depend on hidden state', () => {
    expect(volatileFor('CELL', 'CELL("row",S!A1)')).toBe(false);
    expect(volatileFor('CELL', 'CELL("address",S!A1)')).toBe(false);
    expect(volatileFor('CELL', 'CELL("contents",S!A1)')).toBe(false);
    expect(volatileFor('CELL', 'CELL("row")')).toBe(true);
    expect(volatileFor('CELL', 'CELL("format",S!A1)')).toBe(true);
    expect(volatileFor('CELL', 'CELL(S!B1,S!A1)')).toBe(true);
  });

  it('makes INFO volatile only for the types that change during a session', () => {
    expect(volatileFor('INFO', 'INFO("numfile")')).toBe(true);
    expect(volatileFor('INFO', 'INFO("recalc")')).toBe(true);
    expect(volatileFor('INFO', 'INFO("release")')).toBe(false);
    expect(volatileFor('INFO', 'INFO("system")')).toBe(false);
  });

  it('keeps CELL and INFO off the worker thread', () => {
    expect(registry.get('CELL')!.threadSafe).toBe(false);
    expect(registry.get('INFO')!.threadSafe).toBe(false);
    expect(registry.get('IF')!.threadSafe).toBe(true);
  });

  it('stores the post-2007 names with the _xlfn prefix', () => {
    for (const name of ['IFS', 'XOR', 'SWITCH', 'IFNA', 'SHEET', 'SHEETS']) {
      expect(registry.get(name)!.futureFunction, name).toBe(true);
    }
    for (const name of ['IF', 'AND', 'OR', 'NOT', 'ISBLANK', 'N', 'TYPE', 'CELL']) {
      expect(registry.get(name)!.futureFunction, name).toBe(false);
    }
  });

  it('resolves a name that still carries its storage prefix', () => {
    expect(calc('_xlfn.XOR(TRUE,FALSE)')).toBe(true);
  });

  it('declares the branch parameters lazy, which is what makes IF safe', () => {
    const if_ = registry.get('IF')!;
    expect(if_.params.map((param) => param.kind)).toEqual([0, 3, 3]);
    expect(registry.get('IFERROR')!.params[0]!.errorTransparent).toBe(true);
    expect(registry.get('ISERROR')!.params[0]!.errorTransparent).toBe(true);
    expect(registry.get('ISREF')!.params[0]!.kind).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// Arity
// ---------------------------------------------------------------------------

describe('arity', () => {
  it('rejects calls with too few or too many arguments', () => {
    expect(code('IF(TRUE)')).toBe('#VALUE!');
    expect(code('IF(TRUE,1,2,3)')).toBe('#VALUE!');
    expect(code('NOT()')).toBe('#VALUE!');
    expect(code('TRUE(1)')).toBe('#VALUE!');
    expect(code('SWITCH(1,2)')).toBe('#VALUE!');
    expect(code('IFERROR(1)')).toBe('#VALUE!');
  });
});
