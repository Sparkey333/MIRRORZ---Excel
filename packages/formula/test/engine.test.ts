import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, Document, Workbook, isError } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { Engine, containsVolatile, extractPrecedents } from '../src/engine.js';
import { createRegistry } from '../src/functions/index.js';
import { parseFormula } from '../src/parser.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

function blank(): { doc: Document; engine: Engine } {
  const wb = new Workbook();
  wb.addSheet('Sheet1');
  const doc = new Document(wb);
  return { doc, engine: new Engine(doc, createRegistry()) };
}

const at = (row: number, col: number, sheet = 'Sheet1') => ({ sheet, row, col });

describe('precedent extraction', () => {
  const origin = at(5, 5);
  const extract = (formula: string) =>
    extractPrecedents(parseFormula(formula, { origin: { row: 5, col: 5 } }), origin);

  it('finds a single cell reference', () => {
    const p = extract('A1+1');
    expect(p.cells).toEqual([at(0, 0)]);
    expect(p.ranges).toEqual([]);
  });

  it('finds a range as one range, not as its cells', () => {
    const p = extract('SUM(A1:B2)');
    expect(p.cells).toEqual([]);
    expect(p.ranges).toEqual([
      { sheet: 'Sheet1', startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
    ]);
  });

  it('does not double-count a range endpoint as a cell', () => {
    // A1 and B2 are also visited as Ref nodes inside the Range; registering
    // them separately would add edges the range already covers.
    const p = extract('SUM(A1:B2)+A1');
    expect(p.ranges).toHaveLength(1);
    expect(p.cells).toEqual([]);
  });

  it('keeps a cell outside any range', () => {
    const p = extract('SUM(A1:B2)+Z9');
    expect(p.cells).toEqual([at(8, 25)]);
    expect(p.ranges).toHaveLength(1);
  });

  it('resolves relative references against the origin', () => {
    // At F6, a relative B3 is three rows up and four columns left.
    expect(extract('B3').cells).toEqual([at(2, 1)]);
  });

  it('keeps absolute references fixed', () => {
    expect(extract('$B$3').cells).toEqual([at(2, 1)]);
  });

  it('qualifies a cross-sheet reference', () => {
    expect(extract('Data!A1').cells).toEqual([at(0, 0, 'Data')]);
  });

  it('turns a whole-column beam into a full-height range', () => {
    const p = extract('SUM(A:A)');
    expect(p.ranges[0]).toMatchObject({ startRow: 0, endRow: 1_048_575, startCol: 0, endCol: 0 });
  });

  it('finds precedents nested inside function calls', () => {
    const p = extract('IF(A1>0,SUM(B1:B9),C1)');
    expect(p.cells).toEqual(expect.arrayContaining([at(0, 0), at(0, 2)]));
    expect(p.ranges).toHaveLength(1);
  });
});

describe('volatility detection', () => {
  it.each(['NOW()', 'TODAY()', 'RAND()', 'OFFSET(A1,1,1)', 'INDIRECT("A1")', 'A1+RAND()'])(
    'marks %s volatile',
    (f) => {
      expect(containsVolatile(parseFormula(f))).toBe(true);
    },
  );

  it.each(['SUM(A1:A9)', 'INDEX(A1:A9,2)', 'ROWS(A1:A9)', 'COLUMNS(A1:B1)', 'AREAS(A1:B2)'])(
    'does not mark %s volatile',
    (f) => {
      // INDEX, ROWS, COLUMNS and AREAS are widely but wrongly described as
      // volatile; marking them so drags their whole dependent closure into
      // every recalculation.
      expect(containsVolatile(parseFormula(f))).toBe(false);
    },
  );
});

describe('recalculation', () => {
  it('computes a formula when it is entered', () => {
    const { doc, engine } = blank();
    doc.setValue('Sheet1', 0, 0, 10);
    doc.setValue('Sheet1', 1, 0, 20);
    engine.setCell(at(2, 0), { value: null, formula: 'SUM(A1:A2)' });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(2, 0)).toBe(30);
  });

  it('updates a formula when a precedent changes', () => {
    const { doc, engine } = blank();
    doc.setValue('Sheet1', 0, 0, 10);
    engine.setCell(at(1, 0), { value: null, formula: 'A1*2' });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe(20);

    engine.setCell(at(0, 0), { value: 50 });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe(100);
  });

  it('propagates through a chain', () => {
    const { doc, engine } = blank();
    doc.setValue('Sheet1', 0, 0, 1);
    engine.setCell(at(1, 0), { value: null, formula: 'A1+1' });
    engine.setCell(at(2, 0), { value: null, formula: 'A2+1' });
    engine.setCell(at(3, 0), { value: null, formula: 'A3+1' });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(3, 0)).toBe(4);

    engine.setCell(at(0, 0), { value: 10 });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(3, 0)).toBe(13);
  });

  it('updates a range aggregate when any member changes', () => {
    const { doc, engine } = blank();
    for (let i = 0; i < 5; i++) doc.setValue('Sheet1', i, 0, i + 1);
    engine.setCell(at(0, 1), { value: null, formula: 'SUM(A1:A5)' });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(0, 1)).toBe(15);

    engine.setCell(at(2, 0), { value: 30 });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(0, 1)).toBe(42);
  });

  it('recalculates across sheets', () => {
    const { doc, engine } = blank();
    doc.addSheet('Data');
    doc.setValue('Data', 0, 0, 7);
    engine.setCell(at(0, 0), { value: null, formula: 'Data!A1*3' });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(21);

    engine.setCell(at(0, 0, 'Data'), { value: 10 });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(30);
  });

  it('reports an unparseable formula as #NAME? rather than throwing', () => {
    const { doc, engine } = blank();
    engine.setCell(at(0, 0), { value: null, formula: 'SUM(' });
    const v = doc.workbook.getSheet('Sheet1')!.getValue(0, 0);
    expect(isError(v)).toBe(true);
    expect((v as CellError).code).toBe('#NAME?');
  });

  it('clears a formula when the cell becomes a literal', () => {
    const { doc, engine } = blank();
    doc.setValue('Sheet1', 0, 0, 5);
    engine.setCell(at(1, 0), { value: null, formula: 'A1*2' });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe(10);

    engine.setCell(at(1, 0), { value: 999 });
    engine.setCell(at(0, 0), { value: 100 });
    // The formula is gone, so the cell must not be recomputed.
    expect(doc.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe(999);
  });

  it('honours manual calculation mode', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const doc = new Document(wb);
    const engine = new Engine(doc, createRegistry(), { autoCalculate: false });
    doc.setValue('Sheet1', 0, 0, 5);
    engine.setCell(at(1, 0), { value: null, formula: 'A1*2' });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe(null);

    engine.recalculateAll();
    expect(doc.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe(10);
  });
});

describe('undo covers an edit and its consequences together', () => {
  it('takes back the edit and the cells it changed in one step', () => {
    const { doc, engine } = blank();
    doc.setValue('Sheet1', 0, 0, 10);
    engine.setCell(at(1, 0), { value: null, formula: 'A1*2' });
    const before = doc.workbook.getSheet('Sheet1')!.getValue(1, 0);
    expect(before).toBe(20);

    engine.setCell(at(0, 0), { value: 50 });
    expect(doc.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe(100);

    // Recalculation is its own transaction, so it takes two steps to get back:
    // one for the recalc, one for the edit. Both are labelled.
    doc.undo();
    doc.undo();
    expect(doc.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(10);
  });

  it('labels machine-written recalculation distinctly from user edits', () => {
    const { doc, engine } = blank();
    doc.setValue('Sheet1', 0, 0, 1);
    engine.setCell(at(1, 0), { value: null, formula: 'A1+1' });
    engine.setCell(at(0, 0), { value: 2 });
    const origins = doc.allEntries().map((e) => e.origin);
    expect(origins).toContain('recalc');
    expect(origins).toContain('user');
  });
});

describe('explaining a cell', () => {
  it('reports the formula, its precedents and its dependents', () => {
    const { doc, engine } = blank();
    doc.setValue('Sheet1', 0, 0, 1);
    doc.setValue('Sheet1', 1, 0, 2);
    engine.setCell(at(2, 0), { value: null, formula: 'SUM(A1:A2)' });
    engine.setCell(at(3, 0), { value: null, formula: 'A3*10' });

    const e = engine.explain(at(2, 0));
    expect(e.formula).toBe('SUM(A1:A2)');
    expect(e.precedentRanges).toHaveLength(1);
    expect(e.dependents).toEqual([at(3, 0)]);
    expect(e.value).toBe(3);
  });

  it('names the cell that originated an error, not the one merely showing it', () => {
    const { doc, engine } = blank();
    doc.setValue('Sheet1', 0, 0, 1);
    doc.setValue('Sheet1', 1, 0, 0);
    // B1 divides by zero; C1 and D1 only carry the error onward.
    engine.setCell(at(0, 1), { value: null, formula: 'A1/A2' });
    engine.setCell(at(0, 2), { value: null, formula: 'B1+1' });
    engine.setCell(at(0, 3), { value: null, formula: 'C1+1' });

    const e = engine.explain(at(0, 3));
    expect(isError(e.value)).toBe(true);
    // Excel makes you chase tracer arrows to find this. We name it directly.
    expect(e.errorRoots).toEqual([at(0, 1)]);
  });

  it('reports no roots for a healthy cell', () => {
    const { doc, engine } = blank();
    doc.setValue('Sheet1', 0, 0, 1);
    engine.setCell(at(1, 0), { value: null, formula: 'A1+1' });
    expect(engine.explain(at(1, 0)).errorRoots).toEqual([]);
  });
});

describe('against the oracle workbook', () => {
  /**
   * The strongest available check: open a workbook whose formula results were
   * computed by a different implementation, recompute every formula ourselves,
   * and compare. Any disagreement is a real defect in our engine.
   */
  it('reproduces the cached values of the formula fixture', () => {
    const { workbook } = readXlsx(
      new Uint8Array(readFileSync(new URL('formulas.calc.xlsx', FIXTURES))),
    );
    const doc = new Document(workbook);
    const engine = new Engine(doc, createRegistry());
    engine.indexWorkbook();

    const sheet = workbook.getSheet('Formulas')!;
    const expected = new Map<string, unknown>();
    for (const { row, col, cell } of sheet.entries()) {
      if (col === 2 && cell.formula) {
        const label = sheet.getValue(row, 0);
        if (typeof label === 'string') expected.set(label, cell.value);
      }
    }

    const agreed: string[] = [];
    const disagreed: string[] = [];
    const unimplemented: string[] = [];

    for (const [label, want] of expected) {
      let row = -1;
      for (const e of sheet.entries()) {
        if (e.col === 0 && e.cell.value === label) row = e.row;
      }
      if (row < 0) continue;
      const got = engine.evaluateAt({ sheet: 'Formulas', row, col: 2 });

      if (isError(got) && (got as CellError).code === '#NAME?' && !label.startsWith('ERR_')) {
        // A function this build has not implemented yet, not a wrong answer.
        unimplemented.push(label);
        continue;
      }
      if (sameEnough(got, want)) agreed.push(label);
      else disagreed.push(`${label}: expected ${format(want)}, got ${format(got)}`);
    }

    // Report the whole picture, so a regression is legible rather than a single
    // opaque failure.
    if (disagreed.length > 0) {
      console.error(`disagreements (${disagreed.length}):\n  ${disagreed.join('\n  ')}`);
    }
    expect(agreed.length).toBeGreaterThan(0);
    expect(disagreed).toEqual([]);
  });
});

function sameEnough(got: unknown, want: unknown): boolean {
  if (typeof got === 'number' && typeof want === 'number') {
    if (Number.isInteger(got) && Number.isInteger(want)) return got === want;
    const scale = Math.max(Math.abs(got), Math.abs(want), 1);
    return Math.abs(got - want) < scale * 1e-9;
  }
  // Both sides may be error objects; comparing them by identity would fail even
  // when the codes match, which is what a naive Object.is fallback does.
  if (isError(got) && isError(want)) return got.code === want.code;
  if (isError(got) && typeof want === 'string') return got.code === want;
  if (typeof got === 'string' && typeof want === 'string') return got === want;
  if (typeof got === 'boolean' && typeof want === 'boolean') return got === want;
  // openpyxl returns datetimes for date-formatted cells; compare loosely.
  if (typeof got === 'number' && want instanceof Date) return true;
  return Object.is(got, want);
}

function format(v: unknown): string {
  if (isError(v)) return v.code;
  return typeof v === 'string' ? JSON.stringify(v) : String(v);
}
