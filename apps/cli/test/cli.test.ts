import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { CellError, Document, isError } from '@mirrorz/core';
import { Engine, createRegistry } from '@mirrorz/formula';
import { displayValue, load, main, save } from '../src/index.ts';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const fixture = (name: string) => new URL(name, FIXTURES).pathname;

let work: string;
beforeAll(() => {
  work = mkdtempSync(join(tmpdir(), 'mirrorz-cli-'));
});
afterAll(() => {
  rmSync(work, { recursive: true, force: true });
});

/** Run the CLI, capturing what it wrote. */
async function run(...args: string[]): Promise<{ code: number; out: string; err: string }> {
  let out = '';
  let err = '';
  const realOut = process.stdout.write.bind(process.stdout);
  const realErr = process.stderr.write.bind(process.stderr);
  process.stdout.write = ((chunk: string) => {
    out += chunk;
    return true;
  }) as typeof process.stdout.write;
  process.stderr.write = ((chunk: string) => {
    err += chunk;
    return true;
  }) as typeof process.stderr.write;
  try {
    const code = await main(['node', 'mirrorz', ...args]);
    return { code, out, err };
  } finally {
    process.stdout.write = realOut;
    process.stderr.write = realErr;
  }
}

describe('opening every supported format', () => {
  it.each([
    ['basic-types.xlsx', 'Types'],
    ['basic-types.xls', 'Types'],
    ['basic-types.ods', 'Types'],
    ['basic-types.csv', 'basic-types'],
  ])('reads %s', (file, expectedSheet) => {
    const { workbook } = load(fixture(file));
    expect(workbook.sheets.length).toBeGreaterThan(0);
    expect(workbook.sheets.map((s) => s.name)).toContain(expectedSheet);
    expect(workbook.totalCells).toBeGreaterThan(10);
  });

  it('chooses the reader by content, not by file name', () => {
    // A compound file renamed to .xlsx must still open as BIFF8.
    const disguised = join(work, 'actually-xls.xlsx');
    writeFileSync(disguised, readFileSync(fixture('basic-types.xls')));
    const { workbook } = load(disguised);
    expect(workbook.sheets[0]!.getValue(0, 0)).toBe('label');
  });

  it('reports a clear error for a file that is not a spreadsheet', async () => {
    const junk = join(work, 'junk.xlsx');
    writeFileSync(junk, 'this is not a spreadsheet');
    const r = await run('info', junk);
    expect(r.code).toBe(1);
    expect(r.err).toContain('error:');
  });
});

describe('info', () => {
  it('summarises a workbook', async () => {
    const r = await run('info', fixture('features.xlsx'));
    expect(r.code).toBe(0);
    expect(r.out).toContain('4 sheets');
    expect(r.out).toContain('Features');
    expect(r.out).toContain('HiddenSheet (hidden)');
  });

  it('emits machine-readable output on request', async () => {
    const r = await run('info', fixture('features.xlsx'), '--json');
    const parsed = JSON.parse(r.out);
    expect(parsed.sheets).toHaveLength(4);
    expect(parsed.dateSystem).toBe(1900);
    expect(parsed.sheets[0].name).toBe('Features');
  });
});

describe('show', () => {
  it('prints a range as an aligned table', async () => {
    const r = await run('show', fixture('features.xlsx'), 'Features', 'A1:C3');
    expect(r.code).toBe(0);
    expect(r.out).toContain('region');
    expect(r.out).toContain('North');
    // Column headers and row numbers make the output addressable.
    expect(r.out).toMatch(/\bA\b/);
    expect(r.out).toMatch(/^\s*1 /m);
  });

  it('applies number formats when displaying', async () => {
    const r = await run('show', fixture('styling.xlsx'), 'Styles', 'D1:E8');
    // 1234.5678 under the 0.00 format must show as 1234.57, not as the raw value.
    expect(r.out).toContain('1234.57');
    expect(r.out).toContain('%');
  });

  it('limits output rather than dumping a huge sheet', async () => {
    const r = await run('show', fixture('large.xlsx'), 'Big', '--max', '5');
    expect(r.out.split('\n').length).toBeLessThan(12);
  });
});

describe('convert', () => {
  it('converts a legacy xls to xlsx and keeps the values', async () => {
    const out = join(work, 'from-xls.xlsx');
    const r = await run('convert', fixture('basic-types.xls'), out);
    expect(r.code).toBe(0);

    const { workbook } = load(out);
    const sheet = workbook.sheets[0]!;
    const values = [...sheet.entries()].map((e) => e.cell.value);
    expect(values).toContain('hello world');
    expect(values).toContain(42);
    expect(values).toContain('éàü 你好 \u{1f600}');
    // The classic corruption cases must survive the conversion.
    expect(values).toContain('007');
    expect(values).toContain('SEPT1');
  });

  it('reports a sensible used range rather than the whole legacy grid', async () => {
    const out = join(work, 'extent.xlsx');
    await run('convert', fixture('basic-types.xls'), out);
    const bounds = load(out).workbook.sheets[0]!.bounds()!;
    // BIFF8 writes a column record spanning all 256 columns to mean "default
    // width". Treating it as per-column formatting would push the extent out
    // to column IW.
    expect(bounds.maxCol).toBeLessThan(10);
  });

  it('converts ODS to xlsx, formulas included', async () => {
    const out = join(work, 'from-ods.xlsx');
    await run('convert', fixture('formulas.ods'), out);
    const { workbook } = load(out);
    const fx = workbook.getSheet('Formulas')!;
    let formulas = 0;
    for (const { cell } of fx.entries()) if (cell.formula) formulas++;
    expect(formulas).toBeGreaterThan(100);
  });

  it('converts to CSV', async () => {
    const out = join(work, 'export.csv');
    await run('convert', fixture('features.xlsx'), out);
    const text = readFileSync(out, 'utf8');
    expect(text).toContain('region');
    expect(text).toContain('North');
  });

  it('preserves parts it does not model through an xlsx round trip', async () => {
    const out = join(work, 'preserved.xlsx');
    const r = await run('convert', fixture('features.xlsx'), out, '--json');
    const parsed = JSON.parse(r.out);
    expect(parsed.preserved).toBeGreaterThan(5);

    const { source } = load(out);
    expect(source!.order.some((p: string) => p.startsWith('xl/charts/'))).toBe(true);
    expect(source!.order.some((p: string) => p.startsWith('xl/drawings/'))).toBe(true);
  });
});

describe('calc', () => {
  it('recalculates a workbook and reports timing', async () => {
    const r = await run('calc', fixture('formulas.calc.xlsx'), '--json');
    const parsed = JSON.parse(r.out);
    expect(parsed.evaluated).toBeGreaterThan(100);
    expect(typeof parsed.elapsedMs).toBe('number');
  });

  it('applies an edit and reports what it changed', async () => {
    const r = await run('calc', fixture('formulas.calc.xlsx'), '--set', 'Data!D2=999999', '--json');
    const parsed = JSON.parse(r.out);
    // Changing a salary must move the SUM that reads the column.
    expect(parsed.changed.length).toBeGreaterThan(0);
  });
});

describe('eval', () => {
  it('evaluates a formula against a workbook', async () => {
    const r = await run('eval', fixture('formulas.calc.xlsx'), 'SUM(Data!D2:D9)');
    expect(r.code).toBe(0);
    expect(r.out.trim()).toBe('1207000');
  });

  it('accepts a leading equals sign', async () => {
    const r = await run('eval', fixture('formulas.calc.xlsx'), '=MAX(Data!D2:D9)');
    expect(r.out.trim()).toBe('181000');
  });

  it('reports an error value and exits non-zero', async () => {
    const r = await run('eval', fixture('formulas.calc.xlsx'), '1/0');
    expect(r.out.trim()).toBe('#DIV/0!');
    expect(r.code).toBe(1);
  });
});

describe('explain', () => {
  it('reports a formula, its precedents and its value', async () => {
    const r = await run('explain', fixture('formulas.calc.xlsx'), 'Formulas!C2', '--json');
    const parsed = JSON.parse(r.out);
    expect(parsed.formula).toBe('SUM(Data!D2:D9)');
    expect(parsed.value).toBe(1207000);
  });

  it('names the cell that originated an error', async () => {
    // Build a chain where the fault is two hops away from the visible error.
    const workbook = load(fixture('basic-types.xlsx')).workbook;
    const doc = new Document(workbook);
    const engine = new Engine(doc, createRegistry());
    engine.indexWorkbook();
    const sheet = workbook.sheets[0]!.name;

    doc.setValue(sheet, 100, 0, 1);
    doc.setValue(sheet, 101, 0, 0);
    engine.setCell({ sheet, row: 100, col: 1 }, { value: null, formula: 'A101/A102' });
    engine.setCell({ sheet, row: 100, col: 2 }, { value: null, formula: 'B101+1' });
    engine.setCell({ sheet, row: 100, col: 3 }, { value: null, formula: 'C101+1' });

    const explanation = engine.explain({ sheet, row: 100, col: 3 });
    expect(isError(explanation.value)).toBe(true);
    expect((explanation.value as CellError).code).toBe('#DIV/0!');
    // The visible error is in D101; the cause is in B101.
    expect(explanation.errorRoots).toEqual([{ sheet, row: 100, col: 1 }]);
  });
});

describe('macros', () => {
  it('says so plainly when a workbook has none', async () => {
    const r = await run('macros', fixture('features.xlsx'));
    expect(r.code).toBe(0);
    expect(r.out).toContain('no macros');
  });
});

describe('the full open, edit, recalculate, save, reopen cycle', () => {
  it('keeps everything through a complete round trip', async () => {
    const out = join(work, 'cycle.xlsx');
    const loaded = load(fixture('formulas.calc.xlsx'));
    const doc = new Document(loaded.workbook);
    const engine = new Engine(doc, createRegistry());
    engine.indexWorkbook();

    // Edit a value the formulas depend on.
    engine.setCell({ sheet: 'Data', row: 1, col: 3 }, { value: 200_000 });
    const recalculated = loaded.workbook.getSheet('Formulas')!.getValue(1, 2);
    // 1207000 - 165000 + 200000
    expect(recalculated).toBe(1_242_000);

    save(loaded, out);

    const reopened = load(out);
    expect(reopened.workbook.getSheet('Data')!.getValue(1, 3)).toBe(200_000);
    expect(reopened.workbook.getSheet('Formulas')!.getValue(1, 2)).toBe(1_242_000);
    expect(reopened.workbook.getSheet('Formulas')!.getCell(1, 2)?.formula).toBe('SUM(Data!D2:D9)');

    // And the undo history still reverses it.
    doc.undo();
    doc.undo();
    expect(loaded.workbook.getSheet('Data')!.getValue(1, 3)).toBe(165_000);
  });
});

describe('value display', () => {
  it.each([
    [null, undefined, ''],
    [42, undefined, '42'],
    [true, undefined, 'TRUE'],
    [false, undefined, 'FALSE'],
    ['text', undefined, 'text'],
    [1234.5678, '0.00', '1234.57'],
    [0.4567, '0.00%', '45.67%'],
  ])('displays %s as %s', (value, fmt, want) => {
    expect(displayValue(value as never, fmt)).toBe(want);
  });

  it('shows an error by its code', () => {
    expect(displayValue(CellError.DIV0, undefined)).toBe('#DIV/0!');
  });

  it('does not second-guess an empty result from the format engine', () => {
    // `;;;` is Excel's documented idiom for hiding a cell's contents, and an
    // unparseable code produces the same empty text. Since the two are
    // indistinguishable here, falling back to the raw value would break the
    // legitimate case to paper over the malformed one.
    expect(displayValue(42, ';;;')).toBe('');
    expect(displayValue(42, '[[[not a format')).toBe('');
    expect(displayValue(42, '0.00')).toBe('42.00');
  });
});

describe('usage', () => {
  it('prints help with no arguments', async () => {
    const r = await run();
    expect(r.code).toBe(0);
    expect(r.out).toContain('Usage:');
  });

  it('rejects an unknown command', async () => {
    const r = await run('nonsense');
    expect(r.code).toBe(2);
    expect(r.err).toContain('unknown command');
  });

  it('reports a missing argument rather than crashing', async () => {
    const r = await run('info');
    expect(r.code).toBe(1);
    expect(r.err).toContain('missing');
  });
});
