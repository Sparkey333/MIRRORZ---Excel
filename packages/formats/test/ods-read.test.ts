import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { isError, serialToParts } from '@mirrorz/core';
import { OdsError, looksLikeOds, openFormulaToA1, readOds } from '../src/ods/read.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const bytes = (name: string) => new Uint8Array(readFileSync(new URL(name, FIXTURES)));

describe('detection', () => {
  it('recognises an ODS by its mimetype entry', () => {
    expect(looksLikeOds(bytes('basic-types.ods'))).toBe(true);
  });

  it('does not mistake an xlsx for an ODS', () => {
    // Both are zips, so the mimetype entry is what distinguishes them.
    expect(looksLikeOds(bytes('basic-types.xlsx'))).toBe(false);
  });

  it('does not mistake a non-zip for an ODS', () => {
    expect(looksLikeOds(bytes('basic-types.xls'))).toBe(false);
    expect(looksLikeOds(new TextEncoder().encode('nonsense'))).toBe(false);
  });

  it('refuses a zip with no content.xml', () => {
    expect(() => readOds(bytes('basic-types.xlsx'))).toThrow(OdsError);
    expect(() => readOds(bytes('basic-types.xlsx'))).toThrow(/content\.xml/);
  });
});

describe('OpenFormula translation', () => {
  it.each([
    ['of:=SUM([.A1:.A9])', 'SUM(A1:A9)'],
    ['of:=[.A1]+[.B2]', 'A1+B2'],
    ['of:=[.$A$1]', '$A$1'],
    ['of:=[Sheet2.A1]', 'Sheet2!A1'],
    ['of:=[$Sheet2.$A$1]', 'Sheet2!$A$1'],
    ['of:=SUM([Sheet2.A1:.B2])', 'SUM(Sheet2!A1:B2)'],
    ['of:=IF([.A1]>0;"y";"n")', 'IF(A1>0,"y","n")'],
    ['=[.A1]', 'A1'],
    ['of:=COM.MICROSOFT.IFS([.A1];1)', 'IFS(A1,1)'],
  ])('translates %s to %s', (input, want) => {
    expect(openFormulaToA1(input)).toBe(want);
  });

  it('quotes a sheet name that needs it', () => {
    expect(openFormulaToA1("of:=['My Sheet'.A1]")).toContain('!A1');
  });

  it('returns undefined for an empty formula', () => {
    expect(openFormulaToA1('of:=')).toBeUndefined();
  });

  it('passes through a construct it does not recognise rather than mangling it', () => {
    // A formula that fails to evaluate is recoverable; one silently rewritten
    // into a different meaning is not.
    expect(openFormulaToA1('of:=SOMETHING_ODD(1)')).toBe('SOMETHING_ODD(1)');
  });
});

describe('reading real ODS fixtures', () => {
  // These were produced by LibreOffice from our xlsx fixtures, so the expected
  // contents are known exactly.

  it('reads sheet names', () => {
    const { workbook } = readOds(bytes('features.ods'));
    expect(workbook.sheets.map((s) => s.name)).toEqual(
      expect.arrayContaining(['Features', 'Charts', 'Comments']),
    );
  });

  it('reads primitive values', () => {
    const { workbook } = readOds(bytes('basic-types.ods'));
    const sheet = workbook.sheets[0]!;
    const byLabel = new Map<string, unknown>();
    for (const { row, col, cell } of sheet.entries()) {
      if (col === 1) {
        const label = sheet.getValue(row, 0);
        if (typeof label === 'string') byLabel.set(label, cell.value);
      }
    }
    expect(byLabel.get('integer')).toBe(42);
    expect(byLabel.get('negative')).toBe(-17);
    expect(byLabel.get('string')).toBe('hello world');
    expect(byLabel.get('float')).toBeCloseTo(3.14159265358979, 10);
    expect(byLabel.get('bool_true')).toBe(true);
    expect(byLabel.get('bool_false')).toBe(false);
  });

  it('preserves unicode and embedded quotes', () => {
    const { workbook } = readOds(bytes('basic-types.ods'));
    const values = [...workbook.sheets[0]!.entries()].map((e) => e.cell.value);
    expect(values).toContain('éàü 你好 \u{1f600}');
    expect(values).toContain('he said "hi"');
  });

  it('keeps text that looks numeric as text', () => {
    const { workbook } = readOds(bytes('basic-types.ods'));
    const values = [...workbook.sheets[0]!.entries()].map((e) => e.cell.value);
    expect(values).toContain('007');
    expect(values).toContain('SEPT1');
  });

  it('reads formulas, translated into the A1 dialect', () => {
    const { workbook } = readOds(bytes('formulas.ods'));
    const fx = workbook.getSheet('Formulas')!;
    const formulas: string[] = [];
    for (const { cell } of fx.entries()) if (cell.formula) formulas.push(cell.formula);

    expect(formulas.length).toBeGreaterThan(100);
    // No bracket or dot-prefixed reference may survive the translation.
    expect(formulas.some((f) => f.includes('['))).toBe(false);
    expect(formulas.some((f) => f.startsWith('of:'))).toBe(false);
    expect(formulas.some((f) => f.includes(';'))).toBe(false);
    expect(formulas.some((f) => /^SUM\(/.test(f))).toBe(true);
  });

  it('reads cached formula results', () => {
    const { workbook } = readOds(bytes('formulas.ods'));
    const fx = workbook.getSheet('Formulas')!;
    const byLabel = new Map<string, unknown>();
    for (const { row, col, cell } of fx.entries()) {
      if (col === 2) {
        const label = fx.getValue(row, 0);
        if (typeof label === 'string') byLabel.set(label, cell.value);
      }
    }
    expect(byLabel.get('SUM')).toBe(1_207_000);
    expect(byLabel.get('MAX')).toBe(181_000);
    expect(byLabel.get('UPPER')).toBe('ABC');
  });

  it('reads error values as errors', () => {
    const { workbook } = readOds(bytes('formulas.ods'));
    const fx = workbook.getSheet('Formulas')!;
    const errors: string[] = [];
    for (const { cell } of fx.entries()) if (isError(cell.value)) errors.push(cell.value.code);
    expect(errors).toContain('#DIV/0!');
  });

  it('reads dates as serial numbers', () => {
    const { workbook } = readOds(bytes('formulas.ods'));
    const fx = workbook.getSheet('Formulas')!;
    let dateValue: unknown;
    for (const { row, col, cell } of fx.entries()) {
      if (col === 2 && fx.getValue(row, 0) === 'DATE') dateValue = cell.value;
    }
    expect(typeof dateValue).toBe('number');
    expect(serialToParts(dateValue as number)).toMatchObject({ year: 2024, month: 2, day: 29 });
  });

  it('reads merged ranges from row and column spans', () => {
    const { workbook } = readOds(bytes('styling.ods'));
    expect(workbook.sheets[0]!.merges.length).toBeGreaterThan(0);
  });
});

describe('repeat counts', () => {
  /**
   * The property that matters most: ODS says "the rest of this row is empty"
   * with a repeat count in the thousands, and "the rest of the sheet is empty"
   * with one in the millions. Honouring them for addressing while refusing to
   * materialise them is the difference between opening a file and exhausting
   * memory.
   */
  it('does not materialise an enormous empty run', () => {
    const { workbook } = readOds(bytes('basic-types.ods'));
    const sheet = workbook.sheets[0]!;
    // The fixture has a few dozen populated cells, not tens of thousands.
    expect(sheet.cellCount).toBeLessThan(500);
    expect(sheet.cellCount).toBeGreaterThan(10);
  });

  it('keeps cells at their correct addresses despite repeat runs', () => {
    const { workbook } = readOds(bytes('basic-types.ods'));
    const sheet = workbook.sheets[0]!;
    // Column A row 1 is the header; if repeat counts were mishandled the whole
    // sheet would be shifted.
    expect(sheet.getValue(0, 0)).toBe('label');
    expect(sheet.getValue(0, 1)).toBe('value');
    expect(sheet.getValue(1, 0)).toBe('integer');
    expect(sheet.getValue(1, 1)).toBe(42);
  });

  it('reads a sparse sheet without inventing cells', () => {
    const { workbook } = readOds(bytes('edge-cases.ods'));
    const sparse = workbook.getSheet('Sparse');
    if (sparse) {
      expect(sparse.getValue(0, 0)).toBe('top-left');
      // ODS from LibreOffice clips to its own grid limits, so only assert that
      // what survived is at the right place and that nothing was invented.
      expect(sparse.cellCount).toBeLessThan(100);
    }
  });
});
