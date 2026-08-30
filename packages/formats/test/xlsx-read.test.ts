import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, isError, serialToParts } from '@mirrorz/core';
import { readXlsx, stripFutureFunctionPrefixes } from '../src/xlsx/read.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const open = (name: string) => readXlsx(new Uint8Array(readFileSync(new URL(name, FIXTURES))));

describe('workbook structure', () => {
  it('reads every sheet, in order, with visibility', () => {
    const { workbook, warnings } = open('features.xlsx');
    expect(warnings).toEqual([]);
    expect(workbook.sheets.map((s) => s.name)).toEqual([
      'Features',
      'Charts',
      'Comments',
      'HiddenSheet',
    ]);
    expect(workbook.getSheet('HiddenSheet')!.visibility).toBe('hidden');
    expect(workbook.getSheet('Features')!.visibility).toBe('visible');
  });

  it('identifies the flavour and reports no macros for a plain xlsx', () => {
    const r = open('features.xlsx');
    expect(r.flavour).toBe('xlsx');
    expect(r.workbook.vbaProject).toBeUndefined();
  });

  it('defaults to the 1900 date system', () => {
    expect(open('basic-types.xlsx').workbook.dateSystem).toBe(1900);
  });

  it('reads defined names', () => {
    const { workbook } = open('features.xlsx');
    const total = workbook.definedNames.find((d) => d.name === 'TotalQty');
    expect(total).toBeDefined();
    expect(total!.refersTo).toBe('Features!$B$2:$B$7');
  });

  it('handles sheet names containing quotes', () => {
    const { workbook } = open('edge-cases.xlsx');
    expect(workbook.sheets.map((s) => s.name)).toContain('Edge\'"Case');
  });
});

describe('cell values', () => {
  const byLabel = (name: string) => {
    const { workbook } = open(name);
    const sheet = workbook.sheets[0]!;
    const map = new Map<string, unknown>();
    for (const { row, col, cell } of sheet.entries()) {
      if (col === 1) {
        const label = sheet.getValue(row, 0);
        if (typeof label === 'string') map.set(label, cell.value);
      }
    }
    return map;
  };

  it('reads every primitive type', () => {
    const v = byLabel('basic-types.xlsx');
    expect(v.get('integer')).toBe(42);
    expect(v.get('negative')).toBe(-17);
    expect(v.get('float')).toBeCloseTo(3.14159265358979, 12);
    expect(v.get('string')).toBe('hello world');
    expect(v.get('bool_true')).toBe(true);
    expect(v.get('bool_false')).toBe(false);
  });

  it('preserves unicode, emoji and embedded quotes', () => {
    const v = byLabel('basic-types.xlsx');
    expect(v.get('unicode')).toBe('éàü 你好 \u{1f600}');
    expect(v.get('quote')).toBe('he said "hi"');
  });

  it('keeps text that looks numeric as text', () => {
    const v = byLabel('basic-types.xlsx');
    // These are the classic silent-corruption cases: a spreadsheet that turns
    // them into numbers or dates has destroyed the user's data.
    expect(v.get('leading_zero')).toBe('007');
    expect(v.get('looks_like_date')).toBe('1-2');
    expect(v.get('gene_name')).toBe('SEPT1');
    expect(v.get('long_number')).toBe('1234567890123456789');
  });

  it('reads extreme numeric magnitudes', () => {
    const v = byLabel('basic-types.xlsx');
    expect(v.get('tiny')).toBe(1e-300);
    expect(v.get('huge')).toBe(1e300);
    expect(v.get('zero')).toBe(0);
  });

  it('leaves genuinely empty cells absent from the store', () => {
    const { workbook } = open('basic-types.xlsx');
    const sheet = workbook.sheets[0]!;
    // The "empty" row has a label and a note but no value in column B.
    let emptyRow = -1;
    for (const { row, col, cell } of sheet.entries()) {
      if (col === 0 && cell.value === 'empty') emptyRow = row;
    }
    expect(emptyRow).toBeGreaterThan(0);
    expect(sheet.getCell(emptyRow, 1)).toBeUndefined();
  });
});

describe('the shared vs inline string encodings', () => {
  it('reads inline strings, which openpyxl writes', () => {
    const { workbook } = open('basic-types.xlsx');
    expect(workbook.sheets[0]!.getValue(0, 0)).toBe('label');
  });

  it('reads a shared string table, which LibreOffice writes', () => {
    const { workbook } = open('basic-types.calc.xlsx');
    expect(workbook.sheets[0]!.getValue(0, 0)).toBe('label');
    const values = [...workbook.sheets[0]!.entries()].map((e) => e.cell.value);
    expect(values).toContain('éàü 你好 \u{1f600}');
  });
});

describe('formulas', () => {
  it('reads formula text alongside the cached value', () => {
    const { workbook } = open('formulas.calc.xlsx');
    const fx = workbook.getSheet('Formulas')!;
    let found = 0;
    for (const { cell } of fx.entries()) {
      if (cell.formula) found++;
    }
    expect(found).toBeGreaterThan(130);
  });

  it('pairs each named case with the value LibreOffice computed', () => {
    const { workbook } = open('formulas.calc.xlsx');
    const fx = workbook.getSheet('Formulas')!;
    const cases = new Map<string, { formula?: string; value: unknown }>();
    for (const { row, col, cell } of fx.entries()) {
      if (col !== 2) continue;
      const label = fx.getValue(row, 0);
      if (typeof label === 'string') cases.set(label, { formula: cell.formula, value: cell.value });
    }
    expect(cases.get('SUM')).toMatchObject({ formula: 'SUM(Data!D2:D9)', value: 1_207_000 });
    expect(cases.get('MAX')!.value).toBe(181_000);
    expect(cases.get('COUNT')!.value).toBe(8);
    expect(cases.get('UPPER')!.value).toBe('ABC');
    expect(cases.get('POWER')!.value).toBe(1024);
  });

  it('strips the _xlfn prefix so formulas are displayable and evaluable', () => {
    const { workbook } = open('formulas.calc.xlsx');
    const fx = workbook.getSheet('Formulas')!;
    const formulas: string[] = [];
    for (const { cell } of fx.entries()) if (cell.formula) formulas.push(cell.formula);
    expect(formulas.some((f) => f.includes('_xlfn'))).toBe(false);
    expect(formulas.some((f) => f.startsWith('IFS('))).toBe(true);
  });

  it.each([
    ['_xlfn.IFS(A1>1,"a")', 'IFS(A1>1,"a")'],
    ['_xlfn._xlws.FILTER(A1:A9,B1:B9)', 'FILTER(A1:A9,B1:B9)'],
    ['SUM(A1:A9)', 'SUM(A1:A9)'],
    ['_xlfn.STDEV.S(A1:A9)+_xlfn.XLOOKUP(1,A:A,B:B)', 'STDEV.S(A1:A9)+XLOOKUP(1,A:A,B:B)'],
  ])('stripFutureFunctionPrefixes(%s)', (input, want) => {
    expect(stripFutureFunctionPrefixes(input)).toBe(want);
  });

  it('reads error values as errors, not text', () => {
    const { workbook } = open('formulas.calc.xlsx');
    const fx = workbook.getSheet('Formulas')!;
    const errors = new Map<string, unknown>();
    for (const { row, col, cell } of fx.entries()) {
      if (col !== 2) continue;
      const label = fx.getValue(row, 0);
      if (typeof label === 'string' && label.startsWith('ERR_')) errors.set(label, cell.value);
    }
    expect(isError(errors.get('ERR_DIV0'))).toBe(true);
    expect((errors.get('ERR_DIV0') as CellError).code).toBe('#DIV/0!');
    expect((errors.get('ERR_NAME') as CellError).code).toBe('#NAME?');
  });
});

describe('dates', () => {
  it('reads dates as serial numbers whose date-ness lives in the format', () => {
    const { workbook, styleTables } = open('formulas.calc.xlsx');
    const fx = workbook.getSheet('Formulas')!;
    let dateSerial: number | undefined;
    for (const { row, col, cell } of fx.entries()) {
      if (col === 2 && fx.getValue(row, 0) === 'DATE') dateSerial = cell.value as number;
    }
    expect(typeof dateSerial).toBe('number');
    expect(serialToParts(dateSerial!)).toMatchObject({ year: 2024, month: 2, day: 29 });
    expect(styleTables.cellXfs.length).toBeGreaterThan(0);
  });
});

describe('styles', () => {
  it('resolves fonts, fills and borders through the two-level lookup', () => {
    const { workbook } = open('styling.xlsx');
    const sheet = workbook.sheets[0]!;
    const styleOf = (row: number, col: number) => workbook.styles.get(sheet.getStyle(row, col));

    expect(styleOf(0, 0).font?.bold).toBe(true);
    expect(styleOf(1, 0).font?.italic).toBe(true);
    expect(styleOf(4, 0).font?.size).toBe(20);
    expect(styleOf(4, 0).font?.name).toBe('Arial');
    expect(styleOf(0, 1).fill?.pattern).toBe('solid');
    expect(styleOf(1, 1).border?.top?.style).toBe('thick');
  });

  it('reads alignment including wrap, rotation and indent', () => {
    const { workbook } = open('styling.xlsx');
    const sheet = workbook.sheets[0]!;
    const styleOf = (row: number, col: number) => workbook.styles.get(sheet.getStyle(row, col));
    expect(styleOf(2, 1).alignment?.horizontal).toBe('center');
    expect(styleOf(3, 1).alignment?.wrapText).toBe(true);
    expect(styleOf(4, 1).alignment?.textRotation).toBe(45);
    expect(styleOf(5, 1).alignment?.indent).toBe(3);
  });

  it('resolves number formats, built-in and custom', () => {
    const { workbook } = open('styling.xlsx');
    const sheet = workbook.sheets[0]!;
    const formats = new Map<string, string | undefined>();
    for (const { row, col } of sheet.entries()) {
      if (col !== 4) continue;
      const label = sheet.getValue(row, 3);
      if (typeof label === 'string') {
        formats.set(label, workbook.styles.get(sheet.getStyle(row, col)).numFmt);
      }
    }
    expect(formats.get('2dp')).toBe('0.00');
    expect(formats.get('percent')).toBe('0.00%');
    expect(formats.get('date')).toBe('yyyy-mm-dd');
    expect(formats.get('negred')).toBe('0.00;[Red]-0.00');
    expect(formats.get('text_fmt')).toBe('@');
  });

  it('interns identical styles to a single id', () => {
    const { workbook } = open('styling.xlsx');
    // Far fewer distinct styles than cells: this is the property that keeps a
    // million-row sheet from costing a million format records.
    expect(workbook.styles.size).toBeLessThan(workbook.totalCells);
  });
});

describe('sheet layout', () => {
  it('reads column widths, row heights and hidden flags', () => {
    const { workbook } = open('styling.xlsx');
    const sheet = workbook.sheets[0]!;
    expect(sheet.colWidth(0)).toBeCloseTo(22, 0);
    expect(sheet.rowHeight(3)).toBeCloseTo(48, 0);
    expect(sheet.isColHidden(6)).toBe(true);
    expect(sheet.isRowHidden(9)).toBe(true);
  });

  it('reads frozen panes', () => {
    const { workbook } = open('styling.xlsx');
    const sheet = workbook.sheets[0]!;
    // freeze_panes = "C3" means two frozen columns and two frozen rows.
    expect(sheet.view.frozenCols).toBe(2);
    expect(sheet.view.frozenRows).toBe(2);
  });

  it('reads merged ranges', () => {
    const { workbook } = open('styling.xlsx');
    const sheet = workbook.sheets[0]!;
    expect(sheet.merges).toHaveLength(1);
    const m = sheet.merges[0]!.range;
    expect(m.start).toMatchObject({ row: 0, col: 7 });
    expect(m.end).toMatchObject({ row: 1, col: 9 });
    expect(sheet.mergeAt(0, 8)).toBeDefined();
    expect(sheet.mergeAt(5, 8)).toBeUndefined();
  });

  it('reads the tab colour', () => {
    const { workbook } = open('features.xlsx');
    expect(workbook.getSheet('Comments')!.tabColor).toBe('FF00B050');
  });
});

describe('features we preserve rather than model', () => {
  it('keeps conditional formatting, validation and autofilter as raw XML', () => {
    const { workbook } = open('features.xlsx');
    const sheet = workbook.getSheet('Features')!;
    expect(sheet.preserved['conditionalFormatting']).toContain('cfRule');
    expect(sheet.preserved['dataValidations']).toContain('dataValidation');
    expect(sheet.preserved['autoFilter']).toContain('autoFilter');
  });

  it('keeps chart, drawing and table parts in the package', () => {
    const { pkg } = open('features.xlsx');
    const parts = pkg.order;
    expect(parts.some((p) => p.startsWith('xl/charts/'))).toBe(true);
    expect(parts.some((p) => p.startsWith('xl/drawings/'))).toBe(true);
    expect(parts.some((p) => p.startsWith('xl/tables/'))).toBe(true);
  });
});

describe('sparse and extreme sheets', () => {
  it('reads cells at the far corners of the grid', () => {
    const { workbook } = open('edge-cases.xlsx');
    const sparse = workbook.getSheet('Sparse')!;
    expect(sparse.getValue(0, 0)).toBe('top-left');
    expect(sparse.getValue(0, 16_383)).toBe('last column (16384)');
    expect(sparse.getValue(1_048_575, 0)).toBe('last row');
    expect(sparse.getValue(1_048_575, 16_383)).toBe('bottom-right corner');
  });

  it('stores a sparse sheet sparsely', () => {
    const { workbook } = open('edge-cases.xlsx');
    // Four cells spanning the whole grid must cost four entries, not 17 billion.
    expect(workbook.getSheet('Sparse')!.cellCount).toBe(4);
  });

  it('computes the used bounds', () => {
    const { workbook } = open('edge-cases.xlsx');
    expect(workbook.getSheet('Sparse')!.bounds()).toEqual({
      minRow: 0,
      minCol: 0,
      maxRow: 1_048_575,
      maxCol: 16_383,
    });
  });
});

describe('scale', () => {
  it('opens a 50k-row workbook and reads it correctly', () => {
    const started = performance.now();
    const { workbook } = open('large.xlsx');
    const elapsed = performance.now() - started;
    const sheet = workbook.sheets[0]!;

    expect(sheet.cellCount).toBe(50_001 * 20);
    expect(sheet.getValue(0, 0)).toBe('col1');
    // Data row r (sheet row r+1, since row 0 is the header) holds r*c in
    // columns where c % 3 !== 0, and the text "s<r>-<c>" where it is 0.
    expect(sheet.getValue(1, 1)).toBe(0); // r=0, c=2
    expect(sheet.getValue(2, 1)).toBe(2); // r=1, c=2
    expect(sheet.getValue(1, 2)).toBe('s0-3'); // c=3 is a multiple of 3
    expect(elapsed).toBeLessThan(30_000);
  });

  it('can read structure only, for an instant open', () => {
    const { workbook } = readXlsx(
      new Uint8Array(readFileSync(new URL('large.xlsx', FIXTURES))),
      { structureOnly: true },
    );
    expect(workbook.sheets.map((s) => s.name)).toEqual(['Big']);
    expect(workbook.sheets[0]!.cellCount).toBe(0);
  });

  it('can read a chosen subset of sheets', () => {
    const { workbook } = readXlsx(
      new Uint8Array(readFileSync(new URL('features.xlsx', FIXTURES))),
      { sheets: ['Charts'] },
    );
    expect(workbook.getSheet('Charts')!.cellCount).toBeGreaterThan(0);
    expect(workbook.getSheet('Features')!.cellCount).toBe(0);
  });
});

describe('robustness', () => {
  it('rejects a file that is not a package', () => {
    expect(() => readXlsx(new TextEncoder().encode('not a spreadsheet'))).toThrow();
  });

  it('reports a damaged sheet as a warning rather than failing the whole open', () => {
    // Corrupt one worksheet part in place, leaving the rest of the package valid.
    const { pkg } = open('features.xlsx');
    pkg.putText('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="notanumber"');
    const bytes = pkg.write();
    const result = readXlsx(bytes);
    expect(result.workbook.sheets).toHaveLength(4);
    expect(result.warnings.length).toBeGreaterThan(0);
    // The undamaged sheets still opened.
    expect(result.workbook.getSheet('Charts')!.cellCount).toBeGreaterThan(0);
  });
});
