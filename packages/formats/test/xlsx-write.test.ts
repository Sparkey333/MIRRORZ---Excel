import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, Workbook } from '@mirrorz/core';
import { OpcPackage } from '../src/opc.js';
import { XmlReader, XmlToken } from '../src/xml.js';
import { readXlsx } from '../src/xlsx/read.js';
import { WORKSHEET_CHILD_ORDER, emitInOrder } from '../src/xlsx/order.js';
import { restoreFutureFunctionPrefixes, writeXlsx } from '../src/xlsx/write.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const bytes = (name: string) => new Uint8Array(readFileSync(new URL(name, FIXTURES)));

/** Open, save, reopen - the cycle every trust claim rests on. */
function roundTrip(name: string) {
  const first = readXlsx(bytes(name));
  const saved = writeXlsx(first.workbook, {
    source: first.pkg,
    styleTables: first.styleTables,
    flavour: first.flavour,
    modified: new Date(0),
  });
  return { first, saved, second: readXlsx(saved) };
}

describe('element ordering', () => {
  it('emits fragments in schema sequence regardless of insertion order', () => {
    const fragments = new Map([
      ['extLst', '<extLst/>'],
      ['sheetData', '<sheetData/>'],
      ['dimension', '<dimension/>'],
      ['legacyDrawing', '<legacyDrawing/>'],
      ['drawing', '<drawing/>'],
    ]);
    expect(emitInOrder(WORKSHEET_CHILD_ORDER, fragments)).toBe(
      '<dimension/><sheetData/><drawing/><legacyDrawing/><extLst/>',
    );
  });

  it('places legacyDrawing before picture, tableParts and extLst', () => {
    // This exact ordering defect ships in at least one widely used library and
    // is what the data-driven emitter exists to prevent.
    const i = (n: string) => WORKSHEET_CHILD_ORDER.indexOf(n as never);
    expect(i('legacyDrawing')).toBeGreaterThan(i('drawing'));
    expect(i('legacyDrawing')).toBeLessThan(i('picture'));
    expect(i('legacyDrawing')).toBeLessThan(i('tableParts'));
    expect(i('legacyDrawing')).toBeLessThan(i('extLst'));
  });

  it('appends elements it does not know rather than dropping them', () => {
    const fragments = new Map([
      ['sheetData', '<sheetData/>'],
      ['somethingNew', '<somethingNew/>'],
    ]);
    expect(emitInOrder(WORKSHEET_CHILD_ORDER, fragments)).toBe('<sheetData/><somethingNew/>');
  });

  it('writes a real worksheet in schema order', () => {
    const { saved } = roundTrip('features.xlsx');
    const pkg = OpcPackage.read(saved);
    const xml = pkg.text('xl/worksheets/sheet1.xml');
    const seen: number[] = [];
    const r = new XmlReader(xml);
    for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
      if (t !== XmlToken.Open || r.depth !== 2) continue;
      const idx = WORKSHEET_CHILD_ORDER.indexOf(r.localName as never);
      if (idx >= 0) seen.push(idx);
    }
    expect(seen.length).toBeGreaterThan(3);
    expect([...seen].sort((a, b) => a - b)).toEqual(seen);
  });
});

describe('preserving round trip', () => {
  it.each(['features.xlsx', 'styling.xlsx', 'basic-types.xlsx', 'edge-cases.xlsx'])(
    'keeps every part of %s that it did not rewrite',
    (name) => {
      const { first, saved } = roundTrip(name);
      const after = OpcPackage.read(saved);

      // Parts we model are legitimately rewritten; everything else must survive
      // byte for byte, which is what keeps charts and pivot tables intact.
      const rewritten = /^(\[Content_Types\]\.xml|xl\/workbook\.xml|xl\/worksheets\/|xl\/styles\.xml|xl\/sharedStrings\.xml|.*\.rels$)/;
      for (const part of first.pkg.order) {
        if (rewritten.test(part)) continue;
        expect(after.has(part), `${name}: ${part} was dropped`).toBe(true);
        expect(Array.from(after.bytes(part)!), `${name}: ${part} changed`).toEqual(
          Array.from(first.pkg.bytes(part)!),
        );
      }
    },
  );

  it('keeps chart, drawing and table parts through a save', () => {
    const { saved } = roundTrip('features.xlsx');
    const parts = OpcPackage.read(saved).order;
    expect(parts.some((p) => p.startsWith('xl/charts/'))).toBe(true);
    expect(parts.some((p) => p.startsWith('xl/drawings/'))).toBe(true);
    expect(parts.some((p) => p.startsWith('xl/tables/'))).toBe(true);
    expect(parts.some((p) => p.startsWith('xl/theme/'))).toBe(true);
  });

  it('deletes calcChain, which is stale the moment a formula changes', () => {
    const first = readXlsx(bytes('formulas.calc.xlsx'));
    const saved = writeXlsx(first.workbook, { source: first.pkg, styleTables: first.styleTables });
    const after = OpcPackage.read(saved);
    expect(after.has('xl/calcChain.xml')).toBe(false);
    expect(after.text('xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
  });
});

describe('values survive the round trip', () => {
  it('preserves every primitive type', () => {
    const { first, second } = roundTrip('basic-types.xlsx');
    const a = first.workbook.sheets[0]!;
    const b = second.workbook.sheets[0]!;
    expect(b.cellCount).toBe(a.cellCount);
    for (const { row, col, cell } of a.entries()) {
      expect(b.getValue(row, col), `${row},${col}`).toEqual(cell.value);
    }
  });

  it('preserves unicode, emoji and embedded quotes', () => {
    const { second } = roundTrip('basic-types.xlsx');
    const values = [...second.workbook.sheets[0]!.entries()].map((e) => e.cell.value);
    expect(values).toContain('éàü 你好 \u{1f600}');
    expect(values).toContain('he said "hi"');
  });

  it('keeps text that looks numeric as text', () => {
    const { second } = roundTrip('basic-types.xlsx');
    const sheet = second.workbook.sheets[0]!;
    const byLabel = new Map<string, unknown>();
    for (const { row, col, cell } of sheet.entries()) {
      if (col === 1) {
        const label = sheet.getValue(row, 0);
        if (typeof label === 'string') byLabel.set(label, cell.value);
      }
    }
    expect(byLabel.get('leading_zero')).toBe('007');
    expect(byLabel.get('gene_name')).toBe('SEPT1');
    expect(byLabel.get('looks_like_date')).toBe('1-2');
  });

  it('preserves formulas and their cached values', () => {
    const { first, second } = roundTrip('formulas.calc.xlsx');
    const a = first.workbook.getSheet('Formulas')!;
    const b = second.workbook.getSheet('Formulas')!;
    let compared = 0;
    for (const { row, col, cell } of a.entries()) {
      if (!cell.formula) continue;
      const other = b.getCell(row, col);
      expect(other?.formula, `${row},${col}`).toBe(cell.formula);
      expect(other?.value, `${row},${col}`).toEqual(cell.value);
      compared++;
    }
    expect(compared).toBeGreaterThan(130);
  });

  it('preserves error values as errors', () => {
    const { second } = roundTrip('formulas.calc.xlsx');
    const fx = second.workbook.getSheet('Formulas')!;
    const errors: string[] = [];
    for (const { cell } of fx.entries()) {
      if (cell.value instanceof CellError) errors.push(cell.value.code);
    }
    expect(errors).toContain('#DIV/0!');
    expect(errors).toContain('#NAME?');
  });

  it('preserves styles, merges, panes and sheet layout', () => {
    const { first, second } = roundTrip('styling.xlsx');
    const a = first.workbook.sheets[0]!;
    const b = second.workbook.sheets[0]!;

    expect(b.merges.map((m) => m.range)).toEqual(a.merges.map((m) => m.range));
    expect(b.view.frozenRows).toBe(a.view.frozenRows);
    expect(b.view.frozenCols).toBe(a.view.frozenCols);
    expect(b.colWidth(0)).toBeCloseTo(a.colWidth(0), 2);
    expect(b.rowHeight(3)).toBeCloseTo(a.rowHeight(3), 2);
    expect(b.isColHidden(6)).toBe(a.isColHidden(6));
    expect(b.isRowHidden(9)).toBe(a.isRowHidden(9));

    const styleA = first.workbook.styles.get(a.getStyle(0, 0));
    const styleB = second.workbook.styles.get(b.getStyle(0, 0));
    expect(styleB.font?.bold).toBe(styleA.font?.bold);
    expect(second.workbook.styles.get(b.getStyle(4, 0)).font?.size).toBe(20);
    expect(second.workbook.styles.get(b.getStyle(3, 1)).alignment?.wrapText).toBe(true);
  });

  it('preserves number formats, built-in and custom', () => {
    const { second } = roundTrip('styling.xlsx');
    const sheet = second.workbook.sheets[0]!;
    const formats = new Map<string, string | undefined>();
    for (const { row, col } of sheet.entries()) {
      if (col !== 4) continue;
      const label = sheet.getValue(row, 3);
      if (typeof label === 'string') {
        formats.set(label, second.workbook.styles.get(sheet.getStyle(row, col)).numFmt);
      }
    }
    expect(formats.get('2dp')).toBe('0.00');
    expect(formats.get('date')).toBe('yyyy-mm-dd');
    expect(formats.get('negred')).toBe('0.00;[Red]-0.00');
  });

  it('preserves sheet visibility, tab colour and defined names', () => {
    const { second } = roundTrip('features.xlsx');
    expect(second.workbook.getSheet('HiddenSheet')!.visibility).toBe('hidden');
    expect(second.workbook.getSheet('Comments')!.tabColor).toBe('FF00B050');
    expect(second.workbook.definedNames.find((d) => d.name === 'TotalQty')?.refersTo).toBe(
      'Features!$B$2:$B$7',
    );
  });

  it('preserves cells at the far corners of the grid', () => {
    const { second } = roundTrip('edge-cases.xlsx');
    const sparse = second.workbook.getSheet('Sparse')!;
    expect(sparse.getValue(0, 0)).toBe('top-left');
    expect(sparse.getValue(1_048_575, 16_383)).toBe('bottom-right corner');
    expect(sparse.cellCount).toBe(4);
  });

  it('preserves sheet names containing quotes', () => {
    const { second } = roundTrip('edge-cases.xlsx');
    expect(second.workbook.sheets.map((s) => s.name)).toContain('Edge\'"Case');
  });

  it('is stable across a second save cycle', () => {
    const first = readXlsx(bytes('styling.xlsx'));
    const once = writeXlsx(first.workbook, {
      source: first.pkg,
      styleTables: first.styleTables,
      modified: new Date(0),
    });
    const reopened = readXlsx(once);
    const twice = writeXlsx(reopened.workbook, {
      source: reopened.pkg,
      styleTables: reopened.styleTables,
      modified: new Date(0),
    });
    expect(Array.from(twice)).toEqual(Array.from(once));
  });
});

describe('future-function prefixes', () => {
  it.each([
    ['IFS(A1>1,"a")', '_xlfn.IFS(A1>1,"a")'],
    ['FILTER(A1:A9,B1:B9)', '_xlfn._xlws.FILTER(A1:A9,B1:B9)'],
    ['SUM(A1:A9)', 'SUM(A1:A9)'],
    ['XLOOKUP(1,A:A,B:B)', '_xlfn.XLOOKUP(1,A:A,B:B)'],
    ['STDEV.S(A1:A9)', '_xlfn.STDEV.S(A1:A9)'],
    ['SUM(IFS(A1,1))', 'SUM(_xlfn.IFS(A1,1))'],
  ])('restores %s', (input, want) => {
    expect(restoreFutureFunctionPrefixes(input)).toBe(want);
  });

  it('leaves an already-prefixed name alone', () => {
    expect(restoreFutureFunctionPrefixes('_xlfn.IFS(A1,1)')).toBe('_xlfn.IFS(A1,1)');
  });

  it('does not touch a defined name that merely shares a function name', () => {
    // Without the following paren it is a name, not a call.
    expect(restoreFutureFunctionPrefixes('SORT+1')).toBe('SORT+1');
  });

  it('does not rewrite inside a string literal', () => {
    expect(restoreFutureFunctionPrefixes('IF(A1,"XLOOKUP(x)","")')).toBe('IF(A1,"XLOOKUP(x)","")');
  });

  it('handles doubled quotes inside strings', () => {
    expect(restoreFutureFunctionPrefixes('IF(A1,"say ""SORT(x)""",IFS(B1,2))')).toBe(
      'IF(A1,"say ""SORT(x)""",_xlfn.IFS(B1,2))',
    );
  });

  it('survives a full round trip through the reader and writer', () => {
    const { second } = roundTrip('formulas.calc.xlsx');
    const fx = second.workbook.getSheet('Formulas')!;
    const formulas: string[] = [];
    for (const { cell } of fx.entries()) if (cell.formula) formulas.push(cell.formula);
    // The reader strips the prefix again on the way back in.
    expect(formulas.some((f) => f.includes('_xlfn'))).toBe(false);
    expect(formulas.some((f) => f.startsWith('IFS('))).toBe(true);
  });

  it('actually writes the prefix into the file', () => {
    const { saved } = roundTrip('formulas.calc.xlsx');
    const xml = OpcPackage.read(saved).text('xl/worksheets/sheet2.xml');
    expect(xml).toContain('_xlfn.IFS(');
  });
});

describe('writing a fresh workbook', () => {
  function newWorkbook(): Workbook {
    const wb = new Workbook();
    const sheet = wb.addSheet('Data');
    sheet.setValue(0, 0, 'name');
    sheet.setValue(0, 1, 'qty');
    sheet.setValue(1, 0, 'widget');
    sheet.setValue(1, 1, 12);
    sheet.setValue(2, 0, 'gadget');
    sheet.setValue(2, 1, 7);
    sheet.setFormula(3, 1, 'SUM(B2:B3)', 19);
    sheet.setValue(4, 0, true);
    sheet.setValue(4, 1, CellError.DIV0);
    return wb;
  }

  it('produces a package with the parts Excel requires', () => {
    const saved = writeXlsx(newWorkbook(), { modified: new Date(0) });
    const pkg = OpcPackage.read(saved);
    for (const part of [
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/worksheets/sheet1.xml',
      'xl/styles.xml',
    ]) {
      expect(pkg.has(part), part).toBe(true);
    }
    expect(pkg.mainDocumentPath()).toBe('xl/workbook.xml');
    expect(pkg.flavour()).toBe('xlsx');
  });

  it('reads back everything it wrote', () => {
    const saved = writeXlsx(newWorkbook(), { modified: new Date(0) });
    const { workbook } = readXlsx(saved);
    const sheet = workbook.getSheet('Data')!;
    expect(sheet.getValue(0, 0)).toBe('name');
    expect(sheet.getValue(1, 1)).toBe(12);
    expect(sheet.getValue(4, 0)).toBe(true);
    expect((sheet.getValue(4, 1) as CellError).code).toBe('#DIV/0!');
    expect(sheet.getCell(3, 1)?.formula).toBe('SUM(B2:B3)');
    expect(sheet.getCell(3, 1)?.value).toBe(19);
  });

  it('writes a shared string table with matching counts', () => {
    const saved = writeXlsx(newWorkbook(), { modified: new Date(0) });
    const xml = OpcPackage.read(saved).text('xl/sharedStrings.xml');
    // count is total references, uniqueCount is distinct strings; Excel
    // cross-checks both.
    const count = Number(/count="(\d+)"/.exec(xml)?.[1]);
    const unique = Number(/uniqueCount="(\d+)"/.exec(xml)?.[1]);
    const siCount = (xml.match(/<si>/g) ?? []).length;
    expect(unique).toBe(siCount);
    expect(count).toBeGreaterThanOrEqual(unique);
  });

  it('can write strings inline instead', () => {
    const saved = writeXlsx(newWorkbook(), { inlineStrings: true, modified: new Date(0) });
    const pkg = OpcPackage.read(saved);
    expect(pkg.has('xl/sharedStrings.xml')).toBe(false);
    expect(pkg.text('xl/worksheets/sheet1.xml')).toContain('inlineStr');
    expect(readXlsx(saved).workbook.getSheet('Data')!.getValue(0, 0)).toBe('name');
  });

  it('preserves leading and trailing spaces in text', () => {
    const wb = new Workbook();
    wb.addSheet('S').setValue(0, 0, '  padded  ');
    const saved = writeXlsx(wb, { modified: new Date(0) });
    expect(OpcPackage.read(saved).text('xl/sharedStrings.xml')).toContain('xml:space="preserve"');
    expect(readXlsx(saved).workbook.getSheet('S')!.getValue(0, 0)).toBe('  padded  ');
  });

  it('marks a macro-enabled workbook with the vnd.ms-excel content type', () => {
    const saved = writeXlsx(newWorkbook(), { flavour: 'xlsm', modified: new Date(0) });
    const pkg = OpcPackage.read(saved);
    expect(pkg.contentType('xl/workbook.xml')).toBe(
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    );
    expect(pkg.flavour()).toBe('xlsm');
  });

  it('records the 1904 date system when the workbook uses it', () => {
    const wb = newWorkbook();
    wb.dateSystem = 1904;
    const saved = writeXlsx(wb, { modified: new Date(0) });
    expect(OpcPackage.read(saved).text('xl/workbook.xml')).toContain('date1904="1"');
    expect(readXlsx(saved).workbook.dateSystem).toBe(1904);
  });

  it('is byte-reproducible with a fixed timestamp', () => {
    const a = writeXlsx(newWorkbook(), { modified: new Date(0) });
    const b = writeXlsx(newWorkbook(), { modified: new Date(0) });
    expect(Array.from(a)).toEqual(Array.from(b));
  });
});
