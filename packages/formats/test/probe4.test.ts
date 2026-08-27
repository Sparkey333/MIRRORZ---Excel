import { describe, expect, it } from 'vitest';
import { detectDelimiter, parseDelimited, readCsv, rowsToSheet } from '../src/csv.js';
import { MAX_COLS, MAX_ROWS, Workbook } from '../../core/src/index.js';

describe('hostile input', () => {
  it('handles a huge single line without newlines', () => {
    const text = 'x'.repeat(2_000_000);
    const t0 = Date.now();
    const r = parseDelimited(text);
    console.log('one long line ms', Date.now() - t0, 'rows', r.rows.length, 'delim', JSON.stringify(r.delimiter), 'warnings', r.warnings.length);
    expect(r.rows).toHaveLength(1);
  });

  it('handles a million quotes', () => {
    const text = '"'.repeat(500_000);
    const t0 = Date.now();
    const r = parseDelimited(text, { delimiter: ',' });
    console.log('quotes ms', Date.now() - t0, 'field len', r.rows[0]![0]!.length, 'warnings', r.warnings.length);
  });

  it('handles very many columns on one row', () => {
    const text = 'a'.repeat(1) + ','.repeat(200_000);
    const t0 = Date.now();
    const r = parseDelimited(text, { delimiter: ',' });
    console.log('many cols ms', Date.now() - t0, 'fields', r.rows[0]!.length);
  });

  it('truncates beyond the sheet limits', () => {
    const rows: string[][] = new Array(MAX_ROWS + 3);
    const one = ['v'];
    for (let i = 0; i < rows.length; i++) rows[i] = one;
    const wb = new Workbook();
    const sheet = wb.addSheet('S');
    const warns: any[] = [];
    const t0 = Date.now();
    const out = rowsToSheet(rows, sheet, wb.styles, {}, (w) => warns.push(w));
    console.log('rows ms', Date.now() - t0, JSON.stringify(out), warns.map((w) => w.code));
    expect(out.rowCount).toBe(MAX_ROWS);
    expect(warns.map((w) => w.code)).toContain('sheet-limit');
  });

  it('truncates beyond the column limit', () => {
    const wide = new Array(MAX_COLS + 5).fill('v');
    const wb = new Workbook();
    const sheet = wb.addSheet('S');
    const warns: any[] = [];
    const out = rowsToSheet([wide], sheet, wb.styles, {}, (w) => warns.push(w));
    console.log(JSON.stringify(out), warns.map((w) => w.code));
    expect(out.colCount).toBe(MAX_COLS);
    expect(warns.map((w) => w.code)).toContain('sheet-limit');
  });

  it('detects a delimiter on a big file quickly', () => {
    const text = ('a,b,c,d,e\n').repeat(200_000);
    const t0 = Date.now();
    const d = detectDelimiter(text);
    console.log('detect ms', Date.now() - t0, d.delimiter, d.confident);
  });

  it('warning counting is bounded', () => {
    const src = 'a,b,c\n'.repeat(3) + '1,2\n'.repeat(50_000);
    const r = parseDelimited(src, { delimiter: ',' });
    console.log('warnings', r.warnings.length, r.warnings[r.warnings.length - 1]);
  });
});
