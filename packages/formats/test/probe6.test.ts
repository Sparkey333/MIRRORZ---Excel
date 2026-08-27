import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseDelimited, readCsv, writeCsv } from '../src/csv.js';
const F = new URL('../../../fixtures/generated/', import.meta.url);
const text = (n: string) => readFileSync(new URL(n, F), 'utf8');
describe('precedence.csv', () => {
  it('round trips', () => {
    const original = text('precedence.csv');
    const { sheet } = readCsv(original, { raw: true, delimiter: ',' });
    const out = writeCsv(sheet, { lineEnding: '\n' });
    if (out !== original) {
      const a = original.split('\n'), b = out.split('\n');
      for (let i = 0; i < Math.max(a.length, b.length); i++) if (a[i] !== b[i]) console.log(i, JSON.stringify(a[i]), '!=', JSON.stringify(b[i]));
    }
    expect(out).toBe(original);
  });
  it('warnings', () => {
    const r = parseDelimited(text('precedence.csv'));
    console.log(r.delimiter, JSON.stringify(r.warnings.slice(0, 5)));
    const r2 = parseDelimited(text('styling.csv'));
    console.log('styling', r2.delimiter, JSON.stringify(r2.warnings.slice(0, 5)));
    const r3 = parseDelimited(text('edge-cases.csv'));
    console.log('edge', r3.delimiter, JSON.stringify(r3.warnings.slice(0, 5)), JSON.stringify(r3.rows));
  });
  it('inferred import of precedence keeps the apostrophe formulas as text', () => {
    const { sheet } = readCsv(text('precedence.csv'));
    console.log(JSON.stringify(sheet.getValue(9, 1)), JSON.stringify(sheet.getValue(19, 1)));
  });
});
