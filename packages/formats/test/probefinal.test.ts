import { describe, expect, it } from 'vitest';
import { CsvRowParser, parseDelimited, writeRows, inferValue, formatScalar } from '../src/csv.js';

function reference(text: string, delim = ','): string[][] {
  const rows: string[][] = []; let row: string[] = []; let i = 0; const n = text.length;
  if (n === 0) return rows;
  for (;;) {
    let field = '';
    if (text[i] === '"') {
      i++;
      while (i < n) {
        const ch = text[i]!;
        if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } i++; break; }
        if (ch === '\r') { field += '\n'; if (text[i + 1] === '\n') i += 2; else i++; continue; }
        field += ch; i++;
      }
      while (i < n && text[i] !== delim && text[i] !== '\r' && text[i] !== '\n') { field += text[i]!; i++; }
    } else {
      while (i < n && text[i] !== delim && text[i] !== '\r' && text[i] !== '\n') { field += text[i]!; i++; }
    }
    row.push(field);
    if (i >= n) { rows.push(row); break; }
    if (text[i] === delim) { i++; continue; }
    if (text[i] === '\r' && text[i + 1] === '\n') i += 2; else i++;
    rows.push(row); row = [];
    if (i >= n) break;
  }
  return rows;
}
let seed = 24680;
const rnd = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
const pick = <T>(a: readonly T[]): T => a[Math.floor(rnd() * a.length)]!;
const parse = (text: string, delim: string, sizes: number[]) => {
  const rows: string[][] = [];
  const p = new CsvRowParser(delim, { row: (f) => void rows.push(f) });
  let i = 0, k = 0;
  while (i < text.length) { const n = sizes[k++ % sizes.length]!; p.push(text.slice(i, i + n)); i += n; }
  p.end();
  return rows;
};

describe('final fuzz', () => {
  it('matches an independent reference and is chunk-invariant', () => {
    const alphabet = ['a', ',', ';', '"', '\r', '\n', ' ', '\t'];
    for (let iter = 0; iter < 30000; iter++) {
      const len = 1 + Math.floor(rnd() * 14);
      let s = '';
      for (let k = 0; k < len; k++) s += pick(alphabet);
      const delim = pick([',', ';', '\t']);
      const whole = JSON.stringify(parse(s, delim, [1e9]));
      expect(whole, JSON.stringify(s)).toBe(JSON.stringify(reference(s, delim)));
      for (const sizes of [[1], [2], [3, 1, 4]]) {
        expect(JSON.stringify(parse(s, delim, sizes)), JSON.stringify(s)).toBe(whole);
      }
    }
  });

  it('write then read is now a total identity on field matrices', () => {
    const chars = ['a', ',', '"', '\n', '\r', ' ', ';', '\t', ''];
    for (let iter = 0; iter < 6000; iter++) {
      const rows: string[][] = [];
      const nrows = 1 + Math.floor(rnd() * 4);
      const ncols = 1 + Math.floor(rnd() * 3);
      for (let r = 0; r < nrows; r++) {
        const row: string[] = [];
        for (let c = 0; c < ncols; c++) {
          let f = ''; const len = Math.floor(rnd() * 4);
          for (let k = 0; k < len; k++) f += pick(chars);
          row.push(f);
        }
        rows.push(row);
      }
      const delim = pick([',', ';', '\t']);
      const eol = pick(['\r\n', '\n', '\r'] as const);
      const text = writeRows(rows, { delimiter: delim, lineEnding: eol });
      const back = parseDelimited(text, { delimiter: delim }).rows;
      const expected = rows.map((r) => r.map((f) => f.replace(/\r\n|\r/g, '\n')));
      expect(back, `${JSON.stringify(expected)} wrote ${JSON.stringify(text)}`).toEqual(expected);
    }
  });

  it('numbers survive write then read', () => {
    for (let iter = 0; iter < 20000; iter++) {
      const mantissa = (rnd() - 0.5) * 2;
      const exp = Math.floor(rnd() * 60) - 30;
      const v = mantissa * Math.pow(10, exp);
      if (!Number.isFinite(v) || v === 0) continue;
      const written = formatScalar(v);
      const read = inferValue(written).value;
      expect(typeof read, `${v} wrote ${written}`).toBe('number');
      expect(formatScalar(read as number), `${v} wrote ${written}`).toBe(written);
    }
  });
});
