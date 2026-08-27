import { describe, expect, it } from 'vitest';
import {
  buildImportPlan,
  columnLetter,
  optionsForOverride,
  resolveImport,
} from '../src/renderer/model/import-review.js';

/**
 * The rows that motivate the whole dialogue: a gene symbol, a leading-zero part
 * number, a long order id and a genuine date column.
 */
const rows = [
  ['gene', 'code', 'order', 'when'],
  ['SEPT1', '007', '1234567890123456789', '2024-01-31'],
  ['MARCH2', '0042', '1234567890123456790', '2024-02-01'],
  ['BRCA1', '123', '5', '2024-03-02'],
];

describe('buildImportPlan', () => {
  it('reports a column per source column', () => {
    const plan = buildImportPlan(rows, {}, [], true);
    expect(plan.columns.map((c) => c.name)).toEqual(['gene', 'code', 'order', 'when']);
  });

  it('falls back to column letters without a header row', () => {
    const plan = buildImportPlan(rows, {}, [], false);
    expect(plan.columns.map((c) => c.name)).toEqual(['A', 'B', 'C', 'D']);
  });

  it('protects gene symbols from becoming dates', () => {
    const plan = buildImportPlan(rows, {}, [], true);
    expect(plan.columns[0]!.protectedCount).toBe(2);
    expect(plan.columns[0]!.reasons.join(' ')).toMatch(/gene symbol/);
  });

  it('protects leading zeros', () => {
    const plan = buildImportPlan(rows, {}, [], true);
    expect(plan.columns[1]!.reasons.join(' ')).toMatch(/leading zero/);
  });

  it('protects long digit strings from losing precision', () => {
    const plan = buildImportPlan(rows, {}, [], true);
    expect(plan.columns[2]!.reasons.join(' ')).toMatch(/15 digits/);
  });

  it('reads a real date column as dates', () => {
    const plan = buildImportPlan(rows, {}, [], true);
    expect(plan.columns[3]!.kinds.date).toBe(3);
    expect(plan.columns[3]!.dominant).toBe('date');
  });

  it('states the outcome in one sentence', () => {
    const plan = buildImportPlan(rows, {}, [], true);
    expect(plan.headline).toMatch(/3 cells would be read as dates/);
    expect(plan.headline).toMatch(/kept as text/);
    expect(plan.headline).toMatch(/gene symbols/);
    expect(plan.headline).toMatch(/leading zeros/);
  });

  it('says nothing was changed when nothing would be', () => {
    const plan = buildImportPlan([['alpha', 'beta']], {}, [], false);
    expect(plan.headline).toMatch(/all kept exactly as supplied/);
  });

  it('collects samples of the conversion', () => {
    const plan = buildImportPlan(rows, {}, [], true);
    expect(plan.columns[3]!.samples[0]).toEqual({ input: '2024-01-31', result: '45322' });
  });

  it('honours a text override, so a date column can be kept as typed', () => {
    const plan = buildImportPlan(rows, {}, ['auto', 'auto', 'auto', 'text'], true);
    expect(plan.columns[3]!.kinds.date).toBe(0);
    expect(plan.columns[3]!.kinds.text).toBe(3);
  });

  it('honours a number override, which stops dates being inferred', () => {
    const plan = buildImportPlan([['1/2/2024']], {}, ['number'], false);
    expect(plan.columns[0]!.kinds.date).toBe(0);
  });

  it('flags an ambiguous date rather than resolving it silently', () => {
    const plan = buildImportPlan([['3/4/2024']], {}, [], false);
    expect(plan.columns[0]!.ambiguous).toBe(1);
    expect(plan.review.ambiguous).toHaveLength(1);
  });

  it('does not flag a date whose order is unambiguous', () => {
    const plan = buildImportPlan([['13/4/2024']], { dateOrder: 'dmy' }, [], false);
    expect(plan.columns[0]!.ambiguous).toBe(0);
  });

  it('ignores empty cells in the totals', () => {
    const plan = buildImportPlan([['', '1'], ['', '2']], {}, [], false);
    expect(plan.review.total).toBe(2);
  });

  it('excludes the header row from the counts', () => {
    const withHeader = buildImportPlan([['n'], ['1'], ['2']], {}, [], true);
    expect(withHeader.review.total).toBe(2);
  });

  it('handles ragged rows without losing the wider ones', () => {
    const plan = buildImportPlan([['a'], ['b', 'c', 'd']], {}, [], false);
    expect(plan.columns).toHaveLength(3);
  });
});

describe('optionsForOverride', () => {
  it('locks a column to text', () => {
    expect(optionsForOverride({}, 'text').forceText).toBe(true);
  });

  it('turns off date inference for a number column', () => {
    expect(optionsForOverride({}, 'number').inferDates).toBe(false);
  });

  it('turns off number inference for a date column', () => {
    expect(optionsForOverride({}, 'date').inferNumbers).toBe(false);
  });

  it('passes the base options through unchanged for auto', () => {
    const base = { dateOrder: 'dmy' as const };
    expect(optionsForOverride(base, 'auto')).toBe(base);
  });
});

describe('resolveImport', () => {
  it('produces one entry per non-empty cell, positioned relatively', () => {
    const cells = resolveImport([['a', ''], ['', 'b']], {}, [], false);
    expect(cells).toEqual([
      { value: 'a', row: 0, col: 0 },
      { value: 'b', row: 1, col: 1 },
    ]);
  });

  it('keeps the literal text when the stored value differs', () => {
    const cells = resolveImport([['2024-01-31']], {}, [], false);
    expect(cells[0]!.literal).toBe('2024-01-31');
    expect(cells[0]!.value).toBe(45322);
  });

  it('applies the per-column override', () => {
    const cells = resolveImport([['2024-01-31']], {}, ['text'], false);
    expect(cells[0]!.value).toBe('2024-01-31');
  });

  it('skips the header row when there is one', () => {
    const cells = resolveImport([['h'], ['1']], {}, [], true);
    expect(cells).toEqual([{ value: 1, row: 0, col: 0 }]);
  });
});

describe('columnLetter', () => {
  it('counts in the spreadsheet way', () => {
    expect([0, 25, 26, 27].map(columnLetter)).toEqual(['A', 'Z', 'AA', 'AB']);
  });
});
