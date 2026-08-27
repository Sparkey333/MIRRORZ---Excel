import { describe, expect, it } from 'vitest';
import { CellError, Sheet, a1 } from '@mirrorz/core';
import {
  containsCell,
  describeSelection,
  formatStat,
  normaliseRange,
  rangeArea,
  selectionStats,
  singleCell,
  type Selection,
} from '../src/renderer/model/selection.js';

function sheetWith(values: (number | string | boolean | CellError | null)[][]): Sheet {
  const sheet = new Sheet('Sheet1', 1);
  values.forEach((row, r) => row.forEach((value, c) => sheet.setValue(r, c, value)));
  return sheet;
}

const range = (r1: number, c1: number, r2: number, c2: number): Selection => ({
  sheet: 'Sheet1',
  active: { row: r1, col: c1 },
  ranges: [{ start: { row: r1, col: c1 }, end: { row: r2, col: c2 } }],
});

describe('range helpers', () => {
  it('normalises a range dragged upwards', () => {
    expect(normaliseRange({ start: { row: 5, col: 5 }, end: { row: 1, col: 2 } })).toEqual({
      start: { row: 1, col: 2 },
      end: { row: 5, col: 5 },
    });
  });

  it('computes the area', () => {
    expect(rangeArea({ start: { row: 0, col: 0 }, end: { row: 2, col: 3 } })).toBe(12);
  });

  it('reports containment', () => {
    const selection = range(1, 1, 3, 3);
    expect(containsCell(selection, 2, 2)).toBe(true);
    expect(containsCell(selection, 0, 0)).toBe(false);
  });

  it('makes a single-cell selection', () => {
    expect(singleCell('S', 2, 3)).toEqual({
      sheet: 'S',
      active: { row: 2, col: 3 },
      ranges: [{ start: { row: 2, col: 3 }, end: { row: 2, col: 3 } }],
    });
  });
});

describe('selectionStats', () => {
  const sheet = sheetWith([
    [1, 2, 'text'],
    [3, 4, true],
    [null, 10, null],
  ]);

  it('returns zeroes with no sheet', () => {
    expect(selectionStats(undefined, range(0, 0, 1, 1)).count).toBe(0);
  });

  it('counts everything that holds a value', () => {
    expect(selectionStats(sheet, range(0, 0, 2, 2)).count).toBe(7);
  });

  it('aggregates only the numbers', () => {
    const stats = selectionStats(sheet, range(0, 0, 2, 2));
    expect(stats.numericCount).toBe(5);
    expect(stats.sum).toBe(20);
    expect(stats.average).toBe(4);
    expect(stats.min).toBe(1);
    expect(stats.max).toBe(10);
  });

  it('excludes booleans from the aggregate, as Excel does', () => {
    const booleans = sheetWith([[true, true, true]]);
    const stats = selectionStats(booleans, range(0, 0, 0, 2));
    expect(stats.count).toBe(3);
    expect(stats.numericCount).toBe(0);
    expect(stats.average).toBeNull();
  });

  it('reports an error in the selection rather than quietly summing round it', () => {
    const errored = sheetWith([[1, CellError.REF, 3]]);
    const stats = selectionStats(errored, range(0, 0, 0, 2));
    expect(stats.errorCode).toBe('#REF!');
    expect(stats.count).toBe(3);
  });

  it('counts empty addresses separately from populated ones', () => {
    const stats = selectionStats(sheet, range(0, 0, 2, 2));
    expect(stats.selectedCells).toBe(9);
    expect(stats.count).toBe(7);
  });

  it('does not double-count overlapping ranges', () => {
    const selection: Selection = {
      sheet: 'Sheet1',
      active: { row: 0, col: 0 },
      ranges: [
        { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
        { start: { row: 1, col: 1 }, end: { row: 2, col: 2 } },
      ],
    };
    expect(selectionStats(sheet, selection).sum).toBe(1 + 2 + 3 + 4 + 10);
  });

  it('handles a whole-column selection over a sparse sheet', () => {
    const sparse = new Sheet('Sheet1', 1);
    sparse.setValue(0, 0, 5);
    sparse.setValue(999_999, 0, 5);
    const whole: Selection = {
      sheet: 'Sheet1',
      active: { row: 0, col: 0 },
      ranges: [{ start: { row: 0, col: 0 }, end: { row: 1_048_575, col: 0 } }],
    };
    expect(selectionStats(sparse, whole).sum).toBe(10);
  });
});

describe('formatStat', () => {
  it('renders nothing for a missing statistic', () => {
    expect(formatStat(null)).toBe('');
  });

  it('renders an integer exactly', () => {
    expect(formatStat(1234)).toBe('1234');
  });

  it('trims floating-point noise', () => {
    expect(formatStat(0.1 + 0.2)).toBe('0.3');
  });

  it('passes infinities through', () => {
    expect(formatStat(Number.POSITIVE_INFINITY)).toBe('Infinity');
  });
});

describe('describeSelection', () => {
  it('names a single cell', () => {
    expect(describeSelection(singleCell('S', 0, 0), a1)).toBe('A1');
  });

  it('names a rectangle', () => {
    expect(describeSelection(range(0, 0, 4, 2), a1)).toBe('A1:C5');
  });

  it('counts multiple ranges rather than listing them', () => {
    const selection: Selection = {
      sheet: 'S',
      active: { row: 0, col: 0 },
      ranges: [
        { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
        { start: { row: 2, col: 2 }, end: { row: 3, col: 3 } },
      ],
    };
    expect(describeSelection(selection, a1)).toBe('2 ranges');
  });
});
