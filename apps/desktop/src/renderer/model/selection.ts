/**
 * Selection model and the statistics the status bar shows.
 *
 * The status-bar aggregate is one of the most used features in any spreadsheet -
 * select a column, glance at the sum - and it has one rule that is easy to get
 * wrong: it counts what Excel counts. Text and blanks are excluded from the
 * numeric aggregates but included in the count; booleans are excluded from both,
 * because a column of TRUE/FALSE summing to 4 would be nonsense; and errors
 * poison the aggregate rather than being skipped, since a sum that quietly
 * ignores a `#REF!` is a wrong number presented as a right one.
 */

import { isError, type Scalar, type Sheet } from '@mirrorz/core';

export interface CellPos {
  row: number;
  col: number;
}

export interface SelectionRange {
  start: CellPos;
  end: CellPos;
}

export interface Selection {
  sheet: string;
  /** Where the keyboard focus is; always inside `ranges[0]`. */
  active: CellPos;
  /** Non-empty; multiple ranges for a ctrl-click selection. */
  ranges: SelectionRange[];
}

export function normaliseRange(range: SelectionRange): SelectionRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      col: Math.min(range.start.col, range.end.col),
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      col: Math.max(range.start.col, range.end.col),
    },
  };
}

export function rangeArea(range: SelectionRange): number {
  const n = normaliseRange(range);
  return (n.end.row - n.start.row + 1) * (n.end.col - n.start.col + 1);
}

export function singleCell(sheet: string, row: number, col: number): Selection {
  return { sheet, active: { row, col }, ranges: [{ start: { row, col }, end: { row, col } }] };
}

export function containsCell(selection: Selection, row: number, col: number): boolean {
  return selection.ranges.some((r) => {
    const n = normaliseRange(r);
    return row >= n.start.row && row <= n.end.row && col >= n.start.col && col <= n.end.col;
  });
}

export interface SelectionStats {
  /** Cells holding anything at all. */
  count: number;
  /** Cells holding a number, which is what the aggregates are over. */
  numericCount: number;
  sum: number;
  average: number | null;
  min: number | null;
  max: number | null;
  /** Set when any selected cell holds an error, which invalidates the aggregates. */
  errorCode?: string;
  /** Total addresses selected, including empty ones. */
  selectedCells: number;
}

/**
 * Aggregate the selection.
 *
 * Iteration is over the sheet's populated cells inside each range rather than
 * over every address, because selecting a whole column selects a million
 * addresses and twelve values.
 */
export function selectionStats(sheet: Sheet | undefined, selection: Selection): SelectionStats {
  const stats: SelectionStats = {
    count: 0,
    numericCount: 0,
    sum: 0,
    average: null,
    min: null,
    max: null,
    selectedCells: 0,
  };
  if (!sheet) return stats;

  const seen = new Set<number>();
  for (const raw of selection.ranges) {
    const range = normaliseRange(raw);
    stats.selectedCells += rangeArea(range);
    for (const { row, col, cell } of sheet.entriesIn({
      start: { row: range.start.row, col: range.start.col, rowAbs: false, colAbs: false },
      end: { row: range.end.row, col: range.end.col, rowAbs: false, colAbs: false },
    })) {
      // Overlapping ranges in a multi-selection must not double-count.
      const key = row * 16_384 + col;
      if (seen.has(key)) continue;
      seen.add(key);
      accumulate(stats, cell.value);
    }
  }

  if (stats.numericCount > 0) stats.average = stats.sum / stats.numericCount;
  return stats;
}

function accumulate(stats: SelectionStats, value: Scalar): void {
  if (value === null) return;
  stats.count++;
  if (isError(value)) {
    stats.errorCode ??= value.code;
    return;
  }
  if (typeof value !== 'number') return;
  stats.numericCount++;
  stats.sum += value;
  stats.min = stats.min === null ? value : Math.min(stats.min, value);
  stats.max = stats.max === null ? value : Math.max(stats.max, value);
}

/** Format an aggregate for the status bar, which has no room for 17 digits. */
export function formatStat(value: number | null): string {
  if (value === null) return '';
  if (!Number.isFinite(value)) return String(value);
  if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value);
  const rounded = Number(value.toPrecision(10));
  return String(rounded);
}

/** The A1 description of the selection, e.g. `B2:D9` or `3 ranges`. */
export function describeSelection(selection: Selection, a1: (r: number, c: number) => string): string {
  if (selection.ranges.length > 1) return `${selection.ranges.length} ranges`;
  const range = normaliseRange(selection.ranges[0]!);
  const start = a1(range.start.row, range.start.col);
  const end = a1(range.end.row, range.end.col);
  return start === end ? start : `${start}:${end}`;
}
