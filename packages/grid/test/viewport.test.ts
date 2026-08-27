import { describe, expect, it } from 'vitest';
import { MAX_COLS, MAX_ROWS, Workbook } from '@mirrorz/core';

import {
  AxisIndex,
  DEFAULT_HEADER_HEIGHT,
  DEFAULT_HEADER_WIDTH,
  Viewport,
  buildColAxis,
  buildRowAxis,
  charsToPixels,
  pointsToPixels,
} from '../src/viewport.js';

function uniform(count: number, size: number): AxisIndex {
  return new AxisIndex(count, size, []);
}

function sheet() {
  return new Workbook().addSheet('S');
}

describe('AxisIndex - uniform axis', () => {
  const axis = uniform(MAX_ROWS, 20);

  it('reports the total size without materialising the axis', () => {
    expect(axis.totalSize).toBe(MAX_ROWS * 20);
    expect(axis.exceptionCount).toBe(0);
  });

  it('computes offsets arithmetically', () => {
    expect(axis.offsetOf(0)).toBe(0);
    expect(axis.offsetOf(1)).toBe(20);
    expect(axis.offsetOf(900_000)).toBe(18_000_000);
  });

  it('defines offsetOf at the end of the axis', () => {
    expect(axis.offsetOf(MAX_ROWS)).toBe(axis.totalSize);
  });

  it('inverts the offset for an extreme scroll position', () => {
    expect(axis.indexAt(18_000_000)).toBe(900_000);
    expect(axis.indexAt(18_000_019.9)).toBe(900_000);
    expect(axis.indexAt(18_000_020)).toBe(900_001);
  });

  it('clamps below zero and past the end', () => {
    expect(axis.indexAt(-500)).toBe(0);
    expect(axis.indexAt(axis.totalSize + 10_000)).toBe(MAX_ROWS - 1);
  });

  it('round-trips index -> offset -> index across the whole range', () => {
    for (const i of [0, 1, 999, 65_535, 500_000, 1_048_575]) {
      expect(axis.indexAt(axis.offsetOf(i))).toBe(i);
      expect(axis.indexAt(axis.offsetOf(i) + 19.5)).toBe(i);
    }
  });

  it('reports the default size for every index', () => {
    expect(axis.sizeOf(0)).toBe(20);
    expect(axis.sizeOf(1_048_575)).toBe(20);
    expect(axis.sizeOf(MAX_ROWS)).toBe(0);
  });
});

describe('AxisIndex - exceptions', () => {
  const axis = new AxisIndex(1000, 20, [
    { index: 2, size: 60 },
    { index: 5, size: 0 },
    { index: 7, size: 5 },
  ]);

  it('keeps only indices that differ from the default', () => {
    expect(axis.exceptionCount).toBe(3);
    expect(new AxisIndex(10, 20, [{ index: 1, size: 20 }]).exceptionCount).toBe(0);
  });

  it('sizes exceptions and defaults correctly', () => {
    expect(axis.sizeOf(2)).toBe(60);
    expect(axis.sizeOf(5)).toBe(0);
    expect(axis.sizeOf(7)).toBe(5);
    expect(axis.sizeOf(3)).toBe(20);
  });

  it('accumulates the deltas into the offsets', () => {
    expect(axis.offsetOf(0)).toBe(0);
    expect(axis.offsetOf(2)).toBe(40);
    expect(axis.offsetOf(3)).toBe(100);
    expect(axis.offsetOf(5)).toBe(140);
    expect(axis.offsetOf(6)).toBe(140);
    expect(axis.offsetOf(8)).toBe(165);
  });

  it('adds the deltas to the total size', () => {
    expect(axis.totalSize).toBe(1000 * 20 + 40 - 20 - 15);
  });

  it('finds the index at a position inside a resized span', () => {
    expect(axis.indexAt(40)).toBe(2);
    expect(axis.indexAt(99)).toBe(2);
    expect(axis.indexAt(100)).toBe(3);
  });

  it('never returns a hidden index', () => {
    expect(axis.indexAt(140)).toBe(6);
    expect(axis.sizeOf(axis.indexAt(140))).toBeGreaterThan(0);
  });

  it('resolves positions far past the last exception arithmetically', () => {
    expect(axis.indexAt(axis.offsetOf(900) + 3)).toBe(900);
  });

  it('handles unsorted and duplicate input', () => {
    const a = new AxisIndex(10, 10, [
      { index: 5, size: 30 },
      { index: 1, size: 50 },
      { index: 5, size: 99 },
    ]);
    expect(a.sizeOf(1)).toBe(50);
    expect(a.sizeOf(5)).toBe(30);
    expect(a.offsetOf(6)).toBe(10 + 50 + 30 + 30);
  });

  it('ignores exceptions outside the axis', () => {
    expect(new AxisIndex(4, 10, [{ index: 99, size: 1 }]).totalSize).toBe(40);
  });
});

describe('AxisIndex - hidden runs', () => {
  const hidden: { index: number; size: number }[] = [];
  for (let i = 10; i < 200_000; i++) hidden.push({ index: i, size: 0 });
  const axis = new AxisIndex(MAX_ROWS, 20, hidden);

  it('collapses a huge hidden run out of the offsets', () => {
    expect(axis.offsetOf(10)).toBe(200);
    expect(axis.offsetOf(200_000)).toBe(200);
    expect(axis.offsetOf(200_001)).toBe(220);
  });

  it('skips the whole run in one step', () => {
    expect(axis.firstVisibleAtOrAfter(10)).toBe(200_000);
    expect(axis.firstVisibleAtOrAfter(199_999)).toBe(200_000);
    expect(axis.firstVisibleAtOrAfter(200_000)).toBe(200_000);
  });

  it('walks backwards over the run', () => {
    expect(axis.lastVisibleAtOrBefore(199_999)).toBe(9);
    expect(axis.lastVisibleAtOrBefore(200_000)).toBe(200_000);
  });

  it('lands after the run when the position is at its offset', () => {
    expect(axis.indexAt(200)).toBe(200_000);
  });

  it('returns count when everything from here on is hidden', () => {
    const all = new AxisIndex(5, 10, [
      { index: 3, size: 0 },
      { index: 4, size: 0 },
    ]);
    expect(all.firstVisibleAtOrAfter(3)).toBe(5);
  });

  it('handles a hidden first index without recursing', () => {
    const a = new AxisIndex(5, 10, [
      { index: 0, size: 0 },
      { index: 1, size: 0 },
    ]);
    expect(a.indexAt(0)).toBe(2);
    expect(a.firstVisibleAtOrAfter(0)).toBe(2);
    expect(a.lastVisibleAtOrBefore(1)).toBe(-1);
  });
});

describe('unit conversion', () => {
  it('converts points to pixels at 96 dpi', () => {
    expect(pointsToPixels(15)).toBe(20);
    expect(pointsToPixels(12)).toBe(16);
  });

  it('converts the default column width to Excel s 64 pixels', () => {
    expect(charsToPixels(8.43)).toBe(64);
  });

  it('treats a zero width as hidden', () => {
    expect(charsToPixels(0)).toBe(0);
  });

  it('scales with the max digit width', () => {
    expect(charsToPixels(10, 8)).toBe(85);
  });
});

describe('axes built from a sheet', () => {
  it('uses the sheet defaults', () => {
    const s = sheet();
    const rows = buildRowAxis(s);
    const cols = buildColAxis(s);
    expect(rows.sizeOf(0)).toBe(20);
    expect(cols.sizeOf(0)).toBe(64);
    expect(rows.count).toBe(MAX_ROWS);
    expect(cols.count).toBe(MAX_COLS);
  });

  it('picks up custom heights, widths and hidden flags', () => {
    const s = sheet();
    s.rows.set(3, { height: 30 });
    s.rows.set(4, { hidden: true });
    s.cols.set(1, { width: 20 });
    s.cols.set(2, { hidden: true });
    const rows = buildRowAxis(s);
    const cols = buildColAxis(s);
    expect(rows.sizeOf(3)).toBe(40);
    expect(rows.sizeOf(4)).toBe(0);
    expect(cols.sizeOf(1)).toBe(145);
    expect(cols.sizeOf(2)).toBe(0);
  });

  it('applies zoom to every size', () => {
    const s = sheet();
    expect(buildRowAxis(s, { zoom: 2 }).sizeOf(0)).toBe(40);
    expect(buildColAxis(s, { zoom: 0.5 }).sizeOf(0)).toBe(32);
  });
});

describe('Viewport layout', () => {
  function make(overrides: Record<string, unknown> = {}): Viewport {
    return new Viewport({
      rows: uniform(MAX_ROWS, 20),
      cols: uniform(MAX_COLS, 64),
      width: 1000,
      height: 600,
      ...overrides,
    });
  }

  it('places the grid after the headers by default', () => {
    const v = make();
    expect(v.gridLeft).toBe(DEFAULT_HEADER_WIDTH);
    expect(v.gridTop).toBe(DEFAULT_HEADER_HEIGHT);
    expect(v.xOf(0)).toBe(DEFAULT_HEADER_WIDTH);
    expect(v.yOf(0)).toBe(DEFAULT_HEADER_HEIGHT);
  });

  it('drops the header offset when headers are hidden', () => {
    const v = make({ showHeaders: false });
    expect(v.xOf(0)).toBe(0);
    expect(v.yOf(0)).toBe(0);
  });

  it('produces a bounded number of spans over a full-size sheet', () => {
    const v = make();
    const l = v.layout();
    expect(l.rowSpans.length).toBeLessThanOrEqual(Math.ceil(600 / 20) + 1);
    expect(l.colSpans.length).toBeLessThanOrEqual(Math.ceil(1000 / 64) + 1);
    expect(l.rowSpans.length).toBeGreaterThan(20);
  });

  it('scrolls by whole pixels, not whole rows', () => {
    const v = make();
    v.setScroll(0, 7.5);
    const l = v.layout();
    expect(l.rowSpans[0]?.index).toBe(0);
    expect(l.rowSpans[0]?.start).toBeCloseTo(DEFAULT_HEADER_HEIGHT - 7.5);
    expect(v.scrollY).toBe(7.5);
  });

  it('keeps sub-pixel scroll offsets', () => {
    const v = make();
    v.setScroll(0, 0.25);
    expect(v.scrollY).toBe(0.25);
    expect(v.yOf(0)).toBeCloseTo(DEFAULT_HEADER_HEIGHT - 0.25);
  });

  it('clamps the scroll to the content', () => {
    const v = make();
    v.setScroll(-100, -100);
    expect(v.scrollX).toBe(0);
    expect(v.scrollY).toBe(0);
    v.setScroll(1e12, 1e12);
    expect(v.scrollY).toBe(v.maxScrollY);
    expect(v.scrollY).toBeCloseTo(MAX_ROWS * 20 - (600 - DEFAULT_HEADER_HEIGHT));
  });

  it('finds the first visible row at an extreme scroll position in O(log n)', () => {
    const v = make();
    v.setScroll(0, 18_000_000);
    expect(v.firstScrollRow()).toBe(900_000);
    const l = v.layout();
    expect(l.rowSpans[0]?.index).toBe(900_000);
    expect(l.rowSpans.length).toBeLessThan(40);
  });

  it('ignores NaN scroll input', () => {
    const v = make();
    v.setScroll(Number.NaN, Number.NaN);
    expect(v.scrollX).toBe(0);
    expect(v.scrollY).toBe(0);
  });

  it('re-clamps the scroll on resize', () => {
    const v = make({ rows: uniform(50, 20), cols: uniform(50, 64) });
    v.setScroll(0, v.maxScrollY);
    const before = v.scrollY;
    v.resize(1000, 2000);
    expect(v.scrollY).toBeLessThan(before);
    expect(v.scrollY).toBe(0);
  });
});

describe('Viewport frozen panes', () => {
  function frozen(): Viewport {
    return new Viewport({
      rows: uniform(MAX_ROWS, 20),
      cols: uniform(MAX_COLS, 64),
      width: 800,
      height: 400,
      frozenRows: 2,
      frozenCols: 1,
    });
  }

  it('measures the frozen bands', () => {
    const v = frozen();
    expect(v.frozenWidth).toBe(64);
    expect(v.frozenHeight).toBe(40);
  });

  it('emits four panes', () => {
    const l = frozen().layout();
    expect(l.panes.map((p) => p.id)).toEqual(['topLeft', 'topRight', 'bottomLeft', 'bottomRight']);
  });

  it('keeps frozen rows and columns in place while scrolling', () => {
    const v = frozen();
    v.setScroll(500, 4000);
    expect(v.xOf(0)).toBe(DEFAULT_HEADER_WIDTH);
    expect(v.yOf(0)).toBe(DEFAULT_HEADER_HEIGHT);
    expect(v.yOf(1)).toBe(DEFAULT_HEADER_HEIGHT + 20);
  });

  it('offsets the scrolling pane past the frozen band', () => {
    const v = frozen();
    v.setScroll(0, 0);
    expect(v.xOf(1)).toBe(DEFAULT_HEADER_WIDTH + 64);
    expect(v.yOf(2)).toBe(DEFAULT_HEADER_HEIGHT + 40);
  });

  it('scrolls the scrolling pane only', () => {
    const v = frozen();
    v.setScroll(64, 40);
    expect(v.xOf(1)).toBe(DEFAULT_HEADER_WIDTH + 64 - 64);
    expect(v.layout().panes.find((p) => p.id === 'bottomRight')?.cols[0]?.index).toBe(2);
  });

  it('confines each pane to its own rectangle', () => {
    const l = frozen().layout();
    const topLeft = l.panes.find((p) => p.id === 'topLeft');
    const bottomRight = l.panes.find((p) => p.id === 'bottomRight');
    expect(topLeft?.rect).toEqual({
      x: DEFAULT_HEADER_WIDTH,
      y: DEFAULT_HEADER_HEIGHT,
      width: 64,
      height: 40,
    });
    expect(bottomRight?.rect.x).toBe(DEFAULT_HEADER_WIDTH + 64);
    expect(bottomRight?.rect.y).toBe(DEFAULT_HEADER_HEIGHT + 40);
  });

  it('puts frozen spans before scrolling ones in the header lists', () => {
    const v = frozen();
    v.setScroll(1000, 1000);
    const l = v.layout();
    expect(l.colSpans[0]?.index).toBe(0);
    expect(l.colSpans[1]?.index).toBeGreaterThan(1);
    expect(l.rowSpans[0]?.index).toBe(0);
    expect(l.rowSpans[1]?.index).toBe(1);
  });

  it('stops scrolling with the last row at the bottom edge', () => {
    const v = frozen();
    v.setScroll(0, v.maxScrollY);
    expect(v.yOf(MAX_ROWS - 1) + 20).toBeCloseTo(400);
  });

  it('excludes the frozen band from the scrollable content', () => {
    const v = frozen();
    // The frozen rows never scroll, so they are not part of what can scroll
    // past: the scrollable extent is the content minus the scrolling band.
    expect(v.maxScrollY).toBeCloseTo(MAX_ROWS * 20 - 40 - v.scrollViewHeight);
    expect(v.scrollViewHeight).toBe(400 - DEFAULT_HEADER_HEIGHT - 40);
  });

  it('never lets the frozen band exceed the viewport', () => {
    const v = new Viewport({
      rows: uniform(MAX_ROWS, 20),
      cols: uniform(MAX_COLS, 64),
      width: 200,
      height: 100,
      frozenCols: 40,
      frozenRows: 40,
    });
    expect(v.frozenWidth).toBeLessThanOrEqual(200);
    expect(v.frozenHeight).toBeLessThanOrEqual(100);
  });

  it('drops empty panes', () => {
    const v = new Viewport({
      rows: uniform(MAX_ROWS, 20),
      cols: uniform(MAX_COLS, 64),
      width: 800,
      height: 400,
    });
    expect(v.layout().panes.map((p) => p.id)).toEqual(['bottomRight']);
  });
});

describe('Viewport hit testing', () => {
  function make(): Viewport {
    return new Viewport({
      rows: uniform(MAX_ROWS, 20),
      cols: uniform(MAX_COLS, 64),
      width: 800,
      height: 400,
    });
  }

  it('finds a cell', () => {
    const v = make();
    expect(v.hitTest(DEFAULT_HEADER_WIDTH + 10, DEFAULT_HEADER_HEIGHT + 10)).toEqual({
      kind: 'cell',
      row: 0,
      col: 0,
    });
    expect(v.hitTest(DEFAULT_HEADER_WIDTH + 70, DEFAULT_HEADER_HEIGHT + 25)).toEqual({
      kind: 'cell',
      row: 1,
      col: 1,
    });
  });

  it('finds a cell after scrolling', () => {
    const v = make();
    v.setScroll(0, 18_000_000);
    const hit = v.hitTest(DEFAULT_HEADER_WIDTH + 10, DEFAULT_HEADER_HEIGHT + 5);
    expect(hit).toEqual({ kind: 'cell', row: 900_000, col: 0 });
  });

  it('finds the corner box', () => {
    expect(make().hitTest(2, 2)).toEqual({ kind: 'corner' });
  });

  it('finds a column header', () => {
    expect(make().hitTest(DEFAULT_HEADER_WIDTH + 30, 5)).toEqual({ kind: 'col-header', col: 0 });
  });

  it('finds a row header', () => {
    expect(make().hitTest(10, DEFAULT_HEADER_HEIGHT + 30)).toEqual({ kind: 'row-header', row: 1 });
  });

  it('finds a column resize handle at the edge', () => {
    expect(make().hitTest(DEFAULT_HEADER_WIDTH + 64, 5)).toEqual({ kind: 'col-resize', col: 0 });
  });

  it('maps the left edge of a column to the previous column s handle', () => {
    expect(make().hitTest(DEFAULT_HEADER_WIDTH + 65, 5)).toEqual({ kind: 'col-resize', col: 0 });
  });

  it('finds a row resize handle', () => {
    expect(make().hitTest(10, DEFAULT_HEADER_HEIGHT + 20)).toEqual({ kind: 'row-resize', row: 0 });
  });

  it('reports points outside the canvas', () => {
    expect(make().hitTest(-1, 10)).toEqual({ kind: 'outside' });
    expect(make().hitTest(10, 100_000)).toEqual({ kind: 'outside' });
  });

  it('respects frozen panes when hit testing', () => {
    const v = new Viewport({
      rows: uniform(MAX_ROWS, 20),
      cols: uniform(MAX_COLS, 64),
      width: 800,
      height: 400,
      frozenRows: 2,
      frozenCols: 1,
    });
    v.setScroll(6400, 4000);
    expect(v.hitTest(DEFAULT_HEADER_WIDTH + 10, DEFAULT_HEADER_HEIGHT + 10)).toEqual({
      kind: 'cell',
      row: 0,
      col: 0,
    });
    const hit = v.hitTest(DEFAULT_HEADER_WIDTH + 64 + 10, DEFAULT_HEADER_HEIGHT + 40 + 10);
    expect(hit).toEqual({ kind: 'cell', row: 202, col: 101 });
  });

  it('skips hidden columns when hit testing', () => {
    const v = new Viewport({
      rows: uniform(100, 20),
      cols: new AxisIndex(100, 64, [{ index: 1, size: 0 }]),
      width: 800,
      height: 400,
    });
    expect(v.hitTest(DEFAULT_HEADER_WIDTH + 70, DEFAULT_HEADER_HEIGHT + 5)).toEqual({
      kind: 'cell',
      row: 0,
      col: 2,
    });
  });
});

describe('Viewport reveal and paging', () => {
  function make(): Viewport {
    return new Viewport({
      rows: uniform(MAX_ROWS, 20),
      cols: uniform(MAX_COLS, 64),
      width: 800,
      height: 400,
    });
  }

  it('does nothing when the cell is already visible', () => {
    const v = make();
    v.revealCell(2, 2);
    expect(v.scrollY).toBe(0);
    expect(v.scrollX).toBe(0);
  });

  it('scrolls down by the minimum needed', () => {
    const v = make();
    const visibleRows = Math.floor((400 - DEFAULT_HEADER_HEIGHT) / 20);
    v.revealCell(visibleRows, 0);
    expect(v.scrollY).toBeCloseTo((visibleRows + 1) * 20 - (400 - DEFAULT_HEADER_HEIGHT));
  });

  it('scrolls back up to a cell above the viewport', () => {
    const v = make();
    v.setScroll(0, 10_000);
    v.revealCell(3, 0);
    expect(v.scrollY).toBe(60);
  });

  it('reveals a far cell in one step', () => {
    const v = make();
    v.revealCell(1_000_000, 0);
    expect(v.isFullyVisible(1_000_000, 0)).toBe(true);
  });

  it('never scrolls for a frozen cell', () => {
    const v = new Viewport({
      rows: uniform(MAX_ROWS, 20),
      cols: uniform(MAX_COLS, 64),
      width: 800,
      height: 400,
      frozenRows: 2,
    });
    v.setScroll(0, 5000);
    const before = v.scrollY;
    v.revealCell(0, 0);
    expect(v.scrollY).toBe(before);
  });

  it('reports whole rows per page', () => {
    const v = make();
    expect(v.pageRows()).toBe(Math.floor((400 - DEFAULT_HEADER_HEIGHT) / 20));
  });

  it('reports whole columns per page', () => {
    const v = make();
    expect(v.pageCols()).toBe(Math.floor((800 - DEFAULT_HEADER_WIDTH) / 64));
  });

  it('never reports a page of zero', () => {
    const v = new Viewport({ rows: uniform(100, 500), cols: uniform(100, 500), width: 40, height: 40 });
    expect(v.pageRows()).toBe(1);
    expect(v.pageCols()).toBe(1);
  });
});
