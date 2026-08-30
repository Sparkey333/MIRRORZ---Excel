import { beforeEach, describe, expect, it, vi } from 'vitest';
import { MAX_ROWS, Workbook } from '@mirrorz/core';

import { DARK_THEME, GridView, type GridRenderingContext } from '../src/index.js';
import { FakeCanvas } from './fake-canvas.js';

/**
 * A real CanvasRenderingContext2D must satisfy the structural context the
 * painter asks for; if it stops doing so, this stops compiling.
 */
const _assignable: GridRenderingContext = null as unknown as CanvasRenderingContext2D;
void _assignable;

function build(options: Record<string, unknown> = {}) {
  const wb = new Workbook();
  const sheet = wb.addSheet('Sheet1');
  sheet.setValue(0, 0, 'A1');
  sheet.setValue(4, 4, 42);
  const canvas = new FakeCanvas();
  const view = new GridView(canvas, wb, 'Sheet1', { width: 800, height: 400, ...options });
  return { wb, sheet, canvas, view };
}

describe('GridView construction', () => {
  it('throws for a sheet that does not exist', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    expect(() => new GridView(new FakeCanvas(), wb, 'Nope')).toThrow(/no such sheet/);
  });

  it('sizes the backing store in device pixels and the box in CSS pixels', () => {
    const { canvas } = build({ dpr: 2 });
    expect(canvas.width).toBe(1600);
    expect(canvas.height).toBe(800);
    expect(canvas.style.width).toBe('800px');
    expect(canvas.style.height).toBe('400px');
  });

  it('adopts the sheet view s frozen panes', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('S');
    sheet.view.frozenRows = 2;
    sheet.view.frozenCols = 1;
    const view = new GridView(new FakeCanvas(), wb, 'S', { width: 400, height: 200 });
    expect(view.layout().panes.map((p) => p.id)).toContain('topLeft');
  });

  it('adopts the sheet view s active cell', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('S');
    sheet.view.activeCell = { row: 3, col: 2 };
    const view = new GridView(new FakeCanvas(), wb, 'S', { width: 400, height: 200 });
    expect(view.selection.active).toEqual({ row: 3, col: 2 });
  });

  it('starts at A1 with no stored view', () => {
    expect(build().view.selection.active).toEqual({ row: 0, col: 0 });
  });
});

describe('GridView scrolling', () => {
  let h: ReturnType<typeof build>;
  beforeEach(() => {
    h = build();
  });

  it('clamps and reports scroll changes', () => {
    const spy = vi.fn();
    h.view.on('scroll', spy);
    h.view.setScroll(0, 500);
    expect(h.view.scrollY).toBe(500);
    expect(spy).toHaveBeenCalledWith({ x: 0, y: 500 });
  });

  it('does not fire when the scroll does not move', () => {
    const spy = vi.fn();
    h.view.on('scroll', spy);
    h.view.setScroll(0, 0);
    expect(spy).not.toHaveBeenCalled();
  });

  it('scrolls relatively', () => {
    h.view.setScroll(0, 100);
    h.view.scrollBy(0, -40);
    expect(h.view.scrollY).toBe(60);
  });

  it('scrolls to a far cell in one step', () => {
    h.view.scrollToCell(900_000, 0);
    const layout = h.view.layout();
    expect(layout.rowSpans.some((s) => s.index === 900_000)).toBe(true);
  });

  it('re-clamps the scroll after a resize', () => {
    h.view.setScroll(0, h.view.maxScrollY);
    h.view.resize(800, 4000);
    expect(h.view.scrollY).toBeLessThanOrEqual(h.view.maxScrollY);
  });

  it('exposes the scrollable extent', () => {
    expect(h.view.maxScrollY).toBeCloseTo(MAX_ROWS * 20 - (400 - 20));
  });
});

describe('GridView rendering', () => {
  it('draws and reports statistics', () => {
    const h = build();
    const stats = h.view.render();
    expect(stats.cellsVisited).toBeGreaterThan(0);
    expect(h.canvas.context.ops('fillText').length).toBeGreaterThan(0);
  });

  it('emits a render event with the statistics', () => {
    const h = build();
    const spy = vi.fn();
    h.view.on('render', spy);
    const stats = h.view.render();
    expect(spy).toHaveBeenCalledWith(stats);
  });

  it('draws the cell values of its sheet', () => {
    const h = build();
    h.view.render();
    expect(h.canvas.context.texts()).toContain('A1');
    expect(h.canvas.context.texts()).toContain('42');
  });

  it('repaints in the new theme after a theme change', () => {
    const h = build();
    h.view.setTheme(DARK_THEME);
    h.canvas.context.reset();
    h.view.render();
    expect(h.canvas.context.ops('fillRect')[0]?.fillStyle).toBe(DARK_THEME.background);
  });

  it('clears cached measurements when the theme changes', () => {
    const h = build();
    h.view.render();
    h.view.setTheme(DARK_THEME);
    expect(h.view.measure.stats.size).toBe(0);
  });

  it('never writes to the workbook', () => {
    const h = build();
    const before = h.wb.totalCells;
    h.view.render();
    h.view.keyDown({ key: 'ArrowDown' });
    h.view.pointerDown(200, 100);
    h.view.render();
    expect(h.wb.totalCells).toBe(before);
    expect(h.sheet.getValue(0, 0)).toBe('A1');
  });

  it('throws when the canvas has no 2d context', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    const blind = { width: 0, height: 0, getContext: () => null };
    const view = new GridView(blind, wb, 'S', { width: 100, height: 100 });
    expect(() => view.render()).toThrow(/2d context/);
  });
});

describe('GridView hit testing and pointer input', () => {
  let h: ReturnType<typeof build>;
  beforeEach(() => {
    h = build();
  });

  it('maps a point to a cell', () => {
    expect(h.view.hitTest(50, 25)).toEqual({ kind: 'cell', row: 0, col: 0 });
  });

  it('selects the cell under a press', () => {
    h.view.pointerDown(50 + 64, 25 + 20);
    expect(h.view.selection.active).toEqual({ row: 1, col: 1 });
  });

  it('extends with a shift press', () => {
    h.view.pointerDown(50, 25);
    h.view.pointerDown(50 + 128, 25 + 40, { shift: true });
    expect(h.view.selection.ranges[0]).toEqual({ top: 0, left: 0, bottom: 2, right: 2 });
  });

  it('adds a range with a ctrl press', () => {
    h.view.pointerDown(50, 25);
    h.view.pointerDown(50 + 128, 25 + 40, { ctrl: true });
    expect(h.view.selection.ranges).toHaveLength(2);
  });

  it('selects a whole row from its header', () => {
    const target = h.view.pointerDown(10, 25 + 40);
    expect(target.kind).toBe('row-header');
    expect(h.view.selection.isEntireRowSelected(2)).toBe(true);
  });

  it('selects a whole column from its header', () => {
    const target = h.view.pointerDown(46 + 64 + 32, 5);
    expect(target.kind).toBe('col-header');
    expect(h.view.selection.isEntireColSelected(1)).toBe(true);
  });

  it('selects everything from the corner box', () => {
    h.view.pointerDown(5, 5);
    expect(h.view.selection.ranges[0]?.bottom).toBe(MAX_ROWS - 1);
  });

  it('extends the selection while dragging', () => {
    h.view.pointerDown(50, 25);
    h.view.pointerMove(50 + 192, 25 + 60);
    expect(h.view.selection.ranges[0]).toEqual({ top: 0, left: 0, bottom: 3, right: 3 });
    h.view.pointerUp();
    h.view.pointerMove(50, 25);
    expect(h.view.selection.ranges[0]?.bottom).toBe(3);
  });

  it('reports a resize handle without changing the selection', () => {
    const before = h.view.selection.active;
    const target = h.view.pointerDown(46 + 64, 5);
    expect(target.kind).toBe('col-resize');
    expect(h.view.selection.active).toEqual(before);
  });

  it('asks the host to edit on a double click', () => {
    const spy = vi.fn();
    h.view.on('activate', spy);
    h.view.doubleClick(50 + 64, 25);
    expect(spy).toHaveBeenCalledWith({ row: 0, col: 1 });
  });

  it('reports a header resize as intent, never as an edit', () => {
    const spy = vi.fn();
    h.view.on('resize', spy);
    h.view.requestResize('col', 3, 120);
    expect(spy).toHaveBeenCalledWith({ kind: 'col', index: 3, size: 120 });
    expect(h.sheet.cols.get(3)).toBeUndefined();
  });
});

describe('GridView keyboard', () => {
  let h: ReturnType<typeof build>;
  beforeEach(() => {
    h = build();
  });

  it('routes keys to the selection', () => {
    expect(h.view.keyDown({ key: 'ArrowDown' })).toBe(true);
    expect(h.view.selection.active).toEqual({ row: 1, col: 0 });
  });

  it('reports keys it does not handle', () => {
    expect(h.view.keyDown({ key: 'q' })).toBe(false);
  });

  it('keeps the cursor on screen', () => {
    for (let i = 0; i < 40; i++) h.view.keyDown({ key: 'ArrowDown' });
    expect(h.view.scrollY).toBeGreaterThan(0);
    const layout = h.view.layout();
    expect(layout.rowSpans.some((s) => s.index === h.view.selection.active.row)).toBe(true);
  });

  it('emits selection changes', () => {
    const spy = vi.fn();
    h.view.on('selectionchange', spy);
    h.view.keyDown({ key: 'ArrowRight' });
    expect(spy).toHaveBeenCalled();
    expect(spy.mock.calls[0]?.[0].active).toEqual({ row: 0, col: 1 });
  });

  it('treats F2 as a request to edit', () => {
    const spy = vi.fn();
    h.view.on('activate', spy);
    expect(h.view.keyDown({ key: 'F2' })).toBe(true);
    expect(spy).toHaveBeenCalledWith({ row: 0, col: 0 });
  });

  it('uses the real page size for Page Down', () => {
    h.view.keyDown({ key: 'PageDown' });
    expect(h.view.selection.active.row).toBe(Math.floor((400 - 20) / 20));
  });

  it('stops listening once unsubscribed', () => {
    const spy = vi.fn();
    const off = h.view.on('selectionchange', spy);
    off();
    h.view.keyDown({ key: 'ArrowRight' });
    expect(spy).not.toHaveBeenCalled();
  });
});

describe('GridView model changes', () => {
  it('picks up a new row height after invalidate', () => {
    const h = build();
    expect(h.view.rows.sizeOf(0)).toBe(20);
    h.sheet.rows.set(0, { height: 45 });
    h.view.invalidate();
    expect(h.view.rows.sizeOf(0)).toBe(60);
  });

  it('ignores model changes until told', () => {
    const h = build();
    h.view.layout();
    h.sheet.rows.set(0, { height: 45 });
    expect(h.view.rows.sizeOf(0)).toBe(20);
  });

  it('refreshes the navigation bounds on invalidate', () => {
    const h = build();
    h.sheet.setValue(50, 3, 'new');
    h.view.invalidate();
    h.view.keyDown({ key: 'End', ctrl: true });
    expect(h.view.selection.active).toEqual({ row: 50, col: 4 });
  });

  it('applies zoom to both axes', () => {
    const h = build();
    h.view.setZoom(2);
    expect(h.view.rows.sizeOf(0)).toBe(40);
    expect(h.view.cols.sizeOf(0)).toBe(128);
  });

  it('switches sheets and resets the scroll', () => {
    const h = build();
    h.wb.addSheet('Sheet2').setValue(0, 0, 'other');
    h.view.setScroll(0, 400);
    const spy = vi.fn();
    h.view.on('sheetchange', spy);
    h.view.setSheet('Sheet2');
    expect(spy).toHaveBeenCalledWith({ sheet: 'Sheet2' });
    expect(h.view.scrollY).toBe(0);
    h.view.render();
    expect(h.canvas.context.texts()).toContain('other');
  });

  it('rejects a switch to a sheet that does not exist', () => {
    expect(() => build().view.setSheet('Ghost')).toThrow(/no such sheet/);
  });

  it('reports the rectangle of a cell for an overlaid editor', () => {
    const h = build();
    expect(h.view.cellRect(1, 1)).toEqual({ x: 46 + 64, y: 20 + 20, width: 64, height: 20 });
  });

  it('reports whether a cell is empty', () => {
    const h = build();
    expect(h.view.isCellEmpty(0, 0)).toBe(false);
    expect(h.view.isCellEmpty(9, 9)).toBe(true);
  });

  it('exposes the selected ranges', () => {
    const h = build();
    h.view.selection.selectRange({ top: 1, left: 1, bottom: 2, right: 2 });
    expect(h.view.selectedRanges()).toHaveLength(1);
  });
});
