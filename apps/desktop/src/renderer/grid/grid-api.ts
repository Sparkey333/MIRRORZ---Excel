/**
 * The adapter between the shell and the canvas grid.
 *
 * `@mirrorz/grid` now exists, so this file is no longer a placeholder contract:
 * it is a real translation layer, and it is the only place in the renderer that
 * knows the grid's vocabulary. The grid speaks in `{top,left,bottom,right}`
 * rectangles and emits `selectionchange`; the shell speaks in `{start,end}`
 * ranges on a named sheet. One of them has to convert, and doing it here means a
 * change to either shape is a compile error in one file.
 *
 * Two things this deliberately does NOT do:
 *
 *   It does not load the package through a variable specifier. That was written
 *   while the package was unfinished, and it survived past its usefulness into a
 *   shipping defect: a bare dynamic specifier is left untouched by the bundler,
 *   so the packaged renderer - loaded from file:// with no module resolver -
 *   could never have resolved it and would have shown the fallback grid forever.
 *   A static import is bundled, checked at compile time, and cannot silently
 *   degrade.
 *
 *   It does not write anything. The grid is a read-only view of the Workbook;
 *   every edit the user makes through it comes back as an event and goes through
 *   the controller, and from there through the Document's command log.
 */

import type { Workbook } from '@mirrorz/core';
import {
  DARK_THEME,
  GridView,
  LIGHT_THEME,
  type CanvasLike,
  type GridTheme,
  type GridRange,
  type SelectionState,
} from '@mirrorz/grid';
import { normaliseRange, type Selection } from '../model/selection.js';

export type { GridTheme };

/** Where a pointer landed, in the shell's own vocabulary. */
export interface GridHit {
  row: number;
  col: number;
  region: 'cell' | 'rowHeader' | 'colHeader' | 'corner' | 'outside';
}

export interface GridBridgeOptions {
  dark?: boolean;
  width?: number;
  height?: number;
  dpr?: number;
}

export function gridTheme(dark: boolean): GridTheme {
  return dark ? DARK_THEME : LIGHT_THEME;
}

/**
 * True when this canvas can actually paint.
 *
 * jsdom has no 2d context, and neither does a browser that has run out of GPU
 * contexts. Asking first - rather than constructing a GridView and discovering
 * it at the first paint - is what lets the host fall back to the DOM grid
 * instead of leaving an empty rectangle where the sheet should be.
 */
export function canvasCanPaint(canvas: CanvasLike | null | undefined): boolean {
  if (!canvas) return false;
  try {
    return canvas.getContext('2d') != null;
  } catch {
    return false;
  }
}

/** The grid's selection, as the shell's model. */
export function toSelection(sheet: string, state: SelectionState): Selection {
  const ranges = state.ranges.map((r: GridRange) => ({
    start: { row: r.top, col: r.left },
    end: { row: r.bottom, col: r.right },
  }));
  return {
    sheet,
    active: { row: state.active.row, col: state.active.col },
    ranges: ranges.length > 0 ? ranges : [{ start: state.active, end: state.active }],
  };
}

/** The shell's selection, as the grid's rectangles. */
export function toGridRanges(selection: Selection): GridRange[] {
  return selection.ranges.map((raw) => {
    const r = normaliseRange(raw);
    return { top: r.start.row, left: r.start.col, bottom: r.end.row, right: r.end.col };
  });
}

/** Compare two selections by value, so an echo does not become a feedback loop. */
export function sameSelection(a: Selection, b: Selection): boolean {
  if (a.sheet !== b.sheet) return false;
  if (a.active.row !== b.active.row || a.active.col !== b.active.col) return false;
  if (a.ranges.length !== b.ranges.length) return false;
  for (let i = 0; i < a.ranges.length; i++) {
    const x = normaliseRange(a.ranges[i]!);
    const y = normaliseRange(b.ranges[i]!);
    if (x.start.row !== y.start.row || x.start.col !== y.start.col) return false;
    if (x.end.row !== y.end.row || x.end.col !== y.end.col) return false;
  }
  return true;
}

const REGIONS: Record<string, GridHit['region']> = {
  cell: 'cell',
  'row-header': 'rowHeader',
  'col-header': 'colHeader',
  'row-resize': 'rowHeader',
  'col-resize': 'colHeader',
  corner: 'corner',
  outside: 'outside',
};

/**
 * A small facade over GridView.
 *
 * The facade exists so the host component never touches the grid's own types,
 * and so the one piece of state the two sides share - "this selection change
 * came from the grid, do not push it back" - lives next to the conversion that
 * needs it rather than in a React ref.
 */
export class GridBridge {
  readonly view: GridView;
  private applying = false;

  constructor(
    canvas: CanvasLike,
    workbook: Workbook,
    sheetName: string,
    options: GridBridgeOptions = {},
  ) {
    this.view = new GridView(canvas, workbook, sheetName, {
      theme: gridTheme(options.dark === true),
      ...(options.width !== undefined ? { width: options.width } : {}),
      ...(options.height !== undefined ? { height: options.height } : {}),
      ...(options.dpr !== undefined ? { dpr: options.dpr } : {}),
    });
  }

  get sheetName(): string {
    return this.view.sheetNameValue;
  }

  /** What the grid currently has selected, in the shell's model. */
  currentSelection(): Selection {
    return toSelection(this.view.sheetNameValue, this.view.selection.state);
  }

  /** Selection changes the user made in the grid. Returns an unsubscribe. */
  onSelectionChange(listener: (selection: Selection) => void): () => void {
    return this.view.on('selectionchange', (state) => {
      if (this.applying) return;
      listener(toSelection(this.view.sheetNameValue, state));
    });
  }

  /** The user asked to edit a cell: a double click, F2, or typing over it. */
  onActivate(listener: (cell: { row: number; col: number }) => void): () => void {
    return this.view.on('activate', (cell) => listener({ row: cell.row, col: cell.col }));
  }

  /**
   * Push the shell's selection into the grid without it echoing back as a user
   * action. Sheet mismatches are ignored: the sheet change itself will resync.
   */
  setSelection(selection: Selection): void {
    if (selection.sheet !== this.view.sheetNameValue) return;
    this.applying = true;
    try {
      this.view.selection.setState({
        active: selection.active,
        anchor: selection.active,
        ranges: toGridRanges(selection),
      });
    } finally {
      this.applying = false;
    }
  }

  setSheet(name: string): void {
    if (name === this.view.sheetNameValue) return;
    this.applying = true;
    try {
      this.view.setSheet(name);
    } finally {
      this.applying = false;
    }
  }

  setTheme(dark: boolean): void {
    this.view.setTheme(gridTheme(dark));
  }

  resize(width: number, height: number, dpr?: number): void {
    this.view.resize(width, height, dpr);
  }

  scrollBy(dx: number, dy: number): void {
    if (!GridBridge.finite(dx, dy)) return;
    this.view.scrollBy(dx, dy);
  }

  /**
   * Guard the geometry boundary.
   *
   * `Viewport.hitTest` compares against its bounds, and every comparison with a
   * NaN is false, so a non-finite coordinate walks past the "outside" check and
   * comes back as cell NaN,NaN - which then reaches `a1()` and throws inside a
   * React render, taking the window with it. Pointer events from Chromium always
   * carry numbers; events synthesised by anything else do not always, and this is
   * the boundary where that has to stop.
   */
  private static finite(x: number, y: number): boolean {
    return Number.isFinite(x) && Number.isFinite(y);
  }

  hitTest(x: number, y: number): GridHit {
    if (!GridBridge.finite(x, y)) {
      const active = this.view.selection.active;
      return { row: active.row, col: active.col, region: 'outside' };
    }
    const target = this.view.hitTest(x, y);
    const region = REGIONS[target.kind] ?? 'outside';
    const row = 'row' in target ? target.row : this.view.selection.active.row;
    const col = 'col' in target ? target.col : this.view.selection.active.col;
    return { row, col, region };
  }

  pointerDown(x: number, y: number, mods: { shift?: boolean; ctrl?: boolean; meta?: boolean }): GridHit {
    if (!GridBridge.finite(x, y)) return this.hitTest(x, y);
    this.view.pointerDown(x, y, mods);
    return this.hitTest(x, y);
  }

  pointerMove(x: number, y: number): void {
    if (!GridBridge.finite(x, y)) return;
    this.view.pointerMove(x, y);
  }

  pointerUp(): void {
    this.view.pointerUp();
  }

  doubleClick(x: number, y: number): void {
    if (!GridBridge.finite(x, y)) return;
    this.view.doubleClick(x, y);
  }

  /**
   * The model changed under us. Row heights, hidden state and the used extent
   * are all cached in the view, and this is the only thing that refreshes them -
   * repainting without it draws the new values in the old geometry.
   */
  refresh(): void {
    this.view.invalidate();
    this.render();
  }

  render(): void {
    this.view.render();
  }

  destroy(): void {
    this.view.measure.clear();
  }
}
