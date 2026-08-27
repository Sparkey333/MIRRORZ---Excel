/**
 * Grid coordinate system.
 *
 * The whole design turns on one constraint: nothing here may cost anything
 * proportional to the grid. Excel's sheet is 1,048,576 x 16,384, so a naive
 * prefix-sum array over rows is 8 MB of Float64 per sheet and a linear walk to
 * find "which row is at pixel 14,000,000" is 1M iterations per frame.
 *
 * The observation that makes it cheap: almost every row has the default height
 * and almost every column the default width. Only a handful are resized or
 * hidden. So the index is built over the EXCEPTIONS - the rows that differ -
 * and the offset of index i is
 *
 *     offset(i) = i * defaultSize + (sum of size deltas of exceptions before i)
 *
 * The second term is a prefix sum over the exception list, whose length is the
 * number of customised rows, not the number of rows. Both offsetOf and its
 * inverse indexAt are therefore O(log m) with m = customised rows, typically a
 * few dozen. Scrolling to row 900,000 costs the same as scrolling to row 3.
 *
 * Scroll offsets are plain floats and are never rounded to a row boundary.
 * Excel snapped whole rows for thirty years and "smooth scrolling" was its
 * most-voted feature request; the difference is visible within one flick of a
 * trackpad, so sub-pixel scroll offsets are carried all the way to the painter.
 */

import { MAX_COLS, MAX_ROWS, type Sheet } from '@mirrorz/core';

/** A single index whose size differs from the axis default. */
export interface AxisException {
  readonly index: number;
  /** Size in device-independent pixels. Zero means hidden. */
  readonly size: number;
}

/** Sub-pixel nudge used to keep a position at the very end inside the last span. */
const EPSILON = 1e-6;

function lowerBound(a: Int32Array, n: number, value: number): number {
  let lo = 0;
  let hi = n;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if ((a[mid] as number) < value) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** Index of the last entry of `a` that is <= value, or -1. */
function lastAtOrBelow(a: Float64Array, n: number, value: number): number {
  let lo = 0;
  let hi = n;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if ((a[mid] as number) <= value) lo = mid + 1;
    else hi = mid;
  }
  return lo - 1;
}

/**
 * Cumulative-offset index over one axis (rows or columns).
 *
 * Immutable. Rebuild it when row heights change; building costs O(m log m) for
 * the sort and nothing else, so it is cheap enough to rebuild on every resize
 * of a column.
 */
export class AxisIndex {
  readonly count: number;
  readonly defaultSize: number;
  readonly totalSize: number;

  /** Sorted indices of the exceptions. */
  private readonly xi: Int32Array;
  /** Their sizes, parallel to `xi`. */
  private readonly xs: Float64Array;
  /** Start offset of each exception, parallel to `xi`. */
  private readonly xo: Float64Array;
  /** pre[k] = sum of (size - defaultSize) over exceptions [0, k). Length m+1. */
  private readonly pre: Float64Array;
  private readonly m: number;

  constructor(count: number, defaultSize: number, exceptions: Iterable<AxisException>) {
    this.count = Math.max(0, Math.floor(count));
    this.defaultSize = Math.max(0, defaultSize);

    const kept: AxisException[] = [];
    const seen = new Set<number>();
    for (const e of exceptions) {
      if (e.index < 0 || e.index >= this.count) continue;
      if (e.size === this.defaultSize) continue;
      if (seen.has(e.index)) continue;
      seen.add(e.index);
      kept.push({ index: e.index, size: Math.max(0, e.size) });
    }
    kept.sort((a, b) => a.index - b.index);

    const m = kept.length;
    this.m = m;
    this.xi = new Int32Array(m);
    this.xs = new Float64Array(m);
    this.xo = new Float64Array(m);
    this.pre = new Float64Array(m + 1);

    let delta = 0;
    for (let k = 0; k < m; k++) {
      const e = kept[k] as AxisException;
      this.xi[k] = e.index;
      this.xs[k] = e.size;
      this.pre[k] = delta;
      this.xo[k] = e.index * this.defaultSize + delta;
      delta += e.size - this.defaultSize;
    }
    this.pre[m] = delta;
    this.totalSize = this.count * this.defaultSize + delta;
  }

  /** Number of indices whose size differs from the default. */
  get exceptionCount(): number {
    return this.m;
  }

  sizeOf(index: number): number {
    if (index < 0 || index >= this.count) return 0;
    const k = lowerBound(this.xi, this.m, index);
    if (k < this.m && this.xi[k] === index) return this.xs[k] as number;
    return this.defaultSize;
  }

  isHidden(index: number): boolean {
    return this.sizeOf(index) === 0;
  }

  /**
   * Pixel offset of the start of `index`. Defined for `count` itself, where it
   * equals totalSize, so callers can ask for the end of the last row.
   */
  offsetOf(index: number): number {
    const i = index < 0 ? 0 : index > this.count ? this.count : index;
    const k = lowerBound(this.xi, this.m, i);
    return i * this.defaultSize + (this.pre[k] as number);
  }

  /**
   * The index whose span contains `pos`, clamped into range.
   *
   * Zero-size (hidden) indices are never returned: their span is empty, so a
   * position can never be inside one, and the search naturally lands on the
   * first index after the hidden run.
   */
  indexAt(pos: number): number {
    if (this.count === 0) return 0;
    // Clamping here rather than in the callers keeps this the single place that
    // knows a position past the end belongs to the last row, and it must not
    // route through firstVisibleAtOrAfter, which calls back into this method.
    const p =
      pos <= 0 ? 0 : pos >= this.totalSize ? Math.max(0, this.totalSize - EPSILON) : pos;

    const k = lastAtOrBelow(this.xo, this.m, p);
    if (k < 0) {
      // Before the first exception: pure arithmetic, no search needed.
      if (this.defaultSize <= 0) return 0;
      return Math.min(this.count - 1, Math.floor(p / this.defaultSize));
    }
    const start = this.xo[k] as number;
    const size = this.xs[k] as number;
    if (p < start + size) return this.xi[k] as number;
    if (this.defaultSize <= 0) return Math.min(this.count - 1, (this.xi[k] as number) + 1);
    const rest = p - (start + size);
    const base = (this.xi[k] as number) + 1;
    return Math.min(this.count - 1, base + Math.floor(rest / this.defaultSize));
  }

  /**
   * First index >= `i` with a non-zero size, or `count` when every remaining
   * index is hidden. Lets a span walk skip a million hidden rows in O(log m).
   */
  firstVisibleAtOrAfter(i: number): number {
    let idx = i < 0 ? 0 : i;
    if (idx >= this.count) return this.count;
    if (this.sizeOf(idx) > 0) return idx;
    const off = this.offsetOf(idx);
    if (off >= this.totalSize) return this.count;
    const j = this.indexAt(off);
    return j > idx ? j : idx;
  }

  /** Last index <= `i` with a non-zero size, or -1 when there is none. */
  lastVisibleAtOrBefore(i: number): number {
    let idx = i >= this.count ? this.count - 1 : i;
    if (idx < 0) return -1;
    if (this.sizeOf(idx) > 0) return idx;
    // The end of a hidden index equals its start, so step back one pixel from
    // the offset to land inside the previous visible span.
    const off = this.offsetOf(idx);
    if (off <= 0) return -1;
    return this.indexAt(off - 1);
  }
}

// --- sheet -> axis --------------------------------------------------------

/** CSS pixels per point at the nominal 96 dpi Excel assumes. */
export const PIXELS_PER_POINT = 96 / 72;

/**
 * Width of the '0' glyph of the default font, in pixels. Excel stores column
 * widths as a multiple of this, so the conversion needs it as a constant; 7 is
 * the value for Calibri 11 at 96 dpi, which is what the default styles use.
 */
export const DEFAULT_MAX_DIGIT_WIDTH = 7;

export function pointsToPixels(points: number): number {
  return points * PIXELS_PER_POINT;
}

/**
 * Excel's char-width-to-pixel conversion (ECMA-376 §18.3.1.13, which states the
 * inverse: width = Truncate([{pixels - 5} / MDW * 100 + 0.5] / 100)).
 *
 * The constant 5 is the cell's two-pixel padding plus the gridline, and dropping
 * it is what makes a naive implementation put every column five pixels narrow -
 * enough that the default 8.43 width comes out as 59px instead of Excel's 64 and
 * a whole screen of columns drifts visibly out of alignment.
 */
export function charsToPixels(width: number, maxDigitWidth = DEFAULT_MAX_DIGIT_WIDTH): number {
  if (width <= 0) return 0;
  return Math.round(width * maxDigitWidth) + 5;
}

export interface AxisBuildOptions {
  /** 1 = 100%. Applied to every size so zoom needs no separate transform. */
  zoom?: number;
  maxDigitWidth?: number;
}

export function buildRowAxis(sheet: Sheet, options: AxisBuildOptions = {}): AxisIndex {
  const zoom = options.zoom ?? 1;
  const def = pointsToPixels(sheet.defaultRowHeight) * zoom;
  const exceptions: AxisException[] = [];
  for (const [row, props] of sheet.rows) {
    if (props.hidden) exceptions.push({ index: row, size: 0 });
    else if (props.height !== undefined) {
      exceptions.push({ index: row, size: pointsToPixels(props.height) * zoom });
    }
  }
  return new AxisIndex(MAX_ROWS, def, exceptions);
}

export function buildColAxis(sheet: Sheet, options: AxisBuildOptions = {}): AxisIndex {
  const zoom = options.zoom ?? 1;
  const mdw = options.maxDigitWidth ?? DEFAULT_MAX_DIGIT_WIDTH;
  const def = charsToPixels(sheet.defaultColWidth, mdw) * zoom;
  const exceptions: AxisException[] = [];
  for (const [col, props] of sheet.cols) {
    if (props.hidden) exceptions.push({ index: col, size: 0 });
    else if (props.width !== undefined) {
      exceptions.push({ index: col, size: charsToPixels(props.width, mdw) * zoom });
    }
  }
  return new AxisIndex(MAX_COLS, def, exceptions);
}

// --- viewport -------------------------------------------------------------

export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** One row or column that is at least partly on screen. */
export interface Span {
  readonly index: number;
  /** Left/top edge in canvas CSS pixels. May be negative when partly scrolled. */
  readonly start: number;
  readonly size: number;
}

export type PaneId = 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';

export interface PaneLayout {
  readonly id: PaneId;
  /** Clip rectangle in canvas CSS pixels, excluding the headers. */
  readonly rect: Rect;
  readonly rows: readonly Span[];
  readonly cols: readonly Span[];
  readonly frozenRows: boolean;
  readonly frozenCols: boolean;
}

export interface GridLayout {
  readonly width: number;
  readonly height: number;
  readonly dpr: number;
  readonly headerWidth: number;
  readonly headerHeight: number;
  readonly showHeaders: boolean;
  readonly panes: readonly PaneLayout[];
  /** Frozen rows first, then scrolling rows: the order the row header paints. */
  readonly rowSpans: readonly Span[];
  readonly colSpans: readonly Span[];
  readonly frozenRowCount: number;
  readonly frozenColCount: number;
  readonly frozenWidth: number;
  readonly frozenHeight: number;
  readonly scrollX: number;
  readonly scrollY: number;
  readonly maxScrollX: number;
  readonly maxScrollY: number;
  readonly totalWidth: number;
  readonly totalHeight: number;
  readonly gridRect: Rect;
}

export type HitTarget =
  | { readonly kind: 'cell'; readonly row: number; readonly col: number }
  | { readonly kind: 'row-header'; readonly row: number }
  | { readonly kind: 'col-header'; readonly col: number }
  | { readonly kind: 'row-resize'; readonly row: number }
  | { readonly kind: 'col-resize'; readonly col: number }
  | { readonly kind: 'corner' }
  | { readonly kind: 'outside' };

export interface ViewportOptions {
  rows: AxisIndex;
  cols: AxisIndex;
  width?: number;
  height?: number;
  dpr?: number;
  headerWidth?: number;
  headerHeight?: number;
  showHeaders?: boolean;
  frozenRows?: number;
  frozenCols?: number;
  /** Grab distance in pixels for the resize handles in the headers. */
  resizeHandle?: number;
}

export const DEFAULT_HEADER_WIDTH = 46;
export const DEFAULT_HEADER_HEIGHT = 20;

export class Viewport {
  rows: AxisIndex;
  cols: AxisIndex;
  width: number;
  height: number;
  dpr: number;
  headerWidth: number;
  headerHeight: number;
  showHeaders: boolean;
  frozenRows: number;
  frozenCols: number;
  resizeHandle: number;

  private sx = 0;
  private sy = 0;

  constructor(options: ViewportOptions) {
    this.rows = options.rows;
    this.cols = options.cols;
    this.width = options.width ?? 0;
    this.height = options.height ?? 0;
    this.dpr = options.dpr ?? 1;
    this.showHeaders = options.showHeaders ?? true;
    this.headerWidth = options.headerWidth ?? DEFAULT_HEADER_WIDTH;
    this.headerHeight = options.headerHeight ?? DEFAULT_HEADER_HEIGHT;
    this.frozenRows = Math.max(0, options.frozenRows ?? 0);
    this.frozenCols = Math.max(0, options.frozenCols ?? 0);
    this.resizeHandle = options.resizeHandle ?? 4;
  }

  get scrollX(): number {
    return this.sx;
  }

  get scrollY(): number {
    return this.sy;
  }

  get gridLeft(): number {
    return this.showHeaders ? this.headerWidth : 0;
  }

  get gridTop(): number {
    return this.showHeaders ? this.headerHeight : 0;
  }

  /** Width of the frozen column band, never allowed to eat the whole viewport. */
  get frozenWidth(): number {
    if (this.frozenCols <= 0) return 0;
    const raw = this.cols.offsetOf(this.frozenCols);
    return Math.min(raw, Math.max(0, this.width - this.gridLeft));
  }

  get frozenHeight(): number {
    if (this.frozenRows <= 0) return 0;
    const raw = this.rows.offsetOf(this.frozenRows);
    return Math.min(raw, Math.max(0, this.height - this.gridTop));
  }

  /** Width available to the scrolling column band. */
  get scrollViewWidth(): number {
    return Math.max(0, this.width - this.gridLeft - this.frozenWidth);
  }

  get scrollViewHeight(): number {
    return Math.max(0, this.height - this.gridTop - this.frozenHeight);
  }

  get maxScrollX(): number {
    const origin = this.cols.offsetOf(this.frozenCols);
    return Math.max(0, this.cols.totalSize - origin - this.scrollViewWidth);
  }

  get maxScrollY(): number {
    const origin = this.rows.offsetOf(this.frozenRows);
    return Math.max(0, this.rows.totalSize - origin - this.scrollViewHeight);
  }

  resize(width: number, height: number, dpr?: number): void {
    this.width = Math.max(0, width);
    this.height = Math.max(0, height);
    if (dpr !== undefined) this.dpr = dpr > 0 ? dpr : 1;
    this.setScroll(this.sx, this.sy);
  }

  /** Scroll offsets are floats and are deliberately not snapped to a boundary. */
  setScroll(x: number, y: number): void {
    this.sx = clamp(Number.isFinite(x) ? x : 0, 0, this.maxScrollX);
    this.sy = clamp(Number.isFinite(y) ? y : 0, 0, this.maxScrollY);
  }

  scrollBy(dx: number, dy: number): void {
    this.setScroll(this.sx + dx, this.sy + dy);
  }

  setFrozen(rows: number, cols: number): void {
    this.frozenRows = Math.max(0, Math.floor(rows));
    this.frozenCols = Math.max(0, Math.floor(cols));
    this.setScroll(this.sx, this.sy);
  }

  /** Canvas x of the left edge of a column, wherever it lands. */
  xOf(col: number): number {
    if (col < this.frozenCols) return this.gridLeft + this.cols.offsetOf(col);
    const origin = this.cols.offsetOf(this.frozenCols);
    return this.gridLeft + this.frozenWidth + (this.cols.offsetOf(col) - origin - this.sx);
  }

  yOf(row: number): number {
    if (row < this.frozenRows) return this.gridTop + this.rows.offsetOf(row);
    const origin = this.rows.offsetOf(this.frozenRows);
    return this.gridTop + this.frozenHeight + (this.rows.offsetOf(row) - origin - this.sy);
  }

  /** First scrolling column at least partly visible. */
  firstScrollCol(): number {
    const origin = this.cols.offsetOf(this.frozenCols);
    const idx = this.cols.indexAt(origin + this.sx);
    return Math.max(this.frozenCols, idx);
  }

  firstScrollRow(): number {
    const origin = this.rows.offsetOf(this.frozenRows);
    const idx = this.rows.indexAt(origin + this.sy);
    return Math.max(this.frozenRows, idx);
  }

  /**
   * Compute everything the painter needs for one frame.
   *
   * Cost is proportional to the number of visible rows and columns only. A
   * viewport 1000 px tall over 20 px rows produces at most 51 row spans whether
   * the sheet has ten rows or a million.
   */
  layout(): GridLayout {
    const gridLeft = this.gridLeft;
    const gridTop = this.gridTop;
    const frozenWidth = this.frozenWidth;
    const frozenHeight = this.frozenHeight;
    const right = this.width;
    const bottom = this.height;

    const frozenCols = collectSpans(this.cols, 0, gridLeft, gridLeft + frozenWidth, (c) =>
      gridLeft + this.cols.offsetOf(c),
    ).filter((s) => s.index < this.frozenCols);
    const frozenRows = collectSpans(this.rows, 0, gridTop, gridTop + frozenHeight, (r) =>
      gridTop + this.rows.offsetOf(r),
    ).filter((s) => s.index < this.frozenRows);

    const scrollColStart = this.firstScrollCol();
    const scrollRowStart = this.firstScrollRow();
    const scrollCols = collectSpans(
      this.cols,
      scrollColStart,
      gridLeft + frozenWidth,
      right,
      (c) => this.xOf(c),
    );
    const scrollRows = collectSpans(
      this.rows,
      scrollRowStart,
      gridTop + frozenHeight,
      bottom,
      (r) => this.yOf(r),
    );

    const frozenRect: Rect = { x: gridLeft, y: gridTop, width: frozenWidth, height: frozenHeight };
    const panes: PaneLayout[] = [];
    const push = (
      id: PaneId,
      rect: Rect,
      rowsSpans: Span[],
      colsSpans: Span[],
      fr: boolean,
      fc: boolean,
    ): void => {
      if (rowsSpans.length === 0 || colsSpans.length === 0) return;
      if (rect.width <= 0 || rect.height <= 0) return;
      panes.push({ id, rect, rows: rowsSpans, cols: colsSpans, frozenRows: fr, frozenCols: fc });
    };

    push('topLeft', frozenRect, frozenRows, frozenCols, true, true);
    push(
      'topRight',
      { x: gridLeft + frozenWidth, y: gridTop, width: right - gridLeft - frozenWidth, height: frozenHeight },
      frozenRows,
      scrollCols,
      true,
      false,
    );
    push(
      'bottomLeft',
      { x: gridLeft, y: gridTop + frozenHeight, width: frozenWidth, height: bottom - gridTop - frozenHeight },
      scrollRows,
      frozenCols,
      false,
      true,
    );
    push(
      'bottomRight',
      {
        x: gridLeft + frozenWidth,
        y: gridTop + frozenHeight,
        width: right - gridLeft - frozenWidth,
        height: bottom - gridTop - frozenHeight,
      },
      scrollRows,
      scrollCols,
      false,
      false,
    );

    return {
      width: this.width,
      height: this.height,
      dpr: this.dpr,
      headerWidth: gridLeft,
      headerHeight: gridTop,
      showHeaders: this.showHeaders,
      panes,
      rowSpans: [...frozenRows, ...scrollRows],
      colSpans: [...frozenCols, ...scrollCols],
      frozenRowCount: this.frozenRows,
      frozenColCount: this.frozenCols,
      frozenWidth,
      frozenHeight,
      scrollX: this.sx,
      scrollY: this.sy,
      maxScrollX: this.maxScrollX,
      maxScrollY: this.maxScrollY,
      totalWidth: this.cols.totalSize,
      totalHeight: this.rows.totalSize,
      gridRect: { x: gridLeft, y: gridTop, width: right - gridLeft, height: bottom - gridTop },
    };
  }

  /**
   * Map a canvas point to what is under it. Header resize handles are reported
   * separately so the host can change the cursor without re-deriving geometry.
   */
  hitTest(x: number, y: number): HitTarget {
    if (x < 0 || y < 0 || x > this.width || y > this.height) return { kind: 'outside' };
    const gridLeft = this.gridLeft;
    const gridTop = this.gridTop;

    if (this.showHeaders && x < gridLeft && y < gridTop) return { kind: 'corner' };

    if (this.showHeaders && y < gridTop) {
      const col = this.colAt(x);
      if (col === undefined) return { kind: 'outside' };
      const edge = this.xOf(col) + this.cols.sizeOf(col);
      if (Math.abs(x - edge) <= this.resizeHandle) return { kind: 'col-resize', col };
      const leftEdge = this.xOf(col);
      if (col > 0 && Math.abs(x - leftEdge) <= this.resizeHandle) {
        const prev = this.cols.lastVisibleAtOrBefore(col - 1);
        if (prev >= 0) return { kind: 'col-resize', col: prev };
      }
      return { kind: 'col-header', col };
    }

    if (this.showHeaders && x < gridLeft) {
      const row = this.rowAt(y);
      if (row === undefined) return { kind: 'outside' };
      const edge = this.yOf(row) + this.rows.sizeOf(row);
      if (Math.abs(y - edge) <= this.resizeHandle) return { kind: 'row-resize', row };
      const topEdge = this.yOf(row);
      if (row > 0 && Math.abs(y - topEdge) <= this.resizeHandle) {
        const prev = this.rows.lastVisibleAtOrBefore(row - 1);
        if (prev >= 0) return { kind: 'row-resize', row: prev };
      }
      return { kind: 'row-header', row };
    }

    const col = this.colAt(x);
    const row = this.rowAt(y);
    if (col === undefined || row === undefined) return { kind: 'outside' };
    return { kind: 'cell', row, col };
  }

  /** Column under a canvas x, honouring the frozen band. */
  colAt(x: number): number | undefined {
    const gridLeft = this.gridLeft;
    if (x < gridLeft) return undefined;
    const frozenWidth = this.frozenWidth;
    if (this.frozenCols > 0 && x < gridLeft + frozenWidth) {
      return this.cols.indexAt(x - gridLeft);
    }
    const origin = this.cols.offsetOf(this.frozenCols);
    const pos = origin + this.sx + (x - gridLeft - frozenWidth);
    const idx = this.cols.indexAt(pos);
    return Math.max(this.frozenCols, idx);
  }

  rowAt(y: number): number | undefined {
    const gridTop = this.gridTop;
    if (y < gridTop) return undefined;
    const frozenHeight = this.frozenHeight;
    if (this.frozenRows > 0 && y < gridTop + frozenHeight) {
      return this.rows.indexAt(y - gridTop);
    }
    const origin = this.rows.offsetOf(this.frozenRows);
    const pos = origin + this.sy + (y - gridTop - frozenHeight);
    const idx = this.rows.indexAt(pos);
    return Math.max(this.frozenRows, idx);
  }

  /** True when the cell is fully inside the scrollable area right now. */
  isFullyVisible(row: number, col: number): boolean {
    const x = this.xOf(col);
    const y = this.yOf(row);
    return (
      x >= this.gridLeft + (col >= this.frozenCols ? this.frozenWidth : 0) &&
      y >= this.gridTop + (row >= this.frozenRows ? this.frozenHeight : 0) &&
      x + this.cols.sizeOf(col) <= this.width &&
      y + this.rows.sizeOf(row) <= this.height
    );
  }

  /**
   * Minimal scroll that brings a cell into view, the way Excel nudges by just
   * enough rather than centring. Frozen cells are always visible, so asking to
   * reveal one is a no-op.
   */
  revealCell(row: number, col: number): void {
    let x = this.sx;
    let y = this.sy;

    if (col >= this.frozenCols) {
      const origin = this.cols.offsetOf(this.frozenCols);
      const cellStart = this.cols.offsetOf(col) - origin;
      const cellEnd = cellStart + this.cols.sizeOf(col);
      if (cellStart < x) x = cellStart;
      else if (cellEnd > x + this.scrollViewWidth) x = cellEnd - this.scrollViewWidth;
    }
    if (row >= this.frozenRows) {
      const origin = this.rows.offsetOf(this.frozenRows);
      const cellStart = this.rows.offsetOf(row) - origin;
      const cellEnd = cellStart + this.rows.sizeOf(row);
      if (cellStart < y) y = cellStart;
      else if (cellEnd > y + this.scrollViewHeight) y = cellEnd - this.scrollViewHeight;
    }
    this.setScroll(x, y);
  }

  /** Whole rows that fit in the scrolling band: the Page Up/Down step. */
  pageRows(): number {
    const h = this.scrollViewHeight;
    const first = this.firstScrollRow();
    let n = 0;
    let used = 0;
    let r = first;
    while (used < h && r < this.rows.count) {
      r = this.rows.firstVisibleAtOrAfter(r);
      if (r >= this.rows.count) break;
      const size = this.rows.sizeOf(r);
      if (used + size > h && n > 0) break;
      used += size;
      n++;
      r++;
    }
    return Math.max(1, n);
  }

  pageCols(): number {
    const w = this.scrollViewWidth;
    const first = this.firstScrollCol();
    let n = 0;
    let used = 0;
    let c = first;
    while (used < w && c < this.cols.count) {
      c = this.cols.firstVisibleAtOrAfter(c);
      if (c >= this.cols.count) break;
      const size = this.cols.sizeOf(c);
      if (used + size > w && n > 0) break;
      used += size;
      n++;
      c++;
    }
    return Math.max(1, n);
  }
}

function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

/**
 * Walk visible indices from `from` until the band is covered.
 *
 * The loop is bounded by the pixel width of the band divided by the smallest
 * visible size, never by the axis length, and hidden runs are skipped with a
 * binary search rather than iterated.
 */
function collectSpans(
  axis: AxisIndex,
  from: number,
  bandStart: number,
  bandEnd: number,
  positionOf: (index: number) => number,
): Span[] {
  const spans: Span[] = [];
  if (bandEnd <= bandStart) return spans;
  let i = axis.firstVisibleAtOrAfter(Math.max(0, from));
  if (i >= axis.count) return spans;
  let pos = positionOf(i);
  while (pos < bandEnd && i < axis.count) {
    const size = axis.sizeOf(i);
    if (size > 0) {
      spans.push({ index: i, start: pos, size });
      pos += size;
      i++;
      i = axis.firstVisibleAtOrAfter(i);
    } else {
      i = axis.firstVisibleAtOrAfter(i + 1);
    }
    if (i >= axis.count) break;
  }
  return spans;
}
