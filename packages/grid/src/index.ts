/**
 * GridView: a canvas that draws one sheet and reports what the user pointed at.
 *
 * The deliberate limit of this class is that it owns NOTHING it draws. It holds
 * a Workbook and a sheet name, a scroll position, a selection and a theme, and
 * that is all. It never writes a cell, never touches the Document command log,
 * and never decides what an edit means. A double-click emits an `activate`
 * event; whether that opens an editor, and what the editor does, is the host's
 * problem.
 *
 * Keeping the renderer read-only is what lets the same component draw a live
 * document, a preview of an undo state, a print page and a diff view without any
 * of them being able to corrupt the model by accident.
 */

import type { Workbook } from '@mirrorz/core';
import { type Sheet, packKey } from '@mirrorz/core';

import {
  type GridRenderingContext,
  type RenderStats,
  renderGrid,
} from './render.js';
import {
  type CellPos,
  type GridRange,
  type KeyInput,
  Selection,
  type SelectionSource,
  type SelectionState,
  sheetSource,
} from './selection.js';
import { LIGHT_THEME, type GridTheme } from './theme.js';
import { TextMeasureCache } from './text.js';
import {
  type AxisIndex,
  type GridLayout,
  type HitTarget,
  Viewport,
  buildColAxis,
  buildRowAxis,
} from './viewport.js';

export * from './viewport.js';
export * from './selection.js';
export * from './theme.js';
export * from './text.js';
export * from './render.js';

/**
 * The part of HTMLCanvasElement that is used. Declared structurally so a grid
 * can be rendered headlessly - to an OffscreenCanvas, or to a test double - not
 * only into a DOM element.
 */
export interface CanvasLike {
  width: number;
  height: number;
  style?: { width: string; height: string };
  getContext(contextId: '2d'): GridRenderingContext | null;
}

export interface GridViewOptions {
  theme?: GridTheme;
  dpr?: number;
  width?: number;
  height?: number;
  showHeaders?: boolean;
  showGridLines?: boolean;
  headerWidth?: number;
  headerHeight?: number;
  zoom?: number;
  /** Resolved workbook theme palette, for theme-indexed cell colours. */
  themePalette?: readonly string[];
  measureLimit?: number;
}

export interface GridEvents {
  selectionchange: SelectionState;
  scroll: { x: number; y: number };
  render: RenderStats;
  /** The user asked to start editing a cell - double click, F2, or typing. */
  activate: CellPos;
  /** A header edge was dragged. The host applies it through the command log. */
  resize: { kind: 'row' | 'col'; index: number; size: number };
  sheetchange: { sheet: string };
}

type Listener<K extends keyof GridEvents> = (payload: GridEvents[K]) => void;

export interface PointerModifiers {
  shift?: boolean;
  ctrl?: boolean;
  meta?: boolean;
  alt?: boolean;
}

export class GridView {
  readonly canvas: CanvasLike;
  readonly selection: Selection;
  readonly measure: TextMeasureCache;
  workbook: Workbook;
  theme: GridTheme;
  showGridLines: boolean;
  themePalette?: readonly string[];
  zoom: number;

  private sheetName: string;
  private viewport: Viewport;
  private ctx: GridRenderingContext | null;
  private listeners: { [K in keyof GridEvents]?: Set<Listener<K>> } = {};
  private axesDirty = true;
  private dragging: 'none' | 'cells' | 'rows' | 'cols' = 'none';

  constructor(
    canvas: CanvasLike,
    workbook: Workbook,
    sheetName: string,
    options: GridViewOptions = {},
  ) {
    this.canvas = canvas;
    this.workbook = workbook;
    this.sheetName = sheetName;
    this.theme = options.theme ?? LIGHT_THEME;
    this.showGridLines = options.showGridLines ?? true;
    if (options.themePalette) this.themePalette = options.themePalette;
    this.zoom = options.zoom ?? 1;
    this.measure = new TextMeasureCache(options.measureLimit ?? 100_000);
    this.ctx = canvas.getContext('2d');

    const sheet = this.sheet;
    this.viewport = new Viewport({
      rows: buildRowAxis(sheet, { zoom: this.zoom }),
      cols: buildColAxis(sheet, { zoom: this.zoom }),
      width: options.width ?? 0,
      height: options.height ?? 0,
      dpr: options.dpr ?? 1,
      showHeaders: options.showHeaders ?? true,
      ...(options.headerWidth !== undefined ? { headerWidth: options.headerWidth } : {}),
      ...(options.headerHeight !== undefined ? { headerHeight: options.headerHeight } : {}),
      frozenRows: sheet.view.frozenRows ?? 0,
      frozenCols: sheet.view.frozenCols ?? 0,
    });
    this.axesDirty = false;

    this.selection = new Selection({
      source: sheetSource(sheet),
      pageRows: () => this.viewport.pageRows(),
      pageCols: () => this.viewport.pageCols(),
    });
    this.selection.onChange((state) => this.emit('selectionchange', state));

    const active = sheet.view.activeCell;
    if (active) this.selection.selectCell(active.row, active.col);

    if (options.width !== undefined && options.height !== undefined) {
      this.resize(options.width, options.height, options.dpr);
    }
  }

  get sheet(): Sheet {
    const sheet = this.workbook.getSheet(this.sheetName);
    if (!sheet) throw new Error(`no such sheet: ${this.sheetName}`);
    return sheet;
  }

  get sheetNameValue(): string {
    return this.sheetName;
  }

  get rows(): AxisIndex {
    this.ensureAxes();
    return this.viewport.rows;
  }

  get cols(): AxisIndex {
    this.ensureAxes();
    return this.viewport.cols;
  }

  get scrollX(): number {
    return this.viewport.scrollX;
  }

  get scrollY(): number {
    return this.viewport.scrollY;
  }

  get maxScrollX(): number {
    this.ensureAxes();
    return this.viewport.maxScrollX;
  }

  get maxScrollY(): number {
    this.ensureAxes();
    return this.viewport.maxScrollY;
  }

  // --- events -------------------------------------------------------------

  on<K extends keyof GridEvents>(event: K, listener: Listener<K>): () => void {
    let set = this.listeners[event] as Set<Listener<K>> | undefined;
    if (!set) {
      set = new Set<Listener<K>>();
      (this.listeners as Record<string, unknown>)[event] = set;
    }
    set.add(listener);
    return () => {
      set?.delete(listener);
    };
  }

  private emit<K extends keyof GridEvents>(event: K, payload: GridEvents[K]): void {
    const set = this.listeners[event] as Set<Listener<K>> | undefined;
    if (!set) return;
    for (const l of set) l(payload);
  }

  // --- configuration ------------------------------------------------------

  setSheet(name: string): void {
    if (!this.workbook.getSheet(name)) throw new Error(`no such sheet: ${name}`);
    this.sheetName = name;
    this.axesDirty = true;
    const sheet = this.sheet;
    this.selection.source = sheetSource(sheet);
    this.viewport.setFrozen(sheet.view.frozenRows ?? 0, sheet.view.frozenCols ?? 0);
    this.viewport.setScroll(0, 0);
    this.measure.clear();
    this.emit('sheetchange', { sheet: name });
  }

  setTheme(theme: GridTheme): void {
    this.theme = theme;
    // Font metrics are theme-dependent, so cached widths are no longer valid.
    this.measure.clear();
  }

  setZoom(zoom: number): void {
    this.zoom = zoom > 0 ? zoom : 1;
    this.axesDirty = true;
    this.measure.clear();
  }

  setFrozen(rows: number, cols: number): void {
    this.ensureAxes();
    this.viewport.setFrozen(rows, cols);
  }

  /**
   * Tell the view its model changed. Row heights, column widths, hidden state
   * and the used extent are all cached, and this is the only way they refresh.
   */
  invalidate(): void {
    this.axesDirty = true;
    this.selection.source = sheetSource(this.sheet);
  }

  private ensureAxes(): void {
    if (!this.axesDirty) return;
    const sheet = this.sheet;
    this.viewport.rows = buildRowAxis(sheet, { zoom: this.zoom });
    this.viewport.cols = buildColAxis(sheet, { zoom: this.zoom });
    this.axesDirty = false;
    // Re-clamp: a shrunken sheet can leave the scroll position past the end.
    this.viewport.setScroll(this.viewport.scrollX, this.viewport.scrollY);
  }

  // --- geometry -----------------------------------------------------------

  /**
   * Resize the backing store. The canvas gets device pixels and the CSS box gets
   * logical pixels; conflating the two is what makes canvas text blurry.
   */
  resize(width: number, height: number, dpr?: number): void {
    this.viewport.resize(width, height, dpr);
    const ratio = this.viewport.dpr;
    this.canvas.width = Math.max(0, Math.round(width * ratio));
    this.canvas.height = Math.max(0, Math.round(height * ratio));
    if (this.canvas.style) {
      this.canvas.style.width = `${width}px`;
      this.canvas.style.height = `${height}px`;
    }
  }

  setScroll(x: number, y: number): void {
    this.ensureAxes();
    const before = { x: this.viewport.scrollX, y: this.viewport.scrollY };
    this.viewport.setScroll(x, y);
    if (before.x !== this.viewport.scrollX || before.y !== this.viewport.scrollY) {
      this.emit('scroll', { x: this.viewport.scrollX, y: this.viewport.scrollY });
    }
  }

  scrollBy(dx: number, dy: number): void {
    this.setScroll(this.viewport.scrollX + dx, this.viewport.scrollY + dy);
  }

  /** Minimal scroll that brings a cell fully into view. */
  scrollToCell(row: number, col: number): void {
    this.ensureAxes();
    const before = { x: this.viewport.scrollX, y: this.viewport.scrollY };
    this.viewport.revealCell(row, col);
    if (before.x !== this.viewport.scrollX || before.y !== this.viewport.scrollY) {
      this.emit('scroll', { x: this.viewport.scrollX, y: this.viewport.scrollY });
    }
  }

  layout(): GridLayout {
    this.ensureAxes();
    return this.viewport.layout();
  }

  hitTest(x: number, y: number): HitTarget {
    this.ensureAxes();
    return this.viewport.hitTest(x, y);
  }

  /** Canvas rectangle of a cell, for positioning an editor over it. */
  cellRect(row: number, col: number): { x: number; y: number; width: number; height: number } {
    this.ensureAxes();
    return {
      x: this.viewport.xOf(col),
      y: this.viewport.yOf(row),
      width: this.viewport.cols.sizeOf(col),
      height: this.viewport.rows.sizeOf(row),
    };
  }

  // --- painting -----------------------------------------------------------

  render(): RenderStats {
    this.ensureAxes();
    const ctx = this.ctx ?? this.canvas.getContext('2d');
    if (!ctx) throw new Error('canvas has no 2d context');
    this.ctx = ctx;
    const sheet = this.sheet;
    const stats = renderGrid(ctx, {
      sheet,
      styles: this.workbook.styles,
      layout: this.viewport.layout(),
      geometry: this.viewport,
      theme: this.theme,
      measure: this.measure,
      selection: this.selection.state,
      showGridLines: this.showGridLines && sheet.view.showGridLines !== false,
      dateSystem: this.workbook.dateSystem,
      ...(this.themePalette ? { themePalette: this.themePalette } : {}),
      zoom: this.zoom,
    });
    this.emit('render', stats);
    return stats;
  }

  // --- input --------------------------------------------------------------

  /**
   * A press. Returns what was under the pointer so the host can decide whether
   * this is the start of a drag, a resize, or a context menu.
   */
  pointerDown(x: number, y: number, mods: PointerModifiers = {}): HitTarget {
    const target = this.hitTest(x, y);
    const additive = mods.ctrl === true || mods.meta === true;
    switch (target.kind) {
      case 'cell':
        if (mods.shift) this.selection.extendTo(target.row, target.col);
        else this.selection.selectCell(target.row, target.col, { additive });
        this.dragging = 'cells';
        break;
      case 'row-header':
        this.selection.selectRow(target.row, { additive, extend: mods.shift === true });
        this.dragging = 'rows';
        break;
      case 'col-header':
        this.selection.selectCol(target.col, { additive, extend: mods.shift === true });
        this.dragging = 'cols';
        break;
      case 'corner':
        this.selection.selectAll();
        break;
      default:
        break;
    }
    return target;
  }

  /** Drag: extend whatever the press started. */
  pointerMove(x: number, y: number): void {
    if (this.dragging === 'none') return;
    const target = this.hitTest(x, y);
    if (this.dragging === 'cells' && target.kind === 'cell') {
      this.selection.extendTo(target.row, target.col);
    } else if (this.dragging === 'rows' && (target.kind === 'row-header' || target.kind === 'cell')) {
      this.selection.selectRow(target.row, { extend: true });
    } else if (this.dragging === 'cols' && (target.kind === 'col-header' || target.kind === 'cell')) {
      this.selection.selectCol(target.col, { extend: true });
    }
  }

  pointerUp(): void {
    this.dragging = 'none';
  }

  /** Double click asks the host to edit; the view itself never edits. */
  doubleClick(x: number, y: number): void {
    const target = this.hitTest(x, y);
    if (target.kind === 'cell') this.emit('activate', { row: target.row, col: target.col });
  }

  /**
   * Route a key to the selection, then keep the cursor on screen. Returns true
   * when the key was consumed.
   */
  keyDown(event: KeyInput): boolean {
    this.ensureAxes();
    if (event.key === 'F2') {
      this.emit('activate', this.selection.active);
      return true;
    }
    const handled = this.selection.keyDown(event);
    if (handled) this.scrollToCell(this.selection.active.row, this.selection.active.col);
    return handled;
  }

  /** Report a header drag. The host turns it into an undoable command. */
  requestResize(kind: 'row' | 'col', index: number, size: number): void {
    this.emit('resize', { kind, index, size });
  }

  /** True when a cell holds nothing, which is what overflow painting turns on. */
  isCellEmpty(row: number, col: number): boolean {
    const cell = this.sheet.cells.get(packKey(row, col));
    return cell === undefined || cell.value === null;
  }

  /** Current selection as ranges, for a host that wants to copy or format them. */
  selectedRanges(): readonly GridRange[] {
    return this.selection.ranges;
  }

  get selectionSource(): SelectionSource {
    return this.selection.source;
  }
}
