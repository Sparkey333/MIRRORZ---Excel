/**
 * Sheet and workbook data model.
 *
 * Storage is sparse: only cells that actually hold something exist. Excel sheets
 * are overwhelmingly empty (a 1M x 16k grid is 17 billion addresses), so a dense
 * array is not an option, and a Map keyed by a packed integer is both smaller and
 * faster than one keyed by "A1" strings.
 *
 * Value and format are stored separately. A cell can carry a format with no
 * value (a blank cell that is still shaded yellow) or a value with the default
 * format, and keeping them apart means a fill-colour change over a large range
 * never has to touch the value map.
 */

import {
  MAX_COLS,
  MAX_ROWS,
  type RangeRef,
  packKey,
  unpackCol,
  unpackRow,
} from './address.js';
import { DEFAULT_STYLE_ID, StyleTable, type StyleId } from './style.js';
import type { CellValue, Scalar } from './types.js';

export interface CellData {
  /** The literal value, or the last computed value when `formula` is set. */
  value: Scalar;
  /**
   * The text the user actually typed or pasted, kept when it cannot be
   * reconstructed from `value` alone.
   *
   * This is what makes type inference reversible rather than destructive: a
   * cell that was entered as `007` or `1-2` can be shown, exported and reverted
   * exactly as supplied, instead of the original being lost the moment it was
   * interpreted. xlsx has no slot for it, so on save it goes to a MIRRORZ-owned
   * part that Excel ignores.
   */
  literal?: string;
  /** Formula source without the leading `=`, when this is a formula cell. */
  formula?: string;
  /** Style id, indexing the workbook's StyleTable. */
  style?: StyleId;
  /**
   * For the anchor of a dynamic array formula, the extent it spills over.
   * Cells inside the spill carry `spillParent` instead.
   */
  spill?: { rows: number; cols: number };
  spillParent?: number;
}

export interface RowProps {
  height?: number;
  hidden?: boolean;
  /** Outline level for grouping, 0-7. */
  level?: number;
  collapsed?: boolean;
  /** Style applied to the whole row. */
  style?: StyleId;
  /** True when the height was set explicitly rather than derived from content. */
  customHeight?: boolean;
}

export interface ColProps {
  /** Width in Excel's "number of '0' characters of the default font" unit. */
  width?: number;
  hidden?: boolean;
  level?: number;
  collapsed?: boolean;
  style?: StyleId;
  customWidth?: boolean;
}

export interface MergedRange {
  readonly range: RangeRef;
}

export interface SheetView {
  frozenRows?: number;
  frozenCols?: number;
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  zoomScale?: number;
  /** Top-left visible cell, restored on open. */
  topLeft?: { row: number; col: number };
  selection?: RangeRef[];
  activeCell?: { row: number; col: number };
  rightToLeft?: boolean;
}

export type SheetVisibility = 'visible' | 'hidden' | 'veryHidden';

/** Default sizes, in Excel's units, for the standard Calibri 11 grid. */
export const DEFAULT_ROW_HEIGHT = 15;
export const DEFAULT_COL_WIDTH = 8.43;

export class Sheet {
  readonly cells = new Map<number, CellData>();
  readonly rows = new Map<number, RowProps>();
  readonly cols = new Map<number, ColProps>();
  merges: MergedRange[] = [];
  view: SheetView = { showGridLines: true, showRowColHeaders: true, zoomScale: 100 };
  visibility: SheetVisibility = 'visible';
  tabColor?: string;
  defaultRowHeight = DEFAULT_ROW_HEIGHT;
  defaultColWidth = DEFAULT_COL_WIDTH;

  /**
   * Parts of the sheet XML we parsed but do not yet model, kept verbatim so a
   * save does not silently delete features we have not implemented. This is the
   * single most important property for trust: opening a file in MIRRORZ and
   * saving it must never quietly destroy a pivot table or a chart.
   */
  preserved: Record<string, string> = {};

  /** Cached bounds; invalidated on write. */
  private boundsCache: { minRow: number; minCol: number; maxRow: number; maxCol: number } | null =
    null;

  constructor(
    public name: string,
    /** Stable id, so renaming a sheet does not break references. */
    readonly id: number,
  ) {}

  get cellCount(): number {
    return this.cells.size;
  }

  getCell(row: number, col: number): CellData | undefined {
    return this.cells.get(packKey(row, col));
  }

  getValue(row: number, col: number): Scalar {
    return this.cells.get(packKey(row, col))?.value ?? null;
  }

  getStyle(row: number, col: number): StyleId {
    const cell = this.cells.get(packKey(row, col));
    if (cell?.style !== undefined) return cell.style;
    // Excel's precedence: cell format, then row format, then column format.
    const rowStyle = this.rows.get(row)?.style;
    if (rowStyle !== undefined) return rowStyle;
    return this.cols.get(col)?.style ?? DEFAULT_STYLE_ID;
  }

  setCell(row: number, col: number, data: CellData | undefined): void {
    assertInBounds(row, col);
    const key = packKey(row, col);
    if (data === undefined || (data.value === null && data.formula === undefined && data.style === undefined)) {
      this.cells.delete(key);
    } else {
      this.cells.set(key, data);
    }
    this.boundsCache = null;
  }

  setValue(row: number, col: number, value: Scalar): void {
    assertInBounds(row, col);
    const key = packKey(row, col);
    const existing = this.cells.get(key);
    if (existing) {
      // Writing a literal clears any formula that was there.
      const next: CellData = { ...existing, value };
      delete next.formula;
      this.cells.set(key, next);
    } else if (value !== null) {
      this.cells.set(key, { value });
    }
    this.boundsCache = null;
  }

  setFormula(row: number, col: number, formula: string, cached: Scalar = null): void {
    assertInBounds(row, col);
    const key = packKey(row, col);
    const existing = this.cells.get(key);
    this.cells.set(key, { ...existing, value: cached, formula });
    this.boundsCache = null;
  }

  setStyle(row: number, col: number, style: StyleId): void {
    assertInBounds(row, col);
    const key = packKey(row, col);
    const existing = this.cells.get(key);
    this.cells.set(key, existing ? { ...existing, style } : { value: null, style });
    this.boundsCache = null;
  }

  /**
   * Used extent of the sheet - Excel's `dimension` element. Returns null for a
   * completely empty sheet.
   */
  bounds(): { minRow: number; minCol: number; maxRow: number; maxCol: number } | null {
    if (this.boundsCache) return this.boundsCache;
    if (this.cells.size === 0 && this.rows.size === 0 && this.cols.size === 0) return null;

    let minRow = Number.POSITIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxRow = -1;
    let maxCol = -1;
    for (const key of this.cells.keys()) {
      const r = unpackRow(key);
      const c = unpackCol(key);
      if (r < minRow) minRow = r;
      if (r > maxRow) maxRow = r;
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    }
    // Rows and columns carrying only formatting still count towards the extent.
    for (const r of this.rows.keys()) {
      if (r < minRow) minRow = r;
      if (r > maxRow) maxRow = r;
    }
    for (const c of this.cols.keys()) {
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    }
    if (maxRow < 0 || maxCol < 0) return null;

    this.boundsCache = {
      minRow: minRow === Number.POSITIVE_INFINITY ? 0 : minRow,
      minCol: minCol === Number.POSITIVE_INFINITY ? 0 : minCol,
      maxRow,
      maxCol,
    };
    return this.boundsCache;
  }

  /** Iterate populated cells in row-major order. */
  *entries(): Generator<{ row: number; col: number; cell: CellData }> {
    const keys = [...this.cells.keys()].sort((a, b) => a - b);
    for (const key of keys) {
      yield { row: unpackRow(key), col: unpackCol(key), cell: this.cells.get(key)! };
    }
  }

  /** Iterate populated cells inside a range, in row-major order. */
  *entriesIn(range: RangeRef): Generator<{ row: number; col: number; cell: CellData }> {
    const { start, end } = range;
    const area = (end.row - start.row + 1) * (end.col - start.col + 1);
    if (area <= this.cells.size) {
      // Small window over a dense sheet: walk the window.
      for (let r = start.row; r <= end.row; r++) {
        for (let c = start.col; c <= end.col; c++) {
          const cell = this.cells.get(packKey(r, c));
          if (cell) yield { row: r, col: c, cell };
        }
      }
      return;
    }
    // Huge window (a whole-column reference, say) over a sparse sheet: walking
    // 1M addresses to find 12 cells would be absurd, so walk the cells instead.
    for (const e of this.entries()) {
      if (e.row >= start.row && e.row <= end.row && e.col >= start.col && e.col <= end.col) {
        yield e;
      }
    }
  }

  rowHeight(row: number): number {
    return this.rows.get(row)?.height ?? this.defaultRowHeight;
  }

  colWidth(col: number): number {
    return this.cols.get(col)?.width ?? this.defaultColWidth;
  }

  isRowHidden(row: number): boolean {
    return this.rows.get(row)?.hidden === true;
  }

  isColHidden(col: number): boolean {
    return this.cols.get(col)?.hidden === true;
  }

  /** The merged range containing this cell, if any. */
  mergeAt(row: number, col: number): RangeRef | undefined {
    for (const m of this.merges) {
      const { start, end } = m.range;
      if (row >= start.row && row <= end.row && col >= start.col && col <= end.col) {
        return m.range;
      }
    }
    return undefined;
  }
}

function assertInBounds(row: number, col: number): void {
  if (row < 0 || row >= MAX_ROWS) throw new RangeError(`row out of range: ${row}`);
  if (col < 0 || col >= MAX_COLS) throw new RangeError(`column out of range: ${col}`);
}

export interface DefinedName {
  name: string;
  /** Formula text the name resolves to, e.g. `Sheet1!$A$1:$B$2`. */
  refersTo: string;
  /** Sheet id for a sheet-scoped name; undefined for workbook scope. */
  scope?: number;
  comment?: string;
  hidden?: boolean;
}

/**
 * How the workbook recalculates. `autoNoTable` is Excel's third setting:
 * automatic for everything except data tables, which are the expensive case.
 */
export type CalcMode = 'auto' | 'autoNoTable' | 'manual';

export class Workbook {
  readonly sheets: Sheet[] = [];
  readonly styles = new StyleTable();
  definedNames: DefinedName[] = [];
  dateSystem: 1900 | 1904 = 1900;
  /** Calculation mode, mirroring Excel's `calcPr`. */
  calcMode: CalcMode = 'auto';
  fullCalcOnLoad = false;
  /** Raw parts we did not model, preserved verbatim on save. */
  preserved: Record<string, Uint8Array> = {};
  /** Extracted VBA project, kept byte-identical so macros survive a round trip. */
  vbaProject?: Uint8Array;

  private nextSheetId = 1;

  addSheet(name: string, at?: number): Sheet {
    const unique = this.uniqueSheetName(name);
    const sheet = new Sheet(unique, this.nextSheetId++);
    if (at === undefined || at >= this.sheets.length) this.sheets.push(sheet);
    else this.sheets.splice(Math.max(0, at), 0, sheet);
    return sheet;
  }

  getSheet(name: string): Sheet | undefined {
    // Excel sheet names are case-insensitive for lookup but case-preserving.
    const lower = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === lower);
  }

  getSheetById(id: number): Sheet | undefined {
    return this.sheets.find((s) => s.id === id);
  }

  sheetIndex(sheet: Sheet): number {
    return this.sheets.indexOf(sheet);
  }

  removeSheet(name: string): boolean {
    const i = this.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
    if (i < 0) return false;
    this.sheets.splice(i, 1);
    return true;
  }

  /** Excel appends " (2)", " (3)" ... to keep sheet names unique. */
  private uniqueSheetName(base: string): string {
    if (!this.getSheet(base)) return base;
    for (let n = 2; ; n++) {
      const candidate = `${base} (${n})`;
      if (!this.getSheet(candidate)) return candidate;
    }
  }

  get totalCells(): number {
    let n = 0;
    for (const s of this.sheets) n += s.cellCount;
    return n;
  }
}

export type { CellValue };
