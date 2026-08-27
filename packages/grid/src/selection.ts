/**
 * Selection model and keyboard navigation.
 *
 * Spreadsheet keyboard behaviour is muscle memory, and it is unforgiving: an
 * Excel user hits Ctrl+Down expecting to land on the last row of the block and
 * notices immediately when it lands somewhere else. So the rules below are the
 * real ones rather than plausible ones, including the parts that look odd out of
 * context:
 *
 *   Ctrl+Arrow is not "jump far". It is three different behaviours chosen by
 *   what is under and next to the cursor - run to the end of a block, skip a gap
 *   to the next block, or run to the sheet edge - and using the wrong one is the
 *   most commonly noticed mis-implementation in every spreadsheet clone.
 *
 *   End is a MODE, not a movement. Pressing End and then an arrow does what
 *   Ctrl+Arrow does; End then Home goes to the last used cell. Mapping End
 *   straight to "end of row" is the second most commonly noticed one.
 *
 *   Tab and Enter navigate INSIDE a selected range when one exists, wrapping at
 *   its edges instead of leaving it, and Enter after a run of Tabs returns to
 *   the column the run started in. This is what makes keyboard data entry into a
 *   highlighted block work at all.
 *
 * This module owns no document and performs no edits. It answers "where is the
 * cursor now", and the host decides what that means.
 */

import { MAX_COLS, MAX_ROWS, type Sheet } from '@mirrorz/core';

export interface CellPos {
  readonly row: number;
  readonly col: number;
}

/** Inclusive rectangle, always normalised so top <= bottom and left <= right. */
export interface GridRange {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export type Direction = 'up' | 'down' | 'left' | 'right';

export interface SelectionState {
  readonly active: CellPos;
  /** The corner a Shift+click or Shift+arrow extends away from. */
  readonly anchor: CellPos;
  readonly ranges: readonly GridRange[];
  /** Which range the active cell lives in, for Tab/Enter cycling. */
  readonly activeRange: number;
  readonly endMode: boolean;
}

/**
 * What navigation needs to know about the data. Deliberately narrow: the
 * selection model must work over a test double as easily as over a real Sheet,
 * and it must never be tempted to mutate anything.
 */
export interface SelectionSource {
  isEmpty(row: number, col: number): boolean;
  /** Last used row and column, or -1 when the sheet is empty. */
  readonly lastRow: number;
  readonly lastCol: number;
  isRowHidden?(row: number): boolean;
  isColHidden?(col: number): boolean;
  /** Merged region containing a cell, so the cursor treats one as a unit. */
  mergeAt?(row: number, col: number): GridRange | undefined;
}

/** Adapt a core Sheet. Bounds come from the cached extent, never from a scan. */
export function sheetSource(sheet: Sheet): SelectionSource {
  const bounds = sheet.bounds();
  return {
    isEmpty(row, col) {
      const cell = sheet.getCell(row, col);
      return cell === undefined || (cell.value === null && cell.formula === undefined);
    },
    lastRow: bounds ? bounds.maxRow : -1,
    lastCol: bounds ? bounds.maxCol : -1,
    isRowHidden: (row) => sheet.isRowHidden(row),
    isColHidden: (col) => sheet.isColHidden(col),
    mergeAt(row, col) {
      const m = sheet.mergeAt(row, col);
      if (!m) return undefined;
      return { top: m.start.row, left: m.start.col, bottom: m.end.row, right: m.end.col };
    },
  };
}

/** A source over nothing, for a viewport with no data attached yet. */
export const EMPTY_SOURCE: SelectionSource = {
  isEmpty: () => true,
  lastRow: -1,
  lastCol: -1,
};

export function makeRange(a: CellPos, b: CellPos): GridRange {
  return {
    top: Math.min(a.row, b.row),
    left: Math.min(a.col, b.col),
    bottom: Math.max(a.row, b.row),
    right: Math.max(a.col, b.col),
  };
}

export function rangeContainsCell(r: GridRange, row: number, col: number): boolean {
  return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
}

export function rangeCellCount(r: GridRange): number {
  return (r.bottom - r.top + 1) * (r.right - r.left + 1);
}

export function isFullRowRange(r: GridRange): boolean {
  return r.left === 0 && r.right === MAX_COLS - 1;
}

export function isFullColRange(r: GridRange): boolean {
  return r.top === 0 && r.bottom === MAX_ROWS - 1;
}

export interface KeyInput {
  /** KeyboardEvent.key values: 'ArrowUp', 'Tab', 'a', ... */
  readonly key: string;
  readonly shift?: boolean;
  readonly ctrl?: boolean;
  /** Command on macOS; treated exactly as ctrl. */
  readonly meta?: boolean;
  readonly alt?: boolean;
}

export interface SelectionOptions {
  source?: SelectionSource;
  /** Rows moved by Page Up/Down. The viewport supplies this each frame. */
  pageRows?: () => number;
  pageCols?: () => number;
}

const DELTAS: Record<Direction, { dr: number; dc: number }> = {
  up: { dr: -1, dc: 0 },
  down: { dr: 1, dc: 0 },
  left: { dr: 0, dc: -1 },
  right: { dr: 0, dc: 1 },
};

const ARROW_KEYS: Record<string, Direction> = {
  ArrowUp: 'up',
  ArrowDown: 'down',
  ArrowLeft: 'left',
  ArrowRight: 'right',
  Up: 'up',
  Down: 'down',
  Left: 'left',
  Right: 'right',
};

export class Selection {
  private activeCell: CellPos = { row: 0, col: 0 };
  private anchorCell: CellPos = { row: 0, col: 0 };
  private rangeList: GridRange[] = [{ top: 0, left: 0, bottom: 0, right: 0 }];
  private activeIndex = 0;
  private end = false;
  /**
   * Column a run of Tabs started in. Enter returns here, which is the behaviour
   * that makes tabbing across a row and pressing Enter land under the first
   * field rather than under the last one.
   */
  private entryCol: number | null = null;
  private listeners = new Set<(state: SelectionState) => void>();

  source: SelectionSource;
  pageRows: () => number;
  pageCols: () => number;

  constructor(options: SelectionOptions = {}) {
    this.source = options.source ?? EMPTY_SOURCE;
    this.pageRows = options.pageRows ?? (() => 25);
    this.pageCols = options.pageCols ?? (() => 10);
  }

  get state(): SelectionState {
    return {
      active: this.activeCell,
      anchor: this.anchorCell,
      ranges: this.rangeList.slice(),
      activeRange: this.activeIndex,
      endMode: this.end,
    };
  }

  get active(): CellPos {
    return this.activeCell;
  }

  get anchor(): CellPos {
    return this.anchorCell;
  }

  get ranges(): readonly GridRange[] {
    return this.rangeList;
  }

  get endMode(): boolean {
    return this.end;
  }

  onChange(listener: (state: SelectionState) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    if (this.listeners.size === 0) return;
    const state = this.state;
    for (const l of this.listeners) l(state);
  }

  /** Restore a saved selection, e.g. the one stored in a sheet view. */
  setState(state: {
    active: CellPos;
    anchor?: CellPos;
    ranges?: readonly GridRange[];
    activeRange?: number;
  }): void {
    this.activeCell = clampPos(state.active);
    this.anchorCell = clampPos(state.anchor ?? state.active);
    this.rangeList =
      state.ranges && state.ranges.length > 0
        ? state.ranges.map(normalise)
        : [makeRange(this.activeCell, this.activeCell)];
    this.activeIndex = Math.min(Math.max(0, state.activeRange ?? 0), this.rangeList.length - 1);
    this.end = false;
    this.entryCol = null;
    this.emit();
  }

  isSelected(row: number, col: number): boolean {
    for (const r of this.rangeList) if (rangeContainsCell(r, row, col)) return true;
    return false;
  }

  isRowSelected(row: number): boolean {
    for (const r of this.rangeList) if (row >= r.top && row <= r.bottom) return true;
    return false;
  }

  isColSelected(col: number): boolean {
    for (const r of this.rangeList) if (col >= r.left && col <= r.right) return true;
    return false;
  }

  /** True when the whole row is selected, which the header paints differently. */
  isEntireRowSelected(row: number): boolean {
    for (const r of this.rangeList) {
      if (row >= r.top && row <= r.bottom && isFullRowRange(r)) return true;
    }
    return false;
  }

  isEntireColSelected(col: number): boolean {
    for (const r of this.rangeList) {
      if (col >= r.left && col <= r.right && isFullColRange(r)) return true;
    }
    return false;
  }

  // --- mutation -----------------------------------------------------------

  /**
   * Collapse to a single cell. `additive` starts a second range (Ctrl+click),
   * which is how Excel builds a multi-area selection.
   */
  selectCell(row: number, col: number, opts: { additive?: boolean } = {}): void {
    const pos = clampPos({ row, col });
    const merged = this.source.mergeAt?.(pos.row, pos.col);
    const range = merged ?? makeRange(pos, pos);
    // The cursor sits on a merge's anchor, not on whichever sub-cell was clicked.
    const active = merged ? { row: merged.top, col: merged.left } : pos;
    if (opts.additive) {
      this.rangeList = [...this.rangeList, range];
      this.activeIndex = this.rangeList.length - 1;
    } else {
      this.rangeList = [range];
      this.activeIndex = 0;
    }
    this.activeCell = active;
    this.anchorCell = active;
    this.end = false;
    this.entryCol = null;
    this.emit();
  }

  /** Shift+click / drag: grow the active range away from the anchor. */
  extendTo(row: number, col: number): void {
    const pos = clampPos({ row, col });
    let range = makeRange(this.anchorCell, pos);
    range = this.expandOverMerges(range);
    const next = this.rangeList.slice();
    next[this.activeIndex] = range;
    this.rangeList = next;
    this.activeCell = clampPos(this.activeCell);
    this.end = false;
    this.entryCol = null;
    this.emit();
  }

  selectRange(range: GridRange, opts: { additive?: boolean } = {}): void {
    const r = this.expandOverMerges(normalise(range));
    if (opts.additive) {
      this.rangeList = [...this.rangeList, r];
      this.activeIndex = this.rangeList.length - 1;
    } else {
      this.rangeList = [r];
      this.activeIndex = 0;
    }
    this.activeCell = { row: r.top, col: r.left };
    this.anchorCell = this.activeCell;
    this.end = false;
    this.entryCol = null;
    this.emit();
  }

  selectRow(row: number, opts: { additive?: boolean; extend?: boolean } = {}): void {
    const r = clampRow(row);
    if (opts.extend) {
      const range: GridRange = {
        top: Math.min(this.anchorCell.row, r),
        bottom: Math.max(this.anchorCell.row, r),
        left: 0,
        right: MAX_COLS - 1,
      };
      const next = this.rangeList.slice();
      next[this.activeIndex] = range;
      this.rangeList = next;
      this.emit();
      return;
    }
    const range: GridRange = { top: r, bottom: r, left: 0, right: MAX_COLS - 1 };
    if (opts.additive) {
      this.rangeList = [...this.rangeList, range];
      this.activeIndex = this.rangeList.length - 1;
    } else {
      this.rangeList = [range];
      this.activeIndex = 0;
    }
    this.activeCell = { row: r, col: 0 };
    this.anchorCell = this.activeCell;
    this.end = false;
    this.entryCol = null;
    this.emit();
  }

  selectCol(col: number, opts: { additive?: boolean; extend?: boolean } = {}): void {
    const c = clampCol(col);
    if (opts.extend) {
      const range: GridRange = {
        left: Math.min(this.anchorCell.col, c),
        right: Math.max(this.anchorCell.col, c),
        top: 0,
        bottom: MAX_ROWS - 1,
      };
      const next = this.rangeList.slice();
      next[this.activeIndex] = range;
      this.rangeList = next;
      this.emit();
      return;
    }
    const range: GridRange = { left: c, right: c, top: 0, bottom: MAX_ROWS - 1 };
    if (opts.additive) {
      this.rangeList = [...this.rangeList, range];
      this.activeIndex = this.rangeList.length - 1;
    } else {
      this.rangeList = [range];
      this.activeIndex = 0;
    }
    this.activeCell = { row: 0, col: c };
    this.anchorCell = this.activeCell;
    this.end = false;
    this.entryCol = null;
    this.emit();
  }

  selectAll(): void {
    this.rangeList = [{ top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 }];
    this.activeIndex = 0;
    this.anchorCell = this.activeCell;
    this.end = false;
    this.entryCol = null;
    this.emit();
  }

  /** Drop every range but the active one, keeping the cursor where it is. */
  collapseToActive(): void {
    this.rangeList = [makeRange(this.activeCell, this.activeCell)];
    this.activeIndex = 0;
    this.anchorCell = this.activeCell;
    this.emit();
  }

  // --- movement -----------------------------------------------------------

  moveTo(row: number, col: number, extend = false): void {
    if (extend) {
      const pos = clampPos({ row, col });
      this.activeCell = pos;
      const next = this.rangeList.slice();
      next[this.activeIndex] = this.expandOverMerges(makeRange(this.anchorCell, pos));
      this.rangeList = next;
      this.end = false;
      this.entryCol = null;
      this.emit();
    } else {
      this.selectCell(row, col);
    }
  }

  move(direction: Direction, opts: { extend?: boolean; jump?: boolean } = {}): void {
    const target = opts.jump
      ? this.edgeTarget(this.activeCell, direction)
      : this.step(this.activeCell, direction);
    this.moveTo(target.row, target.col, opts.extend === true);
  }

  /** One cell in a direction, skipping hidden rows and columns as Excel does. */
  step(from: CellPos, direction: Direction): CellPos {
    const { dr, dc } = DELTAS[direction];
    let row = from.row;
    let col = from.col;
    // A merged region moves as a unit: leaving it starts from its far edge.
    const merged = this.source.mergeAt?.(row, col);
    if (merged) {
      if (dr > 0) row = merged.bottom;
      else if (dr < 0) row = merged.top;
      if (dc > 0) col = merged.right;
      else if (dc < 0) col = merged.left;
    }
    row += dr;
    col += dc;
    if (dr !== 0) row = this.skipHiddenRows(row, dr);
    if (dc !== 0) col = this.skipHiddenCols(col, dc);
    if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) return from;
    return { row, col };
  }

  private skipHiddenRows(row: number, dr: number): number {
    const hidden = this.source.isRowHidden;
    if (!hidden) return row;
    let r = row;
    // Bounded so a sheet with every row hidden cannot spin for a million steps.
    for (let guard = 0; guard < 4096 && r >= 0 && r < MAX_ROWS && hidden(r); guard++) r += dr;
    return r;
  }

  private skipHiddenCols(col: number, dc: number): number {
    const hidden = this.source.isColHidden;
    if (!hidden) return col;
    let c = col;
    for (let guard = 0; guard < 4096 && c >= 0 && c < MAX_COLS && hidden(c); guard++) c += dc;
    return c;
  }

  /**
   * Ctrl+Arrow. Three cases, chosen by what is under and next to the cursor:
   *
   *   in a block, next cell also filled -> last filled cell of this run
   *   in a block, next cell empty       -> first filled cell after the gap
   *   on an empty cell                  -> first filled cell in that direction
   *
   * and when nothing is found, the sheet edge. Scanning stops at the used extent
   * rather than walking to row 1,048,575, so a jump off the end of the data is
   * O(size of the data), not O(size of the grid).
   */
  edgeTarget(from: CellPos, direction: Direction): CellPos {
    const { dr, dc } = DELTAS[direction];
    const limit = this.scanLimit(direction);
    const distance = dr !== 0 ? Math.abs(limit - from.row) : Math.abs(limit - from.col);
    if (distance === 0) return from;

    const at = (n: number): CellPos =>
      dr !== 0 ? { row: from.row + n * dr, col: from.col } : { row: from.row, col: from.col + n * dc };

    const src = this.source;
    const currentFilled = !src.isEmpty(from.row, from.col);
    const nextFilled = !src.isEmpty(at(1).row, at(1).col);

    if (currentFilled && nextFilled) {
      let n = 1;
      while (n < distance && !src.isEmpty(at(n + 1).row, at(n + 1).col)) n++;
      return at(n);
    }

    for (let n = 1; n <= distance; n++) {
      const p = at(n);
      if (!src.isEmpty(p.row, p.col)) return p;
    }
    return dr !== 0 ? { row: edgeIndex(direction), col: from.col } : { row: from.row, col: edgeIndex(direction) };
  }

  /**
   * How far a scan may usefully go: one step past the used extent, because
   * everything beyond it is empty by definition.
   */
  private scanLimit(direction: Direction): number {
    const src = this.source;
    switch (direction) {
      case 'up':
        return 0;
      case 'left':
        return 0;
      case 'down':
        return Math.min(MAX_ROWS - 1, Math.max(0, src.lastRow + 1));
      case 'right':
        return Math.min(MAX_COLS - 1, Math.max(0, src.lastCol + 1));
    }
  }

  /** Ctrl+End: the bottom-right of the used range, or A1 for an empty sheet. */
  lastUsedCell(): CellPos {
    const src = this.source;
    if (src.lastRow < 0 || src.lastCol < 0) return { row: 0, col: 0 };
    return { row: src.lastRow, col: src.lastCol };
  }

  /**
   * Tab / Enter inside a selected block.
   *
   * With a real range selected the cursor cycles inside it and wraps at the
   * edges instead of escaping, moving on to the next range when it runs out -
   * this is the whole point of selecting a block before typing into it.
   */
  advance(kind: 'tab' | 'enter', backwards = false): void {
    const multi =
      this.rangeList.length > 1 || (this.rangeList[0] !== undefined && rangeCellCount(this.rangeList[0]) > 1);

    if (multi) {
      this.advanceWithinSelection(kind, backwards);
      return;
    }

    if (kind === 'tab') {
      if (this.entryCol === null) this.entryCol = this.activeCell.col;
      const target = this.step(this.activeCell, backwards ? 'left' : 'right');
      const col = this.entryCol;
      this.selectCell(target.row, target.col);
      this.entryCol = col;
      return;
    }
    // Enter returns to the column the run of Tabs started in, one row on.
    const homeCol = this.entryCol ?? this.activeCell.col;
    const target = this.step({ row: this.activeCell.row, col: homeCol }, backwards ? 'up' : 'down');
    this.selectCell(target.row, target.col);
  }

  private advanceWithinSelection(kind: 'tab' | 'enter', backwards: boolean): void {
    const ranges = this.rangeList;
    let index = this.activeIndex;
    let range = ranges[index];
    if (!range) return;
    let { row, col } = this.activeCell;
    if (!rangeContainsCell(range, row, col)) {
      row = range.top;
      col = range.left;
    }

    const forward = !backwards;
    if (kind === 'tab') {
      col += forward ? 1 : -1;
      if (col > range.right) {
        col = range.left;
        row += 1;
      } else if (col < range.left) {
        col = range.right;
        row -= 1;
      }
      if (row > range.bottom) {
        [index, range] = this.nextRange(index, 1);
        row = range.top;
        col = range.left;
      } else if (row < range.top) {
        [index, range] = this.nextRange(index, -1);
        row = range.bottom;
        col = range.right;
      }
    } else {
      row += forward ? 1 : -1;
      if (row > range.bottom) {
        row = range.top;
        col += 1;
      } else if (row < range.top) {
        row = range.bottom;
        col -= 1;
      }
      if (col > range.right) {
        [index, range] = this.nextRange(index, 1);
        row = range.top;
        col = range.left;
      } else if (col < range.left) {
        [index, range] = this.nextRange(index, -1);
        row = range.bottom;
        col = range.right;
      }
    }

    this.activeIndex = index;
    this.activeCell = { row, col };
    this.end = false;
    this.emit();
  }

  private nextRange(index: number, step: number): [number, GridRange] {
    const n = this.rangeList.length;
    const next = ((index + step) % n + n) % n;
    return [next, this.rangeList[next] as GridRange];
  }

  // --- keyboard -----------------------------------------------------------

  /**
   * Apply one key. Returns true when the key was consumed, so the host can call
   * preventDefault only for keys that actually mean something here.
   */
  keyDown(event: KeyInput): boolean {
    const ctrl = event.ctrl === true || event.meta === true;
    const shift = event.shift === true;
    const key = event.key;

    const direction = ARROW_KEYS[key];
    if (direction) {
      // End mode makes a plain arrow behave as Ctrl+Arrow, then clears itself.
      const jump = ctrl || this.end;
      this.end = false;
      this.move(direction, { extend: shift, jump });
      return true;
    }

    switch (key) {
      case 'Tab':
        this.end = false;
        this.advance('tab', shift);
        return true;

      case 'Enter':
      case 'Return':
        this.end = false;
        this.advance('enter', shift);
        return true;

      case 'Home': {
        if (this.end) {
          // End then Home is Excel's other spelling of Ctrl+End.
          this.end = false;
          const last = this.lastUsedCell();
          this.moveTo(last.row, last.col, shift);
          return true;
        }
        const row = ctrl ? 0 : this.activeCell.row;
        this.moveTo(row, 0, shift);
        return true;
      }

      case 'End': {
        if (ctrl) {
          this.end = false;
          const last = this.lastUsedCell();
          this.moveTo(last.row, last.col, shift);
          return true;
        }
        // A bare End toggles End mode rather than moving anything.
        this.end = !this.end;
        this.emit();
        return true;
      }

      case 'PageDown': {
        this.end = false;
        if (event.alt) {
          const step = Math.max(1, this.pageCols());
          this.moveTo(this.activeCell.row, Math.min(MAX_COLS - 1, this.activeCell.col + step), shift);
        } else {
          const step = Math.max(1, this.pageRows());
          this.moveTo(Math.min(MAX_ROWS - 1, this.activeCell.row + step), this.activeCell.col, shift);
        }
        return true;
      }

      case 'PageUp': {
        this.end = false;
        if (event.alt) {
          const step = Math.max(1, this.pageCols());
          this.moveTo(this.activeCell.row, Math.max(0, this.activeCell.col - step), shift);
        } else {
          const step = Math.max(1, this.pageRows());
          this.moveTo(Math.max(0, this.activeCell.row - step), this.activeCell.col, shift);
        }
        return true;
      }

      case ' ': {
        // Ctrl+Space selects the column, Shift+Space the row, both the sheet.
        if (ctrl && shift) {
          this.selectAll();
          return true;
        }
        if (ctrl) {
          this.selectCol(this.activeCell.col);
          return true;
        }
        if (shift) {
          this.selectRow(this.activeCell.row);
          return true;
        }
        return false;
      }

      case 'a':
      case 'A':
        if (ctrl) {
          this.selectAll();
          return true;
        }
        return false;

      case 'Escape':
        if (this.end) {
          this.end = false;
          this.emit();
          return true;
        }
        return false;

      default:
        return false;
    }
  }

  /**
   * Grow a range so it never cuts a merged region in half, which is what Excel
   * does the moment a selection touches one.
   */
  private expandOverMerges(range: GridRange): GridRange {
    const mergeAt = this.source.mergeAt;
    if (!mergeAt) return range;
    let { top, left, bottom, right } = range;
    // Two passes is enough in practice; each pass can only grow the rectangle,
    // and a fixpoint loop over a pathological merge layout is not worth the risk
    // of spinning inside a paint.
    for (let pass = 0; pass < 2; pass++) {
      let changed = false;
      const area = (bottom - top + 1) * (right - left + 1);
      if (area > 65_536) break;
      for (let r = top; r <= bottom; r++) {
        for (let c = left; c <= right; c++) {
          const m = mergeAt(r, c);
          if (!m) continue;
          if (m.top < top) (top = m.top), (changed = true);
          if (m.left < left) (left = m.left), (changed = true);
          if (m.bottom > bottom) (bottom = m.bottom), (changed = true);
          if (m.right > right) (right = m.right), (changed = true);
        }
      }
      if (!changed) break;
    }
    return { top, left, bottom, right };
  }
}

function edgeIndex(direction: Direction): number {
  switch (direction) {
    case 'up':
    case 'left':
      return 0;
    case 'down':
      return MAX_ROWS - 1;
    case 'right':
      return MAX_COLS - 1;
  }
}

function clampRow(row: number): number {
  return Math.min(MAX_ROWS - 1, Math.max(0, Math.floor(row)));
}

function clampCol(col: number): number {
  return Math.min(MAX_COLS - 1, Math.max(0, Math.floor(col)));
}

function clampPos(p: CellPos): CellPos {
  return { row: clampRow(p.row), col: clampCol(p.col) };
}

function normalise(r: GridRange): GridRange {
  return {
    top: clampRow(Math.min(r.top, r.bottom)),
    bottom: clampRow(Math.max(r.top, r.bottom)),
    left: clampCol(Math.min(r.left, r.right)),
    right: clampCol(Math.max(r.left, r.right)),
  };
}
