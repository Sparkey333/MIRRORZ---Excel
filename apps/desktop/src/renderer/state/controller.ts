/**
 * The single writable object the UI talks to.
 *
 * There is no state library here on purpose. The application already has a
 * store - the Document's command log - and putting a second one beside it would
 * mean two sources of truth for the same cells, which is how undo quietly stops
 * working. So this class owns the Document and the Engine, exposes a small
 * immutable snapshot for `useSyncExternalStore`, and every mutation it offers
 * goes through `Document` or `Engine`. Nothing in `components/` may write to a
 * Workbook directly.
 *
 * The snapshot deliberately holds only small, comparable values. Anything
 * derived and expensive - the history rows, the inspector tree, the sheet
 * list's search - is computed in the component from `version`, so a keystroke in
 * a cell does not rebuild every panel.
 */

import {
  DEFAULT_STYLE_ID,
  MAX_COLS,
  MAX_ROWS,
  a1,
  parseEntry,
  type CellData,
  type CellStyle,
  type ChangeOrigin,
  type Document,
  type EntryOptions,
  type HistoryEntry,
  type Scalar,
  type Sheet,
  type SheetVisibility,
  type StyleId,
} from '@mirrorz/core';
import { format } from '@mirrorz/formats';
import type { CellAddr, CellExplanation, Engine, FunctionRegistry } from '@mirrorz/formula';
import {
  normaliseRange,
  singleCell,
  type Selection,
  type SelectionRange,
} from '../model/selection.js';
import type { ThemePreference } from '../model/theme.js';
import type { ColumnOverride } from '../model/import-review.js';

export interface SheetSummary {
  name: string;
  index: number;
  visibility: SheetVisibility;
  tabColor?: string;
  cellCount: number;
  active: boolean;
}

export interface RecalcSummary {
  elapsedMs: number;
  evaluated: number;
  changed: number;
  circular: number;
  at: number;
}

export type PanelName = 'explorer' | 'history' | 'inspector';

export interface PendingImport {
  /** Where the block lands; the anchor of the paste or the import target. */
  anchor: CellAddr;
  rows: string[][];
  overrides: ColumnOverride[];
  headerRow: boolean;
  source: 'paste' | 'csv';
  fileName?: string;
}

export interface AppSnapshot {
  version: number;
  activeSheet: string;
  selection: Selection;
  canUndo: boolean;
  canRedo: boolean;
  undoLabel: string | null;
  redoLabel: string | null;
  headId: number | null;
  calcMode: 'auto' | 'autoNoTable' | 'manual';
  lastRecalc: RecalcSummary | null;
  theme: ThemePreference;
  panels: Readonly<Record<PanelName, boolean>>;
  paletteOpen: boolean;
  paletteMode: 'all' | 'command';
  /**
   * Find and replace is open. This lives here rather than in the toolbar that
   * draws it because the native menu and the keyboard both have to be able to
   * open it, and a useState inside one component is reachable only by that
   * component's own button.
   */
  findOpen: boolean;
  pendingImport: PendingImport | null;
  fileName: string;
  dirty: boolean;
  /** Transient one-line message for the status bar. */
  message: string | null;
}

/** How many addresses a formatting or fill operation will touch before we refuse. */
const MAX_BULK_CELLS = 250_000;

export interface ControllerOptions {
  entryOptions?: EntryOptions;
  /** Injected so tests and the history panel see stable times. */
  now?: () => number;
}

export class AppController {
  readonly doc: Document;
  readonly engine: Engine;
  readonly registry: FunctionRegistry;

  private snapshot: AppSnapshot;
  private listeners = new Set<() => void>();
  private readonly now: () => number;
  private entryOptions: EntryOptions;
  private disposeDocument: () => void;

  constructor(
    doc: Document,
    engine: Engine,
    registry: FunctionRegistry,
    options: ControllerOptions = {},
  ) {
    this.doc = doc;
    this.engine = engine;
    this.registry = registry;
    this.now = options.now ?? (() => Date.now());
    this.entryOptions = options.entryOptions ?? {};

    // Directly on the workbook, not through the log: an empty workbook's first
    // sheet is initial state, and logging it would make jumping to the start of
    // the history delete the sheet the user is looking at.
    if (!doc.workbook.sheets[0]) doc.workbook.addSheet('Sheet1');
    const activeSheet = doc.workbook.sheets[0]!.name;

    this.snapshot = {
      version: 0,
      activeSheet,
      selection: singleCell(activeSheet, 0, 0),
      canUndo: false,
      canRedo: false,
      undoLabel: null,
      redoLabel: null,
      headId: null,
      calcMode: doc.workbook.calcMode,
      lastRecalc: null,
      theme: 'system',
      panels: { explorer: true, history: false, inspector: false },
      paletteOpen: false,
      paletteMode: 'all',
      findOpen: false,
      pendingImport: null,
      fileName: 'Untitled',
      dirty: false,
      message: null,
    };

    // The Document is the source of truth, so anything that writes to it -
    // including a script or a macro that never touched this class - repaints
    // the UI. That is the whole reason the store is not a second copy.
    this.disposeDocument = doc.onChange(() => this.bump({ dirty: true }));
  }

  dispose(): void {
    this.disposeDocument();
    this.listeners.clear();
  }

  // --- store plumbing -----------------------------------------------------

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  };

  getSnapshot = (): AppSnapshot => this.snapshot;

  /** Replace the snapshot and notify. Every mutation ends here. */
  private bump(patch: Partial<AppSnapshot> = {}): void {
    const undo = this.doc.peekUndo();
    const redo = this.doc.peekRedo();
    this.snapshot = {
      ...this.snapshot,
      version: this.snapshot.version + 1,
      canUndo: this.doc.canUndo,
      canRedo: this.doc.canRedo,
      undoLabel: undo?.label ?? null,
      redoLabel: redo?.label ?? null,
      headId: this.doc.history().at(-1)?.id ?? null,
      calcMode: this.doc.workbook.calcMode,
      ...patch,
    };
    for (const listener of this.listeners) listener();
  }

  // --- reads --------------------------------------------------------------

  get workbook() {
    return this.doc.workbook;
  }

  sheet(name = this.snapshot.activeSheet): Sheet | undefined {
    return this.doc.workbook.getSheet(name);
  }

  sheetSummaries(): SheetSummary[] {
    return this.doc.workbook.sheets.map((sheet, index) => {
      const summary: SheetSummary = {
        name: sheet.name,
        index,
        visibility: sheet.visibility,
        cellCount: sheet.cellCount,
        active: sheet.name === this.snapshot.activeSheet,
      };
      if (sheet.tabColor !== undefined) summary.tabColor = sheet.tabColor;
      return summary;
    });
  }

  historyEntries(): HistoryEntry[] {
    return this.doc.allEntries();
  }

  cellAt(addr: CellAddr): CellData | undefined {
    return this.doc.workbook.getSheet(addr.sheet)?.getCell(addr.row, addr.col);
  }

  activeAddr(): CellAddr {
    const { activeSheet, selection } = this.snapshot;
    return { sheet: activeSheet, row: selection.active.row, col: selection.active.col };
  }

  /**
   * What the formula bar shows.
   *
   * A formula shows as its source, and a value shows as the text the user
   * actually typed when we kept one. Showing the inferred value instead is how
   * a spreadsheet convinces someone that their `007` was always `7`.
   */
  editText(addr: CellAddr = this.activeAddr()): string {
    const cell = this.cellAt(addr);
    if (!cell) return '';
    if (cell.formula !== undefined) return `=${cell.formula}`;
    if (cell.literal !== undefined) return cell.literal;
    if (cell.value === null) return '';
    if (typeof cell.value === 'boolean') return cell.value ? 'TRUE' : 'FALSE';
    return String(cell.value);
  }

  /** The rendered text for a cell, used by the fallback grid and the inspector. */
  displayText(addr: CellAddr): string {
    const cell = this.cellAt(addr);
    if (!cell) return '';
    const style = this.styleOf(addr);
    const code = style.numFmt ?? 'General';
    try {
      return format(cell.value, code, { dateSystem: this.doc.workbook.dateSystem }).text;
    } catch {
      return String(cell.value ?? '');
    }
  }

  styleIdOf(addr: CellAddr): StyleId {
    return this.doc.workbook.getSheet(addr.sheet)?.getStyle(addr.row, addr.col) ?? DEFAULT_STYLE_ID;
  }

  styleOf(addr: CellAddr): CellStyle {
    return this.doc.workbook.styles.get(this.styleIdOf(addr));
  }

  explain(addr: CellAddr): CellExplanation {
    return this.engine.explain(addr);
  }

  /**
   * One-line descriptions for the completion list, straight from the registry.
   *
   * Drawn from the registry rather than a table in the UI so the editor cannot
   * describe a function the engine does not have, or describe one it has
   * differently from the way it behaves.
   */
  functionSummaries(): Record<string, string> {
    const summaries: Record<string, string> = {};
    for (const name of this.registry.names()) {
      const summary = this.registry.get(name)?.summary;
      if (summary !== undefined) summaries[name] = summary;
    }
    return summaries;
  }

  // --- selection and navigation ------------------------------------------

  setActiveSheet(name: string): void {
    if (!this.doc.workbook.getSheet(name)) return;
    this.bump({ activeSheet: name, selection: singleCell(name, 0, 0) });
  }

  setSelection(selection: Selection): void {
    const patch: Partial<AppSnapshot> = { selection };
    if (selection.sheet !== this.snapshot.activeSheet) patch.activeSheet = selection.sheet;
    this.bump(patch);
  }

  selectCell(row: number, col: number, sheet = this.snapshot.activeSheet): void {
    this.setSelection(singleCell(sheet, row, col));
  }

  /** Navigate to a cell on any sheet: the click target of an inspector root. */
  goTo(addr: CellAddr): void {
    if (!this.doc.workbook.getSheet(addr.sheet)) return;
    this.bump({
      activeSheet: addr.sheet,
      selection: singleCell(addr.sheet, addr.row, addr.col),
      panels: { ...this.snapshot.panels, inspector: true },
    });
  }

  // --- editing ------------------------------------------------------------

  /**
   * Commit what the user typed into a cell.
   *
   * Type inference happens here, once, and keeps the literal text when it
   * differs from the stored value, so the conversion stays reversible.
   */
  commitEntry(text: string, addr: CellAddr = this.activeAddr()): void {
    const label = `Edit ${addr.sheet}!${a1(addr.row, addr.col)}`;

    if (text.startsWith('=')) {
      const formula = text.slice(1);
      const existing = this.cellAt(addr);
      const data: CellData = { ...existing, value: null, formula };
      delete data.literal;
      this.applyRecalc(this.engine.setCell(addr, data, label));
      return;
    }

    const parsed = parseEntry(text, this.entryOptions);
    const existing = this.cellAt(addr);
    const data: CellData = { value: parsed.value };
    if (existing?.style !== undefined) data.style = existing.style;
    if (parsed.literal !== undefined) data.literal = parsed.literal;
    // An implied format is applied only where the cell had none, so retyping a
    // date into a column someone deliberately formatted does not reset it.
    if (parsed.impliedFormat !== undefined && existing?.style === undefined) {
      data.style = this.doc.workbook.styles.intern({ numFmt: parsed.impliedFormat });
    }
    const empty = parsed.value === null && data.style === undefined;
    this.applyRecalc(this.engine.setCell(addr, empty ? undefined : data, label));
  }

  clearSelection(): void {
    const cells = this.selectedAddresses();
    if (cells.length === 0) return;
    this.doc.transact({ label: `Clear ${cells.length} cells`, origin: 'user' }, () => {
      for (const addr of cells) this.doc.setCell(addr.sheet, addr.row, addr.col, undefined);
    });
    this.applyRecalc(this.engine.recalculate(cells));
  }

  /**
   * Apply a formatting change to the selection.
   *
   * The patch is merged onto each cell's existing style rather than replacing
   * it, so making a range bold does not also reset everyone's number format -
   * which is the behaviour every "apply style" implementation gets wrong once.
   */
  applyStyle(patch: Partial<CellStyle>, label = 'Format cells'): void {
    const addresses = this.selectedAddresses({ includeEmpty: true });
    if (addresses.length === 0) return;
    if (addresses.length > MAX_BULK_CELLS) {
      this.bump({ message: `Selection is too large to format (${addresses.length} cells)` });
      return;
    }
    this.doc.transact({ label, origin: 'user' }, () => {
      for (const addr of addresses) {
        const sheet = this.doc.workbook.getSheet(addr.sheet);
        if (!sheet) continue;
        const next = this.doc.workbook.styles.derive(sheet.getStyle(addr.row, addr.col), patch);
        this.doc.setStyle(addr.sheet, addr.row, addr.col, next);
      }
    });
  }

  /** Read a style property that is uniform across the selection, else undefined. */
  uniformStyle<T>(pick: (style: CellStyle) => T): T | undefined {
    const addresses = this.selectedAddresses({ includeEmpty: true, cap: 500 });
    if (addresses.length === 0) return undefined;
    let first: T | undefined;
    for (let i = 0; i < addresses.length; i++) {
      const value = pick(this.styleOf(addresses[i]!));
      if (i === 0) first = value;
      else if (JSON.stringify(value) !== JSON.stringify(first)) return undefined;
    }
    return first;
  }

  setNumberFormat(code: string | undefined): void {
    this.applyStyle(
      { numFmt: code === undefined || code.toLowerCase() === 'general' ? undefined : code },
      code === undefined ? 'Clear number format' : `Format as ${code}`,
    );
  }

  // --- structure ----------------------------------------------------------

  insertRows(at: number, count = 1): void {
    const sheet = this.sheet();
    if (!sheet || count < 1) return;
    this.shiftRows(sheet, at, count, 'Insert rows');
  }

  deleteRows(at: number, count = 1): void {
    const sheet = this.sheet();
    if (!sheet || count < 1) return;
    this.shiftRows(sheet, at, -count, 'Delete rows');
  }

  insertCols(at: number, count = 1): void {
    const sheet = this.sheet();
    if (!sheet || count < 1) return;
    this.shiftCols(sheet, at, count, 'Insert columns');
  }

  deleteCols(at: number, count = 1): void {
    const sheet = this.sheet();
    if (!sheet || count < 1) return;
    this.shiftCols(sheet, at, -count, 'Delete columns');
  }

  /**
   * Move every cell at or below `at` by `delta` rows.
   *
   * References inside formulas are NOT rewritten: doing that correctly needs an
   * AST rewrite in the formula package, and a half-correct version that fixes
   * some references and not others would be worse than none. Cells that hold
   * formulas in the shifted region are reported so the user is told rather than
   * finding out later.
   */
  private shiftRows(sheet: Sheet, at: number, delta: number, label: string): void {
    const bounds = sheet.bounds();
    if (!bounds) return;
    // Excel refuses an insert that would push a used cell off the bottom rather
    // than dropping it, and so do we. This is not hypothetical: selecting a
    // column header selects 1,048,576 rows, and "insert rows" over that
    // selection asks to shift every cell a million rows down - past the end of
    // the sheet, where the addresses are not addresses any more.
    if (delta > 0 && bounds.maxRow + delta >= MAX_ROWS) {
      this.bump({ message: 'That would push data off the bottom of the sheet' });
      return;
    }
    const moved: { row: number; col: number; cell: CellData }[] = [];
    let formulas = 0;
    for (const entry of sheet.entries()) {
      if (entry.row < at) continue;
      if (delta < 0 && entry.row < at - delta) continue;
      moved.push(entry);
      if (entry.cell.formula !== undefined) formulas++;
    }
    const removed =
      delta < 0
        ? [...sheet.entries()].filter((e) => e.row >= at && e.row < at - delta)
        : [];

    this.doc.transact({ label, origin: 'user' }, () => {
      for (const entry of [...moved, ...removed]) {
        this.doc.setCell(sheet.name, entry.row, entry.col, undefined);
      }
      for (const entry of moved) {
        this.doc.setCell(sheet.name, entry.row + delta, entry.col, entry.cell);
      }
    });
    this.reindexAndRecalc(formulas, label);
  }

  private shiftCols(sheet: Sheet, at: number, delta: number, label: string): void {
    const bounds = sheet.bounds();
    if (!bounds) return;
    if (delta > 0 && bounds.maxCol + delta >= MAX_COLS) {
      this.bump({ message: 'That would push data off the right of the sheet' });
      return;
    }
    const moved: { row: number; col: number; cell: CellData }[] = [];
    let formulas = 0;
    for (const entry of sheet.entries()) {
      if (entry.col < at) continue;
      if (delta < 0 && entry.col < at - delta) continue;
      moved.push(entry);
      if (entry.cell.formula !== undefined) formulas++;
    }
    const removed =
      delta < 0 ? [...sheet.entries()].filter((e) => e.col >= at && e.col < at - delta) : [];

    this.doc.transact({ label, origin: 'user' }, () => {
      for (const entry of [...moved, ...removed]) {
        this.doc.setCell(sheet.name, entry.row, entry.col, undefined);
      }
      for (const entry of moved) {
        this.doc.setCell(sheet.name, entry.row, entry.col + delta, entry.cell);
      }
    });
    this.reindexAndRecalc(formulas, label);
  }

  private reindexAndRecalc(movedFormulas: number, label: string): void {
    this.engine.indexWorkbook();
    this.applyRecalc(this.engine.recalculateAll());
    if (movedFormulas > 0) {
      this.bump({
        message: `${label}: ${movedFormulas} moved formula${movedFormulas === 1 ? '' : 's'} still point at their original cells`,
      });
    }
  }

  // --- sheets -------------------------------------------------------------

  addSheet(name?: string): void {
    const base = name ?? `Sheet${this.doc.workbook.sheets.length + 1}`;
    const sheet = this.doc.addSheet(base, undefined, {
      label: `Add sheet ${base}`,
      origin: 'user',
      timestamp: this.now(),
    });
    this.bump({ activeSheet: sheet.name, selection: singleCell(sheet.name, 0, 0) });
  }

  renameSheet(from: string, to: string): boolean {
    const trimmed = to.trim();
    if (trimmed === '' || trimmed === from) return false;
    if (this.doc.workbook.getSheet(trimmed)) {
      this.bump({ message: `A sheet called ${trimmed} already exists` });
      return false;
    }
    this.doc.renameSheet(from, trimmed, {
      label: `Rename ${from} to ${trimmed}`,
      origin: 'user',
      timestamp: this.now(),
    });
    this.bump(this.snapshot.activeSheet === from ? { activeSheet: trimmed } : {});
    return true;
  }

  removeSheet(name: string): void {
    const visible = this.doc.workbook.sheets.filter((s) => s.visibility === 'visible');
    if (visible.length <= 1 && visible[0]?.name === name) {
      this.bump({ message: 'A workbook needs at least one visible sheet' });
      return;
    }
    const index = this.doc.workbook.sheets.findIndex((s) => s.name === name);
    this.doc.removeSheet(name, { label: `Delete sheet ${name}`, origin: 'user', timestamp: this.now() });
    if (this.snapshot.activeSheet === name) {
      const next = this.doc.workbook.sheets[Math.min(index, this.doc.workbook.sheets.length - 1)];
      if (next) this.bump({ activeSheet: next.name, selection: singleCell(next.name, 0, 0) });
      else this.bump();
    } else this.bump();
  }

  moveSheet(name: string, to: number): void {
    this.doc.moveSheet(name, to, { label: `Move sheet ${name}`, origin: 'user', timestamp: this.now() });
  }

  setSheetVisibility(name: string, visibility: SheetVisibility): void {
    this.doc.setSheetVisibility(name, visibility, {
      label: `${visibility === 'visible' ? 'Show' : 'Hide'} ${name}`,
      origin: 'user',
      timestamp: this.now(),
    });
    if (visibility !== 'visible' && this.snapshot.activeSheet === name) {
      const next = this.doc.workbook.sheets.find((s) => s.visibility === 'visible');
      if (next) this.bump({ activeSheet: next.name, selection: singleCell(next.name, 0, 0) });
    }
  }

  /**
   * Colour a sheet tab.
   *
   * Goes through the command log like every other edit. It used to be written
   * straight onto the Sheet, because core had no `sheetColor` change kind; that
   * made it the one visible edit in the application that undo silently ignored,
   * which is worse than it sounds - an undo that skips an edit does not just fail
   * to reverse it, it leaves the user's mental model of the history wrong.
   */
  setSheetColor(name: string, color: string | undefined): void {
    if (!this.doc.workbook.getSheet(name)) return;
    this.doc.setSheetColor(name, color, {
      label: color === undefined ? `Clear colour of ${name}` : `Colour ${name}`,
      origin: 'user',
      timestamp: this.now(),
    });
  }

  // --- history ------------------------------------------------------------

  undo(): void {
    const entry = this.doc.undo();
    if (!entry) return;
    if (entry.barrier) {
      this.bump({ message: `"${entry.label}" cannot be reversed exactly` });
      return;
    }
    this.bump({ dirty: true });
  }

  redo(entryId?: number): void {
    this.doc.redo(entryId);
    this.bump({ dirty: true });
  }

  /**
   * Jump to any point in the history, including onto an abandoned branch.
   *
   * `Document.jumpTo` does not fire the change listener - it moves the head
   * without emitting - so the repaint has to be forced here.
   */
  jumpTo(entryId: number | null): void {
    if (!this.doc.jumpTo(entryId)) return;
    this.bump({ dirty: true });
  }

  // --- calculation --------------------------------------------------------

  /**
   * The one workbook property the UI writes without going through a command,
   * because there is no command for it: calculation mode is a setting on the
   * workbook rather than an edit to its contents, and it has no inverse worth
   * putting on the undo stack.
   *
   * It is still a change that gets written into the saved file, so it must mark
   * the document dirty. Without that the close prompt never appears, the change
   * is never saved, and switching to Manual silently reverts the next time the
   * file is opened.
   */
  setCalcMode(mode: 'auto' | 'autoNoTable' | 'manual'): void {
    if (this.doc.workbook.calcMode === mode) return;
    this.doc.workbook.calcMode = mode;
    this.bump({ dirty: true });
  }

  recalculateAll(): void {
    this.applyRecalc(this.engine.recalculateAll());
  }

  private applyRecalc(result: { elapsedMs: number; evaluated: unknown[]; changed: unknown[]; circular: unknown[] } | undefined): void {
    if (!result) {
      this.bump({ dirty: true });
      return;
    }
    this.bump({
      dirty: true,
      lastRecalc: {
        elapsedMs: result.elapsedMs,
        evaluated: result.evaluated.length,
        changed: result.changed.length,
        circular: result.circular.length,
        at: this.now(),
      },
      message:
        result.circular.length > 0
          ? `${result.circular.length} cell${result.circular.length === 1 ? '' : 's'} form a circular reference`
          : null,
    });
  }

  // --- find, sort, filter -------------------------------------------------

  /** Cell addresses matching a search, in row-major order across the sheet. */
  find(query: string, options: { matchCase?: boolean; whole?: boolean; formulas?: boolean } = {}): CellAddr[] {
    const sheet = this.sheet();
    if (!sheet || query === '') return [];
    const needle = options.matchCase ? query : query.toLowerCase();
    const hits: CellAddr[] = [];
    for (const { row, col, cell } of sheet.entries()) {
      const haystackRaw =
        options.formulas && cell.formula !== undefined
          ? `=${cell.formula}`
          : this.displayText({ sheet: sheet.name, row, col });
      const haystack = options.matchCase ? haystackRaw : haystackRaw.toLowerCase();
      const hit = options.whole ? haystack === needle : haystack.includes(needle);
      if (hit) hits.push({ sheet: sheet.name, row, col });
    }
    return hits;
  }

  replaceAll(query: string, replacement: string, options: { matchCase?: boolean; whole?: boolean } = {}): number {
    const hits = this.find(query, options);
    if (hits.length === 0) return 0;
    const sheet = this.sheet();
    if (!sheet) return 0;
    this.doc.transact({ label: `Replace ${hits.length} cells`, origin: 'user', timestamp: this.now() }, () => {
      for (const addr of hits) {
        const current = this.editText(addr);
        const next = options.matchCase
          ? current.split(query).join(replacement)
          : replaceInsensitive(current, query, replacement);
        const parsed = parseEntry(next, this.entryOptions);
        const existing = this.cellAt(addr);
        const data: CellData = { value: parsed.value };
        if (existing?.style !== undefined) data.style = existing.style;
        if (parsed.literal !== undefined) data.literal = parsed.literal;
        this.doc.setCell(addr.sheet, addr.row, addr.col, data);
      }
    });
    this.applyRecalc(this.engine.recalculate(hits));
    return hits.length;
  }

  /**
   * Sort the selected block by one of its columns.
   *
   * Rows move as units, which is the part people expect and the part a naive
   * column sort gets wrong - sorting one column of a table and leaving the rest
   * in place silently destroys every row of the data.
   */
  sortSelection(byColumn: number, direction: 'asc' | 'desc' = 'asc', hasHeader = false): void {
    const sheet = this.sheet();
    if (!sheet) return;
    const raw = normaliseRange(this.snapshot.selection.ranges[0]!);
    // Clamped to the used range, exactly as `selectedAddresses` is. Selecting a
    // whole column selects 1,048,576 addresses; sorting the empty million below
    // the data would build a million row records, write a million empty cells
    // into one transaction, and hang the window for minutes. The user meant the
    // rows with data in them.
    const bounds = sheet.bounds();
    if (!bounds) return;
    const range = {
      start: raw.start,
      end: {
        row: Math.min(raw.end.row, Math.max(bounds.maxRow, raw.start.row)),
        col: Math.min(raw.end.col, Math.max(bounds.maxCol, raw.start.col)),
      },
    };
    const firstRow = range.start.row + (hasHeader ? 1 : 0);
    if (firstRow > range.end.row) return;

    const rows: { key: Scalar; cells: (CellData | undefined)[] }[] = [];
    for (let r = firstRow; r <= range.end.row; r++) {
      const cells: (CellData | undefined)[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) cells.push(sheet.getCell(r, c));
      rows.push({ key: sheet.getValue(r, byColumn), cells });
    }
    rows.sort((a, b) => {
      // Blanks sink to the bottom in both directions, as they do in Excel;
      // a descending sort that floats every empty row to the top is useless.
      const aBlank = a.key === null;
      const bBlank = b.key === null;
      if (aBlank !== bBlank) return aBlank ? 1 : -1;
      return compareScalars(a.key, b.key) * (direction === 'asc' ? 1 : -1);
    });

    this.doc.transact(
      { label: `Sort by column ${a1(0, byColumn).replace(/\d+/, '')}`, origin: 'user', timestamp: this.now() },
      () => {
        rows.forEach((row, i) => {
          row.cells.forEach((cell, j) => {
            this.doc.setCell(sheet.name, firstRow + i, range.start.col + j, cell);
          });
        });
      },
    );
    this.applyRecalc(this.engine.recalculateAll());
  }

  /** Hide rows whose value in `column` does not contain `query`. */
  filterRows(column: number, query: string): number {
    const sheet = this.sheet();
    if (!sheet) return 0;
    const needle = query.toLowerCase();
    const bounds = sheet.bounds();
    if (!bounds) return 0;
    let hidden = 0;
    this.doc.transact({ label: query === '' ? 'Clear filter' : `Filter on ${query}`, origin: 'user', timestamp: this.now() }, () => {
      for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
        const text = this.displayText({ sheet: sheet.name, row: r, col: column }).toLowerCase();
        const keep = query === '' || text.includes(needle);
        const props = { ...sheet.rows.get(r) };
        if (keep) delete props.hidden;
        else {
          props.hidden = true;
          hidden++;
        }
        this.doc.setRowProps(sheet.name, r, Object.keys(props).length > 0 ? props : undefined);
      }
    });
    return hidden;
  }

  // --- import -------------------------------------------------------------

  /** Stage a paste or CSV import for review; nothing is written until confirmed. */
  proposeImport(pending: PendingImport): void {
    this.bump({ pendingImport: pending });
  }

  setImportOverride(column: number, override: ColumnOverride): void {
    const pending = this.snapshot.pendingImport;
    if (!pending) return;
    const overrides = [...pending.overrides];
    overrides[column] = override;
    this.bump({ pendingImport: { ...pending, overrides } });
  }

  setImportHeaderRow(headerRow: boolean): void {
    const pending = this.snapshot.pendingImport;
    if (!pending) return;
    this.bump({ pendingImport: { ...pending, headerRow } });
  }

  cancelImport(): void {
    this.bump({ pendingImport: null });
  }

  /** Write the staged import, honouring the per-column overrides. */
  confirmImport(
    resolve: (
      rows: readonly (readonly string[])[],
      options: EntryOptions,
      overrides: readonly ColumnOverride[],
      headerRow: boolean,
    ) => { value: Scalar; literal?: string; row: number; col: number }[],
  ): number {
    const pending = this.snapshot.pendingImport;
    if (!pending) return 0;
    const cells = resolve(pending.rows, this.entryOptions, pending.overrides, pending.headerRow);
    const anchor = pending.anchor;
    this.doc.transact(
      { label: `Import ${cells.length} cells`, origin: 'import', timestamp: this.now() },
      () => {
        for (const cell of cells) {
          const data: CellData = { value: cell.value };
          if (cell.literal !== undefined) data.literal = cell.literal;
          this.doc.setCell(anchor.sheet, anchor.row + cell.row, anchor.col + cell.col, data);
        }
      },
    );
    this.bump({ pendingImport: null, dirty: true });
    this.applyRecalc(this.engine.recalculateAll());
    return cells.length;
  }

  // --- UI state -----------------------------------------------------------

  setTheme(theme: ThemePreference): void {
    this.bump({ theme });
  }

  togglePanel(panel: PanelName, open?: boolean): void {
    const panels = { ...this.snapshot.panels };
    panels[panel] = open ?? !panels[panel];
    this.bump({ panels });
  }

  setPalette(open: boolean, mode: 'all' | 'command' = 'all'): void {
    this.bump({ paletteOpen: open, paletteMode: mode });
  }

  /**
   * Select the whole used extent of the active sheet.
   *
   * The used extent rather than all 17 billion addresses: selecting the entire
   * grid would make every selection-wide operation - a format, a sort, the
   * status-bar aggregate - walk a range that is almost entirely empty. An empty
   * sheet has no extent, so the selection stays where it is.
   */
  selectAll(): void {
    const sheet = this.sheet();
    const bounds = sheet?.bounds();
    if (!sheet || !bounds) return;
    this.setSelection({
      sheet: sheet.name,
      active: { row: bounds.minRow, col: bounds.minCol },
      ranges: [
        {
          start: { row: bounds.minRow, col: bounds.minCol },
          end: { row: bounds.maxRow, col: bounds.maxCol },
        },
      ],
    });
  }

  /** Open or close find and replace; no argument toggles it. */
  setFind(open?: boolean): void {
    this.bump({ findOpen: open ?? !this.snapshot.findOpen });
  }

  setMessage(message: string | null): void {
    this.bump({ message });
  }

  setFileName(name: string): void {
    this.bump({ fileName: name });
  }

  markSaved(): void {
    this.bump({ dirty: false });
  }

  setEntryOptions(options: EntryOptions): void {
    this.entryOptions = options;
    this.bump();
  }

  getEntryOptions(): EntryOptions {
    return this.entryOptions;
  }

  // --- helpers ------------------------------------------------------------

  /**
   * Addresses covered by the selection.
   *
   * Clamped to the used range by default, because a whole-column selection
   * covers a million addresses and the user meant the twelve with data in them.
   * `includeEmpty` still clamps; it only stops empty cells being skipped inside
   * the clamped rectangle, which is what formatting needs.
   */
  selectedAddresses(options: { includeEmpty?: boolean; cap?: number } = {}): CellAddr[] {
    const sheet = this.sheet();
    if (!sheet) return [];
    const bounds = sheet.bounds() ?? { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 };
    const cap = options.cap ?? MAX_BULK_CELLS;
    const out: CellAddr[] = [];
    const seen = new Set<number>();

    for (const raw of this.snapshot.selection.ranges) {
      const range = normaliseRange(raw);
      const endRow = Math.min(range.end.row, Math.max(bounds.maxRow, range.start.row));
      const endCol = Math.min(range.end.col, Math.max(bounds.maxCol, range.start.col));
      for (let r = range.start.row; r <= endRow; r++) {
        for (let c = range.start.col; c <= endCol; c++) {
          const key = r * 16_384 + c;
          if (seen.has(key)) continue;
          if (!options.includeEmpty && !sheet.getCell(r, c)) continue;
          seen.add(key);
          out.push({ sheet: sheet.name, row: r, col: c });
          if (out.length >= cap) return out;
        }
      }
    }
    return out;
  }

  selectionRange(): SelectionRange {
    return normaliseRange(this.snapshot.selection.ranges[0]!);
  }
}

function replaceInsensitive(source: string, query: string, replacement: string): string {
  if (query === '') return source;
  const lower = source.toLowerCase();
  const needle = query.toLowerCase();
  let out = '';
  let i = 0;
  while (i < source.length) {
    const at = lower.indexOf(needle, i);
    if (at < 0) {
      out += source.slice(i);
      break;
    }
    out += source.slice(i, at) + replacement;
    i = at + needle.length;
  }
  return out;
}

/**
 * Excel's sort order across mixed types: numbers, then text, then logicals, then
 * errors, then blanks. Callers handle blanks separately when they need them
 * pinned to the bottom of a descending sort.
 */
export function compareScalars(a: Scalar, b: Scalar): number {
  const rank = (v: Scalar): number => {
    if (v === null) return 5;
    if (typeof v === 'number') return 0;
    if (typeof v === 'string') return 1;
    if (typeof v === 'boolean') return 2;
    return 3;
  };
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') return a.localeCompare(b);
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  return 0;
}

export type { ChangeOrigin };
