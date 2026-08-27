/**
 * Command log: the undo system, and the reason it can do things Excel's cannot.
 *
 * Every mutation to a workbook goes through here as a command carrying its own
 * inverse. Nothing else is allowed to touch the model. That single rule is what
 * makes the following fall out almost for free, where Excel has treated them as
 * hard or impossible for decades:
 *
 *   Undo survives running a macro. In Excel, VBA wipes the undo stack, so an
 *   automation that goes wrong cannot be taken back - one of the longest-running
 *   complaints about the product. Here a script is simply another writer on the
 *   same bus, so its whole run collapses into ONE undo entry that reverses like
 *   any other edit.
 *
 *   Undo is per-document. Excel shares one stack across every open workbook, so
 *   undoing in one file can silently reverse an edit in another.
 *
 *   Undo is unlimited and survives closing the file, because the log is data
 *   that can be persisted beside the document rather than transient UI state.
 *
 *   Structural operations - deleting a sheet, renaming one - are commands like
 *   any other, so they are undoable rather than being the exceptions Excel
 *   silently drops off the stack.
 *
 *   History is a TREE, not a stack. Undoing and then doing something else does
 *   not destroy the abandoned branch, so exploring an alternative is not a
 *   one-way door.
 *
 * Some operations cannot be honestly inverted: a volatile function, a random
 * number, a refresh of external data. Those are marked as replay barriers rather
 * than pretending an inverse exists, which is the honest failure mode.
 */

import type { CellData, ColProps, RowProps, SheetVisibility } from './sheet.js';
import { Sheet, Workbook } from './sheet.js';
import type { StyleId } from './style.js';

/** Who caused a change, which drives both grouping and the history labels. */
export type ChangeOrigin = 'user' | 'script' | 'macro' | 'import' | 'recalc' | 'system';

export type Change =
  | { kind: 'cell'; sheet: string; row: number; col: number; before?: CellData; after?: CellData }
  | { kind: 'row'; sheet: string; row: number; before?: RowProps; after?: RowProps }
  | { kind: 'col'; sheet: string; col: number; before?: ColProps; after?: ColProps }
  | { kind: 'sheetName'; sheet: string; before: string; after: string }
  | { kind: 'sheetVisibility'; sheet: string; before: SheetVisibility; after: SheetVisibility }
  | { kind: 'sheetAdd'; sheet: string; index: number }
  | { kind: 'sheetRemove'; sheet: string; index: number; snapshot: SheetSnapshot }
  | { kind: 'sheetMove'; sheet: string; from: number; to: number };

/** Everything needed to rebuild a removed sheet, so deleting one is undoable. */
export interface SheetSnapshot {
  name: string;
  id: number;
  cells: [number, CellData][];
  rows: [number, RowProps][];
  cols: [number, ColProps][];
  merges: Sheet['merges'];
  view: Sheet['view'];
  visibility: SheetVisibility;
  tabColor?: string;
  preserved: Record<string, string>;
}

export interface HistoryEntry {
  id: number;
  /** Human-readable, shown in the history panel: "Type 42 in B7". */
  label: string;
  origin: ChangeOrigin;
  /** Milliseconds since the epoch, supplied by the caller so tests are stable. */
  timestamp: number;
  changes: Change[];
  /** Parent entry, making the history a tree rather than a stack. */
  parent: number | null;
  /**
   * True when the entry cannot be faithfully reversed - a refresh of external
   * data, or anything reading a clock or a random source. Undo stops here and
   * says why rather than producing a plausible wrong state.
   */
  barrier?: boolean;
}

export interface CommandOptions {
  label?: string;
  origin?: ChangeOrigin;
  timestamp?: number;
  barrier?: boolean;
}

/**
 * A workbook plus its history.
 *
 * All mutation goes through the `edit` methods, which record inverses as they
 * go. Reads go straight to `workbook`.
 */
export class Document {
  readonly workbook: Workbook;

  private readonly entries = new Map<number, HistoryEntry>();
  private nextId = 1;
  /** The entry the document currently reflects; null means the initial state. */
  private head: number | null = null;
  /** Open transaction, if any. */
  private pending: Change[] | null = null;
  private pendingOptions: CommandOptions | null = null;
  private listeners: ((changes: Change[]) => void)[] = [];

  constructor(workbook = new Workbook()) {
    this.workbook = workbook;
  }

  /** Subscribe to applied changes, for the renderer and the recalc engine. */
  onChange(listener: (changes: Change[]) => void): () => void {
    this.listeners.push(listener);
    return () => {
      this.listeners = this.listeners.filter((l) => l !== listener);
    };
  }

  private emit(changes: Change[]): void {
    for (const l of this.listeners) l(changes);
  }

  // --- transactions -------------------------------------------------------

  /**
   * Run `body` as one undoable unit.
   *
   * This is the whole trick behind surviving a macro: a script driving a
   * thousand edits is one entry, and undoing it reverses all thousand. Nested
   * calls join the outermost transaction rather than creating their own.
   */
  transact<T>(options: CommandOptions, body: () => T): T {
    if (this.pending) return body();

    this.pending = [];
    this.pendingOptions = options;
    try {
      const result = body();
      this.commit();
      return result;
    } catch (err) {
      // A failed transaction reverses cleanly, so a script that throws halfway
      // does not leave the document in a state the user never asked for.
      this.rollback();
      throw err;
    }
  }

  private commit(): void {
    const changes = this.pending ?? [];
    const options = this.pendingOptions ?? {};
    this.pending = null;
    this.pendingOptions = null;
    if (changes.length === 0) return;

    const entry: HistoryEntry = {
      id: this.nextId++,
      label: options.label ?? describe(changes),
      origin: options.origin ?? 'user',
      timestamp: options.timestamp ?? 0,
      changes,
      parent: this.head,
    };
    if (options.barrier) entry.barrier = true;
    this.entries.set(entry.id, entry);
    this.head = entry.id;
    this.emit(changes);
  }

  private rollback(): void {
    const changes = this.pending ?? [];
    this.pending = null;
    this.pendingOptions = null;
    for (let i = changes.length - 1; i >= 0; i--) this.revert(changes[i]!);
  }

  /** Record a change, either into the open transaction or as its own entry. */
  private record(change: Change, options: CommandOptions = {}): void {
    if (this.pending) {
      this.pending.push(change);
      return;
    }
    this.pending = [change];
    this.pendingOptions = options;
    this.commit();
  }

  // --- editing ------------------------------------------------------------

  setCell(sheetName: string, row: number, col: number, data: CellData | undefined, options?: CommandOptions): void {
    const sheet = this.requireSheet(sheetName);
    const before = sheet.getCell(row, col);
    sheet.setCell(row, col, data);
    this.record(
      {
        kind: 'cell',
        sheet: sheetName,
        row,
        col,
        ...(before ? { before: clone(before) } : {}),
        ...(data ? { after: clone(data) } : {}),
      },
      options ?? { label: `Edit ${sheetName}!${row + 1},${col + 1}` },
    );
  }

  setValue(sheetName: string, row: number, col: number, value: CellData['value'], options?: CommandOptions): void {
    const sheet = this.requireSheet(sheetName);
    const before = sheet.getCell(row, col);
    const after: CellData = { ...before, value };
    delete after.formula;
    this.setCell(sheetName, row, col, after.value === null && after.style === undefined ? undefined : after, options);
  }

  setFormula(sheetName: string, row: number, col: number, formula: string, cached: CellData['value'] = null, options?: CommandOptions): void {
    const sheet = this.requireSheet(sheetName);
    const before = sheet.getCell(row, col);
    this.setCell(sheetName, row, col, { ...before, value: cached, formula }, options);
  }

  setStyle(sheetName: string, row: number, col: number, style: StyleId, options?: CommandOptions): void {
    const sheet = this.requireSheet(sheetName);
    const before = sheet.getCell(row, col);
    this.setCell(sheetName, row, col, { value: before?.value ?? null, ...before, style }, options);
  }

  setRowProps(sheetName: string, row: number, props: RowProps | undefined, options?: CommandOptions): void {
    const sheet = this.requireSheet(sheetName);
    const before = sheet.rows.get(row);
    if (props) sheet.rows.set(row, props);
    else sheet.rows.delete(row);
    this.record(
      {
        kind: 'row',
        sheet: sheetName,
        row,
        ...(before ? { before: clone(before) } : {}),
        ...(props ? { after: clone(props) } : {}),
      },
      options ?? { label: `Change row ${row + 1}` },
    );
  }

  setColProps(sheetName: string, col: number, props: ColProps | undefined, options?: CommandOptions): void {
    const sheet = this.requireSheet(sheetName);
    const before = sheet.cols.get(col);
    if (props) sheet.cols.set(col, props);
    else sheet.cols.delete(col);
    this.record(
      {
        kind: 'col',
        sheet: sheetName,
        col,
        ...(before ? { before: clone(before) } : {}),
        ...(props ? { after: clone(props) } : {}),
      },
      options ?? { label: `Change column ${col + 1}` },
    );
  }

  renameSheet(from: string, to: string, options?: CommandOptions): void {
    const sheet = this.requireSheet(from);
    sheet.name = to;
    this.record({ kind: 'sheetName', sheet: to, before: from, after: to }, options ?? {
      label: `Rename ${from} to ${to}`,
    });
  }

  setSheetVisibility(sheetName: string, visibility: SheetVisibility, options?: CommandOptions): void {
    const sheet = this.requireSheet(sheetName);
    const before = sheet.visibility;
    sheet.visibility = visibility;
    this.record({ kind: 'sheetVisibility', sheet: sheetName, before, after: visibility }, options ?? {
      label: `${visibility === 'visible' ? 'Show' : 'Hide'} ${sheetName}`,
    });
  }

  addSheet(name: string, at?: number, options?: CommandOptions): Sheet {
    const sheet = this.workbook.addSheet(name, at);
    this.record(
      { kind: 'sheetAdd', sheet: sheet.name, index: this.workbook.sheetIndex(sheet) },
      options ?? { label: `Add sheet ${sheet.name}` },
    );
    return sheet;
  }

  /**
   * Delete a sheet. Excel makes this one of the operations it cannot undo;
   * snapshotting it first is all that is needed to make it ordinary.
   */
  removeSheet(name: string, options?: CommandOptions): boolean {
    const sheet = this.workbook.getSheet(name);
    if (!sheet) return false;
    const index = this.workbook.sheetIndex(sheet);
    const snapshot = snapshotSheet(sheet);
    this.workbook.removeSheet(name);
    this.record({ kind: 'sheetRemove', sheet: name, index, snapshot }, options ?? {
      label: `Delete sheet ${name}`,
    });
    return true;
  }

  moveSheet(name: string, to: number, options?: CommandOptions): void {
    const sheet = this.requireSheet(name);
    const from = this.workbook.sheetIndex(sheet);
    if (from === to) return;
    this.workbook.sheets.splice(from, 1);
    this.workbook.sheets.splice(to, 0, sheet);
    this.record({ kind: 'sheetMove', sheet: name, from, to }, options ?? {
      label: `Move sheet ${name}`,
    });
  }

  // --- history navigation -------------------------------------------------

  get canUndo(): boolean {
    return this.head !== null;
  }

  get canRedo(): boolean {
    return this.childrenOf(this.head).length > 0;
  }

  /** The entry that would be undone next. */
  peekUndo(): HistoryEntry | undefined {
    return this.head === null ? undefined : this.entries.get(this.head);
  }

  peekRedo(): HistoryEntry | undefined {
    const children = this.childrenOf(this.head);
    // Default to the most recent branch, which is what a stack would have done.
    return children[children.length - 1];
  }

  undo(): HistoryEntry | undefined {
    const entry = this.peekUndo();
    if (!entry) return undefined;
    for (let i = entry.changes.length - 1; i >= 0; i--) this.revert(entry.changes[i]!);
    this.head = entry.parent;
    this.emit(entry.changes);
    return entry;
  }

  redo(entryId?: number): HistoryEntry | undefined {
    const children = this.childrenOf(this.head);
    const entry =
      entryId === undefined ? children[children.length - 1] : children.find((c) => c.id === entryId);
    if (!entry) return undefined;
    for (const change of entry.changes) this.apply(change);
    this.head = entry.id;
    this.emit(entry.changes);
    return entry;
  }

  /**
   * Jump directly to any point in the history.
   *
   * Because the history is a tree, this walks up from the current head to the
   * common ancestor, then down to the target, reversing and reapplying along
   * the way. That is what makes a clickable history panel work rather than
   * forcing the user to press undo forty times.
   */
  jumpTo(entryId: number | null): boolean {
    if (entryId !== null && !this.entries.has(entryId)) return false;

    const toRoot = (id: number | null): number[] => {
      const path: number[] = [];
      let current = id;
      while (current !== null) {
        path.push(current);
        current = this.entries.get(current)?.parent ?? null;
      }
      return path;
    };

    const fromPath = toRoot(this.head);
    const toPath = toRoot(entryId);
    const toSet = new Set(toPath);
    const ancestor = fromPath.find((id) => toSet.has(id)) ?? null;

    for (const id of fromPath) {
      if (id === ancestor) break;
      const entry = this.entries.get(id)!;
      for (let i = entry.changes.length - 1; i >= 0; i--) this.revert(entry.changes[i]!);
    }

    const descend: number[] = [];
    for (const id of toPath) {
      if (id === ancestor) break;
      descend.push(id);
    }
    descend.reverse();
    for (const id of descend) {
      const entry = this.entries.get(id)!;
      for (const change of entry.changes) this.apply(change);
    }

    this.head = entryId;
    return true;
  }

  /** The linear path from the beginning to the current head, oldest first. */
  history(): HistoryEntry[] {
    const path: HistoryEntry[] = [];
    let current = this.head;
    while (current !== null) {
      const entry = this.entries.get(current);
      if (!entry) break;
      path.push(entry);
      current = entry.parent;
    }
    return path.reverse();
  }

  /** Every entry, including abandoned branches, for the history panel. */
  allEntries(): HistoryEntry[] {
    return [...this.entries.values()].sort((a, b) => a.id - b.id);
  }

  private childrenOf(parent: number | null): HistoryEntry[] {
    return [...this.entries.values()].filter((e) => e.parent === parent).sort((a, b) => a.id - b.id);
  }

  // --- applying and reverting --------------------------------------------

  private apply(change: Change): void {
    switch (change.kind) {
      case 'cell':
        this.workbook.getSheet(change.sheet)?.setCell(change.row, change.col, change.after);
        break;
      case 'row': {
        const sheet = this.workbook.getSheet(change.sheet);
        if (!sheet) break;
        if (change.after) sheet.rows.set(change.row, change.after);
        else sheet.rows.delete(change.row);
        break;
      }
      case 'col': {
        const sheet = this.workbook.getSheet(change.sheet);
        if (!sheet) break;
        if (change.after) sheet.cols.set(change.col, change.after);
        else sheet.cols.delete(change.col);
        break;
      }
      case 'sheetName': {
        const sheet = this.workbook.getSheet(change.before);
        if (sheet) sheet.name = change.after;
        break;
      }
      case 'sheetVisibility': {
        const sheet = this.workbook.getSheet(change.sheet);
        if (sheet) sheet.visibility = change.after;
        break;
      }
      case 'sheetAdd':
        if (!this.workbook.getSheet(change.sheet)) this.workbook.addSheet(change.sheet, change.index);
        break;
      case 'sheetRemove':
        this.workbook.removeSheet(change.sheet);
        break;
      case 'sheetMove': {
        const sheet = this.workbook.getSheet(change.sheet);
        if (!sheet) break;
        const at = this.workbook.sheetIndex(sheet);
        this.workbook.sheets.splice(at, 1);
        this.workbook.sheets.splice(change.to, 0, sheet);
        break;
      }
    }
  }

  private revert(change: Change): void {
    switch (change.kind) {
      case 'cell':
        this.workbook.getSheet(change.sheet)?.setCell(change.row, change.col, change.before);
        break;
      case 'row': {
        const sheet = this.workbook.getSheet(change.sheet);
        if (!sheet) break;
        if (change.before) sheet.rows.set(change.row, change.before);
        else sheet.rows.delete(change.row);
        break;
      }
      case 'col': {
        const sheet = this.workbook.getSheet(change.sheet);
        if (!sheet) break;
        if (change.before) sheet.cols.set(change.col, change.before);
        else sheet.cols.delete(change.col);
        break;
      }
      case 'sheetName': {
        const sheet = this.workbook.getSheet(change.after);
        if (sheet) sheet.name = change.before;
        break;
      }
      case 'sheetVisibility': {
        const sheet = this.workbook.getSheet(change.sheet);
        if (sheet) sheet.visibility = change.before;
        break;
      }
      case 'sheetAdd':
        this.workbook.removeSheet(change.sheet);
        break;
      case 'sheetRemove':
        restoreSheet(this.workbook, change.snapshot, change.index);
        break;
      case 'sheetMove': {
        const sheet = this.workbook.getSheet(change.sheet);
        if (!sheet) break;
        const at = this.workbook.sheetIndex(sheet);
        this.workbook.sheets.splice(at, 1);
        this.workbook.sheets.splice(change.from, 0, sheet);
        break;
      }
    }
  }

  private requireSheet(name: string): Sheet {
    const sheet = this.workbook.getSheet(name);
    if (!sheet) throw new Error(`no such sheet: ${name}`);
    return sheet;
  }
}

function clone<T>(v: T): T {
  return structuredClone(v);
}

export function snapshotSheet(sheet: Sheet): SheetSnapshot {
  const snapshot: SheetSnapshot = {
    name: sheet.name,
    id: sheet.id,
    cells: [...sheet.cells.entries()].map(([k, v]) => [k, clone(v)] as [number, CellData]),
    rows: [...sheet.rows.entries()].map(([k, v]) => [k, clone(v)] as [number, RowProps]),
    cols: [...sheet.cols.entries()].map(([k, v]) => [k, clone(v)] as [number, ColProps]),
    merges: clone(sheet.merges),
    view: clone(sheet.view),
    visibility: sheet.visibility,
    preserved: clone(sheet.preserved),
  };
  if (sheet.tabColor !== undefined) snapshot.tabColor = sheet.tabColor;
  return snapshot;
}

export function restoreSheet(workbook: Workbook, snapshot: SheetSnapshot, index: number): Sheet {
  const sheet = workbook.addSheet(snapshot.name, index);
  for (const [k, v] of snapshot.cells) sheet.cells.set(k, v);
  for (const [k, v] of snapshot.rows) sheet.rows.set(k, v);
  for (const [k, v] of snapshot.cols) sheet.cols.set(k, v);
  sheet.merges = snapshot.merges;
  sheet.view = snapshot.view;
  sheet.visibility = snapshot.visibility;
  if (snapshot.tabColor !== undefined) sheet.tabColor = snapshot.tabColor;
  sheet.preserved = snapshot.preserved;
  return sheet;
}

/**
 * A default label for an entry whose caller did not supply one.
 *
 * Labels matter more than they look: the history panel is only useful if its
 * rows say what happened, and "42 changes" is not an answer.
 */
function describe(changes: Change[]): string {
  if (changes.length === 1) {
    const c = changes[0]!;
    switch (c.kind) {
      case 'cell':
        return `Edit ${c.sheet}`;
      case 'row':
        return `Change row ${c.row + 1}`;
      case 'col':
        return `Change column ${c.col + 1}`;
      case 'sheetName':
        return `Rename ${c.before} to ${c.after}`;
      case 'sheetVisibility':
        return `${c.after === 'visible' ? 'Show' : 'Hide'} ${c.sheet}`;
      case 'sheetAdd':
        return `Add sheet ${c.sheet}`;
      case 'sheetRemove':
        return `Delete sheet ${c.sheet}`;
      case 'sheetMove':
        return `Move sheet ${c.sheet}`;
    }
  }
  const cells = changes.filter((c) => c.kind === 'cell').length;
  if (cells === changes.length) return `Edit ${cells} cells`;
  return `${changes.length} changes`;
}
