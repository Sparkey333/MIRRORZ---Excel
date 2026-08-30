import { describe, expect, it } from 'vitest';
import { CellError } from '@mirrorz/core';
import { compareScalars } from '../src/renderer/state/controller.js';
import { createController } from '../src/renderer/state/context.js';
import { singleCell } from '../src/renderer/model/selection.js';

function controllerWith(cells: [number, number, string][] = []) {
  const controller = createController(undefined, { now: () => 1_000 });
  for (const [row, col, text] of cells) controller.commitEntry(text, { sheet: 'Sheet1', row, col });
  return controller;
}

describe('AppController snapshots', () => {
  it('starts with one sheet and a single-cell selection', () => {
    const controller = controllerWith();
    expect(controller.getSnapshot().activeSheet).toBe('Sheet1');
    expect(controller.getSnapshot().selection).toEqual(singleCell('Sheet1', 0, 0));
  });

  it('replaces the snapshot object on every change, so a store can compare by identity', () => {
    const controller = controllerWith();
    const before = controller.getSnapshot();
    controller.selectCell(1, 1);
    expect(controller.getSnapshot()).not.toBe(before);
  });

  it('notifies subscribers', () => {
    const controller = controllerWith();
    let calls = 0;
    const unsubscribe = controller.subscribe(() => calls++);
    controller.selectCell(2, 2);
    expect(calls).toBeGreaterThan(0);
    unsubscribe();
    const after = calls;
    controller.selectCell(3, 3);
    expect(calls).toBe(after);
  });
});

describe('editing', () => {
  it('routes an edit through the document so it is undoable', () => {
    const controller = controllerWith([[0, 0, '42']]);
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe(42);
    expect(controller.getSnapshot().canUndo).toBe(true);
    controller.undo();
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })).toBeUndefined();
  });

  it('keeps the literal text of a value that would otherwise be mangled', () => {
    const controller = controllerWith([[0, 0, '007']]);
    const cell = controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })!;
    expect(cell.value).toBe('007');
    expect(controller.editText({ sheet: 'Sheet1', row: 0, col: 0 })).toBe('007');
  });

  it('stores a formula without its leading equals and shows it back with one', () => {
    const controller = controllerWith([[0, 0, '=1+1']]);
    const cell = controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })!;
    expect(cell.formula).toBe('1+1');
    expect(controller.editText({ sheet: 'Sheet1', row: 0, col: 0 })).toBe('=1+1');
  });

  it('applies an implied number format for a date entry', () => {
    const controller = controllerWith([[0, 0, '2024-01-31']]);
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).numFmt).toBe('yyyy-mm-dd');
  });

  it('does not overwrite an existing format with an implied one', () => {
    const controller = controllerWith();
    controller.selectCell(0, 0);
    controller.setNumberFormat('0.00');
    controller.commitEntry('2024-01-31', { sheet: 'Sheet1', row: 0, col: 0 });
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).numFmt).toBe('0.00');
  });

  it('clears the selection as one undoable step', () => {
    const controller = controllerWith([
      [0, 0, '1'],
      [0, 1, '2'],
    ]);
    controller.setSelection({
      sheet: 'Sheet1',
      active: { row: 0, col: 0 },
      ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }],
    });
    controller.clearSelection();
    expect(controller.sheet()!.cellCount).toBe(0);
    controller.undo();
    expect(controller.sheet()!.cellCount).toBe(2);
  });
});

describe('formatting', () => {
  it('merges a style patch rather than replacing the whole format', () => {
    const controller = controllerWith([[0, 0, '1']]);
    controller.selectCell(0, 0);
    controller.setNumberFormat('0.00');
    controller.applyStyle({ font: { bold: true } });
    const style = controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 });
    expect(style.numFmt).toBe('0.00');
    expect(style.font?.bold).toBe(true);
  });

  it('reports a uniform property across the selection', () => {
    const controller = controllerWith([
      [0, 0, '1'],
      [0, 1, '2'],
    ]);
    controller.setSelection({
      sheet: 'Sheet1',
      active: { row: 0, col: 0 },
      ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }],
    });
    controller.applyStyle({ font: { bold: true } });
    expect(controller.uniformStyle((s) => s.font?.bold)).toBe(true);
  });

  it('reports undefined when the selection disagrees', () => {
    const controller = controllerWith([
      [0, 0, '1'],
      [0, 1, '2'],
    ]);
    controller.selectCell(0, 0);
    controller.applyStyle({ font: { bold: true } });
    controller.setSelection({
      sheet: 'Sheet1',
      active: { row: 0, col: 0 },
      ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }],
    });
    expect(controller.uniformStyle((s) => s.font?.bold)).toBeUndefined();
  });

  it('clamps a whole-column selection to the used range instead of a million writes', () => {
    const controller = controllerWith([[0, 0, '1']]);
    controller.setSelection({
      sheet: 'Sheet1',
      active: { row: 0, col: 0 },
      ranges: [{ start: { row: 0, col: 0 }, end: { row: 1_048_575, col: 0 } }],
    });
    expect(controller.selectedAddresses({ includeEmpty: true }).length).toBe(1);
  });
});

describe('sheets', () => {
  it('adds and activates a sheet', () => {
    const controller = controllerWith();
    controller.addSheet('Data');
    expect(controller.getSnapshot().activeSheet).toBe('Data');
  });

  it('renames a sheet and follows it', () => {
    const controller = controllerWith();
    expect(controller.renameSheet('Sheet1', 'Summary')).toBe(true);
    expect(controller.getSnapshot().activeSheet).toBe('Summary');
  });

  it('refuses a duplicate name and says why', () => {
    const controller = controllerWith();
    controller.addSheet('Data');
    expect(controller.renameSheet('Sheet1', 'Data')).toBe(false);
    expect(controller.getSnapshot().message).toMatch(/already exists/);
  });

  it('refuses to delete the only visible sheet', () => {
    const controller = controllerWith();
    controller.removeSheet('Sheet1');
    expect(controller.workbook.sheets).toHaveLength(1);
    expect(controller.getSnapshot().message).toMatch(/at least one visible sheet/);
  });

  it('deletes a sheet undoably, snapshot and all', () => {
    const controller = controllerWith();
    controller.addSheet('Data');
    controller.commitEntry('9', { sheet: 'Data', row: 0, col: 0 });
    controller.removeSheet('Data');
    expect(controller.workbook.getSheet('Data')).toBeUndefined();
    controller.undo();
    expect(controller.workbook.getSheet('Data')?.getValue(0, 0)).toBe(9);
  });

  it('moves off a sheet it has just hidden', () => {
    const controller = controllerWith();
    controller.addSheet('Data');
    controller.setSheetVisibility('Data', 'hidden');
    expect(controller.getSnapshot().activeSheet).toBe('Sheet1');
  });
});

describe('history navigation', () => {
  it('repaints after jumpTo, which the document does not emit for', () => {
    const controller = controllerWith([
      [0, 0, '1'],
      [0, 1, '2'],
    ]);
    const before = controller.getSnapshot().version;
    controller.jumpTo(null);
    expect(controller.getSnapshot().version).toBeGreaterThan(before);
    expect(controller.sheet()!.cellCount).toBe(0);
  });

  it('reaches an abandoned branch, which a stack could not', () => {
    const controller = controllerWith([[0, 0, 'first']]);
    const branchPoint = controller.historyEntries().at(-1)!.id;
    controller.undo();
    controller.commitEntry('second', { sheet: 'Sheet1', row: 0, col: 0 });
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe('second');
    controller.jumpTo(branchPoint);
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe('first');
  });

  it('exposes the head id for the history panel', () => {
    const controller = controllerWith([[0, 0, '1']]);
    expect(controller.getSnapshot().headId).toBe(controller.historyEntries().at(-1)!.id);
  });
});

describe('find and replace', () => {
  const build = () =>
    controllerWith([
      [0, 0, 'alpha'],
      [1, 0, 'Alphabet'],
      [2, 0, 'beta'],
    ]);

  it('finds case-insensitively by default', () => {
    expect(build().find('alpha')).toHaveLength(2);
  });

  it('honours match case', () => {
    expect(build().find('Alpha', { matchCase: true })).toHaveLength(1);
  });

  it('honours whole-cell matching', () => {
    expect(build().find('alpha', { whole: true })).toHaveLength(1);
  });

  it('searches formulas when asked', () => {
    const controller = controllerWith([[0, 0, '=1+1']]);
    expect(controller.find('1+1', { formulas: true })).toHaveLength(1);
    expect(controller.find('1+1')).toHaveLength(0);
  });

  it('replaces all matches in one undoable step', () => {
    const controller = build();
    expect(controller.replaceAll('alpha', 'gamma')).toBe(2);
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe('gamma');
    controller.undo();
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe('alpha');
  });
});

describe('sort and filter', () => {
  it('moves whole rows, not just the sorted column', () => {
    const controller = controllerWith([
      [0, 0, '3'],
      [0, 1, 'c'],
      [1, 0, '1'],
      [1, 1, 'a'],
      [2, 0, '2'],
      [2, 1, 'b'],
    ]);
    controller.setSelection({
      sheet: 'Sheet1',
      active: { row: 0, col: 0 },
      ranges: [{ start: { row: 0, col: 0 }, end: { row: 2, col: 1 } }],
    });
    controller.sortSelection(0, 'asc');
    expect([0, 1, 2].map((r) => controller.cellAt({ sheet: 'Sheet1', row: r, col: 1 })?.value)).toEqual([
      'a',
      'b',
      'c',
    ]);
  });

  it('keeps a header row in place', () => {
    const controller = controllerWith([
      [0, 0, 'Header'],
      [1, 0, '2'],
      [2, 0, '1'],
    ]);
    controller.setSelection({
      sheet: 'Sheet1',
      active: { row: 0, col: 0 },
      ranges: [{ start: { row: 0, col: 0 }, end: { row: 2, col: 0 } }],
    });
    controller.sortSelection(0, 'asc', true);
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe('Header');
    expect(controller.cellAt({ sheet: 'Sheet1', row: 1, col: 0 })?.value).toBe(1);
  });

  it('hides rows that do not match a filter, and restores them when cleared', () => {
    const controller = controllerWith([
      [0, 0, 'apple'],
      [1, 0, 'banana'],
    ]);
    expect(controller.filterRows(0, 'app')).toBe(1);
    expect(controller.sheet()!.isRowHidden(1)).toBe(true);
    controller.filterRows(0, '');
    expect(controller.sheet()!.isRowHidden(1)).toBe(false);
  });
});

describe('import staging', () => {
  it('writes nothing until the import is confirmed', () => {
    const controller = controllerWith();
    controller.proposeImport({
      anchor: { sheet: 'Sheet1', row: 0, col: 0 },
      rows: [['1', '2']],
      overrides: ['auto', 'auto'],
      headerRow: false,
      source: 'paste',
    });
    expect(controller.sheet()!.cellCount).toBe(0);
    expect(controller.getSnapshot().pendingImport).not.toBeNull();
  });

  it('discards the staged import on cancel', () => {
    const controller = controllerWith();
    controller.proposeImport({
      anchor: { sheet: 'Sheet1', row: 0, col: 0 },
      rows: [['1']],
      overrides: ['auto'],
      headerRow: false,
      source: 'paste',
    });
    controller.cancelImport();
    expect(controller.getSnapshot().pendingImport).toBeNull();
    expect(controller.sheet()!.cellCount).toBe(0);
  });

  it('writes the import at the anchor as a single undoable step', () => {
    const controller = controllerWith();
    controller.selectCell(2, 3);
    controller.proposeImport({
      anchor: { sheet: 'Sheet1', row: 2, col: 3 },
      rows: [['1', '2']],
      overrides: ['auto', 'auto'],
      headerRow: false,
      source: 'paste',
    });
    const written = controller.confirmImport((rows) =>
      rows.flatMap((row, r) => row.map((value, c) => ({ value: Number(value), row: r, col: c }))),
    );
    expect(written).toBe(2);
    expect(controller.cellAt({ sheet: 'Sheet1', row: 2, col: 4 })?.value).toBe(2);
    controller.undo();
    expect(controller.cellAt({ sheet: 'Sheet1', row: 2, col: 4 })).toBeUndefined();
  });
});

describe('every visible edit is undoable', () => {
  it('undoes a tab colour', () => {
    const controller = createController(undefined, { now: () => 1 });
    controller.addSheet('Data');
    controller.setSheetColor('Data', '#e03131');
    expect(controller.workbook.getSheet('Data')!.tabColor).toBe('#e03131');

    // A tab colour used to be written straight onto the Sheet, which left one
    // edit the user could see and undo silently skipped.
    expect(controller.getSnapshot().canUndo).toBe(true);
    controller.undo();
    expect(controller.workbook.getSheet('Data')!.tabColor).toBeUndefined();
  });

  it('sorts only the used rows of a whole-column selection', () => {
    const controller = createController(undefined, { now: () => 1 });
    controller.commitEntry('2', { sheet: 'Sheet1', row: 0, col: 0 });
    controller.commitEntry('1', { sheet: 'Sheet1', row: 1, col: 0 });

    // A whole-column selection covers 1,048,576 addresses. Sorting the empty
    // million below the data would build a million row records and write a
    // million blanks into one transaction.
    controller.setSelection({
      sheet: 'Sheet1',
      active: { row: 0, col: 0 },
      ranges: [{ start: { row: 0, col: 0 }, end: { row: 1_048_575, col: 0 } }],
    });
    const started = Date.now();
    controller.sortSelection(0, 'asc');
    expect(Date.now() - started).toBeLessThan(2000);

    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe(1);
    expect(controller.cellAt({ sheet: 'Sheet1', row: 1, col: 0 })?.value).toBe(2);
    expect(controller.sheet()!.bounds()).toEqual({ minRow: 0, minCol: 0, maxRow: 1, maxCol: 0 });
  });
});

describe('compareScalars', () => {
  it('orders numbers before text', () => {
    expect(compareScalars(1, 'a')).toBeLessThan(0);
  });

  it('orders text before booleans', () => {
    expect(compareScalars('a', true)).toBeLessThan(0);
  });

  it('orders booleans before errors', () => {
    expect(compareScalars(true, CellError.NA)).toBeLessThan(0);
  });

  it('puts blanks last', () => {
    expect(compareScalars(null, CellError.NA)).toBeGreaterThan(0);
  });

  it('compares numbers numerically, not as text', () => {
    expect(compareScalars(9, 10)).toBeLessThan(0);
  });
});
