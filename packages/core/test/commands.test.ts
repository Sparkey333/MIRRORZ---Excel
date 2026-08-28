import { describe, expect, it } from 'vitest';
import { Document } from '../src/commands.js';
import { Workbook } from '../src/sheet.js';

function doc(): Document {
  const wb = new Workbook();
  wb.addSheet('Sheet1');
  return new Document(wb);
}

describe('basic editing and undo', () => {
  it('undoes and redoes a single edit', () => {
    const d = doc();
    d.setValue('Sheet1', 0, 0, 42);
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(42);

    d.undo();
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(null);

    d.redo();
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(42);
  });

  it('restores the previous value rather than clearing', () => {
    const d = doc();
    d.setValue('Sheet1', 0, 0, 'first');
    d.setValue('Sheet1', 0, 0, 'second');
    d.undo();
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe('first');
  });

  it('undoes a long run of edits one at a time', () => {
    const d = doc();
    for (let i = 0; i < 100; i++) d.setValue('Sheet1', i, 0, i);
    for (let i = 99; i >= 0; i--) {
      expect(d.workbook.getSheet('Sheet1')!.getValue(i, 0)).toBe(i);
      d.undo();
      expect(d.workbook.getSheet('Sheet1')!.getValue(i, 0)).toBe(null);
    }
    expect(d.canUndo).toBe(false);
  });

  it('is unlimited: there is no stack depth to exhaust', () => {
    const d = doc();
    for (let i = 0; i < 5000; i++) d.setValue('Sheet1', 0, 0, i);
    for (let i = 0; i < 5000; i++) d.undo();
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(null);
    expect(d.canUndo).toBe(false);
  });

  it('reports what would be undone', () => {
    const d = doc();
    expect(d.canUndo).toBe(false);
    d.setValue('Sheet1', 0, 0, 1, { label: 'Type 1 in A1' });
    expect(d.peekUndo()?.label).toBe('Type 1 in A1');
  });

  it('undoes a formula back to nothing', () => {
    const d = doc();
    d.setFormula('Sheet1', 0, 0, 'SUM(B1:B9)', 10);
    expect(d.workbook.getSheet('Sheet1')!.getCell(0, 0)?.formula).toBe('SUM(B1:B9)');
    d.undo();
    expect(d.workbook.getSheet('Sheet1')!.getCell(0, 0)).toBeUndefined();
  });
});

describe('transactions - the macro-survival property', () => {
  it('collapses a thousand edits into one undo entry', () => {
    const d = doc();
    d.transact({ label: 'Run macro FormatReport', origin: 'macro' }, () => {
      for (let i = 0; i < 1000; i++) d.setValue('Sheet1', i, 0, i);
    });

    expect(d.workbook.getSheet('Sheet1')!.cellCount).toBe(1000);
    expect(d.history()).toHaveLength(1);

    // This is the behaviour Excel has never had: one undo takes back the whole
    // macro run rather than the macro having wiped the stack entirely.
    d.undo();
    expect(d.workbook.getSheet('Sheet1')!.cellCount).toBe(0);

    d.redo();
    expect(d.workbook.getSheet('Sheet1')!.cellCount).toBe(1000);
  });

  it('labels the entry with what ran, and records who ran it', () => {
    const d = doc();
    d.transact({ label: 'Run macro Recalculate', origin: 'macro' }, () => {
      d.setValue('Sheet1', 0, 0, 1);
    });
    const entry = d.peekUndo()!;
    expect(entry.label).toBe('Run macro Recalculate');
    expect(entry.origin).toBe('macro');
  });

  it('nested transactions join the outermost one', () => {
    const d = doc();
    d.transact({ label: 'outer' }, () => {
      d.setValue('Sheet1', 0, 0, 1);
      d.transact({ label: 'inner' }, () => {
        d.setValue('Sheet1', 1, 0, 2);
      });
      d.setValue('Sheet1', 2, 0, 3);
    });
    expect(d.history()).toHaveLength(1);
    expect(d.peekUndo()!.label).toBe('outer');
    d.undo();
    expect(d.workbook.getSheet('Sheet1')!.cellCount).toBe(0);
  });

  it('rolls back cleanly when the body throws', () => {
    const d = doc();
    d.setValue('Sheet1', 0, 0, 'before');
    expect(() =>
      d.transact({ label: 'failing script' }, () => {
        d.setValue('Sheet1', 0, 0, 'during');
        d.setValue('Sheet1', 1, 0, 'also during');
        throw new Error('script failed');
      }),
    ).toThrow('script failed');

    // A script that dies halfway must not leave a state nobody asked for.
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe('before');
    expect(d.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe(null);
    expect(d.history()).toHaveLength(1);
  });

  it('records nothing for a transaction that changed nothing', () => {
    const d = doc();
    d.transact({ label: 'no-op' }, () => {});
    expect(d.canUndo).toBe(false);
  });

  it('returns the body result', () => {
    const d = doc();
    expect(d.transact({ label: 'x' }, () => 7)).toBe(7);
  });
});

describe('structural operations are ordinary commands', () => {
  it('undoes adding a sheet', () => {
    const d = doc();
    d.addSheet('New');
    expect(d.workbook.getSheet('New')).toBeDefined();
    d.undo();
    expect(d.workbook.getSheet('New')).toBeUndefined();
  });

  it('undoes deleting a sheet, contents and all', () => {
    const d = doc();
    const sheet = d.addSheet('Data');
    d.setValue('Data', 0, 0, 'kept');
    d.setValue('Data', 5, 3, 99);
    d.setRowProps('Data', 0, { height: 40 });
    sheet.tabColor = 'FFFF0000';

    d.removeSheet('Data');
    expect(d.workbook.getSheet('Data')).toBeUndefined();

    // Excel treats deleting a sheet as unundoable. Snapshotting first is all it
    // takes to make it ordinary.
    d.undo();
    const restored = d.workbook.getSheet('Data')!;
    expect(restored.getValue(0, 0)).toBe('kept');
    expect(restored.getValue(5, 3)).toBe(99);
    expect(restored.rowHeight(0)).toBe(40);
    expect(restored.tabColor).toBe('FFFF0000');
  });

  it('undoes a tab colour, and clearing one', () => {
    const d = doc();
    d.addSheet('Data');

    d.setSheetColor('Data', 'FFFF0000');
    expect(d.workbook.getSheet('Data')!.tabColor).toBe('FFFF0000');
    d.undo();
    expect(d.workbook.getSheet('Data')!.tabColor).toBeUndefined();

    d.redo();
    expect(d.workbook.getSheet('Data')!.tabColor).toBe('FFFF0000');

    // Clearing is an edit like any other and has to come back the same way.
    d.setSheetColor('Data', undefined);
    expect(d.workbook.getSheet('Data')!.tabColor).toBeUndefined();
    d.undo();
    expect(d.workbook.getSheet('Data')!.tabColor).toBe('FFFF0000');
  });

  it('does not log a tab colour that is already set', () => {
    const d = doc();
    d.addSheet('Data');
    d.setSheetColor('Data', 'FF00FF00');
    const before = d.allEntries().length;
    d.setSheetColor('Data', 'FF00FF00');
    expect(d.allEntries().length).toBe(before);
  });

  it('restores a deleted sheet to its original position', () => {
    const d = doc();
    d.addSheet('A');
    d.addSheet('B');
    d.addSheet('C');
    d.removeSheet('B');
    expect(d.workbook.sheets.map((s) => s.name)).toEqual(['Sheet1', 'A', 'C']);
    d.undo();
    expect(d.workbook.sheets.map((s) => s.name)).toEqual(['Sheet1', 'A', 'B', 'C']);
  });

  it('undoes a rename', () => {
    const d = doc();
    d.renameSheet('Sheet1', 'Renamed');
    expect(d.workbook.getSheet('Renamed')).toBeDefined();
    d.undo();
    expect(d.workbook.getSheet('Sheet1')).toBeDefined();
    expect(d.workbook.getSheet('Renamed')).toBeUndefined();
  });

  it('undoes hiding a sheet', () => {
    const d = doc();
    d.setSheetVisibility('Sheet1', 'hidden');
    expect(d.workbook.getSheet('Sheet1')!.visibility).toBe('hidden');
    d.undo();
    expect(d.workbook.getSheet('Sheet1')!.visibility).toBe('visible');
  });

  it('undoes a sheet move', () => {
    const d = doc();
    d.addSheet('A');
    d.addSheet('B');
    d.moveSheet('B', 0);
    expect(d.workbook.sheets.map((s) => s.name)).toEqual(['B', 'Sheet1', 'A']);
    d.undo();
    expect(d.workbook.sheets.map((s) => s.name)).toEqual(['Sheet1', 'A', 'B']);
  });

  it('undoes row and column property changes', () => {
    const d = doc();
    d.setRowProps('Sheet1', 3, { height: 50, hidden: true });
    d.setColProps('Sheet1', 2, { width: 30 });
    expect(d.workbook.getSheet('Sheet1')!.rowHeight(3)).toBe(50);
    d.undo();
    d.undo();
    expect(d.workbook.getSheet('Sheet1')!.rows.get(3)).toBeUndefined();
    expect(d.workbook.getSheet('Sheet1')!.cols.get(2)).toBeUndefined();
  });
});

describe('history as a tree, not a stack', () => {
  it('keeps the abandoned branch after undoing and doing something else', () => {
    const d = doc();
    d.setValue('Sheet1', 0, 0, 'branch A', { label: 'A' });
    d.undo();
    d.setValue('Sheet1', 0, 0, 'branch B', { label: 'B' });

    // A stack would have discarded A. Both branches are still reachable.
    const labels = d.allEntries().map((e) => e.label);
    expect(labels).toContain('A');
    expect(labels).toContain('B');
  });

  it('can return to the abandoned branch', () => {
    const d = doc();
    d.setValue('Sheet1', 0, 0, 'branch A', { label: 'A' });
    const branchA = d.peekUndo()!.id;
    d.undo();
    d.setValue('Sheet1', 0, 0, 'branch B', { label: 'B' });
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe('branch B');

    d.jumpTo(branchA);
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe('branch A');
  });

  it('jumps directly to any point without stepping', () => {
    const d = doc();
    const ids: number[] = [];
    for (let i = 0; i < 10; i++) {
      d.setValue('Sheet1', 0, 0, i, { label: `step ${i}` });
      ids.push(d.peekUndo()!.id);
    }
    // A history panel needs this: click row three, land on state three.
    d.jumpTo(ids[2]!);
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(2);
    d.jumpTo(ids[8]!);
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(8);
    d.jumpTo(null);
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(null);
  });

  it('jumps across branches through the common ancestor', () => {
    const d = doc();
    d.setValue('Sheet1', 0, 0, 'base', { label: 'base' });
    const base = d.peekUndo()!.id;
    d.setValue('Sheet1', 1, 0, 'A1', { label: 'A1' });
    const a1 = d.peekUndo()!.id;

    d.jumpTo(base);
    d.setValue('Sheet1', 2, 0, 'B1', { label: 'B1' });
    const b1 = d.peekUndo()!.id;

    // On branch B, A's cell must be absent and base's must remain.
    expect(d.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe('base');
    expect(d.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe(null);
    expect(d.workbook.getSheet('Sheet1')!.getValue(2, 0)).toBe('B1');

    d.jumpTo(a1);
    expect(d.workbook.getSheet('Sheet1')!.getValue(1, 0)).toBe('A1');
    expect(d.workbook.getSheet('Sheet1')!.getValue(2, 0)).toBe(null);

    d.jumpTo(b1);
    expect(d.workbook.getSheet('Sheet1')!.getValue(2, 0)).toBe('B1');
  });

  it('rejects a jump to an unknown entry', () => {
    const d = doc();
    expect(d.jumpTo(999)).toBe(false);
  });

  it('lists the current path oldest first', () => {
    const d = doc();
    d.setValue('Sheet1', 0, 0, 1, { label: 'one' });
    d.setValue('Sheet1', 0, 1, 2, { label: 'two' });
    d.setValue('Sheet1', 0, 2, 3, { label: 'three' });
    expect(d.history().map((e) => e.label)).toEqual(['one', 'two', 'three']);
  });
});

describe('labels', () => {
  it('derives a sensible default when none is given', () => {
    const d = doc();
    d.transact({}, () => {
      d.setValue('Sheet1', 0, 0, 1);
      d.setValue('Sheet1', 1, 0, 2);
      d.setValue('Sheet1', 2, 0, 3);
    });
    expect(d.peekUndo()!.label).toBe('Edit 3 cells');
  });

  it('names a structural change specifically', () => {
    const d = doc();
    d.addSheet('Budget');
    expect(d.peekUndo()!.label).toBe('Add sheet Budget');
  });
});

describe('change notification', () => {
  it('notifies subscribers of applied changes', () => {
    const d = doc();
    const seen: number[] = [];
    d.onChange((changes) => seen.push(changes.length));
    d.setValue('Sheet1', 0, 0, 1);
    d.transact({}, () => {
      d.setValue('Sheet1', 1, 0, 2);
      d.setValue('Sheet1', 2, 0, 3);
    });
    expect(seen).toEqual([1, 2]);
  });

  it('notifies on undo as well, so the renderer stays in step', () => {
    const d = doc();
    d.setValue('Sheet1', 0, 0, 1);
    let notified = 0;
    d.onChange(() => notified++);
    d.undo();
    d.redo();
    expect(notified).toBe(2);
  });

  it('stops notifying after unsubscribe', () => {
    const d = doc();
    let notified = 0;
    const off = d.onChange(() => notified++);
    d.setValue('Sheet1', 0, 0, 1);
    off();
    d.setValue('Sheet1', 0, 1, 2);
    expect(notified).toBe(1);
  });
});

describe('replay barriers', () => {
  it('marks an entry that cannot be faithfully reversed', () => {
    const d = doc();
    d.transact({ label: 'Refresh external data', barrier: true }, () => {
      d.setValue('Sheet1', 0, 0, 'fetched');
    });
    expect(d.peekUndo()!.barrier).toBe(true);
  });
});

describe('isolation between documents', () => {
  it('keeps each document history to itself', () => {
    const a = doc();
    const b = doc();
    a.setValue('Sheet1', 0, 0, 'in a');
    b.setValue('Sheet1', 0, 0, 'in b');

    // Excel shares one undo stack across every open workbook, so undoing in one
    // file can silently reverse an edit in another. Ours cannot.
    a.undo();
    expect(a.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe(null);
    expect(b.workbook.getSheet('Sheet1')!.getValue(0, 0)).toBe('in b');
    expect(b.canUndo).toBe(true);
  });
});
