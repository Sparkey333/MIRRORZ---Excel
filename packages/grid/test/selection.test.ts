import { beforeEach, describe, expect, it, vi } from 'vitest';
import { MAX_COLS, MAX_ROWS, Workbook, parseRangeRef } from '@mirrorz/core';

import { Selection, makeRange, rangeCellCount, sheetSource } from '../src/selection.js';

/**
 * Data layout used by the navigation tests, chosen so every Ctrl+Arrow case is
 * reachable: a solid block, a gap, a second block, then nothing.
 *
 *      A   B   C   D   E
 *  1   x   x   x       x
 *  2   x       x
 *  3   x   x   x
 *  4
 *  5   x   x
 */
function dataSheet() {
  const wb = new Workbook();
  const s = wb.addSheet('S');
  const filled: [number, number][] = [
    [0, 0], [0, 1], [0, 2], [0, 4],
    [1, 0], [1, 2],
    [2, 0], [2, 1], [2, 2],
    [4, 0], [4, 1],
  ];
  for (const [r, c] of filled) s.setValue(r, c, 1);
  return s;
}

function sel(): Selection {
  return new Selection({ source: sheetSource(dataSheet()), pageRows: () => 10, pageCols: () => 4 });
}

function pos(s: Selection): [number, number] {
  return [s.active.row, s.active.col];
}

describe('Selection basics', () => {
  let s: Selection;
  beforeEach(() => {
    s = sel();
  });

  it('starts at A1 with a single-cell range', () => {
    expect(pos(s)).toEqual([0, 0]);
    expect(s.ranges).toEqual([{ top: 0, left: 0, bottom: 0, right: 0 }]);
  });

  it('selects a cell and collapses the range', () => {
    s.selectCell(5, 3);
    expect(pos(s)).toEqual([5, 3]);
    expect(rangeCellCount(s.ranges[0]!)).toBe(1);
  });

  it('clamps a selection to the sheet', () => {
    s.selectCell(-4, MAX_COLS + 10);
    expect(pos(s)).toEqual([0, MAX_COLS - 1]);
  });

  it('extends from the anchor', () => {
    s.selectCell(2, 2);
    s.extendTo(4, 5);
    expect(s.ranges[0]).toEqual({ top: 2, left: 2, bottom: 4, right: 5 });
    expect(pos(s)).toEqual([2, 2]);
  });

  it('extends backwards past the anchor', () => {
    s.selectCell(4, 4);
    s.extendTo(1, 1);
    expect(s.ranges[0]).toEqual({ top: 1, left: 1, bottom: 4, right: 4 });
  });

  it('adds a second range with the additive flag', () => {
    s.selectCell(0, 0);
    s.selectCell(3, 3, { additive: true });
    expect(s.ranges).toHaveLength(2);
    expect(s.isSelected(0, 0)).toBe(true);
    expect(s.isSelected(3, 3)).toBe(true);
    expect(s.isSelected(1, 1)).toBe(false);
  });

  it('collapses to the active cell on demand', () => {
    s.selectCell(0, 0);
    s.selectCell(3, 3, { additive: true });
    s.collapseToActive();
    expect(s.ranges).toHaveLength(1);
    expect(s.ranges[0]).toEqual({ top: 3, left: 3, bottom: 3, right: 3 });
  });

  it('selects a whole row', () => {
    s.selectRow(4);
    expect(s.ranges[0]).toEqual({ top: 4, left: 0, bottom: 4, right: MAX_COLS - 1 });
    expect(s.isEntireRowSelected(4)).toBe(true);
    expect(s.isEntireColSelected(0)).toBe(false);
  });

  it('selects a whole column', () => {
    s.selectCol(2);
    expect(s.ranges[0]).toEqual({ top: 0, left: 2, bottom: MAX_ROWS - 1, right: 2 });
    expect(s.isEntireColSelected(2)).toBe(true);
  });

  it('extends a row selection', () => {
    s.selectRow(2);
    s.selectRow(5, { extend: true });
    expect(s.ranges[0]).toEqual({ top: 2, left: 0, bottom: 5, right: MAX_COLS - 1 });
  });

  it('extends a column selection backwards', () => {
    s.selectCol(5);
    s.selectCol(2, { extend: true });
    expect(s.ranges[0]).toEqual({ top: 0, left: 2, bottom: MAX_ROWS - 1, right: 5 });
  });

  it('selects the whole sheet', () => {
    s.selectAll();
    expect(s.ranges[0]).toEqual({ top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 });
  });

  it('reports row and column intersection for header highlighting', () => {
    s.selectRange({ top: 2, left: 1, bottom: 4, right: 3 });
    expect(s.isRowSelected(3)).toBe(true);
    expect(s.isRowSelected(5)).toBe(false);
    expect(s.isColSelected(1)).toBe(true);
    expect(s.isColSelected(0)).toBe(false);
  });

  it('restores a saved state', () => {
    s.setState({ active: { row: 3, col: 4 }, ranges: [{ top: 3, left: 4, bottom: 6, right: 8 }] });
    expect(pos(s)).toEqual([3, 4]);
    expect(s.ranges[0]?.bottom).toBe(6);
  });

  it('normalises an inverted range', () => {
    s.selectRange({ top: 8, left: 9, bottom: 2, right: 1 });
    expect(s.ranges[0]).toEqual({ top: 2, left: 1, bottom: 8, right: 9 });
  });

  it('builds a range from two corners', () => {
    expect(makeRange({ row: 5, col: 5 }, { row: 1, col: 9 })).toEqual({
      top: 1,
      left: 5,
      bottom: 5,
      right: 9,
    });
  });

  it('notifies listeners of changes', () => {
    const spy = vi.fn();
    const off = s.onChange(spy);
    s.selectCell(1, 1);
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.mock.calls[0]?.[0].active).toEqual({ row: 1, col: 1 });
    off();
    s.selectCell(2, 2);
    expect(spy).toHaveBeenCalledTimes(1);
  });
});

describe('Selection - arrow keys', () => {
  let s: Selection;
  beforeEach(() => {
    s = sel();
    s.selectCell(2, 2);
  });

  it('moves one cell per arrow', () => {
    s.keyDown({ key: 'ArrowRight' });
    expect(pos(s)).toEqual([2, 3]);
    s.keyDown({ key: 'ArrowDown' });
    expect(pos(s)).toEqual([3, 3]);
    s.keyDown({ key: 'ArrowLeft' });
    expect(pos(s)).toEqual([3, 2]);
    s.keyDown({ key: 'ArrowUp' });
    expect(pos(s)).toEqual([2, 2]);
  });

  it('stops at the top-left corner', () => {
    s.selectCell(0, 0);
    s.keyDown({ key: 'ArrowUp' });
    s.keyDown({ key: 'ArrowLeft' });
    expect(pos(s)).toEqual([0, 0]);
  });

  it('stops at the bottom-right corner', () => {
    s.selectCell(MAX_ROWS - 1, MAX_COLS - 1);
    s.keyDown({ key: 'ArrowDown' });
    s.keyDown({ key: 'ArrowRight' });
    expect(pos(s)).toEqual([MAX_ROWS - 1, MAX_COLS - 1]);
  });

  it('collapses the selection when moving without shift', () => {
    s.extendTo(6, 6);
    s.keyDown({ key: 'ArrowDown' });
    expect(rangeCellCount(s.ranges[0]!)).toBe(1);
  });

  it('extends with shift and keeps the anchor', () => {
    s.keyDown({ key: 'ArrowDown', shift: true });
    s.keyDown({ key: 'ArrowRight', shift: true });
    expect(s.ranges[0]).toEqual({ top: 2, left: 2, bottom: 3, right: 3 });
    expect(s.anchor).toEqual({ row: 2, col: 2 });
    expect(pos(s)).toEqual([3, 3]);
  });

  it('shrinks a shift-extended range back through the anchor', () => {
    s.keyDown({ key: 'ArrowDown', shift: true });
    s.keyDown({ key: 'ArrowDown', shift: true });
    s.keyDown({ key: 'ArrowUp', shift: true });
    expect(s.ranges[0]).toEqual({ top: 2, left: 2, bottom: 3, right: 2 });
  });

  it('accepts the bare arrow key names too', () => {
    s.keyDown({ key: 'Down' });
    expect(pos(s)).toEqual([3, 2]);
  });

  it('reports unhandled keys', () => {
    expect(s.keyDown({ key: 'x' })).toBe(false);
    expect(s.keyDown({ key: 'ArrowUp' })).toBe(true);
  });
});

describe('Selection - Ctrl+Arrow edge jumps', () => {
  let s: Selection;
  beforeEach(() => {
    s = sel();
  });

  it('runs to the end of a contiguous block', () => {
    s.selectCell(0, 0);
    s.keyDown({ key: 'ArrowDown', ctrl: true });
    expect(pos(s)).toEqual([2, 0]);
  });

  it('skips a gap to the next block', () => {
    s.selectCell(2, 0);
    s.keyDown({ key: 'ArrowDown', ctrl: true });
    expect(pos(s)).toEqual([4, 0]);
  });

  it('runs to the sheet edge when nothing is left', () => {
    s.selectCell(4, 0);
    s.keyDown({ key: 'ArrowDown', ctrl: true });
    expect(pos(s)).toEqual([MAX_ROWS - 1, 0]);
  });

  it('jumps from an empty cell to the first filled one', () => {
    s.selectCell(3, 1);
    s.keyDown({ key: 'ArrowDown', ctrl: true });
    expect(pos(s)).toEqual([4, 1]);
  });

  it('jumps right across a gap', () => {
    s.selectCell(0, 2);
    s.keyDown({ key: 'ArrowRight', ctrl: true });
    expect(pos(s)).toEqual([0, 4]);
  });

  it('runs left to the start of a block', () => {
    s.selectCell(0, 2);
    s.keyDown({ key: 'ArrowLeft', ctrl: true });
    expect(pos(s)).toEqual([0, 0]);
  });

  it('runs to column A when nothing is to the left', () => {
    s.selectCell(1, 1);
    s.keyDown({ key: 'ArrowLeft', ctrl: true });
    expect(pos(s)).toEqual([1, 0]);
  });

  it('runs to row 1 when nothing is above', () => {
    s.selectCell(3, 3);
    s.keyDown({ key: 'ArrowUp', ctrl: true });
    expect(pos(s)).toEqual([0, 3]);
  });

  it('runs to the last column on an empty row', () => {
    s.selectCell(10, 0);
    s.keyDown({ key: 'ArrowRight', ctrl: true });
    expect(pos(s)).toEqual([10, MAX_COLS - 1]);
  });

  it('extends with Ctrl+Shift+Arrow', () => {
    s.selectCell(0, 0);
    s.keyDown({ key: 'ArrowDown', ctrl: true, shift: true });
    expect(s.ranges[0]).toEqual({ top: 0, left: 0, bottom: 2, right: 0 });
    expect(pos(s)).toEqual([2, 0]);
  });

  it('selects a whole block with two Ctrl+Shift jumps', () => {
    s.selectCell(0, 0);
    s.keyDown({ key: 'ArrowDown', ctrl: true, shift: true });
    s.keyDown({ key: 'ArrowRight', ctrl: true, shift: true });
    expect(s.ranges[0]).toEqual({ top: 0, left: 0, bottom: 2, right: 2 });
  });

  it('treats meta as ctrl', () => {
    s.selectCell(0, 0);
    s.keyDown({ key: 'ArrowDown', meta: true });
    expect(pos(s)).toEqual([2, 0]);
  });

  it('does not walk the whole grid when the data ends', () => {
    const spy = vi.fn(() => true);
    const counting = new Selection({
      source: { isEmpty: spy, lastRow: 20, lastCol: 20 },
    });
    counting.selectCell(0, 0);
    counting.keyDown({ key: 'ArrowDown', ctrl: true });
    expect(pos(counting)).toEqual([MAX_ROWS - 1, 0]);
    expect(spy.mock.calls.length).toBeLessThan(64);
  });
});

describe('Selection - Home, End and paging', () => {
  let s: Selection;
  beforeEach(() => {
    s = sel();
    s.selectCell(3, 5);
  });

  it('Home goes to column A of the current row', () => {
    s.keyDown({ key: 'Home' });
    expect(pos(s)).toEqual([3, 0]);
  });

  it('Ctrl+Home goes to A1', () => {
    s.keyDown({ key: 'Home', ctrl: true });
    expect(pos(s)).toEqual([0, 0]);
  });

  it('Shift+Home extends to column A', () => {
    s.keyDown({ key: 'Home', shift: true });
    expect(s.ranges[0]).toEqual({ top: 3, left: 0, bottom: 3, right: 5 });
  });

  it('Ctrl+End goes to the last used cell', () => {
    s.keyDown({ key: 'End', ctrl: true });
    expect(pos(s)).toEqual([4, 4]);
  });

  it('Ctrl+End on an empty sheet goes to A1', () => {
    const empty = new Selection({ source: { isEmpty: () => true, lastRow: -1, lastCol: -1 } });
    empty.selectCell(9, 9);
    empty.keyDown({ key: 'End', ctrl: true });
    expect(pos(empty)).toEqual([0, 0]);
  });

  it('a bare End turns on End mode without moving', () => {
    s.keyDown({ key: 'End' });
    expect(s.endMode).toBe(true);
    expect(pos(s)).toEqual([3, 5]);
  });

  it('End then an arrow behaves as Ctrl+Arrow', () => {
    s.selectCell(0, 0);
    s.keyDown({ key: 'End' });
    s.keyDown({ key: 'ArrowDown' });
    expect(pos(s)).toEqual([2, 0]);
    expect(s.endMode).toBe(false);
  });

  it('End then Home is Ctrl+End', () => {
    s.keyDown({ key: 'End' });
    s.keyDown({ key: 'Home' });
    expect(pos(s)).toEqual([4, 4]);
    expect(s.endMode).toBe(false);
  });

  it('End toggles off when pressed twice', () => {
    s.keyDown({ key: 'End' });
    s.keyDown({ key: 'End' });
    expect(s.endMode).toBe(false);
  });

  it('Escape cancels End mode', () => {
    s.keyDown({ key: 'End' });
    expect(s.keyDown({ key: 'Escape' })).toBe(true);
    expect(s.endMode).toBe(false);
    expect(s.keyDown({ key: 'Escape' })).toBe(false);
  });

  it('Page Down moves by the page height', () => {
    s.keyDown({ key: 'PageDown' });
    expect(pos(s)).toEqual([13, 5]);
  });

  it('Page Up moves back and clamps at the top', () => {
    s.keyDown({ key: 'PageUp' });
    expect(pos(s)).toEqual([0, 5]);
  });

  it('Shift+Page Down extends', () => {
    s.keyDown({ key: 'PageDown', shift: true });
    expect(s.ranges[0]).toEqual({ top: 3, left: 5, bottom: 13, right: 5 });
  });

  it('Alt+Page Down moves a screen to the right', () => {
    s.keyDown({ key: 'PageDown', alt: true });
    expect(pos(s)).toEqual([3, 9]);
  });

  it('Alt+Page Up moves a screen to the left', () => {
    s.keyDown({ key: 'PageUp', alt: true });
    expect(pos(s)).toEqual([3, 1]);
  });
});

describe('Selection - Tab and Enter', () => {
  let s: Selection;
  beforeEach(() => {
    s = sel();
  });

  it('Tab moves right on a single cell', () => {
    s.selectCell(1, 1);
    s.keyDown({ key: 'Tab' });
    expect(pos(s)).toEqual([1, 2]);
  });

  it('Shift+Tab moves left', () => {
    s.selectCell(1, 2);
    s.keyDown({ key: 'Tab', shift: true });
    expect(pos(s)).toEqual([1, 1]);
  });

  it('Enter moves down on a single cell', () => {
    s.selectCell(1, 1);
    s.keyDown({ key: 'Enter' });
    expect(pos(s)).toEqual([2, 1]);
  });

  it('Shift+Enter moves up', () => {
    s.selectCell(2, 1);
    s.keyDown({ key: 'Enter', shift: true });
    expect(pos(s)).toEqual([1, 1]);
  });

  it('Enter returns to the column the Tab run started in', () => {
    s.selectCell(1, 1);
    s.keyDown({ key: 'Tab' });
    s.keyDown({ key: 'Tab' });
    s.keyDown({ key: 'Tab' });
    expect(pos(s)).toEqual([1, 4]);
    s.keyDown({ key: 'Enter' });
    expect(pos(s)).toEqual([2, 1]);
  });

  it('forgets the entry column after an arrow key', () => {
    s.selectCell(1, 1);
    s.keyDown({ key: 'Tab' });
    s.keyDown({ key: 'ArrowRight' });
    s.keyDown({ key: 'Enter' });
    expect(pos(s)).toEqual([2, 3]);
  });

  it('Tab wraps inside a selected block', () => {
    s.selectRange({ top: 1, left: 1, bottom: 2, right: 2 });
    expect(pos(s)).toEqual([1, 1]);
    s.keyDown({ key: 'Tab' });
    expect(pos(s)).toEqual([1, 2]);
    s.keyDown({ key: 'Tab' });
    expect(pos(s)).toEqual([2, 1]);
    s.keyDown({ key: 'Tab' });
    expect(pos(s)).toEqual([2, 2]);
  });

  it('Tab wraps around the end of a block back to its start', () => {
    s.selectRange({ top: 1, left: 1, bottom: 2, right: 2 });
    for (let i = 0; i < 4; i++) s.keyDown({ key: 'Tab' });
    expect(pos(s)).toEqual([1, 1]);
  });

  it('Shift+Tab wraps backwards inside a block', () => {
    s.selectRange({ top: 1, left: 1, bottom: 2, right: 2 });
    s.keyDown({ key: 'Tab', shift: true });
    expect(pos(s)).toEqual([2, 2]);
  });

  it('Enter walks down the columns of a block', () => {
    s.selectRange({ top: 1, left: 1, bottom: 2, right: 2 });
    s.keyDown({ key: 'Enter' });
    expect(pos(s)).toEqual([2, 1]);
    s.keyDown({ key: 'Enter' });
    expect(pos(s)).toEqual([1, 2]);
    s.keyDown({ key: 'Enter' });
    expect(pos(s)).toEqual([2, 2]);
    s.keyDown({ key: 'Enter' });
    expect(pos(s)).toEqual([1, 1]);
  });

  it('Shift+Enter walks back up', () => {
    s.selectRange({ top: 1, left: 1, bottom: 2, right: 2 });
    s.keyDown({ key: 'Enter', shift: true });
    expect(pos(s)).toEqual([2, 2]);
  });

  it('never leaves the block', () => {
    s.selectRange({ top: 3, left: 3, bottom: 4, right: 4 });
    for (let i = 0; i < 25; i++) {
      s.keyDown({ key: 'Tab' });
      expect(s.isSelected(s.active.row, s.active.col)).toBe(true);
    }
  });

  it('moves on to the next range when a block runs out', () => {
    s.selectRange({ top: 0, left: 0, bottom: 0, right: 1 });
    s.selectRange({ top: 5, left: 5, bottom: 5, right: 5 }, { additive: true });
    s.setState({
      active: { row: 0, col: 0 },
      ranges: [
        { top: 0, left: 0, bottom: 0, right: 1 },
        { top: 5, left: 5, bottom: 5, right: 5 },
      ],
      activeRange: 0,
    });
    s.keyDown({ key: 'Tab' });
    expect(pos(s)).toEqual([0, 1]);
    s.keyDown({ key: 'Tab' });
    expect(pos(s)).toEqual([5, 5]);
    s.keyDown({ key: 'Tab' });
    expect(pos(s)).toEqual([0, 0]);
  });

  it('keeps the ranges intact while cycling', () => {
    s.selectRange({ top: 1, left: 1, bottom: 3, right: 3 });
    const before = JSON.stringify(s.ranges);
    s.keyDown({ key: 'Tab' });
    s.keyDown({ key: 'Enter' });
    expect(JSON.stringify(s.ranges)).toBe(before);
  });
});

describe('Selection - shortcuts', () => {
  let s: Selection;
  beforeEach(() => {
    s = sel();
    s.selectCell(4, 3);
  });

  it('Ctrl+Space selects the column', () => {
    expect(s.keyDown({ key: ' ', ctrl: true })).toBe(true);
    expect(s.isEntireColSelected(3)).toBe(true);
  });

  it('Shift+Space selects the row', () => {
    s.keyDown({ key: ' ', shift: true });
    expect(s.isEntireRowSelected(4)).toBe(true);
  });

  it('Ctrl+Shift+Space selects everything', () => {
    s.keyDown({ key: ' ', ctrl: true, shift: true });
    expect(s.ranges[0]).toEqual({ top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 });
  });

  it('a plain space is not consumed', () => {
    expect(s.keyDown({ key: ' ' })).toBe(false);
  });

  it('Ctrl+A selects everything', () => {
    expect(s.keyDown({ key: 'a', ctrl: true })).toBe(true);
    expect(rangeCellCount(s.ranges[0]!)).toBe(MAX_ROWS * MAX_COLS);
  });

  it('a plain a is not consumed', () => {
    expect(s.keyDown({ key: 'a' })).toBe(false);
  });
});

describe('Selection - hidden rows and columns', () => {
  it('skips hidden rows when arrowing', () => {
    const wb = new Workbook();
    const sh = wb.addSheet('S');
    sh.rows.set(1, { hidden: true });
    sh.rows.set(2, { hidden: true });
    const s = new Selection({ source: sheetSource(sh) });
    s.selectCell(0, 0);
    s.keyDown({ key: 'ArrowDown' });
    expect(pos(s)).toEqual([3, 0]);
  });

  it('skips hidden columns in both directions', () => {
    const wb = new Workbook();
    const sh = wb.addSheet('S');
    sh.cols.set(1, { hidden: true });
    const s = new Selection({ source: sheetSource(sh) });
    s.selectCell(0, 0);
    s.keyDown({ key: 'ArrowRight' });
    expect(pos(s)).toEqual([0, 2]);
    s.keyDown({ key: 'ArrowLeft' });
    expect(pos(s)).toEqual([0, 0]);
  });
});

describe('Selection - merged cells', () => {
  function mergedSheet() {
    const wb = new Workbook();
    const sh = wb.addSheet('S');
    sh.merges = [{ range: parseRangeRef('B2:D3')! }];
    return sh;
  }

  it('puts the cursor on the merge anchor wherever it is clicked', () => {
    const s = new Selection({ source: sheetSource(mergedSheet()) });
    s.selectCell(2, 3);
    expect(pos(s)).toEqual([1, 1]);
    expect(s.ranges[0]).toEqual({ top: 1, left: 1, bottom: 2, right: 3 });
  });

  it('steps out of a merge from its far edge', () => {
    const s = new Selection({ source: sheetSource(mergedSheet()) });
    s.selectCell(1, 1);
    s.keyDown({ key: 'ArrowRight' });
    expect(pos(s)).toEqual([1, 4]);
  });

  it('steps down out of a merge from its bottom edge', () => {
    const s = new Selection({ source: sheetSource(mergedSheet()) });
    s.selectCell(1, 1);
    s.keyDown({ key: 'ArrowDown' });
    expect(pos(s)).toEqual([3, 1]);
  });

  it('grows a range so it never cuts a merge in half', () => {
    const s = new Selection({ source: sheetSource(mergedSheet()) });
    s.selectCell(0, 0);
    s.extendTo(2, 2);
    expect(s.ranges[0]).toEqual({ top: 0, left: 0, bottom: 2, right: 3 });
  });
});

/**
 * The cursor may sit anywhere in a 1,048,576-row sheet, and every one of these
 * moves is a single keystroke. A scan that starts from where the cursor happens
 * to be, rather than from where the data is, freezes the app for a keypress.
 */
describe('Selection - navigation never walks the whole grid', () => {
  function countingSource(lastRow: number, lastCol: number) {
    const probes = { n: 0 };
    const source = {
      isEmpty: () => {
        probes.n++;
        return true;
      },
      lastRow,
      lastCol,
    };
    return { probes, source };
  }

  it('does not probe the empty rows below the data on Ctrl+Down', () => {
    const { probes, source } = countingSource(20, 20);
    const s = new Selection({ source });
    s.selectCell(900_000, 0);
    probes.n = 0;
    s.keyDown({ key: 'ArrowDown', ctrl: true });
    expect(probes.n).toBeLessThan(64);
    // Excel runs to the bottom of the sheet from below the data.
    expect(pos(s)).toEqual([MAX_ROWS - 1, 0]);
  });

  it('does not probe the empty rows below the data on Ctrl+Up', () => {
    const { probes, source } = countingSource(20, 20);
    const s = new Selection({ source });
    s.selectCell(900_000, 0);
    probes.n = 0;
    s.keyDown({ key: 'ArrowUp', ctrl: true });
    expect(probes.n).toBeLessThan(64);
    expect(pos(s)).toEqual([0, 0]);
  });

  it('does not probe the empty columns right of the data on Ctrl+Right', () => {
    const { probes, source } = countingSource(20, 20);
    const s = new Selection({ source });
    s.selectCell(0, 10_000);
    probes.n = 0;
    s.keyDown({ key: 'ArrowRight', ctrl: true });
    expect(probes.n).toBeLessThan(64);
    expect(pos(s)).toEqual([0, MAX_COLS - 1]);
  });

  it('still stops on the first filled cell above a far-off cursor', () => {
    const wb = new Workbook();
    const sh = wb.addSheet('S');
    sh.setValue(5, 0, 1);
    const s = new Selection({ source: sheetSource(sh) });
    s.selectCell(500_000, 0);
    s.keyDown({ key: 'ArrowUp', ctrl: true });
    expect(pos(s)).toEqual([5, 0]);
  });
});

describe('Selection - merge expansion cost', () => {
  it('does not probe every selected cell when growing over merges', () => {
    const mergeAt = vi.fn(() => undefined);
    const s = new Selection({
      source: { isEmpty: () => true, lastRow: 500, lastCol: 500, mergeAt },
    });
    s.selectCell(0, 0);
    mergeAt.mockClear();
    s.extendTo(199, 199);
    // The perimeter of a 200x200 rectangle, not its area.
    expect(mergeAt.mock.calls.length).toBeLessThan(1000);
  });

  it('still grows over a merge that only the interior of the drag touches', () => {
    const merge = { top: 100, left: 100, bottom: 100, right: 400 };
    const s = new Selection({
      source: {
        isEmpty: () => true,
        lastRow: 500,
        lastCol: 500,
        mergeAt: (row: number, col: number) =>
          row === merge.top && col >= merge.left && col <= merge.right ? merge : undefined,
      },
    });
    s.selectCell(0, 0);
    s.extendTo(199, 199);
    expect(s.ranges[0]).toEqual({ top: 0, left: 0, bottom: 199, right: 400 });
  });

  it('uses the sheet merge list when one is offered, with no per-cell probing', () => {
    const wb = new Workbook();
    const sh = wb.addSheet('S');
    sh.merges = [{ range: parseRangeRef('C150:H250')! }];
    const source = sheetSource(sh);
    const probe = vi.fn(source.mergeAt!);
    const s = new Selection({ source: { ...source, mergeAt: probe } });
    s.selectCell(0, 0);
    probe.mockClear();
    s.extendTo(199, 199);
    expect(s.ranges[0]).toEqual({ top: 0, left: 0, bottom: 249, right: 199 });
    expect(probe).not.toHaveBeenCalled();
  });
});

describe('Selection - a block run stops at the used extent', () => {
  it('does not walk the sheet when the source claims every cell is filled', () => {
    // A source whose extent and contents disagree - a stale bounds cache, or a
    // test double - must still not cost a million probes.
    const probes = { n: 0 };
    const s = new Selection({
      source: {
        isEmpty: () => {
          probes.n++;
          return false;
        },
        lastRow: 30,
        lastCol: 30,
      },
    });
    s.selectCell(0, 0);
    probes.n = 0;
    s.keyDown({ key: 'ArrowDown', ctrl: true });
    expect(probes.n).toBeLessThan(64);
    expect(pos(s)).toEqual([30, 0]);
  });
});
