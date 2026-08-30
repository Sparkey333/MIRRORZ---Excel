import { describe, expect, it } from 'vitest';
import {
  CellError,
  MAX_COLS,
  MAX_ROWS,
  type Sheet,
  Workbook,
  parseRangeRef,
} from '@mirrorz/core';

import {
  DARK_THEME,
  LIGHT_THEME,
  type GridTheme,
  type RenderStats,
  TextMeasureCache,
  Viewport,
  borderStroke,
  buildColAxis,
  buildRowAxis,
  crisp,
  effectiveHAlign,
  fillColorOf,
  formatCodeOf,
  generalAlign,
  renderGrid,
} from '../src/index.js';
import type { SelectionState } from '../src/selection.js';
import { FakeContext } from './fake-canvas.js';

const HEADER_W = 46;
const HEADER_H = 20;
const ROW_H = 20;
const COL_W = 64;
const PAD = 3;
/** FakeContext measures 0.5 * fontSize per character; the default font is 11px. */
const CHAR = 5.5;

interface Harness {
  wb: Workbook;
  sheet: Sheet;
  view: Viewport;
  ctx: FakeContext;
  measure: TextMeasureCache;
  paint(overrides?: Partial<Parameters<typeof renderGrid>[1]>): RenderStats;
}

function setup(options: {
  width?: number;
  height?: number;
  dpr?: number;
  theme?: GridTheme;
  frozenRows?: number;
  frozenCols?: number;
} = {}): Harness {
  const wb = new Workbook();
  const sheet = wb.addSheet('Sheet1');
  const ctx = new FakeContext();
  const measure = new TextMeasureCache();

  const harness: Harness = {
    wb,
    sheet,
    ctx,
    measure,
    view: null as unknown as Viewport,
    paint(overrides = {}) {
      const view = new Viewport({
        rows: buildRowAxis(sheet),
        cols: buildColAxis(sheet),
        width: options.width ?? 400,
        height: options.height ?? 200,
        dpr: options.dpr ?? 1,
        frozenRows: options.frozenRows ?? 0,
        frozenCols: options.frozenCols ?? 0,
      });
      harness.view = view;
      return renderGrid(ctx, {
        sheet,
        styles: wb.styles,
        layout: view.layout(),
        geometry: view,
        theme: options.theme ?? LIGHT_THEME,
        measure,
        ...overrides,
      });
    },
  };
  return harness;
}

function textCall(ctx: FakeContext, text: string) {
  return ctx.ops('fillText').find((c) => c.args[0] === text);
}

describe('renderGrid - bounded work', () => {
  it('paints only what is visible over a full-size sheet', () => {
    const h = setup({ width: 1000, height: 1000 });
    // A cell every 500 rows down the whole sheet: none of them may be touched
    // except the handful actually on screen.
    for (let r = 0; r < MAX_ROWS; r += 500) h.sheet.setValue(r, 0, r);
    const stats = h.paint();

    const maxRows = Math.ceil((1000 - HEADER_H) / ROW_H) + 1;
    const maxCols = Math.ceil((1000 - HEADER_W) / COL_W) + 1;
    expect(stats.cellsVisited).toBeLessThanOrEqual(maxRows * maxCols);
    expect(stats.cellsVisited).toBeLessThan(1500);
    expect(stats.texts).toBeLessThan(200);
  });

  it('costs the same at row 900,000 as at row 0', () => {
    const h = setup({ width: 1000, height: 1000 });
    for (let r = 0; r < MAX_ROWS; r += 500) h.sheet.setValue(r, 0, r);
    const top = h.paint();
    h.ctx.reset();
    h.view.setScroll(0, 18_000_000);
    const deep = renderGrid(h.ctx, {
      sheet: h.sheet,
      styles: h.wb.styles,
      layout: h.view.layout(),
      geometry: h.view,
      theme: LIGHT_THEME,
      measure: h.measure,
    });
    expect(deep.cellsVisited).toBe(top.cellsVisited);
  });

  it('does not grow with the number of cells off screen', () => {
    const a = setup();
    a.sheet.setValue(0, 0, 1);
    const small = a.paint();
    const b = setup();
    for (let r = 0; r < 20_000; r++) b.sheet.setValue(r + 100, 0, r);
    const large = b.paint();
    expect(large.cellsVisited).toBe(small.cellsVisited);
  });

  it('scales the frame by the device pixel ratio exactly once', () => {
    const h = setup({ dpr: 2 });
    h.paint();
    const transforms = h.ctx.ops('setTransform');
    expect(transforms).toHaveLength(1);
    expect(transforms[0]?.args).toEqual([2, 0, 0, 2, 0, 0]);
  });

  it('paints the background across the whole canvas', () => {
    const h = setup({ width: 400, height: 200 });
    h.paint();
    const first = h.ctx.ops('fillRect')[0];
    expect(first?.args).toEqual([0, 0, 400, 200]);
    expect(first?.fillStyle).toBe(LIGHT_THEME.background);
  });
});

describe('renderGrid - text measurement cache', () => {
  it('reuses measurements across frames', () => {
    const h = setup();
    for (let r = 0; r < 5; r++) h.sheet.setValue(r, 0, 'label');
    const first = h.paint();
    expect(first.measureCalls).toBeGreaterThan(0);
    const before = h.ctx.measureCount;
    const second = h.paint();
    expect(second.measureCalls).toBe(0);
    expect(h.ctx.measureCount).toBe(before);
  });

  it('measures a repeated string only once', () => {
    const repeated = setup();
    for (let r = 0; r < 8; r++) repeated.sheet.setValue(r, 0, 'same');
    const repeatedStats = repeated.paint();

    const distinct = setup();
    for (let r = 0; r < 8; r++) distinct.sheet.setValue(r, 0, `text${r}`);
    const distinctStats = distinct.paint();

    expect(distinctStats.measureCalls - repeatedStats.measureCalls).toBe(7);
    expect(repeated.measure.stats.hits).toBeGreaterThanOrEqual(7);
  });

  it('assigns ctx.font far less often than it measures', () => {
    const h = setup();
    for (let r = 0; r < 8; r++) h.sheet.setValue(r, 0, `text-${r}`);
    h.paint();
    const s = h.measure.stats;
    expect(s.fontSwitches).toBeLessThan(s.hits + s.misses);
  });

  it('drops the cache when the theme changes the font', () => {
    const h = setup();
    h.sheet.setValue(0, 0, 'x');
    h.paint();
    h.measure.clear();
    expect(h.measure.stats.size).toBe(0);
  });
});

describe('renderGrid - alignment', () => {
  it('right-aligns numbers by general alignment', () => {
    const h = setup();
    h.sheet.setValue(0, 0, 42);
    h.paint();
    const call = textCall(h.ctx, '42');
    expect(call?.args[1]).toBeCloseTo(HEADER_W + COL_W - PAD - 2 * CHAR);
  });

  it('left-aligns text by general alignment', () => {
    const h = setup();
    h.sheet.setValue(0, 0, 'hi');
    h.paint();
    expect(textCall(h.ctx, 'hi')?.args[1]).toBeCloseTo(HEADER_W + PAD);
  });

  it('centres booleans by general alignment', () => {
    const h = setup();
    h.sheet.setValue(0, 0, true);
    h.paint();
    const width = 4 * CHAR;
    expect(textCall(h.ctx, 'TRUE')?.args[1]).toBeCloseTo(HEADER_W + PAD + (COL_W - 2 * PAD - width) / 2);
  });

  it('centres errors', () => {
    expect(generalAlign(CellError.DIV0)).toBe('center');
    const h = setup();
    h.sheet.setValue(0, 0, CellError.DIV0);
    h.paint();
    expect(textCall(h.ctx, '#DIV/0!')).toBeDefined();
  });

  it('honours an explicit horizontal alignment over the general rule', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { horizontal: 'left' } });
    h.sheet.setCell(0, 0, { value: 42, style });
    h.paint();
    expect(textCall(h.ctx, '42')?.args[1]).toBeCloseTo(HEADER_W + PAD);
  });

  it('maps centerContinuous to centre and general to the value rule', () => {
    expect(effectiveHAlign('centerContinuous', 'x')).toBe('center');
    expect(effectiveHAlign('general', 1)).toBe('right');
    expect(effectiveHAlign(undefined, 'x')).toBe('left');
    expect(effectiveHAlign('justify', 1)).toBe('justify');
  });

  it('applies an indent from the left edge', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { indent: 2 } });
    h.sheet.setCell(0, 0, { value: 'x', style });
    h.paint();
    // One indent level is three spaces; a space measures 0.5 * 11 here.
    expect(textCall(h.ctx, 'x')?.args[1]).toBeCloseTo(HEADER_W + PAD + 2 * 3 * CHAR);
  });

  it('applies an indent from the right edge for right-aligned text', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { horizontal: 'right', indent: 1 } });
    h.sheet.setCell(0, 0, { value: 'x', style });
    h.paint();
    expect(textCall(h.ctx, 'x')?.args[1]).toBeCloseTo(HEADER_W + COL_W - PAD - 3 * CHAR - CHAR);
  });

  it('bottom-aligns by default', () => {
    const h = setup();
    h.sheet.setValue(0, 0, 'x');
    h.paint();
    const baseline = textCall(h.ctx, 'x')?.args[2] as number;
    expect(baseline).toBeCloseTo(HEADER_H + ROW_H - PAD - 11 * 0.2);
  });

  it('top-aligns when asked', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { vertical: 'top' } });
    h.sheet.setCell(0, 0, { value: 'x', style });
    h.paint();
    expect(textCall(h.ctx, 'x')?.args[2]).toBeCloseTo(HEADER_H + PAD + 11 * 0.8);
  });

  it('centres vertically when asked', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { vertical: 'center' } });
    h.sheet.setCell(0, 0, { value: 'x', style });
    h.paint();
    expect(textCall(h.ctx, 'x')?.args[2]).toBeCloseTo(HEADER_H + (ROW_H - 11) / 2 + 11 * 0.8);
  });
});

describe('renderGrid - overflow and clipping', () => {
  const long = 'this label is far too wide for one column';

  it('runs long text into empty neighbours', () => {
    const h = setup();
    h.sheet.setValue(0, 0, long);
    const stats = h.paint();
    expect(stats.overflowed).toBe(1);
    expect(textCall(h.ctx, long)).toBeDefined();
  });

  it('clips instead when the neighbour is occupied', () => {
    const h = setup();
    h.sheet.setValue(0, 0, long);
    h.sheet.setValue(0, 1, 'stop');
    const stats = h.paint();
    expect(stats.overflowed).toBe(0);
    expect(h.ctx.ops('clip').length).toBeGreaterThan(0);
  });

  it('stops overflowing at a merged neighbour', () => {
    const h = setup();
    h.sheet.setValue(0, 0, long);
    h.sheet.merges = [{ range: parseRangeRef('B1:C1')! }];
    expect(h.paint().overflowed).toBe(0);
  });

  it('overflows leftwards for right-aligned text', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { horizontal: 'right' } });
    h.sheet.setCell(0, 3, { value: long, style });
    h.paint();
    const x = textCall(h.ctx, long)?.args[1] as number;
    expect(x).toBeLessThan(HEADER_W + 3 * COL_W);
  });

  it('overflows both ways for centred text', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { horizontal: 'center' } });
    h.sheet.setCell(0, 3, { value: long, style });
    expect(h.paint().overflowed).toBe(1);
  });

  it('does not overflow out of a merged cell', () => {
    const h = setup();
    h.sheet.merges = [{ range: parseRangeRef('A1:A2')! }];
    h.sheet.setValue(0, 0, long);
    const stats = h.paint();
    expect(stats.overflowed).toBe(0);
    expect(stats.merges).toBe(1);
  });

  it('never overflows a short label', () => {
    const h = setup();
    h.sheet.setValue(0, 0, 'ok');
    expect(h.paint().overflowed).toBe(0);
  });
});

describe('renderGrid - numbers too wide for the column', () => {
  it('renders a row of # instead of a wide number', () => {
    const h = setup();
    h.sheet.setValue(0, 0, 123456789012345);
    const stats = h.paint();
    expect(stats.hashed).toBe(1);
    const drawn = h.ctx.texts().find((t) => t.startsWith('#'));
    expect(drawn).toBe('#'.repeat(Math.floor((COL_W - 2 * PAD) / CHAR)));
  });

  it('leaves a number that fits alone', () => {
    const h = setup();
    h.sheet.setValue(0, 0, 42);
    expect(h.paint().hashed).toBe(0);
  });

  it('does not hash wide text', () => {
    const h = setup();
    h.sheet.setValue(0, 0, 'a very long piece of text indeed');
    expect(h.paint().hashed).toBe(0);
  });

  it('hashes a wide formatted number', () => {
    const h = setup();
    const style = h.wb.styles.intern({ numFmt: '#,##0.0000' });
    h.sheet.setCell(0, 0, { value: 1234567.5, style });
    expect(h.paint().hashed).toBe(1);
  });

  it('widens the # run with the column', () => {
    const code = `0.${'0'.repeat(40)}`;
    const wide = setup();
    wide.sheet.cols.set(0, { width: 30 });
    wide.sheet.setCell(0, 0, { value: 1.5, style: wide.wb.styles.intern({ numFmt: code }) });
    wide.paint();
    const wideRun = wide.ctx.texts().find((t) => t.startsWith('#')) ?? '';

    const narrow = setup();
    narrow.sheet.setCell(0, 0, { value: 1.5, style: narrow.wb.styles.intern({ numFmt: code }) });
    narrow.paint();
    const narrowRun = narrow.ctx.texts().find((t) => t.startsWith('#')) ?? '';

    expect(narrowRun.length).toBe(Math.floor((COL_W - 2 * PAD) / CHAR));
    expect(wideRun.length).toBe(Math.floor((30 * 7 + 5 - 2 * PAD) / CHAR));
    expect(wideRun.length).toBeGreaterThan(narrowRun.length);
  });
});

describe('renderGrid - number formats', () => {
  it('formats through the number-format engine', () => {
    const h = setup();
    const style = h.wb.styles.intern({ numFmtId: 2 });
    h.sheet.setCell(0, 0, { value: 3.14159, style });
    h.paint();
    expect(h.ctx.texts()).toContain('3.14');
  });

  it('resolves a built-in format id', () => {
    expect(formatCodeOf({ numFmtId: 2 })).toBe('0.00');
    expect(formatCodeOf({ numFmt: '0%' })).toBe('0%');
    expect(formatCodeOf({})).toBe('General');
    expect(formatCodeOf({ numFmtId: 9999 })).toBe('General');
  });

  it('applies a colour named by the format', () => {
    const h = setup();
    const style = h.wb.styles.intern({ numFmt: '0.00;[Red]-0.00' });
    h.sheet.setCell(0, 0, { value: -1, style });
    h.paint();
    expect(textCall(h.ctx, '-1.00')?.fillStyle).toBe('#ff0000');
  });

  it('lifts format colours for the dark palette', () => {
    const h = setup({ theme: DARK_THEME });
    const style = h.wb.styles.intern({ numFmt: '0.00;[Red]-0.00' });
    h.sheet.setCell(0, 0, { value: -1, style });
    h.paint();
    expect(textCall(h.ctx, '-1.00')?.fillStyle).toBe('#ff6b5e');
  });

  it('skips a cell whose formatted text is empty', () => {
    const h = setup();
    const style = h.wb.styles.intern({ numFmt: ';;;' });
    h.sheet.setCell(0, 0, { value: 5, style });
    expect(h.paint().texts).toBe(0);
  });
});

describe('renderGrid - wrapping and rotation', () => {
  it('breaks wrapped text into lines', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { wrapText: true } });
    h.sheet.setCell(0, 0, { value: 'alpha beta gamma delta', style });
    h.paint();
    const drawn = h.ctx.texts().filter((t) => t.includes('alpha') || t.includes('beta'));
    expect(drawn.length).toBeGreaterThan(1);
    expect(drawn.join(' ')).not.toContain('alpha beta gamma delta');
  });

  it('keeps a hard line break', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { wrapText: true } });
    h.sheet.setCell(0, 0, { value: 'a\nb', style });
    h.paint();
    expect(h.ctx.texts()).toContain('a');
    expect(h.ctx.texts()).toContain('b');
  });

  it('clips wrapped text to its cell', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { wrapText: true } });
    h.sheet.setCell(0, 0, { value: 'alpha beta gamma delta epsilon', style });
    h.paint();
    expect(h.ctx.ops('clip').length).toBeGreaterThan(0);
  });

  it('rotates counter-clockwise for angles up to 90', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { textRotation: 45 } });
    h.sheet.setCell(0, 0, { value: 'x', style });
    h.paint();
    expect(h.ctx.ops('rotate')[0]?.args[0]).toBeCloseTo((-45 * Math.PI) / 180);
  });

  it('rotates clockwise for angles above 90', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { textRotation: 135 } });
    h.sheet.setCell(0, 0, { value: 'x', style });
    h.paint();
    expect(h.ctx.ops('rotate')[0]?.args[0]).toBeCloseTo((45 * Math.PI) / 180);
  });

  it('stacks characters for rotation 255 instead of rotating', () => {
    const h = setup();
    const style = h.wb.styles.intern({ alignment: { textRotation: 255 } });
    h.sheet.setCell(0, 0, { value: 'abc', style });
    h.paint();
    expect(h.ctx.ops('rotate')).toHaveLength(0);
    expect(h.ctx.texts()).toEqual(expect.arrayContaining(['a', 'b', 'c']));
  });
});

describe('renderGrid - merged cells', () => {
  it('draws a merge as one region', () => {
    const h = setup();
    h.sheet.merges = [{ range: parseRangeRef('A1:B2')! }];
    h.sheet.setValue(0, 0, 'merged');
    const stats = h.paint();
    expect(stats.merges).toBe(1);
    const body = h.ctx
      .ops('fillRect')
      .find((c) => c.args[0] === HEADER_W && c.args[1] === HEADER_H && c.args[2] === 2 * COL_W);
    expect(body?.args[3]).toBe(2 * ROW_H);
  });

  it('draws the anchor text once, not once per covered cell', () => {
    const h = setup();
    h.sheet.merges = [{ range: parseRangeRef('A1:C3')! }];
    h.sheet.setValue(0, 0, 'once');
    h.paint();
    expect(h.ctx.texts().filter((t) => t === 'once')).toHaveLength(1);
  });

  it('does not paint cells hidden under a merge', () => {
    const h = setup();
    h.sheet.merges = [{ range: parseRangeRef('A1:C3')! }];
    h.sheet.setValue(0, 0, 'anchor');
    h.sheet.setValue(1, 1, 'buried');
    h.paint();
    expect(h.ctx.texts()).not.toContain('buried');
  });

  it('draws a merge whose anchor is scrolled off screen', () => {
    const h = setup({ width: 400, height: 200 });
    h.sheet.merges = [{ range: parseRangeRef('A1:D20')! }];
    h.sheet.setValue(0, 0, 'tall');
    h.paint();
    h.ctx.reset();
    h.view.setScroll(0, 100);
    const stats = renderGrid(h.ctx, {
      sheet: h.sheet,
      styles: h.wb.styles,
      layout: h.view.layout(),
      geometry: h.view,
      theme: LIGHT_THEME,
      measure: h.measure,
    });
    expect(stats.merges).toBe(1);
    expect(h.ctx.texts()).toContain('tall');
  });
});

describe('renderGrid - fills, borders and gridlines', () => {
  it('paints a solid fill from its foreground colour', () => {
    const h = setup();
    const style = h.wb.styles.intern({
      fill: { pattern: 'solid', fg: { kind: 'rgb', argb: 'FFFFFF00' } },
    });
    h.sheet.setCell(0, 0, { value: null, style });
    h.paint();
    const rect = h.ctx.ops('fillRect').find((c) => c.fillStyle === '#FFFF00');
    expect(rect?.args).toEqual([HEADER_W, HEADER_H, COL_W, ROW_H]);
  });

  it('reads the foreground of a solid fill, not the background', () => {
    expect(
      fillColorOf({ fill: { pattern: 'solid', fg: { kind: 'rgb', argb: 'FF102030' } } }, undefined),
    ).toBe('#102030');
    expect(fillColorOf({ fill: { pattern: 'none' } }, undefined)).toBeUndefined();
    expect(fillColorOf({}, undefined)).toBeUndefined();
  });

  it('paints each requested border edge', () => {
    const h = setup();
    const style = h.wb.styles.intern({
      border: {
        top: { style: 'thin', color: { kind: 'rgb', argb: 'FF000000' } },
        bottom: { style: 'medium' },
      },
    });
    h.sheet.setCell(0, 0, { value: null, style });
    expect(h.paint().borderSegments).toBe(2);
  });

  it('gives each border style its own weight and dash', () => {
    expect(borderStroke('thin').width).toBe(1);
    expect(borderStroke('medium').width).toBe(2);
    expect(borderStroke('thick').width).toBe(3);
    expect(borderStroke('dashed').dash.length).toBeGreaterThan(0);
    expect(borderStroke('double').double).toBe(true);
  });

  it('draws two lines for a double border', () => {
    const h = setup();
    const style = h.wb.styles.intern({ border: { left: { style: 'double' } } });
    h.sheet.setCell(0, 0, { value: null, style });
    h.paint();
    expect(h.ctx.ops('stroke').length).toBeGreaterThan(1);
  });

  it('draws a gridline per visible row and column edge', () => {
    const h = setup({ width: 400, height: 200 });
    const stats = h.paint();
    const cols = Math.ceil((400 - HEADER_W) / COL_W);
    const rows = Math.ceil((200 - HEADER_H) / ROW_H);
    expect(stats.gridLines).toBe(cols + 1 + rows + 1);
  });

  it('omits gridlines when they are turned off', () => {
    const h = setup();
    expect(h.paint({ showGridLines: false }).gridLines).toBe(0);
  });

  it('snaps a hairline onto a device pixel centre', () => {
    expect(crisp(10, 1)).toBe(10.5);
    expect(crisp(10.2, 2)).toBe(10.25);
  });
});

describe('renderGrid - headers', () => {
  function selectionOf(ranges: SelectionState['ranges'], active = { row: 0, col: 0 }): SelectionState {
    return { active, anchor: active, ranges, activeRange: 0, endMode: false };
  }

  it('labels columns and rows', () => {
    const h = setup();
    h.paint();
    expect(h.ctx.texts()).toContain('A');
    expect(h.ctx.texts()).toContain('B');
    expect(h.ctx.texts()).toContain('1');
    expect(h.ctx.texts()).toContain('2');
  });

  it('counts one paint per header cell plus the corner', () => {
    const h = setup({ width: 400, height: 200 });
    const stats = h.paint();
    expect(stats.headerCells).toBeGreaterThan(4);
  });

  it('highlights the headers of the selection', () => {
    const h = setup();
    h.paint({ selection: selectionOf([{ top: 0, left: 1, bottom: 0, right: 1 }]) });
    const header = textCall(h.ctx, 'B');
    expect(header?.fillStyle).toBe(LIGHT_THEME.headerTextSelected);
  });

  it('uses the strong highlight for a whole selected column', () => {
    const h = setup();
    h.paint({
      selection: selectionOf([{ top: 0, left: 1, bottom: MAX_ROWS - 1, right: 1 }]),
    });
    const fills = h.ctx.ops('fillRect').filter((c) => c.fillStyle === LIGHT_THEME.headerBackgroundActive);
    expect(fills.length).toBeGreaterThan(0);
  });

  it('uses the strong highlight for a whole selected row', () => {
    const h = setup();
    h.paint({
      selection: selectionOf([{ top: 1, left: 0, bottom: 1, right: MAX_COLS - 1 }]),
    });
    const fills = h.ctx.ops('fillRect').filter((c) => c.fillStyle === LIGHT_THEME.headerBackgroundActive);
    expect(fills.length).toBeGreaterThan(0);
  });

  it('leaves unselected headers plain', () => {
    const h = setup();
    h.paint({ selection: selectionOf([{ top: 0, left: 0, bottom: 0, right: 0 }]) });
    expect(textCall(h.ctx, 'C')?.fillStyle).toBe(LIGHT_THEME.headerText);
  });
});

describe('renderGrid - selection overlay', () => {
  const selection: SelectionState = {
    active: { row: 1, col: 1 },
    anchor: { row: 1, col: 1 },
    ranges: [{ top: 1, left: 1, bottom: 3, right: 3 }],
    activeRange: 0,
    endMode: false,
  };

  it('shades the range but not the active cell', () => {
    const h = setup();
    h.paint({ selection });
    const shades = h.ctx.ops('fillRect').filter((c) => c.fillStyle === LIGHT_THEME.selectionFill);
    expect(shades.length).toBeGreaterThan(0);
    const activeX = HEADER_W + COL_W;
    const activeY = HEADER_H + ROW_H;
    expect(shades.some((c) => c.args[0] === activeX && c.args[1] === activeY)).toBe(false);
  });

  it('outlines the range and the active cell', () => {
    const h = setup();
    h.paint({ selection });
    const strokes = h.ctx.ops('strokeRect');
    expect(strokes.some((c) => c.strokeStyle === LIGHT_THEME.selectionBorder)).toBe(true);
    expect(strokes.some((c) => c.strokeStyle === LIGHT_THEME.activeCellBorder)).toBe(true);
  });

  it('shades a range that does not contain the active cell whole', () => {
    const h = setup();
    h.paint({
      selection: {
        ...selection,
        active: { row: 0, col: 0 },
        ranges: [
          { top: 0, left: 0, bottom: 0, right: 0 },
          { top: 2, left: 2, bottom: 3, right: 3 },
        ],
      },
    });
    const shades = h.ctx.ops('fillRect').filter((c) => c.fillStyle === LIGHT_THEME.selectionFill);
    expect(shades.some((c) => c.args[0] === HEADER_W + 2 * COL_W)).toBe(true);
  });

  it('draws a fill handle at the bottom-right of the primary range', () => {
    const h = setup();
    h.paint({ selection });
    const handle = h.ctx.ops('fillRect').find((c) => c.fillStyle === LIGHT_THEME.fillHandle);
    expect(handle?.args[2]).toBe(6);
  });
});

describe('renderGrid - themes', () => {
  it('paints every colour from the theme', () => {
    const h = setup({ theme: DARK_THEME });
    h.sheet.setValue(0, 0, 'x');
    h.paint();
    expect(h.ctx.ops('fillRect')[0]?.fillStyle).toBe(DARK_THEME.background);
    expect(textCall(h.ctx, 'x')?.fillStyle).toBe(DARK_THEME.cellText);
  });

  it('uses the dark gridline, not an inverted light one', () => {
    const h = setup({ theme: DARK_THEME });
    h.paint();
    expect(h.ctx.ops('stroke').some((c) => c.strokeStyle === DARK_THEME.gridLine)).toBe(true);
    expect(DARK_THEME.gridLine).not.toBe(LIGHT_THEME.gridLine);
  });

  it('colours errors distinctly on both palettes', () => {
    for (const theme of [LIGHT_THEME, DARK_THEME]) {
      const h = setup({ theme });
      h.sheet.setValue(0, 0, CellError.REF);
      h.paint();
      expect(textCall(h.ctx, '#REF!')?.fillStyle).toBe(theme.errorText);
    }
  });

  it('leaves an authored cell colour alone in dark mode', () => {
    const h = setup({ theme: DARK_THEME });
    const style = h.wb.styles.intern({
      fill: { pattern: 'solid', fg: { kind: 'rgb', argb: 'FFFFFF00' } },
    });
    h.sheet.setCell(0, 0, { value: null, style });
    h.paint();
    expect(h.ctx.ops('fillRect').some((c) => c.fillStyle === '#FFFF00')).toBe(true);
  });

  it('defines the same keys in both palettes', () => {
    expect(Object.keys(LIGHT_THEME).sort()).toEqual(Object.keys(DARK_THEME).sort());
  });
});

describe('renderGrid - frozen panes', () => {
  it('draws the divider between the frozen and scrolling regions', () => {
    const h = setup({ frozenRows: 1, frozenCols: 1 });
    h.paint();
    expect(h.ctx.ops('stroke').some((c) => c.strokeStyle === LIGHT_THEME.frozenLine)).toBe(true);
  });

  it('clips each pane to its own rectangle', () => {
    const h = setup({ frozenRows: 1, frozenCols: 1 });
    h.paint();
    const clipRects = h.ctx.ops('rect');
    expect(clipRects.some((c) => c.args[0] === HEADER_W && c.args[2] === COL_W)).toBe(true);
  });

  it('paints frozen cells as well as scrolling ones', () => {
    const h = setup({ frozenRows: 1, frozenCols: 1 });
    h.sheet.setValue(0, 0, 'frozen');
    h.sheet.setValue(5, 5, 'scrolled');
    h.paint();
    expect(h.ctx.texts()).toContain('frozen');
    expect(h.ctx.texts()).toContain('scrolled');
  });

  it('draws no divider without frozen panes', () => {
    const h = setup();
    h.paint();
    expect(h.ctx.ops('stroke').some((c) => c.strokeStyle === LIGHT_THEME.frozenLine)).toBe(false);
  });
});
