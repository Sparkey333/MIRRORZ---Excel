import { describe, expect, it } from 'vitest';
import {
  MAX_COLS,
  MAX_ROWS,
  a1,
  colToName,
  formatCellRef,
  formatRangeRef,
  intersectRanges,
  nameToCol,
  offsetCellRef,
  packKey,
  parseCellRef,
  parseRangeRef,
  quoteSheetName,
  rangeContains,
  rangeHeight,
  rangeWidth,
  rangesIntersect,
  unionRanges,
  unpackCol,
  unpackRow,
  unquoteSheetName,
} from '../src/address.js';

describe('column names', () => {
  it.each([
    [0, 'A'],
    [1, 'B'],
    [25, 'Z'],
    [26, 'AA'],
    [27, 'AB'],
    [51, 'AZ'],
    [52, 'BA'],
    [701, 'ZZ'],
    [702, 'AAA'],
    [16383, 'XFD'],
  ])('index %i <-> %s', (idx, name) => {
    expect(colToName(idx)).toBe(name);
    expect(nameToCol(name)).toBe(idx);
  });

  it('round-trips every valid column', () => {
    for (let c = 0; c < MAX_COLS; c++) {
      expect(nameToCol(colToName(c))).toBe(c);
    }
  });

  it('is case-insensitive', () => {
    expect(nameToCol('xfd')).toBe(16383);
    expect(nameToCol('aA')).toBe(26);
  });

  it('rejects out-of-range and malformed names', () => {
    expect(nameToCol('XFE')).toBe(-1); // one past the last column
    expect(nameToCol('AAAA')).toBe(-1);
    expect(nameToCol('')).toBe(-1);
    expect(nameToCol('A1')).toBe(-1);
    expect(() => colToName(MAX_COLS)).toThrow(RangeError);
    expect(() => colToName(-1)).toThrow(RangeError);
  });
});

describe('cell references', () => {
  it('parses relative and absolute forms', () => {
    expect(parseCellRef('B7')).toEqual({ row: 6, col: 1, rowAbs: false, colAbs: false });
    expect(parseCellRef('$B$7')).toEqual({ row: 6, col: 1, rowAbs: true, colAbs: true });
    expect(parseCellRef('B$7')).toEqual({ row: 6, col: 1, rowAbs: true, colAbs: false });
    expect(parseCellRef('$B7')).toEqual({ row: 6, col: 1, rowAbs: false, colAbs: true });
  });

  it('handles the sheet corners', () => {
    expect(parseCellRef('A1')).toMatchObject({ row: 0, col: 0 });
    expect(parseCellRef('XFD1048576')).toMatchObject({ row: MAX_ROWS - 1, col: MAX_COLS - 1 });
  });

  it('rejects addresses past the sheet limits', () => {
    expect(parseCellRef('XFE1')).toBeUndefined();
    expect(parseCellRef('A1048577')).toBeUndefined();
    expect(parseCellRef('A0')).toBeUndefined();
    expect(parseCellRef('1A')).toBeUndefined();
    expect(parseCellRef('')).toBeUndefined();
  });

  it('formats back to the original text', () => {
    for (const s of ['A1', '$A$1', 'A$1', '$A1', 'XFD1048576', 'BC42']) {
      expect(formatCellRef(parseCellRef(s)!)).toBe(s);
    }
  });

  it('a1() is the plain relative shorthand', () => {
    expect(a1(0, 0)).toBe('A1');
    expect(a1(41, 27)).toBe('AB42');
  });
});

describe('range references', () => {
  it('parses a plain range', () => {
    const r = parseRangeRef('B2:D5')!;
    expect(r.start).toMatchObject({ row: 1, col: 1 });
    expect(r.end).toMatchObject({ row: 4, col: 3 });
    expect(rangeWidth(r)).toBe(3);
    expect(rangeHeight(r)).toBe(4);
  });

  it('treats a single cell as a 1x1 range', () => {
    const r = parseRangeRef('C3')!;
    expect(rangeWidth(r)).toBe(1);
    expect(rangeHeight(r)).toBe(1);
    expect(formatRangeRef(r)).toBe('C3');
  });

  it('normalises a reversed range the way Excel does', () => {
    const r = parseRangeRef('D5:B2')!;
    expect(r.start).toMatchObject({ row: 1, col: 1 });
    expect(r.end).toMatchObject({ row: 4, col: 3 });
  });

  it('parses whole columns', () => {
    const r = parseRangeRef('B:D')!;
    expect(r.wholeCol).toBe(true);
    expect(r.start).toMatchObject({ row: 0, col: 1 });
    expect(r.end).toMatchObject({ row: MAX_ROWS - 1, col: 3 });
    expect(formatRangeRef(r)).toBe('B:D');
  });

  it('parses whole rows', () => {
    const r = parseRangeRef('2:5')!;
    expect(r.wholeRow).toBe(true);
    expect(r.start).toMatchObject({ row: 1, col: 0 });
    expect(r.end).toMatchObject({ row: 4, col: MAX_COLS - 1 });
    expect(formatRangeRef(r)).toBe('2:5');
  });

  it('preserves absolute markers on both ends', () => {
    expect(formatRangeRef(parseRangeRef('$A$1:$B$2')!)).toBe('$A$1:$B$2');
    expect(formatRangeRef(parseRangeRef('A$1:$B2')!)).toBe('A$1:$B2');
  });

  it('rejects nonsense', () => {
    expect(parseRangeRef('A1:')).toBeUndefined();
    expect(parseRangeRef('A:2')).toBeUndefined();
    expect(parseRangeRef('::')).toBeUndefined();
  });
});

describe('range algebra', () => {
  const b2d5 = parseRangeRef('B2:D5')!;

  it('tests containment', () => {
    expect(rangeContains(b2d5, 1, 1)).toBe(true);
    expect(rangeContains(b2d5, 4, 3)).toBe(true);
    expect(rangeContains(b2d5, 0, 1)).toBe(false);
    expect(rangeContains(b2d5, 1, 4)).toBe(false);
  });

  it('detects intersection', () => {
    expect(rangesIntersect(b2d5, parseRangeRef('C3:E7')!)).toBe(true);
    expect(rangesIntersect(b2d5, parseRangeRef('F1:G2')!)).toBe(false);
  });

  it('computes the intersection rectangle', () => {
    // This is Excel's space operator: SUM(B2:D5 C3:E7).
    expect(formatRangeRef(intersectRanges(b2d5, parseRangeRef('C3:E7')!)!)).toBe('C3:D5');
    expect(intersectRanges(b2d5, parseRangeRef('F1:G2')!)).toBeUndefined();
  });

  it('computes the bounding union', () => {
    expect(formatRangeRef(unionRanges(b2d5, parseRangeRef('F7:G9')!))).toBe('B2:G9');
  });
});

describe('reference offsetting', () => {
  it('moves relative parts and pins absolute ones', () => {
    const mixed = parseCellRef('$B7')!;
    expect(formatCellRef(offsetCellRef(mixed, 2, 3)!)).toBe('$B9');
    const rel = parseCellRef('B7')!;
    expect(formatCellRef(offsetCellRef(rel, 2, 3)!)).toBe('E9');
  });

  it('returns undefined when pushed off the sheet, which becomes #REF!', () => {
    expect(offsetCellRef(parseCellRef('A1')!, -1, 0)).toBeUndefined();
    expect(offsetCellRef(parseCellRef('A1')!, 0, -1)).toBeUndefined();
    expect(offsetCellRef(parseCellRef('XFD1')!, 0, 1)).toBeUndefined();
  });
});

describe('packed keys', () => {
  it('round-trips every corner', () => {
    for (const [r, c] of [
      [0, 0],
      [0, MAX_COLS - 1],
      [MAX_ROWS - 1, 0],
      [MAX_ROWS - 1, MAX_COLS - 1],
      [524_288, 8192],
    ] as const) {
      const k = packKey(r, c);
      expect(Number.isSafeInteger(k)).toBe(true);
      expect(unpackRow(k)).toBe(r);
      expect(unpackCol(k)).toBe(c);
    }
  });

  it('is collision-free', () => {
    const seen = new Set<number>();
    for (let r = 0; r < 200; r++) {
      for (let c = 0; c < 200; c++) seen.add(packKey(r, c));
    }
    expect(seen.size).toBe(40_000);
  });
});

describe('sheet name quoting', () => {
  it.each([
    ['Sheet1', 'Sheet1'],
    ['Data', 'Data'],
    ['My Sheet', "'My Sheet'"],
    ["Bob's", "'Bob''s'"],
    ['A1', "'A1'"], // looks like a cell address, so must be quoted
    ['R1C1', "'R1C1'"],
    ['2024', "'2024'"],
    ['a-b', "'a-b'"],
  ])('quotes %s as %s', (raw, quoted) => {
    expect(quoteSheetName(raw)).toBe(quoted);
    expect(unquoteSheetName(quoted)).toBe(raw);
  });
});
