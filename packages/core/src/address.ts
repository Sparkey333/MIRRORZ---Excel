/**
 * A1-notation addressing.
 *
 * Everything here is 0-based internally (row 0 == spreadsheet row 1) because
 * that is what array indexing wants, and converts to 1-based only at the
 * A1-string boundary. Mixing the two is the single most common source of
 * off-by-one bugs in spreadsheet code, so the boundary is deliberately narrow.
 */

/** Excel's hard limits, unchanged since the 2007 format. */
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384; // A .. XFD

const A = 'A'.charCodeAt(0);

/** 0 -> "A", 25 -> "Z", 26 -> "AA", 16383 -> "XFD". */
export function colToName(col: number): string {
  if (!Number.isInteger(col) || col < 0 || col >= MAX_COLS) {
    throw new RangeError(`column index out of range: ${col}`);
  }
  let n = col;
  let out = '';
  // Bijective base-26: there is no "zero digit", so we decrement before each step.
  do {
    out = String.fromCharCode(A + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/** "A" -> 0, "XFD" -> 16383. Case-insensitive. Returns -1 if not a column name. */
export function nameToCol(name: string): number {
  if (name.length === 0 || name.length > 3) return -1;
  let n = 0;
  for (let i = 0; i < name.length; i++) {
    const c = name.charCodeAt(i) & ~0x20; // upper-case ASCII letters
    if (c < A || c > A + 25) return -1;
    n = n * 26 + (c - A + 1);
  }
  const col = n - 1;
  return col < MAX_COLS ? col : -1;
}

export interface CellRef {
  readonly row: number;
  readonly col: number;
  readonly rowAbs: boolean;
  readonly colAbs: boolean;
}

export interface RangeRef {
  readonly start: CellRef;
  readonly end: CellRef;
  /** True when written as a whole-column ref such as `A:A`. */
  readonly wholeCol?: boolean;
  /** True when written as a whole-row ref such as `1:1`. */
  readonly wholeRow?: boolean;
}

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;

/** Parse a single A1 cell reference such as `B7`, `$B$7`, `B$7`. */
export function parseCellRef(s: string): CellRef | undefined {
  const m = CELL_RE.exec(s);
  if (!m) return undefined;
  const col = nameToCol(m[2]!);
  if (col < 0) return undefined;
  const row = Number(m[4]) - 1;
  if (row < 0 || row >= MAX_ROWS) return undefined;
  return { row, col, colAbs: m[1] === '$', rowAbs: m[3] === '$' };
}

/** Format a cell reference back to A1, honouring the absolute flags. */
export function formatCellRef(ref: CellRef): string {
  return `${ref.colAbs ? '$' : ''}${colToName(ref.col)}${ref.rowAbs ? '$' : ''}${ref.row + 1}`;
}

/** Shorthand for a plain relative address, e.g. `a1(0, 0) === "A1"`. */
export function a1(row: number, col: number): string {
  return `${colToName(col)}${row + 1}`;
}

const COL_ONLY_RE = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_ONLY_RE = /^(\$?)([0-9]{1,7})$/;

/**
 * Parse an A1 range: `A1:B2`, `$A$1:$B$2`, a single cell `A1`, a whole column
 * `A:C`, or a whole row `2:5`. The returned range is always normalised so that
 * `start` is the top-left corner.
 */
export function parseRangeRef(s: string): RangeRef | undefined {
  const colon = s.indexOf(':');
  if (colon < 0) {
    const c = parseCellRef(s);
    return c ? { start: c, end: c } : undefined;
  }
  const left = s.slice(0, colon);
  const right = s.slice(colon + 1);

  const lc = parseCellRef(left);
  const rc = parseCellRef(right);
  if (lc && rc) return normalise({ start: lc, end: rc });

  // Whole columns: A:C
  const lcol = COL_ONLY_RE.exec(left);
  const rcol = COL_ONLY_RE.exec(right);
  if (lcol && rcol) {
    const c1 = nameToCol(lcol[2]!);
    const c2 = nameToCol(rcol[2]!);
    if (c1 < 0 || c2 < 0) return undefined;
    return normalise({
      start: { row: 0, col: c1, rowAbs: true, colAbs: lcol[1] === '$' },
      end: { row: MAX_ROWS - 1, col: c2, rowAbs: true, colAbs: rcol[1] === '$' },
      wholeCol: true,
    });
  }

  // Whole rows: 2:5
  const lrow = ROW_ONLY_RE.exec(left);
  const rrow = ROW_ONLY_RE.exec(right);
  if (lrow && rrow) {
    const r1 = Number(lrow[2]) - 1;
    const r2 = Number(rrow[2]) - 1;
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS) return undefined;
    return normalise({
      start: { row: r1, col: 0, rowAbs: lrow[1] === '$', colAbs: true },
      end: { row: r2, col: MAX_COLS - 1, rowAbs: rrow[1] === '$', colAbs: true },
      wholeRow: true,
    });
  }

  return undefined;
}

function normalise(r: RangeRef): RangeRef {
  const { start, end } = r;
  if (start.row <= end.row && start.col <= end.col) return r;
  // Swap coordinates independently, carrying each one's absolute flag with it,
  // which is what Excel does when you type B2:A1.
  const top = Math.min(start.row, end.row);
  const bottom = Math.max(start.row, end.row);
  const leftCol = Math.min(start.col, end.col);
  const rightCol = Math.max(start.col, end.col);
  const rowFlags = start.row <= end.row ? [start.rowAbs, end.rowAbs] : [end.rowAbs, start.rowAbs];
  const colFlags = start.col <= end.col ? [start.colAbs, end.colAbs] : [end.colAbs, start.colAbs];
  return {
    ...r,
    start: { row: top, col: leftCol, rowAbs: rowFlags[0]!, colAbs: colFlags[0]! },
    end: { row: bottom, col: rightCol, rowAbs: rowFlags[1]!, colAbs: colFlags[1]! },
  };
}

export function formatRangeRef(r: RangeRef): string {
  if (r.wholeCol) {
    return `${r.start.colAbs ? '$' : ''}${colToName(r.start.col)}:${r.end.colAbs ? '$' : ''}${colToName(r.end.col)}`;
  }
  if (r.wholeRow) {
    return `${r.start.rowAbs ? '$' : ''}${r.start.row + 1}:${r.end.rowAbs ? '$' : ''}${r.end.row + 1}`;
  }
  const a = formatCellRef(r.start);
  const b = formatCellRef(r.end);
  return a === b ? a : `${a}:${b}`;
}

export function rangeWidth(r: RangeRef): number {
  return r.end.col - r.start.col + 1;
}

export function rangeHeight(r: RangeRef): number {
  return r.end.row - r.start.row + 1;
}

export function rangeContains(r: RangeRef, row: number, col: number): boolean {
  return row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col;
}

export function rangesIntersect(a: RangeRef, b: RangeRef): boolean {
  return (
    a.start.row <= b.end.row &&
    b.start.row <= a.end.row &&
    a.start.col <= b.end.col &&
    b.start.col <= a.end.col
  );
}

/** The overlap of two ranges, or undefined when they are disjoint. */
export function intersectRanges(a: RangeRef, b: RangeRef): RangeRef | undefined {
  if (!rangesIntersect(a, b)) return undefined;
  return {
    start: {
      row: Math.max(a.start.row, b.start.row),
      col: Math.max(a.start.col, b.start.col),
      rowAbs: false,
      colAbs: false,
    },
    end: {
      row: Math.min(a.end.row, b.end.row),
      col: Math.min(a.end.col, b.end.col),
      rowAbs: false,
      colAbs: false,
    },
  };
}

/** Smallest range containing both inputs. */
export function unionRanges(a: RangeRef, b: RangeRef): RangeRef {
  return {
    start: {
      row: Math.min(a.start.row, b.start.row),
      col: Math.min(a.start.col, b.start.col),
      rowAbs: false,
      colAbs: false,
    },
    end: {
      row: Math.max(a.end.row, b.end.row),
      col: Math.max(a.end.col, b.end.col),
      rowAbs: false,
      colAbs: false,
    },
  };
}

/**
 * Pack a (row, col) pair into a single integer key.
 *
 * 16384 columns needs 14 bits; 1048576 rows needs 20 bits. 34 bits total fits
 * comfortably inside a JS double's 53-bit integer range, so a plain `Map<number, T>`
 * beats string keys for both memory and lookup speed in the cell store.
 */
export function packKey(row: number, col: number): number {
  return row * MAX_COLS + col;
}

export function unpackRow(key: number): number {
  return Math.floor(key / MAX_COLS);
}

export function unpackCol(key: number): number {
  return key % MAX_COLS;
}

/**
 * Offset a reference, leaving absolute components fixed.
 * Returns undefined when the result falls off the sheet, which is how a
 * relative reference becomes #REF! after a row or column delete.
 */
export function offsetCellRef(ref: CellRef, dRow: number, dCol: number): CellRef | undefined {
  const row = ref.rowAbs ? ref.row : ref.row + dRow;
  const col = ref.colAbs ? ref.col : ref.col + dCol;
  if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) return undefined;
  return { row, col, rowAbs: ref.rowAbs, colAbs: ref.colAbs };
}

/**
 * Sheet names may contain almost anything, and must be single-quoted in
 * references when they do. Inside quotes, a literal apostrophe is doubled.
 */
const BARE_SHEET_RE = /^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/;

export function quoteSheetName(name: string): string {
  // A bare name that happens to look like a cell address (e.g. a sheet called
  // "A1" or "R1C1") must still be quoted, or the parser cannot tell them apart.
  if (BARE_SHEET_RE.test(name) && !parseCellRef(name) && !/^[Rr][0-9]*[Cc][0-9]*$/.test(name)) {
    return name;
  }
  return `'${name.replace(/'/g, "''")}'`;
}

export function unquoteSheetName(token: string): string {
  if (token.startsWith("'") && token.endsWith("'") && token.length >= 2) {
    return token.slice(1, -1).replace(/''/g, "'");
  }
  return token;
}
