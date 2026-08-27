/**
 * Lookup and reference.
 *
 * Five decisions shape this file.
 *
 * First, references stay references. INDEX, OFFSET and INDIRECT take
 * ArgKind.Reference and hand a RefValue back where Excel does, because
 * `A1:INDEX(B1:B9,3)` and `SUM(OFFSET(A1,1,1,3,3))` are ordinary formulas: the
 * result is used as a range operand, not as a value, and dereferencing at the
 * function boundary would turn both into #VALUE!. CHOOSE arrives at the same
 * place by another route - its values come in as thunks, because Microsoft
 * documents that an unselected value is not evaluated, and the chosen one is
 * returned exactly as it evaluated. ROW, COLUMN, ROWS, COLUMNS and AREAS want
 * the reference for the opposite reason: they report on its shape and never
 * look inside it.
 *
 * Second, the volatility declarations follow Microsoft, not folklore. OFFSET
 * and INDIRECT are volatile: their dependencies exist only after the arguments
 * have been evaluated, so the dependency graph cannot know what they read until
 * it has read it. INDEX, ROWS, COLUMNS and AREAS are widely repeated as
 * volatile and are not; marking them so would drag their whole dependent
 * closure into every recalculation for nothing.
 *
 * Third, approximate match is the default. VLOOKUP, HLOOKUP and MATCH with the
 * last argument omitted assume an ascending column and return the largest entry
 * at or below the target; exact match is opt-in. Getting this backwards returns
 * plausible, wrong data rather than an error, which is the worst failure mode a
 * lookup can have. LOOKUP has no exact mode at all.
 *
 * Fourth, ordering is Excel's total order, from compareScalars: number < text <
 * FALSE < TRUE, text compared case-insensitively, numbers compared at fifteen
 * significant digits. Approximate scans therefore behave on mixed-type data the
 * way Excel's binary search does rather than the way a type-filtered scan would.
 * Blank and error cells are skipped by the linear scans: a blank never equals a
 * typed lookup value, which is why MATCH(0, <blank cell>, 0) is #N/A. Where
 * Excel documents a sorted array (MATCH 1/-1, XLOOKUP search modes 2 and -2)
 * duplicates resolve as Excel's do - the last of a run of equals for the
 * ordered scan, the binary-search hit for the binary modes.
 *
 * Fifth, the two places the engine cannot yet reach are reported honestly
 * rather than guessed at:
 *
 *   AREAS counts the union operators in its argument's syntax tree, because the
 *   evaluator collapses `(A1,C3)` into one bounding rectangle and a RefValue
 *   cannot represent two areas. Counting the commas gives Excel's answer for
 *   every literal union; INDEX's area_num, which needs the areas themselves and
 *   not just their number, returns #REF! for anything past the first.
 *
 *   FORMULATEXT returns #N/A. FunctionContext exposes cell values but no cell
 *   formulas, so the text simply is not reachable from here; #N/A is at least
 *   the answer Excel gives for the majority case of a cell holding no formula.
 *   This one is a stub and should be revisited when the context grows a
 *   getFormula.
 *
 * INDIRECT parses its text with the real formula parser rather than a private
 * regular expression, so `Data!$D$2`, `A:A` and `'My Sheet'!A1:B2` all resolve
 * exactly as they would if typed. R1C1 text is rewritten to A1 first, which
 * keeps one parser rather than two.
 */

import {
  CellError,
  MAX_COLS,
  MAX_ROWS,
  type Scalar,
  colToName,
  isError,
  quoteSheetName,
} from '@mirrorz/core';
import type { Ast } from '../ast.js';
import { Node } from '../ast.js';
import { parseFormula } from '../parser.js';
import {
  ArgKind,
  type FunctionContext,
  type FunctionSpec,
  type Thunk,
  p,
} from '../registry.js';
import {
  type ArrayValue,
  type RefValue,
  type Value,
  arrayAt,
  compareScalars,
  isArray,
  isRef,
  makeArray,
  makeRef,
  toBoolean,
  toNumber,
  toText,
  wildcardToRegExp,
} from '../value.js';

// ---------------------------------------------------------------------------
// Argument plumbing
// ---------------------------------------------------------------------------

/** A read-only rectangular view over whatever an argument turned out to be. */
interface Grid {
  readonly rows: number;
  readonly cols: number;
  at(row: number, col: number): Scalar;
}

function arrayGrid(a: ArrayValue): Grid {
  return { rows: a.rows, cols: a.cols, at: (r, c) => arrayAt(a, r, c) };
}

/** Materialise an argument for searching. References are read through ctx. */
function gridOf(v: Value | undefined, ctx: FunctionContext): Grid | CellError {
  if (v === undefined) return { rows: 1, cols: 1, at: () => null };
  if (isError(v)) return v;
  if (isRef(v)) return arrayGrid(ctx.deref(v));
  if (isArray(v)) return arrayGrid(v);
  return { rows: 1, cols: 1, at: () => v };
}

/** The single value behind an argument, whatever shape it arrived in. */
function scalarOf(v: Value | undefined, ctx: FunctionContext): Scalar {
  if (v === undefined) return null;
  if (isRef(v)) return ctx.getScalar(v.sheet, v.startRow, v.startCol);
  if (isArray(v)) return v.data[0] ?? null;
  return v;
}

/** A count or index argument: numeric, then truncated towards zero. */
function intOf(v: Value | undefined, ctx: FunctionContext, whenOmitted: number): number | CellError {
  if (v === undefined) return whenOmitted;
  const n = toNumber(scalarOf(v, ctx));
  if (isError(n)) return n;
  return Math.trunc(n);
}

/**
 * A lookup value. A blank arrives as 0, matching Excel: `VLOOKUP(A1,...)` with
 * A1 empty finds the zeros, not the blanks.
 */
function lookupValueOf(v: Value | undefined, ctx: FunctionContext): Scalar {
  const s = scalarOf(v, ctx);
  return s === null ? 0 : s;
}

/** Omitted means TRUE for VLOOKUP/HLOOKUP; an explicitly blank argument is FALSE. */
function approximateFlag(v: Value | undefined, ctx: FunctionContext): boolean | CellError {
  if (v === undefined) return true;
  return toBoolean(scalarOf(v, ctx));
}

function isThunk(v: unknown): v is Thunk {
  return typeof v === 'object' && v !== null && 'ast' in v && 'evaluate' in v;
}

// ---------------------------------------------------------------------------
// Vectors and matching
// ---------------------------------------------------------------------------

interface Vector {
  readonly length: number;
  at(i: number): Scalar;
}

function rowVector(g: Grid, row: number): Vector {
  return { length: g.cols, at: (i) => g.at(row, i) };
}

function colVector(g: Grid, col: number): Vector {
  return { length: g.rows, at: (i) => g.at(i, col) };
}

/** A grid used as a list, or undefined when it is neither one row nor one column. */
function vectorOf(g: Grid): Vector | undefined {
  if (g.cols === 1) return colVector(g, 0);
  if (g.rows === 1) return rowVector(g, 0);
  return undefined;
}

/**
 * Equality as a lookup sees it: no coercion across types, so the text "3" never
 * matches the number 3, and a blank cell matches only another blank.
 */
function equalValues(a: Scalar, b: Scalar): boolean {
  if (a === null || b === null) return a === null && b === null;
  const cmp = compareScalars(a, b);
  return !isError(cmp) && cmp === 0;
}

/** Ordering, or undefined for a cell a scan must skip. */
function orderValues(a: Scalar, b: Scalar): number | undefined {
  if (a === null) return undefined;
  const cmp = compareScalars(a, b);
  return isError(cmp) ? undefined : cmp;
}

function wildcardMatcher(target: Scalar): RegExp | undefined {
  return typeof target === 'string' && /[*?]/.test(target) ? wildcardToRegExp(target) : undefined;
}

function findExact(vec: Vector, target: Scalar, wildcards: boolean, reverse: boolean): number {
  const pattern = wildcards ? wildcardMatcher(target) : undefined;
  const step = reverse ? -1 : 1;
  for (let i = reverse ? vec.length - 1 : 0; i >= 0 && i < vec.length; i += step) {
    const cell = vec.at(i);
    if (pattern !== undefined) {
      // A wildcard pattern is a text pattern; it never matches a number.
      if (typeof cell === 'string' && pattern.test(cell)) return i;
      continue;
    }
    if (equalValues(cell, target)) return i;
  }
  return -1;
}

/**
 * The sorted-array scan behind MATCH 1/-1 and approximate VLOOKUP: the last
 * entry at or below (or, descending, at or above) the target. Taking the last
 * rather than the first is what makes a run of duplicates resolve the way
 * Excel's binary search resolves it.
 */
function findOrdered(vec: Vector, target: Scalar, ascending: boolean): number {
  let best = -1;
  for (let i = 0; i < vec.length; i++) {
    const cmp = orderValues(vec.at(i), target);
    if (cmp === undefined) continue;
    if (ascending ? cmp <= 0 : cmp >= 0) best = i;
  }
  return best;
}

/**
 * XLOOKUP's next-smaller / next-larger, which unlike MATCH does not assume the
 * array is sorted: the closest value on the requested side wins, and among
 * equals the one reached first in the search direction.
 */
function findClosest(vec: Vector, target: Scalar, smaller: boolean, reverse: boolean): number {
  let best = -1;
  let bestValue: Scalar = null;
  const step = reverse ? -1 : 1;
  for (let i = reverse ? vec.length - 1 : 0; i >= 0 && i < vec.length; i += step) {
    const cell = vec.at(i);
    const cmp = orderValues(cell, target);
    if (cmp === undefined) continue;
    if (smaller ? cmp > 0 : cmp < 0) continue;
    if (best < 0) {
      best = i;
      bestValue = cell;
      continue;
    }
    const better = orderValues(cell, bestValue);
    if (better !== undefined && (smaller ? better > 0 : better < 0)) {
      best = i;
      bestValue = cell;
    }
  }
  return best;
}

/** Binary search for XLOOKUP/XMATCH search modes 2 and -2. */
function findBinary(
  vec: Vector,
  target: Scalar,
  ascending: boolean,
  approx: -1 | 0 | 1,
): number {
  let lo = 0;
  let hi = vec.length - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const cmp = compareScalars(vec.at(mid), target);
    if (isError(cmp)) return -1;
    if (cmp === 0) return mid;
    // Navigation follows the array's sort direction; the fallback candidate is
    // chosen by the value's own relation to the target, which does not flip.
    if (cmp < 0 && approx === -1) best = mid;
    if (cmp > 0 && approx === 1) best = mid;
    if (ascending === (cmp < 0)) lo = mid + 1;
    else hi = mid - 1;
  }
  return approx === 0 ? -1 : best;
}

// ---------------------------------------------------------------------------
// VLOOKUP / HLOOKUP / LOOKUP
// ---------------------------------------------------------------------------

function tableLookup(args: Value[], ctx: FunctionContext, vertical: boolean): Value {
  const target = lookupValueOf(args[0], ctx);
  const table = gridOf(args[1], ctx);
  if (isError(table)) return table;
  const index = intOf(args[2], ctx, 1);
  if (isError(index)) return index;
  const approximate = approximateFlag(args[3], ctx);
  if (isError(approximate)) return approximate;

  const extent = vertical ? table.cols : table.rows;
  if (index < 1) return CellError.VALUE;
  if (index > extent) return CellError.REF;

  const search = vertical ? colVector(table, 0) : rowVector(table, 0);
  // Wildcards are honoured in exact mode only; an approximate search compares
  // ordering, where `*` is just a character.
  const hit = approximate
    ? findOrdered(search, target, true)
    : findExact(search, target, true, false);
  if (hit < 0) return CellError.NA;
  return vertical ? table.at(hit, index - 1) : table.at(index - 1, hit);
}

const VLOOKUP: FunctionSpec = {
  name: 'VLOOKUP',
  params: [
    p.scalar('lookup_value'),
    p.array('table_array'),
    p.scalar('col_index_num'),
    p.scalar('range_lookup', true),
  ],
  summary: 'Looks in the first column of an array and returns a value from a row.',
  impl: (args, ctx) => tableLookup(args, ctx, true),
};

const HLOOKUP: FunctionSpec = {
  name: 'HLOOKUP',
  params: [
    p.scalar('lookup_value'),
    p.array('table_array'),
    p.scalar('row_index_num'),
    p.scalar('range_lookup', true),
  ],
  summary: 'Looks in the top row of an array and returns a value from a column.',
  impl: (args, ctx) => tableLookup(args, ctx, false),
};

const LOOKUP: FunctionSpec = {
  name: 'LOOKUP',
  params: [p.scalar('lookup_value'), p.array('lookup_vector'), p.array('result_vector', true)],
  summary: 'Looks up a value in a one-row or one-column range and returns a value.',
  impl: (args, ctx) => {
    const target = lookupValueOf(args[0], ctx);
    const source = gridOf(args[1], ctx);
    if (isError(source)) return source;

    if (args[2] === undefined) {
      // Array form: the longer axis decides which way the array is read, and
      // the answer comes from the last row or column of the same array.
      const horizontal = source.cols > source.rows;
      const search = horizontal ? rowVector(source, 0) : colVector(source, 0);
      const hit = findOrdered(search, target, true);
      if (hit < 0) return CellError.NA;
      return horizontal ? source.at(source.rows - 1, hit) : source.at(hit, source.cols - 1);
    }

    const search = vectorOf(source);
    if (!search) return CellError.NA;
    const result = gridOf(args[2], ctx);
    if (isError(result)) return result;
    const out = vectorOf(result);
    if (!out) return CellError.NA;
    const hit = findOrdered(search, target, true);
    if (hit < 0) return CellError.NA;
    // The result vector is indexed positionally and may be shorter than the
    // lookup vector, which Excel reports as #N/A rather than reading past it.
    return hit < out.length ? out.at(hit) : CellError.NA;
  },
};

// ---------------------------------------------------------------------------
// MATCH / XMATCH / XLOOKUP
// ---------------------------------------------------------------------------

const MATCH: FunctionSpec = {
  name: 'MATCH',
  params: [p.scalar('lookup_value'), p.array('lookup_array'), p.scalar('match_type', true)],
  summary: 'Returns the relative position of an item in an array.',
  impl: (args, ctx) => {
    const target = lookupValueOf(args[0], ctx);
    const source = gridOf(args[1], ctx);
    if (isError(source)) return source;
    const vec = vectorOf(source);
    if (!vec) return CellError.NA;
    const type = intOf(args[2], ctx, 1);
    if (isError(type)) return type;

    const hit =
      type === 0
        ? findExact(vec, target, true, false)
        : findOrdered(vec, target, type > 0);
    return hit < 0 ? CellError.NA : hit + 1;
  },
};

/** Shared by XLOOKUP and XMATCH: resolve match_mode and search_mode to an index. */
function xSearch(
  vec: Vector,
  target: Scalar,
  matchMode: number,
  searchMode: number,
): number | CellError {
  if (matchMode !== 0 && matchMode !== -1 && matchMode !== 1 && matchMode !== 2) {
    return CellError.VALUE;
  }
  if (searchMode !== 1 && searchMode !== -1 && searchMode !== 2 && searchMode !== -2) {
    return CellError.VALUE;
  }

  if (searchMode === 2 || searchMode === -2) {
    // Binary search assumes the array is sorted the way the mode says it is;
    // wildcards have no meaning there, so mode 2 degrades to an exact search.
    const approx = matchMode === -1 ? -1 : matchMode === 1 ? 1 : 0;
    return findBinary(vec, target, searchMode === 2, approx);
  }

  const reverse = searchMode === -1;
  if (matchMode === 0 || matchMode === 2) {
    return findExact(vec, target, matchMode === 2, reverse);
  }
  const exact = findExact(vec, target, false, reverse);
  if (exact >= 0) return exact;
  return findClosest(vec, target, matchMode === -1, reverse);
}

const XLOOKUP: FunctionSpec = {
  name: 'XLOOKUP',
  params: [
    p.scalar('lookup_value'),
    p.array('lookup_array'),
    p.array('return_array'),
    p.scalar('if_not_found', true),
    p.scalar('match_mode', true),
    p.scalar('search_mode', true),
  ],
  summary: 'Searches a range or array and returns the corresponding item from a second range.',
  impl: (args, ctx) => {
    const target = lookupValueOf(args[0], ctx);
    const source = gridOf(args[1], ctx);
    if (isError(source)) return source;
    const results = gridOf(args[2], ctx);
    if (isError(results)) return results;
    const matchMode = intOf(args[4], ctx, 0);
    if (isError(matchMode)) return matchMode;
    const searchMode = intOf(args[5], ctx, 1);
    if (isError(searchMode)) return searchMode;

    // A single-column lookup array selects a row of the return array, a
    // single-row one selects a column; anything else has no orientation.
    const horizontal = source.rows === 1 && source.cols > 1;
    if (!horizontal && source.cols !== 1) return CellError.VALUE;
    const span = horizontal ? source.cols : source.rows;
    if ((horizontal ? results.cols : results.rows) !== span) return CellError.VALUE;

    const vec = horizontal ? rowVector(source, 0) : colVector(source, 0);
    const hit = xSearch(vec, target, matchMode, searchMode);
    if (isError(hit)) return hit;
    if (hit < 0) return args[3] === undefined ? CellError.NA : (args[3] as Value);

    if (horizontal) {
      if (results.rows === 1) return results.at(0, hit);
      const data: Scalar[] = [];
      for (let r = 0; r < results.rows; r++) data.push(results.at(r, hit));
      return makeArray(results.rows, 1, data);
    }
    if (results.cols === 1) return results.at(hit, 0);
    const data: Scalar[] = [];
    for (let c = 0; c < results.cols; c++) data.push(results.at(hit, c));
    return makeArray(1, results.cols, data);
  },
};

const XMATCH: FunctionSpec = {
  name: 'XMATCH',
  params: [
    p.scalar('lookup_value'),
    p.array('lookup_array'),
    p.scalar('match_mode', true),
    p.scalar('search_mode', true),
  ],
  summary: 'Returns the relative position of an item in an array or range.',
  impl: (args, ctx) => {
    const target = lookupValueOf(args[0], ctx);
    const source = gridOf(args[1], ctx);
    if (isError(source)) return source;
    const vec = vectorOf(source);
    if (!vec) return CellError.VALUE;
    const matchMode = intOf(args[2], ctx, 0);
    if (isError(matchMode)) return matchMode;
    const searchMode = intOf(args[3], ctx, 1);
    if (isError(searchMode)) return searchMode;

    const hit = xSearch(vec, target, matchMode, searchMode);
    if (isError(hit)) return hit;
    return hit < 0 ? CellError.NA : hit + 1;
  },
};

// ---------------------------------------------------------------------------
// INDEX / CHOOSE / OFFSET
// ---------------------------------------------------------------------------

const INDEX: FunctionSpec = {
  name: 'INDEX',
  params: [
    p.ref('array'),
    p.scalar('row_num'),
    p.scalar('column_num', true),
    p.scalar('area_num', true),
  ],
  structural: true,
  summary: 'Returns a value or reference from within a range.',
  impl: (args, ctx) => {
    const source = args[0];
    if (args[3] !== undefined) {
      const area = intOf(args[3], ctx, 1);
      if (isError(area)) return area;
      if (area < 1) return CellError.VALUE;
      // The evaluator folds a union into its bounding rectangle, so there is no
      // second area to reach.
      if (area > 1) return CellError.REF;
    }

    const ref = isRef(source) ? source : undefined;
    const grid = ref ? undefined : gridOf(source, ctx);
    if (grid && isError(grid)) return grid;
    const rows = ref ? ref.endRow - ref.startRow + 1 : (grid as Grid).rows;
    const cols = ref ? ref.endCol - ref.startCol + 1 : (grid as Grid).cols;

    let rowNum = intOf(args[1], ctx, 0);
    if (isError(rowNum)) return rowNum;
    let colNum = intOf(args[2], ctx, 0);
    if (isError(colNum)) return colNum;

    // A one-row source with no column argument reads its single argument as the
    // column, which is what makes INDEX(A1:E1,3) return C1.
    if (args[2] === undefined && rows === 1 && cols > 1) {
      colNum = rowNum;
      rowNum = 0;
    }

    if (rowNum < 0 || colNum < 0) return CellError.VALUE;
    if (rowNum > rows || colNum > cols) return CellError.REF;

    if (ref) {
      const startRow = rowNum === 0 ? ref.startRow : ref.startRow + rowNum - 1;
      const endRow = rowNum === 0 ? ref.endRow : startRow;
      const startCol = colNum === 0 ? ref.startCol : ref.startCol + colNum - 1;
      const endCol = colNum === 0 ? ref.endCol : startCol;
      return makeRef(ref.sheet, startRow, startCol, endRow, endCol);
    }

    const g = grid as Grid;
    if (rowNum > 0 && colNum > 0) return g.at(rowNum - 1, colNum - 1);
    if (rowNum > 0) {
      const data: Scalar[] = [];
      for (let c = 0; c < cols; c++) data.push(g.at(rowNum - 1, c));
      return makeArray(1, cols, data);
    }
    if (colNum > 0) {
      const data: Scalar[] = [];
      for (let r = 0; r < rows; r++) data.push(g.at(r, colNum - 1));
      return makeArray(rows, 1, data);
    }
    const data: Scalar[] = [];
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) data.push(g.at(r, c));
    return makeArray(rows, cols, data);
  },
};

const CHOOSE: FunctionSpec = {
  name: 'CHOOSE',
  params: [p.scalar('index_num'), p.lazy('value1'), p.rest('values', ArgKind.Lazy)],
  summary: 'Chooses a value from a list of values.',
  impl: (args, ctx) => {
    const index = intOf(args[0], ctx, 0);
    if (isError(index)) return index;
    const choices = args.length - 1;
    if (index < 1 || index > choices) return CellError.VALUE;
    const chosen = args[index];
    if (!isThunk(chosen)) return CellError.VALUE;
    // Only the chosen argument is evaluated, which is why an error or a slow
    // expression in an unselected branch costs nothing, and the result is
    // returned untouched so a reference stays a reference: `A1:CHOOSE(2,B9,C9)`
    // still names a range.
    return ctx.force(chosen);
  },
};

const OFFSET: FunctionSpec = {
  name: 'OFFSET',
  params: [
    p.ref('reference'),
    p.scalar('rows'),
    p.scalar('cols'),
    p.scalar('height', true),
    p.scalar('width', true),
  ],
  volatile: true,
  structural: true,
  summary: 'Returns a reference offset from a given reference.',
  impl: (args, ctx) => {
    const base = args[0];
    if (!isRef(base)) return CellError.VALUE;
    const dr = intOf(args[1], ctx, 0);
    if (isError(dr)) return dr;
    const dc = intOf(args[2], ctx, 0);
    if (isError(dc)) return dc;

    const baseHeight = base.endRow - base.startRow + 1;
    const baseWidth = base.endCol - base.startCol + 1;
    const height = args[3] === undefined || args[3] === null ? baseHeight : intOf(args[3], ctx, 0);
    if (isError(height)) return height;
    const width = args[4] === undefined || args[4] === null ? baseWidth : intOf(args[4], ctx, 0);
    if (isError(width)) return width;
    if (height === 0 || width === 0) return CellError.REF;

    // A negative height or width extends backwards from the anchor rather than
    // forwards; Excel accepts both signs even though the documentation only
    // describes the positive case.
    const anchorRow = base.startRow + dr;
    const anchorCol = base.startCol + dc;
    const startRow = height > 0 ? anchorRow : anchorRow + height + 1;
    const endRow = height > 0 ? anchorRow + height - 1 : anchorRow;
    const startCol = width > 0 ? anchorCol : anchorCol + width + 1;
    const endCol = width > 0 ? anchorCol + width - 1 : anchorCol;

    if (startRow < 0 || startCol < 0 || endRow >= MAX_ROWS || endCol >= MAX_COLS) {
      return CellError.REF;
    }
    return makeRef(base.sheet, startRow, startCol, endRow, endCol);
  },
};

// ---------------------------------------------------------------------------
// INDIRECT and ADDRESS
// ---------------------------------------------------------------------------

/** Split `Sheet!rest` at the last separator outside a quoted sheet name. */
function splitSheet(text: string): { sheet?: string; rest: string } {
  const bang = text.lastIndexOf('!');
  if (bang < 0) return { rest: text };
  return { sheet: text.slice(0, bang), rest: text.slice(bang + 1) };
}

const R1C1_PART = /^R(\[-?\d+\]|-?\d+)?C(\[-?\d+\]|-?\d+)?$/i;

/** One R1C1 address to an absolute A1 address, or undefined if it is not one. */
function r1c1ToA1(token: string, ctx: FunctionContext): string | undefined {
  const m = R1C1_PART.exec(token.trim());
  if (!m) return undefined;
  const axis = (raw: string | undefined, origin: number): number | undefined => {
    if (raw === undefined) return origin;
    if (raw.startsWith('[')) return origin + Number(raw.slice(1, -1));
    const abs = Number(raw) - 1;
    return abs;
  };
  const row = axis(m[1], ctx.origin.row);
  const col = axis(m[2], ctx.origin.col);
  if (row === undefined || col === undefined) return undefined;
  if (!Number.isFinite(row) || !Number.isFinite(col)) return undefined;
  if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) return undefined;
  return `$${colToName(col)}$${row + 1}`;
}

/** Rewrite R1C1 reference text as A1 so one parser handles both styles. */
function r1c1TextToA1(text: string, ctx: FunctionContext): string | undefined {
  const { sheet, rest } = splitSheet(text.trim());
  const parts = rest.split(':');
  const converted: string[] = [];
  for (const part of parts) {
    const a1 = r1c1ToA1(part, ctx);
    if (a1 === undefined) return undefined;
    converted.push(a1);
  }
  if (converted.length > 2) return undefined;
  return `${sheet === undefined ? '' : `${sheet}!`}${converted.join(':')}`;
}

/**
 * Turn a parsed reference expression into a RefValue.
 *
 * The AST was parsed at origin (0,0), so a relative node's stored offset is
 * already the absolute index it denotes - INDIRECT text is never relative to
 * the calling cell, whatever `$` signs it carries.
 */
function refFromAst(node: Ast, ctx: FunctionContext): RefValue | CellError {
  switch (node.kind) {
    case Node.Paren:
      return refFromAst(node.inner, ctx);
    case Node.Ref: {
      if (node.book !== undefined) return CellError.REF;
      const sheet = node.sheet ?? ctx.origin.sheet;
      if (!ctx.hasSheet(sheet)) return CellError.REF;
      if (node.row < 0 || node.row >= MAX_ROWS || node.col < 0 || node.col >= MAX_COLS) {
        return CellError.REF;
      }
      return makeRef(sheet, node.row, node.col);
    }
    case Node.Range: {
      const start = refFromAst(node.start, ctx);
      if (isError(start)) return start;
      const end = refFromAst(node.end, ctx);
      if (isError(end)) return end;
      return makeRef(
        start.sheet,
        Math.min(start.startRow, end.startRow),
        Math.min(start.startCol, end.startCol),
        Math.max(start.endRow, end.endRow),
        Math.max(start.endCol, end.endCol),
      );
    }
    case Node.Beam: {
      if (node.book !== undefined) return CellError.REF;
      const sheet = node.sheet ?? ctx.origin.sheet;
      if (!ctx.hasSheet(sheet)) return CellError.REF;
      const from = Math.min(node.from, node.to);
      const to = Math.max(node.from, node.to);
      if (node.axis === 'col') {
        if (from < 0 || to >= MAX_COLS) return CellError.REF;
        return makeRef(sheet, 0, from, MAX_ROWS - 1, to);
      }
      if (from < 0 || to >= MAX_ROWS) return CellError.REF;
      return makeRef(sheet, from, 0, to, MAX_COLS - 1);
    }
    default:
      // Defined names would need the workbook's name table, which the function
      // context does not expose; anything else is not a reference at all.
      return CellError.REF;
  }
}

const INDIRECT: FunctionSpec = {
  name: 'INDIRECT',
  params: [p.scalar('ref_text'), p.scalar('a1', true)],
  volatile: true,
  summary: 'Returns the reference specified by a text string.',
  impl: (args, ctx) => {
    const text = toText(scalarOf(args[0], ctx));
    if (isError(text)) return text;
    const a1 = args[1] === undefined ? true : toBoolean(scalarOf(args[1], ctx));
    if (isError(a1)) return a1;

    const source = a1 ? text.trim() : r1c1TextToA1(text, ctx);
    if (source === undefined || source === '') return CellError.REF;
    try {
      return refFromAst(parseFormula(source), ctx);
    } catch {
      return CellError.REF;
    }
  },
};

const ADDRESS: FunctionSpec = {
  name: 'ADDRESS',
  params: [
    p.scalar('row_num'),
    p.scalar('column_num'),
    p.scalar('abs_num', true),
    p.scalar('a1', true),
    p.scalar('sheet_text', true),
  ],
  summary: 'Returns a reference as text to a single cell in a worksheet.',
  impl: (args, ctx) => {
    const row = intOf(args[0], ctx, 0);
    if (isError(row)) return row;
    const col = intOf(args[1], ctx, 0);
    if (isError(col)) return col;
    // An argument left blank by an empty comma keeps the default here rather
    // than coercing to 0, because ADDRESS(2,3,,FALSE) is a common idiom and
    // Excel answers R2C3 for it.
    const abs = args[2] === null ? 1 : intOf(args[2], ctx, 1);
    if (isError(abs)) return abs;
    const a1 = args[3] === undefined ? true : toBoolean(scalarOf(args[3], ctx));
    if (isError(a1)) return a1;
    if (row < 1 || row > MAX_ROWS || col < 1 || col > MAX_COLS) return CellError.VALUE;
    if (abs < 1 || abs > 4) return CellError.VALUE;

    const rowAbs = abs === 1 || abs === 2;
    const colAbs = abs === 1 || abs === 3;
    let address: string;
    if (a1) {
      address = `${colAbs ? '$' : ''}${colToName(col - 1)}${rowAbs ? '$' : ''}${row}`;
    } else {
      // The relative form brackets the number given rather than an offset from
      // the calling cell: ADDRESS(2,3,2,FALSE) is R2C[3] wherever it sits, which
      // is what Microsoft's own example documents.
      address = `R${rowAbs ? row : `[${row}]`}C${colAbs ? col : `[${col}]`}`;
    }

    const sheet = args[4] === undefined ? null : scalarOf(args[4], ctx);
    if (sheet === null) return address;
    const name = toText(sheet);
    if (isError(name)) return name;
    // An empty sheet name still produces the separator, matching Excel.
    return `${name === '' ? '' : quoteSheetName(name)}!${address}`;
  },
};

// ---------------------------------------------------------------------------
// Shape: ROW, COLUMN, ROWS, COLUMNS, AREAS
// ---------------------------------------------------------------------------

/**
 * ROW and COLUMN over a block return one number per row or column, not a
 * scalar: ROW(A2:A4) is {2;3;4}. Only a single-cell reference collapses.
 */
function axisNumbers(v: Value | undefined, ctx: FunctionContext, wantRow: boolean): Value {
  if (v === undefined) return (wantRow ? ctx.origin.row : ctx.origin.col) + 1;
  if (isError(v)) return v;
  if (!isRef(v)) return CellError.VALUE;
  const from = wantRow ? v.startRow : v.startCol;
  const to = wantRow ? v.endRow : v.endCol;
  if (from === to) return from + 1;
  const data: Scalar[] = [];
  for (let i = from; i <= to; i++) data.push(i + 1);
  return wantRow ? makeArray(data.length, 1, data) : makeArray(1, data.length, data);
}

function axisCount(v: Value | undefined, ctx: FunctionContext, wantRow: boolean): Value {
  if (v === undefined) return CellError.VALUE;
  if (isError(v)) return v;
  if (isRef(v)) return wantRow ? v.endRow - v.startRow + 1 : v.endCol - v.startCol + 1;
  if (isArray(v)) return wantRow ? v.rows : v.cols;
  return CellError.VALUE;
}

const ROW: FunctionSpec = {
  name: 'ROW',
  params: [p.ref('reference', true)],
  structural: true,
  summary: 'Returns the row number of a reference.',
  impl: (args, ctx) => axisNumbers(args[0], ctx, true),
};

const COLUMN: FunctionSpec = {
  name: 'COLUMN',
  params: [p.ref('reference', true)],
  structural: true,
  summary: 'Returns the column number of a reference.',
  impl: (args, ctx) => axisNumbers(args[0], ctx, false),
};

const ROWS: FunctionSpec = {
  name: 'ROWS',
  params: [p.ref('array')],
  structural: true,
  summary: 'Returns the number of rows in a reference or array.',
  impl: (args, ctx) => axisCount(args[0], ctx, true),
};

const COLUMNS: FunctionSpec = {
  name: 'COLUMNS',
  params: [p.ref('array')],
  structural: true,
  summary: 'Returns the number of columns in a reference or array.',
  impl: (args, ctx) => axisCount(args[0], ctx, false),
};

/** Areas separated by the union operator, unwrapping the parentheses around it. */
function countAreas(node: Ast): number {
  if (node.kind === Node.Paren) return countAreas(node.inner);
  if (node.kind === Node.Binary && node.op === ',') {
    return countAreas(node.left) + countAreas(node.right);
  }
  return 1;
}

const AREAS: FunctionSpec = {
  name: 'AREAS',
  params: [p.lazy('reference')],
  structural: true,
  summary: 'Returns the number of areas in a reference.',
  impl: (args, ctx) => {
    const thunk = args[0];
    if (!isThunk(thunk)) return CellError.VALUE;
    const value = ctx.force(thunk);
    if (isError(value)) return value;
    if (!isRef(value)) return CellError.VALUE;
    // The evaluated union is one rectangle, so the count comes from the syntax.
    return countAreas(thunk.ast);
  },
};

// ---------------------------------------------------------------------------
// FORMULATEXT and HYPERLINK
// ---------------------------------------------------------------------------

const FORMULATEXT: FunctionSpec = {
  name: 'FORMULATEXT',
  params: [p.ref('reference')],
  summary: 'Returns the formula at the given reference as text.',
  impl: (args) => {
    if (!isRef(args[0])) return CellError.VALUE;
    // FunctionContext carries values, not formulas. #N/A is Excel's answer for a
    // cell that holds no formula, which is the closest true statement available
    // until the context can be asked for the formula text.
    return CellError.NA;
  },
};

const HYPERLINK: FunctionSpec = {
  name: 'HYPERLINK',
  params: [p.scalar('link_location'), p.scalar('friendly_name', true)],
  summary: 'Creates a shortcut that opens a document or a location on a network.',
  impl: (args, ctx) => {
    // The jump target is a display concern; the cell shows the friendly name, or
    // the location itself when none was given.
    const shown = args[1] === undefined ? scalarOf(args[0], ctx) : scalarOf(args[1], ctx);
    return toText(shown);
  },
};

export const LOOKUP_FUNCTIONS: readonly FunctionSpec[] = [
  VLOOKUP,
  HLOOKUP,
  LOOKUP,
  INDEX,
  MATCH,
  XLOOKUP,
  XMATCH,
  CHOOSE,
  OFFSET,
  INDIRECT,
  ADDRESS,
  ROW,
  COLUMN,
  ROWS,
  COLUMNS,
  AREAS,
  FORMULATEXT,
  HYPERLINK,
];
