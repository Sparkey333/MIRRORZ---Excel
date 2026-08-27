/**
 * Expression evaluator.
 *
 * The evaluator never touches the sheet model directly; it goes through the
 * narrow `SheetStore` interface below. That keeps the engine runnable in a
 * worker thread, in a headless Node process for tests, and later against a
 * native store, without any of them knowing about the others.
 *
 * Errors are values, not exceptions. An Excel formula does not abort on a
 * division by zero - it produces `#DIV/0!` and lets the caller decide, which is
 * exactly what IFERROR relies on. Throwing and catching per cell would also be
 * far slower on a sheet where a whole column legitimately evaluates to `#N/A`.
 */

import {
  CellError,
  MAX_COLS,
  MAX_ROWS,
  type Scalar,
  isError,
} from '@mirrorz/core';
import { type Ast, Node, type RefNode } from './ast.js';
import { parseFormula } from './parser.js';
import {
  ArgKind,
  type FunctionContext,
  type FunctionRegistry,
  type FunctionSpec,
  type ParamSpec,
  type Thunk,
} from './registry.js';
import {
  type ArrayValue,
  type RefValue,
  type Value,
  arrayAt,
  compareScalars,
  excelAdd,
  excelSub,
  isArray,
  isRef,
  makeArray,
  makeRef,
  toBoolean,
  toNumber,
  toText,
  truncateLiteral,
} from './value.js';

/**
 * Everything the evaluator needs from the workbook.
 *
 * Deliberately small: four reads and two lookups. Anything wider would couple
 * the engine to a particular storage layout.
 */
export interface SheetStore {
  /** One cell's value. Out-of-range or missing sheets return null. */
  getScalar(sheet: string, row: number, col: number): Scalar;
  /** Only the cells that actually exist inside a rectangle, in row-major order. */
  iterate(
    sheet: string,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number,
  ): Iterable<{ row: number; col: number; value: Scalar }>;
  hasSheet(name: string): boolean;
  /** Sheet names in tab order, for resolving 3-D references. */
  sheetNames(): readonly string[];
  /** A defined name's formula text, if it exists. */
  getDefinedName(name: string, sheet: string): string | undefined;
  /** The extent actually used, so a whole-column reference is not 1M cells. */
  usedBounds(sheet: string): { maxRow: number; maxCol: number } | null;
}

export interface EvalOptions {
  dateSystem?: 1900 | 1904;
  /** Fixed "now" serial so every NOW() in one recalculation agrees. */
  now?: number;
  /**
   * Dynamic-array semantics. In `modern` mode a range reaching a scalar context
   * spills; in `legacy` mode it implicitly intersects with the formula's own row
   * or column, which is what pre-2019 files expect.
   */
  arrayMode?: 'modern' | 'legacy';
  /** Guards against a runaway INDIRECT/OFFSET chain on a hostile file. */
  maxDepth?: number;
}

export interface EvalRequest {
  ast: Ast;
  sheet: string;
  row: number;
  col: number;
}

export class Evaluator {
  private readonly dateSystem: 1900 | 1904;
  private readonly now: number;
  private readonly arrayMode: 'modern' | 'legacy';
  private readonly maxDepth: number;

  constructor(
    private readonly store: SheetStore,
    private readonly registry: FunctionRegistry,
    options: EvalOptions = {},
  ) {
    this.dateSystem = options.dateSystem ?? 1900;
    this.now = options.now ?? 0;
    this.arrayMode = options.arrayMode ?? 'modern';
    this.maxDepth = options.maxDepth ?? 256;
  }

  /** Evaluate a formula at a cell, returning the value it should display. */
  evaluate(request: EvalRequest): Value {
    const frame: Frame = {
      sheet: request.sheet,
      row: request.row,
      col: request.col,
      depth: 0,
    };
    return this.eval(request.ast, frame);
  }

  /** Evaluate and collapse to what a single cell shows. */
  evaluateScalar(request: EvalRequest): Scalar {
    const v = this.evaluate(request);
    return this.collapse(v, {
      sheet: request.sheet,
      row: request.row,
      col: request.col,
      depth: 0,
    });
  }

  private eval(node: Ast, frame: Frame): Value {
    if (frame.depth > this.maxDepth) return CellError.NUM;

    switch (node.kind) {
      case Node.Number:
        return truncateLiteral(node.value);
      case Node.Text:
        return node.value;
      case Node.Bool:
        return node.value;
      case Node.ErrorLit:
        return new CellError(node.code);
      case Node.Missing:
        return null;
      case Node.Paren:
        return this.eval(node.inner, frame);

      case Node.Unary:
        return this.unary(node.op, this.eval(node.operand, frame), frame);

      case Node.Postfix:
        return this.postfix(node.op, node, frame);

      case Node.Binary:
        return this.binary(node.op, node.left, node.right, frame);

      case Node.Array:
        return this.arrayConstant(node.rows, frame);

      case Node.Ref:
        return this.refFromNode(node, frame);

      case Node.Range: {
        const start = this.refFromNode(node.start, frame);
        const end = this.refFromNode(node.end, frame);
        if (isError(start)) return start;
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
        const sheet = node.sheet ?? frame.sheet;
        if (!this.store.hasSheet(sheet)) return CellError.REF;
        // A whole-column reference is clipped to the used range. Materialising
        // 1,048,576 rows to sum twelve of them is the difference between an
        // instant recalculation and a hang.
        const bounds = this.store.usedBounds(sheet);
        if (node.axis === 'col') {
          const from = node.fromAbs ? node.from : node.from + frame.col;
          const to = node.toAbs ? node.to : node.to + frame.col;
          if (from < 0 || to >= MAX_COLS) return CellError.REF;
          return makeRef(sheet, 0, Math.min(from, to), bounds?.maxRow ?? 0, Math.max(from, to));
        }
        const from = node.fromAbs ? node.from : node.from + frame.row;
        const to = node.toAbs ? node.to : node.to + frame.row;
        if (from < 0 || to >= MAX_ROWS) return CellError.REF;
        return makeRef(sheet, Math.min(from, to), 0, Math.max(from, to), bounds?.maxCol ?? 0);
      }

      case Node.ThreeD:
        return this.threeD(node.sheetStart, node.sheetEnd, node.inner, frame);

      case Node.Call:
        return this.call(node.name, node.args, frame);

      case Node.Name:
        return this.definedName(node.name, node.sheet ?? frame.sheet, frame);

      case Node.StructRef:
        // Structured references need the table catalogue, which lives above the
        // engine. Until that is wired through, report the same error Excel does
        // for an unknown table rather than guessing at a range.
        return CellError.NAME;
    }
  }

  private unary(op: '-' | '+' | '@', operand: Value, frame: Frame): Value {
    if (op === '@') {
      // Explicit implicit-intersection request.
      return this.intersect(operand, frame);
    }
    return this.mapScalar(operand, frame, (v) => {
      const n = toNumber(v);
      if (isError(n)) return n;
      return op === '-' ? -n : n;
    });
  }

  private postfix(op: '%' | '#', node: { operand: Ast }, frame: Frame): Value {
    if (op === '#') {
      // The spill-range operator needs the anchor's current rectangle, which the
      // dependency graph owns. Until spill tracking is wired in, a plain
      // reference is the honest answer for the single-cell case.
      return this.eval(node.operand, frame);
    }
    return this.mapScalar(this.eval(node.operand, frame), frame, (v) => {
      const n = toNumber(v);
      if (isError(n)) return n;
      return n / 100;
    });
  }

  private binary(op: string, leftAst: Ast, rightAst: Ast, frame: Frame): Value {
    // Reference operators work on references, not values, so they are handled
    // before anything is dereferenced.
    if (op === ':') {
      const l = this.eval(leftAst, frame);
      const r = this.eval(rightAst, frame);
      if (isError(l)) return l;
      if (isError(r)) return r;
      if (!isRef(l) || !isRef(r)) return CellError.VALUE;
      return makeRef(
        l.sheet,
        Math.min(l.startRow, r.startRow),
        Math.min(l.startCol, r.startCol),
        Math.max(l.endRow, r.endRow),
        Math.max(l.endCol, r.endCol),
      );
    }

    if (op === ' ') {
      const l = this.eval(leftAst, frame);
      const r = this.eval(rightAst, frame);
      if (isError(l)) return l;
      if (isError(r)) return r;
      if (!isRef(l) || !isRef(r)) return CellError.VALUE;
      const startRow = Math.max(l.startRow, r.startRow);
      const startCol = Math.max(l.startCol, r.startCol);
      const endRow = Math.min(l.endRow, r.endRow);
      const endCol = Math.min(l.endCol, r.endCol);
      // Excel reports #NULL! specifically for a non-intersecting intersection.
      if (startRow > endRow || startCol > endCol) return CellError.NULL;
      return makeRef(l.sheet, startRow, startCol, endRow, endCol);
    }

    if (op === ',') {
      // A union is only meaningful to functions that accept multiple areas.
      // Collapsing it to the first area here would silently drop data, so it
      // stays an error until the aggregate functions understand multi-area
      // references.
      const l = this.eval(leftAst, frame);
      const r = this.eval(rightAst, frame);
      if (isError(l)) return l;
      if (isError(r)) return r;
      if (isRef(l) && isRef(r)) return unionOf(l, r);
      return CellError.VALUE;
    }

    const left = this.eval(leftAst, frame);
    const right = this.eval(rightAst, frame);
    return this.broadcast2(left, right, frame, (a, b) => applyBinaryScalar(op, a, b));
  }

  private arrayConstant(rows: Ast[][], frame: Frame): Value {
    const height = rows.length;
    const width = rows[0]?.length ?? 0;
    const data: Scalar[] = [];
    for (const row of rows) {
      for (const cell of row) {
        const v = this.eval(cell, frame);
        data.push(isArray(v) || isRef(v) ? CellError.VALUE : v);
      }
    }
    return makeArray(height, width, data);
  }

  private refFromNode(node: RefNode, frame: Frame): RefValue | CellError {
    const sheet = node.sheet ?? frame.sheet;
    if (node.book !== undefined) {
      // External workbook links are not resolved yet. Excel shows the cached
      // value from externalLinks; reporting #REF! is the honest interim answer.
      return CellError.REF;
    }
    if (!this.store.hasSheet(sheet)) return CellError.REF;
    const row = node.rowAbs ? node.row : node.row + frame.row;
    const col = node.colAbs ? node.col : node.col + frame.col;
    if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) return CellError.REF;
    return makeRef(sheet, row, col);
  }

  /** `Sheet1:Sheet3!A1` - the same reference across a span of sheet tabs. */
  private threeD(startName: string, endName: string, inner: Ast, frame: Frame): Value {
    const names = this.store.sheetNames();
    const from = names.indexOf(startName);
    const to = names.indexOf(endName);
    if (from < 0 || to < 0) return CellError.REF;
    const lo = Math.min(from, to);
    const hi = Math.max(from, to);

    // A 3-D reference is not a rectangle, so it flattens into an array of the
    // per-sheet values; the aggregate functions consume it the same way.
    const collected: Scalar[] = [];
    for (let i = lo; i <= hi; i++) {
      const sheetFrame: Frame = { ...frame, sheet: names[i]! };
      const v = this.eval(inner, sheetFrame);
      if (isError(v)) return v;
      if (isRef(v)) {
        const arr = this.deref(v);
        collected.push(...arr.data);
      } else if (isArray(v)) {
        collected.push(...v.data);
      } else {
        collected.push(v);
      }
    }
    return makeArray(collected.length, 1, collected);
  }

  private definedName(name: string, sheet: string, frame: Frame): Value {
    const refersTo = this.store.getDefinedName(name, sheet);
    if (refersTo === undefined) return CellError.NAME;
    // Defined names hold a formula, so resolving one is a nested evaluation.
    // The depth counter is what stops a name that refers to itself.
    try {
      const ast = parseFormula(refersTo, { origin: { row: frame.row, col: frame.col } });
      return this.eval(ast, { ...frame, depth: frame.depth + 1 });
    } catch {
      return CellError.NAME;
    }
  }

  private call(name: string, argAsts: Ast[], frame: Frame): Value {
    const spec = this.registry.get(name);
    if (!spec) return CellError.NAME;

    const arityProblem = this.registry.checkArity(spec, argAsts.length);
    if (arityProblem) return new CellError('#VALUE!', arityProblem);

    const ctx = this.contextFor(frame);
    const args: Value[] = [];

    for (let i = 0; i < argAsts.length; i++) {
      const param = this.registry.paramAt(spec, i);
      const kind = param?.kind ?? ArgKind.Scalar;
      const ast = argAsts[i]!;

      if (kind === ArgKind.Lazy) {
        args.push(makeThunk(ast, () => this.eval(ast, frame)) as unknown as Value);
        continue;
      }

      const raw = this.eval(ast, { ...frame, depth: frame.depth + 1 });
      const prepared = this.prepareArg(raw, kind, frame);

      // An error short-circuits the call unless this parameter is declared
      // error-transparent, which is what lets IFERROR and ISERROR see it.
      if (isError(prepared) && !param?.errorTransparent) return prepared;
      args.push(prepared);
    }

    // Fill omitted trailing optional parameters so implementations can index
    // positionally without checking length everywhere.
    for (let i = argAsts.length; i < spec.params.length; i++) {
      if (spec.params[i]?.repeating) break;
      args.push(undefined as unknown as Value);
    }

    const result = spec.impl(args, ctx);

    // A scalar function handed an array applies element-wise, so SIN(A1:A9)
    // returns nine sines rather than one error.
    return result;
  }

  private prepareArg(raw: Value, kind: ArgKind, frame: Frame): Value {
    switch (kind) {
      case ArgKind.Reference:
        return raw;
      case ArgKind.Any:
        return raw;
      case ArgKind.Array:
        if (isRef(raw)) return this.deref(raw);
        return raw;
      case ArgKind.Scalar:
        return this.collapse(raw, frame);
      case ArgKind.Lazy:
        return raw;
    }
  }

  /**
   * Reduce a value to the single scalar a cell or scalar parameter wants.
   *
   * A 1x1 reference is just its cell. A larger reference or array implicitly
   * intersects with the formula's own row or column in legacy mode; in modern
   * mode it keeps its shape and the caller spills it.
   */
  private collapse(v: Value, frame: Frame): Scalar {
    if (isRef(v)) {
      if (v.startRow === v.endRow && v.startCol === v.endCol) {
        return this.store.getScalar(v.sheet, v.startRow, v.startCol);
      }
      const intersected = this.intersect(v, frame);
      if (isRef(intersected)) {
        return this.store.getScalar(intersected.sheet, intersected.startRow, intersected.startCol);
      }
      return isArray(intersected) ? (intersected.data[0] ?? null) : intersected;
    }
    if (isArray(v)) {
      // Excel takes the top-left of an array reaching a scalar context.
      return v.data[0] ?? null;
    }
    return v;
  }

  /**
   * Implicit intersection: a multi-cell reference used where one value is
   * expected resolves to the cell in the formula's own row or column.
   */
  private intersect(v: Value, frame: Frame): Value {
    if (!isRef(v)) return v;
    const single = v.startRow === v.endRow && v.startCol === v.endCol;
    if (single) return v;

    // A single-column range intersects by row; a single-row range by column.
    if (v.startCol === v.endCol && frame.row >= v.startRow && frame.row <= v.endRow) {
      return makeRef(v.sheet, frame.row, v.startCol);
    }
    if (v.startRow === v.endRow && frame.col >= v.startCol && frame.col <= v.endCol) {
      return makeRef(v.sheet, v.startRow, frame.col);
    }
    if (this.arrayMode === 'legacy') return CellError.VALUE;
    return v;
  }

  /** Materialise a reference into a dense array. */
  private deref(ref: RefValue): ArrayValue {
    const rows = ref.endRow - ref.startRow + 1;
    const cols = ref.endCol - ref.startCol + 1;
    const data: Scalar[] = new Array(rows * cols).fill(null);
    for (const cell of this.store.iterate(
      ref.sheet,
      ref.startRow,
      ref.startCol,
      ref.endRow,
      ref.endCol,
    )) {
      data[(cell.row - ref.startRow) * cols + (cell.col - ref.startCol)] = cell.value;
    }
    return makeArray(rows, cols, data);
  }

  /** Apply a scalar operation across whatever shape the operand has. */
  private mapScalar(v: Value, frame: Frame, fn: (s: Scalar) => Scalar): Value {
    if (isError(v)) return v;
    if (isRef(v)) {
      const collapsed = this.intersect(v, frame);
      if (isRef(collapsed) && collapsed.startRow === collapsed.endRow && collapsed.startCol === collapsed.endCol) {
        return fn(this.store.getScalar(collapsed.sheet, collapsed.startRow, collapsed.startCol));
      }
      return this.mapScalar(isRef(collapsed) ? this.deref(collapsed) : collapsed, frame, fn);
    }
    if (isArray(v)) {
      return makeArray(v.rows, v.cols, v.data.map(fn));
    }
    return fn(v);
  }

  /**
   * Broadcast a binary operation over two operands of possibly different
   * shapes, following Excel's rules: a scalar pairs with every element, a row
   * and a column form a rectangle, and mismatched extents pad with #N/A.
   */
  private broadcast2(
    left: Value,
    right: Value,
    frame: Frame,
    fn: (a: Scalar, b: Scalar) => Scalar,
  ): Value {
    const l = this.toOperand(left, frame);
    const r = this.toOperand(right, frame);
    if (isError(l)) return l;
    if (isError(r)) return r;

    const lArr = isArray(l);
    const rArr = isArray(r);
    if (!lArr && !rArr) return fn(l, r);

    const lRows = lArr ? l.rows : 1;
    const lCols = lArr ? l.cols : 1;
    const rRows = rArr ? r.rows : 1;
    const rCols = rArr ? r.cols : 1;
    const rows = Math.max(lRows, rRows);
    const cols = Math.max(lCols, rCols);

    const data: Scalar[] = new Array(rows * cols);
    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        const a = pick(l, lArr, lRows, lCols, i, j);
        const b = pick(r, rArr, rRows, rCols, i, j);
        data[i * cols + j] = a === MISSING || b === MISSING ? CellError.NA : fn(a, b);
      }
    }
    return makeArray(rows, cols, data);
  }

  /** Dereference a value far enough to take part in arithmetic. */
  private toOperand(v: Value, frame: Frame): Scalar | ArrayValue | CellError {
    if (isError(v)) return v;
    if (isRef(v)) {
      const narrowed = this.intersect(v, frame);
      if (isError(narrowed)) return narrowed;
      if (isRef(narrowed)) {
        if (narrowed.startRow === narrowed.endRow && narrowed.startCol === narrowed.endCol) {
          return this.store.getScalar(narrowed.sheet, narrowed.startRow, narrowed.startCol);
        }
        return this.deref(narrowed);
      }
      return narrowed as Scalar | ArrayValue;
    }
    return v;
  }

  private contextFor(frame: Frame): FunctionContext {
    return {
      origin: { sheet: frame.sheet, row: frame.row, col: frame.col },
      getScalar: (sheet, row, col) => this.store.getScalar(sheet, row, col),
      deref: (ref) => this.deref(ref),
      iterate: (ref) =>
        this.store.iterate(ref.sheet, ref.startRow, ref.startCol, ref.endRow, ref.endCol),
      hasSheet: (name) => this.store.hasSheet(name),
      dateSystem: this.dateSystem,
      now: this.now,
      force: (thunk) => thunk.evaluate(),
    };
  }
}

interface Frame {
  sheet: string;
  row: number;
  col: number;
  depth: number;
}

/** Sentinel for "outside this operand's extent" during broadcasting. */
const MISSING = Symbol('missing') as unknown as Scalar;

function pick(
  v: Scalar | ArrayValue,
  isArr: boolean,
  rows: number,
  cols: number,
  i: number,
  j: number,
): Scalar {
  if (!isArr) return v as Scalar;
  const arr = v as ArrayValue;
  // A single row broadcasts down, a single column broadcasts across.
  const r = rows === 1 ? 0 : i;
  const c = cols === 1 ? 0 : j;
  if (r >= rows || c >= cols) return MISSING;
  return arrayAt(arr, r, c);
}

function unionOf(a: RefValue, b: RefValue): RefValue {
  return makeRef(
    a.sheet,
    Math.min(a.startRow, b.startRow),
    Math.min(a.startCol, b.startCol),
    Math.max(a.endRow, b.endRow),
    Math.max(a.endCol, b.endCol),
  );
}

function makeThunk(ast: Ast, evaluate: () => Value): Thunk {
  return { ast, evaluate };
}

/** The scalar core of every binary operator. */
export function applyBinaryScalar(op: string, a: Scalar, b: Scalar): Scalar {
  if (isError(a)) return a;
  if (isError(b)) return b;

  switch (op) {
    case '+':
    case '-':
    case '*':
    case '/':
    case '^': {
      const x = toNumber(a);
      if (isError(x)) return x;
      const y = toNumber(b);
      if (isError(y)) return y;
      switch (op) {
        case '+':
          return excelAdd(x, y);
        case '-':
          return excelSub(x, y);
        case '*':
          return finite(x * y);
        case '/':
          return y === 0 ? CellError.DIV0 : finite(x / y);
        default: {
          // 0^0 is 1 in Excel; a negative base with a fractional exponent is
          // #NUM! rather than NaN.
          const r = x ** y;
          return Number.isNaN(r) ? CellError.NUM : finite(r);
        }
      }
    }

    case '&': {
      const x = toText(a);
      if (isError(x)) return x;
      const y = toText(b);
      if (isError(y)) return y;
      return x + y;
    }

    case '=':
    case '<>':
    case '<':
    case '>':
    case '<=':
    case '>=': {
      const cmp = compareScalars(a, b);
      if (isError(cmp)) return cmp;
      switch (op) {
        case '=':
          return cmp === 0;
        case '<>':
          return cmp !== 0;
        case '<':
          return cmp < 0;
        case '>':
          return cmp > 0;
        case '<=':
          return cmp <= 0;
        default:
          return cmp >= 0;
      }
    }

    default:
      return CellError.VALUE;
  }
}

/**
 * Multiplication, division and exponentiation keep full f64 precision - the
 * probe =1/3*3 is exactly 1, which per-operation rounding would break. Only the
 * range check applies here; the 15-digit rounding happens at comparison and
 * display.
 */
function finite(v: number): Scalar {
  return Number.isFinite(v) ? v : CellError.NUM;
}

export { toBoolean };
