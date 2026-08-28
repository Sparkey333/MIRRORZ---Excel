/**
 * Dynamic arrays, and the lambda calculus that came with them.
 *
 * Seven decisions shape this file.
 *
 * First, spilling is not here. FILTER, SORT, SEQUENCE and the rest return an
 * ArrayValue and stop; deciding where that array lands, whether it collides with
 * occupied cells (#SPILL!), and what `A1#` resolves to are all properties of the
 * sheet and belong to the dependency graph. A function that tried to answer
 * those questions would need to know its own output rectangle, which is exactly
 * the knowledge the graph exists to hold.
 *
 * Second, an empty result is #CALC!, not #N/A and not a zero-sized array. Excel
 * has no representation for an array with no elements, so every path that could
 * produce one - FILTER matching nothing without if_empty, UNIQUE(...,,TRUE)
 * finding no singleton, TAKE(...,0), DROP dropping everything, TOCOL ignoring
 * every cell - reports #CALC! instead. Returning #N/A there is the single most
 * common wrong answer in FILTER clones.
 *
 * Third, LET and LAMBDA get a real lexical environment, built out of the two
 * things FunctionContext already provides: a thunk for each argument, which
 * carries both its unevaluated AST and a closure that evaluates it in the
 * calling cell's frame. Binding a name is therefore not a runtime lookup at all.
 * At the moment a lambda is first invoked (or a LET name first bound) we walk
 * the body once and record the AST nodes that mention that name - its slots -
 * honouring shadowing by nested LET and LAMBDA. Every later invocation patches
 * those recorded nodes to a constant and forces the body thunk, so the cost per
 * invocation is proportional to the number of uses of the parameter rather than
 * to the size of the body, and no string comparison happens during evaluation.
 *
 * The patch is a node rewrite because the evaluator owns evaluation: there is no
 * "evaluate this AST with these bindings" entry point on FunctionContext, and
 * duplicating the evaluator here to get one would guarantee the copy drifts.
 * Rewriting a slot into a call to the internal `__MZ.CONST` (which returns a
 * stashed value verbatim, at full precision - a NumberNode would be truncated to
 * fifteen digits on the way through) means the body is evaluated by the real
 * evaluator, with real broadcasting, real error short-circuiting and real
 * function dispatch. Patches are restored in a finally, in LIFO order, so nested
 * and repeated invocation are both safe. `__MZ.APPLY` is the same trick for the
 * call form: `LET(f,LAMBDA(x,x*2),f(3))` parses `f(3)` as an ordinary call, and
 * the slot rewrite turns it into an application of the bound lambda. Applying a
 * lambda literal directly, `LAMBDA(x,x+1)(5)`, is not reachable from here: the
 * parser has no postfix call form, so that formula never becomes an AST.
 *
 * Fourth, a LAMBDA value is a CellError subclass whose code is #CALC!. That is
 * not a trick to avoid modelling it: `=LAMBDA(x,x+1)` typed into a cell really
 * is #CALC! in Excel, and making the callable value carry that code means it
 * degrades to Excel's answer everywhere it escapes - a cell, an argument to SUM,
 * an operand of `+` - while the higher-order functions, which look for it by
 * identity before they look at errors, still see a lambda.
 *
 * Fifth, sorting. SORT and SORTBY are stable, which Excel's are, and they order
 * with compareScalars so the type ranking is Excel's (number, then text, then
 * FALSE, then TRUE, text compared case-insensitively) rather than JavaScript's.
 * Blanks and errors are the two things compareScalars cannot rank for a sort -
 * it treats a blank as the zero of whatever it meets, and it propagates an error
 * rather than ordering it - so this file adds the rule Excel's sort engine uses:
 * values first, then errors, then blanks, in both directions.
 *
 * Sixth, volatility. RANDARRAY is volatile and SEQUENCE is not; nothing else
 * here is. In particular the reshaping family is pure, and ANCHORARRAY is an
 * ordinary reference dependency, not a volatile one.
 *
 * Seventh, the size guard. SEQUENCE and RANDARRAY are the only functions in the
 * library that can be asked to build an arbitrarily large array from two small
 * numbers. Excel answers #SPILL! because it fails at placement; we have no
 * placement to fail at yet, so a request beyond the sheet's own bounds is #NUM!
 * rather than an allocation that takes the process down.
 *
 * A note for the writer, which does not affect anything here: xlsx stores
 * LAMBDA parameter names with an `_xlpm.` prefix, so `LAMBDA(x,x+1)` is written
 * `_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)`. Names are normalised through that prefix on
 * the way in, so a formula read straight from a file binds the same way one
 * typed by hand does.
 */

import { CellError, MAX_COLS, MAX_ROWS, type Scalar, isError } from '@mirrorz/core';
import { type Ast, type CallNode, Node } from '../ast.js';
import {
  ArgKind,
  type FunctionContext,
  type FunctionSpec,
  type ParamSpec,
  type Thunk,
  p,
} from '../registry.js';
import {
  type ArrayValue,
  type Value,
  arrayAt,
  compareScalars,
  excelAdd,
  isArray,
  isRef,
  makeArray,
  toBoolean,
  toExcelPrecision,
  toNumber,
} from '../value.js';

// ---------------------------------------------------------------------------
// Argument plumbing
// ---------------------------------------------------------------------------

/** Lazy arguments arrive as thunks cast through Value. */
function isThunk(v: unknown): v is Thunk {
  return typeof v === 'object' && v !== null && typeof (v as Thunk).evaluate === 'function';
}

function force(v: Value | undefined, ctx: FunctionContext): Value | undefined {
  return isThunk(v) ? ctx.force(v) : v;
}

/** A read-only rectangular view over whatever an argument turned out to be. */
interface Grid {
  readonly rows: number;
  readonly cols: number;
  at(row: number, col: number): Scalar;
}

const BLANK_GRID: Grid = { rows: 1, cols: 1, at: () => null };

function arrayGrid(a: ArrayValue): Grid {
  return { rows: a.rows, cols: a.cols, at: (r, c) => arrayAt(a, r, c) };
}

/** Materialise an argument. References are read through the context. */
function gridOf(v: Value | undefined, ctx: FunctionContext): Grid {
  if (v === undefined) return BLANK_GRID;
  if (isRef(v)) return arrayGrid(ctx.deref(v));
  if (isArray(v)) return arrayGrid(v);
  return { rows: 1, cols: 1, at: () => v };
}

/** Collapse a 1x1 result to its scalar; anything larger keeps its shape. */
function unwrapSingle(v: Value, ctx: FunctionContext): Value {
  if (isRef(v)) {
    const a = ctx.deref(v);
    return a.rows === 1 && a.cols === 1 ? (a.data[0] ?? null) : a;
  }
  if (isArray(v)) return v.rows === 1 && v.cols === 1 ? (v.data[0] ?? null) : v;
  return v;
}

/** An integer argument, truncated toward zero the way Excel truncates them. */
function intArg(v: Value | undefined, ctx: FunctionContext): number | CellError {
  const s = scalarOf(v, ctx);
  if (isError(s)) return s;
  const n = toNumber(s);
  if (isError(n)) return n;
  return Math.trunc(n);
}

/** The single value behind an argument that arrived in some other shape. */
function scalarOf(v: Value | undefined, ctx: FunctionContext): Scalar {
  if (v === undefined) return null;
  if (isRef(v)) {
    const a = ctx.deref(v);
    return a.data[0] ?? null;
  }
  if (isArray(v)) return v.data[0] ?? null;
  return v;
}

function boolArg(v: Value | undefined, ctx: FunctionContext, fallback: boolean): boolean | CellError {
  if (v === undefined) return fallback;
  const s = scalarOf(v, ctx);
  if (s === null) return fallback;
  return toBoolean(s);
}

/** Was this argument omitted, or written as an empty slot between commas? */
function omitted(v: Value | undefined): boolean {
  return v === undefined || v === null;
}

/**
 * The result an array-building function must not produce.
 * Excel has no empty array, so every caller turns this into #CALC!.
 */
const EMPTY: CellError = CellError.CALC;

/** Guard against an allocation that no sheet could hold. */
function checkSize(rows: number, cols: number): CellError | undefined {
  if (rows > MAX_ROWS || cols > MAX_COLS || rows * cols > MAX_ROWS) {
    return new CellError('#NUM!', 'the result is larger than the sheet');
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Lambdas: values, slots, and the AST patching that binds them
// ---------------------------------------------------------------------------

/** Strip the storage prefixes a name or function may carry in a file. */
function bareName(name: string): string {
  return name
    .replace(/^_xlfn\._xlws\./i, '')
    .replace(/^_xlfn\./i, '')
    .replace(/^_xlpm\./i, '')
    .toUpperCase();
}

/** The name a binder position introduces, or undefined if it is not a name. */
function binderName(ast: Ast): string | undefined {
  return ast.kind === Node.Name ? bareName(ast.name) : undefined;
}

/**
 * One place a bound name is mentioned.
 * `call` distinguishes `f` used as a value from `f(1)` used as an application,
 * which the parser reports as a CallNode and which needs a different rewrite.
 */
interface Slot {
  node: Ast;
  call: boolean;
}

/** A LAMBDA value. See the file header for why this is a CellError. */
class Lambda extends CellError {
  readonly params: readonly string[];
  readonly body: Thunk;
  private slotCache?: Map<string, Slot[]>;

  constructor(params: readonly string[], body: Thunk) {
    super('#CALC!', 'a LAMBDA that was never applied');
    this.params = params;
    this.body = body;
  }

  /**
   * Parameter slots, resolved once. This is the "resolve names to slots at
   * first use" step: after it, invoking the lambda touches only the nodes that
   * actually mention a parameter.
   */
  slots(): Map<string, Slot[]> {
    if (!this.slotCache) {
      const map = new Map<string, Slot[]>();
      for (const name of this.params) {
        const found: Slot[] = [];
        collectSlots(this.body.ast, name, found);
        map.set(name, found);
      }
      this.slotCache = map;
    }
    return this.slotCache;
  }
}

function isLambda(v: unknown): v is Lambda {
  return v instanceof Lambda;
}

/**
 * A lambda parameter that the caller left out.
 * It is an error value so that using an omitted parameter in arithmetic gives
 * #VALUE! as Excel does, and a distinct object so ISOMITTED can recognise it.
 */
const OMITTED = new CellError('#VALUE!', 'an omitted LAMBDA argument');

/**
 * Values injected into a patched AST, addressed by index.
 * The stash is truncated back to its mark when the patch is undone, so it
 * never grows across a recalculation.
 */
const STASH: Value[] = [];

const CONST_FN = '__MZ.CONST';
const APPLY_FN = '__MZ.APPLY';

function constNode(value: Value): CallNode {
  STASH.push(value);
  return {
    kind: Node.Call,
    name: CONST_FN,
    args: [{ kind: Node.Number, value: STASH.length - 1 }],
  };
}

interface Patch {
  node: Record<string, unknown>;
  saved: Record<string, unknown>;
}

/** Rewrite one slot in place, remembering enough to put it back. */
function patchSlot(slot: Slot, value: Value, patches: Patch[]): void {
  const node = slot.node as unknown as Record<string, unknown>;
  const saved = { ...node };
  const replacement: Ast = slot.call
    ? {
        kind: Node.Call,
        name: APPLY_FN,
        // An empty argument slot is how a caller omits a lambda parameter, and
        // it has to survive as something other than a blank.
        args: [
          constNode(value),
          ...(slot.node as CallNode).args.map((a) =>
            a.kind === Node.Missing ? (constNode(OMITTED) as Ast) : a,
          ),
        ],
      }
    : constNode(value);
  for (const key of Object.keys(node)) delete node[key];
  Object.assign(node, replacement);
  patches.push({ node, saved });
}

function undoPatches(patches: Patch[], stashMark: number): void {
  for (let i = patches.length - 1; i >= 0; i--) {
    const { node, saved } = patches[i]!;
    for (const key of Object.keys(node)) delete node[key];
    Object.assign(node, saved);
  }
  STASH.length = stashMark;
}

/**
 * Find every mention of `name` in a body, skipping the parts of it where an
 * inner LET or LAMBDA has rebound the same name.
 */
function collectSlots(node: Ast, name: string, out: Slot[]): void {
  switch (node.kind) {
    case Node.Name:
      if (node.sheet === undefined && node.book === undefined && bareName(node.name) === name) {
        out.push({ node, call: false });
      }
      return;

    case Node.Call: {
      const fn = bareName(node.name);
      if (fn === 'LET') return collectInLet(node, name, out);
      if (fn === 'LAMBDA') return collectInLambda(node, name, out);
      if (fn === name) out.push({ node, call: true });
      for (const a of node.args) collectSlots(a, name, out);
      return;
    }

    case Node.Unary:
    case Node.Postfix:
      collectSlots(node.operand, name, out);
      return;
    case Node.Binary:
      collectSlots(node.left, name, out);
      collectSlots(node.right, name, out);
      return;
    case Node.Paren:
      collectSlots(node.inner, name, out);
      return;
    case Node.Array:
      for (const row of node.rows) for (const cell of row) collectSlots(cell, name, out);
      return;
    case Node.ThreeD:
      collectSlots(node.inner, name, out);
      return;
    default:
      return;
  }
}

/** `LET(n1,v1,n2,v2,calc)`: v1 sees the outer scope, everything after n1 does not. */
function collectInLet(node: CallNode, name: string, out: Slot[]): void {
  const args = node.args;
  const n = args.length;
  if (n < 3 || n % 2 === 0) {
    // Malformed LET; it will report its own error, so just do not lose a slot.
    for (const a of args) collectSlots(a, name, out);
    return;
  }
  for (let i = 0; i + 1 < n; i += 2) {
    collectSlots(args[i + 1]!, name, out);
    if (binderName(args[i]!) === name) return;
  }
  collectSlots(args[n - 1]!, name, out);
}

/** `LAMBDA(p1,p2,calc)`: the parameters are binders, only the body is scanned. */
function collectInLambda(node: CallNode, name: string, out: Slot[]): void {
  const args = node.args;
  if (args.length === 0) return;
  for (let i = 0; i < args.length - 1; i++) {
    if (binderName(args[i]!) === name) return;
  }
  collectSlots(args[args.length - 1]!, name, out);
}

/** Apply a lambda to already-evaluated arguments. */
function invoke(fn: Lambda, argv: (Value | undefined)[], ctx: FunctionContext): Value {
  if (argv.length > fn.params.length) return CellError.VALUE;
  const slots = fn.slots();
  const patches: Patch[] = [];
  const mark = STASH.length;
  try {
    for (let i = 0; i < fn.params.length; i++) {
      const supplied = i < argv.length ? argv[i] : undefined;
      const bound = supplied === undefined ? OMITTED : supplied;
      for (const slot of slots.get(fn.params[i]!) ?? []) patchSlot(slot, bound, patches);
    }
    return ctx.force(fn.body);
  } finally {
    undoPatches(patches, mark);
  }
}

/** A lambda-valued argument, however it was written. */
function lambdaArg(v: Value | undefined, ctx: FunctionContext): Lambda | CellError {
  const forced = force(v, ctx);
  if (isLambda(forced)) return forced;
  if (isError(forced)) return forced;
  return CellError.VALUE;
}

/** The scalar a per-element lambda must have produced. */
function elementResult(v: Value, ctx: FunctionContext): Scalar {
  const single = unwrapSingle(v, ctx);
  if (isArray(single) || isRef(single)) return CellError.CALC;
  return single;
}

// ---------------------------------------------------------------------------
// LET and LAMBDA
// ---------------------------------------------------------------------------

/**
 * LET evaluates a binding only when something later mentions it, which is both
 * the documented "calculated once" behaviour for the names that are used and
 * the reason `LET(x,1/0,5)` is 5 rather than #DIV/0!.
 */
function letImpl(args: Value[], ctx: FunctionContext): Value {
  const n = args.length;
  if (n < 3 || n % 2 === 0) return CellError.VALUE;
  const body = args[n - 1];
  if (!isThunk(body)) return CellError.VALUE;

  const patches: Patch[] = [];
  const mark = STASH.length;
  const bound = new Set<string>();
  try {
    for (let i = 0; i + 1 < n - 1; i += 2) {
      const nameThunk = args[i];
      const valueThunk = args[i + 1];
      if (!isThunk(nameThunk) || !isThunk(valueThunk)) return CellError.VALUE;
      const name = binderName(nameThunk.ast);
      // A name that parsed as anything else - a cell reference, a number - is
      // not a legal LET name.
      if (name === undefined || bound.has(name)) return CellError.NAME;
      bound.add(name);

      // A name is visible to every later binding's value and to the body.
      const slots: Slot[] = [];
      for (let j = i + 3; j < n - 1; j += 2) {
        const later = args[j];
        if (isThunk(later)) collectSlots(later.ast, name, slots);
      }
      collectSlots(body.ast, name, slots);
      if (slots.length === 0) continue;

      const value = ctx.force(valueThunk);
      for (const slot of slots) patchSlot(slot, value, patches);
    }
    return ctx.force(body);
  } finally {
    undoPatches(patches, mark);
  }
}

function lambdaImpl(args: Value[], _ctx: FunctionContext): Value {
  const n = args.length;
  const body = args[n - 1];
  if (!isThunk(body)) return CellError.VALUE;
  const params: string[] = [];
  for (let i = 0; i < n - 1; i++) {
    const t = args[i];
    if (!isThunk(t)) return CellError.VALUE;
    const name = binderName(t.ast);
    if (name === undefined || params.includes(name)) return CellError.VALUE;
    params.push(name);
  }
  return new Lambda(params, body) as unknown as Value;
}

// ---------------------------------------------------------------------------
// Sorting order
// ---------------------------------------------------------------------------

/** Values sort first, then errors, then blanks - in both directions. */
function sortClass(v: Scalar): number {
  if (v === null) return 2;
  if (isError(v)) return 1;
  return 0;
}

function orderedCompare(a: Scalar, b: Scalar, descending: boolean): number {
  const ca = sortClass(a);
  const cb = sortClass(b);
  if (ca !== cb) return ca - cb;
  if (ca !== 0) return 0;
  const cmp = compareScalars(a, b);
  const n = isError(cmp) ? 0 : cmp;
  return descending ? -n : n;
}

/** Equality for UNIQUE: Excel's loose comparison, plus errors matching by code. */
function sameScalar(a: Scalar, b: Scalar): boolean {
  if (isError(a) || isError(b)) return isError(a) && isError(b) && a.code === b.code;
  const cmp = compareScalars(a, b);
  return !isError(cmp) && cmp === 0;
}

// ---------------------------------------------------------------------------
// FILTER, SORT, SORTBY, UNIQUE
// ---------------------------------------------------------------------------

function filterImpl(args: Value[], ctx: FunctionContext): Value {
  const source = gridOf(args[0], ctx);
  const include = gridOf(args[1], ctx);

  for (let r = 0; r < include.rows; r++) {
    for (let c = 0; c < include.cols; c++) {
      const v = include.at(r, c);
      if (isError(v)) return v;
    }
  }

  const byRow = include.cols === 1 && include.rows === source.rows;
  const byCol = include.rows === 1 && include.cols === source.cols;
  // A 1x1 source with a 1x1 mask is both; treating it as a row filter keeps the
  // result a 1x1 array rather than an accident of ordering.
  if (!byRow && !byCol) return CellError.VALUE;

  const keep: number[] = [];
  const count = byRow ? source.rows : source.cols;
  for (let i = 0; i < count; i++) {
    const flag = toBoolean(byRow ? include.at(i, 0) : include.at(0, i));
    if (isError(flag)) return flag;
    if (flag) keep.push(i);
  }

  if (keep.length === 0) return args[2] === undefined ? EMPTY : (args[2] as Value);

  if (byRow) {
    const data: Scalar[] = [];
    for (const r of keep) for (let c = 0; c < source.cols; c++) data.push(source.at(r, c));
    return makeArray(keep.length, source.cols, data);
  }
  const data: Scalar[] = [];
  for (let r = 0; r < source.rows; r++) for (const c of keep) data.push(source.at(r, c));
  return makeArray(source.rows, keep.length, data);
}

/** One sort key: which line to compare on, and in which direction. */
interface SortKey {
  index: number;
  descending: boolean;
}

function readSortKeys(
  indexArg: Value | undefined,
  orderArg: Value | undefined,
  ctx: FunctionContext,
  limit: number,
): SortKey[] | CellError {
  const indexes = omitted(indexArg) ? BLANK_GRID : gridOf(indexArg, ctx);
  const orders = omitted(orderArg) ? BLANK_GRID : gridOf(orderArg, ctx);
  const count = omitted(indexArg) ? 1 : indexes.rows * indexes.cols;
  const keys: SortKey[] = [];
  for (let i = 0; i < count; i++) {
    let index = 1;
    if (!omitted(indexArg)) {
      const raw = toNumber(indexes.at(Math.floor(i / indexes.cols), i % indexes.cols));
      if (isError(raw)) return raw;
      index = Math.trunc(raw);
    }
    if (index < 1 || index > limit) return CellError.VALUE;

    let order = 1;
    if (!omitted(orderArg)) {
      const cells = orders.rows * orders.cols;
      // A single order applies to every key; otherwise they pair up.
      const at = cells === 1 ? 0 : i;
      if (at >= cells) return CellError.VALUE;
      const raw = toNumber(orders.at(Math.floor(at / orders.cols), at % orders.cols));
      if (isError(raw)) return raw;
      order = Math.trunc(raw);
    }
    if (order !== 1 && order !== -1) return CellError.VALUE;
    keys.push({ index, descending: order === -1 });
  }
  return keys;
}

/**
 * A stable ordering of 0..n-1. The index tiebreak is what keeps it stable in
 * the descending direction too, where reversing the comparator would otherwise
 * reverse the order of equal elements.
 */
function stableOrder(n: number, compare: (a: number, b: number) => number): number[] {
  const order = Array.from({ length: n }, (_, i) => i);
  order.sort((a, b) => {
    const c = compare(a, b);
    return c !== 0 ? c : a - b;
  });
  return order;
}

function permute(g: Grid, order: number[], byCol: boolean): ArrayValue {
  const data: Scalar[] = [];
  if (byCol) {
    for (let r = 0; r < g.rows; r++) for (const c of order) data.push(g.at(r, c));
    return makeArray(g.rows, order.length, data);
  }
  for (const r of order) for (let c = 0; c < g.cols; c++) data.push(g.at(r, c));
  return makeArray(order.length, g.cols, data);
}

function sortImpl(args: Value[], ctx: FunctionContext): Value {
  const g = gridOf(args[0], ctx);
  const byCol = boolArg(args[3], ctx, false);
  if (isError(byCol)) return byCol;

  const limit = byCol ? g.rows : g.cols;
  const keys = readSortKeys(args[1], args[2], ctx, limit);
  if (isError(keys)) return keys;

  const lines = byCol ? g.cols : g.rows;
  const order = stableOrder(lines, (a, b) => {
    for (const key of keys) {
      const av = byCol ? g.at(key.index - 1, a) : g.at(a, key.index - 1);
      const bv = byCol ? g.at(key.index - 1, b) : g.at(b, key.index - 1);
      const c = orderedCompare(av, bv, key.descending);
      if (c !== 0) return c;
    }
    return 0;
  });
  return permute(g, order, byCol);
}

function sortByImpl(args: Value[], ctx: FunctionContext): Value {
  const g = gridOf(args[0], ctx);
  if (args.length < 2) return CellError.VALUE;

  interface ByKey {
    values: Grid;
    descending: boolean;
    byCol: boolean;
  }
  const keys: ByKey[] = [];
  for (let i = 1; i < args.length; i += 2) {
    const by = gridOf(args[i], ctx);
    const vertical = by.cols === 1 && by.rows === g.rows;
    const horizontal = by.rows === 1 && by.cols === g.cols;
    if (!vertical && !horizontal) return CellError.VALUE;
    let order = 1;
    if (!omitted(args[i + 1])) {
      const raw = intArg(args[i + 1], ctx);
      if (isError(raw)) return raw;
      order = raw;
    }
    if (order !== 1 && order !== -1) return CellError.VALUE;
    keys.push({ values: by, descending: order === -1, byCol: horizontal && !vertical });
  }

  const byCol = keys[0]!.byCol;
  // Mixing a row key with a column key would ask for two sorts at once.
  if (keys.some((k) => k.byCol !== byCol)) return CellError.VALUE;

  const lines = byCol ? g.cols : g.rows;
  const order = stableOrder(lines, (a, b) => {
    for (const key of keys) {
      const av = byCol ? key.values.at(0, a) : key.values.at(a, 0);
      const bv = byCol ? key.values.at(0, b) : key.values.at(b, 0);
      const c = orderedCompare(av, bv, key.descending);
      if (c !== 0) return c;
    }
    return 0;
  });
  return permute(g, order, byCol);
}

function uniqueImpl(args: Value[], ctx: FunctionContext): Value {
  const g = gridOf(args[0], ctx);
  const byCol = boolArg(args[1], ctx, false);
  if (isError(byCol)) return byCol;
  const exactlyOnce = boolArg(args[2], ctx, false);
  if (isError(exactlyOnce)) return exactlyOnce;

  const lines = byCol ? g.cols : g.rows;
  const width = byCol ? g.rows : g.cols;
  const at = (line: number, k: number): Scalar => (byCol ? g.at(k, line) : g.at(line, k));

  const same = (a: number, b: number): boolean => {
    for (let k = 0; k < width; k++) if (!sameScalar(at(a, k), at(b, k))) return false;
    return true;
  };

  // Quadratic on purpose: Excel's equality here is compareScalars, under which
  // a blank equals 0 and "a" equals "A", and no hash key reproduces that.
  const counts: number[] = new Array(lines).fill(0);
  const representatives: number[] = [];
  for (let i = 0; i < lines; i++) {
    let found = -1;
    for (const r of representatives) {
      if (same(i, r)) {
        found = r;
        break;
      }
    }
    if (found < 0) {
      representatives.push(i);
      counts[i] = 1;
    } else {
      counts[found] = (counts[found] ?? 0) + 1;
    }
  }

  const kept = representatives.filter((r) => !exactlyOnce || counts[r] === 1);
  if (kept.length === 0) return EMPTY;
  return permute(g, kept, byCol);
}

// ---------------------------------------------------------------------------
// SEQUENCE and RANDARRAY
// ---------------------------------------------------------------------------

function sequenceImpl(args: Value[], ctx: FunctionContext): Value {
  const rows = omitted(args[0]) ? 1 : intArg(args[0], ctx);
  if (isError(rows)) return rows;
  const cols = omitted(args[1]) ? 1 : intArg(args[1], ctx);
  if (isError(cols)) return cols;
  const start = omitted(args[2]) ? 1 : toNumber(scalarOf(args[2], ctx));
  if (isError(start)) return start;
  const step = omitted(args[3]) ? 1 : toNumber(scalarOf(args[3], ctx));
  if (isError(step)) return step;

  if (rows < 0 || cols < 0) return CellError.VALUE;
  if (rows === 0 || cols === 0) return EMPTY;
  const tooBig = checkSize(rows, cols);
  if (tooBig) return tooBig;

  const data: Scalar[] = new Array(rows * cols);
  for (let i = 0; i < rows * cols; i++) data[i] = toExcelPrecision(excelAdd(start, i * step));
  return makeArray(rows, cols, data);
}

function randArrayImpl(args: Value[], ctx: FunctionContext): Value {
  const rows = omitted(args[0]) ? 1 : intArg(args[0], ctx);
  if (isError(rows)) return rows;
  const cols = omitted(args[1]) ? 1 : intArg(args[1], ctx);
  if (isError(cols)) return cols;
  const min = omitted(args[2]) ? 0 : toNumber(scalarOf(args[2], ctx));
  if (isError(min)) return min;
  const max = omitted(args[3]) ? 1 : toNumber(scalarOf(args[3], ctx));
  if (isError(max)) return max;
  const whole = boolArg(args[4], ctx, false);
  if (isError(whole)) return whole;

  if (rows < 1 || cols < 1) return CellError.VALUE;
  if (min > max) return CellError.VALUE;
  if (whole && (!Number.isInteger(min) || !Number.isInteger(max))) return CellError.VALUE;
  const tooBig = checkSize(rows, cols);
  if (tooBig) return tooBig;

  const data: Scalar[] = new Array(rows * cols);
  for (let i = 0; i < rows * cols; i++) {
    data[i] = whole
      ? min + Math.floor(Math.random() * (max - min + 1))
      : min + Math.random() * (max - min);
  }
  return makeArray(rows, cols, data);
}

// ---------------------------------------------------------------------------
// The higher-order family
// ---------------------------------------------------------------------------

function byLineImpl(args: Value[], ctx: FunctionContext, byCol: boolean): Value {
  const source = force(args[0], ctx);
  if (isError(source)) return source;
  const g = gridOf(source, ctx);
  const fn = lambdaArg(args[1], ctx);
  if (!isLambda(fn)) return isError(fn) ? fn : CellError.VALUE;
  if (fn.params.length !== 1) return CellError.VALUE;

  const lines = byCol ? g.cols : g.rows;
  const width = byCol ? g.rows : g.cols;
  const out: Scalar[] = new Array(lines);
  for (let i = 0; i < lines; i++) {
    const cells: Scalar[] = new Array(width);
    for (let k = 0; k < width; k++) cells[k] = byCol ? g.at(k, i) : g.at(i, k);
    const line = byCol ? makeArray(width, 1, cells) : makeArray(1, width, cells);
    out[i] = elementResult(invoke(fn, [line], ctx), ctx);
  }
  return byCol ? makeArray(1, lines, out) : makeArray(lines, 1, out);
}

function mapImpl(args: Value[], ctx: FunctionContext): Value {
  const fn = lambdaArg(args[args.length - 1], ctx);
  if (!isLambda(fn)) return isError(fn) ? fn : CellError.VALUE;

  const grids: Grid[] = [];
  for (let i = 0; i < args.length - 1; i++) {
    const v = force(args[i], ctx);
    if (isError(v)) return v;
    grids.push(gridOf(v, ctx));
  }
  if (grids.length === 0 || fn.params.length !== grids.length) return CellError.VALUE;
  const first = grids[0]!;
  if (grids.some((g) => g.rows !== first.rows || g.cols !== first.cols)) return CellError.VALUE;

  const out: Scalar[] = new Array(first.rows * first.cols);
  for (let r = 0; r < first.rows; r++) {
    for (let c = 0; c < first.cols; c++) {
      const argv = grids.map((g) => g.at(r, c) as Value);
      out[r * first.cols + c] = elementResult(invoke(fn, argv, ctx), ctx);
    }
  }
  return makeArray(first.rows, first.cols, out);
}

function reduceImpl(args: Value[], ctx: FunctionContext, scan: boolean): Value {
  const initialThunk = args[0];
  const initial = isThunk(initialThunk) ? ctx.force(initialThunk) : (initialThunk ?? null);
  const arrayValue = force(args[1], ctx);
  if (isError(arrayValue)) return arrayValue;
  const g = gridOf(arrayValue, ctx);
  const fn = lambdaArg(args[2], ctx);
  if (!isLambda(fn)) return isError(fn) ? fn : CellError.VALUE;
  if (fn.params.length !== 2) return CellError.VALUE;

  let acc: Value = initial ?? null;
  const out: Scalar[] = new Array(g.rows * g.cols);
  for (let r = 0; r < g.rows; r++) {
    for (let c = 0; c < g.cols; c++) {
      acc = invoke(fn, [acc, g.at(r, c)], ctx);
      if (scan) out[r * g.cols + c] = elementResult(acc, ctx);
    }
  }
  if (!scan) return unwrapSingle(acc, ctx);
  return makeArray(g.rows, g.cols, out);
}

function makeArrayImpl(args: Value[], ctx: FunctionContext): Value {
  const rows = intArg(args[0], ctx);
  if (isError(rows)) return rows;
  const cols = intArg(args[1], ctx);
  if (isError(cols)) return cols;
  const fn = lambdaArg(args[2], ctx);
  if (!isLambda(fn)) return isError(fn) ? fn : CellError.VALUE;
  if (fn.params.length !== 2) return CellError.VALUE;

  if (rows < 1 || cols < 1) return CellError.VALUE;
  const tooBig = checkSize(rows, cols);
  if (tooBig) return tooBig;

  const out: Scalar[] = new Array(rows * cols);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      out[r * cols + c] = elementResult(invoke(fn, [r + 1, c + 1], ctx), ctx);
    }
  }
  return makeArray(rows, cols, out);
}

// ---------------------------------------------------------------------------
// Stacking and reshaping
// ---------------------------------------------------------------------------

function stackImpl(args: Value[], ctx: FunctionContext, vertical: boolean): Value {
  const grids = args.filter((a) => a !== undefined).map((a) => gridOf(a, ctx));
  if (grids.length === 0) return CellError.VALUE;

  const rows = vertical
    ? grids.reduce((n, g) => n + g.rows, 0)
    : grids.reduce((n, g) => Math.max(n, g.rows), 0);
  const cols = vertical
    ? grids.reduce((n, g) => Math.max(n, g.cols), 0)
    : grids.reduce((n, g) => n + g.cols, 0);

  // Short arguments are padded with #N/A, which is what makes a ragged VSTACK
  // visibly ragged rather than silently squared off.
  const data: Scalar[] = new Array(rows * cols).fill(CellError.NA);
  let offset = 0;
  for (const g of grids) {
    for (let r = 0; r < g.rows; r++) {
      for (let c = 0; c < g.cols; c++) {
        const rr = vertical ? offset + r : r;
        const cc = vertical ? c : offset + c;
        data[rr * cols + cc] = g.at(r, c);
      }
    }
    offset += vertical ? g.rows : g.cols;
  }
  return makeArray(rows, cols, data);
}

function toLineImpl(args: Value[], ctx: FunctionContext, row: boolean): Value {
  const g = gridOf(args[0], ctx);
  const ignore = omitted(args[1]) ? 0 : intArg(args[1], ctx);
  if (isError(ignore)) return ignore;
  const byCol = boolArg(args[2], ctx, false);
  if (isError(byCol)) return byCol;
  if (ignore < 0 || ignore > 3) return CellError.VALUE;

  const dropBlanks = ignore === 1 || ignore === 3;
  const dropErrors = ignore === 2 || ignore === 3;
  const out: Scalar[] = [];
  const push = (v: Scalar): void => {
    if (dropBlanks && v === null) return;
    if (dropErrors && isError(v)) return;
    out.push(v);
  };
  if (byCol) {
    for (let c = 0; c < g.cols; c++) for (let r = 0; r < g.rows; r++) push(g.at(r, c));
  } else {
    for (let r = 0; r < g.rows; r++) for (let c = 0; c < g.cols; c++) push(g.at(r, c));
  }

  if (out.length === 0) return EMPTY;
  return row ? makeArray(1, out.length, out) : makeArray(out.length, 1, out);
}

function wrapImpl(args: Value[], ctx: FunctionContext, rowsFirst: boolean): Value {
  const g = gridOf(args[0], ctx);
  if (g.rows !== 1 && g.cols !== 1) return CellError.VALUE;
  const count = intArg(args[1], ctx);
  if (isError(count)) return count;
  if (count < 1) return CellError.NUM;
  const pad = omitted(args[2]) ? CellError.NA : scalarOf(args[2], ctx);

  const flat: Scalar[] = [];
  for (let r = 0; r < g.rows; r++) for (let c = 0; c < g.cols; c++) flat.push(g.at(r, c));
  const lines = Math.ceil(flat.length / count);
  const tooBig = rowsFirst ? checkSize(lines, count) : checkSize(count, lines);
  if (tooBig) return tooBig;

  const rows = rowsFirst ? lines : count;
  const cols = rowsFirst ? count : lines;
  const data: Scalar[] = new Array(rows * cols).fill(pad);
  for (let i = 0; i < flat.length; i++) {
    // WRAPROWS fills a row at a time, WRAPCOLS a column at a time.
    const r = rowsFirst ? Math.floor(i / count) : i % count;
    const c = rowsFirst ? i % count : Math.floor(i / count);
    data[r * cols + c] = flat[i]!;
  }
  return makeArray(rows, cols, data);
}

/** The half-open slice TAKE and DROP agree on for one axis. */
function takeRange(size: number, n: number): [number, number] {
  if (n >= 0) return [0, Math.min(n, size)];
  return [Math.max(size + n, 0), size];
}

function dropRange(size: number, n: number): [number, number] {
  if (n >= 0) return [Math.min(n, size), size];
  return [0, Math.max(size + n, 0)];
}

function sliceImpl(args: Value[], ctx: FunctionContext, take: boolean): Value {
  const g = gridOf(args[0], ctx);
  const range = take ? takeRange : dropRange;

  let rowRange: [number, number] = [0, g.rows];
  if (!omitted(args[1])) {
    const n = intArg(args[1], ctx);
    if (isError(n)) return n;
    rowRange = range(g.rows, n);
  }
  let colRange: [number, number] = [0, g.cols];
  if (!omitted(args[2])) {
    const n = intArg(args[2], ctx);
    if (isError(n)) return n;
    colRange = range(g.cols, n);
  }

  const rows = rowRange[1] - rowRange[0];
  const cols = colRange[1] - colRange[0];
  if (rows <= 0 || cols <= 0) return EMPTY;

  const data: Scalar[] = new Array(rows * cols);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) data[r * cols + c] = g.at(rowRange[0] + r, colRange[0] + c);
  }
  return makeArray(rows, cols, data);
}

function chooseImpl(args: Value[], ctx: FunctionContext, rows: boolean): Value {
  const g = gridOf(args[0], ctx);
  const size = rows ? g.rows : g.cols;
  const picks: number[] = [];
  for (let i = 1; i < args.length; i++) {
    if (args[i] === undefined) continue;
    const sel = gridOf(args[i], ctx);
    for (let r = 0; r < sel.rows; r++) {
      for (let c = 0; c < sel.cols; c++) {
        const n = toNumber(sel.at(r, c));
        if (isError(n)) return n;
        const k = Math.trunc(n);
        // Negative indices count from the end; zero is never a line.
        const resolved = k > 0 ? k - 1 : k < 0 ? size + k : -1;
        if (resolved < 0 || resolved >= size) return CellError.VALUE;
        picks.push(resolved);
      }
    }
  }
  if (picks.length === 0) return CellError.VALUE;
  return permute(g, picks, !rows);
}

function expandImpl(args: Value[], ctx: FunctionContext): Value {
  const g = gridOf(args[0], ctx);
  let rows = g.rows;
  if (!omitted(args[1])) {
    const n = intArg(args[1], ctx);
    if (isError(n)) return n;
    rows = n;
  }
  let cols = g.cols;
  if (!omitted(args[2])) {
    const n = intArg(args[2], ctx);
    if (isError(n)) return n;
    cols = n;
  }
  const pad = omitted(args[3]) ? CellError.NA : scalarOf(args[3], ctx);

  // EXPAND only grows: asking for fewer rows than the array has is an error
  // rather than a silent truncation, which is what TAKE is for.
  if (rows < g.rows || cols < g.cols) return CellError.VALUE;
  const tooBig = checkSize(rows, cols);
  if (tooBig) return tooBig;

  const data: Scalar[] = new Array(rows * cols).fill(pad);
  for (let r = 0; r < g.rows; r++) {
    for (let c = 0; c < g.cols; c++) data[r * cols + c] = g.at(r, c);
  }
  return makeArray(rows, cols, data);
}

// ---------------------------------------------------------------------------
// The spill-reference pair
// ---------------------------------------------------------------------------

function anchorArrayImpl(args: Value[], _ctx: FunctionContext): Value {
  const v = args[0];
  if (v === undefined) return CellError.REF;
  // The spill rectangle an anchor owns is tracked by the dependency graph, not
  // here, so this is the anchor itself - the same answer the evaluator gives
  // for the `#` operator, and the right one whenever the anchor holds a scalar.
  return v;
}

function singleImpl(args: Value[], ctx: FunctionContext): Value {
  const v = args[0];
  if (v === undefined) return null;
  if (isRef(v)) {
    const { origin } = ctx;
    const single = v.startRow === v.endRow && v.startCol === v.endCol;
    if (single) return ctx.getScalar(v.sheet, v.startRow, v.startCol);
    if (v.sheet === origin.sheet) {
      if (v.startCol === v.endCol && origin.row >= v.startRow && origin.row <= v.endRow) {
        return ctx.getScalar(v.sheet, origin.row, v.startCol);
      }
      if (v.startRow === v.endRow && origin.col >= v.startCol && origin.col <= v.endCol) {
        return ctx.getScalar(v.sheet, v.startRow, origin.col);
      }
    }
    return CellError.VALUE;
  }
  if (isArray(v)) return v.data[0] ?? null;
  return v;
}

// ---------------------------------------------------------------------------
// The table
// ---------------------------------------------------------------------------

/** A parameter that must see error values instead of short-circuiting. */
function inspected(name: string, kind: ArgKind, optional = false): ParamSpec {
  return optional
    ? { name, kind, optional, errorTransparent: true }
    : { name, kind, errorTransparent: true };
}

function restInspected(name: string, kind: ArgKind): ParamSpec {
  return { name, kind, repeating: true, optional: true, errorTransparent: true };
}

export const DYNAMIC_ARRAY_FUNCTIONS: readonly FunctionSpec[] = [
  {
    name: 'FILTER',
    params: [p.array('array'), p.array('include'), p.any('if_empty', true)],
    impl: filterImpl,
    summary: 'Keeps the rows or columns of an array where a condition is TRUE.',
  },
  {
    name: 'SORT',
    params: [
      p.array('array'),
      p.array('sort_index', true),
      p.array('sort_order', true),
      p.scalar('by_col', true),
    ],
    impl: sortImpl,
    summary: 'Sorts an array by one or more of its columns or rows.',
  },
  {
    name: 'SORTBY',
    params: [p.array('array'), p.array('by_array1'), p.rest('by_array_or_order')],
    impl: sortByImpl,
    summary: 'Sorts an array by the values in a corresponding array.',
  },
  {
    name: 'UNIQUE',
    params: [p.array('array'), p.scalar('by_col', true), p.scalar('exactly_once', true)],
    impl: uniqueImpl,
    summary: 'The distinct rows or columns of an array.',
  },
  {
    name: 'SEQUENCE',
    params: [
      p.scalar('rows'),
      p.scalar('columns', true),
      p.scalar('start', true),
      p.scalar('step', true),
    ],
    impl: sequenceImpl,
    summary: 'An array of sequential numbers.',
  },
  {
    name: 'RANDARRAY',
    params: [
      p.scalar('rows', true),
      p.scalar('columns', true),
      p.scalar('min', true),
      p.scalar('max', true),
      p.scalar('whole_number', true),
    ],
    impl: randArrayImpl,
    // Microsoft documents RANDARRAY as recalculating on every pass; SEQUENCE,
    // deliberately, does not.
    volatile: true,
    summary: 'An array of random numbers.',
  },

  {
    name: 'LET',
    params: [p.lazy('name1'), p.lazy('value1'), p.rest('name_value_or_calculation', ArgKind.Lazy)],
    impl: letImpl,
    summary: 'Binds names to values for the duration of one calculation.',
  },
  {
    name: 'LAMBDA',
    params: [p.lazy('parameter_or_calculation'), p.rest('more', ArgKind.Lazy)],
    impl: lambdaImpl,
    summary: 'A reusable function value, applied by the higher-order functions.',
  },
  {
    name: 'ISOMITTED',
    params: [inspected('argument', ArgKind.Any)],
    impl: (args) => args[0] === OMITTED,
    summary: 'TRUE when a LAMBDA parameter was left out at the call site.',
  },
  {
    name: 'BYROW',
    params: [p.lazy('array'), p.lazy('lambda')],
    impl: (args, ctx) => byLineImpl(args, ctx, false),
    summary: 'Applies a LAMBDA to each row, returning a column of results.',
  },
  {
    name: 'BYCOL',
    params: [p.lazy('array'), p.lazy('lambda')],
    impl: (args, ctx) => byLineImpl(args, ctx, true),
    summary: 'Applies a LAMBDA to each column, returning a row of results.',
  },
  {
    name: 'MAP',
    params: [p.lazy('array1'), p.lazy('lambda'), p.rest('more', ArgKind.Lazy)],
    impl: mapImpl,
    summary: 'Applies a LAMBDA element-wise across one or more arrays.',
  },
  {
    name: 'REDUCE',
    params: [p.lazy('initial_value'), p.lazy('array'), p.lazy('lambda')],
    impl: (args, ctx) => reduceImpl(args, ctx, false),
    summary: 'Folds an array into a single value with a LAMBDA.',
  },
  {
    name: 'SCAN',
    params: [p.lazy('initial_value'), p.lazy('array'), p.lazy('lambda')],
    impl: (args, ctx) => reduceImpl(args, ctx, true),
    summary: 'Like REDUCE, but keeps every intermediate value.',
  },
  {
    name: 'MAKEARRAY',
    params: [p.scalar('rows'), p.scalar('columns'), p.lazy('lambda')],
    impl: makeArrayImpl,
    summary: 'Builds an array from a LAMBDA of row and column index.',
  },

  {
    name: 'VSTACK',
    params: [restInspected('array', ArgKind.Array)],
    impl: (args, ctx) => stackImpl(args, ctx, true),
    summary: 'Stacks arrays vertically.',
  },
  {
    name: 'HSTACK',
    params: [restInspected('array', ArgKind.Array)],
    impl: (args, ctx) => stackImpl(args, ctx, false),
    summary: 'Stacks arrays horizontally.',
  },
  {
    name: 'TOROW',
    params: [
      inspected('array', ArgKind.Array),
      p.scalar('ignore', true),
      p.scalar('scan_by_column', true),
    ],
    impl: (args, ctx) => toLineImpl(args, ctx, true),
    summary: 'Flattens an array into a single row.',
  },
  {
    name: 'TOCOL',
    params: [
      inspected('array', ArgKind.Array),
      p.scalar('ignore', true),
      p.scalar('scan_by_column', true),
    ],
    impl: (args, ctx) => toLineImpl(args, ctx, false),
    summary: 'Flattens an array into a single column.',
  },
  {
    name: 'WRAPROWS',
    params: [inspected('vector', ArgKind.Array), p.scalar('wrap_count'), p.any('pad_with', true)],
    impl: (args, ctx) => wrapImpl(args, ctx, true),
    summary: 'Wraps a vector into rows of a given width.',
  },
  {
    name: 'WRAPCOLS',
    params: [inspected('vector', ArgKind.Array), p.scalar('wrap_count'), p.any('pad_with', true)],
    impl: (args, ctx) => wrapImpl(args, ctx, false),
    summary: 'Wraps a vector into columns of a given height.',
  },
  {
    name: 'TAKE',
    params: [inspected('array', ArgKind.Array), p.any('rows', true), p.any('columns', true)],
    impl: (args, ctx) => sliceImpl(args, ctx, true),
    summary: 'The first or last rows and columns of an array.',
  },
  {
    name: 'DROP',
    params: [inspected('array', ArgKind.Array), p.any('rows', true), p.any('columns', true)],
    impl: (args, ctx) => sliceImpl(args, ctx, false),
    summary: 'An array with rows and columns removed from its edges.',
  },
  {
    name: 'CHOOSEROWS',
    params: [inspected('array', ArgKind.Array), p.array('row_num1'), p.rest('more')],
    impl: (args, ctx) => chooseImpl(args, ctx, true),
    summary: 'The rows of an array in a chosen order.',
  },
  {
    name: 'CHOOSECOLS',
    params: [inspected('array', ArgKind.Array), p.array('col_num1'), p.rest('more')],
    impl: (args, ctx) => chooseImpl(args, ctx, false),
    summary: 'The columns of an array in a chosen order.',
  },
  {
    name: 'EXPAND',
    params: [
      inspected('array', ArgKind.Array),
      p.any('rows', true),
      p.any('columns', true),
      p.any('pad_with', true),
    ],
    impl: expandImpl,
    summary: 'Pads an array out to a larger size.',
  },

  {
    name: 'ANCHORARRAY',
    params: [p.ref('reference')],
    impl: anchorArrayImpl,
    structural: true,
    summary: 'The spill range anchored at a cell, as written `A1#`.',
  },
  {
    name: 'SINGLE',
    params: [inspected('value', ArgKind.Any)],
    impl: singleImpl,
    summary: 'Implicit intersection, as written `@`.',
  },

  // Internal, and not a worksheet function: these two exist so that a bound
  // name can be substituted into a formula the real evaluator then evaluates.
  // See the file header. They are registered rather than hidden because
  // dispatch goes through the registry by name.
  {
    name: CONST_FN,
    params: [p.scalar('slot')],
    impl: (args) => {
      const i = args[0];
      return typeof i === 'number' ? (STASH[i] ?? null) : CellError.VALUE;
    },
    summary: 'Internal: a value bound by LET or LAMBDA.',
  },
  {
    name: APPLY_FN,
    params: [inspected('lambda', ArgKind.Any), restInspected('argument', ArgKind.Any)],
    impl: (args, ctx) => {
      const fn = args[0];
      if (!isLambda(fn)) return isError(fn) ? fn : CellError.VALUE;
      return invoke(fn, args.slice(1), ctx);
    },
    summary: 'Internal: applies a LAMBDA bound to a name.',
  },
];
