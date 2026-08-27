/**
 * Logical and information functions.
 *
 * Five decisions shape this file.
 *
 * First, laziness. IF, IFS, SWITCH, IFERROR and IFNA take their branches as
 * `ArgKind.Lazy` thunks and force only the branch actually taken, so
 * `IF(FALSE,1/0,"ok")` is "ok" rather than #DIV/0!. Every other design here is
 * negotiable; that one is the reason the whole ArgKind.Lazy machinery exists.
 *
 * Second, AND, OR and XOR are deliberately NOT lazy. Excel evaluates every
 * argument of AND and OR - `=AND(FALSE,1/0)` is #DIV/0!, which is exactly why
 * spreadsheet authors nest IFs when they need a guard. Taking those arguments
 * eagerly as `ArgKind.Array` buys the correct error behaviour from the
 * evaluator's own short-circuit, and it also gives us the range/argument
 * asymmetry for free: a value that arrived as an ArrayValue came from a range or
 * an array constant and text and blanks in it are ignored, while a bare Scalar
 * was typed into the formula and must coerce or fail.
 *
 * Third, the IS predicates are error-transparent and answer FALSE rather than
 * propagating. `ISNUMBER(SEARCH(...))` is the most widely used idiom in the
 * spreadsheet world and it only works because ISNUMBER sees #VALUE! and says
 * FALSE. ISEVEN and ISODD are the exceptions: they are numeric functions
 * wearing an IS prefix, and they propagate like ABS does.
 *
 * Fourth, honesty about state we do not have. FunctionContext exposes cell
 * values, the date system and a fixed now - no formula text, no column widths,
 * no sheet order, no workbook path. So ISFORMULA, SHEET, the sheet-count form
 * of SHEETS, INFO, and the formatting info_types of CELL return an error saying
 * the answer is unavailable rather than a plausible-looking invention. A wrong
 * FALSE from ISFORMULA is worse than an error, because it is wrong precisely
 * for the cells the function exists to find.
 *
 * Fifth, volatility is a predicate. CELL is not volatile when it is asked for
 * the address, row, column, contents or type of an explicit reference - all of
 * those are ordinary dependencies the graph already tracks. It becomes volatile
 * when the reference is omitted (the answer then depends on which cell is
 * current) or when the info_type is not a literal we can inspect. INFO is
 * volatile only for the runtime types; the release and system strings never
 * change during a session.
 *
 * T is not here although it is an information function: Microsoft files it under
 * Text and text.ts already registers it, and registering it twice would abort
 * the whole registry.
 */

import {
  CellError,
  type Scalar,
  ValueType,
  colToName,
  isError,
  quoteSheetName,
  typeOf,
} from '@mirrorz/core';
import { type Ast, Node } from '../ast.js';
import {
  ArgKind,
  type FunctionContext,
  type FunctionSpec,
  type ParamSpec,
  type Thunk,
  p,
} from '../registry.js';
import {
  type RefValue,
  type Value,
  compareScalars,
  isArray,
  isRef,
  makeArray,
  toBoolean,
  toNumber,
} from '../value.js';

// ---------------------------------------------------------------------------
// Argument plumbing
// ---------------------------------------------------------------------------

/** A parameter that must see error values instead of short-circuiting. */
function inspected(name: string, kind: ArgKind = ArgKind.Scalar, optional = false): ParamSpec {
  return optional
    ? { name, kind, optional, errorTransparent: true }
    : { name, kind, errorTransparent: true };
}

/**
 * Lazy arguments arrive as thunks cast through Value, so every lazy parameter
 * has to be unwrapped before use. An omitted optional parameter arrives as
 * undefined and never as a thunk.
 */
function isThunk(v: unknown): v is Thunk {
  return typeof v === 'object' && v !== null && typeof (v as Thunk).evaluate === 'function';
}

function force(arg: Value | undefined, ctx: FunctionContext): Value | undefined {
  return isThunk(arg) ? ctx.force(arg) : arg;
}

/**
 * The single cell a multi-cell reference stands for in a scalar context.
 * Mirrors the evaluator's implicit intersection so that ArgKind.Any and
 * ArgKind.Reference parameters behave like ArgKind.Scalar ones when they are
 * handed a range.
 */
function intersect(ref: RefValue, ctx: FunctionContext): { row: number; col: number } {
  const { origin } = ctx;
  if (
    ref.startCol === ref.endCol &&
    origin.sheet === ref.sheet &&
    origin.row >= ref.startRow &&
    origin.row <= ref.endRow
  ) {
    return { row: origin.row, col: ref.startCol };
  }
  if (
    ref.startRow === ref.endRow &&
    origin.sheet === ref.sheet &&
    origin.col >= ref.startCol &&
    origin.col <= ref.endCol
  ) {
    return { row: ref.startRow, col: origin.col };
  }
  return { row: ref.startRow, col: ref.startCol };
}

/** Reduce any argument shape to the one scalar a predicate should judge. */
function resolveScalar(v: Value | undefined, ctx: FunctionContext): Scalar {
  if (v === undefined) return null;
  if (isRef(v)) {
    const cell = intersect(v, ctx);
    return ctx.getScalar(v.sheet, cell.row, cell.col);
  }
  if (isArray(v)) return v.data[0] ?? null;
  return v;
}

/**
 * What a taken branch of IF, IFS, SWITCH, IFERROR or IFNA yields.
 *
 * A branch that lands on an empty cell is 0, not blank: `=IF(TRUE,Z1)` with Z1
 * empty shows 0 in Excel. A multi-cell reference keeps its shape, because
 * `=SUM(IF(x,A1:A3,B1:B3))` depends on it.
 */
function branchResult(v: Value | undefined, ctx: FunctionContext): Value {
  if (v === undefined) return 0;
  if (isRef(v)) {
    if (v.startRow !== v.endRow || v.startCol !== v.endCol) return v;
    const cell = ctx.getScalar(v.sheet, v.startRow, v.startCol);
    return cell === null ? 0 : cell;
  }
  return v === null ? 0 : v;
}

// ---------------------------------------------------------------------------
// Logical
// ---------------------------------------------------------------------------

const IF: FunctionSpec = {
  name: 'IF',
  params: [p.scalar('logical_test'), p.lazy('value_if_true'), p.lazy('value_if_false', true)],
  summary: 'One of two values, depending on a condition.',
  impl: (args, ctx) => {
    const test = toBoolean(resolveScalar(args[0], ctx));
    if (isError(test)) return test;
    if (!test && args[2] === undefined) return false;
    return branchResult(force(test ? args[1] : args[2], ctx), ctx);
  },
};

const IFS: FunctionSpec = {
  name: 'IFS',
  params: [
    p.lazy('logical_test1'),
    p.lazy('value_if_true1'),
    p.rest('pairs', ArgKind.Lazy),
  ],
  summary: 'The value paired with the first condition that is TRUE.',
  impl: (args, ctx) => {
    for (let i = 0; i + 1 < args.length; i += 2) {
      const condition = force(args[i], ctx);
      if (isError(condition)) return condition;
      const test = toBoolean(resolveScalar(condition, ctx));
      if (isError(test)) return test;
      if (test) return branchResult(force(args[i + 1], ctx), ctx);
    }
    // A dangling condition with no value, and the all-false case, are both the
    // same #N/A in Excel.
    return CellError.NA;
  },
};

const SWITCH: FunctionSpec = {
  name: 'SWITCH',
  params: [
    p.lazy('expression'),
    p.lazy('value1'),
    p.lazy('result1'),
    p.rest('more', ArgKind.Lazy),
  ],
  summary: 'The result matching the first value equal to an expression.',
  impl: (args, ctx) => {
    const subject = force(args[0], ctx);
    if (isError(subject)) return subject;
    const key = resolveScalar(subject, ctx);

    let i = 1;
    for (; i + 1 < args.length; i += 2) {
      const candidate = force(args[i], ctx);
      if (isError(candidate)) return candidate;
      const cmp = compareScalars(key, resolveScalar(candidate, ctx));
      if (isError(cmp)) return cmp;
      if (cmp === 0) return branchResult(force(args[i + 1], ctx), ctx);
    }
    // An odd argument left over is the default; without one, nothing matched.
    if (i < args.length) return branchResult(force(args[i], ctx), ctx);
    return CellError.NA;
  },
};

/** TRUE and FALSE counted across every argument, with Excel's two rules. */
interface Tally {
  seen: number;
  trues: number;
}

function tally(args: (Value | undefined)[], _ctx: FunctionContext): Tally | CellError {
  let seen = 0;
  let trues = 0;
  for (const arg of args) {
    if (arg === undefined) continue;

    if (isArray(arg)) {
      // Inside a range or an array, text and blanks are simply not logical
      // values and are skipped; an error still propagates.
      for (const cell of arg.data) {
        if (isError(cell)) return cell;
        if (cell === null || typeof cell === 'string') continue;
        seen++;
        if (typeof cell === 'boolean' ? cell : cell !== 0) trues++;
      }
      continue;
    }
    if (isRef(arg)) continue;
    if (isError(arg)) return arg;
    // A direct argument must coerce: AND("TRUE") is TRUE and AND("x") is #VALUE!.
    if (arg === null) continue;
    const b = toBoolean(arg);
    if (isError(b)) return b;
    seen++;
    if (b) trues++;
  }
  if (seen === 0) return CellError.VALUE;
  return { seen, trues };
}

const AND: FunctionSpec = {
  name: 'AND',
  params: [p.array('logical1'), p.rest('logical')],
  summary: 'TRUE when every argument is TRUE.',
  impl: (args, ctx) => {
    const t = tally(args, ctx);
    return isError(t) ? t : t.trues === t.seen;
  },
};

const OR: FunctionSpec = {
  name: 'OR',
  params: [p.array('logical1'), p.rest('logical')],
  summary: 'TRUE when any argument is TRUE.',
  impl: (args, ctx) => {
    const t = tally(args, ctx);
    return isError(t) ? t : t.trues > 0;
  },
};

const XOR: FunctionSpec = {
  name: 'XOR',
  params: [p.array('logical1'), p.rest('logical')],
  summary: 'TRUE when an odd number of arguments are TRUE.',
  impl: (args, ctx) => {
    const t = tally(args, ctx);
    return isError(t) ? t : t.trues % 2 === 1;
  },
};

const NOT: FunctionSpec = {
  name: 'NOT',
  params: [p.scalar('logical')],
  broadcast: true,
  summary: 'The opposite of a logical value.',
  impl: (args, ctx) => {
    const b = toBoolean(resolveScalar(args[0], ctx));
    return isError(b) ? b : !b;
  },
};

const TRUE_FN: FunctionSpec = {
  name: 'TRUE',
  params: [],
  summary: 'The logical value TRUE.',
  impl: () => true,
};

const FALSE_FN: FunctionSpec = {
  name: 'FALSE',
  params: [],
  summary: 'The logical value FALSE.',
  impl: () => false,
};

/** IFERROR and IFNA differ only in which errors they catch. */
function trapping(name: string, caught: (e: CellError) => boolean): FunctionSpec {
  return {
    name,
    params: [inspected('value', ArgKind.Any), p.lazy('value_if_error')],
    summary:
      name === 'IFERROR'
        ? 'The value, or an alternative when it is an error.'
        : 'The value, or an alternative when it is #N/A.',
    impl: (args, ctx) => {
      const value = args[0];
      if (isError(value)) {
        if (!caught(value)) return value;
        return branchResult(force(args[1], ctx), ctx);
      }
      // An array is repaired element-wise, which is what IFERROR does to a
      // dynamic-array result; the fallback is forced once, and only if needed.
      if (isArray(value) && value.data.some((c) => isError(c) && caught(c))) {
        const fallback = resolveScalar(force(args[1], ctx), ctx);
        const data = value.data.map((c) => (isError(c) && caught(c) ? fallback : c));
        return makeArray(value.rows, value.cols, data);
      }
      // A trapped value that is not an error still goes through the branch
      // rules, so a reference to an empty cell reads as 0.
      return branchResult(value, ctx);
    },
  };
}

const IFERROR = trapping('IFERROR', () => true);
const IFNA = trapping('IFNA', (e) => e.code === '#N/A');

// ---------------------------------------------------------------------------
// Information: the IS predicates
// ---------------------------------------------------------------------------

/** An IS predicate over one scalar, error-transparent so it can answer FALSE. */
function predicate(
  name: string,
  summary: string,
  test: (v: Scalar) => boolean,
): FunctionSpec {
  return {
    name,
    params: [inspected('value', ArgKind.Any)],
    broadcast: true,
    summary,
    impl: (args, ctx) => test(resolveScalar(args[0], ctx)),
  };
}

const ISBLANK = predicate(
  'ISBLANK',
  'TRUE when the value is an empty cell.',
  (v) => v === null,
);
const ISERROR = predicate('ISERROR', 'TRUE when the value is any error.', isError);
const ISERR = predicate(
  'ISERR',
  'TRUE when the value is an error other than #N/A.',
  (v) => isError(v) && v.code !== '#N/A',
);
const ISNA = predicate(
  'ISNA',
  'TRUE when the value is #N/A.',
  (v) => isError(v) && v.code === '#N/A',
);
const ISNUMBER = predicate(
  'ISNUMBER',
  'TRUE when the value is a number.',
  (v) => typeof v === 'number',
);
const ISTEXT = predicate('ISTEXT', 'TRUE when the value is text.', (v) => typeof v === 'string');
const ISNONTEXT = predicate(
  'ISNONTEXT',
  'TRUE when the value is not text.',
  (v) => typeof v !== 'string',
);
const ISLOGICAL = predicate(
  'ISLOGICAL',
  'TRUE when the value is TRUE or FALSE.',
  (v) => typeof v === 'boolean',
);

const ISREF: FunctionSpec = {
  name: 'ISREF',
  params: [inspected('value', ArgKind.Reference)],
  summary: 'TRUE when the value is a reference.',
  // The argument must not be dereferenced, or every reference would arrive as
  // its contents and the answer would always be FALSE.
  impl: (args) => isRef(args[0]),
};

/** ISEVEN and ISODD are numeric, so unlike the other IS functions they fail. */
function parity(name: string, wanted: 0 | 1): FunctionSpec {
  return {
    name,
    params: [p.scalar('number')],
    broadcast: true,
    summary: wanted === 0 ? 'TRUE when the number is even.' : 'TRUE when the number is odd.',
    impl: (args, ctx) => {
      const n = toNumber(resolveScalar(args[0], ctx));
      if (isError(n)) return n;
      // Excel truncates towards zero before testing parity.
      return Math.abs(Math.trunc(n)) % 2 === wanted;
    },
  };
}

const ISEVEN = parity('ISEVEN', 0);
const ISODD = parity('ISODD', 1);

const ISFORMULA: FunctionSpec = {
  name: 'ISFORMULA',
  params: [p.ref('reference')],
  structural: true,
  summary: 'TRUE when the reference points at a cell containing a formula.',
  impl: (args) => {
    if (!isRef(args[0])) return CellError.VALUE;
    // FunctionContext exposes cell values but not their formulas. Answering
    // FALSE would be wrong for exactly the cells this function exists to find,
    // so the honest answer is that it is unavailable.
    return new CellError('#N/A', 'ISFORMULA needs cell formulas, which the engine does not expose');
  },
};

// ---------------------------------------------------------------------------
// Information: conversions and types
// ---------------------------------------------------------------------------

const N: FunctionSpec = {
  name: 'N',
  params: [inspected('value')],
  broadcast: true,
  summary: 'A value converted to a number.',
  impl: (args, ctx) => {
    const v = resolveScalar(args[0], ctx);
    if (isError(v)) return v;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    // Text is 0 rather than parsed: N("3") is 0, unlike "3"+0.
    return 0;
  },
};

const TYPE: FunctionSpec = {
  name: 'TYPE',
  params: [inspected('value', ArgKind.Any)],
  summary: "The code for a value's type: 1 number, 2 text, 4 logical, 16 error, 64 array.",
  impl: (args, ctx) => {
    const v = args[0];
    if (isArray(v)) return ValueType.Array;
    return typeOf(resolveScalar(v, ctx));
  },
};

/** ERROR.TYPE's numbering, in the order Microsoft documents it. */
const ERROR_ORDINALS: Readonly<Record<string, number>> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
  '#GETTING_DATA': 8,
  '#SPILL!': 9,
  '#CONNECT!': 10,
  '#BLOCKED!': 11,
  '#UNKNOWN!': 12,
  '#FIELD!': 13,
  '#CALC!': 14,
};

const ERROR_TYPE: FunctionSpec = {
  name: 'ERROR.TYPE',
  params: [inspected('error_val')],
  broadcast: true,
  summary: 'The number identifying an error value.',
  impl: (args, ctx) => {
    const v = resolveScalar(args[0], ctx);
    if (!isError(v)) return CellError.NA;
    return ERROR_ORDINALS[v.code] ?? CellError.NA;
  },
};

const NA: FunctionSpec = {
  name: 'NA',
  params: [],
  summary: 'The error value #N/A.',
  impl: () => CellError.NA,
};

// ---------------------------------------------------------------------------
// Information: cell, workbook and environment
// ---------------------------------------------------------------------------

/** CELL info_types answerable from a reference alone. */
const GEOMETRIC_INFO = new Set(['address', 'col', 'row', 'contents', 'type']);

const CELL: FunctionSpec = {
  name: 'CELL',
  params: [p.scalar('info_type'), p.ref('reference', true)],
  structural: true,
  threadSafe: false,
  volatile: (args: Ast[]) => {
    // Only the geometric forms with an explicit reference are ordinary
    // dependencies; everything else follows state the graph cannot see.
    if (args.length < 2) return true;
    const type = args[0];
    if (!type || type.kind !== Node.Text) return true;
    return !GEOMETRIC_INFO.has(type.value.toLowerCase());
  },
  summary: 'Information about the position, contents or formatting of a cell.',
  impl: (args, ctx) => {
    const info = args[0];
    if (typeof info !== 'string') return CellError.VALUE;
    const type = info.toLowerCase();

    const ref = args[1];
    if (args[1] !== undefined && !isRef(ref)) return CellError.VALUE;
    // With no reference Excel reports on the last cell changed, which needs an
    // editing session. The formula's own cell is the only stable stand-in.
    const sheet = isRef(ref) ? ref.sheet : ctx.origin.sheet;
    const row = isRef(ref) ? ref.startRow : ctx.origin.row;
    const col = isRef(ref) ? ref.startCol : ctx.origin.col;

    switch (type) {
      case 'address': {
        const local = `$${colToName(col)}$${row + 1}`;
        return sheet === ctx.origin.sheet ? local : `${quoteSheetName(sheet)}!${local}`;
      }
      case 'col':
        return col + 1;
      case 'row':
        return row + 1;
      case 'contents':
        return ctx.getScalar(sheet, row, col);
      case 'type': {
        const v = ctx.getScalar(sheet, row, col);
        return v === null ? 'b' : typeof v === 'string' ? 'l' : 'v';
      }
      default:
        // width, format, color, filename, parentheses, prefix and protect all
        // need column metrics, number formats or the workbook path, none of
        // which reach a worksheet function here.
        return new CellError('#VALUE!', `CELL("${type}") needs state the engine does not expose`);
    }
  },
};

/** INFO types whose answer changes during a session. */
const LIVE_INFO = new Set(['numfile', 'recalc', 'memavail', 'memused', 'totmem', 'origin', 'directory']);

const INFO: FunctionSpec = {
  name: 'INFO',
  params: [p.scalar('type_text')],
  threadSafe: false,
  volatile: (args: Ast[]) => {
    const type = args[0];
    if (!type || type.kind !== Node.Text) return true;
    return LIVE_INFO.has(type.value.toLowerCase());
  },
  summary: 'Information about the current operating environment.',
  impl: (args) => {
    const info = args[0];
    if (typeof info !== 'string') return CellError.VALUE;
    // Every INFO type reports on the host application - its version, its
    // memory, its current directory, its calculation mode. Reporting Node's
    // equivalents would be a fabrication dressed as an Excel answer.
    return new CellError(
      '#VALUE!',
      `INFO("${info.toLowerCase()}") needs application state the engine does not expose`,
    );
  },
};

const SHEET: FunctionSpec = {
  name: 'SHEET',
  params: [p.ref('value', true)],
  structural: true,
  summary: 'The sheet number of a reference.',
  impl: (args, ctx) => {
    if (args[0] !== undefined && !isRef(args[0]) && typeof args[0] !== 'string') {
      return CellError.VALUE;
    }
    if (typeof args[0] === 'string' && !ctx.hasSheet(args[0])) return CellError.NA;
    // A sheet number is a position in tab order, and FunctionContext offers no
    // way to enumerate the tabs.
    return new CellError('#N/A', 'SHEET needs the workbook tab order, which the engine does not expose');
  },
};

const SHEETS: FunctionSpec = {
  name: 'SHEETS',
  params: [p.ref('reference', true)],
  structural: true,
  summary: 'The number of sheets a reference spans.',
  impl: (args) => {
    if (args[0] === undefined) {
      return new CellError('#N/A', 'SHEETS needs the workbook tab count, which the engine does not expose');
    }
    // Every reference this engine builds lives on exactly one sheet; a 3-D
    // reference has already collapsed into an array by the time it arrives.
    if (!isRef(args[0])) return CellError.REF;
    return 1;
  },
};

export const LOGICAL_FUNCTIONS: readonly FunctionSpec[] = [
  IF,
  IFS,
  IFERROR,
  IFNA,
  AND,
  OR,
  NOT,
  XOR,
  SWITCH,
  TRUE_FN,
  FALSE_FN,
  ISBLANK,
  ISERROR,
  ISERR,
  ISNA,
  ISNUMBER,
  ISTEXT,
  ISNONTEXT,
  ISLOGICAL,
  ISREF,
  ISEVEN,
  ISODD,
  ISFORMULA,
  N,
  TYPE,
  ERROR_TYPE,
  NA,
  CELL,
  INFO,
  SHEET,
  SHEETS,
];
