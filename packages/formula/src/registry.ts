/**
 * Function registry and metadata.
 *
 * Every worksheet function is described before it is implemented, because the
 * evaluator's behaviour at the argument boundary is driven entirely by this
 * metadata: whether an argument is dereferenced from a reference to a value,
 * whether it is evaluated eagerly or handed over as a thunk, whether an error
 * short-circuits, and whether a scalar function broadcasts over an array.
 * Putting those decisions in data rather than in each function body is what
 * keeps two hundred implementations consistent.
 *
 * The volatility flag is a predicate rather than a boolean because Microsoft
 * documents several functions as volatile only depending on their arguments,
 * and because the folklore list is wrong in both directions - INDEX, ROWS,
 * COLUMNS and AREAS are widely described as volatile and are not. Over-marking
 * is not a harmless conservatism: a volatile function drags its entire
 * dependent closure into every recalculation.
 */

import type { CellError, Scalar } from '@mirrorz/core';
import type { Ast } from './ast.js';
import type { ArrayValue, RefValue, Value } from './value.js';

/** How an argument reaches the implementation. */
export const enum ArgKind {
  /** Dereferenced to a single scalar, with implicit intersection if needed. */
  Scalar,
  /** Dereferenced to a rectangular block of values. */
  Array,
  /** Passed as a reference, undereferenced. ROW, COLUMN, OFFSET need this. */
  Reference,
  /** Passed as an unevaluated thunk, for short-circuiting and lambdas. */
  Lazy,
  /** Whatever the expression produced, untouched. */
  Any,
}

export interface ParamSpec {
  name: string;
  kind: ArgKind;
  /** Optional parameters may be omitted or left empty. */
  optional?: boolean;
  /** This parameter absorbs the rest of the arguments. */
  repeating?: boolean;
  /**
   * Errors in this argument do NOT short-circuit the call. IFERROR's first
   * argument is the obvious case, but ISERROR, ISNA, N and the aggregate
   * functions need it too.
   */
  errorTransparent?: boolean;
}

export interface FunctionContext {
  /** The cell being evaluated, for implicit intersection and ROW/COLUMN. */
  readonly origin: { sheet: string; row: number; col: number };
  /** Read a single cell. */
  getScalar(sheet: string, row: number, col: number): Scalar;
  /** Materialise a reference into an array of values. */
  deref(ref: RefValue): ArrayValue;
  /** Iterate only the cells that actually exist inside a reference. */
  iterate(ref: RefValue): Iterable<{ row: number; col: number; value: Scalar }>;
  /** Resolve a sheet name, or undefined when it does not exist. */
  hasSheet(name: string): boolean;
  /** The workbook's date system, needed by every date function. */
  readonly dateSystem: 1900 | 1904;
  /** A fixed "now" for the whole recalculation, so NOW() is consistent. */
  readonly now: number;
  /** Evaluate a thunk argument. */
  force(thunk: Thunk): Value;
}

/** A deferred argument. */
export interface Thunk {
  readonly ast: Ast;
  evaluate(): Value;
}

export type FunctionImpl = (args: Value[], ctx: FunctionContext) => Value;

export interface FunctionSpec {
  /** Canonical upper-case name, as stored in the file. */
  name: string;
  params: ParamSpec[];
  impl: FunctionImpl;
  /**
   * Stored in xlsx with the `_xlfn.` prefix. Post-2007 functions must round-trip
   * that way or Excel rejects the file.
   */
  futureFunction?: boolean;
  /** Worksheet-namespaced future function, stored as `_xlfn._xlws.NAME`. */
  worksheetScoped?: boolean;
  /**
   * Recalculate on every pass. A predicate lets SUMIF, CELL and INFO be
   * volatile only for the argument shapes that actually make them so.
   */
  volatile?: boolean | ((args: Ast[]) => boolean);
  /** Depends on sheet structure: dirtied by insert, delete, move and rename. */
  structural?: boolean;
  /**
   * Safe to evaluate off the main thread. VBA user-defined functions and the
   * handful of functions that read application state are not.
   */
  threadSafe?: boolean;
  /**
   * A scalar function that should be applied element-wise when handed an array.
   * SIN over a range returns an array of sines; SUM does not broadcast.
   */
  broadcast?: boolean;
  /** Older name kept for file compatibility, e.g. NORMDIST for NORM.DIST. */
  deprecatedAliasOf?: string;
  /** One-line description, surfaced in the formula editor. */
  summary?: string;
}

/**
 * Post-2007 functions, which xlsx stores prefixed.
 *
 * This is the same set the fixture generator applies on write and the reader
 * strips on read; keeping it here as data means the writer can consult one
 * source of truth rather than a second hand-maintained list.
 */
export const FUTURE_FUNCTIONS: ReadonlySet<string> = new Set([
  'IFS', 'XOR', 'TEXTJOIN', 'CONCAT', 'SWITCH', 'MAXIFS', 'MINIFS', 'IFNA',
  'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'PERCENTILE.INC', 'PERCENTILE.EXC',
  'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'RANK.AVG', 'MODE.SNGL', 'MODE.MULT',
  'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'T.TEST', 'F.TEST',
  'CHISQ.TEST', 'COVARIANCE.P', 'COVARIANCE.S', 'BINOM.DIST', 'EXPON.DIST',
  'CEILING.MATH', 'FLOOR.MATH', 'CEILING.PRECISE', 'FLOOR.PRECISE',
  'XLOOKUP', 'XMATCH', 'LET', 'LAMBDA', 'BYROW', 'BYCOL', 'MAP', 'REDUCE',
  'SCAN', 'MAKEARRAY', 'ISOMITTED', 'UNIQUE', 'SORT', 'SORTBY', 'SEQUENCE',
  'RANDARRAY', 'ARRAYTOTEXT', 'VALUETOTEXT', 'TEXTSPLIT', 'TEXTBEFORE',
  'TEXTAFTER', 'VSTACK', 'HSTACK', 'TOCOL', 'TOROW', 'CHOOSECOLS',
  'CHOOSEROWS', 'WRAPROWS', 'WRAPCOLS', 'EXPAND', 'TAKE', 'DROP',
  'REGEXTEST', 'REGEXEXTRACT', 'REGEXREPLACE', 'GROUPBY', 'PIVOTBY',
  'PERCENTOF', 'TRIMRANGE', 'DAYS', 'ISOWEEKNUM', 'NUMBERVALUE', 'UNICHAR',
  'UNICODE', 'BASE', 'DECIMAL', 'COMBINA', 'PERMUTATIONA', 'SEC', 'CSC', 'COT',
  'ACOT', 'SECH', 'CSCH', 'COTH', 'ARABIC', 'BITAND', 'BITOR', 'BITXOR',
  'BITLSHIFT', 'BITRSHIFT', 'PDURATION', 'RRI', 'FORMULATEXT', 'SHEET', 'SHEETS',
  'IMTAN', 'IMCOSH', 'IMSINH', 'IMSEC', 'IMCSC', 'AGGREGATE', 'WORKDAY.INTL',
  'NETWORKDAYS.INTL', 'ERF.PRECISE', 'ERFC.PRECISE', 'GAMMA', 'GAUSS', 'PHI',
  'SKEW.P', 'WEIBULL.DIST', 'Z.TEST', 'F.DIST', 'F.INV', 'T.DIST', 'T.INV',
  'CHISQ.DIST', 'CHISQ.INV', 'BETA.DIST', 'BETA.INV', 'GAMMA.DIST', 'GAMMA.INV',
  'LOGNORM.DIST', 'LOGNORM.INV', 'HYPGEOM.DIST', 'NEGBINOM.DIST', 'POISSON.DIST',
  'CONFIDENCE.NORM', 'CONFIDENCE.T', 'BINOM.INV', 'BINOM.DIST.RANGE',
]);

/** Future functions stored under the worksheet namespace. */
export const WORKSHEET_SCOPED_FUNCTIONS: ReadonlySet<string> = new Set([
  'FILTER',
  'ANCHORARRAY',
  'SINGLE',
]);

/** Re-apply the storage prefix a function name needs in xlsx. */
export function storageName(name: string): string {
  const upper = name.toUpperCase();
  if (WORKSHEET_SCOPED_FUNCTIONS.has(upper)) return `_xlfn._xlws.${upper}`;
  if (FUTURE_FUNCTIONS.has(upper)) return `_xlfn.${upper}`;
  return name;
}

export class FunctionRegistry {
  private readonly byName = new Map<string, FunctionSpec>();

  register(spec: FunctionSpec): this {
    const name = spec.name.toUpperCase();
    if (this.byName.has(name)) {
      throw new Error(`function ${name} is already registered`);
    }
    this.byName.set(name, {
      ...spec,
      name,
      futureFunction: spec.futureFunction ?? FUTURE_FUNCTIONS.has(name),
      worksheetScoped: spec.worksheetScoped ?? WORKSHEET_SCOPED_FUNCTIONS.has(name),
      threadSafe: spec.threadSafe ?? true,
    });
    return this;
  }

  registerAll(specs: readonly FunctionSpec[]): this {
    for (const s of specs) this.register(s);
    return this;
  }

  /**
   * Register a compatibility alias, e.g. NORMDIST for NORM.DIST. Old files use
   * the old names and must not show #NAME?.
   */
  alias(oldName: string, canonical: string): this {
    const target = this.byName.get(canonical.toUpperCase());
    if (!target) throw new Error(`cannot alias ${oldName}: ${canonical} is not registered`);
    return this.register({
      ...target,
      name: oldName.toUpperCase(),
      futureFunction: false,
      worksheetScoped: false,
      deprecatedAliasOf: target.name,
    });
  }

  get(name: string): FunctionSpec | undefined {
    // The `_xlfn.` prefixes are a storage detail; a formula that still carries
    // one must resolve to the same function.
    const bare = name.replace(/^_xlfn\._xlws\./i, '').replace(/^_xlfn\./i, '');
    return this.byName.get(bare.toUpperCase());
  }

  has(name: string): boolean {
    return this.get(name) !== undefined;
  }

  get size(): number {
    return this.byName.size;
  }

  names(): string[] {
    return [...this.byName.keys()].sort();
  }

  /** Arity check, returning a human-readable problem rather than throwing. */
  checkArity(spec: FunctionSpec, count: number): string | undefined {
    const required = spec.params.filter((p) => !p.optional && !p.repeating).length;
    const hasRepeating = spec.params.some((p) => p.repeating);
    const max = hasRepeating ? Number.POSITIVE_INFINITY : spec.params.length;
    if (count < required) {
      return `${spec.name} needs at least ${required} argument${required === 1 ? '' : 's'}, got ${count}`;
    }
    if (count > max) {
      return `${spec.name} takes at most ${max} argument${max === 1 ? '' : 's'}, got ${count}`;
    }
    return undefined;
  }

  /** The parameter spec governing argument `index`, honouring repetition. */
  paramAt(spec: FunctionSpec, index: number): ParamSpec | undefined {
    const direct = spec.params[index];
    if (direct) return direct;
    const last = spec.params[spec.params.length - 1];
    return last?.repeating ? last : undefined;
  }
}

/** Convenience builders, so function tables stay readable. */
export const p = {
  scalar: (name: string, optional = false): ParamSpec =>
    optional ? { name, kind: ArgKind.Scalar, optional } : { name, kind: ArgKind.Scalar },
  array: (name: string, optional = false): ParamSpec =>
    optional ? { name, kind: ArgKind.Array, optional } : { name, kind: ArgKind.Array },
  ref: (name: string, optional = false): ParamSpec =>
    optional ? { name, kind: ArgKind.Reference, optional } : { name, kind: ArgKind.Reference },
  lazy: (name: string, optional = false): ParamSpec =>
    optional ? { name, kind: ArgKind.Lazy, optional } : { name, kind: ArgKind.Lazy },
  any: (name: string, optional = false): ParamSpec =>
    optional ? { name, kind: ArgKind.Any, optional } : { name, kind: ArgKind.Any },
  rest: (name: string, kind: ArgKind = ArgKind.Array): ParamSpec => ({
    name,
    kind,
    repeating: true,
    optional: true,
  }),
} as const;

export type { CellError };
