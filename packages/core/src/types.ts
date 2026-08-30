/**
 * Core value model.
 *
 * A spreadsheet cell holds exactly one of five value kinds. Excel itself models
 * this the same way (its `Variant` for a cell is number / text / boolean / error
 * / empty), and matching that shape one-to-one keeps conversion rules honest all
 * the way from the file format through the formula engine to the renderer.
 */

/** The seven error values Excel can store in a cell. */
export const ERROR_CODES = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  // Newer errors that appear in dynamic-array and data-type workbooks.
  '#SPILL!',
  '#CALC!',
  '#GETTING_DATA',
  '#FIELD!',
  '#BLOCKED!',
  '#CONNECT!',
  '#UNKNOWN!',
  '#BUSY!',
] as const;

export type ErrorCode = (typeof ERROR_CODES)[number];

/** A cell error. Kept as a branded object so it never silently coerces to text. */
export class CellError {
  readonly kind = 'error' as const;
  constructor(
    readonly code: ErrorCode,
    /** Optional human-readable cause, surfaced in our error-explainer UI. */
    readonly detail?: string,
  ) {}

  toString(): string {
    return this.code;
  }

  static readonly NULL = new CellError('#NULL!');
  static readonly DIV0 = new CellError('#DIV/0!');
  static readonly VALUE = new CellError('#VALUE!');
  static readonly REF = new CellError('#REF!');
  static readonly NAME = new CellError('#NAME?');
  static readonly NUM = new CellError('#NUM!');
  static readonly NA = new CellError('#N/A');
  static readonly SPILL = new CellError('#SPILL!');
  static readonly CALC = new CellError('#CALC!');
}

export function isError(v: unknown): v is CellError {
  return v instanceof CellError;
}

/** Look up the canonical singleton for an error string, if it is one. */
export function errorFromCode(code: string): CellError | undefined {
  return (ERROR_CODES as readonly string[]).includes(code)
    ? new CellError(code as ErrorCode)
    : undefined;
}

/**
 * A scalar cell value. `null` means "empty" - distinct from the empty string,
 * which is a real text value that a formula can produce.
 */
export type Scalar = number | string | boolean | CellError | null;

/** A rectangular block of scalars, row-major. Produced by array formulas. */
export type Matrix = Scalar[][];

export type CellValue = Scalar | Matrix;

export function isMatrix(v: CellValue): v is Matrix {
  return Array.isArray(v);
}

/** Excel's own type ordinals, as returned by the TYPE() worksheet function. */
export enum ValueType {
  Number = 1,
  Text = 2,
  Logical = 4,
  Error = 16,
  Array = 64,
}

export function typeOf(v: CellValue): ValueType {
  if (isMatrix(v)) return ValueType.Array;
  if (isError(v)) return ValueType.Error;
  if (typeof v === 'boolean') return ValueType.Logical;
  if (typeof v === 'string') return ValueType.Text;
  // Excel reports an empty cell as a number (0) for TYPE().
  return ValueType.Number;
}
