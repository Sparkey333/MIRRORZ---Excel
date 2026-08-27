/**
 * Formula AST.
 *
 * Node kinds are numeric rather than string-tagged: every evaluation dispatches
 * on this tag, and a numeric switch compiles to a jump table where a string
 * switch does not.
 *
 * References are stored relative to the formula's own address with per-axis
 * absolute flags. That single choice buys three things: filling a formula down a
 * column is a no-op re-render rather than a rewrite, two copies of the same
 * formula produce identical ASTs and can share one cached object, and choosing
 * between A1 and R1C1 display becomes a serializer option rather than a
 * different parse.
 */

import type { ErrorCode } from '@mirrorz/core';

export const enum Node {
  Number,
  Text,
  Bool,
  ErrorLit,
  /** A missing argument, as in `IF(A1,,0)`. */
  Missing,
  Unary,
  Binary,
  Postfix,
  Paren,
  Call,
  Array,
  Ref,
  Range,
  /** Whole column or whole row. */
  Beam,
  /** `Sheet1:Sheet3!A1` */
  ThreeD,
  Name,
  StructRef,
}

export interface Positioned {
  /** Whitespace that preceded this node, preserved so a formula round-trips. */
  ws?: string;
}

/**
 * Whitespace fidelity.
 *
 * Excel stores a formula exactly as the user typed it, spacing included, and
 * reflowing it on every save produces spurious diffs and surprises people who
 * deliberately format long formulas across multiple lines. Leading whitespace
 * rides on each node; the extra slots below cover the positions where
 * whitespace has no following operand to attach to.
 */

export interface NumberNode extends Positioned {
  kind: Node.Number;
  value: number;
}

export interface TextNode extends Positioned {
  kind: Node.Text;
  value: string;
}

export interface BoolNode extends Positioned {
  kind: Node.Bool;
  value: boolean;
}

export interface ErrorNode extends Positioned {
  kind: Node.ErrorLit;
  code: ErrorCode;
}

export interface MissingNode extends Positioned {
  kind: Node.Missing;
}

export type UnaryOp = '-' | '+' | '@';
export type PostfixOp = '%' | '#';
export type BinaryOp =
  | '+'
  | '-'
  | '*'
  | '/'
  | '^'
  | '&'
  | '='
  | '<>'
  | '<'
  | '>'
  | '<='
  | '>='
  /** Range: `A1:B2`, and also `A1:INDEX(...)`. */
  | ':'
  /** Union, only valid inside parentheses. */
  | ','
  /** Intersection - a single space between two references. */
  | ' ';

export interface UnaryNode extends Positioned {
  kind: Node.Unary;
  op: UnaryOp;
  operand: Ast;
}

export interface BinaryNode extends Positioned {
  kind: Node.Binary;
  op: BinaryOp;
  left: Ast;
  right: Ast;
  /** Whitespace between the left operand and the operator. */
  wsOp?: string;
}

export interface PostfixNode extends Positioned {
  kind: Node.Postfix;
  op: PostfixOp;
  operand: Ast;
  /** Whitespace between the operand and the operator. */
  wsOp?: string;
}

export interface ParenNode extends Positioned {
  kind: Node.Paren;
  inner: Ast;
  /** Whitespace before the closing paren. */
  wsClose?: string;
}

export interface CallNode extends Positioned {
  kind: Node.Call;
  name: string;
  args: Ast[];
  /** Whitespace following each argument, before its separator or the close. */
  wsArgs?: string[];
}

export interface ArrayNode extends Positioned {
  kind: Node.Array;
  /** Row-major; every row has the same length. */
  rows: Ast[][];
}

/** A cell reference, stored as an offset from the formula's own position. */
export interface RefNode extends Positioned {
  kind: Node.Ref;
  /** Row offset from the formula, or the absolute row when rowAbs. */
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
  /** Sheet name when the reference was qualified; undefined means "this sheet". */
  sheet?: string;
  /** Workbook name for an external reference. */
  book?: string;
}

export interface RangeNode extends Positioned {
  kind: Node.Range;
  start: RefNode;
  end: RefNode;
}

export interface BeamNode extends Positioned {
  kind: Node.Beam;
  axis: 'col' | 'row';
  from: number;
  to: number;
  fromAbs: boolean;
  toAbs: boolean;
  sheet?: string;
  book?: string;
}

export interface ThreeDNode extends Positioned {
  kind: Node.ThreeD;
  sheetStart: string;
  sheetEnd: string;
  inner: Ast;
}

export interface NameNode extends Positioned {
  kind: Node.Name;
  name: string;
  sheet?: string;
  book?: string;
}

export interface StructRefNode extends Positioned {
  kind: Node.StructRef;
  table: string;
  /** Raw specifier text, e.g. `#Headers`, `@`, `Column1:Column3`. */
  spec: string;
}

export type Ast =
  | NumberNode
  | TextNode
  | BoolNode
  | ErrorNode
  | MissingNode
  | UnaryNode
  | BinaryNode
  | PostfixNode
  | ParenNode
  | CallNode
  | ArrayNode
  | RefNode
  | RangeNode
  | BeamNode
  | ThreeDNode
  | NameNode
  | StructRefNode;

/**
 * Binding powers, matching Excel's documented operator precedence.
 *
 * Two of these are counter-intuitive and both are verified against the fixture
 * oracle rather than assumed:
 *
 *   `^` is LEFT-associative, so `2^3^2` is `(2^3)^2` = 64, not 512.
 *   Unary minus binds TIGHTER than `^`, so `-2^2` is `(-2)^2` = 4, not -4.
 *   Postfix `%` binds tighter still, so `2^2%` is `2^(2%)` = 1.0139...
 *
 * Most parser generators get all three wrong by default, and the resulting
 * errors are silent wrong numbers rather than parse failures.
 */
export const BINDING_POWER: Readonly<Record<string, number>> = Object.freeze({
  ':': 11,
  ' ': 10,
  ',': 9,
  // Prefix operators sit at 7 and postfix at 6; both are applied by the parser
  // directly rather than through this table, which holds infix powers.
  '^': 5,
  '*': 4,
  '/': 4,
  '+': 3,
  '-': 3,
  '&': 2,
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
});

export const PREFIX_BINDING_POWER = 7;
export const POSTFIX_BINDING_POWER = 6;

/** Walk an AST, calling `visit` for every node, parents before children. */
export function walk(node: Ast, visit: (n: Ast) => void): void {
  visit(node);
  switch (node.kind) {
    case Node.Unary:
      walk(node.operand, visit);
      break;
    case Node.Postfix:
      walk(node.operand, visit);
      break;
    case Node.Binary:
      walk(node.left, visit);
      walk(node.right, visit);
      break;
    case Node.Paren:
      walk(node.inner, visit);
      break;
    case Node.Call:
      for (const a of node.args) walk(a, visit);
      break;
    case Node.Array:
      for (const row of node.rows) for (const a of row) walk(a, visit);
      break;
    case Node.Range:
      walk(node.start, visit);
      walk(node.end, visit);
      break;
    case Node.ThreeD:
      walk(node.inner, visit);
      break;
    default:
      break;
  }
}
