/**
 * Turn an AST back into formula text.
 *
 * References are stored relative to the formula's origin, so serialising needs
 * to be told which cell the formula now lives at. That is what makes fill-down
 * work: the same AST rendered at a different origin produces the shifted
 * formula, with no rewriting of the tree.
 */

import { colToName, quoteSheetName } from '@mirrorz/core';
import { type Ast, Node, type RefNode } from './ast.js';

export interface SerializeOptions {
  origin?: { row: number; col: number };
  argSeparator?: ',' | ';';
  /** Emit the leading `=`. Off by default, since xlsx stores formulas without it. */
  withEquals?: boolean;
  /** Re-emit the whitespace captured at parse time. */
  preserveWhitespace?: boolean;
}

export function serializeFormula(ast: Ast, options: SerializeOptions = {}): string {
  const origin = options.origin ?? { row: 0, col: 0 };
  const sep = options.argSeparator ?? ',';
  const ws = options.preserveWhitespace ?? true;

  const emit = (node: Ast): string => {
    const lead = ws && node.ws ? node.ws : '';
    return lead + body(node);
  };

  const body = (node: Ast): string => {
    switch (node.kind) {
      case Node.Number:
        return formatNumber(node.value);
      case Node.Text:
        return `"${node.value.replaceAll('"', '""')}"`;
      case Node.Bool:
        return node.value ? 'TRUE' : 'FALSE';
      case Node.ErrorLit:
        return node.code;
      case Node.Missing:
        return '';
      case Node.Unary:
        return node.op + emit(node.operand);
      case Node.Postfix:
        return emit(node.operand) + (ws && node.wsOp ? node.wsOp : '') + node.op;
      case Node.Binary: {
        const gap = ws && node.wsOp ? node.wsOp : '';
        // The intersection operator IS whitespace, so it must not be doubled.
        const opText = node.op === ' ' ? (gap || ' ') : gap + node.op;
        return emit(node.left) + opText + emit(node.right);
      }
      case Node.Paren:
        return `(${emit(node.inner)}${ws && node.wsClose ? node.wsClose : ''})`;
      case Node.Call: {
        const gaps = ws ? (node.wsArgs ?? []) : [];
        if (node.args.length === 0) return `${node.name}(${gaps[0] ?? ''})`;
        const parts = node.args.map((a, i) => emit(a) + (gaps[i] ?? ''));
        return `${node.name}(${parts.join(sep)})`;
      }
      case Node.Array:
        return `{${node.rows.map((r) => r.map(emit).join(sep)).join(';')}}`;
      case Node.Ref:
        return refText(node, origin);
      case Node.Range: {
        // The sheet prefix belongs to the range as a whole, not to both ends.
        const prefix = contextPrefix(node.start);
        return `${prefix}${bareRef(node.start, origin)}:${bareRef(node.end, origin)}`;
      }
      case Node.Beam: {
        const prefix = contextPrefix(node);
        const abs = (v: boolean) => (v ? '$' : '');
        if (node.axis === 'col') {
          const from = node.fromAbs ? node.from : node.from + origin.col;
          const to = node.toAbs ? node.to : node.to + origin.col;
          return `${prefix}${abs(node.fromAbs)}${colToName(from)}:${abs(node.toAbs)}${colToName(to)}`;
        }
        const from = (node.fromAbs ? node.from : node.from + origin.row) + 1;
        const to = (node.toAbs ? node.to : node.to + origin.row) + 1;
        return `${prefix}${abs(node.fromAbs)}${from}:${abs(node.toAbs)}${to}`;
      }
      case Node.ThreeD: {
        const inner = emit(node.inner);
        // The inner reference already carries the start sheet's prefix; replace
        // it with the span form.
        const bang = inner.indexOf('!');
        const bare = bang < 0 ? inner : inner.slice(bang + 1);
        return `${quoteSheetName(node.sheetStart)}:${quoteSheetName(node.sheetEnd)}!${bare}`;
      }
      case Node.Name:
        return contextPrefix(node) + node.name;
      case Node.StructRef:
        return node.table + node.spec;
    }
  };

  return (options.withEquals ? '=' : '') + emit(ast);
}

function contextPrefix(node: { sheet?: string; book?: string }): string {
  if (node.sheet === undefined) return '';
  const sheet = node.book === undefined ? node.sheet : `[${node.book}]${node.sheet}`;
  return `${quoteSheetName(sheet)}!`;
}

function bareRef(ref: RefNode, origin: { row: number; col: number }): string {
  const col = ref.colAbs ? ref.col : ref.col + origin.col;
  const row = ref.rowAbs ? ref.row : ref.row + origin.row;
  // A reference shifted off the grid is #REF!, which is exactly what Excel shows
  // after deleting the rows a relative reference pointed at.
  if (col < 0 || col >= 16_384 || row < 0 || row >= 1_048_576) return '#REF!';
  return `${ref.colAbs ? '$' : ''}${colToName(col)}${ref.rowAbs ? '$' : ''}${row + 1}`;
}

function refText(ref: RefNode, origin: { row: number; col: number }): string {
  return contextPrefix(ref) + bareRef(ref, origin);
}

/**
 * Render a numeric literal the way Excel does: no exponent for ordinary
 * magnitudes, and no trailing `.0`.
 */
function formatNumber(v: number): string {
  if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
  // JavaScript writes a lower-case exponent; Excel writes an upper-case one,
  // and a formula that changes case on every save produces spurious diffs.
  return String(v).replace('e', 'E');
}
