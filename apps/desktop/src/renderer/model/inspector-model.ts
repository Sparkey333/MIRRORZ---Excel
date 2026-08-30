/**
 * The view model behind the formula inspector.
 *
 * Excel's auditing tools draw one hop of arrows at a time, so finding out where
 * a `#VALUE!` on a summary sheet actually came from is a manual walk backwards
 * through however many intermediate cells there are. The dependency graph knows
 * the answer outright, so the inspector's job is only to say it plainly: here is
 * the value, here is what it reads, and - for an error - here is the cell that
 * started it.
 *
 * Everything in this module is a pure function of an `explain` callback, which
 * is what lets the whole panel be tested without an engine or a DOM.
 */

import { a1, isError, type Scalar } from '@mirrorz/core';
import type { CellAddr, CellExplanation, RangeAddr } from '@mirrorz/formula';

export function formatAddr(addr: CellAddr): string {
  return `${addr.sheet}!${a1(addr.row, addr.col)}`;
}

export function formatRangeAddr(range: RangeAddr): string {
  const start = a1(range.startRow, range.startCol);
  const end = a1(range.endRow, range.endCol);
  return `${range.sheet}!${start === end ? start : `${start}:${end}`}`;
}

export function sameAddr(a: CellAddr, b: CellAddr): boolean {
  return a.sheet === b.sheet && a.row === b.row && a.col === b.col;
}

/** A value rendered for the inspector, which shows types rather than hiding them. */
export function describeValue(value: Scalar): string {
  if (value === null) return 'empty';
  if (isError(value)) return value.detail ? `${value.code} - ${value.detail}` : value.code;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') return `"${value}"`;
  return String(value);
}

export type PrecedentKind = 'cell' | 'range';

export interface PrecedentNode {
  key: string;
  kind: PrecedentKind;
  label: string;
  addr?: CellAddr;
  range?: RangeAddr;
  value?: Scalar;
  formula?: string;
  /** True when this cell holds an error, so the branch can be marked. */
  errored: boolean;
  children: PrecedentNode[];
  /** True when the walk stopped here because it had already visited this cell. */
  cyclic: boolean;
  /** True when the walk stopped here because it hit the depth limit. */
  truncated: boolean;
}

export interface PrecedentTreeOptions {
  maxDepth?: number;
  /** Cap on nodes, so a cell reading ten thousand precedents cannot lock the UI. */
  maxNodes?: number;
}

/**
 * Expand the precedents of a cell into a tree.
 *
 * Ranges are leaves. Expanding a range into its cells is exactly the behaviour
 * that makes Excel's trace-precedents useless on a real model - one arrow per
 * cell of a 5,000-row column - so a range is named as a range and left there.
 */
export function buildPrecedentTree(
  root: CellAddr,
  explain: (addr: CellAddr) => CellExplanation,
  options: PrecedentTreeOptions = {},
): PrecedentNode {
  const maxDepth = options.maxDepth ?? 4;
  const maxNodes = options.maxNodes ?? 200;
  let nodes = 0;

  const visit = (addr: CellAddr, depth: number, seen: Set<string>): PrecedentNode => {
    const key = formatAddr(addr);
    const explanation = explain(addr);
    const node: PrecedentNode = {
      key,
      kind: 'cell',
      label: key,
      addr,
      value: explanation.value,
      errored: isError(explanation.value),
      children: [],
      cyclic: false,
      truncated: false,
    };
    if (explanation.formula !== undefined) node.formula = explanation.formula;
    nodes++;

    if (seen.has(key)) {
      node.cyclic = true;
      return node;
    }
    const hasPrecedents =
      explanation.precedentCells.length > 0 || explanation.precedentRanges.length > 0;
    if (!hasPrecedents) return node;
    if (depth >= maxDepth || nodes >= maxNodes) {
      node.truncated = true;
      return node;
    }

    const nextSeen = new Set(seen).add(key);
    for (const cell of explanation.precedentCells) {
      if (nodes >= maxNodes) {
        node.truncated = true;
        break;
      }
      node.children.push(visit(cell, depth + 1, nextSeen));
    }
    for (const range of explanation.precedentRanges) {
      if (nodes >= maxNodes) {
        node.truncated = true;
        break;
      }
      nodes++;
      node.children.push({
        key: formatRangeAddr(range),
        kind: 'range',
        label: formatRangeAddr(range),
        range,
        errored: false,
        children: [],
        cyclic: false,
        truncated: false,
      });
    }
    return node;
  };

  return visit(root, 0, new Set());
}

export interface ErrorSummary {
  code: string;
  /** The engine's explanation of the error, when it supplied one. */
  detail?: string;
  /** Every cell the graph blames, nearest first. */
  roots: CellAddr[];
  /** The one to lead with, and to jump to on click. */
  primary?: CellAddr;
  /** The sentence the panel shows in the banner. */
  headline: string;
  /** True when the cell is itself the origin, so there is nowhere to jump. */
  selfInflicted: boolean;
}

/**
 * Reduce an explanation to the one sentence worth putting in large type.
 *
 * "The problem starts at Sheet1!B4" is the whole feature. Everything else in the
 * panel is supporting evidence for it.
 */
export function errorSummary(explanation: CellExplanation): ErrorSummary | null {
  const value = explanation.value;
  if (!isError(value)) return null;

  const roots = explanation.errorRoots;
  const selfOnly = roots.length === 0 || roots.every((r) => sameAddr(r, explanation.addr));
  const summary: ErrorSummary = {
    code: value.code,
    roots,
    headline: '',
    selfInflicted: selfOnly,
  };
  if (value.detail !== undefined) summary.detail = value.detail;

  if (selfOnly) {
    summary.headline = `${value.code} originates in this cell`;
    return summary;
  }

  const primary = roots.find((r) => !sameAddr(r, explanation.addr)) ?? roots[0]!;
  summary.primary = primary;
  const others = roots.filter((r) => !sameAddr(r, primary)).length;
  summary.headline =
    others > 0
      ? `The problem starts at ${formatAddr(primary)}, and ${others} other cell${others === 1 ? '' : 's'}`
      : `The problem starts at ${formatAddr(primary)}`;
  return summary;
}

export interface InspectorModel {
  addr: CellAddr;
  label: string;
  value: Scalar;
  valueText: string;
  formula?: string;
  precedents: PrecedentNode;
  dependents: CellAddr[];
  error: ErrorSummary | null;
  /** True when nothing reads this cell and it reads nothing - an island. */
  isolated: boolean;
}

export function buildInspectorModel(
  explanation: CellExplanation,
  explain: (addr: CellAddr) => CellExplanation,
  options?: PrecedentTreeOptions,
): InspectorModel {
  const precedents = buildPrecedentTree(explanation.addr, explain, options);
  const model: InspectorModel = {
    addr: explanation.addr,
    label: formatAddr(explanation.addr),
    value: explanation.value,
    valueText: describeValue(explanation.value),
    precedents,
    dependents: explanation.dependents,
    error: errorSummary(explanation),
    isolated: precedents.children.length === 0 && explanation.dependents.length === 0,
  };
  if (explanation.formula !== undefined) model.formula = explanation.formula;
  return model;
}
