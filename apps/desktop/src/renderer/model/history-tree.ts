/**
 * Turning the command log into something a panel can draw.
 *
 * The document's history is a tree: undo then edit does not destroy the branch
 * you undid, it forks. That is the whole point - exploring an alternative is not
 * a one-way door - but it means the panel cannot just print a list. It has to
 * show which line of work is the live one, which are abandoned but still
 * reachable, and where they diverged.
 *
 * The layout rule is: the path to the current head is drawn as a straight
 * column, and every other branch hangs off it indented. That way the common
 * case, where nobody has ever undone anything, looks exactly like the flat list
 * people expect, and the tree only appears once there is actually a tree.
 */

import type { ChangeOrigin, HistoryEntry } from '@mirrorz/core';

export interface HistoryNode {
  entry: HistoryEntry;
  children: HistoryNode[];
}

export interface HistoryRow {
  id: number;
  label: string;
  origin: ChangeOrigin;
  timestamp: number;
  barrier: boolean;
  changeCount: number;
  /** Indentation level; 0 is the main line of work. */
  depth: number;
  /** True when this entry is an ancestor of, or is, the current head. */
  onPath: boolean;
  isHead: boolean;
  /** True when this entry has more than one child, so the timeline forks here. */
  isForkPoint: boolean;
  /** True when this row starts a branch that was abandoned. */
  startsBranch: boolean;
}

/** Roots of the history forest, children ordered by id. */
export function buildHistoryTree(entries: readonly HistoryEntry[]): HistoryNode[] {
  const nodes = new Map<number, HistoryNode>();
  for (const entry of entries) nodes.set(entry.id, { entry, children: [] });

  const roots: HistoryNode[] = [];
  for (const entry of entries) {
    const node = nodes.get(entry.id)!;
    const parent = entry.parent === null ? undefined : nodes.get(entry.parent);
    // An entry whose parent is missing is treated as a root rather than dropped,
    // because losing history silently is worse than drawing it in the wrong place.
    if (parent) parent.children.push(node);
    else roots.push(node);
  }
  for (const node of nodes.values()) node.children.sort((a, b) => a.entry.id - b.entry.id);
  roots.sort((a, b) => a.entry.id - b.entry.id);
  return roots;
}

/** The ids from the root down to `headId` inclusive. */
export function pathToHead(entries: readonly HistoryEntry[], headId: number | null): Set<number> {
  const byId = new Map(entries.map((e) => [e.id, e]));
  const path = new Set<number>();
  let current = headId;
  // A malformed parent chain would otherwise spin forever; the visited set is
  // cheap insurance on data that arrives from a file.
  while (current !== null && !path.has(current)) {
    path.add(current);
    current = byId.get(current)?.parent ?? null;
  }
  return path;
}

/**
 * Flatten the tree into rows in display order.
 *
 * At every fork the child that leads to the head is visited first, so the live
 * timeline reads top to bottom without interruption and the abandoned branches
 * follow it.
 */
export function flattenHistory(
  entries: readonly HistoryEntry[],
  headId: number | null,
): HistoryRow[] {
  const roots = buildHistoryTree(entries);
  const path = pathToHead(entries, headId);
  const rows: HistoryRow[] = [];

  const visit = (node: HistoryNode, depth: number, startsBranch: boolean): void => {
    rows.push({
      id: node.entry.id,
      label: node.entry.label,
      origin: node.entry.origin,
      timestamp: node.entry.timestamp,
      barrier: node.entry.barrier === true,
      changeCount: node.entry.changes.length,
      depth,
      onPath: path.has(node.entry.id),
      isHead: node.entry.id === headId,
      isForkPoint: node.children.length > 1,
      startsBranch,
    });

    const children = [...node.children].sort((a, b) => {
      const aOn = path.has(a.entry.id) ? 0 : 1;
      const bOn = path.has(b.entry.id) ? 0 : 1;
      return aOn - bOn || a.entry.id - b.entry.id;
    });
    for (let i = 0; i < children.length; i++) {
      // Only a real fork indents; a single child continues its parent's column.
      const childDepth = children.length > 1 && i > 0 ? depth + 1 : depth;
      visit(children[i]!, childDepth, children.length > 1 && i > 0);
    }
  };

  for (const root of roots) visit(root, 0, false);
  return rows;
}

/** How an origin is labelled in the panel. */
export const ORIGIN_LABELS: Readonly<Record<ChangeOrigin, string>> = Object.freeze({
  user: 'you',
  script: 'script',
  macro: 'macro',
  import: 'import',
  recalc: 'recalculation',
  system: 'system',
});

/**
 * A relative time for the history rows.
 *
 * Absolute timestamps are useless at this density - forty rows all saying
 * 14:32 - and the question the panel answers is "how far back is this", so it
 * says that instead.
 */
export function relativeTime(timestamp: number, now: number): string {
  if (timestamp === 0) return '';
  const seconds = Math.round((now - timestamp) / 1000);
  if (seconds < 0) return 'just now';
  if (seconds < 10) return 'just now';
  if (seconds < 60) return `${seconds}s ago`;
  const minutes = Math.round(seconds / 60);
  if (minutes < 60) return `${minutes}m ago`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `${hours}h ago`;
  return `${Math.round(hours / 24)}d ago`;
}

export interface HistorySummary {
  total: number;
  onPath: number;
  abandoned: number;
  branches: number;
}

/** Counts for the panel header, so "3 abandoned branches" can be stated. */
export function summariseHistory(
  entries: readonly HistoryEntry[],
  headId: number | null,
): HistorySummary {
  const path = pathToHead(entries, headId);
  const forks = buildHistoryTree(entries);
  let branches = 0;
  const countForks = (node: HistoryNode): void => {
    if (node.children.length > 1) branches += node.children.length - 1;
    for (const child of node.children) countForks(child);
  };
  for (const root of forks) countForks(root);
  if (forks.length > 1) branches += forks.length - 1;
  return {
    total: entries.length,
    onPath: path.size,
    abandoned: entries.length - path.size,
    branches,
  };
}
