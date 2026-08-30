import { describe, expect, it } from 'vitest';
import type { HistoryEntry } from '@mirrorz/core';
import {
  buildHistoryTree,
  flattenHistory,
  pathToHead,
  relativeTime,
  summariseHistory,
} from '../src/renderer/model/history-tree.js';

function entry(id: number, parent: number | null, extra: Partial<HistoryEntry> = {}): HistoryEntry {
  return {
    id,
    label: `Entry ${id}`,
    origin: 'user',
    timestamp: 1_000 * id,
    changes: [],
    parent,
    ...extra,
  };
}

/**
 * A branching history:
 *   1 -> 2 -> 3
 *        \--> 4 -> 5
 * with the head left on 3, so 4 and 5 are the abandoned branch.
 */
const branching: HistoryEntry[] = [
  entry(1, null),
  entry(2, 1),
  entry(3, 2),
  entry(4, 2),
  entry(5, 4),
];

describe('buildHistoryTree', () => {
  it('builds a single root chain', () => {
    const roots = buildHistoryTree([entry(1, null), entry(2, 1)]);
    expect(roots).toHaveLength(1);
    expect(roots[0]!.children[0]!.entry.id).toBe(2);
  });

  it('attaches both branches to the fork point', () => {
    const roots = buildHistoryTree(branching);
    const fork = roots[0]!.children[0]!;
    expect(fork.entry.id).toBe(2);
    expect(fork.children.map((c) => c.entry.id)).toEqual([3, 4]);
  });

  it('treats an entry with a missing parent as a root rather than dropping it', () => {
    const roots = buildHistoryTree([entry(7, 99)]);
    expect(roots.map((r) => r.entry.id)).toEqual([7]);
  });

  it('returns nothing for an empty history', () => {
    expect(buildHistoryTree([])).toEqual([]);
  });
});

describe('pathToHead', () => {
  it('walks from the head back to the root', () => {
    expect([...pathToHead(branching, 3)]).toEqual([3, 2, 1]);
  });

  it('is empty at the initial state', () => {
    expect(pathToHead(branching, null).size).toBe(0);
  });

  it('terminates on a cyclic parent chain', () => {
    const cyclic: HistoryEntry[] = [entry(1, 2), entry(2, 1)];
    expect(pathToHead(cyclic, 1).size).toBe(2);
  });
});

describe('flattenHistory', () => {
  it('renders a linear history flat, with no indentation', () => {
    const rows = flattenHistory([entry(1, null), entry(2, 1), entry(3, 2)], 3);
    expect(rows.map((r) => r.depth)).toEqual([0, 0, 0]);
    expect(rows.map((r) => r.id)).toEqual([1, 2, 3]);
  });

  it('puts the live branch first and indents the abandoned one', () => {
    const rows = flattenHistory(branching, 3);
    expect(rows.map((r) => r.id)).toEqual([1, 2, 3, 4, 5]);
    expect(rows.map((r) => r.depth)).toEqual([0, 0, 0, 1, 1]);
  });

  it('follows the head when the other branch is the live one', () => {
    const rows = flattenHistory(branching, 5);
    expect(rows.map((r) => r.id)).toEqual([1, 2, 4, 5, 3]);
    expect(rows.find((r) => r.id === 3)!.depth).toBe(1);
  });

  it('marks the fork point', () => {
    const rows = flattenHistory(branching, 3);
    expect(rows.filter((r) => r.isForkPoint).map((r) => r.id)).toEqual([2]);
  });

  it('marks which rows are on the current path', () => {
    const rows = flattenHistory(branching, 3);
    expect(rows.filter((r) => r.onPath).map((r) => r.id)).toEqual([1, 2, 3]);
  });

  it('marks the head', () => {
    const rows = flattenHistory(branching, 3);
    expect(rows.filter((r) => r.isHead).map((r) => r.id)).toEqual([3]);
  });

  it('flags the first row of an abandoned branch', () => {
    const rows = flattenHistory(branching, 3);
    expect(rows.filter((r) => r.startsBranch).map((r) => r.id)).toEqual([4]);
  });

  it('carries the origin, timestamp, barrier flag and change count through', () => {
    const rows = flattenHistory(
      [entry(1, null, { origin: 'macro', barrier: true, changes: [{} as never, {} as never] })],
      1,
    );
    expect(rows[0]).toMatchObject({ origin: 'macro', barrier: true, changeCount: 2, timestamp: 1000 });
  });

  it('handles an empty history', () => {
    expect(flattenHistory([], null)).toEqual([]);
  });
});

describe('summariseHistory', () => {
  it('counts entries on and off the path', () => {
    expect(summariseHistory(branching, 3)).toEqual({
      total: 5,
      onPath: 3,
      abandoned: 2,
      branches: 1,
    });
  });

  it('reports no branches for a linear history', () => {
    expect(summariseHistory([entry(1, null), entry(2, 1)], 2).branches).toBe(0);
  });

  it('counts multiple forks from one entry', () => {
    const wide = [entry(1, null), entry(2, 1), entry(3, 1), entry(4, 1)];
    expect(summariseHistory(wide, 2).branches).toBe(2);
  });
});

describe('relativeTime', () => {
  const now = 1_000_000;

  it('renders an unset timestamp as nothing', () => {
    expect(relativeTime(0, now)).toBe('');
  });

  it('renders a few seconds as just now', () => {
    expect(relativeTime(now - 3_000, now)).toBe('just now');
  });

  it('renders seconds', () => {
    expect(relativeTime(now - 42_000, now)).toBe('42s ago');
  });

  it('renders minutes', () => {
    expect(relativeTime(now - 300_000, now)).toBe('5m ago');
  });

  it('renders hours', () => {
    expect(relativeTime(now - 7_200_000, now)).toBe('2h ago');
  });

  it('renders days', () => {
    expect(relativeTime(now - 172_800_000, now)).toBe('2d ago');
  });

  it('does not render a future timestamp as negative', () => {
    expect(relativeTime(now + 10_000, now)).toBe('just now');
  });
});
