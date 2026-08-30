/**
 * The recent-files list is small enough to be tested exhaustively, and its bugs
 * are all the visible kind: duplicates, a menu that grows without bound, a list
 * that disappears because someone hand-edited the settings file.
 */

import { describe, expect, it } from 'vitest';
import { RECENT_LIMIT, RecentFiles, foldsCase, parseRecent, recentMenuLabels } from '../../src/main/recent.js';
import { memoryStore } from '../../src/main/store.js';

function make(initial?: unknown, platform: NodeJS.Platform = 'linux'): { recent: RecentFiles; store: ReturnType<typeof memoryStore> } {
  const store = memoryStore(initial);
  return { recent: new RecentFiles(store, platform), store };
}

describe('RecentFiles', () => {
  it('starts empty when there is nothing on disk', () => {
    expect(make().recent.list()).toEqual([]);
  });

  it('puts the most recently opened file first', () => {
    const { recent } = make();
    recent.add('/home/user/a.xlsx', 1);
    recent.add('/home/user/b.xlsx', 2);
    expect(recent.paths()).toEqual(['/home/user/b.xlsx', '/home/user/a.xlsx']);
  });

  it('moves a re-opened file to the top instead of duplicating it', () => {
    const { recent } = make();
    recent.add('/home/user/a.xlsx', 1);
    recent.add('/home/user/b.xlsx', 2);
    recent.add('/home/user/a.xlsx', 3);
    expect(recent.paths()).toEqual(['/home/user/a.xlsx', '/home/user/b.xlsx']);
  });

  it('caps the list so the File menu cannot outgrow the screen', () => {
    const { recent } = make();
    for (let i = 0; i < RECENT_LIMIT + 8; i++) recent.add(`/home/user/f${i}.xlsx`, i);
    expect(recent.list()).toHaveLength(RECENT_LIMIT);
    expect(recent.paths()[0]).toBe(`/home/user/f${RECENT_LIMIT + 7}.xlsx`);
  });

  it('treats paths case-insensitively on Windows and macOS', () => {
    const { recent } = make(undefined, 'darwin');
    recent.add('/Users/u/Budget.xlsx', 1);
    recent.add('/Users/u/budget.xlsx', 2);
    expect(recent.list()).toHaveLength(1);
  });

  it('treats paths case-sensitively on Linux, where they are different files', () => {
    const { recent } = make(undefined, 'linux');
    recent.add('/home/u/Budget.xlsx', 1);
    recent.add('/home/u/budget.xlsx', 2);
    expect(recent.list()).toHaveLength(2);
  });

  it('reports which platforms fold case', () => {
    expect(foldsCase('win32')).toBe(true);
    expect(foldsCase('darwin')).toBe(true);
    expect(foldsCase('linux')).toBe(false);
    expect(foldsCase('freebsd')).toBe(false);
  });

  it('ignores a path that would not pass the open check', () => {
    const { recent } = make();
    recent.add('relative.xlsx', 1);
    recent.add('/home/user/payload.exe', 2);
    recent.add('/home/user/ok.xlsx', 3);
    expect(recent.paths()).toEqual(['/home/user/ok.xlsx']);
  });

  it('persists through the store, most recent first', () => {
    const { recent, store } = make();
    recent.add('/home/user/a.xlsx', 10);
    const raw = store.read() as { version: number; entries: { path: string; openedAt: number }[] };
    expect(raw.version).toBe(1);
    expect(raw.entries[0]).toEqual({ path: '/home/user/a.xlsx', openedAt: 10 });
  });

  it('reloads what a previous run persisted', () => {
    const store = memoryStore();
    new RecentFiles(store, 'linux').add('/home/user/a.xlsx', 10);
    expect(new RecentFiles(store, 'linux').paths()).toEqual(['/home/user/a.xlsx']);
  });

  it('removes a single entry', () => {
    const { recent } = make();
    recent.add('/home/user/a.xlsx', 1);
    recent.add('/home/user/b.xlsx', 2);
    recent.remove('/home/user/a.xlsx');
    expect(recent.paths()).toEqual(['/home/user/b.xlsx']);
  });

  it('clears the whole list', () => {
    const { recent, store } = make();
    recent.add('/home/user/a.xlsx', 1);
    recent.clear();
    expect(recent.list()).toEqual([]);
    expect((store.read() as { entries: unknown[] }).entries).toEqual([]);
  });

  it('prunes files that have gone, and reports how many', () => {
    const { recent } = make();
    recent.add('/home/user/gone.xlsx', 1);
    recent.add('/home/user/here.xlsx', 2);
    const dropped = recent.prune((p) => p.endsWith('here.xlsx'));
    expect(dropped).toBe(1);
    expect(recent.paths()).toEqual(['/home/user/here.xlsx']);
  });

  it('hands out copies, so a caller cannot mutate the list in place', () => {
    const { recent } = make();
    recent.add('/home/user/a.xlsx', 1);
    const list = recent.list();
    const first = list[0];
    if (first) first.path = '/home/user/hacked.xlsx';
    expect(recent.paths()).toEqual(['/home/user/a.xlsx']);
  });
});

describe('parseRecent', () => {
  it('survives every shape a corrupt settings file can take', () => {
    expect(parseRecent(undefined)).toEqual([]);
    expect(parseRecent(null)).toEqual([]);
    expect(parseRecent('not json')).toEqual([]);
    expect(parseRecent({ version: 2, entries: [] })).toEqual([]);
    expect(parseRecent({ version: 1, entries: 'nope' })).toEqual([]);
    expect(parseRecent({ version: 1 })).toEqual([]);
  });

  it('drops entries that are not usable and keeps the rest', () => {
    const parsed = parseRecent({
      version: 1,
      entries: [
        null,
        { path: 42 },
        { path: 'relative.xlsx', openedAt: 1 },
        { path: '/home/user/a.xlsx', openedAt: 'soon' },
        { path: '/home/user/b.xlsx', openedAt: 7 },
      ],
    });
    expect(parsed).toEqual([
      { path: '/home/user/a.xlsx', openedAt: 0 },
      { path: '/home/user/b.xlsx', openedAt: 7 },
    ]);
  });

  it('deduplicates a hand-edited file that lists the same path twice', () => {
    const store = memoryStore({
      version: 1,
      entries: [
        { path: '/home/user/a.xlsx', openedAt: 2 },
        { path: '/home/user/a.xlsx', openedAt: 1 },
      ],
    });
    expect(new RecentFiles(store, 'linux').paths()).toEqual(['/home/user/a.xlsx']);
  });
});

describe('recentMenuLabels', () => {
  it('shows the bare file name when it is unambiguous', () => {
    expect(recentMenuLabels(['/home/user/a.xlsx', '/home/user/b.xlsx'])).toEqual(['a.xlsx', 'b.xlsx']);
  });

  it('adds the parent directory when two files share a name', () => {
    expect(recentMenuLabels(['/home/user/2024/q3.xlsx', '/home/user/2025/q3.xlsx'])).toEqual([
      'q3.xlsx — 2024',
      'q3.xlsx — 2025',
    ]);
  });

  it('leaves an unrelated third entry alone', () => {
    const labels = recentMenuLabels(['/a/1/q.xlsx', '/a/2/q.xlsx', '/a/3/other.xlsx']);
    expect(labels[2]).toBe('other.xlsx');
  });
});
