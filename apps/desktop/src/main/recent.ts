/**
 * The Recent Files list.
 *
 * Deliberately more careful than the feature deserves, because the failure
 * modes are all user-visible: duplicates that differ only in path separator, an
 * entry that silently reopens a file from a network share that is no longer
 * mounted, and a list that grows until the File menu is taller than the screen.
 *
 * Path comparison is case-sensitive on POSIX and case-insensitive on Windows
 * and macOS, matching how those filesystems actually behave. Getting this wrong
 * is what produces two "Budget.xlsx" entries pointing at the same file.
 *
 * The store is injected, so the whole class is testable without touching disk.
 */

import { basename, dirname } from 'node:path';
import type { JsonStore } from './store.js';
import { checkPath } from './paths.js';

export interface RecentEntry {
  path: string;
  /** Milliseconds since the epoch, supplied by the caller so tests are stable. */
  openedAt: number;
}

export const RECENT_LIMIT = 12;

/** Whether path comparison folds case, decided by platform rather than guessed. */
export function foldsCase(platform: NodeJS.Platform): boolean {
  return platform === 'win32' || platform === 'darwin';
}

function key(path: string, platform: NodeJS.Platform): string {
  return foldsCase(platform) ? path.toLowerCase() : path;
}

interface Persisted {
  version: 1;
  entries: RecentEntry[];
}

/** Accept only what we recognise; anything else is treated as an empty list. */
export function parseRecent(raw: unknown): RecentEntry[] {
  if (typeof raw !== 'object' || raw === null) return [];
  const record = raw as Partial<Persisted>;
  if (record.version !== 1 || !Array.isArray(record.entries)) return [];
  const out: RecentEntry[] = [];
  for (const entry of record.entries) {
    if (typeof entry !== 'object' || entry === null) continue;
    const candidate = entry as Partial<RecentEntry>;
    // The list is data that has been on disk, so it goes through the same path
    // check as anything arriving over IPC: a hand-edited file is untrusted input.
    const path = checkPath(candidate.path, { write: false });
    if (!path.ok) continue;
    const openedAt = typeof candidate.openedAt === 'number' && Number.isFinite(candidate.openedAt)
      ? candidate.openedAt
      : 0;
    out.push({ path: path.path, openedAt });
  }
  return out;
}

export class RecentFiles {
  private entries: RecentEntry[];

  constructor(
    private readonly store: JsonStore,
    private readonly platform: NodeJS.Platform = process.platform,
    private readonly limit: number = RECENT_LIMIT,
  ) {
    this.entries = dedupe(parseRecent(store.read()), platform).slice(0, limit);
  }

  list(): RecentEntry[] {
    return this.entries.map((e) => ({ ...e }));
  }

  paths(): string[] {
    return this.entries.map((e) => e.path);
  }

  /** Most recent first; re-adding an existing path moves it to the top. */
  add(path: string, openedAt: number): void {
    const checked = checkPath(path, { write: false });
    if (!checked.ok) return;
    const k = key(checked.path, this.platform);
    this.entries = [
      { path: checked.path, openedAt },
      ...this.entries.filter((e) => key(e.path, this.platform) !== k),
    ].slice(0, this.limit);
    this.persist();
  }

  remove(path: string): void {
    const k = key(path, this.platform);
    const before = this.entries.length;
    this.entries = this.entries.filter((e) => key(e.path, this.platform) !== k);
    if (this.entries.length !== before) this.persist();
  }

  clear(): void {
    this.entries = [];
    this.persist();
  }

  /**
   * Drop entries whose file has gone. Called at startup with a real `exists`,
   * because an entry that opens an error dialog is worse than a shorter menu.
   */
  prune(exists: (path: string) => boolean): number {
    const before = this.entries.length;
    this.entries = this.entries.filter((e) => exists(e.path));
    const dropped = before - this.entries.length;
    if (dropped > 0) this.persist();
    return dropped;
  }

  private persist(): void {
    const payload: Persisted = { version: 1, entries: this.entries };
    this.store.write(payload);
  }
}

function dedupe(entries: RecentEntry[], platform: NodeJS.Platform): RecentEntry[] {
  const seen = new Set<string>();
  const out: RecentEntry[] = [];
  for (const entry of entries) {
    const k = key(entry.path, platform);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(entry);
  }
  return out;
}

/**
 * Menu labels. Two files with the same name in different directories are
 * distinguished by their parent directory rather than by the full path, which
 * would make the menu unreadably wide on the deep paths people actually have.
 */
export function recentMenuLabels(paths: string[]): string[] {
  const counts = new Map<string, number>();
  for (const p of paths) counts.set(basename(p), (counts.get(basename(p)) ?? 0) + 1);
  return paths.map((p) => {
    const name = basename(p);
    if ((counts.get(name) ?? 0) < 2) return name;
    const parent = basename(dirname(p));
    return parent ? `${name} — ${parent}` : name;
  });
}
