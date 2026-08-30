/**
 * The command palette's item model and search.
 *
 * A spreadsheet's navigation problem is not "which menu is that under", it is
 * "which of my sixty sheets was the assumptions one". So the palette searches
 * three kinds of thing in one box - commands, sheets and defined names - and
 * ranks them together, because forcing a user to pick the right search mode
 * before typing is asking them to answer the question they came to ask.
 *
 * Ranking puts sheets and names slightly above commands on an equal text match:
 * commands are discoverable from the toolbar, and a sheet buried at tab 47 is
 * not discoverable at all.
 */

import { fuzzyFilter, type FuzzyResult } from './fuzzy.js';

export type PaletteItemKind = 'command' | 'sheet' | 'name';

export interface PaletteItem {
  id: string;
  kind: PaletteItemKind;
  title: string;
  /** Shown to the right: a category, a sheet's position, a name's target. */
  detail?: string;
  /** Extra words that should match without appearing in the title. */
  keywords?: string;
  shortcut?: string;
  /** False for a command that cannot run right now, e.g. undo with no history. */
  enabled?: boolean;
}

const KIND_BOOST: Readonly<Record<PaletteItemKind, number>> = Object.freeze({
  sheet: 12,
  name: 8,
  command: 0,
});

export interface PaletteSearchOptions {
  limit?: number;
  /** Item ids used recently, most recent first; they float up on an equal match. */
  recent?: readonly string[];
  /** Restrict to one kind, for the `>` and `@` prefixes. */
  kind?: PaletteItemKind;
}

const RECENT_BOOST = 20;
const RECENT_DECAY = 4;

/**
 * Search the palette.
 *
 * A leading `>` restricts to commands and `@` to sheets, matching the convention
 * people already know from code editors, and the prefix is stripped before it
 * reaches the matcher so `>sum` does not try to find a command containing `>`.
 */
export function searchPalette(
  items: readonly PaletteItem[],
  rawQuery: string,
  options: PaletteSearchOptions = {},
): FuzzyResult<PaletteItem>[] {
  let query = rawQuery;
  let kind = options.kind;
  if (query.startsWith('>')) {
    kind = 'command';
    query = query.slice(1);
  } else if (query.startsWith('@')) {
    kind = 'sheet';
    query = query.slice(1);
  }
  query = query.trim();

  const pool = kind ? items.filter((i) => i.kind === kind) : items;
  const recent = options.recent ?? [];

  return fuzzyFilter(pool, query, {
    key: (item) => item.title,
    extra: (item) => [item.keywords, item.detail].filter(Boolean).join(' ') || undefined,
    ...(options.limit === undefined ? {} : { limit: options.limit }),
    boost: (item) => {
      let boost = KIND_BOOST[item.kind];
      const recentIndex = recent.indexOf(item.id);
      if (recentIndex >= 0) boost += Math.max(0, RECENT_BOOST - recentIndex * RECENT_DECAY);
      if (item.enabled === false) boost -= 50;
      return boost;
    },
  });
}

/** Move a selection index inside a list, wrapping at both ends. */
export function moveIndex(current: number, delta: number, length: number): number {
  if (length === 0) return 0;
  return (((current + delta) % length) + length) % length;
}

/** Push an id onto a most-recently-used list, de-duplicated and capped. */
export function pushRecent(recent: readonly string[], id: string, cap = 12): string[] {
  return [id, ...recent.filter((r) => r !== id)].slice(0, cap);
}
