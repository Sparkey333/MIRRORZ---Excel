/**
 * Subsequence fuzzy matching with ranking.
 *
 * Three different surfaces need this - the sheet explorer, the command palette
 * and the function autocomplete - and they need it to agree, because a user who
 * learns that "pl" finds "Profit & Loss" in one list expects the same in the
 * next. So the scorer lives here once, and every list is a thin call over it.
 *
 * The ranking rules are chosen from how people actually type into a filter box:
 * they type the initials of a multi-word name, or the start of one word, and
 * they type lower case. So a match at a word boundary is worth far more than a
 * match in the middle of a word, a run of adjacent characters is worth more than
 * the same characters scattered, and an exact prefix beats everything. Length is
 * a tie-breaker rather than a term, because otherwise short irrelevant names
 * outrank the long name the user was clearly aiming at.
 */

export interface FuzzyMatch {
  score: number;
  /** Indices in the target that the query matched, for highlighting. */
  positions: number[];
}

/** Characters after which the next character starts a new "word". */
function isBoundary(ch: string): boolean {
  return ch === ' ' || ch === '_' || ch === '-' || ch === '.' || ch === '!' || ch === '/' || ch === '(';
}

const SCORE_EXACT = 200;
const SCORE_PREFIX = 120;
const SCORE_BOUNDARY = 30;
const SCORE_CAMEL = 25;
const SCORE_CONSECUTIVE = 18;
const SCORE_CASE_MATCH = 3;
const PENALTY_LEADING = 2;
const PENALTY_GAP = 3;
const MAX_LEADING_PENALTY = 20;

/**
 * Score `query` against `target`.
 *
 * Returns null when the query is not a subsequence of the target at all; an
 * empty query matches everything with score 0, which is what a filter box wants
 * before the user has typed.
 */
export function fuzzyMatch(query: string, target: string): FuzzyMatch | null {
  if (query === '') return { score: 0, positions: [] };
  if (target === '') return null;

  const q = query.toLowerCase();
  const t = target.toLowerCase();

  // Greedy left-to-right is wrong for targets that repeat a query character
  // ("Sales" against "ss"), so walk from the left but prefer boundary matches
  // when the same character appears again shortly after.
  const positions: number[] = [];
  let ti = 0;
  for (let qi = 0; qi < q.length; qi++) {
    const ch = q[qi]!;
    let found = -1;
    for (let i = ti; i < t.length; i++) {
      if (t[i] !== ch) continue;
      if (found < 0) found = i;
      // A later occurrence that starts a word beats an earlier one buried
      // mid-word: "pl" should land on Profit / Loss, not Profit / "l" of Profit.
      if (isWordStart(target, i)) {
        found = i;
        break;
      }
      // Only look a little way ahead; scanning the whole string for a boundary
      // would let a distant match win over an obviously adjacent one.
      if (i > found + 8) break;
    }
    if (found < 0) return null;
    positions.push(found);
    ti = found + 1;
  }

  let score = 0;
  if (t === q) score += SCORE_EXACT;
  if (t.startsWith(q)) score += SCORE_PREFIX;

  score -= Math.min(positions[0]! * PENALTY_LEADING, MAX_LEADING_PENALTY);

  for (let i = 0; i < positions.length; i++) {
    const pos = positions[i]!;
    if (isWordStart(target, pos)) {
      score += SCORE_BOUNDARY;
      if (pos > 0 && isUpper(target[pos]!) && isLower(target[pos - 1]!)) score += SCORE_CAMEL;
    }
    if (query[i] === target[pos]) score += SCORE_CASE_MATCH;
    if (i > 0) {
      const gap = pos - positions[i - 1]! - 1;
      if (gap === 0) score += SCORE_CONSECUTIVE;
      // Skipping a whole word to land on the start of the next one is not a
      // gap, it is how initials work: "sa" meaning Sensitivity Analysis must not
      // be punished for the eleven characters it stepped over.
      else if (!isWordStart(target, pos)) score -= Math.min(gap, 6) * PENALTY_GAP;
    }
  }

  return { score, positions };
}

function isWordStart(target: string, i: number): boolean {
  if (i === 0) return true;
  const prev = target[i - 1]!;
  if (isBoundary(prev)) return true;
  // A capital following a lower-case letter starts a word in CamelCase.
  return isUpper(target[i]!) && isLower(prev);
}

function isUpper(ch: string): boolean {
  return ch >= 'A' && ch <= 'Z';
}

function isLower(ch: string): boolean {
  return ch >= 'a' && ch <= 'z';
}

export interface FuzzyResult<T> {
  item: T;
  score: number;
  positions: number[];
}

export interface FuzzyFilterOptions<T> {
  /** The text to match against. */
  key: (item: T) => string;
  /**
   * Extra text that may match but scores lower - a command's category, a
   * function's summary. Keeps "money" finding CURRENCY without letting a
   * description outrank a name.
   */
  extra?: (item: T) => string | undefined;
  limit?: number;
  /** Added to the score, for surfaces that want to pin recent items. */
  boost?: (item: T) => number;
}

const EXTRA_MATCH_FACTOR = 0.25;

/**
 * Rank a list. Ties break on the shorter target and then on the original order,
 * so a filtered list does not reshuffle itself for reasons the user cannot see.
 */
export function fuzzyFilter<T>(
  items: readonly T[],
  query: string,
  options: FuzzyFilterOptions<T>,
): FuzzyResult<T>[] {
  const { key, extra, limit, boost } = options;
  const scored: (FuzzyResult<T> & { index: number; length: number })[] = [];

  for (let index = 0; index < items.length; index++) {
    const item = items[index]!;
    const target = key(item);
    let match = fuzzyMatch(query, target);
    let positions = match?.positions ?? [];
    let score = match?.score ?? 0;

    if (!match && extra) {
      const secondary = extra(item);
      if (secondary) {
        const alt = fuzzyMatch(query, secondary);
        if (alt) {
          match = alt;
          positions = [];
          score = alt.score * EXTRA_MATCH_FACTOR;
        }
      }
    }
    if (!match) continue;
    if (boost) score += boost(item);
    scored.push({ item, score, positions, index, length: target.length });
  }

  scored.sort((a, b) => b.score - a.score || a.length - b.length || a.index - b.index);
  const trimmed = limit === undefined ? scored : scored.slice(0, limit);
  return trimmed.map(({ item, score, positions }) => ({ item, score, positions }));
}

/**
 * Split a target into alternating plain and matched segments, so a component can
 * render the highlight without knowing how the scoring worked.
 */
export function highlightSegments(
  target: string,
  positions: readonly number[],
): { text: string; match: boolean }[] {
  if (positions.length === 0) return target === '' ? [] : [{ text: target, match: false }];
  const set = new Set(positions);
  const out: { text: string; match: boolean }[] = [];
  let current = '';
  let currentMatch = set.has(0);
  for (let i = 0; i < target.length; i++) {
    const match = set.has(i);
    if (match !== currentMatch) {
      if (current !== '') out.push({ text: current, match: currentMatch });
      current = '';
      currentMatch = match;
    }
    current += target[i];
  }
  if (current !== '') out.push({ text: current, match: currentMatch });
  return out;
}
