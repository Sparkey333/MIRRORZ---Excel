import { describe, expect, it } from 'vitest';
import { fuzzyFilter, fuzzyMatch, highlightSegments } from '../src/renderer/model/fuzzy.js';

describe('fuzzyMatch', () => {
  it('matches an empty query against anything with a neutral score', () => {
    expect(fuzzyMatch('', 'Anything')).toEqual({ score: 0, positions: [] });
  });

  it('rejects a query that is not a subsequence', () => {
    expect(fuzzyMatch('xyz', 'Summary')).toBeNull();
  });

  it('rejects any query against an empty target', () => {
    expect(fuzzyMatch('a', '')).toBeNull();
  });

  it('reports the matched positions', () => {
    expect(fuzzyMatch('sm', 'Summary')?.positions).toEqual([0, 2]);
  });

  it('is case insensitive', () => {
    expect(fuzzyMatch('SUM', 'summary')).not.toBeNull();
  });

  it('scores an exact match above a prefix match', () => {
    const exact = fuzzyMatch('sum', 'sum')!.score;
    const prefix = fuzzyMatch('sum', 'summary')!.score;
    expect(exact).toBeGreaterThan(prefix);
  });

  it('scores a prefix match above a mid-word match', () => {
    const prefix = fuzzyMatch('sum', 'summary')!.score;
    const middle = fuzzyMatch('sum', 'the sum of it')!.score;
    expect(prefix).toBeGreaterThan(middle);
  });

  it('prefers word boundaries when the same letters appear twice', () => {
    // "pl" should land on Profit and Loss, not Profit and the l inside it.
    expect(fuzzyMatch('pl', 'Profit and Loss')?.positions).toEqual([0, 11]);
  });

  it('rewards consecutive characters over scattered ones', () => {
    const consecutive = fuzzyMatch('abc', 'abcdefgh')!.score;
    const scattered = fuzzyMatch('abc', 'axbxcxgh')!.score;
    expect(consecutive).toBeGreaterThan(scattered);
  });

  it('finds camel-case initials', () => {
    expect(fuzzyMatch('cf', 'cashFlow')?.positions).toEqual([0, 4]);
  });

  it('penalises a match that starts late', () => {
    const early = fuzzyMatch('x', 'xylophone')!.score;
    const late = fuzzyMatch('x', 'aaaaaaaax')!.score;
    expect(early).toBeGreaterThan(late);
  });
});

describe('fuzzyFilter', () => {
  const sheets = ['Summary', 'Q1 Sales', 'Q2 Sales', 'Assumptions', 'Sensitivity Analysis'];

  it('returns everything for an empty query, in the original order', () => {
    const results = fuzzyFilter(sheets, '', { key: (s) => s });
    expect(results.map((r) => r.item)).toEqual(sheets);
  });

  it('drops non-matching items', () => {
    const results = fuzzyFilter(sheets, 'zzz', { key: (s) => s });
    expect(results).toEqual([]);
  });

  it('ranks the closest match first', () => {
    const results = fuzzyFilter(sheets, 'sum', { key: (s) => s });
    expect(results[0]?.item).toBe('Summary');
  });

  it('finds a multi-word target by its initials', () => {
    const results = fuzzyFilter(sheets, 'sa', { key: (s) => s });
    expect(results[0]?.item).toBe('Sensitivity Analysis');
  });

  it('breaks ties on the shorter target', () => {
    const results = fuzzyFilter(['Sales Detail', 'Sales'], 'sales', { key: (s) => s });
    expect(results[0]?.item).toBe('Sales');
  });

  it('honours the limit', () => {
    const results = fuzzyFilter(sheets, 's', { key: (s) => s, limit: 2 });
    expect(results).toHaveLength(2);
  });

  it('matches secondary text but ranks it below a title match', () => {
    const items = [
      { title: 'Bold', hint: 'nothing' },
      { title: 'Weight', hint: 'bold heavy' },
    ];
    const results = fuzzyFilter(items, 'bold', {
      key: (i) => i.title,
      extra: (i) => i.hint,
    });
    expect(results.map((r) => r.item.title)).toEqual(['Bold', 'Weight']);
  });

  it('applies a boost', () => {
    const results = fuzzyFilter(['alpha', 'beta'], 'a', {
      key: (s) => s,
      boost: (s) => (s === 'beta' ? 1000 : 0),
    });
    expect(results[0]?.item).toBe('beta');
  });
});

describe('highlightSegments', () => {
  it('splits a target into matched and unmatched runs', () => {
    expect(highlightSegments('Summary', [0, 1, 2])).toEqual([
      { text: 'Sum', match: true },
      { text: 'mary', match: false },
    ]);
  });

  it('returns one plain segment when nothing matched', () => {
    expect(highlightSegments('Summary', [])).toEqual([{ text: 'Summary', match: false }]);
  });

  it('handles a match at the end', () => {
    expect(highlightSegments('ab', [1])).toEqual([
      { text: 'a', match: false },
      { text: 'b', match: true },
    ]);
  });

  it('returns nothing for an empty target', () => {
    expect(highlightSegments('', [])).toEqual([]);
  });
});
