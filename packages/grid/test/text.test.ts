import { describe, expect, it } from 'vitest';

import { TextMeasureCache, fontHeights, fontSizeOf, fontString } from '../src/text.js';
import { FakeContext } from './fake-canvas.js';

const FONT = '10px sans-serif';
/** FakeContext measures 0.5 * fontSize per character. */
const CHAR = 5;

describe('TextMeasureCache', () => {
  it('measures once and serves the rest from the cache', () => {
    const ctx = new FakeContext();
    const cache = new TextMeasureCache();
    expect(cache.measure(ctx, 'abc', FONT)).toBe(3 * CHAR);
    cache.measure(ctx, 'abc', FONT);
    cache.measure(ctx, 'abc', FONT);
    expect(ctx.measureCount).toBe(1);
    expect(cache.stats).toMatchObject({ hits: 2, misses: 1 });
  });

  it('keys the cache by font as well as text', () => {
    const ctx = new FakeContext();
    const cache = new TextMeasureCache();
    expect(cache.measure(ctx, 'abc', '10px x')).toBe(15);
    expect(cache.measure(ctx, 'abc', '20px x')).toBe(30);
    expect(ctx.measureCount).toBe(2);
  });

  it('never confuses a font/text split', () => {
    const ctx = new FakeContext();
    const cache = new TextMeasureCache();
    const a = cache.measure(ctx, 'b', '10px a');
    const b = cache.measure(ctx, '', '10px ab');
    expect(a).toBe(CHAR);
    expect(b).toBe(0);
  });

  it('costs nothing for the empty string', () => {
    const ctx = new FakeContext();
    const cache = new TextMeasureCache();
    expect(cache.measure(ctx, '', FONT)).toBe(0);
    expect(ctx.measureCount).toBe(0);
  });

  it('assigns ctx.font only when it changes', () => {
    const ctx = new FakeContext();
    const cache = new TextMeasureCache();
    cache.measure(ctx, 'a', FONT);
    cache.measure(ctx, 'b', FONT);
    cache.measure(ctx, 'c', '12px other');
    expect(cache.stats.fontSwitches).toBe(2);
  });

  it('re-assigns the font after an external invalidation', () => {
    const ctx = new FakeContext();
    const cache = new TextMeasureCache();
    cache.useFont(ctx, FONT);
    cache.invalidateFont();
    cache.useFont(ctx, FONT);
    expect(cache.stats.fontSwitches).toBe(2);
  });

  it('evicts wholesale when it reaches its limit', () => {
    const ctx = new FakeContext();
    const cache = new TextMeasureCache(4);
    for (let i = 0; i < 6; i++) cache.measure(ctx, `t${i}`, FONT);
    expect(cache.stats.clears).toBe(1);
    expect(cache.stats.size).toBeLessThanOrEqual(4);
  });

  it('resets its counters on clear', () => {
    const ctx = new FakeContext();
    const cache = new TextMeasureCache();
    cache.measure(ctx, 'a', FONT);
    cache.clear();
    expect(cache.stats).toMatchObject({ hits: 0, misses: 0, size: 0 });
  });
});

describe('TextMeasureCache.wrap', () => {
  const ctx = new FakeContext();
  const cache = new TextMeasureCache();

  it('returns a single line when it fits', () => {
    expect(cache.wrap(ctx, 'abc', FONT, 100)).toEqual(['abc']);
  });

  it('breaks on spaces', () => {
    expect(cache.wrap(ctx, 'aaa bbb ccc', FONT, 6 * CHAR)).toEqual(['aaa', 'bbb', 'ccc']);
  });

  it('fits as many words per line as it can', () => {
    expect(cache.wrap(ctx, 'aa bb cc dd', FONT, 6 * CHAR)).toEqual(['aa bb', 'cc dd']);
  });

  it('breaks a word that is wider than the column', () => {
    const lines = cache.wrap(ctx, 'abcdefghij', FONT, 4 * CHAR);
    expect(lines).toEqual(['abcd', 'efgh', 'ij']);
  });

  it('honours a hard line break', () => {
    expect(cache.wrap(ctx, 'a\nb', FONT, 100)).toEqual(['a', 'b']);
  });

  it('keeps an empty paragraph as an empty line', () => {
    expect(cache.wrap(ctx, 'a\n\nb', FONT, 100)).toEqual(['a', '', 'b']);
  });

  it('gives up rather than looping on a zero-width column', () => {
    expect(cache.wrap(ctx, 'abc', FONT, 0)).toEqual(['abc']);
  });

  it('returns one empty line for empty text', () => {
    expect(cache.wrap(ctx, '', FONT, 50)).toEqual(['']);
  });

  it('finds the longest prefix that fits', () => {
    expect(cache.breakPoint(ctx, 'abcdefgh', FONT, 3 * CHAR)).toBe(3);
    expect(cache.breakPoint(ctx, 'abcdefgh', FONT, 0)).toBe(1);
  });
});

describe('fontString', () => {
  const defaults = { family: 'Calibri, sans-serif', size: 11 };

  it('emits the CSS shorthand in grammar order', () => {
    expect(fontString({ bold: true, italic: true, size: 12, family: 'Arial' }, defaults)).toBe(
      'italic bold 12px Arial',
    );
  });

  it('falls back to the theme defaults', () => {
    expect(fontString({}, defaults)).toBe('11px Calibri, sans-serif');
  });

  it('quotes a family name that needs it', () => {
    expect(fontString({ family: 'Times New Roman' }, defaults)).toBe('11px "Times New Roman"');
  });

  it('applies zoom to the size', () => {
    expect(fontString({ size: 10 }, defaults, 1.5)).toBe('15px Calibri, sans-serif');
  });

  it('rounds a fractional size to two places', () => {
    expect(fontString({ size: 11 }, defaults, 1.333)).toBe('14.66px Calibri, sans-serif');
  });

  it('reads the size back out of a shorthand', () => {
    expect(fontSizeOf('bold 12.5px Arial', 11)).toBe(12.5);
    expect(fontSizeOf('inherit', 11)).toBe(11);
  });

  it('derives ascent, descent and line height from the size', () => {
    const h = fontHeights(10);
    expect(h.ascent).toBe(8);
    expect(h.descent).toBe(2);
    expect(h.line).toBeGreaterThan(h.ascent + h.descent);
  });
});
