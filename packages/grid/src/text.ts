/**
 * Text measurement and wrapping.
 *
 * measureText is the single dominant cost in a canvas grid. A 40x20 viewport is
 * 800 cells; measuring each one, plus each wrapped line, plus each width probe
 * while fitting text to a column, is several thousand calls per frame, and the
 * browser cannot cache them for us because the answer depends on the current
 * ctx.font. Profiles of naive canvas grids routinely show most of the frame
 * spent inside measureText alone.
 *
 * So every measurement goes through a cache keyed by (font, text). Two further
 * details matter as much as the cache itself:
 *
 *   Setting ctx.font is not free either - it re-parses a CSS font shorthand - so
 *   the current font is tracked and assigned only when it actually changes.
 *
 *   The cache is bounded and evicts wholesale rather than keeping an LRU list.
 *   Scrolling through a large sheet inserts thousands of distinct strings, and
 *   an unbounded map is a slow leak; a generational clear costs one frame's
 *   re-measurement every few hundred thousand entries, which is invisible, while
 *   an LRU costs a pointer update on every hit, which is not.
 */

/** The slice of CanvasRenderingContext2D that measurement needs. */
export interface MeasuringContext {
  font: string;
  measureText(text: string): { width: number };
}

export interface MeasureStats {
  readonly hits: number;
  readonly misses: number;
  readonly size: number;
  readonly fontSwitches: number;
  readonly clears: number;
}

/** ASCII unit separator: cannot occur in a font shorthand, so keys are unambiguous. */
const SEPARATOR = String.fromCharCode(31);

export class TextMeasureCache {
  private readonly widths = new Map<string, number>();
  private currentFont = '';
  private hitCount = 0;
  private missCount = 0;
  private fontSwitchCount = 0;
  private clearCount = 0;

  constructor(readonly limit = 100_000) {}

  get stats(): MeasureStats {
    return {
      hits: this.hitCount,
      misses: this.missCount,
      size: this.widths.size,
      fontSwitches: this.fontSwitchCount,
      clears: this.clearCount,
    };
  }

  /** Assign ctx.font only when it differs, and remember that it was assigned. */
  useFont(ctx: MeasuringContext, font: string): void {
    if (this.currentFont === font) return;
    ctx.font = font;
    this.currentFont = font;
    this.fontSwitchCount++;
  }

  /**
   * Call before a frame when something outside this cache may have changed
   * ctx.font - a save/restore pair, or another painter sharing the context.
   */
  invalidateFont(): void {
    this.currentFont = '';
  }

  measure(ctx: MeasuringContext, text: string, font: string): number {
    if (text === '') return 0;
    const key = font + SEPARATOR + text;
    const hit = this.widths.get(key);
    if (hit !== undefined) {
      this.hitCount++;
      return hit;
    }
    this.missCount++;
    this.useFont(ctx, font);
    const width = ctx.measureText(text).width;
    if (this.widths.size >= this.limit) {
      this.widths.clear();
      this.clearCount++;
    }
    this.widths.set(key, width);
    return width;
  }

  clear(): void {
    this.widths.clear();
    this.currentFont = '';
    this.hitCount = 0;
    this.missCount = 0;
    this.fontSwitchCount = 0;
    this.clearCount = 0;
  }

  /**
   * Greedy word wrap. Excel breaks on spaces, then breaks a word that is still
   * too wide mid-character rather than letting it escape the cell, and keeps
   * explicit newlines (entered with Alt+Enter) as hard breaks.
   */
  wrap(ctx: MeasuringContext, text: string, font: string, maxWidth: number): string[] {
    if (text === '') return [''];
    const out: string[] = [];
    for (const paragraph of text.split('\n')) {
      if (paragraph === '') {
        out.push('');
        continue;
      }
      if (maxWidth <= 0 || this.measure(ctx, paragraph, font) <= maxWidth) {
        out.push(paragraph);
        continue;
      }
      let line = '';
      for (const word of splitWords(paragraph)) {
        const candidate = line === '' ? word : line + word;
        if (this.measure(ctx, candidate, font) <= maxWidth) {
          line = candidate;
          continue;
        }
        if (line !== '') {
          out.push(line.trimEnd());
          line = word.trimStart();
          if (this.measure(ctx, line, font) <= maxWidth) continue;
        } else {
          line = word;
        }
        // A single word wider than the column: break it character by character.
        while (line.length > 1 && this.measure(ctx, line, font) > maxWidth) {
          const cut = this.breakPoint(ctx, line, font, maxWidth);
          out.push(line.slice(0, cut));
          line = line.slice(cut);
        }
      }
      if (line !== '') out.push(line.trimEnd());
    }
    return out.length > 0 ? out : [''];
  }

  /** Largest prefix length of `text` that fits, never less than one character. */
  breakPoint(ctx: MeasuringContext, text: string, font: string, maxWidth: number): number {
    let lo = 1;
    let hi = text.length;
    while (lo < hi) {
      const mid = (lo + hi + 1) >>> 1;
      if (this.measure(ctx, text.slice(0, mid), font) <= maxWidth) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}

/** Split keeping trailing spaces attached to their word, so widths stay faithful. */
function splitWords(s: string): string[] {
  const parts = s.match(/\S+\s*|\s+/g);
  return parts ?? [s];
}

export interface FontSpec {
  family?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
}

/**
 * Build the CSS font shorthand. Order is fixed by the CSS grammar: style,
 * weight, size, family. Getting it wrong makes the whole declaration invalid and
 * the canvas silently falls back to 10px sans-serif, which reads as a font bug
 * rather than the syntax bug it is.
 */
export function fontString(
  spec: FontSpec,
  defaults: { family: string; size: number },
  zoom = 1,
): string {
  const size = (spec.size ?? defaults.size) * zoom;
  const family = spec.family ? quoteFamily(spec.family) : defaults.family;
  let prefix = '';
  if (spec.italic) prefix += 'italic ';
  if (spec.bold) prefix += 'bold ';
  return `${prefix}${round2(size)}px ${family}`;
}

function quoteFamily(name: string): string {
  return /^[A-Za-z0-9_-]+$/.test(name) ? name : `"${name.replace(/"/g, '')}"`;
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

/**
 * Approximate ascent/descent from the font size.
 *
 * TextMetrics carries real font bounding boxes, but only in browsers new enough
 * to have shipped them and never in a headless test context, so layout cannot
 * depend on them. The 0.8/0.2 split is the usual approximation for Latin faces
 * and is what the vertical-alignment maths uses.
 */
export function fontHeights(sizePx: number): { ascent: number; descent: number; line: number } {
  return { ascent: sizePx * 0.8, descent: sizePx * 0.2, line: sizePx * 1.28 };
}

/** Pixel size out of a CSS font shorthand, for line-height maths. */
export function fontSizeOf(font: string, fallback: number): number {
  const m = /(\d+(?:\.\d+)?)px/.exec(font);
  return m ? Number(m[1]) : fallback;
}
