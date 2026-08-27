/**
 * Cell formatting model and style interning.
 *
 * Excel does not store formatting per cell; it stores a table of formats and
 * gives each cell an index into it (the `s` attribute on `<c>`, pointing into
 * `cellXfs`). A 1M-row sheet with one uniform format therefore costs one style
 * record, not a million. We keep that design: styles are immutable, deduplicated
 * by structural key, and referenced by a small integer everywhere else.
 */

export type HorizontalAlign =
  | 'general'
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';

export type VerticalAlign = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export type BorderStyle =
  | 'none'
  | 'thin'
  | 'medium'
  | 'thick'
  | 'double'
  | 'hair'
  | 'dotted'
  | 'dashed'
  | 'dashDot'
  | 'dashDotDot'
  | 'mediumDashed'
  | 'mediumDashDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export type PatternType =
  | 'none'
  | 'solid'
  | 'gray125'
  | 'darkGray'
  | 'mediumGray'
  | 'lightGray'
  | 'darkHorizontal'
  | 'darkVertical'
  | 'darkDown'
  | 'darkUp'
  | 'darkGrid'
  | 'darkTrellis'
  | 'lightHorizontal'
  | 'lightVertical'
  | 'lightDown'
  | 'lightUp'
  | 'lightGrid'
  | 'lightTrellis';

/**
 * A colour, in whichever of Excel's four encodings the file used.
 *
 * We keep the original encoding rather than flattening everything to RGB on
 * read, because theme and indexed colours must survive a round trip: resolving
 * a theme colour to RGB and writing that back would silently break the file's
 * response to a theme change.
 */
export type Color =
  | { kind: 'rgb'; argb: string }
  | { kind: 'theme'; theme: number; tint?: number }
  | { kind: 'indexed'; index: number }
  | { kind: 'auto' };

export interface Font {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  strike?: boolean;
  color?: Color;
  vertAlign?: 'baseline' | 'superscript' | 'subscript';
  /** Font family / charset / scheme, preserved for faithful round-tripping. */
  family?: number;
  charset?: number;
  scheme?: 'none' | 'major' | 'minor';
}

export interface Fill {
  pattern?: PatternType;
  fg?: Color;
  bg?: Color;
  gradient?: {
    type: 'linear' | 'path';
    degree?: number;
    stops: { position: number; color: Color }[];
  };
}

export interface BorderEdge {
  style?: BorderStyle;
  color?: Color;
}

export interface Border {
  left?: BorderEdge;
  right?: BorderEdge;
  top?: BorderEdge;
  bottom?: BorderEdge;
  diagonal?: BorderEdge;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

export interface Alignment {
  horizontal?: HorizontalAlign;
  vertical?: VerticalAlign;
  wrapText?: boolean;
  /** Degrees, 0-180, where 255 means "stacked vertically" in the file format. */
  textRotation?: number;
  indent?: number;
  shrinkToFit?: boolean;
  readingOrder?: 0 | 1 | 2;
}

export interface Protection {
  locked?: boolean;
  hidden?: boolean;
}

/** A complete cell format: the thing `cellXfs` entries describe. */
export interface CellStyle {
  /** Index into the number-format table, or a literal format code. */
  numFmtId?: number;
  numFmt?: string;
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  protection?: Protection;
  /** Index of the named cell style (`cellStyleXfs`) this format is based on. */
  xfId?: number;
}

/** Opaque handle into the style table. 0 is always the default format. */
export type StyleId = number;

export const DEFAULT_STYLE_ID: StyleId = 0;

/**
 * Deduplicating store of cell formats.
 *
 * Interning is by a canonical JSON key. Building a string per lookup sounds
 * expensive, but it happens only when a *new* format is applied, not per cell,
 * and it makes structural equality exact without hand-writing a deep comparison
 * for every nested option.
 */
export class StyleTable {
  private readonly styles: CellStyle[] = [];
  private readonly index = new Map<string, StyleId>();

  constructor() {
    // Slot 0 is the empty default format, matching Excel's own convention.
    this.styles.push({});
    this.index.set(canonicalKey({}), DEFAULT_STYLE_ID);
  }

  get size(): number {
    return this.styles.length;
  }

  /** Intern a format, returning its id. Identical formats share an id. */
  intern(style: CellStyle): StyleId {
    const key = canonicalKey(style);
    const existing = this.index.get(key);
    if (existing !== undefined) return existing;
    const id = this.styles.length;
    this.styles.push(Object.freeze(structuredClone(style)));
    this.index.set(key, id);
    return id;
  }

  get(id: StyleId): CellStyle {
    return this.styles[id] ?? {};
  }

  /** Apply a partial change on top of an existing style, returning the new id. */
  derive(base: StyleId, patch: Partial<CellStyle>): StyleId {
    const merged: CellStyle = { ...this.get(base) };
    for (const [k, v] of Object.entries(patch)) {
      if (v === undefined) {
        delete (merged as Record<string, unknown>)[k];
      } else if (isPlainObject(v) && isPlainObject((merged as Record<string, unknown>)[k])) {
        // Merge one level down so `derive(id, { font: { bold: true } })` keeps
        // the existing font name and size instead of replacing the whole font.
        (merged as Record<string, unknown>)[k] = {
          ...((merged as Record<string, unknown>)[k] as object),
          ...v,
        };
      } else {
        (merged as Record<string, unknown>)[k] = v;
      }
    }
    return this.intern(merged);
  }

  /** All interned styles in id order, for serialisation. */
  all(): readonly CellStyle[] {
    return this.styles;
  }
}

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

/**
 * Stable structural key: object keys sorted, undefined dropped, so that
 * `{bold: true, size: 11}` and `{size: 11, bold: true, italic: undefined}`
 * intern to the same id.
 */
function canonicalKey(value: unknown): string {
  return JSON.stringify(canonicalise(value));
}

function canonicalise(value: unknown): unknown {
  if (value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return value.map(canonicalise);
  const out: Record<string, unknown> = {};
  for (const key of Object.keys(value as Record<string, unknown>).sort()) {
    const v = (value as Record<string, unknown>)[key];
    if (v !== undefined) out[key] = canonicalise(v);
  }
  return out;
}
