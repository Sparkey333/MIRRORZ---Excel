/**
 * styles.xml: reading and writing Excel's format tables.
 *
 * Excel's style model is two levels of indirection, and getting it wrong
 * produces files that open but look wrong. A cell's `s` attribute indexes
 * `cellXfs`. That `xf` names a font, fill, border and number format by index,
 * and also carries an `xfId` pointing into `cellStyleXfs` - the named-style
 * layer behind "Normal", "Heading 1" and friends. The `applyFont`,
 * `applyFill`, `applyNumberFormat` ... booleans decide, per attribute, whether
 * the xf's own choice wins or the named style's does.
 *
 * We flatten that into a single resolved CellStyle when reading, because the
 * rest of the app wants one answer per cell, and we keep the original table
 * around so a save can put the indirection back rather than exploding every
 * cell into its own format record.
 */

import type { Border, BorderStyle, CellStyle, Color, Fill, Font, PatternType } from '@mirrorz/core';
import { XmlReader, XmlToken, XmlWriter } from '../xml.js';

/**
 * Built-in number formats. Excel implies these rather than storing them, and
 * writing them back explicitly can make Excel reject or mis-render the file, so
 * this table is read-only knowledge: we resolve ids through it and never emit
 * a `numFmt` element for an id below 164.
 *
 * The gaps (5-8, 23-36, 41-44, 50-58) are locale-specific or reserved; Excel
 * fills them differently per language, so an unknown id resolves to General
 * rather than to a guess.
 */
export const BUILTIN_NUMBER_FORMATS: Readonly<Record<number, string>> = Object.freeze({
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
});

/** Custom formats must use ids from here up; below is reserved for built-ins. */
export const FIRST_CUSTOM_NUMFMT_ID = 164;

/**
 * The legacy 56-colour indexed palette, used by `indexed` colours and by all of
 * BIFF8. Positions 0-7 duplicate 8-15 for historical reasons, and 64/65 are the
 * "system foreground/background" sentinels rather than real colours.
 */
export const INDEXED_COLORS: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

/** Raw style tables as they appear in the file, kept for faithful writing. */
export interface StyleTables {
  numFmts: Map<number, string>;
  fonts: Font[];
  fills: Fill[];
  borders: Border[];
  /** The named-style layer. */
  cellStyleXfs: Xf[];
  /** The per-cell layer; a cell's `s` indexes this. */
  cellXfs: Xf[];
  /** Differential formats, used by conditional formatting and table styles. */
  dxfs: string[];
  /** Named cell styles (`<cellStyles>`), preserved verbatim. */
  cellStylesXml?: string;
  /** Anything else in styles.xml we did not model, kept for round-tripping. */
  preserved: Map<string, string>;
}

/** One `xf` record: indices plus the apply flags that select between layers. */
export interface Xf {
  numFmtId?: number;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  xfId?: number;
  applyNumberFormat?: boolean;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
  applyProtection?: boolean;
  alignment?: CellStyle['alignment'];
  protection?: CellStyle['protection'];
  quotePrefix?: boolean;
}

export function emptyStyleTables(): StyleTables {
  return {
    numFmts: new Map(),
    fonts: [],
    fills: [],
    borders: [],
    cellStyleXfs: [],
    cellXfs: [],
    dxfs: [],
    preserved: new Map(),
  };
}

export function parseStyles(xml: string): StyleTables {
  const tables = emptyStyleTables();
  const r = new XmlReader(xml);

  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t !== XmlToken.Open) continue;
    switch (r.localName) {
      case 'numFmts':
        readNumFmts(r, tables);
        break;
      case 'fonts':
        tables.fonts = readList(r, 'font', readFont);
        break;
      case 'fills':
        tables.fills = readList(r, 'fill', readFill);
        break;
      case 'borders':
        tables.borders = readList(r, 'border', readBorder);
        break;
      case 'cellStyleXfs':
        tables.cellStyleXfs = readList(r, 'xf', readXf);
        break;
      case 'cellXfs':
        tables.cellXfs = readList(r, 'xf', readXf);
        break;
      case 'dxfs':
        tables.dxfs = readList(r, 'dxf', (rr) => rr.readRaw());
        break;
      case 'cellStyles':
        tables.cellStylesXml = r.readRaw();
        break;
      case 'styleSheet':
        // The root; keep walking into it.
        break;
      default:
        // Anything else at the top level - tableStyles, colors, extLst - is
        // preserved verbatim so features we do not model survive a save.
        if (r.depth === 2) tables.preserved.set(r.localName, r.readRaw());
        break;
    }
  }
  return tables;
}

function readNumFmts(r: XmlReader, tables: StyleTables): void {
  if (r.isSelfClosing) return;
  const target = r.depth - 1;
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Close && r.depth === target) return;
    if (t !== XmlToken.Open || r.localName !== 'numFmt') continue;
    const id = Number(r.attr('numFmtId'));
    const code = r.attr('formatCode');
    if (Number.isFinite(id) && code !== undefined) tables.numFmts.set(id, code);
  }
}

/** Read a container of repeated child elements into an array. */
function readList<T>(r: XmlReader, childName: string, read: (r: XmlReader) => T): T[] {
  const out: T[] = [];
  if (r.isSelfClosing) return out;
  const target = r.depth - 1;
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Close && r.depth === target) return out;
    if (t === XmlToken.Open && r.localName === childName) out.push(read(r));
  }
  return out;
}

function readColor(r: XmlReader): Color | undefined {
  const rgb = r.attr('rgb');
  if (rgb) return { kind: 'rgb', argb: rgb.toUpperCase() };
  const theme = r.attr('theme');
  if (theme !== undefined) {
    const tint = r.attr('tint');
    // Theme and indexed colours are kept in their original encoding rather than
    // resolved to RGB: flattening them here would break the file's response to
    // a theme change, which is a silent data loss the user cannot see coming.
    return tint === undefined
      ? { kind: 'theme', theme: Number(theme) }
      : { kind: 'theme', theme: Number(theme), tint: Number(tint) };
  }
  const indexed = r.attr('indexed');
  if (indexed !== undefined) return { kind: 'indexed', index: Number(indexed) };
  if (r.attr('auto') !== undefined) return { kind: 'auto' };
  return undefined;
}

function readFont(r: XmlReader): Font {
  const font: Font = {};
  if (r.isSelfClosing) return font;
  const target = r.depth - 1;
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Close && r.depth === target) break;
    if (t !== XmlToken.Open) continue;
    switch (r.localName) {
      case 'name':
      case 'rFont':
        font.name = r.attr('val');
        break;
      case 'sz':
        font.size = Number(r.attr('val'));
        break;
      // A bare <b/> means true; <b val="0"/> means false.
      case 'b':
        font.bold = boolAttr(r, true);
        break;
      case 'i':
        font.italic = boolAttr(r, true);
        break;
      case 'strike':
        font.strike = boolAttr(r, true);
        break;
      case 'u':
        font.underline = (r.attr('val') as Font['underline']) ?? 'single';
        break;
      case 'color': {
        const c = readColor(r);
        if (c) font.color = c;
        break;
      }
      case 'family':
        font.family = Number(r.attr('val'));
        break;
      case 'charset':
        font.charset = Number(r.attr('val'));
        break;
      case 'scheme':
        font.scheme = r.attr('val') as Font['scheme'];
        break;
      case 'vertAlign':
        font.vertAlign = r.attr('val') as Font['vertAlign'];
        break;
      default:
        break;
    }
  }
  return font;
}

/**
 * A boolean carried by a child element's `val` attribute, as in `<b/>` (true by
 * omission) or `<b val="0"/>`.
 */
function boolAttr(r: XmlReader, whenAbsent: boolean): boolean {
  return parseBool(r.attr('val'), whenAbsent);
}

/**
 * A boolean carried by a named attribute on the element itself, as in
 * `<alignment wrapText="1"/>`. Distinct from `boolAttr` above: conflating the
 * two makes every such flag read as false, which is a silent formatting loss
 * rather than a visible error.
 */
function attrBool(r: XmlReader, name: string, whenAbsent: boolean): boolean {
  return parseBool(r.attr(name), whenAbsent);
}

function parseBool(v: string | undefined, whenAbsent: boolean): boolean {
  if (v === undefined) return whenAbsent;
  return v !== '0' && v !== 'false';
}

function readFill(r: XmlReader): Fill {
  const fill: Fill = {};
  if (r.isSelfClosing) return fill;
  const target = r.depth - 1;
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Close && r.depth === target) break;
    if (t !== XmlToken.Open) continue;
    // Read the name into a local each time: the reader mutates localName on
    // every next(), and comparing the live property inside a nested loop would
    // let the type checker narrow it to a value it can no longer have.
    const kind = r.localName;
    if (kind === 'patternFill') {
      fill.pattern = (r.attr('patternType') as PatternType) ?? 'none';
      if (r.isSelfClosing) continue;
      const inner = r.depth - 1;
      for (let u = r.next(); u !== XmlToken.EOF; u = r.next()) {
        if (u === XmlToken.Close && r.depth === inner) break;
        if (u !== XmlToken.Open) continue;
        const child = r.localName;
        if (child === 'fgColor') {
          const c = readColor(r);
          if (c) fill.fg = c;
        } else if (child === 'bgColor') {
          const c = readColor(r);
          if (c) fill.bg = c;
        }
      }
    } else if (kind === 'gradientFill') {
      const stops: { position: number; color: Color }[] = [];
      const type = (r.attr('type') as 'linear' | 'path') ?? 'linear';
      const degreeAttr = r.attr('degree');
      if (!r.isSelfClosing) {
        const inner = r.depth - 1;
        for (let u = r.next(); u !== XmlToken.EOF; u = r.next()) {
          if (u === XmlToken.Close && r.depth === inner) break;
          if (u !== XmlToken.Open || r.localName !== 'stop') continue;
          const position = Number(r.attr('position') ?? 0);
          let color: Color | undefined;
          if (!r.isSelfClosing) {
            const stopDepth = r.depth - 1;
            for (let v = r.next(); v !== XmlToken.EOF; v = r.next()) {
              if (v === XmlToken.Close && r.depth === stopDepth) break;
              const stopChild: string = r.localName;
              if (v === XmlToken.Open && stopChild === 'color') color = readColor(r);
            }
          }
          stops.push({ position, color: color ?? { kind: 'auto' } });
        }
      }
      fill.gradient =
        degreeAttr === undefined ? { type, stops } : { type, degree: Number(degreeAttr), stops };
    }
  }
  return fill;
}

const BORDER_EDGES = ['left', 'right', 'top', 'bottom', 'diagonal'] as const;

function readBorder(r: XmlReader): Border {
  const border: Border = {};
  if (r.attr('diagonalUp') !== undefined) border.diagonalUp = attrBool(r, 'diagonalUp', false);
  if (r.attr('diagonalDown') !== undefined) border.diagonalDown = attrBool(r, 'diagonalDown', false);
  if (r.isSelfClosing) return border;
  const target = r.depth - 1;
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Close && r.depth === target) break;
    if (t !== XmlToken.Open) continue;
    const edge = BORDER_EDGES.find((e) => e === r.localName);
    if (!edge) continue;
    const style = r.attr('style') as BorderStyle | undefined;
    let color: Color | undefined;
    if (!r.isSelfClosing) {
      const inner = r.depth - 1;
      for (let u = r.next(); u !== XmlToken.EOF; u = r.next()) {
        if (u === XmlToken.Close && r.depth === inner) break;
        if (u === XmlToken.Open && r.localName === 'color') color = readColor(r);
      }
    }
    if (style || color) {
      border[edge] = color ? { style: style ?? 'thin', color } : { style: style ?? 'thin' };
    }
  }
  return border;
}

function readXf(r: XmlReader): Xf {
  const xf: Xf = {};
  const num = (name: string): number | undefined => {
    const v = r.attr(name);
    return v === undefined ? undefined : Number(v);
  };
  const flag = (name: string): boolean | undefined => {
    const v = r.attr(name);
    return v === undefined ? undefined : v !== '0' && v !== 'false';
  };
  xf.numFmtId = num('numFmtId');
  xf.fontId = num('fontId');
  xf.fillId = num('fillId');
  xf.borderId = num('borderId');
  xf.xfId = num('xfId');
  xf.applyNumberFormat = flag('applyNumberFormat');
  xf.applyFont = flag('applyFont');
  xf.applyFill = flag('applyFill');
  xf.applyBorder = flag('applyBorder');
  xf.applyAlignment = flag('applyAlignment');
  xf.applyProtection = flag('applyProtection');
  xf.quotePrefix = flag('quotePrefix');

  if (r.isSelfClosing) return xf;
  const target = r.depth - 1;
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Close && r.depth === target) break;
    if (t !== XmlToken.Open) continue;
    if (r.localName === 'alignment') {
      const a: NonNullable<CellStyle['alignment']> = {};
      const h = r.attr('horizontal');
      if (h) a.horizontal = h as NonNullable<CellStyle['alignment']>['horizontal'];
      const v = r.attr('vertical');
      if (v) a.vertical = v as NonNullable<CellStyle['alignment']>['vertical'];
      if (r.attr('wrapText') !== undefined) a.wrapText = attrBool(r, 'wrapText', false);
      if (r.attr('shrinkToFit') !== undefined) a.shrinkToFit = attrBool(r, 'shrinkToFit', false);
      const rot = r.attr('textRotation');
      if (rot !== undefined) a.textRotation = Number(rot);
      const indent = r.attr('indent');
      if (indent !== undefined) a.indent = Number(indent);
      const ro = r.attr('readingOrder');
      if (ro !== undefined) a.readingOrder = Number(ro) as 0 | 1 | 2;
      xf.alignment = a;
    } else if (r.localName === 'protection') {
      const p: NonNullable<CellStyle['protection']> = {};
      if (r.attr('locked') !== undefined) p.locked = attrBool(r, 'locked', true);
      if (r.attr('hidden') !== undefined) p.hidden = attrBool(r, 'hidden', false);
      xf.protection = p;
    }
  }
  return xf;
}

/**
 * Flatten one `cellXfs` entry into a single CellStyle, applying the two-level
 * lookup. An `apply*` flag that is explicitly false means "take the named
 * style's value"; absent means "use mine", which is what Excel does in practice
 * even though the schema treats the flags as optional hints.
 */
export function resolveXf(tables: StyleTables, index: number): CellStyle {
  const xf = tables.cellXfs[index];
  if (!xf) return {};
  const parent = xf.xfId !== undefined ? tables.cellStyleXfs[xf.xfId] : undefined;

  const pick = <K extends keyof Xf>(key: K, applied: boolean | undefined): Xf[K] => {
    if (applied === false && parent && parent[key] !== undefined) return parent[key];
    return xf[key] !== undefined ? xf[key] : parent?.[key];
  };

  const numFmtId = pick('numFmtId', xf.applyNumberFormat) as number | undefined;
  const fontId = pick('fontId', xf.applyFont) as number | undefined;
  const fillId = pick('fillId', xf.applyFill) as number | undefined;
  const borderId = pick('borderId', xf.applyBorder) as number | undefined;

  const style: CellStyle = {};
  if (numFmtId !== undefined) {
    style.numFmtId = numFmtId;
    const code = tables.numFmts.get(numFmtId) ?? BUILTIN_NUMBER_FORMATS[numFmtId];
    if (code !== undefined) style.numFmt = code;
  }
  if (fontId !== undefined && tables.fonts[fontId]) style.font = tables.fonts[fontId];
  if (fillId !== undefined && tables.fills[fillId]) style.fill = tables.fills[fillId];
  if (borderId !== undefined && tables.borders[borderId]) style.border = tables.borders[borderId];

  const alignment = xf.applyAlignment === false ? parent?.alignment : (xf.alignment ?? parent?.alignment);
  if (alignment) style.alignment = alignment;
  const protection =
    xf.applyProtection === false ? parent?.protection : (xf.protection ?? parent?.protection);
  if (protection) style.protection = protection;
  if (xf.xfId !== undefined) style.xfId = xf.xfId;
  return style;
}

/** Resolve a colour to a concrete #RRGGBB, given a theme palette. */
export function resolveColor(color: Color | undefined, themePalette?: readonly string[]): string | undefined {
  if (!color) return undefined;
  switch (color.kind) {
    case 'rgb':
      // Stored as AARRGGBB; drop the alpha for CSS-style output.
      return `#${color.argb.length === 8 ? color.argb.slice(2) : color.argb}`;
    case 'indexed': {
      const hex = INDEXED_COLORS[color.index];
      return hex ? `#${hex}` : undefined;
    }
    case 'theme': {
      const base = themePalette?.[color.theme];
      if (!base) return undefined;
      return color.tint ? `#${applyTint(base, color.tint)}` : `#${base}`;
    }
    case 'auto':
      return undefined;
  }
}

/**
 * Excel's tint: a positive value lightens towards white, a negative one darkens
 * towards black, applied to the luminance of the HSL representation. The
 * approximation on the RGB channels below matches Excel closely enough that the
 * difference is not visible, and avoids a full RGB->HSL->RGB round trip.
 */
function applyTint(hex: string, tint: number): string {
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  const parts = [0, 2, 4].map((i) => Number.parseInt(rgb.slice(i, i + 2), 16));
  const out = parts.map((c) => {
    const v = tint < 0 ? c * (1 + tint) : c * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(v)));
  });
  return out.map((c) => c.toString(16).toUpperCase().padStart(2, '0')).join('');
}

/**
 * Serialise style tables back to styles.xml.
 *
 * Child order follows the CT_Stylesheet sequence exactly. OOXML uses xs:sequence
 * rather than xs:all, so an element in the wrong position is the single most
 * common cause of Excel's "we found a problem with some content" repair dialog -
 * and the recovery log names only the part, never the element, so it is a
 * miserable bug to chase after the fact.
 */
export function writeStyles(tables: StyleTables): string {
  const w = new XmlWriter();
  w.open('styleSheet', {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  });

  // Only custom formats are written; built-in ids are implied by the spec, and
  // redefining them can make Excel reject the file.
  const custom = [...tables.numFmts].filter(([id]) => id >= FIRST_CUSTOM_NUMFMT_ID);
  if (custom.length > 0) {
    w.open('numFmts', { count: custom.length });
    for (const [id, code] of custom) w.empty('numFmt', { numFmtId: id, formatCode: code });
    w.close();
  }

  // Every `count` attribute below is load-bearing: Excel cross-checks them and
  // a mismatch triggers silently-lossy "repaired records" recovery.
  w.open('fonts', { count: tables.fonts.length });
  for (const f of tables.fonts) writeFont(w, f);
  w.close();

  w.open('fills', { count: tables.fills.length });
  for (const f of tables.fills) writeFill(w, f);
  w.close();

  w.open('borders', { count: tables.borders.length });
  for (const b of tables.borders) writeBorder(w, b);
  w.close();

  w.open('cellStyleXfs', { count: tables.cellStyleXfs.length });
  for (const xf of tables.cellStyleXfs) writeXf(w, xf);
  w.close();

  w.open('cellXfs', { count: tables.cellXfs.length });
  for (const xf of tables.cellXfs) writeXf(w, xf);
  w.close();

  if (tables.cellStylesXml) w.raw(tables.cellStylesXml);

  w.open('dxfs', { count: tables.dxfs.length });
  for (const d of tables.dxfs) w.raw(d);
  w.close();

  // Preserved tail elements, in CT_Stylesheet order.
  for (const name of ['tableStyles', 'colors', 'extLst']) {
    const raw = tables.preserved.get(name);
    if (raw) w.raw(raw);
  }

  w.close();
  return w.toString();
}

function writeColor(w: XmlWriter, tag: string, color: Color | undefined): void {
  if (!color) return;
  switch (color.kind) {
    case 'rgb':
      w.empty(tag, { rgb: color.argb });
      break;
    case 'theme':
      w.empty(tag, { theme: color.theme, tint: color.tint });
      break;
    case 'indexed':
      w.empty(tag, { indexed: color.index });
      break;
    case 'auto':
      w.empty(tag, { auto: '1' });
      break;
  }
}

function writeFont(w: XmlWriter, f: Font): void {
  w.open('font');
  if (f.bold) w.empty('b');
  if (f.italic) w.empty('i');
  if (f.strike) w.empty('strike');
  if (f.underline && f.underline !== 'none') {
    w.empty('u', f.underline === 'single' ? undefined : { val: f.underline });
  }
  if (f.vertAlign) w.empty('vertAlign', { val: f.vertAlign });
  if (f.size !== undefined) w.empty('sz', { val: f.size });
  writeColor(w, 'color', f.color);
  if (f.name) w.empty('name', { val: f.name });
  if (f.family !== undefined) w.empty('family', { val: f.family });
  if (f.charset !== undefined) w.empty('charset', { val: f.charset });
  if (f.scheme) w.empty('scheme', { val: f.scheme });
  w.close();
}

function writeFill(w: XmlWriter, f: Fill): void {
  w.open('fill');
  if (f.gradient) {
    w.open('gradientFill', { type: f.gradient.type, degree: f.gradient.degree });
    for (const s of f.gradient.stops) {
      w.open('stop', { position: s.position });
      writeColor(w, 'color', s.color);
      w.close();
    }
    w.close();
  } else {
    const pattern = f.pattern ?? 'none';
    if (!f.fg && !f.bg) {
      w.empty('patternFill', { patternType: pattern });
    } else {
      w.open('patternFill', { patternType: pattern });
      writeColor(w, 'fgColor', f.fg);
      writeColor(w, 'bgColor', f.bg);
      w.close();
    }
  }
  w.close();
}

function writeBorder(w: XmlWriter, b: Border): void {
  w.open('border', { diagonalUp: b.diagonalUp, diagonalDown: b.diagonalDown });
  for (const edge of BORDER_EDGES) {
    const e = b[edge];
    if (!e || (!e.style && !e.color)) {
      w.empty(edge);
      continue;
    }
    if (!e.color) {
      w.empty(edge, { style: e.style });
    } else {
      w.open(edge, { style: e.style });
      writeColor(w, 'color', e.color);
      w.close();
    }
  }
  w.close();
}

function writeXf(w: XmlWriter, xf: Xf): void {
  const attrs = {
    numFmtId: xf.numFmtId ?? 0,
    fontId: xf.fontId ?? 0,
    fillId: xf.fillId ?? 0,
    borderId: xf.borderId ?? 0,
    xfId: xf.xfId,
    applyNumberFormat: xf.applyNumberFormat ? 1 : undefined,
    applyFont: xf.applyFont ? 1 : undefined,
    applyFill: xf.applyFill ? 1 : undefined,
    applyBorder: xf.applyBorder ? 1 : undefined,
    applyAlignment: xf.applyAlignment ? 1 : undefined,
    applyProtection: xf.applyProtection ? 1 : undefined,
    quotePrefix: xf.quotePrefix ? 1 : undefined,
  };
  if (!xf.alignment && !xf.protection) {
    w.empty('xf', attrs);
    return;
  }
  w.open('xf', attrs);
  if (xf.alignment) {
    w.empty('alignment', {
      horizontal: xf.alignment.horizontal,
      vertical: xf.alignment.vertical,
      wrapText: xf.alignment.wrapText ? 1 : undefined,
      textRotation: xf.alignment.textRotation,
      indent: xf.alignment.indent,
      shrinkToFit: xf.alignment.shrinkToFit ? 1 : undefined,
      readingOrder: xf.alignment.readingOrder,
    });
  }
  if (xf.protection) {
    w.empty('protection', {
      locked: xf.protection.locked === false ? 0 : undefined,
      hidden: xf.protection.hidden ? 1 : undefined,
    });
  }
  w.close();
}

/**
 * The style tables Excel expects in a brand-new workbook.
 *
 * Excel is particular about a few of these: fill index 1 must be the gray125
 * pattern, and the default font must name the minor theme scheme, or files can
 * render with the wrong default face.
 */
export function defaultStyleTables(): StyleTables {
  const tables = emptyStyleTables();
  tables.fonts.push({
    size: 11,
    color: { kind: 'theme', theme: 1 },
    name: 'Calibri',
    family: 2,
    scheme: 'minor',
  });
  tables.fills.push({ pattern: 'none' }, { pattern: 'gray125' });
  tables.borders.push({});
  tables.cellStyleXfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
  tables.cellXfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 });
  return tables;
}
