/**
 * Excel number formatting: a stored value plus a format code in, display text out.
 *
 * The format code is a small, badly documented, thoroughly load-bearing language.
 * Every date, currency and percentage a user sees comes out of it, and the file
 * format itself relies on it for one thing that is nowhere else in the XML: a
 * cell holding 45000 is a date if and only if its number format says so.
 *
 * The engine is split in two on purpose. `parseFormat` compiles a code into an
 * immutable description; `formatCompiled` walks that description. Parsing is the
 * expensive half and a workbook has a handful of distinct codes for a million
 * cells, so the compiled form is cached and reused. The formatter itself is pure
 * - no column widths, no fonts, no locale lookups - which keeps it testable and
 * lets the renderer decide separately when a value is too wide and must become
 * a run of '#'.
 *
 * Two deliberate limits. Locale is not implemented: '.' and ',' are always the
 * decimal point and the group separator, and month and day names are en-US,
 * because that is what the code points in a stored `formatCode` mean regardless
 * of who opens the file. And '*' fill is consumed but not expanded, since
 * repeating a character to the column width is a layout job, not a formatting
 * one.
 *
 * Behaviour was checked against Microsoft's numFmt documentation (ECMA-376
 * §18.8.30/§18.8.31) and cross-checked against LibreOffice's rendering of
 * fixtures/generated/styling.xlsx. Where the two disagree the Excel-documented
 * rule wins, and the divergence is called out at the point it is decided.
 */

import type { DateSystem, Scalar } from '@mirrorz/core';
import { CellError, dateToSerial, serialToParts } from '@mirrorz/core';

/** Anything a cell can hold that we might have to draw. */
export type FormatValue = Scalar | Date | undefined;

/**
 * The eight colour names the mini-language accepts, plus the indexed form.
 * `[Color1]` is index 1 into the legacy palette, which is the palette entry
 * stored as `indexed="8"` - the offset is a quirk of the file format, so we hand
 * the raw 1-56 index back and let the style layer resolve it.
 */
export type ColorName =
  | 'Black'
  | 'Blue'
  | 'Cyan'
  | 'Green'
  | 'Magenta'
  | 'Red'
  | 'White'
  | 'Yellow';

export interface FormatColor {
  name?: ColorName;
  /** 1-56, when the code used `[ColorN]` rather than a name. */
  index?: number;
}

export interface FormatResult {
  text: string;
  color?: FormatColor;
  /** True for values a renderer may replace with '#' when the column is narrow. */
  numeric: boolean;
  /**
   * True when the format cannot render this value at any width - a negative
   * date, or a conditional format where no section matched. Excel fills the
   * cell with '#' in both cases.
   */
  overflow: boolean;
}

export interface FormatOptions {
  /** Which epoch serial numbers are measured from. Workbook-level in xlsx. */
  dateSystem?: DateSystem;
}

/**
 * Format codes Excel implies rather than writing out: a cell may carry
 * `numFmtId="14"` with no matching `<numFmt>` element anywhere in styles.xml.
 *
 * Transcribed from ECMA-376 §18.8.30. Only the language-neutral ids are here.
 * The gaps (5-8, 23-36, 41-44, 50-58) are either locale-specific or reserved,
 * and a file that uses one must supply its own formatCode.
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

/** Ids below this are reserved by Excel; a workbook's own formats start here. */
export const FIRST_CUSTOM_NUMFMT_ID = 164;

export function builtinFormatCode(id: number): string | undefined {
  return BUILTIN_NUMBER_FORMATS[id];
}

const MONTHS_FULL = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];
const DAYS_FULL = [
  'Sunday',
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday',
];

const COLOR_NAMES: readonly ColorName[] = [
  'Black',
  'Blue',
  'Cyan',
  'Green',
  'Magenta',
  'Red',
  'White',
  'Yellow',
];

// --- lexing ---------------------------------------------------------------

type Lexeme =
  | { k: 'lit'; s: string }
  | { k: 'digit'; s: '0' | '#' | '?' }
  | { k: 'dot' }
  | { k: 'comma' }
  | { k: 'percent' }
  | { k: 'exp'; sign: '+' | '-' }
  | { k: 'slash' }
  | { k: 'at' }
  | { k: 'fill'; s: string }
  | { k: 'skip'; s: string }
  | { k: 'general' }
  | { k: 'date'; s: string; width: number }
  | { k: 'ampm'; long: boolean; upper: boolean }
  | { k: 'elapsed'; unit: 'h' | 'm' | 's'; width: number };

interface RawSection {
  lexemes: Lexeme[];
  color?: FormatColor;
  condition?: Condition;
}

interface Condition {
  op: '<' | '<=' | '>' | '>=' | '=' | '<>';
  value: number;
}

/**
 * Split on top-level semicolons.
 *
 * Quotes, brackets and backslash escapes all hide a semicolon, which is why this
 * cannot be a `String.split`: `"a;b"` is one literal, not two sections.
 */
function splitSections(code: string): string[] {
  const out: string[] = [];
  let start = 0;
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '\\') {
      i++;
    } else if (c === '"') {
      while (++i < code.length && code[i] !== '"') {
        /* consume */
      }
    } else if (c === '[') {
      while (++i < code.length && code[i] !== ']') {
        /* consume */
      }
    } else if (c === ';') {
      out.push(code.slice(start, i));
      start = i + 1;
    }
  }
  out.push(code.slice(start));
  return out;
}

const DATE_LETTERS = 'ymdhs';

function lexSection(src: string): RawSection {
  const lexemes: Lexeme[] = [];
  const section: RawSection = { lexemes };
  const push = (l: Lexeme): void => {
    // Merging adjacent literals keeps a fixed fraction denominator such as the
    // "16" in "# ?/16" readable as one run rather than two stray characters.
    const last = lexemes[lexemes.length - 1];
    if (l.k === 'lit' && last && last.k === 'lit') last.s += l.s;
    else lexemes.push(l);
  };

  for (let i = 0; i < src.length; ) {
    const c = src[i] as string;

    if (c === '"') {
      const end = src.indexOf('"', i + 1);
      push({ k: 'lit', s: src.slice(i + 1, end < 0 ? src.length : end) });
      i = end < 0 ? src.length : end + 1;
      continue;
    }
    if (c === '\\') {
      if (i + 1 < src.length) push({ k: 'lit', s: src[i + 1] as string });
      i += 2;
      continue;
    }
    if (c === '[') {
      const end = src.indexOf(']', i);
      const inner = src.slice(i + 1, end < 0 ? src.length : end);
      applyBracket(inner, section, push);
      i = end < 0 ? src.length : end + 1;
      continue;
    }
    if (c === '_' || c === '*') {
      const next = src[i + 1];
      push({ k: c === '_' ? 'skip' : 'fill', s: next ?? ' ' });
      i += next === undefined ? 1 : 2;
      continue;
    }
    if (c === '0' || c === '#' || c === '?') {
      push({ k: 'digit', s: c });
      i++;
      continue;
    }
    if (c === '.') {
      push({ k: 'dot' });
      i++;
      continue;
    }
    if (c === ',') {
      push({ k: 'comma' });
      i++;
      continue;
    }
    if (c === '%') {
      push({ k: 'percent' });
      i++;
      continue;
    }
    if (c === '/') {
      push({ k: 'slash' });
      i++;
      continue;
    }
    if (c === '@') {
      push({ k: 'at' });
      i++;
      continue;
    }
    if (/^general/i.test(src.slice(i))) {
      push({ k: 'general' });
      i += 7;
      continue;
    }
    if (/^am\/pm/i.test(src.slice(i))) {
      push({ k: 'ampm', long: true, upper: src[i] === 'A' });
      i += 5;
      continue;
    }
    if (/^a\/p/i.test(src.slice(i))) {
      push({ k: 'ampm', long: false, upper: src[i] === 'A' });
      i += 3;
      continue;
    }
    const lower = c.toLowerCase();
    if (DATE_LETTERS.includes(lower)) {
      let n = 1;
      while (src[i + n]?.toLowerCase() === lower) n++;
      push({ k: 'date', s: lower, width: n });
      i += n;
      continue;
    }
    if (lower === 'e') {
      const next = src[i + 1];
      if (next === '+' || next === '-') {
        push({ k: 'exp', sign: next });
        i += 2;
        continue;
      }
      // A bare 'e' is the Japanese-era year token; a bare 'E' is not a token.
      if (c === 'e') {
        let n = 1;
        while (src[i + n] === 'e') n++;
        push({ k: 'date', s: 'e', width: n });
        i += n;
        continue;
      }
    }
    if (lower === 'g') {
      let n = 1;
      while (src[i + n]?.toLowerCase() === 'g') n++;
      push({ k: 'date', s: 'g', width: n });
      i += n;
      continue;
    }
    if (lower === 'b' && (src[i + 1] === '1' || src[i + 1] === '2')) {
      // Calendar selector. We only implement the Gregorian calendar, so both
      // forms are dropped rather than mis-rendered.
      i += 2;
      continue;
    }
    push({ k: 'lit', s: c });
    i++;
  }
  return section;
}

function applyBracket(
  inner: string,
  section: RawSection,
  push: (l: Lexeme) => void,
): void {
  const named = COLOR_NAMES.find((n) => n.toLowerCase() === inner.toLowerCase());
  if (named) {
    section.color = { name: named };
    return;
  }
  const indexed = /^color\s*(\d{1,2})$/i.exec(inner);
  if (indexed) {
    const n = Number(indexed[1]);
    if (n >= 1 && n <= 56) section.color = { index: n };
    return;
  }
  const cond = /^(<=|>=|<>|=|<|>)(.+)$/.exec(inner);
  if (cond) {
    const value = Number(cond[2]);
    if (!Number.isNaN(value)) {
      section.condition = { op: cond[1] as Condition['op'], value };
    }
    return;
  }
  if (inner.startsWith('$')) {
    // [$-409] carries only locale; [$USD-409] and [$<eur>-407] also carry a
    // currency string, which is printed. The locale id is the tail after the
    // last '-', so a currency containing a dash still survives.
    const body = inner.slice(1);
    const dash = body.lastIndexOf('-');
    const currency = dash < 0 ? body : body.slice(0, dash);
    if (currency) push({ k: 'lit', s: currency });
    return;
  }
  const elapsed = /^(h+|m+|s+)$/i.exec(inner);
  if (elapsed) {
    const text = elapsed[1] as string;
    push({
      k: 'elapsed',
      unit: text[0]?.toLowerCase() as 'h' | 'm' | 's',
      width: text.length,
    });
    return;
  }
  // [DBNum1], [ENG], [t] and friends: recognised as brackets, then ignored.
}

// --- compiled shape -------------------------------------------------------

type Role =
  | 'int'
  | 'frac'
  | 'exp'
  | 'num'
  | 'den'
  | 'slash'
  | 'point'
  | 'drop'
  | 'plain';

interface Piece {
  lex: Lexeme;
  role: Role;
  /** Ordinal of this placeholder within its group, for positional filling. */
  slot: number;
}

interface FractionSpec {
  numPattern: string;
  denPattern: string;
  /** Set when the denominator is a literal such as the 16 in "# ?/16". */
  fixedDen?: number;
  hasInteger: boolean;
}

interface NumberSection {
  kind: 'number';
  color?: FormatColor;
  condition?: Condition;
  pieces: Piece[];
  intPattern: string;
  fracPattern: string;
  grouping: boolean;
  /** Trailing commas, each dividing by 1000. */
  scale: number;
  percent: boolean;
  exponent?: { sign: '+' | '-'; digits: number };
  fraction?: FractionSpec;
  /** False for literal-only sections such as `"big"`, which never take a sign. */
  hasPlaceholders: boolean;
}

type DatePiece =
  | { p: 'year'; width: number }
  | { p: 'month'; width: number }
  | { p: 'day'; width: number }
  | { p: 'hour'; width: number }
  | { p: 'minute'; width: number }
  | { p: 'second'; width: number }
  | { p: 'subsecond'; width: number }
  | { p: 'ampm'; long: boolean; upper: boolean }
  | { p: 'elapsed'; unit: 'h' | 'm' | 's'; width: number }
  | { p: 'era'; width: number }
  | { p: 'eraYear'; width: number }
  | { p: 'literal'; s: string };

interface DateSection {
  kind: 'date';
  color?: FormatColor;
  condition?: Condition;
  pieces: DatePiece[];
  twelveHour: boolean;
  /** Digits of a second the format shows; drives display rounding. */
  subsecond: number;
  smallest: 'day' | 'hour' | 'minute' | 'second';
}

interface GeneralSection {
  kind: 'general';
  color?: FormatColor;
  condition?: Condition;
}

type Section = NumberSection | DateSection | GeneralSection;

export interface CompiledFormat {
  readonly code: string;
  readonly sections: readonly Section[];
  /** Sections 1-3: the ones a number can land in. */
  readonly numericSections: readonly Section[];
  readonly textSection?: Section;
  readonly isDate: boolean;
}

// --- compilation ----------------------------------------------------------

const cache = new Map<string, CompiledFormat>();
const CACHE_LIMIT = 512;

export function parseFormat(code: string): CompiledFormat {
  const hit = cache.get(code);
  if (hit) return hit;
  const compiled = compile(code);
  if (cache.size >= CACHE_LIMIT) cache.clear();
  cache.set(code, compiled);
  return compiled;
}

function compile(code: string): CompiledFormat {
  const sections = splitSections(code).map((s) => compileSection(lexSection(s)));
  const last = sections[sections.length - 1] as Section;
  // The text section is normally the fourth. It is also any *final* section that
  // carries an '@', because the two-section idiom "0.00;@" is everywhere in real
  // files - Excel and openpyxl both emit it - and it means "numbers here, text
  // there", not "positives here, negatives there".
  const trailingText = sections.length > 1 && sectionHasText(last);
  const pool = trailingText ? sections.slice(0, -1) : sections;
  const numericSections = pool.slice(0, 3);
  const textSection = trailingText
    ? last
    : sections.length >= 4
      ? sections[3]
      : sections.length === 1 && sectionHasText(last)
        ? last
        : undefined;
  return {
    code,
    sections,
    numericSections,
    textSection,
    isDate: numericSections.some((s) => s.kind === 'date'),
  };
}

function sectionHasText(section: Section): boolean {
  return (
    section.kind === 'number' && section.pieces.some((p) => p.lex.k === 'at')
  );
}

function compileSection(raw: RawSection): Section {
  const { lexemes } = raw;
  const isDate = lexemes.some(
    (l) => l.k === 'date' || l.k === 'ampm' || l.k === 'elapsed',
  );
  if (isDate) return compileDateSection(raw);
  // Only a bare "General" is the General format. `General" units"` keeps the
  // literal, so it compiles as a number section whose sole placeholder happens
  // to be the General rendering of the value.
  if (lexemes.length === 1 && lexemes[0]?.k === 'general') {
    return { kind: 'general', color: raw.color, condition: raw.condition };
  }
  return compileNumberSection(raw);
}

/**
 * The month-versus-minute rule, which is where most naive implementations break.
 *
 * ECMA-376: an `m`/`mm` immediately after `h`/`hh`, or immediately before
 * `s`/`ss`, is minutes; otherwise it is a month. "Immediately" means with only
 * literals in between - `h:mm` has a colon between the two and is still minutes
 * - so the test is against the neighbouring *date* tokens, not the neighbouring
 * characters. Three or more `m`s are always a month name, so only the one- and
 * two-letter forms are ambiguous.
 */
function isMinuteToken(lexemes: Lexeme[], index: number): boolean {
  const self = lexemes[index];
  if (!self || self.k !== 'date' || self.s !== 'm' || self.width > 2) return false;
  for (let i = index - 1; i >= 0; i--) {
    const l = lexemes[i];
    if (!l) continue;
    if (l.k === 'elapsed') return l.unit === 'h';
    if (l.k !== 'date') continue;
    if (l.s === 'h') return true;
    break;
  }
  for (let i = index + 1; i < lexemes.length; i++) {
    const l = lexemes[i];
    if (!l) continue;
    if (l.k === 'elapsed') return l.unit === 's';
    if (l.k !== 'date') continue;
    if (l.s === 's') return true;
    break;
  }
  return false;
}

function compileDateSection(raw: RawSection): DateSection {
  const { lexemes } = raw;
  const pieces: DatePiece[] = [];
  const twelveHour = lexemes.some((l) => l.k === 'ampm');
  let subsecond = 0;

  for (let i = 0; i < lexemes.length; i++) {
    const l = lexemes[i] as Lexeme;
    switch (l.k) {
      case 'date':
        switch (l.s) {
          case 'y':
            pieces.push({ p: 'year', width: l.width <= 2 ? 2 : 4 });
            break;
          case 'd':
            pieces.push({ p: 'day', width: l.width });
            break;
          case 'h':
            pieces.push({ p: 'hour', width: l.width });
            break;
          case 's':
            pieces.push({ p: 'second', width: l.width });
            break;
          case 'e':
            pieces.push({ p: 'eraYear', width: l.width });
            break;
          case 'g':
            pieces.push({ p: 'era', width: l.width });
            break;
          default:
            pieces.push(
              isMinuteToken(lexemes, i)
                ? { p: 'minute', width: l.width }
                : { p: 'month', width: l.width },
            );
        }
        break;
      case 'ampm':
        pieces.push({ p: 'ampm', long: l.long, upper: l.upper });
        break;
      case 'elapsed':
        pieces.push({ p: 'elapsed', unit: l.unit, width: l.width });
        break;
      case 'dot': {
        // ".0" is fractional seconds only where seconds precede it; anywhere
        // else a dot in a date format is just a separator.
        const last = pieces[pieces.length - 1];
        const secondsBefore =
          last !== undefined &&
          (last.p === 'second' || (last.p === 'elapsed' && last.unit === 's'));
        let width = 0;
        while (lexemes[i + 1 + width]?.k === 'digit') {
          const d = lexemes[i + 1 + width] as { k: 'digit'; s: string };
          if (d.s !== '0') break;
          width++;
        }
        if (secondsBefore && width > 0) {
          pieces.push({ p: 'subsecond', width });
          subsecond = Math.max(subsecond, width);
          i += width;
        } else {
          pieces.push({ p: 'literal', s: '.' });
        }
        break;
      }
      case 'lit':
        pieces.push({ p: 'literal', s: l.s });
        break;
      case 'skip':
        pieces.push({ p: 'literal', s: ' ' });
        break;
      case 'fill':
        break;
      case 'slash':
        pieces.push({ p: 'literal', s: '/' });
        break;
      case 'comma':
        pieces.push({ p: 'literal', s: ',' });
        break;
      case 'percent':
        pieces.push({ p: 'literal', s: '%' });
        break;
      case 'digit':
        pieces.push({ p: 'literal', s: l.s });
        break;
      default:
        break;
    }
  }

  const has = (p: DatePiece['p']): boolean => pieces.some((x) => x.p === p);
  const elapsedUnit = pieces.find((x) => x.p === 'elapsed');
  let smallest: DateSection['smallest'] = 'day';
  if (has('second') || (elapsedUnit?.p === 'elapsed' && elapsedUnit.unit === 's')) {
    smallest = 'second';
  } else if (
    has('minute') ||
    (elapsedUnit?.p === 'elapsed' && elapsedUnit.unit === 'm')
  ) {
    smallest = 'minute';
  } else if (has('hour') || elapsedUnit !== undefined) {
    smallest = 'hour';
  }

  return {
    kind: 'date',
    color: raw.color,
    condition: raw.condition,
    pieces,
    twelveHour,
    subsecond,
    smallest,
  };
}

function compileNumberSection(raw: RawSection): NumberSection {
  const { lexemes } = raw;
  const pieces: Piece[] = [];

  const dotAt = lexemes.findIndex((l) => l.k === 'dot');
  const expAt = lexemes.findIndex((l) => l.k === 'exp');
  const slashAt = lexemes.findIndex((l) => l.k === 'slash');

  // The integer region ends at the first decimal point or exponent marker.
  const intEnd = Math.min(
    dotAt < 0 ? lexemes.length : dotAt,
    expAt < 0 ? lexemes.length : expAt,
  );

  const fraction =
    slashAt >= 0 && dotAt < 0 && expAt < 0
      ? describeFraction(lexemes, slashAt)
      : undefined;

  let intPattern = '';
  let fracPattern = '';
  let expDigits = 0;
  let grouping = false;
  let scale = 0;
  let percent = false;

  let numSlots = 0;
  let denSlots = 0;

  for (let i = 0; i < lexemes.length; i++) {
    const l = lexemes[i] as Lexeme;
    let role: Role = 'plain';
    let slot = 0;
    switch (l.k) {
      case 'digit':
        if (fraction) {
          if (i > slashAt) {
            role = 'den';
            slot = denSlots++;
          } else if (i >= fraction.numStart) {
            role = 'num';
            slot = numSlots++;
          } else {
            role = 'int';
            slot = intPattern.length;
            intPattern += l.s;
          }
        } else if (expAt >= 0 && i > expAt) {
          role = 'drop';
          expDigits++;
        } else if (dotAt >= 0 && i > dotAt) {
          role = 'frac';
          slot = fracPattern.length;
          fracPattern += l.s;
        } else {
          role = 'int';
          slot = intPattern.length;
          intPattern += l.s;
        }
        break;
      case 'dot':
        role = 'point';
        break;
      case 'comma': {
        // A comma with a digit still to come inside the integer region groups
        // thousands; anywhere else it scales the value down by 1000.
        let groupsHere = false;
        for (let j = i + 1; j < intEnd; j++) {
          if (lexemes[j]?.k === 'digit') {
            groupsHere = true;
            break;
          }
        }
        if (groupsHere) grouping = true;
        else scale++;
        role = 'drop';
        break;
      }
      case 'percent':
        percent = true;
        role = 'plain';
        break;
      case 'exp':
        role = 'exp';
        break;
      case 'slash':
        role = fraction ? 'slash' : 'plain';
        break;
      default:
        role = 'plain';
    }
    pieces.push({ lex: l, role, slot });
  }

  const section: NumberSection = {
    kind: 'number',
    color: raw.color,
    condition: raw.condition,
    pieces,
    intPattern,
    fracPattern,
    grouping,
    scale,
    percent,
    hasPlaceholders: lexemes.some((l) => l.k === 'digit'),
  };
  if (expAt >= 0) {
    section.exponent = {
      sign: (lexemes[expAt] as { k: 'exp'; sign: '+' | '-' }).sign,
      digits: Math.max(1, expDigits),
    };
  }
  if (fraction) {
    section.fraction = {
      numPattern: fraction.numPattern,
      denPattern: fraction.denPattern,
      hasInteger: intPattern.length > 0,
      ...(fraction.fixedDen !== undefined ? { fixedDen: fraction.fixedDen } : {}),
    };
  }
  return section;
}

/**
 * Work out which placeholders belong to a fraction.
 *
 * The numerator is the last unbroken run of placeholders before the slash; any
 * placeholder further left is the whole-number part. The denominator is either
 * placeholders after the slash or a literal number, as in "# ?/16".
 */
function describeFraction(
  lexemes: Lexeme[],
  slashAt: number,
): {
  numStart: number;
  numPattern: string;
  denPattern: string;
  fixedDen?: number;
} | undefined {
  let numStart = slashAt;
  let numPattern = '';
  while (numStart > 0 && lexemes[numStart - 1]?.k === 'digit') {
    numStart--;
    numPattern = (lexemes[numStart] as { k: 'digit'; s: string }).s + numPattern;
  }
  if (numPattern === '') return undefined;

  let denPattern = '';
  for (let i = slashAt + 1; i < lexemes.length; i++) {
    const l = lexemes[i];
    if (l?.k !== 'digit') break;
    denPattern += l.s;
  }
  if (denPattern !== '') return { numStart, numPattern, denPattern };

  const after = lexemes[slashAt + 1];
  const fixed = after?.k === 'lit' ? /^\d+/.exec(after.s) : null;
  if (!fixed) return undefined;
  return { numStart, numPattern, denPattern: '', fixedDen: Number(fixed[0]) };
}

// --- decimal helpers ------------------------------------------------------

/**
 * Render a double as a plain decimal string, clamped to 15 significant digits.
 *
 * Excel carries doubles but never shows more than 15 significant decimal
 * digits, and that single rule is what makes 0.1+0.2 display as 0.3 rather than
 * 0.30000000000000004. Doing the clamp here, once, means every later step -
 * rounding, General, scientific - inherits it for free.
 */
export function toPlainDecimal(value: number, significant = 15): string {
  if (!Number.isFinite(value)) return String(value);
  const s = value.toPrecision(significant);
  const m = /^(-?)(\d+)(?:\.(\d+))?e([+-]\d+)$/i.exec(s);
  if (!m) return s;
  const sign = m[1] ?? '';
  const digits = (m[2] ?? '0') + (m[3] ?? '');
  const point = (m[2] ?? '0').length + Number(m[4]);
  if (point <= 0) return `${sign}0.${'0'.repeat(-point)}${digits}`;
  if (point >= digits.length) return sign + digits + '0'.repeat(point - digits.length);
  return `${sign}${digits.slice(0, point)}.${digits.slice(point)}`;
}

function incrementDigits(digits: string): string {
  const out = digits.split('');
  for (let i = out.length - 1; i >= 0; i--) {
    if (out[i] === '9') {
      out[i] = '0';
    } else {
      out[i] = String(Number(out[i]) + 1);
      return out.join('');
    }
  }
  return '1' + out.join('');
}

/**
 * Round a non-negative value to `places` decimals, half away from zero.
 *
 * `toFixed` cannot be used: it rounds the binary double, so (1.005).toFixed(2)
 * is "1.00" where Excel shows 1.01. Rounding the 15-digit decimal expansion
 * instead reproduces Excel because Excel makes the same 15-digit cut first.
 */
function roundDecimal(value: number, places: number): { int: string; frac: string } {
  const plain = toPlainDecimal(Math.abs(value));
  const dot = plain.indexOf('.');
  const intPart = dot < 0 ? plain : plain.slice(0, dot);
  const fracPart = dot < 0 ? '' : plain.slice(dot + 1);
  if (fracPart.length <= places) {
    return { int: stripLeadingZeros(intPart), frac: fracPart.padEnd(places, '0') };
  }
  let digits = intPart + fracPart.slice(0, places);
  if ((fracPart.charCodeAt(places) - 48) >= 5) digits = incrementDigits(digits);
  const frac = places > 0 ? digits.slice(digits.length - places) : '';
  const int = digits.slice(0, digits.length - places);
  return { int: stripLeadingZeros(int), frac };
}

function stripLeadingZeros(s: string): string {
  const t = s.replace(/^0+/, '');
  return t === '' ? '0' : t;
}

function decimalExponent(value: number): number {
  if (value === 0 || !Number.isFinite(value)) return 0;
  const parts = Math.abs(value).toExponential(14).split('e');
  return Number(parts[1]);
}

// --- General --------------------------------------------------------------

/**
 * Excel's "General" rendering.
 *
 * The documented rule is a width budget of 11 characters (leading zeros and the
 * decimal separator count, the minus sign does not), trailing zeros dropped, and
 * a switch to exponential notation once fixed notation cannot fit. The banding
 * below is the practical form of that: small magnitudes lose budget to leading
 * zeros so fewer significant digits are asked for, and beyond the budget the
 * value goes exponential with a two-digit exponent.
 *
 * LibreOffice widens the column instead of trimming, so its rendering of, say,
 * 1/3 keeps 15 digits. Excel's 11-character rule is the one implemented here.
 */
export function formatGeneral(value: number): string {
  if (!Number.isFinite(value)) return '#NUM!';
  if (value === 0) return '0';
  const budget = value < 0 ? 12 : 11;
  const exp = decimalExponent(value);
  let out: string;
  if (exp >= -4 && exp <= -1) {
    out = value.toPrecision(10 + exp);
  } else if (exp >= -9 && exp <= 9) {
    out = stripTrailingZeros(toPlainDecimal(value, 15));
    if (out.length > budget) out = value.toPrecision(10);
    if (stripTrailingZeros(out).length > budget) out = value.toExponential(5);
  } else if (exp === 10) {
    out = toPlainDecimal(value, 11);
  } else {
    out = value.toExponential(5);
  }
  return normaliseExponent(stripTrailingZeros(out));
}

function stripTrailingZeros(s: string): string {
  if (!s.includes('.')) return s;
  if (/e/i.test(s)) {
    return s.replace(/\.?0+(?=[eE])/, '');
  }
  return s.replace(/\.0*$/, '').replace(/(\.\d*[1-9])0+$/, '$1');
}

function normaliseExponent(s: string): string {
  const m = /^(-?[\d.]+)e([+-])(\d+)$/i.exec(s);
  if (!m) return s;
  return `${m[1]}E${m[2]}${(m[3] as string).padStart(2, '0')}`;
}

// --- fractions ------------------------------------------------------------

/**
 * Best rational approximation of `x` with a denominator no larger than
 * `maxDen`, by continued fractions.
 *
 * Rounding `x * maxDen` would be wrong: with a one-digit denominator, 0.7 is
 * closer to 5/7 than to the 6/9 that naive rounding gives. The convergents of
 * the continued fraction are the candidates, plus - when the next convergent
 * overshoots the limit - the largest semiconvergent that still fits, which is
 * the only other value that can win.
 */
export function bestFraction(x: number, maxDen: number): { num: number; den: number } {
  if (!Number.isFinite(x) || maxDen < 1) return { num: 0, den: 1 };
  let h0 = 0;
  let h1 = 1;
  let k0 = 1;
  let k1 = 0;
  let b = x;
  for (let step = 0; step < 40; step++) {
    const a = Math.floor(b);
    const h2 = a * h1 + h0;
    const k2 = a * k1 + k0;
    if (k2 > maxDen) {
      if (k1 > 0) {
        const j = Math.floor((maxDen - k0) / k1);
        const hs = j * h1 + h0;
        const ks = j * k1 + k0;
        if (ks >= 1 && Math.abs(x - hs / ks) < Math.abs(x - h1 / k1)) {
          return { num: hs, den: ks };
        }
      }
      break;
    }
    h0 = h1;
    h1 = h2;
    k0 = k1;
    k1 = k2;
    const rem = b - a;
    if (rem < 1e-12) break;
    b = 1 / rem;
    if (!Number.isFinite(b)) break;
  }
  return k1 === 0 ? { num: Math.round(x), den: 1 } : { num: h1, den: k1 };
}

// --- number rendering -----------------------------------------------------

function padInteger(digits: string, pattern: string): string {
  // '#' and '?' both drop a value of zero; only a literal '0' forces it to show.
  // This is why the accounting zero section `_("$"* "-"??_)` prints a dash and
  // blanks rather than a dash and "0".
  let text = digits === '0' && !pattern.includes('0') ? '' : digits;
  if (text.length >= pattern.length) return text;
  let pad = '';
  for (let i = 0; i < pattern.length - text.length; i++) {
    const ch = pattern[i];
    pad += ch === '0' ? '0' : ch === '?' ? ' ' : '';
  }
  return pad + text;
}

function groupThousands(text: string): string {
  const m = /^( *)(\d*)$/.exec(text);
  if (!m) return text;
  const digits = m[2] as string;
  if (digits.length <= 3) return text;
  let out = '';
  for (let i = 0; i < digits.length; i++) {
    if (i > 0 && (digits.length - i) % 3 === 0) out += ',';
    out += digits[i];
  }
  return (m[1] as string) + out;
}

function trimFraction(digits: string, pattern: string): string {
  let end = digits.length;
  while (end > 0 && pattern[end - 1] !== '0' && digits[end - 1] === '0') end--;
  let out = digits.slice(0, end);
  for (let i = end; i < pattern.length; i++) if (pattern[i] === '?') out += ' ';
  return out;
}

function padLeft(text: string, width: number): string {
  return text.length >= width ? text : ' '.repeat(width - text.length) + text;
}

function padRight(text: string, width: number): string {
  return text.length >= width ? text : text + ' '.repeat(width - text.length);
}

interface NumberRender {
  text: string;
  /** False when everything rendered was zero, so no minus sign is warranted. */
  significant: boolean;
}

function renderNumberSection(section: NumberSection, magnitude: number): NumberRender {
  // Percent scales once however many '%' the section carries, and each trailing
  // comma divides by another thousand.
  let n = magnitude;
  if (section.percent) n *= 100;
  for (let i = 0; i < section.scale; i++) n /= 1000;

  const r: Rendered = {
    intText: '',
    fracText: '',
    expText: '',
    numText: '',
    denText: '',
    fractionBlank: false,
    magnitude,
  };
  let significant: boolean;

  if (section.fraction) {
    significant = renderFraction(section, section.fraction, n, r);
  } else if (section.exponent) {
    significant = renderScientific(section, section.exponent, n, r);
  } else {
    const rounded = roundDecimal(n, section.fracPattern.length);
    r.intText = padInteger(rounded.int, section.intPattern);
    r.fracText = trimFraction(rounded.frac, section.fracPattern);
    significant = /[1-9]/.test(rounded.int + rounded.frac);
  }
  if (section.grouping) r.intText = groupThousands(r.intText);
  return { text: assemble(section, r), significant };
}

function renderFraction(
  section: NumberSection,
  spec: FractionSpec,
  n: number,
  r: Rendered,
): boolean {
  const whole = spec.hasInteger ? Math.floor(n) : 0;
  const rest = n - whole;
  const maxDen =
    spec.fixedDen !== undefined ? spec.fixedDen : Math.pow(10, spec.denPattern.length) - 1;
  let { num, den } =
    spec.fixedDen !== undefined
      ? { num: Math.round(rest * spec.fixedDen), den: spec.fixedDen }
      : bestFraction(rest, maxDen);
  if (den < 1) den = 1;
  let intValue = whole;
  if (spec.hasInteger && num >= den) {
    // The approximation rounded the remainder up to a whole unit: 1.9999 with
    // "# ?/?" is 2, not "1 1/1".
    intValue += Math.floor(num / den);
    num %= den;
  }

  if (num === 0 && spec.hasInteger) {
    // Excel blanks the fraction rather than printing "0/1", keeping the column
    // aligned; the integer is forced to show so the cell is not left empty.
    r.fractionBlank = true;
    r.intText = padInteger(String(intValue), section.intPattern) || '0';
    r.numText = ' '.repeat(spec.numPattern.length);
    r.denText = ' '.repeat(spec.denPattern.length);
    return intValue !== 0;
  }
  r.intText = spec.hasInteger ? padInteger(String(intValue), section.intPattern) : '';
  r.numText = padLeft(String(num), spec.numPattern.length);
  // The denominator pads on the right, so that "5   1/4  " and "5   3/10 " line
  // up under "# ???/???" the way Excel aligns them.
  r.denText =
    spec.fixedDen !== undefined ? '' : padRight(String(den), spec.denPattern.length);
  return true;
}

function renderScientific(
  section: NumberSection,
  spec: { sign: '+' | '-'; digits: number },
  n: number,
  r: Rendered,
): boolean {
  const width = Math.max(1, section.intPattern.length);
  // A mantissa pattern of bare zeros pins the integer digit count; anything with
  // '#' or '?' steps the exponent in multiples of that width, which is how
  // "##0.0E+0" produces engineering notation.
  const stepped = /[#?]/.test(section.intPattern);
  let exp = n === 0 ? 0 : decimalExponent(n);
  exp = stepped ? Math.floor(exp / width) * width : exp - (width - 1);
  let rounded = roundDecimal(n === 0 ? 0 : n / Math.pow(10, exp), section.fracPattern.length);
  if (rounded.int.length > width && n !== 0) {
    // 9.99 with "0.0E+0" rounds to 10.0; shift a decade and try again.
    exp += stepped ? width : 1;
    rounded = roundDecimal(n / Math.pow(10, exp), section.fracPattern.length);
  }
  r.intText = padInteger(rounded.int, section.intPattern);
  r.fracText = trimFraction(rounded.frac, section.fracPattern);
  const sign = exp < 0 ? '-' : spec.sign === '+' ? '+' : '';
  r.expText = sign + String(Math.abs(exp)).padStart(spec.digits, '0');
  return n !== 0;
}

interface Rendered {
  intText: string;
  fracText: string;
  expText: string;
  numText: string;
  denText: string;
  fractionBlank: boolean;
  magnitude: number;
}

/**
 * Spread rendered digits over the placeholder positions, filling from the right
 * and giving any overflow to the leftmost slot.
 *
 * Placeholders keep their positions relative to the literals around them, which
 * is what makes "000-0000" turn 5551234 into 555-1234 rather than 5551234-.
 */
function distributeRight(text: string, count: number): string[] {
  if (count === 0) return [];
  const out = new Array<string>(count).fill('');
  let i = text.length;
  for (let slot = count - 1; slot >= 1 && i > 0; slot--) out[slot] = text[--i] as string;
  out[0] = text.slice(0, i);
  return out;
}

/** The same, for the fraction side, where overflow belongs to the last slot. */
function distributeLeft(text: string, count: number): string[] {
  if (count === 0) return [];
  const out = new Array<string>(count).fill('');
  let i = 0;
  for (let slot = 0; slot < count - 1 && i < text.length; slot++) out[slot] = text[i++] as string;
  out[count - 1] = text.slice(i);
  return out;
}

function assemble(section: NumberSection, r: Rendered): string {
  const intSlots = distributeRight(r.intText, section.intPattern.length);
  const fracSlots = distributeLeft(r.fracText, section.fracPattern.length);
  const numSlots = distributeRight(r.numText, section.fraction?.numPattern.length ?? 0);
  const denSlots = distributeLeft(r.denText, section.fraction?.denPattern.length ?? 0);
  let out = '';
  for (const piece of section.pieces) {
    switch (piece.role) {
      case 'int':
        out += intSlots[piece.slot] ?? '';
        break;
      case 'frac':
        out += fracSlots[piece.slot] ?? '';
        break;
      case 'num':
        out += numSlots[piece.slot] ?? '';
        break;
      case 'den':
        out += denSlots[piece.slot] ?? '';
        break;
      case 'slash':
        out += r.fractionBlank ? ' ' : '/';
        break;
      case 'exp':
        out += 'E' + r.expText;
        break;
      case 'point':
        // Excel writes the separator only when something follows it, so "0."
        // shows 5 rather than "5.", while "0.??" keeps "5.  " for alignment.
        if (r.fracText !== '') out += '.';
        break;
      case 'drop':
        break;
      default:
        out += plainText(piece.lex, r);
    }
  }
  return out;
}

function plainText(lex: Lexeme, r: Rendered): string {
  switch (lex.k) {
    case 'lit':
      return lex.s;
    case 'percent':
      return '%';
    case 'skip':
      // A width-skip should be as wide as the character it names; without font
      // metrics a single space is the honest approximation.
      return ' ';
    case 'fill':
      return '';
    case 'at':
    case 'general':
      return formatGeneral(r.magnitude);
    case 'slash':
      return '/';
    case 'dot':
      return '.';
    case 'comma':
      return ',';
    default:
      return '';
  }
}

// --- date rendering -------------------------------------------------------

const MS_PER_DAY = 86_400_000;

function renderDateSection(
  section: DateSection,
  serial: number,
  system: DateSystem,
): string | undefined {
  // Excel rounds a date/time to the precision it is about to show, which is why
  // 23:59:59.9 formatted as "h:mm:ss" rolls over to the next day rather than
  // truncating. LibreOffice truncates instead; we follow Excel.
  const perDay =
    section.subsecond > 0
      ? 86_400 * Math.pow(10, section.subsecond)
      : section.smallest === 'second'
        ? 86_400
        : section.smallest === 'minute'
          ? 1_440
          : section.smallest === 'hour'
            ? 24
            : 0;
  const rounded = perDay === 0 ? serial : Math.round(serial * perDay) / perDay;
  // A date or time outside the representable range is a '#' fill in Excel, not
  // a wrapped-around calendar date, so refuse rather than invent one.
  if (rounded < 0) return undefined;

  const parts = serialToParts(rounded, system);
  if (parts.year > 9999 || Number.isNaN(parts.year)) return undefined;
  const totalMs = Math.round(rounded * MS_PER_DAY);
  const hour12 = parts.hour % 12 === 0 ? 12 : parts.hour % 12;
  const pm = parts.hour >= 12;

  let out = '';
  for (const p of section.pieces) {
    switch (p.p) {
      case 'year':
        out +=
          p.width === 2
            ? String(parts.year % 100).padStart(2, '0')
            : String(parts.year).padStart(4, '0');
        break;
      case 'month':
        if (p.width === 1) out += String(parts.month);
        else if (p.width === 2) out += String(parts.month).padStart(2, '0');
        else if (p.width === 3) out += (MONTHS_FULL[parts.month - 1] as string).slice(0, 3);
        else if (p.width === 4) out += MONTHS_FULL[parts.month - 1] as string;
        else out += (MONTHS_FULL[parts.month - 1] as string).slice(0, 1);
        break;
      case 'day': {
        if (p.width === 1) out += String(parts.day);
        else if (p.width === 2) out += String(parts.day).padStart(2, '0');
        else {
          const name = DAYS_FULL[weekdayIndex(rounded, system)] as string;
          out += p.width === 3 ? name.slice(0, 3) : name;
        }
        break;
      }
      case 'hour': {
        const h = section.twelveHour ? hour12 : parts.hour;
        out += p.width >= 2 ? String(h).padStart(2, '0') : String(h);
        break;
      }
      case 'minute':
        out += p.width >= 2 ? String(parts.minute).padStart(2, '0') : String(parts.minute);
        break;
      case 'second':
        out += p.width >= 2 ? String(parts.second).padStart(2, '0') : String(parts.second);
        break;
      case 'subsecond': {
        const scaled = Math.round(parts.millisecond / Math.pow(10, 3 - p.width));
        out += '.' + String(scaled).padStart(p.width, '0');
        break;
      }
      case 'ampm': {
        // The code's own case is kept: "AM/PM" prints PM, "am/pm" prints pm.
        const word = p.long ? (pm ? 'PM' : 'AM') : pm ? 'P' : 'A';
        out += p.upper ? word : word.toLowerCase();
        break;
      }
      case 'elapsed': {
        const total =
          p.unit === 'h'
            ? Math.floor(totalMs / 3_600_000)
            : p.unit === 'm'
              ? Math.floor(totalMs / 60_000)
              : Math.floor(totalMs / 1000);
        out += String(total).padStart(p.width, '0');
        break;
      }
      case 'eraYear':
        // Japanese-era years are not implemented; falling back to the Gregorian
        // year keeps the rest of the format readable instead of crashing.
        out += p.width >= 2 ? String(parts.year).padStart(4, '0') : String(parts.year);
        break;
      case 'era':
        break;
      default:
        out += p.s;
    }
  }
  return out;
}

/** Sunday = 0, taken from the serial so the 1900 leap-year bug is preserved. */
function weekdayIndex(serial: number, system: DateSystem): number {
  const days = Math.floor(serial);
  if (system === 1904) return ((days % 7) + 5) % 7;
  return (((days - 1) % 7) + 7) % 7;
}

// --- section selection ----------------------------------------------------

function conditionMatches(c: Condition, v: number): boolean {
  switch (c.op) {
    case '<':
      return v < c.value;
    case '<=':
      return v <= c.value;
    case '>':
      return v > c.value;
    case '>=':
      return v >= c.value;
    case '=':
      return v === c.value;
    default:
      return v !== c.value;
  }
}

/**
 * Pick the section a number belongs in.
 *
 * One section covers every number. Two split at zero, with zero joining the
 * positives. Three give zero its own. Conditions override all of that: the
 * first section whose test passes wins, and an unconditioned section acts as
 * the "otherwise" arm.
 */
function chooseSection(compiled: CompiledFormat, value: number): Section | undefined {
  const sections = compiled.numericSections;
  if (sections.length === 0) return undefined;
  if (sections.some((s) => s.condition)) {
    const matched = sections.find(
      (s) => s.condition && conditionMatches(s.condition, value),
    );
    if (matched) return matched;
    // With no "otherwise" arm the value meets no criterion, and ECMA-376 says
    // to fill the cell with '#'. LibreOffice quietly falls back to General.
    return sections.find((s) => !s.condition);
  }
  if (sections.length === 1) return sections[0];
  if (sections.length === 2) return sections[value < 0 ? 1 : 0];
  return sections[value > 0 ? 0 : value < 0 ? 1 : 2];
}

// --- public entry points --------------------------------------------------

export function format(
  value: FormatValue,
  code: string,
  options: FormatOptions = {},
): FormatResult {
  return formatCompiled(value, parseFormat(code), options);
}

export function formatCompiled(
  value: FormatValue,
  compiled: CompiledFormat,
  options: FormatOptions = {},
): FormatResult {
  const system = options.dateSystem ?? 1900;

  if (value === null || value === undefined) return { text: '', numeric: false, overflow: false };
  if (value instanceof CellError) return { text: value.code, numeric: false, overflow: false };
  // Excel renders logicals as their uppercase names whatever the format says.
  if (typeof value === 'boolean') {
    return { text: value ? 'TRUE' : 'FALSE', numeric: false, overflow: false };
  }
  if (typeof value === 'string') return formatText(value, compiled);

  const serial = value instanceof Date ? dateToSerial(value, system) : value;
  if (!Number.isFinite(serial)) {
    return { text: '#NUM!', numeric: true, overflow: false };
  }

  const section = chooseSection(compiled, serial);
  if (!section) return { text: '', numeric: true, overflow: true };
  const color = section.color;

  if (section.kind === 'general') {
    return withColor(formatGeneral(serial), color, true, false);
  }
  if (section.kind === 'date') {
    const text = renderDateSection(section, serial, system);
    if (text === undefined) return { text: '', numeric: true, overflow: true, ...(color ? { color } : {}) };
    return withColor(text, color, true, false);
  }

  const rendered = renderNumberSection(section, Math.abs(serial));
  // The automatic minus appears only when the format offers no negative section
  // of its own, and never when everything rounded away to zero.
  const needsSign =
    serial < 0 &&
    rendered.significant &&
    section.hasPlaceholders &&
    compiled.numericSections.length === 1;
  return withColor((needsSign ? '-' : '') + rendered.text, color, true, false);
}

function formatText(value: string, compiled: CompiledFormat): FormatResult {
  const section = compiled.textSection;
  if (!section || section.kind !== 'number') {
    return { text: value, numeric: false, overflow: false };
  }
  let out = '';
  for (const piece of section.pieces) {
    const lex = piece.lex;
    if (lex.k === 'at' || lex.k === 'general') out += value;
    else if (lex.k === 'lit') out += lex.s;
    else if (lex.k === 'skip') out += ' ';
  }
  return withColor(out, section.color, false, false);
}

function withColor(
  text: string,
  color: FormatColor | undefined,
  numeric: boolean,
  overflow: boolean,
): FormatResult {
  return color ? { text, color, numeric, overflow } : { text, numeric, overflow };
}

/**
 * Whether a format code makes its cell a date.
 *
 * In xlsx this is the only place date-ness is recorded: the cell itself just
 * holds a number, so a reader that skips this check turns every date column
 * into five-digit integers.
 */
export function isDateFormat(code: string): boolean {
  return parseFormat(code).isDate;
}

// --- column-width overflow ------------------------------------------------

/**
 * The '#' fill Excel paints when a number will not fit its column.
 *
 * Kept out of `format` on purpose: the formatter has no business knowing about
 * column widths, and the renderer is the only layer that knows the real one.
 */
export function overflowText(widthChars: number): string {
  return '#'.repeat(Math.max(1, Math.floor(widthChars)));
}

/**
 * Apply the '####' rule to an already-formatted result.
 *
 * Only numbers overflow this way. Text simply spills into the next cell or is
 * clipped, which is a layout decision, so text is handed back untouched.
 */
export function fitToWidth(result: FormatResult, widthChars: number): string {
  if (!result.numeric) return result.text;
  if (result.overflow) return overflowText(widthChars);
  return result.text.length > widthChars ? overflowText(widthChars) : result.text;
}
