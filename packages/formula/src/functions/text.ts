/**
 * Text functions.
 *
 * Design decisions worth stating once, because they recur through the file:
 *
 * Lengths and offsets are UTF-16 code units, not code points. Excel counts an
 * emoji as two characters and LEFT("<emoji>x",2) hands back the emoji, and a
 * formula that disagrees with Excel about LEN is a worse bug than one that
 * disagrees with Unicode. Every index here is therefore a plain JS string index.
 *
 * FIND and SEARCH are deliberately different functions, not one with a flag:
 * FIND is case-sensitive and literal, SEARCH is case-insensitive and honours the
 * `*` and `?` wildcards. Both are 1-based and report #VALUE! - not #N/A - when
 * the needle is absent, which is the error Excel actually returns.
 *
 * CHAR and CODE work in Windows-1252, not Latin-1. Excel's CHAR(128) is the euro
 * sign because the Windows ANSI code page fills 0x80-0x9F, and a workbook that
 * built a bullet with CHAR(149) must keep rendering one.
 *
 * UPPER and LOWER use simple, one-to-one case mapping. JavaScript applies full
 * Unicode case folding, so "ß".toUpperCase() is "SS" - two characters where
 * Excel leaves one. Any mapping that changes the character count is rejected and
 * the original kept.
 *
 * TEXT carries a deliberately small number-format engine. The complete one lives
 * in packages/formats/src/numfmt.ts, and importing it here would point the
 * dependency arrow the wrong way: the formula package must not know about file
 * formats. `setTextFormatter` below is the seam for the application layer to
 * inject the full engine; until it does, the local subset covers the digit
 * placeholders, decimals, thousands separators and scaling commas, percent,
 * scientific and fraction forms, section splitting and the common date and time
 * tokens. Conditions, colours, elapsed time ([h]), fractional seconds and
 * locale identifiers are the full engine's job.
 *
 * The REGEX* functions are specified against RE2, which has no backreferences
 * and no lookbehind. They are implemented over JavaScript's RegExp, which has
 * both, so patterns using either are rejected with #VALUE! rather than silently
 * evaluated under different semantics. Everything RE2 and JavaScript agree on
 * behaves identically; the difference that remains is that JavaScript may
 * backtrack exponentially where RE2 would not.
 */

import {
  CellError,
  type Scalar,
  isError,
  serialToParts,
  weekdayFromSerial,
  partsToSerial,
} from '@mirrorz/core';
import { ArgKind, type FunctionContext, type FunctionSpec, p } from '../registry.js';
import {
  MAX_TEXT_LENGTH,
  type ArrayValue,
  type Value,
  formatNumberForConcat,
  isArray,
  makeArray,
  parseNumericText,
  toBoolean,
  toExcelPrecision,
  toNumber,
  toText,
  wildcardToRegExp,
} from '../value.js';

/* -------------------------------------------------------------------------- */
/* Argument helpers                                                           */
/* -------------------------------------------------------------------------- */

/**
 * An omitted trailing argument arrives as `undefined`; an argument written but
 * left empty (`MID(a,1,)`) arrives as `null`. The two are not the same - Excel
 * reads the first as "use the default" and the second as blank, so
 * `LEFT("abc")` is "a" while `LEFT("abc",)` is "". Every optional parameter
 * below tests for `undefined` rather than falsiness for that reason.
 */
function omitted(v: Value | undefined): boolean {
  return v === undefined;
}

function text(v: Value | undefined): string | CellError {
  return toText((v ?? null) as Scalar);
}

function num(v: Value | undefined): number | CellError {
  return toNumber((v ?? null) as Scalar);
}

/** Excel truncates a fractional count argument towards zero rather than rounding. */
function truncTowardZero(v: number): number {
  return v < 0 ? Math.ceil(v) : Math.floor(v);
}

/** An integer count argument, or the error Excel reports for the argument. */
function count(v: Value | undefined, fallback: number): number | CellError {
  if (omitted(v)) return fallback;
  const n = num(v);
  if (isError(n)) return n;
  if (!Number.isFinite(n)) return CellError.VALUE;
  return truncTowardZero(n);
}

/** Text results are capped; Excel reports #VALUE! rather than truncating. */
function capped(s: string): string | CellError {
  return s.length > MAX_TEXT_LENGTH ? CellError.VALUE : s;
}

/** Flatten an array-shaped argument into the scalars it holds. */
function flatten(v: Value | undefined): Scalar[] {
  if (omitted(v)) return [];
  if (isArray(v)) return [...(v as ArrayValue).data];
  return [v as Scalar];
}

/* -------------------------------------------------------------------------- */
/* Concatenation                                                              */
/* -------------------------------------------------------------------------- */

function concatScalars(values: Scalar[], skipBlanks: boolean): Value {
  let out = '';
  for (const v of values) {
    if (skipBlanks && v === null) continue;
    const t = toText(v);
    if (isError(t)) return t;
    out += t;
  }
  return capped(out);
}

const CONCATENATE: FunctionSpec = {
  name: 'CONCATENATE',
  params: [p.scalar('text1'), p.rest('text', ArgKind.Scalar)],
  broadcast: true,
  summary: 'Join several text values into one.',
  impl: (args) => concatScalars(args.map((a) => (a ?? null) as Scalar), false),
};

const CONCAT: FunctionSpec = {
  name: 'CONCAT',
  params: [p.array('text1'), p.rest('text')],
  summary: 'Join text from ranges and values into one string.',
  impl: (args) => concatScalars(args.flatMap((a) => flatten(a)), true),
};

const TEXTJOIN: FunctionSpec = {
  name: 'TEXTJOIN',
  params: [p.array('delimiter'), p.scalar('ignore_empty'), p.array('text1'), p.rest('text')],
  summary: 'Join text with a delimiter, optionally skipping empty values.',
  impl: (args) => {
    // A range of delimiters is used cyclically, which is how Excel builds
    // "a, b; c, d" style joins from a single call.
    const delimiters: string[] = [];
    for (const d of flatten(args[0])) {
      const t = toText(d);
      if (isError(t)) return t;
      delimiters.push(t);
    }
    if (delimiters.length === 0) delimiters.push('');

    const ignoreEmpty = toBoolean((args[1] ?? null) as Scalar);
    if (isError(ignoreEmpty)) return ignoreEmpty;

    const pieces: string[] = [];
    for (const v of args.slice(2).flatMap((a) => flatten(a))) {
      const t = toText(v);
      if (isError(t)) return t;
      if (ignoreEmpty && t === '') continue;
      pieces.push(t);
    }

    let out = '';
    for (let i = 0; i < pieces.length; i++) {
      if (i > 0) out += delimiters[(i - 1) % delimiters.length]!;
      out += pieces[i]!;
    }
    return capped(out);
  },
};

/* -------------------------------------------------------------------------- */
/* Extraction and measurement                                                 */
/* -------------------------------------------------------------------------- */

const LEFT: FunctionSpec = {
  name: 'LEFT',
  params: [p.scalar('text'), p.scalar('num_chars', true)],
  broadcast: true,
  summary: 'The leftmost characters of a text value.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    const n = count(args[1], 1);
    if (isError(n)) return n;
    if (n < 0) return CellError.VALUE;
    return s.slice(0, n);
  },
};

const RIGHT: FunctionSpec = {
  name: 'RIGHT',
  params: [p.scalar('text'), p.scalar('num_chars', true)],
  broadcast: true,
  summary: 'The rightmost characters of a text value.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    const n = count(args[1], 1);
    if (isError(n)) return n;
    if (n < 0) return CellError.VALUE;
    return n === 0 ? '' : s.slice(Math.max(0, s.length - n));
  },
};

const MID: FunctionSpec = {
  name: 'MID',
  params: [p.scalar('text'), p.scalar('start_num'), p.scalar('num_chars')],
  broadcast: true,
  summary: 'A run of characters taken from the middle of a text value.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    const start = count(args[1], 1);
    if (isError(start)) return start;
    const n = count(args[2], 0);
    if (isError(n)) return n;
    if (start < 1 || n < 0) return CellError.VALUE;
    if (start > s.length) return '';
    return s.slice(start - 1, start - 1 + n);
  },
};

const LEN: FunctionSpec = {
  name: 'LEN',
  params: [p.scalar('text')],
  broadcast: true,
  summary: 'The number of characters in a text value.',
  impl: (args) => {
    const s = text(args[0]);
    return isError(s) ? s : s.length;
  },
};

/* -------------------------------------------------------------------------- */
/* Case and cleanup                                                           */
/* -------------------------------------------------------------------------- */

function simpleCase(s: string, upper: boolean): string {
  const mapped = upper ? s.toUpperCase() : s.toLowerCase();
  if (mapped.length === s.length) return mapped;
  // Full case folding changed the character count somewhere. Redo it one code
  // point at a time and keep the original wherever the mapping is not 1:1.
  let out = '';
  for (const ch of s) {
    const one = upper ? ch.toUpperCase() : ch.toLowerCase();
    out += [...one].length === 1 ? one : ch;
  }
  return out;
}

const UPPER: FunctionSpec = {
  name: 'UPPER',
  params: [p.scalar('text')],
  broadcast: true,
  summary: 'Convert text to upper case.',
  impl: (args) => {
    const s = text(args[0]);
    return isError(s) ? s : simpleCase(s, true);
  },
};

const LOWER: FunctionSpec = {
  name: 'LOWER',
  params: [p.scalar('text')],
  broadcast: true,
  summary: 'Convert text to lower case.',
  impl: (args) => {
    const s = text(args[0]);
    return isError(s) ? s : simpleCase(s, false);
  },
};

const IS_LETTER = /\p{L}/u;

const PROPER: FunctionSpec = {
  name: 'PROPER',
  params: [p.scalar('text')],
  broadcast: true,
  summary: 'Capitalise the first letter of each word.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    // Excel starts a new word after *any* non-letter, not just whitespace, so
    // "o'brien" becomes "O'Brien" and "3rd" becomes "3Rd".
    let out = '';
    let atWordStart = true;
    for (const ch of s) {
      const letter = IS_LETTER.test(ch);
      out += letter ? simpleCase(ch, atWordStart) : ch;
      atWordStart = !letter;
    }
    return out;
  },
};

const TRIM: FunctionSpec = {
  name: 'TRIM',
  params: [p.scalar('text')],
  broadcast: true,
  summary: 'Remove leading, trailing and repeated spaces.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    // Only U+0020. Excel's TRIM leaves tabs and non-breaking spaces alone, which
    // is why pasted web data so often still looks padded after a TRIM.
    return s.split(' ').filter((part) => part !== '').join(' ');
  },
};

const CLEAN: FunctionSpec = {
  name: 'CLEAN',
  params: [p.scalar('text')],
  broadcast: true,
  summary: 'Remove non-printable characters from text.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    // The documented set is the first 32 codes of 7-bit ASCII; DEL (127) stays.
    let out = '';
    for (let i = 0; i < s.length; i++) {
      if (s.charCodeAt(i) >= 32) out += s[i];
    }
    return out;
  },
};

/* -------------------------------------------------------------------------- */
/* Search and replace                                                         */
/* -------------------------------------------------------------------------- */

const SUBSTITUTE: FunctionSpec = {
  name: 'SUBSTITUTE',
  params: [
    p.scalar('text'),
    p.scalar('old_text'),
    p.scalar('new_text'),
    p.scalar('instance_num', true),
  ],
  broadcast: true,
  summary: 'Replace occurrences of one piece of text with another.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    const oldText = text(args[1]);
    if (isError(oldText)) return oldText;
    const newText = text(args[2]);
    if (isError(newText)) return newText;
    if (oldText === '') return s;

    if (omitted(args[3]) || args[3] === null) {
      return capped(s.split(oldText).join(newText));
    }
    const instance = count(args[3], 1);
    if (isError(instance)) return instance;
    if (instance < 1) return CellError.VALUE;

    let from = 0;
    for (let seen = 1; ; seen++) {
      const at = s.indexOf(oldText, from);
      if (at < 0) return s;
      if (seen === instance) {
        return capped(s.slice(0, at) + newText + s.slice(at + oldText.length));
      }
      from = at + oldText.length;
    }
  },
};

const REPLACE: FunctionSpec = {
  name: 'REPLACE',
  params: [
    p.scalar('old_text'),
    p.scalar('start_num'),
    p.scalar('num_chars'),
    p.scalar('new_text'),
  ],
  broadcast: true,
  summary: 'Replace a run of characters identified by position.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    const start = count(args[1], 1);
    if (isError(start)) return start;
    const n = count(args[2], 0);
    if (isError(n)) return n;
    const replacement = text(args[3]);
    if (isError(replacement)) return replacement;
    if (start < 1 || n < 0) return CellError.VALUE;
    // A start beyond the end appends rather than failing.
    const head = s.slice(0, start - 1);
    const tail = s.slice(Math.min(s.length, start - 1 + n));
    return capped(head + replacement + tail);
  },
};

const FIND: FunctionSpec = {
  name: 'FIND',
  params: [p.scalar('find_text'), p.scalar('within_text'), p.scalar('start_num', true)],
  broadcast: true,
  summary: 'Position of one text value inside another, case-sensitive.',
  impl: (args) => {
    const needle = text(args[0]);
    if (isError(needle)) return needle;
    const haystack = text(args[1]);
    if (isError(haystack)) return haystack;
    const start = count(args[2], 1);
    if (isError(start)) return start;
    // Documented: "If start_num is not greater than 0 (zero) or is greater than
    // the length of within_text, the #VALUE! error value is returned." That
    // bound applies to the empty needle too, so FIND("","abc",4) is an error.
    if (start < 1 || start > haystack.length) return CellError.VALUE;
    const at = haystack.indexOf(needle, start - 1);
    return at < 0 ? CellError.VALUE : at + 1;
  },
};

/**
 * Unanchored, case-insensitive matcher for a SEARCH pattern.
 *
 * `wildcardToRegExp` owns the escaping rules and the `~` escape, so this reuses
 * it rather than re-deriving them, and strips the anchors it adds - SEARCH looks
 * for the pattern anywhere, not for a whole-string match. The `^` and `$` it
 * appends are the only unescaped ones in the source, since a literal `^` or `$`
 * in the pattern comes back backslash-escaped.
 */
function searchPattern(pattern: string): RegExp {
  const anchored = wildcardToRegExp(pattern).source;
  return new RegExp(anchored.slice(1, -1), 'i');
}

const SEARCH: FunctionSpec = {
  name: 'SEARCH',
  params: [p.scalar('find_text'), p.scalar('within_text'), p.scalar('start_num', true)],
  broadcast: true,
  summary: 'Position of one text value inside another, ignoring case, with wildcards.',
  impl: (args) => {
    const needle = text(args[0]);
    if (isError(needle)) return needle;
    const haystack = text(args[1]);
    if (isError(haystack)) return haystack;
    const start = count(args[2], 1);
    if (isError(start)) return start;
    // Same bound as FIND: past the last character is #VALUE!, not a miss.
    if (start < 1 || start > haystack.length) return CellError.VALUE;
    const hit = searchPattern(needle).exec(haystack.slice(start - 1));
    return hit === null ? CellError.VALUE : start + hit.index;
  },
};

const REPT: FunctionSpec = {
  name: 'REPT',
  params: [p.scalar('text'), p.scalar('number_times')],
  broadcast: true,
  summary: 'Repeat text a given number of times.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    const n = count(args[1], 0);
    if (isError(n)) return n;
    if (n < 0) return CellError.VALUE;
    if (n === 0 || s === '') return '';
    if (s.length * n > MAX_TEXT_LENGTH) return CellError.VALUE;
    return s.repeat(n);
  },
};

const EXACT: FunctionSpec = {
  name: 'EXACT',
  params: [p.scalar('text1'), p.scalar('text2')],
  broadcast: true,
  summary: 'Whether two text values are identical, including case.',
  impl: (args) => {
    const a = text(args[0]);
    if (isError(a)) return a;
    const b = text(args[1]);
    if (isError(b)) return b;
    return a === b;
  },
};

/* -------------------------------------------------------------------------- */
/* Conversion                                                                 */
/* -------------------------------------------------------------------------- */

/** Currency symbols VALUE tolerates in front of a number. */
const CURRENCY = /^[$£¥€¢]\s*/;

const ISO_DATE = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
const CLOCK = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(AM|PM)?$/i;

const VALUE: FunctionSpec = {
  name: 'VALUE',
  params: [p.scalar('text')],
  broadcast: true,
  summary: 'Convert text that represents a number into a number.',
  impl: (args, ctx) => {
    const raw = args[0] ?? null;
    if (typeof raw === 'number') return raw;
    if (typeof raw === 'boolean') return CellError.VALUE;
    const s = text(raw as Scalar);
    if (isError(s)) return s;

    // A blank cell coerces to zero, but empty *text* is not a number: Excel
    // reports #VALUE! for VALUE(""), which is why =VALUE(TRIM(A1)) fails on an
    // empty cell while =VALUE(A1) does not. parseNumericText maps "" to 0 for
    // arithmetic coercion, so the emptiness has to be caught before it.
    if (raw === null) return 0;
    const body = s.trim().replace(CURRENCY, '');
    if (body === '') return CellError.VALUE;
    const asNumber = parseNumericText(body);
    if (asNumber !== undefined) return asNumber;

    // Excel's VALUE also accepts the date and time text its locale would accept.
    // Only the unambiguous ISO and clock forms are recognised here; the wider
    // locale grammar belongs with DATEVALUE and TIMEVALUE in the date module.
    const date = ISO_DATE.exec(body);
    if (date) {
      return partsToSerial(
        Number(date[1]),
        Number(date[2]),
        Number(date[3]),
        0,
        0,
        0,
        ctx.dateSystem,
      );
    }
    const clock = CLOCK.exec(body);
    if (clock) {
      let hour = Number(clock[1]);
      const minute = Number(clock[2]);
      const second = clock[3] === undefined ? 0 : Number(clock[3]);
      const meridiem = clock[4]?.toUpperCase();
      if (meridiem !== undefined) {
        if (hour < 1 || hour > 12) return CellError.VALUE;
        hour = (hour % 12) + (meridiem === 'PM' ? 12 : 0);
      }
      if (minute > 59 || second >= 60) return CellError.VALUE;
      return toExcelPrecision((hour * 3600 + minute * 60 + second) / 86_400);
    }
    return CellError.VALUE;
  },
};

const NUMBERVALUE: FunctionSpec = {
  name: 'NUMBERVALUE',
  params: [p.scalar('text'), p.scalar('decimal_separator', true), p.scalar('group_separator', true)],
  broadcast: true,
  summary: 'Convert text to a number using explicit decimal and group separators.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;

    const separator = (arg: Value | undefined, fallback: string): string | CellError => {
      if (omitted(arg) || arg === null) return fallback;
      const t = text(arg);
      if (isError(t)) return t;
      // Only the first character counts, and an explicitly empty separator is
      // an error rather than a silent fall back to the default.
      return t === '' ? CellError.VALUE : t[0]!;
    };
    const decimal = separator(args[1], '.');
    if (isError(decimal)) return decimal;
    const group = separator(args[2], ',');
    if (isError(group)) return group;
    if (decimal === group) return CellError.VALUE;

    // Spaces are ignored anywhere in the argument, so " 3 000 " is 3000.
    let body = s.replace(/\s+/g, '');
    if (body === '') return 0;

    // Trailing percent signs compound, exactly as they do written in a formula.
    let percents = 0;
    while (body.endsWith('%')) {
      percents++;
      body = body.slice(0, -1);
    }

    const parts = body.split(decimal);
    if (parts.length > 2) return CellError.VALUE;
    // A group separator to the left of the decimal point is decoration; one to
    // the right of it is malformed.
    if (parts.length === 2 && parts[1]!.includes(group)) return CellError.VALUE;
    const whole = parts[0]!.split(group).join('');
    const fraction = parts[1] ?? '';

    const candidate = `${whole}${parts.length === 2 ? '.' : ''}${fraction}`;
    if (!/^[+-]?(\d+\.?\d*|\.\d+|\d+\.)$/.test(candidate)) return CellError.VALUE;
    const n = Number(candidate);
    if (!Number.isFinite(n)) return CellError.VALUE;
    return toExcelPrecision(n / 100 ** percents);
  },
};

const T: FunctionSpec = {
  name: 'T',
  params: [p.scalar('value')],
  broadcast: true,
  summary: 'The value if it is text, otherwise empty text.',
  impl: (args) => (typeof args[0] === 'string' ? args[0] : ''),
};

/* -------------------------------------------------------------------------- */
/* Character codes                                                            */
/* -------------------------------------------------------------------------- */

/** Windows-1252 fills 0x80-0x9F, where Latin-1 has control characters. */
const CP1252_HIGH = [
  0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021,
  0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008d, 0x017d, 0x008f,
  0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
  0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
] as const;

const CHAR: FunctionSpec = {
  name: 'CHAR',
  params: [p.scalar('number')],
  broadcast: true,
  summary: 'The character with the given code in the Windows ANSI character set.',
  impl: (args) => {
    const n = count(args[0], 0);
    if (isError(n)) return n;
    if (n < 1 || n > 255) return CellError.VALUE;
    return String.fromCharCode(n >= 0x80 && n <= 0x9f ? CP1252_HIGH[n - 0x80]! : n);
  },
};

const CODE: FunctionSpec = {
  name: 'CODE',
  params: [p.scalar('text')],
  broadcast: true,
  summary: 'The Windows ANSI code of the first character of a text value.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    if (s === '') return CellError.VALUE;
    const cp = s.charCodeAt(0);
    if (cp < 0x80 || (cp >= 0xa0 && cp <= 0xff)) return cp;
    const high = CP1252_HIGH.indexOf(cp as (typeof CP1252_HIGH)[number]);
    // Characters outside the code page come back as '?', which is what Excel
    // substitutes when it cannot represent them in the workbook's ANSI set.
    return high >= 0 ? 0x80 + high : 63;
  },
};

const UNICHAR: FunctionSpec = {
  name: 'UNICHAR',
  params: [p.scalar('number')],
  broadcast: true,
  summary: 'The character with the given Unicode code point.',
  impl: (args) => {
    const n = count(args[0], 0);
    if (isError(n)) return n;
    // Lone surrogates are not characters, and Excel refuses them.
    if (n < 1 || n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)) return CellError.VALUE;
    return String.fromCodePoint(n);
  },
};

const UNICODE: FunctionSpec = {
  name: 'UNICODE',
  params: [p.scalar('text')],
  broadcast: true,
  summary: 'The Unicode code point of the first character of a text value.',
  impl: (args) => {
    const s = text(args[0]);
    if (isError(s)) return s;
    if (s === '') return CellError.VALUE;
    return s.codePointAt(0)!;
  },
};

/* -------------------------------------------------------------------------- */
/* Fixed-point rendering shared by DOLLAR, FIXED and TEXT                     */
/* -------------------------------------------------------------------------- */

/**
 * Decimal digits of `abs` rounded to `decimals` places, half away from zero.
 *
 * Rounding runs through `toExcelPrecision` first because Excel decides at
 * fifteen significant digits: TEXT(1.005,"0.00") is "1.01" in Excel, while
 * (1.005).toFixed(2) is "1.00" because the stored double is a hair below.
 */
function fixedDigits(abs: number, decimals: number): { whole: string; fraction: string } {
  const scale = 10 ** decimals;
  const scaled = toExcelPrecision(abs * scale);
  if (!Number.isSafeInteger(Math.round(scaled))) {
    // Beyond 2^53 the digits are no longer meaningful anyway; toFixed keeps the
    // magnitude right, which is all that is left to preserve.
    const s = abs.toFixed(Math.max(0, Math.min(100, decimals)));
    const [w = '0', f = ''] = s.split('.');
    return { whole: w, fraction: f };
  }
  const digits = String(Math.round(scaled)).padStart(decimals + 1, '0');
  return {
    whole: digits.slice(0, digits.length - decimals) || '0',
    fraction: decimals > 0 ? digits.slice(digits.length - decimals) : '',
  };
}

function group3(digits: string): string {
  let out = '';
  for (let i = 0; i < digits.length; i++) {
    if (i > 0 && (digits.length - i) % 3 === 0) out += ',';
    out += digits[i];
  }
  return out;
}

/** Round to a multiple of 10^-decimals, half away from zero. */
function roundTo(value: number, decimals: number): number {
  const scale = 10 ** decimals;
  const scaled = toExcelPrecision(Math.abs(value) * scale);
  const rounded = Math.round(scaled) / scale;
  return value < 0 ? -rounded : rounded;
}

function fixedText(value: number, decimals: number, commas: boolean): string {
  // Negative decimals round to the left of the point and display none.
  const rounded = decimals < 0 ? roundTo(value, decimals) : value;
  const places = Math.max(0, decimals);
  const { whole, fraction } = fixedDigits(Math.abs(rounded), places);
  const body = (commas ? group3(whole) : whole) + (places > 0 ? `.${fraction}` : '');
  return body;
}

const DOLLAR: FunctionSpec = {
  name: 'DOLLAR',
  params: [p.scalar('number'), p.scalar('decimals', true)],
  broadcast: true,
  summary: 'Format a number as currency text.',
  impl: (args) => {
    const v = num(args[0]);
    if (isError(v)) return v;
    const decimals = count(args[1], 2);
    if (isError(decimals)) return decimals;
    if (decimals > 127) return CellError.VALUE;
    const body = fixedText(v, decimals, true);
    // DOLLAR's built-in format parenthesises negatives; the currency symbol is
    // the locale's, which is the dollar sign until locale plumbing exists.
    const negative = roundTo(v, Math.max(0, decimals)) < 0;
    return negative ? `($${body})` : `$${body}`;
  },
};

const FIXED: FunctionSpec = {
  name: 'FIXED',
  params: [p.scalar('number'), p.scalar('decimals', true), p.scalar('no_commas', true)],
  broadcast: true,
  summary: 'Format a number with a fixed number of decimals.',
  impl: (args) => {
    const v = num(args[0]);
    if (isError(v)) return v;
    const decimals = count(args[1], 2);
    if (isError(decimals)) return decimals;
    if (decimals > 127) return CellError.VALUE;
    const noCommas = omitted(args[2]) ? false : toBoolean((args[2] ?? null) as Scalar);
    if (isError(noCommas)) return noCommas;
    const body = fixedText(v, decimals, !noCommas);
    const negative = roundTo(v, Math.max(0, decimals)) < 0;
    return negative ? `-${body}` : body;
  },
};

/* -------------------------------------------------------------------------- */
/* TEXT and its local number-format subset                                    */
/* -------------------------------------------------------------------------- */

/**
 * Injection seam for the complete number-format engine.
 *
 * packages/formats/src/numfmt.ts implements the whole grammar - conditions,
 * colours, elapsed time, locale identifiers. The formula package
 * cannot import it without inverting the dependency, so the application wires it
 * in at startup instead. Everything below is the fallback used until it does.
 */
export type TextFormatter = (
  value: Scalar,
  formatCode: string,
  dateSystem: 1900 | 1904,
) => string | CellError;

let injectedFormatter: TextFormatter | undefined;

export function setTextFormatter(formatter: TextFormatter | undefined): void {
  injectedFormatter = formatter;
}

type NumToken =
  | { t: 'lit'; s: string }
  | { t: 'ph'; c: '0' | '#' | '?' }
  | { t: 'dot' }
  | { t: 'comma' }
  | { t: 'pct' }
  | { t: 'at' };

/** Split a format code into its `;`-separated sections, honouring quotes. */
function splitSections(code: string): string[] {
  const sections: string[] = [];
  let current = '';
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!;
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const stop = end < 0 ? code.length : end;
      current += code.slice(i, stop + 1);
      i = stop;
      continue;
    }
    if (ch === '\\') {
      current += code.slice(i, i + 2);
      i++;
      continue;
    }
    if (ch === ';') {
      sections.push(current);
      current = '';
      continue;
    }
    current += ch;
  }
  sections.push(current);
  return sections;
}

function tokenizeNumber(section: string): NumToken[] {
  const out: NumToken[] = [];
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]!;
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      const stop = end < 0 ? section.length : end;
      out.push({ t: 'lit', s: section.slice(i + 1, stop) });
      i = stop;
    } else if (ch === '\\' || ch === '_') {
      // A backslash escapes the next character; an underscore reserves the width
      // of it, which without proportional metrics is one space.
      out.push({ t: 'lit', s: ch === '\\' ? (section[i + 1] ?? '') : ' ' });
      i++;
    } else if (ch === '*') {
      // Repeat-to-fill is a rendering concern, not a text one.
      i++;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      i = end < 0 ? section.length : end;
    } else if (ch === '0' || ch === '#' || ch === '?') {
      out.push({ t: 'ph', c: ch });
    } else if (ch === '.') {
      out.push({ t: 'dot' });
    } else if (ch === ',') {
      out.push({ t: 'comma' });
    } else if (ch === '%') {
      out.push({ t: 'pct' });
    } else if (ch === '@') {
      out.push({ t: 'at' });
    } else {
      out.push({ t: 'lit', s: ch });
    }
  }
  return out;
}

/**
 * Scientific formats: an `E` or `e` immediately followed by `+` or `-` splits
 * the section into a mantissa and an exponent. The number of integer
 * placeholders in the mantissa sets the exponent's step, which is what makes
 * "##0.0E+0" an engineering format (exponents in multiples of three) while
 * "0.00E+00" is the ordinary one.
 */
function scientificSplit(tokens: NumToken[]): { at: number; sign: '+' | '-' } | undefined {
  for (let i = 0; i < tokens.length - 1; i++) {
    const t = tokens[i]!;
    const next = tokens[i + 1]!;
    if (t.t !== 'lit' || (t.s !== 'E' && t.s !== 'e')) continue;
    if (next.t !== 'lit' || (next.s !== '+' && next.s !== '-')) continue;
    // Only a mantissa with digit placeholders makes this an exponent form.
    if (!tokens.slice(0, i).some((x) => x.t === 'ph')) continue;
    if (!tokens.slice(i + 2).some((x) => x.t === 'ph')) continue;
    return { at: i, sign: next.s };
  }
  return undefined;
}

function countPlaceholders(tokens: NumToken[], integerPart: boolean): number {
  const dotAt = tokens.findIndex((t) => t.t === 'dot');
  const intEnd = dotAt < 0 ? tokens.length : dotAt;
  let n = 0;
  tokens.forEach((t, i) => {
    if (t.t === 'ph' && (i < intEnd) === integerPart) n++;
  });
  return n;
}

function formatScientific(
  value: number,
  tokens: NumToken[],
  split: { at: number; sign: '+' | '-' },
  generalText: string,
): string {
  const mantissaTokens = tokens.slice(0, split.at);
  const exponentTokens = tokens.slice(split.at + 2);
  const intPh = mantissaTokens.filter((t): t is { t: 'ph'; c: '0' | '#' | '?' } => t.t === 'ph');
  const dotAt = mantissaTokens.findIndex((t) => t.t === 'dot');
  const width = Math.max(
    1,
    mantissaTokens.filter((t, i) => t.t === 'ph' && (dotAt < 0 || i < dotAt)).length,
  );
  // A mantissa of bare zeros pins the integer digit count; a '#' or '?' among
  // them steps the exponent in multiples of that width, which is what makes
  // "##0.0E+0" engineering notation. Same rule as packages/formats/numfmt.ts.
  const stepped = intPh.slice(0, width).some((t) => t.c !== '0');
  const fracPlaces = countPlaceholders(mantissaTokens, false);

  let exponent = 0;
  let mantissa = 0;
  if (value !== 0) {
    // toExponential rather than log10: the decimal exponent has to agree with
    // the digits that will actually be printed, and log10(1000) is not exactly 3
    // for every double.
    exponent = Number(Math.abs(value).toExponential().split('e')[1]);
    exponent = stepped ? Math.floor(exponent / width) * width : exponent - (width - 1);
    mantissa = toExcelPrecision(Math.abs(value) / 10 ** exponent);
    // Rounding the mantissa can carry it into the next power of ten.
    const scale = 10 ** fracPlaces;
    if (Math.round(toExcelPrecision(mantissa * scale)) / scale >= 10 ** width) {
      exponent += stepped ? width : 1;
      mantissa = toExcelPrecision(Math.abs(value) / 10 ** exponent);
    }
  }

  const body = formatNumberSection(mantissa, mantissaTokens, generalText);
  const digits = formatNumberSection(Math.abs(exponent), exponentTokens, generalText);
  const sign = exponent < 0 ? '-' : split.sign === '+' ? '+' : '';
  const e = (tokens[split.at] as { t: 'lit'; s: string }).s;
  return `${body}${e}${sign}${digits}`;
}

/**
 * Fraction formats: "# ?/?" and friends. The numerator is the last unbroken run
 * of placeholders before the slash, anything further left is the whole part, and
 * the denominator is either placeholders or a literal number as in "# ?/16".
 */
function fractionSplit(
  tokens: NumToken[],
): { slash: number; numStart: number; numLen: number; denLen: number; fixedDen?: number } | undefined {
  const slash = tokens.findIndex((t) => t.t === 'lit' && t.s === '/');
  if (slash < 0) return undefined;

  let numStart = slash;
  while (numStart > 0 && tokens[numStart - 1]!.t === 'ph') numStart--;
  const numLen = slash - numStart;
  if (numLen === 0) return undefined;

  let denLen = 0;
  while (tokens[slash + 1 + denLen]?.t === 'ph') denLen++;
  if (denLen > 0) return { slash, numStart, numLen, denLen };

  // A literal denominator such as "# ?/16" arrives as one lit token per digit.
  let literal = '';
  for (let i = slash + 1; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (t.t !== 'lit' || !/^\d$/.test(t.s)) break;
    literal += t.s;
  }
  if (literal === '') return undefined;
  return { slash, numStart, numLen, denLen: 0, fixedDen: Number(literal) };
}

/**
 * Best rational approximation of `x` with a denominator no larger than `maxDen`,
 * by continued fractions - the same routine as packages/formats/src/numfmt.ts.
 *
 * Rounding `x * maxDen` would be wrong: with a one-digit denominator, 0.7 is
 * closer to 5/7 than to the 6/9 naive rounding gives.
 */
function bestFraction(x: number, maxDen: number): { num: number; den: number } {
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
        if (ks >= 1 && Math.abs(x - hs / ks) < Math.abs(x - h1 / k1)) return { num: hs, den: ks };
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

function formatFractionSection(
  value: number,
  tokens: NumToken[],
  spec: { slash: number; numStart: number; numLen: number; denLen: number; fixedDen?: number },
  generalText: string,
): string {
  const intPh: ('0' | '#' | '?')[] = [];
  for (let i = 0; i < spec.numStart; i++) {
    const t = tokens[i]!;
    if (t.t === 'ph') intPh.push(t.c);
  }
  const hasInteger = intPh.length > 0;
  let whole = hasInteger ? Math.floor(value) : 0;
  const rest = value - whole;

  let { num, den } =
    spec.fixedDen !== undefined
      ? { num: Math.round(rest * spec.fixedDen), den: spec.fixedDen }
      : bestFraction(rest, 10 ** spec.denLen - 1);
  if (den < 1) den = 1;
  if (hasInteger && num >= den) {
    // The approximation rounded the remainder up to a whole unit: 1.9999 with
    // "# ?/?" is 2, not "1 1/1".
    whole += Math.floor(num / den);
    num %= den;
  }
  // Excel blanks the fraction rather than printing "0/1", keeping a column of
  // fractions aligned.
  const blank = num === 0 && hasInteger;

  const digits = hasInteger ? String(whole) : '';
  const intOut = new Array<string>(intPh.length).fill('');
  if (intPh.length > 0) {
    const forced = digits === '0' && !intPh.includes('0') && !blank ? '' : digits;
    let pos = forced.length;
    for (let k = intPh.length - 1; k >= 1 && pos > 0; k--) {
      intOut[k] = forced[pos - 1]!;
      pos--;
    }
    intOut[0] = forced.slice(0, pos);
    for (let k = 0; k < intPh.length; k++) {
      if (intOut[k] === '' && intPh[k] === '?') intOut[k] = ' ';
    }
  }

  // The numerator is right-aligned in its placeholders and the denominator left-
  // aligned, so "5   1/4  " and "5   3/10 " line up under "# ???/???".
  const numText = blank ? '' : String(num);
  const denText = blank || spec.fixedDen !== undefined ? '' : String(den);
  // One entry per placeholder; a numerator too wide for its placeholders spills
  // into the leftmost one, so "?/?" can still print 13/3.
  const numSlots = new Array<string>(spec.numLen).fill(' ');
  if (numText.length >= spec.numLen) {
    numSlots[0] = numText.slice(0, numText.length - spec.numLen + 1);
    for (let k = 1; k < spec.numLen; k++) {
      numSlots[k] = numText[numText.length - spec.numLen + k]!;
    }
  } else {
    for (let k = 0; k < numText.length; k++) numSlots[spec.numLen - numText.length + k] = numText[k]!;
  }
  if (blank) numSlots.fill(' ');
  const denSlots = new Array<string>(spec.denLen).fill(' ');
  for (let k = 0; k < denText.length && k < spec.denLen; k++) denSlots[k] = denText[k]!;

  let out = '';
  let intSeen = 0;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (i >= spec.numStart && i < spec.slash) {
      out += numSlots[i - spec.numStart] ?? '';
      continue;
    }
    if (i > spec.slash && i <= spec.slash + spec.denLen) {
      out += denSlots[i - spec.slash - 1] ?? '';
      continue;
    }
    switch (t.t) {
      case 'lit':
        out += t.s === '/' && blank ? ' ' : t.s;
        break;
      case 'at':
        out += generalText;
        break;
      case 'ph':
        out += intOut[intSeen++] ?? '';
        break;
      case 'pct':
        out += '%';
        break;
      default:
        break;
    }
  }
  return out;
}

function formatNumberSection(value: number, tokens: NumToken[], generalText: string): string {
  const scientific = scientificSplit(tokens);
  if (scientific) return formatScientific(value, tokens, scientific, generalText);
  const fractionSpec = fractionSplit(tokens);
  if (fractionSpec) return formatFractionSection(value, tokens, fractionSpec, generalText);

  const dotAt = tokens.findIndex((t) => t.t === 'dot');
  const intEnd = dotAt < 0 ? tokens.length : dotAt;
  const lastPh = tokens.reduce((last, t, i) => (t.t === 'ph' ? i : last), -1);

  // A comma is a grouping mark when a digit placeholder still follows it;
  // a comma past the last placeholder - including one after the decimal
  // placeholders, as in "0.0," - scales the value down by a thousand.
  let scale = 0;
  let grouped = false;
  tokens.forEach((t, i) => {
    if (t.t !== 'comma') return;
    if (i > lastPh) scale++;
    else if (i < intEnd) grouped = true;
  });

  const percents = tokens.filter((t) => t.t === 'pct').length;
  let v = value * 100 ** percents;
  if (scale > 0) v /= 1000 ** scale;

  const intPh: ('0' | '#' | '?')[] = [];
  const fracPh: ('0' | '#' | '?')[] = [];
  tokens.forEach((t, i) => {
    if (t.t !== 'ph') return;
    (i < intEnd ? intPh : fracPh).push(t.c);
  });

  const { whole, fraction } = fixedDigits(Math.abs(v), fracPh.length);
  // Zeros are forced only as far left as the leftmost '0' placeholder; '#' and
  // '?' leave an insignificant digit out, '?' holding its width with a space.
  const firstZero = intPh.indexOf('0');
  const minWhole = firstZero < 0 ? 0 : intPh.length - firstZero;
  let digits = whole.replace(/^0+(?=\d)/, '');
  if (digits === '0' && minWhole === 0) digits = '';
  digits = digits.padStart(minWhole, '0');

  // Trailing places that would only show zeros disappear unless a literal '0'
  // holds them, and the decimal point goes with them when nothing at all is left
  // of the fraction. A '?' drops the digit but keeps its width, below.
  let cut = fracPh.length;
  while (cut > 0 && fracPh[cut - 1] !== '0' && fraction[cut - 1] === '0') cut--;

  const intText = grouped && digits !== '' ? group3(digits) : digits;

  // Integer digits fill the placeholders right to left, so "000-0000" lays a
  // seven-digit number out as 555-1234 and the leftmost placeholder absorbs the
  // overflow. A grouped format puts the whole grouped run at the leftmost
  // placeholder instead: a thousands separator has no single placeholder it
  // could belong to.
  const intOut = new Array<string>(intPh.length).fill('');
  if (intPh.length > 0) {
    if (grouped) {
      intOut[0] = intText;
    } else {
      let pos = intText.length;
      for (let k = intPh.length - 1; k >= 1 && pos > 0; k--) {
        intOut[k] = intText[pos - 1]!;
        pos--;
      }
      intOut[0] = intText.slice(0, pos);
    }
    // A '?' that received no digit still reserves its width.
    for (let k = 0; k < intPh.length; k++) {
      if (intOut[k] === '' && intPh[k] === '?') intOut[k] = ' ';
    }
  }

  let out = '';
  let intSeen = 0;
  let fracSeen = 0;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    switch (t.t) {
      case 'lit':
        out += t.s;
        break;
      case 'pct':
        out += '%';
        break;
      case 'at':
        // The text placeholder in a numeric section shows the number the way
        // General would: TEXT(1234.567,"@") is "1234.567".
        out += generalText;
        break;
      case 'comma':
        break;
      case 'dot':
        if (cut > 0 || fracPh.length === 0) out += '.';
        break;
      case 'ph':
        if (i < intEnd) {
          out += intOut[intSeen++] ?? '';
        } else {
          const idx = fracSeen++;
          if (idx < cut) out += fraction[idx] ?? '';
          else if (t.c === '?') out += ' ';
        }
        break;
    }
  }
  return out;
}

const DATE_LETTERS = 'ymdhs';
const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
] as const;
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'] as const;

type DateToken = { t: 'lit'; s: string } | { t: 'run'; s: string } | { t: 'ampm'; s: string };

function tokenizeDate(section: string): DateToken[] {
  const out: DateToken[] = [];
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]!;
    const lower = ch.toLowerCase();
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      const stop = end < 0 ? section.length : end;
      out.push({ t: 'lit', s: section.slice(i + 1, stop) });
      i = stop;
      continue;
    }
    if (ch === '\\') {
      out.push({ t: 'lit', s: section[i + 1] ?? '' });
      i++;
      continue;
    }
    if (ch === '[') {
      const end = section.indexOf(']', i);
      i = end < 0 ? section.length : end;
      continue;
    }
    const rest = section.slice(i);
    const meridiem = /^(AM\/PM|A\/P)/i.exec(rest);
    if (meridiem) {
      out.push({ t: 'ampm', s: meridiem[0]! });
      i += meridiem[0]!.length - 1;
      continue;
    }
    if (DATE_LETTERS.includes(lower)) {
      let run = lower;
      while (section[i + 1]?.toLowerCase() === lower) {
        run += lower;
        i++;
      }
      out.push({ t: 'run', s: run });
      continue;
    }
    out.push({ t: 'lit', s: ch });
  }
  return out;
}

/**
 * `m` means months or minutes depending on its neighbours: directly after an
 * hour token or directly before a seconds token it is minutes. This is the one
 * genuinely context-sensitive rule in the format grammar.
 */
function isMinuteRun(tokens: DateToken[], index: number): boolean {
  for (let i = index - 1; i >= 0; i--) {
    const t = tokens[i]!;
    if (t.t === 'lit') continue;
    if (t.t === 'run' && t.s[0] === 'h') return true;
    break;
  }
  for (let i = index + 1; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (t.t === 'lit') continue;
    if (t.t === 'run' && t.s[0] === 's') return true;
    break;
  }
  return false;
}

function looksLikeDate(section: string): boolean {
  const runs = tokenizeDate(section).filter(
    (t): t is { t: 'run'; s: string } => t.t === 'run',
  );
  if (runs.length === 0) return false;
  if (runs.some((r) => r.s[0] !== 'm')) return true;
  // A lone `m` is a month only when no digit placeholder claims the section,
  // which is what separates "m/d" from "0.0 m" style codes.
  return !tokenizeNumber(section).some((t) => t.t === 'ph');
}

function pad2(n: number): string {
  return String(n).padStart(2, '0');
}

function formatDateSection(serial: number, section: string, system: 1900 | 1904): string | CellError {
  if (serial < 0) return CellError.VALUE;
  const tokens = tokenizeDate(section);
  const parts = serialToParts(serial, system);
  const twelveHour = tokens.some((t) => t.t === 'ampm');
  const dow = weekdayFromSerial(serial, system);

  let out = '';
  tokens.forEach((token, index) => {
    if (token.t === 'lit') {
      out += token.s;
      return;
    }
    if (token.t === 'ampm') {
      const pm = parts.hour >= 12;
      const marker = token.s.toUpperCase() === 'A/P' ? (pm ? 'P' : 'A') : pm ? 'PM' : 'AM';
      // "AM/PM" prints upper case, "am/pm" lower: the code's own case wins.
      out += token.s === token.s.toLowerCase() ? marker.toLowerCase() : marker;
      return;
    }
    const run = token.s;
    switch (run[0]) {
      case 'y':
        out += run.length <= 2 ? pad2(parts.year % 100) : String(parts.year).padStart(4, '0');
        break;
      case 'd':
        if (run.length === 1) out += String(parts.day);
        else if (run.length === 2) out += pad2(parts.day);
        else if (run.length === 3) out += DAYS[dow]!.slice(0, 3);
        else out += DAYS[dow]!;
        break;
      case 'h': {
        const hour = twelveHour ? parts.hour % 12 || 12 : parts.hour;
        out += run.length === 1 ? String(hour) : pad2(hour);
        break;
      }
      case 's':
        out += run.length === 1 ? String(parts.second) : pad2(parts.second);
        break;
      case 'm':
        if (isMinuteRun(tokens, index)) {
          out += run.length === 1 ? String(parts.minute) : pad2(parts.minute);
        } else if (run.length === 1) out += String(parts.month);
        else if (run.length === 2) out += pad2(parts.month);
        else if (run.length === 3) out += MONTHS[parts.month - 1]!.slice(0, 3);
        else if (run.length === 4) out += MONTHS[parts.month - 1]!;
        else out += MONTHS[parts.month - 1]!.slice(0, 1);
        break;
    }
  });
  return out;
}

function applyFormat(value: Scalar, code: string, system: 1900 | 1904): string | CellError {
  if (injectedFormatter) return injectedFormatter(value, code, system);

  // Excel accepts at most four sections and drops the rest.
  const sections = splitSections(code).slice(0, 4);
  const carriesAt = (sec: string): boolean => tokenizeNumber(sec).some((t) => t.t === 'at');
  const last = sections[sections.length - 1] ?? '';
  // The text section is normally the fourth. It is also any *final* section
  // carrying an '@', because "0.00;@" - which openpyxl and Excel both emit -
  // means "numbers here, text there", not "positives here, negatives there".
  // The same rule as packages/formats/src/numfmt.ts, which will replace this.
  const trailingText = sections.length > 1 && carriesAt(last);
  const numeric = trailingText ? sections.slice(0, -1) : sections;
  const textSection = trailingText
    ? last
    : (sections[3] ?? (sections.length === 1 && carriesAt(last) ? last : undefined));

  let subject: Scalar = value;
  if (typeof subject === 'string') {
    // Numeric text is formatted as the number it spells: TEXT("5","0.00") is
    // "5.00", the same coercion every other numeric argument gets.
    const asNumber = subject === '' ? undefined : parseNumericText(subject);
    if (asNumber === undefined) {
      // Without a text section the text passes through untouched, which is why
      // TEXT("abc","0.00") is "abc" rather than an error.
      if (textSection === undefined) return subject;
      const body = subject;
      return tokenizeNumber(textSection)
        .map((t) => (t.t === 'at' ? body : t.t === 'lit' ? t.s : t.t === 'pct' ? '%' : ''))
        .join('');
    }
    subject = asNumber;
  }

  const n = toNumber(subject);
  if (isError(n)) return n;

  let section = numeric[0] ?? '';
  let magnitude = n;
  if (n < 0 && numeric.length > 1) {
    section = numeric[1]!;
    // The negative section supplies its own sign, if it wants one.
    magnitude = -n;
  } else if (n === 0 && numeric.length > 2) {
    section = numeric[2]!;
  }

  if (/^\s*general\s*$/i.test(section)) return formatNumberForConcat(magnitude);
  if (looksLikeDate(section)) return formatDateSection(magnitude, section, system);

  const body = formatNumberSection(
    Math.abs(magnitude),
    tokenizeNumber(section),
    formatNumberForConcat(Math.abs(magnitude)),
  );
  return magnitude < 0 ? `-${body}` : body;
}

const TEXT: FunctionSpec = {
  name: 'TEXT',
  params: [p.scalar('value'), p.scalar('format_text')],
  broadcast: true,
  summary: 'Format a value as text using a number-format code.',
  impl: (args, ctx: FunctionContext) => {
    const code = text(args[1]);
    if (isError(code)) return code;
    const value = (args[0] ?? null) as Scalar;
    if (isError(value)) return value;
    return applyFormat(value === null ? 0 : value, code, ctx.dateSystem);
  },
};

/* -------------------------------------------------------------------------- */
/* Splitting                                                                  */
/* -------------------------------------------------------------------------- */

interface DelimiterHit {
  start: number;
  end: number;
}

function delimiterList(v: Value | undefined): string[] | CellError {
  const out: string[] = [];
  for (const item of flatten(v)) {
    const t = toText(item);
    if (isError(t)) return t;
    // An empty delimiter can never match; keeping it would make every position a
    // hit and every result empty.
    if (t !== '') out.push(t);
  }
  return out;
}

/** Left-to-right, non-overlapping delimiter positions. */
function findDelimiters(body: string, delimiters: string[], ignoreCase: boolean): DelimiterHit[] {
  const hay = ignoreCase ? body.toLowerCase() : body;
  const needles = ignoreCase ? delimiters.map((d) => d.toLowerCase()) : delimiters;
  const hits: DelimiterHit[] = [];
  for (let i = 0; i < hay.length; ) {
    let matched = 0;
    for (const needle of needles) {
      if (needle.length > 0 && hay.startsWith(needle, i)) {
        matched = needle.length;
        break;
      }
    }
    if (matched > 0) {
      hits.push({ start: i, end: i + matched });
      i += matched;
    } else {
      i++;
    }
  }
  return hits;
}

function splitOn(body: string, delimiters: string[], ignoreCase: boolean): string[] {
  if (delimiters.length === 0) return [body];
  const hits = findDelimiters(body, delimiters, ignoreCase);
  const parts: string[] = [];
  let from = 0;
  for (const hit of hits) {
    parts.push(body.slice(from, hit.start));
    from = hit.end;
  }
  parts.push(body.slice(from));
  return parts;
}

function aroundDelimiter(args: Value[], before: boolean): Value {
  const body = text(args[0]);
  if (isError(body)) return body;
  const delimiters = delimiterList(args[1]);
  if (isError(delimiters)) return delimiters;
  const instance = count(args[2], 1);
  if (isError(instance)) return instance;
  if (instance === 0) return CellError.VALUE;
  const matchMode = count(args[3], 0);
  if (isError(matchMode)) return matchMode;
  const matchEnd = count(args[4], 0);
  if (isError(matchEnd)) return matchEnd;
  const notFound = omitted(args[5]) ? CellError.NA : ((args[5] ?? null) as Scalar);

  if (body === '') return '';

  const hits = findDelimiters(body, delimiters, matchMode !== 0);
  // match_end treats the far end of the text as one more delimiter, so
  // TEXTBEFORE("a-b","-",2,,1) yields the whole string.
  if (matchEnd !== 0) hits.push({ start: body.length, end: body.length });

  const index = instance > 0 ? instance - 1 : hits.length + instance;
  const hit = hits[index];
  if (hit === undefined) return notFound;
  return before ? body.slice(0, hit.start) : body.slice(hit.end);
}

const TEXTBEFORE: FunctionSpec = {
  name: 'TEXTBEFORE',
  params: [
    p.scalar('text'),
    p.array('delimiter'),
    p.scalar('instance_num', true),
    p.scalar('match_mode', true),
    p.scalar('match_end', true),
    p.scalar('if_not_found', true),
  ],
  broadcast: true,
  summary: 'The text before a given occurrence of a delimiter.',
  impl: (args) => aroundDelimiter(args, true),
};

const TEXTAFTER: FunctionSpec = {
  name: 'TEXTAFTER',
  params: [
    p.scalar('text'),
    p.array('delimiter'),
    p.scalar('instance_num', true),
    p.scalar('match_mode', true),
    p.scalar('match_end', true),
    p.scalar('if_not_found', true),
  ],
  broadcast: true,
  summary: 'The text after a given occurrence of a delimiter.',
  impl: (args) => aroundDelimiter(args, false),
};

const TEXTSPLIT: FunctionSpec = {
  name: 'TEXTSPLIT',
  params: [
    p.scalar('text'),
    p.array('col_delimiter'),
    p.array('row_delimiter', true),
    p.scalar('ignore_empty', true),
    p.scalar('match_mode', true),
    p.scalar('pad_with', true),
  ],
  summary: 'Split text into a two-dimensional array of values.',
  impl: (args) => {
    const body = text(args[0]);
    if (isError(body)) return body;
    const colDelims = delimiterList(args[1]);
    if (isError(colDelims)) return colDelims;
    const rowDelims = delimiterList(args[2]);
    if (isError(rowDelims)) return rowDelims;
    if (colDelims.length === 0 && rowDelims.length === 0) return CellError.VALUE;
    const ignoreEmpty = omitted(args[3]) ? false : toBoolean((args[3] ?? null) as Scalar);
    if (isError(ignoreEmpty)) return ignoreEmpty;
    const matchMode = count(args[4], 0);
    if (isError(matchMode)) return matchMode;
    const pad = omitted(args[5]) ? CellError.NA : ((args[5] ?? null) as Scalar);

    const ignoreCase = matchMode !== 0;
    const lines = splitOn(body, rowDelims, ignoreCase);
    let grid = lines.map((line) => splitOn(line, colDelims, ignoreCase));
    if (ignoreEmpty) {
      grid = grid.map((row) => row.filter((cell) => cell !== '')).filter((row) => row.length > 0);
    }
    if (grid.length === 0) grid = [['']];

    const cols = Math.max(...grid.map((row) => row.length));
    const data: Scalar[] = [];
    for (const row of grid) {
      for (let c = 0; c < cols; c++) data.push(c < row.length ? row[c]! : pad);
    }
    return makeArray(grid.length, cols, data);
  },
};

/* -------------------------------------------------------------------------- */
/* Value rendering                                                            */
/* -------------------------------------------------------------------------- */

function renderScalar(v: Scalar, strict: boolean): string {
  if (isError(v)) return v.code;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return formatNumberForConcat(v);
  return strict ? `"${v.replaceAll('"', '""')}"` : v;
}

const VALUETOTEXT: FunctionSpec = {
  name: 'VALUETOTEXT',
  params: [{ name: 'value', kind: ArgKind.Scalar, errorTransparent: true }, p.scalar('format', true)],
  broadcast: true,
  summary: 'Render a value as text, optionally in strict form.',
  impl: (args) => {
    const mode = count(args[1], 0);
    if (isError(mode)) return mode;
    if (mode !== 0 && mode !== 1) return CellError.VALUE;
    return renderScalar((args[0] ?? null) as Scalar, mode === 1);
  },
};

const ARRAYTOTEXT: FunctionSpec = {
  name: 'ARRAYTOTEXT',
  params: [{ name: 'array', kind: ArgKind.Array, errorTransparent: true }, p.scalar('format', true)],
  summary: 'Render an array as text, optionally in strict array-literal form.',
  impl: (args) => {
    const mode = count(args[1], 0);
    if (isError(mode)) return mode;
    if (mode !== 0 && mode !== 1) return CellError.VALUE;
    const strict = mode === 1;

    const arg = args[0];
    if (!isArray(arg)) return renderScalar((arg ?? null) as Scalar, strict);

    const a = arg as ArrayValue;
    if (!strict) {
      return capped(a.data.map((v) => renderScalar(v, false)).join(', '));
    }
    const rows: string[] = [];
    for (let r = 0; r < a.rows; r++) {
      const cells: string[] = [];
      for (let c = 0; c < a.cols; c++) cells.push(renderScalar(a.data[r * a.cols + c] ?? null, true));
      rows.push(cells.join(','));
    }
    return capped(`{${rows.join(';')}}`);
  },
};

/* -------------------------------------------------------------------------- */
/* Regular expressions                                                        */
/* -------------------------------------------------------------------------- */

/**
 * Compile a pattern under RE2-compatible restrictions.
 *
 * RE2 has no backreferences and no lookbehind, so a pattern using either would
 * mean something in our engine that it cannot mean in Excel's. Rejecting is the
 * honest answer; silently accepting would make the divergence invisible.
 */
function compilePattern(pattern: string, ignoreCase: boolean, global: boolean): RegExp | CellError {
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]!;
    if (ch === '\\') {
      const next = pattern[i + 1];
      if (next !== undefined && (/[1-9]/.test(next) || next === 'k')) return CellError.VALUE;
      i++;
      continue;
    }
    if (ch === '(' && (pattern.startsWith('(?<=', i) || pattern.startsWith('(?<!', i))) {
      return CellError.VALUE;
    }
  }
  try {
    return new RegExp(pattern, `${ignoreCase ? 'i' : ''}${global ? 'g' : ''}`);
  } catch {
    return CellError.VALUE;
  }
}

function caseFlag(v: Value | undefined): boolean | CellError {
  const mode = count(v, 0);
  if (isError(mode)) return mode;
  if (mode !== 0 && mode !== 1) return CellError.VALUE;
  return mode === 1;
}

const REGEXTEST: FunctionSpec = {
  name: 'REGEXTEST',
  params: [p.scalar('text'), p.scalar('pattern'), p.scalar('case_sensitivity', true)],
  broadcast: true,
  summary: 'Whether text matches a regular expression.',
  impl: (args) => {
    const body = text(args[0]);
    if (isError(body)) return body;
    const pattern = text(args[1]);
    if (isError(pattern)) return pattern;
    const ignoreCase = caseFlag(args[2]);
    if (isError(ignoreCase)) return ignoreCase;
    const re = compilePattern(pattern, ignoreCase, false);
    if (isError(re)) return re;
    return re.test(body);
  },
};

const REGEXEXTRACT: FunctionSpec = {
  name: 'REGEXEXTRACT',
  params: [
    p.scalar('text'),
    p.scalar('pattern'),
    p.scalar('return_mode', true),
    p.scalar('case_sensitivity', true),
  ],
  summary: 'Extract the parts of text that match a regular expression.',
  impl: (args) => {
    const body = text(args[0]);
    if (isError(body)) return body;
    const pattern = text(args[1]);
    if (isError(pattern)) return pattern;
    const mode = count(args[2], 0);
    if (isError(mode)) return mode;
    if (mode < 0 || mode > 2) return CellError.VALUE;
    const ignoreCase = caseFlag(args[3]);
    if (isError(ignoreCase)) return ignoreCase;

    const re = compilePattern(pattern, ignoreCase, mode === 1);
    if (isError(re)) return re;

    if (mode === 1) {
      const all = [...body.matchAll(re)].map((m) => m[0]);
      if (all.length === 0) return CellError.NA;
      return makeArray(all.length, 1, all);
    }
    const hit = re.exec(body);
    if (hit === null) return CellError.NA;
    if (mode === 0) return hit[0];
    const groups = hit.slice(1).map((g) => g ?? '');
    if (groups.length === 0) return CellError.NA;
    return makeArray(1, groups.length, groups);
  },
};

const REGEXREPLACE: FunctionSpec = {
  name: 'REGEXREPLACE',
  params: [
    p.scalar('text'),
    p.scalar('pattern'),
    p.scalar('replacement'),
    p.scalar('occurrence', true),
    p.scalar('case_sensitivity', true),
  ],
  broadcast: true,
  summary: 'Replace the parts of text that match a regular expression.',
  impl: (args) => {
    const body = text(args[0]);
    if (isError(body)) return body;
    const pattern = text(args[1]);
    if (isError(pattern)) return pattern;
    const replacement = text(args[2]);
    if (isError(replacement)) return replacement;
    const occurrence = count(args[3], 0);
    if (isError(occurrence)) return occurrence;
    const ignoreCase = caseFlag(args[4]);
    if (isError(ignoreCase)) return ignoreCase;

    const re = compilePattern(pattern, ignoreCase, true);
    if (isError(re)) return re;

    const matches = [...body.matchAll(re)];
    if (matches.length === 0) return body;

    // occurrence 0 replaces every match; a negative one counts from the end.
    const chosen =
      occurrence === 0
        ? matches
        : (() => {
            const index = occurrence > 0 ? occurrence - 1 : matches.length + occurrence;
            const one = matches[index];
            return one === undefined ? [] : [one];
          })();
    if (chosen.length === 0) return body;

    let out = '';
    let from = 0;
    for (const m of chosen) {
      const at = m.index ?? 0;
      out += body.slice(from, at);
      out += expandReplacement(replacement, m);
      from = at + m[0]!.length;
    }
    out += body.slice(from);
    return capped(out);
  },
};

/** `$1`-style group references in a replacement, plus `$$` for a literal dollar. */
function expandReplacement(replacement: string, match: RegExpMatchArray): string {
  let out = '';
  for (let i = 0; i < replacement.length; i++) {
    const ch = replacement[i]!;
    if (ch !== '$') {
      out += ch;
      continue;
    }
    const next = replacement[i + 1];
    if (next === '$') {
      out += '$';
      i++;
      continue;
    }
    const digits = /^\d{1,2}/.exec(replacement.slice(i + 1));
    if (digits === null) {
      out += ch;
      continue;
    }
    out += match[Number(digits[0])] ?? '';
    i += digits[0].length;
  }
  return out;
}

export const TEXT_FUNCTIONS: readonly FunctionSpec[] = [
  CONCATENATE,
  CONCAT,
  TEXTJOIN,
  LEFT,
  RIGHT,
  MID,
  LEN,
  LOWER,
  UPPER,
  PROPER,
  TRIM,
  CLEAN,
  SUBSTITUTE,
  REPLACE,
  FIND,
  SEARCH,
  REPT,
  EXACT,
  VALUE,
  NUMBERVALUE,
  T,
  CHAR,
  CODE,
  UNICHAR,
  UNICODE,
  DOLLAR,
  FIXED,
  TEXTBEFORE,
  TEXTAFTER,
  TEXTSPLIT,
  ARRAYTOTEXT,
  VALUETOTEXT,
  REGEXTEST,
  REGEXEXTRACT,
  REGEXREPLACE,
  TEXT,
];
