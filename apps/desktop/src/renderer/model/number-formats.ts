/**
 * The number-format menu.
 *
 * Excel's format dialogue is a category list, a sample box and a code field, and
 * the code field is where almost everyone eventually ends up because the
 * presets never quite cover the case at hand. So the presets here are a short
 * list of the ones people actually pick, every one of them shows a live sample
 * rendered by the real format engine rather than a hard-coded string, and the
 * custom code is a first-class entry rather than a dialogue three clicks deep.
 *
 * The sample matters more than it looks: the difference between `0.00` and
 * `#,##0.00` is invisible in a menu of names and obvious in a menu of samples.
 */

import { format, isDateFormat } from '@mirrorz/formats';
import type { Scalar } from '@mirrorz/core';

export type FormatCategory = 'general' | 'number' | 'currency' | 'date' | 'time' | 'text' | 'custom';

export interface FormatPreset {
  id: string;
  label: string;
  category: FormatCategory;
  /** The format code, or undefined for General, which has no code. */
  code?: string;
  /** The value the sample renders, chosen to show what the format does. */
  sample?: number;
}

/** A date serial and a time of day that exercise every field of a date format. */
const SAMPLE_DATE = 45_000.5 + 1 / 24;
const SAMPLE_NUMBER = 1234.567;
const SAMPLE_NEGATIVE = -1234.567;

export const FORMAT_PRESETS: readonly FormatPreset[] = [
  { id: 'general', label: 'General', category: 'general', sample: SAMPLE_NUMBER },
  { id: 'integer', label: 'Number', category: 'number', code: '0', sample: SAMPLE_NUMBER },
  { id: 'decimal2', label: 'Number, 2 decimals', category: 'number', code: '0.00', sample: SAMPLE_NUMBER },
  { id: 'thousands', label: 'Thousands separator', category: 'number', code: '#,##0.00', sample: SAMPLE_NUMBER },
  { id: 'percent', label: 'Percent', category: 'number', code: '0.00%', sample: 0.1234 },
  { id: 'scientific', label: 'Scientific', category: 'number', code: '0.00E+00', sample: SAMPLE_NUMBER },
  { id: 'fraction', label: 'Fraction', category: 'number', code: '# ?/?', sample: 1.25 },
  { id: 'currency', label: 'Currency', category: 'currency', code: '"$"#,##0.00', sample: SAMPLE_NUMBER },
  {
    id: 'currency-red',
    label: 'Currency, negatives in red',
    category: 'currency',
    code: '"$"#,##0.00;[Red]-"$"#,##0.00',
    sample: SAMPLE_NEGATIVE,
  },
  {
    id: 'accounting',
    label: 'Accounting',
    category: 'currency',
    code: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    sample: SAMPLE_NUMBER,
  },
  { id: 'date-short', label: 'Short date', category: 'date', code: 'yyyy-mm-dd', sample: SAMPLE_DATE },
  { id: 'date-long', label: 'Long date', category: 'date', code: 'dddd, d mmmm yyyy', sample: SAMPLE_DATE },
  { id: 'date-month', label: 'Month and year', category: 'date', code: 'mmm yyyy', sample: SAMPLE_DATE },
  { id: 'time', label: 'Time', category: 'time', code: 'h:mm:ss', sample: SAMPLE_DATE },
  { id: 'time-ampm', label: 'Time, 12 hour', category: 'time', code: 'h:mm AM/PM', sample: SAMPLE_DATE },
  { id: 'datetime', label: 'Date and time', category: 'date', code: 'yyyy-mm-dd hh:mm', sample: SAMPLE_DATE },
  { id: 'duration', label: 'Elapsed hours', category: 'time', code: '[h]:mm:ss', sample: 1.5 },
  { id: 'text', label: 'Text', category: 'text', code: '@', sample: SAMPLE_NUMBER },
];

/**
 * Render a preset's sample. A format that throws is a bug in the format engine
 * rather than in the menu, but a menu that throws takes the whole toolbar with
 * it, so a failed sample degrades to the raw code.
 */
export function presetSample(preset: FormatPreset): string {
  const value: Scalar = preset.sample ?? SAMPLE_NUMBER;
  if (preset.code === undefined) return formatSafely(value, 'General');
  return formatSafely(value, preset.code);
}

export function formatSafely(value: Scalar, code: string): string {
  try {
    return format(value, code).text;
  } catch {
    return code;
  }
}

/** The preset a format code corresponds to, or undefined for a custom code. */
export function presetForCode(code: string | undefined): FormatPreset | undefined {
  if (code === undefined || code === '' || code.toLowerCase() === 'general') {
    return FORMAT_PRESETS[0];
  }
  return FORMAT_PRESETS.find((p) => p.code === code);
}

/** Menu entries grouped by category, in the order the menu shows them. */
export function presetsByCategory(): { category: FormatCategory; presets: FormatPreset[] }[] {
  const order: FormatCategory[] = ['general', 'number', 'currency', 'date', 'time', 'text'];
  return order
    .map((category) => ({ category, presets: FORMAT_PRESETS.filter((p) => p.category === category) }))
    .filter((group) => group.presets.length > 0);
}

export interface FormatValidation {
  valid: boolean
  /** Why it was rejected, in one line, for the field's error text. */
  problem?: string;
  /** How the sample value renders under this code, when it is valid. */
  preview?: string;
  isDate?: boolean;
}

/**
 * Check a custom format code by running it.
 *
 * Validating format codes by pattern is a losing game - the grammar has four
 * sections, colour and condition prefixes, fill and skip-width operators, and
 * locale tags - so the honest check is to compile it and see. An empty code is
 * rejected explicitly because the format engine treats it as General, and
 * silently doing something other than what the field says is worse than an
 * error message.
 */
export function validateFormatCode(code: string): FormatValidation {
  const trimmed = code.trim();
  if (trimmed === '') return { valid: false, problem: 'Enter a format code, or choose General' };
  // More than four sections is the one structural rule worth stating, because
  // the error the engine gives for it is otherwise obscure.
  if (countSections(trimmed) > 4) {
    return { valid: false, problem: 'A format has at most four sections: positive; negative; zero; text' };
  }
  try {
    const isDate = isDateFormat(trimmed);
    const preview = format(isDate ? SAMPLE_DATE : SAMPLE_NUMBER, trimmed).text;
    return { valid: true, preview, isDate };
  } catch (err) {
    return { valid: false, problem: err instanceof Error ? err.message : 'Not a valid format code' };
  }
}

/** Count top-level `;` separators, ignoring ones inside quotes or brackets. */
function countSections(code: string): number {
  let sections = 1;
  let inQuote = false;
  let bracket = 0;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!;
    if (ch === '"') inQuote = !inQuote;
    else if (inQuote) continue;
    else if (ch === '[') bracket++;
    else if (ch === ']') bracket = Math.max(0, bracket - 1);
    else if (ch === '\\') i++;
    else if (ch === ';' && bracket === 0) sections++;
  }
  return sections;
}

/**
 * Step a format's decimal places up or down, for the two toolbar buttons.
 *
 * The buttons have to work on whatever code the cell already has, including a
 * custom one, so this edits the code's existing decimal run rather than
 * replacing the code with a preset and throwing away the user's currency symbol
 * or thousands separator.
 */
export function adjustDecimals(code: string | undefined, delta: number): string {
  const base = code === undefined || code === '' || code.toLowerCase() === 'general' ? '0' : code;
  const dot = base.indexOf('.');
  if (dot < 0) {
    if (delta <= 0) return base;
    return insertDecimals(base, delta);
  }
  let end = dot + 1;
  while (end < base.length && base[end] === '0') end++;
  const current = end - dot - 1;
  const next = Math.max(0, Math.min(30, current + delta));
  const zeros = '0'.repeat(next);
  const head = base.slice(0, dot);
  const tail = base.slice(end);
  return next === 0 ? head + tail : `${head}.${zeros}${tail}`;
}

/** Put a decimal run after the last digit placeholder of the integer part. */
function insertDecimals(code: string, places: number): string {
  const match = /[0#?](?![\s\S]*[0#?])/.exec(code);
  if (!match) return `${code}.${'0'.repeat(places)}`;
  const at = match.index + 1;
  return `${code.slice(0, at)}.${'0'.repeat(places)}${code.slice(at)}`;
}
