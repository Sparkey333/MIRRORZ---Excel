/**
 * What the import dialogue shows, computed before anything is committed.
 *
 * `reviewImport` in core answers "what would happen", per cell. What a person
 * needs in order to decide is coarser and per column: this column is going to
 * become dates, that one has leading zeros we are keeping as text, and here is
 * the switch to change either decision. So this module aggregates the per-cell
 * review into per-column facts, and holds the overrides the user sets before the
 * paste happens rather than offering an undo afterwards.
 *
 * The reason this dialogue exists at all: the destructive conversions - dates
 * from gene names, integers from leading-zero codes, floats from long ids - are
 * silent and irreversible in every spreadsheet that ships. Showing them first
 * costs one dialogue and saves the data.
 */

import { parseEntry, reviewImport, type EntryKind, type EntryOptions, type ImportReview } from '@mirrorz/core';

/** What the user can force a column to be, over the top of the inference. */
export type ColumnOverride = 'auto' | 'text' | 'number' | 'date';

export interface ColumnPlan {
  index: number;
  /** Header text when the first row is a header, otherwise the column letter. */
  name: string;
  override: ColumnOverride;
  /** How many cells of each kind the inference produced. */
  kinds: Record<EntryKind, number>;
  /** The kind the column is mostly made of, ignoring blanks. */
  dominant: EntryKind;
  /** Cells this column would convert to something other than text. */
  converted: number;
  ambiguous: number;
  protectedCount: number;
  /** Distinct reasons cells in this column were kept as text. */
  reasons: string[];
  samples: { input: string; result: string }[];
}

export interface ImportPlan {
  review: ImportReview;
  columns: ColumnPlan[];
  totals: Record<EntryKind, number>;
  /** The sentence at the top of the dialogue. */
  headline: string;
}

const EMPTY_KINDS = (): Record<EntryKind, number> => ({
  text: 0,
  number: 0,
  date: 0,
  time: 0,
  datetime: 0,
  boolean: 0,
  error: 0,
  blank: 0,
});

/** Entry options implied by a column override. */
export function optionsForOverride(base: EntryOptions, override: ColumnOverride): EntryOptions {
  switch (override) {
    case 'text':
      return { ...base, forceText: true };
    case 'number':
      return { ...base, inferNumbers: true, inferDates: false, inferBooleans: false };
    case 'date':
      return { ...base, inferDates: true, inferNumbers: false };
    case 'auto':
      return base;
  }
}

const MAX_SAMPLES = 3;

/**
 * Build the plan for a block of incoming text.
 *
 * Overrides are per column and are applied when the plan is computed, so the
 * dialogue can be recomputed on every change of a dropdown and show the real
 * consequence rather than a promise about it.
 */
export function buildImportPlan(
  rows: readonly (readonly string[])[],
  options: EntryOptions = {},
  overrides: readonly ColumnOverride[] = [],
  headerRow = false,
): ImportPlan {
  const body = headerRow ? rows.slice(1) : rows;
  const width = rows.reduce((max, row) => Math.max(max, row.length), 0);

  const columns: ColumnPlan[] = [];
  const totals = EMPTY_KINDS();

  for (let c = 0; c < width; c++) {
    const override = overrides[c] ?? 'auto';
    const columnOptions = optionsForOverride(options, override);
    const plan: ColumnPlan = {
      index: c,
      name: headerRow ? (rows[0]?.[c] ?? columnLetter(c)) : columnLetter(c),
      override,
      kinds: EMPTY_KINDS(),
      dominant: 'blank',
      converted: 0,
      ambiguous: 0,
      protectedCount: 0,
      reasons: [],
      samples: [],
    };

    for (const row of body) {
      const input = row[c] ?? '';
      if (input === '') continue;
      const result = parseEntry(input, columnOptions);
      plan.kinds[result.kind]++;
      totals[result.kind]++;
      if (result.kind !== 'text' && result.kind !== 'blank') plan.converted++;
      if (result.confidence === 'ambiguous') plan.ambiguous++;
      if (result.confidence === 'risky') {
        plan.protectedCount++;
        if (result.note && !plan.reasons.includes(result.note)) plan.reasons.push(result.note);
      }
      if (plan.samples.length < MAX_SAMPLES) {
        plan.samples.push({ input, result: renderScalar(result.value) });
      }
    }

    plan.dominant = dominantKind(plan.kinds);
    columns.push(plan);
  }

  const review = reviewImport(body, options);
  return { review, columns, totals, headline: headlineFor(totals, review) };
}

function dominantKind(kinds: Record<EntryKind, number>): EntryKind {
  let best: EntryKind = 'blank';
  let bestCount = 0;
  for (const [kind, count] of Object.entries(kinds) as [EntryKind, number][]) {
    if (kind === 'blank') continue;
    if (count > bestCount) {
      best = kind;
      bestCount = count;
    }
  }
  return best;
}

function renderScalar(value: unknown): string {
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

export function columnLetter(index: number): string {
  let n = index;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

/**
 * The one-line summary.
 *
 * It leads with the conversions because those are the ones that change the data,
 * and mentions the protected cells second because those are the ones that
 * reassure - "18 were kept as text" is the sentence that stops someone
 * cancelling an import they should accept.
 */
export function headlineFor(totals: Record<EntryKind, number>, review: ImportReview): string {
  const parts: string[] = [];
  const dates = totals.date + totals.datetime + totals.time;
  if (dates > 0) parts.push(`${dates} cell${dates === 1 ? '' : 's'} would be read as dates`);
  if (totals.number > 0) {
    parts.push(`${totals.number} as number${totals.number === 1 ? '' : 's'}`);
  }
  if (totals.boolean > 0) parts.push(`${totals.boolean} as TRUE/FALSE`);
  if (review.protected.length > 0) {
    parts.push(
      `${review.protected.length} kept as text (${summariseReasons(review)})`,
    );
  }
  if (parts.length === 0) return `${review.total} cells, all kept exactly as supplied`;
  return `${parts.join(', ')}.`;
}

/** The short reason list in the headline: "leading zeros, gene symbols". */
function summariseReasons(review: ImportReview): string {
  const kinds = new Set<string>();
  for (const issue of review.protected) {
    const note = issue.result.note ?? '';
    if (note.includes('leading zero')) kinds.add('leading zeros');
    else if (note.includes('gene symbol')) kinds.add('gene symbols');
    else if (note.includes('15 digits')) kinds.add('long ids');
    else if (note.includes('range or score')) kinds.add('ranges');
    else if (note.includes('version number')) kinds.add('version numbers');
    else kinds.add('other');
  }
  return [...kinds].join(', ');
}

/**
 * Apply the plan, producing the values to write.
 *
 * Returned rather than written, because the caller owns the transaction: a
 * paste and a CSV import land in the document differently but must agree on
 * exactly what the user approved.
 */
export function resolveImport(
  rows: readonly (readonly string[])[],
  options: EntryOptions,
  overrides: readonly ColumnOverride[],
  headerRow = false,
): { value: ReturnType<typeof parseEntry>['value']; literal?: string; row: number; col: number }[] {
  const body = headerRow ? rows.slice(1) : rows;
  const out: { value: ReturnType<typeof parseEntry>['value']; literal?: string; row: number; col: number }[] = [];
  for (let r = 0; r < body.length; r++) {
    const row = body[r]!;
    for (let c = 0; c < row.length; c++) {
      const input = row[c] ?? '';
      if (input === '') continue;
      const result = parseEntry(input, optionsForOverride(options, overrides[c] ?? 'auto'));
      const entry: { value: typeof result.value; literal?: string; row: number; col: number } = {
        value: result.value,
        row: r,
        col: c,
      };
      if (result.literal !== undefined) entry.literal = result.literal;
      out.push(entry);
    }
  }
  return out;
}
