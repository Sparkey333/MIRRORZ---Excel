/**
 * Reading OpenDocument spreadsheets (.ods).
 *
 * ODS is a zip like xlsx but organised very differently, and two of those
 * differences dominate the reader:
 *
 * Repetition instead of addressing. An xlsx cell carries its own address; an ODS
 * cell does not. Position is implied by order, and runs of identical cells
 * collapse into a `number-columns-repeated` count. A row of one value followed
 * by `number-columns-repeated="16384"` is how ODS says "and the rest of the row
 * is empty", so the counts must be honoured for addressing while emphatically
 * not being materialised - a naive reader allocates sixteen thousand blank cells
 * per row and then a million rows.
 *
 * A different formula dialect. Formulas are stored as OpenFormula with a
 * namespace prefix and bracketed dot-prefixed references: `of:=SUM([.A1:.A9])`
 * rather than `SUM(A1:A9)`. Translating is mechanical but must happen, or every
 * formula in the file is unparseable.
 *
 * This is a read-only importer. The useful thing to do with an .ods file here is
 * open it and save it as .xlsx, which the rest of the package already does well;
 * writing ODS would be a second full serializer for no additional user benefit.
 */

import {
  type CellData,
  type Scalar,
  Workbook,
  errorFromCode,
  partsToSerial,
} from '@mirrorz/core';
import { XmlReader, XmlToken } from '../xml.js';
import { looksLikeZip, readZip } from '../zip.js';

export class OdsError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'OdsError';
  }
}

export interface OdsReadResult {
  workbook: Workbook;
  warnings: string[];
}

/**
 * ODS declares its type in an uncompressed `mimetype` entry, which the
 * specification requires to be the first entry in the archive.
 */
export function looksLikeOds(bytes: Uint8Array): boolean {
  if (!looksLikeZip(bytes)) return false;
  try {
    const entry = readZip(bytes).get('mimetype');
    if (!entry) return false;
    return new TextDecoder().decode(entry.data()).includes('opendocument.spreadsheet');
  } catch {
    return false;
  }
}

/** Guard against a file whose repeat counts would materialise the whole grid. */
const MAX_MATERIALISED_COLUMNS = 16_384;
const MAX_MATERIALISED_ROWS = 1_048_576;

export function readOds(bytes: Uint8Array): OdsReadResult {
  if (!looksLikeZip(bytes)) throw new OdsError('not a zip archive, so not an .ods file');
  const entries = readZip(bytes);
  const content = entries.get('content.xml');
  if (!content) throw new OdsError('archive has no content.xml, so it is not an ODS document');

  const warnings: string[] = [];
  const workbook = new Workbook();
  parseContent(new TextDecoder().decode(content.data()), workbook, warnings);
  return { workbook, warnings };
}

function parseContent(xml: string, workbook: Workbook, warnings: string[]): void {
  const r = new XmlReader(xml);
  let sheet: import('@mirrorz/core').Sheet | undefined;
  let row = -1;
  let col = 0;
  /** Rows the current row element stands in for, from its repeat count. */
  let rowRepeat = 1;

  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Close) {
      if (r.localName === 'table-row' && sheet) {
        // A repeated row advances the cursor without duplicating content: an
        // empty run of a million rows must cost nothing.
        row += rowRepeat;
        rowRepeat = 1;
      }
      continue;
    }
    if (t !== XmlToken.Open) continue;

    switch (r.localName) {
      case 'table': {
        const name = r.attr('table:name') ?? r.attr('name') ?? `Sheet${workbook.sheets.length + 1}`;
        sheet = workbook.addSheet(name);
        row = -1;
        break;
      }

      case 'table-row': {
        if (!sheet) break;
        row += 1;
        col = 0;
        const repeated = Number(r.attr('table:number-rows-repeated') ?? '1');
        rowRepeat = Number.isFinite(repeated) && repeated > 0 ? repeated : 1;
        // The trailing "rest of the sheet is empty" row carries an enormous
        // count; cap it so the cursor arithmetic stays sane.
        if (rowRepeat > MAX_MATERIALISED_ROWS) rowRepeat = MAX_MATERIALISED_ROWS;
        // Subtract one because the row itself was already counted above.
        rowRepeat -= 1;
        const height = parseLength(r.attr('style:row-height'));
        if (height !== undefined) sheet.rows.set(row, { height });
        break;
      }

      case 'covered-table-cell':
      case 'table-cell': {
        if (!sheet) break;
        // Every attribute must be read BEFORE the element's children are
        // consumed: readCell walks into the text nodes, and after that the
        // reader is positioned on a different element entirely.
        const repeatAttr = Number(r.attr('table:number-columns-repeated') ?? '1');
        const repeat = Number.isFinite(repeatAttr) && repeatAttr > 0 ? repeatAttr : 1;
        const spanRows = Number(r.attr('table:number-rows-spanned') ?? '1');
        const spanCols = Number(r.attr('table:number-columns-spanned') ?? '1');

        const cell = readCell(r, workbook);
        if (cell) {
          // Only a cell with content is written, and a repeated run of content
          // is written out; a repeated run of nothing just advances the cursor.
          const limit = Math.min(repeat, MAX_MATERIALISED_COLUMNS - col);
          for (let i = 0; i < limit; i++) {
            if (row >= 0 && row < MAX_MATERIALISED_ROWS && col + i < MAX_MATERIALISED_COLUMNS) {
              sheet.setCell(row, col + i, { ...cell });
            }
          }
        }

        if (spanRows > 1 || spanCols > 1) {
          sheet.merges.push({
            range: {
              start: { row, col, rowAbs: false, colAbs: false },
              end: {
                row: row + Math.max(1, spanRows) - 1,
                col: col + Math.max(1, spanCols) - 1,
                rowAbs: false,
                colAbs: false,
              },
            },
          });
        }

        col += repeat;
        break;
      }

      case 'table-column': {
        if (!sheet) break;
        const repeated = Number(r.attr('table:number-columns-repeated') ?? '1');
        const width = parseLength(r.attr('style:column-width'));
        if (width !== undefined) {
          const count = Math.min(
            Number.isFinite(repeated) && repeated > 0 ? repeated : 1,
            1024,
          );
          const first = sheet.cols.size;
          for (let i = 0; i < count; i++) sheet.cols.set(first + i, { width });
        }
        break;
      }

      default:
        break;
    }
  }

  if (workbook.sheets.length === 0) {
    warnings.push('content.xml contained no tables');
  }
}

/** Read one cell element's value and formula. */
function readCell(r: XmlReader, workbook: Workbook): CellData | undefined {
  const valueType = r.attr('office:value-type') ?? r.attr('value-type');
  const formulaAttr = r.attr('table:formula') ?? r.attr('formula');
  const selfClosing = r.isSelfClosing;
  // An error result is written as an empty string value with the real error
  // text only in the display paragraph, flagged by a LibreOffice extension
  // attribute. Trusting office:string-value alone turns every error into "".
  const isErrorCell = (r.attr('calcext:value-type') ?? '') === 'error';

  let value: Scalar = null;

  switch (valueType) {
    case 'float':
    case 'percentage':
    case 'currency': {
      const raw = r.attr('office:value') ?? r.attr('value');
      value = raw === undefined ? null : Number(raw);
      break;
    }
    case 'boolean': {
      const raw = r.attr('office:boolean-value') ?? r.attr('boolean-value');
      value = raw === 'true';
      break;
    }
    case 'date': {
      const raw = r.attr('office:date-value') ?? r.attr('date-value');
      value = raw === undefined ? null : isoToSerial(raw, workbook.dateSystem);
      break;
    }
    case 'time': {
      const raw = r.attr('office:time-value') ?? r.attr('time-value');
      value = raw === undefined ? null : durationToFraction(raw);
      break;
    }
    case 'string': {
      // A string's text lives in the child paragraphs, not in an attribute.
      const raw = r.attr('office:string-value') ?? r.attr('string-value');
      // For an error cell the attribute is deliberately empty, so the display
      // text is the only place the error code appears.
      value = raw !== undefined && !isErrorCell ? raw : selfClosing ? '' : readText(r);
      break;
    }
    default: {
      // No value type: either genuinely empty, or a formula whose result the
      // producer stored only as display text.
      if (!selfClosing) {
        const text = readText(r);
        value = text === '' ? null : text;
      }
      break;
    }
  }

  // An error result is written as a string that looks like an Excel error.
  if (typeof value === 'string' && value.startsWith('#')) {
    const asError = errorFromCode(value);
    if (asError) value = asError;
  }

  if (value === null && formulaAttr === undefined) return undefined;

  const cell: CellData = { value };
  if (formulaAttr !== undefined) {
    const converted = openFormulaToA1(formulaAttr);
    if (converted !== undefined) cell.formula = converted;
  }
  return cell;
}

/** Concatenate the text of an element's paragraph children. */
function readText(r: XmlReader): string {
  const target = r.depth - 1;
  const parts: string[] = [];
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Close && r.depth === target) break;
    if (t === XmlToken.Text || t === XmlToken.CData) parts.push(r.text);
    // Each paragraph is a line; ODS has no other newline representation.
    else if (t === XmlToken.Close && r.localName === 'p' && parts.length > 0) parts.push('\n');
    // A run of spaces is encoded as an element rather than as literal spaces.
    else if (t === XmlToken.Open && r.localName === 's') {
      const count = Number(r.attr('text:c') ?? '1');
      parts.push(' '.repeat(Number.isFinite(count) && count > 0 ? Math.min(count, 4096) : 1));
    } else if (t === XmlToken.Open && r.localName === 'tab') parts.push('\t');
  }
  return parts.join('').replace(/\n$/, '');
}

/**
 * Translate an OpenFormula expression into the A1 dialect the rest of the
 * engine speaks.
 *
 * `of:=SUM([.A1:.A9])` becomes `SUM(A1:A9)`, and `[Sheet2.A1]` becomes
 * `Sheet2!A1`. The transformation is textual and deliberately conservative: a
 * construct it does not recognise is passed through rather than mangled, since
 * a formula that fails to evaluate is recoverable and one silently rewritten
 * into a different meaning is not.
 */
export function openFormulaToA1(formula: string): string | undefined {
  // Strip the namespace prefix and the leading '='.
  let text = formula.replace(/^[A-Za-z0-9_]+:/, '');
  if (text.startsWith('=')) text = text.slice(1);
  if (text === '') return undefined;

  // Bracketed references: [.A1], [.A1:.B2], [Sheet2.A1], [$Sheet2.$A$1].
  text = text.replace(/\[([^\]]*)\]/g, (_match, inner: string) => {
    const parts = inner.split(':');
    const converted = parts.map(convertReference);
    return converted.join(':');
  });

  // ODS separates arguments with semicolons.
  text = text.replace(/;/g, ',');

  // A few function names differ.
  text = text.replace(/\bCOM\.MICROSOFT\./gi, '');
  text = text.replace(/\bLEGACY\./gi, '');

  return text;
}

function convertReference(ref: string): string {
  let text = ref.trim();
  if (text === '') return text;
  // A leading $ on the sheet name marks an absolute sheet, which A1 has no
  // notation for; dropping it changes nothing about which cell is meant.
  if (text.startsWith('$')) text = text.slice(1);

  const dot = text.indexOf('.');
  if (dot < 0) return text;

  const sheet = text.slice(0, dot);
  const cell = text.slice(dot + 1);
  if (sheet === '') return cell;

  // A sheet name needing quotes already carries them in ODS.
  const needsQuotes = !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) && !sheet.startsWith("'");
  return `${needsQuotes ? `'${sheet.replace(/'/g, "''")}'` : sheet}!${cell}`;
}

/** ODS lengths carry a unit: `2.5cm`, `0.89in`, `48pt`. Our model uses points. */
function parseLength(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(value.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  switch (m[2]) {
    case 'cm':
      return (n / 2.54) * 72;
    case 'mm':
      return (n / 25.4) * 72;
    case 'in':
      return n * 72;
    case 'pc':
      return n * 12;
    case 'px':
      return n * 0.75;
    default:
      return n;
  }
}

/** `2024-02-29` or `2024-02-29T13:45:30` to a spreadsheet serial. */
function isoToSerial(iso: string, system: 1900 | 1904): number | null {
  const m = /^(-?\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(iso);
  if (!m) return null;
  return partsToSerial(
    Number(m[1]),
    Number(m[2]),
    Number(m[3]),
    Number(m[4] ?? 0),
    Number(m[5] ?? 0),
    Math.floor(Number(m[6] ?? 0)),
    system,
  );
}

/** An ISO 8601 duration such as `PT13H45M30S`, as a fraction of a day. */
function durationToFraction(duration: string): number | null {
  const m = /^-?PT(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?$/.exec(duration);
  if (!m) return null;
  const hours = Number(m[1] ?? 0);
  const minutes = Number(m[2] ?? 0);
  const seconds = Number(m[3] ?? 0);
  const fraction = (hours * 3600 + minutes * 60 + seconds) / 86_400;
  return duration.startsWith('-') ? -fraction : fraction;
}
