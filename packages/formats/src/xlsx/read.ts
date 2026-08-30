/**
 * Reading .xlsx / .xlsm / .xltx / .xltm / .xlam into the core workbook model.
 *
 * The reader keeps the OpcPackage alongside the model. Everything it does not
 * understand stays in the package untouched, so a later save can put it back;
 * the model is a projection, not a replacement. That is what makes it safe to
 * open a workbook full of pivot tables and charts we cannot yet edit.
 */

import {
  type CellData,
  type ColProps,
  type DefinedName,
  type RowProps,
  type Scalar,
  type Sheet,
  Workbook,
  errorFromCode,
  nameToCol,
  parseRangeRef,
} from '@mirrorz/core';
import { OpcPackage, RelType, type WorkbookFlavour } from '../opc.js';
import { XmlReader, XmlToken, unescapeSharedString } from '../xml.js';
import { type StyleTables, emptyStyleTables, parseStyles, resolveXf } from './styles.js';

export interface ReadResult {
  workbook: Workbook;
  /** The package the workbook came from, retained so a save can preserve it. */
  pkg: OpcPackage;
  styleTables: StyleTables;
  flavour: WorkbookFlavour;
  /** Non-fatal problems: we open what we can and tell the user what we skipped. */
  warnings: string[];
}

export interface ReadOptions {
  /**
   * Skip cell values and read only structure. Used by the "open instantly, fill
   * in behind" path so a 200 MB workbook shows its sheet tabs immediately.
   */
  structureOnly?: boolean;
  /** Limit which sheets are populated, by name. */
  sheets?: string[];
}

export class XlsxError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'XlsxError';
  }
}

export function readXlsx(bytes: Uint8Array, options: ReadOptions = {}): ReadResult {
  const pkg = OpcPackage.read(bytes);
  return readFromPackage(pkg, options);
}

export function readFromPackage(pkg: OpcPackage, options: ReadOptions = {}): ReadResult {
  const warnings: string[] = [];
  const workbookPath = pkg.mainDocumentPath();
  if (!pkg.has(workbookPath)) {
    throw new XlsxError(`workbook part is missing: ${workbookPath}`);
  }

  const workbook = new Workbook();
  const flavour = pkg.flavour();

  const styleTables = readStyleTables(pkg, workbookPath, warnings);
  const sharedStrings = readSharedStrings(pkg, workbookPath, warnings);

  const wbInfo = parseWorkbookPart(pkg.text(workbookPath));
  workbook.dateSystem = wbInfo.date1904 ? 1904 : 1900;
  if (wbInfo.calcMode) workbook.calcMode = wbInfo.calcMode;
  workbook.fullCalcOnLoad = wbInfo.fullCalcOnLoad;
  workbook.definedNames = wbInfo.definedNames;

  // The VBA project is carried through as opaque bytes. Round-tripping it
  // byte-identically is what keeps macros working in a file we saved; we do not
  // regenerate it, and reading it is not executing it.
  if (pkg.has('xl/vbaProject.bin')) {
    workbook.vbaProject = pkg.bytes('xl/vbaProject.bin');
  }

  const wanted = options.sheets ? new Set(options.sheets) : undefined;

  for (const entry of wbInfo.sheets) {
    const sheet = workbook.addSheet(entry.name);
    sheet.visibility = entry.state;
    if (entry.rId) {
      const path = pkg.resolve(workbookPath, entry.rId);
      if (!path || !pkg.has(path)) {
        warnings.push(`sheet "${entry.name}" points at a missing part (${entry.rId})`);
        continue;
      }
      if (options.structureOnly || (wanted && !wanted.has(entry.name))) continue;
      try {
        readWorksheet(pkg.text(path), sheet, sharedStrings, styleTables, workbook);
      } catch (err) {
        // One damaged sheet should not stop the other sheets opening. Users
        // overwhelmingly prefer a partly-recovered workbook to a refusal.
        warnings.push(
          `sheet "${entry.name}" could not be fully read: ${err instanceof Error ? err.message : String(err)}`,
        );
      }
    }
  }

  return { workbook, pkg, styleTables, flavour, warnings };
}

function readStyleTables(pkg: OpcPackage, workbookPath: string, warnings: string[]): StyleTables {
  const rel = pkg.related(workbookPath, RelType.styles)[0];
  if (!rel || !pkg.has(rel.path)) return emptyStyleTables();
  try {
    return parseStyles(pkg.text(rel.path));
  } catch (err) {
    warnings.push(`styles could not be read: ${err instanceof Error ? err.message : String(err)}`);
    return emptyStyleTables();
  }
}

function readSharedStrings(pkg: OpcPackage, workbookPath: string, warnings: string[]): string[] {
  const rel = pkg.related(workbookPath, RelType.sharedStrings)[0];
  if (!rel || !pkg.has(rel.path)) return [];
  try {
    return parseSharedStrings(pkg.text(rel.path));
  } catch (err) {
    warnings.push(
      `shared strings could not be read: ${err instanceof Error ? err.message : String(err)}`,
    );
    return [];
  }
}

/**
 * The shared string table.
 *
 * Each `<si>` is one string, but it may be split across several `<r>` runs with
 * their own formatting. We concatenate the text and drop the run formatting for
 * now; rich text within a cell is a v2 feature, and the raw part is preserved in
 * the package so nothing is lost on save.
 */
export function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  const r = new XmlReader(xml);
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Open && r.localName === 'si') {
      out.push(unescapeSharedString(r.readText()));
    }
  }
  return out;
}

interface SheetEntry {
  name: string;
  sheetId: number;
  rId?: string;
  state: 'visible' | 'hidden' | 'veryHidden';
}

interface WorkbookInfo {
  sheets: SheetEntry[];
  definedNames: DefinedName[];
  date1904: boolean;
  calcMode?: 'auto' | 'autoNoTable' | 'manual';
  fullCalcOnLoad: boolean;
}

export function parseWorkbookPart(xml: string): WorkbookInfo {
  const info: WorkbookInfo = {
    sheets: [],
    definedNames: [],
    date1904: false,
    fullCalcOnLoad: false,
  };
  const r = new XmlReader(xml);
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t !== XmlToken.Open) continue;
    const name = r.localName;
    if (name === 'sheet') {
      const state = r.attr('state');
      info.sheets.push({
        name: r.attr('name') ?? `Sheet${info.sheets.length + 1}`,
        sheetId: Number(r.attr('sheetId') ?? info.sheets.length + 1),
        rId: r.attr('r:id') ?? r.attr('id'),
        state: state === 'hidden' ? 'hidden' : state === 'veryHidden' ? 'veryHidden' : 'visible',
      });
    } else if (name === 'workbookPr') {
      const d = r.attr('date1904');
      info.date1904 = d === '1' || d === 'true';
    } else if (name === 'calcPr') {
      const mode = r.attr('calcMode');
      if (mode === 'manual' || mode === 'autoNoTable' || mode === 'auto') info.calcMode = mode;
      const full = r.attr('fullCalcOnLoad');
      info.fullCalcOnLoad = full === '1' || full === 'true';
    } else if (name === 'definedName') {
      const dn: DefinedName = {
        name: r.attr('name') ?? '',
        refersTo: '',
      };
      const localSheetId = r.attr('localSheetId');
      if (localSheetId !== undefined) dn.scope = Number(localSheetId);
      const hidden = r.attr('hidden');
      if (hidden === '1' || hidden === 'true') dn.hidden = true;
      const comment = r.attr('comment');
      if (comment) dn.comment = comment;
      dn.refersTo = r.readText();
      if (dn.name) info.definedNames.push(dn);
    }
  }
  return info;
}

/**
 * Parse one worksheet part into a Sheet.
 *
 * This is the hot path - a real workbook can hold millions of `<c>` elements -
 * so it stays a single forward pass with no intermediate objects beyond the
 * CellData we actually keep.
 */
export function readWorksheet(
  xml: string,
  sheet: Sheet,
  sharedStrings: readonly string[],
  styleTables: StyleTables,
  workbook: Workbook,
): void {
  const r = new XmlReader(xml);
  const styleCache = new Map<number, number>();

  /** Map a file style index to an interned StyleId, memoised per sheet. */
  const styleFor = (raw: string | undefined): number | undefined => {
    if (raw === undefined) return undefined;
    const idx = Number(raw);
    if (!Number.isFinite(idx)) return undefined;
    const cached = styleCache.get(idx);
    if (cached !== undefined) return cached;
    const id = workbook.styles.intern(resolveXf(styleTables, idx));
    styleCache.set(idx, id);
    return id;
  };

  let currentRow = -1;

  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t !== XmlToken.Open) continue;
    const name = r.localName;

    switch (name) {
      case 'row': {
        const rAttr = r.attr('r');
        currentRow = rAttr === undefined ? currentRow + 1 : Number(rAttr) - 1;
        const props = readRowProps(r, styleFor);
        if (props) sheet.rows.set(currentRow, props);
        break;
      }

      case 'c': {
        const cell = readCell(r, currentRow, sharedStrings, styleFor, styleTables);
        if (cell) sheet.setCell(cell.row, cell.col, cell.data);
        break;
      }

      case 'col': {
        // A single <col> element can span a range of columns.
        const min = Number(r.attr('min') ?? 1) - 1;
        const max = Number(r.attr('max') ?? min + 1) - 1;
        const props: ColProps = {};
        const width = r.attr('width');
        if (width !== undefined) props.width = Number(width);
        if (r.attr('hidden') === '1' || r.attr('hidden') === 'true') props.hidden = true;
        if (r.attr('customWidth') === '1') props.customWidth = true;
        const level = r.attr('outlineLevel');
        if (level !== undefined) props.level = Number(level);
        if (r.attr('collapsed') === '1') props.collapsed = true;
        const style = styleFor(r.attr('style'));
        if (style !== undefined) props.style = style;
        if (Object.keys(props).length > 0) {
          // Excel writes max="16384" for a format applied to the whole sheet.
          // Materialising 16384 entries for that would be absurd, so cap the
          // span at what the sheet actually uses plus a small margin.
          const limit = Math.min(max, min + 1024);
          for (let c = min; c <= limit; c++) sheet.cols.set(c, { ...props });
        }
        break;
      }

      case 'mergeCell': {
        const ref = r.attr('ref');
        if (ref) {
          const range = parseRangeRef(ref);
          if (range) sheet.merges.push({ range });
        }
        break;
      }

      case 'sheetFormatPr': {
        const dh = r.attr('defaultRowHeight');
        if (dh !== undefined) sheet.defaultRowHeight = Number(dh);
        const dw = r.attr('defaultColWidth');
        if (dw !== undefined) sheet.defaultColWidth = Number(dw);
        break;
      }

      case 'sheetView': {
        if (r.attr('showGridLines') !== undefined) {
          sheet.view.showGridLines = r.attr('showGridLines') !== '0';
        }
        if (r.attr('showRowColHeaders') !== undefined) {
          sheet.view.showRowColHeaders = r.attr('showRowColHeaders') !== '0';
        }
        const zoom = r.attr('zoomScale');
        if (zoom !== undefined) sheet.view.zoomScale = Number(zoom);
        if (r.attr('rightToLeft') === '1') sheet.view.rightToLeft = true;
        break;
      }

      case 'pane': {
        // Frozen panes are expressed as the split position in cell units.
        const state = r.attr('state');
        if (state === 'frozen' || state === 'frozenSplit') {
          const x = r.attr('xSplit');
          const y = r.attr('ySplit');
          if (x !== undefined) sheet.view.frozenCols = Number(x);
          if (y !== undefined) sheet.view.frozenRows = Number(y);
        }
        break;
      }

      case 'selection': {
        const active = r.attr('activeCell');
        if (active) {
          const ref = parseRangeRef(active);
          if (ref) sheet.view.activeCell = { row: ref.start.row, col: ref.start.col };
        }
        break;
      }

      case 'sheetPr': {
        // codeName is how VBA addresses the sheet. Dropping it leaves an .xlsm
        // that opens fine and whose macros silently no longer bind.
        const codeName = r.attr('codeName');
        if (codeName) sheet.preserved['codeName'] = codeName;
        break;
      }

      case 'tabColor': {
        const rgb = r.attr('rgb');
        if (rgb) sheet.tabColor = rgb;
        break;
      }

      // Subtrees we do not model yet are preserved verbatim at the sheet level
      // so a save puts them back rather than silently dropping the feature.
      case 'conditionalFormatting':
      case 'dataValidations':
      case 'autoFilter':
      case 'hyperlinks':
      case 'drawing':
      case 'legacyDrawing':
      case 'tableParts':
      case 'extLst':
      case 'pageMargins':
      case 'pageSetup':
      case 'printOptions':
      case 'headerFooter':
      case 'rowBreaks':
      case 'colBreaks':
      case 'sheetProtection': {
        const existing = sheet.preserved[name];
        const raw = r.readRaw();
        sheet.preserved[name] = existing === undefined ? raw : existing + raw;
        break;
      }

      default:
        break;
    }
  }
}

function readRowProps(
  r: XmlReader,
  styleFor: (raw: string | undefined) => number | undefined,
): RowProps | undefined {
  const props: RowProps = {};
  const ht = r.attr('ht');
  if (ht !== undefined) props.height = Number(ht);
  if (r.attr('customHeight') === '1') props.customHeight = true;
  if (r.attr('hidden') === '1' || r.attr('hidden') === 'true') props.hidden = true;
  const level = r.attr('outlineLevel');
  if (level !== undefined) props.level = Number(level);
  if (r.attr('collapsed') === '1') props.collapsed = true;
  // A row style only applies when customFormat says so.
  if (r.attr('customFormat') === '1') {
    const style = styleFor(r.attr('s'));
    if (style !== undefined) props.style = style;
  }
  return Object.keys(props).length > 0 ? props : undefined;
}

/**
 * Parse a `<c>` element.
 *
 * The `t` attribute selects the interpretation of `<v>`:
 *   absent or "n" - number
 *   "s"           - index into the shared string table
 *   "str"         - a formula that returned text; `<v>` holds the text itself
 *   "inlineStr"   - text in a nested `<is>`, with no `<v>` at all
 *   "b"           - boolean, "1" or "0"
 *   "e"           - error, with the error string in `<v>`
 *   "d"           - ISO 8601 date; defined in later ISO editions, and Excel does
 *                   not write it, but LibreOffice and other producers may.
 */
function readCell(
  r: XmlReader,
  fallbackRow: number,
  sharedStrings: readonly string[],
  styleFor: (raw: string | undefined) => number | undefined,
  styleTables: StyleTables,
): { row: number; col: number; data: CellData } | undefined {
  const ref = r.attr('r');
  let row = fallbackRow;
  let col = -1;
  if (ref) {
    const parsed = parseCellAddress(ref);
    if (parsed) {
      row = parsed.row;
      col = parsed.col;
    }
  }
  if (col < 0 || row < 0) return undefined;

  const type = r.attr('t') ?? 'n';
  const style = styleFor(r.attr('s'));
  const isSelfClosing = r.isSelfClosing;

  let value: Scalar = null;
  let formula: string | undefined;
  let rawV: string | undefined;
  let inlineText: string | undefined;
  let spill: { rows: number; cols: number } | undefined;

  if (!isSelfClosing) {
    const target = r.depth - 1;
    for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
      if (t === XmlToken.Close && r.depth === target) break;
      if (t !== XmlToken.Open) continue;
      const child = r.localName;
      if (child === 'v') {
        rawV = r.readText();
      } else if (child === 'f') {
        const fType = r.attr('t');
        const refAttr = r.attr('ref');
        // A dynamic array formula declares the rectangle it spills over.
        if (fType === 'array' && refAttr) {
          const range = parseRangeRef(refAttr);
          if (range) {
            spill = {
              rows: range.end.row - range.start.row + 1,
              cols: range.end.col - range.start.col + 1,
            };
          }
        }
        const text = r.readText();
        // Shared formulas store the expression only on the master cell; the
        // followers carry just a `si` index. Expanding those needs the master,
        // which we may not have seen yet, so an empty follower keeps its cached
        // value and is recomputed on first calculation.
        if (text) formula = stripFutureFunctionPrefixes(text);
      } else if (child === 'is') {
        inlineText = unescapeSharedString(r.readText());
      }
    }
  } else {
    // A self-closing <c/> still consumed only its own tag; the reader owes a
    // Close token, which the caller's loop will see and ignore.
  }

  switch (type) {
    case 's': {
      const idx = rawV === undefined ? Number.NaN : Number(rawV);
      value = Number.isFinite(idx) ? (sharedStrings[idx] ?? '') : '';
      break;
    }
    case 'inlineStr':
      value = inlineText ?? '';
      break;
    case 'str':
      value = rawV === undefined ? '' : unescapeSharedString(rawV);
      break;
    case 'b':
      value = rawV === '1' || rawV === 'true';
      break;
    case 'e':
      value = (rawV !== undefined ? errorFromCode(rawV) : undefined) ?? null;
      break;
    case 'd':
      value = rawV ?? null;
      break;
    default: {
      if (rawV === undefined || rawV === '') {
        value = null;
      } else {
        const n = Number(rawV);
        value = Number.isFinite(n) ? n : rawV;
      }
      break;
    }
  }

  if (value === null && formula === undefined && style === undefined) return undefined;

  const data: CellData = { value };
  if (formula !== undefined) data.formula = formula;
  if (style !== undefined) data.style = style;
  if (spill) data.spill = spill;
  void styleTables;
  return { row, col, data };
}

/**
 * Strip the `_xlfn.` and `_xlfn._xlws.` prefixes OOXML puts on functions newer
 * than Excel 2007. They are a storage detail; showing them to the user, or
 * feeding them to the evaluator, produces #NAME? for perfectly valid formulas.
 * The writer puts them back.
 */
export function stripFutureFunctionPrefixes(formula: string): string {
  if (formula.indexOf('_xlfn.') < 0) return formula;
  return formula.replaceAll('_xlfn._xlws.', '').replaceAll('_xlfn.', '');
}

/** Parse the `r` attribute of a cell: an A1 address, possibly with $ markers. */
function parseCellAddress(ref: string): { row: number; col: number } | undefined {
  let i = 0;
  if (ref.charCodeAt(i) === 36) i++; // $
  const colStart = i;
  while (i < ref.length) {
    const c = ref.charCodeAt(i) & ~0x20;
    if (c < 65 || c > 90) break;
    i++;
  }
  if (i === colStart) return undefined;
  const col = nameToCol(ref.slice(colStart, i));
  if (col < 0) return undefined;
  if (ref.charCodeAt(i) === 36) i++;
  const row = Number(ref.slice(i));
  if (!Number.isFinite(row) || row < 1) return undefined;
  return { row: row - 1, col };
}
