/**
 * Writing .xlsx / .xlsm / .xltx / .xltm.
 *
 * Two modes, and the distinction matters more than any other decision in this
 * file:
 *
 *   Round-trip save. The workbook came from a package, so we rewrite only the
 *   parts we actually modelled and leave every other byte exactly as it arrived.
 *   Charts, pivot tables, drawings, slicers and everything else we cannot yet
 *   edit survive untouched. This is the difference between a tool people trust
 *   with real files and one whose documentation has to warn that charts are lost
 *   on save.
 *
 *   Fresh save. There is no source package, so we build the minimum part set
 *   Excel accepts without a repair prompt.
 *
 * The recurring failure mode this file guards against is Excel's "we found a
 * problem with some content" dialog, whose two most common causes are an element
 * in the wrong sequence position (see order.ts) and a count attribute that does
 * not match the number of children it counts. Excel cross-checks those counts
 * and silently drops records rather than reporting a clean error.
 */

import {
  type CellData,
  type Sheet,
  type Workbook,
  colToName,
  formatRangeRef,
  isError,
  quoteSheetName,
} from '@mirrorz/core';
import { ContentType, OpcPackage, RelType, type WorkbookFlavour } from '../opc.js';
import { XmlWriter, escapeSharedString } from '../xml.js';
import { WORKSHEET_CHILD_ORDER, emitInOrder } from './order.js';
import { type StyleTables, defaultStyleTables, writeStyles } from './styles.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface WriteOptions {
  /** The package the workbook was read from, for a preserving round-trip save. */
  source?: OpcPackage;
  styleTables?: StyleTables;
  flavour?: WorkbookFlavour;
  /**
   * Emit strings inline rather than through a shared string table. Costs bytes
   * on repetitive data but avoids a second pass, which matters for an export of
   * a very large sheet.
   */
  inlineStrings?: boolean;
  /** Fixed timestamp, making output reproducible for tests. */
  modified?: Date;
}

export function writeXlsx(workbook: Workbook, options: WriteOptions = {}): Uint8Array {
  const pkg = options.source ?? freshPackage(options.flavour ?? 'xlsx');
  const styleTables = options.styleTables ?? defaultStyleTables();
  const workbookPath = pkg.mainDocumentPath();

  // The calculation chain is a cache of the order Excel last evaluated formulas
  // in. Once any formula has changed it is stale, and a stale chain makes Excel
  // show a repair prompt naming the part. Deleting it is the supported fix, and
  // the one deliberate exception to preserving what we do not understand.
  pkg.dropCalcChain();

  const sst = options.inlineStrings ? undefined : buildSharedStrings(workbook);

  writeWorkbookPart(pkg, workbook, workbookPath);
  writeWorksheets(pkg, workbook, workbookPath, sst, options.inlineStrings ?? false);

  pkg.putText(stylesPath(pkg, workbookPath), writeStyles(styleTables), ContentType.styles);

  if (sst) {
    pkg.putText(sharedStringsPath(pkg, workbookPath), renderSharedStrings(sst), ContentType.sharedStrings);
  }

  return pkg.write(options.modified === undefined ? {} : { modified: options.modified });
}

/** The minimum package Excel opens without complaint. */
function freshPackage(flavour: WorkbookFlavour): OpcPackage {
  const pkg = new OpcPackage();
  const workbookType =
    flavour === 'xlsm'
      ? ContentType.xlsm
      : flavour === 'xltx'
        ? ContentType.xltx
        : flavour === 'xltm'
          ? ContentType.xltm
          : flavour === 'xlam'
            ? ContentType.xlam
            : ContentType.xlsx;

  const rels = new XmlWriter();
  rels.open('Relationships', {
    xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
  });
  rels.empty('Relationship', {
    Id: 'rId1',
    Type: RelType.officeDocument,
    Target: 'xl/workbook.xml',
  });
  rels.close();
  pkg.putText('_rels/.rels', rels.toString());

  pkg.putText('xl/workbook.xml', '<workbook/>', workbookType);
  return pkg;
}

function stylesPath(pkg: OpcPackage, workbookPath: string): string {
  return pkg.related(workbookPath, RelType.styles)[0]?.path ?? 'xl/styles.xml';
}

function sharedStringsPath(pkg: OpcPackage, workbookPath: string): string {
  const existing = pkg.related(workbookPath, RelType.sharedStrings)[0]?.path;
  if (existing) return existing;
  pkg.addRelationship(workbookPath, RelType.sharedStrings, 'sharedStrings.xml');
  return 'xl/sharedStrings.xml';
}

/** Shared string table: unique strings, in first-use order. */
interface SharedStrings {
  index: Map<string, number>;
  order: string[];
  /** Total references, which is what the `count` attribute reports. */
  total: number;
}

function buildSharedStrings(workbook: Workbook): SharedStrings {
  const sst: SharedStrings = { index: new Map(), order: [], total: 0 };
  for (const sheet of workbook.sheets) {
    for (const { cell } of sheet.entries()) {
      if (typeof cell.value === 'string') {
        sst.total++;
        if (!sst.index.has(cell.value)) {
          sst.index.set(cell.value, sst.order.length);
          sst.order.push(cell.value);
        }
      }
    }
  }
  return sst;
}

function renderSharedStrings(sst: SharedStrings): string {
  const w = new XmlWriter();
  // Both counts are load-bearing: Excel cross-checks them and recovers lossily
  // rather than reporting an error when they disagree.
  w.open('sst', { xmlns: NS_MAIN, count: sst.total, uniqueCount: sst.order.length });
  for (const s of sst.order) {
    w.open('si');
    // xml:space="preserve" is required or Excel trims leading and trailing
    // spaces out of the user's data.
    w.leaf('t', escapeSharedString(s), needsSpacePreserve(s) ? { 'xml:space': 'preserve' } : undefined);
    w.close();
  }
  w.close();
  return w.toString();
}

function needsSpacePreserve(s: string): boolean {
  return s !== s.trim() || s.includes('\n');
}

function writeWorkbookPart(pkg: OpcPackage, workbook: Workbook, workbookPath: string): void {
  // Sheet relationship ids are reused where the sheet already had one. Minting
  // fresh ids for existing sheets would break every other part that references
  // them by id.
  const existing = pkg.related(workbookPath, RelType.worksheet);
  const w = new XmlWriter();
  w.open('workbook', { xmlns: NS_MAIN, 'xmlns:r': NS_REL });

  if (workbook.dateSystem === 1904) w.empty('workbookPr', { date1904: 1 });

  w.open('sheets');
  workbook.sheets.forEach((sheet, i) => {
    const rel = existing[i];
    const rId = rel?.rel.id ?? pkg.addRelationship(workbookPath, RelType.worksheet, `worksheets/sheet${i + 1}.xml`);
    w.empty('sheet', {
      name: sheet.name,
      sheetId: sheet.id,
      state: sheet.visibility === 'visible' ? undefined : sheet.visibility,
      'r:id': rId,
    });
  });
  w.close();

  if (workbook.definedNames.length > 0) {
    w.open('definedNames');
    for (const dn of workbook.definedNames) {
      w.leaf('definedName', dn.refersTo, {
        name: dn.name,
        localSheetId: dn.scope,
        hidden: dn.hidden ? 1 : undefined,
        comment: dn.comment,
      });
    }
    w.close();
  }

  // Every formula must be recalculated on open: we have just deleted the
  // calculation chain, and our cached values may be from a different engine.
  w.empty('calcPr', {
    calcId: 191029,
    calcMode: workbook.calcMode === 'auto' ? undefined : workbook.calcMode,
    fullCalcOnLoad: 1,
  });

  w.close();
  pkg.putText(workbookPath, w.toString(), pkg.contentType(workbookPath) ?? ContentType.xlsx);
}

function writeWorksheets(
  pkg: OpcPackage,
  workbook: Workbook,
  workbookPath: string,
  sst: SharedStrings | undefined,
  inlineStrings: boolean,
): void {
  const existing = pkg.related(workbookPath, RelType.worksheet);
  workbook.sheets.forEach((sheet, i) => {
    const path = existing[i]?.path ?? `xl/worksheets/sheet${i + 1}.xml`;
    pkg.putText(path, renderWorksheet(sheet, sst, inlineStrings), ContentType.worksheet);
  });
}

function renderWorksheet(
  sheet: Sheet,
  sst: SharedStrings | undefined,
  inlineStrings: boolean,
): string {
  // Fragments are collected by element name and then emitted in schema order,
  // so no amount of restructuring here can produce an out-of-sequence document.
  const fragments = new Map<string, string>();

  if (sheet.preserved['codeName']) {
    const w = new XmlWriter(false);
    w.empty('sheetPr', { codeName: sheet.preserved['codeName'] });
    fragments.set('sheetPr', w.toString());
  }

  const bounds = sheet.bounds();
  {
    const w = new XmlWriter(false);
    w.empty('dimension', {
      ref: bounds
        ? formatRangeRef({
            start: { row: bounds.minRow, col: bounds.minCol, rowAbs: false, colAbs: false },
            end: { row: bounds.maxRow, col: bounds.maxCol, rowAbs: false, colAbs: false },
          })
        : 'A1',
    });
    fragments.set('dimension', w.toString());
  }

  {
    const w = new XmlWriter(false);
    w.open('sheetViews');
    w.open('sheetView', {
      workbookViewId: 0,
      showGridLines: sheet.view.showGridLines === false ? 0 : undefined,
      showRowColHeaders: sheet.view.showRowColHeaders === false ? 0 : undefined,
      zoomScale: sheet.view.zoomScale !== 100 ? sheet.view.zoomScale : undefined,
      rightToLeft: sheet.view.rightToLeft ? 1 : undefined,
    });
    const fr = sheet.view.frozenRows ?? 0;
    const fc = sheet.view.frozenCols ?? 0;
    if (fr > 0 || fc > 0) {
      w.empty('pane', {
        xSplit: fc > 0 ? fc : undefined,
        ySplit: fr > 0 ? fr : undefined,
        topLeftCell: `${colToName(fc)}${fr + 1}`,
        activePane: fc > 0 && fr > 0 ? 'bottomRight' : fc > 0 ? 'topRight' : 'bottomLeft',
        state: 'frozen',
      });
    }
    w.close();
    w.close();
    fragments.set('sheetViews', w.toString());
  }

  {
    const w = new XmlWriter(false);
    w.empty('sheetFormatPr', {
      defaultRowHeight: sheet.defaultRowHeight,
      defaultColWidth: sheet.defaultColWidth !== 8.43 ? sheet.defaultColWidth : undefined,
    });
    fragments.set('sheetFormatPr', w.toString());
  }

  if (sheet.cols.size > 0) {
    const w = new XmlWriter(false);
    w.open('cols');
    // Consecutive columns with identical properties collapse into one element,
    // which is what Excel itself writes and keeps the part small when a format
    // spans many columns.
    for (const run of runsOf(sheet.cols)) {
      w.empty('col', {
        min: run.from + 1,
        max: run.to + 1,
        width: run.props.width ?? sheet.defaultColWidth,
        customWidth: run.props.width !== undefined ? 1 : undefined,
        hidden: run.props.hidden ? 1 : undefined,
        outlineLevel: run.props.level,
        collapsed: run.props.collapsed ? 1 : undefined,
        style: run.props.style,
      });
    }
    w.close();
    fragments.set('cols', w.toString());
  }

  fragments.set('sheetData', renderSheetData(sheet, sst, inlineStrings));

  if (sheet.merges.length > 0) {
    const w = new XmlWriter(false);
    // The count must equal the number of children exactly.
    w.open('mergeCells', { count: sheet.merges.length });
    for (const m of sheet.merges) w.empty('mergeCell', { ref: formatRangeRef(m.range) });
    w.close();
    fragments.set('mergeCells', w.toString());
  }

  if (sheet.tabColor) {
    // tabColor lives inside sheetPr, so merge it with any codeName we kept.
    const w = new XmlWriter(false);
    w.open('sheetPr', { codeName: sheet.preserved['codeName'] });
    w.empty('tabColor', { rgb: sheet.tabColor });
    w.close();
    fragments.set('sheetPr', w.toString());
  }

  // Subtrees we captured verbatim on read go back in at their own positions.
  for (const [name, raw] of Object.entries(sheet.preserved)) {
    if (name === 'codeName') continue;
    fragments.set(name, raw);
  }

  const w = new XmlWriter();
  w.open('worksheet', { xmlns: NS_MAIN, 'xmlns:r': NS_REL });
  w.raw(emitInOrder(WORKSHEET_CHILD_ORDER, fragments));
  w.close();
  return w.toString();
}

interface ColRun<T> {
  from: number;
  to: number;
  props: T;
}

/** Collapse consecutive identical column properties into runs. */
function runsOf<T>(map: Map<number, T>): ColRun<T>[] {
  const keys = [...map.keys()].sort((a, b) => a - b);
  const runs: ColRun<T>[] = [];
  for (const key of keys) {
    const props = map.get(key)!;
    const last = runs[runs.length - 1];
    if (last && last.to === key - 1 && sameProps(last.props, props)) {
      last.to = key;
    } else {
      runs.push({ from: key, to: key, props });
    }
  }
  return runs;
}

function sameProps(a: unknown, b: unknown): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}

function renderSheetData(
  sheet: Sheet,
  sst: SharedStrings | undefined,
  inlineStrings: boolean,
): string {
  const w = new XmlWriter(false);
  w.open('sheetData');

  // Group cells by row in a single pass over the sorted keys, so a million-cell
  // sheet never builds an intermediate structure larger than one row.
  let currentRow = -1;
  let rowOpen = false;
  for (const { row, col, cell } of sheet.entries()) {
    if (row !== currentRow) {
      if (rowOpen) w.close();
      const props = sheet.rows.get(row);
      w.open('row', {
        r: row + 1,
        ht: props?.height,
        customHeight: props?.customHeight ? 1 : undefined,
        hidden: props?.hidden ? 1 : undefined,
        outlineLevel: props?.level,
        collapsed: props?.collapsed ? 1 : undefined,
        s: props?.style,
        customFormat: props?.style !== undefined ? 1 : undefined,
      });
      rowOpen = true;
      currentRow = row;
    }
    writeCell(w, row, col, cell, sst, inlineStrings);
  }
  if (rowOpen) w.close();

  // Rows carrying only formatting still need an element, or the formatting is
  // lost even though no cell in them holds a value.
  w.close();
  return w.toString();
}

function writeCell(
  w: XmlWriter,
  row: number,
  col: number,
  cell: CellData,
  sst: SharedStrings | undefined,
  inlineStrings: boolean,
): void {
  const ref = `${colToName(col)}${row + 1}`;
  const value = cell.value;
  const style = cell.style;

  if (value === null && cell.formula === undefined) {
    if (style !== undefined) w.empty('c', { r: ref, s: style });
    return;
  }

  // The `t` attribute is omitted for numbers, which is the default and by far
  // the most common case.
  if (cell.formula !== undefined) {
    const t = formulaValueType(value);
    w.open('c', { r: ref, s: style, t });
    w.leaf('f', restoreFutureFunctionPrefixes(cell.formula));
    if (value !== null) w.leaf('v', formulaValueText(value));
    w.close();
    return;
  }

  if (typeof value === 'number') {
    w.open('c', { r: ref, s: style });
    w.leaf('v', String(value));
    w.close();
    return;
  }

  if (typeof value === 'boolean') {
    w.open('c', { r: ref, s: style, t: 'b' });
    w.leaf('v', value ? 1 : 0);
    w.close();
    return;
  }

  if (isError(value)) {
    w.open('c', { r: ref, s: style, t: 'e' });
    w.leaf('v', value.code);
    w.close();
    return;
  }

  // Only text remains. The null case returned above, but the formula branch
  // also returns, so the compiler needs the guard spelled out.
  if (typeof value !== 'string') return;

  // Text: either an index into the shared table, or inline.
  if (!inlineStrings && sst) {
    const index = sst.index.get(value);
    if (index !== undefined) {
      w.open('c', { r: ref, s: style, t: 's' });
      w.leaf('v', index);
      w.close();
      return;
    }
  }
  w.open('c', { r: ref, s: style, t: 'inlineStr' });
  w.open('is');
  w.leaf('t', escapeSharedString(value), needsSpacePreserve(value) ? { 'xml:space': 'preserve' } : undefined);
  w.close();
  w.close();
}

function formulaValueType(value: CellData['value']): string | undefined {
  if (typeof value === 'string') return 'str';
  if (typeof value === 'boolean') return 'b';
  if (isError(value)) return 'e';
  return undefined;
}

function formulaValueText(value: CellData['value']): string | number {
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (isError(value)) return value.code;
  if (typeof value === 'number') return String(value);
  return String(value ?? '');
}

/**
 * Re-apply the `_xlfn.` prefixes OOXML requires on post-2007 functions.
 *
 * The reader strips them so formulas display and evaluate normally; writing them
 * back is not optional, since Excel rejects a file that stores XLOOKUP under its
 * bare name.
 *
 * The rewrite is deliberately conservative: it only touches an identifier
 * immediately followed by an opening parenthesis, and skips anything already
 * prefixed or appearing inside a string literal.
 */
export function restoreFutureFunctionPrefixes(formula: string): string {
  let out = '';
  let i = 0;
  let inString = false;

  while (i < formula.length) {
    const ch = formula[i]!;
    if (inString) {
      out += ch;
      if (ch === '"') {
        // A doubled quote is an escaped quote, not the end of the string.
        if (formula[i + 1] === '"') {
          out += '"';
          i += 2;
          continue;
        }
        inString = false;
      }
      i++;
      continue;
    }
    if (ch === '"') {
      inString = true;
      out += ch;
      i++;
      continue;
    }
    if (!/[A-Za-z_]/.test(ch)) {
      out += ch;
      i++;
      continue;
    }

    let j = i;
    while (j < formula.length && /[A-Za-z0-9_.]/.test(formula[j]!)) j++;
    const word = formula.slice(i, j);
    // Only a name directly followed by '(' is a function call. The prefix is
    // prepended to the word as the author wrote it: Excel matches function
    // names case-insensitively, so rewriting the case would be a gratuitous
    // change to every formula in the file.
    if (formula[j] === '(' && !word.startsWith('_xlfn')) {
      out += (prefixFor(word.toUpperCase()) ?? '') + word;
    } else {
      out += word;
    }
    i = j;
  }
  return out;
}

/**
 * The set of names that must carry a storage prefix. Kept here rather than
 * imported from the formula package so that formats does not depend on formula;
 * the two lists are checked against each other by a test.
 */
const XLFN_NAMES: ReadonlySet<string> = new Set([
  'IFS', 'XOR', 'TEXTJOIN', 'CONCAT', 'SWITCH', 'MAXIFS', 'MINIFS', 'IFNA',
  'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'PERCENTILE.INC', 'PERCENTILE.EXC',
  'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'RANK.AVG', 'MODE.SNGL', 'MODE.MULT',
  'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'T.TEST', 'F.TEST',
  'CHISQ.TEST', 'COVARIANCE.P', 'COVARIANCE.S', 'BINOM.DIST', 'EXPON.DIST',
  'CEILING.MATH', 'FLOOR.MATH', 'CEILING.PRECISE', 'FLOOR.PRECISE',
  'XLOOKUP', 'XMATCH', 'LET', 'LAMBDA', 'BYROW', 'BYCOL', 'MAP', 'REDUCE',
  'SCAN', 'MAKEARRAY', 'ISOMITTED', 'UNIQUE', 'SORT', 'SORTBY', 'SEQUENCE',
  'RANDARRAY', 'ARRAYTOTEXT', 'VALUETOTEXT', 'TEXTSPLIT', 'TEXTBEFORE',
  'TEXTAFTER', 'VSTACK', 'HSTACK', 'TOCOL', 'TOROW', 'CHOOSECOLS',
  'CHOOSEROWS', 'WRAPROWS', 'WRAPCOLS', 'EXPAND', 'TAKE', 'DROP',
  'REGEXTEST', 'REGEXEXTRACT', 'REGEXREPLACE', 'GROUPBY', 'PIVOTBY',
  'PERCENTOF', 'TRIMRANGE', 'DAYS', 'ISOWEEKNUM', 'NUMBERVALUE', 'UNICHAR',
  'UNICODE', 'BASE', 'DECIMAL', 'COMBINA', 'PERMUTATIONA', 'SEC', 'CSC', 'COT',
  'ACOT', 'SECH', 'CSCH', 'COTH', 'ARABIC', 'BITAND', 'BITOR', 'BITXOR',
  'BITLSHIFT', 'BITRSHIFT', 'PDURATION', 'RRI', 'FORMULATEXT', 'SHEET', 'SHEETS',
  'IMTAN', 'IMCOT', 'IMCOSH', 'IMSINH', 'IMSEC', 'IMCSC', 'AGGREGATE', 'WORKDAY.INTL',
  'NETWORKDAYS.INTL', 'ERF.PRECISE', 'ERFC.PRECISE', 'GAMMA', 'GAUSS', 'PHI',
  'SKEW.P', 'WEIBULL.DIST', 'Z.TEST', 'F.DIST', 'F.INV', 'T.DIST', 'T.INV',
  'CHISQ.DIST', 'CHISQ.INV', 'BETA.DIST', 'BETA.INV', 'GAMMA.DIST', 'GAMMA.INV',
  'LOGNORM.DIST', 'LOGNORM.INV', 'HYPGEOM.DIST', 'NEGBINOM.DIST', 'POISSON.DIST',
  'CONFIDENCE.NORM', 'CONFIDENCE.T', 'BINOM.INV', 'BINOM.DIST.RANGE',
]);

/** Future functions stored under the worksheet namespace instead. */
const XLWS_NAMES: ReadonlySet<string> = new Set(['FILTER', 'ANCHORARRAY', 'SINGLE']);

/** The storage prefix a function name needs, or undefined when it needs none. */
function prefixFor(upper: string): string | undefined {
  if (XLWS_NAMES.has(upper)) return '_xlfn._xlws.';
  if (XLFN_NAMES.has(upper)) return '_xlfn.';
  return undefined;
}

export { XLFN_NAMES, XLWS_NAMES };
export { quoteSheetName };
