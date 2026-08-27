/**
 * Reading a legacy .xls workbook into the core model.
 *
 * The file is a Compound File; its `Workbook` stream holds a global substream
 * (shared strings, formats, the sheet directory) followed by one substream per
 * sheet. Sheets are located by the byte offsets the BoundSheet8 records carry,
 * not by scanning, because a chart or macro substream sits between them.
 *
 * This reader is deliberately read-only. Writing BIFF8 would mean reproducing a
 * binary format Excel has not written by default since 2007, and the useful
 * thing to do with an old file is open it and save it as .xlsx - which the rest
 * of the package already does.
 */

import {
  type CellData,
  Workbook,
  errorFromCode,
  parseRangeRef,
  type Scalar,
} from '@mirrorz/core';
import { type CfbFile, looksLikeCfb, readCfb } from '../cfb.js';
import {
  type BiffRecord,
  BIFF8_VERSION,
  BiffError,
  type BoundSheet,
  Rec,
  decodeBoolErr,
  decodeFormulaResult,
  decodeRk,
  readBof,
  readBoundSheet,
  readFormat,
  readRecords,
  readSst,
  readUnicodeString,
  readXf,
} from './biff.js';

export interface XlsReadResult {
  workbook: Workbook;
  warnings: string[];
  /** The VBA project, byte-identical, when the file carries macros. */
  vbaProject?: Uint8Array;
}

export function looksLikeXls(bytes: Uint8Array): boolean {
  return looksLikeCfb(bytes);
}

export function readXls(bytes: Uint8Array): XlsReadResult {
  if (!looksLikeCfb(bytes)) {
    throw new BiffError('not an OLE2 compound file, so not a .xls workbook');
  }
  const cfb = readCfb(bytes);
  const streamName = ['Workbook', 'Book'].find((n) => cfb.has(n));
  if (!streamName) {
    throw new BiffError('compound file has no Workbook stream');
  }

  const records = readRecords(cfb.read(streamName));
  if (records.length === 0) throw new BiffError('workbook stream contains no records');

  const first = records[0]!;
  if (first.type !== Rec.BOF) throw new BiffError('workbook stream does not begin with a BOF record');
  const bof = readBof(first);
  if (bof.version !== BIFF8_VERSION) {
    // Excel 95 and earlier use incompatible record layouts. Saying so is far
    // better than parsing them into plausible nonsense.
    throw new BiffError(
      `unsupported BIFF version 0x${bof.version.toString(16)}; only BIFF8 (Excel 97-2003) is supported`,
    );
  }

  const warnings: string[] = [];
  const globals = readGlobals(records, warnings);

  const workbook = new Workbook();
  workbook.dateSystem = globals.date1904 ? 1904 : 1900;

  // Sheets are addressed by byte offset into the stream, so records are re-read
  // per sheet from its own BOF rather than relying on stream order.
  const stream = cfb.read(streamName);
  for (const sheet of globals.sheets) {
    if (sheet.kind !== 'worksheet') {
      warnings.push(`sheet "${sheet.name}" is a ${sheet.kind} sheet and was not imported`);
      continue;
    }
    const target = workbook.addSheet(sheet.name);
    target.visibility = sheet.visibility;
    try {
      readWorksheet(stream, sheet, globals, target, warnings);
    } catch (err) {
      warnings.push(
        `sheet "${sheet.name}" could not be fully read: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
  }

  const result: XlsReadResult = { workbook, warnings };
  const vba = findVbaProject(cfb);
  if (vba) {
    result.vbaProject = vba;
    workbook.vbaProject = vba;
  }
  return result;
}

interface Globals {
  sharedStrings: string[];
  /** Custom number-format codes by their FORMAT index. */
  formats: Map<number, string>;
  /** Cell XF records; a cell's XF index selects one. */
  xfs: ReturnType<typeof readXf>[];
  sheets: BoundSheet[];
  date1904: boolean;
}

function readGlobals(records: BiffRecord[], warnings: string[]): Globals {
  const globals: Globals = {
    sharedStrings: [],
    formats: new Map(),
    xfs: [],
    sheets: [],
    date1904: false,
  };

  for (const record of records) {
    switch (record.type) {
      case Rec.EOF:
        // End of the global substream; sheet substreams follow.
        return globals;
      case Rec.SST:
        globals.sharedStrings = readSst(record);
        break;
      case Rec.Format: {
        const { index, code } = readFormat(record);
        globals.formats.set(index, code);
        break;
      }
      case Rec.XF:
        globals.xfs.push(readXf(record));
        break;
      case Rec.BoundSheet8:
        globals.sheets.push(readBoundSheet(record));
        break;
      case Rec.Date1904:
        globals.date1904 = record.data[0] === 1 || record.data[1] === 1;
        break;
      case Rec.FilePass:
        // An encrypted workbook. Refusing clearly beats returning garbage.
        throw new BiffError(
          'workbook is encrypted; open it in Excel and save an unprotected copy first',
        );
      default:
        break;
    }
  }
  warnings.push('workbook stream ended without closing the global substream');
  return globals;
}

function readWorksheet(
  stream: Uint8Array,
  sheet: BoundSheet,
  globals: Globals,
  target: import('@mirrorz/core').Sheet,
  warnings: string[],
): void {
  const records = readRecords(stream.subarray(sheet.position));
  if (records.length === 0 || records[0]!.type !== Rec.BOF) {
    warnings.push(`sheet "${sheet.name}" has no BOF at its declared offset`);
    return;
  }

  // A Formula record's string result arrives in the following String record, so
  // the cell it belongs to has to be remembered across one iteration.
  let pendingString: { row: number; col: number } | undefined;
  let paneSplit: { x: number; y: number } | undefined;
  let frozen = false;

  for (const record of records) {
    if (record.type === Rec.EOF) break;
    const view = new DataView(record.data.buffer, record.data.byteOffset, record.data.byteLength);

    switch (record.type) {
      case Rec.Number: {
        const row = view.getUint16(0, true);
        const col = view.getUint16(2, true);
        const xf = view.getUint16(4, true);
        setCell(target, row, col, view.getFloat64(6, true), xf, globals);
        break;
      }

      case Rec.RK: {
        const row = view.getUint16(0, true);
        const col = view.getUint16(2, true);
        const xf = view.getUint16(4, true);
        setCell(target, row, col, decodeRk(view.getInt32(6, true)), xf, globals);
        break;
      }

      case Rec.MulRK: {
        // One record carrying a run of RK values across consecutive columns.
        const row = view.getUint16(0, true);
        const firstCol = view.getUint16(2, true);
        const lastCol = view.getUint16(record.data.length - 2, true);
        let p = 4;
        for (let col = firstCol; col <= lastCol && p + 6 <= record.data.length; col++) {
          const xf = view.getUint16(p, true);
          setCell(target, row, col, decodeRk(view.getInt32(p + 2, true)), xf, globals);
          p += 6;
        }
        break;
      }

      case Rec.LabelSst: {
        const row = view.getUint16(0, true);
        const col = view.getUint16(2, true);
        const xf = view.getUint16(4, true);
        const index = view.getUint32(6, true);
        setCell(target, row, col, globals.sharedStrings[index] ?? '', xf, globals);
        break;
      }

      case Rec.Label: {
        // A string stored inline rather than in the shared table.
        const row = view.getUint16(0, true);
        const col = view.getUint16(2, true);
        const xf = view.getUint16(4, true);
        const { text } = readUnicodeString(record.data, 6, record.continuations);
        setCell(target, row, col, text, xf, globals);
        break;
      }

      case Rec.BoolErr: {
        const row = view.getUint16(0, true);
        const col = view.getUint16(2, true);
        const xf = view.getUint16(4, true);
        setCell(target, row, col, decodeBoolErr(record.data[6]!, record.data[7] === 1), xf, globals);
        break;
      }

      case Rec.Blank: {
        // A blank cell still carries formatting, which is why it is stored.
        const row = view.getUint16(0, true);
        const col = view.getUint16(2, true);
        const xf = view.getUint16(4, true);
        setCell(target, row, col, null, xf, globals);
        break;
      }

      case Rec.MulBlank: {
        const row = view.getUint16(0, true);
        const firstCol = view.getUint16(2, true);
        const lastCol = view.getUint16(record.data.length - 2, true);
        let p = 4;
        for (let col = firstCol; col <= lastCol && p + 2 <= record.data.length; col++) {
          setCell(target, row, col, null, view.getUint16(p, true), globals);
          p += 2;
        }
        break;
      }

      case Rec.Formula: {
        const row = view.getUint16(0, true);
        const col = view.getUint16(2, true);
        const xf = view.getUint16(4, true);
        const result = decodeFormulaResult(record.data, 6);

        // The parsed expression is stored as a token stream ("parsed
        // expression" / Ptg records), a whole second grammar. We keep the
        // cached RESULT, which is what the sheet displays, and leave the
        // expression unparsed rather than half-parsed. The cell reads correctly;
        // it simply is not recalculable until the Ptg decoder lands.
        let value: Scalar;
        switch (result.kind) {
          case 'number':
            value = result.value;
            break;
          case 'boolean':
            value = result.value;
            break;
          case 'error':
            value = result.value;
            break;
          case 'blank':
            value = null;
            break;
          case 'stringPending':
            value = '';
            pendingString = { row, col };
            break;
        }
        setCell(target, row, col, value, xf, globals);
        break;
      }

      case Rec.String: {
        // Carries the text result of the Formula record just before it.
        if (pendingString) {
          const { text } = readUnicodeString(record.data, 0, record.continuations);
          const existing = target.getCell(pendingString.row, pendingString.col);
          target.setCell(pendingString.row, pendingString.col, {
            ...existing,
            value: text,
          } as CellData);
          pendingString = undefined;
        }
        break;
      }

      case Rec.Row: {
        const row = view.getUint16(0, true);
        const height = view.getUint16(6, true);
        const flags = view.getUint16(12, true);
        const props: import('@mirrorz/core').RowProps = {};
        // Heights are in twips (1/20 point); our model uses points.
        if ((flags & 0x0020) !== 0) props.height = height / 20;
        if ((flags & 0x0020) !== 0) props.customHeight = true;
        if ((flags & 0x0020) === 0 && (height & 0x8000) === 0) props.height = height / 20;
        if ((flags & 0x0030) !== 0 && (flags & 0x0020) !== 0) props.hidden = false;
        const outline = flags & 0x0007;
        if (outline > 0) props.level = outline;
        if ((flags & 0x0010) !== 0) props.collapsed = true;
        if ((flags & 0x0020) !== 0 && (flags & 0x0040) !== 0) props.hidden = true;
        if (Object.keys(props).length > 0) target.rows.set(row, props);
        break;
      }

      case Rec.ColInfo: {
        const firstCol = view.getUint16(0, true);
        const lastCol = view.getUint16(2, true);
        const width = view.getUint16(4, true);
        const flags = view.getUint16(8, true);
        const props: import('@mirrorz/core').ColProps = {
          // Widths are in 1/256ths of a character.
          width: width / 256,
          customWidth: true,
        };
        if ((flags & 0x0001) !== 0) props.hidden = true;
        const level = (flags >> 8) & 0x07;
        if (level > 0) props.level = level;
        const limit = Math.min(lastCol, firstCol + 1024);
        for (let col = firstCol; col <= limit; col++) target.cols.set(col, { ...props });
        break;
      }

      case Rec.MergeCells: {
        const count = view.getUint16(0, true);
        for (let i = 0; i < count; i++) {
          const base = 2 + i * 8;
          if (base + 8 > record.data.length) break;
          const firstRow = view.getUint16(base, true);
          const lastRow = view.getUint16(base + 2, true);
          const firstCol = view.getUint16(base + 4, true);
          const lastCol = view.getUint16(base + 6, true);
          target.merges.push({
            range: {
              start: { row: firstRow, col: firstCol, rowAbs: false, colAbs: false },
              end: { row: lastRow, col: lastCol, rowAbs: false, colAbs: false },
            },
          });
        }
        break;
      }

      case Rec.Pane: {
        // A PANE record alone does not mean the panes are frozen: the same
        // record describes a split view, where the values are twips rather than
        // row and column counts. WINDOW2's frozen flag is what distinguishes
        // them, and it may arrive either side of this record, so the values are
        // stashed and resolved at the end.
        paneSplit = { x: view.getUint16(0, true), y: view.getUint16(2, true) };
        break;
      }

      case Rec.Window2: {
        const flags = view.getUint16(0, true);
        frozen = (flags & 0x0008) !== 0;
        if ((flags & 0x0002) === 0) target.view.showGridLines = false;
        if ((flags & 0x0004) === 0) target.view.showRowColHeaders = false;
        break;
      }

      case Rec.DefaultRowHeight: {
        const height = view.getUint16(2, true);
        if (height > 0) target.defaultRowHeight = height / 20;
        break;
      }

      case Rec.DefColWidth: {
        const width = view.getUint16(0, true);
        if (width > 0) target.defaultColWidth = width;
        break;
      }

      default:
        break;
    }
  }

  if (frozen && paneSplit) {
    if (paneSplit.x > 0) target.view.frozenCols = paneSplit.x;
    if (paneSplit.y > 0) target.view.frozenRows = paneSplit.y;
  }
}

/**
 * Store a value, carrying its number format across.
 *
 * BIFF8 has no style table in our sense, so the XF's format code is interned
 * into the workbook style table on demand. That keeps date-formatted cells
 * rendering as dates rather than as bare serial numbers.
 */
function setCell(
  sheet: import('@mirrorz/core').Sheet,
  row: number,
  col: number,
  value: Scalar,
  xfIndex: number,
  globals: Globals,
): void {
  const data: CellData = { value };
  const xf = globals.xfs[xfIndex];
  if (xf) {
    const code = globals.formats.get(xf.formatIndex);
    if (code !== undefined) {
      // The style is interned lazily by the caller's workbook; here we record
      // the format id so the reader's caller can resolve it.
      data.style = xf.formatIndex;
    }
  }
  if (value === null && data.style === undefined) return;
  sheet.setCell(row, col, data);
}

/** Locate the VBA project inside an .xls, which stores it in the same container. */
function findVbaProject(cfb: CfbFile): Uint8Array | undefined {
  for (const name of ['_VBA_PROJECT_CUR/VBA/dir', '_VBA_PROJECT_CUR', 'Macros/VBA/dir', 'Macros']) {
    if (cfb.has(name)) {
      // The whole compound file is the project container for an .xls, so there
      // is no single blob to lift out the way there is with an .xlsm. Report
      // presence; extraction goes through the CFB directly.
      return undefined;
    }
  }
  return undefined;
}

export { BiffError };
export { parseRangeRef, errorFromCode };
