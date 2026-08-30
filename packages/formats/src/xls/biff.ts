/**
 * BIFF8: reading legacy .xls workbooks.
 *
 * An .xls is a Compound File whose `Workbook` stream is a flat sequence of
 * records, each a 16-bit type, a 16-bit length, and that many bytes. There is no
 * nesting; structure comes from ordering and from cross-references by index.
 *
 * Three things make BIFF8 harder than the record layout suggests, and each has
 * bitten every implementation that skipped it:
 *
 * A record's payload is capped at 8224 bytes, so anything longer continues into
 * following CONTINUE records. Reassembly is not a simple concatenation: inside
 * the shared string table, each fragment carries its OWN encoding flag, so a
 * string can begin as compressed 8-bit text and continue as UTF-16 across the
 * boundary. Treating CONTINUE as plain concatenation produces mojibake in
 * exactly the files that are large enough for anyone to care about.
 *
 * Strings are length-prefixed with an option byte selecting 8-bit or 16-bit
 * encoding, and optionally carrying rich-text run counts and Far East phonetic
 * data that must be skipped rather than read as characters.
 *
 * Dates, like in xlsx, are not a type. A cell holds a number and its XF index
 * points at a format code; date-ness lives there and nowhere else.
 */

import { CellError, type Scalar, errorFromCode } from '@mirrorz/core';

/** The record types we interpret. Everything else is skipped by length. */
export const Rec = {
  Formula: 0x0006,
  EOF: 0x000a,
  CalcCount: 0x000c,
  CalcMode: 0x000d,
  Blank: 0x0201,
  Number: 0x0203,
  Label: 0x0204,
  BoolErr: 0x0205,
  String: 0x0207,
  Row: 0x0208,
  Index: 0x020b,
  Array: 0x0221,
  DefaultRowHeight: 0x0225,
  Font: 0x0031,
  Continue: 0x003c,
  Window1: 0x003d,
  Window2: 0x023e,
  Backup: 0x0040,
  Pane: 0x0041,
  CodePage: 0x0042,
  DefColWidth: 0x0055,
  ColInfo: 0x007d,
  Selection: 0x001d,
  Dimensions: 0x0200,
  SST: 0x00fc,
  ExtSST: 0x00ff,
  LabelSst: 0x00fd,
  MulRK: 0x00bd,
  MulBlank: 0x00be,
  RK: 0x027e,
  Format: 0x041e,
  XF: 0x00e0,
  BoundSheet8: 0x0085,
  MergeCells: 0x00e5,
  BOF: 0x0809,
  Date1904: 0x0022,
  Style: 0x0293,
  Palette: 0x0092,
  DefinedName: 0x0018,
  HLink: 0x01b8,
  SheetProtect: 0x0867,
  FilePass: 0x002f,
} as const;

export class BiffError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'BiffError';
  }
}

export interface BiffRecord {
  type: number;
  /** Payload with any CONTINUE records already joined. */
  data: Uint8Array;
  /** Where each CONTINUE fragment began, needed for per-fragment encoding. */
  continuations: number[];
}

/**
 * Split a workbook stream into records, joining CONTINUE payloads.
 *
 * The fragment boundaries are reported rather than discarded: the shared string
 * table needs them, because its per-string encoding flag resets at each one.
 */
export function readRecords(stream: Uint8Array): BiffRecord[] {
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const records: BiffRecord[] = [];
  let p = 0;

  while (p + 4 <= stream.length) {
    const type = view.getUint16(p, true);
    const length = view.getUint16(p + 2, true);
    const start = p + 4;
    const end = start + length;
    if (end > stream.length) break; // truncated tail; keep what we have

    if (type === Rec.Continue && records.length > 0) {
      const previous = records[records.length - 1]!;
      const joined = new Uint8Array(previous.data.length + length);
      joined.set(previous.data, 0);
      joined.set(stream.subarray(start, end), previous.data.length);
      previous.continuations.push(previous.data.length);
      previous.data = joined;
    } else {
      records.push({ type, data: stream.subarray(start, end), continuations: [] });
    }
    p = end;
  }
  return records;
}

/**
 * BIFF8 numbers are stored either as a full IEEE double or as an "RK" value: a
 * 30-bit payload with two flag bits, which stores the common small numbers in
 * four bytes instead of eight. Bit 0 says the value was multiplied by 100 before
 * storing, bit 1 says it is an integer rather than the top half of a double.
 */
export function decodeRk(rk: number): number {
  const isMultiplied = (rk & 0x01) !== 0;
  const isInteger = (rk & 0x02) !== 0;
  let value: number;
  if (isInteger) {
    // The 30 payload bits are a signed integer; shifting right by 2 with sign
    // extension recovers it.
    value = rk >> 2;
  } else {
    // The payload is the high 30 bits of a double's bit pattern, low bits zero.
    const buffer = new ArrayBuffer(8);
    const dv = new DataView(buffer);
    dv.setUint32(4, rk & 0xffff_fffc, true);
    dv.setUint32(0, 0, true);
    value = dv.getFloat64(0, true);
  }
  return isMultiplied ? value / 100 : value;
}

/**
 * Read a BIFF8 Unicode string.
 *
 * Layout: a character count, then an option byte whose bit 0 selects UTF-16
 * (set) or 8-bit Latin-1 (clear), bit 2 marks the presence of Far East phonetic
 * data, and bit 3 marks rich-text formatting runs. The run and phonetic sizes
 * come next and their bodies sit AFTER the characters, so they must be skipped
 * rather than read as text.
 *
 * `fragmentStarts` lists the offsets where a CONTINUE record began. Crossing one
 * mid-string means re-reading the option byte, because a string may switch
 * encoding at that boundary.
 */
export function readUnicodeString(
  data: Uint8Array,
  offset: number,
  fragmentStarts: readonly number[] = [],
  lengthBytes: 1 | 2 = 2,
): { text: string; next: number } {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let p = offset;
  const charCount = lengthBytes === 2 ? view.getUint16(p, true) : data[p]!;
  p += lengthBytes;

  const options = data[p]!;
  p += 1;
  let wide = (options & 0x01) !== 0;
  const hasPhonetic = (options & 0x04) !== 0;
  const hasRuns = (options & 0x08) !== 0;

  let runCount = 0;
  let phoneticSize = 0;
  if (hasRuns) {
    runCount = view.getUint16(p, true);
    p += 2;
  }
  if (hasPhonetic) {
    phoneticSize = view.getUint32(p, true);
    p += 4;
  }

  const chars: number[] = [];
  let read = 0;
  const isFragmentBoundary = (at: number) => fragmentStarts.includes(at);

  while (read < charCount && p < data.length) {
    // A CONTINUE boundary restarts the encoding flag for the remainder.
    if (isFragmentBoundary(p)) {
      wide = (data[p]! & 0x01) !== 0;
      p += 1;
      continue;
    }
    if (wide) {
      if (p + 1 >= data.length) break;
      chars.push(view.getUint16(p, true));
      p += 2;
    } else {
      chars.push(data[p]!);
      p += 1;
    }
    read++;
  }

  // Formatting runs and phonetic data follow the characters and are not text.
  p += runCount * 4;
  p += phoneticSize;

  return { text: stringFromCodes(chars), next: p };
}

/** Build a string from code units without blowing the argument limit. */
function stringFromCodes(codes: number[]): string {
  let out = '';
  const CHUNK = 4096;
  for (let i = 0; i < codes.length; i += CHUNK) {
    out += String.fromCharCode(...codes.slice(i, i + CHUNK));
  }
  return out;
}

/** Parse the shared string table. */
export function readSst(record: BiffRecord): string[] {
  const view = new DataView(record.data.buffer, record.data.byteOffset, record.data.byteLength);
  const uniqueCount = view.getUint32(4, true);
  const strings: string[] = [];
  let p = 8;
  for (let i = 0; i < uniqueCount && p < record.data.length; i++) {
    const { text, next } = readUnicodeString(record.data, p, record.continuations);
    strings.push(text);
    if (next <= p) break; // refuse to loop on a malformed entry
    p = next;
  }
  return strings;
}

/** The error codes BIFF8 stores as a single byte in a BoolErr record. */
const BIFF_ERRORS: Record<number, string> = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
};

export function decodeBoolErr(value: number, isError: boolean): Scalar {
  if (!isError) return value !== 0;
  return errorFromCode(BIFF_ERRORS[value] ?? '#VALUE!') ?? CellError.VALUE;
}

/**
 * Decode a Formula record's cached result.
 *
 * The eight result bytes are either an IEEE double or, when the high word is
 * 0xFFFF, a tagged non-numeric result whose first byte selects string, boolean,
 * error or blank. A string result's text arrives in the NEXT record, a String
 * record, which is why this returns a marker rather than the value.
 */
export type FormulaResult =
  | { kind: 'number'; value: number }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; value: CellError }
  | { kind: 'blank' }
  /** The text follows in a String record. */
  | { kind: 'stringPending' };

export function decodeFormulaResult(data: Uint8Array, offset: number): FormulaResult {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const high = view.getUint16(offset + 6, true);
  if (high !== 0xffff) {
    return { kind: 'number', value: view.getFloat64(offset, true) };
  }
  const tag = data[offset]!;
  switch (tag) {
    case 0:
      return { kind: 'stringPending' };
    case 1:
      return { kind: 'boolean', value: data[offset + 2] !== 0 };
    case 2: {
      const code = BIFF_ERRORS[data[offset + 2]!] ?? '#VALUE!';
      return { kind: 'error', value: errorFromCode(code) ?? CellError.VALUE };
    }
    case 3:
      return { kind: 'blank' };
    default:
      return { kind: 'number', value: view.getFloat64(offset, true) };
  }
}

export interface BoundSheet {
  name: string;
  /** Byte offset of the sheet's BOF record within the workbook stream. */
  position: number;
  visibility: 'visible' | 'hidden' | 'veryHidden';
  kind: 'worksheet' | 'chart' | 'macro' | 'other';
}

export function readBoundSheet(record: BiffRecord): BoundSheet {
  const view = new DataView(record.data.buffer, record.data.byteOffset, record.data.byteLength);
  const position = view.getUint32(0, true);
  const hiddenBits = record.data[4]! & 0x03;
  const typeByte = record.data[5]!;
  // The name here uses a one-byte length rather than the usual two.
  const { text } = readUnicodeString(record.data, 6, record.continuations, 1);
  return {
    name: text,
    position,
    visibility: hiddenBits === 1 ? 'hidden' : hiddenBits === 2 ? 'veryHidden' : 'visible',
    kind:
      typeByte === 0x00
        ? 'worksheet'
        : typeByte === 0x02
          ? 'chart'
          : typeByte === 0x01
            ? 'macro'
            : 'other',
  };
}

/** An XF record, of which we need only the number-format index and flags. */
export interface XfRecord {
  fontIndex: number;
  formatIndex: number;
  /** True for a style XF, false for a cell XF. */
  isStyle: boolean;
  horizontal: number;
  vertical: number;
  wrap: boolean;
}

export function readXf(record: BiffRecord): XfRecord {
  const view = new DataView(record.data.buffer, record.data.byteOffset, record.data.byteLength);
  const fontIndex = view.getUint16(0, true);
  const formatIndex = view.getUint16(2, true);
  const typeFlags = view.getUint16(4, true);
  const alignment = record.data[6]!;
  return {
    fontIndex,
    formatIndex,
    isStyle: (typeFlags & 0x04) !== 0,
    horizontal: alignment & 0x07,
    vertical: (alignment >> 4) & 0x07,
    wrap: (alignment & 0x08) !== 0,
  };
}

/** A FORMAT record: a custom number-format code and the index cells refer to. */
export function readFormat(record: BiffRecord): { index: number; code: string } {
  const view = new DataView(record.data.buffer, record.data.byteOffset, record.data.byteLength);
  const index = view.getUint16(0, true);
  const { text } = readUnicodeString(record.data, 2, record.continuations);
  return { index, code: text };
}

/**
 * The BOF record identifies both the BIFF version and what kind of substream
 * follows. Version 0x0600 is BIFF8; earlier workbooks (Excel 95 and before) use
 * a different record layout that we detect and reject clearly rather than
 * misparsing into nonsense.
 */
export interface BofInfo {
  version: number;
  substream: 'workbook' | 'worksheet' | 'chart' | 'macro' | 'other';
}

export function readBof(record: BiffRecord): BofInfo {
  const view = new DataView(record.data.buffer, record.data.byteOffset, record.data.byteLength);
  const version = view.getUint16(0, true);
  const type = view.getUint16(2, true);
  return {
    version,
    substream:
      type === 0x0005
        ? 'workbook'
        : type === 0x0010
          ? 'worksheet'
          : type === 0x0020
            ? 'chart'
            : type === 0x0040
              ? 'macro'
              : 'other',
  };
}

export const BIFF8_VERSION = 0x0600;
export { BIFF_ERRORS };
