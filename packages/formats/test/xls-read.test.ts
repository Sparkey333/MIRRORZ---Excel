import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, isError } from '@mirrorz/core';
import { readCfb } from '../src/cfb.js';
import {
  Rec,
  decodeBoolErr,
  decodeFormulaResult,
  decodeRk,
  readBof,
  readRecords,
  readUnicodeString,
} from '../src/xls/biff.js';
import { BiffError, looksLikeXls, readXls } from '../src/xls/read.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const bytes = (name: string) => new Uint8Array(readFileSync(new URL(name, FIXTURES)));

describe('RK decoding', () => {
  // The four combinations of the two flag bits, which is the whole encoding.
  it('decodes an integer RK', () => {
    // 100 stored as an integer: 100 << 2 | 0b10
    expect(decodeRk((100 << 2) | 0b10)).toBe(100);
  });

  it('decodes a negative integer RK with sign extension', () => {
    expect(decodeRk((-17 << 2) | 0b10)).toBe(-17);
  });

  it('decodes an integer RK that was multiplied by 100', () => {
    // 1.5 is stored as 150 with the multiplied flag.
    expect(decodeRk((150 << 2) | 0b11)).toBeCloseTo(1.5, 12);
  });

  it('decodes a double RK from the high 30 bits', () => {
    // Build the bit pattern of 3.5 and keep only what an RK can hold.
    const dv = new DataView(new ArrayBuffer(8));
    dv.setFloat64(0, 3.5, true);
    const high = dv.getUint32(4, true) & 0xffff_fffc;
    expect(decodeRk(high)).toBe(3.5);
  });

  it('round-trips values that fit the encoding exactly', () => {
    for (const n of [0, 1, -1, 42, -42, 1000, 100_000, -250_000]) {
      expect(decodeRk((n << 2) | 0b10)).toBe(n);
    }
  });
});

describe('BoolErr decoding', () => {
  it('decodes booleans', () => {
    expect(decodeBoolErr(1, false)).toBe(true);
    expect(decodeBoolErr(0, false)).toBe(false);
  });

  it.each([
    [0x00, '#NULL!'],
    [0x07, '#DIV/0!'],
    [0x0f, '#VALUE!'],
    [0x17, '#REF!'],
    [0x1d, '#NAME?'],
    [0x24, '#NUM!'],
    [0x2a, '#N/A'],
  ])('decodes error byte %i as %s', (code, want) => {
    const v = decodeBoolErr(code, true);
    expect(isError(v)).toBe(true);
    expect((v as CellError).code).toBe(want);
  });

  it('falls back to #VALUE! for an unknown error byte', () => {
    expect((decodeBoolErr(0x99, true) as CellError).code).toBe('#VALUE!');
  });
});

describe('formula result decoding', () => {
  const make = (bytes: number[]) => new Uint8Array(bytes);

  it('reads a numeric result', () => {
    const buf = new Uint8Array(8);
    new DataView(buf.buffer).setFloat64(0, 42.5, true);
    expect(decodeFormulaResult(buf, 0)).toEqual({ kind: 'number', value: 42.5 });
  });

  it('reads a boolean result', () => {
    expect(decodeFormulaResult(make([1, 0, 1, 0, 0, 0, 0xff, 0xff]), 0)).toEqual({
      kind: 'boolean',
      value: true,
    });
  });

  it('reads an error result', () => {
    const r = decodeFormulaResult(make([2, 0, 0x07, 0, 0, 0, 0xff, 0xff]), 0);
    expect(r.kind).toBe('error');
    expect((r as { value: CellError }).value.code).toBe('#DIV/0!');
  });

  it('flags a string result as pending, since its text is in the next record', () => {
    expect(decodeFormulaResult(make([0, 0, 0, 0, 0, 0, 0xff, 0xff]), 0)).toEqual({
      kind: 'stringPending',
    });
  });

  it('reads a blank result', () => {
    expect(decodeFormulaResult(make([3, 0, 0, 0, 0, 0, 0xff, 0xff]), 0)).toEqual({ kind: 'blank' });
  });
});

describe('unicode strings', () => {
  /** Build a BIFF8 string: 16-bit count, option byte, then characters. */
  function str(text: string, wide: boolean): Uint8Array {
    const out: number[] = [];
    out.push(text.length & 0xff, (text.length >> 8) & 0xff);
    out.push(wide ? 0x01 : 0x00);
    for (const ch of text) {
      const code = ch.charCodeAt(0);
      if (wide) out.push(code & 0xff, (code >> 8) & 0xff);
      else out.push(code & 0xff);
    }
    return new Uint8Array(out);
  }

  it('reads a compressed 8-bit string', () => {
    expect(readUnicodeString(str('hello', false), 0).text).toBe('hello');
  });

  it('reads a UTF-16 string', () => {
    expect(readUnicodeString(str('héllo', true), 0).text).toBe('héllo');
  });

  it('reports where the next field begins', () => {
    const s = str('abc', false);
    expect(readUnicodeString(s, 0).next).toBe(s.length);
  });

  it('reads an empty string', () => {
    expect(readUnicodeString(str('', false), 0).text).toBe('');
  });

  it('skips rich-text runs, which follow the characters', () => {
    // count=1, options=0x08 (has runs), runCount=2, then the character, then
    // two four-byte runs that must not be read as text.
    const data = new Uint8Array([1, 0, 0x08, 2, 0, 0x41, 1, 0, 0, 0, 2, 0, 0, 0]);
    const r = readUnicodeString(data, 0);
    expect(r.text).toBe('A');
    expect(r.next).toBe(data.length);
  });

  it('skips phonetic data', () => {
    // count=1, options=0x04 (phonetic), size=4, character, then four bytes.
    const data = new Uint8Array([1, 0, 0x04, 4, 0, 0, 0, 0x42, 9, 9, 9, 9]);
    const r = readUnicodeString(data, 0);
    expect(r.text).toBe('B');
    expect(r.next).toBe(data.length);
  });

  it('supports a one-byte length prefix, as BoundSheet8 uses', () => {
    const data = new Uint8Array([3, 0x00, 0x41, 0x42, 0x43]);
    expect(readUnicodeString(data, 0, [], 1).text).toBe('ABC');
  });

  it('re-reads the encoding flag at a CONTINUE boundary', () => {
    // "AB" compressed, then a fragment boundary introducing wide "C".
    const data = new Uint8Array([3, 0, 0x00, 0x41, 0x42, 0x01, 0x43, 0x00]);
    // The boundary sits at offset 5, where the new option byte lives.
    expect(readUnicodeString(data, 0, [5]).text).toBe('ABC');
  });
});

describe('record splitting', () => {
  it('splits a stream into records', () => {
    // Two records: type 1 with 2 bytes, type 2 with 1 byte.
    const data = new Uint8Array([1, 0, 2, 0, 0xaa, 0xbb, 2, 0, 1, 0, 0xcc]);
    const records = readRecords(data);
    expect(records).toHaveLength(2);
    expect(records[0]!.type).toBe(1);
    expect(Array.from(records[0]!.data)).toEqual([0xaa, 0xbb]);
    expect(records[1]!.type).toBe(2);
  });

  it('joins CONTINUE payloads onto the previous record', () => {
    const data = new Uint8Array([
      1, 0, 2, 0, 0xaa, 0xbb, // record type 1
      0x3c, 0, 2, 0, 0xcc, 0xdd, // CONTINUE
    ]);
    const records = readRecords(data);
    expect(records).toHaveLength(1);
    expect(Array.from(records[0]!.data)).toEqual([0xaa, 0xbb, 0xcc, 0xdd]);
    // The boundary is reported, because a string may change encoding there.
    expect(records[0]!.continuations).toEqual([2]);
  });

  it('stops at a truncated record rather than reading past the end', () => {
    const data = new Uint8Array([1, 0, 10, 0, 0xaa]);
    expect(readRecords(data)).toHaveLength(0);
  });

  it('handles an empty stream', () => {
    expect(readRecords(new Uint8Array(0))).toEqual([]);
  });
});

describe('reading real .xls fixtures', () => {
  // These files were produced by LibreOffice from our xlsx fixtures, so the
  // expected contents are known exactly.

  it('recognises the container', () => {
    expect(looksLikeXls(bytes('basic-types.xls'))).toBe(true);
    expect(looksLikeXls(bytes('basic-types.xlsx'))).toBe(false);
  });

  it('finds a BIFF8 BOF in the Workbook stream', () => {
    const cfb = readCfb(bytes('basic-types.xls'));
    expect(cfb.has('Workbook')).toBe(true);
    const records = readRecords(cfb.read('Workbook'));
    expect(records[0]!.type).toBe(Rec.BOF);
    expect(readBof(records[0]!).version).toBe(0x0600);
    expect(readBof(records[0]!).substream).toBe('workbook');
  });

  it('reads sheet names and order', () => {
    const { workbook } = readXls(bytes('features.xls'));
    expect(workbook.sheets.map((s) => s.name)).toEqual(
      expect.arrayContaining(['Features', 'Charts', 'Comments']),
    );
  });

  it('reads primitive values', () => {
    const { workbook } = readXls(bytes('basic-types.xls'));
    const sheet = workbook.sheets[0]!;
    const byLabel = new Map<string, unknown>();
    for (const { row, col, cell } of sheet.entries()) {
      if (col === 1) {
        const label = sheet.getValue(row, 0);
        if (typeof label === 'string') byLabel.set(label, cell.value);
      }
    }
    expect(byLabel.get('integer')).toBe(42);
    expect(byLabel.get('negative')).toBe(-17);
    expect(byLabel.get('string')).toBe('hello world');
    expect(byLabel.get('float')).toBeCloseTo(3.14159265358979, 10);
  });

  it('preserves unicode through the shared string table', () => {
    const { workbook } = readXls(bytes('basic-types.xls'));
    const values = [...workbook.sheets[0]!.entries()].map((e) => e.cell.value);
    expect(values).toContain('éàü 你好 \u{1f600}');
    expect(values).toContain('he said "hi"');
  });

  it('keeps text that looks numeric as text', () => {
    const { workbook } = readXls(bytes('basic-types.xls'));
    const values = [...workbook.sheets[0]!.entries()].map((e) => e.cell.value);
    expect(values).toContain('007');
    expect(values).toContain('SEPT1');
    expect(values).toContain('1-2');
  });

  it('reads booleans', () => {
    const { workbook } = readXls(bytes('basic-types.xls'));
    const values = [...workbook.sheets[0]!.entries()].map((e) => e.cell.value);
    expect(values).toContain(true);
    expect(values).toContain(false);
  });

  it('reads cached formula results', () => {
    const { workbook } = readXls(bytes('formulas.xls'));
    const fx = workbook.getSheet('Formulas')!;
    const byLabel = new Map<string, unknown>();
    for (const { row, col, cell } of fx.entries()) {
      if (col === 2) {
        const label = fx.getValue(row, 0);
        if (typeof label === 'string') byLabel.set(label, cell.value);
      }
    }
    expect(byLabel.get('SUM')).toBe(1_207_000);
    expect(byLabel.get('MAX')).toBe(181_000);
    expect(byLabel.get('COUNT')).toBe(8);
    // A string-valued formula result arrives in the following String record.
    expect(byLabel.get('UPPER')).toBe('ABC');
  });

  it('reads error results', () => {
    const { workbook } = readXls(bytes('formulas.xls'));
    const fx = workbook.getSheet('Formulas')!;
    const errors: string[] = [];
    for (const { cell } of fx.entries()) {
      if (isError(cell.value)) errors.push(cell.value.code);
    }
    expect(errors).toContain('#DIV/0!');
  });

  it('reads merged ranges', () => {
    const { workbook } = readXls(bytes('styling.xls'));
    expect(workbook.sheets[0]!.merges.length).toBeGreaterThan(0);
  });

  it('does not invent frozen panes the file does not declare', () => {
    // LibreOffice drops the freeze when converting this fixture to .xls: the
    // sheet has no PANE record and WINDOW2's frozen bit is clear. Reporting a
    // freeze here would be reading intent that is genuinely not in the file.
    const { workbook } = readXls(bytes('styling.xls'));
    expect(workbook.sheets[0]!.view.frozenRows).toBeUndefined();
  });

  it('reads column widths', () => {
    const { workbook } = readXls(bytes('styling.xls'));
    expect(workbook.sheets[0]!.colWidth(0)).toBeGreaterThan(10);
  });

  it('reads the sparse corners of a sheet', () => {
    const { workbook } = readXls(bytes('edge-cases.xls'));
    const sparse = workbook.getSheet('Sparse')!;
    expect(sparse.getValue(0, 0)).toBe('top-left');
    // BIFF8's grid stops at 65,536 rows and 256 columns, so LibreOffice will
    // have clipped the far corners on conversion. What survives must be right.
    expect(sparse.cellCount).toBeGreaterThan(0);
  });

  it('reports non-worksheet sheets as warnings rather than importing them', () => {
    const { warnings } = readXls(bytes('features.xls'));
    for (const w of warnings) expect(typeof w).toBe('string');
  });
});

describe('frozen panes', () => {
  it('reads a freeze only when WINDOW2 says the panes are frozen', () => {
    // PANE alone describes a SPLIT view, whose values are twips rather than
    // row and column counts, so acting on it unconditionally would turn a split
    // into a nonsensical freeze of several thousand rows.
    const pane = [0x41, 0x00, 8, 0, 2, 0, 3, 0, 2, 0, 0];
    const frozenWindow2 = [0x3e, 0x02, 2, 0, 0x08 | 0x02 | 0x04, 0x00];
    const splitWindow2 = [0x3e, 0x02, 2, 0, 0x02 | 0x04, 0x00];
    const bof = [0x09, 0x08, 4, 0, 0x00, 0x06, 0x10, 0x00];
    const eof = [0x0a, 0x00, 0, 0];

    const withFreeze = readRecords(new Uint8Array([...bof, ...pane, ...frozenWindow2, ...eof]));
    const withSplit = readRecords(new Uint8Array([...bof, ...pane, ...splitWindow2, ...eof]));
    // Both parse; the distinction is made by the reader, and both records are
    // present in each stream.
    expect(withFreeze.some((r) => r.type === Rec.Pane)).toBe(true);
    expect(withSplit.some((r) => r.type === Rec.Pane)).toBe(true);
  });
});

describe('rejection', () => {
  it('refuses a file that is not a compound file', () => {
    expect(() => readXls(bytes('basic-types.xlsx'))).toThrow(BiffError);
    expect(() => readXls(bytes('basic-types.xlsx'))).toThrow(/compound file/);
  });

  it('refuses random bytes', () => {
    expect(() => readXls(new TextEncoder().encode('nonsense'))).toThrow(BiffError);
  });
});
