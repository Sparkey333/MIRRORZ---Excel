import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError, MAX_COLS, Workbook, partsToSerial } from '../../core/src/index.js';
import {
  CsvRowParser,
  decodeCsvBytes,
  detectDelimiter,
  formatScalar,
  inferValue,
  looksLikeGeneName,
  needsQuoting,
  parseDelimited,
  readCsv,
  writeCsv,
  writeCsvBytes,
  writeRows,
  type CsvInferenceOptions,
} from '../src/csv.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const fixtureText = (name: string) => readFileSync(new URL(name, FIXTURES), 'utf8');
const fixtureBytes = (name: string) => new Uint8Array(readFileSync(new URL(name, FIXTURES)));

/** Fields only, for the many cases where warnings are not the point. */
const fields = (text: string, delimiter = ','): string[][] =>
  parseDelimited(text, { delimiter }).rows;

const val = (text: string, options: CsvInferenceOptions = {}) => inferValue(text, options).value;
const fmt = (text: string, options: CsvInferenceOptions = {}) => inferValue(text, options).numFmt;

/** A one-sheet workbook whose sheet is ready to be written out. */
function sheetOf(rows: (string | number | boolean | null)[][]) {
  const wb = new Workbook();
  const sheet = wb.addSheet('S');
  rows.forEach((row, r) =>
    row.forEach((v, c) => {
      if (v !== null) sheet.setValue(r, c, v);
    }),
  );
  return { wb, sheet };
}

describe('field splitting', () => {
  it('splits a plain row', () => {
    expect(fields('a,b,c')).toEqual([['a', 'b', 'c']]);
  });

  it('keeps empty fields', () => {
    expect(fields('a,,c')).toEqual([['a', '', 'c']]);
  });

  it('keeps a trailing empty field', () => {
    expect(fields('a,b,')).toEqual([['a', 'b', '']]);
  });

  it('keeps a leading empty field', () => {
    expect(fields(',b')).toEqual([['', 'b']]);
  });

  it('reads a single field with no delimiter', () => {
    expect(fields('solo')).toEqual([['solo']]);
  });

  it('returns no rows for empty input', () => {
    expect(fields('')).toEqual([]);
  });

  it('drops the row implied by a trailing newline', () => {
    expect(fields('a,b\n')).toEqual([['a', 'b']]);
  });

  it('distinguishes one empty row from no rows at all', () => {
    // An empty input is an empty file: zero rows, matching what a spreadsheet
    // does when it opens a zero-byte CSV. A single row holding one empty field
    // is a different thing, and writing it emits a line terminator so the two
    // stay distinguishable and the round trip is an identity.
    expect(writeRows([['']])).toBe('\r\n');
    expect(parseDelimited('\r\n').rows).toEqual([['']]);
    expect(parseDelimited('').rows).toEqual([]);
    expect(parseDelimited(writeRows([['']])).rows).toEqual([['']]);
  });

  it('keeps an interior blank line as an empty row', () => {
    expect(fields('a\n\nb')).toEqual([['a'], [''], ['b']]);
  });

  it('preserves spaces around fields', () => {
    expect(fields('a , b ,c')).toEqual([['a ', ' b ', 'c']]);
  });

  it('splits on tabs when told to', () => {
    expect(fields('a\tb\tc', '\t')).toEqual([['a', 'b', 'c']]);
  });

  it('rejects a multi-character delimiter', () => {
    expect(() => new CsvRowParser('::', { row() {} })).toThrow(RangeError);
  });

  it('rejects a delimiter that would break the grammar', () => {
    expect(() => new CsvRowParser('"', { row() {} })).toThrow(RangeError);
    expect(() => new CsvRowParser('\n', { row() {} })).toThrow(RangeError);
  });
});

describe('RFC 4180 quoting', () => {
  it('unwraps a quoted field', () => {
    expect(fields('"a",b')).toEqual([['a', 'b']]);
  });

  it('keeps a delimiter inside quotes', () => {
    expect(fields('"a,b",c')).toEqual([['a,b', 'c']]);
  });

  it('unescapes doubled quotes', () => {
    expect(fields('"he said ""hi"""')).toEqual([['he said "hi"']]);
  });

  it('reads a field that is only a doubled quote', () => {
    expect(fields('""""')).toEqual([['"']]);
  });

  it('distinguishes an empty quoted field from a missing one', () => {
    expect(fields('"",')).toEqual([['', '']]);
  });

  it('keeps an embedded newline', () => {
    expect(fields('"line1\nline2",b')).toEqual([['line1\nline2', 'b']]);
  });

  it('normalises an embedded CRLF to LF', () => {
    expect(fields('"line1\r\nline2"')).toEqual([['line1\nline2']]);
  });

  it('normalises an embedded lone CR to LF', () => {
    expect(fields('"line1\rline2"')).toEqual([['line1\nline2']]);
  });

  it('does not let a CR swallow a newline that is not its own', () => {
    // Only the LF immediately after a CR is that CR's other half. Here a
    // doubled quote sits between them, so they are two separate line breaks and
    // the second one must survive.
    expect(fields('"a\r""b""\nc"')).toEqual([['a\n"b"\nc']]);
    expect(fields('"a\r"",\nb"')).toEqual([['a\n",\nb']]);
    // The genuine CRLF pair still collapses to one break.
    expect(fields('"a\r\nb"')).toEqual([['a\nb']]);
    expect(fields('"a\r\r\nb"')).toEqual([['a\n\nb']]);
  });

  it('keeps leading and trailing spaces inside quotes', () => {
    expect(fields('"  padded  ",x')).toEqual([['  padded  ', 'x']]);
  });

  it('treats a quote inside an unquoted field as data, and says so', () => {
    const r = parseDelimited('a"b,c', { delimiter: ',' });
    expect(r.rows).toEqual([['a"b', 'c']]);
    // RFC 4180 has no reading for a bare quote in a non-escaped field. Keeping
    // it as data is the tolerant choice; keeping it silently is not, because
    // the more likely story is a field that was meant to be quoted and lost a
    // delimiter's worth of structure on the way.
    expect(r.warnings.map((w) => w.code)).toEqual(['quote-in-unquoted-field']);
    expect(r.warnings[0]!.row).toBe(0);
    expect(r.warnings[0]!.col).toBe(0);
  });

  it('warns once per field about a stray quote', () => {
    const r = parseDelimited('a"b"c"d', { delimiter: ',' });
    expect(r.rows).toEqual([['a"b"c"d']]);
    expect(r.warnings.filter((w) => w.code === 'quote-in-unquoted-field')).toHaveLength(1);
  });

  it('does not confuse a stray quote with text after a closing quote', () => {
    const r = parseDelimited('"a"x,b"c', { delimiter: ',' });
    expect(r.rows).toEqual([['ax', 'b"c']]);
    expect(r.warnings.map((w) => w.code)).toEqual(['text-after-quote', 'quote-in-unquoted-field']);
  });

  it('says nothing about a properly quoted file', () => {
    expect(parseDelimited('"a","b,c"\n"d",""', { delimiter: ',' }).warnings).toEqual([]);
  });

  it('treats a quote after a space as data', () => {
    expect(fields(' "a",b')).toEqual([[' "a"', 'b']]);
  });

  it('appends text found after a closing quote and warns', () => {
    const r = parseDelimited('"a"x,b', { delimiter: ',' });
    expect(r.rows).toEqual([['ax', 'b']]);
    expect(r.warnings.map((w) => w.code)).toContain('text-after-quote');
  });

  it('warns once per field about text after a closing quote', () => {
    const r = parseDelimited('"a"xyz', { delimiter: ',' });
    expect(r.warnings.filter((w) => w.code === 'text-after-quote')).toHaveLength(1);
  });

  it('recovers from an unterminated quote and warns', () => {
    const r = parseDelimited('a,"b,c', { delimiter: ',' });
    expect(r.rows).toEqual([['a', 'b,c']]);
    expect(r.warnings.map((w) => w.code)).toContain('unterminated-quote');
  });

  it('quotes are not special in the middle of the file', () => {
    expect(fields('a,b\n"c",d')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });
});

describe('line endings', () => {
  it('splits on LF', () => {
    expect(fields('a\nb')).toEqual([['a'], ['b']]);
  });

  it('splits on CRLF', () => {
    expect(fields('a\r\nb')).toEqual([['a'], ['b']]);
  });

  it('splits on a lone CR', () => {
    expect(fields('a\rb')).toEqual([['a'], ['b']]);
  });

  it('handles mixed terminators in one file', () => {
    expect(fields('a\r\nb\nc\rd')).toEqual([['a'], ['b'], ['c'], ['d']]);
  });

  it('drops the row implied by a trailing CRLF', () => {
    expect(fields('a,b\r\n')).toEqual([['a', 'b']]);
  });

  it('does not lose a field ending at CRLF', () => {
    expect(fields('a,b\r\nc,d')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });
});

describe('streaming', () => {
  const collect = (chunks: string[]): string[][] => {
    const rows: string[][] = [];
    const p = new CsvRowParser(',', { row: (f) => void rows.push(f) });
    for (const c of chunks) p.push(c);
    p.end();
    return rows;
  };

  it('gives the same result whatever the chunking', () => {
    const src = 'a,"b,1"\r\n"c\nd",e\nf,\n';
    const whole = collect([src]);
    const perChar = collect([...src]);
    expect(perChar).toEqual(whole);
    expect(whole).toEqual([['a', 'b,1'], ['c\nd', 'e'], ['f', '']]);
  });

  it('survives a split inside a quoted field', () => {
    expect(collect(['"ab', 'cd",e'])).toEqual([['abcd', 'e']]);
  });

  it('survives a split between the halves of a doubled quote', () => {
    expect(collect(['"a"', '"b"'])).toEqual([['a"b']]);
  });

  it('survives a split between CR and LF', () => {
    expect(collect(['a\r', '\nb'])).toEqual([['a'], ['b']]);
  });

  it('keeps a quoted CR and a later LF apart across a chunk boundary', () => {
    expect(collect(['"a\r', '""b\n', 'c"'])).toEqual([['a\n"b\nc']]);
  });

  it('stops when the sink asks it to', () => {
    const rows: string[][] = [];
    const p = new CsvRowParser(',', {
      row(f) {
        rows.push(f);
        return rows.length < 2;
      },
    });
    p.push('1\n2\n3\n4\n');
    p.end();
    expect(rows).toHaveLength(2);
    expect(p.stopped).toBe(true);
  });

  it('reports row indices to the sink', () => {
    const seen: number[] = [];
    const p = new CsvRowParser(',', { row: (_f, i) => void seen.push(i) });
    p.push('a\nb\nc');
    p.end();
    expect(seen).toEqual([0, 1, 2]);
  });
});

describe('encoding', () => {
  const utf16le = (s: string, bom = true): Uint8Array => {
    const text = (bom ? '\uFEFF' : '') + s;
    const out = new Uint8Array(text.length * 2);
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      out[i * 2] = c & 0xff;
      out[i * 2 + 1] = c >> 8;
    }
    return out;
  };
  const swap = (b: Uint8Array): Uint8Array => {
    const out = new Uint8Array(b.length);
    for (let i = 0; i + 1 < b.length; i += 2) {
      out[i] = b[i + 1]!;
      out[i + 1] = b[i]!;
    }
    return out;
  };

  it('strips a UTF-8 BOM', () => {
    const bytes = new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode('a,b')]);
    const r = parseDelimited(bytes);
    expect(r.encoding).toBe('utf-8');
    expect(r.rows).toEqual([['a', 'b']]);
  });

  it('strips a BOM that survived into a string source', () => {
    expect(parseDelimited('\uFEFFa,b').rows).toEqual([['a', 'b']]);
  });

  it('decodes UTF-16LE from its BOM', () => {
    const r = parseDelimited(utf16le('a,é\nb,c'));
    expect(r.encoding).toBe('utf-16le');
    expect(r.rows).toEqual([
      ['a', 'é'],
      ['b', 'c'],
    ]);
  });

  it('decodes UTF-16BE from its BOM', () => {
    const r = parseDelimited(swap(utf16le('a,é')));
    expect(r.encoding).toBe('utf-16be');
    expect(r.rows).toEqual([['a', 'é']]);
  });

  it('honours a forced encoding over a missing BOM', () => {
    const r = decodeCsvBytes(utf16le('hi', false), 'utf-16le');
    expect(r.text).toBe('hi');
  });

  it('decodes windows-1252 when asked', () => {
    const r = decodeCsvBytes(new Uint8Array([0x41, 0xe9, 0x80]), 'windows-1252');
    expect(r.text).toBe('Aé€');
  });

  it('decodes every windows-1252 byte from its own table', () => {
    // Not delegated to TextDecoder: the "windows-1252" label only resolves in
    // builds carrying full ICU data, so leaning on it would make a CSV either
    // open or throw depending on how Node was compiled.
    const bytes = new Uint8Array(256);
    for (let i = 0; i < 256; i++) bytes[i] = i;
    const text = decodeCsvBytes(bytes, 'windows-1252').text;
    expect(text).toHaveLength(256);
    // ASCII and the Latin-1 half are the identity.
    for (const i of [0x00, 0x41, 0x7f, 0xa0, 0xe9, 0xff]) {
      expect(text.codePointAt(i)).toBe(i);
    }
    // The 0x80-0x9F window is where the code page differs from Latin-1.
    const high = [
      0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160,
      0x2039, 0x0152, 0x008d, 0x017d, 0x008f, 0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022,
      0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
    ];
    expect([...text.slice(0x80, 0xa0)].map((c) => c.codePointAt(0))).toEqual(high);
  });

  it('decodes a buffer past the chunking threshold', () => {
    // Built one chunk at a time, so a file bigger than the argument limit does
    // not turn into a RangeError from String.fromCharCode.
    const pattern = [0x41, 0x80, 0xe9, 0x9d, 0xff];
    const bytes = new Uint8Array(200_000);
    for (let i = 0; i < bytes.length; i++) bytes[i] = pattern[i % pattern.length]!;
    const text = decodeCsvBytes(bytes, 'windows-1252').text;
    expect(text).toHaveLength(bytes.length);
    expect(text.slice(0, 5)).toBe('A\u20ac\u00e9\u009d\u00ff');
    expect(text.slice(-5)).toBe('A\u20ac\u00e9\u009d\u00ff');
  });

  it('agrees with the platform decoder where the platform has one', () => {
    let expected: string | undefined;
    try {
      expected = new TextDecoder('windows-1252').decode(new Uint8Array([...Array(256).keys()]));
    } catch {
      expected = undefined; // A build without full ICU: nothing to compare against.
    }
    if (expected !== undefined) {
      const bytes = new Uint8Array(256);
      for (let i = 0; i < 256; i++) bytes[i] = i;
      expect(decodeCsvBytes(bytes, 'windows-1252').text).toBe(expected);
    }
  });

  it('leaves a UTF-16 BOM alone when UTF-8 is forced', () => {
    // Nonsense in, nonsense out - but no crash, and the mark is not silently
    // eaten as though the file had been understood.
    const r = decodeCsvBytes(new Uint8Array([0xff, 0xfe, 0x41, 0x00]), 'utf-8');
    expect(r.encoding).toBe('utf-8');
    expect(r.text.length).toBeGreaterThan(0);
  });

  it('reports utf-8 for a plain string source', () => {
    expect(parseDelimited('a').encoding).toBe('utf-8');
  });
});

describe('delimiter detection', () => {
  const pick = (text: string) => detectDelimiter(text).delimiter;

  it('finds commas', () => {
    expect(pick('a,b,c\n1,2,3')).toBe(',');
  });

  it('finds semicolons', () => {
    expect(pick('a;b;c\n1;2;3')).toBe(';');
  });

  it('finds tabs', () => {
    expect(pick('a\tb\tc\n1\t2\t3')).toBe('\t');
  });

  it('finds pipes', () => {
    expect(pick('a|b|c\n1|2|3')).toBe('|');
  });

  it('prefers the delimiter that makes rows consistent', () => {
    // Commas are more numerous, but only the semicolons form a table.
    expect(pick('one, two, three;4\nalpha, beta;5')).toBe(';');
  });

  it('ignores delimiters inside quoted fields', () => {
    expect(pick('"a;b;c";x\n"d;e";y')).toBe(';');
  });

  it('falls back to a comma for a single-column file', () => {
    const d = detectDelimiter('alpha\nbeta\ngamma');
    expect(d.delimiter).toBe(',');
    expect(d.confident).toBe(false);
  });

  it('is not confident when two candidates tie', () => {
    const d = detectDelimiter('a,b\tc\n1,2\t3');
    expect(d.confident).toBe(false);
  });

  it('looks at a bounded prefix, however long the line', () => {
    // A file with no line terminator never reaches the row cap, so without a
    // character cap every candidate would scan the whole input - four passes
    // over an arbitrarily large string before a single row has been read.
    const text = 'x'.repeat(1 << 20) + ',,,,\n';
    const d = detectDelimiter(text);
    expect(d.scores.find((s) => s.delimiter === ',')!.fields).toBe(1);
  });

  it('samples only the first rows', () => {
    const head = 'a;b\n'.repeat(5);
    const tail = 'x,y,z\n'.repeat(50);
    expect(detectDelimiter(head + tail, { sampleRows: 5 }).delimiter).toBe(';');
  });

  it('warns when the choice was a coin toss', () => {
    const r = parseDelimited('a,b\tc\n1,2\t3');
    expect(r.warnings.map((w) => w.code)).toContain('delimiter-ambiguous');
  });

  it('does not warn about a one-column file', () => {
    expect(parseDelimited('alpha\nbeta').warnings).toEqual([]);
  });

  it('does not warn when one delimiter clearly wins', () => {
    expect(parseDelimited('a;b\n1;2').warnings).toEqual([]);
  });

  it('an explicit delimiter overrides detection', () => {
    expect(parseDelimited('a;b;c', { delimiter: ',' }).rows).toEqual([['a;b;c']]);
  });

  it('reports scores for every candidate', () => {
    const d = detectDelimiter('a,b\n1,2');
    expect(d.scores).toHaveLength(4);
    expect(d.scores.find((s) => s.delimiter === ',')).toMatchObject({
      fields: 2,
      consistency: 1,
      rows: 2,
    });
  });
});

describe('number inference', () => {
  it('reads integers', () => {
    expect(val('42')).toBe(42);
  });

  it('reads zero', () => {
    expect(val('0')).toBe(0);
  });

  it('reads negative integers', () => {
    expect(val('-17')).toBe(-17);
  });

  it('reads an explicit plus sign', () => {
    expect(val('+17')).toBe(17);
  });

  it('reads decimals', () => {
    expect(val('3.14159265358979')).toBe(3.14159265358979);
  });

  it('reads a bare fractional part', () => {
    expect(val('.5')).toBe(0.5);
  });

  it('reads a trailing decimal point', () => {
    expect(val('5.')).toBe(5);
  });

  it('reads scientific notation in both cases', () => {
    expect(val('1E-300')).toBe(1e-300);
    expect(val('1e+300')).toBe(1e300);
    expect(val('-2.5e3')).toBe(-2500);
  });

  it('ignores surrounding whitespace', () => {
    expect(val('  42\t')).toBe(42);
  });

  it('reads thousands separators', () => {
    expect(val('1,234')).toBe(1234);
    expect(val('1,234,567.25')).toBe(1234567.25);
  });

  it('rejects misplaced thousands separators', () => {
    expect(val('1,23')).toBe('1,23');
    expect(val('12,34,567')).toBe('12,34,567');
  });

  it('reads parenthesised negatives', () => {
    expect(val('(1,234)')).toBe(-1234);
    expect(val('(0.5)')).toBe(-0.5);
  });

  it('rejects a sign inside parentheses', () => {
    expect(val('(-5)')).toBe('(-5)');
  });

  it('rejects unbalanced parentheses', () => {
    expect(val('(5')).toBe('(5');
  });

  it('reads percentages as fractions and formats them', () => {
    expect(val('45.67%')).toBeCloseTo(0.4567, 12);
    expect(fmt('45.67%')).toBe('0.00%');
    expect(fmt('50%')).toBe('0%');
  });

  it('keeps the decimals a percentage was written with', () => {
    // "50.00%" is a measurement quoted to two places; a format of 0% would
    // redisplay it as "50%" and lose the precision the writer asserted.
    expect(fmt('50.00%')).toBe('0.00%');
    expect(fmt('12.50%')).toBe('0.00%');
    expect(fmt('12.5%')).toBe('0.0%');
    expect(fmt('0.000%')).toBe('0.000%');
    expect(val('50.00%')).toBe(0.5);
  });

  it("counts a percentage's decimals after applying its exponent", () => {
    expect(fmt('1e2%')).toBe('0%');
    expect(val('1e2%')).toBe(1);
  });

  it('reads a negative percentage', () => {
    expect(val('-10%')).toBeCloseTo(-0.1, 12);
  });

  it('keeps a leading zero as text', () => {
    expect(val('007')).toBe('007');
    expect(val('0123456')).toBe('0123456');
  });

  it('keeps a leading-zero decimal as a number', () => {
    expect(val('0.1')).toBe(0.1);
  });

  it('keeps a doubled zero as text', () => {
    expect(val('00')).toBe('00');
  });

  it('keeps a leading zero as text inside parentheses too', () => {
    expect(val('(007)')).toBe('(007)');
  });

  it('keeps more than fifteen significant digits as text', () => {
    expect(val('1234567890123456789')).toBe('1234567890123456789');
    expect(val('1.2345678901234567')).toBe('1.2345678901234567');
  });

  it('accepts exactly fifteen significant digits', () => {
    expect(val('123456789012345')).toBe(123456789012345);
  });

  it('does not count trailing zeros towards precision', () => {
    expect(val('1000000000000000000')).toBe(1e18);
  });

  it('keeps an overflowing exponent as text', () => {
    expect(val('1E+400')).toBe('1E+400');
  });

  it('keeps an underflowing exponent as text', () => {
    expect(val('1E-400')).toBe('1E-400');
  });

  it('does not read Infinity or NaN as numbers', () => {
    expect(val('Infinity')).toBe('Infinity');
    expect(val('NaN')).toBe('NaN');
    expect(val('-Infinity')).toBe('-Infinity');
  });

  it('does not read hex or octal literals', () => {
    expect(val('0x1F')).toBe('0x1F');
    expect(val('1_000')).toBe('1_000');
  });

  it('does not read a currency symbol', () => {
    expect(val('$1,234')).toBe('$1,234');
    expect(val('1234 EUR')).toBe('1234 EUR');
  });

  it('does not read a bare sign', () => {
    expect(val('-')).toBe('-');
    expect(val('+')).toBe('+');
  });

  it('normalises negative zero', () => {
    expect(Object.is(val('-0'), 0)).toBe(true);
  });

  it('can be switched off', () => {
    expect(val('42', { inferNumbers: false })).toBe('42');
  });
});

describe('date inference', () => {
  const serial = (y: number, m: number, d: number, h = 0, mi = 0, s = 0) =>
    partsToSerial(y, m, d, h, mi, s);

  it('reads an ISO date', () => {
    expect(val('2019-03-01')).toBe(43525);
    expect(fmt('2019-03-01')).toBe('yyyy-mm-dd');
  });

  it('reads an ISO date-time with a T separator', () => {
    expect(val('2023-03-15T13:37')).toBeCloseTo(serial(2023, 3, 15, 13, 37), 9);
    expect(fmt('2023-03-15T13:37')).toBe('yyyy-mm-dd hh:mm');
  });

  it('reads an ISO date-time with a space separator', () => {
    expect(val('2023-03-15 13:37:37')).toBeCloseTo(serial(2023, 3, 15, 13, 37, 37), 9);
    expect(fmt('2023-03-15 13:37:37')).toBe('yyyy-mm-dd hh:mm:ss');
  });

  it('reads fractional seconds', () => {
    const v = val('2023-03-15 13:37:37.5') as number;
    expect(v - serial(2023, 3, 15, 13, 37, 37)).toBeCloseTo(0.5 / 86400, 11);
  });

  it('rejects an impossible day', () => {
    expect(val('2023-02-30')).toBe('2023-02-30');
    expect(val('2023-13-01')).toBe('2023-13-01');
    expect(val('2023-00-10')).toBe('2023-00-10');
  });

  it('accepts the phantom 29 February 1900 that Excel carries', () => {
    // Excel inherited Lotus 1-2-3's belief that 1900 was a leap year, so serial
    // 60 is a date that never happened. Files written by Excel contain it, and
    // core renders serial 60 as 1900-02-29, so refusing it on the way in would
    // lose the value on a round trip.
    expect(val('1900-02-29')).toBe(60);
    expect(fmt('1900-02-29')).toBe('yyyy-mm-dd');
    expect(val('1900-02-28')).toBe(59);
    expect(val('1900-03-01')).toBe(61);
    expect(val('02/29/1900', { dateOrder: 'mdy' })).toBe(60);
    // Only the 1900 calendar has the bug, and 1900 predates the 1904 epoch.
    expect(val('1900-02-29', { dateSystem: 1904 })).toBe('1900-02-29');
    // No other February gains a 29th.
    expect(val('1901-02-29')).toBe('1901-02-29');
    expect(val('2100-02-29')).toBe('2100-02-29');
  });

  it('honours leap years', () => {
    expect(typeof val('2024-02-29')).toBe('number');
    expect(val('2023-02-29')).toBe('2023-02-29');
  });

  it('keeps dates before the epoch as text', () => {
    expect(val('1899-12-30')).toBe('1899-12-30');
  });

  it('keeps the day before serial 1 as text', () => {
    // Excel's calendar runs from 1/1/1900 (serial 1) to 31/12/9999. Serial 0 is
    // the fictitious "January 0, 1900", so 1899-12-31 has no serial to become:
    // importing it as 0 would render as a date nobody wrote.
    expect(val('1899-12-31')).toBe('1899-12-31');
    expect(val('1899-12-31 12:00')).toBe('1899-12-31 12:00');
    expect(val('1900-01-01')).toBe(1);
  });

  it('holds the last date Excel can represent', () => {
    expect(val('9999-12-31')).toBe(2958465);
  });

  it('starts the 1904 system at its own epoch', () => {
    expect(val('1904-01-01', { dateSystem: 1904 })).toBe(0);
    expect(val('1903-12-31', { dateSystem: 1904 })).toBe('1903-12-31');
  });

  it('keeps a zoned timestamp as text', () => {
    // A spreadsheet serial has no timezone; converting one would silently move
    // the wall-clock time the user typed.
    expect(val('2023-03-15T13:37:00Z')).toBe('2023-03-15T13:37:00Z');
    expect(val('2023-03-15T13:37:00+01:00')).toBe('2023-03-15T13:37:00+01:00');
  });

  it('leaves ambiguous forms as text by default', () => {
    expect(val('1/2/2024')).toBe('1/2/2024');
    expect(val('01/01/1900')).toBe('01/01/1900');
    expect(val('15.03.2023')).toBe('15.03.2023');
  });

  it('reads day-first order when asked', () => {
    expect(val('1/2/2024', { dateOrder: 'dmy' })).toBe(serial(2024, 2, 1));
  });

  it('reads month-first order when asked', () => {
    expect(val('1/2/2024', { dateOrder: 'mdy' })).toBe(serial(2024, 1, 2));
  });

  it('reads year-first order when asked', () => {
    expect(val('2024/1/2', { dateOrder: 'ymd' })).toBe(serial(2024, 1, 2));
  });

  it('accepts dot and dash separators for ambiguous dates', () => {
    expect(val('15.03.2023', { dateOrder: 'dmy' })).toBe(serial(2023, 3, 15));
    expect(val('15-03-2023', { dateOrder: 'dmy' })).toBe(serial(2023, 3, 15));
  });

  it('requires the separators to match', () => {
    expect(val('15-03.2023', { dateOrder: 'dmy' })).toBe('15-03.2023');
  });

  it('maps two-digit years the way Excel does', () => {
    expect(val('1/2/29', { dateOrder: 'dmy' })).toBe(serial(2029, 2, 1));
    expect(val('1/2/30', { dateOrder: 'dmy' })).toBe(serial(1930, 2, 1));
  });

  it('rejects a three-digit year', () => {
    expect(val('1/2/999', { dateOrder: 'dmy' })).toBe('1/2/999');
  });

  it('never reads a two-component value as a date', () => {
    // "1-2" is the canonical Excel casualty: it becomes 2 January.
    expect(val('1-2', { dateOrder: 'dmy' })).toBe('1-2');
    expect(val('3/4', { dateOrder: 'mdy' })).toBe('3/4');
  });

  it('gives ISO precedence over the stated order', () => {
    expect(val('2024-01-02', { dateOrder: 'dmy' })).toBe(serial(2024, 1, 2));
  });

  it('can be switched off', () => {
    expect(val('2019-03-01', { inferDates: false })).toBe('2019-03-01');
  });

  it('honours the 1904 date system', () => {
    expect(val('2019-03-01', { dateSystem: 1904 })).toBe(43525 - 1462);
  });

  it('keeps a 1900-system date that predates the 1904 epoch as text', () => {
    expect(val('1901-01-01', { dateSystem: 1904 })).toBe('1901-01-01');
  });
});

describe('time inference', () => {
  it('leaves bare times as text by default', () => {
    expect(val('13:37')).toBe('13:37');
    expect(val('1:30')).toBe('1:30');
  });

  it('reads times when asked', () => {
    expect(val('13:37', { inferTimes: true })).toBeCloseTo((13 * 60 + 37) / 1440, 12);
    expect(fmt('13:37', { inferTimes: true })).toBe('hh:mm');
  });

  it('reads seconds and their fraction', () => {
    expect(val('13:37:37.92', { inferTimes: true })).toBeCloseTo(
      (13 * 3600 + 37 * 60 + 37.92) / 86400,
      12,
    );
    expect(fmt('13:37:37', { inferTimes: true })).toBe('hh:mm:ss');
  });

  it('keeps an elapsed duration as text', () => {
    expect(val('36:00:00', { inferTimes: true })).toBe('36:00:00');
  });

  it('rejects out-of-range components', () => {
    expect(val('12:60', { inferTimes: true })).toBe('12:60');
    expect(val('12:30:61', { inferTimes: true })).toBe('12:30:61');
  });

  it('reads midnight as zero', () => {
    expect(val('00:00', { inferTimes: true })).toBe(0);
  });
});

describe('boolean and text inference', () => {
  it('reads TRUE and FALSE', () => {
    expect(val('TRUE')).toBe(true);
    expect(val('FALSE')).toBe(false);
  });

  it('is case-insensitive', () => {
    expect(val('true')).toBe(true);
    expect(val('False')).toBe(false);
  });

  it('does not read other truthy words', () => {
    expect(val('yes')).toBe('yes');
    expect(val('T')).toBe('T');
    expect(val('1')).toBe(1);
  });

  it('can be switched off', () => {
    expect(val('TRUE', { inferBooleans: false })).toBe('TRUE');
  });

  it('reads an empty field as an empty cell', () => {
    expect(val('')).toBeNull();
  });

  it('keeps a whitespace-only field as text, verbatim', () => {
    expect(val('   ')).toBe('   ');
  });

  it('keeps unrecognised text verbatim, spaces included', () => {
    expect(val('  hello  ')).toBe('  hello  ');
  });
});

describe('false-positive guards', () => {
  it('recognises the gene names spreadsheets are known to eat', () => {
    expect(looksLikeGeneName('SEPT1')).toBe(true);
    expect(looksLikeGeneName('MARCH1')).toBe(true);
    expect(looksLikeGeneName('DEC1')).toBe(true);
    expect(looksLikeGeneName('MAR-1')).toBe(true);
    expect(looksLikeGeneName('sept2')).toBe(true);
  });

  it('does not claim ordinary words', () => {
    expect(looksLikeGeneName('SEPTEMBER')).toBe(false);
    expect(looksLikeGeneName('DECIMAL')).toBe(false);
    expect(looksLikeGeneName('SEPT')).toBe(false);
  });

  it('keeps gene names as text whatever the options', () => {
    expect(val('SEPT1', { dateOrder: 'mdy' })).toBe('SEPT1');
    expect(val('MARCH1', { dateOrder: 'dmy' })).toBe('MARCH1');
  });

  it('strips a leading apostrophe and keeps the rest as text', () => {
    expect(val("'007")).toBe('007');
    expect(val("'42")).toBe('42');
    expect(val("'2019-03-01")).toBe('2019-03-01');
  });

  it('can keep the apostrophe when asked', () => {
    expect(val("'42", { stripLeadingApostrophe: false })).toBe("'42");
  });

  it('raw mode imports everything as text', () => {
    const o = { raw: true } as const;
    expect(val('42', o)).toBe('42');
    expect(val('2019-03-01', o)).toBe('2019-03-01');
    expect(val('TRUE', o)).toBe('TRUE');
    expect(val("'007", o)).toBe("'007");
    expect(val('', o)).toBeNull();
  });
});

describe('warnings', () => {
  it('reports a ragged row against the modal width', () => {
    const r = parseDelimited('a,b,c\n1,2,3\n4,5\n6,7,8', { delimiter: ',' });
    const ragged = r.warnings.filter((w) => w.code === 'ragged-row');
    expect(ragged).toHaveLength(1);
    expect(ragged[0]!.row).toBe(2);
    expect(ragged[0]!.message).toContain('expected 3');
  });

  it('treats a title line above a table as the anomaly', () => {
    const r = parseDelimited('report\na,b\n1,2\n3,4', { delimiter: ',' });
    expect(r.warnings.filter((w) => w.code === 'ragged-row').map((w) => w.row)).toEqual([0]);
  });

  it('says nothing about a rectangular file', () => {
    expect(parseDelimited('a,b\n1,2', { delimiter: ',' }).warnings).toEqual([]);
  });

  it('ignores blank lines when judging raggedness', () => {
    const r = parseDelimited('a,b\n\n1,2\n', { delimiter: ',' });
    expect(r.warnings.filter((w) => w.code === 'ragged-row')).toHaveLength(0);
  });

  it('caps the warning list and says how many it dropped', () => {
    const src = 'a,b,c\n'.repeat(20) + '1,2\n'.repeat(20);
    const r = parseDelimited(src, { delimiter: ',', maxWarnings: 5 });
    expect(r.warnings.filter((w) => w.code === 'ragged-row')).toHaveLength(5);
    const limit = r.warnings.find((w) => w.code === 'warning-limit');
    expect(limit?.message).toContain('15 further');
  });

  it('never throws on malformed input', () => {
    expect(() => parseDelimited('"unclosed,\n\n""",,,\r\r\n')).not.toThrow();
  });
});

describe('importing into a sheet', () => {
  it('populates cells row-major from A1', () => {
    const { sheet } = readCsv('a,b\n1,2');
    expect(sheet.getValue(0, 0)).toBe('a');
    expect(sheet.getValue(0, 1)).toBe('b');
    expect(sheet.getValue(1, 0)).toBe(1);
    expect(sheet.getValue(1, 1)).toBe(2);
  });

  it('leaves empty fields as empty cells', () => {
    const { sheet } = readCsv('a,,c');
    expect(sheet.getCell(0, 1)).toBeUndefined();
    expect(sheet.cellCount).toBe(2);
  });

  it('reports the imported extent', () => {
    const r = readCsv('a,b,c\n1,2,3\n4,5,6');
    expect(r.rowCount).toBe(3);
    expect(r.colCount).toBe(3);
  });

  it('attaches a number format to inferred dates', () => {
    const r = readCsv('when\n2019-03-01');
    const style = r.workbook.styles.get(r.sheet.getStyle(1, 0));
    expect(style.numFmt).toBe('yyyy-mm-dd');
    expect(r.sheet.getValue(1, 0)).toBe(43525);
  });

  it('interns one style per distinct format', () => {
    const before = new Workbook().styles.size;
    const r = readCsv('2019-03-01\n2019-03-02\n2019-03-03');
    expect(r.workbook.styles.size).toBe(before + 1);
  });

  it('keeps header cells as text', () => {
    const r = readCsv('2019,007,TRUE\n1,2,3', { headerRow: true });
    expect(r.header).toEqual(['2019', '007', 'TRUE']);
    expect(r.sheet.getValue(0, 0)).toBe('2019');
    expect(r.sheet.getValue(1, 0)).toBe(1);
  });

  it('names the sheet as asked', () => {
    expect(readCsv('a', { sheetName: 'Import' }).sheet.name).toBe('Import');
  });

  it('carries the detected delimiter and encoding through', () => {
    const r = readCsv('a;b\n1;2');
    expect(r.delimiter).toBe(';');
    expect(r.encoding).toBe('utf-8');
  });

  it('surfaces parse warnings on the result', () => {
    const r = readCsv('a,b,c\n1,2\n3,4,5', { delimiter: ',' });
    expect(r.warnings.map((w) => w.code)).toContain('ragged-row');
  });
});

describe('writer quoting', () => {
  const line = (row: string[], delimiter = ',') =>
    writeRows([row], { delimiter, lineEnding: '\n', trailingNewline: false });

  it('leaves plain fields bare', () => {
    expect(line(['a', 'b'])).toBe('a,b');
  });

  it('quotes a field containing the delimiter', () => {
    expect(line(['a,b', 'c'])).toBe('"a,b",c');
  });

  it('only quotes for the delimiter actually in use', () => {
    expect(line(['a,b', 'c'], ';')).toBe('a,b;c');
    expect(line(['a;b', 'c'], ';')).toBe('"a;b";c');
  });

  it('quotes and doubles an embedded quote', () => {
    expect(line(['he said "hi"'])).toBe('"he said ""hi"""');
  });

  it('quotes a field containing a newline', () => {
    expect(line(['a\nb'])).toBe('"a\nb"');
    expect(line(['a\rb'])).toBe('"a\rb"');
  });

  it('quotes leading and trailing whitespace', () => {
    expect(line([' a'])).toBe('" a"');
    expect(line(['a '])).toBe('"a "');
    expect(line(['a b'])).toBe('a b');
  });

  it('leaves an empty field unquoted', () => {
    expect(line(['', 'a'])).toBe(',a');
  });

  it('quotes everything when told to', () => {
    expect(writeRows([['a', 'b']], { quoteAll: true, trailingNewline: false })).toBe('"a","b"');
  });

  it('exposes the quoting predicate', () => {
    expect(needsQuoting('plain', ',')).toBe(false);
    expect(needsQuoting('a,b', ',')).toBe(true);
    expect(needsQuoting('a"b', ',')).toBe(true);
    expect(needsQuoting(' a', ',')).toBe(true);
    expect(needsQuoting('', ',')).toBe(false);
  });

  it('defaults to CRLF and a trailing terminator', () => {
    expect(writeRows([['a'], ['b']])).toBe('a\r\nb\r\n');
  });

  it('can end without a terminator', () => {
    expect(writeRows([['a'], ['b']], { trailingNewline: false })).toBe('a\r\nb');
  });

  it('can emit a BOM', () => {
    expect(writeRows([['a']], { bom: true, lineEnding: '\n' })).toBe('\uFEFFa\n');
  });

  it('writes nothing for no rows', () => {
    expect(writeRows([])).toBe('');
  });
});

describe('writing a sheet', () => {
  const opts = { lineEnding: '\n' } as const;

  it('writes the used range', () => {
    const { sheet } = sheetOf([
      ['a', 'b'],
      ['c', 'd'],
    ]);
    expect(writeCsv(sheet, opts)).toBe('a,b\nc,d\n');
  });

  it('pads short rows to the used width', () => {
    const { sheet } = sheetOf([['a', 'b', 'c'], ['d']]);
    expect(writeCsv(sheet, opts)).toBe('a,b,c\nd,,\n');
  });

  it('anchors the export at A1, keeping leading blanks', () => {
    const wb = new Workbook();
    const sheet = wb.addSheet('S');
    sheet.setValue(1, 1, 'x');
    expect(writeCsv(sheet, opts)).toBe(',\n,x\n');
  });

  it('round-trips a sheet whose first column is empty', () => {
    const source = ',a\n,b\n';
    const { sheet } = readCsv(source, { raw: true, delimiter: ',' });
    expect(writeCsv(sheet, opts)).toBe(source);
  });

  it('writes an empty sheet as nothing', () => {
    const { sheet } = sheetOf([]);
    expect(writeCsv(sheet, opts)).toBe('');
  });

  it('renders booleans as Excel does', () => {
    const { sheet } = sheetOf([[true, false]]);
    expect(writeCsv(sheet, opts)).toBe('TRUE,FALSE\n');
  });

  it('renders numbers with an upper-case exponent', () => {
    const { sheet } = sheetOf([[1e300, 1e-300, 0.1]]);
    expect(writeCsv(sheet, opts)).toBe('1E+300,1E-300,0.1\n');
  });

  it('renders numbers to Excel\'s fifteen significant digits', () => {
    // The shortest round-tripping form of a computed double runs to seventeen
    // digits, which the reader on the other side of this file refuses to take
    // as a number - a value would leave the sheet numeric and come back text.
    expect(formatScalar(0.1 + 0.2)).toBe('0.3');
    expect(formatScalar(1 / 3)).toBe('0.333333333333333');
    expect(formatScalar(2 / 3)).toBe('0.666666666666667');
    expect(formatScalar(0.1)).toBe('0.1');
    expect(formatScalar(-0)).toBe('0');
    expect(formatScalar(43525)).toBe('43525');
  });

  it('writes numbers its own reader reads back unchanged', () => {
    const values = [0.1 + 0.2, 1 / 3, 2 / 3, 1e300, 1e-300, 1e21, 43525, 0.1, -17, 1e-7];
    for (const v of values) {
      const written = formatScalar(v);
      const read = inferValue(written).value;
      expect(typeof read, `${written} should read back as a number`).toBe('number');
      expect(formatScalar(read as number)).toBe(written);
    }
  });

  it('renders errors as their codes', () => {
    const { sheet } = sheetOf([]);
    sheet.setFormula(0, 0, '1/0', CellError.DIV0);
    sheet.setValue(0, 1, 'ok');
    expect(writeCsv(sheet, opts)).toBe('#DIV/0!,ok\n');
  });

  it('renders an empty cell as an empty field', () => {
    expect(formatScalar(null)).toBe('');
    expect(formatScalar(CellError.NA)).toBe('#N/A');
  });

  it('restricts output to a given range', () => {
    const { sheet } = sheetOf([
      ['a', 'b', 'c'],
      ['d', 'e', 'f'],
    ]);
    expect(writeCsv(sheet, { ...opts, range: { minRow: 0, minCol: 1, maxRow: 1, maxCol: 2 } })).toBe(
      'b,c\ne,f\n',
    );
  });

  it('clamps a range to the sheet limits instead of aliasing', () => {
    // Addresses pack as row * MAX_COLS + col, so a column past the last one is
    // the next row's territory: an over-wide range must stop, not wrap.
    const wb = new Workbook();
    const sheet = wb.addSheet('S');
    sheet.setValue(0, 0, 'a');
    sheet.setValue(1, 3, 'next row');
    const text = writeCsv(sheet, {
      ...opts,
      range: { minRow: 0, minCol: 0, maxRow: 0, maxCol: MAX_COLS + 5 },
    });
    expect(text).not.toContain('next row');
    expect(text.split('\n')[0]!.split(',')).toHaveLength(MAX_COLS);
  });

  it('uses a formatter callback when one is given', () => {
    const { wb, sheet } = sheetOf([[43525]]);
    sheet.setStyle(0, 0, wb.styles.intern({ numFmt: 'yyyy-mm-dd' }));
    const text = writeCsv(sheet, {
      ...opts,
      styles: wb.styles,
      format: (value, numFmt) => (numFmt === 'yyyy-mm-dd' ? `date:${String(value)}` : undefined),
    });
    expect(text).toBe('date:43525\n');
  });

  it('falls back to the default rendering when the formatter declines', () => {
    const { wb, sheet } = sheetOf([[1, 'x']]);
    const text = writeCsv(sheet, {
      ...opts,
      styles: wb.styles,
      format: (value) => (typeof value === 'number' ? 'N' : undefined),
    });
    expect(text).toBe('N,x\n');
  });

  it('passes the cell position to the formatter', () => {
    const { sheet } = sheetOf([['a', 'b']]);
    const seen: string[] = [];
    writeCsv(sheet, {
      ...opts,
      format: (_v, _f, ctx) => {
        seen.push(`${ctx.row},${ctx.col}`);
        return undefined;
      },
    });
    expect(seen).toEqual(['0,0', '0,1']);
  });

  it('encodes to UTF-8 bytes, BOM included', () => {
    const { sheet } = sheetOf([['é']]);
    const bytes = writeCsvBytes(sheet, { ...opts, bom: true });
    expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes)).toBe('\uFEFFé\n');
  });
});

describe('injection sanitising', () => {
  const write = (text: string, sanitise: boolean) => {
    const { sheet } = sheetOf([[text]]);
    return writeCsv(sheet, { lineEnding: '\n', trailingNewline: false, sanitise });
  };

  it('leaves formula-shaped text alone by default', () => {
    expect(write('=1+1', false)).toBe('=1+1');
  });

  it('prefixes the risky leading characters when switched on', () => {
    for (const prefix of ['=', '+', '-', '@']) {
      expect(write(`${prefix}cmd`, true)).toBe(`'${prefix}cmd`);
    }
  });

  it('prefixes leading tab and CR too', () => {
    expect(write('\tx', true)).toBe("'\tx");
  });

  it('leaves innocent text alone', () => {
    expect(write('hello', true)).toBe('hello');
    expect(write('a=b', true)).toBe('a=b');
  });

  it('never touches numbers, so negatives survive', () => {
    const { sheet } = sheetOf([[-5]]);
    expect(writeCsv(sheet, { lineEnding: '\n', trailingNewline: false, sanitise: true })).toBe('-5');
  });

  it('round-trips through the reader, which strips the apostrophe', () => {
    const { sheet } = sheetOf([['=SUM(A1)']]);
    const text = writeCsv(sheet, { lineEnding: '\n', sanitise: true });
    expect(readCsv(text).sheet.getValue(0, 0)).toBe('=SUM(A1)');
  });
});

describe('package surface', () => {
  it('is reachable from the package entry point', async () => {
    // The module is useless to a consumer of @mirrorz/formats until index.ts
    // re-exports it, and nothing else in the package imports it.
    const index = (await import('../src/index.js')) as Record<string, unknown>;
    for (const name of ['readCsv', 'writeCsv', 'parseDelimited', 'inferValue', 'needsQuoting']) {
      expect(typeof index[name], name).toBe('function');
    }
  });
});

describe('fixtures', () => {
  const CSVS = [
    'basic-types.csv',
    'formulas.csv',
    'features.csv',
    'styling.csv',
    'edge-cases.csv',
    'precedence.csv',
  ];

  it.each(CSVS)('%s parses with a comma delimiter', (name) => {
    const r = parseDelimited(fixtureBytes(name));
    expect(r.delimiter).toBe(',');
    expect(r.rows.length).toBeGreaterThan(0);
  });

  it.each(CSVS)('%s round-trips byte for byte in raw mode', (name) => {
    const original = fixtureText(name);
    const { sheet } = readCsv(original, { raw: true, delimiter: ',' });
    expect(writeCsv(sheet, { lineEnding: '\n' })).toBe(original);
  });

  it('reads the basic-types sheet the way the fixture demands', () => {
    const { sheet } = readCsv(fixtureBytes('basic-types.csv'));
    const byLabel = new Map<string, unknown>();
    for (let r = 1; r < 40; r++) {
      const label = sheet.getValue(r, 0);
      if (typeof label === 'string') byLabel.set(label, sheet.getValue(r, 1));
    }
    expect(byLabel.get('integer')).toBe(42);
    expect(byLabel.get('negative')).toBe(-17);
    expect(byLabel.get('zero')).toBe(0);
    expect(byLabel.get('float')).toBe(3.14159265358979);
    expect(byLabel.get('tiny')).toBe(1e-300);
    expect(byLabel.get('huge')).toBe(1e300);
    expect(byLabel.get('string')).toBe('hello world');
    expect(byLabel.get('quote')).toBe('he said "hi"');
    expect(byLabel.get('bool_true')).toBe(true);
    expect(byLabel.get('bool_false')).toBe(false);
    expect(byLabel.get('empty')).toBeNull();
    expect(byLabel.get('leading_zero')).toBe('007');
    expect(byLabel.get('looks_like_date')).toBe('1-2');
    expect(byLabel.get('gene_name')).toBe('SEPT1');
    expect(byLabel.get('long_number')).toBe('1234567890123456789');
  });

  it('keeps the fixture unicode intact', () => {
    const { sheet } = readCsv(fixtureBytes('basic-types.csv'));
    const row = fixtureText('basic-types.csv')
      .split('\n')
      .findIndex((l) => l.startsWith('unicode,'));
    expect(sheet.getValue(row, 1)).toBe('éàü 你好 😀');
    expect(sheet.getValue(row, 2)).toBe('accents, CJK, emoji');
  });

  it('leaves the fixture locale dates as text by default', () => {
    const { sheet } = readCsv(fixtureBytes('basic-types.csv'));
    const lines = fixtureText('basic-types.csv').split('\n');
    const row = lines.findIndex((l) => l.startsWith('date serial 1,'));
    expect(sheet.getValue(row, 1)).toBe('01/01/1900');
  });

  it('reads them as dates once the order is stated', () => {
    const { sheet } = readCsv(fixtureBytes('basic-types.csv'), { dateOrder: 'mdy' });
    const lines = fixtureText('basic-types.csv').split('\n');
    expect(sheet.getValue(lines.indexOf('date serial 1,01/01/1900,'), 1)).toBe(1);
    // Excel's phantom 29 February 1900 puts 1 March on serial 61.
    expect(sheet.getValue(lines.indexOf('leap bug serial 60,03/01/1900,'), 1)).toBe(61);
  });

  it('reads the ISO date column in formulas.csv', () => {
    const r = readCsv(fixtureBytes('formulas.csv'), { headerRow: true });
    expect(r.header).toEqual(['n', 'name', 'dept', 'salary', 'hired']);
    expect(r.sheet.getValue(1, 4)).toBe(partsToSerial(2019, 3, 1));
    expect(r.workbook.styles.get(r.sheet.getStyle(1, 4)).numFmt).toBe('yyyy-mm-dd');
    expect(r.sheet.getValue(1, 3)).toBe(165000);
  });

  it('handles the number-format column of styling.csv without mangling it', () => {
    const { sheet } = readCsv(fixtureBytes('styling.csv'));
    const lines = fixtureText('styling.csv').split('\n');
    const row = lines.findIndex((l) => l.includes('#,##0.00'));
    expect(sheet.getValue(row, 4)).toBe(1234567.891);
    expect(sheet.getValue(row, 5)).toBe('#,##0.00');
  });

  it('reads the accounting format string with its escaped quotes intact', () => {
    const { sheet } = readCsv(fixtureBytes('styling.csv'));
    const lines = fixtureText('styling.csv').split('\n');
    const row = lines.findIndex((l) => l.includes(',accounting,'));
    expect(sheet.getValue(row, 5)).toContain('_("$"* #,##0.00_)');
  });

  it('imports features.csv with its trailing empty column', () => {
    const r = readCsv(fixtureBytes('features.csv'), { headerRow: true });
    expect(r.header).toEqual(['region', 'qty', 'score', '', 'https://example.com']);
    expect(r.sheet.getValue(1, 1)).toBe(12);
    expect(r.sheet.getValue(1, 2)).toBe(0.91);
    expect(r.warnings).toEqual([]);
  });

  it('re-exports an inferred import with a formatter, unchanged', () => {
    const source = 'a,b\n1,2\n3,4\n';
    const r = readCsv(source);
    expect(writeCsv(r.sheet, { lineEnding: '\n', styles: r.workbook.styles })).toBe(source);
  });
});
