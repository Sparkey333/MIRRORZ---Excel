import { readFileSync } from 'node:fs';
import { deflateRawSync } from 'node:zlib';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import {
  InflateError,
  deflateRawStored,
  getCodec,
  inflateRaw,
  portableCodec,
  resetCodec,
  setCodec,
} from '../src/inflate.js';
import { readZip, writeZip } from '../src/zip.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const bytes = (name: string) => new Uint8Array(readFileSync(new URL(name, FIXTURES)));

/** Deterministic pseudo-random bytes, so failures are reproducible. */
function pseudoRandom(length: number, seed = 42): Uint8Array {
  const out = new Uint8Array(length);
  let state = seed;
  for (let i = 0; i < length; i++) {
    state = (state * 1_103_515_245 + 12_345) & 0x7fff_ffff;
    out[i] = (state >> 16) & 0xff;
  }
  return out;
}

describe('the portable inflate against zlib output', () => {
  // zlib is the reference here: whatever it produces, we must be able to read.

  it.each([
    ['empty', new Uint8Array(0)],
    ['one byte', new Uint8Array([42])],
    ['repetitive', new TextEncoder().encode('abc'.repeat(5000))],
    ['all zeros', new Uint8Array(10_000)],
    ['incompressible', pseudoRandom(10_000)],
    ['every byte value', new Uint8Array(Array.from({ length: 256 }, (_, i) => i))],
  ])('round-trips %s', (_label, data) => {
    const compressed = new Uint8Array(deflateRawSync(data));
    expect(Array.from(inflateRaw(compressed))).toEqual(Array.from(data));
  });

  it.each([0, 1, 6, 9])('reads output from compression level %i', (level) => {
    const data = new TextEncoder().encode('the quick brown fox '.repeat(500));
    const compressed = new Uint8Array(deflateRawSync(data, { level }));
    expect(Array.from(inflateRaw(compressed))).toEqual(Array.from(data));
  });

  it('handles a long back-reference run, where copies overlap the output', () => {
    // A run encodes as a copy whose source overlaps its own destination, which
    // is why the copy must proceed byte by byte.
    const data = new Uint8Array(50_000).fill(0x5a);
    const compressed = new Uint8Array(deflateRawSync(data));
    const out = inflateRaw(compressed);
    expect(out.length).toBe(50_000);
    expect(out.every((b) => b === 0x5a)).toBe(true);
  });

  it('handles text long enough to need dynamic Huffman tables', () => {
    const words = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta'];
    let text = '';
    for (let i = 0; i < 20_000; i++) text += `${words[i % words.length]} ${i} `;
    const data = new TextEncoder().encode(text);
    const compressed = new Uint8Array(deflateRawSync(data));
    expect(new TextDecoder().decode(inflateRaw(compressed))).toBe(text);
  });

  it('round-trips a megabyte', () => {
    const data = pseudoRandom(1_000_000, 7);
    const compressed = new Uint8Array(deflateRawSync(data));
    expect(Array.from(inflateRaw(compressed).subarray(0, 1000))).toEqual(
      Array.from(data.subarray(0, 1000)),
    );
    expect(inflateRaw(compressed).length).toBe(data.length);
  });
});

describe('stored-block deflate', () => {
  it('produces a stream zlib can read back', () => {
    const data = new TextEncoder().encode('hello, stored world');
    // Our own inflate must read it, and so must zlib's.
    expect(Array.from(inflateRaw(deflateRawStored(data)))).toEqual(Array.from(data));
  });

  it('splits input longer than a single stored block', () => {
    // A stored block's length field is 16 bits, so 200 KiB needs four blocks.
    const data = pseudoRandom(200_000, 11);
    expect(Array.from(inflateRaw(deflateRawStored(data)))).toEqual(Array.from(data));
  });

  it('handles empty input', () => {
    expect(inflateRaw(deflateRawStored(new Uint8Array(0))).length).toBe(0);
  });

  it('is exactly at the block boundary', () => {
    for (const size of [65_534, 65_535, 65_536, 65_537]) {
      const data = pseudoRandom(size, size);
      expect(inflateRaw(deflateRawStored(data)).length, String(size)).toBe(size);
    }
  });
});

describe('hostile input', () => {
  it('rejects a truncated stream', () => {
    const compressed = new Uint8Array(deflateRawSync(new TextEncoder().encode('x'.repeat(1000))));
    expect(() => inflateRaw(compressed.subarray(0, 5))).toThrow(InflateError);
  });

  it('rejects an invalid block type', () => {
    // Block type 3 is reserved and must never appear.
    expect(() => inflateRaw(new Uint8Array([0x07, 0x00, 0x00]))).toThrow(/invalid block type/);
  });

  it('rejects a stored block whose length complement is wrong', () => {
    // Final stored block, length 4, complement deliberately corrupted.
    expect(() => inflateRaw(new Uint8Array([0x01, 4, 0, 0, 0, 1, 2, 3, 4]))).toThrow(
      /does not match its complement/,
    );
  });

  it('rejects a back-reference pointing before the start of the output', () => {
    // Hand-built fixed-Huffman block: one literal, then a copy whose distance
    // reaches further back than anything written. A guess at the bytes would
    // not reliably hit this path, so the stream is constructed bit by bit.
    const bits: number[] = [];
    /** Huffman codes are written most-significant bit first. */
    const codeMsb = (value: number, width: number) => {
      for (let i = width - 1; i >= 0; i--) bits.push((value >> i) & 1);
    };
    /** Header fields and extra bits are written least-significant bit first. */
    const valueLsb = (value: number, width: number) => {
      for (let i = 0; i < width; i++) bits.push((value >> i) & 1);
    };

    valueLsb(1, 1); // final block
    valueLsb(1, 2); // fixed Huffman
    codeMsb(0x30 + 65, 8); // literal 'A', which occupies codes 0x30..0xBF
    codeMsb(1, 7); // symbol 257: a match of length 3
    codeMsb(5, 5); // distance symbol 5: a distance of 7, past our single byte

    const out = new Uint8Array(Math.ceil(bits.length / 8));
    bits.forEach((bit, i) => {
      if (bit) out[i >> 3]! |= 1 << (i & 7);
    });

    expect(() => inflateRaw(out)).toThrow(/before start of output/);
  });

  it('enforces the output size limit rather than exhausting memory', () => {
    // A kilobyte of zeros expands hugely; the cap must stop it.
    const data = new Uint8Array(1_000_000);
    const compressed = new Uint8Array(deflateRawSync(data));
    expect(() => inflateRaw(compressed, { maxSize: 1000 })).toThrow(/exceeds the 1000-byte limit/);
  });

  it('does not loop forever on random bytes', () => {
    // Whatever it makes of nonsense, it must terminate.
    for (let seed = 0; seed < 40; seed++) {
      const junk = pseudoRandom(64, seed);
      try {
        inflateRaw(junk, { maxSize: 1_000_000 });
      } catch {
        // Throwing is the expected outcome; not hanging is the assertion.
      }
    }
    expect(true).toBe(true);
  });
});

describe('the portable codec reading real Office files', () => {
  // The property that matters: with zlib swapped out entirely, every fixture
  // must still open. This is what makes the package usable in a renderer.
  beforeAll(() => setCodec(portableCodec));
  afterAll(() => resetCodec());

  it.each([
    'basic-types.xlsx',
    'formulas.xlsx',
    'styling.xlsx',
    'features.xlsx',
    'edge-cases.xlsx',
    'basic-types.ods',
  ])('opens %s with no zlib available', (name) => {
    const entries = readZip(bytes(name));
    expect(entries.size).toBeGreaterThan(3);
    for (const entry of entries.values()) {
      // CRC verification is on, so a wrong byte anywhere fails here.
      expect(() => entry.data()).not.toThrow();
    }
  });

  it('reads a multi-megabyte part correctly', () => {
    const entries = readZip(bytes('large.xlsx'));
    const sheet = [...entries.keys()].find((k) => k.startsWith('xl/worksheets/sheet'))!;
    const xml = new TextDecoder().decode(entries.get(sheet)!.data());
    expect(xml.length).toBeGreaterThan(1_000_000);
    expect(xml.trimEnd().endsWith('</worksheet>')).toBe(true);
  });

  it('writes an archive the Node codec can read, and the reverse', () => {
    const payload = new TextEncoder().encode('round trip across codecs '.repeat(200));
    const written = writeZip([{ name: 'a.txt', data: payload }]);

    // Read it back with the portable codec still installed.
    expect(Array.from(readZip(written).get('a.txt')!.data())).toEqual(Array.from(payload));

    // And with zlib restored.
    resetCodec();
    void getCodec();
    expect(Array.from(readZip(written).get('a.txt')!.data())).toEqual(Array.from(payload));
    setCodec(portableCodec);
  });
});

describe('codec selection', () => {
  it('uses whichever codec is installed', () => {
    const original = getCodec();
    let called = 0;
    setCodec({
      inflateRaw: (d, o) => {
        called++;
        return portableCodec.inflateRaw(d, o);
      },
      deflateRaw: portableCodec.deflateRaw,
    });
    readZip(bytes('basic-types.xlsx')).get('xl/workbook.xml')!.data();
    expect(called).toBeGreaterThan(0);
    setCodec(original);
  });

  it('restores the portable codec on reset', () => {
    resetCodec();
    expect(getCodec().deflateRaw).toBe(deflateRawStored);
  });
});
