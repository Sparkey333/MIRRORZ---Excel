import { describe, expect, it } from 'vitest';
import { VbaCompressionError, compress, decompress } from '../src/compression.js';

const hex = (text: string): Uint8Array =>
  new Uint8Array(text.trim().split(/\s+/).map((byte) => Number.parseInt(byte, 16)));

const ascii = (text: string): Uint8Array => new Uint8Array([...text].map((c) => c.charCodeAt(0)));

const toAscii = (bytes: Uint8Array): string => String.fromCharCode(...bytes);

/** A compressed chunk carrying `body`, framed as the decompressor expects. */
function chunk(body: Uint8Array, compressed = true): Uint8Array {
  const header = 0b011 * 0x1000 + (body.length + 2 - 3) + (compressed ? 0x8000 : 0);
  return new Uint8Array([header & 0xff, header >>> 8, ...body]);
}

function container(...chunks: Uint8Array[]): Uint8Array {
  const out: number[] = [0x01];
  for (const c of chunks) out.push(...c);
  return new Uint8Array(out);
}

/** Deterministic pseudo-random bytes, so a failure is reproducible. */
function pseudoRandom(length: number, seed = 1): Uint8Array {
  const out = new Uint8Array(length);
  let state = seed >>> 0;
  for (let i = 0; i < length; i++) {
    state = (state * 1_664_525 + 1_013_904_223) >>> 0;
    out[i] = (state >>> 24) & 0xff;
  }
  return out;
}

// The two worked examples from [MS-OVBA] section 3.2. These are the only bytes
// in this file that were not produced by our own code, which makes them the
// only real check that we agree with Microsoft's implementation rather than
// merely with ourselves.
const NORMAL_TEXT = '#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa';
const NORMAL_COMPRESSED = hex(`
  01 2F B0 00 23 61 61 61 62 63 64 65 82 66 00 70
  61 67 68 69 6A 01 38 08 61 6B 6C 00 20 6D 6E 6F
  70 06 71 02 70 04 00 72 73 74 75 76 10 77 78 79
  7A 00 2C
`);
const MAXIMUM_TEXT = 'a'.repeat(73);
const MAXIMUM_COMPRESSED = hex('01 03 B0 02 61 45 00');

describe('specification examples, section 3.2', () => {
  it('decompresses the normal compression example', () => {
    expect(toAscii(decompress(NORMAL_COMPRESSED))).toBe(NORMAL_TEXT);
  });

  it('compresses the normal compression example to the published bytes', () => {
    expect(Array.from(compress(ascii(NORMAL_TEXT)))).toEqual(Array.from(NORMAL_COMPRESSED));
  });

  it('decompresses the maximum compression example', () => {
    expect(toAscii(decompress(MAXIMUM_COMPRESSED))).toBe(MAXIMUM_TEXT);
  });

  it('compresses the maximum compression example to the published bytes', () => {
    expect(Array.from(compress(ascii(MAXIMUM_TEXT)))).toEqual(Array.from(MAXIMUM_COMPRESSED));
  });

  it('reads the maximum example as one literal and one 72 byte overlapping copy', () => {
    // 0x0045 at position 1: four offset bits, so length is 0x45 + 3 and offset 1.
    expect(decompress(MAXIMUM_COMPRESSED).length).toBe(1 + 72);
  });
});

describe('decompress: container framing', () => {
  it('rejects an empty buffer', () => {
    expect(() => decompress(new Uint8Array())).toThrow(VbaCompressionError);
  });

  it('rejects a container whose signature byte is not 0x01', () => {
    expect(() => decompress(new Uint8Array([0x00, 0x03, 0xb0, 0x61]))).toThrow(/signature byte/);
  });

  it('accepts a container holding nothing but the signature byte', () => {
    expect(decompress(new Uint8Array([0x01])).length).toBe(0);
  });

  it('rejects a chunk whose signature bits are not 0b011', () => {
    const bad = new Uint8Array([0x01, 0x03, 0xa0, 0x00, 0x61]);
    expect(() => decompress(bad)).toThrow(/expected 0b011/);
  });

  it('stops cleanly on a header truncated to one byte', () => {
    const truncated = new Uint8Array([...compress(ascii('hello')), 0x03]);
    expect(toAscii(decompress(truncated))).toBe('hello');
  });

  it('reads a raw chunk verbatim', () => {
    const raw = pseudoRandom(4096, 7);
    expect(Array.from(decompress(container(chunk(raw, false))))).toEqual(Array.from(raw));
  });

  it('tolerates a raw chunk truncated by the end of the buffer', () => {
    const raw = ascii('partial raw chunk');
    const framed = container(chunk(raw, false)).subarray(0, 3 + raw.length);
    expect(toAscii(decompress(framed))).toBe('partial raw chunk');
  });

  it('joins several chunks into one output', () => {
    const first = new Uint8Array([0x00, 0x41, 0x42, 0x43]);
    const second = new Uint8Array([0x00, 0x44, 0x45]);
    expect(toAscii(decompress(container(chunk(first), chunk(second))))).toBe('ABCDE');
  });

  it('ends a token sequence at the chunk boundary even when its flag byte promises eight', () => {
    // The final TokenSequence in a chunk may hold as few as one token, so the
    // seven unused bits of its flag byte must not pull in the next chunk.
    const body = new Uint8Array([0x00, 0x41]);
    expect(toAscii(decompress(container(chunk(body), chunk(new Uint8Array([0x00, 0x42])))))).toBe(
      'AB',
    );
  });

  it('rejects a copy token that reaches before the start of its chunk', () => {
    const body = new Uint8Array([0x02, 0x41, 0xff, 0xff]);
    expect(() => decompress(container(chunk(body)))).toThrow(/before the start/);
  });

  it('rejects a copy token cut off by the end of its chunk', () => {
    const body = new Uint8Array([0x02, 0x41, 0x00]);
    expect(() => decompress(container(chunk(body)))).toThrow(/runs past the end/);
  });

  it('rejects a chunk that would decode to more than 4096 bytes', () => {
    // 'a' then a maximal copy, repeated: each token sequence adds ~4098 bytes.
    const body = new Uint8Array([0x02, 0x61, 0xff, 0x0f, 0x02, 0x61, 0xff, 0x0f]);
    expect(() => decompress(container(chunk(body)))).toThrow(/more than 4096/);
  });

  it('honours the output ceiling', () => {
    const big = compress(ascii('x'.repeat(20_000)));
    expect(() => decompress(big, { maxOutput: 1024 })).toThrow(/limit/);
  });
});

describe('decompress: copy token bit split', () => {
  // The offset/length split changes as the decoder moves through a chunk. These
  // cases pin the boundary: at 16 bytes in it is still four offset bits, at 17
  // it becomes five, and the same 16 bits therefore mean different things.
  const literalRun = (text: string): number[] => {
    const out: number[] = [];
    for (let i = 0; i < text.length; i += 8) {
      out.push(0x00, ...ascii(text.slice(i, i + 8)));
    }
    return out;
  };

  it('uses four offset bits at exactly 16 decompressed bytes', () => {
    const body = new Uint8Array([...literalRun('0123456789abcdef'), 0x01, 0x00, 0xf0]);
    // 0xF000 with a four bit count is offset 16, length 3.
    expect(toAscii(decompress(container(chunk(body))))).toBe('0123456789abcdef012');
  });

  it('uses five offset bits at 17 decompressed bytes', () => {
    // The seventeenth literal shares a token sequence with the copy that
    // follows it, so the flag byte is 0b00000010.
    const body = new Uint8Array([
      ...literalRun('0123456789abcdef'),
      0x02,
      ...ascii('g'),
      0x00,
      0x80,
    ]);
    // 0x8000 with a five bit count is offset 17, length 3.
    expect(toAscii(decompress(container(chunk(body))))).toBe('0123456789abcdefg012');
  });

  it('encodes the same offset differently either side of the boundary', () => {
    const before = compress(ascii(`${'0123456789abcde'}0123`));
    const after = compress(ascii(`${'0123456789abcdefg'}0123`));
    // Both end in a copy of "0123", but the token bytes differ because the bit
    // split moved. If they were equal, the derivation would be a constant.
    expect(before.subarray(before.length - 2)).not.toEqual(after.subarray(after.length - 2));
  });

  it('round trips a copy near the far end of a full chunk', () => {
    // At 4096 bytes in, twelve bits carry the offset and length caps at 18.
    const filler = pseudoRandom(4000, 3);
    const input = new Uint8Array([...filler, ...filler.subarray(0, 90)]);
    expect(Array.from(decompress(compress(input)))).toEqual(Array.from(input));
  });
});

describe('compress: framing', () => {
  it('emits only the signature byte for empty input', () => {
    expect(Array.from(compress(new Uint8Array()))).toEqual([0x01]);
  });

  it('marks compressed chunks with the flag bit and the 0b011 signature', () => {
    const out = compress(ascii('hello world'));
    const header = out[1]! | (out[2]! << 8);
    expect((header >> 12) & 0b111).toBe(0b011);
    expect(header & 0x8000).toBe(0x8000);
  });

  it('records the chunk size as the chunk length minus three', () => {
    const out = compress(ascii('hello world'));
    const header = out[1]! | (out[2]! << 8);
    expect((header & 0x0fff) + 3).toBe(out.length - 1);
  });

  it('splits input longer than 4096 bytes into several chunks', () => {
    const input = pseudoRandom(10_000, 11);
    const out = compress(input);
    let chunks = 0;
    for (let pos = 1; pos < out.length; ) {
      const header = out[pos]! | (out[pos + 1]! << 8);
      expect((header >> 12) & 0b111).toBe(0b011);
      pos += (header & 0x0fff) + 3;
      chunks++;
    }
    expect(chunks).toBe(3);
  });

  it('stores an incompressible full chunk raw rather than growing it', () => {
    const input = pseudoRandom(4096, 13);
    const out = compress(input);
    const header = out[1]! | (out[2]! << 8);
    expect(header & 0x8000).toBe(0);
    expect(header & 0x0fff).toBe(4095);
    expect(out.length).toBe(1 + 2 + 4096);
  });

  it('never pads a short incompressible tail with nulls', () => {
    // A raw chunk must decode to exactly 4096 bytes, so the reference encoder
    // would pad here. We split into compressed chunks instead and stay exact.
    const input = pseudoRandom(3800, 17);
    expect(Array.from(decompress(compress(input)))).toEqual(Array.from(input));
  });

  it('produces a smaller stream for repetitive input than for random input', () => {
    const repetitive = ascii('Sub Test()\n'.repeat(300));
    const random = pseudoRandom(repetitive.length, 19);
    expect(compress(repetitive).length).toBeLessThan(compress(random).length / 4);
  });
});

describe('compress: literalsOnly', () => {
  it('round trips', () => {
    const input = ascii('aaaaaaaaaaaaaaaaaaaaaaaaaaaa');
    expect(Array.from(decompress(compress(input, { literalsOnly: true })))).toEqual(
      Array.from(input),
    );
  });

  it('emits no copy tokens', () => {
    const out = compress(ascii('abcabcabcabcabcabc'), { literalsOnly: true });
    expect(out[3]).toBe(0x00);
    expect(out.length).toBe(1 + 2 + 18 + 3);
  });

  it('is larger than the matching encoder on repetitive input', () => {
    const input = ascii('x'.repeat(2000));
    expect(compress(input, { literalsOnly: true }).length).toBeGreaterThan(compress(input).length);
  });
});

describe('round trips', () => {
  const cases: Array<[string, Uint8Array]> = [
    ['empty', new Uint8Array()],
    ['one byte', ascii('a')],
    ['two bytes', ascii('ab')],
    ['three identical bytes', ascii('aaa')],
    ['every byte value', new Uint8Array(Array.from({ length: 256 }, (_, i) => i))],
    ['all zeroes, 5000 bytes', new Uint8Array(5000)],
    ['all 0xff, 4096 bytes', new Uint8Array(4096).fill(0xff)],
    ['exactly one chunk', pseudoRandom(4096, 23)],
    ['one chunk plus one byte', pseudoRandom(4097, 29)],
    ['one chunk minus one byte', pseudoRandom(4095, 31)],
    ['two chunks exactly', pseudoRandom(8192, 37)],
    ['three chunks and a tail', pseudoRandom(9000, 41)],
    ['ascending run', new Uint8Array(Array.from({ length: 3000 }, (_, i) => i & 0xff))],
    ['long run of one byte', new Uint8Array(9000).fill(0x61)],
    ['alternating pair', new Uint8Array(Array.from({ length: 6000 }, (_, i) => (i % 2) * 0xff))],
    ['vba source', ascii(`Sub Demo()\r\n  Dim i As Long\r\n  For i = 1 To 10\r\n  Next i\r\nEnd Sub\r\n`)],
    ['repeated vba source', ascii('Sub Demo()\r\n  MsgBox "hi"\r\nEnd Sub\r\n'.repeat(200))],
    ['high bytes', new Uint8Array(Array.from({ length: 1000 }, (_, i) => 0x80 | (i & 0x7f)))],
    ['single null', new Uint8Array([0])],
    ['nulls and text', ascii('a b c '.repeat(100))],
  ];

  it.each(cases)('%s', (_name, input) => {
    const compressed = compress(input);
    expect(compressed[0]).toBe(0x01);
    expect(Array.from(decompress(compressed))).toEqual(Array.from(input));
  });

  it.each(cases)('%s, literals only', (_name, input) => {
    expect(Array.from(decompress(compress(input, { literalsOnly: true })))).toEqual(
      Array.from(input),
    );
  });

  it('round trips every length from 0 to 300', () => {
    const source = ascii('the quick brown fox jumps over the lazy dog. '.repeat(10));
    for (let length = 0; length <= 300; length++) {
      const input = source.subarray(0, length);
      expect(Array.from(decompress(compress(input))), `length ${length}`).toEqual(
        Array.from(input),
      );
    }
  });

  it('round trips lengths straddling the 4096 byte chunk boundary', () => {
    const source = pseudoRandom(4200, 43);
    for (const length of [4093, 4094, 4095, 4096, 4097, 4098, 4099, 4100]) {
      const input = source.subarray(0, length);
      expect(Array.from(decompress(compress(input))), `length ${length}`).toEqual(
        Array.from(input),
      );
    }
  });

  it('round trips a hundred pseudo-random inputs of varying length', () => {
    for (let seed = 1; seed <= 100; seed++) {
      const input = pseudoRandom(seed * 37, seed);
      expect(Array.from(decompress(compress(input))), `seed ${seed}`).toEqual(Array.from(input));
    }
  });

  it('round trips text with long repeated sections across a chunk boundary', () => {
    const block = ascii('Option Explicit\r\n'.repeat(500));
    const input = new Uint8Array([...block, ...pseudoRandom(200, 47), ...block]);
    expect(Array.from(decompress(compress(input)))).toEqual(Array.from(input));
  });
});
