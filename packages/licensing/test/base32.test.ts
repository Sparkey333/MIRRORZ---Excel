import { describe, expect, it } from 'vitest';
import {
  ALPHABET,
  CHECKSUM_BYTES,
  checksum,
  decodeBase32,
  decodeChecked,
  encodeBase32,
  encodeChecked,
  formatGroups,
} from '../src/base32.js';

function bytes(...values: number[]): Uint8Array {
  return Uint8Array.from(values);
}

function randomBytes(length: number, seed: number): Uint8Array {
  // Deterministic pseudo-random so a failure is reproducible.
  const out = new Uint8Array(length);
  let state = seed >>> 0;
  for (let i = 0; i < length; i += 1) {
    state = (state * 1_664_525 + 1_013_904_223) >>> 0;
    out[i] = state >>> 24;
  }
  return out;
}

describe('alphabet', () => {
  it('has 32 characters', () => {
    expect(ALPHABET).toHaveLength(32);
  });

  it('omits the characters that are confusable when read aloud', () => {
    for (const character of ['I', 'L', 'O', 'U']) {
      expect(ALPHABET).not.toContain(character);
    }
  });

  it('has no repeated characters', () => {
    expect(new Set(ALPHABET).size).toBe(32);
  });
});

describe('base32 round trip', () => {
  it('encodes nothing as nothing', () => {
    expect(encodeBase32(new Uint8Array())).toBe('');
  });

  it.each([1, 2, 3, 4, 5, 6, 7, 8, 16, 31, 64, 97, 128])('round trips %i bytes', (length) => {
    const input = randomBytes(length, length + 1);
    const decoded = decodeBase32(encodeBase32(input));
    expect(decoded.ok).toBe(true);
    if (decoded.ok) expect(Array.from(decoded.bytes)).toEqual(Array.from(input));
  });

  it('encodes to the expected character count', () => {
    expect(encodeBase32(bytes(0, 0, 0, 0, 0))).toBe('00000000');
  });

  it('ignores case', () => {
    const text = encodeBase32(randomBytes(20, 7));
    const upper = decodeBase32(text);
    const lower = decodeBase32(text.toLowerCase());
    expect(upper.ok && lower.ok).toBe(true);
    if (upper.ok && lower.ok) expect(Array.from(lower.bytes)).toEqual(Array.from(upper.bytes));
  });
});

describe('reading a key over the phone', () => {
  it('reads O as zero', () => {
    const zero = decodeBase32('00000000');
    const oh = decodeBase32('OOOOOOOO');
    expect(zero.ok && oh.ok).toBe(true);
    if (zero.ok && oh.ok) expect(Array.from(oh.bytes)).toEqual(Array.from(zero.bytes));
  });

  it('reads I and L as one', () => {
    const one = decodeBase32('11111111');
    const eye = decodeBase32('IIIIIIII');
    const ell = decodeBase32('LLLLLLLL');
    expect(one.ok && eye.ok && ell.ok).toBe(true);
    if (one.ok && eye.ok && ell.ok) {
      expect(Array.from(eye.bytes)).toEqual(Array.from(one.bytes));
      expect(Array.from(ell.bytes)).toEqual(Array.from(one.bytes));
    }
  });

  it('ignores dashes, spaces and line breaks', () => {
    const text = encodeBase32(randomBytes(30, 11));
    const grouped = formatGroups(text, { group: 5, perLine: 3 });
    const spaced = grouped.replace(/-/g, ' ');
    const plain = decodeBase32(text);
    const withSeparators = decodeBase32(grouped);
    const withSpaces = decodeBase32(spaced);
    expect(plain.ok && withSeparators.ok && withSpaces.ok).toBe(true);
    if (plain.ok && withSeparators.ok && withSpaces.ok) {
      expect(Array.from(withSeparators.bytes)).toEqual(Array.from(plain.bytes));
      expect(Array.from(withSpaces.bytes)).toEqual(Array.from(plain.bytes));
    }
  });

  it('rejects a character outside the alphabet and reports where', () => {
    const result = decodeBase32('ABCD$EFG');
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.reason).toBe('charset');
      expect(result.at).toBe(4);
    }
  });

  it('rejects empty input', () => {
    const result = decodeBase32('   ');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('empty');
  });

  it('rejects a dropped character that leaves stray bits', () => {
    const result = decodeBase32('A');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('padding');
  });
});

describe('checksum', () => {
  it('is two bytes', () => {
    expect(checksum(bytes(1, 2, 3))).toHaveLength(CHECKSUM_BYTES);
  });

  it('round trips through encodeChecked', () => {
    const input = randomBytes(48, 3);
    const result = decodeChecked(encodeChecked(input));
    expect(result.ok).toBe(true);
    if (result.ok) expect(Array.from(result.bytes)).toEqual(Array.from(input));
  });

  it('rejects a single mistyped character as a checksum failure, not a forgery', () => {
    const text = encodeChecked(randomBytes(40, 5));
    const index = 3;
    const wrong = text[index] === 'A' ? 'B' : 'A';
    const mistyped = `${text.slice(0, index)}${wrong}${text.slice(index + 1)}`;
    const result = decodeChecked(mistyped);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('checksum');
  });

  it('rejects a transposition', () => {
    const text = encodeChecked(randomBytes(40, 9));
    const swapped = `${text.slice(0, 6)}${text[7]}${text[6]}${text.slice(8)}`;
    const result = decodeChecked(swapped);
    // A transposition of two identical characters is a no-op, so only assert
    // when the text actually changed.
    if (swapped !== text) expect(result.ok).toBe(false);
  });

  it('rejects input too short to contain a checksum', () => {
    const result = decodeChecked(encodeBase32(bytes(1)));
    expect(result.ok).toBe(false);
  });

  it('reports a bad character before it reports a bad checksum', () => {
    const result = decodeChecked('!!!!');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('charset');
  });
});

describe('formatGroups', () => {
  it('groups into blocks of five', () => {
    expect(formatGroups('ABCDEFGHIJ', { group: 5, perLine: 0 })).toBe('ABCDE-FGHIJ');
  });

  it('breaks lines after the requested number of blocks', () => {
    const formatted = formatGroups('A'.repeat(60), { group: 5, perLine: 6 });
    expect(formatted.split('\n')).toHaveLength(2);
  });

  it('keeps a short remainder as its own block', () => {
    expect(formatGroups('ABCDEFG', { group: 5, perLine: 0 })).toBe('ABCDE-FG');
  });
});
