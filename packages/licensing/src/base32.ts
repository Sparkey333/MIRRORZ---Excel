/**
 * Human-typable text encoding for licence keys.
 *
 * A licence key gets read aloud over the phone, copied off a printed invoice and
 * retyped by someone who is annoyed. Base64 is hostile to all three: it is case
 * sensitive and contains characters that are ambiguous in most fonts. This module
 * uses Crockford's base32 alphabet, which drops I, L, O and U, and decodes the
 * shapes people actually type instead of the intended character (O as 0, I and L
 * as 1). Case is ignored and separators are ignored, so "abcd-efgh" and
 * "ABCD EFGH" decode identically.
 *
 * A two-byte checksum rides along so a mistyped key can be reported as a typo
 * rather than as a forgery. That distinction is the whole point: the Ed25519
 * signature already rejects every corruption, but "check digit 3 of your key" is
 * a kinder message than "your licence is invalid", and the difference matters
 * when the person is a paying customer sitting on a plane with no way to ask.
 */

import { createHash } from 'node:crypto';

/** Crockford base32: no I, L, O or U, so no 0/O or 1/I ambiguity when spoken. */
export const ALPHABET = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';

const DECODE = new Map<string, number>();
for (let i = 0; i < ALPHABET.length; i += 1) DECODE.set(ALPHABET[i]!, i);
// The confusable shapes, mapped to what the reader meant rather than rejected.
DECODE.set('O', 0);
DECODE.set('I', 1);
DECODE.set('L', 1);

/** Characters allowed as visual separators and silently dropped on decode. */
const SEPARATORS = /[\s\-_.]/g;

export const CHECKSUM_BYTES = 2;

export type Base32Failure = 'charset' | 'padding' | 'empty';

export type DecodeResult =
  | { readonly ok: true; readonly bytes: Uint8Array }
  | { readonly ok: false; readonly reason: Base32Failure; readonly at?: number };

/** Pack bytes big-endian into 5-bit groups. No padding: length is implied. */
export function encodeBase32(bytes: Uint8Array): string {
  let out = '';
  let acc = 0;
  let bits = 0;
  for (const byte of bytes) {
    acc = (acc << 8) | byte;
    bits += 8;
    while (bits >= 5) {
      bits -= 5;
      out += ALPHABET[(acc >>> bits) & 31];
    }
  }
  if (bits > 0) out += ALPHABET[(acc << (5 - bits)) & 31];
  return out;
}

export function decodeBase32(text: string): DecodeResult {
  const cleaned = text.replace(SEPARATORS, '').toUpperCase();
  if (cleaned.length === 0) return { ok: false, reason: 'empty' };

  const bytes: number[] = [];
  let acc = 0;
  let bits = 0;
  for (let i = 0; i < cleaned.length; i += 1) {
    const value = DECODE.get(cleaned[i]!);
    if (value === undefined) return { ok: false, reason: 'charset', at: i };
    acc = (acc << 5) | value;
    bits += 5;
    if (bits >= 8) {
      bits -= 8;
      bytes.push((acc >>> bits) & 0xff);
    }
  }
  // Leftover bits are the tail of the last character and must be zero, otherwise
  // a character was dropped or added and the byte string is not what was signed.
  if (bits >= 5 || (acc & ((1 << bits) - 1)) !== 0) return { ok: false, reason: 'padding' };
  return { ok: true, bytes: Uint8Array.from(bytes) };
}

/** Truncated SHA-256. Not a security control - the signature is. This catches typos. */
export function checksum(bytes: Uint8Array): Uint8Array {
  const digest = createHash('sha256').update(bytes).digest();
  return new Uint8Array(digest.subarray(0, CHECKSUM_BYTES));
}

/** Append the checksum and encode. The inverse of {@link decodeChecked}. */
export function encodeChecked(bytes: Uint8Array): string {
  const tag = checksum(bytes);
  const framed = new Uint8Array(bytes.length + tag.length);
  framed.set(bytes, 0);
  framed.set(tag, bytes.length);
  return encodeBase32(framed);
}

export type CheckedResult =
  | { readonly ok: true; readonly bytes: Uint8Array }
  | { readonly ok: false; readonly reason: Base32Failure | 'checksum' };

export function decodeChecked(text: string): CheckedResult {
  const decoded = decodeBase32(text);
  if (!decoded.ok) return { ok: false, reason: decoded.reason };
  if (decoded.bytes.length <= CHECKSUM_BYTES) return { ok: false, reason: 'checksum' };

  const split = decoded.bytes.length - CHECKSUM_BYTES;
  const body = decoded.bytes.subarray(0, split);
  const tag = decoded.bytes.subarray(split);
  const expected = checksum(body);
  for (let i = 0; i < CHECKSUM_BYTES; i += 1) {
    if (tag[i] !== expected[i]) return { ok: false, reason: 'checksum' };
  }
  return { ok: true, bytes: new Uint8Array(body) };
}

export interface GroupOptions {
  /** Characters per dash-separated block. */
  readonly group?: number;
  /** Blocks per line. Zero puts everything on one line. */
  readonly perLine?: number;
}

/** Break an encoded string into blocks a person can track with a finger. */
export function formatGroups(text: string, options: GroupOptions = {}): string {
  const group = options.group ?? 5;
  const perLine = options.perLine ?? 6;
  const blocks: string[] = [];
  for (let i = 0; i < text.length; i += group) blocks.push(text.slice(i, i + group));
  if (perLine <= 0) return blocks.join('-');

  const lines: string[] = [];
  for (let i = 0; i < blocks.length; i += perLine) {
    lines.push(blocks.slice(i, i + perLine).join('-'));
  }
  return lines.join('\n');
}
