/**
 * Compact binary serialization of a licence payload.
 *
 * JSON would be easier, but the whole payload is signed and then rendered as
 * something a human retypes, and every byte costs two typed characters. The
 * Ed25519 signature alone is 64 bytes (103 characters of base32) and there is no
 * way to shrink that without a server to shorten against - which we refuse to
 * have. So the payload itself is packed: LEB128 varints, day-granularity dates,
 * and plan names as one-byte codes. A typical licence lands near 60 bytes rather
 * than the ~200 the equivalent JSON would take.
 *
 * The format carries its own version byte. A licence signed today must still
 * decode in five years, so decoding is additive-only: fields may be appended,
 * never reordered or repurposed.
 */

import type { PlanId } from './plans.js';
import { PLAN_CODES, planFromCode } from './plans.js';

export const PAYLOAD_FORMAT = 1;

export const DAY_MS = 86_400_000;

/** Kind of grant. Perpetual licences never stop running - see license.ts. */
export type LicenseKind = 'perpetual' | 'subscription';

export interface LicensePayload {
  /** Payload format version, not the product version. */
  readonly version: number;
  /** Opaque order/licence identifier, printed on the receipt. */
  readonly id: string;
  readonly email: string;
  readonly plan: PlanId;
  readonly kind: LicenseKind;
  /** Issue date, epoch milliseconds truncated to a whole day (UTC). */
  readonly issued: number;
  /**
   * End of the right to run, epoch ms. Subscriptions only. A perpetual licence
   * MUST carry null here; see the honesty rules in license.ts.
   */
  readonly expires: number | null;
  /** End of update coverage, epoch ms. Limits future builds, never this one. */
  readonly maintenanceExpires: number | null;
  readonly seats: number;
  /** Extra capability grants beyond the plan, for one-off deals and beta thanks. */
  readonly features: readonly string[];
  /** Major product version a perpetual licence covers. 0 means "any". */
  readonly major: number;
}

export type CodecFailure = 'truncated' | 'format' | 'range' | 'utf8';

export type DecodePayloadResult =
  | { readonly ok: true; readonly payload: LicensePayload }
  | { readonly ok: false; readonly reason: CodecFailure };

// Sanity ceilings. A licence larger than this is either corrupt or hostile, and
// the parser runs on untrusted text pasted in by anyone.
const MAX_STRING_BYTES = 512;
const MAX_FEATURES = 32;
const MAX_SEATS = 100_000;

class Writer {
  private readonly bytes: number[] = [];

  byte(value: number): void {
    this.bytes.push(value & 0xff);
  }

  varint(value: number): void {
    if (!Number.isFinite(value) || value < 0) throw new RangeError('varint must be a non-negative integer');
    let rest = Math.trunc(value);
    while (rest >= 0x80) {
      this.bytes.push((rest & 0x7f) | 0x80);
      rest = Math.floor(rest / 128);
    }
    this.bytes.push(rest);
  }

  text(value: string): void {
    const encoded = new TextEncoder().encode(value);
    if (encoded.length > MAX_STRING_BYTES) throw new RangeError('string too long for a licence payload');
    this.varint(encoded.length);
    for (const byte of encoded) this.bytes.push(byte);
  }

  finish(): Uint8Array {
    return Uint8Array.from(this.bytes);
  }
}

class Reader {
  private offset = 0;
  constructor(private readonly bytes: Uint8Array) {}

  get exhausted(): boolean {
    return this.offset >= this.bytes.length;
  }

  byte(): number {
    if (this.offset >= this.bytes.length) throw new CodecError('truncated');
    return this.bytes[this.offset++]!;
  }

  varint(): number {
    let result = 0;
    let shift = 1;
    for (let i = 0; i < 8; i += 1) {
      const byte = this.byte();
      result += (byte & 0x7f) * shift;
      if ((byte & 0x80) === 0) return result;
      shift *= 128;
    }
    throw new CodecError('range');
  }

  text(): string {
    const length = this.varint();
    if (length > MAX_STRING_BYTES) throw new CodecError('range');
    if (this.offset + length > this.bytes.length) throw new CodecError('truncated');
    const slice = this.bytes.subarray(this.offset, this.offset + length);
    this.offset += length;
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(slice);
    } catch {
      throw new CodecError('utf8');
    }
  }
}

class CodecError extends Error {
  constructor(readonly reason: CodecFailure) {
    super(`licence payload ${reason}`);
  }
}

/** Whole UTC days since the epoch. Licences do not need minute precision. */
export function toDays(epochMs: number): number {
  return Math.floor(epochMs / DAY_MS);
}

export function fromDays(days: number): number {
  return days * DAY_MS;
}

// Day 0 is 1970-01-01, which no licence can legitimately carry, so zero is free
// to mean "absent" for the two optional dates.
function encodeOptionalDay(epochMs: number | null): number {
  if (epochMs === null) return 0;
  return Math.max(1, toDays(epochMs));
}

function decodeOptionalDay(days: number): number | null {
  return days === 0 ? null : fromDays(days);
}

export function encodePayload(payload: LicensePayload): Uint8Array {
  const writer = new Writer();
  writer.byte(PAYLOAD_FORMAT);
  writer.byte(payload.kind === 'subscription' ? 1 : 0);
  writer.varint(PLAN_CODES[payload.plan]);
  writer.varint(payload.seats);
  writer.varint(toDays(payload.issued));
  writer.varint(encodeOptionalDay(payload.expires));
  writer.varint(encodeOptionalDay(payload.maintenanceExpires));
  writer.varint(payload.major);
  writer.text(payload.id);
  writer.text(payload.email);
  writer.varint(payload.features.length);
  for (const feature of payload.features) writer.text(feature);
  return writer.finish();
}

export function decodePayload(bytes: Uint8Array): DecodePayloadResult {
  try {
    const reader = new Reader(bytes);
    const version = reader.byte();
    if (version !== PAYLOAD_FORMAT) return { ok: false, reason: 'format' };

    const kind: LicenseKind = (reader.byte() & 1) === 1 ? 'subscription' : 'perpetual';
    const plan = planFromCode(reader.varint());
    if (plan === undefined) return { ok: false, reason: 'format' };

    const seats = reader.varint();
    if (seats < 1 || seats > MAX_SEATS) return { ok: false, reason: 'range' };

    const issued = fromDays(reader.varint());
    const expires = decodeOptionalDay(reader.varint());
    const maintenanceExpires = decodeOptionalDay(reader.varint());
    const major = reader.varint();

    const id = reader.text();
    const email = reader.text();

    const featureCount = reader.varint();
    if (featureCount > MAX_FEATURES) return { ok: false, reason: 'range' };
    const features: string[] = [];
    for (let i = 0; i < featureCount; i += 1) features.push(reader.text());

    // Trailing bytes mean the text is not what the signer produced, even if the
    // prefix happens to parse. Refuse rather than guess.
    if (!reader.exhausted) return { ok: false, reason: 'format' };

    return {
      ok: true,
      payload: { version, id, email, plan, kind, issued, expires, maintenanceExpires, seats, features, major },
    };
  } catch (error) {
    if (error instanceof CodecError) return { ok: false, reason: error.reason };
    return { ok: false, reason: 'format' };
  }
}
