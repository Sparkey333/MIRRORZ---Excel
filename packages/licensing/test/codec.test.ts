import { describe, expect, it } from 'vitest';
import { DAY_MS, PAYLOAD_FORMAT, decodePayload, encodePayload, fromDays, toDays } from '../src/codec.js';
import type { LicensePayload } from '../src/codec.js';

const BASE: LicensePayload = {
  version: PAYLOAD_FORMAT,
  id: 'MZ-000123',
  email: 'buyer@example.com',
  plan: 'pro',
  kind: 'perpetual',
  issued: Date.UTC(2026, 0, 15),
  expires: null,
  maintenanceExpires: Date.UTC(2027, 0, 15),
  seats: 2,
  features: [],
  major: 1,
};

function roundTrip(payload: LicensePayload): LicensePayload {
  const result = decodePayload(encodePayload(payload));
  if (!result.ok) throw new Error(`decode failed: ${result.reason}`);
  return result.payload;
}

describe('payload codec', () => {
  it('round trips a perpetual licence', () => {
    expect(roundTrip(BASE)).toEqual(BASE);
  });

  it('round trips a subscription with a term', () => {
    const payload: LicensePayload = {
      ...BASE,
      kind: 'subscription',
      expires: Date.UTC(2027, 0, 15),
      maintenanceExpires: Date.UTC(2027, 0, 15),
    };
    expect(roundTrip(payload)).toEqual(payload);
  });

  it('round trips feature grants', () => {
    const payload = { ...BASE, features: ['macro.execute', 'diff.semantic'] };
    expect(roundTrip(payload).features).toEqual(['macro.execute', 'diff.semantic']);
  });

  it('round trips non-ASCII email addresses', () => {
    const payload = { ...BASE, email: 'zoë@exämple.test' };
    expect(roundTrip(payload).email).toBe('zoë@exämple.test');
  });

  it('keeps null dates null', () => {
    const payload = { ...BASE, expires: null, maintenanceExpires: null };
    const decoded = roundTrip(payload);
    expect(decoded.expires).toBeNull();
    expect(decoded.maintenanceExpires).toBeNull();
  });

  it('truncates dates to whole days', () => {
    const payload = { ...BASE, issued: Date.UTC(2026, 0, 15) + 3_600_000 };
    expect(roundTrip(payload).issued).toBe(Date.UTC(2026, 0, 15));
  });

  it('stays compact enough to retype', () => {
    expect(encodePayload(BASE).length).toBeLessThan(80);
  });

  it.each([1, 3, 50, 5000])('round trips %i seats', (seats) => {
    expect(roundTrip({ ...BASE, seats }).seats).toBe(seats);
  });
});

describe('payload codec rejects what it cannot trust', () => {
  it('rejects a future payload format', () => {
    const bytes = encodePayload(BASE);
    bytes[0] = PAYLOAD_FORMAT + 1;
    const result = decodePayload(bytes);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('format');
  });

  it('rejects an unknown plan code', () => {
    const bytes = encodePayload(BASE);
    bytes[2] = 99;
    const result = decodePayload(bytes);
    expect(result.ok).toBe(false);
  });

  it('rejects trailing bytes', () => {
    const bytes = encodePayload(BASE);
    const padded = new Uint8Array(bytes.length + 1);
    padded.set(bytes, 0);
    const result = decodePayload(padded);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('format');
  });

  it('rejects a truncated payload', () => {
    const bytes = encodePayload(BASE);
    const result = decodePayload(bytes.subarray(0, bytes.length - 4));
    expect(result.ok).toBe(false);
  });

  it('rejects an empty buffer without throwing', () => {
    expect(decodePayload(new Uint8Array()).ok).toBe(false);
  });

  it('rejects zero seats', () => {
    const bytes = encodePayload(BASE);
    bytes[3] = 0;
    const result = decodePayload(bytes);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('range');
  });

  it('refuses to encode an absurdly long field', () => {
    expect(() => encodePayload({ ...BASE, id: 'x'.repeat(600) })).toThrow();
  });

  it('never throws on random bytes', () => {
    for (let seed = 0; seed < 200; seed += 1) {
      const bytes = new Uint8Array(24);
      let state = seed * 2_654_435_761;
      for (let i = 0; i < bytes.length; i += 1) {
        state = (state * 1_664_525 + 1_013_904_223) >>> 0;
        bytes[i] = state >>> 24;
      }
      expect(() => decodePayload(bytes)).not.toThrow();
    }
  });
});

describe('dates a Date object cannot hold', () => {
  // Regression: a day count large enough to put `new Date(...)` outside its
  // representable range decoded happily, and then threw inside the assessment's
  // explanation - a throw in the one module that promises never to throw.
  it('refuses to encode a date past the year 9999', () => {
    expect(() => encodePayload({ ...BASE, maintenanceExpires: 8.64e15 })).toThrow();
  });

  it('refuses to encode a date before the epoch', () => {
    expect(() => encodePayload({ ...BASE, issued: Date.UTC(1969, 0, 1) })).toThrow();
  });

  it('refuses to decode an absurd day count', () => {
    // With all three dates at day 0 they occupy one byte each, at offsets 4, 5
    // and 6, so the maintenance varint can be swapped for an oversized one
    // without knowing anything else about the layout.
    const flat = encodePayload({ ...BASE, issued: 0, expires: null, maintenanceExpires: null });
    expect(decodePayload(flat).ok).toBe(true);
    const spliced = Uint8Array.from([
      ...flat.subarray(0, 6),
      ...leb(9_000_000), // maintenance, past the ceiling
      ...flat.subarray(7),
    ]);
    const result = decodePayload(spliced);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.reason).toBe('range');
  });

  it('still round trips a licence dated in the far but sane future', () => {
    const payload = { ...BASE, maintenanceExpires: Date.UTC(2200, 0, 1) };
    expect(roundTrip(payload).maintenanceExpires).toBe(Date.UTC(2200, 0, 1));
  });
});

/** LEB128, matching the encoder, so the splice above is a real payload byte. */
function leb(value: number): number[] {
  const out: number[] = [];
  let rest = value;
  while (rest >= 0x80) {
    out.push((rest & 0x7f) | 0x80);
    rest = Math.floor(rest / 128);
  }
  out.push(rest);
  return out;
}

describe('day conversion', () => {
  it('round trips through days', () => {
    expect(fromDays(toDays(Date.UTC(2026, 5, 1)))).toBe(Date.UTC(2026, 5, 1));
  });

  it('uses whole UTC days', () => {
    expect(DAY_MS).toBe(86_400_000);
  });
});
