import { generateKeyPairSync } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { ALPHABET } from '../src/base32.js';
import type { LicensePayload } from '../src/codec.js';
import {
  derivePublicKey,
  generateKeyPair,
  makePayload,
  signLicense,
  toPrivateKey,
  toPublicKey,
  verifyLicense,
} from '../src/license.js';

const KEYS = generateKeyPair();
const OTHER = generateKeyPair();

const PAYLOAD: LicensePayload = makePayload({
  id: 'MZ-ABC123',
  email: 'buyer@example.com',
  plan: 'pro',
  kind: 'perpetual',
  issued: Date.UTC(2026, 0, 1),
  maintenanceExpires: Date.UTC(2027, 0, 1),
  seats: 1,
  major: 1,
});

const KEY_TEXT = signLicense(PAYLOAD, KEYS.privateKey);

function mutateOneCharacter(text: string, index: number): string {
  const current = text[index]!;
  const replacement = current === ALPHABET[0] ? ALPHABET[1]! : ALPHABET[0]!;
  return `${text.slice(0, index)}${replacement}${text.slice(index + 1)}`;
}

describe('key generation', () => {
  it('produces a 32-byte public key in base64', () => {
    expect(Buffer.from(KEYS.publicKey, 'base64')).toHaveLength(32);
  });

  it('produces a 32-byte private seed in base64', () => {
    expect(Buffer.from(KEYS.privateKey, 'base64')).toHaveLength(32);
  });

  it('produces PEM forms of both halves', () => {
    expect(KEYS.publicKeyPem).toContain('BEGIN PUBLIC KEY');
    expect(KEYS.privateKeyPem).toContain('BEGIN PRIVATE KEY');
  });

  it('produces a different pair each time', () => {
    expect(KEYS.publicKey).not.toBe(OTHER.publicKey);
  });
});

describe('signing and verifying', () => {
  it('verifies a licence it signed', () => {
    const result = verifyLicense(KEY_TEXT, KEYS.publicKey);
    expect(result.valid).toBe(true);
    expect(result.reason).toBe('ok');
  });

  it('returns the payload that was signed', () => {
    const result = verifyLicense(KEY_TEXT, KEYS.publicKey);
    expect(result.payload).toEqual(PAYLOAD);
  });

  it('formats the key in typable blocks', () => {
    expect(KEY_TEXT).toMatch(/^[0-9A-Z-\n]+$/);
    for (const line of KEY_TEXT.split('\n')) {
      for (const block of line.split('-')) expect(block.length).toBeLessThanOrEqual(5);
    }
  });

  it('verifies after the key is retyped in lower case with spaces', () => {
    const retyped = KEY_TEXT.replace(/-/g, ' ').replace(/\n/g, ' ').toLowerCase();
    expect(verifyLicense(retyped, KEYS.publicKey).valid).toBe(true);
  });

  it('verifies after the key is pasted as one unbroken run', () => {
    const flattened = KEY_TEXT.replace(/[-\n]/g, '');
    expect(verifyLicense(flattened, KEYS.publicKey).valid).toBe(true);
  });

  it('accepts a PEM public key', () => {
    expect(verifyLicense(KEY_TEXT, KEYS.publicKeyPem).valid).toBe(true);
  });

  it('accepts a raw byte public key', () => {
    const raw = new Uint8Array(Buffer.from(KEYS.publicKey, 'base64'));
    expect(verifyLicense(KEY_TEXT, raw).valid).toBe(true);
  });

  it('accepts a KeyObject public key', () => {
    expect(verifyLicense(KEY_TEXT, toPublicKey(KEYS.publicKey)).valid).toBe(true);
  });

  it('accepts a private key given as PEM when signing', () => {
    const text = signLicense(PAYLOAD, KEYS.privateKeyPem);
    expect(verifyLicense(text, KEYS.publicKey).valid).toBe(true);
  });

  it('is deterministic: Ed25519 signs the same payload identically', () => {
    expect(signLicense(PAYLOAD, KEYS.privateKey)).toBe(KEY_TEXT);
  });
});

describe('tampering', () => {
  it('rejects a licence signed by somebody else', () => {
    const forged = signLicense(PAYLOAD, OTHER.privateKey);
    const result = verifyLicense(forged, KEYS.publicKey);
    expect(result.valid).toBe(false);
    expect(result.reason).toBe('signature');
    expect(result.payload).toBeNull();
  });

  it('rejects an upgraded plan', () => {
    const upgraded = signLicense({ ...PAYLOAD, plan: 'team' }, OTHER.privateKey);
    expect(verifyLicense(upgraded, KEYS.publicKey).reason).toBe('signature');
  });

  it('rejects every single-character mutation', () => {
    const flat = KEY_TEXT.replace(/[-\n]/g, '');
    for (let i = 0; i < flat.length; i += 7) {
      const mutated = mutateOneCharacter(flat, i);
      const result = verifyLicense(mutated, KEYS.publicKey);
      expect(result.valid).toBe(false);
      expect(['checksum', 'signature', 'malformed']).toContain(result.reason);
    }
  });

  it('reports a mistyped character as a checksum problem, not a forgery', () => {
    const flat = KEY_TEXT.replace(/[-\n]/g, '');
    const result = verifyLicense(mutateOneCharacter(flat, 10), KEYS.publicKey);
    expect(result.reason).toBe('checksum');
    expect(result.message).toContain('character');
  });

  it('rejects a truncated key', () => {
    const flat = KEY_TEXT.replace(/[-\n]/g, '');
    expect(verifyLicense(flat.slice(0, flat.length - 8), KEYS.publicKey).valid).toBe(false);
  });

  it('rejects a key with an extra block appended', () => {
    const flat = KEY_TEXT.replace(/[-\n]/g, '');
    expect(verifyLicense(`${flat}ABCDE`, KEYS.publicKey).valid).toBe(false);
  });

  it('rejects a payload spliced onto another licence signature', () => {
    const other = signLicense({ ...PAYLOAD, plan: 'team', id: 'MZ-OTHER' }, KEYS.privateKey);
    const flatA = KEY_TEXT.replace(/[-\n]/g, '');
    const flatB = other.replace(/[-\n]/g, '');
    const spliced = flatB.slice(0, 20) + flatA.slice(20);
    expect(verifyLicense(spliced, KEYS.publicKey).valid).toBe(false);
  });
});

describe('a build defect is never reported as a forgery', () => {
  // Regression: any object that was not a Uint8Array used to be passed straight
  // through as if it were a KeyObject. `verify` then threw, the throw was
  // swallowed as "signature", and a paying customer was told their key "was not
  // issued by us" because OUR build shipped a broken key.
  it('rejects a key of the wrong algorithm as a build problem', () => {
    const rsa = generateKeyPairSync('rsa', { modulusLength: 2048 });
    const result = verifyLicense(KEY_TEXT, rsa.publicKey);
    expect(result.reason).toBe('public-key');
    expect(result.message).toContain('free tier');
  });

  it('rejects a private key handed in where the public one belongs', () => {
    const result = verifyLicense(KEY_TEXT, toPrivateKey(KEYS.privateKey));
    expect(result.reason).toBe('public-key');
  });

  it.each([null, undefined, 42, {}, [], true])('reports %p as a build problem, not a forgery', (key) => {
    const result = verifyLicense(KEY_TEXT, key as never);
    expect(result.reason).toBe('public-key');
    expect(result.valid).toBe(false);
  });

  it('still calls a genuine forgery a forgery', () => {
    expect(verifyLicense(signLicense(PAYLOAD, OTHER.privateKey), KEYS.publicKey).reason).toBe('signature');
  });
});

describe('deriving the public half', () => {
  it('recovers the public key from the private key', () => {
    expect(derivePublicKey(KEYS.privateKey)).toBe(KEYS.publicKey);
  });

  it('recovers it from the PEM form too', () => {
    expect(derivePublicKey(KEYS.privateKeyPem)).toBe(KEYS.publicKey);
  });

  it('verifies a freshly minted licence against the derived key', () => {
    // This is the check the minting tool runs on every licence it issues.
    expect(verifyLicense(KEY_TEXT, derivePublicKey(KEYS.privateKey)).valid).toBe(true);
  });

  it('does not silently derive from the wrong pair', () => {
    expect(verifyLicense(KEY_TEXT, derivePublicKey(OTHER.privateKey)).reason).toBe('signature');
  });
});

describe('missing and malformed input', () => {
  it.each([null, undefined, '', '   ', '\n'])('treats %p as no licence at all', (text) => {
    const result = verifyLicense(text as string | null, KEYS.publicKey);
    expect(result.reason).toBe('empty');
    expect(result.valid).toBe(false);
  });

  it('reports prose as malformed', () => {
    expect(verifyLicense('this is not a licence key', KEYS.publicKey).reason).toBe('malformed');
  });

  it('reports an unusable public key without blaming the user', () => {
    const result = verifyLicense(KEY_TEXT, 'not-a-key');
    expect(result.reason).toBe('public-key');
    expect(result.message).toContain('free tier');
  });

  it('never throws, whatever it is handed', () => {
    const inputs = ['', '0', 'A'.repeat(500), '!!!', KEY_TEXT.slice(3), '01234-56789'];
    for (const input of inputs) {
      expect(() => verifyLicense(input, KEYS.publicKey)).not.toThrow();
      expect(() => verifyLicense(input, 'garbage')).not.toThrow();
    }
  });

  it('never throws on fuzzed base32 text', () => {
    let state = 12_345;
    for (let round = 0; round < 200; round += 1) {
      let text = '';
      const length = 20 + (round % 90);
      for (let i = 0; i < length; i += 1) {
        state = (state * 1_664_525 + 1_013_904_223) >>> 0;
        text += ALPHABET[state % 32];
      }
      const result = verifyLicense(text, KEYS.publicKey);
      expect(result.valid).toBe(false);
      expect(result.payload).toBeNull();
    }
  });

  it('carries a message for every failure it can report', () => {
    for (const text of [null, 'prose', KEY_TEXT]) {
      const result = verifyLicense(text, KEYS.publicKey);
      expect(result.message.length).toBeGreaterThan(0);
      expect(result.message).not.toContain('!');
    }
  });
});
