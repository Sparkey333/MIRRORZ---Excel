/**
 * Offline licence keys: Ed25519 signatures verified locally, with no network.
 *
 * THE APPLICATION MAKES NO NETWORK REQUEST TO CHECK A LICENCE. There is no
 * activation call, no heartbeat, no telemetry, no deferred "we will phone home
 * later". The shipped binary contains a 32-byte public key and nothing else; the
 * private key lives on the machine that mints licences and never leaves it. That
 * design is not a convenience, it is the product promise: a spreadsheet must open
 * on a plane, in a SCIF, on a laptop that has never seen a network, in 2040 when
 * whatever server we might have run is long gone.
 *
 * Three honesty rules are encoded here, and they matter more than the crypto:
 *
 *   1. A PERPETUAL LICENCE NEVER STOPS WORKING for the version it covers. What
 *      expires is `maintenanceExpires` - the right to future updates - and that
 *      is checked against the RUNNING BUILD's compile-time date, never against
 *      the wall clock. A perpetual licence therefore has no wall-clock condition
 *      at all: the same build keeps running forever, offline, unchecked.
 *   2. A SUBSCRIPTION THAT CANNOT BE VERIFIED GETS A LONG GRACE PERIOD. We cannot
 *      distinguish "did not renew" from "has been offline for two months", so we
 *      assume the customer and not the thief, for {@link SUBSCRIPTION_GRACE_DAYS}
 *      days past the term.
 *   3. A TAMPERED OR ABSENT LICENCE DEGRADES TO THE FREE TIER. It never destroys
 *      data, never locks a file, never refuses to save. Verification failure is
 *      reported as a downgrade, not as an accusation. Their files are theirs.
 *
 * Nothing in this module can fail closed: every path returns a value, no path
 * throws, and every failure resolves to the free tier, which can open, edit and
 * save every format the product supports.
 */

import { KeyObject, createPrivateKey, createPublicKey, generateKeyPairSync, sign, verify } from 'node:crypto';

import { decodeChecked, encodeChecked, formatGroups } from './base32.js';
import type { GroupOptions } from './base32.js';
import { DAY_MS, decodePayload, encodePayload } from './codec.js';
import type { LicenseKind, LicensePayload } from './codec.js';
import type { PlanId } from './plans.js';

export type { LicenseKind, LicensePayload } from './codec.js';

/**
 * Domain separation. Signing raw payload bytes would let a signature made for
 * some other MIRRORZ artefact be replayed as a licence, so every signature
 * commits to this context string first.
 */
const CONTEXT = new TextEncoder().encode('mirrorz-license-v1');

const SIGNATURE_BYTES = 64;
const RAW_KEY_BYTES = 32;

// DER wrappers, so a bare 32-byte key can be handed to node:crypto. These are
// fixed by RFC 8410 for Ed25519 and never vary.
const SPKI_PREFIX = Buffer.from('302a300506032b6570032100', 'hex');
const PKCS8_PREFIX = Buffer.from('302e020100300506032b657004220420', 'hex');

/**
 * How long a lapsed subscription keeps working. Deliberately long: two months
 * covers a sabbatical, a long flight, an expired card the customer has not
 * noticed, and a bank that declined a renewal for its own reasons. Locking
 * someone out of their own spreadsheet to protect a $12 charge is a bad trade
 * for them and for us.
 */
export const SUBSCRIPTION_GRACE_DAYS = 60;

export interface KeyPairText {
  /** Base64 of the raw 32-byte public key. This is what ships in the app. */
  readonly publicKey: string;
  /** Base64 of the raw 32-byte private seed. This NEVER ships. */
  readonly privateKey: string;
  readonly publicKeyPem: string;
  readonly privateKeyPem: string;
}

/**
 * Build-time only. This function exists so the release engineer can mint a key
 * pair once; it is never called by the application, which only ever verifies.
 * Keep the private half in an offline password manager or an HSM - if it leaks,
 * every licence ever issued becomes forgeable and the only remedy is a new key
 * in a new build, which by design cannot be pushed to installed copies.
 */
export function generateKeyPair(): KeyPairText {
  const { publicKey, privateKey } = generateKeyPairSync('ed25519');
  const spki = publicKey.export({ format: 'der', type: 'spki' });
  const pkcs8 = privateKey.export({ format: 'der', type: 'pkcs8' });
  return {
    publicKey: Buffer.from(spki.subarray(spki.length - RAW_KEY_BYTES)).toString('base64'),
    privateKey: Buffer.from(pkcs8.subarray(pkcs8.length - RAW_KEY_BYTES)).toString('base64'),
    publicKeyPem: publicKey.export({ format: 'pem', type: 'spki' }).toString(),
    privateKeyPem: privateKey.export({ format: 'pem', type: 'pkcs8' }).toString(),
  };
}

export type KeyInput = KeyObject | string | Uint8Array;

function rawBytes(key: string | Uint8Array): Uint8Array | null {
  if (typeof key !== 'string') return key;
  const trimmed = key.trim();
  if (trimmed.length === 0) return null;
  const decoded = Buffer.from(trimmed, 'base64');
  // Buffer.from silently discards junk, so length is the only real validation.
  return decoded.length === RAW_KEY_BYTES ? new Uint8Array(decoded) : null;
}

function isPem(key: KeyInput): key is string {
  return typeof key === 'string' && key.includes('-----BEGIN');
}

/**
 * A key of the wrong shape must be rejected HERE, where the caller can be told
 * it is a build defect, and not left to fail later inside `verify`. `verify`
 * signals a bad signature by returning false and a bad key by throwing, so an
 * unusable key that reaches it is indistinguishable from a forged licence
 * unless it is caught first - and telling a paying customer their key "was not
 * issued by us" because our own build shipped a broken public key is the worst
 * message this module can produce.
 */
function assertEd25519(key: KeyObject, want: 'public' | 'private', label: string): KeyObject {
  if (key.type !== want || key.asymmetricKeyType !== 'ed25519') {
    throw new TypeError(`${label} must be an Ed25519 ${want} key`);
  }
  return key;
}

/** Accept a KeyObject, a PEM block, or the bare 32 bytes (raw or base64). */
export function toPublicKey(key: KeyInput): KeyObject {
  if (key instanceof KeyObject) return assertEd25519(key, 'public', 'public key');
  if (typeof key !== 'string' && !(key instanceof Uint8Array)) {
    throw new TypeError('public key must be a KeyObject, PEM, base64 or 32 raw bytes');
  }
  if (isPem(key)) return assertEd25519(createPublicKey(key), 'public', 'public key');
  const raw = rawBytes(key);
  if (raw === null) throw new TypeError('public key must be 32 raw bytes, base64 or PEM');
  return createPublicKey({ key: Buffer.concat([SPKI_PREFIX, raw]), format: 'der', type: 'spki' });
}

export function toPrivateKey(key: KeyInput): KeyObject {
  if (key instanceof KeyObject) return assertEd25519(key, 'private', 'private key');
  if (typeof key !== 'string' && !(key instanceof Uint8Array)) {
    throw new TypeError('private key must be a KeyObject, PEM, base64 or 32 raw bytes');
  }
  if (isPem(key)) return assertEd25519(createPrivateKey(key), 'private', 'private key');
  const raw = rawBytes(key);
  if (raw === null) throw new TypeError('private key must be 32 raw bytes, base64 or PEM');
  return createPrivateKey({ key: Buffer.concat([PKCS8_PREFIX, raw]), format: 'der', type: 'pkcs8' });
}

/**
 * The public half of a private key, base64, in the form that ships in a build.
 *
 * Build-time only, and it exists so the minter can self-check every licence it
 * issues without being handed the public key separately - a minter that can
 * only check when an extra environment variable happens to be set is a minter
 * that does not check.
 */
export function derivePublicKey(privateKey: KeyInput): string {
  const spki = createPublicKey(toPrivateKey(privateKey)).export({ format: 'der', type: 'spki' });
  return Buffer.from(spki.subarray(spki.length - RAW_KEY_BYTES)).toString('base64');
}

function signedBytes(payloadBytes: Uint8Array): Buffer {
  return Buffer.concat([CONTEXT, payloadBytes]);
}

/**
 * Build-time only, same as {@link generateKeyPair}: this is what the licence
 * minter runs after a payment webhook fires. Returns the typable key text.
 */
export function signLicense(payload: LicensePayload, privateKey: KeyInput, options: GroupOptions = {}): string {
  const payloadBytes = encodePayload(payload);
  const signature = sign(null, signedBytes(payloadBytes), toPrivateKey(privateKey));
  const framed = new Uint8Array(payloadBytes.length + signature.length);
  framed.set(payloadBytes, 0);
  framed.set(signature, payloadBytes.length);
  return formatGroups(encodeChecked(framed), options);
}

export type VerifyReason =
  | 'ok'
  /** No licence text at all. The ordinary state of a free-tier user. */
  | 'empty'
  /** Not decodable as a licence: wrong characters, truncated, or not a key. */
  | 'malformed'
  /** Decodes, but the check digits disagree - almost always a typo. */
  | 'checksum'
  /** Decodes cleanly and is signed by somebody who is not us. */
  | 'signature'
  /** A newer licence format than this build understands. */
  | 'unsupported'
  /** Our own public key is unusable. A build defect, not the user's fault. */
  | 'public-key';

export interface VerifyResult {
  readonly valid: boolean;
  readonly payload: LicensePayload | null;
  readonly reason: VerifyReason;
  /** Message fit to show a human. Never accusatory: a typo is the likely cause. */
  readonly message: string;
}

const MESSAGES: Readonly<Record<VerifyReason, string>> = {
  ok: 'Licence verified offline.',
  empty: 'No licence key entered.',
  malformed: 'That does not look like a licence key. Check for a missing or extra block.',
  checksum: 'The key did not check out - one character is probably wrong. Compare it with your receipt.',
  signature: 'This key was not issued by us. If you bought it from a reseller, ask them for the original receipt.',
  unsupported: 'This key needs a newer version of MIRRORZ Sheets. Your files are unaffected.',
  'public-key': 'This build cannot check licences. Continuing on the free tier.',
};

function fail(reason: VerifyReason): VerifyResult {
  return { valid: false, payload: null, reason, message: MESSAGES[reason] };
}

/**
 * Verify a licence key against the shipped public key. Pure, offline, total: it
 * never throws and never touches the network.
 */
export function verifyLicense(text: string | null | undefined, publicKey: KeyInput): VerifyResult {
  if (text === null || text === undefined || text.trim().length === 0) return fail('empty');

  let key: KeyObject;
  try {
    key = toPublicKey(publicKey);
  } catch {
    return fail('public-key');
  }

  const decoded = decodeChecked(text);
  if (!decoded.ok) return fail(decoded.reason === 'checksum' ? 'checksum' : 'malformed');
  if (decoded.bytes.length <= SIGNATURE_BYTES) return fail('malformed');

  const split = decoded.bytes.length - SIGNATURE_BYTES;
  const payloadBytes = decoded.bytes.subarray(0, split);
  const signature = decoded.bytes.subarray(split);

  let signatureOk = false;
  try {
    signatureOk = verify(null, signedBytes(payloadBytes), key, signature);
  } catch {
    // A usable Ed25519 key answers a bad signature with `false`. A throw means
    // the key itself is unusable, which is our defect and not a forgery, so it
    // is reported as one - see assertEd25519 above.
    return fail('public-key');
  }
  if (!signatureOk) return fail('signature');

  const payload = decodePayload(payloadBytes);
  if (!payload.ok) {
    // The signature already proved we issued these bytes, so a decode failure
    // here means a payload format from the future, not tampering.
    return fail(payload.reason === 'format' ? 'unsupported' : 'malformed');
  }

  return { valid: true, payload: payload.payload, reason: 'ok', message: MESSAGES.ok };
}

export type LicenseState =
  /** No licence installed. Free tier, which is a real tier and not a nag screen. */
  | 'none'
  /** Present but not ours, or unreadable. Free tier, data untouched. */
  | 'invalid'
  /** Subscription inside its term. */
  | 'active'
  /** Perpetual licence covering this build. Has no expiry and never will. */
  | 'perpetual'
  /** Subscription past its term, inside the offline grace window. Full features. */
  | 'grace'
  /** Subscription past its term and past grace. Free tier. */
  | 'lapsed'
  /**
   * A valid perpetual licence, but this BUILD was released after the update
   * coverage ended. The licence is not expired and the covered build still runs
   * with everything paid for; only this newer build falls back to the free tier.
   */
  | 'update-not-covered';

export interface RuntimeContext {
  /** Wall clock. Injected so tests and the app share one code path. */
  readonly now?: number;
  /**
   * Release date of the RUNNING build, baked in at compile time. Update coverage
   * is judged against this and never against the clock, which is what makes a
   * perpetual licence perpetual.
   *
   * Omitting it means "unknown build", which is read as COVERED, never as the
   * current time. Bake a real one in for the update-coverage rule to have any
   * effect; the default is deliberately the one that cannot revoke a plan.
   */
  readonly buildDate?: number;
  /** Major version of the running build, matched against a perpetual `major`. */
  readonly buildMajor?: number;
  /**
   * Highest timestamp the installation has ever observed, from the trial's
   * monotonic mark. Stops a clock rollback resurrecting a lapsed subscription,
   * and is only ever used to move `now` FORWARD.
   */
  readonly clockFloor?: number;
  /** Override for {@link SUBSCRIPTION_GRACE_DAYS}, mostly for tests. */
  readonly graceDays?: number;
}

export interface LicenseAssessment {
  readonly state: LicenseState;
  /** The plan the user actually gets right now. Always a real, usable plan. */
  readonly plan: PlanId;
  readonly kind: LicenseKind | null;
  readonly seats: number;
  readonly features: readonly string[];
  readonly payload: LicensePayload | null;
  readonly reason: VerifyReason;
  /** True when the user is getting less than the licence nominally grants. */
  readonly degraded: boolean;
  readonly updatesCoveredUntil: number | null;
  /** Whether the running build falls inside the licence's update coverage. */
  readonly buildCovered: boolean;
  readonly graceEndsAt: number | null;
  readonly graceDaysRemaining: number;
  /** One sentence for the status line. Plain, never a threat. */
  readonly explanation: string;
}

function daysBetween(from: number, to: number): number {
  return Math.max(0, Math.ceil((to - from) / DAY_MS));
}

function freeAssessment(state: LicenseState, reason: VerifyReason, explanation: string): LicenseAssessment {
  return {
    state,
    plan: 'free',
    kind: null,
    seats: 0,
    features: [],
    payload: null,
    reason,
    degraded: state === 'invalid',
    updatesCoveredUntil: null,
    buildCovered: true,
    graceEndsAt: null,
    graceDaysRemaining: 0,
    explanation,
  };
}

/**
 * Turn a licence key into what the user may do right now. This is where the
 * honesty rules live; the cryptography above only decides whether we issued the
 * key, and this decides what that means.
 */
export function assessLicense(
  text: string | null | undefined,
  publicKey: KeyInput,
  context: RuntimeContext = {},
): LicenseAssessment {
  const verified = verifyLicense(text, publicKey);
  if (!verified.valid || verified.payload === null) {
    const state: LicenseState = verified.reason === 'empty' ? 'none' : 'invalid';
    const explanation =
      state === 'none'
        ? 'Running the free tier. Opening, editing and saving every supported format are never restricted.'
        : `${verified.message} Continuing on the free tier - your files are untouched and still fully editable.`;
    return freeAssessment(state, verified.reason, explanation);
  }

  const payload = verified.payload;
  // A non-finite clock is a caller bug, and the arithmetic below would answer it
  // by silently reporting every subscription as lapsed - a fail-CLOSED path in a
  // module whose whole contract is to fail open. Fall back to the real clock.
  const wall = Number.isFinite(context.now) ? (context.now as number) : Date.now();
  // Only ever forward. A clock behind the high-water mark is a rollback (or a
  // dead CMOS battery); either way we refuse to hand out time that already passed.
  const floor = Number.isFinite(context.clockFloor) ? (context.clockFloor as number) : undefined;
  const now = floor !== undefined ? Math.max(wall, floor) : wall;
  /**
   * NEVER `?? now`.
   *
   * `buildDate` is the compile-time release date of the running build, and it is
   * the ONLY thing update coverage is judged against. Defaulting it to the wall
   * clock would mean that a caller who forgot to bake one in gets a perpetual
   * licence that silently stops granting its plan on the day maintenance runs
   * out - a wall-clock expiry on a perpetual licence, which is precisely what
   * rule 1 at the top of this file forbids and what `docs/pricing.md` calls a
   * bug of the highest severity. The honest default is the one that cannot take
   * away something already paid for: an unknown build is treated as covered.
   */
  const buildDate = Number.isFinite(context.buildDate) ? (context.buildDate as number) : 0;
  const buildMajor = Number.isFinite(context.buildMajor) ? (context.buildMajor as number) : 0;

  const base = {
    kind: payload.kind,
    seats: payload.seats,
    features: payload.features,
    payload,
    reason: 'ok' as const,
  };

  if (payload.kind === 'perpetual') {
    // A perpetual licence that carries a hard expiry is an issuing mistake we
    // refuse to enforce: read it as update coverage, which is what it can only
    // honestly have meant.
    const coverage = payload.maintenanceExpires ?? payload.expires;
    const dateCovered = coverage === null || buildDate <= coverage;
    const majorCovered = payload.major === 0 || buildMajor === 0 || buildMajor <= payload.major;

    if (dateCovered && majorCovered) {
      return {
        ...base,
        state: 'perpetual',
        plan: payload.plan,
        degraded: false,
        updatesCoveredUntil: coverage,
        buildCovered: true,
        graceEndsAt: null,
        graceDaysRemaining: 0,
        explanation:
          coverage === null
            ? 'Perpetual licence. This copy keeps working, offline, with no expiry.'
            : `Perpetual licence, with updates through ${isoDay(coverage)}. The version you have keeps working after that date.`,
      };
    }

    return {
      ...base,
      state: 'update-not-covered',
      plan: 'free',
      degraded: true,
      updatesCoveredUntil: coverage,
      buildCovered: false,
      graceEndsAt: null,
      graceDaysRemaining: 0,
      explanation: majorCovered
        ? `Your licence covers updates released up to ${isoDay(coverage ?? buildDate)}; this build is newer. Your licensed build still runs with everything you paid for - reinstall it, or renew updates to use this one.`
        : `Your licence covers version ${payload.major}; this is version ${buildMajor}. Version ${payload.major} still runs with everything you paid for.`,
    };
  }

  const graceDays = Number.isFinite(context.graceDays) ? (context.graceDays as number) : SUBSCRIPTION_GRACE_DAYS;
  const graceMs = Math.max(0, graceDays) * DAY_MS;

  // A subscription with no term is an issuing mistake; fail open and let it run.
  if (payload.expires === null) {
    return {
      ...base,
      state: 'active',
      plan: payload.plan,
      degraded: false,
      updatesCoveredUntil: payload.maintenanceExpires,
      buildCovered: true,
      graceEndsAt: null,
      graceDaysRemaining: 0,
      explanation: 'Subscription active.',
    };
  }

  const graceEndsAt = payload.expires + graceMs;

  if (now <= payload.expires) {
    return {
      ...base,
      state: 'active',
      plan: payload.plan,
      degraded: false,
      updatesCoveredUntil: payload.maintenanceExpires ?? payload.expires,
      buildCovered: true,
      graceEndsAt,
      graceDaysRemaining: daysBetween(now, graceEndsAt),
      explanation: `Subscription active through ${isoDay(payload.expires)}.`,
    };
  }

  if (now <= graceEndsAt) {
    return {
      ...base,
      state: 'grace',
      plan: payload.plan,
      degraded: false,
      updatesCoveredUntil: payload.maintenanceExpires ?? payload.expires,
      buildCovered: true,
      graceEndsAt,
      graceDaysRemaining: daysBetween(now, graceEndsAt),
      explanation: `We could not confirm your renewal, so everything stays on for another ${daysBetween(now, graceEndsAt)} days. No network check is required or performed.`,
    };
  }

  return {
    ...base,
    state: 'lapsed',
    plan: 'free',
    degraded: true,
    updatesCoveredUntil: payload.maintenanceExpires ?? payload.expires,
    buildCovered: true,
    graceEndsAt,
    graceDaysRemaining: 0,
    explanation:
      'Your subscription ended and the offline grace period is over, so this is now the free tier. Every file you have stays open, editable and saveable.',
  };
}

/**
 * Total: `Date#toISOString` throws on a timestamp outside the representable
 * range, and this runs inside the explanation of an assessment that promises
 * never to throw. The codec already refuses absurd dates, so this only ever
 * fires on a payload we ourselves mis-minted - which is when the app must still
 * come up.
 */
function isoDay(epochMs: number): string {
  if (!Number.isFinite(epochMs)) return 'an unknown date';
  try {
    return new Date(epochMs).toISOString().slice(0, 10);
  } catch {
    return 'an unknown date';
  }
}

/** Convenience for the licence-minting tool and for tests. */
export function makePayload(input: Partial<LicensePayload> & Pick<LicensePayload, 'id' | 'email' | 'plan'>): LicensePayload {
  return {
    version: 1,
    kind: 'perpetual',
    issued: Date.now(),
    expires: null,
    maintenanceExpires: null,
    seats: 1,
    features: [],
    major: 1,
    ...input,
  };
}
