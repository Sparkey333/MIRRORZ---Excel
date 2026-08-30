/**
 * A local, offline trial.
 *
 * The trial is recorded on disk and nowhere else. There is no trial server, no
 * device fingerprint, no hidden registry key and no phone-home: someone
 * determined to reset the trial can delete one file, and we have decided that is
 * fine. The alternative - fingerprinting machines and hiding state where users
 * cannot find it - treats every customer as a suspect to catch a handful of
 * people who were never going to pay.
 *
 * What we do defend against is the accidental and the casual: a clock nudged
 * backwards. The record carries a monotonic high-water mark, the highest
 * timestamp this installation has ever seen, and time is only ever read forward
 * from it. Rolling the system clock back to 2019 does not extend a trial; it just
 * makes the app act as if no time has passed.
 *
 * Every failure path FAILS OPEN. An unreadable file, a corrupt record, a
 * read-only disk, a full disk, a tampered digest - all of them resolve to "grant
 * access", never to "lock out". A person whose disk is failing needs their
 * spreadsheet more than we need $12, and in the worst case the trial ends where
 * every other path ends: at the free tier, which opens, edits and saves
 * everything.
 */

import { createHash, randomUUID } from 'node:crypto';

import { DAY_MS } from './codec.js';
import type { PlanId } from './plans.js';
import { isPlanId } from './plans.js';

export const DEFAULT_TRIAL_DAYS = 14;

/** Trials run at the top tier: a trial of a crippled version proves nothing. */
export const DEFAULT_TRIAL_PLAN: PlanId = 'pro';

/**
 * Tolerance before a backwards clock is called an anomaly rather than noise.
 * NTP steps, timezone-confused hardware clocks and virtual-machine suspends all
 * move time by small amounts in normal operation.
 */
export const CLOCK_SLACK_MS = 5 * 60_000;

const RECORD_VERSION = 1;

// Not a secret and not pretending to be one: it makes the digest specific to
// this file's purpose, so an unrelated JSON blob cannot be mistaken for a trial
// record. Anyone reading this source can forge the digest, which is why forging
// it is treated as an ordinary missing record rather than as an attack.
const DIGEST_SALT = 'mirrorz-trial-v1';

export interface TrialRecord {
  readonly version: number;
  readonly id: string;
  readonly plan: PlanId;
  readonly started: number;
  readonly expires: number;
  /** Highest timestamp ever observed by this installation. Never decreases. */
  readonly highWater: number;
}

/**
 * Storage is injected so the core stays pure and testable: the desktop app
 * passes a file-backed store, tests pass an in-memory one. Implementations may
 * throw; every caller here treats a throw as "no record".
 */
export interface TrialStore {
  read(): string | null;
  write(text: string): void;
}

export class MemoryTrialStore implements TrialStore {
  constructor(private text: string | null = null) {}

  read(): string | null {
    return this.text;
  }

  write(text: string): void {
    this.text = text;
  }
}

export type TrialStatus = 'none' | 'active' | 'expired';

export interface TrialState {
  readonly status: TrialStatus;
  readonly active: boolean;
  readonly plan: PlanId;
  readonly started: number | null;
  readonly expires: number | null;
  readonly daysRemaining: number;
  /** True when the clock was found behind the high-water mark. */
  readonly clockAnomaly: boolean;
  /** The timestamp actually used, after the monotonic floor is applied. */
  readonly effectiveNow: number;
  /** False when the record could not be written. The trial still runs. */
  readonly persisted: boolean;
  readonly record: TrialRecord | null;
  readonly explanation: string;
}

export interface TrialOptions {
  readonly now?: number;
  readonly days?: number;
  readonly plan?: PlanId;
}

function digest(record: TrialRecord): string {
  const canonical = [
    DIGEST_SALT,
    record.version,
    record.id,
    record.plan,
    record.started,
    record.expires,
    record.highWater,
  ].join('|');
  return createHash('sha256').update(canonical).digest('base64url').slice(0, 22);
}

export function serializeRecord(record: TrialRecord): string {
  return JSON.stringify({ ...record, digest: digest(record) });
}

/** Total: any malformed input is a missing record, which is the fail-open answer. */
export function parseRecord(text: string | null | undefined): TrialRecord | null {
  if (typeof text !== 'string' || text.trim().length === 0) return null;
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch {
    return null;
  }
  if (typeof raw !== 'object' || raw === null) return null;

  const value = raw as Record<string, unknown>;
  const record: TrialRecord = {
    version: typeof value.version === 'number' ? value.version : 0,
    id: typeof value.id === 'string' ? value.id : '',
    plan: isPlanId(value.plan) ? value.plan : DEFAULT_TRIAL_PLAN,
    started: typeof value.started === 'number' ? value.started : Number.NaN,
    expires: typeof value.expires === 'number' ? value.expires : Number.NaN,
    highWater: typeof value.highWater === 'number' ? value.highWater : Number.NaN,
  };

  if (record.version !== RECORD_VERSION) return null;
  if (!Number.isFinite(record.started) || !Number.isFinite(record.expires) || !Number.isFinite(record.highWater)) {
    return null;
  }
  if (record.id.length === 0) return null;
  if (value.digest !== digest(record)) return null;
  return record;
}

function readRecord(store: TrialStore): TrialRecord | null {
  try {
    return parseRecord(store.read());
  } catch {
    return null;
  }
}

function writeRecord(store: TrialStore, record: TrialRecord): boolean {
  try {
    store.write(serializeRecord(record));
    return true;
  } catch {
    // A trial that cannot be written is still a trial. Losing the write means we
    // may grant a few extra days; refusing to run would cost someone their work.
    return false;
  }
}

function daysLeft(from: number, to: number): number {
  return Math.max(0, Math.ceil((to - from) / DAY_MS));
}

const NO_TRIAL: TrialState = {
  status: 'none',
  active: false,
  plan: DEFAULT_TRIAL_PLAN,
  started: null,
  expires: null,
  daysRemaining: 0,
  clockAnomaly: false,
  effectiveNow: 0,
  persisted: true,
  record: null,
  explanation: 'No trial started.',
};

/**
 * Evaluate a record against a clock, without touching storage. The monotonic
 * floor is applied here, so this is the single place where "what time is it"
 * gets decided for the whole trial.
 */
export function evaluateTrial(record: TrialRecord | null, now: number): TrialState {
  if (record === null) return { ...NO_TRIAL, effectiveNow: now };

  const clockAnomaly = now < record.highWater - CLOCK_SLACK_MS;
  const effectiveNow = Math.max(now, record.highWater);
  const active = effectiveNow < record.expires;
  const remaining = daysLeft(effectiveNow, record.expires);

  return {
    status: active ? 'active' : 'expired',
    active,
    plan: record.plan,
    started: record.started,
    expires: record.expires,
    daysRemaining: remaining,
    clockAnomaly,
    effectiveNow,
    persisted: true,
    record,
    explanation: active
      ? `Trial of the ${record.plan} features: ${remaining} day${remaining === 1 ? '' : 's'} left. No account, no network check.`
      : 'The trial has ended. Everything you have made stays open, editable and saveable on the free tier.',
  };
}

/**
 * Start a trial, or return the existing one. Starting twice is not an error and
 * never resets the clock; the first record wins for as long as it survives.
 */
export function startTrial(store: TrialStore, options: TrialOptions = {}): TrialState {
  const now = options.now ?? Date.now();
  const existing = readRecord(store);
  if (existing !== null) return checkTrial(store, { ...options, now });

  const days = options.days ?? DEFAULT_TRIAL_DAYS;
  const record: TrialRecord = {
    version: RECORD_VERSION,
    id: randomUUID(),
    plan: options.plan ?? DEFAULT_TRIAL_PLAN,
    started: now,
    expires: now + days * DAY_MS,
    highWater: now,
  };
  const persisted = writeRecord(store, record);
  return { ...evaluateTrial(record, now), persisted };
}

/**
 * Read the trial and advance its high-water mark. Call this on every launch: the
 * mark is what makes a clock rollback inert, and it only moves forward.
 */
export function checkTrial(store: TrialStore, options: TrialOptions = {}): TrialState {
  const now = options.now ?? Date.now();
  const record = readRecord(store);
  if (record === null) return { ...NO_TRIAL, effectiveNow: now };

  const state = evaluateTrial(record, now);
  if (state.effectiveNow > record.highWater) {
    const advanced: TrialRecord = { ...record, highWater: state.effectiveNow };
    const persisted = writeRecord(store, advanced);
    return { ...evaluateTrial(advanced, now), persisted };
  }
  return state;
}

/**
 * The high-water mark, for {@link RuntimeContext.clockFloor} in license.ts, so a
 * subscription cannot be revived by moving the clock back either. Returns
 * undefined when there is no trial record, in which case the wall clock is all
 * we have and we trust it.
 */
export function clockFloor(store: TrialStore): number | undefined {
  const record = readRecord(store);
  return record === null ? undefined : record.highWater;
}
