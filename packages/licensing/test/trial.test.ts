import { describe, expect, it } from 'vitest';
import { DAY_MS } from '../src/codec.js';
import {
  CLOCK_SLACK_MS,
  DEFAULT_TRIAL_DAYS,
  DEFAULT_TRIAL_PLAN,
  MemoryTrialStore,
  checkTrial,
  clockFloor,
  evaluateTrial,
  parseRecord,
  serializeRecord,
  startTrial,
} from '../src/trial.js';
import type { TrialRecord, TrialStore } from '../src/trial.js';

const T0 = Date.UTC(2026, 2, 1);

class ThrowingStore implements TrialStore {
  constructor(
    private readonly onRead: boolean,
    private readonly onWrite: boolean,
    private text: string | null = null,
  ) {}

  read(): string | null {
    if (this.onRead) throw new Error('disk is on fire');
    return this.text;
  }

  write(text: string): void {
    if (this.onWrite) throw new Error('read-only filesystem');
    this.text = text;
  }
}

describe('starting a trial', () => {
  it('starts active', () => {
    const state = startTrial(new MemoryTrialStore(), { now: T0 });
    expect(state.status).toBe('active');
    expect(state.active).toBe(true);
  });

  it('runs at the top tier, because a crippled trial proves nothing', () => {
    expect(DEFAULT_TRIAL_PLAN).toBe('pro');
    expect(startTrial(new MemoryTrialStore(), { now: T0 }).plan).toBe('pro');
  });

  it('lasts the default number of days', () => {
    const state = startTrial(new MemoryTrialStore(), { now: T0 });
    expect(state.expires).toBe(T0 + DEFAULT_TRIAL_DAYS * DAY_MS);
    expect(state.daysRemaining).toBe(DEFAULT_TRIAL_DAYS);
  });

  it('honours a custom length and plan', () => {
    const state = startTrial(new MemoryTrialStore(), { now: T0, days: 3, plan: 'personal' });
    expect(state.daysRemaining).toBe(3);
    expect(state.plan).toBe('personal');
  });

  it('persists the record', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0 });
    expect(parseRecord(store.read())).not.toBeNull();
  });

  it('does not restart an existing trial', () => {
    const store = new MemoryTrialStore();
    const first = startTrial(store, { now: T0 });
    const second = startTrial(store, { now: T0 + 5 * DAY_MS });
    expect(second.started).toBe(first.started);
    expect(second.expires).toBe(first.expires);
    expect(second.daysRemaining).toBe(DEFAULT_TRIAL_DAYS - 5);
  });

  it('reports no trial before one is started', () => {
    const state = checkTrial(new MemoryTrialStore(), { now: T0 });
    expect(state.status).toBe('none');
    expect(state.active).toBe(false);
    expect(state.record).toBeNull();
  });
});

describe('trial expiry', () => {
  it('expires after the last day', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0, days: 14 });
    const state = checkTrial(store, { now: T0 + 15 * DAY_MS });
    expect(state.status).toBe('expired');
    expect(state.active).toBe(false);
  });

  it('is still active one hour before the end', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0, days: 14 });
    expect(checkTrial(store, { now: T0 + 14 * DAY_MS - 3_600_000 }).active).toBe(true);
  });

  it('describes the end as a downgrade, not a lockout', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0, days: 1 });
    const state = checkTrial(store, { now: T0 + 10 * DAY_MS });
    expect(state.explanation).toContain('free tier');
    expect(state.explanation).toContain('saveable');
  });

  it('reports zero days remaining once over', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0, days: 1 });
    expect(checkTrial(store, { now: T0 + 10 * DAY_MS }).daysRemaining).toBe(0);
  });
});

describe('clock rollback', () => {
  it('records the highest timestamp ever seen', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0 });
    checkTrial(store, { now: T0 + 5 * DAY_MS });
    expect(clockFloor(store)).toBe(T0 + 5 * DAY_MS);
  });

  it('does not give days back when the clock moves backwards', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0, days: 14 });
    checkTrial(store, { now: T0 + 10 * DAY_MS });
    const rolledBack = checkTrial(store, { now: T0 - 300 * DAY_MS });
    expect(rolledBack.daysRemaining).toBe(4);
    expect(rolledBack.effectiveNow).toBe(T0 + 10 * DAY_MS);
  });

  it('cannot resurrect an expired trial', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0, days: 7 });
    checkTrial(store, { now: T0 + 30 * DAY_MS });
    const rolledBack = checkTrial(store, { now: T0 + DAY_MS });
    expect(rolledBack.active).toBe(false);
  });

  it('flags the rollback so the app can say something honest about it', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0 });
    checkTrial(store, { now: T0 + 5 * DAY_MS });
    expect(checkTrial(store, { now: T0 }).clockAnomaly).toBe(true);
  });

  it('tolerates small clock corrections without calling them an anomaly', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0 });
    checkTrial(store, { now: T0 + 5 * DAY_MS });
    const nudged = checkTrial(store, { now: T0 + 5 * DAY_MS - CLOCK_SLACK_MS / 2 });
    expect(nudged.clockAnomaly).toBe(false);
  });

  it('never moves the high-water mark backwards', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0 });
    checkTrial(store, { now: T0 + 9 * DAY_MS });
    checkTrial(store, { now: T0 - 900 * DAY_MS });
    expect(clockFloor(store)).toBe(T0 + 9 * DAY_MS);
  });

  it('advances the mark as ordinary time passes', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0 });
    checkTrial(store, { now: T0 + DAY_MS });
    checkTrial(store, { now: T0 + 2 * DAY_MS });
    expect(clockFloor(store)).toBe(T0 + 2 * DAY_MS);
  });
});

describe('the trial fails open', () => {
  it('treats corrupt JSON as no trial rather than as an expired one', () => {
    const state = checkTrial(new MemoryTrialStore('{not json'), { now: T0 });
    expect(state.status).toBe('none');
  });

  it('lets a fresh trial start after a corrupt record', () => {
    const store = new MemoryTrialStore('garbage');
    expect(startTrial(store, { now: T0 }).active).toBe(true);
  });

  it('treats a hand-edited record as no trial', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0, days: 7 });
    const tampered = JSON.parse(store.read() ?? '{}') as Record<string, unknown>;
    tampered.expires = T0 + 4000 * DAY_MS;
    store.write(JSON.stringify(tampered));
    expect(checkTrial(store, { now: T0 }).status).toBe('none');
  });

  it('rejects a record whose digest was removed', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0 });
    const record = JSON.parse(store.read() ?? '{}') as Record<string, unknown>;
    delete record.digest;
    expect(parseRecord(JSON.stringify(record))).toBeNull();
  });

  it('rejects a record from an unknown version', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: T0 });
    const record = JSON.parse(store.read() ?? '{}') as Record<string, unknown>;
    record.version = 99;
    expect(parseRecord(JSON.stringify(record))).toBeNull();
  });

  it('survives a store that cannot be read', () => {
    const state = checkTrial(new ThrowingStore(true, false), { now: T0 });
    expect(state.status).toBe('none');
  });

  it('grants the trial even when the record cannot be written', () => {
    const state = startTrial(new ThrowingStore(false, true), { now: T0 });
    expect(state.active).toBe(true);
    expect(state.persisted).toBe(false);
  });

  it('keeps working when the high-water write fails mid-trial', () => {
    const store = new ThrowingStore(false, false);
    startTrial(store, { now: T0 });
    const failing = new ThrowingStore(false, true, store.read());
    expect(checkTrial(failing, { now: T0 + DAY_MS }).active).toBe(true);
  });

  it('never throws for any stored text', () => {
    for (const text of ['', 'null', '[]', '{}', '{"version":1}', 'ï¿½']) {
      expect(() => checkTrial(new MemoryTrialStore(text), { now: T0 })).not.toThrow();
    }
  });

  it('reports no clock floor when there is no record', () => {
    expect(clockFloor(new MemoryTrialStore())).toBeUndefined();
  });
});

describe('record serialization', () => {
  const record: TrialRecord = {
    version: 1,
    id: 'a2b1c0d9-0000-4000-8000-000000000000',
    plan: 'pro',
    started: T0,
    expires: T0 + 14 * DAY_MS,
    highWater: T0,
  };

  it('round trips', () => {
    expect(parseRecord(serializeRecord(record))).toEqual(record);
  });

  it('rejects null and empty text', () => {
    expect(parseRecord(null)).toBeNull();
    expect(parseRecord('')).toBeNull();
  });

  it('rejects a record with non-numeric dates', () => {
    const broken = JSON.parse(serializeRecord(record)) as Record<string, unknown>;
    broken.expires = 'soon';
    expect(parseRecord(JSON.stringify(broken))).toBeNull();
  });
});

describe('evaluateTrial is pure', () => {
  it('reports none for a missing record', () => {
    const state = evaluateTrial(null, T0);
    expect(state.status).toBe('none');
    expect(state.effectiveNow).toBe(T0);
  });

  it('does not touch storage', () => {
    const record: TrialRecord = {
      version: 1,
      id: 'x',
      plan: 'personal',
      started: T0,
      expires: T0 + DAY_MS,
      highWater: T0,
    };
    expect(evaluateTrial(record, T0).plan).toBe('personal');
    expect(evaluateTrial(record, T0 + 2 * DAY_MS).active).toBe(false);
  });
});
