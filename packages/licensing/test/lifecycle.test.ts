/**
 * End-to-end: the states a real installation moves through, and the invariant
 * that must hold in every one of them - the user can always open, edit and save.
 */
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, describe, expect, it } from 'vitest';
import { DAY_MS } from '../src/codec.js';
import { entitlementsFor } from '../src/entitlements.js';
import type { Capability } from '../src/entitlements.js';
import { SUBSCRIPTION_GRACE_DAYS, assessLicense, generateKeyPair, makePayload, signLicense } from '../src/license.js';
import { FileTextStore, configDir, licensePath, trialPath } from '../src/store.js';
import { MemoryTrialStore, checkTrial, clockFloor, startTrial } from '../src/trial.js';

const KEYS = generateKeyPair();
const START = Date.UTC(2026, 0, 1);
const DATA_CAPABILITIES: Capability[] = ['file.open', 'file.save', 'file.saveAs', 'edit.cells', 'formula.calculate'];

const roots: string[] = [];
function scratch(): string {
  const dir = mkdtempSync(join(tmpdir(), 'mirrorz-licensing-'));
  roots.push(dir);
  return dir;
}

afterAll(() => {
  for (const dir of roots) rmSync(dir, { recursive: true, force: true });
});

const subscription = signLicense(
  makePayload({
    id: 'MZ-LIFE',
    email: 'owner@example.com',
    plan: 'pro',
    kind: 'subscription',
    issued: START,
    expires: START + 365 * DAY_MS,
    maintenanceExpires: START + 365 * DAY_MS,
  }),
  KEYS.privateKey,
);

function entitlementsAt(text: string | null, now: number, trialStore?: MemoryTrialStore) {
  const trial = trialStore ? checkTrial(trialStore, { now }) : null;
  const assessment = assessLicense(text, KEYS.publicKey, {
    now,
    buildDate: START,
    buildMajor: 1,
    clockFloor: trialStore ? clockFloor(trialStore) : undefined,
  });
  return entitlementsFor(assessment, trial);
}

describe('the whole lifecycle', () => {
  it('starts free and fully able to work', () => {
    const entitlements = entitlementsAt(null, START);
    expect(entitlements.plan).toBe('free');
    for (const capability of DATA_CAPABILITIES) expect(entitlements.can(capability)).toBe(true);
  });

  it('upgrades to pro during a trial', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: START });
    expect(entitlementsAt(null, START + DAY_MS, store).plan).toBe('pro');
  });

  it('returns to free when the trial ends, without losing file access', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: START, days: 14 });
    const entitlements = entitlementsAt(null, START + 30 * DAY_MS, store);
    expect(entitlements.plan).toBe('free');
    for (const capability of DATA_CAPABILITIES) expect(entitlements.can(capability)).toBe(true);
  });

  it('becomes pro when a subscription is entered', () => {
    expect(entitlementsAt(subscription, START + 100 * DAY_MS).plan).toBe('pro');
  });

  it('stays pro through the offline grace window', () => {
    expect(entitlementsAt(subscription, START + (365 + 30) * DAY_MS).plan).toBe('pro');
  });

  it('returns to free after grace, with the data capabilities intact', () => {
    const entitlements = entitlementsAt(subscription, START + (365 + SUBSCRIPTION_GRACE_DAYS + 2) * DAY_MS);
    expect(entitlements.plan).toBe('free');
    for (const capability of DATA_CAPABILITIES) expect(entitlements.can(capability)).toBe(true);
  });

  it('cannot be revived by winding the clock back', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: START });
    checkTrial(store, { now: START + (365 + SUBSCRIPTION_GRACE_DAYS + 10) * DAY_MS });
    expect(entitlementsAt(subscription, START + 10 * DAY_MS, store).plan).toBe('free');
  });

  it('keeps the data capabilities in every state it can reach', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: START, days: 7 });
    const moments = [START, START + 3 * DAY_MS, START + 200 * DAY_MS, START + 1000 * DAY_MS];
    const licences = [null, 'nonsense', subscription];
    for (const now of moments) {
      for (const licence of licences) {
        const entitlements = entitlementsAt(licence, now, store);
        for (const capability of DATA_CAPABILITIES) expect(entitlements.can(capability)).toBe(true);
      }
    }
  });
});

describe('file-backed storage', () => {
  it('reads a missing file as nothing', () => {
    const store = new FileTextStore(join(scratch(), 'nope.key'));
    expect(store.read()).toBeNull();
  });

  it('round trips a licence key through disk', () => {
    const store = new FileTextStore(join(scratch(), 'license.key'));
    store.write(subscription);
    expect(store.read()).toBe(subscription);
    expect(assessLicense(store.read(), KEYS.publicKey, { now: START }).state).toBe('active');
  });

  it('creates the directory it needs', () => {
    const store = new FileTextStore(join(scratch(), 'nested', 'deeper', 'trial.json'));
    store.write('{}');
    expect(store.read()).toBe('{}');
  });

  it('carries a trial across restarts', () => {
    const store = new FileTextStore(join(scratch(), 'trial.json'));
    startTrial(store, { now: START, days: 5 });
    const reopened = new FileTextStore(store.path);
    expect(checkTrial(reopened, { now: START + DAY_MS }).daysRemaining).toBe(4);
  });

  it('lets the user delete their licence', () => {
    const store = new FileTextStore(join(scratch(), 'license.key'));
    store.write(subscription);
    store.clear();
    expect(store.read()).toBeNull();
  });

  it('treats an unreadable directory-as-file as nothing rather than throwing', () => {
    const dir = scratch();
    const store = new FileTextStore(dir);
    expect(store.read()).toBeNull();
  });

  it('reads a hand-pasted licence file with stray whitespace', () => {
    const path = join(scratch(), 'license.key');
    writeFileSync(path, `\n  ${subscription}  \n`, 'utf8');
    expect(assessLicense(new FileTextStore(path).read(), KEYS.publicKey, { now: START }).state).toBe('active');
  });

  it('puts both files in the same visible per-user config directory', () => {
    expect(licensePath()).toContain(configDir());
    expect(trialPath()).toContain(configDir());
  });
});
