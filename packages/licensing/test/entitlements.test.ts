import { describe, expect, it } from 'vitest';
import type { Capability } from '../src/entitlements.js';
import {
  NEVER_GATED,
  PLAN_CAPABILITIES,
  PREMIUM,
  capabilitiesFor,
  entitlementsFor,
  freeEntitlements,
} from '../src/entitlements.js';
import { assessLicense, generateKeyPair, makePayload, signLicense } from '../src/license.js';
import { PLANS } from '../src/plans.js';
import { MemoryTrialStore, evaluateTrial, startTrial } from '../src/trial.js';

const KEYS = generateKeyPair();
const NOW = Date.UTC(2026, 5, 1);

function licenseFor(plan: 'personal' | 'pro' | 'team'): string {
  return signLicense(
    makePayload({ id: `MZ-${plan}`, email: 'a@b.c', plan, kind: 'perpetual', issued: NOW, maintenanceExpires: null, major: 0 }),
    KEYS.privateKey,
  );
}

function assess(text: string | null): ReturnType<typeof assessLicense> {
  return assessLicense(text, KEYS.publicKey, { now: NOW, buildDate: NOW, buildMajor: 1 });
}

describe('the free tier is a real tier', () => {
  it.each(PLANS)('gives %s every never-gated capability', (plan) => {
    for (const capability of NEVER_GATED) {
      expect(PLAN_CAPABILITIES[plan].has(capability)).toBe(true);
    }
  });

  it.each(['file.open', 'file.save', 'file.saveAs', 'edit.cells', 'edit.structure', 'formula.calculate'] as Capability[])(
    'never gates %s',
    (capability) => {
      expect(capabilitiesFor('free').has(capability)).toBe(true);
    },
  );

  it('lets the free tier print and export', () => {
    const free = capabilitiesFor('free');
    expect(free.has('file.print')).toBe(true);
    expect(free.has('export.csv')).toBe(true);
    expect(free.has('export.pdf')).toBe(true);
  });

  it('lets the free tier preserve macros and pivots it cannot run', () => {
    const free = capabilitiesFor('free');
    expect(free.has('macro.preserve')).toBe(true);
    expect(free.has('pivot.preserve')).toBe(true);
    expect(free.has('macro.execute')).toBe(false);
  });

  it('leaves the local scripting runtime ungated', () => {
    expect(capabilitiesFor('free').has('automation.scripts')).toBe(true);
  });

  it('never sells anything that is on the never-gated list', () => {
    for (const gate of PREMIUM) {
      expect(NEVER_GATED).not.toContain(gate.capability);
    }
  });

  it('gives free exactly the never-gated set and nothing else', () => {
    expect(capabilitiesFor('free').size).toBe(NEVER_GATED.length);
  });
});

describe('the premium catalogue is explicit', () => {
  it('states a reason for every gated capability', () => {
    for (const gate of PREMIUM) {
      expect(gate.why.length).toBeGreaterThan(30);
    }
  });

  it('lists each capability once', () => {
    const names = PREMIUM.map((gate) => gate.capability);
    expect(new Set(names).size).toBe(names.length);
  });

  it('assigns every gate to a paid plan', () => {
    for (const gate of PREMIUM) {
      expect(gate.plan).not.toBe('free');
    }
  });
});

describe('plans stack', () => {
  it('gives personal charts and history but not macro execution', () => {
    const personal = capabilitiesFor('personal');
    expect(personal.has('chart.author')).toBe(true);
    expect(personal.has('history.local')).toBe(true);
    expect(personal.has('macro.execute')).toBe(false);
  });

  it('gives pro the expensive engineering', () => {
    const pro = capabilitiesFor('pro');
    for (const capability of ['macro.execute', 'pivot.interactive', 'diff.semantic', 'merge.threeWay', 'formula.inspector', 'batch.cli'] as Capability[]) {
      expect(pro.has(capability)).toBe(true);
    }
  });

  it('gives team everything pro has', () => {
    for (const capability of capabilitiesFor('pro')) {
      expect(capabilitiesFor('team').has(capability)).toBe(true);
    }
  });

  it('adds seat management only at team', () => {
    expect(capabilitiesFor('pro').has('seats.manage')).toBe(false);
    expect(capabilitiesFor('team').has('seats.manage')).toBe(true);
  });

  it('grants a per-licence extra beyond the plan', () => {
    expect(capabilitiesFor('personal', ['macro.execute']).has('macro.execute')).toBe(true);
  });

  it('ignores an unknown extra rather than failing', () => {
    expect(() => capabilitiesFor('personal', ['not.a.capability'])).not.toThrow();
    expect(capabilitiesFor('personal', ['not.a.capability']).has('chart.author')).toBe(true);
  });
});

describe('resolving entitlements', () => {
  it('gives the licensed plan', () => {
    const entitlements = entitlementsFor(assess(licenseFor('pro')));
    expect(entitlements.plan).toBe('pro');
    expect(entitlements.source).toBe('license');
    expect(entitlements.can('macro.execute')).toBe(true);
  });

  it('falls back to free with no licence, still able to save', () => {
    const entitlements = entitlementsFor(assess(null));
    expect(entitlements.plan).toBe('free');
    expect(entitlements.can('file.save')).toBe(true);
    expect(entitlements.can('macro.execute')).toBe(false);
  });

  it('falls back to free on a tampered licence, still able to save', () => {
    const other = generateKeyPair();
    const forged = signLicense(makePayload({ id: 'X', email: 'a@b.c', plan: 'team' }), other.privateKey);
    const entitlements = entitlementsFor(assess(forged));
    expect(entitlements.plan).toBe('free');
    expect(entitlements.can('file.open')).toBe(true);
    expect(entitlements.can('file.saveAs')).toBe(true);
  });

  it('uses the trial when there is no licence', () => {
    const store = new MemoryTrialStore();
    const trial = startTrial(store, { now: NOW });
    const entitlements = entitlementsFor(assess(null), trial);
    expect(entitlements.plan).toBe('pro');
    expect(entitlements.source).toBe('trial');
  });

  it('ignores an expired trial', () => {
    const store = new MemoryTrialStore();
    startTrial(store, { now: NOW, days: 1 });
    const trial = evaluateTrial(null, NOW);
    const entitlements = entitlementsFor(assess(null), trial);
    expect(entitlements.plan).toBe('free');
  });

  it('keeps the higher of licence and trial', () => {
    const store = new MemoryTrialStore();
    const trial = startTrial(store, { now: NOW, plan: 'pro' });
    const entitlements = entitlementsFor(assess(licenseFor('personal')), trial);
    expect(entitlements.plan).toBe('pro');
    expect(entitlements.source).toBe('trial');
  });

  it('does not downgrade a licence to the trial plan', () => {
    const store = new MemoryTrialStore();
    const trial = startTrial(store, { now: NOW, plan: 'personal' });
    const entitlements = entitlementsFor(assess(licenseFor('pro')), trial);
    expect(entitlements.plan).toBe('pro');
    expect(entitlements.source).toBe('license');
  });

  it('explains why a capability is unavailable', () => {
    const entitlements = entitlementsFor(assess(null));
    const gate = entitlements.gateFor('macro.execute');
    expect(gate?.plan).toBe('pro');
    expect(gate?.why).toContain('VBA');
  });

  it('returns no gate for something already granted', () => {
    expect(entitlementsFor(assess(null)).gateFor('file.save')).toBeNull();
  });

  it('carries the seat count from the licence', () => {
    const key = signLicense(
      makePayload({ id: 'MZ-SEATS', email: 'a@b.c', plan: 'team', kind: 'subscription', issued: NOW, expires: NOW + 31_536_000_000, seats: 5 }),
      KEYS.privateKey,
    );
    expect(entitlementsFor(assess(key)).seats).toBe(5);
  });

  it('offers a free entitlement for the boot path', () => {
    const free = freeEntitlements();
    expect(free.plan).toBe('free');
    expect(free.can('file.open')).toBe(true);
    expect(free.explanation).toContain('never restricted');
  });
});
