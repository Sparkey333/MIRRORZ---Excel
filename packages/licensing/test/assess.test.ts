import { describe, expect, it } from 'vitest';
import { DAY_MS } from '../src/codec.js';
import type { LicensePayload } from '../src/codec.js';
import {
  SUBSCRIPTION_GRACE_DAYS,
  assessLicense,
  generateKeyPair,
  makePayload,
  signLicense,
} from '../src/license.js';

const KEYS = generateKeyPair();
const OTHER = generateKeyPair();

const JAN_2026 = Date.UTC(2026, 0, 1);
const JAN_2027 = Date.UTC(2027, 0, 1);
const JAN_2030 = Date.UTC(2030, 0, 1);

function key(payload: Partial<LicensePayload> & Pick<LicensePayload, 'id' | 'email' | 'plan'>): string {
  return signLicense(makePayload({ issued: JAN_2026, ...payload }), KEYS.privateKey);
}

const perpetual = key({
  id: 'MZ-PERP',
  email: 'owner@example.com',
  plan: 'pro',
  kind: 'perpetual',
  maintenanceExpires: JAN_2027,
  major: 1,
});

const subscription = key({
  id: 'MZ-SUB',
  email: 'owner@example.com',
  plan: 'pro',
  kind: 'subscription',
  expires: JAN_2027,
  maintenanceExpires: JAN_2027,
});

describe('a perpetual licence never stops working', () => {
  it('is active for a build released inside the coverage window', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, {
      now: JAN_2026 + 30 * DAY_MS,
      buildDate: JAN_2026 + 20 * DAY_MS,
      buildMajor: 1,
    });
    expect(result.state).toBe('perpetual');
    expect(result.plan).toBe('pro');
  });

  it('is still active decades after maintenance ended', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, {
      now: Date.UTC(2099, 0, 1),
      buildDate: JAN_2026 + 20 * DAY_MS,
      buildMajor: 1,
    });
    expect(result.state).toBe('perpetual');
    expect(result.plan).toBe('pro');
    expect(result.degraded).toBe(false);
  });

  it('ignores the wall clock entirely', () => {
    const context = { buildDate: JAN_2026, buildMajor: 1 } as const;
    const early = assessLicense(perpetual, KEYS.publicKey, { ...context, now: JAN_2026 });
    const late = assessLicense(perpetual, KEYS.publicKey, { ...context, now: JAN_2030 });
    expect(early.state).toBe(late.state);
    expect(early.plan).toBe(late.plan);
  });

  it('has no grace period because it has nothing to lapse from', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, { now: JAN_2030, buildDate: JAN_2026, buildMajor: 1 });
    expect(result.graceEndsAt).toBeNull();
  });

  it('reports when updates are covered until', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, { now: JAN_2026, buildDate: JAN_2026, buildMajor: 1 });
    expect(result.updatesCoveredUntil).toBe(JAN_2027);
    expect(result.explanation).toContain('2027-01-01');
  });

  it('runs forever with no maintenance date at all', () => {
    const forever = key({ id: 'MZ-F', email: 'a@b.c', plan: 'pro', kind: 'perpetual', maintenanceExpires: null, major: 0 });
    const result = assessLicense(forever, KEYS.publicKey, { now: JAN_2030, buildDate: JAN_2030, buildMajor: 9 });
    expect(result.state).toBe('perpetual');
    expect(result.explanation).toContain('no expiry');
  });

  it('reads a hard expiry on a perpetual licence as update coverage instead', () => {
    // An issuing mistake we refuse to enforce against the customer.
    const mistake = key({
      id: 'MZ-MISTAKE',
      email: 'a@b.c',
      plan: 'pro',
      kind: 'perpetual',
      expires: JAN_2027,
      maintenanceExpires: null,
      major: 1,
    });
    const result = assessLicense(mistake, KEYS.publicKey, { now: JAN_2030, buildDate: JAN_2026, buildMajor: 1 });
    expect(result.state).toBe('perpetual');
    expect(result.plan).toBe('pro');
  });
});

describe('the default context cannot revoke a perpetual licence', () => {
  // Regression: `buildDate` used to default to the wall clock, which put a
  // wall-clock expiry on exactly the licence that is sold as having none. Any
  // caller that forgot to bake a build date in - and nothing in the repo passes
  // one yet - downgraded every perpetual customer on their maintenance date.
  it('stays perpetual decades later when no build date is baked in', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, { now: Date.UTC(2099, 0, 1) });
    expect(result.state).toBe('perpetual');
    expect(result.plan).toBe('pro');
    expect(result.degraded).toBe(false);
  });

  it('stays perpetual on the day after maintenance ends with no build date', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, { now: JAN_2027 + DAY_MS });
    expect(result.state).toBe('perpetual');
    expect(result.plan).toBe('pro');
  });

  it('is not swayed by a clock floor either', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, {
      now: JAN_2026,
      clockFloor: Date.UTC(2099, 0, 1),
    });
    expect(result.state).toBe('perpetual');
  });

  it('treats an unknown build major as covered', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, { now: JAN_2030, buildDate: JAN_2026 });
    expect(result.state).toBe('perpetual');
  });
});

describe('a broken clock never costs a customer their plan', () => {
  it.each([Number.NaN, Number.POSITIVE_INFINITY, undefined])(
    'falls back to the real clock when now is %p',
    (now) => {
      const result = assessLicense(subscription, KEYS.publicKey, {
        now: now as number | undefined,
        clockFloor: JAN_2026 + 10 * DAY_MS,
      });
      // The floor is inside the term, so whatever the wall clock says the licence
      // must not read as lapsed on account of an unusable `now`.
      expect(['active', 'grace', 'lapsed']).toContain(result.state);
      expect(result.reason).toBe('ok');
    },
  );

  it('ignores a non-finite clock floor rather than poisoning the comparison', () => {
    const result = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2026 + 100 * DAY_MS,
      clockFloor: Number.NaN,
    });
    expect(result.state).toBe('active');
  });

  it('ignores a non-finite grace override', () => {
    const result = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2027 + DAY_MS,
      graceDays: Number.NaN,
    });
    expect(result.state).toBe('grace');
  });

  it('never throws on a non-finite clock', () => {
    for (const now of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      expect(() => assessLicense(perpetual, KEYS.publicKey, { now, buildDate: now })).not.toThrow();
    }
  });
});

describe('maintenance limits updates, not the software', () => {
  it('drops a build released after coverage ended to the free tier', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, {
      now: JAN_2027 + 10 * DAY_MS,
      buildDate: JAN_2027 + 5 * DAY_MS,
      buildMajor: 1,
    });
    expect(result.state).toBe('update-not-covered');
    expect(result.plan).toBe('free');
    expect(result.buildCovered).toBe(false);
    expect(result.degraded).toBe(true);
  });

  it('says plainly that the licensed build still runs', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, {
      now: JAN_2030,
      buildDate: JAN_2030,
      buildMajor: 1,
    });
    expect(result.explanation).toContain('still runs');
  });

  it('keeps the payload so the app can show what was bought', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, { now: JAN_2030, buildDate: JAN_2030, buildMajor: 1 });
    expect(result.payload?.id).toBe('MZ-PERP');
  });

  it('covers a build released exactly on the maintenance date', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, { now: JAN_2027, buildDate: JAN_2027, buildMajor: 1 });
    expect(result.state).toBe('perpetual');
  });

  it('does not cover a later major version', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, { now: JAN_2026, buildDate: JAN_2026, buildMajor: 2 });
    expect(result.state).toBe('update-not-covered');
    expect(result.explanation).toContain('version 1');
  });

  it('covers any major version when the licence says so', () => {
    const any = key({ id: 'MZ-ANY', email: 'a@b.c', plan: 'pro', kind: 'perpetual', maintenanceExpires: null, major: 0 });
    const result = assessLicense(any, KEYS.publicKey, { now: JAN_2030, buildDate: JAN_2030, buildMajor: 7 });
    expect(result.state).toBe('perpetual');
  });
});

describe('subscription grace', () => {
  it('is active inside the term', () => {
    const result = assessLicense(subscription, KEYS.publicKey, { now: JAN_2026 + 100 * DAY_MS });
    expect(result.state).toBe('active');
    expect(result.plan).toBe('pro');
  });

  it('is active on the last day of the term', () => {
    expect(assessLicense(subscription, KEYS.publicKey, { now: JAN_2027 }).state).toBe('active');
  });

  it('keeps every feature for the whole grace window', () => {
    const result = assessLicense(subscription, KEYS.publicKey, { now: JAN_2027 + 30 * DAY_MS });
    expect(result.state).toBe('grace');
    expect(result.plan).toBe('pro');
    expect(result.degraded).toBe(false);
  });

  it('grants a long grace window by default', () => {
    expect(SUBSCRIPTION_GRACE_DAYS).toBeGreaterThanOrEqual(30);
  });

  it('is still in grace on the final grace day', () => {
    const result = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2027 + SUBSCRIPTION_GRACE_DAYS * DAY_MS,
    });
    expect(result.state).toBe('grace');
  });

  it('counts down the days remaining', () => {
    const result = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2027 + (SUBSCRIPTION_GRACE_DAYS - 10) * DAY_MS,
    });
    expect(result.graceDaysRemaining).toBe(10);
  });

  it('explains that no network check is performed', () => {
    const result = assessLicense(subscription, KEYS.publicKey, { now: JAN_2027 + DAY_MS });
    expect(result.explanation).toContain('No network check');
  });

  it('falls back to the free tier once grace is over', () => {
    const result = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2027 + (SUBSCRIPTION_GRACE_DAYS + 1) * DAY_MS,
    });
    expect(result.state).toBe('lapsed');
    expect(result.plan).toBe('free');
  });

  it('promises the files stay editable when it lapses', () => {
    const result = assessLicense(subscription, KEYS.publicKey, { now: JAN_2030 });
    expect(result.explanation).toContain('editable');
  });

  it('honours a shorter grace window when configured', () => {
    const result = assessLicense(subscription, KEYS.publicKey, { now: JAN_2027 + 5 * DAY_MS, graceDays: 2 });
    expect(result.state).toBe('lapsed');
  });

  it('fails open on a subscription minted with no term', () => {
    const broken = key({ id: 'MZ-NOTERM', email: 'a@b.c', plan: 'pro', kind: 'subscription', expires: null });
    expect(assessLicense(broken, KEYS.publicKey, { now: JAN_2030 }).state).toBe('active');
  });
});

describe('clock rollback', () => {
  it('cannot revive a lapsed subscription', () => {
    const rolledBack = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2026 + 10 * DAY_MS,
      clockFloor: JAN_2030,
    });
    expect(rolledBack.state).toBe('lapsed');
  });

  it('cannot extend the grace window', () => {
    const result = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2027,
      clockFloor: JAN_2027 + (SUBSCRIPTION_GRACE_DAYS + 5) * DAY_MS,
    });
    expect(result.state).toBe('lapsed');
  });

  it('only ever moves time forward, so a stale floor is harmless', () => {
    const result = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2026 + 200 * DAY_MS,
      clockFloor: JAN_2026,
    });
    expect(result.state).toBe('active');
  });

  it('says the clock is why, rather than implying non-payment', () => {
    // The high-water mark is sticky: one launch with a broken clock poisons it
    // permanently. A customer whose subscription is fine must be told which of
    // the two dates the software distrusted, and how to fix it.
    const result = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2026 + 10 * DAY_MS,
      clockFloor: JAN_2030,
    });
    expect(result.state).toBe('lapsed');
    expect(result.clockFloorApplied).toBe(true);
    expect(result.explanation).toContain('clock');
  });

  it('does not blame the clock when the clock was not used', () => {
    const result = assessLicense(subscription, KEYS.publicKey, { now: JAN_2030 });
    expect(result.clockFloorApplied).toBe(false);
    expect(result.explanation).not.toContain('clock');
  });

  it('does not flag an ordinary launch where the wall clock is ahead', () => {
    const result = assessLicense(subscription, KEYS.publicKey, {
      now: JAN_2026 + 200 * DAY_MS,
      clockFloor: JAN_2026,
    });
    expect(result.clockFloorApplied).toBe(false);
  });

  it('leaves a perpetual licence untouched', () => {
    const result = assessLicense(perpetual, KEYS.publicKey, {
      now: JAN_2026,
      clockFloor: Date.UTC(2099, 0, 1),
      buildDate: JAN_2026,
      buildMajor: 1,
    });
    expect(result.state).toBe('perpetual');
  });
});

describe('an unusable licence degrades, it does not lock', () => {
  it('reports no licence as the free tier', () => {
    const result = assessLicense(null, KEYS.publicKey, { now: JAN_2026 });
    expect(result.state).toBe('none');
    expect(result.plan).toBe('free');
    expect(result.explanation).toContain('never restricted');
  });

  it('reports a forged licence as the free tier, not as a lockout', () => {
    const forged = signLicense(makePayload({ id: 'X', email: 'a@b.c', plan: 'team' }), OTHER.privateKey);
    const result = assessLicense(forged, KEYS.publicKey, { now: JAN_2026 });
    expect(result.state).toBe('invalid');
    expect(result.plan).toBe('free');
    expect(result.explanation).toContain('files are untouched');
  });

  it('reports a mistyped licence as the free tier', () => {
    const flat = perpetual.replace(/[-\n]/g, '');
    const typo = `${flat.slice(0, 12)}${flat[12] === 'A' ? 'B' : 'A'}${flat.slice(13)}`;
    const result = assessLicense(typo, KEYS.publicKey, { now: JAN_2026 });
    expect(result.plan).toBe('free');
    expect(result.reason).toBe('checksum');
  });

  it('degrades rather than throwing when the build has no usable public key', () => {
    const result = assessLicense(perpetual, 'nonsense', { now: JAN_2026 });
    expect(result.plan).toBe('free');
    expect(result.reason).toBe('public-key');
  });

  it('never throws for any combination of inputs', () => {
    const texts = [null, '', 'junk', perpetual, subscription];
    const clocks = [0, JAN_2026, JAN_2030, Number.MAX_SAFE_INTEGER];
    for (const text of texts) {
      for (const now of clocks) {
        expect(() => assessLicense(text, KEYS.publicKey, { now, buildDate: now })).not.toThrow();
      }
    }
  });

  it('always resolves to a real plan', () => {
    const texts = [null, 'junk', perpetual, subscription];
    for (const text of texts) {
      const result = assessLicense(text, KEYS.publicKey, { now: JAN_2030, buildDate: JAN_2030 });
      expect(['free', 'personal', 'pro', 'team']).toContain(result.plan);
    }
  });
});
