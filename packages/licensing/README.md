# @mirrorz/licensing

Offline licence verification, feature gating and the local trial.

**No network. No telemetry. No account.** The shipped application contains a
32-byte Ed25519 public key and verification code, and nothing that can reach a
server. There is no activation call, no heartbeat and no deferred phone-home.

```ts
import {
  assessLicense, entitlementsFor, checkTrial, clockFloor,
  fileLicenseStore, fileTrialStore,
} from '@mirrorz/licensing';

const LICENSE_PUBLIC_KEY = '...'; // baked in at build time
const BUILD_DATE = 1_767_225_600_000; // baked in at build time
const BUILD_MAJOR = 1;

const trialStore = fileTrialStore();
const trial = checkTrial(trialStore);
const assessment = assessLicense(fileLicenseStore().read(), LICENSE_PUBLIC_KEY, {
  buildDate: BUILD_DATE,
  buildMajor: BUILD_MAJOR,
  clockFloor: clockFloor(trialStore),
});

const entitlements = entitlementsFor(assessment, trial);
if (entitlements.can('macro.execute')) { /* ... */ }
```

Three rules the code enforces, in order of importance:

1. **Opening, editing and saving every supported format is never gated** - not by
   plan, not by an expired licence, not by a failed signature check. See
   `NEVER_GATED` in `src/entitlements.ts`.
2. **A perpetual licence never stops working.** `maintenanceExpires` limits which
   *builds* are covered, checked against the build's compile-time date, never
   against the wall clock.
3. **Everything fails open, to the free tier.** No path throws, and no path can
   lock a user out of their own files.

`tools/mint-license.ts` is the build-time key generator and licence minter. It is
not part of the shipped application and must never be imported by it.

Docs: [`docs/pricing.md`](../../docs/pricing.md),
[`docs/naming-and-trademark.md`](../../docs/naming-and-trademark.md).
