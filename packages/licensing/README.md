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
   against the wall clock. Omitting `buildDate` means "unknown build", which is
   read as *covered* - never as the current time, because that would silently put
   a wall-clock expiry back on a perpetual licence.
3. **Everything fails open, to the free tier.** No path throws, and no path can
   lock a user out of their own files. A non-finite clock, an unusable public key
   and a payload from the future are all downgrades with a plain explanation, not
   errors.

`test/no-network.test.ts` enforces rule zero structurally: the shipped sources
are read and checked to import nothing but `node:` builtins and their own
siblings, and to name no network global. The promise fails a test rather than a
customer.

`tools/mint-license.ts` is the build-time key generator and licence minter. It is
not part of the shipped application and must never be imported by it. It derives
the public half from the private key and self-checks every licence it issues with
`verifyLicense` - the same code the application runs - so a key that would fail on
a customer's machine cannot be printed. Setting `MIRRORZ_LICENSE_PUBLIC_KEY` adds
a second check that the licence was minted with the key the build actually ships.

Docs: [`docs/pricing.md`](../../docs/pricing.md),
[`docs/naming-and-trademark.md`](../../docs/naming-and-trademark.md).
