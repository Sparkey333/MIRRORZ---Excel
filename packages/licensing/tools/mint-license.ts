/**
 * Build-time licence minting. NOT SHIPPED.
 *
 * This is the only program that touches the private key, and it runs on the
 * release machine after a payment webhook (or by hand, for a comp licence). The
 * application binary contains the public half and the verification code, and
 * nothing else - there is no code path in the shipped product that can sign a
 * licence, which is the property that makes the whole scheme safe to run offline.
 *
 * Compile it, then run it:
 *   npx tsc -p packages/licensing/tsconfig.tools.json
 *   node packages/licensing/dist/tools-build/tools/mint-license.js keygen
 *   MIRRORZ_LICENSE_KEY=<base64 private key> \
 *     node packages/licensing/dist/tools-build/tools/mint-license.js mint \
 *       --email a@b.c --plan pro --kind perpetual --id ORD-1234 --maintenance-days 365 --major 1
 *
 * The private key is read from the environment rather than from a flag so it does
 * not land in shell history or in a process listing.
 */

import { derivePublicKey, generateKeyPair, makePayload, signLicense, verifyLicense } from '../src/license.js';
import type { LicenseKind } from '../src/codec.js';
import { DAY_MS } from '../src/codec.js';
import { isPlanId } from '../src/plans.js';

function flags(argv: readonly string[]): Map<string, string> {
  const out = new Map<string, string>();
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === undefined || !token.startsWith('--')) continue;
    const next = argv[i + 1];
    out.set(token.slice(2), next !== undefined && !next.startsWith('--') ? next : 'true');
  }
  return out;
}

function fail(message: string): never {
  process.stderr.write(`${message}\n`);
  process.exit(1);
}

/**
 * Numeric flags, validated rather than coerced.
 *
 * `Number('abc')` is NaN, and a NaN that reaches the payload encoder either
 * throws a RangeError with no context or - worse, before the encoder grew its
 * range checks - mints a key whose dates are nonsense. A licence is a thing a
 * customer pays for and then retypes off a receipt; it is not the place to find
 * out that a flag was a typo.
 */
function integer(options: Map<string, string>, name: string, fallback: number, min: number, max: number): number {
  const raw = options.get(name);
  if (raw === undefined) return fallback;
  const value = Number(raw);
  if (!Number.isFinite(value) || !Number.isInteger(value) || value < min || value > max) {
    fail(`--${name} must be a whole number between ${min} and ${max} (got ${JSON.stringify(raw)})`);
  }
  return value;
}

function main(): void {
  const [command, ...rest] = process.argv.slice(2);

  if (command === 'keygen') {
    const pair = generateKeyPair();
    process.stdout.write(
      [
        '# Keep the private key offline. Anyone holding it can mint licences.',
        `MIRRORZ_LICENSE_KEY=${pair.privateKey}`,
        '',
        '# Compile this into the application:',
        `export const LICENSE_PUBLIC_KEY = '${pair.publicKey}';`,
        '',
      ].join('\n'),
    );
    return;
  }

  if (command !== 'mint') fail('usage: mint-license <keygen|mint> [--flags]');

  const privateKey = process.env.MIRRORZ_LICENSE_KEY;
  if (!privateKey) fail('set MIRRORZ_LICENSE_KEY to the base64 private key');

  const options = flags(rest);
  const email = options.get('email');
  const plan = options.get('plan') ?? 'pro';
  const kind = (options.get('kind') ?? 'perpetual') as LicenseKind;
  if (!email) fail('--email is required');
  if (!isPlanId(plan)) fail(`--plan must be one of free, personal, pro, team`);
  if (kind !== 'perpetual' && kind !== 'subscription') fail('--kind must be perpetual or subscription');

  const now = Date.now();
  // ~100 years, which is longer than any licence anyone should write and short
  // enough to stay inside the range the payload codec will carry.
  const MAX_TERM_DAYS = 36_500;
  const days = (name: string, fallback: number | null): number | null => {
    if (!options.has(name)) return fallback;
    return now + integer(options, name, 0, 1, MAX_TERM_DAYS) * DAY_MS;
  };

  const payload = makePayload({
    id: options.get('id') ?? `MZ-${now.toString(36).toUpperCase()}`,
    email,
    plan,
    kind,
    issued: now,
    // A perpetual licence carries no expiry, ever. That is the product promise
    // and the minter refuses to write one even if asked.
    expires: kind === 'subscription' ? days('term-days', now + 365 * DAY_MS) : null,
    maintenanceExpires: days('maintenance-days', kind === 'perpetual' ? now + 365 * DAY_MS : null),
    seats: integer(options, 'seats', 1, 1, 100_000),
    major: integer(options, 'major', 1, 0, 1_000),
    features: (options.get('features') ?? '').split(',').filter((f) => f.length > 0),
  });

  let key: string;
  try {
    key = signLicense(payload, privateKey);
  } catch (error) {
    fail(`could not sign: ${error instanceof Error ? error.message : String(error)}`);
  }

  /**
   * Self-check every licence with the same code the application runs, always.
   *
   * This used to run only when MIRRORZ_LICENSE_PUBLIC_KEY happened to be set,
   * which is not set by the usage documented at the top of this file - so the
   * check that exists to stop a broken key reaching a customer was, in the
   * normal case, not running at all. The public half is derived from the private
   * half we already hold, so there is nothing left to forget to set.
   */
  const publicKey = derivePublicKey(privateKey);
  const check = verifyLicense(key, publicKey);
  if (!check.valid) fail(`self-check failed (${check.reason}): this licence must not be sent to a customer`);

  // When the build's public key is supplied, also confirm we minted with the
  // matching private key. Minting with last year's key produces a licence that
  // verifies here and fails on the customer's machine.
  const shipped = process.env.MIRRORZ_LICENSE_PUBLIC_KEY;
  if (shipped && shipped.trim() !== publicKey) {
    fail('MIRRORZ_LICENSE_KEY does not match MIRRORZ_LICENSE_PUBLIC_KEY: this licence would not verify in that build');
  }

  process.stdout.write(`${key}\n`);
}

main();
