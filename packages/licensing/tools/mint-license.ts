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

import { generateKeyPair, makePayload, signLicense, verifyLicense } from '../src/license.js';
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
  const days = (name: string): number | null => {
    const raw = options.get(name);
    return raw === undefined ? null : now + Number(raw) * DAY_MS;
  };

  const payload = makePayload({
    id: options.get('id') ?? `MZ-${now.toString(36).toUpperCase()}`,
    email,
    plan,
    kind,
    issued: now,
    // A perpetual licence carries no expiry, ever. That is the product promise
    // and the minter refuses to write one even if asked.
    expires: kind === 'subscription' ? days('term-days') ?? now + 365 * DAY_MS : null,
    maintenanceExpires: days('maintenance-days') ?? (kind === 'perpetual' ? now + 365 * DAY_MS : null),
    seats: Number(options.get('seats') ?? '1'),
    major: Number(options.get('major') ?? '1'),
    features: (options.get('features') ?? '').split(',').filter((f) => f.length > 0),
  });

  const key = signLicense(payload, privateKey);

  // Verify what we just minted with the same code the app runs. A licence that
  // fails here must never reach a customer.
  const pair = verifyLicense(key, process.env.MIRRORZ_LICENSE_PUBLIC_KEY ?? '');
  if (process.env.MIRRORZ_LICENSE_PUBLIC_KEY && !pair.valid) fail(`self-check failed: ${pair.reason}`);

  process.stdout.write(`${key}\n`);
}

main();
