/**
 * Offline licensing for MIRRORZ Sheets.
 *
 * No network. No telemetry. No account. Verification is an Ed25519 signature
 * check against a public key compiled into the binary, and every failure path
 * lands on the free tier, which opens, edits and saves every supported format.
 */

export * from './base32.js';
export * from './codec.js';
export * from './entitlements.js';
export * from './license.js';
export * from './plans.js';
export * from './store.js';
export * from './trial.js';
