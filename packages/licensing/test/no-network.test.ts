/**
 * The promise, enforced instead of asserted in prose.
 *
 * "No network" is the central claim of this package and of `docs/pricing.md`,
 * and it is the kind of claim that decays quietly: someone adds a "check for
 * updates" helper, or an import that drags in a fetch polyfill, and nothing
 * fails. So the shipped sources are read and checked directly. This is a
 * structural test, not a behavioural one - it is meant to fail in the pull
 * request that introduces the call, not in the field.
 */
import { readdirSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const SRC = new URL('../src/', import.meta.url);

const files = readdirSync(SRC).filter((name) => name.endsWith('.ts'));

/** Every module Node offers that can open a socket, plus the browser's. */
const NETWORK_MODULES = [
  'node:http',
  'node:https',
  'node:http2',
  'node:net',
  'node:tls',
  'node:dgram',
  'node:dns',
  'node:cluster',
  'undici',
  'node-fetch',
  'axios',
];

const NETWORK_GLOBALS = [
  'fetch(',
  'XMLHttpRequest',
  'WebSocket',
  'EventSource',
  'sendBeacon',
  'navigator.onLine',
];

function source(name: string): string {
  return readFileSync(new URL(name, SRC), 'utf8');
}

/**
 * Comments stripped, because these modules explain at length what they refuse to
 * do and naming a thing is not doing it - `from "has been offline for two
 * months"` inside a doc block is prose, not an import.
 */
function code(name: string): string {
  return source(name)
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/(^|[^:])\/\/.*$/gm, '$1');
}

function importsOf(name: string): string[] {
  // The lookbehind keeps `Buffer.from('302a...', 'hex')` - a DER prefix, not an
  // import - out of the list.
  return [...code(name).matchAll(/(?<![.\w])(?:from|import|require)\s*\(?\s*['"]([^'"]+)['"]/g)].map(
    (match) => match[1] as string,
  );
}

describe('the shipped licensing code cannot reach the network', () => {
  it('ships some source to check', () => {
    expect(files.length).toBeGreaterThan(5);
  });

  it.each(files)('%s imports no network-capable module', (name) => {
    for (const specifier of importsOf(name)) {
      expect(NETWORK_MODULES).not.toContain(specifier);
    }
  });

  it.each(files)('%s calls no network global', (name) => {
    const text = code(name);
    for (const global of NETWORK_GLOBALS) {
      expect(text).not.toContain(global);
    }
  });

  it.each(files)('%s imports only node: builtins and its own siblings', (name) => {
    for (const specifier of importsOf(name)) {
      // A third-party dependency is how a network call arrives without anyone
      // writing one, so the package has none and this is what keeps it that way.
      expect(specifier.startsWith('node:') || specifier.startsWith('./')).toBe(true);
    }
  });
});
