/**
 * The settings store. Its only real requirement is that nothing it does can
 * stop the application from starting, however damaged the file on disk is.
 */

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileStore, memoryStore } from '../../src/main/store.js';

let dir: string;

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'mirrorz-store-'));
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

describe('memoryStore', () => {
  it('reads back what was written', () => {
    const store = memoryStore();
    store.write({ a: 1 });
    expect(store.read()).toEqual({ a: 1 });
  });

  it('starts from the value it was seeded with', () => {
    expect(memoryStore({ a: 1 }).read()).toEqual({ a: 1 });
  });

  it('does not alias the value it was given, matching the file store', () => {
    const store = memoryStore();
    const value = { entries: [1] };
    store.write(value);
    value.entries.push(2);
    expect(store.read()).toEqual({ entries: [1] });
  });
});

describe('fileStore', () => {
  it('returns undefined when the file does not exist', () => {
    expect(fileStore(join(dir, 'settings.json')).read()).toBeUndefined();
  });

  it('round-trips a value', () => {
    const store = fileStore(join(dir, 'settings.json'));
    store.write({ version: 1, entries: ['a'] });
    expect(store.read()).toEqual({ version: 1, entries: ['a'] });
  });

  it('creates the directory it was pointed at', () => {
    const store = fileStore(join(dir, 'nested', 'deeper', 'settings.json'));
    store.write({ a: 1 });
    expect(existsSync(join(dir, 'nested', 'deeper', 'settings.json'))).toBe(true);
  });

  it('returns undefined rather than throwing on invalid JSON', () => {
    const path = join(dir, 'settings.json');
    writeFileSync(path, '{ truncated');
    expect(fileStore(path).read()).toBeUndefined();
  });

  it('overwrites a damaged file on the next write', () => {
    const path = join(dir, 'settings.json');
    writeFileSync(path, 'garbage');
    const store = fileStore(path);
    store.write({ ok: true });
    expect(JSON.parse(readFileSync(path, 'utf8')) as unknown).toEqual({ ok: true });
  });

  it('leaves no temporary file behind', () => {
    const store = fileStore(join(dir, 'settings.json'));
    store.write({ a: 1 });
    expect(readdirSync(dir)).toEqual(['settings.json']);
  });

  it('swallows a write it cannot perform', () => {
    const locked = join(dir, 'locked');
    mkdirSync(locked);
    chmodSync(locked, 0o500);
    try {
      expect(() => fileStore(join(locked, 'settings.json')).write({ a: 1 })).not.toThrow();
    } finally {
      chmodSync(locked, 0o700);
    }
  });
});
