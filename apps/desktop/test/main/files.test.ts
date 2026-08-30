/**
 * Reading and writing documents, against a real filesystem in a temp directory.
 *
 * A fake would not test the thing that matters here, which is that a save
 * replaces the file rather than truncating it, and that it leaves nothing
 * behind when it fails.
 */

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { readDocument, writeDocument } from '../../src/main/files.js';

let dir: string;

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'mirrorz-files-'));
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

describe('readDocument', () => {
  it('reads a file and reports its name', () => {
    const path = join(dir, 'Budget.xlsx');
    writeFileSync(path, Buffer.from([0x50, 0x4b, 0x03, 0x04]));
    const result = readDocument(path);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.file.name).toBe('Budget.xlsx');
    expect(result.file.path).toBe(path);
    expect([...result.file.data]).toEqual([0x50, 0x4b, 0x03, 0x04]);
  });

  it('explains a missing file in a sentence, not an errno', () => {
    const result = readDocument(join(dir, 'nope.xlsx'));
    expect(result.ok).toBe(false);
    expect(result.ok === false && result.error).toMatch(/could not be found/);
    expect(result.ok === false && result.error).not.toMatch(/ENOENT/);
  });

  it('refuses a directory', () => {
    mkdirSync(join(dir, 'folder.xlsx'));
    const result = readDocument(join(dir, 'folder.xlsx'));
    expect(result.ok === false && result.error).toMatch(/not a file/);
  });

  it('refuses an empty file, which no reader can make sense of', () => {
    const path = join(dir, 'empty.xlsx');
    writeFileSync(path, '');
    expect(readDocument(path).ok).toBe(false);
  });

  it('hands over a copy, so the renderer cannot alias the read buffer', () => {
    const path = join(dir, 'a.csv');
    writeFileSync(path, 'a,b\n1,2\n');
    const first = readDocument(path);
    const second = readDocument(path);
    expect(first.ok && second.ok && first.file.data).not.toBe(second.ok && second.file.data);
  });
});

describe('writeDocument', () => {
  it('writes a new file', () => {
    const path = join(dir, 'New.xlsx');
    expect(writeDocument(path, new Uint8Array([1, 2, 3]))).toEqual({ ok: true });
    expect([...readFileSync(path)]).toEqual([1, 2, 3]);
  });

  it('replaces an existing file completely, leaving no tail of the old one', () => {
    const path = join(dir, 'Budget.xlsx');
    writeFileSync(path, Buffer.alloc(4096, 0xaa));
    expect(writeDocument(path, new Uint8Array([1, 2, 3]))).toEqual({ ok: true });
    expect(readFileSync(path).length).toBe(3);
  });

  it('leaves no temporary file behind on success', () => {
    const path = join(dir, 'Budget.xlsx');
    writeDocument(path, new Uint8Array([1, 2, 3]));
    expect(readdirSync(dir)).toEqual(['Budget.xlsx']);
  });

  it('reports a failure it cannot recover from, rather than throwing', () => {
    const locked = join(dir, 'locked');
    mkdirSync(locked);
    chmodSync(locked, 0o500);
    try {
      const result = writeDocument(join(locked, 'Budget.xlsx'), new Uint8Array([1]));
      // Running as root defeats the permission bits, and then the write does
      // succeed; the assertion is that neither outcome throws.
      if (!result.ok) expect(result.error).toMatch(/could not be saved/);
    } finally {
      chmodSync(locked, 0o700);
    }
  });

  it('leaves no temporary file behind on failure', () => {
    const locked = join(dir, 'locked2');
    mkdirSync(locked);
    chmodSync(locked, 0o500);
    try {
      writeDocument(join(locked, 'Budget.xlsx'), new Uint8Array([1]));
      expect(readdirSync(locked).filter((n) => n.endsWith('.tmp'))).toEqual([]);
    } finally {
      chmodSync(locked, 0o700);
    }
  });

  it('round-trips through read', () => {
    const path = join(dir, 'Round.csv');
    const bytes = new Uint8Array([...'a,b\n1,2\n'].map((c) => c.charCodeAt(0)));
    writeDocument(path, bytes);
    const back = readDocument(path);
    expect(back.ok && [...back.file.data]).toEqual([...bytes]);
  });
});
