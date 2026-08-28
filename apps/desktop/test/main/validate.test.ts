/**
 * The IPC boundary treats the renderer as an attacker, so these tests pass the
 * things an attacker would pass: lookalike typed arrays, oversized payloads,
 * paths that are not paths, and names designed to escape a directory.
 */

import { describe, expect, it } from 'vitest';
import {
  MAX_FILE_BYTES,
  MAX_JOURNAL_BYTES,
  validateBytes,
  validateDocumentId,
  validateJournalPayload,
  validateOpenPath,
  validateSaveAsRequest,
  validateSaveRequest,
} from '../../src/main/validate.js';

const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);

describe('validateBytes', () => {
  it('accepts a Uint8Array', () => {
    const result = validateBytes(bytes);
    expect(result.ok && result.value).toEqual(bytes);
  });

  it('accepts a raw ArrayBuffer', () => {
    const result = validateBytes(bytes.buffer);
    expect(result.ok && result.value.byteLength).toBe(4);
  });

  it('rejects a number array pretending to be bytes', () => {
    expect(validateBytes([80, 75, 3, 4]).ok).toBe(false);
  });

  it('rejects a string', () => {
    expect(validateBytes('PK').ok).toBe(false);
  });

  it('rejects an object with a forged length', () => {
    expect(validateBytes({ length: 4, byteLength: 4, 0: 80 }).ok).toBe(false);
  });

  it('refuses to write an empty file', () => {
    const result = validateBytes(new Uint8Array(0));
    expect(result.ok).toBe(false);
    expect(result.ok === false && result.reason).toContain('empty');
  });

  it('refuses a payload over the size cap', () => {
    // The cap is checked against byteLength, so a fake view is enough and no
    // half-gigabyte allocation is needed for the test.
    const huge = { byteLength: MAX_FILE_BYTES + 1 };
    Object.setPrototypeOf(huge, Uint8Array.prototype);
    const result = validateBytes(huge);
    expect(result.ok).toBe(false);
    expect(result.ok === false && result.reason).toContain('larger than');
  });
});

describe('validateSaveRequest', () => {
  it('accepts bytes with an absolute writable path', () => {
    const result = validateSaveRequest(bytes, '/home/user/Budget.xlsx');
    expect(result.ok && result.value.path).toBe('/home/user/Budget.xlsx');
  });

  it('rejects a relative path', () => {
    expect(validateSaveRequest(bytes, 'Budget.xlsx').ok).toBe(false);
  });

  it('rejects a format the writer cannot produce', () => {
    const result = validateSaveRequest(bytes, '/home/user/Budget.xls');
    expect(result.ok).toBe(false);
  });

  it('checks the bytes before the path, so an empty write never reaches fs', () => {
    const result = validateSaveRequest(new Uint8Array(0), '/home/user/Budget.xlsx');
    expect(result.ok === false && result.reason).toContain('empty');
  });
});

describe('validateSaveAsRequest', () => {
  it('sanitizes the suggested name rather than rejecting it', () => {
    const result = validateSaveAsRequest(bytes, '../../../etc/passwd.xlsx');
    expect(result.ok && result.value.suggestedName).toBe('passwd.xlsx');
  });

  it('supplies a fallback name when the renderer sends nonsense', () => {
    const result = validateSaveAsRequest(bytes, { toString: () => 'x' });
    expect(result.ok && result.value.suggestedName).toBe('Untitled.xlsx');
  });

  it('still rejects an unusable payload', () => {
    expect(validateSaveAsRequest(null, 'Budget.xlsx').ok).toBe(false);
  });
});

describe('validateOpenPath', () => {
  it('accepts a readable format', () => {
    expect(validateOpenPath('/home/user/data.csv')).toEqual({ ok: true, value: '/home/user/data.csv' });
  });

  it('rejects a file:// URL', () => {
    expect(validateOpenPath('file:///home/user/data.csv').ok).toBe(false);
  });

  it('rejects a directory-looking path with no extension', () => {
    expect(validateOpenPath('/home/user/Documents').ok).toBe(false);
  });
});

describe('journal and document identifiers', () => {
  it('accepts an ordinary payload', () => {
    expect(validateJournalPayload('{"entries":[]}').ok).toBe(true);
  });

  it('rejects a non-string payload', () => {
    expect(validateJournalPayload({ entries: [] }).ok).toBe(false);
  });

  it('rejects an empty payload', () => {
    expect(validateJournalPayload('').ok).toBe(false);
  });

  it('rejects a payload over the cap', () => {
    expect(validateJournalPayload('x'.repeat(MAX_JOURNAL_BYTES + 1)).ok).toBe(false);
  });

  it('accepts a hex document id', () => {
    expect(validateDocumentId('a3f9c1d24b6e4f0a').ok).toBe(true);
  });

  it('rejects a document id that could name a file elsewhere', () => {
    expect(validateDocumentId('../../etc/passwd').ok).toBe(false);
    expect(validateDocumentId('doc id').ok).toBe(false);
    expect(validateDocumentId('').ok).toBe(false);
    expect(validateDocumentId('a'.repeat(65)).ok).toBe(false);
  });
});
