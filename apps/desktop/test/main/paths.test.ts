/**
 * Path safety is the first thing a hostile renderer would attack, so these
 * tests are written as the attacks rather than as the happy path: NUL
 * truncation, URLs wearing a path's clothes, relative escapes, and names that
 * are legal on one platform and a device on another.
 */

import { describe, expect, it } from 'vitest';
import {
  MAX_PATH_LENGTH,
  READABLE_EXTENSIONS,
  WRITABLE_EXTENSIONS,
  checkPath,
  extensionOf,
  isPathWithin,
  isReadableExtension,
  isWritableExtension,
  openDialogFilters,
  sanitizeSuggestedName,
  saveDialogFilters,
} from '../../src/main/paths.js';

describe('extension handling', () => {
  it('lower-cases the extension so BUDGET.XLSX is a workbook', () => {
    expect(extensionOf('/home/a/BUDGET.XLSX')).toBe('.xlsx');
    expect(isReadableExtension('/home/a/BUDGET.XLSX')).toBe(true);
  });

  it('accepts every format the reader supports', () => {
    for (const ext of READABLE_EXTENSIONS) {
      expect(isReadableExtension(`/home/a/file${ext}`)).toBe(true);
    }
  });

  it('refuses to write formats the writer cannot produce', () => {
    expect(isReadableExtension('/home/a/old.xls')).toBe(true);
    expect(isWritableExtension('/home/a/old.xls')).toBe(false);
    expect(isWritableExtension('/home/a/sheet.ods')).toBe(false);
  });

  it('writes only formats in the writable list', () => {
    for (const ext of WRITABLE_EXTENSIONS) {
      expect(isWritableExtension(`/home/a/file${ext}`)).toBe(true);
    }
  });

  it('rejects an executable dressed as a document', () => {
    expect(checkPath('/home/a/payload.exe').ok).toBe(false);
    expect(checkPath('/home/a/payload.xlsx.exe').ok).toBe(false);
  });
});

describe('checkPath', () => {
  it('accepts an absolute path to a supported file', () => {
    const result = checkPath('/home/user/Budget.xlsx');
    expect(result).toEqual({ ok: true, path: '/home/user/Budget.xlsx' });
  });

  it('rejects anything that is not a string', () => {
    for (const value of [undefined, null, 42, {}, [], Symbol('x')]) {
      expect(checkPath(value).ok).toBe(false);
    }
  });

  it('rejects an empty path', () => {
    expect(checkPath('')).toEqual({ ok: false, reason: 'path must not be empty' });
  });

  it('rejects an embedded NUL, which truncates inside libc', () => {
    const result = checkPath('/home/user/safe.xlsx\0/../../etc/shadow');
    expect(result.ok).toBe(false);
    expect(result.ok === false && result.reason).toContain('NUL');
  });

  it('rejects a URL, however plausible it looks', () => {
    expect(checkPath('file:///home/user/Budget.xlsx').ok).toBe(false);
    expect(checkPath('https://example.test/Budget.xlsx').ok).toBe(false);
  });

  it('rejects a relative path', () => {
    expect(checkPath('Budget.xlsx').ok).toBe(false);
    expect(checkPath('../../etc/passwd.xlsx').ok).toBe(false);
  });

  it('collapses traversal segments inside an absolute path', () => {
    const result = checkPath('/home/user/docs/../Budget.xlsx');
    expect(result).toEqual({ ok: true, path: '/home/user/Budget.xlsx' });
  });

  it('rejects a path longer than the limit', () => {
    const long = `/${'a'.repeat(MAX_PATH_LENGTH)}.xlsx`;
    expect(checkPath(long).ok).toBe(false);
  });

  it('applies the writable list when the caller is saving', () => {
    expect(checkPath('/home/user/old.xls', { write: true }).ok).toBe(false);
    expect(checkPath('/home/user/new.xlsx', { write: true }).ok).toBe(true);
  });

  it('names the offending extension so the message is actionable', () => {
    const result = checkPath('/home/user/notes.docx');
    expect(result.ok === false && result.reason).toContain('.docx');
  });
});

describe('isPathWithin', () => {
  it('accepts a file inside the directory', () => {
    expect(isPathWithin('/var/app', '/var/app/recovery/a.json')).toBe(true);
  });

  it('accepts the directory itself', () => {
    expect(isPathWithin('/var/app', '/var/app')).toBe(true);
  });

  it('rejects an escape through ..', () => {
    expect(isPathWithin('/var/app', '/var/app/../etc/passwd')).toBe(false);
  });

  it('rejects a sibling with a shared prefix', () => {
    expect(isPathWithin('/var/app', '/var/application/a.json')).toBe(false);
  });
});

describe('sanitizeSuggestedName', () => {
  it('keeps an ordinary name unchanged', () => {
    expect(sanitizeSuggestedName('Q3 Budget.xlsx')).toBe('Q3 Budget.xlsx');
  });

  it('strips any directory the renderer tried to smuggle in', () => {
    expect(sanitizeSuggestedName('../../etc/cron.d/evil.xlsx')).toBe('evil.xlsx');
    expect(sanitizeSuggestedName('C:\\Windows\\System32\\drivers\\etc\\hosts.csv')).toBe('hosts.csv');
  });

  it('removes control characters', () => {
    expect(sanitizeSuggestedName('bud\u0007get\u0000.xlsx')).toBe('budget.xlsx');
  });

  it('replaces characters Windows refuses in a file name', () => {
    expect(sanitizeSuggestedName('a<b>c:d"e|f?g*h.xlsx')).toBe('a_b_c_d_e_f_g_h.xlsx');
  });

  it('does not produce a dotfile', () => {
    expect(sanitizeSuggestedName('...secret.xlsx')).toBe('secret.xlsx');
  });

  it('trims trailing dots and spaces, which Windows silently drops', () => {
    expect(sanitizeSuggestedName('report.xlsx . ')).toBe('report.xlsx');
  });

  it('escapes reserved device names', () => {
    expect(sanitizeSuggestedName('CON.xlsx')).toBe('_CON.xlsx');
    expect(sanitizeSuggestedName('lpt9.csv')).toBe('_lpt9.csv');
    expect(sanitizeSuggestedName('console.xlsx')).toBe('console.xlsx');
  });

  it('falls back when nothing usable is left', () => {
    expect(sanitizeSuggestedName('')).toBe('Untitled.xlsx');
    expect(sanitizeSuggestedName('/')).toBe('Untitled.xlsx');
    expect(sanitizeSuggestedName(undefined)).toBe('Untitled.xlsx');
    expect(sanitizeSuggestedName(12345)).toBe('Untitled.xlsx');
  });

  it('caps the length while keeping the extension', () => {
    const name = sanitizeSuggestedName(`${'x'.repeat(500)}.xlsx`);
    expect(name.length).toBeLessThanOrEqual(200);
    expect(name.endsWith('.xlsx')).toBe(true);
  });
});

describe('dialog filters', () => {
  it('offers every readable extension in the first filter', () => {
    const first = openDialogFilters()[0];
    expect(first?.extensions).toEqual(READABLE_EXTENSIONS.map((e) => e.slice(1)));
  });

  it('never offers a format the writer cannot produce', () => {
    const offered = new Set(saveDialogFilters().flatMap((f) => f.extensions));
    expect(offered.has('xls')).toBe(false);
    expect(offered.has('ods')).toBe(false);
    expect(offered.has('xlsx')).toBe(true);
  });
});
