/**
 * Path safety, kept pure so it can be tested without launching Electron.
 *
 * Everything here exists because the renderer is treated as hostile. It runs
 * sandboxed content derived from files that arrived by email, and a spreadsheet
 * is one of the most effective phishing carriers there is. So a path arriving
 * over IPC is a string of unknown provenance, not a path: it has to be proven
 * absolute, free of embedded NULs (which truncate inside libc and let a name
 * pass validation while a different file gets opened), and pointing at an
 * extension we actually handle before any syscall is made with it.
 *
 * The rules deliberately reject rather than repair. A path we cannot vouch for
 * is an error the user can see, not a silently rewritten one that opens
 * something they did not ask for.
 */

import { basename, extname, isAbsolute, normalize, relative, resolve, sep } from 'node:path';

/** Formats the shell will open. Anything else is refused before it reaches fs. */
export const READABLE_EXTENSIONS = [
  '.xlsx',
  '.xlsm',
  '.xltx',
  '.xltm',
  '.xls',
  '.csv',
  '.tsv',
  '.ods',
  '.txt',
] as const;

/**
 * Formats the shell will write. `.xls` and `.ods` are absent on purpose: the
 * format layer reads them and does not write them, and offering a save that
 * silently produces a different format is worse than not offering it.
 */
export const WRITABLE_EXTENSIONS = ['.xlsx', '.xlsm', '.xltx', '.xltm', '.csv', '.tsv', '.txt'] as const;

/** A path longer than this is a bug or an attack; no real file needs it. */
export const MAX_PATH_LENGTH = 4096;

export type Extension = string;

export function extensionOf(path: string): Extension {
  return extname(path).toLowerCase();
}

export function isReadableExtension(path: string): boolean {
  return (READABLE_EXTENSIONS as readonly string[]).includes(extensionOf(path));
}

export function isWritableExtension(path: string): boolean {
  return (WRITABLE_EXTENSIONS as readonly string[]).includes(extensionOf(path));
}

export interface PathOk {
  ok: true;
  path: string;
}

export interface PathErr {
  ok: false;
  reason: string;
}

export type PathCheck = PathOk | PathErr;

/**
 * Validate a path handed to us by the renderer or by the command line.
 *
 * `mustExist` is not checked here - existence is a filesystem question and this
 * module stays pure. The caller does the stat, after the shape is known good.
 */
export function checkPath(input: unknown, opts: { write?: boolean } = {}): PathCheck {
  if (typeof input !== 'string') return { ok: false, reason: 'path must be a string' };
  if (input.length === 0) return { ok: false, reason: 'path must not be empty' };
  if (input.length > MAX_PATH_LENGTH) return { ok: false, reason: 'path is too long' };
  if (input.includes('\0')) return { ok: false, reason: 'path contains a NUL byte' };

  // A URL is not a path. file:// in particular would round-trip through
  // normalize() looking plausible while meaning something else entirely.
  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(input)) return { ok: false, reason: 'path must not be a URL' };

  if (!isAbsolute(input)) return { ok: false, reason: 'path must be absolute' };

  const normalized = normalize(input);
  const allowed = opts.write ? isWritableExtension(normalized) : isReadableExtension(normalized);
  if (!allowed) {
    return { ok: false, reason: `unsupported file type: ${extensionOf(normalized) || '(none)'}` };
  }
  return { ok: true, path: normalized };
}

/** True when `child` resolves inside `parent`, used to fence the recovery store. */
export function isPathWithin(parent: string, child: string): boolean {
  const rel = relative(resolve(parent), resolve(child));
  if (rel === '') return true;
  return rel !== '' && !rel.startsWith('..') && !isAbsolute(rel);
}

/** Windows reserves these stems in every directory, extension notwithstanding. */
const RESERVED_STEMS = new Set([
  'con',
  'prn',
  'aux',
  'nul',
  ...Array.from({ length: 9 }, (_, i) => `com${i + 1}`),
  ...Array.from({ length: 9 }, (_, i) => `lpt${i + 1}`),
]);

/**
 * Turn a renderer-supplied "suggested name" into something safe to hand to a
 * save dialog. The renderer proposes; it never gets to choose a directory, so
 * every separator is stripped rather than resolved.
 */
export function sanitizeSuggestedName(input: unknown, fallback = 'Untitled.xlsx'): string {
  if (typeof input !== 'string') return fallback;

  // basename twice: POSIX separators survive basename() on Windows and vice
  // versa, and a name like "a/b\\c.xlsx" must not leave a separator behind.
  let name = basename(input.replace(/[\\/]+/g, sep));
  name = name.replace(/[\0-\x1f\x7f]/g, '');
  name = name.replace(/[<>:"|?*]/g, '_');
  name = name.replace(/^\.+/, '');
  name = name.replace(/[. ]+$/, '');
  name = name.trim();
  if (name.length === 0) return fallback;

  const stem = name.slice(0, name.length - extname(name).length).toLowerCase();
  if (RESERVED_STEMS.has(stem)) name = `_${name}`;

  if (name.length > 200) {
    const ext = extname(name);
    name = name.slice(0, 200 - ext.length) + ext;
  }
  return name;
}

/** The filter list for the open dialog, in the order Excel users expect. */
export function openDialogFilters(): { name: string; extensions: string[] }[] {
  return [
    {
      name: 'Spreadsheets',
      extensions: READABLE_EXTENSIONS.map((e) => e.slice(1)),
    },
    { name: 'Excel workbooks', extensions: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xls'] },
    { name: 'Text and delimited', extensions: ['csv', 'tsv', 'txt'] },
    { name: 'OpenDocument', extensions: ['ods'] },
    { name: 'All files', extensions: ['*'] },
  ];
}

/** The filter list for the save dialog: writable formats only. */
export function saveDialogFilters(): { name: string; extensions: string[] }[] {
  return [
    { name: 'Excel workbook', extensions: ['xlsx'] },
    { name: 'Excel macro-enabled workbook', extensions: ['xlsm'] },
    { name: 'Excel template', extensions: ['xltx', 'xltm'] },
    { name: 'Comma separated values', extensions: ['csv'] },
    { name: 'Tab separated values', extensions: ['tsv'] },
  ];
}
