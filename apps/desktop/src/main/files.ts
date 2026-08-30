/**
 * Reading and writing the user's documents.
 *
 * Split out of `ipc.ts` so it imports no Electron: these two functions are
 * where the shell actually touches somebody's data, and they should be testable
 * against a real filesystem rather than only through a running application.
 *
 * Both return a message rather than throwing. An exception crossing IPC arrives
 * in the renderer as a stack trace containing local paths, which is a leak and
 * is also not something a person can act on.
 */

import { basename, dirname, join } from 'node:path';
import { readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { MAX_FILE_BYTES } from './validate.js';
import type { OpenedDocument } from './channels.js';

export type ReadResult = { ok: true; file: OpenedDocument } | { ok: false; error: string };
export type WriteResult = { ok: true } | { ok: false; error: string };

/**
 * The size is checked with stat before the read, so a file too large to handle
 * is refused rather than being pulled into memory first and refused after.
 */
export function readDocument(path: string): ReadResult {
  let size: number;
  try {
    const stat = statSync(path);
    if (!stat.isFile()) return { ok: false, error: 'That is not a file.' };
    size = stat.size;
  } catch {
    return { ok: false, error: 'The file could not be found. It may have been moved or deleted.' };
  }
  if (size === 0) return { ok: false, error: 'The file is empty.' };
  if (size > MAX_FILE_BYTES) {
    return {
      ok: false,
      error: `The file is larger than the ${Math.floor(MAX_FILE_BYTES / (1024 * 1024))} MB limit.`,
    };
  }
  try {
    const buffer = readFileSync(path);
    return { ok: true, file: { path, name: basename(path), data: new Uint8Array(buffer) } };
  } catch {
    return {
      ok: false,
      error: 'The file could not be read. Check that you have permission to open it.',
    };
  }
}

/**
 * Write bytes as atomically as the platform allows.
 *
 * The temporary file is created in the destination directory, because rename(2)
 * is only atomic within a filesystem and the system temp directory is routinely
 * on a different one. A crash halfway through then leaves the previous version
 * of the spreadsheet intact rather than a truncated one - which for a zip
 * container like xlsx is the difference between an old file and no file.
 *
 * If the rename cannot be done - a synced folder holding a lock, a filesystem
 * with unusual semantics - it falls back to writing in place, because refusing
 * to save at all is worse than saving the way everything else does.
 */
export function writeDocument(path: string, bytes: Uint8Array): WriteResult {
  const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.tmp`);
  try {
    writeFileSync(tmp, bytes);
    renameSync(tmp, path);
    return { ok: true };
  } catch {
    try {
      rmSync(tmp, { force: true });
    } catch {
      // The temporary file is already gone, or the directory is unreadable.
    }
    try {
      writeFileSync(path, bytes);
      return { ok: true };
    } catch {
      return {
        ok: false,
        error:
          'The file could not be saved. Check that it is not open elsewhere and that you have permission to write to it.',
      };
    }
  }
}
