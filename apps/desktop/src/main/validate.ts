/**
 * Validation of everything that crosses the IPC boundary, in one pure module.
 *
 * The preload bridge is deliberately thin: it forwards, it does not decide.
 * Every check lives here, on the main-process side, because a preload script
 * shares a process with the page it serves and a renderer compromise that can
 * reach the bridge can also reach anything the bridge trusted. So the rule is
 * that main assumes the arguments were written by an attacker, and the preload
 * is only a convenience for the honest case.
 *
 * The functions return a discriminated result rather than throwing. An IPC
 * handler that throws sends a stack trace back into the renderer, which is both
 * a leak of local paths and a much less useful thing for the UI to show than a
 * sentence.
 */

import { checkPath, sanitizeSuggestedName } from './paths.js';

/**
 * 512 MiB. Large enough for any real workbook - the biggest xlsx files seen in
 * the wild are tens of megabytes - and small enough that a renderer cannot use
 * a save call to fill the disk or exhaust main-process memory in one shot.
 */
export const MAX_FILE_BYTES = 512 * 1024 * 1024;

export type Validated<T> = { ok: true; value: T } | { ok: false; reason: string };

function fail(reason: string): { ok: false; reason: string } {
  return { ok: false, reason };
}

/**
 * Bytes arrive from the renderer as a Uint8Array over the structured clone
 * algorithm. Anything else - a number, an array of numbers, a string, a
 * SharedArrayBuffer - is refused rather than coerced, because coercion is how
 * a length field ends up controlled by the caller.
 */
export function validateBytes(input: unknown): Validated<Uint8Array> {
  let bytes: Uint8Array;
  if (input instanceof Uint8Array) {
    bytes = input;
  } else if (input instanceof ArrayBuffer) {
    bytes = new Uint8Array(input);
  } else {
    return fail('expected file contents as a Uint8Array');
  }
  if (bytes.byteLength === 0) return fail('refusing to write an empty file');
  if (bytes.byteLength > MAX_FILE_BYTES) {
    return fail(`file is larger than the ${Math.floor(MAX_FILE_BYTES / (1024 * 1024))} MB limit`);
  }
  return { ok: true, value: bytes };
}

export interface SaveRequest {
  bytes: Uint8Array;
  path: string;
}

/** `saveFile(bytes, path)`: an overwrite of a file the user already chose. */
export function validateSaveRequest(bytesInput: unknown, pathInput: unknown): Validated<SaveRequest> {
  const bytes = validateBytes(bytesInput);
  if (!bytes.ok) return bytes;
  const path = checkPath(pathInput, { write: true });
  if (!path.ok) return fail(path.reason);
  return { ok: true, value: { bytes: bytes.value, path: path.path } };
}

export interface SaveAsRequest {
  bytes: Uint8Array;
  suggestedName: string;
}

/**
 * `saveFileAs(bytes, suggestedName)`: the renderer proposes a file name and
 * nothing else. The directory comes from the native dialog, so a hostile
 * suggestion can at worst produce an odd-looking default in a box the user is
 * looking at.
 */
export function validateSaveAsRequest(
  bytesInput: unknown,
  nameInput: unknown,
): Validated<SaveAsRequest> {
  const bytes = validateBytes(bytesInput);
  if (!bytes.ok) return bytes;
  return { ok: true, value: { bytes: bytes.value, suggestedName: sanitizeSuggestedName(nameInput) } };
}

/** `openPath(path)`: reading a specific file, typically from Recent Files. */
export function validateOpenPath(pathInput: unknown): Validated<string> {
  const path = checkPath(pathInput, { write: false });
  if (!path.ok) return fail(path.reason);
  return { ok: true, value: path.path };
}

/**
 * The autosave journal. Serialization is the renderer's business - main stores
 * an opaque string - but its size is main's business, since it is written to
 * disk on a timer and an unbounded one would fill the volume between saves.
 */
export const MAX_JOURNAL_BYTES = 64 * 1024 * 1024;

export function validateJournalPayload(input: unknown): Validated<string> {
  if (typeof input !== 'string') return fail('journal payload must be a string');
  if (input.length === 0) return fail('journal payload must not be empty');
  // Length in UTF-16 units understates bytes by at most a factor of three, and
  // the cheap check is the point: it runs on every autosave tick.
  if (input.length > MAX_JOURNAL_BYTES) return fail('journal payload is too large');
  return { ok: true, value: input };
}

/**
 * Document identifiers name recovery files on disk, so they are restricted to
 * a character set that cannot escape a directory or collide with a dotfile.
 */
export function validateDocumentId(input: unknown): Validated<string> {
  if (typeof input !== 'string') return fail('document id must be a string');
  if (!/^[A-Za-z0-9_-]{1,64}$/.test(input)) return fail('document id has an unexpected shape');
  return { ok: true, value: input };
}
