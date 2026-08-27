/**
 * A tiny JSON-file store, and the seam that makes everything above it testable.
 *
 * Nothing in this shell talks to `fs` directly except this file and the
 * autosave store. Everything else takes a `JsonStore`, so the recent-files
 * list, the window-state memory and the settings can all be exercised with an
 * in-memory implementation and no temporary directories.
 *
 * Reads never throw. A settings file that has been truncated by a power cut, or
 * hand-edited into invalid JSON, must not stop the application from starting -
 * losing the recent-files list is an annoyance, failing to launch is not.
 */

import { mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';

export interface JsonStore {
  read(): unknown;
  write(value: unknown): void;
}

export function memoryStore(initial?: unknown): JsonStore {
  let current: unknown = initial;
  return {
    read: () => current,
    write: (value) => {
      // Round-trip through JSON so the in-memory store has the same aliasing
      // behaviour as the file one; otherwise tests pass on shared references
      // that would not survive a real write.
      current = JSON.parse(JSON.stringify(value)) as unknown;
    },
  };
}

/**
 * Writes go to a sibling temporary file and are then renamed. rename(2) is
 * atomic within a filesystem, so a crash mid-write leaves either the old file
 * or the new one, never a half-written one that fails to parse on next launch.
 */
export function fileStore(path: string): JsonStore {
  return {
    read(): unknown {
      try {
        return JSON.parse(readFileSync(path, 'utf8')) as unknown;
      } catch {
        return undefined;
      }
    },
    write(value: unknown): void {
      try {
        mkdirSync(dirname(path), { recursive: true });
        const tmp = `${path}.${process.pid}.tmp`;
        writeFileSync(tmp, JSON.stringify(value, null, 2), 'utf8');
        renameSync(tmp, path);
      } catch {
        // A store that cannot persist still has to let the app run.
      }
    },
  };
}
