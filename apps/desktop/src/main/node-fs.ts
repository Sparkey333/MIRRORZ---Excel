/**
 * The real filesystem behind the autosave store's seam.
 *
 * Separate from `autosave.ts` so that file stays importable - and testable -
 * without `node:fs` being involved at all. Writes go via a temporary file and a
 * rename, for the same reason as the settings store: a journal half-written at
 * the moment of a crash is exactly the journal most likely to matter.
 */

import { mkdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import type { AutosaveFs, FileFacts } from './autosave.js';

export function nodeAutosaveFs(): AutosaveFs {
  return {
    readFile(path: string): string | null {
      try {
        return readFileSync(path, 'utf8');
      } catch {
        return null;
      }
    },
    writeFile(path: string, contents: string): void {
      const tmp = `${path}.${process.pid}.tmp`;
      writeFileSync(tmp, contents, 'utf8');
      renameSync(tmp, path);
    },
    remove(path: string): void {
      rmSync(path, { force: true });
    },
    mkdirp(path: string): void {
      mkdirSync(path, { recursive: true });
    },
    stat(path: string): FileFacts | null {
      try {
        const s = statSync(path);
        return { exists: s.isFile(), mtimeMs: s.mtimeMs };
      } catch {
        return null;
      }
    },
  };
}
