/**
 * Filesystem storage for the licence key and the trial record.
 *
 * Kept in its own module so the licence and trial logic stay pure and importable
 * from a renderer, a worker or a test with no filesystem at all. Both files live
 * in the ordinary per-user config directory, in plain text, where the user can
 * read them, back them up, copy them to another machine and delete them. Nothing
 * is hidden, obfuscated or written outside the user's own profile: a program that
 * scatters hidden state around someone's computer to enforce a payment has
 * already lost the argument about whose computer it is.
 */

import { mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs';
import { homedir, platform } from 'node:os';
import { dirname, join } from 'node:path';

import type { TrialStore } from './trial.js';

export const APP_DIR_NAME = 'MIRRORZ Sheets';

/** Per-user config directory, following each platform's own convention. */
export function configDir(appName: string = APP_DIR_NAME, env: NodeJS.ProcessEnv = process.env): string {
  switch (platform()) {
    case 'win32':
      return join(env.APPDATA ?? join(homedir(), 'AppData', 'Roaming'), appName);
    case 'darwin':
      return join(homedir(), 'Library', 'Application Support', appName);
    default:
      return join(env.XDG_CONFIG_HOME ?? join(homedir(), '.config'), appName);
  }
}

export function licensePath(appName?: string): string {
  return join(configDir(appName), 'license.key');
}

export function trialPath(appName?: string): string {
  return join(configDir(appName), 'trial.json');
}

/**
 * A text file that behaves like a {@link TrialStore}: missing or unreadable
 * reads as null rather than throwing, and writes are atomic through a rename so
 * a crash mid-write cannot leave a half-written record that reads as tampering.
 */
export class FileTextStore implements TrialStore {
  constructor(readonly path: string) {}

  read(): string | null {
    try {
      return readFileSync(this.path, 'utf8');
    } catch {
      return null;
    }
  }

  write(text: string): void {
    mkdirSync(dirname(this.path), { recursive: true });
    const temporary = `${this.path}.tmp`;
    try {
      writeFileSync(temporary, text, { encoding: 'utf8', mode: 0o600 });
      renameSync(temporary, this.path);
    } catch (error) {
      try {
        unlinkSync(temporary);
      } catch {
        // Nothing useful to do; the original file is still intact.
      }
      throw error;
    }
  }

  /** Removing a licence is the user's right and must always be one call. */
  clear(): void {
    try {
      unlinkSync(this.path);
    } catch {
      // Already gone.
    }
  }
}

export function fileTrialStore(appName?: string): FileTextStore {
  return new FileTextStore(trialPath(appName));
}

export function fileLicenseStore(appName?: string): FileTextStore {
  return new FileTextStore(licensePath(appName));
}
