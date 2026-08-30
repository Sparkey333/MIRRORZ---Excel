/**
 * Working out which files the operating system wants us to open.
 *
 * Three separate mechanisms deliver the same intent, and each one is shaped
 * differently:
 *
 *   Windows and Linux pass paths in argv when the user double-clicks a document
 *   or drops one on the application. The awkward part is that argv also
 *   contains the executable, Chromium's own switches, and - when running from
 *   source - the script path, so a naive `argv.slice(1)` opens the project
 *   directory as a spreadsheet.
 *
 *   macOS does not use argv at all. It sends an `open-file` event, sometimes
 *   before the app is ready, which is why the caller queues them.
 *
 *   The second launch of an already-running instance forwards its argv to the
 *   first through the single-instance lock, which is the same parsing problem
 *   with a different source.
 *
 * All of it reduces to one pure function over a string array.
 */

import { checkPath } from './paths.js';

export interface ArgvOptions {
  /**
   * False when running from source (`electron .`), where argv[1] is the app
   * directory or the entry script rather than a document.
   */
  packaged: boolean;
}

/**
 * Chromium and Electron both accept switches that take a separate value
 * argument. Consuming the value as if it were a path is how a launch flag ends
 * up in the Recent Files list.
 */
const SWITCHES_WITH_VALUES = new Set([
  '--user-data-dir',
  '--remote-debugging-port',
  '--inspect',
  '--inspect-brk',
  '--lang',
  '--proxy-server',
  '--js-flags',
]);

export function filesFromArgv(argv: readonly string[], opts: ArgvOptions): string[] {
  const start = opts.packaged ? 1 : 2;
  const out: string[] = [];
  for (let i = start; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === undefined) continue;
    if (arg === '--') continue;
    if (arg.startsWith('-')) {
      // `--switch=value` carries its value inline; `--switch value` does not.
      const bare = arg.split('=')[0] ?? arg;
      if (!arg.includes('=') && SWITCHES_WITH_VALUES.has(bare)) i++;
      continue;
    }
    // `.` is what `electron .` passes, and is never a document.
    if (arg === '.' || arg === './') continue;
    const checked = checkPath(arg, { write: false });
    if (!checked.ok) continue;
    if (!out.includes(checked.path)) out.push(checked.path);
  }
  return out;
}

/**
 * Paths waiting to be opened once a window exists.
 *
 * macOS delivers `open-file` before `ready` on a cold start, and Electron
 * explicitly documents that the event must be handled synchronously, so the
 * only correct shape is a queue that the ready handler drains.
 */
export class PendingOpenQueue {
  private queue: string[] = [];
  private drain: ((path: string) => void) | null = null;

  push(path: string): void {
    const checked = checkPath(path, { write: false });
    if (!checked.ok) return;
    if (this.drain) {
      this.drain(checked.path);
      return;
    }
    if (!this.queue.includes(checked.path)) this.queue.push(checked.path);
  }

  pushAll(paths: readonly string[]): void {
    for (const path of paths) this.push(path);
  }

  /** Attach the real handler; everything queued so far is delivered at once. */
  attach(handler: (path: string) => void): void {
    this.drain = handler;
    const pending = this.queue;
    this.queue = [];
    for (const path of pending) handler(path);
  }

  pending(): string[] {
    return [...this.queue];
  }
}
