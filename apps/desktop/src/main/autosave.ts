/**
 * Crash recovery: an autosave that never touches the user's file.
 *
 * The behaviour people complain about in other spreadsheets is AutoSave writing
 * over the document while they are still deciding, so an accidental edit
 * becomes permanent and the only way back is version history in a cloud they
 * may not use. This does the opposite. The user's file changes when, and only
 * when, they save it. What we persist on a timer is the command log - the
 * document's history - into a separate journal file, and on a clean exit that
 * journal is deleted.
 *
 * That gives crash detection for free and with no heuristics: a journal still
 * on disk at startup means the process that owned it did not shut down
 * cleanly. There is no "was it a crash" guess, no lock file to go stale, and
 * nothing to clean up after a normal quit.
 *
 * The journal lives beside the working file, so it travels with the document
 * and is found even if the application's own data directory is wiped. When that
 * directory cannot be written - a read-only mount, a share with no create
 * permission, a file opened from a disk image - it falls back to the
 * application's recovery directory, which is also where journals for
 * never-saved documents go. An index in that directory records where each
 * journal went, because a startup scan cannot search the whole filesystem.
 *
 * Everything except the two thin fs seams is pure and tested directly.
 */

import { basename, dirname, join } from 'node:path';
import { isPathWithin } from './paths.js';

export const JOURNAL_VERSION = 1;
export const JOURNAL_SUFFIX = '.mirrorz-journal.json';

/** Autosave cadence. Frequent enough to lose little, rare enough to be free. */
export const AUTOSAVE_INTERVAL_MS = 20_000;

/** Journals older than this are assumed abandoned and are not offered back. */
export const MAX_JOURNAL_AGE_MS = 30 * 24 * 60 * 60 * 1000;

/**
 * Clock skew between a network filesystem and this machine routinely runs to
 * seconds, so a file mtime is only treated as newer than a journal when it is
 * clearly newer.
 */
export const SUPERSEDED_GRACE_MS = 2_000;

export interface Journal {
  version: number;
  /** Stable per open document, and the name of the journal file. */
  docId: string;
  /** The file being edited, or null for a document never saved anywhere. */
  filePath: string | null;
  /** Display name for the recovery prompt, so it works even if the file moved. */
  displayName: string;
  savedAt: number;
  appVersion: string;
  /** Opaque to the main process: the renderer serializes its own command log. */
  payload: string;
}

export function encodeJournal(journal: Journal): string {
  return JSON.stringify(journal);
}

/**
 * Tolerant by design. A journal is read exactly once, at startup, from a file
 * that a crash may have truncated mid-write; the only correct response to
 * anything unexpected is to ignore that journal and carry on booting.
 */
export function decodeJournal(text: string): Journal | null {
  let raw: unknown;
  try {
    raw = JSON.parse(text) as unknown;
  } catch {
    return null;
  }
  if (typeof raw !== 'object' || raw === null) return null;
  const j = raw as Partial<Journal>;
  if (j.version !== JOURNAL_VERSION) return null;
  if (typeof j.docId !== 'string' || j.docId.length === 0) return null;
  if (typeof j.payload !== 'string' || j.payload.length === 0) return null;
  if (typeof j.savedAt !== 'number' || !Number.isFinite(j.savedAt)) return null;
  const filePath = typeof j.filePath === 'string' && j.filePath.length > 0 ? j.filePath : null;
  return {
    version: JOURNAL_VERSION,
    docId: j.docId,
    filePath,
    displayName:
      typeof j.displayName === 'string' && j.displayName.length > 0
        ? j.displayName
        : filePath
          ? basename(filePath)
          : 'Untitled',
    savedAt: j.savedAt,
    appVersion: typeof j.appVersion === 'string' ? j.appVersion : 'unknown',
    payload: j.payload,
  };
}

/**
 * Where a journal goes. A leading dot keeps it out of the way on POSIX; on
 * Windows it is merely an odd name, which is preferable to a visible one.
 */
export function journalPathBesideFile(filePath: string): string {
  return join(dirname(filePath), `.${basename(filePath)}${JOURNAL_SUFFIX}`);
}

export function journalPathInRecoveryDir(recoveryDir: string, docId: string): string {
  return join(recoveryDir, `${docId}${JOURNAL_SUFFIX}`);
}

export type RecoveryVerdict = 'recover' | 'stale' | 'superseded' | 'source-missing';

export interface FileFacts {
  exists: boolean;
  mtimeMs: number;
}

export interface ClassifyOptions {
  now: number;
  maxAgeMs?: number;
  graceMs?: number;
}

/**
 * Decide what a leftover journal means.
 *
 * The interesting case is `superseded`: the file on disk is newer than the last
 * autosave, so the user did save after the journal was written and the journal
 * describes history they already have. Offering it back would invite them to
 * overwrite good work with old work.
 */
export function classifyJournal(
  journal: Journal,
  file: FileFacts | null,
  opts: ClassifyOptions,
): RecoveryVerdict {
  const maxAge = opts.maxAgeMs ?? MAX_JOURNAL_AGE_MS;
  const grace = opts.graceMs ?? SUPERSEDED_GRACE_MS;
  if (opts.now - journal.savedAt > maxAge) return 'stale';
  if (journal.filePath === null) return 'recover';
  if (!file || !file.exists) return 'source-missing';
  if (file.mtimeMs > journal.savedAt + grace) return 'superseded';
  return 'recover';
}

export interface RecoveryCandidate {
  journal: Journal;
  journalPath: string;
  verdict: RecoveryVerdict;
}

/** Only these are worth putting in front of the user. */
export function offerable(candidates: RecoveryCandidate[]): RecoveryCandidate[] {
  return candidates.filter((c) => c.verdict === 'recover' || c.verdict === 'source-missing');
}

/** The narrow filesystem surface the store needs, so tests can supply a fake. */
export interface AutosaveFs {
  readFile(path: string): string | null;
  writeFile(path: string, contents: string): void;
  remove(path: string): void;
  mkdirp(path: string): void;
  stat(path: string): FileFacts | null;
}

interface IndexRecord {
  docId: string;
  journalPath: string;
}

interface PersistedIndex {
  version: 1;
  records: IndexRecord[];
}

export function parseIndex(raw: string | null): IndexRecord[] {
  if (raw === null) return [];
  let value: unknown;
  try {
    value = JSON.parse(raw) as unknown;
  } catch {
    return [];
  }
  if (typeof value !== 'object' || value === null) return [];
  const index = value as Partial<PersistedIndex>;
  if (index.version !== 1 || !Array.isArray(index.records)) return [];
  const out: IndexRecord[] = [];
  for (const record of index.records) {
    if (typeof record !== 'object' || record === null) continue;
    const r = record as Partial<IndexRecord>;
    if (typeof r.docId !== 'string' || typeof r.journalPath !== 'string') continue;
    if (r.docId.length === 0 || r.journalPath.length === 0) continue;
    out.push({ docId: r.docId, journalPath: r.journalPath });
  }
  return out;
}

/**
 * The journal store. One instance per application run; the recovery directory
 * is the application's own, and is the only place this class will create files
 * other than beside a document the user already has open.
 */
export class AutosaveStore {
  private readonly indexPath: string;

  constructor(
    private readonly fs: AutosaveFs,
    private readonly recoveryDir: string,
    private readonly appVersion: string,
  ) {
    this.indexPath = join(recoveryDir, 'journals.json');
  }

  /**
   * Write a journal, preferring the document's own directory.
   *
   * Returns the path written, or null if even the fallback failed - in which
   * case the caller carries on rather than interrupting the user, since a
   * failed autosave is not a reason to stop them editing.
   */
  save(input: {
    docId: string;
    filePath: string | null;
    displayName: string;
    payload: string;
    now: number;
  }): string | null {
    const journal: Journal = {
      version: JOURNAL_VERSION,
      docId: input.docId,
      filePath: input.filePath,
      displayName: input.displayName,
      savedAt: input.now,
      appVersion: this.appVersion,
      payload: input.payload,
    };
    const encoded = encodeJournal(journal);

    const preferred = input.filePath ? journalPathBesideFile(input.filePath) : null;
    if (preferred) {
      try {
        this.fs.writeFile(preferred, encoded);
        this.record(input.docId, preferred);
        return preferred;
      } catch {
        // Read-only directory, or a share that refuses dotfiles: fall through.
      }
    }
    const fallback = journalPathInRecoveryDir(this.recoveryDir, input.docId);
    try {
      this.fs.mkdirp(this.recoveryDir);
      this.fs.writeFile(fallback, encoded);
      this.record(input.docId, fallback);
      return fallback;
    } catch {
      return null;
    }
  }

  /** Called on a clean close: the absence of a journal is what says "no crash". */
  discard(docId: string): void {
    const records = this.readIndex();
    for (const record of records) {
      if (record.docId !== docId) continue;
      try {
        this.fs.remove(record.journalPath);
      } catch {
        // Already gone, or on a volume that has been unmounted since.
      }
    }
    this.writeIndex(records.filter((r) => r.docId !== docId));
  }

  /**
   * Everything left behind by a previous run, classified.
   *
   * Journals that cannot be read or parsed are removed from the index here
   * rather than being reported, so a corrupt file does not reappear in the
   * recovery prompt on every subsequent launch.
   */
  scan(now: number): RecoveryCandidate[] {
    const records = this.readIndex();
    const kept: IndexRecord[] = [];
    const candidates: RecoveryCandidate[] = [];
    for (const record of records) {
      const text = this.fs.readFile(record.journalPath);
      const journal = text === null ? null : decodeJournal(text);
      if (!journal) continue;
      const file = journal.filePath ? this.fs.stat(journal.filePath) : null;
      const verdict = classifyJournal(journal, file, { now });
      if (verdict === 'stale' || verdict === 'superseded') {
        try {
          this.fs.remove(record.journalPath);
        } catch {
          // Nothing useful to do; it will be retried next launch.
        }
        continue;
      }
      kept.push(record);
      candidates.push({ journal, journalPath: record.journalPath, verdict });
    }
    this.writeIndex(kept);
    return candidates;
  }

  private record(docId: string, journalPath: string): void {
    const records = this.readIndex().filter((r) => r.docId !== docId);
    records.push({ docId, journalPath });
    this.writeIndex(records);
  }

  private readIndex(): IndexRecord[] {
    return parseIndex(this.fs.readFile(this.indexPath));
  }

  private writeIndex(records: IndexRecord[]): void {
    const payload: PersistedIndex = { version: 1, records };
    try {
      this.fs.mkdirp(this.recoveryDir);
      this.fs.writeFile(this.indexPath, JSON.stringify(payload));
    } catch {
      // The index is a convenience; losing it costs a recovery offer, not data.
    }
  }

  /** Guard for the fallback location: never write outside our own directory. */
  ownsPath(path: string): boolean {
    return isPathWithin(this.recoveryDir, path);
  }
}
