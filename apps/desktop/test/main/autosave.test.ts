/**
 * Crash recovery, tested against a fake filesystem.
 *
 * The two properties that matter most are asserted directly: the user's file is
 * never written by autosave, and a journal that describes work the user has
 * already saved is never offered back.
 */

import { describe, expect, it } from 'vitest';
import {
  AutosaveStore,
  JOURNAL_SUFFIX,
  MAX_JOURNAL_AGE_MS,
  SUPERSEDED_GRACE_MS,
  classifyJournal,
  decodeJournal,
  encodeJournal,
  journalPathBesideFile,
  journalPathInRecoveryDir,
  looksLikeJournalPath,
  offerable,
  parseIndex,
  type AutosaveFs,
  type FileFacts,
  type Journal,
} from '../../src/main/autosave.js';

function fakeFs(seed: Record<string, string> = {}, opts: { readOnly?: string[] } = {}): AutosaveFs & {
  files: Map<string, string>;
  stats: Map<string, FileFacts>;
} {
  const files = new Map(Object.entries(seed));
  const stats = new Map<string, FileFacts>();
  return {
    files,
    stats,
    readFile: (path) => files.get(path) ?? null,
    writeFile: (path, contents) => {
      if (opts.readOnly?.some((dir) => path.startsWith(dir))) {
        throw new Error('EACCES');
      }
      files.set(path, contents);
    },
    remove: (path) => {
      files.delete(path);
    },
    mkdirp: () => {},
    stat: (path) => stats.get(path) ?? null,
  };
}

const journal: Journal = {
  version: 1,
  docId: 'doc1',
  filePath: '/home/user/Budget.xlsx',
  displayName: 'Budget.xlsx',
  savedAt: 1_000_000,
  appVersion: '0.1.0',
  payload: '{"entries":[1,2,3]}',
};

describe('journal encoding', () => {
  it('round-trips', () => {
    expect(decodeJournal(encodeJournal(journal))).toEqual(journal);
  });

  it('returns null for anything that is not a journal', () => {
    expect(decodeJournal('')).toBeNull();
    expect(decodeJournal('{')).toBeNull();
    expect(decodeJournal('null')).toBeNull();
    expect(decodeJournal('[]')).toBeNull();
    expect(decodeJournal('"a string"')).toBeNull();
  });

  it('rejects a journal from a future format version', () => {
    expect(decodeJournal(JSON.stringify({ ...journal, version: 99 }))).toBeNull();
  });

  it('rejects a truncated journal with no payload', () => {
    expect(decodeJournal(JSON.stringify({ ...journal, payload: '' }))).toBeNull();
    expect(decodeJournal(JSON.stringify({ ...journal, savedAt: 'yesterday' }))).toBeNull();
  });

  it('recovers a display name from the file path when it is missing', () => {
    const decoded = decodeJournal(JSON.stringify({ ...journal, displayName: undefined }));
    expect(decoded?.displayName).toBe('Budget.xlsx');
  });

  it('treats a document with no file as Untitled', () => {
    const decoded = decodeJournal(JSON.stringify({ ...journal, filePath: null, displayName: '' }));
    expect(decoded?.filePath).toBeNull();
    expect(decoded?.displayName).toBe('Untitled');
  });
});

describe('journal locations', () => {
  it('sits beside the document as a dotfile, never over it', () => {
    const path = journalPathBesideFile('/home/user/Budget.xlsx');
    expect(path).toBe(`/home/user/.Budget.xlsx${JOURNAL_SUFFIX}`);
    expect(path).not.toBe('/home/user/Budget.xlsx');
  });

  it('falls back to the application directory for a document with no file', () => {
    expect(journalPathInRecoveryDir('/var/app/recovery', 'doc1')).toBe(
      `/var/app/recovery/doc1${JOURNAL_SUFFIX}`,
    );
  });
});

describe('classifyJournal', () => {
  const now = 2_000_000;

  it('offers recovery when the file is untouched since the crash', () => {
    expect(classifyJournal(journal, { exists: true, mtimeMs: 900_000 }, { now })).toBe('recover');
  });

  it('discards a journal describing work the user has since saved', () => {
    const file = { exists: true, mtimeMs: journal.savedAt + SUPERSEDED_GRACE_MS + 1 };
    expect(classifyJournal(journal, file, { now })).toBe('superseded');
  });

  it('tolerates clock skew inside the grace window', () => {
    const file = { exists: true, mtimeMs: journal.savedAt + SUPERSEDED_GRACE_MS - 1 };
    expect(classifyJournal(journal, file, { now })).toBe('recover');
  });

  it('still offers recovery when the source file has been deleted', () => {
    expect(classifyJournal(journal, { exists: false, mtimeMs: 0 }, { now })).toBe('source-missing');
    expect(classifyJournal(journal, null, { now })).toBe('source-missing');
  });

  it('always offers a document that was never saved anywhere', () => {
    const untitled = { ...journal, filePath: null };
    expect(classifyJournal(untitled, null, { now })).toBe('recover');
  });

  it('gives up on a journal nobody came back for', () => {
    expect(classifyJournal(journal, null, { now: journal.savedAt + MAX_JOURNAL_AGE_MS + 1 })).toBe('stale');
  });

  it('offers only the verdicts worth showing a user', () => {
    const kinds = (['recover', 'stale', 'superseded', 'source-missing'] as const).map((verdict) => ({
      journal,
      journalPath: '/x',
      verdict,
    }));
    expect(offerable(kinds).map((c) => c.verdict)).toEqual(['recover', 'source-missing']);
  });
});

describe('parseIndex', () => {
  it('survives a missing or damaged index', () => {
    expect(parseIndex(null)).toEqual([]);
    expect(parseIndex('{')).toEqual([]);
    expect(parseIndex('[]')).toEqual([]);
    expect(parseIndex(JSON.stringify({ version: 2, records: [] }))).toEqual([]);
  });

  it('drops records that are not usable', () => {
    const good = `/var/app/recovery/b${JOURNAL_SUFFIX}`;
    const raw = JSON.stringify({
      version: 1,
      records: [null, { docId: 'a' }, { docId: 'b', journalPath: good }],
    });
    expect(parseIndex(raw)).toEqual([{ docId: 'b', journalPath: good }]);
  });

  /**
   * Every path in this index is one the store later hands to `remove()`, so a
   * record naming something that is not a journal is the one way a damaged file
   * in the application's own data directory turns into an unlink of the user's
   * work. The shape test is what stops it.
   */
  it('refuses a record that names something which is not a journal', () => {
    const raw = JSON.stringify({
      version: 1,
      records: [
        { docId: 'a', journalPath: '/home/user/Budget.xlsx' },
        { docId: 'b', journalPath: '/' },
        { docId: 'c', journalPath: `/home/user/.Budget.xlsx\u0000${JOURNAL_SUFFIX}` },
      ],
    });
    expect(parseIndex(raw)).toEqual([]);
  });
});

describe('looksLikeJournalPath', () => {
  it('accepts both places a journal is ever written', () => {
    expect(looksLikeJournalPath(journalPathBesideFile('/home/user/Budget.xlsx'))).toBe(true);
    expect(looksLikeJournalPath(journalPathInRecoveryDir('/var/app/recovery', 'doc1'))).toBe(true);
  });

  it('rejects a document, an empty path and a NUL', () => {
    expect(looksLikeJournalPath('/home/user/Budget.xlsx')).toBe(false);
    expect(looksLikeJournalPath('')).toBe(false);
    expect(looksLikeJournalPath(`/a\u0000${JOURNAL_SUFFIX}`)).toBe(false);
  });
});

describe('AutosaveStore', () => {
  it('writes beside the document and never touches the document itself', () => {
    const fs = fakeFs();
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    const written = store.save({
      docId: 'doc1',
      filePath: '/home/user/Budget.xlsx',
      displayName: 'Budget.xlsx',
      payload: 'LOG',
      now: 1_000,
    });
    expect(written).toBe(`/home/user/.Budget.xlsx${JOURNAL_SUFFIX}`);
    expect(fs.files.has('/home/user/Budget.xlsx')).toBe(false);
  });

  it('falls back to the application directory when the document folder is read-only', () => {
    const fs = fakeFs({}, { readOnly: ['/mnt/readonly'] });
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    const written = store.save({
      docId: 'doc1',
      filePath: '/mnt/readonly/Budget.xlsx',
      displayName: 'Budget.xlsx',
      payload: 'LOG',
      now: 1_000,
    });
    expect(written).toBe(`/var/app/recovery/doc1${JOURNAL_SUFFIX}`);
    expect(store.ownsPath(written ?? '')).toBe(true);
  });

  it('reports failure rather than throwing when nowhere is writable', () => {
    const fs = fakeFs({}, { readOnly: ['/'] });
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    expect(store.save({ docId: 'd', filePath: null, displayName: 'Untitled', payload: 'L', now: 1 })).toBeNull();
  });

  it('finds what an unclean exit left behind', () => {
    const fs = fakeFs();
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    store.save({ docId: 'doc1', filePath: '/home/user/Budget.xlsx', displayName: 'Budget.xlsx', payload: 'LOG', now: 1_000 });
    fs.stats.set('/home/user/Budget.xlsx', { exists: true, mtimeMs: 500 });

    const found = new AutosaveStore(fs, '/var/app/recovery', '0.1.0').scan(2_000);
    expect(found).toHaveLength(1);
    expect(found[0]?.journal.payload).toBe('LOG');
    expect(found[0]?.verdict).toBe('recover');
  });

  it('finds nothing after a clean close, which is how a crash is detected', () => {
    const fs = fakeFs();
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    store.save({ docId: 'doc1', filePath: '/home/user/Budget.xlsx', displayName: 'Budget.xlsx', payload: 'LOG', now: 1_000 });
    store.discard('doc1');

    expect(new AutosaveStore(fs, '/var/app/recovery', '0.1.0').scan(2_000)).toEqual([]);
    expect(fs.files.has(`/home/user/.Budget.xlsx${JOURNAL_SUFFIX}`)).toBe(false);
  });

  it('keeps one journal per document rather than one per save', () => {
    const fs = fakeFs();
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    for (let i = 0; i < 5; i++) {
      store.save({ docId: 'doc1', filePath: '/home/user/Budget.xlsx', displayName: 'Budget.xlsx', payload: `LOG${i}`, now: 1_000 + i });
    }
    const journals = [...fs.files.keys()].filter((k) => k.endsWith(JOURNAL_SUFFIX));
    expect(journals).toHaveLength(1);
    expect(fs.files.get(journals[0] ?? '')).toContain('LOG4');
  });

  it('cleans up a journal the user has already made redundant by saving', () => {
    const fs = fakeFs();
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    store.save({ docId: 'doc1', filePath: '/home/user/Budget.xlsx', displayName: 'Budget.xlsx', payload: 'LOG', now: 1_000 });
    fs.stats.set('/home/user/Budget.xlsx', { exists: true, mtimeMs: 1_000_000 });

    expect(store.scan(2_000_000)).toEqual([]);
    expect(fs.files.has(`/home/user/.Budget.xlsx${JOURNAL_SUFFIX}`)).toBe(false);
  });

  it('forgets a corrupt journal instead of offering it on every launch', () => {
    const fs = fakeFs();
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    store.save({ docId: 'doc1', filePath: '/home/user/Budget.xlsx', displayName: 'Budget.xlsx', payload: 'LOG', now: 1_000 });
    fs.files.set(`/home/user/.Budget.xlsx${JOURNAL_SUFFIX}`, 'truncated{{{');

    expect(store.scan(2_000)).toEqual([]);
    expect(store.scan(2_000)).toEqual([]);
  });

  it('recovers several documents from one crash', () => {
    const fs = fakeFs();
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    store.save({ docId: 'doc1', filePath: '/home/user/a.xlsx', displayName: 'a.xlsx', payload: 'A', now: 1_000 });
    store.save({ docId: 'doc2', filePath: null, displayName: 'Untitled', payload: 'B', now: 1_100 });
    fs.stats.set('/home/user/a.xlsx', { exists: true, mtimeMs: 100 });

    const found = store.scan(2_000);
    expect(found.map((c) => c.journal.docId).sort()).toEqual(['doc1', 'doc2']);
  });

  it('discards only the document asked for', () => {
    const fs = fakeFs();
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    store.save({ docId: 'doc1', filePath: '/home/user/a.xlsx', displayName: 'a.xlsx', payload: 'A', now: 1_000 });
    store.save({ docId: 'doc2', filePath: '/home/user/b.xlsx', displayName: 'b.xlsx', payload: 'B', now: 1_000 });
    store.discard('doc1');
    fs.stats.set('/home/user/b.xlsx', { exists: true, mtimeMs: 100 });

    expect(store.scan(2_000).map((c) => c.journal.docId)).toEqual(['doc2']);
  });

  it('records the running application version, so an old journal is identifiable', () => {
    const fs = fakeFs();
    new AutosaveStore(fs, '/var/app/recovery', '9.9.9').save({
      docId: 'doc1',
      filePath: null,
      displayName: 'Untitled',
      payload: 'L',
      now: 5,
    });
    const raw = fs.files.get(`/var/app/recovery/doc1${JOURNAL_SUFFIX}`) ?? '';
    expect(decodeJournal(raw)?.appVersion).toBe('9.9.9');
  });

  it('moves the journal when an untitled document is first saved to a file', () => {
    const fs = fakeFs();
    const store = new AutosaveStore(fs, '/var/app/recovery', '0.1.0');
    store.save({ docId: 'doc1', filePath: null, displayName: 'Untitled', payload: 'A', now: 1_000 });
    store.save({ docId: 'doc1', filePath: '/home/user/Budget.xlsx', displayName: 'Budget.xlsx', payload: 'B', now: 1_100 });

    expect(fs.files.has(`/var/app/recovery/doc1${JOURNAL_SUFFIX}`)).toBe(false);
    expect(fs.files.has(`/home/user/.Budget.xlsx${JOURNAL_SUFFIX}`)).toBe(true);
    expect(store.scan(2_000)).toHaveLength(1);
  });

  it('refuses to claim a path outside its own recovery directory', () => {
    const store = new AutosaveStore(fakeFs(), '/var/app/recovery', '0.1.0');
    expect(store.ownsPath('/var/app/recovery/doc1.json')).toBe(true);
    expect(store.ownsPath('/home/user/Budget.xlsx')).toBe(false);
  });
});
