/**
 * Turning a launch into a list of documents.
 *
 * Most of these cases are real argv shapes: a double-click on Windows, a drop
 * onto a Linux launcher, `electron .` in development, and the argv a second
 * instance forwards to the first.
 */

import { describe, expect, it } from 'vitest';
import { PendingOpenQueue, filesFromArgv } from '../../src/main/cli.js';

const packaged = { packaged: true };
const fromSource = { packaged: false };

describe('filesFromArgv', () => {
  it('finds nothing in a bare launch', () => {
    expect(filesFromArgv(['/opt/mirrorz/mirrorz'], packaged)).toEqual([]);
  });

  it('finds a double-clicked document', () => {
    expect(filesFromArgv(['/opt/mirrorz/mirrorz', '/home/user/Budget.xlsx'], packaged)).toEqual([
      '/home/user/Budget.xlsx',
    ]);
  });

  it('finds several dropped documents', () => {
    const argv = ['/opt/mirrorz/mirrorz', '/home/user/a.xlsx', '/home/user/b.csv'];
    expect(filesFromArgv(argv, packaged)).toEqual(['/home/user/a.xlsx', '/home/user/b.csv']);
  });

  it('skips the script path when running from source', () => {
    const argv = ['/usr/bin/electron', '/home/user/project', '/home/user/Budget.xlsx'];
    expect(filesFromArgv(argv, fromSource)).toEqual(['/home/user/Budget.xlsx']);
  });

  it('ignores the dot that electron . passes', () => {
    expect(filesFromArgv(['electron', '.', '/home/user/a.xlsx'], packaged)).toEqual([
      '/home/user/a.xlsx',
    ]);
  });

  it('ignores Chromium switches', () => {
    const argv = [
      '/opt/mirrorz/mirrorz',
      '--no-sandbox',
      '--disable-gpu',
      '/home/user/a.xlsx',
      '--enable-logging',
    ];
    expect(filesFromArgv(argv, packaged)).toEqual(['/home/user/a.xlsx']);
  });

  it('does not mistake a switch value for a document', () => {
    const argv = ['/opt/mirrorz/mirrorz', '--user-data-dir', '/tmp/profile', '/home/user/a.xlsx'];
    expect(filesFromArgv(argv, packaged)).toEqual(['/home/user/a.xlsx']);
  });

  it('handles a switch that carries its value inline', () => {
    const argv = ['/opt/mirrorz/mirrorz', '--user-data-dir=/tmp/profile', '/home/user/a.xlsx'];
    expect(filesFromArgv(argv, packaged)).toEqual(['/home/user/a.xlsx']);
  });

  it('ignores paths it would refuse to open anyway', () => {
    const argv = ['/opt/mirrorz/mirrorz', '/home/user/notes.docx', 'relative.xlsx', '/home/user/ok.xlsx'];
    expect(filesFromArgv(argv, packaged)).toEqual(['/home/user/ok.xlsx']);
  });

  it('deduplicates the same file passed twice', () => {
    const argv = ['/opt/mirrorz/mirrorz', '/home/user/a.xlsx', '/home/user/a.xlsx'];
    expect(filesFromArgv(argv, packaged)).toEqual(['/home/user/a.xlsx']);
  });

  it('normalises before comparing, so one file is not opened twice', () => {
    const argv = ['/opt/mirrorz/mirrorz', '/home/user/a.xlsx', '/home/user/./a.xlsx'];
    expect(filesFromArgv(argv, packaged)).toEqual(['/home/user/a.xlsx']);
  });

  it('accepts every format the shell can open', () => {
    const argv = [
      '/opt/mirrorz/mirrorz',
      '/a/w.xlsx',
      '/a/m.xlsm',
      '/a/t.xltx',
      '/a/tm.xltm',
      '/a/old.xls',
      '/a/d.csv',
      '/a/d.tsv',
      '/a/o.ods',
    ];
    expect(filesFromArgv(argv, packaged)).toHaveLength(8);
  });
});

describe('PendingOpenQueue', () => {
  it('holds files that arrive before there is anywhere to put them', () => {
    const queue = new PendingOpenQueue();
    queue.push('/home/user/a.xlsx');
    expect(queue.pending()).toEqual(['/home/user/a.xlsx']);
  });

  it('delivers everything queued as soon as a handler attaches', () => {
    const queue = new PendingOpenQueue();
    queue.pushAll(['/home/user/a.xlsx', '/home/user/b.xlsx']);
    const seen: string[] = [];
    queue.attach((path) => seen.push(path));
    expect(seen).toEqual(['/home/user/a.xlsx', '/home/user/b.xlsx']);
    expect(queue.pending()).toEqual([]);
  });

  it('passes later arrivals straight through', () => {
    const queue = new PendingOpenQueue();
    const seen: string[] = [];
    queue.attach((path) => seen.push(path));
    queue.push('/home/user/c.xlsx');
    expect(seen).toEqual(['/home/user/c.xlsx']);
  });

  it('validates what the operating system hands it', () => {
    const queue = new PendingOpenQueue();
    queue.push('not-a-path');
    queue.push('/home/user/payload.exe');
    expect(queue.pending()).toEqual([]);
  });

  it('does not queue the same document twice before the window exists', () => {
    const queue = new PendingOpenQueue();
    queue.push('/home/user/a.xlsx');
    queue.push('/home/user/a.xlsx');
    expect(queue.pending()).toEqual(['/home/user/a.xlsx']);
  });
});
