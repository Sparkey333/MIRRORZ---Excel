/**
 * The main process: lifecycle, and nothing else.
 *
 * Everything with a decision in it lives in a sibling module that can be tested
 * without Electron - path safety, IPC validation, the recent-files list, the
 * autosave journal, the menu template, argv parsing. What is left here is the
 * wiring: what happens when the app is ready, when a file arrives from the
 * operating system, when a window closes, when the last one does.
 *
 * The order of the startup sequence matters and is not arbitrary. The sandbox
 * is enabled and the session policy installed before any window can exist, so
 * there is no window that ever ran without them.
 */

import { app, dialog, Menu, session } from 'electron';
import { join } from 'node:path';
import { existsSync } from 'node:fs';
import { AutosaveStore, AUTOSAVE_INTERVAL_MS, offerable } from './autosave.js';
import { nodeAutosaveFs } from './node-fs.js';
import { RecentFiles, foldsCase } from './recent.js';
import { fileStore } from './store.js';
import { filesFromArgv, PendingOpenQueue } from './cli.js';
import { buildMenuTemplate, type MenuCommand } from './menu.js';
import { registerIpc } from './ipc.js';
import { readDocument } from './files.js';
import { installSessionPolicies, WindowManager, type DocumentWindow } from './windows.js';
import { CHANNEL } from './channels.js';

/**
 * In development the renderer comes from Vite so hot reload works; in a
 * packaged build there is no dev server and the files are on disk. Deciding by
 * `app.isPackaged` rather than NODE_ENV means a packaged build can never be
 * talked into loading from the network by an environment variable.
 */
const devServer = app.isPackaged
  ? undefined
  : (process.env['MIRRORZ_DEV_SERVER'] ?? 'http://localhost:5273');

const pendingOpens = new PendingOpenQueue();

// macOS sends this before `ready` on a cold start and it must be handled
// synchronously, so it is registered at module scope rather than after setup.
app.on('open-file', (event, path) => {
  event.preventDefault();
  pendingOpens.push(path);
});

app.enableSandbox();

/**
 * A second launch - double-clicking another spreadsheet - hands its argv to the
 * running instance instead of starting a rival one. Two instances would each
 * have their own recent-files list and their own idea of what is open.
 */
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  pendingOpens.pushAll(filesFromArgv(process.argv, { packaged: app.isPackaged }));
  app.on('second-instance', (_event, argv) => {
    pendingOpens.pushAll(filesFromArgv(argv, { packaged: true }));
  });
  void start();
}

async function start(): Promise<void> {
  await app.whenReady();

  /**
   * The autosave clock lives in main, so the cadence is one policy in one place
   * rather than a timer in every renderer that could drift or be starved by a
   * long recalculation.
   */
  const tickers = new Map<number, NodeJS.Timeout>();

  const userData = app.getPath('userData');
  const recent = new RecentFiles(fileStore(join(userData, 'recent-files.json')));
  recent.prune((path) => existsSync(path));

  const autosave = new AutosaveStore(nodeAutosaveFs(), join(userData, 'recovery'), app.getVersion());

  installSessionPolicies(session.defaultSession, devServer);

  const appRoot = app.getAppPath();
  const windows = new WindowManager({
    preloadPath: join(appRoot, 'dist/preload/index.cjs'),
    rendererFile: join(appRoot, 'dist/renderer/index.html'),
    devServer,
    onClosed: (doc) => {
      // A window that closed through the normal path is a clean exit for that
      // document, and its journal is what would otherwise look like a crash.
      autosave.discard(doc.docId);
      stopAutosaveTicker(doc);
    },
  });

  registerIpc({
    app,
    windows,
    recent,
    autosave,
    now: () => Date.now(),
    onRecentChanged: () => installMenu(),
  });

  const fold = foldsCase(process.platform);

  function installMenu(): void {
    Menu.setApplicationMenu(
      Menu.buildFromTemplate(
        buildMenuTemplate({
          platform: process.platform,
          appName: 'MIRRORZ Sheets',
          recentPaths: recent.paths(),
          isDev: !app.isPackaged,
          dispatch: (command, arg) => handleMenuCommand(command, arg),
        }),
      ),
    );
  }

  function handleMenuCommand(command: MenuCommand, arg?: string): void {
    if (command === 'file.new') {
      newWindow();
      return;
    }
    if (command === 'file.openRecent' && arg) {
      void openPathInWindow(arg);
      return;
    }
    if (command === 'file.clearRecent') {
      recent.clear();
      installMenu();
      return;
    }
    // Everything else belongs to the document, and the focused window is the
    // one that owns it.
    const focused = windows.all().find((doc) => doc.window.isFocused()) ?? windows.all()[0];
    if (focused) {
      focused.window.webContents.send(CHANNEL.command, command, arg);
      return;
    }
    // macOS keeps the menu bar up with every window closed, so Open… and the
    // rest are still clickable with nothing to send them to. Dropping the
    // command there makes the menu look broken; opening a window for it is what
    // every other document application does.
    const created = newWindow();
    void created.ready.then(() => {
      if (created.window.isDestroyed()) return;
      created.window.webContents.send(CHANNEL.command, command, arg);
    });
  }

  function newWindow(docId?: string): DocumentWindow {
    const doc = windows.create(docId);
    startAutosaveTicker(doc);
    attachClosePrompt(doc);
    return doc;
  }

  /**
   * The unsaved-changes prompt, in main rather than as a `beforeunload` in the
   * renderer, because a renderer that is wedged - a runaway recalculation, a
   * script that will not yield - is exactly when the prompt matters most and is
   * also when the renderer cannot draw it.
   *
   * Choosing Save cannot close the window here: the bytes only exist in the
   * renderer. The close is cancelled, a save is requested, and the save handler
   * closes the window once the file is actually on disk.
   */
  function attachClosePrompt(doc: DocumentWindow): void {
    doc.window.on('close', (event) => {
      if (!doc.dirty) return;
      event.preventDefault();
      const choice = dialog.showMessageBoxSync(doc.window, {
        type: 'warning',
        buttons: ['Save', "Don't Save", 'Cancel'],
        defaultId: 0,
        cancelId: 2,
        message: `Save changes to ${doc.displayName}?`,
        detail: 'If you do not save, the changes since your last save are lost. A recovery journal is kept until this window closes.',
      });
      if (choice === 0) {
        doc.closeAfterSave = true;
        doc.window.webContents.send(CHANNEL.command, 'file.save');
        return;
      }
      if (choice === 1) {
        doc.dirty = false;
        autosave.discard(doc.docId);
        doc.window.close();
      }
    });
  }

  /** Open a path, reusing an untouched blank window rather than stacking one. */
  async function openPathInWindow(path: string): Promise<void> {
    const existing = windows.forPath(path, fold);
    if (existing) {
      existing.window.focus();
      return;
    }
    const read = readDocument(path);
    if (!read.ok) {
      recent.remove(path);
      installMenu();
      dialog.showErrorBox('Cannot open that file', read.error);
      return;
    }
    const blank = windows.all().find((doc) => doc.filePath === null && !doc.dirty);
    const target = blank ?? newWindow();
    await target.ready;
    windows.setDocumentFile(target, path);
    recent.add(path, Date.now());
    installMenu();
    target.window.webContents.send(CHANNEL.openRequest, read.file);
  }

  installMenu();
  pendingOpens.attach((path) => void openPathInWindow(path));

  const recovered = await offerRecovery();
  if (windows.count() === 0 && !recovered) newWindow();

  app.on('activate', () => {
    // macOS keeps the application running with no windows; clicking the dock
    // icon is a request for one.
    if (windows.count() === 0) newWindow();
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
  });

  /**
   * Offer back anything a previous run left behind.
   *
   * A journal on disk means the process that owned it did not shut down
   * cleanly; there is no guesswork and no lock file to go stale. The user is
   * asked rather than told, because the recovered state can be older than what
   * they remember and silently applying it would be its own kind of data loss.
   */
  async function offerRecovery(): Promise<boolean> {
    const candidates = offerable(autosave.scan(Date.now()));
    if (candidates.length === 0) return false;

    const names = candidates.map((c) => `  • ${c.journal.displayName}`).join('\n');
    const answer = await dialog.showMessageBox({
      type: 'question',
      buttons: ['Recover', 'Discard'],
      defaultId: 0,
      cancelId: 1,
      message:
        candidates.length === 1
          ? 'MIRRORZ Sheets closed unexpectedly with unsaved work.'
          : `MIRRORZ Sheets closed unexpectedly with unsaved work in ${candidates.length} documents.`,
      detail: `${names}\n\nRecovering opens the documents with their unsaved changes. Your files on disk have not been altered.`,
    });
    if (answer.response !== 0) {
      for (const candidate of candidates) autosave.discard(candidate.journal.docId);
      return false;
    }

    for (const candidate of candidates) {
      // The recovered window IS the document the journal describes, so it takes
      // the journal's id. Without that the journal has no owner, is never
      // discarded on a clean close, and reappears in this prompt at every launch.
      const doc = newWindow(candidate.journal.docId);
      await doc.ready;
      if (candidate.journal.filePath && existsSync(candidate.journal.filePath)) {
        const read = readDocument(candidate.journal.filePath);
        if (read.ok) {
          windows.setDocumentFile(doc, candidate.journal.filePath);
          doc.window.webContents.send(CHANNEL.openRequest, read.file);
        }
      } else {
        doc.displayName = candidate.journal.displayName;
      }
      doc.window.webContents.send(CHANNEL.command, 'app.recover', candidate.journal.payload);
    }
    return true;
  }

  function startAutosaveTicker(doc: DocumentWindow): void {
    const timer = setInterval(() => {
      if (doc.window.isDestroyed()) return;
      if (!doc.dirty) return;
      doc.window.webContents.send(CHANNEL.command, 'app.autosaveTick');
    }, AUTOSAVE_INTERVAL_MS);
    // An autosave timer is not a reason to keep the process alive at exit.
    timer.unref?.();
    tickers.set(doc.window.id, timer);
  }

  function stopAutosaveTicker(doc: DocumentWindow): void {
    const timer = tickers.get(doc.window.id);
    if (timer) clearInterval(timer);
    tickers.delete(doc.window.id);
  }
}

export const APP_NAME = 'MIRRORZ Sheets';
