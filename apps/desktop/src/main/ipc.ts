/**
 * The IPC surface: seven file operations, plus the shell channel.
 *
 * Every handler starts the same way - find the document window the message came
 * from, and refuse if there isn't one. A message whose sender is not a window
 * we created is either a bug or something that should not be happening at all,
 * and in neither case is running the request the right answer.
 *
 * Validation is delegated to `validate.ts` so it can be tested on its own; this
 * file is the wiring, the filesystem calls and the dialogs.
 */

import { dialog, ipcMain } from 'electron';
import type { App, IpcMainInvokeEvent } from 'electron';
import { basename, dirname, extname, join } from 'node:path';
import {
  CHANNEL,
  type AppVersionInfo,
  type OpenedDocument,
  type RecentFileInfo,
  type SaveOutcome,
} from './channels.js';
import { readDocument, writeDocument } from './files.js';
import {
  validateJournalPayload,
  validateOpenPath,
  validateSaveAsRequest,
  validateSaveRequest,
} from './validate.js';
import { openDialogFilters, saveDialogFilters } from './paths.js';
import type { RecentFiles } from './recent.js';
import type { AutosaveStore } from './autosave.js';
import { AUTOSAVE_INTERVAL_MS } from './autosave.js';
import type { DocumentWindow, WindowManager } from './windows.js';
import { foldsCase } from './recent.js';

export interface IpcDependencies {
  app: App;
  windows: WindowManager;
  recent: RecentFiles;
  autosave: AutosaveStore;
  /** Injected so tests and the autosave rate limiter share one clock. */
  now: () => number;
  onRecentChanged: () => void;
}

export function registerIpc(deps: IpcDependencies): void {
  const { app, windows, recent, autosave } = deps;
  const fold = foldsCase(process.platform);

  function senderDoc(event: IpcMainInvokeEvent): DocumentWindow | undefined {
    return windows.forWebContentsId(event.sender.id);
  }

  function adopt(doc: DocumentWindow, file: OpenedDocument): void {
    windows.setDocumentFile(doc, file.path);
    recent.add(file.path, deps.now());
    deps.onRecentChanged();
  }

  ipcMain.handle(CHANNEL.openFile, async (event): Promise<OpenedDocument | null> => {
    const doc = senderDoc(event);
    if (!doc) return null;
    const result = await dialog.showOpenDialog(doc.window, {
      title: 'Open',
      properties: ['openFile'],
      filters: openDialogFilters(),
    });
    const chosen = result.filePaths[0];
    if (result.canceled || chosen === undefined) return null;

    const checked = validateOpenPath(chosen);
    if (!checked.ok) {
      await dialog.showMessageBox(doc.window, { type: 'error', message: 'Cannot open that file', detail: checked.reason });
      return null;
    }
    const read = readDocument(checked.value);
    if (!read.ok) {
      await dialog.showMessageBox(doc.window, { type: 'error', message: 'Cannot open that file', detail: read.error });
      return null;
    }
    adopt(doc, read.file);
    return read.file;
  });

  ipcMain.handle(CHANNEL.openPath, async (event, rawPath: unknown): Promise<OpenedDocument | null> => {
    const doc = senderDoc(event);
    if (!doc) return null;
    const checked = validateOpenPath(rawPath);
    if (!checked.ok) return null;

    // Already open somewhere: focus it rather than opening a second copy whose
    // edits would silently race the first.
    const existing = windows.forPath(checked.value, fold);
    if (existing && existing !== doc) {
      existing.window.focus();
      return null;
    }
    const read = readDocument(checked.value);
    if (!read.ok) {
      // A stale Recent Files entry is the common case; drop it so the menu
      // stops offering a file that is not there.
      recent.remove(checked.value);
      deps.onRecentChanged();
      await dialog.showMessageBox(doc.window, { type: 'error', message: 'Cannot open that file', detail: read.error });
      return null;
    }
    adopt(doc, read.file);
    return read.file;
  });

  ipcMain.handle(CHANNEL.saveFile, async (event, rawBytes: unknown, rawPath: unknown): Promise<SaveOutcome> => {
    const doc = senderDoc(event);
    if (!doc) return { ok: false, error: 'No document window' };
    const checked = validateSaveRequest(rawBytes, rawPath);
    if (!checked.ok) return abandonPendingClose(doc, { ok: false, error: checked.reason });

    const written = writeDocument(checked.value.path, checked.value.bytes);
    if (!written.ok) return abandonPendingClose(doc, { ok: false, error: written.error });

    windows.setDocumentFile(doc, checked.value.path);
    windows.setDirty(doc, false);
    recent.add(checked.value.path, deps.now());
    deps.onRecentChanged();
    // The file on disk is now the truth, so the journal describes nothing the
    // user would want back.
    autosave.discard(doc.docId);
    finishPendingClose(doc);
    return { ok: true, path: checked.value.path };
  });

  ipcMain.handle(CHANNEL.saveFileAs, async (event, rawBytes: unknown, rawName: unknown): Promise<SaveOutcome> => {
    const doc = senderDoc(event);
    if (!doc) return { ok: false, error: 'No document window' };
    const checked = validateSaveAsRequest(rawBytes, rawName);
    if (!checked.ok) return abandonPendingClose(doc, { ok: false, error: checked.reason });

    const suggested = checked.value.suggestedName;
    const result = await dialog.showSaveDialog(doc.window, {
      title: 'Save As',
      defaultPath: doc.filePath ? join(dirname(doc.filePath), suggested) : suggested,
      filters: saveDialogFilters(),
    });
    if (result.canceled || !result.filePath) return abandonPendingClose(doc, { ok: false });

    // The dialog can return a path with no extension when the user types one;
    // fill it in from the suggestion rather than writing an extensionless file.
    const chosen = extname(result.filePath) === ''
      ? `${result.filePath}${extname(suggested) || '.xlsx'}`
      : result.filePath;

    const revalidated = validateSaveRequest(checked.value.bytes, chosen);
    if (!revalidated.ok) return abandonPendingClose(doc, { ok: false, error: revalidated.reason });

    const written = writeDocument(revalidated.value.path, revalidated.value.bytes);
    if (!written.ok) return abandonPendingClose(doc, { ok: false, error: written.error });

    windows.setDocumentFile(doc, revalidated.value.path);
    windows.setDirty(doc, false);
    recent.add(revalidated.value.path, deps.now());
    deps.onRecentChanged();
    autosave.discard(doc.docId);
    finishPendingClose(doc);
    return { ok: true, path: revalidated.value.path };
  });

  // Sender-checked like every other handler. The recent list is a list of the
  // paths this person has been working in, which is worth something to anything
  // that should not be talking to us in the first place.
  ipcMain.handle(CHANNEL.recentFiles, (event): RecentFileInfo[] => {
    if (!senderDoc(event)) return [];
    return recent
      .list()
      .map((entry) => ({ path: entry.path, name: basename(entry.path), openedAt: entry.openedAt }));
  });

  ipcMain.handle(CHANNEL.version, (): AppVersionInfo => ({
    app: app.getVersion(),
    electron: process.versions.electron ?? 'unknown',
    chrome: process.versions.chrome ?? 'unknown',
    node: process.versions.node,
    platform: process.platform,
  }));

  ipcMain.handle(CHANNEL.autosave, (event, rawPayload: unknown): boolean => {
    const doc = senderDoc(event);
    if (!doc) return false;
    const checked = validateJournalPayload(rawPayload);
    if (!checked.ok) return false;

    // Rate limited in main rather than trusting the renderer's timer: a
    // renderer in a tight loop must not be able to hammer the disk.
    const now = deps.now();
    if (now - doc.lastAutosaveAt < AUTOSAVE_INTERVAL_MS / 2) return false;
    doc.lastAutosaveAt = now;

    const written = autosave.save({
      docId: doc.docId,
      filePath: doc.filePath,
      displayName: doc.displayName,
      payload: checked.value,
      now,
    });
    return written !== null;
  });

  ipcMain.on(CHANNEL.documentState, (event, state: unknown) => {
    const doc = windows.forWebContentsId(event.sender.id);
    if (!doc) return;
    if (typeof state !== 'object' || state === null) return;
    const record = state as { dirty?: unknown; displayName?: unknown };
    windows.setDirty(doc, record.dirty === true);
    if (typeof record.displayName === 'string' && record.displayName.length > 0 && doc.filePath === null) {
      doc.displayName = record.displayName.slice(0, 200);
      doc.window.setTitle(`${doc.displayName} — MIRRORZ Sheets`);
    }
  });

  function finishPendingClose(doc: DocumentWindow): void {
    if (!doc.closeAfterSave) return;
    doc.closeAfterSave = false;
    doc.window.close();
  }

  /**
   * The other half of `finishPendingClose`, and the reason both exist.
   *
   * `closeAfterSave` is set when the user answered Save to the close prompt, and
   * the window is closed by whichever save call succeeds next. If a save then
   * fails - or the user cancels the Save As dialog they were just handed - the
   * flag has to be cleared, or it stays armed: the user carries on working, and
   * the next ordinary Ctrl+S closes the window out from under them.
   */
  function abandonPendingClose(doc: DocumentWindow, outcome: SaveOutcome): SaveOutcome {
    doc.closeAfterSave = false;
    return outcome;
  }
}
