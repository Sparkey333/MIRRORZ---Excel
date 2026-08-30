/**
 * One window per open document, and the cage each one runs in.
 *
 * Per-document windows rather than tabs in a single window is a deliberate
 * departure from what Excel does on Windows. Two long-standing complaints fall
 * out of the single-window model: two workbooks with the same file name become
 * indistinguishable once they are tabs, and putting one workbook on each of two
 * monitors is either impossible or requires a second instance with its own
 * quirks. A window per document is what the operating system's own window
 * management already knows how to arrange, and the file name in the title bar
 * is then never ambiguous because the full path is one hover away.
 *
 * The security posture is applied here, in one place, per window and per
 * session, so there is no path by which a window is created without it.
 */

import { BrowserWindow, shell } from 'electron';
import type { Session } from 'electron';
import { basename } from 'node:path';
import { randomUUID } from 'node:crypto';
import {
  SECURE_WEB_PREFERENCES,
  contentSecurityPolicy,
  decideExternalOpen,
  isAllowedNavigation,
} from './security.js';

export interface DocumentWindow {
  window: BrowserWindow;
  /** Identity of the document for autosave purposes; outlives its file path. */
  docId: string;
  filePath: string | null;
  displayName: string;
  dirty: boolean;
  lastAutosaveAt: number;
  /** Set when the user chose Save in the close prompt; closes once saved. */
  closeAfterSave: boolean;
  /**
   * Resolves once the renderer has loaded, or the window has gone. Anything
   * pushed to a renderer - an open request, a recovered journal - has to wait
   * on this, or it is sent to a page that has not yet subscribed and is lost
   * with no error anywhere.
   */
  ready: Promise<void>;
}

export interface WindowManagerOptions {
  preloadPath: string;
  /** The built renderer's index.html, used when there is no dev server. */
  rendererFile: string;
  devServer?: string | undefined;
  onClosed?: (doc: DocumentWindow) => void;
}

/**
 * Session-wide policy. Applied to the default session before any window opens,
 * because a policy installed after the first load is a policy the first load
 * did not have.
 */
export function installSessionPolicies(session: Session, devServer?: string): void {
  session.webRequest.onHeadersReceived((details, callback) => {
    callback({
      responseHeaders: {
        ...details.responseHeaders,
        'Content-Security-Policy': [contentSecurityPolicy({ devServer })],
      },
    });
  });

  // A spreadsheet has no business asking for the camera, the microphone, the
  // user's location or notifications. Denying by default means a future feature
  // has to ask for its permission explicitly rather than inheriting one.
  session.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
  session.setPermissionCheckHandler(() => false);

  // Belt and braces with `spellcheck: false` in SECURE_WEB_PREFERENCES. The
  // spellchecker fetches Hunspell dictionaries from Google's CDN from the
  // browser process, where no CSP applies, so it is the one component that
  // could make this application talk to the network behind the user's back.
  // Turning it off at the session as well means a window created with different
  // webPreferences by some future code path still cannot start that download.
  session.setSpellCheckerEnabled?.(false);
}

export class WindowManager {
  private readonly docs = new Map<number, DocumentWindow>();

  constructor(private readonly options: WindowManagerOptions) {}

  all(): DocumentWindow[] {
    return [...this.docs.values()];
  }

  count(): number {
    return this.docs.size;
  }

  forWebContentsId(id: number): DocumentWindow | undefined {
    for (const doc of this.docs.values()) {
      if (doc.window.webContents.id === id) return doc;
    }
    return undefined;
  }

  /** An already-open window for this file, so opening twice focuses instead. */
  forPath(path: string, foldCase: boolean): DocumentWindow | undefined {
    const wanted = foldCase ? path.toLowerCase() : path;
    for (const doc of this.docs.values()) {
      if (doc.filePath === null) continue;
      const held = foldCase ? doc.filePath.toLowerCase() : doc.filePath;
      if (held === wanted) return doc;
    }
    return undefined;
  }

  /**
   * `docId` is supplied only when recovering: the window continues a document
   * that already has a journal on disk, and inheriting that identity is what
   * lets the journal be discarded when this window closes cleanly. A recovered
   * window given a fresh id would leave the old journal in the index with
   * nothing that will ever delete it, and it would be offered back at every
   * launch until it aged out.
   */
  create(docId?: string): DocumentWindow {
    const window = new BrowserWindow({
      width: 1280,
      height: 820,
      minWidth: 640,
      minHeight: 420,
      // Nothing is shown until the first paint; otherwise the window flashes
      // the background colour before the grid exists.
      show: false,
      backgroundColor: '#ffffff',
      title: 'MIRRORZ Sheets',
      webPreferences: {
        ...SECURE_WEB_PREFERENCES,
        preload: this.options.preloadPath,
      },
    });

    let settle = (): void => {};
    const ready = new Promise<void>((resolve) => {
      settle = resolve;
    });

    const doc: DocumentWindow = {
      window,
      docId: docId ?? randomUUID().replace(/-/g, ''),
      filePath: null,
      displayName: 'Untitled',
      dirty: false,
      lastAutosaveAt: 0,
      closeAfterSave: false,
      ready,
    };
    this.docs.set(window.id, doc);

    this.harden(window);

    window.webContents.once('did-finish-load', () => settle());
    window.once('ready-to-show', () => window.show());
    window.on('closed', () => {
      this.docs.delete(window.id);
      settle();
      this.options.onClosed?.(doc);
    });

    if (this.options.devServer) {
      void window.loadURL(this.options.devServer);
    } else {
      void window.loadFile(this.options.rendererFile);
    }
    return doc;
  }

  /** Navigation and child-window policy for one window's contents. */
  private harden(window: BrowserWindow): void {
    const devServer = this.options.devServer;

    const policy = { devServer, rendererFile: this.options.rendererFile };

    window.webContents.on('will-navigate', (event, url) => {
      if (!isAllowedNavigation(url, policy)) {
        event.preventDefault();
        const decision = decideExternalOpen(url);
        if (decision.action === 'external') void shell.openExternal(decision.url);
      }
    });

    // Redirects land here rather than in will-navigate, and a link that
    // redirects offsite is the interesting case.
    window.webContents.on('will-redirect', (event, url) => {
      if (!isAllowedNavigation(url, policy)) event.preventDefault();
    });

    window.webContents.setWindowOpenHandler(({ url }) => {
      const decision = decideExternalOpen(url);
      if (decision.action === 'external') void shell.openExternal(decision.url);
      // Always deny: a second Electron window for foreign content would run
      // with this application's preload attached.
      return { action: 'deny' };
    });

    window.webContents.on('will-attach-webview', (event) => event.preventDefault());
  }

  setDocumentFile(doc: DocumentWindow, filePath: string): void {
    doc.filePath = filePath;
    doc.displayName = basename(filePath);
    doc.window.setTitle(`${doc.displayName} — MIRRORZ Sheets`);
    // Gives macOS the proxy icon and the drag-out behaviour for the real file.
    doc.window.setRepresentedFilename(filePath);
  }

  setDirty(doc: DocumentWindow, dirty: boolean): void {
    doc.dirty = dirty;
    doc.window.setDocumentEdited(dirty);
  }
}
