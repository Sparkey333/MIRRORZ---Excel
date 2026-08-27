/**
 * The preload bridge: the only thing the renderer can reach outside itself.
 *
 * Rules this file follows, all of them load-bearing:
 *
 *   It exposes functions, never objects with methods that could be walked, and
 *   never anything from Electron itself. Handing over `ipcRenderer` - even
 *   partially applied - is the classic contextBridge mistake, because a page
 *   that can call `send` with an arbitrary channel can talk to every handler in
 *   the main process rather than the seven it is meant to have.
 *
 *   It validates nothing of consequence. All checking happens in main, because
 *   the preload shares a process with the page: a renderer compromise that
 *   reaches the bridge would sail past any check made here. What it does do is
 *   normalise shapes, so main receives a copy of a Uint8Array rather than
 *   whatever exotic object the page passed.
 *
 *   Callbacks registered by the renderer are wrapped: the raw IpcRendererEvent
 *   never reaches the page, since it carries `sender` and with it a route back
 *   into the main process.
 *
 * `window.mirrorzHost` is a small compatibility shim for the renderer's
 * existing FileHost interface, expressed entirely in terms of the same seven
 * functions - it adds no capability of its own.
 */

import { contextBridge, ipcRenderer } from 'electron';
import type { IpcRendererEvent } from 'electron';
import {
  CHANNEL,
  type AppVersionInfo,
  type OpenedDocument,
  type RecentFileInfo,
  type SaveOutcome,
} from '../main/channels.js';

/**
 * Copy bytes into a plain Uint8Array before they cross.
 *
 * The structured clone algorithm rejects some objects outright and, worse,
 * silently accepts an object whose `length` disagrees with its buffer. Copying
 * here means main always receives something whose length it can trust before it
 * even starts validating.
 */
function toBytes(value: unknown): Uint8Array {
  if (value instanceof Uint8Array) return new Uint8Array(value);
  if (value instanceof ArrayBuffer) return new Uint8Array(value.slice(0));
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength));
  }
  // Deliberately not coerced: main will reject an empty payload with a message.
  return new Uint8Array(0);
}

const api = {
  /** Show the open dialog. Resolves null when the user cancels. */
  openFile(): Promise<OpenedDocument | null> {
    return ipcRenderer.invoke(CHANNEL.openFile) as Promise<OpenedDocument | null>;
  },

  /** Open a specific path, as used by the Recent Files list. */
  openPath(path: string): Promise<OpenedDocument | null> {
    return ipcRenderer.invoke(CHANNEL.openPath, String(path)) as Promise<OpenedDocument | null>;
  },

  /** Overwrite a file the user has already chosen. */
  saveFile(bytes: Uint8Array, path: string): Promise<SaveOutcome> {
    return ipcRenderer.invoke(CHANNEL.saveFile, toBytes(bytes), String(path)) as Promise<SaveOutcome>;
  },

  /** Show the save dialog, seeded with a name the renderer suggests. */
  saveFileAs(bytes: Uint8Array, suggestedName: string): Promise<SaveOutcome> {
    return ipcRenderer.invoke(
      CHANNEL.saveFileAs,
      toBytes(bytes),
      String(suggestedName),
    ) as Promise<SaveOutcome>;
  },

  recentFiles(): Promise<RecentFileInfo[]> {
    return ipcRenderer.invoke(CHANNEL.recentFiles) as Promise<RecentFileInfo[]>;
  },

  /**
   * Files the operating system wants opened in this window: a double-clicked
   * document, a second launch forwarded by the single-instance lock, or a
   * Recent Files pick from the native menu. Returns an unsubscribe function.
   */
  onOpenRequest(callback: (file: OpenedDocument) => void): () => void {
    const listener = (_event: IpcRendererEvent, file: OpenedDocument): void => {
      callback(file);
    };
    ipcRenderer.on(CHANNEL.openRequest, listener);
    return () => {
      ipcRenderer.removeListener(CHANNEL.openRequest, listener);
    };
  },

  getVersion(): Promise<AppVersionInfo> {
    return ipcRenderer.invoke(CHANNEL.version) as Promise<AppVersionInfo>;
  },
};

/**
 * The shell channel, kept separate from the file API above so the file surface
 * stays exactly seven functions.
 *
 * This carries menu commands, the periodic autosave of the command log, and the
 * recovery handoff after an unclean exit. None of it can read or write a file
 * the user did not choose: `autosave` hands main an opaque string that main
 * writes to a journal path of its own choosing.
 */
const shell = {
  /** Menu items and shell events. Returns an unsubscribe function. */
  onCommand(callback: (command: string, arg?: string) => void): () => void {
    const listener = (_event: IpcRendererEvent, command: string, arg?: string): void => {
      callback(command, arg);
    };
    ipcRenderer.on(CHANNEL.command, listener);
    return () => {
      ipcRenderer.removeListener(CHANNEL.command, listener);
    };
  },

  /** Hand main a serialized command log to journal. Rate-limited in main. */
  autosave(payload: string): Promise<boolean> {
    return ipcRenderer.invoke(CHANNEL.autosave, String(payload)) as Promise<boolean>;
  },

  /** Lets main show the modified indicator and prompt before closing. */
  setDocumentState(state: { dirty: boolean; displayName?: string }): void {
    ipcRenderer.send(CHANNEL.documentState, {
      dirty: Boolean(state?.dirty),
      displayName: typeof state?.displayName === 'string' ? state.displayName : undefined,
    });
  },
};

/** The renderer's existing FileHost shape, in terms of the API above. */
const legacyHost = {
  async openFile(): Promise<{ name: string; data: Uint8Array } | null> {
    const file = await api.openFile();
    return file ? { name: file.name, data: file.data } : null;
  },
  async saveFile(name: string, data: Uint8Array): Promise<boolean> {
    const recent = await api.recentFiles();
    const known = recent.find((entry) => entry.name === name);
    // Without a known path this is a first save, which has to go through the
    // dialog: silently inventing a location is how files end up in a directory
    // the user cannot find.
    const outcome = known
      ? await api.saveFile(data, known.path)
      : await api.saveFileAs(data, name);
    return outcome.ok;
  },
  async saveFileAs(name: string, data: Uint8Array): Promise<string | null> {
    const outcome = await api.saveFileAs(data, name);
    return outcome.ok ? (outcome.path ?? null) : null;
  },
};

contextBridge.exposeInMainWorld('mirrorz', api);
contextBridge.exposeInMainWorld('mirrorzShell', shell);
contextBridge.exposeInMainWorld('mirrorzHost', legacyHost);

export type MirrorzApi = typeof api;
export type MirrorzShell = typeof shell;
