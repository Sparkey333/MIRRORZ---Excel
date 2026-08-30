/**
 * The complete list of IPC channels, shared by the main process and the preload.
 *
 * One file, imported by both sides, so a channel cannot be renamed on one side
 * and silently stop working on the other. The renderer never sees these names:
 * it sees the seven functions the preload exposes.
 *
 * This module must stay free of Electron and Node imports - the preload is
 * bundled for a sandboxed context where neither exists.
 */

export const CHANNEL = {
  /** Renderer to main, invoke/handle. */
  openFile: 'mirrorz:file/open',
  openPath: 'mirrorz:file/open-path',
  saveFile: 'mirrorz:file/save',
  saveFileAs: 'mirrorz:file/save-as',
  recentFiles: 'mirrorz:file/recent',
  version: 'mirrorz:app/version',
  autosave: 'mirrorz:shell/autosave',
  documentState: 'mirrorz:shell/document-state',

  /** Main to renderer, send. */
  openRequest: 'mirrorz:file/open-request',
  command: 'mirrorz:shell/command',
} as const;

export type Channel = (typeof CHANNEL)[keyof typeof CHANNEL];

/** A file handed to the renderer. `data` crosses as a structured-cloned copy. */
export interface OpenedDocument {
  path: string;
  name: string;
  data: Uint8Array;
}

export interface SaveOutcome {
  ok: boolean;
  /** Where it went, absent when the user cancelled the dialog. */
  path?: string;
  /** Present only on failure, and safe to show: never a stack trace. */
  error?: string;
}

export interface RecentFileInfo {
  path: string;
  name: string;
  openedAt: number;
}

export interface AppVersionInfo {
  app: string;
  electron: string;
  chrome: string;
  node: string;
  platform: string;
}
