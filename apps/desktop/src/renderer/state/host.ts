/**
 * The boundary between the renderer and whatever is hosting it.
 *
 * In the packaged application that is the Electron main process, reached through
 * a preload bridge on `window.mirrorzHost`. In a browser, and in tests, it is
 * the fallbacks here. Keeping the renderer ignorant of which one it has is what
 * lets the whole shell be developed and tested without launching Electron, and
 * it is also the reason no component anywhere calls `require` or touches `fs`.
 */

import {
  looksLikeOds,
  looksLikeXls,
  parseDelimited,
  readOds,
  readXls,
  readXlsx,
  writeXlsx,
} from '@mirrorz/formats';
import { Document, Workbook } from '@mirrorz/core';
import { Engine, createRegistry } from '@mirrorz/formula';
import { AppController } from './controller.js';

export interface OpenedFile {
  name: string;
  data: Uint8Array;
}

export interface FileHost {
  openFile(): Promise<OpenedFile | null>;
  saveFile(name: string, data: Uint8Array): Promise<boolean>;
  saveFileAs(name: string, data: Uint8Array): Promise<string | null>;
}

declare global {
  interface Window {
    mirrorzHost?: Partial<FileHost>;
  }
}

/** A host that asks the browser, used when no Electron bridge is present. */
export const browserHost: FileHost = {
  async openFile() {
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.xlsx,.xlsm,.xls,.csv,.tsv,.txt';
      input.onchange = () => {
        const file = input.files?.[0];
        if (!file) {
          resolve(null);
          return;
        }
        void file.arrayBuffer().then((buffer) => {
          resolve({ name: file.name, data: new Uint8Array(buffer) });
        });
      };
      input.click();
    });
  },
  async saveFile(name, data) {
    downloadBytes(name, data);
    return true;
  },
  async saveFileAs(name, data) {
    downloadBytes(name, data);
    return name;
  },
};

function downloadBytes(name: string, data: Uint8Array): void {
  const blob = new Blob([data as BlobPart], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  URL.revokeObjectURL(url);
}

/** The Electron bridge when it exists, falling back method by method. */
export function resolveHost(): FileHost {
  const bridge = typeof window === 'undefined' ? undefined : window.mirrorzHost;
  if (!bridge) return browserHost;
  return {
    openFile: bridge.openFile?.bind(bridge) ?? browserHost.openFile,
    saveFile: bridge.saveFile?.bind(bridge) ?? browserHost.saveFile,
    saveFileAs: bridge.saveFileAs?.bind(bridge) ?? browserHost.saveFileAs,
  };
}

export function isDelimitedName(name: string): boolean {
  return /\.(csv|tsv|txt)$/i.test(name);
}

export interface OpenResult {
  controller: AppController;
  warnings: string[];
}

/**
 * Open a workbook file into a fresh controller.
 *
 * A new Document rather than a reset one, because the command log is the
 * document's identity: reusing it would leave the previous file's history
 * pointing at cells that no longer exist.
 */
export function openWorkbook(file: OpenedFile): OpenResult {
  const result = readWorkbookBytes(file.data);
  const doc = new Document(result.workbook);
  const registry = createRegistry();
  const engine = new Engine(doc, registry);
  engine.indexWorkbook();
  const controller = new AppController(doc, engine, registry);
  controller.setFileName(file.name);
  controller.markSaved();
  return { controller, warnings: result.warnings };
}

/**
 * Decide what a workbook actually is, by looking rather than by its name.
 *
 * The shell will hand us anything its dialog accepts, and that includes `.xls`
 * and `.ods`. Reading every one of them as xlsx would throw on the two formats
 * this application specifically advertises support for. The container is
 * sniffed rather than taken from the extension for the same reason the command
 * line does it: a `.xls` that is really a zip and an `.xlsx` that is really a
 * compound file are both common, and the bytes cannot be wrong about it.
 */
function readWorkbookBytes(data: Uint8Array): { workbook: Workbook; warnings: string[] } {
  if (looksLikeXls(data)) {
    const read = readXls(data);
    return { workbook: read.workbook, warnings: read.warnings };
  }
  if (looksLikeOds(data)) {
    const read = readOds(data);
    return { workbook: read.workbook, warnings: read.warnings };
  }
  const read = readXlsx(data);
  return { workbook: read.workbook, warnings: read.warnings };
}

/** Parse a delimited file into rows, for the import review to inspect. */
export function readDelimited(file: OpenedFile): { rows: string[][]; warnings: string[] } {
  // parseDelimited sniffs the byte-order mark and the delimiter itself, so the
  // bytes go straight in: guessing either here would only get it wrong twice.
  const parsed = parseDelimited(file.data);
  return { rows: parsed.rows, warnings: parsed.warnings.map((w) => w.message) };
}

export function serializeWorkbook(workbook: Workbook): Uint8Array {
  return writeXlsx(workbook);
}
