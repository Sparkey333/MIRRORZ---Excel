/**
 * The renderer's half of the shell channel.
 *
 * The main process already sends menu commands, pushes files the operating
 * system asked us to open, and runs the autosave clock. Nothing in the renderer
 * listened, and the consequences were not cosmetic:
 *
 *   Every item in the native menu bar did nothing at all.
 *
 *   A double-clicked spreadsheet opened a window that stayed blank, because the
 *   file arrives on `openRequest` and nobody was subscribed to it.
 *
 *   Worst: main learns a document has unsaved changes only from
 *   `setDocumentState`, and BOTH the close prompt and the autosave ticker are
 *   gated on that flag. With nothing reporting it, closing a window with
 *   unsaved work discarded it without asking, and no recovery journal was ever
 *   written to get it back. Two halves of a working contract were shipped as
 *   one half and a comment claiming they connected.
 *
 * Everything here is a no-op when the bridge is absent, which is the case in a
 * browser build and in every test.
 */

import { useEffect, useRef } from 'react';
import type { AppController } from './controller.js';
import { buildCommands } from './commands.js';

/** A file the shell hands us, unasked. */
export interface PushedFile {
  path?: string;
  name: string;
  data: Uint8Array;
}

declare global {
  interface Window {
    mirrorz?: {
      onOpenRequest?(callback: (file: PushedFile) => void): () => void;
    };
    mirrorzShell?: {
      onCommand?(callback: (command: string, arg?: string) => void): () => void;
      setDocumentState?(state: { dirty: boolean; displayName?: string }): void;
      autosave?(payload: string): Promise<boolean>;
    };
  }
}

/** What the shell can ask the application to do that the controller cannot. */
export interface ShellActions {
  openFile: () => void;
  save: (as: boolean) => void;
  /** A workbook file that arrived from outside, however it arrived. */
  acceptFile: (file: { name: string; data: Uint8Array }) => void;
  /** The bytes to journal, which are the bytes a save would have written. */
  serialize: () => Uint8Array;
}

/**
 * The recovery journal crosses as a string, and the honest payload for a
 * spreadsheet is the workbook written through the same writer a save uses: a
 * recovery that travels a different code path from a save is a recovery nobody
 * has tested. `btoa` over a binary string is the encoder a sandboxed renderer
 * has - there is no Buffer in one.
 */
export function bytesToBase64(bytes: Uint8Array): string {
  let binary = '';
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  return btoa(binary);
}

export function base64ToBytes(text: string): Uint8Array {
  const binary = atob(text);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

export type MenuOutcome = 'ran' | 'unimplemented' | 'unknown';

/**
 * Menu items that have an exact twin in the command catalogue.
 *
 * Dispatched through the catalogue rather than reimplemented, so a menu item
 * and its palette twin cannot drift into meaning different things - which is
 * the failure the catalogue exists to prevent in the first place.
 */
const MENU_TO_COMMAND: Readonly<Record<string, string>> = {
  'edit.undo': 'edit.undo',
  'edit.redo': 'edit.redo',
  'insert.sheet': 'sheet.add',
  'view.toggleHistory': 'view.history',
  'data.sortAscending': 'data.sort.asc',
  'data.sortDescending': 'data.sort.desc',
  'data.recalculate': 'calc.now',
};

/** Menu items this build draws but does not yet perform. */
const NOT_YET_IMPLEMENTED: ReadonlySet<string> = new Set([
  'view.freezePanes',
  'view.toggleGridlines',
  'view.toggleFormulaBar',
  'insert.function',
  'format.cells',
  'data.filter',
  'data.removeDuplicates',
  'data.textToColumns',
  'help.documentation',
  'help.shortcuts',
  'help.about',
]);

/**
 * Perform one command from the native menu.
 *
 * Exported and pure over the controller so the whole table can be asserted
 * without a window; a menu that can only be tested by clicking it is a menu
 * that quietly rots.
 */
export function runMenuCommand(
  controller: AppController,
  actions: ShellActions,
  command: string,
  arg?: string,
): MenuOutcome {
  const mapped = MENU_TO_COMMAND[command];
  if (mapped !== undefined) {
    const entry = buildCommands(controller, controller.getSnapshot()).find((c) => c.id === mapped);
    if (!entry) return 'unknown';
    entry.run();
    return 'ran';
  }

  const range = controller.selectionRange();
  const font = () => controller.uniformStyle((s) => s.font);

  switch (command) {
    case 'file.open':
      actions.openFile();
      return 'ran';
    case 'file.save':
      actions.save(false);
      return 'ran';
    case 'file.saveAs':
      actions.save(true);
      return 'ran';

    case 'edit.find':
    case 'edit.replace':
      controller.setFind(true);
      return 'ran';
    case 'edit.goTo':
      controller.setPalette(true, 'all');
      return 'ran';
    case 'view.commandPalette':
      controller.setPalette(true, 'command');
      return 'ran';

    // Above and below differ only in where the block lands: inserting below the
    // selection means starting one past its end.
    case 'insert.rowsAbove':
      controller.insertRows(range.start.row, range.end.row - range.start.row + 1);
      return 'ran';
    case 'insert.rowsBelow':
      controller.insertRows(range.end.row + 1, range.end.row - range.start.row + 1);
      return 'ran';
    case 'insert.columnsLeft':
      controller.insertCols(range.start.col, range.end.col - range.start.col + 1);
      return 'ran';
    case 'insert.columnsRight':
      controller.insertCols(range.end.col + 1, range.end.col - range.start.col + 1);
      return 'ran';

    case 'format.bold':
      controller.applyStyle({ font: { bold: !font()?.bold } }, 'Bold');
      return 'ran';
    case 'format.italic':
      controller.applyStyle({ font: { italic: !font()?.italic } }, 'Italic');
      return 'ran';
    case 'format.underline':
      controller.applyStyle(
        { font: { underline: font()?.underline === 'single' ? 'none' : 'single' } },
        'Underline',
      );
      return 'ran';
    case 'format.numberGeneral':
      controller.setNumberFormat(undefined);
      return 'ran';
    case 'format.numberCurrency':
      controller.setNumberFormat('$#,##0.00');
      return 'ran';
    case 'format.numberPercent':
      controller.setNumberFormat('0.00%');
      return 'ran';
    case 'format.numberDate':
      controller.setNumberFormat('yyyy-mm-dd');
      return 'ran';

    default:
      // `file.new`, `file.close`, `file.openRecent` and `file.clearRecent` never
      // reach the renderer: main acts on those itself, because they are about
      // the window and the recent-files list rather than the document.
      void arg;
      return NOT_YET_IMPLEMENTED.has(command) ? 'unimplemented' : 'unknown';
  }
}

function shellBridge(): Window['mirrorzShell'] {
  return typeof window === 'undefined' ? undefined : window.mirrorzShell;
}

function fileBridge(): Window['mirrorz'] {
  return typeof window === 'undefined' ? undefined : window.mirrorz;
}

/**
 * Subscribe to the shell for as long as the controller lives.
 *
 * `actions` is read through a ref rather than being an effect dependency: the
 * callbacks are rebuilt on every render, and re-registering the listeners on
 * every keystroke is how a menu command lands in the gap and does nothing.
 */
export function useShell(controller: AppController, actions: ShellActions): void {
  const latest = useRef(actions);
  latest.current = actions;

  useEffect(() => {
    const shell = shellBridge();
    const files = fileBridge();
    const stops: (() => void)[] = [];

    const onCommand = shell?.onCommand;
    if (onCommand) {
      stops.push(
        onCommand((command, arg) => {
          // Two things on this channel are not menu items and must not be
          // reported as unrecognised ones.
          if (command === 'app.autosaveTick') {
            const autosave = shellBridge()?.autosave;
            if (!autosave) return;
            try {
              void autosave(bytesToBase64(latest.current.serialize())).catch(() => {
                // Main rate-limits and asks again; a failed journal is not a
                // reason to interrupt whoever is typing.
              });
            } catch {
              // Serialising a workbook can fail on a file we cannot yet write.
            }
            return;
          }
          if (command === 'app.recover') {
            if (arg === undefined) return;
            try {
              // Recovered bytes take the same path as an opened file, so what
              // comes back is exactly what a save would have produced.
              latest.current.acceptFile({ name: 'Recovered', data: base64ToBytes(arg) });
            } catch {
              controller.setMessage('The recovered copy could not be read.');
            }
            return;
          }

          const outcome = runMenuCommand(controller, latest.current, command, arg);
          // Silence is indistinguishable from a broken menu, so a command this
          // build cannot perform says so rather than doing nothing.
          if (outcome === 'unimplemented') controller.setMessage('That is not built yet');
          else if (outcome === 'unknown') {
            controller.setMessage(`Unrecognised menu command: ${command}`);
          }
        }),
      );
    }

    const onOpenRequest = files?.onOpenRequest;
    if (onOpenRequest) {
      stops.push(
        onOpenRequest((file) => {
          latest.current.acceptFile({ name: file.name, data: file.data });
        }),
      );
    }

    return () => {
      for (const stop of stops) stop();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [controller]);
}

/**
 * Keep main's idea of this document in step with the renderer's.
 *
 * This is the flag the unsaved-changes prompt and the autosave ticker are both
 * gated on, so it is pushed on every change rather than on a timer: the moment
 * it has to be right is the moment somebody reaches for the close button.
 */
export function useDocumentState(dirty: boolean, displayName: string): void {
  useEffect(() => {
    shellBridge()?.setDocumentState?.({ dirty, displayName });
  }, [dirty, displayName]);
}
