/**
 * The native menu, built as data.
 *
 * `buildMenuTemplate` returns a plain array and imports nothing from Electron at
 * runtime - the Electron types are erased at compile time - so the entire menu,
 * including which accelerators are bound and what each item dispatches, is
 * checked by ordinary unit tests. Only `installMenu` touches `Menu`.
 *
 * Two decisions worth stating:
 *
 * Undo and Redo are NOT the built-in `undo`/`redo` roles. Those roles reach the
 * focused text field's editing history, which is not the document's history.
 * Ours has to run through the command log, or undo would mean one thing while
 * typing in a cell and another thing everywhere else - the exact confusion that
 * makes undo in a spreadsheet feel untrustworthy.
 *
 * Cut, Copy and Paste ARE roles, for the opposite reason: a sandboxed renderer
 * has no direct clipboard access, and the roles are what make the native
 * clipboard, including its system-wide keyboard shortcuts, work at all.
 */

import type { MenuItemConstructorOptions } from 'electron';
import { recentMenuLabels } from './recent.js';

/**
 * Everything the menu can ask the renderer to do. A closed union rather than
 * free strings, so a typo in a menu item is a compile error and the renderer
 * has an exhaustive list to switch over.
 */
export type MenuCommand =
  | 'file.new'
  | 'file.open'
  | 'file.openRecent'
  | 'file.clearRecent'
  | 'file.save'
  | 'file.saveAs'
  | 'file.close'
  | 'edit.undo'
  | 'edit.redo'
  | 'edit.find'
  | 'edit.replace'
  | 'edit.goTo'
  | 'edit.selectAll'
  | 'view.freezePanes'
  | 'view.toggleGridlines'
  | 'view.toggleFormulaBar'
  | 'view.toggleHistory'
  | 'view.commandPalette'
  | 'insert.rowsAbove'
  | 'insert.rowsBelow'
  | 'insert.columnsLeft'
  | 'insert.columnsRight'
  | 'insert.sheet'
  | 'insert.function'
  | 'format.bold'
  | 'format.italic'
  | 'format.underline'
  | 'format.numberGeneral'
  | 'format.numberCurrency'
  | 'format.numberPercent'
  | 'format.numberDate'
  | 'format.cells'
  | 'data.sortAscending'
  | 'data.sortDescending'
  | 'data.filter'
  | 'data.removeDuplicates'
  | 'data.textToColumns'
  | 'data.recalculate'
  | 'help.documentation'
  | 'help.shortcuts'
  | 'help.about';

export type MenuDispatch = (command: MenuCommand, arg?: string) => void;

export interface MenuOptions {
  platform: NodeJS.Platform;
  appName: string;
  recentPaths: readonly string[];
  dispatch: MenuDispatch;
  /** Adds the developer-tools item; absent from shipped builds. */
  isDev?: boolean;
}

function item(
  label: string,
  command: MenuCommand,
  dispatch: MenuDispatch,
  accelerator?: string,
): MenuItemConstructorOptions {
  const options: MenuItemConstructorOptions = { label, click: () => dispatch(command) };
  if (accelerator) options.accelerator = accelerator;
  return options;
}

function recentSubmenu(
  paths: readonly string[],
  dispatch: MenuDispatch,
): MenuItemConstructorOptions[] {
  if (paths.length === 0) {
    return [{ label: 'No recent files', enabled: false }];
  }
  const labels = recentMenuLabels([...paths]);
  const entries: MenuItemConstructorOptions[] = paths.map((path, index) => ({
    label: labels[index] ?? path,
    // The full path in the tooltip, since the label is deliberately short.
    toolTip: path,
    click: () => dispatch('file.openRecent', path),
  }));
  entries.push({ type: 'separator' });
  entries.push({ label: 'Clear Recent Files', click: () => dispatch('file.clearRecent') });
  return entries;
}

export function buildMenuTemplate(opts: MenuOptions): MenuItemConstructorOptions[] {
  const { platform, dispatch } = opts;
  const isMac = platform === 'darwin';
  const template: MenuItemConstructorOptions[] = [];

  if (isMac) {
    // macOS puts About, Preferences, Services, Hide and Quit in an application
    // menu named after the app; omitting it does not remove those commands, it
    // just leaves them in a menu labelled "Electron".
    template.push({
      label: opts.appName,
      submenu: [
        { label: `About ${opts.appName}`, click: () => dispatch('help.about') },
        { type: 'separator' },
        { role: 'services' },
        { type: 'separator' },
        { role: 'hide' },
        { role: 'hideOthers' },
        { role: 'unhide' },
        { type: 'separator' },
        { role: 'quit' },
      ],
    });
  }

  template.push({
    label: '&File',
    submenu: [
      item('New', 'file.new', dispatch, 'CmdOrCtrl+N'),
      item('Open…', 'file.open', dispatch, 'CmdOrCtrl+O'),
      { label: 'Open Recent', submenu: recentSubmenu(opts.recentPaths, dispatch) },
      { type: 'separator' },
      item('Save', 'file.save', dispatch, 'CmdOrCtrl+S'),
      item('Save As…', 'file.saveAs', dispatch, 'CmdOrCtrl+Shift+S'),
      { type: 'separator' },
      { role: 'close', label: 'Close Window', accelerator: 'CmdOrCtrl+W' },
      ...(isMac ? [] : [{ role: 'quit' as const, label: 'Exit' }]),
    ],
  });

  template.push({
    label: '&Edit',
    submenu: [
      item('Undo', 'edit.undo', dispatch, 'CmdOrCtrl+Z'),
      item('Redo', 'edit.redo', dispatch, isMac ? 'Cmd+Shift+Z' : 'Ctrl+Y'),
      { type: 'separator' },
      { role: 'cut' },
      { role: 'copy' },
      { role: 'paste' },
      { role: 'pasteAndMatchStyle', label: 'Paste Values Only' },
      { type: 'separator' },
      item('Select All', 'edit.selectAll', dispatch, 'CmdOrCtrl+A'),
      { type: 'separator' },
      item('Find…', 'edit.find', dispatch, 'CmdOrCtrl+F'),
      item('Replace…', 'edit.replace', dispatch, isMac ? 'Cmd+Alt+F' : 'Ctrl+H'),
      item('Go To…', 'edit.goTo', dispatch, 'F5'),
    ],
  });

  template.push({
    label: '&View',
    submenu: [
      item('Command Palette…', 'view.commandPalette', dispatch, 'CmdOrCtrl+Shift+P'),
      { type: 'separator' },
      item('Freeze Panes', 'view.freezePanes', dispatch),
      item('Gridlines', 'view.toggleGridlines', dispatch),
      item('Formula Bar', 'view.toggleFormulaBar', dispatch),
      item('History', 'view.toggleHistory', dispatch, 'CmdOrCtrl+Shift+H'),
      { type: 'separator' },
      { role: 'resetZoom' },
      { role: 'zoomIn' },
      { role: 'zoomOut' },
      { type: 'separator' },
      { role: 'togglefullscreen' },
      ...(opts.isDev ? [{ role: 'toggleDevTools' as const }] : []),
    ],
  });

  template.push({
    label: '&Insert',
    submenu: [
      item('Rows Above', 'insert.rowsAbove', dispatch),
      item('Rows Below', 'insert.rowsBelow', dispatch),
      item('Columns Left', 'insert.columnsLeft', dispatch),
      item('Columns Right', 'insert.columnsRight', dispatch),
      { type: 'separator' },
      item('Sheet', 'insert.sheet', dispatch),
      item('Function…', 'insert.function', dispatch, 'CmdOrCtrl+Shift+A'),
    ],
  });

  template.push({
    label: 'F&ormat',
    submenu: [
      item('Bold', 'format.bold', dispatch, 'CmdOrCtrl+B'),
      item('Italic', 'format.italic', dispatch, 'CmdOrCtrl+I'),
      item('Underline', 'format.underline', dispatch, 'CmdOrCtrl+U'),
      { type: 'separator' },
      {
        label: 'Number',
        submenu: [
          item('General', 'format.numberGeneral', dispatch, 'CmdOrCtrl+Shift+~'),
          item('Currency', 'format.numberCurrency', dispatch, 'CmdOrCtrl+Shift+4'),
          item('Percent', 'format.numberPercent', dispatch, 'CmdOrCtrl+Shift+5'),
          item('Date', 'format.numberDate', dispatch, 'CmdOrCtrl+Shift+3'),
        ],
      },
      { type: 'separator' },
      item('Cells…', 'format.cells', dispatch, 'CmdOrCtrl+1'),
    ],
  });

  template.push({
    label: '&Data',
    submenu: [
      item('Sort A to Z', 'data.sortAscending', dispatch),
      item('Sort Z to A', 'data.sortDescending', dispatch),
      item('Filter', 'data.filter', dispatch, 'CmdOrCtrl+Shift+L'),
      { type: 'separator' },
      item('Remove Duplicates…', 'data.removeDuplicates', dispatch),
      item('Text to Columns…', 'data.textToColumns', dispatch),
      { type: 'separator' },
      item('Recalculate', 'data.recalculate', dispatch, 'F9'),
    ],
  });

  template.push({
    role: 'help',
    label: '&Help',
    submenu: [
      item('Documentation', 'help.documentation', dispatch),
      item('Keyboard Shortcuts', 'help.shortcuts', dispatch, 'CmdOrCtrl+/'),
      ...(isMac ? [] : [{ type: 'separator' as const }, item('About', 'help.about', dispatch)]),
    ],
  });

  return template;
}
