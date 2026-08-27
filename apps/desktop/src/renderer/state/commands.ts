/**
 * The command catalogue.
 *
 * Every action the application can perform is described here once, as data, and
 * the toolbar, the keyboard map and the command palette all read from this list.
 * The alternative - a button that calls a method and a palette entry that calls
 * the same method - is how the two drift apart until half the features are
 * reachable from only one of them.
 */

import type { AppController, AppSnapshot } from './controller.js';
import type { PaletteItem } from '../model/palette.js';

export interface AppCommand {
  id: string;
  title: string;
  category: string;
  keywords?: string;
  shortcut?: string;
  /** False greys the entry out; it still appears, so it is still discoverable. */
  enabled?: boolean;
  run: () => void;
}

export function buildCommands(controller: AppController, snapshot: AppSnapshot): AppCommand[] {
  const commands: AppCommand[] = [
    {
      id: 'edit.undo',
      title: snapshot.undoLabel ? `Undo ${snapshot.undoLabel}` : 'Undo',
      category: 'Edit',
      shortcut: 'Ctrl+Z',
      keywords: 'revert back',
      enabled: snapshot.canUndo,
      run: () => controller.undo(),
    },
    {
      id: 'edit.redo',
      title: snapshot.redoLabel ? `Redo ${snapshot.redoLabel}` : 'Redo',
      category: 'Edit',
      shortcut: 'Ctrl+Y',
      enabled: snapshot.canRedo,
      run: () => controller.redo(),
    },
    {
      id: 'edit.clear',
      title: 'Clear cells',
      category: 'Edit',
      shortcut: 'Delete',
      keywords: 'delete empty',
      run: () => controller.clearSelection(),
    },
    {
      id: 'view.history',
      title: snapshot.panels.history ? 'Hide history' : 'Show history',
      category: 'View',
      keywords: 'undo tree timeline branches',
      shortcut: 'Ctrl+Shift+H',
      run: () => controller.togglePanel('history'),
    },
    {
      id: 'view.inspector',
      title: snapshot.panels.inspector ? 'Hide formula inspector' : 'Show formula inspector',
      category: 'View',
      keywords: 'precedents dependents trace error audit',
      shortcut: 'Ctrl+Shift+I',
      run: () => controller.togglePanel('inspector'),
    },
    {
      id: 'view.explorer',
      title: snapshot.panels.explorer ? 'Hide sheet explorer' : 'Show sheet explorer',
      category: 'View',
      keywords: 'sidebar sheets list',
      shortcut: 'Ctrl+Shift+E',
      run: () => controller.togglePanel('explorer'),
    },
    {
      id: 'view.theme.light',
      title: 'Use light theme',
      category: 'View',
      keywords: 'appearance colour',
      run: () => controller.setTheme('light'),
    },
    {
      id: 'view.theme.dark',
      title: 'Use dark theme',
      category: 'View',
      keywords: 'appearance colour night',
      run: () => controller.setTheme('dark'),
    },
    {
      id: 'view.theme.system',
      title: 'Follow the system theme',
      category: 'View',
      keywords: 'appearance auto os',
      run: () => controller.setTheme('system'),
    },
    {
      id: 'sheet.add',
      title: 'Add sheet',
      category: 'Sheet',
      keywords: 'new tab worksheet',
      run: () => controller.addSheet(),
    },
    {
      id: 'sheet.delete',
      title: `Delete sheet ${snapshot.activeSheet}`,
      category: 'Sheet',
      keywords: 'remove',
      run: () => controller.removeSheet(snapshot.activeSheet),
    },
    {
      id: 'sheet.hide',
      title: `Hide sheet ${snapshot.activeSheet}`,
      category: 'Sheet',
      run: () => controller.setSheetVisibility(snapshot.activeSheet, 'hidden'),
    },
    {
      id: 'rows.insert',
      title: 'Insert rows',
      category: 'Structure',
      keywords: 'add row above',
      run: () => controller.insertRows(controller.selectionRange().start.row, rowSpan(controller)),
    },
    {
      id: 'rows.delete',
      title: 'Delete rows',
      category: 'Structure',
      keywords: 'remove row',
      run: () => controller.deleteRows(controller.selectionRange().start.row, rowSpan(controller)),
    },
    {
      id: 'cols.insert',
      title: 'Insert columns',
      category: 'Structure',
      keywords: 'add column left',
      run: () => controller.insertCols(controller.selectionRange().start.col, colSpan(controller)),
    },
    {
      id: 'cols.delete',
      title: 'Delete columns',
      category: 'Structure',
      keywords: 'remove column',
      run: () => controller.deleteCols(controller.selectionRange().start.col, colSpan(controller)),
    },
    {
      id: 'data.sort.asc',
      title: 'Sort ascending',
      category: 'Data',
      keywords: 'order a to z',
      run: () => controller.sortSelection(controller.selectionRange().start.col, 'asc'),
    },
    {
      id: 'data.sort.desc',
      title: 'Sort descending',
      category: 'Data',
      keywords: 'order z to a',
      run: () => controller.sortSelection(controller.selectionRange().start.col, 'desc'),
    },
    {
      id: 'data.filter.clear',
      title: 'Clear filter',
      category: 'Data',
      keywords: 'unhide rows show all',
      run: () => controller.filterRows(controller.selectionRange().start.col, ''),
    },
    {
      id: 'calc.now',
      title: 'Recalculate now',
      category: 'Formulas',
      shortcut: 'F9',
      keywords: 'refresh compute',
      run: () => controller.recalculateAll(),
    },
    {
      id: 'calc.manual',
      title:
        snapshot.calcMode === 'manual' ? 'Switch to automatic calculation' : 'Switch to manual calculation',
      category: 'Formulas',
      run: () => controller.setCalcMode(snapshot.calcMode === 'manual' ? 'auto' : 'manual'),
    },
  ];
  return commands;
}

function rowSpan(controller: AppController): number {
  const range = controller.selectionRange();
  return range.end.row - range.start.row + 1;
}

function colSpan(controller: AppController): number {
  const range = controller.selectionRange();
  return range.end.col - range.start.col + 1;
}

/**
 * Everything the palette can find: commands, sheets and defined names, in one
 * list so a single query searches all three.
 */
export function buildPaletteItems(
  controller: AppController,
  snapshot: AppSnapshot,
  commands = buildCommands(controller, snapshot),
): PaletteItem[] {
  const items: PaletteItem[] = commands.map((command) => {
    const item: PaletteItem = {
      id: command.id,
      kind: 'command',
      title: command.title,
      detail: command.category,
    };
    if (command.keywords !== undefined) item.keywords = command.keywords;
    if (command.shortcut !== undefined) item.shortcut = command.shortcut;
    if (command.enabled !== undefined) item.enabled = command.enabled;
    return item;
  });

  for (const sheet of controller.sheetSummaries()) {
    items.push({
      id: `sheet:${sheet.name}`,
      kind: 'sheet',
      title: sheet.name,
      detail:
        sheet.visibility === 'visible'
          ? `Sheet ${sheet.index + 1}`
          : `Sheet ${sheet.index + 1}, hidden`,
      keywords: 'worksheet tab',
    });
  }

  for (const name of controller.workbook.definedNames) {
    if (name.hidden) continue;
    items.push({
      id: `name:${name.name}`,
      kind: 'name',
      title: name.name,
      detail: name.refersTo,
      keywords: 'named range defined name',
    });
  }

  return items;
}

/** Run whatever a palette item stands for. */
export function runPaletteItem(
  controller: AppController,
  item: PaletteItem,
  commands: readonly AppCommand[],
): void {
  if (item.kind === 'command') {
    commands.find((c) => c.id === item.id)?.run();
    return;
  }
  if (item.kind === 'sheet') {
    controller.setActiveSheet(item.title);
    return;
  }
  // A defined name jumps to the sheet its target names; resolving the full
  // reference is the formula package's job and is not needed to navigate.
  const sheetName = item.detail?.split('!')[0]?.replace(/^'|'$/g, '');
  if (sheetName && controller.workbook.getSheet(sheetName)) controller.setActiveSheet(sheetName);
}
