/**
 * The menu is built as data, so it can be inspected without a display.
 *
 * What is worth asserting is not the shape of the tree but the decisions
 * embedded in it: that Undo goes through the document's command log rather than
 * the text field's, that clipboard items stay as native roles, and that the
 * developer tools do not appear in a shipped build.
 */

import { describe, expect, it, vi } from 'vitest';
import type { MenuItemConstructorOptions } from 'electron';
import { buildMenuTemplate, type MenuCommand } from '../../src/main/menu.js';

type Item = MenuItemConstructorOptions;

function build(overrides: Partial<Parameters<typeof buildMenuTemplate>[0]> = {}): {
  template: Item[];
  dispatch: ReturnType<typeof vi.fn>;
} {
  const dispatch = vi.fn();
  const template = buildMenuTemplate({
    platform: 'linux',
    appName: 'MIRRORZ Sheets',
    recentPaths: [],
    dispatch: dispatch as unknown as (command: MenuCommand, arg?: string) => void,
    ...overrides,
  });
  return { template, dispatch };
}

function menu(template: Item[], label: string): Item[] {
  const found = template.find((item) => item.label?.replace('&', '') === label);
  return (found?.submenu as Item[] | undefined) ?? [];
}

function find(items: Item[], label: string): Item | undefined {
  for (const item of items) {
    if (item.label === label) return item;
    if (Array.isArray(item.submenu)) {
      const nested = find(item.submenu as Item[], label);
      if (nested) return nested;
    }
  }
  return undefined;
}

function click(items: Item[], label: string): void {
  const item = find(items, label);
  // The Electron signature carries a MenuItem and an event neither this menu
  // nor its handlers look at, so calling with none is faithful enough.
  (item?.click as undefined | (() => void))?.();
}

describe('menu structure', () => {
  it('has the seven top-level menus users expect from a spreadsheet', () => {
    const { template } = build();
    expect(template.map((item) => item.label?.replace('&', ''))).toEqual([
      'File',
      'Edit',
      'View',
      'Insert',
      'Format',
      'Data',
      'Help',
    ]);
  });

  it('adds the application menu on macOS, where the platform requires one', () => {
    const { template } = build({ platform: 'darwin' });
    expect(template[0]?.label).toBe('MIRRORZ Sheets');
    expect(template.map((item) => item.label?.replace('&', ''))).toContain('File');
  });

  it('puts Exit in the File menu everywhere except macOS', () => {
    expect(find(menu(build().template, 'File'), 'Exit')).toBeDefined();
    expect(find(menu(build({ platform: 'darwin' }).template, 'File'), 'Exit')).toBeUndefined();
  });
});

describe('menu commands', () => {
  it('dispatches the file commands', () => {
    const { template, dispatch } = build();
    const file = menu(template, 'File');
    click(file, 'New');
    click(file, 'Open…');
    click(file, 'Save');
    click(file, 'Save As…');
    expect(dispatch.mock.calls.map((c) => c[0])).toEqual([
      'file.new',
      'file.open',
      'file.save',
      'file.saveAs',
    ]);
  });

  it('routes Undo and Redo through the document, not the focused text field', () => {
    const { template, dispatch } = build();
    const edit = menu(template, 'Edit');
    expect(find(edit, 'Undo')?.role).toBeUndefined();
    expect(find(edit, 'Redo')?.role).toBeUndefined();
    click(edit, 'Undo');
    click(edit, 'Redo');
    expect(dispatch.mock.calls.map((c) => c[0])).toEqual(['edit.undo', 'edit.redo']);
  });

  it('keeps clipboard items as native roles, which the sandbox needs', () => {
    const edit = menu(build().template, 'Edit');
    const roles = edit.map((item) => item.role).filter(Boolean);
    expect(roles).toContain('cut');
    expect(roles).toContain('copy');
    expect(roles).toContain('paste');
  });

  it('binds the accelerators people arrive with', () => {
    const { template } = build();
    expect(find(menu(template, 'File'), 'Save')?.accelerator).toBe('CmdOrCtrl+S');
    expect(find(menu(template, 'File'), 'Open…')?.accelerator).toBe('CmdOrCtrl+O');
    expect(find(menu(template, 'Edit'), 'Undo')?.accelerator).toBe('CmdOrCtrl+Z');
    expect(find(menu(template, 'Data'), 'Recalculate')?.accelerator).toBe('F9');
    expect(find(menu(template, 'Format'), 'Cells…')?.accelerator).toBe('CmdOrCtrl+1');
  });

  it('uses each platform Redo shortcut', () => {
    expect(find(menu(build().template, 'Edit'), 'Redo')?.accelerator).toBe('Ctrl+Y');
    expect(find(menu(build({ platform: 'darwin' }).template, 'Edit'), 'Redo')?.accelerator).toBe(
      'Cmd+Shift+Z',
    );
  });

  it('offers the format, data and insert commands', () => {
    const { template, dispatch } = build();
    click(menu(template, 'Insert'), 'Rows Above');
    click(menu(template, 'Format'), 'Bold');
    click(menu(template, 'Format'), 'Percent');
    click(menu(template, 'Data'), 'Sort A to Z');
    click(menu(template, 'Data'), 'Recalculate');
    expect(dispatch.mock.calls.map((c) => c[0])).toEqual([
      'insert.rowsAbove',
      'format.bold',
      'format.numberPercent',
      'data.sortAscending',
      'data.recalculate',
    ]);
  });
});

describe('recent files submenu', () => {
  it('says so when there is nothing, rather than showing an empty menu', () => {
    const file = menu(build().template, 'File');
    const recent = find(file, 'Open Recent')?.submenu as Item[];
    expect(recent).toHaveLength(1);
    expect(recent[0]?.label).toBe('No recent files');
    expect(recent[0]?.enabled).toBe(false);
  });

  it('lists files by name and dispatches the full path', () => {
    const { template, dispatch } = build({
      recentPaths: ['/home/user/Budget.xlsx', '/home/user/Q3.csv'],
    });
    const recent = find(menu(template, 'File'), 'Open Recent')?.submenu as Item[];
    expect(recent.map((item) => item.label)).toEqual([
      'Budget.xlsx',
      'Q3.csv',
      undefined,
      'Clear Recent Files',
    ]);
    (recent[0]?.click as undefined | (() => void))?.();
    expect(dispatch).toHaveBeenCalledWith('file.openRecent', '/home/user/Budget.xlsx');
  });

  it('carries the full path in the tooltip, since the label is shortened', () => {
    const { template } = build({ recentPaths: ['/home/user/deep/Budget.xlsx'] });
    const recent = find(menu(template, 'File'), 'Open Recent')?.submenu as Item[];
    expect(recent[0]?.toolTip).toBe('/home/user/deep/Budget.xlsx');
  });

  it('offers a way to clear the list', () => {
    const { template, dispatch } = build({ recentPaths: ['/home/user/a.xlsx'] });
    click(menu(template, 'File'), 'Clear Recent Files');
    expect(dispatch).toHaveBeenCalledWith('file.clearRecent');
  });
});

describe('developer tools', () => {
  it('are absent from a shipped build', () => {
    const view = menu(build().template, 'View');
    expect(view.some((item) => item.role === 'toggleDevTools')).toBe(false);
  });

  it('are present when running from source', () => {
    const view = menu(build({ isDev: true }).template, 'View');
    expect(view.some((item) => item.role === 'toggleDevTools')).toBe(true);
  });
});
