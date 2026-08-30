import { describe, expect, it } from 'vitest';
import { fireEvent, screen } from '@testing-library/react';
import { CommandPalette } from '../src/renderer/components/CommandPalette.js';
import { enter, renderWith } from './render.js';
import type { AppController } from '../src/renderer/state/controller.js';

const open = (controller: AppController): void => controller.setPalette(true);

function type(value: string): void {
  fireEvent.change(screen.getByLabelText('Search commands, sheets and names'), {
    target: { value },
  });
}

describe('CommandPalette', () => {
  it('renders nothing while closed', () => {
    renderWith(() => <CommandPalette />);
    expect(screen.queryByRole('dialog')).toBeNull();
  });

  it('opens with every kind of item listed', () => {
    renderWith(() => <CommandPalette />, open);
    const kinds = [...document.querySelectorAll('.mz-palette-kind')].map((n) => n.textContent);
    expect(new Set(kinds)).toContain('command');
    expect(new Set(kinds)).toContain('sheet');
  });

  it('finds a sheet by a partial name', () => {
    renderWith(() => <CommandPalette />, (c) => {
      c.addSheet('Sensitivity Analysis');
      open(c);
    });
    type('sa');
    const first = document.querySelector('.mz-palette-item')!;
    expect(first.textContent).toContain('Sensitivity Analysis');
  });

  it('finds a command by name', () => {
    renderWith(() => <CommandPalette />, open);
    type('recalc');
    expect(document.querySelector('.mz-palette-item')!.textContent).toContain('Recalculate now');
  });

  it('finds a named range', () => {
    renderWith(() => <CommandPalette />, (c) => {
      c.workbook.definedNames = [{ name: 'TaxRate', refersTo: 'Sheet1!$B$2' }];
      open(c);
    });
    type('taxr');
    expect(document.querySelector('.mz-palette-item')!.textContent).toContain('TaxRate');
  });

  it('restricts to commands after a > prefix', () => {
    renderWith(() => <CommandPalette />, (c) => {
      c.addSheet('Undoubtedly');
      open(c);
    });
    type('>undo');
    const kinds = [...document.querySelectorAll('.mz-palette-kind')].map((n) => n.textContent);
    expect(new Set(kinds)).toEqual(new Set(['command']));
  });

  it('restricts to sheets after an @ prefix', () => {
    renderWith(() => <CommandPalette />, (c) => {
      c.addSheet('Data');
      open(c);
    });
    type('@');
    const kinds = [...document.querySelectorAll('.mz-palette-kind')].map((n) => n.textContent);
    expect(new Set(kinds)).toEqual(new Set(['sheet']));
  });

  it('says so when nothing matches', () => {
    renderWith(() => <CommandPalette />, open);
    type('zzzzzz');
    expect(screen.getByText(/Nothing matches/)).toBeDefined();
  });

  it('moves the selection with the arrow keys', () => {
    renderWith(() => <CommandPalette />, open);
    const input = screen.getByLabelText('Search commands, sheets and names');
    fireEvent.keyDown(input, { key: 'ArrowDown' });
    const selected = document.querySelector('.mz-palette-item.mz-selected')!;
    expect(selected).toBe(document.querySelectorAll('.mz-palette-item')[1]);
  });

  it('wraps the selection past the top', () => {
    renderWith(() => <CommandPalette />, open);
    const input = screen.getByLabelText('Search commands, sheets and names');
    fireEvent.keyDown(input, { key: 'ArrowUp' });
    const items = document.querySelectorAll('.mz-palette-item');
    expect(document.querySelector('.mz-palette-item.mz-selected')).toBe(items[items.length - 1]);
  });

  it('runs a command on Enter and closes', () => {
    const { controller } = renderWith(() => <CommandPalette />, (c) => {
      enter(c, 0, 0, '42');
      open(c);
    });
    type('undo');
    fireEvent.keyDown(screen.getByLabelText('Search commands, sheets and names'), { key: 'Enter' });
    expect(controller.getSnapshot().paletteOpen).toBe(false);
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })).toBeUndefined();
  });

  it('switches sheets when a sheet is chosen', () => {
    const { controller } = renderWith(() => <CommandPalette />, (c) => {
      c.addSheet('Data');
      c.setActiveSheet('Sheet1');
      open(c);
    });
    type('@data');
    fireEvent.mouseDown(document.querySelector('.mz-palette-item')!);
    expect(controller.getSnapshot().activeSheet).toBe('Data');
  });

  it('closes on Escape without running anything', () => {
    const { controller } = renderWith(() => <CommandPalette />, (c) => {
      enter(c, 0, 0, '42');
      open(c);
    });
    fireEvent.keyDown(screen.getByLabelText('Search commands, sheets and names'), { key: 'Escape' });
    expect(controller.getSnapshot().paletteOpen).toBe(false);
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe(42);
  });

  it('greys out a command that cannot run yet', () => {
    renderWith(() => <CommandPalette />, open);
    type('undo');
    const item = document.querySelector('.mz-palette-item')!;
    expect(item.getAttribute('aria-disabled')).toBe('true');
  });

  it('shows the keyboard shortcut for a command that has one', () => {
    renderWith(() => <CommandPalette />, open);
    type('recalculate now');
    expect(screen.getByText('F9')).toBeDefined();
  });
});
