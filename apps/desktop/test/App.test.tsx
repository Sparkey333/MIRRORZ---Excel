import { describe, expect, it } from 'vitest';
import { act, fireEvent, screen, waitFor } from '@testing-library/react';
import { App } from '../src/renderer/App.js';
import { FormulaBar } from '../src/renderer/components/FormulaBar.js';
import { enter, renderWith } from './render.js';

describe('FormulaBar', () => {
  it('shows the active cell reference in the name box', () => {
    renderWith(() => <FormulaBar />, (c) => c.selectCell(2, 3));
    expect(screen.getByText('D3')).toBeDefined();
  });

  it('shows a formula with its leading equals', () => {
    renderWith(() => <FormulaBar />, (c) => {
      enter(c, 0, 0, '=1+1');
      c.selectCell(0, 0);
    });
    expect((screen.getByLabelText('Formula') as HTMLTextAreaElement).value).toBe('=1+1');
  });

  it('shows the text the user typed, not the value inferred from it', () => {
    renderWith(() => <FormulaBar />, (c) => {
      enter(c, 0, 0, '007');
      c.selectCell(0, 0);
    });
    expect((screen.getByLabelText('Formula') as HTMLTextAreaElement).value).toBe('007');
  });

  it('commits an edit to the document on Enter', () => {
    const { controller } = renderWith(() => <FormulaBar />);
    const input = screen.getByLabelText('Formula');
    fireEvent.change(input, { target: { value: '=2*21' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe(42);
  });

  it('does not touch the document while the draft is being typed', () => {
    const { controller } = renderWith(() => <FormulaBar />);
    fireEvent.change(screen.getByLabelText('Formula'), { target: { value: '=2*2' } });
    expect(controller.historyEntries()).toHaveLength(0);
  });

  it('abandons the draft when the selection moves', () => {
    const { controller } = renderWith(() => <FormulaBar />);
    fireEvent.change(screen.getByLabelText('Formula'), { target: { value: 'draft' } });
    act(() => controller.selectCell(1, 1));
    expect((screen.getByLabelText('Formula') as HTMLTextAreaElement).value).toBe('');
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })).toBeUndefined();
  });

  it('commits from the tick button', () => {
    const { controller } = renderWith(() => <FormulaBar />);
    fireEvent.change(screen.getByLabelText('Formula'), { target: { value: '5' } });
    fireEvent.click(screen.getByLabelText('Commit edit'));
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe(5);
  });

  it('reverts from the cross button', () => {
    const { controller } = renderWith(() => <FormulaBar />);
    fireEvent.change(screen.getByLabelText('Formula'), { target: { value: '5' } });
    fireEvent.click(screen.getByLabelText('Cancel edit'));
    expect((screen.getByLabelText('Formula') as HTMLTextAreaElement).value).toBe('');
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })).toBeUndefined();
  });

  it('completes from the engine registry, so it can only offer real functions', () => {
    const { controller } = renderWith(() => <FormulaBar />);
    const input = screen.getByLabelText('Formula') as HTMLTextAreaElement;
    fireEvent.change(input, { target: { value: '=SU' } });
    input.setSelectionRange(3, 3);
    fireEvent.keyUp(input, { key: 'U' });
    const offered = screen
      .queryAllByRole('option')
      .map((o) => o.querySelector('.mz-completion-name')?.textContent?.replace(/\s/g, ''));
    expect(offered.length).toBeGreaterThan(0);
    // Every suggestion is a name the engine will actually resolve, which is the
    // point of drawing the list from the registry rather than a hand-kept list.
    const known = new Set(controller.registry.names());
    expect(offered.every((name) => name !== undefined && known.has(name))).toBe(true);
  });
});

describe('App', () => {
  it('renders the whole shell', async () => {
    renderWith(() => <App />);
    expect(screen.getByRole('toolbar', { name: 'Main toolbar' })).toBeDefined();
    expect(screen.getByLabelText('Formula')).toBeDefined();
    expect(screen.getByRole('tablist', { name: 'Sheets' })).toBeDefined();
    expect(screen.getByRole('status', { name: 'Status bar' })).toBeDefined();
    expect(screen.getByLabelText('Sheet explorer')).toBeDefined();
  });

  it('falls back to a DOM grid when the canvas cannot paint', async () => {
    renderWith(() => <App />);
    await waitFor(() => expect(screen.getByLabelText('Sheet1 cells')).toBeDefined());
  });

  it('applies the theme as custom properties on the root', () => {
    renderWith(() => <App />);
    expect(document.documentElement.style.getPropertyValue('--mz-grid-line')).not.toBe('');
    expect(document.documentElement.dataset['mzTheme']).toBe('light');
  });

  it('switches the palette on Ctrl+P', () => {
    const { controller } = renderWith(() => <App />);
    act(() => {
      fireEvent.keyDown(window, { key: 'p', ctrlKey: true });
    });
    expect(controller.getSnapshot().paletteOpen).toBe(true);
    expect(screen.getByRole('dialog', { name: 'Command palette' })).toBeDefined();
  });

  it('undoes on Ctrl+Z', () => {
    const { controller } = renderWith(() => <App />, (c) => enter(c, 0, 0, '42'));
    act(() => {
      fireEvent.keyDown(window, { key: 'z', ctrlKey: true });
    });
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })).toBeUndefined();
  });

  it('does not steal a shortcut while the user is typing in a field', () => {
    const { controller } = renderWith(() => <App />, (c) => enter(c, 0, 0, '42'));
    act(() => {
      fireEvent.keyDown(screen.getByLabelText('Formula'), { key: 'z', ctrlKey: true });
    });
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe(42);
  });

  it('moves the selection with the arrow keys', () => {
    const { controller } = renderWith(() => <App />);
    act(() => {
      fireEvent.keyDown(window, { key: 'ArrowDown' });
      fireEvent.keyDown(window, { key: 'ArrowRight' });
    });
    expect(controller.getSnapshot().selection.active).toEqual({ row: 1, col: 1 });
  });

  it('opens the history panel from its shortcut', () => {
    const { controller } = renderWith(() => <App />);
    act(() => {
      fireEvent.keyDown(window, { key: 'H', ctrlKey: true, shiftKey: true });
    });
    expect(controller.getSnapshot().panels.history).toBe(true);
    expect(screen.getByLabelText('History')).toBeDefined();
  });

  it('selects a cell in the fallback grid', async () => {
    const { controller } = renderWith(() => <App />);
    await waitFor(() => expect(screen.getByLabelText('Sheet1 cells')).toBeDefined());
    fireEvent.mouseDown(document.querySelector('[data-address="C4"]')!);
    expect(controller.getSnapshot().selection.active).toEqual({ row: 3, col: 2 });
  });

  it('shows a formatted value in the grid', async () => {
    renderWith(() => <App />, (c) => {
      enter(c, 0, 0, '1234.5');
      c.selectCell(0, 0);
      c.setNumberFormat('#,##0.00');
    });
    await waitFor(() => expect(screen.getByLabelText('Sheet1 cells')).toBeDefined());
    expect(document.querySelector('[data-address="A1"]')!.textContent).toBe('1,234.50');
  });
});
