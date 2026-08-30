import { describe, expect, it } from 'vitest';
import { fireEvent, screen, within } from '@testing-library/react';
import { Toolbar } from '../src/renderer/components/Toolbar.js';
import { enter, renderWith } from './render.js';
import type { AppController } from '../src/renderer/state/controller.js';

const openFormatMenu = (): void => {
  const group = screen.getByRole('group', { name: 'Number' });
  fireEvent.click(within(group).getByRole('button', { expanded: false }));
};

const withCell = (c: AppController): void => {
  enter(c, 0, 0, '1234.567');
  c.selectCell(0, 0);
};

describe('Toolbar', () => {
  it('groups its controls and labels each group', () => {
    renderWith(() => <Toolbar />);
    for (const label of ['File', 'Clipboard', 'Font', 'Alignment', 'Number', 'Cells', 'Data', 'View']) {
      expect(screen.getByRole('group', { name: label })).toBeDefined();
    }
  });

  it('disables the file buttons when no host supplied them', () => {
    renderWith(() => <Toolbar />);
    expect((screen.getByText('Open') as HTMLButtonElement).disabled).toBe(true);
  });

  it('applies bold through the document, so it can be undone', () => {
    const { controller } = renderWith(() => <Toolbar />, withCell);
    fireEvent.click(screen.getByLabelText('Bold'));
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).font?.bold).toBe(true);
    controller.undo();
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).font?.bold).toBeUndefined();
  });

  it('shows bold as pressed once the selection is bold', () => {
    renderWith(() => <Toolbar />, withCell);
    fireEvent.click(screen.getByLabelText('Bold'));
    expect(screen.getByLabelText('Bold').getAttribute('aria-pressed')).toBe('true');
  });

  it('toggles bold off again', () => {
    const { controller } = renderWith(() => <Toolbar />, withCell);
    fireEvent.click(screen.getByLabelText('Bold'));
    fireEvent.click(screen.getByLabelText('Bold'));
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).font?.bold).toBe(false);
  });

  it('changes the font family without discarding the rest of the format', () => {
    const { controller } = renderWith(() => <Toolbar />, withCell);
    fireEvent.click(screen.getByLabelText('Bold'));
    fireEvent.change(screen.getByLabelText('Font family'), { target: { value: 'Georgia' } });
    const font = controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).font!;
    expect(font.name).toBe('Georgia');
    expect(font.bold).toBe(true);
  });

  it('sets a fill colour', () => {
    const { controller } = renderWith(() => <Toolbar />, withCell);
    fireEvent.click(screen.getByLabelText('Fill colour'));
    fireEvent.click(screen.getByLabelText('Fill colour #fff3bf'));
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).fill?.fg).toEqual({
      kind: 'rgb',
      argb: 'FFFFF3BF',
    });
  });

  it('sets an alignment', () => {
    const { controller } = renderWith(() => <Toolbar />, withCell);
    fireEvent.click(screen.getByLabelText('Align center'));
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).alignment?.horizontal).toBe('center');
  });

  it('opens the format menu showing a live sample for each preset', () => {
    renderWith(() => <Toolbar />, withCell);
    openFormatMenu();
    expect(screen.getByText('1,234.57')).toBeDefined();
    expect(screen.getByText('12.34%')).toBeDefined();
  });

  it('applies a preset format and shows it on the button', () => {
    const { controller } = renderWith(() => <Toolbar />, withCell);
    openFormatMenu();
    fireEvent.click(screen.getByText('Thousands separator'));
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).numFmt).toBe('#,##0.00');
    expect(screen.getByText('#,##0.00')).toBeDefined();
  });

  it('marks the current format as checked in the menu', () => {
    renderWith(() => <Toolbar />, withCell);
    openFormatMenu();
    fireEvent.click(screen.getByText('Percent'));
    openFormatMenu();
    const item = screen.getByText('Percent').closest('button')!;
    expect(item.getAttribute('aria-checked')).toBe('true');
  });

  it('previews a valid custom format code', () => {
    renderWith(() => <Toolbar />, withCell);
    openFormatMenu();
    fireEvent.change(screen.getByLabelText('Custom format code'), { target: { value: '0.0' } });
    expect(screen.getByText('Sample: 1234.6')).toBeDefined();
  });

  it('explains an invalid custom format code instead of applying it', () => {
    renderWith(() => <Toolbar />, withCell);
    openFormatMenu();
    fireEvent.change(screen.getByLabelText('Custom format code'), { target: { value: '0;0;0;0;0' } });
    expect(screen.getByText(/at most four sections/)).toBeDefined();
  });

  it('applies a custom format on Enter', () => {
    const { controller } = renderWith(() => <Toolbar />, withCell);
    openFormatMenu();
    const input = screen.getByLabelText('Custom format code');
    fireEvent.change(input, { target: { value: '0.000' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).numFmt).toBe('0.000');
  });

  it('does not apply an invalid custom format on Enter', () => {
    const { controller } = renderWith(() => <Toolbar />, withCell);
    openFormatMenu();
    const input = screen.getByLabelText('Custom format code');
    fireEvent.change(input, { target: { value: '0;0;0;0;0' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).numFmt).toBeUndefined();
  });

  it('steps the decimal places of the existing code', () => {
    const { controller } = renderWith(() => <Toolbar />, withCell);
    openFormatMenu();
    fireEvent.click(screen.getByText('Thousands separator'));
    fireEvent.click(screen.getByLabelText('Add decimal place'));
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).numFmt).toBe('#,##0.000');
    fireEvent.click(screen.getByLabelText('Remove decimal place'));
    expect(controller.styleOf({ sheet: 'Sheet1', row: 0, col: 0 }).numFmt).toBe('#,##0.00');
  });

  it('sorts the selection', () => {
    const { controller } = renderWith(() => <Toolbar />, (c) => {
      enter(c, 0, 0, '3');
      enter(c, 1, 0, '1');
      c.setSelection({
        sheet: 'Sheet1',
        active: { row: 0, col: 0 },
        ranges: [{ start: { row: 0, col: 0 }, end: { row: 1, col: 0 } }],
      });
    });
    fireEvent.click(screen.getByLabelText('Sort ascending'));
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe(1);
  });

  it('opens find and replace', () => {
    renderWith(() => <Toolbar />);
    fireEvent.click(screen.getByText('Find'));
    expect(screen.getByRole('dialog', { name: 'Find and replace' })).toBeDefined();
  });

  it('toggles the history panel', () => {
    const { controller } = renderWith(() => <Toolbar />);
    fireEvent.click(screen.getByText('History'));
    expect(controller.getSnapshot().panels.history).toBe(true);
  });

  it('cycles the theme preference', () => {
    const { controller } = renderWith(() => <Toolbar />);
    fireEvent.click(screen.getByLabelText('Theme: system'));
    expect(controller.getSnapshot().theme).toBe('light');
    fireEvent.click(screen.getByLabelText('Theme: light'));
    expect(controller.getSnapshot().theme).toBe('dark');
  });

  it('inserts a row through the document', () => {
    const { controller } = renderWith(() => <Toolbar />, (c) => {
      enter(c, 0, 0, 'first');
      c.selectCell(0, 0);
    });
    fireEvent.click(screen.getByText('Insert row'));
    expect(controller.cellAt({ sheet: 'Sheet1', row: 1, col: 0 })?.value).toBe('first');
    controller.undo();
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe('first');
  });
});
