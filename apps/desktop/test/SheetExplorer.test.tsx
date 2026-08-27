import { describe, expect, it } from 'vitest';
import { fireEvent, screen } from '@testing-library/react';
import { SheetExplorer } from '../src/renderer/components/SheetExplorer.js';
import { SheetTabs } from '../src/renderer/components/SheetTabs.js';
import { renderWith } from './render.js';
import type { AppController } from '../src/renderer/state/controller.js';

function manySheets(controller: AppController): void {
  for (const name of ['Summary', 'Q1 Sales', 'Q2 Sales', 'Assumptions', 'Sensitivity Analysis']) {
    controller.addSheet(name);
  }
  controller.setActiveSheet('Sheet1');
}

const search = (value: string): void => {
  fireEvent.change(screen.getByLabelText('Search sheets'), { target: { value } });
};

const names = (): string[] =>
  [...document.querySelectorAll('.mz-explorer-name')].map((n) => n.textContent ?? '');

describe('SheetExplorer', () => {
  it('lists every sheet', () => {
    renderWith(() => <SheetExplorer />, manySheets);
    expect(names()).toHaveLength(6);
  });

  it('filters fuzzily as the user types', () => {
    renderWith(() => <SheetExplorer />, manySheets);
    search('sa');
    expect(names()[0]).toBe('Sensitivity Analysis');
  });

  it('highlights the matched characters', () => {
    renderWith(() => <SheetExplorer />, manySheets);
    search('sum');
    expect(document.querySelector('.mz-explorer-name .mz-match')).not.toBeNull();
  });

  it('says so when nothing matches', () => {
    renderWith(() => <SheetExplorer />, manySheets);
    search('zzzz');
    expect(screen.getByText(/No sheets match/)).toBeDefined();
  });

  it('switches sheets on click', () => {
    const { controller } = renderWith(() => <SheetExplorer />, manySheets);
    fireEvent.click(screen.getByText('Assumptions'));
    expect(controller.getSnapshot().activeSheet).toBe('Assumptions');
  });

  it('lists hidden sheets, which the tab strip cannot', () => {
    const { controller } = renderWith(() => <SheetExplorer />, (c) => {
      manySheets(c);
      c.setSheetVisibility('Q1 Sales', 'hidden');
    });
    expect(names()).toContain('Q1 Sales');
    expect(controller.workbook.getSheet('Q1 Sales')!.visibility).toBe('hidden');
  });

  it('can filter the hidden sheets back out', () => {
    renderWith(() => <SheetExplorer />, (c) => {
      manySheets(c);
      c.setSheetVisibility('Q1 Sales', 'hidden');
    });
    fireEvent.click(screen.getByLabelText('Show hidden sheets', { selector: 'input' }));
    expect(names()).not.toContain('Q1 Sales');
  });

  it('shows a hidden sheet again', () => {
    const { controller } = renderWith(() => <SheetExplorer />, (c) => {
      manySheets(c);
      c.setSheetVisibility('Q1 Sales', 'hidden');
    });
    fireEvent.click(screen.getByLabelText('Show Q1 Sales'));
    expect(controller.workbook.getSheet('Q1 Sales')!.visibility).toBe('visible');
  });

  it('reorders a sheet', () => {
    const { controller } = renderWith(() => <SheetExplorer />, manySheets);
    fireEvent.click(screen.getByLabelText('Move Summary up'));
    expect(controller.workbook.sheets[0]!.name).toBe('Summary');
  });

  it('renames a sheet on double click and Enter', () => {
    const { controller } = renderWith(() => <SheetExplorer />, manySheets);
    fireEvent.doubleClick(screen.getByText('Summary'));
    const input = screen.getByLabelText('Rename Summary');
    fireEvent.change(input, { target: { value: 'Overview' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    expect(controller.workbook.getSheet('Overview')).toBeDefined();
  });

  it('abandons a rename on Escape', () => {
    const { controller } = renderWith(() => <SheetExplorer />, manySheets);
    fireEvent.doubleClick(screen.getByText('Summary'));
    const input = screen.getByLabelText('Rename Summary');
    fireEvent.change(input, { target: { value: 'Overview' } });
    fireEvent.keyDown(input, { key: 'Escape' });
    expect(controller.workbook.getSheet('Summary')).toBeDefined();
  });

  it('colours a sheet', () => {
    const { controller } = renderWith(() => <SheetExplorer />, manySheets);
    fireEvent.click(screen.getByLabelText('Colour Summary #e03131'));
    expect(controller.workbook.getSheet('Summary')!.tabColor).toBe('#e03131');
  });

  it('adds a sheet from its footer', () => {
    const { controller } = renderWith(() => <SheetExplorer />, manySheets);
    fireEvent.click(screen.getByText('Add sheet'));
    expect(controller.workbook.sheets).toHaveLength(7);
  });
});

describe('SheetTabs', () => {
  it('shows only the visible sheets, and counts the hidden ones', () => {
    renderWith(() => <SheetTabs />, (c) => {
      manySheets(c);
      c.setSheetVisibility('Q1 Sales', 'hidden');
    });
    expect(screen.getByText('1 hidden')).toBeDefined();
    expect(screen.queryByText('Q1 Sales')).toBeNull();
  });

  it('marks the active tab', () => {
    renderWith(() => <SheetTabs />, manySheets);
    const tab = screen.getByText('Sheet1').closest('[role="tab"]')!;
    expect(tab.getAttribute('aria-selected')).toBe('true');
  });

  it('switches sheets on click', () => {
    const { controller } = renderWith(() => <SheetTabs />, manySheets);
    fireEvent.click(screen.getByText('Summary'));
    expect(controller.getSnapshot().activeSheet).toBe('Summary');
  });

  it('renames on double click', () => {
    const { controller } = renderWith(() => <SheetTabs />, manySheets);
    fireEvent.doubleClick(screen.getByText('Summary'));
    const input = screen.getByLabelText('Rename Summary');
    fireEvent.change(input, { target: { value: 'Overview' } });
    fireEvent.keyDown(input, { key: 'Enter' });
    expect(controller.workbook.getSheet('Overview')).toBeDefined();
  });

  it('adds a sheet', () => {
    const { controller } = renderWith(() => <SheetTabs />, manySheets);
    fireEvent.click(screen.getByLabelText('Add sheet'));
    expect(controller.workbook.sheets).toHaveLength(7);
  });

  it('deletes a sheet from its menu, undoably', () => {
    const { controller } = renderWith(() => <SheetTabs />, manySheets);
    fireEvent.click(screen.getByLabelText('Options for Summary'));
    fireEvent.click(screen.getByText('Delete'));
    expect(controller.workbook.getSheet('Summary')).toBeUndefined();
    controller.undo();
    expect(controller.workbook.getSheet('Summary')).toBeDefined();
  });
});
