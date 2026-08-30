import { describe, expect, it } from 'vitest';
import { fireEvent, screen } from '@testing-library/react';
import { ImportReviewDialog } from '../src/renderer/components/ImportReviewDialog.js';
import { renderWith } from './render.js';
import type { AppController } from '../src/renderer/state/controller.js';

const rows = [
  ['gene', 'code', 'when'],
  ['SEPT1', '007', '2024-01-31'],
  ['MARCH2', '0042', '2024-02-01'],
];

function stage(controller: AppController): void {
  controller.proposeImport({
    anchor: { sheet: 'Sheet1', row: 0, col: 0 },
    rows,
    overrides: ['auto', 'auto', 'auto'],
    headerRow: true,
    source: 'paste',
  });
}

describe('ImportReviewDialog', () => {
  it('renders nothing when no import is staged', () => {
    renderWith(() => <ImportReviewDialog />);
    expect(screen.queryByRole('dialog')).toBeNull();
  });

  it('states what would happen before anything is written', () => {
    renderWith(() => <ImportReviewDialog />, stage);
    const headline = document.querySelector('.mz-import-headline')!.textContent!;
    expect(headline).toMatch(/2 cells would be read as dates/);
    expect(headline).toMatch(/4 kept as text \(gene symbols, leading zeros\)/);
  });

  it('names the landing address', () => {
    renderWith(() => <ImportReviewDialog />, stage);
    expect(screen.getByText(/Landing at Sheet1!A1/)).toBeDefined();
  });

  it('uses the header row for column names', () => {
    renderWith(() => <ImportReviewDialog />, stage);
    expect(screen.getByLabelText('Type for column gene')).toBeDefined();
  });

  it('explains why each protected cell was kept as text', () => {
    renderWith(() => <ImportReviewDialog />, stage);
    expect(screen.getAllByText(/looks like a gene symbol/)).toHaveLength(2);
    expect(screen.getAllByText(/has a leading zero/)).toHaveLength(2);
  });

  it('recomputes the counts when a column is overridden to text', () => {
    const { controller } = renderWith(() => <ImportReviewDialog />, stage);
    fireEvent.change(screen.getByLabelText('Type for column when'), { target: { value: 'text' } });
    expect(controller.getSnapshot().pendingImport?.overrides[2]).toBe('text');
    expect(document.querySelector('.mz-import-headline')!.textContent).not.toMatch(/dates/);
  });

  it('re-reads the block when the header toggle changes', () => {
    const { controller } = renderWith(() => <ImportReviewDialog />, stage);
    fireEvent.click(screen.getByLabelText('First row is a header', { selector: 'input' }));
    expect(controller.getSnapshot().pendingImport?.headerRow).toBe(false);
    expect(screen.getByLabelText('Type for column A')).toBeDefined();
  });

  it('writes nothing on cancel', () => {
    const { controller } = renderWith(() => <ImportReviewDialog />, stage);
    fireEvent.click(screen.getByText('Cancel'));
    expect(controller.sheet()!.cellCount).toBe(0);
    expect(controller.getSnapshot().pendingImport).toBeNull();
  });

  it('writes the reviewed values on confirm', () => {
    const { controller } = renderWith(() => <ImportReviewDialog />, stage);
    fireEvent.click(screen.getByText(/^Import \d+ cells$/));
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe('SEPT1');
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 1 })?.value).toBe('007');
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 2 })?.value).toBe(45322);
  });

  it('honours an override when it writes', () => {
    const { controller } = renderWith(() => <ImportReviewDialog />, stage);
    fireEvent.change(screen.getByLabelText('Type for column when'), { target: { value: 'text' } });
    fireEvent.click(screen.getByText(/^Import \d+ cells$/));
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 2 })?.value).toBe('2024-01-31');
  });

  it('makes the whole import one undoable step', () => {
    const { controller } = renderWith(() => <ImportReviewDialog />, stage);
    fireEvent.click(screen.getByText(/^Import \d+ cells$/));
    expect(controller.sheet()!.cellCount).toBeGreaterThan(1);
    controller.undo();
    expect(controller.sheet()!.cellCount).toBe(0);
  });

  it('flags an ambiguous date rather than resolving it quietly', () => {
    renderWith(() => <ImportReviewDialog />, (c) =>
      c.proposeImport({
        anchor: { sheet: 'Sheet1', row: 0, col: 0 },
        rows: [['3/4/2024']],
        overrides: ['auto'],
        headerRow: false,
        source: 'csv',
        fileName: 'dates.csv',
      }),
    );
    expect(screen.getByText('Read one of two ways')).toBeDefined();
    expect(screen.getByText(/could be 4\/3 or 3\/4/)).toBeDefined();
  });

  it('names the file for a CSV import', () => {
    renderWith(() => <ImportReviewDialog />, (c) =>
      c.proposeImport({
        anchor: { sheet: 'Sheet1', row: 0, col: 0 },
        rows: [['1']],
        overrides: ['auto'],
        headerRow: false,
        source: 'csv',
        fileName: 'orders.csv',
      }),
    );
    expect(screen.getByText(/orders\.csv/)).toBeDefined();
  });
});
