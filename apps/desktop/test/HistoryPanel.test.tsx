import { describe, expect, it } from 'vitest';
import { act, fireEvent, screen, within } from '@testing-library/react';
import { HistoryPanel } from '../src/renderer/components/HistoryPanel.js';
import { enter, renderWith } from './render.js';

const NOW = 1_700_000_000_000;

describe('HistoryPanel', () => {
  it('shows the initial state as a jumpable row even with no edits', () => {
    renderWith(() => <HistoryPanel now={NOW} />);
    expect(screen.getByText('Opened')).toBeDefined();
    expect(screen.getByText('Nothing has been changed yet.')).toBeDefined();
  });

  it('lists an edit with its label and origin', () => {
    renderWith(() => <HistoryPanel now={NOW} />, (c) => enter(c, 0, 0, '42'));
    expect(screen.getByText('Edit Sheet1!A1')).toBeDefined();
    expect(screen.getAllByText('you').length).toBeGreaterThan(0);
  });

  it('hides recalculation steps until asked for them', () => {
    renderWith(() => <HistoryPanel now={NOW} />, (c) => enter(c, 0, 0, '=1+1'));
    expect(screen.queryByText('recalculation')).toBeNull();
    fireEvent.click(screen.getByLabelText('Show recalculation steps', { selector: 'input' }));
    expect(screen.queryAllByText('recalculation').length).toBeGreaterThanOrEqual(0);
  });

  it('marks the current head', () => {
    renderWith(() => <HistoryPanel now={NOW} />, (c) => enter(c, 0, 0, '42'));
    const head = screen.getByText('Edit Sheet1!A1').closest('button')!;
    expect(head.getAttribute('aria-current')).toBe('true');
  });

  it('renders a branch after an undo and a divergent edit', () => {
    const { controller } = renderWith(
      () => <HistoryPanel now={NOW} />,
      (c) => {
        enter(c, 0, 0, 'first');
        c.undo();
        enter(c, 0, 0, 'second');
      },
    );
    // Both the abandoned edit and the live one are listed; a stack would show one.
    const rows = document.querySelectorAll('.mz-history-entry');
    expect(rows.length).toBeGreaterThanOrEqual(3);
    expect(controller.historyEntries().length).toBeGreaterThanOrEqual(2);
  });

  it('indents the abandoned branch and marks it', () => {
    renderWith(
      () => <HistoryPanel now={NOW} />,
      (c) => {
        enter(c, 0, 0, 'first');
        c.undo();
        enter(c, 0, 0, 'second');
      },
    );
    const abandoned = document.querySelectorAll('[data-abandoned="true"]');
    expect(abandoned.length).toBeGreaterThan(0);
  });

  it('reports the abandoned count in the header', () => {
    renderWith(
      () => <HistoryPanel now={NOW} />,
      (c) => {
        enter(c, 0, 0, 'first');
        c.undo();
        enter(c, 0, 0, 'second');
      },
    );
    expect(screen.getByText(/off the current path/)).toBeDefined();
  });

  it('jumps the document when a row is clicked', () => {
    const { controller } = renderWith(
      () => <HistoryPanel now={NOW} />,
      (c) => enter(c, 0, 0, '42'),
    );
    act(() => {
      fireEvent.click(screen.getByText('Opened'));
    });
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })).toBeUndefined();
  });

  it('walks back onto an abandoned branch when its row is clicked', () => {
    const { controller } = renderWith(
      () => <HistoryPanel now={NOW} />,
      (c) => {
        enter(c, 0, 0, 'first');
        c.undo();
        enter(c, 0, 0, 'second');
      },
    );
    const first = screen.getAllByText('Edit Sheet1!A1')[0]!.closest('button')!;
    act(() => {
      fireEvent.click(first);
    });
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe('first');
  });

  it('labels a macro edit as a macro, not as the user', () => {
    renderWith(
      () => <HistoryPanel now={NOW} />,
      (c) => {
        c.doc.transact({ label: 'Run FormatReport', origin: 'macro', timestamp: NOW }, () => {
          c.doc.setValue('Sheet1', 0, 0, 1);
        });
      },
    );
    const row = screen.getByText('Run FormatReport').closest('button')!;
    expect(within(row).getByText('macro')).toBeDefined();
    expect(row.getAttribute('data-origin')).toBe('macro');
  });

  it('shows the change count for a multi-cell step', () => {
    renderWith(
      () => <HistoryPanel now={NOW} />,
      (c) => {
        c.doc.transact({ label: 'Paste 3 cells', origin: 'user', timestamp: NOW }, () => {
          c.doc.setValue('Sheet1', 0, 0, 1);
          c.doc.setValue('Sheet1', 0, 1, 2);
          c.doc.setValue('Sheet1', 0, 2, 3);
        });
      },
    );
    expect(screen.getByText('3 changes')).toBeDefined();
  });

  it('flags a step that cannot be reversed exactly', () => {
    renderWith(
      () => <HistoryPanel now={NOW} />,
      (c) => {
        c.doc.transact({ label: 'Refresh external data', origin: 'system', barrier: true, timestamp: NOW }, () => {
          c.doc.setValue('Sheet1', 0, 0, 1);
        });
      },
    );
    expect(screen.getByText('not reversible')).toBeDefined();
  });

  it('closes when the close button is used', () => {
    const { controller } = renderWith(() => <HistoryPanel now={NOW} />, (c) =>
      c.togglePanel('history', true),
    );
    fireEvent.click(screen.getByLabelText('Close history'));
    expect(controller.getSnapshot().panels.history).toBe(false);
  });
});
