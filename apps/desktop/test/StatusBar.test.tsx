import { describe, expect, it } from 'vitest';
import { act, fireEvent, screen } from '@testing-library/react';
import { StatusBar } from '../src/renderer/components/StatusBar.js';
import { enter, renderWith } from './render.js';
import type { AppController } from '../src/renderer/state/controller.js';

const NOW = 1_700_000_000_000;

function selectBlock(controller: AppController): void {
  controller.setSelection({
    sheet: 'Sheet1',
    active: { row: 0, col: 0 },
    ranges: [{ start: { row: 0, col: 0 }, end: { row: 2, col: 0 } }],
  });
}

describe('StatusBar', () => {
  it('names the selection in A1 terms', () => {
    renderWith(() => <StatusBar now={NOW} />, selectBlock);
    expect(screen.getByText('A1:A3')).toBeDefined();
  });

  it('counts the populated cells in the selection', () => {
    renderWith(() => <StatusBar now={NOW} />, (c) => {
      enter(c, 0, 0, '1');
      enter(c, 1, 0, '2');
      selectBlock(c);
    });
    expect(screen.getByText('Count 2')).toBeDefined();
  });

  it('shows the aggregates for numbers', () => {
    renderWith(() => <StatusBar now={NOW} />, (c) => {
      enter(c, 0, 0, '1');
      enter(c, 1, 0, '2');
      enter(c, 2, 0, '6');
      selectBlock(c);
    });
    expect(screen.getByText('Sum 9')).toBeDefined();
    expect(screen.getByText('Average 3')).toBeDefined();
    expect(screen.getByText('Min 1')).toBeDefined();
    expect(screen.getByText('Max 6')).toBeDefined();
  });

  it('omits the aggregates when nothing numeric is selected', () => {
    renderWith(() => <StatusBar now={NOW} />, (c) => {
      enter(c, 0, 0, 'alpha');
      selectBlock(c);
    });
    expect(screen.queryByText(/^Sum/)).toBeNull();
  });

  it('warns about an error in the selection rather than summing round it', () => {
    renderWith(() => <StatusBar now={NOW} />, (c) => {
      enter(c, 0, 0, '=1/0');
      selectBlock(c);
    });
    expect(screen.getByText('#DIV/0! in selection')).toBeDefined();
  });

  it('reports the sheet dimensions', () => {
    renderWith(() => <StatusBar now={NOW} />, (c) => enter(c, 4, 2, '1'));
    expect(screen.getByText('5 rows x 3 columns')).toBeDefined();
  });

  it('says the sheet is empty when it is', () => {
    renderWith(() => <StatusBar now={NOW} />);
    expect(screen.getByText('empty')).toBeDefined();
  });

  it('shows the calculation mode, and flags manual', () => {
    const { controller } = renderWith(() => <StatusBar now={NOW} />);
    expect(screen.getByText('Automatic calculation')).toBeDefined();
    act(() => controller.setCalcMode('manual'));
    expect(screen.getByText('Manual calculation')).toBeDefined();
  });

  it('reports the last recalculation', () => {
    renderWith(() => <StatusBar now={NOW} />, (c) => enter(c, 0, 0, '=1+1'));
    expect(screen.getByText(/Recalculated \d+ cells in/)).toBeDefined();
  });

  it('says so before anything has been recalculated', () => {
    renderWith(() => <StatusBar now={NOW} />);
    expect(screen.getByText('Not yet recalculated')).toBeDefined();
  });

  it('flags unsaved changes', () => {
    renderWith(() => <StatusBar now={NOW} />, (c) => enter(c, 0, 0, '1'));
    expect(screen.getByText('Unsaved changes')).toBeDefined();
  });

  it('shows and dismisses a transient message', () => {
    const { controller } = renderWith(() => <StatusBar now={NOW} />, (c) =>
      c.setMessage('Something happened'),
    );
    fireEvent.click(screen.getByText('Something happened'));
    expect(controller.getSnapshot().message).toBeNull();
  });
});
