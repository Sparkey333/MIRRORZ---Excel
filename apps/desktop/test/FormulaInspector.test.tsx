import { describe, expect, it } from 'vitest';
import { fireEvent, screen } from '@testing-library/react';
import { FormulaInspector } from '../src/renderer/components/FormulaInspector.js';
import { enter, renderWith } from './render.js';

/**
 * A chain that breaks two hops away: A1 = 0, B1 = 1/A1, C1 = B1 + 1. Selecting
 * C1 should name B1 as the origin rather than making the user walk the arrows.
 */
function brokenChain(): (c: import('../src/renderer/state/controller.js').AppController) => void {
  return (controller) => {
    enter(controller, 0, 0, '0');
    enter(controller, 0, 1, '=1/A1');
    enter(controller, 0, 2, '=B1+1');
    controller.selectCell(0, 2);
  };
}

describe('FormulaInspector', () => {
  it('shows the selected address and its value', () => {
    renderWith(() => <FormulaInspector />, (c) => {
      enter(c, 0, 0, '42');
      c.selectCell(0, 0);
    });
    expect(screen.getByText('Sheet1!A1')).toBeDefined();
    expect(screen.getByText('42')).toBeDefined();
  });

  it('says plainly when a cell holds a literal', () => {
    renderWith(() => <FormulaInspector />, (c) => enter(c, 0, 0, '42'));
    expect(screen.getByText('This cell holds a literal value.')).toBeDefined();
  });

  it('shows the formula source', () => {
    renderWith(() => <FormulaInspector />, (c) => {
      enter(c, 0, 0, '1');
      enter(c, 0, 1, '=A1*2');
      c.selectCell(0, 1);
    });
    expect(screen.getByText('=A1*2')).toBeDefined();
  });

  it('lists the precedents', () => {
    renderWith(() => <FormulaInspector />, (c) => {
      enter(c, 0, 0, '1');
      enter(c, 0, 1, '=A1*2');
      c.selectCell(0, 1);
    });
    expect(screen.getByRole('button', { name: 'Sheet1!A1' })).toBeDefined();
  });

  it('says so when a cell reads nothing', () => {
    renderWith(() => <FormulaInspector />, (c) => enter(c, 0, 0, '42'));
    expect(screen.getByText('Nothing. This cell does not read any other cell.')).toBeDefined();
  });

  it('lists what reads the cell', () => {
    renderWith(() => <FormulaInspector />, (c) => {
      enter(c, 0, 0, '1');
      enter(c, 0, 1, '=A1*2');
      c.selectCell(0, 0);
    });
    expect(screen.getByRole('button', { name: 'Sheet1!B1' })).toBeDefined();
  });

  it('calls out an island cell', () => {
    renderWith(() => <FormulaInspector />, (c) => enter(c, 0, 0, '42'));
    expect(screen.getByText(/it is an island/)).toBeDefined();
  });

  it('names the cell where an error actually starts', () => {
    renderWith(() => <FormulaInspector />, brokenChain());
    const alert = screen.getByRole('alert');
    expect(alert.textContent).toContain('The problem starts at Sheet1!B1');
  });

  it('offers a jump to the originating cell', () => {
    const { controller } = renderWith(() => <FormulaInspector />, brokenChain());
    fireEvent.click(screen.getByRole('button', { name: 'Go to Sheet1!B1' }));
    expect(controller.getSnapshot().selection.active).toEqual({ row: 0, col: 1 });
  });

  it('shows the error value itself in the header', () => {
    renderWith(() => <FormulaInspector />, brokenChain());
    expect(document.querySelector('.mz-inspector-value[data-error="true"]')).not.toBeNull();
  });

  it('expands an errored precedent branch by default so the path is visible', () => {
    renderWith(() => <FormulaInspector />, brokenChain());
    expect(screen.getByRole('button', { name: 'Collapse Sheet1!B1' })).toBeDefined();
  });

  it('collapses a branch when its twisty is clicked', () => {
    renderWith(() => <FormulaInspector />, brokenChain());
    fireEvent.click(screen.getByRole('button', { name: 'Collapse Sheet1!B1' }));
    expect(screen.getByRole('button', { name: 'Expand Sheet1!B1' })).toBeDefined();
  });

  it('jumps to a precedent when its label is clicked', () => {
    const { controller } = renderWith(() => <FormulaInspector />, (c) => {
      enter(c, 4, 0, '1');
      enter(c, 0, 1, '=A5*2');
      c.selectCell(0, 1);
    });
    fireEvent.click(screen.getByRole('button', { name: 'Sheet1!A5' }));
    expect(controller.getSnapshot().selection.active).toEqual({ row: 4, col: 0 });
  });

  it('shows a range precedent as a range rather than its cells', () => {
    renderWith(() => <FormulaInspector />, (c) => {
      enter(c, 0, 0, '=SUM(A2:A100)');
      c.selectCell(0, 0);
    });
    const range = screen.getByRole('button', { name: 'Sheet1!A2:A100' });
    expect(range.hasAttribute('disabled')).toBe(true);
  });

  it('closes when asked', () => {
    const { controller } = renderWith(() => <FormulaInspector />, (c) =>
      c.togglePanel('inspector', true),
    );
    fireEvent.click(screen.getByLabelText('Close inspector'));
    expect(controller.getSnapshot().panels.inspector).toBe(false);
  });
});
