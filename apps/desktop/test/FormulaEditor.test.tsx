import { describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { FormulaEditor } from '../src/renderer/components/FormulaEditor.js';

const FUNCTIONS = ['SUM', 'SUMIF', 'SUMPRODUCT', 'IF', 'IFERROR', 'VLOOKUP', 'XLOOKUP'];

/** A wrapper that owns the draft, as the formula bar does. */
function Harness({
  initial = '',
  onCommit = () => {},
  onCancel = () => {},
}: {
  initial?: string;
  onCommit?: (value: string) => void;
  onCancel?: () => void;
}) {
  const [value, setValue] = useState(initial);
  const [height, setHeight] = useState(28);
  return (
    <FormulaEditor
      value={value}
      onChange={setValue}
      onCommit={onCommit}
      onCancel={onCancel}
      functionNames={FUNCTIONS}
      summaries={{ SUM: 'Adds its arguments' }}
      height={height}
      onHeightChange={setHeight}
    />
  );
}

const textarea = (): HTMLTextAreaElement => screen.getByLabelText('Formula') as HTMLTextAreaElement;

function typeInto(value: string, caret = value.length): void {
  const el = textarea();
  fireEvent.change(el, { target: { value } });
  el.setSelectionRange(caret, caret);
  fireEvent.keyUp(el, { key: 'a' });
}

describe('FormulaEditor', () => {
  it('is a textarea, so a formula can span lines', () => {
    render(<Harness />);
    expect(textarea().tagName).toBe('TEXTAREA');
  });

  it('highlights function names and references separately', () => {
    render(<Harness initial="=SUM(A1:B2)" />);
    expect(document.querySelector('[data-token="function"]')?.textContent).toBe('SUM');
    expect(document.querySelector('[data-token="ref"]')?.textContent).toBe('A1');
  });

  it('reproduces the source exactly in the highlight layer', () => {
    const src = '=IF(A1>0, "yes", "no")';
    render(<Harness initial={src} />);
    const spans = [...document.querySelectorAll('.mz-formula-highlight span')];
    expect(spans.map((s) => s.textContent).join('')).toBe(src);
  });

  it('colours brackets by nesting depth', () => {
    render(<Harness initial="=IF(SUM(A1),1,2)" />);
    expect(document.querySelector('.mz-depth-0')).not.toBeNull();
    expect(document.querySelector('.mz-depth-1')).not.toBeNull();
  });

  it('marks an unmatched bracket and counts it', () => {
    render(<Harness initial="=SUM(A1" />);
    expect(document.querySelector('.mz-bracket-unmatched')).not.toBeNull();
    expect(screen.getByText('1 unmatched bracket')).toBeDefined();
  });

  it('says nothing about brackets when they all match', () => {
    render(<Harness initial="=SUM(A1)" />);
    expect(screen.queryByText(/unmatched/)).toBeNull();
  });

  it('highlights the pair the caret is touching', () => {
    render(<Harness />);
    typeInto('=SUM(A1)', 5);
    expect(document.querySelectorAll('.mz-bracket-match')).toHaveLength(2);
  });

  it('shows which argument of which function the caret is in', () => {
    render(<Harness />);
    typeInto('=SUM(A1,');
    expect(screen.getByText('SUM, argument 2')).toBeDefined();
  });

  it('offers completions for a partial function name', () => {
    render(<Harness />);
    typeInto('=SUMP');
    const options = screen.getAllByRole('option').map((o) => o.textContent);
    expect(options.some((text) => text?.replace(/\s/g, '').includes('SUMPRODUCT'))).toBe(true);
  });

  it('ranks the closest completion first', () => {
    render(<Harness />);
    typeInto('=SUM');
    expect(screen.getAllByRole('option')[0]!.textContent).toContain('SUM');
  });

  it('shows the summary beside a completion', () => {
    render(<Harness />);
    typeInto('=SUM');
    expect(screen.getByText('Adds its arguments')).toBeDefined();
  });

  it('offers no completions outside a formula', () => {
    render(<Harness />);
    typeInto('SUM');
    expect(screen.queryAllByRole('option')).toHaveLength(0);
  });

  it('offers no completions inside a string literal', () => {
    render(<Harness />);
    typeInto('="SUM');
    expect(screen.queryAllByRole('option')).toHaveLength(0);
  });

  it('accepts a completion on Tab and opens the bracket', () => {
    render(<Harness />);
    typeInto('=SUMP');
    fireEvent.keyDown(textarea(), { key: 'Tab' });
    expect(textarea().value).toBe('=SUMPRODUCT(');
  });

  it('moves through the completion list with the arrow keys', () => {
    render(<Harness />);
    typeInto('=SUM');
    fireEvent.keyDown(textarea(), { key: 'ArrowDown' });
    fireEvent.keyDown(textarea(), { key: 'Enter' });
    expect(textarea().value).toBe('=SUMIF(');
  });

  it('dismisses the completion list on Escape without cancelling the edit', () => {
    const onCancel = vi.fn();
    render(<Harness onCancel={onCancel} />);
    typeInto('=SUM');
    fireEvent.keyDown(textarea(), { key: 'Escape' });
    expect(screen.queryAllByRole('option')).toHaveLength(0);
    expect(onCancel).not.toHaveBeenCalled();
  });

  it('commits on Enter when no completion is open', () => {
    const onCommit = vi.fn();
    render(<Harness onCommit={onCommit} />);
    typeInto('=1+1');
    fireEvent.keyDown(textarea(), { key: 'Escape' });
    fireEvent.keyDown(textarea(), { key: 'Enter' });
    expect(onCommit).toHaveBeenCalledWith('=1+1');
  });

  it('does not commit on Alt+Enter, which adds a line instead', () => {
    const onCommit = vi.fn();
    render(<Harness onCommit={onCommit} />);
    typeInto('=1+1');
    fireEvent.keyDown(textarea(), { key: 'Escape' });
    fireEvent.keyDown(textarea(), { key: 'Enter', altKey: true });
    expect(onCommit).not.toHaveBeenCalled();
  });

  it('cancels on Escape when no completion is open', () => {
    const onCancel = vi.fn();
    render(<Harness initial="=1+1" onCancel={onCancel} />);
    fireEvent.keyDown(textarea(), { key: 'Escape' });
    expect(onCancel).toHaveBeenCalled();
  });

  it('resizes from the keyboard', () => {
    render(<Harness />);
    const handle = screen.getByLabelText('Resize formula editor');
    const body = document.querySelector('.mz-formula-editor-body') as HTMLElement;
    const before = body.style.height;
    fireEvent.keyDown(handle, { key: 'ArrowDown' });
    expect(body.style.height).not.toBe(before);
  });
});
