import { describe, expect, it } from 'vitest';
import { CellError } from '@mirrorz/core';
import type { CellAddr, CellExplanation } from '@mirrorz/formula';
import {
  buildInspectorModel,
  buildPrecedentTree,
  describeValue,
  errorSummary,
  formatAddr,
  formatRangeAddr,
  sameAddr,
} from '../src/renderer/model/inspector-model.js';

const addr = (sheet: string, row: number, col: number): CellAddr => ({ sheet, row, col });

function explanation(partial: Partial<CellExplanation> & { addr: CellAddr }): CellExplanation {
  return {
    value: null,
    precedentCells: [],
    precedentRanges: [],
    dependents: [],
    errorRoots: [],
    ...partial,
  };
}

/**
 * A small model: D1 = C1 * 2, C1 = B4 + 1, B4 = 1/0, so the error at D1
 * originates two hops away at B4.
 */
function chainExplain(a: CellAddr): CellExplanation {
  const key = formatAddr(a);
  switch (key) {
    case 'Sheet1!D1':
      return explanation({
        addr: a,
        value: CellError.DIV0,
        formula: 'C1*2',
        precedentCells: [addr('Sheet1', 0, 2)],
        errorRoots: [addr('Sheet1', 3, 1)],
      });
    case 'Sheet1!C1':
      return explanation({
        addr: a,
        value: CellError.DIV0,
        formula: 'B4+1',
        precedentCells: [addr('Sheet1', 3, 1)],
        dependents: [addr('Sheet1', 0, 3)],
      });
    case 'Sheet1!B4':
      return explanation({
        addr: a,
        value: CellError.DIV0,
        formula: '1/0',
        dependents: [addr('Sheet1', 0, 2)],
      });
    default:
      return explanation({ addr: a, value: 1 });
  }
}

describe('address formatting', () => {
  it('renders a cell address', () => {
    expect(formatAddr(addr('Sheet1', 3, 1))).toBe('Sheet1!B4');
  });

  it('renders a range address', () => {
    expect(
      formatRangeAddr({ sheet: 'Data', startRow: 0, startCol: 0, endRow: 9, endCol: 2 }),
    ).toBe('Data!A1:C10');
  });

  it('renders a one-cell range without a colon', () => {
    expect(
      formatRangeAddr({ sheet: 'Data', startRow: 4, startCol: 4, endRow: 4, endCol: 4 }),
    ).toBe('Data!E5');
  });

  it('compares addresses across sheets', () => {
    expect(sameAddr(addr('A', 0, 0), addr('A', 0, 0))).toBe(true);
    expect(sameAddr(addr('A', 0, 0), addr('B', 0, 0))).toBe(false);
  });
});

describe('describeValue', () => {
  it('names an empty cell rather than showing nothing', () => {
    expect(describeValue(null)).toBe('empty');
  });

  it('quotes text so an empty string is visible', () => {
    expect(describeValue('')).toBe('""');
  });

  it('renders booleans in Excel form', () => {
    expect(describeValue(true)).toBe('TRUE');
  });

  it('renders an error with its detail when there is one', () => {
    expect(describeValue(new CellError('#VALUE!', 'text where a number was needed'))).toBe(
      '#VALUE! - text where a number was needed',
    );
  });

  it('renders a number plainly', () => {
    expect(describeValue(42.5)).toBe('42.5');
  });
});

describe('buildPrecedentTree', () => {
  it('walks the chain of precedents', () => {
    const tree = buildPrecedentTree(addr('Sheet1', 0, 3), chainExplain);
    expect(tree.label).toBe('Sheet1!D1');
    expect(tree.children[0]!.label).toBe('Sheet1!C1');
    expect(tree.children[0]!.children[0]!.label).toBe('Sheet1!B4');
  });

  it('marks errored cells all the way down', () => {
    const tree = buildPrecedentTree(addr('Sheet1', 0, 3), chainExplain);
    expect(tree.children[0]!.children[0]!.errored).toBe(true);
  });

  it('leaves ranges as leaves rather than exploding them', () => {
    const tree = buildPrecedentTree(
      addr('S', 0, 0),
      (a) =>
        formatAddr(a) === 'S!A1'
          ? explanation({
              addr: a,
              formula: 'SUM(B1:B5000)',
              precedentRanges: [{ sheet: 'S', startRow: 0, startCol: 1, endRow: 4999, endCol: 1 }],
            })
          : explanation({ addr: a }),
    );
    expect(tree.children).toHaveLength(1);
    expect(tree.children[0]).toMatchObject({ kind: 'range', label: 'S!B1:B5000', children: [] });
  });

  it('stops at the depth limit and says so', () => {
    const tree = buildPrecedentTree(addr('Sheet1', 0, 3), chainExplain, { maxDepth: 1 });
    expect(tree.children[0]!.truncated).toBe(true);
    expect(tree.children[0]!.children).toEqual([]);
  });

  it('stops on a cycle rather than recursing forever', () => {
    const cyclic = (a: CellAddr): CellExplanation =>
      explanation({
        addr: a,
        formula: 'other',
        precedentCells: [addr('S', 0, a.col === 0 ? 1 : 0)],
      });
    const tree = buildPrecedentTree(addr('S', 0, 0), cyclic, { maxDepth: 10 });
    let node = tree;
    let depth = 0;
    while (node.children.length > 0 && depth < 20) {
      node = node.children[0]!;
      depth++;
    }
    expect(node.cyclic).toBe(true);
  });

  it('honours the node cap', () => {
    const wide = (a: CellAddr): CellExplanation =>
      explanation({
        addr: a,
        formula: 'x',
        precedentCells: Array.from({ length: 50 }, (_, i) => addr('S', i + 1, 0)),
      });
    const tree = buildPrecedentTree(addr('S', 0, 0), wide, { maxNodes: 10 });
    expect(tree.children.length).toBeLessThanOrEqual(10);
  });
});

describe('errorSummary', () => {
  it('returns null for a cell that is not in error', () => {
    expect(errorSummary(explanation({ addr: addr('S', 0, 0), value: 1 }))).toBeNull();
  });

  it('names the originating cell in the headline', () => {
    const summary = errorSummary(chainExplain(addr('Sheet1', 0, 3)))!;
    expect(summary.headline).toBe('The problem starts at Sheet1!B4');
    expect(summary.selfInflicted).toBe(false);
  });

  it('offers the originating cell as the jump target', () => {
    const summary = errorSummary(chainExplain(addr('Sheet1', 0, 3)))!;
    expect(formatAddr(summary.primary!)).toBe('Sheet1!B4');
  });

  it('says so when the cell is its own source', () => {
    const summary = errorSummary(
      explanation({ addr: addr('S', 0, 0), value: CellError.DIV0, formula: '1/0' }),
    )!;
    expect(summary.selfInflicted).toBe(true);
    expect(summary.headline).toBe('#DIV/0! originates in this cell');
    expect(summary.primary).toBeUndefined();
  });

  it('treats a root that is the cell itself as self-inflicted', () => {
    const self = addr('S', 0, 0);
    const summary = errorSummary(
      explanation({ addr: self, value: CellError.NA, errorRoots: [self] }),
    )!;
    expect(summary.selfInflicted).toBe(true);
  });

  it('counts the other roots when there are several', () => {
    const summary = errorSummary(
      explanation({
        addr: addr('S', 0, 0),
        value: CellError.VALUE,
        errorRoots: [addr('S', 1, 0), addr('S', 2, 0), addr('S', 3, 0)],
      }),
    )!;
    expect(summary.headline).toBe('The problem starts at S!A2, and 2 other cells');
  });

  it('keeps the engine detail when there is one', () => {
    const summary = errorSummary(
      explanation({
        addr: addr('S', 0, 0),
        value: new CellError('#VALUE!', 'MID needs a number'),
      }),
    )!;
    expect(summary.detail).toBe('MID needs a number');
  });
});

describe('buildInspectorModel', () => {
  it('assembles the value, formula, tree and error together', () => {
    const model = buildInspectorModel(chainExplain(addr('Sheet1', 0, 3)), chainExplain);
    expect(model.label).toBe('Sheet1!D1');
    expect(model.formula).toBe('C1*2');
    expect(model.valueText).toBe('#DIV/0!');
    expect(model.error?.headline).toContain('Sheet1!B4');
    expect(model.precedents.children).toHaveLength(1);
  });

  it('reports an island cell', () => {
    const model = buildInspectorModel(
      explanation({ addr: addr('S', 0, 0), value: 7 }),
      () => explanation({ addr: addr('S', 0, 0), value: 7 }),
    );
    expect(model.isolated).toBe(true);
    expect(model.error).toBeNull();
  });

  it('is not isolated when something reads the cell', () => {
    const model = buildInspectorModel(
      explanation({ addr: addr('S', 0, 0), value: 7, dependents: [addr('S', 1, 0)] }),
      (a) => explanation({ addr: a, value: 7 }),
    );
    expect(model.isolated).toBe(false);
  });
});
