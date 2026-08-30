import { describe, expect, it } from 'vitest';
import { CellError, type Scalar } from '@mirrorz/core';
import { type CellAddr, DependencyGraph, cellKey } from '../src/graph.js';

const at = (row: number, col: number, sheet = 'S'): CellAddr => ({ sheet, row, col });
const range = (
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
  sheet = 'S',
) => ({ sheet, startRow, startCol, endRow, endCol });

/** A tiny sheet plus a graph, so tests read like the model they describe. */
function harness() {
  const values = new Map<string, Scalar>();
  const formulas = new Map<string, () => Scalar>();
  const graph = new DependencyGraph();
  let computeCount = 0;

  const key = (a: CellAddr) => cellKey(a.sheet, a.row, a.col);
  const getValue = (a: CellAddr): Scalar => values.get(key(a)) ?? null;
  const setValue = (a: CellAddr, v: Scalar) => void values.set(key(a), v);
  const compute = (a: CellAddr): Scalar => {
    computeCount++;
    return formulas.get(key(a))?.() ?? null;
  };

  return {
    graph,
    values,
    getValue,
    setValue,
    setLiteral(a: CellAddr, v: Scalar) {
      values.set(key(a), v);
    },
    setFormula(
      a: CellAddr,
      precedents: { cells?: CellAddr[]; ranges?: ReturnType<typeof range>[] },
      fn: () => Scalar,
      volatile = false,
    ) {
      graph.setFormula(
        a,
        { cells: precedents.cells ?? [], ranges: precedents.ranges ?? [] },
        volatile,
      );
      formulas.set(key(a), fn);
    },
    recalc(seeds: CellAddr[]) {
      computeCount = 0;
      return graph.recalculate(seeds, compute, getValue, setValue);
    },
    get computeCount() {
      return computeCount;
    },
  };
}

describe('registration', () => {
  it('tracks direct cell dependents', () => {
    const h = harness();
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => 1);
    expect(h.graph.directDependents('S', 0, 0)).toEqual([at(1, 0)]);
    expect(h.graph.directDependents('S', 5, 5)).toEqual([]);
  });

  it('tracks dependents through a range without one edge per cell', () => {
    const h = harness();
    h.setFormula(at(0, 1), { ranges: [range(0, 0, 99_999, 0)] }, () => 1);
    // Any cell inside the range finds the formula.
    expect(h.graph.directDependents('S', 0, 0)).toEqual([at(0, 1)]);
    expect(h.graph.directDependents('S', 50_000, 0)).toEqual([at(0, 1)]);
    expect(h.graph.directDependents('S', 99_999, 0)).toEqual([at(0, 1)]);
    // A cell outside it does not.
    expect(h.graph.directDependents('S', 100_000, 0)).toEqual([]);
    expect(h.graph.directDependents('S', 0, 5)).toEqual([]);
  });

  it('keeps sheets separate', () => {
    const h = harness();
    h.setFormula(at(1, 0, 'A'), { cells: [at(0, 0, 'A')] }, () => 1);
    expect(h.graph.directDependents('A', 0, 0)).toHaveLength(1);
    expect(h.graph.directDependents('B', 0, 0)).toHaveLength(0);
  });

  it('replaces precedents when a formula is edited', () => {
    const h = harness();
    h.setFormula(at(2, 0), { cells: [at(0, 0)] }, () => 1);
    h.setFormula(at(2, 0), { cells: [at(1, 0)] }, () => 1);
    expect(h.graph.directDependents('S', 0, 0)).toEqual([]);
    expect(h.graph.directDependents('S', 1, 0)).toEqual([at(2, 0)]);
  });

  it('removes a formula cleanly', () => {
    const h = harness();
    h.setFormula(at(1, 0), { cells: [at(0, 0)], ranges: [range(0, 0, 10, 10)] }, () => 1);
    h.graph.removeFormula(at(1, 0));
    expect(h.graph.directDependents('S', 0, 0)).toEqual([]);
    expect(h.graph.size).toBe(0);
  });

  it('reports precedents for tracing', () => {
    const h = harness();
    h.setFormula(at(2, 0), { cells: [at(0, 0)], ranges: [range(0, 1, 5, 1)] }, () => 1);
    const p = h.graph.precedentsOf(at(2, 0))!;
    expect(p.cells).toEqual([at(0, 0)]);
    expect(p.ranges).toEqual([range(0, 1, 5, 1)]);
    expect(h.graph.precedentsOf(at(9, 9))).toBeUndefined();
  });
});

describe('recalculation order', () => {
  it('evaluates a chain from the bottom up', () => {
    const h = harness();
    h.setLiteral(at(0, 0), 1);
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => (h.getValue(at(0, 0)) as number) + 1);
    h.setFormula(at(2, 0), { cells: [at(1, 0)] }, () => (h.getValue(at(1, 0)) as number) + 1);
    h.setFormula(at(3, 0), { cells: [at(2, 0)] }, () => (h.getValue(at(2, 0)) as number) + 1);

    const result = h.recalc([at(0, 0)]);
    expect(h.getValue(at(3, 0))).toBe(4);
    // Each cell computed exactly once, in dependency order.
    expect(result.evaluated).toEqual([at(1, 0), at(2, 0), at(3, 0)]);
  });

  it('evaluates a diamond once per node', () => {
    const h = harness();
    h.setLiteral(at(0, 0), 10);
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => (h.getValue(at(0, 0)) as number) * 2);
    h.setFormula(at(1, 1), { cells: [at(0, 0)] }, () => (h.getValue(at(0, 0)) as number) * 3);
    h.setFormula(at(2, 0), { cells: [at(1, 0), at(1, 1)] }, () =>
      (h.getValue(at(1, 0)) as number) + (h.getValue(at(1, 1)) as number),
    );

    const result = h.recalc([at(0, 0)]);
    expect(h.getValue(at(2, 0))).toBe(50);
    expect(result.evaluated.filter((a) => a.row === 2)).toHaveLength(1);
    // The join must come after both branches.
    const order = result.evaluated.map((a) => `${a.row},${a.col}`);
    expect(order.indexOf('2,0')).toBeGreaterThan(order.indexOf('1,0'));
    expect(order.indexOf('2,0')).toBeGreaterThan(order.indexOf('1,1'));
  });

  it('handles a deep chain without exhausting the stack', () => {
    const h = harness();
    const depth = 20_000;
    h.setLiteral(at(0, 0), 1);
    for (let i = 1; i <= depth; i++) {
      const prev = at(i - 1, 0);
      h.setFormula(at(i, 0), { cells: [prev] }, () => (h.getValue(prev) as number) + 1);
    }
    // A recursive topological sort would overflow well before this depth.
    const result = h.recalc([at(0, 0)]);
    expect(h.getValue(at(depth, 0))).toBe(depth + 1);
    expect(result.evaluated).toHaveLength(depth);
  });

  it('recalculates through a range dependency', () => {
    const h = harness();
    for (let i = 0; i < 5; i++) h.setLiteral(at(i, 0), i + 1);
    h.setFormula(at(0, 1), { ranges: [range(0, 0, 4, 0)] }, () => {
      let sum = 0;
      for (let i = 0; i < 5; i++) sum += (h.getValue(at(i, 0)) as number) ?? 0;
      return sum;
    });

    h.recalc([at(0, 0)]);
    expect(h.getValue(at(0, 1))).toBe(15);

    h.setLiteral(at(2, 0), 30);
    h.recalc([at(2, 0)]);
    expect(h.getValue(at(0, 1))).toBe(42);
  });
});

describe('change-based propagation', () => {
  it('stops propagating when a value does not change', () => {
    const h = harness();
    h.setLiteral(at(0, 0), 5);
    // A formula whose result is constant regardless of its precedent.
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => 42);
    h.setFormula(at(2, 0), { cells: [at(1, 0)] }, () => (h.getValue(at(1, 0)) as number) + 1);

    h.recalc([at(0, 0)]);
    expect(h.getValue(at(2, 0))).toBe(43);

    // Second pass: the middle cell recomputes to the same 42, so the cell above
    // it is evaluated but reports no change.
    const second = h.recalc([at(0, 0)]);
    expect(second.changed).toHaveLength(0);
  });

  it('treats an unchanged error as unchanged', () => {
    const h = harness();
    h.setLiteral(at(0, 0), 1);
    // A fresh CellError object each time must not look like a new value.
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => new CellError('#DIV/0!'));
    h.recalc([at(0, 0)]);
    const second = h.recalc([at(0, 0)]);
    expect(second.changed).toHaveLength(0);
  });

  it('reports what actually changed', () => {
    const h = harness();
    h.setLiteral(at(0, 0), 1);
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => (h.getValue(at(0, 0)) as number) * 10);
    h.recalc([at(0, 0)]);
    h.setLiteral(at(0, 0), 2);
    const result = h.recalc([at(0, 0)]);
    expect(result.changed).toEqual([at(1, 0)]);
    expect(h.getValue(at(1, 0))).toBe(20);
  });
});

describe('volatile formulas', () => {
  it('recalculates on every pass even with an unrelated seed', () => {
    const h = harness();
    let ticks = 0;
    h.setFormula(at(0, 0), {}, () => ++ticks, true);
    h.setLiteral(at(9, 9), 1);

    h.recalc([at(9, 9)]);
    expect(h.getValue(at(0, 0))).toBe(1);
    h.recalc([at(9, 9)]);
    expect(h.getValue(at(0, 0))).toBe(2);
  });

  it('drags its dependents along', () => {
    const h = harness();
    let ticks = 0;
    h.setFormula(at(0, 0), {}, () => ++ticks, true);
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => (h.getValue(at(0, 0)) as number) * 100);
    h.recalc([]);
    h.recalc([]);
    expect(h.getValue(at(1, 0))).toBe(200);
  });
});

describe('circular references', () => {
  it('detects a two-cell cycle and reports it', () => {
    const h = harness();
    h.setFormula(at(0, 0), { cells: [at(1, 0)] }, () => (h.getValue(at(1, 0)) as number) + 1);
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => (h.getValue(at(0, 0)) as number) + 1);
    const result = h.recalc([at(0, 0)]);
    expect(result.circular).toHaveLength(2);
  });

  it('leaves circular cells at zero, which is what Excel shows', () => {
    const h = harness();
    h.setFormula(at(0, 0), { cells: [at(1, 0)] }, () => 1);
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => 1);
    h.recalc([at(0, 0)]);
    expect(h.getValue(at(0, 0))).toBe(0);
    expect(h.getValue(at(1, 0))).toBe(0);
  });

  it('detects a self-reference', () => {
    const h = harness();
    h.setFormula(at(0, 0), { cells: [at(0, 0)] }, () => 1);
    const result = h.recalc([at(0, 0)]);
    expect(result.circular.length).toBeGreaterThan(0);
  });

  it('does not mistake a diamond for a cycle', () => {
    const h = harness();
    h.setLiteral(at(0, 0), 1);
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => 2);
    h.setFormula(at(1, 1), { cells: [at(0, 0)] }, () => 3);
    h.setFormula(at(2, 0), { cells: [at(1, 0), at(1, 1)] }, () => 5);
    expect(h.recalc([at(0, 0)]).circular).toHaveLength(0);
  });

  it('converges under iterative calculation', () => {
    const values = new Map<string, Scalar>();
    const graph = new DependencyGraph({ iterative: true, maxIterations: 200, maxChange: 1e-6 });
    const key = (a: CellAddr) => cellKey(a.sheet, a.row, a.col);
    const get = (a: CellAddr): Scalar => values.get(key(a)) ?? null;
    const set = (a: CellAddr, v: Scalar) => void values.set(key(a), v);

    // x = (x + 2/x) / 2 converges on the square root of 2.
    graph.setFormula(at(0, 0), { cells: [at(0, 0)], ranges: [] });
    const compute = (a: CellAddr): Scalar => {
      const x = (get(a) as number) || 1;
      return (x + 2 / x) / 2;
    };
    const result = graph.recalculate([at(0, 0)], compute, get, set);
    expect(result.circular).toHaveLength(1);
    expect(get(at(0, 0)) as number).toBeCloseTo(Math.SQRT2, 8);
  });
});

describe('error root-cause search', () => {
  it('names the cell that originated an error', () => {
    const h = harness();
    // A1 is a literal error; B1 reads it; C1 reads B1.
    h.setLiteral(at(0, 0), new CellError('#DIV/0!'));
    h.setFormula(at(0, 1), { cells: [at(0, 0)] }, () => new CellError('#DIV/0!'));
    h.setFormula(at(0, 2), { cells: [at(0, 1)] }, () => new CellError('#DIV/0!'));
    h.recalc([at(0, 0)]);

    const roots = h.graph.findErrorRoots(at(0, 2), h.getValue);
    expect(roots).toEqual([at(0, 0)]);
  });

  it('names a formula that produced the error itself', () => {
    const h = harness();
    h.setLiteral(at(0, 0), 0);
    // B1 divides by a healthy zero, so B1 is the root, not A1.
    h.setFormula(at(0, 1), { cells: [at(0, 0)] }, () => new CellError('#DIV/0!'));
    h.setFormula(at(0, 2), { cells: [at(0, 1)] }, () => new CellError('#DIV/0!'));
    h.recalc([at(0, 0)]);

    expect(h.graph.findErrorRoots(at(0, 2), h.getValue)).toEqual([at(0, 1)]);
  });

  it('finds a root through a range precedent', () => {
    const h = harness();
    for (let i = 0; i < 5; i++) h.setLiteral(at(i, 0), i);
    h.setLiteral(at(3, 0), new CellError('#VALUE!'));
    h.setFormula(at(0, 1), { ranges: [range(0, 0, 4, 0)] }, () => new CellError('#VALUE!'));
    h.recalc([at(0, 0)]);

    expect(h.graph.findErrorRoots(at(0, 1), h.getValue)).toEqual([at(3, 0)]);
  });

  it('returns nothing for a healthy cell', () => {
    const h = harness();
    h.setLiteral(at(0, 0), 1);
    h.setFormula(at(0, 1), { cells: [at(0, 0)] }, () => 2);
    h.recalc([at(0, 0)]);
    expect(h.graph.findErrorRoots(at(0, 1), h.getValue)).toEqual([]);
  });

  it('reports several roots when an error has more than one source', () => {
    const h = harness();
    h.setLiteral(at(0, 0), new CellError('#REF!'));
    h.setLiteral(at(1, 0), new CellError('#NAME?'));
    h.setFormula(at(2, 0), { cells: [at(0, 0), at(1, 0)] }, () => new CellError('#REF!'));
    h.recalc([at(0, 0)]);
    const roots = h.graph.findErrorRoots(at(2, 0), h.getValue);
    expect(roots).toHaveLength(2);
  });

  it('terminates on a cyclic error chain', () => {
    const h = harness();
    h.setFormula(at(0, 0), { cells: [at(1, 0)] }, () => new CellError('#VALUE!'));
    h.setFormula(at(1, 0), { cells: [at(0, 0)] }, () => new CellError('#VALUE!'));
    h.setValue(at(0, 0), new CellError('#VALUE!'));
    h.setValue(at(1, 0), new CellError('#VALUE!'));
    expect(() => h.graph.findErrorRoots(at(0, 0), h.getValue)).not.toThrow();
  });
});

describe('scale', () => {
  it('answers the reverse lookup quickly with many overlapping ranges', () => {
    const h = harness();
    // A thousand column-sized ranges, the shape a real model has.
    for (let c = 0; c < 1000; c++) {
      h.setFormula(at(0, c + 1), { ranges: [range(0, c, 100_000, c)] }, () => 1);
    }
    const started = performance.now();
    for (let i = 0; i < 1000; i++) {
      h.graph.directDependents('S', i * 37, i % 1000);
    }
    const elapsed = performance.now() - started;
    // Without the block index this degrades to scanning all 1000 ranges per
    // lookup; the bound is generous but catches that regression.
    expect(elapsed).toBeLessThan(2000);
  });

  it('recalculates a wide fan-out once per dependent', () => {
    const h = harness();
    h.setLiteral(at(0, 0), 1);
    for (let i = 1; i <= 5000; i++) {
      h.setFormula(at(i, 0), { cells: [at(0, 0)] }, () => (h.getValue(at(0, 0)) as number) + 1);
    }
    const result = h.recalc([at(0, 0)]);
    expect(result.evaluated).toHaveLength(5000);
    expect(h.computeCount).toBe(5000);
  });
});
