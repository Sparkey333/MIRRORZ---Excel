/**
 * Dependency graph and incremental recalculation.
 *
 * Three design decisions carry most of the weight here:
 *
 * A formula that reads a range gets ONE edge to a range vertex, not one edge per
 * cell. `SUM(A1:A100000)` must not create a hundred thousand edges, and the
 * reverse lookup - "which formulas depend on the cell that just changed?" - has
 * to be answered without scanning every range in the workbook. A block index
 * over the grid does that: each range registers in the coarse blocks it
 * overlaps, and a dirty cell only exact-tests the ranges in its own block.
 *
 * Topological ordering is iterative, not recursive. Real workbooks have
 * dependency chains tens of thousands deep, and a recursive Tarjan blows the
 * JavaScript stack on files that Excel opens without complaint.
 *
 * Dirtiness propagates only when a value actually CHANGES. On real models this
 * cuts the recalculated set dramatically, because most edits do not alter the
 * result of the formulas above them.
 *
 * The graph is also what makes formula debugging possible: precedents and
 * dependents are already here, so "which cell made this one an error" is a walk
 * rather than a new subsystem.
 */

import { MAX_COLS, type Scalar, isError } from '@mirrorz/core';

/** A cell address qualified by sheet. */
export interface CellAddr {
  sheet: string;
  row: number;
  col: number;
}

export interface RangeAddr {
  sheet: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export function cellKey(sheet: string, row: number, col: number): string {
  return `${sheet}!${row * MAX_COLS + col}`;
}

export function rangeKey(r: RangeAddr): string {
  return `${r.sheet}!${r.startRow}:${r.startCol}:${r.endRow}:${r.endCol}`;
}

/**
 * Grid blocks for the reverse lookup.
 *
 * 128 x 128 is a deliberate compromise: small enough that a dirty cell's block
 * holds few candidate ranges, large enough that a range spanning a whole column
 * registers in ~8000 blocks rather than a million.
 */
const BLOCK_SIZE = 128;

function blocksOf(r: RangeAddr): string[] {
  const out: string[] = [];
  const r0 = Math.floor(r.startRow / BLOCK_SIZE);
  const r1 = Math.floor(r.endRow / BLOCK_SIZE);
  const c0 = Math.floor(r.startCol / BLOCK_SIZE);
  const c1 = Math.floor(r.endCol / BLOCK_SIZE);
  for (let br = r0; br <= r1; br++) {
    for (let bc = c0; bc <= c1; bc++) out.push(`${r.sheet}#${br}#${bc}`);
  }
  return out;
}

function blockOfCell(sheet: string, row: number, col: number): string {
  return `${sheet}#${Math.floor(row / BLOCK_SIZE)}#${Math.floor(col / BLOCK_SIZE)}`;
}

function rangeContains(r: RangeAddr, sheet: string, row: number, col: number): boolean {
  return (
    r.sheet === sheet &&
    row >= r.startRow &&
    row <= r.endRow &&
    col >= r.startCol &&
    col <= r.endCol
  );
}

interface FormulaVertex {
  key: string;
  addr: CellAddr;
  /** Cells this formula reads directly. */
  cellPrecedents: string[];
  /** Ranges this formula reads. */
  rangePrecedents: RangeAddr[];
  volatile: boolean;
  /** Formulas that read this cell. */
  dependents: Set<string>;
}

export interface RecalcResult {
  /** Cells recomputed, in evaluation order. */
  evaluated: CellAddr[];
  /** Cells whose value actually changed. */
  changed: CellAddr[];
  /** Cells that form circular references. */
  circular: CellAddr[];
  /** Wall-clock milliseconds, for the status bar and for benchmarks. */
  elapsedMs: number;
}

export interface GraphOptions {
  /**
   * Iterative calculation, matching Excel's option of the same name. Without it
   * a cycle is reported and its cells hold zero, which is Excel's default
   * behaviour rather than an error value.
   */
  iterative?: boolean;
  maxIterations?: number;
  maxChange?: number;
}

export class DependencyGraph {
  private readonly formulas = new Map<string, FormulaVertex>();
  /** Cell key -> formula keys that read that cell directly. */
  private readonly cellDependents = new Map<string, Set<string>>();
  /** Block key -> formula keys whose ranges overlap that block. */
  private readonly blockIndex = new Map<string, Set<string>>();
  private readonly volatiles = new Set<string>();

  constructor(private readonly options: GraphOptions = {}) {}

  get size(): number {
    return this.formulas.size;
  }

  /**
   * Record a formula cell and what it reads. Replaces any previous registration
   * for the same cell, which is what an edit does.
   */
  setFormula(
    addr: CellAddr,
    precedents: { cells: CellAddr[]; ranges: RangeAddr[] },
    volatile = false,
  ): void {
    const key = cellKey(addr.sheet, addr.row, addr.col);
    this.removeFormula(addr);

    const vertex: FormulaVertex = {
      key,
      addr,
      cellPrecedents: precedents.cells.map((c) => cellKey(c.sheet, c.row, c.col)),
      rangePrecedents: precedents.ranges,
      volatile,
      dependents: new Set(),
    };
    this.formulas.set(key, vertex);

    for (const p of vertex.cellPrecedents) {
      let set = this.cellDependents.get(p);
      if (!set) {
        set = new Set();
        this.cellDependents.set(p, set);
      }
      set.add(key);
    }
    for (const r of vertex.rangePrecedents) {
      for (const b of blocksOf(r)) {
        let set = this.blockIndex.get(b);
        if (!set) {
          set = new Set();
          this.blockIndex.set(b, set);
        }
        set.add(key);
      }
    }
    if (volatile) this.volatiles.add(key);
  }

  removeFormula(addr: CellAddr): void {
    const key = cellKey(addr.sheet, addr.row, addr.col);
    const existing = this.formulas.get(key);
    if (!existing) return;

    for (const p of existing.cellPrecedents) {
      const set = this.cellDependents.get(p);
      if (set) {
        set.delete(key);
        if (set.size === 0) this.cellDependents.delete(p);
      }
    }
    for (const r of existing.rangePrecedents) {
      for (const b of blocksOf(r)) {
        const set = this.blockIndex.get(b);
        if (set) {
          set.delete(key);
          if (set.size === 0) this.blockIndex.delete(b);
        }
      }
    }
    this.volatiles.delete(key);
    this.formulas.delete(key);
  }

  /** Formulas that read this cell, directly or through a range. */
  directDependents(sheet: string, row: number, col: number): CellAddr[] {
    const out = new Set<string>();
    const direct = this.cellDependents.get(cellKey(sheet, row, col));
    if (direct) for (const k of direct) out.add(k);

    // Only the ranges registered in this cell's own block need exact-testing,
    // which is what keeps the reverse lookup cheap on a large workbook.
    const candidates = this.blockIndex.get(blockOfCell(sheet, row, col));
    if (candidates) {
      for (const k of candidates) {
        const vertex = this.formulas.get(k);
        if (!vertex) continue;
        if (vertex.rangePrecedents.some((r) => rangeContains(r, sheet, row, col))) out.add(k);
      }
    }
    return [...out].map((k) => this.formulas.get(k)!.addr);
  }

  /** Cells and ranges a formula reads. The precedent half of formula tracing. */
  precedentsOf(addr: CellAddr): { cells: CellAddr[]; ranges: RangeAddr[] } | undefined {
    const vertex = this.formulas.get(cellKey(addr.sheet, addr.row, addr.col));
    if (!vertex) return undefined;
    return {
      cells: vertex.cellPrecedents.map(parseCellKey),
      ranges: vertex.rangePrecedents,
    };
  }

  /**
   * Walk precedents to the cells that originate an error.
   *
   * This is the "which cell broke my formula" answer. Excel makes users chase
   * tracer arrows one hop at a time; because the graph already exists, naming
   * the root cause is a bounded search rather than a new subsystem.
   */
  findErrorRoots(addr: CellAddr, valueOf: (a: CellAddr) => Scalar, limit = 64): CellAddr[] {
    const roots: CellAddr[] = [];
    const seen = new Set<string>();
    const stack: CellAddr[] = [addr];

    while (stack.length > 0 && roots.length < limit) {
      const current = stack.pop()!;
      const key = cellKey(current.sheet, current.row, current.col);
      if (seen.has(key)) continue;
      seen.add(key);

      if (!isError(valueOf(current))) continue;

      const precedents = this.precedentsOf(current);
      if (!precedents) {
        // A non-formula cell holding an error is where the error was entered.
        roots.push(current);
        continue;
      }

      const erroring: CellAddr[] = [];
      for (const c of precedents.cells) {
        if (isError(valueOf(c))) erroring.push(c);
      }
      for (const r of precedents.ranges) {
        for (const c of cellsOfRange(r)) {
          if (isError(valueOf(c))) erroring.push(c);
        }
      }

      // No erroring precedent means this formula produced the error itself -
      // a division by zero, a bad argument - so it IS the root.
      if (erroring.length === 0) roots.push(current);
      else stack.push(...erroring);
    }
    return roots;
  }

  /**
   * Recalculate everything reachable from the seed cells.
   *
   * `compute` returns the new value for a cell; `getValue` reads the current
   * one. Returning the same value stops propagation, which is the single
   * largest saving on a real model.
   */
  recalculate(
    seeds: CellAddr[],
    compute: (addr: CellAddr) => Scalar,
    getValue: (addr: CellAddr) => Scalar,
    setValue: (addr: CellAddr, value: Scalar) => void,
  ): RecalcResult {
    const started = Date.now();
    const evaluated: CellAddr[] = [];
    const changed: CellAddr[] = [];

    const dirty = new Set<string>();
    const enqueue = (a: CellAddr) => {
      const k = cellKey(a.sheet, a.row, a.col);
      if (this.formulas.has(k)) dirty.add(k);
    };

    for (const seed of seeds) {
      enqueue(seed);
      for (const d of this.directDependents(seed.sheet, seed.row, seed.col)) enqueue(d);
    }
    // Volatile formulas recalculate on every pass by definition.
    for (const k of this.volatiles) dirty.add(k);

    // Expand to the full transitive closure before ordering, so the topological
    // sort sees every vertex it will need to place.
    const closure = new Set(dirty);
    const stack = [...dirty];
    while (stack.length > 0) {
      const key = stack.pop()!;
      const vertex = this.formulas.get(key);
      if (!vertex) continue;
      for (const d of this.directDependents(vertex.addr.sheet, vertex.addr.row, vertex.addr.col)) {
        const dk = cellKey(d.sheet, d.row, d.col);
        if (!closure.has(dk)) {
          closure.add(dk);
          stack.push(dk);
        }
      }
    }

    const { order, cycles } = this.topologicalOrder(closure);

    for (const key of order) {
      const vertex = this.formulas.get(key);
      if (!vertex) continue;
      const before = getValue(vertex.addr);
      const after = compute(vertex.addr);
      evaluated.push(vertex.addr);
      if (!sameValue(before, after)) {
        setValue(vertex.addr, after);
        changed.push(vertex.addr);
      }
    }

    const circular: CellAddr[] = [];
    for (const group of cycles) {
      for (const key of group) {
        const vertex = this.formulas.get(key);
        if (vertex) circular.push(vertex.addr);
      }
    }

    if (circular.length > 0 && this.options.iterative) {
      this.iterate(cycles, compute, getValue, setValue, evaluated, changed);
    } else {
      // Excel's default: a circular reference warns and the cells hold zero,
      // rather than showing an error value.
      for (const addr of circular) setValue(addr, 0);
    }

    return { evaluated, changed, circular, elapsedMs: Date.now() - started };
  }

  /** Excel's iterative-calculation mode for deliberate cycles. */
  private iterate(
    cycles: string[][],
    compute: (addr: CellAddr) => Scalar,
    getValue: (addr: CellAddr) => Scalar,
    setValue: (addr: CellAddr, value: Scalar) => void,
    evaluated: CellAddr[],
    changed: CellAddr[],
  ): void {
    const maxIterations = this.options.maxIterations ?? 100;
    const maxChange = this.options.maxChange ?? 0.001;

    for (const group of cycles) {
      const addrs = group.map((k) => this.formulas.get(k)?.addr).filter(Boolean) as CellAddr[];
      // Seed the cycle at zero so the first pass has something to read.
      for (const a of addrs) if (getValue(a) === null) setValue(a, 0);

      for (let i = 0; i < maxIterations; i++) {
        let delta = 0;
        for (const addr of addrs) {
          const before = getValue(addr);
          const after = compute(addr);
          evaluated.push(addr);
          if (!sameValue(before, after)) {
            if (typeof before === 'number' && typeof after === 'number') {
              delta = Math.max(delta, Math.abs(after - before));
            } else {
              delta = Number.POSITIVE_INFINITY;
            }
            setValue(addr, after);
            changed.push(addr);
          }
        }
        if (delta < maxChange) break;
      }
    }
  }

  /**
   * Iterative Tarjan: topological order plus the strongly-connected components
   * that cannot be ordered.
   *
   * Written with an explicit stack rather than recursion because dependency
   * chains in real workbooks run far deeper than the JavaScript stack allows,
   * and a stack overflow on open is indistinguishable to the user from a crash.
   */
  private topologicalOrder(keys: ReadonlySet<string>): { order: string[]; cycles: string[][] } {
    const index = new Map<string, number>();
    const lowlink = new Map<string, number>();
    const onStack = new Set<string>();
    const sccStack: string[] = [];
    const cycles: string[][] = [];
    const order: string[] = [];
    let counter = 0;

    // Each frame tracks which successor it is up to, so the walk resumes
    // exactly where recursion would have returned.
    interface Frame {
      key: string;
      successors: string[];
      next: number;
      /**
       * A cell that reads itself is a cycle, but Tarjan reports it as an
       * ordinary component of one. The self-edge has to be noted separately or
       * =A1+1 in A1 is silently treated as acyclic.
       */
      selfLoop: boolean;
    }

    const successorsOf = (key: string): string[] => {
      const vertex = this.formulas.get(key);
      if (!vertex) return [];
      const out: string[] = [];
      for (const d of this.directDependents(vertex.addr.sheet, vertex.addr.row, vertex.addr.col)) {
        const dk = cellKey(d.sheet, d.row, d.col);
        if (keys.has(dk)) out.push(dk);
      }
      return out;
    };

    for (const root of keys) {
      if (index.has(root)) continue;
      const rootSuccessors = successorsOf(root);
      const frames: Frame[] = [
        { key: root, successors: rootSuccessors, next: 0, selfLoop: rootSuccessors.includes(root) },
      ];
      index.set(root, counter);
      lowlink.set(root, counter);
      counter++;
      sccStack.push(root);
      onStack.add(root);

      while (frames.length > 0) {
        const frame = frames[frames.length - 1]!;
        if (frame.next < frame.successors.length) {
          const succ = frame.successors[frame.next++]!;
          if (!index.has(succ)) {
            index.set(succ, counter);
            lowlink.set(succ, counter);
            counter++;
            sccStack.push(succ);
            onStack.add(succ);
            const succSuccessors = successorsOf(succ);
            frames.push({
              key: succ,
              successors: succSuccessors,
              next: 0,
              selfLoop: succSuccessors.includes(succ),
            });
          } else if (onStack.has(succ)) {
            lowlink.set(frame.key, Math.min(lowlink.get(frame.key)!, index.get(succ)!));
          }
          continue;
        }

        frames.pop();
        const parent = frames[frames.length - 1];
        if (parent) {
          lowlink.set(parent.key, Math.min(lowlink.get(parent.key)!, lowlink.get(frame.key)!));
        }

        if (lowlink.get(frame.key) === index.get(frame.key)) {
          const component: string[] = [];
          for (;;) {
            const w = sccStack.pop()!;
            onStack.delete(w);
            component.push(w);
            if (w === frame.key) break;
          }
          if (component.length > 1 || frame.selfLoop) cycles.push(component);
          else order.push(component[0]!);
        }
      }
    }

    // Tarjan emits components in reverse topological order; dependents come out
    // before their precedents, so evaluation walks the list backwards.
    order.reverse();
    return { order, cycles };
  }
}

function parseCellKey(key: string): CellAddr {
  const bang = key.lastIndexOf('!');
  const sheet = key.slice(0, bang);
  const packed = Number(key.slice(bang + 1));
  return { sheet, row: Math.floor(packed / MAX_COLS), col: packed % MAX_COLS };
}

/** Bounded expansion of a range into addresses, for error tracing. */
function* cellsOfRange(r: RangeAddr, limit = 4096): Generator<CellAddr> {
  let n = 0;
  for (let row = r.startRow; row <= r.endRow; row++) {
    for (let col = r.startCol; col <= r.endCol; col++) {
      if (n++ >= limit) return;
      yield { sheet: r.sheet, row, col };
    }
  }
}

/**
 * Value equality for the propagation short-circuit.
 *
 * Errors compare by code, so a cell that stays #DIV/0! does not re-dirty its
 * dependents just because a fresh error object was allocated.
 */
function sameValue(a: Scalar, b: Scalar): boolean {
  if (isError(a) && isError(b)) return a.code === b.code;
  return Object.is(a, b);
}

export { sameValue };
