/**
 * The calculation engine: the piece that makes edits propagate.
 *
 * It joins the four parts that already exist - the parser, the evaluator, the
 * dependency graph and the document's command log - and owns the one thing none
 * of them can do alone: deciding what to recompute when a cell changes, and
 * writing the results back without those writes looking like user edits.
 *
 * Two details are worth stating up front.
 *
 * Parsed formulas are cached by their normalised text. A column of fifty
 * thousand copies of the same formula differs only in the origin it is rendered
 * at, and because references are stored relative to that origin, all fifty
 * thousand share one AST. Parsing is otherwise the dominant cost of opening a
 * large workbook.
 *
 * Recalculation writes through the document inside a single transaction tagged
 * as machine-originated. That keeps the undo history readable - undoing an edit
 * takes back the edit and its consequences together, rather than requiring the
 * user to step back through every recomputed cell.
 */

import {
  type CellData,
  type Document,
  type Scalar,
  type Workbook,
  isError,
} from '@mirrorz/core';
import { type Ast, Node, walk } from './ast.js';
import { Evaluator, type EvalOptions } from './evaluator.js';
import { type CellAddr, DependencyGraph, type RangeAddr, type RecalcResult } from './graph.js';
import { ParseError, parseFormula } from './parser.js';
import type { FunctionRegistry } from './registry.js';
import { WorkbookStore } from './store.js';
import { CellError } from '@mirrorz/core';

export interface EngineOptions extends EvalOptions {
  /** Recalculate automatically after every edit. Off mirrors Excel's manual mode. */
  autoCalculate?: boolean;
  iterative?: boolean;
  maxIterations?: number;
  maxChange?: number;
}

export class Engine {
  readonly graph: DependencyGraph;
  private readonly store: WorkbookStore;
  private readonly evaluator: Evaluator;
  /** Normalised formula text -> parsed AST, shared across every copy. */
  private readonly astCache = new Map<string, Ast | ParseError>();
  private readonly options: EngineOptions;

  constructor(
    private readonly doc: Document,
    registry: FunctionRegistry,
    options: EngineOptions = {},
  ) {
    this.options = options;
    this.store = new WorkbookStore(doc.workbook);
    this.graph = new DependencyGraph({
      ...(options.iterative === undefined ? {} : { iterative: options.iterative }),
      ...(options.maxIterations === undefined ? {} : { maxIterations: options.maxIterations }),
      ...(options.maxChange === undefined ? {} : { maxChange: options.maxChange }),
    });
    this.evaluator = new Evaluator(this.store, registry, {
      dateSystem: doc.workbook.dateSystem,
      ...options,
    });
  }

  get workbook(): Workbook {
    return this.doc.workbook;
  }

  /**
   * Register every formula in the workbook. Called once after opening a file,
   * before the first recalculation.
   */
  indexWorkbook(): void {
    for (const sheet of this.doc.workbook.sheets) {
      for (const { row, col, cell } of sheet.entries()) {
        if (cell.formula !== undefined) {
          this.registerFormula({ sheet: sheet.name, row, col }, cell.formula);
        }
      }
    }
  }

  /** Parse a formula, sharing the AST with every identical formula. */
  parse(formula: string, origin: { row: number; col: number }): Ast | ParseError {
    // The cache key includes the origin because references are stored relative
    // to it, so two textually identical formulas at different cells legitimately
    // produce different trees.
    const key = `${origin.row}:${origin.col}:${formula}`;
    const cached = this.astCache.get(key);
    if (cached) return cached;
    let result: Ast | ParseError;
    try {
      result = parseFormula(formula, { origin });
    } catch (err) {
      result = err instanceof ParseError ? err : new ParseError(String(err), 0);
    }
    this.astCache.set(key, result);
    return result;
  }

  /** Tell the graph what a formula reads. */
  registerFormula(addr: CellAddr, formula: string): void {
    const ast = this.parse(formula, addr);
    if (ast instanceof ParseError) {
      // An unparseable formula depends on nothing and evaluates to #NAME?.
      this.graph.setFormula(addr, { cells: [], ranges: [] });
      return;
    }
    const precedents = extractPrecedents(ast, addr);
    this.graph.setFormula(addr, precedents, containsVolatile(ast));
  }

  /**
   * Set a cell's contents from user input that has already been interpreted.
   * Returns the recalculation result so a caller can report timing.
   */
  setCell(addr: CellAddr, data: CellData | undefined, label?: string): RecalcResult | undefined {
    this.doc.transact({ label: label ?? `Edit ${addr.sheet}`, origin: 'user' }, () => {
      this.doc.setCell(addr.sheet, addr.row, addr.col, data);
      if (data?.formula !== undefined) this.registerFormula(addr, data.formula);
      else this.graph.removeFormula(addr);
    });
    return this.options.autoCalculate === false ? undefined : this.recalculate([addr]);
  }

  /** Evaluate one cell without touching the document. Used by the inspector. */
  evaluateAt(addr: CellAddr): Scalar {
    const cell = this.doc.workbook.getSheet(addr.sheet)?.getCell(addr.row, addr.col);
    if (!cell?.formula) return cell?.value ?? null;
    const ast = this.parse(cell.formula, addr);
    if (ast instanceof ParseError) return new CellError('#NAME?', ast.message);
    return this.evaluator.evaluateScalar({
      ast,
      sheet: addr.sheet,
      row: addr.row,
      col: addr.col,
    });
  }

  /**
   * Recalculate everything affected by the seed cells.
   *
   * The writes go through the document as one machine-originated transaction, so
   * the undo history shows the user's edit rather than the hundred cells it
   * happened to change.
   */
  recalculate(seeds: CellAddr[] = []): RecalcResult {
    let result!: RecalcResult;
    this.doc.transact({ label: 'Recalculate', origin: 'recalc' }, () => {
      result = this.graph.recalculate(
        seeds,
        (addr) => this.evaluateAt(addr),
        (addr) => this.doc.workbook.getSheet(addr.sheet)?.getValue(addr.row, addr.col) ?? null,
        (addr, value) => {
          const sheet = this.doc.workbook.getSheet(addr.sheet);
          const existing = sheet?.getCell(addr.row, addr.col);
          if (!existing) return;
          this.doc.setCell(addr.sheet, addr.row, addr.col, { ...existing, value });
        },
      );
    });
    return result;
  }

  /** Recalculate everything, as Excel's Ctrl+Alt+F9 does. */
  recalculateAll(): RecalcResult {
    const seeds: CellAddr[] = [];
    for (const sheet of this.doc.workbook.sheets) {
      for (const { row, col, cell } of sheet.entries()) {
        if (cell.formula !== undefined) seeds.push({ sheet: sheet.name, row, col });
      }
    }
    return this.recalculate(seeds);
  }

  /**
   * Explain a cell: its formula, its precedents, and - when it holds an error -
   * the cells that actually originated it.
   *
   * Excel makes users chase tracer arrows one hop at a time. The graph already
   * knows the answer, so naming it is a lookup rather than a feature.
   */
  explain(addr: CellAddr): CellExplanation {
    const cell = this.doc.workbook.getSheet(addr.sheet)?.getCell(addr.row, addr.col);
    const value = cell?.value ?? null;
    const precedents = this.graph.precedentsOf(addr);
    const explanation: CellExplanation = {
      addr,
      value,
      dependents: this.graph.directDependents(addr.sheet, addr.row, addr.col),
      precedentCells: precedents?.cells ?? [],
      precedentRanges: precedents?.ranges ?? [],
      errorRoots: [],
    };
    if (cell?.formula !== undefined) explanation.formula = cell.formula;
    if (isError(value)) {
      explanation.errorRoots = this.graph.findErrorRoots(
        addr,
        (a) => this.doc.workbook.getSheet(a.sheet)?.getValue(a.row, a.col) ?? null,
      );
    }
    return explanation;
  }
}

export interface CellExplanation {
  addr: CellAddr;
  value: Scalar;
  formula?: string;
  precedentCells: CellAddr[];
  precedentRanges: RangeAddr[];
  dependents: CellAddr[];
  /** Cells that originated the error, when this cell holds one. */
  errorRoots: CellAddr[];
}

/**
 * Walk an AST for everything it reads.
 *
 * Ranges are collected separately from single cells so the graph can register
 * one range vertex rather than an edge per cell.
 */
export function extractPrecedents(
  ast: Ast,
  origin: CellAddr,
): { cells: CellAddr[]; ranges: RangeAddr[] } {
  const cells: CellAddr[] = [];
  const ranges: RangeAddr[] = [];

  walk(ast, (node) => {
    switch (node.kind) {
      case Node.Ref: {
        // A reference nested inside a Range node is handled by that node, so
        // only free-standing references are collected here.
        const row = node.rowAbs ? node.row : node.row + origin.row;
        const col = node.colAbs ? node.col : node.col + origin.col;
        if (row >= 0 && col >= 0) {
          cells.push({ sheet: node.sheet ?? origin.sheet, row, col });
        }
        break;
      }
      case Node.Range: {
        const startRow = node.start.rowAbs ? node.start.row : node.start.row + origin.row;
        const startCol = node.start.colAbs ? node.start.col : node.start.col + origin.col;
        const endRow = node.end.rowAbs ? node.end.row : node.end.row + origin.row;
        const endCol = node.end.colAbs ? node.end.col : node.end.col + origin.col;
        if (startRow >= 0 && startCol >= 0 && endRow >= 0 && endCol >= 0) {
          ranges.push({
            sheet: node.start.sheet ?? origin.sheet,
            startRow: Math.min(startRow, endRow),
            startCol: Math.min(startCol, endCol),
            endRow: Math.max(startRow, endRow),
            endCol: Math.max(startCol, endCol),
          });
        }
        break;
      }
      case Node.Beam: {
        const sheet = node.sheet ?? origin.sheet;
        if (node.axis === 'col') {
          const from = node.fromAbs ? node.from : node.from + origin.col;
          const to = node.toAbs ? node.to : node.to + origin.col;
          ranges.push({
            sheet,
            startRow: 0,
            startCol: Math.min(from, to),
            endRow: 1_048_575,
            endCol: Math.max(from, to),
          });
        } else {
          const from = node.fromAbs ? node.from : node.from + origin.row;
          const to = node.toAbs ? node.to : node.to + origin.row;
          ranges.push({
            sheet,
            startRow: Math.min(from, to),
            startCol: 0,
            endRow: Math.max(from, to),
            endCol: 16_383,
          });
        }
        break;
      }
      default:
        break;
    }
  });

  // A Range node's own endpoints were also visited as Ref nodes; drop the
  // duplicates so a range does not additionally register two cell edges.
  const covered = (c: CellAddr) =>
    ranges.some(
      (r) =>
        r.sheet === c.sheet &&
        c.row >= r.startRow &&
        c.row <= r.endRow &&
        c.col >= r.startCol &&
        c.col <= r.endCol,
    );

  return { cells: cells.filter((c) => !covered(c)), ranges };
}

/**
 * Does this formula contain a function whose value can change without any of
 * its precedents changing?
 *
 * The list is deliberately short. Over-marking is not harmless conservatism: a
 * volatile formula drags its entire dependent closure into every recalculation,
 * and the widely repeated claim that INDEX, ROWS, COLUMNS and AREAS are volatile
 * is simply wrong.
 */
const VOLATILE_FUNCTIONS = new Set([
  'NOW',
  'TODAY',
  'RAND',
  'RANDBETWEEN',
  'RANDARRAY',
  'OFFSET',
  'INDIRECT',
  'INFO',
  'CELL',
]);

export function containsVolatile(ast: Ast): boolean {
  let volatileFound = false;
  walk(ast, (node) => {
    if (node.kind === Node.Call && VOLATILE_FUNCTIONS.has(node.name.toUpperCase())) {
      volatileFound = true;
    }
  });
  return volatileFound;
}
