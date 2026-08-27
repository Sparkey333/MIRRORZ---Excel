/**
 * A SheetStore backed by the core Workbook model.
 *
 * This is the adapter that lets the engine, which knows nothing about our sheet
 * representation, run against it. Keeping it in one small file is what makes the
 * same engine usable from a worker thread or against a different store later.
 */

import { type Scalar, type Workbook, packKey } from '@mirrorz/core';
import type { SheetStore } from './evaluator.js';

export class WorkbookStore implements SheetStore {
  constructor(private readonly workbook: Workbook) {}

  getScalar(sheet: string, row: number, col: number): Scalar {
    return this.workbook.getSheet(sheet)?.getValue(row, col) ?? null;
  }

  *iterate(
    sheet: string,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number,
  ): Iterable<{ row: number; col: number; value: Scalar }> {
    const s = this.workbook.getSheet(sheet);
    if (!s) return;
    const area = (endRow - startRow + 1) * (endCol - startCol + 1);
    if (area <= s.cellCount) {
      // A small window over a dense sheet: walking the window is cheaper.
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const cell = s.cells.get(packKey(r, c));
          if (cell && cell.value !== null) yield { row: r, col: c, value: cell.value };
        }
      }
      return;
    }
    // A large window over a sparse sheet: walk the cells instead, so summing a
    // whole column costs the number of cells that exist, not 1,048,576.
    for (const e of s.entries()) {
      if (e.row >= startRow && e.row <= endRow && e.col >= startCol && e.col <= endCol) {
        if (e.cell.value !== null) yield { row: e.row, col: e.col, value: e.cell.value };
      }
    }
  }

  hasSheet(name: string): boolean {
    return this.workbook.getSheet(name) !== undefined;
  }

  sheetNames(): readonly string[] {
    return this.workbook.sheets.map((s) => s.name);
  }

  getDefinedName(name: string, sheet: string): string | undefined {
    const lower = name.toLowerCase();
    const sheetIndex = this.workbook.sheets.findIndex(
      (s) => s.name.toLowerCase() === sheet.toLowerCase(),
    );
    // A sheet-scoped name shadows a workbook-scoped one of the same name.
    const scoped = this.workbook.definedNames.find(
      (d) => d.name.toLowerCase() === lower && d.scope === sheetIndex,
    );
    const found =
      scoped ?? this.workbook.definedNames.find((d) => d.name.toLowerCase() === lower && d.scope === undefined);
    if (!found) return undefined;
    // Stored names carry a leading '=' in some producers.
    return found.refersTo.startsWith('=') ? found.refersTo.slice(1) : found.refersTo;
  }

  usedBounds(sheet: string): { maxRow: number; maxCol: number } | null {
    const b = this.workbook.getSheet(sheet)?.bounds();
    return b ? { maxRow: b.maxRow, maxCol: b.maxCol } : null;
  }
}
