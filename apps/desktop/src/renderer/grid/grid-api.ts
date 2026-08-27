/**
 * The contract the shell expects from the canvas grid.
 *
 * The grid package is being written alongside this one, so the shell codes
 * against this interface rather than importing the package's types directly.
 * That is not just scheduling convenience: it keeps the boundary explicit, so a
 * change on either side shows up here as a compile error in one file instead of
 * scattered through twenty components.
 *
 * The interface is deliberately the smallest thing the shell needs: construct
 * with a canvas, a workbook and a sheet name; tell it where the viewport is and
 * how big; ask it to paint; ask it what is under a point; and hear about
 * selection changes. Everything else - freezing, merges, overflow - is the
 * grid's own business and the shell has no opinion about it.
 */

import type { Workbook } from '@mirrorz/core';
import type { Selection } from '../model/selection.js';

export interface GridHit {
  row: number;
  col: number;
  /** What part of the grid the point landed on. */
  region: 'cell' | 'rowHeader' | 'colHeader' | 'corner';
}

export interface GridViewLike {
  setScroll(x: number, y: number): void;
  resize(width: number, height: number): void;
  render(): void;
  hitTest(x: number, y: number): GridHit | null;
  /** Returns an unsubscribe function, as every other listener in this app does. */
  onSelectionChange(listener: (selection: Selection) => void): () => void;
  setSelection?(selection: Selection): void;
  setSheet?(name: string): void;
  destroy?(): void;
}

export interface GridViewConstructor {
  new (canvas: HTMLCanvasElement, workbook: Workbook, sheetName: string): GridViewLike;
}

/**
 * Load the grid package at runtime.
 *
 * The specifier goes through a variable so the bundler leaves it alone: the
 * package may not exist yet, and a static import of a missing module fails the
 * whole build rather than degrading to the fallback grid.
 */
export async function loadGridView(): Promise<GridViewConstructor | null> {
  const specifier = '@mirrorz/grid';
  try {
    const module = (await import(/* @vite-ignore */ specifier)) as Record<string, unknown>;
    const candidate = module['GridView'];
    return typeof candidate === 'function' ? (candidate as GridViewConstructor) : null;
  } catch {
    return null;
  }
}
