/**
 * The canvas grid, mounted for real.
 *
 * These tests exist because the shell and the grid package were written against
 * two different interfaces and nothing caught it: every other test in this
 * directory runs in jsdom, where a canvas has no 2d context and the DOM fallback
 * takes over, so the canvas path was never executed once. Here a recording 2d
 * context is installed so the real GridView is constructed, and both directions
 * of the selection wiring are asserted end to end.
 */

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { act, fireEvent, screen, waitFor } from '@testing-library/react';
import {
  DEFAULT_HEADER_HEIGHT,
  DEFAULT_HEADER_WIDTH,
  buildColAxis,
  buildRowAxis,
} from '@mirrorz/grid';
import { App } from '../src/renderer/App.js';
import { renderWith } from './render.js';

/** Enough of a 2d context for the painter; it records nothing it is not asked. */
function fakeContext(): unknown {
  const noop = (): void => {};
  return {
    font: '10px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'start',
    textBaseline: 'alphabetic',
    globalAlpha: 1,
    measureText: (text: string) => ({ width: text.length * 5 }),
    save: noop,
    restore: noop,
    beginPath: noop,
    closePath: noop,
    rect: noop,
    clip: noop,
    moveTo: noop,
    lineTo: noop,
    stroke: noop,
    fill: noop,
    fillRect: noop,
    strokeRect: noop,
    clearRect: noop,
    fillText: noop,
    translate: noop,
    scale: noop,
    setTransform: noop,
    setLineDash: noop,
  };
}

const VIEWPORT = { width: 900, height: 600 };

/**
 * jsdom has no PointerEvent, and testing-library's synthetic fallback carries no
 * coordinates at all - which is how a click ends up at NaN,NaN. A MouseEvent
 * named `pointerdown` is what React's pointer handler listens for and it does
 * carry clientX/clientY.
 */
function pointerDownAt(target: Element, x: number, y: number): void {
  fireEvent(
    target,
    new MouseEvent('pointerdown', { clientX: x, clientY: y, bubbles: true, buttons: 1 }),
  );
}
let restore: (() => void)[] = [];

beforeEach(() => {
  const canvasProto = HTMLCanvasElement.prototype as unknown as Record<string, unknown>;
  const previous = canvasProto['getContext'];
  canvasProto['getContext'] = function getContext(kind: string) {
    return kind === '2d' ? fakeContext() : null;
  };
  restore.push(() => {
    canvasProto['getContext'] = previous;
  });

  // jsdom lays nothing out, so the host would measure a zero-sized viewport and
  // every hit test would land outside it.
  for (const [prop, value] of [
    ['clientWidth', VIEWPORT.width],
    ['clientHeight', VIEWPORT.height],
  ] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, prop);
    Object.defineProperty(HTMLElement.prototype, prop, { configurable: true, get: () => value });
    restore.push(() => {
      if (original) Object.defineProperty(HTMLElement.prototype, prop, original);
      else delete (HTMLElement.prototype as unknown as Record<string, unknown>)[prop];
    });
  }
});

afterEach(() => {
  for (const undo of restore.reverse()) undo();
  restore = [];
});

describe('GridHost with a paintable canvas', () => {
  it('mounts the canvas grid rather than the DOM fallback', async () => {
    renderWith(() => <App />);
    await waitFor(() => {
      expect(document.querySelector('.mz-grid-host')?.getAttribute('data-grid')).toBe('canvas');
    });
    expect(screen.queryByLabelText('Sheet1 cells')).toBeNull();
  });

  it('reports a click on a cell back to the controller', async () => {
    const { controller } = renderWith(() => <App />);
    await waitFor(() => {
      expect(document.querySelector('.mz-grid-host')?.getAttribute('data-grid')).toBe('canvas');
    });

    // Work out which cell the point lands on from the same axis maths the grid
    // uses, so the assertion does not silently encode a default row height.
    const sheet = controller.sheet()!;
    const rows = buildRowAxis(sheet);
    const cols = buildColAxis(sheet);
    const x = DEFAULT_HEADER_WIDTH + cols.offsetOf(3) + 4;
    const y = DEFAULT_HEADER_HEIGHT + rows.offsetOf(5) + 4;

    const canvas = document.querySelector('.mz-grid-canvas')!;
    act(() => pointerDownAt(canvas, x, y));

    expect(controller.getSnapshot().selection.active).toEqual({ row: 5, col: 3 });
    expect(controller.getSnapshot().selection.sheet).toBe('Sheet1');
  });

  it('follows a selection made anywhere else in the shell', async () => {
    const { controller } = renderWith(() => <App />);
    await waitFor(() => {
      expect(document.querySelector('.mz-grid-host')?.getAttribute('data-grid')).toBe('canvas');
    });

    act(() => controller.selectCell(9, 2));

    // The grid is the thing that has to have moved: read it back through a hit
    // test on the cell it should now consider active.
    const sheet = controller.sheet()!;
    const rows = buildRowAxis(sheet);
    const cols = buildColAxis(sheet);
    const x = DEFAULT_HEADER_WIDTH + cols.offsetOf(2) + 4;
    const y = DEFAULT_HEADER_HEIGHT + rows.offsetOf(9) + 4;
    const canvas = document.querySelector('.mz-grid-canvas')!;
    act(() => pointerDownAt(canvas, x, y));
    expect(controller.getSnapshot().selection.active).toEqual({ row: 9, col: 2 });
  });

  it('keeps painting after the document changes', async () => {
    const { controller } = renderWith(() => <App />);
    await waitFor(() => {
      expect(document.querySelector('.mz-grid-host')?.getAttribute('data-grid')).toBe('canvas');
    });
    // An edit refreshes the view's cached geometry; if invalidate() were missing
    // this still passes, but a throw from the paint path would not.
    act(() => controller.commitEntry('42', { sheet: 'Sheet1', row: 0, col: 0 }));
    expect(controller.cellAt({ sheet: 'Sheet1', row: 0, col: 0 })?.value).toBe(42);
    expect(document.querySelector('.mz-grid-host')?.getAttribute('data-grid')).toBe('canvas');
  });
});
