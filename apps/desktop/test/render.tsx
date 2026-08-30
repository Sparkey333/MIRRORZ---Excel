/**
 * Shared harness for the component tests.
 *
 * Every component reaches the model through the controller context, so the
 * harness builds a real controller over a real Document and Engine rather than
 * mocking either. Mocking the document would test the mock; the whole claim
 * being made about this UI is that it reflects the command log faithfully.
 */

import { render, type RenderResult } from '@testing-library/react';
import type { ReactElement } from 'react';
import { AppProvider, createController } from '../src/renderer/state/context.js';
import type { AppController } from '../src/renderer/state/controller.js';

export interface Harness {
  controller: AppController;
  view: RenderResult;
}

export function renderWith(
  ui: (controller: AppController) => ReactElement,
  setup?: (controller: AppController) => void,
): Harness {
  const controller = createController(undefined, { now: () => 1_700_000_000_000 });
  setup?.(controller);
  const view = render(<AppProvider controller={controller}>{ui(controller)}</AppProvider>);
  return { controller, view };
}

/** Type a value into a cell through the controller, as the grid would. */
export function enter(controller: AppController, row: number, col: number, text: string): void {
  controller.commitEntry(text, { sheet: 'Sheet1', row, col });
}
