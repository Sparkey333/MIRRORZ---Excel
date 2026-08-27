/**
 * React bindings for the controller.
 *
 * `useSyncExternalStore` over the controller's snapshot is the whole of the
 * state management. It is enough because the controller already guarantees the
 * two things a store has to: the snapshot object identity changes exactly when
 * something changed, and nothing else in the app is allowed to mutate the model.
 */

import { createContext, useContext, useMemo, useSyncExternalStore, type ReactNode } from 'react';
import { Document, Workbook } from '@mirrorz/core';
import { Engine, createRegistry } from '@mirrorz/formula';
import { AppController, type AppSnapshot, type ControllerOptions } from './controller.js';

const ControllerContext = createContext<AppController | null>(null);

export function AppProvider({
  controller,
  children,
}: {
  controller: AppController;
  children: ReactNode;
}) {
  return <ControllerContext.Provider value={controller}>{children}</ControllerContext.Provider>;
}

export function useController(): AppController {
  const controller = useContext(ControllerContext);
  if (!controller) throw new Error('useController used outside AppProvider');
  return controller;
}

export function useApp(): AppSnapshot {
  const controller = useController();
  return useSyncExternalStore(controller.subscribe, controller.getSnapshot, controller.getSnapshot);
}

/**
 * Recompute a derived value only when the document has actually changed.
 *
 * The panels all derive something expensive - a history tree, a precedent walk,
 * a filtered sheet list - and the snapshot's version is the one number that says
 * whether recomputing is necessary.
 */
export function useDerived<T>(compute: () => T, extraDeps: unknown[] = []): T {
  const { version } = useApp();
  // The dependency list is intentionally the version plus whatever the caller
  // adds; `compute` is recreated every render and is not a useful dependency.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  return useMemo(compute, [version, ...extraDeps]);
}

/** Build a controller with a fresh workbook, for the app entry and for tests. */
export function createController(
  workbook = new Workbook(),
  options: ControllerOptions = {},
): AppController {
  const doc = new Document(workbook);
  const registry = createRegistry();
  const engine = new Engine(doc, registry);
  if (workbook.sheets.length === 0) {
    doc.addSheet('Sheet1', undefined, { label: 'New workbook', origin: 'system' });
  }
  engine.indexWorkbook();
  return new AppController(doc, engine, registry, options);
}
