/**
 * The grid's mounting point.
 *
 * The canvas grid lives in its own package; this component owns the canvas
 * element, tells the view about size and scroll, and translates its selection
 * events into controller calls. It knows nothing about how cells are drawn, and
 * the grid knows nothing about panels or the command palette.
 *
 * When the grid package is not present the fallback below renders a small DOM
 * grid instead. That is not a stub for its own sake: without it the whole shell
 * would be untestable and undemonstrable until another package lands, and a
 * shell you cannot run is a shell you cannot review.
 */

import { useEffect, useRef, useState } from 'react';
import { a1, colToName } from '@mirrorz/core';
import { useApp, useController } from '../state/context.js';
import { loadGridView, type GridViewLike } from '../grid/grid-api.js';
import { containsCell, singleCell } from '../model/selection.js';

export function GridHost() {
  const controller = useController();
  const snapshot = useApp();
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<GridViewLike | null>(null);
  const [available, setAvailable] = useState<boolean | null>(null);

  useEffect(() => {
    let disposed = false;
    let unsubscribe: (() => void) | undefined;

    void loadGridView().then((GridView) => {
      if (disposed) return;
      const canvas = canvasRef.current;
      if (!GridView || !canvas) {
        setAvailable(false);
        return;
      }
      const view = new GridView(canvas, controller.workbook, snapshot.activeSheet);
      viewRef.current = view;
      unsubscribe = view.onSelectionChange((selection) => controller.setSelection(selection));
      setAvailable(true);
      view.render();
    });

    return () => {
      disposed = true;
      unsubscribe?.();
      viewRef.current?.destroy?.();
      viewRef.current = null;
    };
    // The view is constructed once; sheet changes go through setSheet below.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [controller]);

  useEffect(() => {
    viewRef.current?.setSheet?.(snapshot.activeSheet);
    viewRef.current?.render();
  }, [snapshot.activeSheet, snapshot.version]);

  useEffect(() => {
    const container = containerRef.current;
    const view = viewRef.current;
    if (!container || !view || typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(() => {
      view.resize(container.clientWidth, container.clientHeight);
      view.render();
    });
    observer.observe(container);
    return () => observer.disconnect();
  }, [available]);

  return (
    <div className="mz-grid-host" ref={containerRef} data-grid={available ? 'canvas' : 'fallback'}>
      <canvas
        ref={canvasRef}
        className="mz-grid-canvas"
        hidden={available !== true}
        onPointerDown={(e) => {
          const view = viewRef.current;
          if (!view) return;
          const rect = e.currentTarget.getBoundingClientRect();
          const hit = view.hitTest(e.clientX - rect.left, e.clientY - rect.top);
          if (hit?.region === 'cell') controller.selectCell(hit.row, hit.col);
        }}
      />
      {available === false ? <FallbackGrid /> : null}
    </div>
  );
}

const FALLBACK_ROWS = 40;
const FALLBACK_COLS = 16;

/** A plain DOM grid, enough to drive and demonstrate the shell. */
function FallbackGrid() {
  const controller = useController();
  const snapshot = useApp();
  const sheet = controller.sheet();

  return (
    <table className="mz-fallback-grid" aria-label={`${snapshot.activeSheet} cells`}>
      <thead>
        <tr>
          <th className="mz-fallback-corner" />
          {Array.from({ length: FALLBACK_COLS }, (_, c) => (
            <th key={c} scope="col">
              {colToName(c)}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {Array.from({ length: FALLBACK_ROWS }, (_, r) => (
          <tr key={r} hidden={sheet?.isRowHidden(r)}>
            <th scope="row">{r + 1}</th>
            {Array.from({ length: FALLBACK_COLS }, (_, c) => {
              const selected = containsCell(snapshot.selection, r, c);
              const active = snapshot.selection.active.row === r && snapshot.selection.active.col === c;
              return (
                <td
                  key={c}
                  className={selected ? 'mz-selected-cell' : undefined}
                  aria-selected={active}
                  data-address={a1(r, c)}
                  onMouseDown={(e) => {
                    if (e.shiftKey) {
                      const anchor = snapshot.selection.active;
                      controller.setSelection({
                        sheet: snapshot.activeSheet,
                        active: anchor,
                        ranges: [{ start: anchor, end: { row: r, col: c } }],
                      });
                    } else {
                      controller.setSelection(singleCell(snapshot.activeSheet, r, c));
                    }
                  }}
                >
                  {controller.displayText({ sheet: snapshot.activeSheet, row: r, col: c })}
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
