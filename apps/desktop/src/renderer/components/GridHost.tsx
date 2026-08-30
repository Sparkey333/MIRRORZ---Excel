/**
 * The grid's mounting point.
 *
 * The canvas grid lives in its own package; this component owns the canvas
 * element, tells the view about size, theme and scroll, and translates its
 * events into controller calls. It knows nothing about how cells are drawn, and
 * the grid knows nothing about panels or the command palette.
 *
 * Every path out of here that changes data goes through the controller, and so
 * through the Document's command log. The grid is given the Workbook to read and
 * is never asked to write to it.
 *
 * The DOM fallback below is not a stub. A canvas with no 2d context - a test
 * environment, a browser that has exhausted its contexts - would otherwise leave
 * an empty rectangle where the sheet should be, which is the worst of the
 * available failures because it looks like data loss.
 */

import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { a1, colToName } from '@mirrorz/core';
import { useApp, useController } from '../state/context.js';
import { GridBridge, canvasCanPaint, sameSelection } from '../grid/grid-api.js';
import { containsCell, singleCell } from '../model/selection.js';
import { resolveTheme } from '../model/theme.js';
import { useSystemPrefersDark } from '../state/useTheme.js';

export function GridHost() {
  const controller = useController();
  const snapshot = useApp();
  const prefersDark = useSystemPrefersDark();
  const dark = resolveTheme(snapshot.theme, prefersDark) === 'dark';

  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const bridgeRef = useRef<GridBridge | null>(null);
  const [available, setAvailable] = useState<boolean | null>(null);

  // Layout effect, not effect: the canvas has to be measured and painted before
  // the browser shows the frame, or the first paint is a blank grid that fills
  // in a tick later.
  useLayoutEffect(() => {
    const canvas = canvasRef.current;
    if (!canvasCanPaint(canvas) || !canvas) {
      setAvailable(false);
      return;
    }

    let bridge: GridBridge;
    try {
      const container = containerRef.current;
      bridge = new GridBridge(canvas, controller.workbook, controller.getSnapshot().activeSheet, {
        dark,
        width: container?.clientWidth ?? 0,
        height: container?.clientHeight ?? 0,
        dpr: typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1,
      });
    } catch {
      // A grid that cannot be constructed is a reason to show the fallback, not
      // a reason to show nothing.
      setAvailable(false);
      return;
    }

    bridgeRef.current = bridge;
    const stopSelection = bridge.onSelectionChange((selection) => {
      controller.setSelection(selection);
    });
    // Double click and F2 mean "edit this cell"; the formula bar is where the
    // editing happens, so activation moves the selection and focuses it.
    const stopActivate = bridge.onActivate((cell) => {
      controller.selectCell(cell.row, cell.col);
      const editor = document.querySelector<HTMLTextAreaElement>('.mz-formula-input');
      editor?.focus();
    });

    bridge.setSelection(controller.getSnapshot().selection);
    setAvailable(true);
    bridge.render();

    return () => {
      stopSelection();
      stopActivate();
      bridge.destroy();
      bridgeRef.current = null;
    };
  }, [controller, dark]);

  // The document changed: refresh the cached geometry, then repaint. Version is
  // the one number that says something actually changed.
  useEffect(() => {
    const bridge = bridgeRef.current;
    if (!bridge) return;
    bridge.setSheet(snapshot.activeSheet);
    bridge.refresh();
  }, [snapshot.activeSheet, snapshot.version, available]);

  // Selection moved somewhere else - a palette jump, an inspector root, the
  // keyboard map - and the grid has to follow it. The value comparison is what
  // stops the grid's own event echoing back into it as a second selection.
  useEffect(() => {
    const bridge = bridgeRef.current;
    if (!bridge) return;
    if (sameSelection(bridge.currentSelection(), snapshot.selection)) return;
    bridge.setSelection(snapshot.selection);
    bridge.render();
  }, [snapshot.selection]);

  useEffect(() => {
    const bridge = bridgeRef.current;
    if (!bridge) return;
    bridge.setTheme(dark);
    bridge.render();
  }, [dark, available]);

  useEffect(() => {
    const container = containerRef.current;
    if (!container || typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(() => {
      const bridge = bridgeRef.current;
      if (!bridge) return;
      bridge.resize(
        container.clientWidth,
        container.clientHeight,
        typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1,
      );
      bridge.render();
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
          const bridge = bridgeRef.current;
          if (!bridge) return;
          const rect = e.currentTarget.getBoundingClientRect();
          e.currentTarget.setPointerCapture?.(e.pointerId);
          bridge.pointerDown(e.clientX - rect.left, e.clientY - rect.top, {
            shift: e.shiftKey,
            ctrl: e.ctrlKey,
            meta: e.metaKey,
          });
          bridge.render();
        }}
        onPointerMove={(e) => {
          const bridge = bridgeRef.current;
          if (!bridge || e.buttons === 0) return;
          const rect = e.currentTarget.getBoundingClientRect();
          bridge.pointerMove(e.clientX - rect.left, e.clientY - rect.top);
          bridge.render();
        }}
        onPointerUp={(e) => {
          e.currentTarget.releasePointerCapture?.(e.pointerId);
          bridgeRef.current?.pointerUp();
        }}
        onDoubleClick={(e) => {
          const bridge = bridgeRef.current;
          if (!bridge) return;
          const rect = e.currentTarget.getBoundingClientRect();
          bridge.doubleClick(e.clientX - rect.left, e.clientY - rect.top);
        }}
        onWheel={(e) => {
          const bridge = bridgeRef.current;
          if (!bridge) return;
          // Pixel deltas straight through: the grid carries sub-pixel scroll
          // offsets, and rounding to a row here would throw that away.
          bridge.scrollBy(e.deltaX, e.deltaY);
          bridge.render();
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
