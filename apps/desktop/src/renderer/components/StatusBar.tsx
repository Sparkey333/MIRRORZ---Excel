/**
 * The status bar.
 *
 * The selection aggregate is the most-used read-only feature of any spreadsheet,
 * so it goes first. Calculation mode and the last recalculation time go beside
 * it because manual calculation mode is a silent trap - a workbook can show
 * stale numbers for hours with nothing on screen to say so - and putting the
 * mode and the age of the last pass in permanent view is the cheapest possible
 * fix for it.
 */

import { a1 } from '@mirrorz/core';
import { useApp, useController, useDerived } from '../state/context.js';
import { describeSelection, formatStat, selectionStats } from '../model/selection.js';

export function StatusBar({ now }: { now?: number }) {
  const controller = useController();
  const snapshot = useApp();

  const stats = useDerived(() => selectionStats(controller.sheet(), snapshot.selection), [
    snapshot.selection,
  ]);
  const dimensions = useDerived(() => {
    const bounds = controller.sheet()?.bounds();
    if (!bounds) return 'empty';
    return `${bounds.maxRow + 1} rows x ${bounds.maxCol + 1} columns`;
  });

  const reference = now ?? Date.now();
  const recalc = snapshot.lastRecalc;

  return (
    <footer className="mz-status" role="status" aria-label="Status bar">
      <span className="mz-status-selection">{describeSelection(snapshot.selection, a1)}</span>
      <span className="mz-status-stat">Count {stats.count}</span>
      {stats.numericCount > 0 ? (
        <>
          <span className="mz-status-stat">Sum {formatStat(stats.sum)}</span>
          <span className="mz-status-stat">Average {formatStat(stats.average)}</span>
          <span className="mz-status-stat">Min {formatStat(stats.min)}</span>
          <span className="mz-status-stat">Max {formatStat(stats.max)}</span>
        </>
      ) : null}
      {stats.errorCode ? (
        <span className="mz-status-error">{stats.errorCode} in selection</span>
      ) : null}

      <span className="mz-status-spacer" />

      <span className="mz-status-dimensions">{dimensions}</span>
      <span className="mz-status-calc" data-mode={snapshot.calcMode}>
        {snapshot.calcMode === 'manual' ? 'Manual calculation' : 'Automatic calculation'}
      </span>
      <span className="mz-status-recalc">
        {recalc
          ? `Recalculated ${recalc.evaluated} cells in ${recalc.elapsedMs.toFixed(1)}ms, ${ageOf(recalc.at, reference)}`
          : 'Not yet recalculated'}
      </span>
      {snapshot.dirty ? <span className="mz-status-dirty">Unsaved changes</span> : null}
      {snapshot.message ? (
        <button type="button" className="mz-status-message" onClick={() => controller.setMessage(null)}>
          {snapshot.message}
        </button>
      ) : null}
    </footer>
  );
}

function ageOf(at: number, now: number): string {
  const seconds = Math.max(0, Math.round((now - at) / 1000));
  if (seconds < 5) return 'just now';
  if (seconds < 60) return `${seconds}s ago`;
  return `${Math.round(seconds / 60)}m ago`;
}
