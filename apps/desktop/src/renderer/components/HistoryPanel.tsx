/**
 * The undo history panel.
 *
 * Excel has nothing like this, and the reason is structural rather than a
 * missing feature: its undo is a stack of UI actions that macros clear and
 * closing the file discards, so there is nothing to draw. Here the history is a
 * tree of labelled, origin-tagged, timestamped entries that outlives the
 * session, so drawing it is mostly a matter of not getting in its way.
 *
 * Two things this panel does that a linear list cannot. It shows abandoned
 * branches, so undoing five steps and trying something else does not lose the
 * five steps - they are still there, indented, one click away. And it labels the
 * origin, so a row that says "macro" tells you that the thing that changed your
 * numbers was not you.
 */

import { useMemo, useState } from 'react';
import { useApp, useController, useDerived } from '../state/context.js';
import {
  ORIGIN_LABELS,
  flattenHistory,
  relativeTime,
  summariseHistory,
  type HistoryRow,
} from '../model/history-tree.js';

export interface HistoryPanelProps {
  /** Injected so the relative times are stable in tests. */
  now?: number;
}

export function HistoryPanel({ now }: HistoryPanelProps) {
  const controller = useController();
  const snapshot = useApp();
  const [showRecalc, setShowRecalc] = useState(false);

  const entries = useDerived(() => controller.historyEntries());
  const rows = useMemo(() => flattenHistory(entries, snapshot.headId), [entries, snapshot.headId]);
  const summary = useMemo(() => summariseHistory(entries, snapshot.headId), [entries, snapshot.headId]);

  // Recalculation entries are real history - they are how a value got there -
  // but they double the row count, so they are collapsed by default.
  const visible = showRecalc ? rows : rows.filter((row) => row.origin !== 'recalc');
  const reference = now ?? Date.now();

  return (
    <aside className="mz-history" aria-label="History">
      <header className="mz-panel-header">
        <h2>History</h2>
        <button type="button" aria-label="Close history" onClick={() => controller.togglePanel('history', false)}>
          &#215;
        </button>
      </header>
      <div className="mz-history-summary">
        <span>{summary.total} entries</span>
        {summary.abandoned > 0 ? (
          <span className="mz-history-abandoned">
            {summary.abandoned} off the current path in {summary.branches} branch
            {summary.branches === 1 ? '' : 'es'}
          </span>
        ) : null}
      </div>
      <label className="mz-history-toggle">
        <input type="checkbox" checked={showRecalc} onChange={(e) => setShowRecalc(e.target.checked)} />
        Show recalculation steps
      </label>
      <ol className="mz-history-list">
        <li className="mz-history-row" data-depth={0}>
          <button
            type="button"
            className={snapshot.headId === null ? 'mz-history-entry mz-head' : 'mz-history-entry'}
            aria-current={snapshot.headId === null}
            onClick={() => controller.jumpTo(null)}
          >
            <span className="mz-history-label">Opened</span>
            <span className="mz-history-origin">start</span>
          </button>
        </li>
        {visible.map((row) => (
          <HistoryRowView
            key={row.id}
            row={row}
            now={reference}
            onJump={() => controller.jumpTo(row.id)}
          />
        ))}
      </ol>
      {visible.length === 0 ? <p className="mz-history-empty">Nothing has been changed yet.</p> : null}
    </aside>
  );
}

function HistoryRowView({ row, now, onJump }: { row: HistoryRow; now: number; onJump: () => void }) {
  const classes = ['mz-history-entry'];
  if (row.isHead) classes.push('mz-head');
  if (!row.onPath) classes.push('mz-abandoned');
  if (row.barrier) classes.push('mz-barrier');

  return (
    <li className="mz-history-row" data-depth={row.depth} style={{ paddingLeft: 8 + row.depth * 14 }}>
      {row.startsBranch ? (
        <span className="mz-history-branch-marker" aria-hidden="true">
          &#8627;
        </span>
      ) : null}
      <button
        type="button"
        className={classes.join(' ')}
        aria-current={row.isHead}
        data-origin={row.origin}
        data-abandoned={!row.onPath}
        onClick={onJump}
        title={row.barrier ? 'This step cannot be reversed exactly' : undefined}
      >
        <span className="mz-history-label">{row.label}</span>
        <span className="mz-history-origin">{ORIGIN_LABELS[row.origin]}</span>
        <span className="mz-history-time">{relativeTime(row.timestamp, now)}</span>
        {row.changeCount > 1 ? (
          <span className="mz-history-count">{row.changeCount} changes</span>
        ) : null}
        {row.isForkPoint ? <span className="mz-history-fork">branches here</span> : null}
        {row.barrier ? <span className="mz-history-barrier-flag">not reversible</span> : null}
      </button>
    </li>
  );
}
