/**
 * The import review, shown BEFORE anything is written.
 *
 * This is the dialogue that stops the classic silent corruptions: gene symbols
 * turning into dates, leading zeros vanishing from postcodes and part numbers,
 * nineteen-digit order ids being rounded into scientific notation. Every one of
 * those is irreversible once committed and invisible when it happens, which is
 * why the check belongs before the paste rather than in an undo afterwards.
 *
 * The per-column override is what makes it a decision rather than a warning. A
 * column that the inference read as dates can be locked to text in one click,
 * and the counts update to show what that would do.
 */

import { useMemo } from 'react';
import { a1 } from '@mirrorz/core';
import { useApp, useController } from '../state/context.js';
import {
  buildImportPlan,
  resolveImport,
  type ColumnOverride,
  type ColumnPlan,
} from '../model/import-review.js';

const OVERRIDES: { value: ColumnOverride; label: string }[] = [
  { value: 'auto', label: 'Detect' },
  { value: 'text', label: 'Text' },
  { value: 'number', label: 'Number' },
  { value: 'date', label: 'Date' },
];

const PREVIEW_ROWS = 6;

export function ImportReviewDialog() {
  const controller = useController();
  const snapshot = useApp();
  const pending = snapshot.pendingImport;

  const plan = useMemo(
    () =>
      pending
        ? buildImportPlan(pending.rows, controller.getEntryOptions(), pending.overrides, pending.headerRow)
        : null,
    [pending, controller],
  );

  if (!pending || !plan) return null;

  const previewRows = pending.rows.slice(0, PREVIEW_ROWS);

  return (
    <div className="mz-modal-backdrop">
      <div className="mz-modal mz-import" role="dialog" aria-modal="true" aria-label="Review import">
        <header className="mz-modal-header">
          <h2>
            Review {pending.source === 'csv' ? `import of ${pending.fileName ?? 'file'}` : 'paste'}
          </h2>
          <p className="mz-import-headline">{plan.headline}</p>
          <p className="mz-import-target">
            Landing at {pending.anchor.sheet}!{a1(pending.anchor.row, pending.anchor.col)}, {plan.review.total}{' '}
            non-empty cells
          </p>
        </header>

        <label className="mz-import-header-toggle">
          <input
            type="checkbox"
            checked={pending.headerRow}
            onChange={(e) => controller.setImportHeaderRow(e.target.checked)}
          />
          First row is a header
        </label>

        <div className="mz-import-table-wrap">
          <table className="mz-import-table">
            <thead>
              <tr>
                {plan.columns.map((column) => (
                  <th key={column.index} scope="col">
                    <div className="mz-import-col-name">{column.name}</div>
                    <select
                      aria-label={`Type for column ${column.name}`}
                      value={column.override}
                      onChange={(e) =>
                        controller.setImportOverride(column.index, e.target.value as ColumnOverride)
                      }
                    >
                      {OVERRIDES.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                    <ColumnSummary column={column} />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {previewRows.map((row, r) => (
                <tr key={r}>
                  {plan.columns.map((column) => (
                    <td key={column.index}>{row[column.index] ?? ''}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {plan.review.protected.length > 0 ? (
          <section className="mz-import-protected">
            <h3>Kept as text</h3>
            <ul>
              {plan.review.protected.slice(0, 8).map((issue) => (
                <li key={`${issue.row}:${issue.col}`}>
                  <code>{issue.input}</code> {issue.result.note}
                </li>
              ))}
            </ul>
            {plan.review.protected.length > 8 ? (
              <p>and {plan.review.protected.length - 8} more</p>
            ) : null}
          </section>
        ) : null}

        {plan.review.ambiguous.length > 0 ? (
          <section className="mz-import-ambiguous">
            <h3>Read one of two ways</h3>
            <ul>
              {plan.review.ambiguous.slice(0, 8).map((issue) => (
                <li key={`${issue.row}:${issue.col}`}>
                  <code>{issue.input}</code> {issue.result.note}
                </li>
              ))}
            </ul>
          </section>
        ) : null}

        <footer className="mz-modal-footer">
          <button type="button" onClick={() => controller.cancelImport()}>
            Cancel
          </button>
          <button
            type="button"
            className="mz-primary"
            onClick={() => controller.confirmImport(resolveImport)}
          >
            Import {plan.review.total} cells
          </button>
        </footer>
      </div>
    </div>
  );
}

function ColumnSummary({ column }: { column: ColumnPlan }) {
  const parts: string[] = [];
  const dates = column.kinds.date + column.kinds.datetime + column.kinds.time;
  if (dates > 0) parts.push(`${dates} date${dates === 1 ? '' : 's'}`);
  if (column.kinds.number > 0) parts.push(`${column.kinds.number} number${column.kinds.number === 1 ? '' : 's'}`);
  if (column.kinds.text > 0) parts.push(`${column.kinds.text} text`);
  if (column.kinds.boolean > 0) parts.push(`${column.kinds.boolean} boolean`);

  return (
    <div className="mz-import-col-summary">
      <span>{parts.join(', ') || 'empty'}</span>
      {column.protectedCount > 0 ? (
        <span className="mz-import-protected-flag" title={column.reasons.join('; ')}>
          {column.protectedCount} protected
        </span>
      ) : null}
      {column.ambiguous > 0 ? (
        <span className="mz-import-ambiguous-flag">{column.ambiguous} ambiguous</span>
      ) : null}
    </div>
  );
}
