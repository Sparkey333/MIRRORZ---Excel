/**
 * The formula bar: a name box, the editor, and the commit controls.
 *
 * The bar owns the draft text rather than the controller, because a half-typed
 * formula is not a document change and putting it in the command log would fill
 * the undo history with keystrokes. It becomes a document change exactly once,
 * on commit.
 */

import { useEffect, useMemo, useState } from 'react';
import { a1 } from '@mirrorz/core';
import { useApp, useController } from '../state/context.js';
import { describeSelection } from '../model/selection.js';
import { FormulaEditor } from './FormulaEditor.js';

export function FormulaBar() {
  const controller = useController();
  const snapshot = useApp();
  const addr = controller.activeAddr();
  const committed = controller.editText(addr);

  const [draft, setDraft] = useState(committed);
  const [height, setHeight] = useState(28);
  const [editing, setEditing] = useState(false);

  // Moving the selection abandons an uncommitted draft, which is what every
  // spreadsheet does and what users expect: the bar always shows the cell.
  useEffect(() => {
    setDraft(committed);
    setEditing(false);
  }, [committed, addr.sheet, addr.row, addr.col]);

  const functionNames = useMemo(() => controller.registry.names(), [controller]);

  const label = describeSelection(snapshot.selection, a1);

  return (
    <div className="mz-formula-bar">
      <div className="mz-name-box" title="Selection">
        <span className="mz-name-box-text">{label}</span>
      </div>
      <div className="mz-formula-bar-actions">
        <button
          type="button"
          className="mz-icon-button"
          aria-label="Cancel edit"
          disabled={!editing}
          onClick={() => {
            setDraft(committed);
            setEditing(false);
          }}
        >
          &#215;
        </button>
        <button
          type="button"
          className="mz-icon-button"
          aria-label="Commit edit"
          disabled={!editing}
          onClick={() => {
            controller.commitEntry(draft, addr);
            setEditing(false);
          }}
        >
          &#10003;
        </button>
      </div>
      <FormulaEditor
        value={draft}
        height={height}
        onHeightChange={setHeight}
        functionNames={functionNames}
        placeholder={`${addr.sheet}!${a1(addr.row, addr.col)}`}
        onChange={(next) => {
          setDraft(next);
          setEditing(true);
        }}
        onCommit={(next) => {
          controller.commitEntry(next, addr);
          setEditing(false);
        }}
        onCancel={() => {
          setDraft(committed);
          setEditing(false);
        }}
      />
    </div>
  );
}
