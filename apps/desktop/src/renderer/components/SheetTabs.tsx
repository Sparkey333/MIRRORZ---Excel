/**
 * The tab strip.
 *
 * Tabs are the right control for five sheets and the wrong one for fifty, which
 * is why this strip exists alongside the explorer sidebar rather than instead of
 * it. Here the strip stays as it is - familiar, draggable, double-click to
 * rename - and the sidebar takes over when the count makes tabs useless.
 */

import { useState } from 'react';
import { useApp, useController, useDerived } from '../state/context.js';

const TAB_COLORS = ['#e03131', '#f08c00', '#2f9e44', '#1971c2', '#7048e8', '#495057'];

export function SheetTabs() {
  const controller = useController();
  const snapshot = useApp();
  const sheets = useDerived(() => controller.sheetSummaries());
  const [renaming, setRenaming] = useState<string | null>(null);
  const [draftName, setDraftName] = useState('');
  const [dragging, setDragging] = useState<string | null>(null);
  const [colorFor, setColorFor] = useState<string | null>(null);

  const visible = sheets.filter((s) => s.visibility === 'visible');
  const hiddenCount = sheets.length - visible.length;

  return (
    <div className="mz-tabs" role="tablist" aria-label="Sheets">
      <button type="button" className="mz-tab-add" aria-label="Add sheet" onClick={() => controller.addSheet()}>
        +
      </button>
      {visible.map((sheet) => (
        <div
          key={sheet.name}
          className={sheet.active ? 'mz-tab mz-tab-active' : 'mz-tab'}
          role="tab"
          aria-selected={sheet.active}
          tabIndex={0}
          draggable
          onDragStart={() => setDragging(sheet.name)}
          onDragOver={(e) => e.preventDefault()}
          onDrop={() => {
            if (dragging && dragging !== sheet.name) controller.moveSheet(dragging, sheet.index);
            setDragging(null);
          }}
          onClick={() => controller.setActiveSheet(sheet.name)}
          onDoubleClick={() => {
            setRenaming(sheet.name);
            setDraftName(sheet.name);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter' || e.key === ' ') controller.setActiveSheet(sheet.name);
            if (e.key === 'F2') {
              setRenaming(sheet.name);
              setDraftName(sheet.name);
            }
          }}
          style={sheet.tabColor ? { borderBottomColor: sheet.tabColor } : undefined}
        >
          {renaming === sheet.name ? (
            <input
              aria-label={`Rename ${sheet.name}`}
              value={draftName}
              autoFocus
              onChange={(e) => setDraftName(e.target.value)}
              onBlur={() => setRenaming(null)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  controller.renameSheet(sheet.name, draftName);
                  setRenaming(null);
                }
                if (e.key === 'Escape') setRenaming(null);
              }}
            />
          ) : (
            <span className="mz-tab-label">{sheet.name}</span>
          )}
          <button
            type="button"
            className="mz-tab-menu"
            aria-label={`Options for ${sheet.name}`}
            onClick={(e) => {
              e.stopPropagation();
              setColorFor(colorFor === sheet.name ? null : sheet.name);
            }}
          >
            &#8942;
          </button>
          {colorFor === sheet.name ? (
            <div className="mz-tab-popover" role="menu" aria-label={`${sheet.name} options`}>
              <div className="mz-swatches">
                {TAB_COLORS.map((color) => (
                  <button
                    key={color}
                    type="button"
                    aria-label={`Colour ${sheet.name} ${color}`}
                    style={{ background: color }}
                    onClick={() => {
                      controller.setSheetColor(sheet.name, color);
                      setColorFor(null);
                    }}
                  />
                ))}
                <button
                  type="button"
                  aria-label={`Clear colour of ${sheet.name}`}
                  onClick={() => {
                    controller.setSheetColor(sheet.name, undefined);
                    setColorFor(null);
                  }}
                >
                  &#215;
                </button>
              </div>
              <button type="button" role="menuitem" onClick={() => controller.setSheetVisibility(sheet.name, 'hidden')}>
                Hide
              </button>
              <button type="button" role="menuitem" onClick={() => controller.removeSheet(sheet.name)}>
                Delete
              </button>
            </div>
          ) : null}
        </div>
      ))}
      {hiddenCount > 0 ? (
        <button
          type="button"
          className="mz-tab-hidden-count"
          onClick={() => controller.togglePanel('explorer', true)}
        >
          {hiddenCount} hidden
        </button>
      ) : null}
      <span className="mz-tab-count">{visible.length} sheets</span>
      {snapshot.panels.explorer ? null : (
        <button type="button" onClick={() => controller.togglePanel('explorer', true)}>
          Explorer
        </button>
      )}
    </div>
  );
}
