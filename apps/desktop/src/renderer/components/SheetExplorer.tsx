/**
 * The sheet explorer.
 *
 * A workbook with sixty sheets is normal in finance and reporting, and at that
 * count the tab strip has stopped being navigation - it is a horizontal scroll
 * bar with words on it. The explorer is the answer: a searchable, reorderable
 * list that shows every sheet including the hidden ones, which the tab strip by
 * definition cannot.
 *
 * Hidden sheets being invisible is a real hazard, not a tidiness feature: a
 * workbook can carry a `veryHidden` sheet that the Excel UI gives no way to
 * reveal at all. Listing them is the honest behaviour.
 */

import { useMemo, useState } from 'react';
import { useController, useDerived } from '../state/context.js';
import { fuzzyFilter, highlightSegments } from '../model/fuzzy.js';
import type { SheetSummary } from '../state/controller.js';

const TAB_COLORS = ['#e03131', '#f08c00', '#2f9e44', '#1971c2', '#7048e8', '#495057'];

export function SheetExplorer() {
  const controller = useController();
  const [query, setQuery] = useState('');
  const [renaming, setRenaming] = useState<string | null>(null);
  const [draftName, setDraftName] = useState('');
  const [showHidden, setShowHidden] = useState(true);

  const sheets = useDerived(() => controller.sheetSummaries());
  const pool = useMemo(
    () => (showHidden ? sheets : sheets.filter((s) => s.visibility === 'visible')),
    [sheets, showHidden],
  );
  const results = useMemo(
    () => fuzzyFilter(pool, query, { key: (sheet: SheetSummary) => sheet.name }),
    [pool, query],
  );

  return (
    <aside className="mz-explorer" aria-label="Sheet explorer">
      <header className="mz-panel-header">
        <h2>Sheets</h2>
        <button type="button" aria-label="Close sheet explorer" onClick={() => controller.togglePanel('explorer', false)}>
          &#215;
        </button>
      </header>
      <input
        className="mz-explorer-search"
        aria-label="Search sheets"
        placeholder="Search sheets"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
      />
      <label className="mz-explorer-toggle">
        <input type="checkbox" checked={showHidden} onChange={(e) => setShowHidden(e.target.checked)} />
        Show hidden sheets
      </label>
      <ul className="mz-explorer-list">
        {results.map(({ item: sheet, positions }) => (
          <li
            key={sheet.name}
            className={sheet.active ? 'mz-explorer-item mz-active' : 'mz-explorer-item'}
            data-visibility={sheet.visibility}
          >
            <span className="mz-explorer-color" style={{ background: sheet.tabColor ?? 'transparent' }} />
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
              <button
                type="button"
                className="mz-explorer-name"
                onClick={() => controller.setActiveSheet(sheet.name)}
                onDoubleClick={() => {
                  setRenaming(sheet.name);
                  setDraftName(sheet.name);
                }}
              >
                {highlightSegments(sheet.name, positions).map((segment, i) => (
                  <span key={i} className={segment.match ? 'mz-match' : undefined}>
                    {segment.text}
                  </span>
                ))}
              </button>
            )}
            <span className="mz-explorer-meta">
              {sheet.visibility === 'visible' ? `${sheet.cellCount} cells` : sheet.visibility}
            </span>
            <div className="mz-explorer-actions">
              <button
                type="button"
                aria-label={`Move ${sheet.name} up`}
                disabled={sheet.index === 0}
                onClick={() => controller.moveSheet(sheet.name, sheet.index - 1)}
              >
                &#8593;
              </button>
              <button
                type="button"
                aria-label={`Move ${sheet.name} down`}
                disabled={sheet.index === sheets.length - 1}
                onClick={() => controller.moveSheet(sheet.name, sheet.index + 1)}
              >
                &#8595;
              </button>
              <button
                type="button"
                aria-label={
                  sheet.visibility === 'visible' ? `Hide ${sheet.name}` : `Show ${sheet.name}`
                }
                onClick={() =>
                  controller.setSheetVisibility(
                    sheet.name,
                    sheet.visibility === 'visible' ? 'hidden' : 'visible',
                  )
                }
              >
                {sheet.visibility === 'visible' ? 'Hide' : 'Show'}
              </button>
              <details className="mz-explorer-colors">
                <summary aria-label={`Colour ${sheet.name}`}>Colour</summary>
                <div className="mz-swatches">
                  {TAB_COLORS.map((color) => (
                    <button
                      key={color}
                      type="button"
                      aria-label={`Colour ${sheet.name} ${color}`}
                      style={{ background: color }}
                      onClick={() => controller.setSheetColor(sheet.name, color)}
                    />
                  ))}
                </div>
              </details>
            </div>
          </li>
        ))}
        {results.length === 0 ? <li className="mz-explorer-empty">No sheets match “{query}”</li> : null}
      </ul>
      <footer className="mz-explorer-footer">
        <button type="button" onClick={() => controller.addSheet()}>
          Add sheet
        </button>
        <span>{sheets.length} total</span>
      </footer>
    </aside>
  );
}
