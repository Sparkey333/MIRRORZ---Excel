/**
 * The command palette.
 *
 * One box that searches commands, sheets and named ranges together. The reason
 * to search all three at once is that the user's question is "take me to the
 * assumptions", and whether that is a sheet, a named range or a command is our
 * problem, not theirs.
 */

import { useEffect, useMemo, useRef, useState } from 'react';
import { useApp, useController, useDerived } from '../state/context.js';
import { moveIndex, pushRecent, searchPalette } from '../model/palette.js';
import { highlightSegments } from '../model/fuzzy.js';
import { buildCommands, buildPaletteItems, runPaletteItem } from '../state/commands.js';

export function CommandPalette() {
  const controller = useController();
  const snapshot = useApp();
  const [query, setQuery] = useState('');
  const [index, setIndex] = useState(0);
  const [recent, setRecent] = useState<string[]>([]);
  const inputRef = useRef<HTMLInputElement>(null);

  const commands = useDerived(() => buildCommands(controller, snapshot), [snapshot.paletteOpen]);
  const items = useDerived(
    () => buildPaletteItems(controller, snapshot, commands),
    [snapshot.paletteOpen, commands],
  );

  const results = useMemo(
    () =>
      searchPalette(items, snapshot.paletteMode === 'command' && !query.startsWith('>') ? `>${query}` : query, {
        limit: 40,
        recent,
      }),
    [items, query, recent, snapshot.paletteMode],
  );

  useEffect(() => {
    if (snapshot.paletteOpen) {
      setQuery('');
      setIndex(0);
      inputRef.current?.focus();
    }
  }, [snapshot.paletteOpen]);

  useEffect(() => setIndex(0), [query]);

  if (!snapshot.paletteOpen) return null;

  const choose = (at: number): void => {
    const item = results[at]?.item;
    if (!item) return;
    setRecent((r) => pushRecent(r, item.id));
    controller.setPalette(false);
    runPaletteItem(controller, item, commands);
  };

  return (
    <div
      className="mz-palette-backdrop"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) controller.setPalette(false);
      }}
    >
      <div className="mz-palette" role="dialog" aria-modal="true" aria-label="Command palette">
        <input
          ref={inputRef}
          className="mz-palette-input"
          aria-label="Search commands, sheets and names"
          placeholder="Search commands, sheets and names. > for commands, @ for sheets"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'ArrowDown') {
              e.preventDefault();
              setIndex((i) => moveIndex(i, 1, results.length));
            } else if (e.key === 'ArrowUp') {
              e.preventDefault();
              setIndex((i) => moveIndex(i, -1, results.length));
            } else if (e.key === 'Enter') {
              e.preventDefault();
              choose(index);
            } else if (e.key === 'Escape') {
              e.preventDefault();
              controller.setPalette(false);
            }
          }}
        />
        <ul className="mz-palette-list" role="listbox" aria-label="Results">
          {results.map((result, i) => (
            <li
              key={result.item.id}
              role="option"
              aria-selected={i === index}
              aria-disabled={result.item.enabled === false}
              className={i === index ? 'mz-palette-item mz-selected' : 'mz-palette-item'}
              onMouseDown={(e) => {
                e.preventDefault();
                choose(i);
              }}
              onMouseEnter={() => setIndex(i)}
            >
              <span className="mz-palette-kind" data-kind={result.item.kind}>
                {result.item.kind}
              </span>
              <span className="mz-palette-title">
                {highlightSegments(result.item.title, result.positions).map((segment, s) => (
                  <span key={s} className={segment.match ? 'mz-match' : undefined}>
                    {segment.text}
                  </span>
                ))}
              </span>
              {result.item.detail ? <span className="mz-palette-detail">{result.item.detail}</span> : null}
              {result.item.shortcut ? (
                <kbd className="mz-palette-shortcut">{result.item.shortcut}</kbd>
              ) : null}
            </li>
          ))}
          {results.length === 0 ? <li className="mz-palette-empty">Nothing matches “{query}”</li> : null}
        </ul>
      </div>
    </div>
  );
}
