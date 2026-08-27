/**
 * Find and replace.
 *
 * The hit list is shown rather than only stepped through, because "3 of 47" with
 * no way to see the other 46 is how a replace-all goes wrong. Replace-all is one
 * transaction, so it is one undo.
 */

import { useMemo, useState } from 'react';
import { a1 } from '@mirrorz/core';
import { useController, useDerived } from '../state/context.js';

export function FindReplace({ onClose }: { onClose: () => void }) {
  const controller = useController();
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [inFormulas, setInFormulas] = useState(false);
  const [index, setIndex] = useState(0);

  const hits = useDerived(
    () => controller.find(query, { matchCase, whole, formulas: inFormulas }),
    [query, matchCase, whole, inFormulas],
  );

  const current = hits[Math.min(index, Math.max(0, hits.length - 1))];
  const position = useMemo(() => (hits.length === 0 ? 0 : Math.min(index, hits.length - 1) + 1), [hits, index]);

  const step = (delta: number): void => {
    if (hits.length === 0) return;
    const next = (((index + delta) % hits.length) + hits.length) % hits.length;
    setIndex(next);
    const target = hits[next];
    if (target) controller.goTo(target);
  };

  return (
    <div className="mz-find" role="dialog" aria-label="Find and replace">
      <input
        aria-label="Find"
        placeholder="Find"
        value={query}
        autoFocus
        onChange={(e) => {
          setQuery(e.target.value);
          setIndex(0);
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') step(e.shiftKey ? -1 : 1);
          if (e.key === 'Escape') onClose();
        }}
      />
      <input
        aria-label="Replace with"
        placeholder="Replace with"
        value={replacement}
        onChange={(e) => setReplacement(e.target.value)}
      />
      <span className="mz-find-count" role="status">
        {hits.length === 0 ? 'No matches' : `${position} of ${hits.length}`}
        {current ? ` (${a1(current.row, current.col)})` : ''}
      </span>
      <button type="button" onClick={() => step(-1)} aria-label="Previous match">
        &#8593;
      </button>
      <button type="button" onClick={() => step(1)} aria-label="Next match">
        &#8595;
      </button>
      <button
        type="button"
        disabled={hits.length === 0}
        onClick={() => controller.replaceAll(query, replacement, { matchCase, whole })}
      >
        Replace all
      </button>
      <label>
        <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
        Match case
      </label>
      <label>
        <input type="checkbox" checked={whole} onChange={(e) => setWhole(e.target.checked)} />
        Whole cell
      </label>
      <label>
        <input type="checkbox" checked={inFormulas} onChange={(e) => setInFormulas(e.target.checked)} />
        Search formulas
      </label>
      <button type="button" onClick={onClose} aria-label="Close find">
        &#215;
      </button>
    </div>
  );
}
