/**
 * The formula editor: a real text editor, not a one-line box.
 *
 * Excel's formula bar has been a single line with a tiny expander for thirty
 * years, and editing a nested IF in it is the most complained-about part of the
 * product after the paperclip. Everything here follows from refusing that:
 *
 *   It is a textarea, so newlines inside a formula are ordinary text and the
 *   editor grows to fit. Excel accepts Alt+Enter in formulas too, but gives you
 *   nowhere to see the result.
 *
 *   Highlighting is a `<pre>` layer behind a transparent textarea, with both
 *   sharing one font metric. Rich-text editing inside a contenteditable would
 *   break the caret, the IME and undo, all for the same visual result.
 *
 *   The matching bracket is shown for the caret's own bracket, and unmatched
 *   ones are marked, because "missing a parenthesis" is the single most common
 *   formula error and Excel's only feedback is a modal after you press Enter.
 *
 *   Completion is over the engine's own registry, so it can never list a
 *   function the engine does not have.
 */

import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  activeCall,
  applyCompletion,
  bracketPairAtCaret,
  completionContext,
  matchBrackets,
  tokenizeFormula,
  type FormulaToken,
} from '../model/formula-tokens.js';
import { fuzzyFilter, highlightSegments } from '../model/fuzzy.js';

export interface FormulaEditorProps {
  value: string;
  onChange: (value: string) => void;
  onCommit: (value: string) => void;
  onCancel: () => void;
  /** Function names from the engine's registry. */
  functionNames: readonly string[];
  /** Summaries keyed by upper-case function name, for the completion detail. */
  summaries?: Readonly<Record<string, string>>;
  height: number;
  onHeightChange: (height: number) => void;
  placeholder?: string;
  disabled?: boolean;
}

const MIN_HEIGHT = 26;
const MAX_HEIGHT = 480;
const MAX_COMPLETIONS = 12;

export function FormulaEditor({
  value,
  onChange,
  onCommit,
  onCancel,
  functionNames,
  summaries,
  height,
  onHeightChange,
  placeholder,
  disabled,
}: FormulaEditorProps) {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const highlightRef = useRef<HTMLPreElement>(null);
  const [caret, setCaret] = useState(0);
  const [completionIndex, setCompletionIndex] = useState(0);
  const [completionOpen, setCompletionOpen] = useState(false);

  const tokens = useMemo(() => tokenizeFormula(value), [value]);
  const brackets = useMemo(() => matchBrackets(tokens), [tokens]);
  const pair = useMemo(() => bracketPairAtCaret(tokens, caret), [tokens, caret]);
  const call = useMemo(() => activeCall(tokens, caret), [tokens, caret]);

  // Completion only makes sense inside a formula; a plain value should not
  // suggest SUMPRODUCT because it happens to start with an S.
  const context = useMemo(
    () => (value.startsWith('=') ? completionContext(value, caret) : null),
    [value, caret],
  );

  const completions = useMemo(() => {
    if (!context) return [];
    return fuzzyFilter(functionNames, context.prefix, {
      key: (name) => name,
      ...(summaries ? { extra: (name: string) => summaries[name] } : {}),
      limit: MAX_COMPLETIONS,
    });
  }, [context, functionNames, summaries]);

  useEffect(() => {
    setCompletionIndex(0);
    setCompletionOpen(completions.length > 0);
  }, [completions]);

  const syncCaret = useCallback(() => {
    const el = textareaRef.current;
    if (el) setCaret(el.selectionStart);
  }, []);

  // The highlight layer must scroll with the textarea or the two drift apart
  // the moment a formula is longer than the box.
  const syncScroll = useCallback(() => {
    const el = textareaRef.current;
    const pre = highlightRef.current;
    if (el && pre) {
      pre.scrollTop = el.scrollTop;
      pre.scrollLeft = el.scrollLeft;
    }
  }, []);

  useLayoutEffect(syncScroll, [value, syncScroll]);

  const accept = useCallback(
    (index: number) => {
      if (!context) return;
      const chosen = completions[index]?.item;
      if (chosen === undefined) return;
      const applied = applyCompletion(value, context, chosen, true);
      onChange(applied.text);
      setCompletionOpen(false);
      // Restoring the caret has to wait for React to write the new value back
      // into the textarea, or the browser puts it at the end of the old text.
      queueMicrotask(() => {
        const el = textareaRef.current;
        if (!el) return;
        el.setSelectionRange(applied.caret, applied.caret);
        setCaret(applied.caret);
      });
    },
    [completions, context, onChange, value],
  );

  const onKeyDown = (event: React.KeyboardEvent<HTMLTextAreaElement>): void => {
    if (completionOpen && completions.length > 0) {
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        setCompletionIndex((i) => (i + 1) % completions.length);
        return;
      }
      if (event.key === 'ArrowUp') {
        event.preventDefault();
        setCompletionIndex((i) => (i - 1 + completions.length) % completions.length);
        return;
      }
      if (event.key === 'Tab' || (event.key === 'Enter' && !event.shiftKey && !event.altKey)) {
        event.preventDefault();
        accept(completionIndex);
        return;
      }
      if (event.key === 'Escape') {
        event.preventDefault();
        setCompletionOpen(false);
        return;
      }
    }

    // Alt+Enter and Shift+Enter add a line, which is the point of a multi-line
    // editor; plain Enter commits.
    if (event.key === 'Enter' && (event.altKey || event.shiftKey)) return;
    if (event.key === 'Enter') {
      event.preventDefault();
      onCommit(value);
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      onCancel();
    }
  };

  const startResize = (event: React.PointerEvent<HTMLDivElement>): void => {
    event.preventDefault();
    const startY = event.clientY;
    const startHeight = height;
    const move = (e: PointerEvent): void => {
      const next = Math.min(MAX_HEIGHT, Math.max(MIN_HEIGHT, startHeight + (e.clientY - startY)));
      onHeightChange(next);
    };
    const up = (): void => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  return (
    <div className="mz-formula-editor" data-testid="formula-editor">
      <div className="mz-formula-editor-body" style={{ height }}>
        <pre className="mz-formula-highlight" ref={highlightRef} aria-hidden="true">
          {renderTokens(tokens, brackets.unmatched, pair)}
          {'\n'}
        </pre>
        <textarea
          ref={textareaRef}
          className="mz-formula-input"
          value={value}
          spellCheck={false}
          disabled={disabled}
          placeholder={placeholder}
          aria-label="Formula"
          onChange={(e) => {
            onChange(e.target.value);
            setCaret(e.target.selectionStart);
          }}
          onKeyDown={onKeyDown}
          onKeyUp={syncCaret}
          onClick={syncCaret}
          onSelect={syncCaret}
          onScroll={syncScroll}
        />
        {completionOpen && completions.length > 0 && context ? (
          <ul className="mz-completions" role="listbox" aria-label="Function suggestions">
            {completions.map((result, i) => (
              <li
                key={result.item}
                role="option"
                aria-selected={i === completionIndex}
                className={i === completionIndex ? 'mz-completion mz-selected' : 'mz-completion'}
                onMouseDown={(e) => {
                  // mousedown, not click: a click would blur the textarea first
                  // and the editor would commit before the completion applied.
                  e.preventDefault();
                  accept(i);
                }}
              >
                <span className="mz-completion-name">
                  {highlightSegments(result.item, result.positions).map((segment, s) => (
                    <span key={s} className={segment.match ? 'mz-match' : undefined}>
                      {segment.text}
                    </span>
                  ))}
                </span>
                {summaries?.[result.item] ? (
                  <span className="mz-completion-detail">{summaries[result.item]}</span>
                ) : null}
              </li>
            ))}
          </ul>
        ) : null}
      </div>
      <div className="mz-formula-status">
        {call ? (
          <span className="mz-signature">
            {call.name}, argument {call.argIndex + 1}
          </span>
        ) : null}
        {brackets.unmatched.length > 0 ? (
          <span className="mz-unmatched" role="status">
            {brackets.unmatched.length} unmatched bracket
            {brackets.unmatched.length === 1 ? '' : 's'}
          </span>
        ) : null}
      </div>
      <div
        className="mz-formula-resize"
        role="separator"
        aria-label="Resize formula editor"
        aria-orientation="horizontal"
        onPointerDown={startResize}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown') onHeightChange(Math.min(MAX_HEIGHT, height + 16));
          if (e.key === 'ArrowUp') onHeightChange(Math.max(MIN_HEIGHT, height - 16));
        }}
        tabIndex={0}
      />
    </div>
  );
}

/**
 * Turn tokens into spans.
 *
 * Every character of the source must end up in exactly one span, including
 * whitespace, or the highlight layer stops lining up with the textarea.
 */
function renderTokens(
  tokens: readonly FormulaToken[],
  unmatched: readonly number[],
  pair: { open: number; close: number } | null,
) {
  const unmatchedSet = new Set(unmatched);
  return tokens.map((token, i) => {
    const classes = [`mz-tok-${token.kind}`];
    if (token.kind === 'paren' || token.kind === 'brace') {
      classes.push(`mz-depth-${token.depth % 4}`);
      if (unmatchedSet.has(i)) classes.push('mz-bracket-unmatched');
      else if (pair && (i === pair.open || i === pair.close)) classes.push('mz-bracket-match');
    }
    return (
      <span key={i} className={classes.join(' ')} data-token={token.kind}>
        {token.text}
      </span>
    );
  });
}
