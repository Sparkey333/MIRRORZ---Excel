/**
 * A tokeniser for the formula EDITOR, deliberately separate from the engine's.
 *
 * The engine's lexer exists to feed a parser: it must be exact, and it is
 * allowed to fail. An editor's tokeniser has the opposite job. It runs on every
 * keystroke over text that is, by definition, usually incomplete - `=SUM(A1:` is
 * what a half-typed formula looks like - and it must never throw, never lose a
 * character, and always produce a complete cover of the source so the highlight
 * overlay lines up with the textarea underneath it exactly.
 *
 * Everything the editor needs is derived here rather than in the component:
 * which parenthesis matches which (so both can be highlighted, and unmatched
 * ones flagged), which function call the caret sits inside and at which argument
 * (for signature help), and what word the caret is on (for autocomplete). Those
 * are all pure functions of text and a caret offset, which is what makes them
 * testable without rendering anything.
 */

export type TokenKind =
  | 'equals'
  | 'function'
  | 'sheet'
  | 'ref'
  | 'number'
  | 'string'
  | 'boolean'
  | 'error'
  | 'name'
  | 'operator'
  | 'paren'
  | 'brace'
  | 'separator'
  | 'whitespace'
  | 'unknown';

export interface FormulaToken {
  kind: TokenKind;
  start: number;
  end: number;
  text: string;
  /**
   * Nesting depth, used to colour matching pairs. An opening bracket carries the
   * depth it opens at, and its closer carries the same number, so a component
   * can colour a pair identically without re-deriving the nesting.
   */
  depth: number;
}

const OPERATOR_CHARS = '+-*/^&=<>%:';
const ERROR_LITERALS = [
  '#GETTING_DATA',
  '#BLOCKED!',
  '#CONNECT!',
  '#UNKNOWN!',
  '#EXTERNAL!',
  '#DIV/0!',
  '#SPILL!',
  '#VALUE!',
  '#NAME?',
  '#FIELD!',
  '#NULL!',
  '#BUSY!',
  '#CALC!',
  '#NUM!',
  '#REF!',
  '#N/A',
];

const CELL_SHAPE = /^\$?[A-Za-z]{1,3}\$?\d{1,7}$/;
const COLUMN_SHAPE = /^\$?[A-Za-z]{1,3}$/;
const ROW_SHAPE = /^\$?\d{1,7}$/;

/** Is this word shaped like a cell reference? `A1`, `$B$7`, `XFD1048576`. */
export function looksLikeCell(word: string): boolean {
  return CELL_SHAPE.test(word);
}

function isNameStart(ch: string): boolean {
  return /[A-Za-z_\\À-￿]/.test(ch);
}

function isNamePart(ch: string): boolean {
  return /[A-Za-z0-9_.\\À-￿]/.test(ch);
}

/**
 * Tokenise a formula for display. Never throws; anything unrecognised comes back
 * as a single `unknown` character so the output always covers the input.
 */
export function tokenizeFormula(src: string): FormulaToken[] {
  const tokens: FormulaToken[] = [];
  let i = 0;
  let depth = 0;

  const push = (kind: TokenKind, start: number, end: number, d = depth): void => {
    tokens.push({ kind, start, end, text: src.slice(start, end), depth: d });
  };

  while (i < src.length) {
    const ch = src[i]!;
    const start = i;

    if (ch === '=' && tokens.length === 0) {
      push('equals', start, ++i);
      continue;
    }

    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      while (i < src.length && /\s/.test(src[i]!)) i++;
      push('whitespace', start, i);
      continue;
    }

    if (ch === '"') {
      i++;
      // Excel escapes a quote inside a string by doubling it.
      while (i < src.length) {
        if (src[i] === '"') {
          if (src[i + 1] === '"') i += 2;
          else {
            i++;
            break;
          }
        } else i++;
      }
      push('string', start, i);
      continue;
    }

    if (ch === "'") {
      // A quoted sheet name. Unterminated ones still tokenise, as `sheet`, so a
      // half-typed `='My Sh` does not turn the rest of the line into garbage.
      i++;
      while (i < src.length) {
        if (src[i] === "'") {
          if (src[i + 1] === "'") i += 2;
          else {
            i++;
            break;
          }
        } else i++;
      }
      if (src[i] === '!') i++;
      push('sheet', start, i);
      continue;
    }

    if (ch === '#') {
      const upper = src.slice(i).toUpperCase();
      const literal = ERROR_LITERALS.find((e) => upper.startsWith(e));
      if (literal) {
        i += literal.length;
        push('error', start, i);
        continue;
      }
      push('unknown', start, ++i);
      continue;
    }

    if (ch === '(') {
      push('paren', start, ++i, depth);
      depth++;
      continue;
    }
    if (ch === ')') {
      depth = Math.max(0, depth - 1);
      push('paren', start, ++i, depth);
      continue;
    }
    if (ch === '{') {
      push('brace', start, ++i, depth);
      depth++;
      continue;
    }
    if (ch === '}') {
      depth = Math.max(0, depth - 1);
      push('brace', start, ++i, depth);
      continue;
    }
    if (ch === ',' || ch === ';') {
      push('separator', start, ++i);
      continue;
    }

    if (/\d/.test(ch) || (ch === '.' && /\d/.test(src[i + 1] ?? ''))) {
      while (i < src.length && /[\d.]/.test(src[i]!)) i++;
      // Scientific notation, but only when a digit or sign actually follows, so
      // `1E` stays a number followed by a name rather than swallowing the E.
      if (/[eE]/.test(src[i] ?? '') && /[\d+-]/.test(src[i + 1] ?? '')) {
        i += 2;
        while (i < src.length && /\d/.test(src[i]!)) i++;
      }
      const text = src.slice(start, i);
      // `1:1` is a whole-row beam, not a number followed by an operator, but the
      // row half alone is still just a number until the colon is seen; the
      // colon operator carries that meaning, so nothing special is needed here.
      push(ROW_SHAPE.test(text) && src[i] === ':' ? 'ref' : 'number', start, i);
      continue;
    }

    if (ch === '$' || isNameStart(ch)) {
      if (ch === '$') i++;
      while (i < src.length && (isNamePart(src[i]!) || src[i] === '$')) i++;
      const text = src.slice(start, i);

      // An unquoted sheet prefix: the `!` binds to the name before it.
      if (src[i] === '!') {
        i++;
        push('sheet', start, i);
        continue;
      }

      // A table reference: Table1[Column].
      if (src[i] === '[') {
        let bracket = 0;
        do {
          if (src[i] === '[') bracket++;
          else if (src[i] === ']') bracket--;
          i++;
        } while (i < src.length && bracket > 0);
        push('ref', start, i);
        continue;
      }

      if (src[i] === '(') {
        push('function', start, i);
        continue;
      }

      const upper = text.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        push('boolean', start, i);
        continue;
      }
      if (CELL_SHAPE.test(text)) {
        push('ref', start, i);
        continue;
      }
      // A bare column name is only a reference in a beam like `A:A`; on its own
      // it is a defined name, and guessing otherwise would paint every single
      // letter variable as a reference.
      if (COLUMN_SHAPE.test(text) && (src[i] === ':' || tokens[tokens.length - 1]?.text === ':')) {
        push('ref', start, i);
        continue;
      }
      push('name', start, i);
      continue;
    }

    if (OPERATOR_CHARS.includes(ch)) {
      // Two-character comparisons: <=, >=, <>.
      if ((ch === '<' || ch === '>') && (src[i + 1] === '=' || src[i + 1] === '>')) i += 2;
      else i++;
      push('operator', start, i);
      continue;
    }

    push('unknown', start, ++i);
  }

  return tokens;
}

export interface BracketPairs {
  /** Token index -> the index of its partner, for every matched bracket. */
  partner: Map<number, number>;
  /** Token indices of brackets with no partner. */
  unmatched: number[];
}

/** Pair up brackets. Unmatched ones are reported rather than guessed at. */
export function matchBrackets(tokens: readonly FormulaToken[]): BracketPairs {
  const partner = new Map<number, number>();
  const unmatched: number[] = [];
  const stack: number[] = [];

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]!;
    if (token.kind !== 'paren' && token.kind !== 'brace') continue;
    if (token.text === '(' || token.text === '{') {
      stack.push(i);
      continue;
    }
    const open = stack.pop();
    if (open === undefined) {
      unmatched.push(i);
      continue;
    }
    const opener = tokens[open]!;
    // A `)` closing a `{` is a genuine mistake; pairing them anyway would draw a
    // highlight that tells the user their formula is fine when it is not.
    if ((opener.text === '(') !== (token.text === ')')) {
      unmatched.push(open, i);
      continue;
    }
    partner.set(open, i);
    partner.set(i, open);
  }
  unmatched.push(...stack);
  unmatched.sort((a, b) => a - b);
  return { partner, unmatched };
}

/**
 * The bracket pair to highlight for a caret position.
 *
 * Editors differ on whether the caret "is on" the bracket before it or after it.
 * Both, is the answer users expect: a caret touching either side of a bracket
 * highlights it, with the one immediately to the left winning when both apply.
 */
export function bracketPairAtCaret(
  tokens: readonly FormulaToken[],
  caret: number,
): { open: number; close: number } | null {
  const pairs = matchBrackets(tokens);
  const candidates: number[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (t.kind !== 'paren' && t.kind !== 'brace') continue;
    if (t.end === caret) candidates.unshift(i);
    else if (t.start === caret) candidates.push(i);
  }
  for (const index of candidates) {
    const other = pairs.partner.get(index);
    if (other === undefined) continue;
    return index < other ? { open: index, close: other } : { open: other, close: index };
  }
  return null;
}

export interface ActiveCall {
  name: string;
  /** Token index of the function name. */
  nameToken: number;
  /** Zero-based index of the argument the caret is in. */
  argIndex: number;
  /** Source offset just after the opening parenthesis. */
  argsStart: number;
}

/**
 * The innermost function call the caret sits inside, and which argument it is
 * on, for signature help. Returns null when the caret is not inside a call.
 */
export function activeCall(tokens: readonly FormulaToken[], caret: number): ActiveCall | null {
  const open: { nameToken: number; parenToken: number; args: number }[] = [];
  let result: ActiveCall | null = null;

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]!;
    // Once the caret is passed, the enclosing frames are whatever is still open.
    if (token.start >= caret) break;

    if (token.kind === 'paren' && token.text === '(') {
      const prev = tokens[i - 1];
      open.push({ nameToken: prev?.kind === 'function' ? i - 1 : -1, parenToken: i, args: 0 });
    } else if (token.kind === 'paren' && token.text === ')') {
      // A close bracket at or after the caret does not end the frame the caret
      // is in, which is why the loop stops at the caret rather than the end.
      open.pop();
    } else if (token.kind === 'separator' && open.length > 0) {
      open[open.length - 1]!.args++;
    }
  }

  for (let i = open.length - 1; i >= 0; i--) {
    const frame = open[i]!;
    if (frame.nameToken < 0) continue;
    result = {
      name: tokens[frame.nameToken]!.text,
      nameToken: frame.nameToken,
      argIndex: frame.args,
      argsStart: tokens[frame.parenToken]!.end,
    };
    break;
  }
  return result;
}

export interface CompletionContext {
  /** The partial word under the caret. */
  prefix: string;
  /** Source range the completion should replace. */
  start: number;
  end: number;
}

/**
 * What the caret is in the middle of typing.
 *
 * Only identifier-shaped text is completable, and only when the caret is at its
 * end: completing a word the user is editing the middle of would replace text
 * they can see and did not ask to change.
 */
export function completionContext(src: string, caret: number): CompletionContext | null {
  if (caret < 0 || caret > src.length) return null;
  let start = caret;
  while (start > 0 && isNamePart(src[start - 1]!)) start--;
  if (start === caret) return null;
  const prefix = src.slice(start, caret);
  // A cell reference is not something to offer function names for.
  if (looksLikeCell(prefix) || /^\d/.test(prefix)) return null;
  // Inside a string literal there is nothing to complete either.
  if (insideString(src, caret)) return null;
  return { prefix, start, end: caret };
}

function insideString(src: string, caret: number): boolean {
  let open = false;
  for (let i = 0; i < caret; i++) {
    if (src[i] !== '"') continue;
    if (open && src[i + 1] === '"') {
      i++;
      continue;
    }
    open = !open;
  }
  return open;
}

/** Replace the completion range with a chosen name, and say where the caret goes. */
export function applyCompletion(
  src: string,
  context: CompletionContext,
  name: string,
  isFunction: boolean,
): { text: string; caret: number } {
  // Typing the opening parenthesis too is the whole point of completing a
  // function name; the closer is left to the editor's own bracket handling.
  const inserted = isFunction ? `${name}(` : name;
  return {
    text: src.slice(0, context.start) + inserted + src.slice(context.end),
    caret: context.start + inserted.length,
  };
}
