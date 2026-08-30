import { describe, expect, it } from 'vitest';
import {
  activeCall,
  applyCompletion,
  bracketPairAtCaret,
  completionContext,
  looksLikeCell,
  matchBrackets,
  tokenizeFormula,
} from '../src/renderer/model/formula-tokens.js';

const kinds = (src: string) => tokenizeFormula(src).map((t) => `${t.kind}:${t.text}`);

describe('tokenizeFormula', () => {
  it('covers every character of the source', () => {
    const src = '=SUM(A1:B2, "hi", 3.5%) + Sheet2!C3';
    const tokens = tokenizeFormula(src);
    expect(tokens.map((t) => t.text).join('')).toBe(src);
  });

  it('never throws on incomplete input', () => {
    for (const src of ['=', '=SUM(', '="unterminated', "='My Sh", '=#', '=A1:', '={1,2']) {
      expect(() => tokenizeFormula(src)).not.toThrow();
    }
  });

  it('marks the leading equals', () => {
    expect(tokenizeFormula('=1')[0]).toMatchObject({ kind: 'equals', text: '=' });
  });

  it('treats a later equals as an operator', () => {
    expect(kinds('=A1=B1')).toContain('operator:=');
  });

  it('recognises a function name only when a parenthesis follows', () => {
    expect(kinds('=SUM(1)')).toContain('function:SUM');
    expect(kinds('=SUM')).toContain('name:SUM');
  });

  it('recognises cell references and absolute forms', () => {
    expect(kinds('=$A$1+B22')).toEqual(
      expect.arrayContaining(['ref:$A$1', 'ref:B22', 'operator:+']),
    );
  });

  it('recognises a sheet prefix, quoted or not', () => {
    expect(kinds('=Sheet1!A1')).toContain('sheet:Sheet1!');
    expect(kinds("='My Sheet'!A1")).toContain("sheet:'My Sheet'!");
  });

  it('keeps a doubled quote inside a string', () => {
    expect(kinds('="a""b"')).toContain('string:"a""b"');
  });

  it('recognises error literals', () => {
    expect(kinds('=#REF!+1')).toContain('error:#REF!');
    expect(kinds('=#N/A')).toContain('error:#N/A');
  });

  it('reads scientific notation as one number, but not a bare trailing E', () => {
    expect(kinds('=1.5E+3')).toContain('number:1.5E+3');
    expect(kinds('=1E')).toEqual(expect.arrayContaining(['number:1', 'name:E']));
  });

  it('keeps whitespace as its own token, since a space is the intersection operator', () => {
    expect(kinds('=SUM(A1:C3 B1:B5)')).toContain('whitespace: ');
  });

  it('reads a beam as references around a colon', () => {
    expect(kinds('=SUM(A:A)')).toEqual(expect.arrayContaining(['ref:A', 'operator::']));
  });

  it('reads a bare single letter as a name, not a reference', () => {
    expect(kinds('=x+1')).toContain('name:x');
  });

  it('reads a structured reference as one token', () => {
    expect(kinds('=Table1[Amount]')).toContain('ref:Table1[Amount]');
  });

  it('assigns nesting depth to brackets', () => {
    const parens = tokenizeFormula('=IF(SUM(A1),1,2)').filter((t) => t.kind === 'paren');
    expect(parens.map((t) => `${t.text}${t.depth}`)).toEqual(['(0', '(1', ')1', ')0']);
  });
});

describe('matchBrackets', () => {
  it('pairs matching parentheses', () => {
    const tokens = tokenizeFormula('=IF(A1,SUM(B1:B2),0)');
    const { partner, unmatched } = matchBrackets(tokens);
    expect(unmatched).toEqual([]);
    const opens = tokens.map((t, i) => [t, i] as const).filter(([t]) => t.text === '(');
    expect(partner.get(opens[0]![1])).toBeGreaterThan(opens[1]![1]);
  });

  it('reports an unclosed parenthesis', () => {
    const { unmatched } = matchBrackets(tokenizeFormula('=SUM(A1'));
    expect(unmatched).toHaveLength(1);
  });

  it('reports a surplus closing parenthesis', () => {
    const { unmatched } = matchBrackets(tokenizeFormula('=SUM(A1))'));
    expect(unmatched).toHaveLength(1);
  });

  it('refuses to pair a brace with a parenthesis', () => {
    const { unmatched, partner } = matchBrackets(tokenizeFormula('={1,2)'));
    expect(partner.size).toBe(0);
    expect(unmatched).toHaveLength(2);
  });

  it('pairs braces', () => {
    const { unmatched } = matchBrackets(tokenizeFormula('={1,2}'));
    expect(unmatched).toEqual([]);
  });
});

describe('bracketPairAtCaret', () => {
  const src = '=IF(A1,1,2)';

  it('finds the pair when the caret is just after the opening bracket', () => {
    const tokens = tokenizeFormula(src);
    const pair = bracketPairAtCaret(tokens, src.indexOf('(') + 1);
    expect(pair).not.toBeNull();
    expect(tokens[pair!.close]!.start).toBe(src.length - 1);
  });

  it('finds the pair when the caret is on the closing bracket', () => {
    const tokens = tokenizeFormula(src);
    const pair = bracketPairAtCaret(tokens, src.length - 1);
    expect(pair).not.toBeNull();
    expect(tokens[pair!.open]!.text).toBe('(');
  });

  it('returns null away from any bracket', () => {
    expect(bracketPairAtCaret(tokenizeFormula(src), 2)).toBeNull();
  });

  it('returns null for an unmatched bracket', () => {
    const open = '=SUM(';
    expect(bracketPairAtCaret(tokenizeFormula(open), open.length)).toBeNull();
  });
});

describe('activeCall', () => {
  it('names the call the caret is inside and the argument index', () => {
    const src = '=SUM(A1,B1';
    const call = activeCall(tokenizeFormula(src), src.length);
    expect(call).toMatchObject({ name: 'SUM', argIndex: 1 });
  });

  it('reports argument zero straight after the bracket', () => {
    const src = '=SUM(';
    expect(activeCall(tokenizeFormula(src), src.length)?.argIndex).toBe(0);
  });

  it('reports the innermost call', () => {
    const src = '=IF(SUM(A1,';
    expect(activeCall(tokenizeFormula(src), src.length)?.name).toBe('SUM');
  });

  it('returns to the outer call after the inner one closes', () => {
    const src = '=IF(SUM(A1),';
    expect(activeCall(tokenizeFormula(src), src.length)?.name).toBe('IF');
  });

  it('returns null outside any call', () => {
    const src = '=A1+B1';
    expect(activeCall(tokenizeFormula(src), src.length)).toBeNull();
  });
});

describe('completionContext', () => {
  it('finds the partial word under the caret', () => {
    expect(completionContext('=SUMP', 5)).toEqual({ prefix: 'SUMP', start: 1, end: 5 });
  });

  it('returns null when the caret is not on a word', () => {
    expect(completionContext('=SUM(', 5)).toBeNull();
  });

  it('does not offer completions for a cell reference', () => {
    expect(completionContext('=A1', 3)).toBeNull();
  });

  it('does not offer completions for a number', () => {
    expect(completionContext('=123', 4)).toBeNull();
  });

  it('does not offer completions inside a string', () => {
    expect(completionContext('="SUM', 5)).toBeNull();
  });

  it('offers completions after a closed string', () => {
    expect(completionContext('="a"&SU', 7)?.prefix).toBe('SU');
  });
});

describe('applyCompletion', () => {
  it('replaces the partial word and opens a bracket for a function', () => {
    const context = completionContext('=SUMP', 5)!;
    expect(applyCompletion('=SUMP', context, 'SUMPRODUCT', true)).toEqual({
      text: '=SUMPRODUCT(',
      caret: 12,
    });
  });

  it('inserts a plain name without a bracket', () => {
    const context = completionContext('=Tot', 4)!;
    expect(applyCompletion('=Tot', context, 'Total', false)).toEqual({ text: '=Total', caret: 6 });
  });

  it('keeps text after the caret', () => {
    const context = completionContext('=SU+1', 3)!;
    expect(applyCompletion('=SU+1', context, 'SUM', true).text).toBe('=SUM(+1');
  });
});

describe('looksLikeCell', () => {
  it('accepts cell shapes', () => {
    expect(looksLikeCell('A1')).toBe(true);
    expect(looksLikeCell('$XFD$1048576')).toBe(true);
  });

  it('rejects non-cell shapes', () => {
    expect(looksLikeCell('SUM')).toBe(false);
    expect(looksLikeCell('ABCD1')).toBe(false);
  });
});
