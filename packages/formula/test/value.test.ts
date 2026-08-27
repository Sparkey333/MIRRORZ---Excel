import { describe, expect, it } from 'vitest';
import { CellError } from '@mirrorz/core';
import {
  compareScalars,
  excelAdd,
  excelNumbersEqual,
  excelSub,
  formatNumberForConcat,
  matchesCriterion,
  parseCriterion,
  parseNumericText,
  toBoolean,
  toExcelPrecision,
  toNumber,
  toText,
  truncateLiteral,
  wildcardToRegExp,
} from '../src/value.js';

describe('15-significant-digit precision', () => {
  it('collapses the classic cancellation residues onto the value Excel shows', () => {
    expect(toExcelPrecision(0.5 - 0.4)).toBe(0.1);
    expect(toExcelPrecision(0.1 + 0.2)).toBe(0.3);
    expect(toExcelPrecision(Math.SQRT2 ** 2)).toBe(2);
  });

  it('truncates literals on entry, as the oracle shows', () => {
    // =1.0000000000000002 stores as 1; =9007199254740993 stores as
    // 9007199254740990.
    expect(truncateLiteral(1.0000000000000002)).toBe(1);
    expect(truncateLiteral(0.9999999999999999)).toBe(1);
    expect(truncateLiteral(9007199254740993)).toBe(9007199254740990);
  });

  it('leaves ordinary values untouched', () => {
    for (const v of [0, 1, -1, 42, 3.14, 1e300, 1e-300, -2.5]) {
      expect(toExcelPrecision(v)).toBe(v);
    }
  });

  it('renders 1/3 at fifteen digits, matching the oracle', () => {
    expect(formatNumberForConcat(1 / 3)).toBe('0.333333333333333');
  });

  it('does not round after every operation', () => {
    // =1/3*3 is exactly 1. Rounding 1/3 to 15 digits first would give
    // 0.999999999999999, which Excel does not produce.
    expect((1 / 3) * 3).toBe(1);
    expect(toExcelPrecision((1 / 3) * 3)).toBe(1);
  });
});

describe('cancellation snap on add and subtract', () => {
  it.each([
    ['0.1+0.2-0.3', () => excelSub(excelAdd(0.1, 0.2), 0.3)],
    ['0.5-0.4-0.1', () => excelSub(excelSub(0.5, 0.4), 0.1)],
    ['0.3*3-0.9', () => excelSub(0.3 * 3, 0.9)],
  ])('%s is exactly zero', (_label, compute) => {
    expect(compute()).toBe(0);
  });

  it('snaps at the operation, not cosmetically at the end', () => {
    // The oracle probe =(0.1+0.2-0.3)*1E20 returns 0, which is only possible if
    // the residue is gone before the multiplication.
    expect(excelSub(excelAdd(0.1, 0.2), 0.3) * 1e20).toBe(0);
  });

  it('does not snap a genuine small difference', () => {
    // 0.5-0.4 is a real result near 0.1, nowhere near zero relative to 0.5.
    expect(excelSub(0.5, 0.4)).not.toBe(0);
    expect(excelSub(1, 0.9999)).toBeCloseTo(0.0001, 10);
  });

  it('does not snap when both operands are themselves tiny', () => {
    expect(excelAdd(1e-20, 1e-20)).toBe(2e-20);
    expect(excelSub(3e-300, 1e-300)).toBeCloseTo(2e-300, 310);
  });

  it('leaves exact zero and infinities alone', () => {
    expect(excelAdd(0, 0)).toBe(0);
    expect(excelAdd(1, -1)).toBe(0);
  });
});

describe('comparison at display precision', () => {
  it('makes the cases the oracle reports TRUE actually true', () => {
    expect(excelNumbersEqual(0.1 + 0.2, 0.3)).toBe(true);
    expect(excelNumbersEqual(0.5 - 0.4, 0.1)).toBe(true);
    expect(excelNumbersEqual(0.1 * 3, 0.3)).toBe(true);
    expect(excelNumbersEqual(Math.SQRT2 ** 2, 2)).toBe(true);
  });

  it('still separates genuinely different numbers', () => {
    expect(excelNumbersEqual(0.1, 0.2)).toBe(false);
    expect(excelNumbersEqual(1, 1.0000001)).toBe(false);
    expect(excelNumbersEqual(0, 1e-300)).toBe(false);
  });
});

describe('type ordering', () => {
  // Excel ranks number < text < FALSE < TRUE, so any text beats any number and
  // any boolean beats any text. LibreOffice disagrees on the boolean case; we
  // follow Excel. See docs/oracle-divergences.md.
  it('ranks text above numbers', () => {
    expect(compareScalars('a', 1)).toBeGreaterThan(0);
    expect(compareScalars(1, 'a')).toBeLessThan(0);
  });

  it('ranks booleans above text', () => {
    expect(compareScalars(true, 'z')).toBeGreaterThan(0);
    expect(compareScalars(false, 'z')).toBeGreaterThan(0);
    expect(compareScalars(true, 1)).toBeGreaterThan(0);
  });

  it('ranks TRUE above FALSE', () => {
    expect(compareScalars(true, false)).toBeGreaterThan(0);
  });

  it('compares text case-insensitively', () => {
    expect(compareScalars('abc', 'ABC')).toBe(0);
    expect(compareScalars('a', 'B')).toBeLessThan(0);
  });

  it('treats a blank as the zero of whatever it meets', () => {
    // Matches the oracle: =Z99="" and =Z99=0 are both TRUE.
    expect(compareScalars(null, '')).toBe(0);
    expect(compareScalars(null, 0)).toBe(0);
    expect(compareScalars(null, false)).toBe(0);
    expect(compareScalars(null, null)).toBe(0);
    expect(compareScalars(null, 5)).toBeLessThan(0);
  });

  it('propagates errors rather than ordering them', () => {
    expect(compareScalars(CellError.DIV0, 1)).toBe(CellError.DIV0);
    expect(compareScalars(1, CellError.NA)).toBe(CellError.NA);
  });
});

describe('coercion to number', () => {
  it.each([
    [null, 0],
    [42, 42],
    [true, 1],
    [false, 0],
    ['1', 1],
    ['  2.5  ', 2.5],
    ['-3', -3],
    ['+3', 3],
    ['1e3', 1000],
    ['50%', 0.5],
    ['(100)', -100],
    ['1,234', 1234],
    ['1,234,567', 1234567],
    ['', 0],
  ])('toNumber(%s) === %s', (input, want) => {
    expect(toNumber(input as never)).toBe(want);
  });

  it.each(['abc', '1abc', '1.2.3', '1,2,3', '--1', '1 2'])(
    'toNumber(%s) is #VALUE!',
    (input) => {
      expect(toNumber(input)).toBe(CellError.VALUE);
    },
  );

  it('passes errors straight through', () => {
    expect(toNumber(CellError.REF)).toBe(CellError.REF);
  });
});

describe('coercion to text', () => {
  it.each([
    [null, ''],
    ['x', 'x'],
    [true, 'TRUE'],
    [false, 'FALSE'],
    [42, '42'],
    [-0.5, '-0.5'],
    [1 / 3, '0.333333333333333'],
  ])('toText(%s) === %s', (input, want) => {
    expect(toText(input as never)).toBe(want);
  });
});

describe('coercion to boolean', () => {
  it.each([
    [null, false],
    [true, true],
    [0, false],
    [1, true],
    [-1, true],
    ['TRUE', true],
    ['false', false],
  ])('toBoolean(%s) === %s', (input, want) => {
    expect(toBoolean(input as never)).toBe(want);
  });

  it('refuses a numeric string, unlike arithmetic coercion', () => {
    // =IF("1",...) is #VALUE! in Excel even though ="1"+1 is 2.
    expect(toBoolean('1')).toBe(CellError.VALUE);
    expect(toBoolean('yes')).toBe(CellError.VALUE);
  });
});

describe('numeric text parsing', () => {
  it('accepts well-formed thousands groups only', () => {
    expect(parseNumericText('1,234')).toBe(1234);
    expect(parseNumericText('12,345,678')).toBe(12345678);
    expect(parseNumericText('1,23')).toBeUndefined();
    expect(parseNumericText('1,2345')).toBeUndefined();
  });

  it('handles nested percent and parens', () => {
    expect(parseNumericText('(50%)')).toBe(-0.5);
  });
});

describe('criteria', () => {
  it.each([
    ['5', '=', 5],
    ['>5', '>', 5],
    ['>=5', '>=', 5],
    ['<5', '<', 5],
    ['<=5', '<=', 5],
    ['<>5', '<>', 5],
    ['=5', '=', 5],
    ['apple', '=', 'apple'],
    ['<>apple', '<>', 'apple'],
    ['TRUE', '=', true],
  ])('parses %s', (raw, op, value) => {
    expect(parseCriterion(raw)).toMatchObject({ op, value });
  });

  it('matches numeric comparisons', () => {
    const gt5 = parseCriterion('>5');
    expect(matchesCriterion(6, gt5)).toBe(true);
    expect(matchesCriterion(5, gt5)).toBe(false);
    expect(matchesCriterion(4, gt5)).toBe(false);
  });

  it('matches text equality case-insensitively', () => {
    const apple = parseCriterion('apple');
    expect(matchesCriterion('Apple', apple)).toBe(true);
    expect(matchesCriterion('APPLE', apple)).toBe(true);
    expect(matchesCriterion('apples', apple)).toBe(false);
  });

  it('supports wildcards', () => {
    const starts = parseCriterion('app*');
    expect(matchesCriterion('apple', starts)).toBe(true);
    expect(matchesCriterion('application', starts)).toBe(true);
    expect(matchesCriterion('banana', starts)).toBe(false);

    const oneChar = parseCriterion('a?c');
    expect(matchesCriterion('abc', oneChar)).toBe(true);
    expect(matchesCriterion('ac', oneChar)).toBe(false);
    expect(matchesCriterion('abbc', oneChar)).toBe(false);
  });

  it('escapes wildcards with ~', () => {
    const literalStar = parseCriterion('a~*b');
    expect(matchesCriterion('a*b', literalStar)).toBe(true);
    expect(matchesCriterion('axb', literalStar)).toBe(false);
  });

  it('treats regex metacharacters as literals', () => {
    // A criterion is not a regular expression; "a.c" must not match "abc".
    const dotted = parseCriterion('a.c');
    expect(matchesCriterion('a.c', dotted)).toBe(true);
    expect(matchesCriterion('abc', dotted)).toBe(false);

    const parens = parseCriterion('(x)');
    expect(matchesCriterion('(x)', parens)).toBe(true);
  });

  it('matches blanks with a bare = criterion', () => {
    const blank = parseCriterion('=');
    expect(matchesCriterion(null, blank)).toBe(true);
    expect(matchesCriterion('', blank)).toBe(true);
    expect(matchesCriterion('x', blank)).toBe(false);
  });

  it('negates wildcard matches', () => {
    const notApp = parseCriterion('<>app*');
    expect(matchesCriterion('apple', notApp)).toBe(false);
    expect(matchesCriterion('banana', notApp)).toBe(true);
  });
});

describe('wildcard compilation', () => {
  it.each([
    ['*', 'anything', true],
    ['?', 'a', true],
    ['?', 'ab', false],
    ['a*', 'abc', true],
    ['*c', 'abc', true],
    ['*b*', 'abc', true],
    ['a?c', 'abc', true],
    ['', '', true],
    ['', 'x', false],
  ])('%s against %s is %s', (pattern, text, want) => {
    expect(wildcardToRegExp(pattern).test(text)).toBe(want);
  });

  it('matches across newlines, which a bare . would not', () => {
    expect(wildcardToRegExp('a*b').test('a\nb')).toBe(true);
  });
});
