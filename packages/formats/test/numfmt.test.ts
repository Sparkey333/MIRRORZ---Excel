import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CellError } from '@mirrorz/core';
import {
  BUILTIN_NUMBER_FORMATS,
  FIRST_CUSTOM_NUMFMT_ID,
  bestFraction,
  builtinFormatCode,
  fitToWidth,
  format,
  formatGeneral,
  isDateFormat,
  overflowText,
  parseFormat,
  toPlainDecimal,
} from '../src/numfmt.js';
import { XmlReader, XmlToken } from '../src/xml.js';
import { readZip } from '../src/zip.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

/** Shorthand: the display text only, which is what most assertions care about. */
const t = (value: unknown, code: string, dateSystem?: 1900 | 1904): string =>
  format(value as never, code, dateSystem ? { dateSystem } : {}).text;

// A Wednesday, chosen because make-fixtures.py uses it throughout styling.xlsx.
const MAR_15_2023 = 45000;

describe('built-in format table', () => {
  it('maps the language-neutral ids ECMA-376 reserves', () => {
    expect(BUILTIN_NUMBER_FORMATS[0]).toBe('General');
    expect(BUILTIN_NUMBER_FORMATS[1]).toBe('0');
    expect(BUILTIN_NUMBER_FORMATS[2]).toBe('0.00');
    expect(BUILTIN_NUMBER_FORMATS[3]).toBe('#,##0');
    expect(BUILTIN_NUMBER_FORMATS[4]).toBe('#,##0.00');
  });

  it('maps the percent, scientific and fraction ids', () => {
    expect(BUILTIN_NUMBER_FORMATS[9]).toBe('0%');
    expect(BUILTIN_NUMBER_FORMATS[10]).toBe('0.00%');
    expect(BUILTIN_NUMBER_FORMATS[11]).toBe('0.00E+00');
    expect(BUILTIN_NUMBER_FORMATS[12]).toBe('# ?/?');
    expect(BUILTIN_NUMBER_FORMATS[13]).toBe('# ??/??');
  });

  it('maps the date and time ids', () => {
    expect(BUILTIN_NUMBER_FORMATS[14]).toBe('mm-dd-yy');
    expect(BUILTIN_NUMBER_FORMATS[15]).toBe('d-mmm-yy');
    expect(BUILTIN_NUMBER_FORMATS[16]).toBe('d-mmm');
    expect(BUILTIN_NUMBER_FORMATS[17]).toBe('mmm-yy');
    expect(BUILTIN_NUMBER_FORMATS[18]).toBe('h:mm AM/PM');
    expect(BUILTIN_NUMBER_FORMATS[19]).toBe('h:mm:ss AM/PM');
    expect(BUILTIN_NUMBER_FORMATS[20]).toBe('h:mm');
    expect(BUILTIN_NUMBER_FORMATS[21]).toBe('h:mm:ss');
    expect(BUILTIN_NUMBER_FORMATS[22]).toBe('m/d/yy h:mm');
  });

  it('maps the accounting-style ids 37-40', () => {
    expect(BUILTIN_NUMBER_FORMATS[37]).toBe('#,##0 ;(#,##0)');
    expect(BUILTIN_NUMBER_FORMATS[38]).toBe('#,##0 ;[Red](#,##0)');
    expect(BUILTIN_NUMBER_FORMATS[39]).toBe('#,##0.00;(#,##0.00)');
    expect(BUILTIN_NUMBER_FORMATS[40]).toBe('#,##0.00;[Red](#,##0.00)');
  });

  it('maps the elapsed-time and text ids 45-49', () => {
    expect(BUILTIN_NUMBER_FORMATS[45]).toBe('mm:ss');
    expect(BUILTIN_NUMBER_FORMATS[46]).toBe('[h]:mm:ss');
    expect(BUILTIN_NUMBER_FORMATS[47]).toBe('mmss.0');
    expect(BUILTIN_NUMBER_FORMATS[48]).toBe('##0.0E+0');
    expect(BUILTIN_NUMBER_FORMATS[49]).toBe('@');
  });

  it('leaves the locale-specific and reserved ids unmapped', () => {
    for (const id of [5, 6, 7, 8, 23, 27, 41, 44, 50, 58]) {
      expect(builtinFormatCode(id)).toBeUndefined();
    }
  });

  it('starts custom ids at 164', () => {
    expect(FIRST_CUSTOM_NUMFMT_ID).toBe(164);
    expect(builtinFormatCode(FIRST_CUSTOM_NUMFMT_ID)).toBeUndefined();
  });

  it('renders every built-in code without throwing', () => {
    for (const code of Object.values(BUILTIN_NUMBER_FORMATS)) {
      expect(typeof t(1234.5678, code)).toBe('string');
      expect(typeof t(-0.5, code)).toBe('string');
      expect(typeof t(0, code)).toBe('string');
    }
  });
});

describe('General', () => {
  it('prints small integers as themselves', () => {
    expect(formatGeneral(0)).toBe('0');
    expect(formatGeneral(42)).toBe('42');
    expect(formatGeneral(-42)).toBe('-42');
  });

  it('drops trailing zeros', () => {
    expect(formatGeneral(1.5)).toBe('1.5');
    expect(formatGeneral(2.0)).toBe('2');
    expect(formatGeneral(1234.5678)).toBe('1234.5678');
  });

  it('shows 0.1 + 0.2 as 0.3', () => {
    expect(formatGeneral(0.1 + 0.2)).toBe('0.3');
    expect(t(0.1 + 0.2, 'General')).toBe('0.3');
  });

  it('trims to the 11-character budget', () => {
    expect(formatGeneral(1 / 3)).toBe('0.333333333');
    expect(formatGeneral(2 / 3)).toBe('0.666666667');
    expect(formatGeneral(123456789.123)).toBe('123456789.1');
  });

  it('keeps 11 digits before switching to exponential', () => {
    expect(formatGeneral(1e10)).toBe('10000000000');
    expect(formatGeneral(99999999999)).toBe('99999999999');
    expect(formatGeneral(1e11)).toBe('1E+11');
    expect(formatGeneral(123456789012)).toBe('1.23457E+11');
  });

  it('keeps small magnitudes fixed while they fit', () => {
    expect(formatGeneral(0.0001)).toBe('0.0001');
    expect(formatGeneral(1e-6)).toBe('0.000001');
    expect(formatGeneral(1e-9)).toBe('0.000000001');
  });

  it('goes exponential once leading zeros overflow the budget', () => {
    expect(formatGeneral(1e-10)).toBe('1E-10');
    expect(formatGeneral(1e-11)).toBe('1E-11');
    expect(formatGeneral(0.0000123456789)).toBe('1.23457E-05');
  });

  it('pads the exponent to two digits and uppercases E', () => {
    expect(formatGeneral(1e16)).toBe('1E+16');
    expect(formatGeneral(1e100)).toBe('1E+100');
    expect(formatGeneral(1e-100)).toBe('1E-100');
  });

  it('gives negatives one extra character of budget for the sign', () => {
    expect(formatGeneral(-1234.5678)).toBe('-1234.5678');
    expect(formatGeneral(-1e-9)).toBe('-0.000000001');
  });

  it('is reachable through the format entry point', () => {
    expect(t(1234.5678, 'General')).toBe('1234.5678');
    expect(t(1234.5678, 'general')).toBe('1234.5678');
  });

  it('keeps literals that sit alongside General', () => {
    expect(t(1.5, 'General" units"')).toBe('1.5 units');
  });

  it('reports #NUM! rather than Infinity', () => {
    expect(t(Number.POSITIVE_INFINITY, 'General')).toBe('#NUM!');
    expect(t(Number.NaN, '0.00')).toBe('#NUM!');
  });
});

describe('digit placeholders', () => {
  it('pads with zeros for 0 and spaces for ?', () => {
    expect(t(5, '000')).toBe('005');
    expect(t(5, '??0')).toBe('  5');
    expect(t(5, '###0')).toBe('5');
  });

  it('drops a zero value under # and ? but not under 0', () => {
    expect(t(0, '#')).toBe('');
    expect(t(0, '?')).toBe(' ');
    expect(t(0, '??')).toBe('  ');
    expect(t(0, '0')).toBe('0');
  });

  it('never truncates digits that overflow the pattern', () => {
    expect(t(12345, '0')).toBe('12345');
    expect(t(12345, '#')).toBe('12345');
  });

  it('drops insignificant decimals under # and blanks them under ?', () => {
    expect(t(0.5, '0.00')).toBe('0.50');
    expect(t(0.5, '0.##')).toBe('0.5');
    expect(t(0.5, '0.??')).toBe('0.5 ');
  });

  it('omits the separator when nothing follows it', () => {
    expect(t(5, '0.')).toBe('5');
    expect(t(5, '#.##')).toBe('5');
    expect(t(5, '0.??')).toBe('5.  ');
  });

  it('starts with the point when the integer part is all #', () => {
    expect(t(0.5, '#.##')).toBe('.5');
    expect(t(0.5, '?.?')).toBe(' .5');
    expect(t(0, '#.##')).toBe('');
  });

  it('matches the documented placeholder examples', () => {
    expect(t(1234.59, '####.#')).toBe('1234.6');
    expect(t(8.9, '#.000')).toBe('8.900');
    expect(t(0.631, '0.#')).toBe('0.6');
    expect(t(12.01234, '#.0#')).toBe('12.01');
  });

  it('fills placeholders in position, not all at the first one', () => {
    expect(t(5551234, '000-0000')).toBe('555-1234');
    expect(t(12, '(000)')).toBe('(012)');
    expect(t(1, '00-00')).toBe('00-01');
  });

  it('gives leftover digits to the leftmost placeholder', () => {
    expect(t(1234, '0-0')).toBe('123-4');
    expect(t(12, '0-0')).toBe('1-2');
  });
});

describe('thousands separator and scaling', () => {
  it('groups when a placeholder follows the comma', () => {
    expect(t(1234567.891, '#,##0.00')).toBe('1,234,567.89');
    expect(t(1000, '#,###')).toBe('1,000');
    expect(t(100, '#,###')).toBe('100');
  });

  it('groups the padded zeros too', () => {
    expect(t(5, '0,000')).toBe('0,005');
  });

  it('divides by 1000 for each trailing comma', () => {
    expect(t(12000, '#,')).toBe('12');
    expect(t(12200000, '0.0,,')).toBe('12.2');
    expect(t(1234.5, '0,,,')).toBe('0');
  });

  it('combines grouping with scaling', () => {
    expect(t(1500000, '#,##0,')).toBe('1,500');
    expect(t(1500000, '#,##0.0,')).toBe('1,500.0');
  });
});

describe('percent', () => {
  it('multiplies by 100 and shows the sign', () => {
    expect(t(0.05, '0%')).toBe('5%');
    expect(t(0.4567, '0.00%')).toBe('45.67%');
    expect(t(0.5, '0%')).toBe('50%');
  });

  it('scales once however many percent signs appear', () => {
    expect(t(0.015, '0%%')).toBe('2%%');
    expect(t(1.5, '0.0%%')).toBe('150.0%%');
  });

  it('routes negatives through the negative section', () => {
    expect(t(-0.5, '0.0%;-0.0%')).toBe('-50.0%');
  });
});

describe('scientific notation', () => {
  it('pins the mantissa when the integer pattern is all zeros', () => {
    expect(t(12345.678, '0.00E+00')).toBe('1.23E+04');
    expect(t(255, '0.00E+00')).toBe('2.55E+02');
    expect(t(0.000255, '0.00E+00')).toBe('2.55E-04');
  });

  it('steps the exponent in multiples when the pattern has # or ?', () => {
    expect(t(12200000, '##0.0E+0')).toBe('12.2E+6');
    expect(t(1234, '#0.0E+0')).toBe('12.3E+2');
  });

  it('omits the plus for E- formats', () => {
    expect(t(1234, '0.0E-0')).toBe('1.2E3');
    expect(t(0.001234, '0.0E-0')).toBe('1.2E-3');
    expect(t(0.001234, '0.0E+0')).toBe('1.2E-3');
  });

  it('shifts a decade when rounding overflows the mantissa', () => {
    expect(t(9.99, '0.0E+0')).toBe('1.0E+1');
  });

  it('renders zero without a fake exponent', () => {
    expect(t(0, '0.00E+00')).toBe('0.00E+00');
  });

  it('carries the sign for negatives', () => {
    expect(t(-1, '0.00E+00')).toBe('-1.00E+00');
  });
});

describe('fractions', () => {
  it('finds the best approximation for the denominator width', () => {
    expect(bestFraction(0.7, 9)).toEqual({ num: 5, den: 7 });
    expect(bestFraction(0.3, 99)).toEqual({ num: 3, den: 10 });
    expect(bestFraction(1 / 3, 9)).toEqual({ num: 1, den: 3 });
    expect(bestFraction(1 / 7, 999)).toEqual({ num: 1, den: 7 });
    expect(bestFraction(0, 9)).toEqual({ num: 0, den: 1 });
  });

  it('renders single-digit denominators', () => {
    expect(t(0.75, '# ?/?')).toBe(' 3/4');
    expect(t(5.25, '# ?/?')).toBe('5 1/4');
    expect(t(0.7, '# ?/?')).toBe(' 5/7');
  });

  it('renders two- and three-digit denominators with alignment padding', () => {
    expect(t(0.3, '# ??/??')).toBe('  3/10');
    expect(t(5.25, '# ???/???')).toBe('5   1/4  ');
    expect(t(5.3, '# ???/???')).toBe('5   3/10 ');
  });

  it('supports a fixed denominator', () => {
    expect(t(2.71828, '# ?/16')).toBe('2 11/16');
    expect(t(100.5, '# ?/2')).toBe('100 1/2');
    expect(t(0.26, '# ?/4')).toBe(' 1/4');
  });

  it('blanks the fraction when the remainder rounds away', () => {
    expect(t(5, '# ?/?')).toBe('5    ');
    expect(t(7, '# ?/?')).toBe('7    ');
    expect(t(0, '# ?/?')).toBe('0    ');
  });

  it('carries into the whole part rather than printing n/n', () => {
    expect(t(1.9999, '# ?/?')).toBe('2    ');
  });

  it('produces an improper fraction with no whole-number placeholder', () => {
    expect(t(0.75, '?/?')).toBe('3/4');
    expect(t(5.25, '?/?')).toBe('21/4');
  });

  it('takes a sign from the single section', () => {
    expect(t(-2.5, '# ?/?')).toBe('-2 1/2');
  });
});

describe('section fallback rules', () => {
  it('applies one section to every number', () => {
    expect(t(5, '0.00')).toBe('5.00');
    expect(t(-5, '0.00')).toBe('-5.00');
    expect(t(0, '0.00')).toBe('0.00');
  });

  it('splits two sections at zero, with zero taking the positive one', () => {
    expect(t(5, '0.00;(0.00)')).toBe('5.00');
    expect(t(-5, '0.00;(0.00)')).toBe('(5.00)');
    expect(t(0, '0.00;(0.00)')).toBe('0.00');
  });

  it('adds no minus when a negative section supplies the sign itself', () => {
    expect(t(-1234, '#,##0 ;(#,##0)')).toBe('(1,234)');
    expect(t(-42, '0.00;[Red]-0.00')).toBe('-42.00');
    expect(t(-1234.5678, '#,##0.00_);[Red](#,##0.00)')).toBe('(1,234.57)');
  });

  it('gives zero its own section when three are present', () => {
    expect(t(5, '"a";"b";"c"')).toBe('a');
    expect(t(-5, '"a";"b";"c"')).toBe('b');
    expect(t(0, '"a";"b";"c"')).toBe('c');
  });

  it('renders an empty section as nothing', () => {
    expect(t(-1234.5, '#,##0.00;;"zed"')).toBe('');
    expect(t(0, '#,##0.00;;"zed"')).toBe('zed');
  });

  it('treats a trailing @ section as the text section, not the negative one', () => {
    expect(t(-5, '0.00;@')).toBe('-5.00');
    expect(t('raw', '0.00;@')).toBe('raw');
    expect(t(MAR_15_2023, 'mm-dd-yy;@')).toBe('03-15-23');
  });

  it('suppresses the minus once the value rounds to zero', () => {
    expect(t(-0.0001, '0.00')).toBe('0.00');
    expect(t(-0.5, '#,##0.00')).toBe('-0.50');
  });

  it('never signs a literal-only section', () => {
    expect(t(-5, '"big"')).toBe('big');
  });
});

describe('conditional sections', () => {
  it('picks the first section whose condition holds', () => {
    expect(t(150, '[>100]"big";[<=100]"small"')).toBe('big');
    expect(t(50, '[>100]"big";[<=100]"small"')).toBe('small');
    expect(t(-5, '[>100]"big";[<=100]"small"')).toBe('small');
  });

  it('understands every comparison operator', () => {
    expect(t(5, '[=5]"eq";"ne"')).toBe('eq');
    expect(t(6, '[=5]"eq";"ne"')).toBe('ne');
    expect(t(6, '[<>5]"ne";"eq"')).toBe('ne');
    expect(t(5, '[>=5]"ge";"lt"')).toBe('ge');
    expect(t(4, '[<5]"lt";"ge"')).toBe('lt');
  });

  it('falls back to the first unconditioned section', () => {
    expect(t(5, '[>100]"big";0.00')).toBe('5.00');
  });

  it('overflows to # when nothing matches', () => {
    const r = format(5, '[>100]0.00');
    expect(r.overflow).toBe(true);
    expect(fitToWidth(r, 6)).toBe('######');
  });

  it('leaves the sign to the section once conditions choose it', () => {
    expect(t(-5, '[>100]0.00;[<=100]0.00')).toBe('5.00');
  });

  it('keeps the automatic minus for a lone conditional section', () => {
    expect(t(-5, '[<=100]0.00;@')).toBe('-5.00');
  });
});

describe('colours', () => {
  it('reads the eight names', () => {
    expect(format(-1, '0.00;[Red]-0.00').color).toEqual({ name: 'Red' });
    expect(format(1, '[Blue]0.00').color).toEqual({ name: 'Blue' });
    expect(format(1, '[green]0').color).toEqual({ name: 'Green' });
    expect(format(1, '[MAGENTA]0').color).toEqual({ name: 'Magenta' });
  });

  it('reads the indexed form', () => {
    expect(format(1, '[Color12]0').color).toEqual({ index: 12 });
    expect(format(1, '[Color56]0').color).toEqual({ index: 56 });
  });

  it('ignores an out-of-range index', () => {
    expect(format(1, '[Color57]0').color).toBeUndefined();
  });

  it('reports no colour when the section has none', () => {
    expect(format(1, '0.00').color).toBeUndefined();
    expect(format(1, '0.00;[Red]-0.00').color).toBeUndefined();
  });

  it('does not leak the colour bracket into the output', () => {
    expect(t(-1234, '#,##0 ;[Red](#,##0)')).toBe('(1,234)');
  });
});

describe('literals, escapes and spacing', () => {
  it('prints quoted text', () => {
    expect(t(1234.5, '"$"#,##0.00')).toBe('$1,234.50');
    expect(t(0.5, '0.0" ea"')).toBe('0.5 ea');
  });

  it('prints a backslash-escaped character without the backslash', () => {
    expect(t(1234.5, '0\\!')).toBe('1235!');
    expect(t(MAR_15_2023, 'd\\-mmm\\-yy')).toBe('15-Mar-23');
    expect(t(MAR_15_2023, '\\y\\ymmdd')).toBe('yy0315');
  });

  it('prints the self-literal characters', () => {
    expect(t(12, '(000)')).toBe('(012)');
    expect(t(5, '+0')).toBe('+5');
    expect(t(5, '0:0')).toBe('0:5');
  });

  it('emits one space for a width skip', () => {
    expect(t(1234.5, '_(0.0_)')).toBe(' 1234.5 ');
  });

  it('consumes the fill character without corrupting the output', () => {
    expect(t(3, '0*x')).toBe('3');
    expect(t(3, '0*x0')).toBe('03');
  });

  it('prints the currency string from a locale bracket', () => {
    expect(t(1234.5, '[$USD-409]#,##0.00')).toBe('USD1,234.50');
    expect(t(1234.5, '[$€-407]#,##0.00')).toBe('€1,234.50');
    expect(t(MAR_15_2023, '[$-409]yyyy')).toBe('2023');
  });

  it('keeps non-ASCII literals', () => {
    expect(t(MAR_15_2023, 'yyyy"年"m"月"')).toBe('2023年3月');
  });

  it('hides a semicolon inside quotes from the section splitter', () => {
    expect(parseFormat('"a;b"').sections).toHaveLength(1);
    expect(t(1, '"a;b"0')).toBe('a;b1');
  });
});

describe('month versus minute', () => {
  it('reads a lone m as a month', () => {
    expect(t(MAR_15_2023, 'm')).toBe('3');
    expect(t(MAR_15_2023, 'mm')).toBe('03');
  });

  it('reads m after an hour token as minutes', () => {
    expect(t(0.5678, 'h:mm')).toBe('13:38');
    expect(t(0.5678, 'h m s')).toBe('13 37 38');
  });

  it('reads m before a seconds token as minutes', () => {
    expect(t(0.5678, 'mm:ss')).toBe('37:38');
    expect(t(0.5678, 'mmss.0')).toBe('3737.9');
  });

  it('reads m between date tokens as a month', () => {
    expect(t(MAR_15_2023, 'yyyy-mm-dd')).toBe('2023-03-15');
    expect(t(MAR_15_2023, 'm/d/yy')).toBe('3/15/23');
  });

  it('resolves both meanings in one format', () => {
    expect(t(45000.5678, 'yyyy-mm-dd hh:mm')).toBe('2023-03-15 13:38');
    expect(t(45000.5678, 'm/d/yy h:mm')).toBe('3/15/23 13:38');
  });

  it('treats three or more m as a month name regardless of neighbours', () => {
    expect(t(45000.5, 'h:mmm')).toBe('12:Mar');
    expect(t(MAR_15_2023, 'mmm-yy')).toBe('Mar-23');
    expect(t(MAR_15_2023, 'mmmm')).toBe('March');
    expect(t(MAR_15_2023, 'mmmmm')).toBe('M');
  });

  it('looks past literals when deciding', () => {
    expect(t(0.5678, 'h"x"mm')).toBe('13x38');
  });

  it('reads m after an elapsed-hour bracket as minutes', () => {
    expect(t(1.5, '[h]:mm')).toBe('36:00');
  });

  it('reads m before an elapsed-second bracket as minutes', () => {
    expect(t(1.5, 'm:[s]')).toBe('0:129600');
  });

  it('is not fooled by a quoted m', () => {
    expect(isDateFormat('0.0"m"')).toBe(false);
    expect(t(1.5, '0.0"m"')).toBe('1.5m');
  });
});

describe('date and time tokens', () => {
  it('renders years at two or four digits', () => {
    expect(t(MAR_15_2023, 'yy')).toBe('23');
    expect(t(MAR_15_2023, 'yyyy')).toBe('2023');
  });

  it('renders day numbers and day names', () => {
    expect(t(MAR_15_2023, 'd')).toBe('15');
    expect(t(MAR_15_2023, 'dd')).toBe('15');
    expect(t(MAR_15_2023, 'ddd')).toBe('Wed');
    expect(t(MAR_15_2023, 'dddd')).toBe('Wednesday');
  });

  it('renders the built-in date formats', () => {
    expect(t(MAR_15_2023, 'mm-dd-yy')).toBe('03-15-23');
    expect(t(MAR_15_2023, 'd-mmm-yy')).toBe('15-Mar-23');
    expect(t(MAR_15_2023, 'd-mmm')).toBe('15-Mar');
    expect(t(MAR_15_2023, 'dddd, mmmm d, yyyy')).toBe('Wednesday, March 15, 2023');
  });

  it('uses a 24-hour clock without AM/PM', () => {
    expect(t(0.5678, 'h:mm:ss')).toBe('13:37:38');
    expect(t(0.5678, 'hh:mm:ss')).toBe('13:37:38');
    expect(t(0.25, 'h:mm:ss')).toBe('6:00:00');
  });

  it('switches to a 12-hour clock for AM/PM', () => {
    expect(t(0.5678, 'h:mm AM/PM')).toBe('1:38 PM');
    expect(t(0.5678, 'hh AM/PM')).toBe('02 PM');
    expect(t(0.5, 'h:mm AM/PM')).toBe('12:00 PM');
    expect(t(0, 'h:mm AM/PM')).toBe('12:00 AM');
  });

  it('switches to a 12-hour clock for A/P', () => {
    expect(t(0.5, 'h:mm A/P')).toBe('12:00 P');
    expect(t(0.2, 'h:mm A/P')).toBe('4:48 A');
    expect(t(0.2, 'h:mm a/p')).toBe('4:48 a');
  });

  it('renders fractional seconds', () => {
    expect(t(0.5678, 'h:mm:ss.0')).toBe('13:37:37.9');
    expect(t(0.123456, 'hh:mm:ss.00')).toBe('02:57:46.60');
    expect(t(0.5678, 'h:mm:ss.000')).toBe('13:37:37.920');
  });

  it('rounds the time to the precision it shows', () => {
    expect(t(0.9999999, 'hh:mm:ss')).toBe('00:00:00');
    expect(t(0.9999999, 'hh:mm')).toBe('00:00');
    expect(t(45000.99998, 'yyyy-mm-dd hh:mm:ss')).toBe('2023-03-15 23:59:58');
  });

  it('reproduces the 1900 phantom leap day', () => {
    expect(t(59, 'yyyy-mm-dd')).toBe('1900-02-28');
    expect(t(60, 'yyyy-mm-dd')).toBe('1900-02-29');
    expect(t(61, 'yyyy-mm-dd')).toBe('1900-03-01');
  });

  it('honours the 1904 date system', () => {
    expect(t(0, 'yyyy-mm-dd', 1904)).toBe('1904-01-01');
    expect(t(MAR_15_2023, 'yyyy-mm-dd', 1904)).toBe('2027-03-16');
  });

  it('overflows for a negative serial', () => {
    const r = format(-1, 'yyyy-mm-dd');
    expect(r.overflow).toBe(true);
    expect(fitToWidth(r, 4)).toBe('####');
  });

  it('overflows past year 9999', () => {
    expect(format(3_000_000, 'yyyy-mm-dd').overflow).toBe(true);
  });

  it('accepts a Date object', () => {
    expect(t(new Date(Date.UTC(2023, 2, 15)), 'yyyy-mm-dd')).toBe('2023-03-15');
  });

  it('renders era tokens without crashing', () => {
    expect(t(MAR_15_2023, '[$-411]ge.m.d')).toBe('2023.3.15');
    expect(() => t(MAR_15_2023, 'ggge"y"')).not.toThrow();
  });

  it('drops the calendar selector rather than printing it', () => {
    expect(t(MAR_15_2023, 'b1yyyy')).toBe('2023');
  });
});

describe('elapsed time brackets', () => {
  it('does not wrap at the natural boundary', () => {
    expect(t(1.5, '[h]:mm:ss')).toBe('36:00:00');
    expect(t(1.5, '[mm]:ss')).toBe('2160:00');
    expect(t(1.5, '[ss].00')).toBe('129600.00');
  });

  it('pads to the bracket width but never truncates', () => {
    expect(t(0.5, '[h]:mm')).toBe('12:00');
    expect(t(0.04, '[hh]:mm')).toBe('00:58');
  });

  it('handles very long durations', () => {
    expect(t(100000, '[h]:mm:ss')).toBe('2400000:00:00');
  });

  it('reports elapsed seconds', () => {
    expect(t(0.5678, '[s]')).toBe('49058');
  });

  it('counts as a date format', () => {
    expect(isDateFormat('[h]:mm:ss')).toBe(true);
  });
});

describe('rounding and 15 significant digits', () => {
  it('rounds half away from zero on the decimal expansion', () => {
    expect(t(1.005, '0.00')).toBe('1.01');
    expect(t(2.675, '0.00')).toBe('2.68');
    expect(t(0.45, '0.0')).toBe('0.5');
  });

  it('rounds the integer form too', () => {
    expect(t(1234.5678, '0')).toBe('1235');
    expect(t(0.5, '0')).toBe('1');
  });

  it('clamps to 15 significant digits before padding', () => {
    expect(t(0.1 + 0.2, '0.00000000000000000')).toBe('0.30000000000000000');
    expect(t(1234.5678, '0.000000000000000000')).toBe('1234.567800000000000000');
  });

  it('expands a double to a plain decimal string', () => {
    expect(toPlainDecimal(0.1 + 0.2)).toBe('0.300000000000000');
    expect(toPlainDecimal(1e-7)).toBe('0.000000100000000000000');
    expect(toPlainDecimal(1e21)).toBe('1000000000000000000000');
  });

  it('carries a rounding overflow through the integer part', () => {
    expect(t(9.999, '0.00')).toBe('10.00');
    expect(t(0.999, '0.00')).toBe('1.00');
  });
});

describe('text values', () => {
  it('leaves text untouched when the format has no text section', () => {
    expect(t('raw', '0.00')).toBe('raw');
    expect(t('txt', '"pre"')).toBe('txt');
  });

  it('substitutes @ with the text', () => {
    expect(t('raw', '@')).toBe('raw');
    expect(t('June', '"gross receipts for "@')).toBe('gross receipts for June');
  });

  it('uses the fourth section for text', () => {
    expect(t('raw', '0.00;[Red]-0.00;"zero";@')).toBe('raw');
    expect(t('raw', '0.00;[Red]-0.00;"zero";"was: "@')).toBe('was: raw');
  });

  it('shows only the literals when the text section omits @', () => {
    expect(t('txt', '0.00;-0.00;"z";"lit"')).toBe('lit');
  });

  it('formats a number under the Text format as General', () => {
    expect(t(123, '@')).toBe('123');
    expect(t(0.1 + 0.2, '@')).toBe('0.3');
  });

  it('reports text as non-numeric so it never becomes ####', () => {
    const r = format('a long piece of text', '@');
    expect(r.numeric).toBe(false);
    expect(fitToWidth(r, 3)).toBe('a long piece of text');
  });
});

describe('non-numeric values', () => {
  it('renders booleans as their uppercase names', () => {
    expect(t(true, '0.00')).toBe('TRUE');
    expect(t(false, 'yyyy-mm-dd')).toBe('FALSE');
  });

  it('renders errors as their codes', () => {
    expect(t(CellError.DIV0, '0.00')).toBe('#DIV/0!');
    expect(t(new CellError('#N/A'), '@')).toBe('#N/A');
  });

  it('renders empty as the empty string', () => {
    expect(t(null, '0.00')).toBe('');
    expect(t(undefined, 'General')).toBe('');
  });

  it('marks non-numbers as non-numeric', () => {
    expect(format(true, '0.00').numeric).toBe(false);
    expect(format(CellError.NA, '0.00').numeric).toBe(false);
  });
});

describe('isDateFormat', () => {
  it('is true for date and time codes', () => {
    expect(isDateFormat('yyyy-mm-dd')).toBe(true);
    expect(isDateFormat('h:mm:ss')).toBe(true);
    expect(isDateFormat('mm-dd-yy;@')).toBe(true);
    expect(isDateFormat('[h]:mm:ss')).toBe(true);
    expect(isDateFormat('d-mmm')).toBe(true);
  });

  it('is false for numeric and text codes', () => {
    expect(isDateFormat('General')).toBe(false);
    expect(isDateFormat('0.00')).toBe(false);
    expect(isDateFormat('#,##0.00')).toBe(false);
    expect(isDateFormat('0.00E+00')).toBe(false);
    expect(isDateFormat('# ??/??')).toBe(false);
    expect(isDateFormat('@')).toBe(false);
  });

  it('is not fooled by quoted or escaped date letters', () => {
    expect(isDateFormat('0.00" days"')).toBe(false);
    expect(isDateFormat('0\\d')).toBe(false);
  });

  it('is not fooled by a locale bracket', () => {
    expect(isDateFormat('[$-409]#,##0.00')).toBe(false);
    expect(isDateFormat('[$USD-409]#,##0.00')).toBe(false);
  });

  it('agrees with every built-in id', () => {
    const dateIds = [14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47];
    for (const [id, code] of Object.entries(BUILTIN_NUMBER_FORMATS)) {
      expect(isDateFormat(code)).toBe(dateIds.includes(Number(id)));
    }
  });
});

describe('column-width overflow', () => {
  it('builds a fill of the requested width', () => {
    expect(overflowText(5)).toBe('#####');
    expect(overflowText(0)).toBe('#');
    expect(overflowText(3.7)).toBe('###');
  });

  it('replaces a number that does not fit', () => {
    const r = format(1234567.891, '#,##0.00');
    expect(r.text).toBe('1,234,567.89');
    expect(fitToWidth(r, 12)).toBe('1,234,567.89');
    expect(fitToWidth(r, 8)).toBe('########');
  });

  it('never replaces text', () => {
    expect(fitToWidth(format('hello', '@'), 2)).toBe('hello');
  });

  it('is not baked into format itself', () => {
    expect(format(1234567.891, '#,##0.00').text).not.toContain('#');
  });
});

describe('parser robustness', () => {
  it('caches and reuses a compiled format', () => {
    expect(parseFormat('0.00')).toBe(parseFormat('0.00'));
  });

  it('survives an unterminated quote', () => {
    expect(() => t(1, '0"abc')).not.toThrow();
    expect(t(1, '0"abc')).toBe('1abc');
  });

  it('survives an unterminated bracket', () => {
    expect(() => t(1, '[Red0.00')).not.toThrow();
  });

  it('survives a trailing backslash', () => {
    expect(() => t(1, '0\\')).not.toThrow();
  });

  it('survives an empty code', () => {
    expect(t(1, '')).toBe('');
    expect(isDateFormat('')).toBe(false);
  });

  it('ignores brackets it does not implement', () => {
    expect(t(1, '[DBNum1]0')).toBe('1');
    expect(t(1, '[ENG]0')).toBe('1');
  });

  it('handles more than four sections without crashing', () => {
    expect(() => t(1, '0;0;0;@;0')).not.toThrow();
  });
});

describe('styling.xlsx fixture', () => {
  /** The number formats make-fixtures.py deliberately spread across column E. */
  const numFmts = (): Map<number, string> => {
    const zip = readZip(
      new Uint8Array(readFileSync(new URL('styling.xlsx', FIXTURES))),
    );
    const xml = new TextDecoder().decode(zip.get('xl/styles.xml')!.data());
    const reader = new XmlReader(xml);
    const out = new Map<number, string>();
    for (let tok = reader.next(); tok !== XmlToken.EOF; tok = reader.next()) {
      if (tok === XmlToken.Open && reader.localName === 'numFmt') {
        out.set(Number(reader.attr('numFmtId')), reader.attr('formatCode') ?? '');
      }
    }
    return out;
  };

  it('carries the custom formats the fixture was authored with', () => {
    const codes = [...numFmts().values()];
    expect(codes).toContain('"$"#,##0.00');
    expect(codes).toContain('yyyy-mm-dd');
    expect(codes).toContain('dddd, mmmm d, yyyy');
    expect(codes).toContain('0.00;[Red]-0.00;"zero";@');
    expect(codes).toContain('[>100]"big";[<=100]"small"');
  });

  it('renders each fixture value the way the fixture intends', () => {
    const codes = numFmts();
    const code = (id: number): string => codes.get(id)!;
    expect(t(1234.5, code(164))).toBe('$1,234.50');
    expect(t(-1234.5, code(165))).toBe(' $(1,234.50)');
    expect(t(MAR_15_2023, code(166))).toBe('2023-03-15');
    expect(t(MAR_15_2023, code(167))).toBe('Wednesday, March 15, 2023');
    expect(t(0.5678, code(168))).toBe('13:37:38');
    expect(t(45000.5678, code(169))).toBe('2023-03-15 13:38');
    expect(t(-42, code(170))).toBe('-42.00');
    expect(t(0, code(171))).toBe('zero');
    expect(t(150, code(172))).toBe('big');
  });

  it('classifies the fixture formats by date-ness', () => {
    const codes = numFmts();
    expect(isDateFormat(codes.get(166)!)).toBe(true);
    expect(isDateFormat(codes.get(167)!)).toBe(true);
    expect(isDateFormat(codes.get(168)!)).toBe(true);
    expect(isDateFormat(codes.get(169)!)).toBe(true);
    expect(isDateFormat(codes.get(164)!)).toBe(false);
    expect(isDateFormat(codes.get(165)!)).toBe(false);
  });

  it('agrees with the recalculated twin on the same codes', () => {
    const zip = readZip(
      new Uint8Array(readFileSync(new URL('styling.calc.xlsx', FIXTURES))),
    );
    const xml = new TextDecoder().decode(zip.get('xl/styles.xml')!.data());
    const reader = new XmlReader(xml);
    const codes: string[] = [];
    for (let tok = reader.next(); tok !== XmlToken.EOF; tok = reader.next()) {
      if (tok === XmlToken.Open && reader.localName === 'numFmt') {
        codes.push(reader.attr('formatCode') ?? '');
      }
    }
    // LibreOffice rewrites the table, so only the codes it kept are compared.
    expect(codes.length).toBeGreaterThan(0);
    for (const c of codes) expect(() => format(45000.5, c)).not.toThrow();
  });
});
