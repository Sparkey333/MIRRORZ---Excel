import { describe, expect, it } from 'vitest';
import { originalText, parseEntry, reviewImport } from '../src/entry.js';
import { serialToParts } from '../src/serial.js';

const parse = (s: string, o = {}) => parseEntry(s, o);

describe('the corruption cases', () => {
  // Each of these is a real, documented way Excel destroys data on entry. They
  // are the reason this module exists, so they are the first tests.

  it('keeps gene symbols as text', () => {
    for (const gene of ['SEPT1', 'MARCH1', 'DEC1', 'OCT4', 'MARC1', 'SEP15']) {
      const r = parse(gene);
      expect(r.value, gene).toBe(gene);
      expect(r.kind).toBe('text');
      expect(r.confidence).toBe('risky');
      expect(r.note).toContain('gene');
    }
  });

  it('keeps leading zeros', () => {
    for (const v of ['007', '01234', '0001']) {
      const r = parse(v);
      expect(r.value, v).toBe(v);
      expect(r.kind).toBe('text');
      expect(r.note).toContain('leading zero');
    }
  });

  it('keeps long digit strings whole', () => {
    // A nineteen-digit order number loses its low digits to a float. Excel
    // shows 1.23457E+18 and the original is gone for good.
    const id = '1234567890123456789';
    const r = parse(id);
    expect(r.value).toBe(id);
    expect(r.kind).toBe('text');
    expect(r.note).toContain('15 digits');
  });

  it('keeps ranges and score lines as text', () => {
    for (const v of ['1-2', '3-4', '10-12']) {
      const r = parse(v);
      expect(r.value, v).toBe(v);
      expect(r.kind).toBe('text');
    }
  });

  it('keeps version numbers as text', () => {
    expect(parse('1.2.3').value).toBe('1.2.3');
    expect(parse('10.15.7').kind).toBe('text');
  });

  it('honours a leading apostrophe, and keeps it in the literal', () => {
    const r = parse("'123");
    expect(r.value).toBe('123');
    expect(r.kind).toBe('text');
    expect(r.literal).toBe("'123");
  });

  it('honours a column type lock', () => {
    const r = parse('2024-01-15', { forceText: true });
    expect(r.value).toBe('2024-01-15');
    expect(r.kind).toBe('text');
    expect(r.note).toContain('locked');
  });
});

describe('numbers', () => {
  it.each([
    ['42', 42],
    ['-17', -17],
    ['3.14', 3.14],
    ['.5', 0.5],
    ['1e3', 1000],
    ['1.5E-10', 1.5e-10],
    ['+7', 7],
  ])('reads %s as %s', (input, want) => {
    const r = parse(input);
    expect(r.value).toBe(want);
    expect(r.kind).toBe('number');
  });

  it('reads percentages and implies the format', () => {
    const r = parse('45%');
    expect(r.value).toBeCloseTo(0.45, 12);
    expect(r.impliedFormat).toBe('0.00%');
  });

  it('reads accounting parentheses as negative', () => {
    expect(parse('(100)').value).toBe(-100);
  });

  it('reads well-formed thousands groups', () => {
    expect(parse('1,234').value).toBe(1234);
    expect(parse('12,345,678').value).toBe(12345678);
  });

  it('refuses malformed thousands groups rather than guessing', () => {
    // "1,2,3" quietly becoming 123 would be exactly the kind of silent damage
    // this module exists to prevent.
    expect(parse('1,2,3').kind).toBe('text');
    expect(parse('1,23').kind).toBe('text');
  });

  it('reads a currency prefix and implies a currency format', () => {
    const r = parse('$1,234.50');
    expect(r.value).toBe(1234.5);
    expect(r.impliedFormat).toContain('$');
  });

  it('keeps the literal when the number does not render back identically', () => {
    expect(parse('1,234').literal).toBe('1,234');
    expect(parse('45%').literal).toBe('45%');
    // A plain number needs no literal, since the value reproduces it exactly.
    expect(parse('42').literal).toBeUndefined();
  });

  it('can be switched off entirely', () => {
    expect(parse('42', { inferNumbers: false }).kind).toBe('text');
  });
});

describe('dates', () => {
  it('reads an ISO date', () => {
    const r = parse('2024-02-29');
    expect(r.kind).toBe('date');
    expect(serialToParts(r.value as number)).toMatchObject({ year: 2024, month: 2, day: 29 });
    expect(r.impliedFormat).toBe('yyyy-mm-dd');
    expect(r.literal).toBe('2024-02-29');
  });

  it('reads an ISO datetime', () => {
    const r = parse('2024-02-29 13:45:30');
    expect(r.kind).toBe('datetime');
    const p = serialToParts(r.value as number);
    expect(p).toMatchObject({ year: 2024, month: 2, day: 29, hour: 13, minute: 45, second: 30 });
  });

  it('reads a time of day as a fraction', () => {
    expect(parse('12:00').value).toBeCloseTo(0.5, 10);
    expect(parse('06:00').value).toBeCloseTo(0.25, 10);
  });

  it('reads a 12-hour clock', () => {
    expect(parse('1:00 PM').value).toBeCloseTo(13 / 24, 10);
    expect(parse('12:00 AM').value).toBeCloseTo(0, 10);
    expect(parse('12:00 PM').value).toBeCloseTo(0.5, 10);
  });

  it('rejects an impossible date rather than rolling it over', () => {
    // Excel's DATE() rolls 2023-02-30 into March. On ENTRY that is a typo, and
    // silently moving it is worse than leaving the text alone.
    expect(parse('2023-02-30').kind).toBe('text');
    expect(parse('2024-13-01').kind).toBe('text');
  });

  it('flags an ambiguous numeric date instead of guessing silently', () => {
    // 3/4/2024 is 3 April in most of the world and 4 March in the US.
    const r = parse('3/4/2024');
    expect(r.kind).toBe('date');
    expect(r.confidence).toBe('ambiguous');
    expect(r.note).toContain('could be');
  });

  it('resolves confidently once the order is stated', () => {
    const us = parse('3/4/2024', { dateOrder: 'mdy' });
    expect(us.confidence).toBe('certain');
    expect(serialToParts(us.value as number)).toMatchObject({ month: 3, day: 4 });

    const uk = parse('3/4/2024', { dateOrder: 'dmy' });
    expect(serialToParts(uk.value as number)).toMatchObject({ month: 4, day: 3 });
  });

  it('is not ambiguous when only one reading is possible', () => {
    // 25 cannot be a month.
    expect(parse('25/12/2024', { dateOrder: 'dmy' }).confidence).toBe('certain');
    expect(parse('12/25/2024').confidence).toBe('certain');
  });

  it('can be switched off entirely', () => {
    expect(parse('2024-01-15', { inferDates: false }).kind).toBe('text');
  });
});

describe('booleans and blanks', () => {
  it.each([
    ['TRUE', true],
    ['true', true],
    ['FALSE', false],
    ['False', false],
  ])('reads %s', (input, want) => {
    expect(parse(input).value).toBe(want);
  });

  it('reads an empty string as blank', () => {
    expect(parse('').value).toBe(null);
    expect(parse('').kind).toBe('blank');
  });

  it('can be switched off', () => {
    expect(parse('TRUE', { inferBooleans: false }).kind).toBe('text');
  });
});

describe('reversibility', () => {
  it('recovers the original text from the stored literal', () => {
    for (const input of ['1,234', '45%', '2024-02-29', "'007", '$1,234.50']) {
      const r = parse(input);
      expect(originalText(r.value, r.literal), input).toBe(input);
    }
  });

  it('recovers text that needed no literal', () => {
    const r = parse('42');
    expect(originalText(r.value, r.literal)).toBe('42');
  });

  it('recovers protected text unchanged', () => {
    const r = parse('SEPT1');
    expect(originalText(r.value, r.literal)).toBe('SEPT1');
  });
});

describe('import review', () => {
  const rows = [
    ['id', 'gene', 'date', 'amount'],
    ['007', 'SEPT1', '2024-01-15', '1,234.50'],
    ['008', 'MARCH1', '3/4/2024', '2,000'],
    ['009', 'BRCA1', '2024-03-01', '42'],
  ];

  it('counts what would be converted before anything is committed', () => {
    const review = reviewImport(rows);
    expect(review.total).toBe(16);
    expect(review.converted).toBeGreaterThan(0);
  });

  it('lists the values it protected, with reasons', () => {
    const review = reviewImport(rows);
    const inputs = review.protected.map((p) => p.input);
    expect(inputs).toContain('007');
    expect(inputs).toContain('SEPT1');
    expect(inputs).toContain('MARCH1');
    expect(review.protected[0]!.result.note).toBeTruthy();
  });

  it('lists ambiguous conversions for review', () => {
    const review = reviewImport(rows);
    expect(review.ambiguous.map((a) => a.input)).toContain('3/4/2024');
  });

  it('reports positions so the UI can highlight them', () => {
    const review = reviewImport(rows);
    const gene = review.protected.find((p) => p.input === 'SEPT1')!;
    expect(gene.row).toBe(1);
    expect(gene.col).toBe(1);
  });

  it('has nothing to report once the order is stated and text is locked', () => {
    const review = reviewImport(rows, { dateOrder: 'dmy', forceText: true });
    expect(review.ambiguous).toHaveLength(0);
    expect(review.protected).toHaveLength(0);
  });

  it('skips empty cells', () => {
    expect(reviewImport([['a', '', 'b']]).total).toBe(2);
  });
});

describe('ordinary text', () => {
  it.each(['hello', 'Widget A', 'N/A', 'a1', 'BRCA1', '2024 budget', 'x-ray'])(
    'leaves %s alone with no fuss',
    (input) => {
      const r = parse(input);
      expect(r.value).toBe(input);
      expect(r.kind).toBe('text');
      expect(r.confidence).toBe('certain');
    },
  );
});
