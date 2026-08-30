/**
 * Text function tests.
 *
 * The oracle block asserts against the values LibreOffice actually computed in
 * fixtures/generated/formulas.calc.xlsx, read back through our own xlsx reader,
 * so those expectations are not written from memory. Everything after it covers
 * the edges the fixture does not reach: blanks against zeros, numeric-looking
 * text, error propagation, boundary counts and surrogate pairs.
 */

import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { CellError, type Scalar, type Workbook, isError } from '@mirrorz/core';
import { readXlsx } from '../../formats/src/xlsx/read.js';
import { Evaluator, type SheetStore } from '../src/evaluator.js';
import { parseFormula } from '../src/parser.js';
import { FunctionRegistry } from '../src/registry.js';
import { TEXT_FUNCTIONS, setTextFormatter } from '../src/functions/text.js';
import { type ArrayValue, type Value, isArray } from '../src/value.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);

function storeFor(workbook: Workbook): SheetStore {
  const byName = new Map(workbook.sheets.map((s) => [s.name.toLowerCase(), s]));
  const sheet = (name: string) => byName.get(name.toLowerCase());
  return {
    getScalar: (name, row, col) => sheet(name)?.getValue(row, col) ?? null,
    *iterate(name, startRow, startCol, endRow, endCol) {
      const s = sheet(name);
      if (!s) return;
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const cell = s.getCell(r, c);
          if (cell !== undefined) yield { row: r, col: c, value: cell.value };
        }
      }
    },
    hasSheet: (name) => byName.has(name.toLowerCase()),
    sheetNames: () => workbook.sheets.map((s) => s.name),
    getDefinedName: () => undefined,
    usedBounds(name) {
      const b = sheet(name)?.bounds();
      return b ? { maxRow: b.maxRow, maxCol: b.maxCol } : null;
    },
  };
}

let workbook: Workbook;
let evaluator: Evaluator;

beforeAll(() => {
  workbook = readXlsx(
    new Uint8Array(readFileSync(new URL('formulas.calc.xlsx', FIXTURES))),
  ).workbook;
  // Only the text module is registered: the sibling categories are being built
  // in parallel, and a test that depends on them would fail for reasons that
  // have nothing to do with this file.
  const registry = new FunctionRegistry().registerAll(TEXT_FUNCTIONS);
  // Test-only scaffolding. The oracle stores `TRUE` as `TRUE()`, and several
  // error-propagation cases need an #N/A source; both live in categories other
  // agents are writing, so stubs stand in rather than a cross-module dependency.
  registry.registerAll([
    { name: 'TRUE', params: [], impl: () => true },
    { name: 'FALSE', params: [], impl: () => false },
    { name: 'NA', params: [], impl: () => CellError.NA },
  ]);
  evaluator = new Evaluator(storeFor(workbook), registry, { dateSystem: 1900 });
});

/** Evaluate a formula (written without the leading `=`) on the Formulas sheet. */
function calc(formula: string): Value {
  const ast = parseFormula(formula, { origin: { row: 0, col: 0 } });
  return evaluator.evaluate({ ast, sheet: 'Formulas', row: 0, col: 0 });
}

function code(v: Value): string {
  return isError(v) ? v.code : `not an error: ${JSON.stringify(v)}`;
}

function grid(v: Value): Scalar[][] {
  if (!isArray(v)) throw new Error(`expected an array, got ${JSON.stringify(v)}`);
  const a = v as ArrayValue;
  const rows: Scalar[][] = [];
  for (let r = 0; r < a.rows; r++) rows.push(a.data.slice(r * a.cols, (r + 1) * a.cols));
  return rows;
}

/* -------------------------------------------------------------------------- */

describe('the oracle fixture', () => {
  /** Case name -> { formula, value } from the Formulas sheet. */
  function oracleCases(): Map<string, { formula: string; value: Scalar }> {
    const sheet = workbook.getSheet('Formulas')!;
    const cases = new Map<string, { formula: string; value: Scalar }>();
    const bounds = sheet.bounds()!;
    for (let r = 1; r <= bounds.maxRow; r++) {
      const name = sheet.getValue(r, 0);
      const cell = sheet.getCell(r, 2);
      if (typeof name === 'string' && cell?.formula) {
        cases.set(name, { formula: cell.formula, value: cell.value });
      }
    }
    return cases;
  }

  const covered = [
    'CONCAT', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'UPPER', 'LOWER',
    'PROPER', 'TRIM', 'SUBSTITUTE', 'REPLACE', 'FIND', 'SEARCH', 'TEXT',
    'VALUE', 'REPT', 'EXACT', 'T',
  ];

  it('has a live value for every text case we claim to cover', () => {
    const cases = oracleCases();
    for (const name of covered) expect(cases.has(name), name).toBe(true);
    expect(cases.size).toBeGreaterThan(130);
  });

  it.each(
    // Vitest needs the table at collection time, so the fixture is read twice
    // rather than reaching into `beforeAll` state here.
    (() => {
      const wb = readXlsx(
        new Uint8Array(readFileSync(new URL('formulas.calc.xlsx', FIXTURES))),
      ).workbook;
      const sheet = wb.getSheet('Formulas')!;
      const rows: [string, string, Scalar][] = [];
      const bounds = sheet.bounds()!;
      for (let r = 1; r <= bounds.maxRow; r++) {
        const name = sheet.getValue(r, 0);
        const cell = sheet.getCell(r, 2);
        if (typeof name === 'string' && cell?.formula && covered.includes(name)) {
          rows.push([name, cell.formula, cell.value]);
        }
      }
      return rows;
    })(),
  )('%s reproduces the value the oracle computed', (_name, formula, expected) => {
    expect(calc(formula)).toEqual(expected);
  });
});

/* -------------------------------------------------------------------------- */

describe('CONCATENATE, CONCAT and TEXTJOIN', () => {
  it('joins scalars, coercing each the way & does', () => {
    expect(calc('CONCATENATE("a",1,TRUE)')).toBe('a1TRUE');
    expect(calc('CONCAT("a","b","c")')).toBe('abc');
  });

  it('treats a blank as empty text rather than a zero', () => {
    expect(calc('CONCATENATE("a",Data!Z1,"b")')).toBe('ab');
    expect(calc('CONCAT(Data!Z1:Z9)')).toBe('');
  });

  it('flattens a range, which CONCATENATE cannot do', () => {
    expect(calc('CONCAT(Data!B2:B4)')).toBe('AdaGraceKatherine');
  });

  it('propagates an error from any argument', () => {
    expect(code(calc('CONCAT("a",1/0)'))).toBe('#DIV/0!');
    expect(code(calc('CONCATENATE("a",NA())'))).toBe('#N/A');
  });

  it('joins with a delimiter and honours ignore_empty', () => {
    expect(calc('TEXTJOIN("-",TRUE,"x","","z")')).toBe('x-z');
    expect(calc('TEXTJOIN("-",FALSE,"x","","z")')).toBe('x--z');
    expect(calc('TEXTJOIN(", ",TRUE,Data!C2:C4)')).toBe('Eng, Eng, Science');
  });

  it('cycles a range of delimiters', () => {
    expect(calc('TEXTJOIN({"-","+"},TRUE,"a","b","c")')).toBe('a-b+c');
  });

  it('drops every value when the range is empty and blanks are ignored', () => {
    expect(calc('TEXTJOIN(",",TRUE,Data!Z1:Z9)')).toBe('');
    // Without ignore_empty the blanks still separate.
    expect(calc('TEXTJOIN(",",FALSE,Data!Z1:Z3)')).toBe(',,');
  });

  it('reports #VALUE! rather than truncating past the cell text limit', () => {
    expect(code(calc('CONCAT(REPT("a",30000),REPT("b",30000))'))).toBe('#VALUE!');
  });
});

describe('LEFT, RIGHT, MID and LEN', () => {
  it('counts from one', () => {
    expect(calc('LEFT("abcdef",2)')).toBe('ab');
    expect(calc('RIGHT("abcdef",2)')).toBe('ef');
    expect(calc('MID("abcdef",2,3)')).toBe('bcd');
  });

  it('defaults num_chars to one only when the argument is absent', () => {
    expect(calc('LEFT("abc")')).toBe('a');
    expect(calc('RIGHT("abc")')).toBe('c');
    // An argument written but left empty is blank, which coerces to zero.
    expect(calc('LEFT("abc",)')).toBe('');
    expect(calc('RIGHT("abc",)')).toBe('');
  });

  it('clamps a count past the end instead of failing', () => {
    expect(calc('LEFT("abc",99)')).toBe('abc');
    expect(calc('RIGHT("abc",99)')).toBe('abc');
    expect(calc('MID("abc",2,99)')).toBe('bc');
  });

  it('returns empty text when MID starts past the end', () => {
    expect(calc('MID("abc",4,2)')).toBe('');
    expect(calc('MID("abc",99,2)')).toBe('');
    expect(calc('MID("",1,2)')).toBe('');
  });

  it('rejects a negative count and a start below one', () => {
    expect(code(calc('LEFT("abc",-1)'))).toBe('#VALUE!');
    expect(code(calc('RIGHT("abc",-1)'))).toBe('#VALUE!');
    expect(code(calc('MID("abc",0,1)'))).toBe('#VALUE!');
    expect(code(calc('MID("abc",1,-1)'))).toBe('#VALUE!');
  });

  it('truncates a fractional count towards zero', () => {
    expect(calc('LEFT("abcdef",2.9)')).toBe('ab');
    expect(calc('MID("abcdef",2.9,2.9)')).toBe('bc');
  });

  it('accepts a numeric-looking count given as text', () => {
    expect(calc('LEFT("abcdef","3")')).toBe('abc');
    expect(code(calc('LEFT("abcdef","three")'))).toBe('#VALUE!');
  });

  it('coerces a non-text first argument the way & does', () => {
    expect(calc('LEFT(1234,2)')).toBe('12');
    expect(calc('LEN(TRUE)')).toBe(4);
    expect(calc('LEN(Data!Z1)')).toBe(0);
    expect(calc('LEN("")')).toBe(0);
  });

  it('counts UTF-16 code units, so an emoji is two characters', () => {
    // Excel agrees with UTF-16 here and not with Unicode; matching Excel is the
    // point, because a workbook's LEN column has to keep its values.
    expect(calc('LEN(UNICHAR(128169))')).toBe(2);
    expect(calc('LEN(UNICHAR(128169)&"x")')).toBe(3);
    expect((calc('LEFT(UNICHAR(128169)&"x",2)') as string).length).toBe(2);
  });

  it('propagates errors', () => {
    expect(code(calc('LEN(1/0)'))).toBe('#DIV/0!');
    expect(code(calc('MID(NA(),1,1)'))).toBe('#N/A');
  });
});

describe('UPPER, LOWER, PROPER, TRIM and CLEAN', () => {
  it('changes case', () => {
    expect(calc('UPPER("straße")')).toBe('STRAßE');
    expect(calc('LOWER("ÀÉÎ")')).toBe('àéî');
  });

  it('keeps a character whose case mapping is not one-to-one', () => {
    // JavaScript would give "SS" here; Excel keeps the sharp s, and a LEN
    // downstream of an UPPER must not change.
    expect(calc('LEN(UPPER("straße"))')).toBe(calc('LEN("straße")'));
  });

  it('starts a new word after any non-letter', () => {
    expect(calc('PROPER("o\'brien")')).toBe("O'Brien");
    expect(calc('PROPER("3rd place")')).toBe('3Rd Place');
    expect(calc('PROPER("HELLO world")')).toBe('Hello World');
    expect(calc('PROPER("a-b_c")')).toBe('A-B_C');
  });

  it('trims only the space character, and collapses internal runs', () => {
    expect(calc('TRIM("  a   b  ")')).toBe('a b');
    expect(calc('TRIM("   ")')).toBe('');
    // A non-breaking space is not a space as far as TRIM is concerned.
    expect(calc('TRIM(UNICHAR(160)&"a"&UNICHAR(160))')).toBe('\u00A0a\u00A0');
  });

  it('strips the first 32 ASCII codes and nothing else', () => {
    expect(calc('CLEAN(CHAR(9)&"a"&CHAR(10))')).toBe('a');
    expect(calc('LEN(CLEAN(CHAR(127)))')).toBe(1);
  });
});

describe('SUBSTITUTE and REPLACE', () => {
  it('replaces every occurrence by default and one when told to', () => {
    expect(calc('SUBSTITUTE("a-b-c","-","+")')).toBe('a+b+c');
    expect(calc('SUBSTITUTE("aaa","a","b",2)')).toBe('aba');
    expect(calc('SUBSTITUTE("aaa","a","b",4)')).toBe('aaa');
  });

  it('is case-sensitive and leaves an empty needle alone', () => {
    expect(calc('SUBSTITUTE("aA","a","x")')).toBe('xA');
    expect(calc('SUBSTITUTE("abc","","x")')).toBe('abc');
  });

  it('rejects an instance number below one', () => {
    expect(code(calc('SUBSTITUTE("aaa","a","b",0)'))).toBe('#VALUE!');
    expect(code(calc('SUBSTITUTE("aaa","a","b",-1)'))).toBe('#VALUE!');
  });

  it('replaces by position, appending when the start is past the end', () => {
    expect(calc('REPLACE("abcdef",2,3,"XY")')).toBe('aXYef');
    expect(calc('REPLACE("abc",10,2,"X")')).toBe('abcX');
    expect(calc('REPLACE("abc",1,0,"X")')).toBe('Xabc');
    expect(calc('REPLACE("abc",2,99,"X")')).toBe('aX');
  });

  it('rejects an out-of-range position', () => {
    expect(code(calc('REPLACE("abc",0,1,"X")'))).toBe('#VALUE!');
    expect(code(calc('REPLACE("abc",1,-1,"X")'))).toBe('#VALUE!');
  });
});

describe('FIND and SEARCH', () => {
  it('FIND is case-sensitive and literal', () => {
    expect(calc('FIND("b","abc")')).toBe(2);
    expect(code(calc('FIND("B","abc")'))).toBe('#VALUE!');
    // No wildcards: the star is just a star.
    expect(calc('FIND("*","a*b")')).toBe(2);
    expect(code(calc('FIND("*","ab")'))).toBe('#VALUE!');
  });

  it('SEARCH ignores case and honours wildcards', () => {
    expect(calc('SEARCH("B","abc")')).toBe(2);
    expect(calc('SEARCH("?c","abc")')).toBe(2);
    expect(calc('SEARCH("a*c","xxabc")')).toBe(3);
    // A tilde escapes the wildcard back into a literal.
    expect(calc('SEARCH("~*","a*b")')).toBe(2);
    expect(code(calc('SEARCH("~*","ab")'))).toBe('#VALUE!');
  });

  it('does not read the needle as a regular expression', () => {
    expect(calc('SEARCH("a.c","xa.c")')).toBe(2);
    expect(code(calc('SEARCH("a.c","xabc")'))).toBe('#VALUE!');
  });

  it('reports #VALUE! rather than #N/A when the needle is absent', () => {
    expect(code(calc('FIND("z","abc")'))).toBe('#VALUE!');
    expect(code(calc('SEARCH("z","abc")'))).toBe('#VALUE!');
  });

  it('honours start_num and validates it', () => {
    expect(calc('FIND("a","abca",2)')).toBe(4);
    expect(calc('SEARCH("A","abca",2)')).toBe(4);
    expect(code(calc('FIND("a","abc",0)'))).toBe('#VALUE!');
    expect(code(calc('SEARCH("a","abc",0)'))).toBe('#VALUE!');
    expect(code(calc('FIND("a","abc",99)'))).toBe('#VALUE!');
  });

  it('finds the empty needle at the start position', () => {
    expect(calc('FIND("","abc")')).toBe(1);
    expect(calc('FIND("","abc",3)')).toBe(3);
    expect(calc('SEARCH("","abc",2)')).toBe(2);
  });

  it('rejects a start_num past the last character, empty needle included', () => {
    // Microsoft documents one rule for both functions: "If start_num is not
    // greater than 0 (zero) or is greater than the length of within_text, the
    // #VALUE! error value is returned." Nothing exempts an empty find_text, so
    // the last legal start is LEN(within_text).
    expect(code(calc('FIND("","abc",4)'))).toBe('#VALUE!');
    expect(code(calc('SEARCH("","abc",4)'))).toBe('#VALUE!');
    expect(code(calc('FIND("","")'))).toBe('#VALUE!');
    expect(calc('FIND("c","abc",3)')).toBe(3);
  });
});

describe('REPT and EXACT', () => {
  it('repeats, and treats zero and empty text as empty', () => {
    expect(calc('REPT("ab",3)')).toBe('ababab');
    expect(calc('REPT("ab",0)')).toBe('');
    expect(calc('REPT("",5)')).toBe('');
    expect(calc('REPT("ab",2.9)')).toBe('abab');
  });

  it('rejects a negative count and an over-long result', () => {
    expect(code(calc('REPT("ab",-1)'))).toBe('#VALUE!');
    expect(code(calc('REPT("ab",20000)'))).toBe('#VALUE!');
  });

  it('compares case-sensitively after coercing both sides to text', () => {
    expect(calc('EXACT("a","A")')).toBe(false);
    expect(calc('EXACT("a","a")')).toBe(true);
    expect(calc('EXACT(1,"1")')).toBe(true);
    expect(calc('EXACT(Data!Z1,"")')).toBe(true);
  });
});

describe('VALUE and NUMBERVALUE', () => {
  it('parses the numeric text forms Excel accepts', () => {
    expect(calc('VALUE("123.45")')).toBe(123.45);
    expect(calc('VALUE("  -12  ")')).toBe(-12);
    expect(calc('VALUE("1,234")')).toBe(1234);
    expect(calc('VALUE("50%")')).toBe(0.5);
    expect(calc('VALUE("(100)")')).toBe(-100);
    expect(calc('VALUE("$1,234.50")')).toBe(1234.5);
    expect(calc('VALUE("1E3")')).toBe(1000);
  });

  it('separates a blank cell, which is zero, from empty text, which is not', () => {
    // =VALUE(A1) on an empty cell is 0 because the blank coerces to zero, but
    // empty text spells no number at all - which is why =VALUE(TRIM(A1)) is the
    // classic #VALUE! on an empty cell.
    expect(calc('VALUE(Data!Z1)')).toBe(0);
    expect(code(calc('VALUE("")'))).toBe('#VALUE!');
    expect(code(calc('VALUE("   ")'))).toBe('#VALUE!');
  });

  it('passes a number through and refuses a boolean', () => {
    expect(calc('VALUE(42)')).toBe(42);
    expect(code(calc('VALUE(TRUE)'))).toBe('#VALUE!');
    expect(code(calc('VALUE("abc")'))).toBe('#VALUE!');
    expect(code(calc('VALUE("12abc")'))).toBe('#VALUE!');
  });

  it('reads the unambiguous date and clock forms as serials', () => {
    expect(calc('VALUE("2024-02-29")')).toBe(45351);
    expect(calc('VALUE("06:00")')).toBe(0.25);
    expect(calc('VALUE("12:00 PM")')).toBe(0.5);
  });

  it('applies explicit separators, and rejects a conflicting pair', () => {
    expect(calc('NUMBERVALUE("2.500,50",",",".")')).toBe(2500.5);
    expect(calc('NUMBERVALUE("1,234.5")')).toBe(1234.5);
    expect(code(calc('NUMBERVALUE("1,0",",",",")'))).toBe('#VALUE!');
  });

  it('ignores spaces anywhere and compounds trailing percent signs', () => {
    expect(calc('NUMBERVALUE(" 3 000 ")')).toBe(3000);
    expect(calc('NUMBERVALUE("9%")')).toBe(0.09);
    expect(calc('NUMBERVALUE("9%%")')).toBe(0.0009);
    expect(calc('NUMBERVALUE("")')).toBe(0);
  });

  it('rejects a second decimal separator and a misplaced group separator', () => {
    expect(code(calc('NUMBERVALUE("3.5.5")'))).toBe('#VALUE!');
    expect(code(calc('NUMBERVALUE("3.5,5")'))).toBe('#VALUE!');
    expect(code(calc('NUMBERVALUE("abc")'))).toBe('#VALUE!');
  });
});

describe('T', () => {
  it('keeps text and blanks everything else', () => {
    expect(calc('T("txt")')).toBe('txt');
    expect(calc('T(1)')).toBe('');
    expect(calc('T(TRUE)')).toBe('');
    expect(calc('T(Data!Z1)')).toBe('');
    expect(calc('T(Data!B2)')).toBe('Ada');
  });

  it('lets an error through', () => {
    expect(code(calc('T(NA())'))).toBe('#N/A');
  });
});

describe('CHAR, CODE, UNICHAR and UNICODE', () => {
  it('round-trips the ASCII range', () => {
    expect(calc('CHAR(65)')).toBe('A');
    expect(calc('CODE("A")')).toBe(65);
    expect(calc('CODE("Abc")')).toBe(65);
  });

  it('uses Windows-1252 for the 128-159 block, not Latin-1', () => {
    expect(calc('CHAR(128)')).toBe('€');
    expect(calc('CHAR(149)')).toBe('•');
    expect(calc('CODE(CHAR(128))')).toBe(128);
    expect(calc('CODE(CHAR(149))')).toBe(149);
  });

  it('rejects codes outside 1-255', () => {
    expect(code(calc('CHAR(0)'))).toBe('#VALUE!');
    expect(code(calc('CHAR(256)'))).toBe('#VALUE!');
    expect(code(calc('CHAR(-1)'))).toBe('#VALUE!');
    expect(code(calc('CODE("")'))).toBe('#VALUE!');
  });

  it('substitutes a question mark for a character outside the code page', () => {
    expect(calc('CODE(UNICHAR(23383))')).toBe(63);
  });

  it('handles astral code points, which need a surrogate pair', () => {
    expect(calc('UNICODE(UNICHAR(128169))')).toBe(128169);
    expect(calc('UNICODE("A")')).toBe(65);
    expect(code(calc('UNICHAR(0)'))).toBe('#VALUE!');
    expect(code(calc('UNICHAR(1114112)'))).toBe('#VALUE!');
    // A lone surrogate is the one case the documentation gives #N/A rather than
    // #VALUE!: "If Unicode numbers are partial surrogates and data types that
    // are not valid, UNICHAR returns the #N/A error value."
    expect(code(calc('UNICHAR(55296)'))).toBe('#N/A');
    expect(code(calc('UNICHAR(57343)'))).toBe('#N/A');
    expect(code(calc('UNICODE("")'))).toBe('#VALUE!');
  });
});

describe('DOLLAR and FIXED', () => {
  it('formats currency with grouped thousands and parenthesised negatives', () => {
    expect(calc('DOLLAR(1234.567)')).toBe('$1,234.57');
    expect(calc('DOLLAR(-1234.567)')).toBe('($1,234.57)');
    expect(calc('DOLLAR(0)')).toBe('$0.00');
    expect(calc('DOLLAR(1234.567,0)')).toBe('$1,235');
  });

  it('rounds to the left of the point for a negative decimals argument', () => {
    expect(calc('DOLLAR(-1234.567,-2)')).toBe('($1,200)');
    expect(calc('FIXED(1234.567,-1)')).toBe('1,230');
  });

  it('uses a minus sign and can drop the separators', () => {
    expect(calc('FIXED(1234.567,1)')).toBe('1,234.6');
    expect(calc('FIXED(-1234.567,1)')).toBe('-1,234.6');
    expect(calc('FIXED(1234.567,1,TRUE)')).toBe('1234.6');
    expect(calc('FIXED(1234.567)')).toBe('1,234.57');
  });

  it('rounds half away from zero at display precision', () => {
    expect(calc('FIXED(1.005,2,TRUE)')).toBe('1.01');
    expect(calc('FIXED(2.5,0,TRUE)')).toBe('3');
    expect(calc('FIXED(-2.5,0,TRUE)')).toBe('-3');
  });

  it('propagates a non-numeric argument as #VALUE!', () => {
    expect(code(calc('DOLLAR("abc")'))).toBe('#VALUE!');
    expect(code(calc('FIXED("abc")'))).toBe('#VALUE!');
  });
});

describe('TEXT', () => {
  it('applies digit placeholders, decimals and thousands separators', () => {
    expect(calc('TEXT(1234.567,"#,##0.00")')).toBe('1,234.57');
    expect(calc('TEXT(0.5,"0.00")')).toBe('0.50');
    expect(calc('TEXT(0.5,"#.##")')).toBe('.5');
    expect(calc('TEXT(7,"000")')).toBe('007');
    expect(calc('TEXT(1234567,"#,##0")')).toBe('1,234,567');
  });

  it('drops trailing hash places but keeps zero places', () => {
    expect(calc('TEXT(1.5,"0.##")')).toBe('1.5');
    expect(calc('TEXT(1.5,"0.00")')).toBe('1.50');
    expect(calc('TEXT(1,"0.##")')).toBe('1');
  });

  it('rounds at fifteen significant digits, as Excel does', () => {
    expect(calc('TEXT(1.005,"0.00")')).toBe('1.01');
    expect(calc('TEXT(2.5,"0")')).toBe('3');
  });

  it('scales by a thousand for each trailing comma and by a hundred for percent', () => {
    expect(calc('TEXT(1234567,"#,##0,,")')).toBe('1');
    expect(calc('TEXT(0.285,"0.0%")')).toBe('28.5%');
  });

  it('scales by a thousand for a comma past the last placeholder, wherever it sits', () => {
    // The scaling comma does not have to be in the integer part: any comma after
    // the last digit placeholder divides by another thousand.
    expect(calc('TEXT(1234.567,"0.0,")')).toBe('1.2');
    expect(calc('TEXT(1234567,"#,##0.0,,")')).toBe('1.2');
    expect(calc('TEXT(1000,"0,")')).toBe('1');
  });

  it('gives a ? placeholder the width of the digit it is standing in for', () => {
    expect(calc('TEXT(0.5,"?.?")')).toBe(' .5');
    expect(calc('TEXT(5,"??0.0")')).toBe('  5.0');
    expect(calc('TEXT(1.5,"0.??")')).toBe('1.5 ');
    expect(calc('TEXT(1.25,"0.??")')).toBe('1.25');
  });

  it('formats numeric text as the number it spells', () => {
    expect(calc('TEXT("5","0.00")')).toBe('5.00');
    expect(calc('TEXT("1,234.5","0.0")')).toBe('1234.5');
    expect(calc('TEXT("abc","0.00")')).toBe('abc');
    // Blank text is not a zero: it stays text and passes through.
    expect(calc('TEXT("","0.00")')).toBe('');
    expect(calc('TEXT(" ","0.00")')).toBe(' ');
  });

  it('renders the @ placeholder, in a numeric section as well as a text one', () => {
    expect(calc('TEXT(1234.567,"@")')).toBe('1234.567');
    expect(calc('TEXT("abc","(@)")')).toBe('(abc)');
    // A trailing section carrying @ is the text section, not the negative one:
    // "0.00;@" means numbers here, text there.
    expect(calc('TEXT(1.5,"0.00;@")')).toBe('1.50');
    expect(calc('TEXT(-1.5,"0.00;@")')).toBe('-1.50');
    expect(calc('TEXT("abc","0.00;@")')).toBe('abc');
  });

  it('writes scientific and engineering notation', () => {
    expect(calc('TEXT(123456789,"0.00E+00")')).toBe('1.23E+08');
    expect(calc('TEXT(0.000123,"0.00E+00")')).toBe('1.23E-04');
    expect(calc('TEXT(-12345,"0.00E+00")')).toBe('-1.23E+04');
    expect(calc('TEXT(0,"0.00E+00")')).toBe('0.00E+00');
    // A '-' exponent sign shows nothing for a positive exponent.
    expect(calc('TEXT(1234.5,"0.0E-0")')).toBe('1.2E3');
    // A '#' or '?' among the mantissa's integer places steps the exponent in
    // multiples of their width, which is engineering notation.
    expect(calc('TEXT(123456789,"##0.0E+0")')).toBe('123.5E+6');
    // Rounding the mantissa up a decade moves the exponent with it.
    expect(calc('TEXT(9.99,"0.0E+0")')).toBe('1.0E+1');
    expect(calc('TEXT(1,"0.0e+00")')).toBe('1.0e+00');
  });

  it('writes fraction formats', () => {
    // The examples from Microsoft's own TEXT article.
    expect(calc('TEXT(4.34,"# ?/?")')).toBe('4 1/3');
    expect(calc('TEXT(0.34,"# ?/?")')).toBe(' 1/3');
    expect(calc('TEXT(12200000,"#,###.0,")')).toBe('12,200.0');
    expect(calc('TEXT(12200000,"0.00E+00")')).toBe('1.22E+07');
    // The denominator is a best rational approximation, not a rounded one: 0.7
    // is 5/7 with one digit to spend, never 6/9.
    expect(calc('TEXT(0.7,"?/?")')).toBe('5/7');
    expect(calc('TEXT(4.34,"?/?")')).toBe('13/3');
    // A fixed denominator, and the padding that keeps a column aligned.
    expect(calc('TEXT(4.34,"# ?/16")')).toBe('4 5/16');
    expect(calc('TEXT(5.25,"# ???/???")')).toBe('5   1/4  ');
    expect(calc('TEXT(5.3,"# ???/???")')).toBe('5   3/10 ');
    // A whole number blanks the fraction rather than printing 0/1, and a
    // remainder that rounds up carries into the integer.
    expect(calc('TEXT(2,"# ?/?")')).toBe('2    ');
    expect(calc('TEXT(1.9999,"# ?/?")')).toBe('2    ');
    expect(calc('TEXT(-4.34,"# ?/?")')).toBe('-4 1/3');
    // A date format's slashes are not a fraction.
    expect(calc('TEXT(45351,"m/d/yyyy")')).toBe('2/29/2024');
  });

  it('takes the case of AM/PM from the format code', () => {
    expect(calc('TEXT(0.75,"h AM/PM")')).toBe('6 PM');
    expect(calc('TEXT(0.75,"h am/pm")')).toBe('6 pm');
    expect(calc('TEXT(0.75,"h A/P")')).toBe('6 P');
    expect(calc('TEXT(0.25,"h a/p")')).toBe('6 a');
  });

  it('spreads digits across placeholders separated by a literal', () => {
    expect(calc('TEXT(5551234,"000-0000")')).toBe('555-1234');
  });

  it('picks the section by sign, letting the negative section supply its sign', () => {
    expect(calc('TEXT(-5,"0.00;(0.00)")')).toBe('(5.00)');
    expect(calc('TEXT(-5,"0.00")')).toBe('-5.00');
    expect(calc(`TEXT(0,"0.00;(0.00);""nil""")`)).toBe('nil');
  });

  it('leaves text alone unless the format has a text section', () => {
    expect(calc('TEXT("abc","0.00")')).toBe('abc');
    expect(calc('TEXT("abc","0.00;;;(@)")')).toBe('(abc)');
  });

  it('renders the common date and time tokens', () => {
    expect(calc('TEXT(45351,"yyyy-mm-dd")')).toBe('2024-02-29');
    expect(calc('TEXT(45351,"yy/m/d")')).toBe('24/2/29');
    expect(calc('TEXT(45351,"mmm d, yyyy")')).toBe('Feb 29, 2024');
    expect(calc('TEXT(45351,"mmmm")')).toBe('February');
    expect(calc('TEXT(45351,"dddd")')).toBe('Thursday');
    expect(calc('TEXT(45351,"ddd")')).toBe('Thu');
  });

  it('reads m as minutes next to an hour or a seconds token', () => {
    expect(calc('TEXT(0.573263888888889,"h:mm:ss")')).toBe('13:45:30');
    expect(calc('TEXT(0.5,"h:mm AM/PM")')).toBe('12:00 PM');
    expect(calc('TEXT(0.25,"h:mm AM/PM")')).toBe('6:00 AM');
    expect(calc('TEXT(0.25,"hh:mm")')).toBe('06:00');
  });

  it('reproduces the phantom 29 February 1900', () => {
    expect(calc('TEXT(60,"yyyy-mm-dd")')).toBe('1900-02-29');
    expect(calc('TEXT(61,"yyyy-mm-dd")')).toBe('1900-03-01');
  });

  it('rejects a negative serial in a date format', () => {
    // There is no date before the epoch to render, so the date branch reports
    // the same error Excel does for an out-of-range serial.
    expect(code(calc('TEXT(-1,"yyyy")'))).toBe('#VALUE!');
    expect(code(calc('TEXT(-1,"mmm d")'))).toBe('#VALUE!');
  });

  it('keeps very large and very small magnitudes intact', () => {
    expect(calc('TEXT(1E20,"0")')).toBe('100000000000000000000');
    expect(calc('TEXT(1E-20,"0.00")')).toBe('0.00');
    expect(calc('TEXT(0.000001,"0.000000")')).toBe('0.000001');
  });

  it('falls back to the General rendering', () => {
    expect(calc('TEXT(1/3,"General")')).toBe('0.333333333333333');
    expect(calc('TEXT(12,"General")')).toBe('12');
  });

  it('treats a blank value as zero and propagates errors', () => {
    expect(calc('TEXT(Data!Z1,"0.00")')).toBe('0.00');
    expect(code(calc('TEXT(1/0,"0.00")'))).toBe('#DIV/0!');
    // An empty text section hides the value, which is what Excel shows too.
    expect(calc('TEXT("abc","0.00;;;")')).toBe('');
  });
});

describe('TEXTBEFORE and TEXTAFTER', () => {
  it('splits around the nth delimiter, counting from either end', () => {
    expect(calc('TEXTBEFORE("a-b-c","-")')).toBe('a');
    expect(calc('TEXTAFTER("a-b-c","-")')).toBe('b-c');
    expect(calc('TEXTBEFORE("a-b-c","-",2)')).toBe('a-b');
    expect(calc('TEXTAFTER("a-b-c","-",2)')).toBe('c');
    expect(calc('TEXTBEFORE("a-b-c","-",-1)')).toBe('a-b');
    expect(calc('TEXTAFTER("a-b-c","-",-2)')).toBe('b-c');
  });

  it('returns #N/A, or the caller-supplied value, when the delimiter is absent', () => {
    expect(code(calc('TEXTBEFORE("abc","x")'))).toBe('#N/A');
    expect(calc('TEXTBEFORE("abc","x",1,0,0,"none")')).toBe('none');
    expect(code(calc('TEXTAFTER("a-b","-",5)'))).toBe('#N/A');
  });

  it('rejects instance zero', () => {
    expect(code(calc('TEXTBEFORE("a-b","-",0)'))).toBe('#VALUE!');
  });

  it('can ignore case and can treat the end of the text as a delimiter', () => {
    expect(calc('TEXTBEFORE("aXbXc","x",1,1)')).toBe('a');
    expect(calc('TEXTBEFORE("a-b","-",2,0,1)')).toBe('a-b');
    expect(calc('TEXTAFTER("a-b","-",2,0,1)')).toBe('');
  });

  it('accepts several delimiters and returns empty text for empty input', () => {
    expect(calc('TEXTBEFORE("a;b,c",{",",";"},2)')).toBe('a;b');
    expect(calc('TEXTBEFORE("","-")')).toBe('');
  });
});

describe('TEXTSPLIT', () => {
  it('splits into a row, and into a grid when a row delimiter is given', () => {
    expect(grid(calc('TEXTSPLIT("a,b,c",",")'))).toEqual([['a', 'b', 'c']]);
    expect(grid(calc('TEXTSPLIT("a,b;c,d",",",";")'))).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('pads a short row with #N/A by default', () => {
    const rows = grid(calc('TEXTSPLIT("a,b;c",",",";")'));
    expect(rows[0]).toEqual(['a', 'b']);
    expect(rows[1]![0]).toBe('c');
    expect(code(rows[1]![1] as Value)).toBe('#N/A');
  });

  it('uses an explicit pad value', () => {
    expect(grid(calc('TEXTSPLIT("a,b;c",",",";",FALSE,0,"-")'))).toEqual([
      ['a', 'b'],
      ['c', '-'],
    ]);
  });

  it('keeps empty pieces unless told to ignore them', () => {
    expect(grid(calc('TEXTSPLIT("a,,b",",")'))).toEqual([['a', '', 'b']]);
    expect(grid(calc('TEXTSPLIT("a,,b",",",,TRUE)'))).toEqual([['a', 'b']]);
  });

  it('needs at least one delimiter', () => {
    expect(code(calc('TEXTSPLIT("abc","")'))).toBe('#VALUE!');
  });
});

describe('ARRAYTOTEXT and VALUETOTEXT', () => {
  it('renders concisely by default', () => {
    expect(calc('ARRAYTOTEXT({1,2;3,4})')).toBe('1, 2, 3, 4');
    expect(calc('ARRAYTOTEXT(Data!B2:B4)')).toBe('Ada, Grace, Katherine');
    expect(calc('VALUETOTEXT("a")')).toBe('a');
    expect(calc('VALUETOTEXT(TRUE)')).toBe('TRUE');
    expect(calc('VALUETOTEXT(1/3)')).toBe('0.333333333333333');
  });

  it('renders strictly as an array literal with quoted text', () => {
    expect(calc('ARRAYTOTEXT({1,"a"},1)')).toBe('{1,"a"}');
    expect(calc('ARRAYTOTEXT({1,2;3,4},1)')).toBe('{1,2;3,4}');
    expect(calc('VALUETOTEXT("a",1)')).toBe('"a"');
    expect(calc('VALUETOTEXT("say ""hi""",1)')).toBe('"say ""hi"""');
  });

  it('renders an error as its code rather than propagating it', () => {
    expect(calc('VALUETOTEXT(NA())')).toBe('#N/A');
    expect(calc('ARRAYTOTEXT(NA())')).toBe('#N/A');
    expect(calc('ARRAYTOTEXT({1,2}/0)')).toBe('#DIV/0!, #DIV/0!');
  });

  it('rejects a format other than 0 or 1', () => {
    expect(code(calc('VALUETOTEXT("a",2)'))).toBe('#VALUE!');
    expect(code(calc('ARRAYTOTEXT({1},2)'))).toBe('#VALUE!');
  });
});

describe('the REGEX family', () => {
  it('tests, honouring the case-sensitivity flag', () => {
    expect(calc('REGEXTEST("abc","b")')).toBe(true);
    expect(calc('REGEXTEST("ABC","b")')).toBe(false);
    expect(calc('REGEXTEST("ABC","b",1)')).toBe(true);
    expect(calc('REGEXTEST("a1","^[a-z]\\d$")')).toBe(true);
  });

  it('extracts the first match, all matches, or the capture groups', () => {
    expect(calc('REGEXEXTRACT("a1b2","\\d")')).toBe('1');
    expect(grid(calc('REGEXEXTRACT("a1b2","\\d",1)'))).toEqual([['1'], ['2']]);
    expect(grid(calc('REGEXEXTRACT("2024-05","(\\d+)-(\\d+)",2)'))).toEqual([['2024', '05']]);
  });

  it('reports #N/A when nothing matches', () => {
    expect(code(calc('REGEXEXTRACT("abc","z")'))).toBe('#N/A');
    expect(code(calc('REGEXEXTRACT("abc","z",1)'))).toBe('#N/A');
    expect(code(calc('REGEXEXTRACT("abc","b",2)'))).toBe('#N/A');
  });

  it('replaces all matches, or one chosen from either end', () => {
    expect(calc('REGEXREPLACE("a1b2","\\d","#")')).toBe('a#b#');
    expect(calc('REGEXREPLACE("a1b2","\\d","#",1)')).toBe('a#b2');
    expect(calc('REGEXREPLACE("a1b2","\\d","#",-1)')).toBe('a1b#');
    expect(calc('REGEXREPLACE("a1b2","\\d","#",5)')).toBe('a1b2');
    expect(calc('REGEXREPLACE("abc","z","#")')).toBe('abc');
  });

  it('expands group references in the replacement', () => {
    expect(calc('REGEXREPLACE("john smith","(\\w+) (\\w+)","$2 $1")')).toBe('smith john');
    expect(calc('REGEXREPLACE("a","a","$$")')).toBe('$');
  });

  it('refuses the constructs RE2 does not have, rather than quietly differing', () => {
    // Backreference.
    expect(code(calc('REGEXTEST("aa","(a)\\1")'))).toBe('#VALUE!');
    expect(code(calc('REGEXREPLACE("aa","(a)\\1","x")'))).toBe('#VALUE!');
    // Lookbehind.
    expect(code(calc('REGEXTEST("ab","(?<=a)b")'))).toBe('#VALUE!');
    expect(code(calc('REGEXEXTRACT("ab","(?<!a)b")'))).toBe('#VALUE!');
    // Lookahead is in RE2's syntax set for our purposes and stays allowed.
    expect(calc('REGEXTEST("ab","a(?=b)")')).toBe(true);
  });

  it('reports a malformed pattern as #VALUE!', () => {
    expect(code(calc('REGEXTEST("a","(")'))).toBe('#VALUE!');
  });
});

describe('the injected-formatter seam', () => {
  it('hands TEXT over to the full engine once one is supplied', () => {
    setTextFormatter((value, formatCode, dateSystem) => `${String(value)}|${formatCode}|${dateSystem}`);
    try {
      expect(calc('TEXT(5,"0.00")')).toBe('5|0.00|1900');
    } finally {
      setTextFormatter(undefined);
    }
    // Removing it restores the local subset.
    expect(calc('TEXT(5,"0.00")')).toBe('5.00');
  });
});

describe('boundary magnitudes', () => {
  it('holds text right up to the cell limit and refuses one character more', () => {
    expect(calc('LEN(REPT("a",32767))')).toBe(32767);
    expect(code(calc('REPT("a",32768)'))).toBe('#VALUE!');
  });

  it('accepts an enormous count without allocating past the source', () => {
    expect(calc('MID("abc",1,1000000000)')).toBe('abc');
    expect(calc('LEFT("abc",1000000000)')).toBe('abc');
  });

  it('coerces numeric arguments to text before searching', () => {
    expect(calc('FIND(2,123)')).toBe(2);
    expect(calc('SEARCH(2,123)')).toBe(2);
  });

  it('folds case beyond ASCII in SEARCH', () => {
    expect(calc('SEARCH("É","xéy")')).toBe(2);
  });

  it('rejects more than 127 decimal places', () => {
    expect(code(calc('FIXED(1,128)'))).toBe('#VALUE!');
    expect(code(calc('DOLLAR(1,128)'))).toBe('#VALUE!');
  });
});

describe('registration metadata', () => {
  it('registers every declared function exactly once', () => {
    const registry = new FunctionRegistry().registerAll(TEXT_FUNCTIONS);
    expect(registry.size).toBe(TEXT_FUNCTIONS.length);
    expect(registry.size).toBe(36);
  });

  it('marks the post-2007 additions as future functions, and nothing as volatile', () => {
    const registry = new FunctionRegistry().registerAll(TEXT_FUNCTIONS);
    expect(registry.get('TEXTJOIN')?.futureFunction).toBe(true);
    expect(registry.get('TEXTSPLIT')?.futureFunction).toBe(true);
    expect(registry.get('LEFT')?.futureFunction).toBe(false);
    for (const spec of TEXT_FUNCTIONS) {
      expect(spec.volatile, spec.name).toBeUndefined();
      expect(spec.structural, spec.name).toBeUndefined();
    }
  });
});

describe('errors as values', () => {
  it('never throws, whatever the argument', () => {
    const formulas = [
      'LEFT(1/0,1)', 'MID(NA(),0,0)', 'TEXT(NA(),"0")', 'FIND(1/0,"a")',
      'REPT(NA(),1)', 'SUBSTITUTE(1/0,"a","b")', 'TEXTSPLIT(1/0,",")',
      'CONCAT(NA())', 'UNICHAR(NA())', 'VALUE(1/0)',
    ];
    for (const f of formulas) {
      expect(() => calc(f), f).not.toThrow();
      expect(isError(calc(f)), f).toBe(true);
    }
  });

  it('reports #VALUE! and not #NUM! for a non-numeric count', () => {
    expect(code(calc('LEFT("abc","x")'))).toBe('#VALUE!');
    expect(code(calc('REPT("a","x")'))).toBe('#VALUE!');
    expect(code(calc('CHAR("x")'))).toBe('#VALUE!');
  });
});

describe('exports', () => {
  it('names every function it claims to implement', () => {
    expect(TEXT_FUNCTIONS.map((f) => f.name).sort()).toEqual(
      [
        'ARRAYTOTEXT', 'CHAR', 'CLEAN', 'CODE', 'CONCAT', 'CONCATENATE',
        'DOLLAR', 'EXACT', 'FIND', 'FIXED', 'LEFT', 'LEN', 'LOWER', 'MID',
        'NUMBERVALUE', 'PROPER', 'REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST',
        'REPLACE', 'REPT', 'RIGHT', 'SEARCH', 'SUBSTITUTE', 'T', 'TEXT',
        'TEXTAFTER', 'TEXTBEFORE', 'TEXTJOIN', 'TEXTSPLIT', 'TRIM', 'UNICHAR',
        'UNICODE', 'UPPER', 'VALUE', 'VALUETOTEXT',
      ].sort(),
    );
  });
});
