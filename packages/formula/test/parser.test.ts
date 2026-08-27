import { describe, expect, it } from 'vitest';
import { type Ast, Node } from '../src/ast.js';
import { Lexer, Tok, isCellShape, isColumnShape, tokenText } from '../src/lexer.js';
import { ParseError, parseContext, parseFormula } from '../src/parser.js';
import { serializeFormula } from '../src/serialize.js';

/**
 * Render an AST as a fully-parenthesised s-expression. Comparing these makes
 * precedence and associativity bugs obvious, which comparing formula text does
 * not - `2^3^2` serialises identically under either associativity.
 */
function sexp(n: Ast): string {
  switch (n.kind) {
    case Node.Number:
      return String(n.value);
    case Node.Text:
      return JSON.stringify(n.value);
    case Node.Bool:
      return n.value ? 'TRUE' : 'FALSE';
    case Node.ErrorLit:
      return n.code;
    case Node.Missing:
      return '_';
    case Node.Unary:
      return `(${n.op === ' ' ? 'isect' : n.op}u ${sexp(n.operand)})`;
    case Node.Postfix:
      return `(${n.op}p ${sexp(n.operand)})`;
    case Node.Binary:
      return `(${n.op === ' ' ? 'isect' : n.op} ${sexp(n.left)} ${sexp(n.right)})`;
    case Node.Paren:
      return sexp(n.inner);
    case Node.Call:
      return `(${n.name.toUpperCase()}${n.args.map((a) => ` ${sexp(a)}`).join('')})`;
    case Node.Array:
      return `{${n.rows.map((r) => r.map(sexp).join(',')).join(';')}}`;
    case Node.Ref:
      return `ref(${n.sheet ? `${n.sheet}!` : ''}${n.col}${n.colAbs ? 'a' : 'r'},${n.row}${n.rowAbs ? 'a' : 'r'})`;
    case Node.Range:
      return `range(${sexp(n.start)},${sexp(n.end)})`;
    case Node.Beam:
      return `beam(${n.axis},${n.from},${n.to})`;
    case Node.ThreeD:
      return `3d(${n.sheetStart}..${n.sheetEnd},${sexp(n.inner)})`;
    case Node.Name:
      return `name(${n.sheet ? `${n.sheet}!` : ''}${n.name})`;
    case Node.StructRef:
      return `struct(${n.table}${n.spec})`;
  }
}

const parse = (src: string) => parseFormula(src, { origin: { row: 0, col: 0 } });

describe('lexer', () => {
  const kinds = (src: string) =>
    Lexer.tokenize(src)
      .filter((t) => t.type !== Tok.Whitespace)
      .map((t) => `${Tok[t.type]}:${tokenText(src, t)}`);

  it('tokenizes a simple formula', () => {
    expect(kinds('=SUM(A1:B2)*2')).toEqual([
      'Equals:=',
      'Function:SUM',
      'OpenParen:(',
      'Ref:A1:B2',
      'CloseParen:)',
      'Operator:*',
      'Number:2',
    ]);
  });

  it('keeps whitespace as a token, because a space is the intersection operator', () => {
    const all = Lexer.tokenize('A1:C3 B1:B5').map((t) => t.type);
    expect(all).toContain(Tok.Whitespace);
  });

  it('distinguishes a function from a defined name by the following paren', () => {
    expect(kinds('SUM(1)')).toEqual(['Function:SUM', 'OpenParen:(', 'Number:1', 'CloseParen:)']);
    expect(kinds('SUM')).toEqual(['Name:SUM']);
    // A space before the paren means it is not a call.
    expect(kinds('SUM (1)')).toEqual(['Name:SUM', 'OpenParen:(', 'Number:1', 'CloseParen:)']);
  });

  it('lexes off-grid addresses as names, not references', () => {
    // These are legal defined names; treating them as references would turn a
    // workbook's own named range into a broken address.
    expect(kinds('A1048577')).toEqual(['Name:A1048577']);
    expect(kinds('XFE1')).toEqual(['Name:XFE1']);
    // TAX2024 is deliberately NOT in this list: column TAX is 13570, inside the
    // grid, so TAX2024 genuinely is a cell address and must lex as one.
    expect(kinds('TAX2024')).toEqual(['Ref:TAX2024']);
    expect(kinds('REVENUE2024')).toEqual(['Name:REVENUE2024']);
    expect(kinds('A1')).toEqual(['Ref:A1']);
    expect(kinds('XFD1048576')).toEqual(['Ref:XFD1048576']);
  });

  it('handles doubled quotes inside strings', () => {
    expect(kinds('"he said ""hi"""')).toEqual(['String:"he said ""hi"""']);
  });

  it('recognises every error literal', () => {
    for (const e of ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!']) {
      expect(kinds(e)).toEqual([`ErrorLit:${e}`]);
    }
  });

  it('reads a bare # as the spill-range operator', () => {
    expect(kinds('A1#')).toEqual(['Ref:A1', 'Operator:#']);
  });

  it('reads quoted and unquoted sheet contexts', () => {
    expect(kinds("'My Sheet'!A1")).toEqual(["Context:'My Sheet'!", 'Ref:A1']);
    expect(kinds('Sheet1!A1')).toEqual(['Context:Sheet1!', 'Ref:A1']);
    expect(kinds('Sheet1:Sheet3!A1')).toEqual(['Context:Sheet1:Sheet3!', 'Ref:A1']);
  });

  it('reads beams and structured references', () => {
    expect(kinds('A:A')).toEqual(['Beam:A:A']);
    expect(kinds('1:1')).toEqual(['Number:1', 'Operator::', 'Number:1']);
    expect(kinds('Table1[Column]')).toEqual(['StructRef:Table1[Column]']);
  });

  it.each([
    ['1', true],
    ['1.5', true],
    ['.5', true],
    ['1e3', true],
    ['1E-3', true],
    ['1.5e+10', true],
  ])('lexes %s as a number', (src, ok) => {
    expect(kinds(src)[0]?.startsWith('Number:')).toBe(ok);
  });

  it.each([
    ['A1', true],
    ['$A$1', true],
    ['XFD1048576', true],
    ['XFE1', false],
    ['A1048577', false],
    ['AAAA1', false],
    ['A', false],
    ['1', false],
  ])('isCellShape(%s) === %s', (word, want) => {
    expect(isCellShape(word)).toBe(want);
  });

  it.each([
    ['A', true],
    ['$XFD', true],
    ['XFE', false],
    ['A1', false],
  ])('isColumnShape(%s) === %s', (word, want) => {
    expect(isColumnShape(word)).toBe(want);
  });
});

describe('precedence - the cases Excel gets counter-intuitively right', () => {
  // Every expectation in this block was verified against the LibreOffice-
  // recalculated fixture fixtures/generated/precedence.calc.xlsx, not assumed.

  it('exponentiation is LEFT-associative: 2^3^2 is 64, not 512', () => {
    expect(sexp(parse('=2^3^2'))).toBe('(^ (^ 2 3) 2)');
  });

  it('unary minus binds tighter than ^: -2^2 is 4, not -4', () => {
    expect(sexp(parse('=-2^2'))).toBe('(^ (-u 2) 2)');
  });

  it('an explicit paren still gives -4', () => {
    expect(sexp(parse('=-(2^2)'))).toBe('(-u (^ 2 2))');
  });

  it('postfix % binds tighter than ^: 2^2% is 2^(2%)', () => {
    expect(sexp(parse('=2^2%'))).toBe('(^ 2 (%p 2))');
  });

  it('* and / bind tighter than + and -', () => {
    expect(sexp(parse('=1+2*3'))).toBe('(+ 1 (* 2 3))');
    expect(sexp(parse('=1-6/3'))).toBe('(- 1 (/ 6 3))');
  });

  it('^ binds tighter than *', () => {
    expect(sexp(parse('=2*3^2'))).toBe('(* 2 (^ 3 2))');
  });

  it('& binds looser than arithmetic but tighter than comparison', () => {
    expect(sexp(parse('="a"&"b"="ab"'))).toBe('(= (& "a" "b") "ab")');
    expect(sexp(parse('=1+1&"x"'))).toBe('(& (+ 1 1) "x")');
  });

  it('comparison binds loosest', () => {
    expect(sexp(parse('=1+1=2'))).toBe('(= (+ 1 1) 2)');
  });

  it('subtraction and division are left-associative', () => {
    expect(sexp(parse('=10-3-2'))).toBe('(- (- 10 3) 2)');
    expect(sexp(parse('=100/10/2'))).toBe('(/ (/ 100 10) 2)');
  });

  it('chains unary operators', () => {
    expect(sexp(parse('=--3'))).toBe('(-u (-u 3))');
    // Postfix % binds looser than unary minus (6 versus 7), so the negation
    // happens first and the percent applies to the result. Both readings give
    // -0.5, so the oracle cannot distinguish them; the binding powers can.
    expect(sexp(parse('=-50%'))).toBe('(%p (-u 50))');
  });

  it('the range operator binds tightest of all', () => {
    expect(sexp(parse('=SUM(A1:B2*2)'))).toBe('(SUM (* range(ref(0r,0r),ref(1r,1r)) 2))');
  });
});

describe('reference operators', () => {
  it('parses a space between references as intersection', () => {
    expect(sexp(parse('=SUM(A1:C3 B1:B5)'))).toBe(
      '(SUM (isect range(ref(0r,0r),ref(2r,2r)) range(ref(1r,0r),ref(1r,4r))))',
    );
  });

  it('parses a parenthesised comma list as union', () => {
    expect(sexp(parse('=SUM((A1:A2,B1:B2))'))).toBe(
      '(SUM (, range(ref(0r,0r),ref(0r,1r)) range(ref(1r,0r),ref(1r,1r))))',
    );
  });

  it('treats a comma inside a call as an argument separator, not a union', () => {
    expect(sexp(parse('=SUM(A1,B1)'))).toBe('(SUM ref(0r,0r) ref(1r,0r))');
  });

  it('keeps a dynamic range as a binary : node', () => {
    // The right operand is only known at evaluation time, so it cannot collapse.
    expect(sexp(parse('=A1:INDEX(B:B,3)'))).toBe(
      '(: ref(0r,0r) (INDEX beam(col,1,1) 3))',
    );
  });

  it('intersection does not fire across an operator', () => {
    expect(sexp(parse('=A1 + B1'))).toBe('(+ ref(0r,0r) ref(1r,0r))');
  });
});

describe('references', () => {
  it('stores relative references as offsets from the origin', () => {
    const ast = parseFormula('=B3', { origin: { row: 4, col: 2 } }) as Extract<Ast, { kind: Node.Ref }>;
    // B3 is col 1 row 2; the formula sits at col 2 row 4.
    expect(ast).toMatchObject({ col: -1, row: -2, colAbs: false, rowAbs: false });
  });

  it('stores absolute references as addresses', () => {
    const ast = parseFormula('=$B$3', { origin: { row: 4, col: 2 } }) as Extract<Ast, { kind: Node.Ref }>;
    expect(ast).toMatchObject({ col: 1, row: 2, colAbs: true, rowAbs: true });
  });

  it('handles mixed absoluteness per axis', () => {
    const ast = parseFormula('=$B3', { origin: { row: 4, col: 2 } }) as Extract<Ast, { kind: Node.Ref }>;
    expect(ast).toMatchObject({ col: 1, row: -2, colAbs: true, rowAbs: false });
  });

  it('parses sheet-qualified references', () => {
    expect(sexp(parse('=Sheet1!A1'))).toBe('ref(Sheet1!0r,0r)');
    expect(sexp(parse("='My Sheet'!A1"))).toBe('ref(My Sheet!0r,0r)');
  });

  it('parses 3-D references', () => {
    expect(sexp(parse('=SUM(Sheet1:Sheet3!A1)'))).toBe('(SUM 3d(Sheet1..Sheet3,ref(Sheet1!0r,0r)))');
  });

  it('parses whole-column and whole-row beams', () => {
    expect(sexp(parse('=SUM(A:A)'))).toBe('(SUM beam(col,0,0))');
    expect(sexp(parse('=SUM(B:D)'))).toBe('(SUM beam(col,1,3))');
  });

  it('parses external workbook references', () => {
    const ast = parseFormula("='[Book 1.xlsx]Sheet 1'!A1") as Extract<Ast, { kind: Node.Ref }>;
    expect(ast.book).toBe('Book 1.xlsx');
    expect(ast.sheet).toBe('Sheet 1');
  });

  it('parses structured table references', () => {
    expect(sexp(parse('=SUM(Table1[Amount])'))).toBe('(SUM struct(Table1[Amount]))');
    expect(sexp(parse('=Table1[@Amount]'))).toBe('struct(Table1[@Amount])');
  });
});

describe('literals and calls', () => {
  it('parses strings with escaped quotes', () => {
    expect(sexp(parse('="he said ""hi"""'))).toBe('"he said \\"hi\\""');
  });

  it('parses booleans and errors', () => {
    expect(sexp(parse('=TRUE'))).toBe('TRUE');
    expect(sexp(parse('=#DIV/0!'))).toBe('#DIV/0!');
  });

  it('parses nested calls', () => {
    expect(sexp(parse('=IF(SUM(A1:A3)>0,"y","n")'))).toBe(
      '(IF (> (SUM range(ref(0r,0r),ref(0r,2r))) 0) "y" "n")',
    );
  });

  it('parses a call with no arguments', () => {
    expect(sexp(parse('=PI()'))).toBe('(PI)');
    expect(sexp(parse('=TODAY()'))).toBe('(TODAY)');
  });

  it('parses omitted arguments', () => {
    expect(sexp(parse('=IF(A1,,0)'))).toBe('(IF ref(0r,0r) _ 0)');
  });

  it('parses array constants', () => {
    expect(sexp(parse('={1,2;3,4}'))).toBe('{1,2;3,4}');
    expect(sexp(parse('={1,2,3}'))).toBe('{1,2,3}');
    expect(sexp(parse('={"a";"b"}'))).toBe('{"a";"b"}');
  });

  it('accepts a semicolon separator, since files cross locales', () => {
    expect(sexp(parseFormula('=SUM(1;2)'))).toBe('(SUM 1 2)');
  });

  it('parses without a leading equals, as xlsx stores formulas', () => {
    expect(sexp(parse('SUM(A1:A2)'))).toBe('(SUM range(ref(0r,0r),ref(0r,1r)))');
  });
});

describe('errors', () => {
  it.each([
    ['=SUM(', 'unterminated'],
    ['=1+', 'unexpected end'],
    ['=(1', 'expected'],
    ['={1,2;3}', 'different lengths'],
  ])('rejects %s', (src, message) => {
    expect(() => parse(src)).toThrow(ParseError);
    expect(() => parse(src)).toThrow(new RegExp(message, 'i'));
  });

  it('reports the offset of the problem', () => {
    try {
      parse('=1+*2');
      expect.unreachable();
    } catch (e) {
      expect((e as ParseError).offset).toBeGreaterThan(0);
    }
  });
});

describe('context parsing', () => {
  it.each([
    ['Sheet1!', { sheet: 'Sheet1' }],
    ["'My Sheet'!", { sheet: 'My Sheet' }],
    ['Sheet1:Sheet3!', { sheet: 'Sheet1', sheetEnd: 'Sheet3' }],
    ["'[Book 1.xlsx]Sheet 1'!", { sheet: 'Sheet 1', book: 'Book 1.xlsx' }],
  ])('parses %s', (token, want) => {
    expect(parseContext(token)).toMatchObject(want);
  });
});

describe('round-tripping', () => {
  const cases = [
    'SUM(A1:B2)',
    'IF(A1>0,"yes","no")',
    '2^3^2',
    '-2^2',
    'A1+$B$2*C3',
    'Sheet1!A1',
    "'My Sheet'!A1:B2",
    'SUM(A:A)',
    '{1,2;3,4}',
    'CONCATENATE("a","b")',
    'Table1[Amount]',
    'SUM((A1:A2,B1:B2))',
    'IF(A1,,0)',
    'PI()',
    '"quote""inside"',
    '#DIV/0!',
    '1.5E-10',
  ];

  it.each(cases)('serialises %s back to itself', (src) => {
    expect(serializeFormula(parse(src))).toBe(src);
  });

  it('is idempotent under reparse', () => {
    for (const src of cases) {
      const once = serializeFormula(parse(src));
      expect(serializeFormula(parse(once))).toBe(once);
    }
  });

  it('shifts relative references when rendered at a new origin, which is fill-down', () => {
    // Parsed at A1, rendered at A2: the relative ref moves with the formula.
    const ast = parseFormula('=B1+$C$1', { origin: { row: 0, col: 0 } });
    expect(serializeFormula(ast, { origin: { row: 0, col: 0 } })).toBe('B1+$C$1');
    expect(serializeFormula(ast, { origin: { row: 1, col: 0 } })).toBe('B2+$C$1');
    expect(serializeFormula(ast, { origin: { row: 5, col: 3 } })).toBe('E6+$C$1');
  });

  it('renders a reference shifted off the grid as #REF!', () => {
    const ast = parseFormula('=A1', { origin: { row: 5, col: 5 } });
    expect(serializeFormula(ast, { origin: { row: 0, col: 0 } })).toBe('#REF!');
  });

  it('preserves the user whitespace inside a formula', () => {
    expect(serializeFormula(parse('=SUM( A1 , B1 )'))).toBe('SUM( A1 , B1 )');
  });

  it('can drop whitespace when asked', () => {
    expect(serializeFormula(parse('=SUM( A1 , B1 )'), { preserveWhitespace: false })).toBe(
      'SUM(A1,B1)',
    );
  });
});
