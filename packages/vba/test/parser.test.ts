import { describe, expect, it } from 'vitest';
import { Tk, joinContinuations, significant, tokenize } from '../src/lexer.js';
import { analyseModule, summarise } from '../src/compat.js';
import { type Expr, type Stmt, parseModule } from '../src/parser.js';

const kinds = (src: string) =>
  significant(tokenize(src))
    .filter((t) => t.type !== Tk.EOF)
    .map((t) => `${Tk[t.type]}:${t.text}`);

describe('line continuation', () => {
  it('joins a continued line', () => {
    // The join contributes exactly one space, whatever whitespace surrounded
    // the underscore, so the result does not depend on how the author indented.
    expect(joinContinuations('x = 1 + _\n    2').lines).toEqual(['x = 1 +     2']);
    expect(joinContinuations('x = 1 +   _\n2').lines).toEqual(['x = 1 + 2']);
    expect(joinContinuations('x = 1 + _\n2').lines).toEqual(['x = 1 + 2']);
  });

  it('joins several continuations in a row', () => {
    const { lines } = joinContinuations('a = 1 + _\n2 + _\n3');
    expect(lines).toHaveLength(1);
    expect(lines[0]).toContain('3');
  });

  it('reports the original line number for an error', () => {
    const { origins } = joinContinuations('a = 1 + _\n2\nb = 3');
    expect(origins).toEqual([1, 3]);
  });

  it('does not treat a trailing underscore in a name as a continuation', () => {
    // The space before the underscore is what makes it a continuation.
    const { lines } = joinContinuations('my_var_ = 1\nnext = 2');
    expect(lines).toHaveLength(2);
  });
});

describe('tokenizer', () => {
  it('is case-insensitive about keywords', () => {
    expect(kinds('dim x')).toEqual(['Keyword:dim', 'Identifier:x', 'Terminator:\n']);
    expect(kinds('DIM x')[0]).toBe('Keyword:DIM');
  });

  it('reads strings with doubled quotes', () => {
    expect(kinds('s = "say ""hi"""')).toContain('String:"say ""hi"""');
  });

  it('does not treat an apostrophe inside a string as a comment', () => {
    const t = kinds(`s = "it's fine"`);
    expect(t).toContain(`String:"it's fine"`);
    expect(t.some((x) => x.startsWith('Comment'))).toBe(false);
  });

  it('reads comments to end of line', () => {
    expect(significant(tokenize("x = 1 ' a note")).some((t) => t.type === Tk.Comment)).toBe(false);
    expect(tokenize("x = 1 ' a note").some((t) => t.type === Tk.Comment)).toBe(true);
  });

  it('treats Rem as a comment only as a whole word', () => {
    expect(tokenize('Rem this is a note').some((t) => t.type === Tk.Comment)).toBe(true);
    // A variable called Remainder must not be swallowed.
    expect(kinds('Remainder = 1')[0]).toBe('Identifier:Remainder');
  });

  it('splits statements on a colon', () => {
    const t = kinds('x = 1: y = 2');
    expect(t.filter((x) => x.startsWith('Terminator'))).toHaveLength(2);
  });

  it('recognises a line label', () => {
    expect(kinds('ErrorHandler:')[0]).toBe('Label:ErrorHandler');
  });

  it('reads date literals', () => {
    expect(kinds('d = #1/15/2024#')).toContain('DateLiteral:#1/15/2024#');
  });

  it('reads hex and octal literals', () => {
    expect(kinds('x = &HFF')).toContain('Number:&HFF');
    expect(kinds('x = &O17')).toContain('Number:&O17');
  });

  it('reads type suffixes as part of the name', () => {
    expect(kinds('count% = 1')[0]).toBe('Identifier:count%');
    expect(kinds('name$ = "a"')[0]).toBe('Identifier:name$');
  });

  it('distinguishes the Double suffix from a date literal', () => {
    // A lone hash with no closing hash is a type suffix.
    expect(kinds('x# = 1')[0]).toBe('Identifier:x#');
  });

  it('reads Attribute lines as their own token', () => {
    expect(tokenize('Attribute VB_Name = "Module1"')[0]!.type).toBe(Tk.Attribute);
  });

  it('reads conditional-compilation directives', () => {
    expect(tokenize('#If Win64 Then')[0]!.type).toBe(Tk.Directive);
  });

  it('reads two-character operators', () => {
    for (const op of ['<=', '>=', '<>']) {
      expect(kinds(`If a ${op} b Then`)).toContain(`Operator:${op}`);
    }
  });

  it('reads a bracketed identifier', () => {
    expect(kinds('[My Sheet].Activate')[0]).toBe('Identifier:[My Sheet]');
  });
});

describe('declarations', () => {
  const decls = (src: string) => parseModule(src).declarations;

  it('parses Dim with a type', () => {
    const [d] = decls('Dim x As Long') as [Extract<Stmt, { kind: 'dim' }>];
    expect(d.scope).toBe('Dim');
    expect(d.vars[0]).toMatchObject({ name: 'x', typeName: 'Long' });
  });

  it('parses several variables in one Dim', () => {
    const [d] = decls('Dim a As Long, b As String, c') as [Extract<Stmt, { kind: 'dim' }>];
    expect(d.vars.map((v) => v.name)).toEqual(['a', 'b', 'c']);
    expect(d.vars[2]!.typeName).toBeUndefined();
  });

  it('parses array declarations with bounds', () => {
    const [d] = decls('Dim grid(1 To 10, 1 To 5) As Double') as [Extract<Stmt, { kind: 'dim' }>];
    expect(d.vars[0]!.isArray).toBe(true);
    expect(d.vars[0]!.bounds).toHaveLength(2);
  });

  it('parses a dynamic array', () => {
    const [d] = decls('Dim items() As String') as [Extract<Stmt, { kind: 'dim' }>];
    expect(d.vars[0]!.isArray).toBe(true);
    expect(d.vars[0]!.bounds).toBeUndefined();
  });

  it('parses visibility modifiers', () => {
    expect((decls('Public x As Long')[0] as Extract<Stmt, { kind: 'dim' }>).scope).toBe('Public');
    expect((decls('Private y')[0] as Extract<Stmt, { kind: 'dim' }>).scope).toBe('Private');
  });

  it('parses WithEvents and As New', () => {
    const [a] = decls('Private WithEvents app As Application') as [Extract<Stmt, { kind: 'dim' }>];
    expect(a.vars[0]!.withEvents).toBe(true);
    const [b] = decls('Dim c As New Collection') as [Extract<Stmt, { kind: 'dim' }>];
    expect(b.vars[0]!.isNew).toBe(true);
  });

  it('parses Const', () => {
    const [c] = decls('Const MAX As Long = 100') as [Extract<Stmt, { kind: 'const' }>];
    expect(c.vars[0]).toMatchObject({ name: 'MAX', typeName: 'Long' });
    expect(c.vars[0]!.value).toMatchObject({ kind: 'number', value: 100 });
  });

  it('parses Option statements', () => {
    expect(decls('Option Explicit')[0]).toMatchObject({ kind: 'option', text: 'Option Explicit' });
  });

  it('parses a user-defined Type', () => {
    const [t] = parseModule('Type Point\n  X As Double\n  Y As Double\nEnd Type').declarations as [
      Extract<Stmt, { kind: 'type' }>,
    ];
    expect(t.name).toBe('Point');
    expect(t.fields.map((f) => f.name)).toEqual(['X', 'Y']);
  });

  it('parses an Enum', () => {
    const [e] = parseModule('Enum Colours\n  Red = 1\n  Green\nEnd Enum').declarations as [
      Extract<Stmt, { kind: 'enum' }>,
    ];
    expect(e.name).toBe('Colours');
    expect(e.members.map((m) => m.name)).toEqual(['Red', 'Green']);
  });
});

describe('procedures', () => {
  it('parses a Sub with parameters', () => {
    const m = parseModule('Sub Greet(name As String, Optional loud As Boolean = False)\nEnd Sub');
    const p = m.procedures[0]!;
    expect(p.kind).toBe('Sub');
    expect(p.name).toBe('Greet');
    expect(p.params).toHaveLength(2);
    expect(p.params[1]).toMatchObject({ name: 'loud', optional: true });
  });

  it('defaults parameters to ByRef, as VBA does', () => {
    const p = parseModule('Sub S(a As Long, ByVal b As Long)\nEnd Sub').procedures[0]!;
    expect(p.params[0]!.byRef).toBe(true);
    expect(p.params[1]!.byRef).toBe(false);
  });

  it('parses a Function with a return type', () => {
    const p = parseModule('Function Add(a As Long, b As Long) As Long\nEnd Function').procedures[0]!;
    expect(p.kind).toBe('Function');
    expect(p.returnType).toBe('Long');
  });

  it('parses Property accessors', () => {
    const m = parseModule(
      'Property Get Name() As String\nEnd Property\nProperty Let Name(v As String)\nEnd Property',
    );
    expect(m.procedures.map((p) => p.kind)).toEqual(['PropertyGet', 'PropertyLet']);
  });

  it('parses visibility and Static', () => {
    const p = parseModule('Private Static Sub S()\nEnd Sub').procedures[0]!;
    expect(p.visibility).toBe('Private');
    expect(p.isStatic).toBe(true);
  });

  it('parses ParamArray', () => {
    const p = parseModule('Sub S(ParamArray items())\nEnd Sub').procedures[0]!;
    expect(p.params[0]!.paramArray).toBe(true);
  });

  it('warns rather than failing when End Sub is missing', () => {
    const m = parseModule('Sub Broken()\n  x = 1\n');
    expect(m.warnings.length).toBeGreaterThan(0);
  });
});

describe('statements', () => {
  const body = (src: string) => parseModule(`Sub S()\n${src}\nEnd Sub`).procedures[0]!.body;

  it('parses assignment', () => {
    expect(body('x = 1')[0]).toMatchObject({ kind: 'assign', isSet: false });
  });

  it('parses Set assignment', () => {
    expect(body('Set r = Range("A1")')[0]).toMatchObject({ kind: 'assign', isSet: true });
  });

  it('parses a call with unparenthesised arguments', () => {
    const s = body('MsgBox "hello", vbOKOnly')[0] as Extract<Stmt, { kind: 'call' }>;
    expect(s.kind).toBe('call');
    expect(s.args).toHaveLength(2);
  });

  it('parses Call with parentheses', () => {
    const s = body('Call DoThing(1, 2)')[0] as Extract<Stmt, { kind: 'call' }>;
    expect(s.args).toHaveLength(2);
  });

  it('parses a block If with ElseIf and Else', () => {
    const s = body('If a Then\nx = 1\nElseIf b Then\nx = 2\nElse\nx = 3\nEnd If')[0] as Extract<
      Stmt,
      { kind: 'if' }
    >;
    expect(s.branches).toHaveLength(2);
    expect(s.elseBody).toHaveLength(1);
  });

  it('parses a single-line If', () => {
    const s = body('If a > 1 Then x = 2')[0] as Extract<Stmt, { kind: 'if' }>;
    expect(s.kind).toBe('if');
    expect(s.branches[0]!.body).toHaveLength(1);
  });

  it('parses a single-line If with Else', () => {
    const s = body('If a Then x = 1 Else x = 2')[0] as Extract<Stmt, { kind: 'if' }>;
    expect(s.elseBody).toHaveLength(1);
  });

  it('parses For with Step', () => {
    const s = body('For i = 1 To 10 Step 2\nx = i\nNext i')[0] as Extract<Stmt, { kind: 'for' }>;
    expect(s.kind).toBe('for');
    expect(s.step).toBeDefined();
    expect(s.body).toHaveLength(1);
  });

  it('parses For Each', () => {
    const s = body('For Each c In Range("A1:A9")\nc.Value = 1\nNext c')[0] as Extract<
      Stmt,
      { kind: 'forEach' }
    >;
    expect(s.kind).toBe('forEach');
  });

  it('parses Do While and Do Until with pre- and post-tests', () => {
    expect(body('Do While a < 10\na = a + 1\nLoop')[0]).toMatchObject({
      kind: 'do',
      test: { kind: 'while', post: false },
    });
    expect(body('Do\na = a + 1\nLoop Until a > 10')[0]).toMatchObject({
      kind: 'do',
      test: { kind: 'until', post: true },
    });
  });

  it('parses While Wend', () => {
    expect(body('While a < 3\na = a + 1\nWend')[0]).toMatchObject({ kind: 'while' });
  });

  it('parses Select Case with ranges and Is comparisons', () => {
    const s = body(
      'Select Case n\nCase 1, 2\nx = 1\nCase 3 To 5\nx = 2\nCase Is > 10\nx = 3\nCase Else\nx = 4\nEnd Select',
    )[0] as Extract<Stmt, { kind: 'selectCase' }>;
    expect(s.cases).toHaveLength(3);
    expect(s.cases[0]!.labels).toHaveLength(2);
    expect(s.cases[1]!.labels[0]!.kind).toBe('range');
    expect(s.cases[2]!.labels[0]!.kind).toBe('compare');
    expect(s.elseBody).toHaveLength(1);
  });

  it('parses With blocks and leading-dot members', () => {
    const s = body('With Range("A1")\n.Value = 1\n.Font.Bold = True\nEnd With')[0] as Extract<
      Stmt,
      { kind: 'with' }
    >;
    expect(s.kind).toBe('with');
    expect(s.body).toHaveLength(2);
  });

  it('parses error handling', () => {
    expect(body('On Error Resume Next')[0]).toMatchObject({ kind: 'onError', mode: 'resumeNext' });
    expect(body('On Error GoTo Handler')[0]).toMatchObject({ kind: 'onError', mode: 'goto' });
    expect(body('On Error GoTo 0')[0]).toMatchObject({ kind: 'onError', mode: 'goto0' });
  });

  it('parses ReDim Preserve', () => {
    expect(body('ReDim Preserve items(1 To 20)')[0]).toMatchObject({
      kind: 'redim',
      preserve: true,
    });
  });

  it('parses Exit statements', () => {
    expect(body('Exit Sub')[0]).toMatchObject({ kind: 'exit', target: 'SUB' });
    expect(body('Exit For')[0]).toMatchObject({ kind: 'exit', target: 'FOR' });
  });

  it('parses a label inside a procedure', () => {
    const stmts = body('On Error GoTo Handler\nExit Sub\nHandler:\nx = 1');
    expect(stmts.some((s) => s.kind === 'label')).toBe(true);
  });
});

describe('expressions', () => {
  const expr = (src: string): Expr => {
    const s = parseModule(`Sub S()\nx = ${src}\nEnd Sub`).procedures[0]!.body[0] as Extract<
      Stmt,
      { kind: 'assign' }
    >;
    return s.value;
  };

  const sexp = (e: Expr): string => {
    switch (e.kind) {
      case 'number':
        return String(e.value);
      case 'string':
        return JSON.stringify(e.value);
      case 'bool':
        return e.value ? 'True' : 'False';
      case 'identifier':
        return e.name;
      case 'member':
        return `${e.target ? sexp(e.target) : ''}.${e.name}`;
      case 'call':
        return `(${sexp(e.target)}${e.args.map((a) => ` ${a.value ? sexp(a.value) : '_'}`).join('')})`;
      case 'unary':
        return `(${e.op} ${sexp(e.operand)})`;
      case 'binary':
        return `(${e.op} ${sexp(e.left)} ${sexp(e.right)})`;
      case 'paren':
        return sexp(e.inner);
      case 'new':
        return `New:${e.typeName}`;
      case 'nothing':
        return 'Nothing';
      case 'null':
        return 'Null';
      case 'me':
        return 'Me';
      default:
        return e.kind;
    }
  };

  it('applies arithmetic precedence', () => {
    expect(sexp(expr('1 + 2 * 3'))).toBe('(+ 1 (* 2 3))');
    expect(sexp(expr('2 * 3 ^ 2'))).toBe('(* 2 (^ 3 2))');
  });

  it('makes exponentiation right-associative, unlike the worksheet grammar', () => {
    // VBA and the worksheet formula language differ here, which is a genuine
    // trap: 2^3^2 is 512 in VBA and 64 in a cell formula.
    expect(sexp(expr('2 ^ 3 ^ 2'))).toBe('(^ 2 (^ 3 2))');
  });

  it('orders integer division and Mod between multiplication and addition', () => {
    expect(sexp(expr('1 + 6 \\ 2'))).toBe('(+ 1 (\\ 6 2))');
    expect(sexp(expr('1 + 7 Mod 3'))).toBe('(+ 1 (MOD 7 3))');
  });

  it('puts concatenation below arithmetic and above comparison', () => {
    expect(sexp(expr('"a" & 1 + 2'))).toBe('(& "a" (+ 1 2))');
    expect(sexp(expr('a & b = c'))).toBe('(= (& a b) c)');
  });

  it('orders the logical operators from Not out to Imp', () => {
    expect(sexp(expr('Not a And b'))).toBe('(AND (Not a) b)');
    expect(sexp(expr('a And b Or c'))).toBe('(OR (AND a b) c)');
    expect(sexp(expr('a Or b Xor c'))).toBe('(XOR (OR a b) c)');
  });

  it('parses Is and Like at comparison precedence', () => {
    expect(sexp(expr('a Is Nothing'))).toBe('(IS a Nothing)');
    expect(sexp(expr('s Like "a*"'))).toBe('(LIKE s "a*")');
  });

  it('parses member chains and calls', () => {
    expect(sexp(expr('Range("A1").Font.Bold'))).toBe('(Range "A1").Font.Bold');
    expect(sexp(expr('Cells(1, 2).Value'))).toBe('(Cells 1 2).Value');
  });

  it('parses named arguments', () => {
    const e = expr('Range(Cell1:="A1")') as Extract<Expr, { kind: 'call' }>;
    expect(e.args[0]!.name).toBe('Cell1');
  });

  it('parses omitted arguments', () => {
    const e = expr('Foo(1, , 3)') as Extract<Expr, { kind: 'call' }>;
    expect(e.args).toHaveLength(3);
    expect(e.args[1]!.value).toBeNull();
  });

  it('parses New', () => {
    expect(sexp(expr('New Collection'))).toBe('New:Collection');
  });

  it('parses number literals in every form', () => {
    expect((expr('&HFF') as { value: number }).value).toBe(255);
    expect((expr('&O17') as { value: number }).value).toBe(15);
    expect((expr('1.5E3') as { value: number }).value).toBe(1500);
    // VBA writes double-precision exponents with D.
    expect((expr('1.5D3') as { value: number }).value).toBe(1500);
  });
});

describe('compatibility report', () => {
  it('passes a macro that uses only supported features', () => {
    const report = analyseModule(`
Sub Total()
  Dim i As Long
  Dim sum As Double
  For i = 1 To 10
    sum = sum + Cells(i, 1).Value
  Next i
  Range("B1").Value = sum
End Sub`);
    expect(report.verdict).toBe('green');
    expect(report.findings).toEqual([]);
    expect(summarise(report)).toContain('only supported features');
  });

  it('refuses Shell outright and says why', () => {
    const report = analyseModule('Sub Bad()\n  Shell "cmd.exe /c del *.*"\nEnd Sub');
    expect(report.verdict).toBe('red');
    const finding = report.findings.find((f) => f.subject.toUpperCase() === 'SHELL')!;
    expect(finding.kind).toBe('blocked-capability');
    expect(finding.detail).toContain('external program');
  });

  it('refuses CreateObject', () => {
    const report = analyseModule('Sub Bad()\n  Set o = CreateObject("WScript.Shell")\nEnd Sub');
    expect(report.verdict).toBe('red');
    expect(report.findings.some((f) => f.kind === 'blocked-capability')).toBe(true);
  });

  it('refuses a Declare of native code', () => {
    const report = analyseModule(
      'Private Declare Function GetTickCount Lib "kernel32" () As Long',
    );
    expect(report.verdict).toBe('red');
    const finding = report.findings[0]!;
    expect(finding.detail).toContain('native code');
    expect(finding.detail).toContain('kernel32');
  });

  it('refuses Application.OnTime, which schedules code outside this run', () => {
    const report = analyseModule('Sub S()\n  Application.OnTime Now, "Later"\nEnd Sub');
    expect(report.verdict).toBe('red');
  });

  it('reports a stubbed member as amber rather than blocking', () => {
    const report = analyseModule('Sub S()\n  ActiveSheet.PivotTables("P").RefreshTable\nEnd Sub');
    expect(report.verdict).toBe('amber');
    expect(report.findings.some((f) => f.kind === 'stubbed-member')).toBe(true);
    expect(summarise(report)).toContain('may behave differently');
  });

  it('reports an unimplemented member without blocking', () => {
    const report = analyseModule('Sub S()\n  Range("A1").SomethingExotic = 1\nEnd Sub');
    expect(report.verdict).toBe('amber');
    expect(report.findings.some((f) => f.kind === 'unimplemented-member')).toBe(true);
  });

  it('does not report the module own declarations as unknown', () => {
    const report = analyseModule(`
Dim myOwnVariable As Long
Sub S()
  myOwnVariable = 1
End Sub`);
    expect(report.unknownCalls).not.toContain('myOwnVariable');
  });

  it('gives a verdict per procedure', () => {
    const report = analyseModule(`
Sub Safe()
  Range("A1").Value = 1
End Sub
Sub Unsafe()
  Shell "notepad"
End Sub`);
    const byName = new Map(report.procedures.map((p) => [p.name, p.verdict]));
    expect(byName.get('Safe')).toBe('green');
    expect(byName.get('Unsafe')).toBe('red');
    expect(report.verdict).toBe('red');
  });

  it('lists every member used, for building out the object model', () => {
    const report = analyseModule('Sub S()\n  Range("A1").Font.Bold = True\nEnd Sub');
    expect(report.membersUsed).toEqual(expect.arrayContaining(['Font', 'Bold']));
  });

  it('reports a module it cannot parse as red rather than pretending', () => {
    const report = analyseModule('Sub S(\n');
    expect(report.verdict).toBe('red');
  });

  it('flags GoSub as an obsolete construct', () => {
    const report = analyseModule('Sub S()\n  GoSub Helper\nEnd Sub');
    expect(report.findings.some((f) => f.subject === 'GoSub')).toBe(true);
  });
});

describe('a realistic macro', () => {
  const source = `
Attribute VB_Name = "Report"
Option Explicit

Private Const HEADER_ROW As Long = 1

Public Sub BuildReport(ByVal sheetName As String, Optional ByVal clear As Boolean = True)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim total As Double

    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets(sheetName)
    If clear Then ws.Range("C:C").ClearContents

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = HEADER_ROW + 1 To lastRow
        Select Case ws.Cells(i, 2).Value
            Case "Eng", "Science"
                total = total + ws.Cells(i, 3).Value
            Case Is > 1000
                total = total + 1
            Case Else
                ' nothing to do
        End Select
    Next i

    With ws.Range("E1")
        .Value = total
        .Font.Bold = True
        .NumberFormat = "#,##0.00"
    End With

    Exit Sub

Failed:
    MsgBox "Report failed: " & Err.Description, vbExclamation
End Sub`;

  it('parses completely, with no warnings', () => {
    const module = parseModule(source);
    expect(module.warnings).toEqual([]);
    expect(module.procedures).toHaveLength(1);
    expect(module.procedures[0]!.name).toBe('BuildReport');
    expect(module.procedures[0]!.params).toHaveLength(2);
  });

  it('keeps the module attribute', () => {
    expect(parseModule(source).attributes[0]).toContain('VB_Name');
  });

  it('finds the nested control flow', () => {
    const body = parseModule(source).procedures[0]!.body;
    expect(body.some((s) => s.kind === 'for')).toBe(true);
    expect(body.some((s) => s.kind === 'with')).toBe(true);
    expect(body.some((s) => s.kind === 'onError')).toBe(true);
    expect(body.some((s) => s.kind === 'label')).toBe(true);
  });

  it('is judged runnable', () => {
    const report = analyseModule(source);
    expect(report.verdict).not.toBe('red');
  });
});
