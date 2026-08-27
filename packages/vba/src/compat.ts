/**
 * Static compatibility analysis for a VBA project.
 *
 * The defining failure of every other non-Microsoft VBA implementation is silent
 * breakage: the macro runs, hits something unsupported, and either stops
 * halfway or quietly does the wrong thing to the user's data. By then the damage
 * is done and the user has no idea which line caused it.
 *
 * So before running anything, we walk the parsed tree and produce a verdict:
 *
 *   green  - every construct and every member used is supported
 *   amber  - it will run, but some members are stubs and behaviour may differ
 *   red    - it uses something we refuse to run, or cannot parse
 *
 * The report names the specific construct and line, so a user deciding whether
 * to trust a macro is looking at facts rather than at a shrug. This is useful on
 * its own, before a single line can be executed.
 *
 * Some things are red on purpose rather than through lack of implementation.
 * `Declare ... Lib` calls arbitrary native code; `Shell`, `CreateObject` and the
 * FileSystemObject reach outside the workbook entirely. A macro arriving by
 * email is the classic malware vector, so those are refused by policy and the
 * report says so in those terms.
 */

import type { Expr, Module, Procedure, Stmt } from './parser.js';
import { parseModule } from './parser.js';

export type Verdict = 'green' | 'amber' | 'red';

export type FindingKind =
  | 'blocked-capability'
  | 'unsupported-construct'
  | 'unimplemented-member'
  | 'parse-failure'
  | 'stubbed-member';

export interface CompatFinding {
  kind: FindingKind;
  /** The identifier or construct at fault. */
  subject: string;
  /** Why it matters, in one sentence a user can act on. */
  detail: string;
  severity: 'red' | 'amber';
  line?: number;
  procedure?: string;
}

export interface CompatReport {
  verdict: Verdict;
  findings: CompatFinding[];
  /** Every distinct member accessed, for building out the object model. */
  membersUsed: string[];
  /** Every function or sub called that we did not recognise. */
  unknownCalls: string[];
  procedures: { name: string; kind: string; verdict: Verdict }[];
}

/**
 * Capabilities refused by policy, not by omission.
 *
 * Each of these can reach outside the workbook. A spreadsheet macro that wants
 * them is either doing something genuinely unusual or is hostile, and in both
 * cases refusing and saying why beats running it and hoping.
 */
export const BLOCKED_CAPABILITIES: Readonly<Record<string, string>> = {
  SHELL: 'runs an external program',
  CREATEOBJECT: 'creates an arbitrary COM object, which can reach anything on the machine',
  GETOBJECT: 'attaches to an arbitrary COM object',
  'SCRIPTING.FILESYSTEMOBJECT': 'reads and writes arbitrary files',
  'WSCRIPT.SHELL': 'runs an external program and edits the registry',
  KILL: 'deletes files',
  FILECOPY: 'copies files outside the workbook',
  MKDIR: 'creates directories',
  RMDIR: 'removes directories',
  SETATTR: 'changes file attributes',
  OPEN: 'opens a file for direct read or write',
  SENDKEYS: 'synthesises keystrokes into other applications',
  APPACTIVATE: 'brings another application to the foreground',
  DDEINITIATE: 'opens a channel to another application',
  ENVIRON: 'reads environment variables',
  'APPLICATION.REGISTRYREAD': 'reads the registry',
  'APPLICATION.REGISTRYWRITE': 'writes the registry',
  'APPLICATION.ONTIME': 'schedules code to run later, outside the current run',
  'APPLICATION.EXECUTEEXCEL4MACRO': 'runs a legacy macro language with no sandbox',
};

/**
 * The Excel object-model surface we implement, or intend to. A member outside
 * this list is reported so the user knows what will be missing, and so the list
 * itself is driven by what real macros actually use.
 */
export const SUPPORTED_MEMBERS: ReadonlySet<string> = new Set([
  // Application
  'APPLICATION', 'SCREENUPDATING', 'CALCULATION', 'DISPLAYALERTS', 'ENABLEEVENTS',
  'WORKSHEETFUNCTION', 'ACTIVEWORKBOOK', 'ACTIVESHEET', 'ACTIVECELL', 'SELECTION',
  'THISWORKBOOK', 'WORKBOOKS', 'CALCULATE',
  // Workbook
  'WORKSHEETS', 'SHEETS', 'NAME', 'FULLNAME', 'PATH', 'SAVE', 'SAVEAS', 'CLOSE',
  'ADD', 'COUNT', 'ITEM', 'ACTIVATE', 'NAMES',
  // Worksheet
  'RANGE', 'CELLS', 'ROWS', 'COLUMNS', 'USEDRANGE', 'SELECT', 'VISIBLE', 'INDEX',
  'DELETE', 'COPY', 'MOVE', 'PROTECT', 'UNPROTECT', 'AUTOFILTER', 'SORT',
  // Range
  'VALUE', 'VALUE2', 'FORMULA', 'FORMULAR1C1', 'TEXT', 'ADDRESS', 'ROW', 'COLUMN',
  'OFFSET', 'RESIZE', 'ENTIREROW', 'ENTIRECOLUMN', 'CLEAR', 'CLEARCONTENTS',
  'CLEARFORMATS', 'INSERT', 'MERGE', 'UNMERGE', 'NUMBERFORMAT', 'INTERIOR',
  'FONT', 'BORDERS', 'HORIZONTALALIGNMENT', 'VERTICALALIGNMENT', 'WRAPTEXT',
  'COLUMNWIDTH', 'ROWHEIGHT', 'AUTOFIT', 'FIND', 'END', 'CURRENTREGION',
  'SPECIALCELLS', 'PASTESPECIAL', 'COLOR', 'COLORINDEX', 'BOLD', 'ITALIC',
  'UNDERLINE', 'SIZE', 'FORMULAHIDDEN', 'LOCKED', 'HIDDEN', 'COUNTLARGE',
  // Common VBA library
  'MSGBOX', 'INPUTBOX', 'LEN', 'LEFT', 'RIGHT', 'MID', 'TRIM', 'LTRIM', 'RTRIM',
  'UCASE', 'LCASE', 'INSTR', 'INSTRREV', 'REPLACE', 'SPLIT', 'JOIN', 'CSTR',
  'CINT', 'CLNG', 'CDBL', 'CBOOL', 'CDATE', 'CVAR', 'VAL', 'STR', 'FORMAT',
  'ISNUMERIC', 'ISEMPTY', 'ISNULL', 'ISERROR', 'ISDATE', 'ISARRAY', 'ISOBJECT',
  'ARRAY', 'UBOUND', 'LBOUND', 'ABS', 'INT', 'FIX', 'SGN', 'SQR', 'RND', 'ROUND',
  'NOW', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND',
  'DATEADD', 'DATEDIFF', 'DATEPART', 'DATESERIAL', 'TIMESERIAL', 'WEEKDAY',
  'CHR', 'ASC', 'SPACE', 'STRING', 'TYPENAME', 'VARTYPE', 'ERR', 'NUMBER',
  'DESCRIPTION', 'RAISE', 'CLEAR', 'DEBUG', 'PRINT',
]);

/** Members that will parse and run but whose behaviour is not yet complete. */
export const STUBBED_MEMBERS: ReadonlySet<string> = new Set([
  'PIVOTTABLES', 'PIVOTCACHES', 'CHARTOBJECTS', 'CHARTS', 'SHAPES', 'PICTURES',
  'QUERYTABLES', 'LISTOBJECTS', 'SPARKLINEGROUPS', 'SLICERS', 'COMMENTS',
  'HYPERLINKS', 'VALIDATION', 'FORMATCONDITIONS', 'OUTLINE', 'PAGESETUP',
  'PRINTOUT', 'EXPORTASFIXEDFORMAT', 'ONKEY', 'WINDOWS', 'PANES',
]);

export function analyseModule(source: string, moduleName = 'Module'): CompatReport {
  let module: Module;
  try {
    module = parseModule(source);
  } catch (err) {
    return {
      verdict: 'red',
      findings: [
        {
          kind: 'parse-failure',
          subject: moduleName,
          detail: `could not be parsed: ${err instanceof Error ? err.message : String(err)}`,
          severity: 'red',
        },
      ],
      membersUsed: [],
      unknownCalls: [],
      procedures: [],
    };
  }
  return analyse(module, moduleName);
}

export function analyse(module: Module, moduleName = 'Module'): CompatReport {
  const findings: CompatFinding[] = [];
  const membersUsed = new Set<string>();
  const unknownCalls = new Set<string>();
  const procedures: CompatReport['procedures'] = [];

  for (const warning of module.warnings) {
    findings.push({
      kind: 'parse-failure',
      subject: moduleName,
      detail: warning,
      severity: 'amber',
    });
  }

  const declared = new Set<string>();
  for (const p of module.procedures) declared.add(p.name.toUpperCase());
  for (const stmt of module.declarations) collectDeclared(stmt, declared);

  const inspectStatements = (statements: Stmt[], procedureName: string | undefined) => {
    for (const stmt of statements) {
      walkStatement(stmt, (s) => inspectStatement(s, procedureName, findings));
      walkStatementExpressions(stmt, (expr) =>
        inspectExpression(expr, procedureName, findings, membersUsed, unknownCalls, declared),
      );
    }
  };

  inspectStatements(module.declarations, undefined);

  for (const procedure of module.procedures) {
    const before = findings.length;
    inspectStatements(procedure.body, procedure.name);
    const own = findings.slice(before);
    procedures.push({
      name: procedure.name,
      kind: procedure.kind,
      verdict: own.some((f) => f.severity === 'red')
        ? 'red'
        : own.length > 0
          ? 'amber'
          : 'green',
    });
  }

  const verdict: Verdict = findings.some((f) => f.severity === 'red')
    ? 'red'
    : findings.length > 0
      ? 'amber'
      : 'green';

  return {
    verdict,
    findings,
    membersUsed: [...membersUsed].sort(),
    unknownCalls: [...unknownCalls].sort(),
    procedures,
  };
}

/** Names the module defines itself, which must not be reported as unknown. */
function collectDeclared(stmt: Stmt, into: Set<string>): void {
  switch (stmt.kind) {
    case 'dim':
      for (const v of stmt.vars) into.add(v.name.replace(/[%&@!#$]$/, '').toUpperCase());
      break;
    case 'const':
      for (const v of stmt.vars) into.add(v.name.replace(/[%&@!#$]$/, '').toUpperCase());
      break;
    case 'type':
      into.add(stmt.name.toUpperCase());
      break;
    case 'enum':
      into.add(stmt.name.toUpperCase());
      for (const m of stmt.members) into.add(m.name.toUpperCase());
      break;
    case 'declare':
      into.add(stmt.name.toUpperCase());
      break;
    default:
      break;
  }
}

function inspectStatement(
  stmt: Stmt,
  procedure: string | undefined,
  findings: CompatFinding[],
): void {
  if (stmt.kind === 'declare') {
    findings.push({
      kind: 'blocked-capability',
      subject: `Declare ${stmt.isFunction ? 'Function' : 'Sub'} ${stmt.name}`,
      detail: `calls native code in ${stmt.lib || 'an external library'}, which cannot be sandboxed and will not be run`,
      severity: 'red',
      ...(procedure === undefined ? {} : { procedure }),
    });
  }
  if (stmt.kind === 'gosub' || stmt.kind === 'return') {
    findings.push({
      kind: 'unsupported-construct',
      subject: stmt.kind === 'gosub' ? 'GoSub' : 'Return',
      detail: 'GoSub and Return are obsolete and not implemented',
      severity: 'amber',
      ...(procedure === undefined ? {} : { procedure }),
    });
  }
  if (stmt.kind === 'unknown') {
    findings.push({
      kind: 'unsupported-construct',
      subject: stmt.text.slice(0, 60),
      detail: 'this statement was not recognised',
      severity: 'amber',
      ...(procedure === undefined ? {} : { procedure }),
    });
  }
}

function inspectExpression(
  expr: Expr,
  procedure: string | undefined,
  findings: CompatFinding[],
  membersUsed: Set<string>,
  unknownCalls: Set<string>,
  declared: ReadonlySet<string>,
): void {
  walkExpression(expr, (e) => {
    if (e.kind === 'member') {
      const upper = e.name.replace(/[%&@!#$]$/, '').toUpperCase();
      membersUsed.add(e.name);

      // A qualified blocked capability, such as Application.OnTime.
      if (e.target?.kind === 'identifier') {
        const qualified = `${e.target.name.toUpperCase()}.${upper}`;
        const blocked = BLOCKED_CAPABILITIES[qualified];
        if (blocked) {
          findings.push({
            kind: 'blocked-capability',
            subject: qualified,
            detail: `${blocked}; refused by policy`,
            severity: 'red',
            ...(procedure === undefined ? {} : { procedure }),
          });
          return;
        }
      }

      if (STUBBED_MEMBERS.has(upper)) {
        findings.push({
          kind: 'stubbed-member',
          subject: e.name,
          detail: 'recognised but not fully implemented; behaviour may differ',
          severity: 'amber',
          ...(procedure === undefined ? {} : { procedure }),
        });
      } else if (!SUPPORTED_MEMBERS.has(upper) && !declared.has(upper)) {
        findings.push({
          kind: 'unimplemented-member',
          subject: e.name,
          detail: 'this member is not implemented',
          severity: 'amber',
          ...(procedure === undefined ? {} : { procedure }),
        });
      }
      return;
    }

    if (e.kind === 'identifier') {
      const upper = e.name.replace(/[%&@!#$]$/, '').toUpperCase();
      const blocked = BLOCKED_CAPABILITIES[upper];
      if (blocked) {
        findings.push({
          kind: 'blocked-capability',
          subject: e.name,
          detail: `${blocked}; refused by policy`,
          severity: 'red',
          ...(procedure === undefined ? {} : { procedure }),
        });
        return;
      }
      if (!SUPPORTED_MEMBERS.has(upper) && !declared.has(upper)) {
        unknownCalls.add(e.name);
      }
      return;
    }

    if (e.kind === 'new') {
      const upper = e.typeName.toUpperCase();
      const blocked = BLOCKED_CAPABILITIES[upper];
      if (blocked) {
        findings.push({
          kind: 'blocked-capability',
          subject: `New ${e.typeName}`,
          detail: `${blocked}; refused by policy`,
          severity: 'red',
          ...(procedure === undefined ? {} : { procedure }),
        });
      }
    }
  });
}

/** Visit a statement and everything nested inside it. */
export function walkStatement(stmt: Stmt, visit: (s: Stmt) => void): void {
  visit(stmt);
  switch (stmt.kind) {
    case 'if':
      for (const b of stmt.branches) for (const s of b.body) walkStatement(s, visit);
      if (stmt.elseBody) for (const s of stmt.elseBody) walkStatement(s, visit);
      break;
    case 'for':
    case 'forEach':
    case 'do':
    case 'while':
    case 'with':
      for (const s of stmt.body) walkStatement(s, visit);
      break;
    case 'selectCase':
      for (const c of stmt.cases) for (const s of c.body) walkStatement(s, visit);
      if (stmt.elseBody) for (const s of stmt.elseBody) walkStatement(s, visit);
      break;
    default:
      break;
  }
}

/** Visit every expression a statement contains, nested statements included. */
export function walkStatementExpressions(stmt: Stmt, visit: (e: Expr) => void): void {
  walkStatement(stmt, (s) => {
    switch (s.kind) {
      case 'assign':
        visit(s.target);
        visit(s.value);
        break;
      case 'call':
        visit(s.target);
        for (const a of s.args) if (a.value) visit(a.value);
        break;
      case 'if':
        for (const b of s.branches) visit(b.condition);
        break;
      case 'for':
        visit(s.variable);
        visit(s.from);
        visit(s.to);
        if (s.step) visit(s.step);
        break;
      case 'forEach':
        visit(s.variable);
        visit(s.collection);
        break;
      case 'do':
        if (s.test) visit(s.test.condition);
        break;
      case 'while':
        visit(s.condition);
        break;
      case 'selectCase':
        visit(s.subject);
        for (const c of s.cases) {
          for (const label of c.labels) {
            if (label.kind === 'value') visit(label.value);
            else if (label.kind === 'range') {
              visit(label.from);
              visit(label.to);
            } else visit(label.value);
          }
        }
        break;
      case 'with':
        visit(s.subject);
        break;
      case 'const':
        for (const v of s.vars) visit(v.value);
        break;
      case 'erase':
        for (const t of s.targets) visit(t);
        break;
      case 'raiseEvent':
        for (const a of s.args) if (a.value) visit(a.value);
        break;
      default:
        break;
    }
  });
}

export function walkExpression(expr: Expr, visit: (e: Expr) => void): void {
  visit(expr);
  switch (expr.kind) {
    case 'member':
      if (expr.target) walkExpression(expr.target, visit);
      break;
    case 'call':
    case 'index':
      walkExpression(expr.target, visit);
      for (const a of expr.args) if (a.value) walkExpression(a.value, visit);
      break;
    case 'unary':
      walkExpression(expr.operand, visit);
      break;
    case 'binary':
      walkExpression(expr.left, visit);
      walkExpression(expr.right, visit);
      break;
    case 'paren':
      walkExpression(expr.inner, visit);
      break;
    default:
      break;
  }
}

/** A one-line summary suitable for a status bar or a dialog heading. */
export function summarise(report: CompatReport): string {
  const red = report.findings.filter((f) => f.severity === 'red').length;
  const amber = report.findings.filter((f) => f.severity === 'amber').length;
  if (report.verdict === 'green') {
    return 'This macro uses only supported features.';
  }
  if (report.verdict === 'red') {
    return `This macro will not run: ${red} blocked or unsupported item${red === 1 ? '' : 's'}.`;
  }
  return `This macro should run, but ${amber} item${amber === 1 ? '' : 's'} may behave differently.`;
}

export type { Procedure };
