/**
 * MIRRORZ command line.
 *
 * The desktop application is the product, but a command line makes the whole
 * engine usable now, scriptable, and testable end to end without a display -
 * which is also how the conversion and inspection paths get exercised in CI.
 *
 * Every command is read-then-write with the same preservation guarantees the
 * app has: an unmodelled part of a workbook survives a conversion untouched.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { basename, extname } from 'node:path';
import {
  CellError,
  Document,
  type Scalar,
  Workbook,
  a1,
  colToName,
  isError,
  parseRangeRef,
} from '@mirrorz/core';
import {
  format as formatNumber,
  isDateFormat,
  looksLikeOds,
  looksLikeXls,
  parseDelimited,
  readOds,
  readXls,
  readXlsx,
  writeRows,
  writeXlsx,
} from '@mirrorz/formats';
import { type CellAddr, Engine, type RangeAddr, createRegistry } from '@mirrorz/formula';
import { analyseModule, parseVbaProject, summarise } from '@mirrorz/vba';

const USAGE = `MIRRORZ Sheets - a spreadsheet that opens Excel files

Usage:
  mirrorz info <file>                 summarise a workbook
  mirrorz show <file> [sheet] [range] print cells as a table
  mirrorz convert <in> <out>          convert between formats
  mirrorz calc <file> [--set A1=x]    recalculate and report
  mirrorz eval <file> <formula>       evaluate a formula against a workbook
  mirrorz explain <file> <sheet!A1>   why does this cell hold that value
  mirrorz macros <file>               list macros and check what they would do

Formats read:  .xlsx .xlsm .xltx .xltm .xls .csv .tsv .ods
Formats written: .xlsx .xlsm .xltx .xltm .csv .tsv

Options:
  --sheet <name>    restrict to one sheet
  --json            machine-readable output
  --max <n>         limit rows printed (default 50)
`;

export interface LoadedWorkbook {
  workbook: Workbook;
  warnings: string[];
  /** The package, when the source was an OPC file, so a save can preserve it. */
  source?: ReturnType<typeof readXlsx>['pkg'];
  styleTables?: ReturnType<typeof readXlsx>['styleTables'];
  flavour?: ReturnType<typeof readXlsx>['flavour'];
}

/** Open any supported file, choosing the reader by content rather than by name. */
export function load(path: string): LoadedWorkbook {
  const bytes = new Uint8Array(readFileSync(path));
  const ext = extname(path).toLowerCase();

  if (ext === '.csv' || ext === '.tsv') {
    const text = new TextDecoder().decode(bytes);
    const parsed = parseDelimited(text, ext === '.tsv' ? { delimiter: '\t' } : {});
    const workbook = new Workbook();
    const sheet = workbook.addSheet(basename(path, ext));
    parsed.rows.forEach((row, r) => {
      row.forEach((value, c) => {
        if (value !== null && value !== '') sheet.setValue(r, c, value as Scalar);
      });
    });
    // CSV warnings are structured; flatten them for the shared string list.
    return {
      workbook,
      warnings: (parsed.warnings ?? []).map(
        (w) => `${w.message}${w.row === undefined ? '' : ` (row ${w.row + 1})`}`,
      ),
    };
  }

  // Sniff the container rather than trusting the extension: a .xls that is
  // really a zip, or an .xlsx that is really a compound file, are both common
  // enough that guessing from the name is worse than looking.
  if (looksLikeXls(bytes)) {
    const r = readXls(bytes);
    return { workbook: r.workbook, warnings: r.warnings };
  }
  if (looksLikeOds(bytes)) {
    const r = readOds(bytes);
    return { workbook: r.workbook, warnings: r.warnings };
  }

  const r = readXlsx(bytes);
  return {
    workbook: r.workbook,
    warnings: r.warnings,
    source: r.pkg,
    styleTables: r.styleTables,
    flavour: r.flavour,
  };
}

/** Save a workbook, choosing the writer by the target extension. */
export function save(loaded: LoadedWorkbook, path: string): void {
  const ext = extname(path).toLowerCase();

  if (ext === '.csv' || ext === '.tsv') {
    const sheet = loaded.workbook.sheets[0];
    if (!sheet) throw new Error('workbook has no sheets to export');
    const bounds = sheet.bounds();
    const rows: string[][] = [];
    if (bounds) {
      for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
        const row: string[] = [];
        for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
          row.push(displayValue(sheet.getValue(r, c), undefined));
        }
        rows.push(row);
      }
    }
    writeFileSync(path, writeRows(rows, { delimiter: ext === '.tsv' ? '\t' : ',' }));
    return;
  }

  const bytes = writeXlsx(loaded.workbook, {
    ...(loaded.source ? { source: loaded.source } : {}),
    ...(loaded.styleTables ? { styleTables: loaded.styleTables } : {}),
    ...(ext === '.xlsm' ? { flavour: 'xlsm' as const } : {}),
  });
  writeFileSync(path, bytes);
}

/** Render a value the way the grid would, honouring its number format. */
export function displayValue(value: Scalar, numFmt: string | undefined): string {
  if (value === null) return '';
  if (isError(value)) return value.code;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number' && numFmt) {
    try {
      // The format engine's answer is taken as given, empty included: `;;;` is
      // Excel's idiom for hiding a cell, and second-guessing an empty result
      // would break it. Only an outright throw falls back.
      return formatNumber(value, numFmt).text;
    } catch {
      return String(value);
    }
  }
  return String(value);
}

export async function main(argv: string[]): Promise<number> {
  const args = argv.slice(2);
  const command = args[0];
  if (!command || command === '--help' || command === '-h') {
    process.stdout.write(USAGE);
    return 0;
  }

  const flags = new Map<string, string>();
  const positional: string[] = [];
  for (let i = 1; i < args.length; i++) {
    const a = args[i]!;
    if (a.startsWith('--')) {
      const eq = a.indexOf('=');
      if (eq > 0) flags.set(a.slice(2, eq), a.slice(eq + 1));
      else if (args[i + 1] && !args[i + 1]!.startsWith('--')) flags.set(a.slice(2), args[++i]!);
      else flags.set(a.slice(2), 'true');
    } else {
      positional.push(a);
    }
  }
  const json = flags.has('json');

  try {
    switch (command) {
      case 'info':
        return info(positional[0], json);
      case 'show':
        return show(positional[0], positional[1], positional[2], flags, json);
      case 'convert':
        return convert(positional[0], positional[1], json);
      case 'calc':
        return calc(positional[0], flags, json);
      case 'eval':
        return evaluate(positional[0], positional[1], flags, json);
      case 'explain':
        return explain(positional[0], positional[1], json);
      case 'macros':
        return macros(positional[0], json);
      default:
        process.stderr.write(`unknown command: ${command}\n\n${USAGE}`);
        return 2;
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    if (json) process.stdout.write(`${JSON.stringify({ error: message })}\n`);
    else process.stderr.write(`error: ${message}\n`);
    return 1;
  }
}

function requirePath(path: string | undefined, what: string): string {
  if (!path) throw new Error(`missing ${what}`);
  return path;
}

function info(path: string | undefined, json: boolean): number {
  const file = requirePath(path, 'file');
  const { workbook, warnings } = load(file);

  const sheets = workbook.sheets.map((s) => {
    const b = s.bounds();
    let formulas = 0;
    for (const { cell } of s.entries()) if (cell.formula) formulas++;
    return {
      name: s.name,
      visibility: s.visibility,
      cells: s.cellCount,
      formulas,
      extent: b ? `${a1(b.minRow, b.minCol)}:${a1(b.maxRow, b.maxCol)}` : 'empty',
    };
  });

  const summary = {
    file: basename(file),
    sheets,
    definedNames: workbook.definedNames.length,
    dateSystem: workbook.dateSystem,
    styles: workbook.styles.size,
    hasMacros: workbook.vbaProject !== undefined,
    warnings,
  };

  if (json) {
    process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
    return 0;
  }

  process.stdout.write(`${summary.file}\n`);
  process.stdout.write(`  ${sheets.length} sheet${sheets.length === 1 ? '' : 's'}`);
  process.stdout.write(`, ${workbook.totalCells} cells`);
  process.stdout.write(`, ${summary.definedNames} defined name${summary.definedNames === 1 ? '' : 's'}`);
  process.stdout.write(`, ${summary.dateSystem} date system`);
  if (summary.hasMacros) process.stdout.write(', contains macros');
  process.stdout.write('\n\n');
  for (const s of sheets) {
    const hidden = s.visibility === 'visible' ? '' : ` (${s.visibility})`;
    process.stdout.write(
      `  ${s.name}${hidden}  ${s.extent}  ${s.cells} cells, ${s.formulas} formulas\n`,
    );
  }
  if (warnings.length > 0) {
    process.stdout.write('\nwarnings:\n');
    for (const w of warnings) process.stdout.write(`  ${w}\n`);
  }
  return 0;
}

function show(
  path: string | undefined,
  sheetName: string | undefined,
  rangeText: string | undefined,
  flags: Map<string, string>,
  json: boolean,
): number {
  const file = requirePath(path, 'file');
  const { workbook } = load(file);
  const sheet = sheetName ? workbook.getSheet(sheetName) : workbook.sheets[0];
  if (!sheet) throw new Error(`no such sheet: ${sheetName ?? '(first)'}`);

  const bounds = sheet.bounds();
  if (!bounds) {
    process.stdout.write('(empty sheet)\n');
    return 0;
  }

  const range = rangeText ? parseRangeRef(rangeText) : undefined;
  const maxRows = Number(flags.get('max') ?? '50');
  const startRow = range?.start.row ?? bounds.minRow;
  const startCol = range?.start.col ?? bounds.minCol;
  const endRow = Math.min(range?.end.row ?? bounds.maxRow, startRow + maxRows - 1);
  const endCol = Math.min(range?.end.col ?? bounds.maxCol, startCol + 63);

  const grid: string[][] = [];
  for (let r = startRow; r <= endRow; r++) {
    const row: string[] = [];
    for (let c = startCol; c <= endCol; c++) {
      const style = workbook.styles.get(sheet.getStyle(r, c));
      row.push(displayValue(sheet.getValue(r, c), style.numFmt));
    }
    grid.push(row);
  }

  if (json) {
    process.stdout.write(`${JSON.stringify({ sheet: sheet.name, rows: grid }, null, 2)}\n`);
    return 0;
  }

  // Column widths from the content, so the table lines up.
  const widths: number[] = [];
  for (let c = 0; c <= endCol - startCol; c++) {
    let w = colToName(startCol + c).length;
    for (const row of grid) w = Math.max(w, (row[c] ?? '').length);
    widths.push(Math.min(w, 24));
  }

  const rowLabelWidth = String(endRow + 1).length;
  process.stdout.write(`${' '.repeat(rowLabelWidth)} `);
  widths.forEach((w, c) => process.stdout.write(`${colToName(startCol + c).padEnd(w)} `));
  process.stdout.write('\n');

  grid.forEach((row, i) => {
    process.stdout.write(`${String(startRow + i + 1).padStart(rowLabelWidth)} `);
    row.forEach((cell, c) => {
      const w = widths[c]!;
      const text = cell.length > w ? `${cell.slice(0, w - 1)}…` : cell;
      process.stdout.write(`${text.padEnd(w)} `);
    });
    process.stdout.write('\n');
  });
  return 0;
}

function convert(from: string | undefined, to: string | undefined, json: boolean): number {
  const input = requirePath(from, 'input file');
  const output = requirePath(to, 'output file');
  const loaded = load(input);
  save(loaded, output);

  const message = {
    converted: basename(input),
    to: basename(output),
    sheets: loaded.workbook.sheets.length,
    cells: loaded.workbook.totalCells,
    preserved: loaded.source ? loaded.source.order.length : 0,
    warnings: loaded.warnings,
  };
  if (json) process.stdout.write(`${JSON.stringify(message, null, 2)}\n`);
  else {
    process.stdout.write(`${message.converted} -> ${message.to}\n`);
    process.stdout.write(`  ${message.sheets} sheets, ${message.cells} cells`);
    if (message.preserved) process.stdout.write(`, ${message.preserved} parts preserved`);
    process.stdout.write('\n');
    for (const w of message.warnings) process.stdout.write(`  warning: ${w}\n`);
  }
  return 0;
}

function calc(path: string | undefined, flags: Map<string, string>, json: boolean): number {
  const file = requirePath(path, 'file');
  const loaded = load(file);
  const doc = new Document(loaded.workbook);
  const engine = new Engine(doc, createRegistry());
  engine.indexWorkbook();

  // `--set Sheet1!A1=42` applies an edit before recalculating, which is what
  // makes the command useful for checking a model rather than just re-running it.
  const set = flags.get('set');
  if (set) {
    const eq = set.indexOf('=');
    if (eq < 0) throw new Error('--set expects CELL=VALUE');
    const target = set.slice(0, eq);
    const raw = set.slice(eq + 1);
    const bang = target.indexOf('!');
    const sheetName = bang < 0 ? loaded.workbook.sheets[0]!.name : target.slice(0, bang);
    const ref = parseRangeRef(bang < 0 ? target : target.slice(bang + 1));
    if (!ref) throw new Error(`--set: ${target} is not a cell reference`);
    const numeric = Number(raw);
    engine.setCell(
      { sheet: sheetName, row: ref.start.row, col: ref.start.col },
      raw.startsWith('=')
        ? { value: null, formula: raw.slice(1) }
        : { value: raw === '' ? null : Number.isFinite(numeric) && raw.trim() !== '' ? numeric : raw },
    );
  }

  const result = engine.recalculateAll();
  const changed = result.changed.map((c: CellAddr) => `${c.sheet}!${a1(c.row, c.col)}`);

  if (json) {
    process.stdout.write(
      `${JSON.stringify(
        {
          evaluated: result.evaluated.length,
          changed,
          circular: result.circular.map((c: CellAddr) => `${c.sheet}!${a1(c.row, c.col)}`),
          elapsedMs: result.elapsedMs,
        },
        null,
        2,
      )}\n`,
    );
    return 0;
  }

  process.stdout.write(
    `recalculated ${result.evaluated.length} formula${result.evaluated.length === 1 ? '' : 's'} in ${result.elapsedMs}ms\n`,
  );
  if (changed.length > 0) {
    process.stdout.write(`  ${changed.length} changed\n`);
    for (const c of changed.slice(0, 20)) process.stdout.write(`    ${c}\n`);
    if (changed.length > 20) process.stdout.write(`    ... and ${changed.length - 20} more\n`);
  }
  if (result.circular.length > 0) {
    process.stdout.write(`  ${result.circular.length} cells form circular references\n`);
  }
  return 0;
}

function evaluate(
  path: string | undefined,
  formula: string | undefined,
  flags: Map<string, string>,
  json: boolean,
): number {
  const file = requirePath(path, 'file');
  const expression = requirePath(formula, 'formula');
  const loaded = load(file);
  const doc = new Document(loaded.workbook);
  const engine = new Engine(doc, createRegistry());
  engine.indexWorkbook();

  const sheetName = flags.get('sheet') ?? loaded.workbook.sheets[0]?.name;
  if (!sheetName) throw new Error('workbook has no sheets');

  // Evaluate in a scratch cell well outside the used range, so the expression
  // cannot collide with real data.
  const bounds = loaded.workbook.getSheet(sheetName)?.bounds();
  const row = (bounds?.maxRow ?? 0) + 2;
  const col = (bounds?.maxCol ?? 0) + 2;
  const text = expression.startsWith('=') ? expression.slice(1) : expression;
  engine.setCell({ sheet: sheetName, row, col }, { value: null, formula: text });
  const value = loaded.workbook.getSheet(sheetName)!.getValue(row, col);

  if (json) {
    process.stdout.write(
      `${JSON.stringify({ formula: text, value: isError(value) ? value.code : value })}\n`,
    );
  } else {
    process.stdout.write(`${displayValue(value, undefined)}\n`);
  }
  return isError(value) ? 1 : 0;
}

function explain(path: string | undefined, target: string | undefined, json: boolean): number {
  const file = requirePath(path, 'file');
  const address = requirePath(target, 'cell reference');
  const loaded = load(file);
  const doc = new Document(loaded.workbook);
  const engine = new Engine(doc, createRegistry());
  engine.indexWorkbook();

  const bang = address.indexOf('!');
  const sheetName = bang < 0 ? loaded.workbook.sheets[0]!.name : address.slice(0, bang);
  const ref = parseRangeRef(bang < 0 ? address : address.slice(bang + 1));
  if (!ref) throw new Error(`${address} is not a cell reference`);

  const explanation = engine.explain({
    sheet: sheetName,
    row: ref.start.row,
    col: ref.start.col,
  });

  const label = (c: { sheet: string; row: number; col: number }) =>
    `${c.sheet}!${a1(c.row, c.col)}`;

  if (json) {
    process.stdout.write(
      `${JSON.stringify(
        {
          cell: label(explanation.addr),
          value: isError(explanation.value) ? explanation.value.code : explanation.value,
          formula: explanation.formula,
          precedents: explanation.precedentCells.map(label),
          dependents: explanation.dependents.map(label),
          errorRoots: explanation.errorRoots.map(label),
        },
        null,
        2,
      )}\n`,
    );
    return 0;
  }

  process.stdout.write(`${label(explanation.addr)}\n`);
  if (explanation.formula) process.stdout.write(`  formula   =${explanation.formula}\n`);
  process.stdout.write(`  value     ${displayValue(explanation.value, undefined)}\n`);
  if (explanation.precedentCells.length > 0) {
    process.stdout.write(`  reads     ${explanation.precedentCells.map(label).join(', ')}\n`);
  }
  if (explanation.precedentRanges.length > 0) {
    const ranges = explanation.precedentRanges.map(
      (r: RangeAddr) => `${r.sheet}!${a1(r.startRow, r.startCol)}:${a1(r.endRow, r.endCol)}`,
    );
    process.stdout.write(`  ranges    ${ranges.join(', ')}\n`);
  }
  if (explanation.dependents.length > 0) {
    process.stdout.write(`  used by   ${explanation.dependents.map(label).join(', ')}\n`);
  }
  if (explanation.errorRoots.length > 0) {
    // The whole point: name the cell that caused it, not the chain that carried it.
    process.stdout.write(
      `\n  the problem starts at ${explanation.errorRoots.map(label).join(' and ')}\n`,
    );
  }
  return 0;
}

function macros(path: string | undefined, json: boolean): number {
  const file = requirePath(path, 'file');
  const { workbook } = load(file);
  if (!workbook.vbaProject) {
    if (json) process.stdout.write(`${JSON.stringify({ macros: [] })}\n`);
    else process.stdout.write('This workbook contains no macros.\n');
    return 0;
  }

  const project = parseVbaProject(workbook.vbaProject);
  const reports = project.modules.map((m: { name: string; type: string; source: string }) => ({
    name: m.name,
    type: m.type,
    lines: m.source.split('\n').length,
    report: analyseModule(m.source, m.name),
  }));

  if (json) {
    process.stdout.write(
      `${JSON.stringify(
        reports.map((r: (typeof reports)[number]) => ({
          name: r.name,
          type: r.type,
          lines: r.lines,
          verdict: r.report.verdict,
          findings: r.report.findings,
        })),
        null,
        2,
      )}\n`,
    );
    return 0;
  }

  process.stdout.write(`${reports.length} module${reports.length === 1 ? '' : 's'}`);
  if (project.protected) process.stdout.write(' (project is locked)');
  process.stdout.write('\n\n');

  for (const r of reports) {
    const mark = r.report.verdict === 'green' ? 'ok  ' : r.report.verdict === 'amber' ? 'warn' : 'stop';
    process.stdout.write(`  [${mark}] ${r.name} (${r.type}, ${r.lines} lines)\n`);
    process.stdout.write(`         ${summarise(r.report)}\n`);
    for (const f of r.report.findings.slice(0, 8)) {
      process.stdout.write(`         - ${f.subject}: ${f.detail}\n`);
    }
    if (r.report.findings.length > 8) {
      process.stdout.write(`         - ... and ${r.report.findings.length - 8} more\n`);
    }
  }
  return 0;
}

export { CellError };
