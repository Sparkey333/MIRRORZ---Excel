import { readFileSync } from 'node:fs';
import { isError } from '/home/user/MIRRORZ---Excel/packages/core/src/index.js';
import { readXlsx } from '/home/user/MIRRORZ---Excel/packages/formats/src/xlsx/read.js';
import { Evaluator } from '/home/user/MIRRORZ---Excel/packages/formula/src/evaluator.js';
import { createRegistry } from '/home/user/MIRRORZ---Excel/packages/formula/src/functions/index.js';
import { parseFormula } from '/home/user/MIRRORZ---Excel/packages/formula/src/parser.js';
import { WorkbookStore } from '/home/user/MIRRORZ---Excel/packages/formula/src/store.js';
import { isArray, toExcelPrecision } from '/home/user/MIRRORZ---Excel/packages/formula/src/value.js';

const { workbook } = readXlsx(new Uint8Array(readFileSync('scratch-tmp/probe.calc.xlsx')));
const sheet = workbook.getSheet('C')!;
const ev = new Evaluator(new WorkbookStore(workbook), createRegistry(), { dateSystem: workbook.dateSystem });
const show = (v: unknown): string => {
  if (isArray(v as never)) {
    const a = v as { rows: number; cols: number; data: unknown[] };
    return `[${a.rows}x${a.cols} ${a.data.map((x) => show(x)).join('|')}]`;
  }
  if (isError(v)) return (v as { code: string }).code;
  if (typeof v === 'number') return String(toExcelPrecision(v));
  return JSON.stringify(v);
};
let bad = 0;
for (const { row, col, cell } of sheet.entries()) {
  if (col !== 1 || !cell.formula) continue;
  const mine = ev.evaluate({
    ast: parseFormula(cell.formula.replaceAll('~', ','), { origin: { row, col: 1 } }),
    sheet: 'C', row, col: 1,
  });
  const a = show(mine);
  const b = show(cell.value);
  const near = typeof mine === 'number' && typeof cell.value === 'number'
    && (Math.abs(mine - cell.value) <= 1e-10 * Math.max(1, Math.abs(cell.value)));
  if (a !== b && !near) {
    bad++;
    console.log(`MISMATCH ${cell.formula}\n   mine=${a}\n     lo=${b}`);
  }
}
console.log(`checked, mismatches=${bad}`);
