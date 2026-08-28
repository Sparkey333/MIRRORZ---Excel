import { readFileSync } from 'node:fs';
import { readXlsx } from '../packages/formats/src/xlsx/read.js';
const { workbook } = readXlsx(new Uint8Array(readFileSync('fixtures/generated/formulas.calc.xlsx')));
const sh = workbook.getSheet('Formulas')!;
for (const { row, col, cell } of sh.entries()) {
  if (col !== 2 || !cell.formula) continue;
  const name = sh.getValue(row, 0);
  console.log(JSON.stringify({ name, row, f: cell.formula, v: cell.value }));
}
