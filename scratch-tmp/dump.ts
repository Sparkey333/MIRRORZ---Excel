import { readFileSync } from 'node:fs';
import { readXlsx } from '/home/user/MIRRORZ---Excel/packages/formats/src/xlsx/read.js';
const { workbook } = readXlsx(new Uint8Array(readFileSync('fixtures/generated/precedence.calc.xlsx')));
for (const name of workbook.sheets.map(s=>s.name)) {
  const sh = workbook.getSheet(name)!;
  console.log('=== SHEET', name);
  const rows = new Map<number, any[]>();
  for (const { row, col, cell } of sh.entries()) {
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row)![col] = cell.formula ? `{${cell.formula}} => ${JSON.stringify(cell.value)}` : cell.value;
  }
  for (const r of [...rows.keys()].sort((a,b)=>a-b)) {
    console.log(r, JSON.stringify(rows.get(r)));
  }
}
