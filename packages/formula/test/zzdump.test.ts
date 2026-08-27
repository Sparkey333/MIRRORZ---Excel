import { readFileSync } from 'node:fs';
import { it } from 'vitest';
import { readXlsx } from '/home/user/MIRRORZ---Excel/packages/formats/src/xlsx/read.js';

it('dump', () => {
  for (const file of ['formulas.calc.xlsx', 'precedence.calc.xlsx']) {
    const { workbook } = readXlsx(new Uint8Array(readFileSync(`/home/user/MIRRORZ---Excel/fixtures/generated/${file}`)));
    console.log('=== FILE', file);
    for (const sheet of workbook.sheets) {
      console.log('--- SHEET', sheet.name);
      const rows = new Map<number, string[]>();
      for (const { row, col, cell } of sheet.entries()) {
        if (!rows.has(row)) rows.set(row, []);
        rows.get(row)![col] = `${cell.formula ? '=' + cell.formula + ' -> ' : ''}${JSON.stringify(cell.value)}`;
      }
      for (const r of [...rows.keys()].sort((a, b) => a - b)) {
        console.log(r + 1, JSON.stringify(rows.get(r)));
      }
    }
  }
});
