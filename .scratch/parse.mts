import { parseFormula } from '../packages/formula/src/parser.js';
for (const t of ['1:1','A:A','$1:$3','A1:A3','Sheet1!1:1','R1C1','A1']) {
  try { console.log(t, '=>', JSON.stringify(parseFormula(t))); } catch(e){ console.log(t,'THREW',(e as Error).message); }
}
