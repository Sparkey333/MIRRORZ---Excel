import { Workbook, CellError } from '../packages/core/src/index.js';
import { Evaluator } from '../packages/formula/src/evaluator.js';
import { LOOKUP_FUNCTIONS } from '../packages/formula/src/functions/lookup.js';
import { parseFormula } from '../packages/formula/src/parser.js';
import { FunctionRegistry } from '../packages/formula/src/registry.js';
import { WorkbookStore } from '../packages/formula/src/store.js';

const registry = new FunctionRegistry().registerAll(LOOKUP_FUNCTIONS);
const wb = new Workbook();
const s = wb.addSheet('S');
const numbers = [1, 2, 4, 8, 16];
const letters = ['a','b','c','d','e'];
for (let i=0;i<5;i++){ s.setValue(i,0,numbers[i]!); s.setValue(i,1,letters[i]!); }
s.setValue(0,2,'apple'); s.setValue(1,2,'banana'); s.setValue(2,2,'cherry');
s.setValue(0,3,30); s.setValue(1,3,20); s.setValue(2,3,10);
s.setValue(0,4,10); s.setValue(1,4,10); s.setValue(2,4,20); s.setValue(3,4,30); s.setValue(4,4,30);
s.setValue(0,5,CellError.DIV0);
s.setValue(0,6,1); s.setValue(0,7,2); s.setValue(0,8,3);
s.setValue(1,6,'x'); s.setValue(1,7,'y'); s.setValue(1,8,'z');
wb.addSheet('Other');
const ev = new Evaluator(new WorkbookStore(wb), registry, { dateSystem: wb.dateSystem });
function shape(f: string, row=20, col=10) {
  return ev.evaluate({ ast: parseFormula(f, { origin: { row, col } }), sheet:'S', row, col });
}
function run(f: string, row=20, col=10) {
  return ev.evaluateScalar({ ast: parseFormula(f, { origin: { row, col } }), sheet:'S', row, col });
}
const cases = process.argv.slice(2);
for (const c of cases) {
  let out: unknown;
  const t0 = Date.now();
  try { out = c.startsWith('#') ? shape(c.slice(1)) : run(c); } catch (e) { out = `THREW ${(e as Error).message}`; }
  console.log(`${c}  =>  ${JSON.stringify(out)}   [${Date.now()-t0}ms]`);
}
