import { describe, it } from 'vitest';
import { Evaluator, type SheetStore } from '../src/evaluator.js';
import { DATETIME_FUNCTIONS } from '../src/functions/datetime.js';
import { parseFormula } from '../src/parser.js';
import { FunctionRegistry } from '../src/registry.js';
import type { Value } from '../src/value.js';

const store: SheetStore = {
  getScalar: () => null,
  *iterate() {},
  hasSheet: () => true,
  sheetNames: () => ['Formulas'],
  getDefinedName: () => undefined,
  usedBounds: () => null,
};
const reg = new FunctionRegistry().registerAll(DATETIME_FUNCTIONS).registerAll([
  { name: 'TRUE', params: [], impl: () => true },
  { name: 'FALSE', params: [], impl: () => false },
]);
const ev = new Evaluator(store, reg, { dateSystem: 1900, now: 45352.75 });
function calc(f: string): Value {
  return ev.evaluate({ ast: parseFormula(f, { origin: { row: 0, col: 0 } }), sheet: 'Formulas', row: 0, col: 0 });
}
const PROBES = (process.env.PROBES ?? '').split('\n').filter(Boolean);
describe('probe', () => {
  it('runs', () => {
    for (const f of PROBES) {
      let out: unknown;
      try { out = calc(f); } catch (e) { out = `THROW ${(e as Error).message}`; }
      console.log(f, '=>', JSON.stringify(out));
    }
  });
});
