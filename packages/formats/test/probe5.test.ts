import { describe, expect, it } from 'vitest';
import { inferValue, formatScalar } from '../src/csv.js';
describe('anchors', () => {
  it('microsoft-documented serials', () => {
    console.log('2008-01-01 ->', inferValue('2008-01-01').value, '(docs: 39448)');
    console.log('1900-01-01 ->', inferValue('1900-01-01').value, '(docs: 1)');
    console.log('9999-12-31 ->', inferValue('9999-12-31').value, '(max)');
    console.log('1904 offset ->', (inferValue('2019-03-01').value as number) - (inferValue('2019-03-01', { dateSystem: 1904 }).value as number));
  });
  it('writer round trip through the reader', () => {
    const vals = [0.1 + 0.2, 1 / 3, 2 / 3, 1e300, 1e-300, 43525, 9007199254740993, 123456789012345678, 1e21, -0.0000001, 1e18];
    for (const v of vals) {
      const text = formatScalar(v);
      const back = inferValue(text).value;
      console.log(String(v).padEnd(24), '->', text.padEnd(24), '-> ', JSON.stringify(back), back === v ? 'SAME' : typeof back === 'number' ? 'DIFFERENT NUMBER' : 'NOT A NUMBER');
    }
  });
});
