import { describe, expect, it } from 'vitest';
import { compress, decompress } from '../src/compression.js';

function prng(seed: number) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1_664_525 + 1_013_904_223) >>> 0;
    return (s >>> 16) & 0xffff;
  };
}
// Alphabet of 64 symbols: incompressible enough that a 4kB tail overflows a
// chunk, repetitive enough that three-byte matches are common.
function alpha(length: number, seed: number, symbols: number): Uint8Array {
  const r = prng(seed);
  const out = new Uint8Array(length);
  for (let i = 0; i < length; i++) out[i] = r() % symbols;
  return out;
}

describe('probe', () => {
  it('finds the overflow witness', () => {
    const hits: string[] = [];
    const bad: string[] = [];
    outer: for (const symbols of [48, 64, 80]) {
      for (let seed = 1; seed <= 4; seed++) {
        const base = alpha(4095, seed, symbols);
        for (let len = 3700; len <= 4095; len++) {
          (globalThis as any).__tokenOverflow = false;
          const input = base.subarray(0, len);
          const round = decompress(compress(input));
          if ((globalThis as any).__tokenOverflow) hits.push(`sym ${symbols} seed ${seed} len ${len}`);
          if (round.length !== len) {
            bad.push(`sym ${symbols} seed ${seed} len ${len} -> ${round.length}`);
            break outer;
          }
        }
      }
    }
    console.log('HITS', hits.slice(0, 10), hits.length, 'BAD', bad);
    expect(true).toBe(true);
  }, 600_000);
});
