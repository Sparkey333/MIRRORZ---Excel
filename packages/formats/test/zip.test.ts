import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { crc32, looksLikeZip, readZip, writeZip } from '../src/zip.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const read = (name: string) => new Uint8Array(readFileSync(new URL(name, FIXTURES)));

describe('crc32', () => {
  it('matches the known IEEE check value', () => {
    // "123456789" -> 0xCBF43926 is the standard CRC-32 test vector.
    expect(crc32(new TextEncoder().encode('123456789'))).toBe(0xcbf4_3926);
  });

  it('is zero for empty input', () => {
    expect(crc32(new Uint8Array(0))).toBe(0);
  });
});

describe('reading real Office archives', () => {
  it.each(['basic-types.xlsx', 'formulas.xlsx', 'styling.xlsx', 'features.xlsx', 'edge-cases.xlsx'])(
    'opens %s and finds the OOXML parts',
    (name) => {
      const entries = readZip(read(name));
      expect(entries.has('[Content_Types].xml')).toBe(true);
      expect(entries.has('xl/workbook.xml')).toBe(true);
      expect(entries.has('_rels/.rels')).toBe(true);
      const wb = entries.get('xl/workbook.xml')!.data();
      expect(new TextDecoder().decode(wb)).toContain('<workbook');
    },
  );

  it('opens an ODS archive', () => {
    const entries = readZip(read('basic-types.ods'));
    expect(entries.has('content.xml')).toBe(true);
    expect(entries.has('mimetype')).toBe(true);
    // ODS stores the mimetype entry uncompressed by specification.
    expect(entries.get('mimetype')!.method).toBe(0);
    expect(new TextDecoder().decode(entries.get('mimetype')!.data())).toBe(
      'application/vnd.oasis.opendocument.spreadsheet',
    );
  });

  it('reads a multi-megabyte archive without truncation', () => {
    const entries = readZip(read('large.xlsx'));
    const sheet = [...entries.keys()].find((k) => k.startsWith('xl/worksheets/sheet'))!;
    const xml = entries.get(sheet)!.data();
    expect(xml.length).toBeGreaterThan(1_000_000);
    expect(new TextDecoder().decode(xml.subarray(xml.length - 40))).toContain('</worksheet>');
  });

  it('verifies stored CRCs across every part of every fixture', () => {
    for (const name of ['basic-types.xlsx', 'features.xlsx', 'large.xlsx']) {
      for (const entry of readZip(read(name)).values()) {
        if (entry.size === 0) continue;
        expect(crc32(entry.data()), `${name}:${entry.name}`).toBe(entry.crc32);
      }
    }
  });

  it('decompresses lazily and caches', () => {
    const entries = readZip(read('basic-types.xlsx'));
    const e = entries.get('xl/workbook.xml')!;
    expect(e.data()).toBe(e.data());
  });
});

describe('writing', () => {
  const fixed = new Date(Date.UTC(2024, 0, 1, 12, 0, 0));

  it('round-trips text and binary entries', () => {
    const payloads = [
      { name: 'a.txt', data: new TextEncoder().encode('hello') },
      { name: 'nested/dir/b.xml', data: new TextEncoder().encode('<x>'.repeat(500)) },
      { name: 'empty.bin', data: new Uint8Array(0) },
      { name: 'binary.bin', data: new Uint8Array(Array.from({ length: 5000 }, (_, i) => i & 0xff)) },
    ];
    const zip = writeZip(payloads, { modified: fixed });
    expect(looksLikeZip(zip)).toBe(true);

    const back = readZip(zip);
    expect([...back.keys()]).toEqual(payloads.map((p) => p.name));
    for (const p of payloads) {
      expect(Array.from(back.get(p.name)!.data())).toEqual(Array.from(p.data));
    }
  });

  it('compresses compressible data and stores incompressible data', () => {
    const compressible = new TextEncoder().encode('a'.repeat(10_000));
    // Deterministic pseudo-random bytes: a linear congruential generator, so the
    // test does not depend on Math.random and stays reproducible.
    const random = new Uint8Array(10_000);
    let seed = 12_345;
    for (let i = 0; i < random.length; i++) {
      seed = (seed * 1_103_515_245 + 12_345) & 0x7fff_ffff;
      random[i] = (seed >> 16) & 0xff;
    }
    const zip = readZip(
      writeZip([
        { name: 'compressible', data: compressible },
        { name: 'random', data: random },
      ]),
    );
    expect(zip.get('compressible')!.compressedSize).toBeLessThan(compressible.length / 10);
    expect(zip.get('random')!.compressedSize).toBeLessThanOrEqual(random.length);
  });

  it('honours the store flag', () => {
    const data = new TextEncoder().encode('x'.repeat(1000));
    const zip = readZip(writeZip([{ name: 'raw', data, store: true }]));
    expect(zip.get('raw')!.method).toBe(0);
    expect(zip.get('raw')!.compressedSize).toBe(1000);
  });

  it('is byte-for-byte reproducible with a fixed timestamp', () => {
    const make = () =>
      writeZip([{ name: 'a', data: new TextEncoder().encode('same') }], { modified: fixed });
    expect(Array.from(make())).toEqual(Array.from(make()));
  });

  it('preserves UTF-8 entry names', () => {
    const name = 'xl/worksheets/ünïcødé-名前.xml';
    const zip = readZip(writeZip([{ name, data: new TextEncoder().encode('v') }]));
    expect(zip.has(name)).toBe(true);
  });

  it('round-trips an entire real workbook through unzip and rezip', () => {
    const original = readZip(read('features.xlsx'));
    const rebuilt = readZip(
      writeZip([...original.values()].map((e) => ({ name: e.name, data: e.data() }))),
    );
    expect([...rebuilt.keys()].sort()).toEqual([...original.keys()].sort());
    for (const [name, entry] of original) {
      expect(Array.from(rebuilt.get(name)!.data()), name).toEqual(Array.from(entry.data()));
    }
  });
});

describe('error handling', () => {
  it('rejects a buffer that is not a ZIP', () => {
    expect(() => readZip(new TextEncoder().encode('this is not a zip file at all'))).toThrow(
      /end of central directory/,
    );
  });

  it('rejects a truncated archive', () => {
    const zip = writeZip([{ name: 'a', data: new TextEncoder().encode('hello world') }]);
    expect(() => readZip(zip.subarray(0, zip.length - 30))).toThrow();
  });

  it('detects corrupted entry data via the CRC', () => {
    // 'abcabcabc...' deflates to a stream long enough that flipping a byte
    // inside it still inflates to the right length - which is exactly the case
    // a size check alone would wave through, and the CRC catches.
    const payload = new TextEncoder().encode('abc'.repeat(400));
    const zip = writeZip([{ name: 'a', data: payload }]);
    expect(readZip(zip).get('a')!.data().length).toBe(payload.length);

    let detected = 0;
    for (let i = 34; i < 60; i++) {
      const damaged = zip.slice();
      damaged[i] = damaged[i]! ^ 0x55;
      try {
        readZip(damaged).get('a')!.data();
      } catch {
        detected++;
      }
    }
    expect(detected).toBeGreaterThan(0);
  });

  it('can skip CRC verification when the caller opts out', () => {
    const zip = writeZip([{ name: 'a', data: new TextEncoder().encode('hello') }]);
    expect(readZip(zip, { verifyCrc: false }).get('a')!.data().length).toBe(5);
  });
});
