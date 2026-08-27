import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { CfbError, looksLikeCfb, readCfb } from '../src/cfb.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const read = (name: string) => new Uint8Array(readFileSync(new URL(name, FIXTURES)));

const XLS = [
  'basic-types.xls',
  'edge-cases.xls',
  'features.xls',
  'formulas.xls',
  'precedence.xls',
  'styling.xls',
];

// ---------------------------------------------------------------------------
// A minimal version 3/4 compound file builder, so the guards can be tested
// against files that no producer would ever emit. It lays sectors out in the
// order they are allocated and does not attempt to balance anything: the reader
// treats the directory as a plain binary tree, which is what this builds.
// ---------------------------------------------------------------------------

const ENDOFCHAIN = 0xffff_fffe;
const FREESECT = 0xffff_ffff;
const FATSECT = 0xffff_fffd;
const NOSTREAM = 0xffff_ffff;

interface BuildStream {
  name: string;
  data: Uint8Array;
  /** Force allocation from the FAT even when the stream is below the cutoff. */
  big?: boolean;
}

interface BuiltFile {
  bytes: Uint8Array;
  view: DataView;
  sectorSize: number;
  /** Byte offset of a data sector, i.e. where sector n's contents begin. */
  offsetOf(sector: number): number;
  /** Byte offset of FAT entry n. */
  fatEntryOffset(n: number): number;
  dirEntryOffset(n: number): number;
}

function buildCfb(streams: BuildStream[], opts: { major?: 3 | 4 } = {}): BuiltFile {
  const major = opts.major ?? 3;
  const sectorShift = major === 3 ? 9 : 12;
  const sectorSize = 1 << sectorShift;
  const cutoff = 4096;

  const sectors: Uint8Array[] = [];
  const fat: number[] = [];
  const allocate = (): number => {
    sectors.push(new Uint8Array(sectorSize));
    fat.push(FREESECT);
    return sectors.length - 1;
  };
  const chainFor = (byteLength: number): number[] => {
    const n = Math.max(1, Math.ceil(byteLength / sectorSize));
    const ids: number[] = [];
    for (let i = 0; i < n; i++) ids.push(allocate());
    for (let i = 0; i < n; i++) fat[ids[i]!] = i === n - 1 ? ENDOFCHAIN : ids[i + 1]!;
    return ids;
  };
  const writeChain = (ids: number[], data: Uint8Array): void => {
    for (let i = 0; i < ids.length; i++) {
      sectors[ids[i]!]!.set(data.subarray(i * sectorSize, (i + 1) * sectorSize), 0);
    }
  };

  // Small streams accumulate into the mini stream; large ones get their own chain.
  const mini: number[] = [];
  const miniBytes: Uint8Array[] = [];
  let miniSectorCount = 0;
  const entries: Array<{ name: string; type: number; start: number; size: number }> = [];

  const big: Array<{ name: string; data: Uint8Array }> = [];
  for (const s of streams) {
    if (!s.big && s.data.length < cutoff && s.data.length > 0) {
      const start = miniSectorCount;
      const count = Math.ceil(s.data.length / 64);
      for (let i = 0; i < count; i++) {
        mini.push(miniSectorCount + i + 1 === miniSectorCount + count ? ENDOFCHAIN : miniSectorCount + i + 1);
      }
      const padded = new Uint8Array(count * 64);
      padded.set(s.data);
      miniBytes.push(padded);
      miniSectorCount += count;
      entries.push({ name: s.name, type: 2, start, size: s.data.length });
    } else {
      big.push({ name: s.name, data: s.data });
      entries.push({ name: s.name, type: 2, start: 0, size: s.data.length });
    }
  }

  for (const b of big) {
    const ids = chainFor(b.data.length);
    writeChain(ids, b.data);
    const e = entries.find((x) => x.name === b.name)!;
    e.start = ids[0]!;
  }

  const miniStream = concat(miniBytes);
  let miniStreamStart = ENDOFCHAIN;
  if (miniStream.length > 0) {
    const ids = chainFor(miniStream.length);
    writeChain(ids, miniStream);
    miniStreamStart = ids[0]!;
  }

  let firstMiniFat = ENDOFCHAIN;
  let miniFatSectors = 0;
  if (mini.length > 0) {
    const miniFatBytes = new Uint8Array(Math.ceil((mini.length * 4) / sectorSize) * sectorSize).fill(0xff);
    const mv = new DataView(miniFatBytes.buffer);
    for (let i = 0; i < mini.length; i++) mv.setUint32(i * 4, mini[i]!, true);
    const ids = chainFor(miniFatBytes.length);
    writeChain(ids, miniFatBytes);
    firstMiniFat = ids[0]!;
    miniFatSectors = ids.length;
  }

  // Directory: entry 0 is the root, and every stream hangs off it as a right-leaning chain.
  const dirCount = entries.length + 1;
  const dirBytes = new Uint8Array(Math.ceil((dirCount * 128) / sectorSize) * sectorSize);
  const dv = new DataView(dirBytes.buffer);
  const writeEntry = (
    index: number,
    name: string,
    type: number,
    left: number,
    right: number,
    child: number,
    start: number,
    size: number,
  ): void => {
    const o = index * 128;
    for (let i = 0; i < name.length; i++) dv.setUint16(o + i * 2, name.charCodeAt(i), true);
    dv.setUint16(o + name.length * 2, 0, true);
    dv.setUint16(o + 64, name.length * 2 + 2, true);
    dv.setUint8(o + 66, type);
    dv.setUint8(o + 67, 1);
    dv.setUint32(o + 68, left, true);
    dv.setUint32(o + 72, right, true);
    dv.setUint32(o + 76, child, true);
    dv.setUint32(o + 116, start, true);
    dv.setUint32(o + 120, size, true);
  };
  writeEntry(0, 'Root Entry', 5, NOSTREAM, NOSTREAM, entries.length > 0 ? 1 : NOSTREAM, miniStreamStart, miniStream.length);
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i]!;
    writeEntry(i + 1, e.name, e.type, NOSTREAM, i + 2 <= entries.length ? i + 2 : NOSTREAM, NOSTREAM, e.start, e.size);
  }
  const dirIds = chainFor(dirBytes.length);
  writeChain(dirIds, dirBytes);

  // The FAT must describe its own sectors, so allocating them grows the table.
  const entriesPerSector = sectorSize / 4;
  let fatSectorCount = 1;
  for (;;) {
    const total = sectors.length + fatSectorCount;
    const needed = Math.ceil(total / entriesPerSector);
    if (needed <= fatSectorCount) break;
    fatSectorCount = needed;
  }
  const fatIds: number[] = [];
  for (let i = 0; i < fatSectorCount; i++) fatIds.push(allocate());
  for (const id of fatIds) fat[id] = FATSECT;

  const fatBytes = new Uint8Array(fatSectorCount * sectorSize).fill(0xff);
  const fv = new DataView(fatBytes.buffer);
  for (let i = 0; i < fat.length; i++) fv.setUint32(i * 4, fat[i]!, true);
  for (let i = 0; i < fatSectorCount; i++) {
    sectors[fatIds[i]!]!.set(fatBytes.subarray(i * sectorSize, (i + 1) * sectorSize), 0);
  }

  const header = new Uint8Array(sectorSize);
  const hv = new DataView(header.buffer);
  header.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  hv.setUint16(24, 0x003e, true);
  hv.setUint16(26, major, true);
  hv.setUint16(28, 0xfffe, true);
  hv.setUint16(30, sectorShift, true);
  hv.setUint16(32, 6, true);
  hv.setUint32(40, major === 4 ? dirIds.length : 0, true);
  hv.setUint32(44, fatSectorCount, true);
  hv.setUint32(48, dirIds[0]!, true);
  hv.setUint32(56, cutoff, true);
  hv.setUint32(60, firstMiniFat, true);
  hv.setUint32(64, miniFatSectors, true);
  hv.setUint32(68, ENDOFCHAIN, true);
  hv.setUint32(72, 0, true);
  for (let i = 0; i < 109; i++) hv.setUint32(76 + i * 4, i < fatSectorCount ? fatIds[i]! : FREESECT, true);

  const bytes = concat([header, ...sectors]);
  return {
    bytes,
    view: new DataView(bytes.buffer),
    sectorSize,
    offsetOf: (sector: number) => (sector + 1) * sectorSize,
    fatEntryOffset: (n: number) => (fatIds[Math.floor(n / entriesPerSector)]! + 1) * sectorSize + (n % entriesPerSector) * 4,
    dirEntryOffset: (n: number) => (dirIds[Math.floor((n * 128) / sectorSize)]! + 1) * sectorSize + ((n * 128) % sectorSize),
  };
}

function concat(parts: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const p of parts) total += p.length;
  const out = new Uint8Array(total);
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

const bytes = (n: number, seed = 0) =>
  new Uint8Array(Array.from({ length: n }, (_, i) => (i * 7 + seed) & 0xff));

describe('looksLikeCfb', () => {
  it('accepts every .xls fixture', () => {
    for (const name of XLS) expect(looksLikeCfb(read(name)), name).toBe(true);
  });

  it('rejects a zip archive', () => {
    expect(looksLikeCfb(read('basic-types.xlsx'))).toBe(false);
  });

  it('rejects a buffer shorter than the header', () => {
    const short = read('basic-types.xls').subarray(0, 511);
    expect(looksLikeCfb(short)).toBe(false);
  });

  it('rejects a buffer whose signature is one byte wrong', () => {
    const buf = read('basic-types.xls').slice();
    buf[7] = 0xe0;
    expect(looksLikeCfb(buf)).toBe(false);
  });

  it('rejects an all-zero buffer', () => {
    expect(looksLikeCfb(new Uint8Array(1024))).toBe(false);
  });
});

describe('reading real .xls containers', () => {
  it.each(XLS)('%s exposes a Workbook stream starting with a BIFF BOF record', (name) => {
    const cfb = readCfb(read(name));
    expect(cfb.has('Workbook')).toBe(true);
    const wb = cfb.read('Workbook');
    const view = new DataView(wb.buffer, wb.byteOffset, wb.byteLength);
    // BOF: record type 0x0809, then the record length, then BIFF version 0x0600.
    expect(view.getUint16(0, true)).toBe(0x0809);
    // A BIFF8 BOF carries a 16-byte body naming version 0x0600 and substream type 5.
    expect(view.getUint16(2, true)).toBe(16);
    expect(view.getUint16(4, true)).toBe(0x0600);
    expect(view.getUint16(6, true)).toBe(0x0005);
  });

  it.each(XLS)('%s has the directory entries LibreOffice writes', (name) => {
    const cfb = readCfb(read(name));
    const names = [...cfb.root.children].map((c) => c.name).sort();
    expect(names).toEqual(
      [
        'Workbook',
        'CompObj',
        'Ole',
        'SummaryInformation',
        'DocumentSummaryInformation',
      ].sort(),
    );
  });

  it('parses the header of a version 3 file', () => {
    const cfb = readCfb(read('basic-types.xls'));
    expect(cfb.header.majorVersion).toBe(3);
    expect(cfb.header.sectorSize).toBe(512);
    expect(cfb.header.miniSectorSize).toBe(64);
    expect(cfb.header.miniStreamCutoff).toBe(4096);
    // Version 3 does not use the directory sector count field.
    expect(cfb.header.directorySectorCount).toBe(0);
    expect(cfb.header.fatSectorCount).toBeGreaterThan(0);
  });

  it('marks the first directory entry as the root', () => {
    const cfb = readCfb(read('basic-types.xls'));
    expect(cfb.root.type).toBe('root');
    expect(cfb.root.name).toBe('Root Entry');
    expect(cfb.root.path).toBe('');
    expect(cfb.root.id).toBe(0);
    // The root's own stream is the mini stream, so its size is a multiple of 64.
    expect(cfb.root.size % 64).toBe(0);
  });

  it('reads a stream that lives in the mini stream', () => {
    // Ole is 20 bytes, far below the 4096-byte cutoff.
    const cfb = readCfb(read('basic-types.xls'));
    const ole = cfb.entries.get('Ole')!;
    expect(ole.size).toBe(20);
    expect(ole.data()).toHaveLength(20);
    // The OLE stream opens with its version field, 0x02000001.
    expect(new DataView(ole.data().buffer).getUint32(0, true)).toBe(0x0200_0001);
  });

  it('reads a stream that lives in the FAT', () => {
    // formulas.xls has a Workbook well above the cutoff.
    const cfb = readCfb(read('formulas.xls'));
    const wb = cfb.entries.get('Workbook')!;
    expect(wb.size).toBeGreaterThan(4096);
    expect(wb.data()).toHaveLength(wb.size);
  });

  it('handles a Workbook below the cutoff, which basic-types.xls has', () => {
    const cfb = readCfb(read('basic-types.xls'));
    const wb = cfb.entries.get('Workbook')!;
    expect(wb.size).toBeLessThan(4096);
    expect(wb.data()).toHaveLength(wb.size);
    expect(new DataView(wb.data().buffer).getUint16(0, true)).toBe(0x0809);
  });

  it('reads the property-set streams', () => {
    const cfb = readCfb(read('features.xls'));
    const summary = cfb.read('SummaryInformation');
    expect(summary.length).toBe(200);
    // A property set stream starts with byte order 0xFFFE, same mark as the header.
    expect(new DataView(summary.buffer, summary.byteOffset).getUint16(0, true)).toBe(0xfffe);
  });

  it('caches stream data rather than re-reading it', () => {
    const cfb = readCfb(read('styling.xls'));
    const wb = cfb.entries.get('Workbook')!;
    expect(wb.data()).toBe(wb.data());
  });

  it('records the entry size, id and starting sector', () => {
    const cfb = readCfb(read('styling.xls'));
    for (const child of cfb.root.children) {
      expect(child.id).toBeGreaterThan(0);
      expect(child.type).toBe('stream');
      expect(child.data().length).toBe(child.size);
    }
  });

  it('leaves the CLSID undefined when the entry has none', () => {
    const cfb = readCfb(read('basic-types.xls'));
    expect(cfb.entries.get('Workbook')!.clsid).toBeUndefined();
  });

  it('decodes the root storage CLSID', () => {
    // Every .xls carries the Excel worksheet class GUID on its root storage.
    for (const name of XLS) {
      expect(readCfb(read(name)).root.clsid, name).toBe('00020810-0000-0000-c000-000000000046');
    }
  });

  it('decodes FILETIME stamps into Dates in a plausible range', () => {
    for (const name of XLS) {
      for (const entry of readCfb(read(name)).entries.values()) {
        for (const t of [entry.created, entry.modified]) {
          if (t === undefined) continue;
          expect(t.getUTCFullYear(), `${name}:${entry.name}`).toBeGreaterThan(1980);
          expect(t.getUTCFullYear()).toBeLessThan(2200);
        }
      }
    }
  });

  it('sorts children by the specification ordering: length, then case-insensitive', () => {
    const cfb = readCfb(read('basic-types.xls'));
    const names = cfb.root.children.map((c) => c.name);
    for (let i = 1; i < names.length; i++) {
      expect(names[i - 1]!.length).toBeLessThanOrEqual(names[i]!.length);
    }
  });
});

describe('lookups', () => {
  it('has() accepts a leading slash', () => {
    const cfb = readCfb(read('basic-types.xls'));
    expect(cfb.has('/Workbook')).toBe(true);
    expect(cfb.read('/Workbook')).toEqual(cfb.read('Workbook'));
  });

  it('has() is false for an unknown name', () => {
    expect(readCfb(read('basic-types.xls')).has('NoSuchStream')).toBe(false);
  });

  it('read() throws for an unknown name', () => {
    const cfb = readCfb(read('basic-types.xls'));
    expect(() => cfb.read('NoSuchStream')).toThrow(CfbError);
    expect(() => cfb.read('NoSuchStream')).toThrow(/no such stream/);
  });

  it('read() refuses a storage', () => {
    const cfb = readCfb(buildNested().bytes);
    expect(() => cfb.read('VBA')).toThrow(/is a storage, not a stream/);
  });
});

// A container with a nested storage, which no .xls fixture has: the shape a
// vbaProject.bin takes.
function buildNested(): BuiltFile {
  const built = buildCfb([
    { name: 'PROJECT', data: new TextEncoder().encode('ID="{deadbeef}"') },
    { name: 'Module1', data: new TextEncoder().encode('Sub Main()\r\nEnd Sub\r\n') },
    { name: 'dir', data: bytes(300, 3) },
  ]);
  // Re-parent Module1 and dir under a new storage called VBA, by rewriting the
  // directory in place: entry 1 becomes the storage, its child chain the rest.
  const { view, dirEntryOffset } = built;
  const setName = (index: number, name: string) => {
    const o = dirEntryOffset(index);
    for (let i = 0; i < 32; i++) view.setUint16(o + i * 2, 0, true);
    for (let i = 0; i < name.length; i++) view.setUint16(o + i * 2, name.charCodeAt(i), true);
    view.setUint16(o + 64, name.length * 2 + 2, true);
  };
  // Entry 1: turn PROJECT into the VBA storage holding entries 2 and 3.
  setName(1, 'VBA');
  view.setUint8(dirEntryOffset(1) + 66, 1);
  view.setUint32(dirEntryOffset(1) + 72, NOSTREAM, true); // no right sibling
  view.setUint32(dirEntryOffset(1) + 76, 2, true); // child = Module1
  view.setUint32(dirEntryOffset(1) + 116, 0, true);
  view.setUint32(dirEntryOffset(1) + 120, 0, true);
  return built;
}

describe('synthetic containers', () => {
  it('round-trips small streams through the mini FAT', () => {
    const payloads = [
      { name: 'tiny', data: bytes(1) },
      { name: 'exact', data: bytes(64, 1) },
      { name: 'spans', data: bytes(200, 2) },
      { name: 'nearCutoff', data: bytes(4095, 3) },
    ];
    const cfb = readCfb(buildCfb(payloads).bytes);
    for (const p of payloads) {
      expect(cfb.read(p.name), p.name).toEqual(p.data);
    }
    // Everything below the cutoff lives in the mini stream, so the root has one.
    expect(cfb.root.size).toBeGreaterThan(0);
  });

  it('round-trips large streams through the FAT', () => {
    const payloads = [
      { name: 'atCutoff', data: bytes(4096, 5) },
      { name: 'multiSector', data: bytes(5000, 6) },
      { name: 'many', data: bytes(40_000, 7) },
    ];
    const cfb = readCfb(buildCfb(payloads).bytes);
    for (const p of payloads) expect(cfb.read(p.name), p.name).toEqual(p.data);
  });

  it('handles a mix of mini and FAT streams in one file', () => {
    const payloads = [
      { name: 'small', data: bytes(30, 1) },
      { name: 'large', data: bytes(9000, 2) },
      { name: 'small2', data: bytes(1000, 3) },
    ];
    const cfb = readCfb(buildCfb(payloads).bytes);
    for (const p of payloads) expect(cfb.read(p.name), p.name).toEqual(p.data);
  });

  it('reads a zero-length stream as empty', () => {
    const cfb = readCfb(buildCfb([{ name: 'empty', data: new Uint8Array(0) }]).bytes);
    expect(cfb.read('empty')).toEqual(new Uint8Array(0));
  });

  it('reads a version 4 container with 4096-byte sectors', () => {
    const payloads = [
      { name: 'small', data: bytes(100, 1) },
      { name: 'large', data: bytes(20_000, 2) },
    ];
    const built = buildCfb(payloads, { major: 4 });
    const cfb = readCfb(built.bytes);
    expect(cfb.header.majorVersion).toBe(4);
    expect(cfb.header.sectorSize).toBe(4096);
    for (const p of payloads) expect(cfb.read(p.name), p.name).toEqual(p.data);
  });

  it('walks into nested storages and keys entries by path', () => {
    const cfb = readCfb(buildNested().bytes);
    expect(cfb.has('VBA')).toBe(true);
    expect(cfb.entries.get('VBA')!.type).toBe('storage');
    expect(cfb.has('VBA/Module1')).toBe(true);
    expect(cfb.has('VBA/dir')).toBe(true);
    expect(new TextDecoder().decode(cfb.read('VBA/Module1'))).toContain('Sub Main()');
    // The bare name is registered as well, so callers need not know the storage.
    expect(cfb.read('Module1')).toEqual(cfb.read('VBA/Module1'));
  });

  it('reports a storage as zero-length with no data', () => {
    const cfb = readCfb(buildNested().bytes);
    const vba = cfb.entries.get('VBA')!;
    expect(vba.size).toBe(0);
    expect(vba.data()).toEqual(new Uint8Array(0));
    expect(vba.children.map((c) => c.name).sort()).toEqual(['Module1', 'dir']);
  });

  it('gives nested entries a path relative to the root', () => {
    const cfb = readCfb(buildNested().bytes);
    expect(cfb.entries.get('VBA/Module1')!.path).toBe('VBA/Module1');
    expect(cfb.entries.get('VBA')!.path).toBe('VBA');
  });

  it('spans more than one directory sector', () => {
    // Ten streams need three 512-byte directory sectors at four entries each.
    const payloads = Array.from({ length: 10 }, (_, i) => ({
      name: `stream${i}`,
      data: bytes(50 + i, i),
    }));
    const cfb = readCfb(buildCfb(payloads).bytes);
    expect(cfb.root.children).toHaveLength(10);
    for (const p of payloads) expect(cfb.read(p.name), p.name).toEqual(p.data);
  });

  it('spans more than one FAT sector', () => {
    // A 512-byte sector holds 128 FAT entries, so 200 KB needs several.
    const cfb = readCfb(buildCfb([{ name: 'huge', data: bytes(200_000, 9) }]).bytes);
    expect(cfb.header.fatSectorCount).toBeGreaterThan(1);
    expect(cfb.read('huge')).toEqual(bytes(200_000, 9));
  });

  it('keeps names containing control characters intact', () => {
    const cfb = readCfb(buildCfb([{ name: 'SummaryInformation', data: bytes(10) }]).bytes);
    expect(cfb.has('SummaryInformation')).toBe(true);
    expect(cfb.root.children[0]!.name.charCodeAt(0)).toBe(5);
  });

  it('accepts a chain terminated with FREESECT rather than ENDOFCHAIN', () => {
    const built = buildCfb([{ name: 'big', data: bytes(1500) }]);
    const cfb = readCfb(built.bytes);
    const start = cfb.entries.get('big')!.startSector;
    // Walk to the last sector of the chain and slacken its terminator.
    let s = start;
    for (;;) {
      const next = built.view.getUint32(built.fatEntryOffset(s), true);
      if (next === ENDOFCHAIN) break;
      s = next;
    }
    built.view.setUint32(built.fatEntryOffset(s), FREESECT, true);
    expect(readCfb(built.bytes).read('big')).toEqual(bytes(1500));
  });

  it('ignores unallocated directory slots', () => {
    const built = buildCfb([{ name: 'a', data: bytes(20) }, { name: 'b', data: bytes(20, 1) }]);
    // Slot 3 exists but nothing links to it; fill it with junk.
    const o = built.dirEntryOffset(3);
    for (let i = 0; i < 128; i++) built.view.setUint8(o + i, 0xab);
    const cfb = readCfb(built.bytes);
    expect(cfb.root.children.map((c) => c.name)).toEqual(['a', 'b']);
  });
});

describe('rejecting malformed input', () => {
  const corrupt = (mutate: (b: BuiltFile) => void): Uint8Array => {
    const built = buildCfb([
      { name: 'small', data: bytes(100) },
      { name: 'large', data: bytes(6000, 1) },
    ]);
    mutate(built);
    return built.bytes;
  };

  it('rejects a buffer that is not a compound file', () => {
    expect(() => readCfb(read('basic-types.xlsx'))).toThrow(CfbError);
    expect(() => readCfb(read('basic-types.xlsx'))).toThrow(/signature/);
  });

  it('rejects a buffer shorter than a header', () => {
    expect(() => readCfb(new Uint8Array(100))).toThrow(/signature/);
  });

  it('rejects a bad byte order mark', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint16(28, 0xfeff, true)))).toThrow(
      /byte order mark/,
    );
  });

  it('rejects an unsupported major version', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint16(26, 5, true)))).toThrow(
      /major version 5/,
    );
  });

  it('rejects an invalid sector shift', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint16(30, 10, true)))).toThrow(
      /invalid sector shift 10/,
    );
  });

  it('rejects an invalid mini sector shift', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint16(32, 7, true)))).toThrow(
      /mini sector shift 7/,
    );
  });

  it('rejects a file too short to hold a sector after the header', () => {
    const built = buildCfb([{ name: 'a', data: bytes(10) }]);
    expect(() => readCfb(built.bytes.subarray(0, 512))).toThrow(/too short/);
  });

  it('rejects a FAT sector location past the end of the file', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint32(76, 9999, true)))).toThrow(
      /FAT sector 9999 is past the end/,
    );
  });

  it('rejects a first directory sector past the end of the file', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint32(48, 5000, true)))).toThrow(
      /directory sector 5000 is past the end/,
    );
  });

  it('rejects a circular FAT chain for the directory', () => {
    expect(() =>
      readCfb(
        corrupt((b) => {
          const first = b.view.getUint32(48, true);
          b.view.setUint32(b.fatEntryOffset(first), first, true);
        }),
      ),
    ).toThrow(/does not terminate/);
  });

  it('rejects a circular FAT chain for a stream', () => {
    const built = buildCfb([{ name: 'large', data: bytes(6000, 1) }]);
    const cfb = readCfb(built.bytes);
    const start = cfb.entries.get('large')!.startSector;
    built.view.setUint32(built.fatEntryOffset(start), start, true);
    expect(() => readCfb(built.bytes).read('large')).toThrow(/circular FAT/);
  });

  it('rejects a reserved sector id inside a chain', () => {
    const built = buildCfb([{ name: 'large', data: bytes(6000, 1) }]);
    const cfb = readCfb(built.bytes);
    const start = cfb.entries.get('large')!.startSector;
    built.view.setUint32(built.fatEntryOffset(start), FATSECT, true);
    expect(() => readCfb(built.bytes).read('large')).toThrow(/reserved sector id/);
  });

  it('rejects a circular DIFAT chain', () => {
    expect(() =>
      readCfb(
        corrupt((b) => {
          // Turn sector 0 into a DIFAT sector holding no FAT locations whose
          // "next" link points back at itself.
          for (let i = 0; i < b.sectorSize; i++) b.view.setUint8(b.offsetOf(0) + i, 0xff);
          b.view.setUint32(b.offsetOf(0) + b.sectorSize - 4, 0, true);
          b.view.setUint32(68, 0, true);
        }),
      ),
    ).toThrow(/DIFAT chain does not terminate/);
  });

  it('rejects a DIFAT sector past the end of the file', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint32(68, 100_000, true)))).toThrow(
      /DIFAT sector 100000 is past the end/,
    );
  });

  it('rejects a stream declaring more bytes than the file holds', () => {
    expect(() =>
      readCfb(corrupt((b) => b.view.setUint32(b.dirEntryOffset(2) + 120, 0x7fff_0000, true))),
    ).toThrow(/but the whole file is only/);
  });

  it('rejects a stream above the configured size limit', () => {
    const built = buildCfb([{ name: 'large', data: bytes(6000, 1) }]);
    expect(() => readCfb(built.bytes, { maxStreamSize: 1000 })).toThrow(/above the 1000-byte limit/);
  });

  it('rejects a stream whose chain supplies fewer bytes than it declares', () => {
    const built = buildCfb([
      { name: 'large', data: bytes(6000, 1) },
      { name: 'pad', data: bytes(40_000, 2) },
    ]);
    // A size that fits inside the file, but not inside this stream's own chain.
    built.view.setUint32(built.dirEntryOffset(1) + 120, 20_000, true);
    expect(() => readCfb(built.bytes).read('large')).toThrow(/supplies only/);
  });

  // The padding stream only exists to make the file comfortably larger than the
  // bogus sizes below, so that the mini-stream guards are what fires and not the
  // cheaper "bigger than the whole file" check.
  const miniCase = () =>
    buildCfb([
      { name: 'small', data: bytes(500) },
      { name: 'pad', data: bytes(20_000, 4) },
    ]);

  it('rejects a mini stream chain that ends early', () => {
    const built = miniCase();
    built.view.setUint32(built.dirEntryOffset(1) + 120, 4000, true);
    expect(() => readCfb(built.bytes).read('small')).toThrow(/its chain ends after/);
  });

  it('rejects a mini sector past the end of the mini stream', () => {
    const built = miniCase();
    built.view.setUint32(built.dirEntryOffset(1) + 116, 400, true);
    expect(() => readCfb(built.bytes).read('small')).toThrow(/past the end of the mini stream/);
  });

  it('rejects a circular mini FAT chain', () => {
    const built = miniCase();
    const miniFatSector = readCfb(built.bytes).header.firstMiniFatSector;
    // Every mini sector points at itself, so the chain never advances past the
    // declared length.
    for (let i = 0; i < 8; i++) built.view.setUint32(built.offsetOf(miniFatSector) + i * 4, i, true);
    built.view.setUint32(built.dirEntryOffset(1) + 120, 4000, true);
    expect(() => readCfb(built.bytes).read('small')).toThrow(/circular mini FAT/);
  });

  it('rejects a first directory entry that is not the root', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint8(b.dirEntryOffset(0) + 66, 2)))).toThrow(
      /expected root/,
    );
  });

  it('rejects a directory link past the end of the directory', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint32(b.dirEntryOffset(1) + 72, 900, true)))).toThrow(
      /past the end of the .* directory/,
    );
  });

  it('rejects a reserved stream id in a directory link', () => {
    expect(() =>
      readCfb(corrupt((b) => b.view.setUint32(b.dirEntryOffset(1) + 68, 0xffff_fffb, true))),
    ).toThrow(/reserved stream id/);
  });

  it('rejects a cyclic directory tree', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint32(b.dirEntryOffset(2) + 68, 1, true)))).toThrow(
      /reachable twice/,
    );
  });

  it('rejects a self-referential directory entry', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint32(b.dirEntryOffset(1) + 68, 1, true)))).toThrow(
      /reachable twice/,
    );
  });

  it('rejects an odd name length on a live entry', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint16(b.dirEntryOffset(1) + 64, 9, true)))).toThrow(
      /invalid name length of 9/,
    );
  });

  it('rejects a name length beyond the 64-byte field', () => {
    expect(() => readCfb(corrupt((b) => b.view.setUint16(b.dirEntryOffset(1) + 64, 200, true)))).toThrow(
      /invalid name length of 200/,
    );
  });

  it('rejects a truncated file whose sectors run out', () => {
    const built = buildCfb([{ name: 'large', data: bytes(6000, 1) }]);
    expect(() => readCfb(built.bytes.subarray(0, built.bytes.length - 2048))).toThrow(CfbError);
  });

  it('throws CfbError, not a generic Error, for every malformation', () => {
    const cases: Array<(b: BuiltFile) => void> = [
      (b) => b.view.setUint16(28, 0, true),
      (b) => b.view.setUint16(30, 3, true),
      (b) => b.view.setUint32(48, 0xffff_0000, true),
      (b) => b.view.setUint8(b.dirEntryOffset(0) + 66, 0),
    ];
    for (const mutate of cases) {
      let thrown: unknown;
      try {
        readCfb(corrupt(mutate));
      } catch (e) {
        thrown = e;
      }
      expect(thrown).toBeInstanceOf(CfbError);
    }
  });
});
