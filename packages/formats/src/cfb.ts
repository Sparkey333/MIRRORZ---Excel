/**
 * Compound File Binary (OLE2) container reader, per [MS-CFB].
 *
 * CFB is a small filesystem inside a single file: a header, a sector allocation
 * table, and a directory of named storages and streams. Three things we need
 * sit inside one, so implementing the container once pays for itself three
 * times over: legacy .xls workbooks (a BIFF8 stream called 'Workbook'),
 * vbaProject.bin inside a macro-enabled .xlsm, and the outer wrapper of an
 * encrypted OOXML package.
 *
 * This is a reader only. Writing CFB means allocating sectors and rebalancing a
 * red-black tree, and we have no use for it: we open legacy workbooks and save
 * them as .xlsx, and a vbaProject.bin is copied through byte-for-byte rather
 * than rebuilt.
 *
 * The reading posture is deliberately suspicious. These files arrive by email
 * from strangers, and the format is a graph of 32-bit indices with no
 * checksums, so a corrupt or hostile file can trivially describe a sector chain
 * that loops forever or a stream that claims to be four gigabytes. Every index
 * is range-checked against the real file length, every chain walk is capped by
 * the sector count, and a declared stream size is cross-checked against the
 * bytes its chain can actually supply before anything is allocated.
 */

/** Sector ids above this are reserved markers rather than real locations. */
const MAXREGSECT = 0xffff_fffa;
/** Storage reserved for a DIFAT sector; not part of any chain. */
const DIFSECT = 0xffff_fffc;
/** Storage reserved for a FAT sector; not part of any chain. */
const FATSECT = 0xffff_fffd;
const ENDOFCHAIN = 0xffff_fffe;
const FREESECT = 0xffff_ffff;

/** Directory stream ids above this are reserved; NOSTREAM terminates a link. */
const MAXREGSID = 0xffff_fffa;
const NOSTREAM = 0xffff_ffff;

const OBJ_UNALLOCATED = 0x00;
const OBJ_STORAGE = 0x01;
const OBJ_STREAM = 0x02;
const OBJ_ROOT = 0x05;

const HEADER_SIZE = 512;
const DIR_ENTRY_SIZE = 128;
/** Header DIFAT slots, holding the first 109 FAT sector locations. */
const HEADER_DIFAT_COUNT = 109;
const HEADER_DIFAT_OFFSET = 76;
/** The mini stream is always divided into 64-byte sectors (mini sector shift 6). */
const MINI_SECTOR_SIZE = 64;

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export class CfbError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CfbError';
  }
}

export type CfbEntryType = 'root' | 'storage' | 'stream';

export interface CfbEntry {
  /**
   * Name exactly as stored. Leading control characters are kept: Office names
   * its property-set streams 'SummaryInformation' and its OLE metadata
   * 'CompObj', and stripping them would make those streams unfindable.
   */
  name: string;
  /** Path from the root, e.g. 'VBA/Module1'. Empty for the root entry itself. */
  path: string;
  /** Index of this entry in the directory array. */
  id: number;
  type: CfbEntryType;
  /** Stream length in bytes. Zero for storages; the mini stream length for the root. */
  size: number;
  /** First sector of the stream, in the FAT or the mini FAT depending on size. */
  startSector: number;
  /** Object class GUID, lower-case hyphenated, or undefined when all zeroes. */
  clsid: string | undefined;
  /** User-defined flags on a storage; zero for streams. */
  stateBits: number;
  created: Date | undefined;
  modified: Date | undefined;
  /** Immediate children, sorted by the specification's name ordering. */
  children: CfbEntry[];
  /** Stream contents, read on first access and cached. Empty for a storage. */
  data(): Uint8Array;
}

export interface CfbHeader {
  minorVersion: number;
  majorVersion: number;
  /** Bytes per sector: 512 for version 3, 4096 for version 4. */
  sectorSize: number;
  miniSectorSize: number;
  directorySectorCount: number;
  fatSectorCount: number;
  firstDirectorySector: number;
  transactionSignature: number;
  /** Streams shorter than this live in the mini stream. Always 4096 in practice. */
  miniStreamCutoff: number;
  firstMiniFatSector: number;
  miniFatSectorCount: number;
  firstDifatSector: number;
  difatSectorCount: number;
}

export interface CfbFile {
  header: CfbHeader;
  root: CfbEntry;
  /**
   * Every entry, keyed by full path ('VBA/Module1') and additionally by bare
   * name where that name is unambiguous. Callers looking for a well-known
   * stream such as 'Workbook' should not have to know which storage it sits in.
   */
  entries: Map<string, CfbEntry>;
  read(path: string): Uint8Array;
  has(path: string): boolean;
}

export interface CfbReadOptions {
  /**
   * Reject any stream declaring more than this many bytes. The 64-bit size
   * field is attacker-controlled and a bogus value would otherwise become an
   * allocation. Default 2 GiB, which is the version 3 ceiling anyway.
   */
  maxStreamSize?: number;
}

/** True when the buffer starts with the OLE2 signature. */
export function looksLikeCfb(buf: Uint8Array): boolean {
  if (buf.length < HEADER_SIZE) return false;
  for (let i = 0; i < SIGNATURE.length; i++) {
    if (buf[i] !== SIGNATURE[i]) return false;
  }
  return true;
}

export function readCfb(buf: Uint8Array, options: CfbReadOptions = {}): CfbFile {
  const maxStreamSize = options.maxStreamSize ?? 2 * 1024 * 1024 * 1024;
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const header = parseHeader(buf, view);
  const { sectorSize } = header;

  // Sector 0 begins one sector into the file, so the last addressable sector is
  // bounded by the real file length. Every id we follow is checked against this
  // rather than against anything the file claims about itself.
  const sectorCount = Math.max(0, Math.floor(buf.length / sectorSize) - 1);

  const fat = readFat(view, header, sectorCount);
  const dirSectors = followChain(fat, header.firstDirectorySector, sectorCount, 'directory');
  const directory = gatherSectors(buf, dirSectors, sectorSize, dirSectors.length * sectorSize);
  const raws = parseDirectory(directory, buf.length, maxStreamSize, header.majorVersion);
  if (raws.length === 0) throw new CfbError('compound file has an empty directory');
  const rootRaw = raws[0]!;
  if (rootRaw.type !== OBJ_ROOT) {
    throw new CfbError(`first directory entry has object type ${rootRaw.type}, expected root (5)`);
  }

  // The mini FAT and the mini stream are only built if some stream actually
  // needs them; a workbook whose every stream clears the cutoff should not pay
  // to materialise a mini stream it will never read.
  let miniFat: Uint32Array | undefined;
  const getMiniFat = (): Uint32Array => {
    if (!miniFat) {
      const chain = followChain(fat, header.firstMiniFatSector, sectorCount, 'mini FAT');
      const bytes = gatherSectors(buf, chain, sectorSize, chain.length * sectorSize);
      miniFat = toSectorIds(bytes);
    }
    return miniFat;
  };

  let miniStream: Uint8Array | undefined;
  const getMiniStream = (): Uint8Array => {
    // The root entry's own stream is the mini stream, and it is always chained
    // through the FAT however short it is.
    if (!miniStream) miniStream = readFatStream(buf, fat, rootRaw, sectorSize, sectorCount);
    return miniStream;
  };

  const readData = (raw: RawEntry): Uint8Array => {
    if (raw.type === OBJ_STORAGE) return new Uint8Array(0);
    if (raw.size === 0) return new Uint8Array(0);
    if (raw.type === OBJ_ROOT || raw.size >= header.miniStreamCutoff) {
      return readFatStream(buf, fat, raw, sectorSize, sectorCount);
    }
    return readMiniStream(getMiniStream(), getMiniFat(), raw);
  };

  const root = buildTree(raws, readData);
  const entries = indexEntries(root);

  return {
    header,
    root,
    entries,
    has(path: string): boolean {
      return entries.has(normalisePath(path));
    },
    read(path: string): Uint8Array {
      const entry = entries.get(normalisePath(path));
      if (!entry) throw new CfbError(`no such stream: ${path}`);
      if (entry.type === 'storage') throw new CfbError(`${path} is a storage, not a stream`);
      return entry.data();
    },
  };
}

function parseHeader(buf: Uint8Array, view: DataView): CfbHeader {
  if (!looksLikeCfb(buf)) {
    throw new CfbError('not a compound file: missing the D0 CF 11 E0 A1 B1 1A E1 signature');
  }
  const byteOrder = view.getUint16(28, true);
  if (byteOrder !== 0xfffe) {
    throw new CfbError(
      `unsupported byte order mark 0x${byteOrder.toString(16)}; only little-endian (0xFFFE) exists`,
    );
  }
  const majorVersion = view.getUint16(26, true);
  if (majorVersion !== 3 && majorVersion !== 4) {
    throw new CfbError(`unsupported compound file major version ${majorVersion}`);
  }
  // The sector shift, not the version field, is what actually governs the
  // layout, so it is the field we validate and trust.
  const sectorShift = view.getUint16(30, true);
  if (sectorShift !== 9 && sectorShift !== 12) {
    throw new CfbError(`invalid sector shift ${sectorShift}; expected 9 (512 bytes) or 12 (4096)`);
  }
  const miniSectorShift = view.getUint16(32, true);
  if (miniSectorShift !== 6) {
    throw new CfbError(`invalid mini sector shift ${miniSectorShift}; the format fixes it at 6`);
  }
  const sectorSize = 1 << sectorShift;
  if (buf.length < sectorSize * 2) {
    throw new CfbError(`file is ${buf.length} bytes, too short to hold a header and one sector`);
  }
  // MS-CFB 2.2 fixes this field at 0x00001000, and it is not decorative: it
  // decides, for every stream in the file, whether that stream is read from the
  // FAT or from the mini FAT. A wrong value does not fail loudly, it silently
  // routes reads through the wrong allocator and returns some other stream's
  // bytes, so it has to be checked rather than trusted.
  const miniStreamCutoff = view.getUint32(56, true);
  if (miniStreamCutoff !== 0x1000) {
    throw new CfbError(
      `invalid mini stream cutoff ${miniStreamCutoff}; the format fixes it at 4096`,
    );
  }
  return {
    minorVersion: view.getUint16(24, true),
    majorVersion,
    sectorSize,
    miniSectorSize: 1 << miniSectorShift,
    directorySectorCount: view.getUint32(40, true),
    fatSectorCount: view.getUint32(44, true),
    firstDirectorySector: view.getUint32(48, true),
    transactionSignature: view.getUint32(52, true),
    miniStreamCutoff,
    firstMiniFatSector: view.getUint32(60, true),
    miniFatSectorCount: view.getUint32(64, true),
    firstDifatSector: view.getUint32(68, true),
    difatSectorCount: view.getUint32(72, true),
  };
}

/** Byte offset of a sector. Sector 0 sits immediately after the header sector. */
function sectorOffset(sector: number, sectorSize: number): number {
  return (sector + 1) * sectorSize;
}

/**
 * Assemble the FAT from the sectors the DIFAT names.
 *
 * The header carries the first 109 FAT sector locations inline, which covers
 * every file below about 7 MB. Past that the DIFAT continues into its own chain
 * of sectors, each holding sector-size/4 - 1 locations plus a link to the next.
 * The declared FAT sector count is not used as the loop bound: producers get it
 * wrong often enough, and the DIFAT's own terminators are authoritative.
 */
function readFat(view: DataView, header: CfbHeader, sectorCount: number): Uint32Array {
  const { sectorSize } = header;
  const fatSectors: number[] = [];

  for (let i = 0; i < HEADER_DIFAT_COUNT; i++) {
    const sector = view.getUint32(HEADER_DIFAT_OFFSET + i * 4, true);
    if (sector > MAXREGSECT) continue;
    checkSector(sector, sectorCount, 'FAT');
    fatSectors.push(sector);
  }

  const perDifatSector = sectorSize / 4 - 1;
  let next = header.firstDifatSector;
  let walked = 0;
  while (next !== ENDOFCHAIN && next !== FREESECT) {
    if (next > MAXREGSECT) {
      throw new CfbError(`DIFAT chain contains reserved sector id 0x${next.toString(16)}`);
    }
    checkSector(next, sectorCount, 'DIFAT');
    if (++walked > sectorCount) {
      throw new CfbError('DIFAT chain does not terminate (circular?)');
    }
    const base = sectorOffset(next, sectorSize);
    for (let i = 0; i < perDifatSector; i++) {
      const sector = view.getUint32(base + i * 4, true);
      if (sector > MAXREGSECT) continue;
      checkSector(sector, sectorCount, 'FAT');
      fatSectors.push(sector);
    }
    if (fatSectors.length > sectorCount) {
      throw new CfbError('DIFAT names more FAT sectors than the file contains');
    }
    next = view.getUint32(base + perDifatSector * 4, true);
  }

  // The FAT is indexed by sector number, so entries past the last sector the
  // file actually has can never be reached. Clamping keeps the allocation
  // proportional to the file instead of to the number of FAT sectors the DIFAT
  // claims: 109 header slots all naming one sector would otherwise turn an 8 KB
  // version 4 file into a 446 KB table of repeated garbage.
  const entriesPerSector = sectorSize / 4;
  const fat = new Uint32Array(Math.min(fatSectors.length * entriesPerSector, sectorCount));
  let p = 0;
  for (const sector of fatSectors) {
    if (p >= fat.length) break;
    const base = sectorOffset(sector, sectorSize);
    for (let i = 0; i < entriesPerSector && p < fat.length; i++) {
      fat[p++] = view.getUint32(base + i * 4, true);
    }
  }
  return fat;
}

function checkSector(sector: number, sectorCount: number, what: string): void {
  if (sector >= sectorCount) {
    throw new CfbError(
      `${what} sector ${sector} is past the end of the file (only ${sectorCount} sectors)`,
    );
  }
}

/**
 * Walk a sector chain to its terminator.
 *
 * The cap is the total sector count: a chain cannot legitimately visit more
 * sectors than the file has, so exceeding it means the FAT loops. That is
 * cheaper than a visited set and catches every cycle, since any cycle must
 * revisit within one full pass.
 */
function followChain(
  fat: Uint32Array,
  start: number,
  sectorCount: number,
  what: string,
): number[] {
  const out: number[] = [];
  let sector = start;
  // A well-formed chain ends in ENDOFCHAIN, but writers exist that leave the
  // last link FREESECT, and refusing to open those buys nothing.
  while (sector !== ENDOFCHAIN && sector !== FREESECT) {
    if (sector === FATSECT || sector === DIFSECT || sector > MAXREGSECT) {
      throw new CfbError(`${what} chain contains reserved sector id 0x${sector.toString(16)}`);
    }
    checkSector(sector, sectorCount, what);
    if (sector >= fat.length) {
      throw new CfbError(`${what} sector ${sector} has no FAT entry`);
    }
    out.push(sector);
    if (out.length > sectorCount) {
      throw new CfbError(`${what} sector chain does not terminate (circular FAT?)`);
    }
    sector = fat[sector]!;
  }
  return out;
}

function gatherSectors(
  buf: Uint8Array,
  chain: number[],
  sectorSize: number,
  size: number,
): Uint8Array {
  if (chain.length * sectorSize < size) {
    throw new CfbError(
      `stream declares ${size} bytes but its chain of ${chain.length} sectors supplies only ${chain.length * sectorSize}`,
    );
  }
  const out = new Uint8Array(size);
  let p = 0;
  for (const sector of chain) {
    if (p >= size) break;
    const start = sectorOffset(sector, sectorSize);
    const n = Math.min(sectorSize, size - p);
    if (start + n > buf.length) {
      throw new CfbError(`sector ${sector} extends past the end of the file`);
    }
    out.set(buf.subarray(start, start + n), p);
    p += n;
  }
  return out;
}

function readFatStream(
  buf: Uint8Array,
  fat: Uint32Array,
  raw: RawEntry,
  sectorSize: number,
  sectorCount: number,
): Uint8Array {
  if (raw.size === 0) return new Uint8Array(0);
  const chain = followChain(fat, raw.startSector, sectorCount, `stream ${raw.name}`);
  return gatherSectors(buf, chain, sectorSize, raw.size);
}

function readMiniStream(mini: Uint8Array, miniFat: Uint32Array, raw: RawEntry): Uint8Array {
  const capacity = Math.floor(mini.length / MINI_SECTOR_SIZE);
  // Check what the mini stream can supply before allocating, so a bogus size
  // in the directory cannot become a large Uint8Array we then throw away. The
  // FAT path gets the same treatment inside gatherSectors.
  if (raw.size > capacity * MINI_SECTOR_SIZE) {
    throw new CfbError(
      `mini stream ${raw.name} declares ${raw.size} bytes but the whole mini stream holds only ${capacity * MINI_SECTOR_SIZE}`,
    );
  }
  const out = new Uint8Array(raw.size);
  // A mini chain ends when the declared size is satisfied, not when it runs out
  // of links, so a cycle would not lengthen the walk - it would quietly repeat
  // one 64-byte sector for the whole stream. Detecting it needs a real visited
  // set, which costs one byte per mini sector and is bounded by the mini
  // stream we already hold.
  const visited = new Uint8Array(capacity);
  let sector = raw.startSector;
  let p = 0;
  while (p < raw.size) {
    if (sector === ENDOFCHAIN || sector === FREESECT) {
      throw new CfbError(
        `mini stream ${raw.name} declares ${raw.size} bytes but its chain ends after ${p}`,
      );
    }
    if (sector > MAXREGSECT) {
      throw new CfbError(`mini chain for ${raw.name} contains reserved id 0x${sector.toString(16)}`);
    }
    if (sector >= capacity) {
      throw new CfbError(
        `mini sector ${sector} of ${raw.name} is past the end of the mini stream`,
      );
    }
    if (visited[sector] === 1) {
      throw new CfbError(`mini chain for ${raw.name} does not terminate (circular mini FAT?)`);
    }
    visited[sector] = 1;
    const start = sector * MINI_SECTOR_SIZE;
    const n = Math.min(MINI_SECTOR_SIZE, raw.size - p);
    out.set(mini.subarray(start, start + n), p);
    p += n;
    if (sector >= miniFat.length) {
      throw new CfbError(`mini sector ${sector} of ${raw.name} has no mini FAT entry`);
    }
    sector = miniFat[sector]!;
  }
  return out;
}

function toSectorIds(bytes: Uint8Array): Uint32Array {
  const out = new Uint32Array(Math.floor(bytes.length / 4));
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let i = 0; i < out.length; i++) out[i] = view.getUint32(i * 4, true);
  return out;
}

interface RawEntry {
  name: string;
  type: number;
  leftId: number;
  rightId: number;
  childId: number;
  clsid: string | undefined;
  stateBits: number;
  created: Date | undefined;
  modified: Date | undefined;
  startSector: number;
  size: number;
}

const utf16Decoder = new TextDecoder('utf-16le');

function parseDirectory(
  dir: Uint8Array,
  fileLength: number,
  maxStreamSize: number,
  majorVersion: number,
): RawEntry[] {
  const view = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  const count = Math.floor(dir.length / DIR_ENTRY_SIZE);
  const out: RawEntry[] = [];
  for (let i = 0; i < count; i++) {
    const o = i * DIR_ENTRY_SIZE;
    let type = dir[o + 66]!;
    // Producers do not all zero the slots they are not using, and a stray byte
    // in a slot nothing links to is no reason to refuse a workbook. Anything we
    // do not recognise becomes unallocated, and unallocated entries are never
    // linked into the tree.
    if (type !== OBJ_STORAGE && type !== OBJ_STREAM && type !== OBJ_ROOT) type = OBJ_UNALLOCATED;

    const nameLength = view.getUint16(o + 64, true);
    // The length counts the mandatory terminating null, so an allocated entry
    // cannot declare fewer than two bytes; zero means the name is not there.
    if (nameLength > 64 || nameLength % 2 !== 0 || (nameLength < 2 && type !== OBJ_UNALLOCATED)) {
      if (type === OBJ_UNALLOCATED) {
        out.push(unallocated());
        continue;
      }
      throw new CfbError(
        `directory entry ${i} has an invalid name length of ${nameLength} bytes`,
      );
    }
    // The length counts the UTF-16 terminating null, which we do not want.
    const nameBytes = nameLength >= 2 ? dir.subarray(o, o + nameLength - 2) : dir.subarray(o, o);
    const name = trimAtNull(utf16Decoder.decode(nameBytes));

    // MS-CFB 2.6.1: a storage object's stream size MUST be zero. Reporting the
    // raw field would hand callers a length that data() can never return.
    const size = type === OBJ_STORAGE ? 0 : readStreamSize(view, o + 120, majorVersion);
    if (type !== OBJ_UNALLOCATED && type !== OBJ_STORAGE) {
      if (size > maxStreamSize) {
        throw new CfbError(
          `stream ${name} declares ${size} bytes, above the ${maxStreamSize}-byte limit`,
        );
      }
      if (size > fileLength) {
        throw new CfbError(
          `stream ${name} declares ${size} bytes but the whole file is only ${fileLength}`,
        );
      }
    }

    out.push({
      name,
      type,
      leftId: view.getUint32(o + 68, true),
      rightId: view.getUint32(o + 72, true),
      childId: view.getUint32(o + 76, true),
      clsid: readClsid(view, o + 80),
      stateBits: view.getUint32(o + 96, true),
      created: readFileTime(view, o + 100),
      modified: readFileTime(view, o + 108),
      startSector: view.getUint32(o + 116, true),
      size,
    });
  }
  return out;
}

function unallocated(): RawEntry {
  return {
    name: '',
    type: OBJ_UNALLOCATED,
    leftId: NOSTREAM,
    rightId: NOSTREAM,
    childId: NOSTREAM,
    clsid: undefined,
    stateBits: 0,
    created: undefined,
    modified: undefined,
    startSector: ENDOFCHAIN,
    size: 0,
  };
}

/**
 * Read the 64-bit stream size.
 *
 * A version 3 stream cannot exceed 2 GB, so its top 32 bits MUST be zero; MS-CFB
 * 2.6.1 warns that older writers left them uninitialised anyway and explicitly
 * recommends parsers ignore them - but that licence is granted for version 3
 * only. Dropping the high half unconditionally would silently turn a version 4
 * stream declaring 4 GiB + 100 bytes into a 100-byte read, which is a wrong
 * answer rather than an error. Keep it for version 4 and let the maxStreamSize
 * and file-length guards refuse what we cannot serve.
 */
function readStreamSize(view: DataView, offset: number, majorVersion: number): number {
  const low = view.getUint32(offset, true);
  const high = view.getUint32(offset + 4, true);
  if (high === 0 || majorVersion === 3) return low;
  return high * 0x1_0000_0000 + low;
}

function trimAtNull(name: string): string {
  const nul = name.indexOf('\u0000');
  return nul < 0 ? name : name.slice(0, nul);
}

function readClsid(view: DataView, offset: number): string | undefined {
  let allZero = true;
  for (let i = 0; i < 16; i++) {
    if (view.getUint8(offset + i) !== 0) {
      allZero = false;
      break;
    }
  }
  if (allZero) return undefined;
  const hex = (n: number, width: number) => n.toString(16).padStart(width, '0');
  const tail: string[] = [];
  for (let i = 8; i < 16; i++) tail.push(hex(view.getUint8(offset + i), 2));
  return (
    `${hex(view.getUint32(offset, true), 8)}-${hex(view.getUint16(offset + 4, true), 4)}-` +
    `${hex(view.getUint16(offset + 6, true), 4)}-${tail.slice(0, 2).join('')}-${tail.slice(2).join('')}`
  );
}

/** Windows FILETIME: 100-nanosecond ticks since 1601-01-01 UTC. */
const FILETIME_EPOCH_OFFSET_MS = 11_644_473_600_000n;

function readFileTime(view: DataView, offset: number): Date | undefined {
  const ticks = view.getBigUint64(offset, true);
  if (ticks === 0n) return undefined;
  const ms = ticks / 10_000n - FILETIME_EPOCH_OFFSET_MS;
  const asNumber = Number(ms);
  return Number.isSafeInteger(asNumber) ? new Date(asNumber) : undefined;
}

/**
 * Turn the directory array into a tree.
 *
 * Siblings are stored as a red-black tree, but colour is irrelevant to a
 * reader: left/right are simply an ordinary binary search tree. Rather than
 * walking it in order, children are collected in any order and then sorted by
 * the specification's own comparison - shorter name first, then uppercased
 * UTF-16 code point by code point. That yields the same sequence without a
 * recursive in-order walk, which matters because a degenerate tree from a
 * malformed file would otherwise be a stack overflow.
 */
function buildTree(raws: RawEntry[], readData: (raw: RawEntry) => Uint8Array): CfbEntry {
  const rootRaw = raws[0]!;
  const root = makeEntry(rootRaw, 0, '', readData);
  const seen = new Set<number>([0]);
  const pending: Array<{ parent: CfbEntry; childId: number }> = [
    { parent: root, childId: rootRaw.childId },
  ];

  while (pending.length > 0) {
    const { parent, childId } = pending.pop()!;
    const stack: number[] = [childId];
    while (stack.length > 0) {
      const id = stack.pop()!;
      if (id === NOSTREAM) continue;
      if (id > MAXREGSID) {
        throw new CfbError(`directory link 0x${id.toString(16)} is a reserved stream id`);
      }
      if (id >= raws.length) {
        throw new CfbError(`directory link ${id} is past the end of the ${raws.length}-entry directory`);
      }
      if (seen.has(id)) {
        throw new CfbError(`directory entry ${id} is reachable twice (cyclic directory tree)`);
      }
      seen.add(id);
      const raw = raws[id]!;
      if (raw.type === OBJ_UNALLOCATED) continue;

      const path = parent.path === '' ? raw.name : `${parent.path}/${raw.name}`;
      const entry = makeEntry(raw, id, path, readData);
      parent.children.push(entry);
      stack.push(raw.leftId, raw.rightId);
      if (raw.type === OBJ_STORAGE && raw.childId !== NOSTREAM) {
        pending.push({ parent: entry, childId: raw.childId });
      }
    }
    parent.children.sort((a, b) => compareEntryNames(a.name, b.name));
  }
  return root;
}

function makeEntry(
  raw: RawEntry,
  id: number,
  path: string,
  readData: (raw: RawEntry) => Uint8Array,
): CfbEntry {
  let cached: Uint8Array | undefined;
  return {
    name: raw.name,
    path,
    id,
    type: raw.type === OBJ_ROOT ? 'root' : raw.type === OBJ_STORAGE ? 'storage' : 'stream',
    size: raw.size,
    startSector: raw.startSector,
    clsid: raw.clsid,
    stateBits: raw.stateBits,
    created: raw.created,
    modified: raw.modified,
    children: [],
    data(): Uint8Array {
      if (!cached) cached = readData(raw);
      return cached;
    },
  };
}

/**
 * The sibling ordering MS-CFB defines: length first, then case-insensitive
 * comparison of single UTF-16 code points. JavaScript's toUpperCase on one code
 * unit is a close enough stand-in for the simple case folding the spec names,
 * and this ordering is only used to present children, never to find them.
 */
function compareEntryNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  for (let i = 0; i < a.length; i++) {
    const x = a.charAt(i).toUpperCase().charCodeAt(0);
    const y = b.charAt(i).toUpperCase().charCodeAt(0);
    if (x !== y) return x - y;
  }
  return 0;
}

/**
 * Index every entry by path, then by bare name where the name is still free.
 *
 * Paths are registered first so that a stream whose name collides with another
 * entry's path never shadows it, and the first entry to claim a bare name keeps
 * it, which makes lookups stable regardless of tree shape.
 */
function indexEntries(root: CfbEntry): Map<string, CfbEntry> {
  const all: CfbEntry[] = [];
  const stack: CfbEntry[] = [root];
  while (stack.length > 0) {
    const entry = stack.pop()!;
    all.push(entry);
    for (const child of entry.children) stack.push(child);
  }

  const entries = new Map<string, CfbEntry>();
  for (const entry of all) {
    if (entry.path !== '' && !entries.has(entry.path)) entries.set(entry.path, entry);
  }
  for (const entry of all) {
    if (!entries.has(entry.name)) entries.set(entry.name, entry);
  }
  return entries;
}

function normalisePath(path: string): string {
  return path.startsWith('/') ? path.slice(1) : path;
}

export { ENDOFCHAIN, FREESECT, FATSECT, DIFSECT, MAXREGSECT, NOSTREAM, MAXREGSID };
