/**
 * ZIP container reader and writer.
 *
 * Every modern Office format is a ZIP of XML parts, so this is the floor the
 * whole xlsx path stands on. It is written from scratch rather than pulled from
 * npm for three reasons: no supply-chain surface on the one code path that
 * touches every file a user opens, exact control over how unknown entries are
 * preserved byte-for-byte, and no licence entanglement in a product we intend
 * to sell.
 *
 * Supports store (method 0) and deflate (method 8), which is everything Excel,
 * LibreOffice, and Google Sheets emit, plus ZIP64 for archives past 4 GiB or
 * 65535 entries.
 */

import { deflateRawSync, inflateRawSync } from 'node:zlib';

const SIG_LOCAL = 0x0403_4b50;
const SIG_CENTRAL = 0x0201_4b50;
const SIG_EOCD = 0x0605_4b50;
const SIG_ZIP64_EOCD = 0x0606_4b50;
const SIG_ZIP64_LOCATOR = 0x0706_4b50;
const SIG_DATA_DESCRIPTOR = 0x0807_4b50;

const METHOD_STORE = 0;
const METHOD_DEFLATE = 8;

export interface ZipEntry {
  name: string;
  /** Uncompressed bytes. Decompressed lazily on first access. */
  data(): Uint8Array;
  size: number;
  compressedSize: number;
  method: number;
  crc32: number;
  /** DOS timestamp, decoded to a Date. */
  modified: Date;
  comment?: string;
}

export class ZipError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ZipError';
  }
}

/** Precomputed CRC-32 table (IEEE polynomial, reflected). */
const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb8_8320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c;
  }
  return table;
})();

export function crc32(data: Uint8Array): number {
  let c = -1;
  for (let i = 0; i < data.length; i++) {
    c = CRC_TABLE[(c ^ data[i]!) & 0xff]! ^ (c >>> 8);
  }
  return (c ^ -1) >>> 0;
}

/**
 * Read a ZIP archive from a buffer.
 *
 * Parsing goes through the central directory rather than by scanning local
 * headers forward. The central directory is authoritative: local headers may
 * carry zeroed sizes when a streaming writer used a data descriptor, and
 * scanning forward through those is guesswork.
 */
export interface ZipReadOptions {
  /**
   * Reject any single entry that inflates beyond this many bytes.
   *
   * A desktop app opens files that arrive by email, so a zip bomb - a few
   * kilobytes that inflate to gigabytes - is a real denial-of-service surface,
   * not a theoretical one. Default is 2 GiB, comfortably above any genuine
   * worksheet part and far below what would exhaust memory.
   */
  maxEntrySize?: number;
  /**
   * Reject entries whose compression ratio exceeds this. Legitimate XML
   * compresses roughly 10-20x; 1000x is not a spreadsheet.
   */
  maxCompressionRatio?: number;
  /**
   * Verify each entry's CRC-32 when its data is decompressed. On by default:
   * returning silently-corrupted cell data is far worse than the few
   * milliseconds a table-driven CRC costs, and a damaged workbook is exactly
   * the case where a clear error beats mangled output.
   */
  verifyCrc?: boolean;
}

export function readZip(buf: Uint8Array, options: ZipReadOptions = {}): Map<string, ZipEntry> {
  const verifyCrc = options.verifyCrc ?? true;
  const maxEntrySize = options.maxEntrySize ?? 2 * 1024 * 1024 * 1024;
  const maxRatio = options.maxCompressionRatio ?? 1000;
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const eocdOffset = findEocd(buf);
  if (eocdOffset < 0) throw new ZipError('not a ZIP archive: end of central directory not found');

  let entryCount = view.getUint16(eocdOffset + 10, true);
  let centralSize = view.getUint32(eocdOffset + 12, true);
  let centralOffset = view.getUint32(eocdOffset + 16, true);

  // ZIP64: the 32-bit fields saturate and the real values live in the ZIP64 record.
  if (entryCount === 0xffff || centralOffset === 0xffff_ffff || centralSize === 0xffff_ffff) {
    const locatorOffset = eocdOffset - 20;
    if (locatorOffset >= 0 && view.getUint32(locatorOffset, true) === SIG_ZIP64_LOCATOR) {
      const z64Offset = Number(view.getBigUint64(locatorOffset + 8, true));
      if (view.getUint32(z64Offset, true) !== SIG_ZIP64_EOCD) {
        throw new ZipError('ZIP64 locator does not point at a ZIP64 end of central directory');
      }
      entryCount = Number(view.getBigUint64(z64Offset + 32, true));
      centralSize = Number(view.getBigUint64(z64Offset + 40, true));
      centralOffset = Number(view.getBigUint64(z64Offset + 48, true));
    }
  }

  if (centralOffset + centralSize > buf.length) {
    throw new ZipError('central directory extends past end of file (truncated archive?)');
  }

  const entries = new Map<string, ZipEntry>();
  let p = centralOffset;
  for (let i = 0; i < entryCount; i++) {
    if (p + 46 > buf.length) throw new ZipError(`central directory entry ${i} is truncated`);
    if (view.getUint32(p, true) !== SIG_CENTRAL) {
      throw new ZipError(`bad central directory signature at offset ${p}`);
    }
    const flags = view.getUint16(p + 8, true);
    const method = view.getUint16(p + 10, true);
    const dosTime = view.getUint16(p + 12, true);
    const dosDate = view.getUint16(p + 14, true);
    const crc = view.getUint32(p + 16, true);
    let compressedSize = view.getUint32(p + 20, true);
    let uncompressedSize = view.getUint32(p + 24, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    let localOffset = view.getUint32(p + 42, true);

    const nameBytes = buf.subarray(p + 46, p + 46 + nameLen);
    // Bit 11 declares UTF-8 names. Older archives use CP437, but every Office
    // producer sets bit 11, and decoding as UTF-8 is right far more often than
    // not for the rest, so we do not carry a CP437 table.
    const name = decodeUtf8(nameBytes);
    const comment =
      commentLen > 0
        ? decodeUtf8(buf.subarray(p + 46 + nameLen + extraLen, p + 46 + nameLen + extraLen + commentLen))
        : undefined;

    // ZIP64 extended information overrides any saturated 32-bit field, in a
    // fixed order, and only for the fields that actually saturated.
    if (
      compressedSize === 0xffff_ffff ||
      uncompressedSize === 0xffff_ffff ||
      localOffset === 0xffff_ffff
    ) {
      const extraStart = p + 46 + nameLen;
      let e = extraStart;
      const extraEnd = extraStart + extraLen;
      while (e + 4 <= extraEnd) {
        const headerId = view.getUint16(e, true);
        const dataSize = view.getUint16(e + 2, true);
        if (headerId === 0x0001) {
          let q = e + 4;
          if (uncompressedSize === 0xffff_ffff) {
            uncompressedSize = Number(view.getBigUint64(q, true));
            q += 8;
          }
          if (compressedSize === 0xffff_ffff) {
            compressedSize = Number(view.getBigUint64(q, true));
            q += 8;
          }
          if (localOffset === 0xffff_ffff) {
            localOffset = Number(view.getBigUint64(q, true));
            q += 8;
          }
          break;
        }
        e += 4 + dataSize;
      }
    }

    if (uncompressedSize > maxEntrySize) {
      throw new ZipError(
        `entry ${name} declares ${uncompressedSize} bytes, above the ${maxEntrySize}-byte limit`,
      );
    }
    if (compressedSize > 0 && uncompressedSize / compressedSize > maxRatio) {
      throw new ZipError(
        `entry ${name} has a compression ratio of ${Math.round(uncompressedSize / compressedSize)}:1, ` +
          `above the ${maxRatio}:1 limit (possible zip bomb)`,
      );
    }

    const entry = makeEntry(buf, view, verifyCrc, {
      name,
      method,
      crc,
      compressedSize,
      uncompressedSize,
      localOffset,
      modified: dosDateTimeToDate(dosDate, dosTime),
      comment,
      hasDataDescriptor: (flags & 0x08) !== 0,
    });
    entries.set(name, entry);

    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

interface RawEntry {
  name: string;
  method: number;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  localOffset: number;
  modified: Date;
  comment?: string;
  hasDataDescriptor: boolean;
}

function makeEntry(buf: Uint8Array, view: DataView, verifyCrc: boolean, raw: RawEntry): ZipEntry {
  let cached: Uint8Array | undefined;
  return {
    name: raw.name,
    size: raw.uncompressedSize,
    compressedSize: raw.compressedSize,
    method: raw.method,
    crc32: raw.crc,
    modified: raw.modified,
    comment: raw.comment,
    data(): Uint8Array {
      if (cached) return cached;
      const lo = raw.localOffset;
      if (lo + 30 > buf.length) throw new ZipError(`local header for ${raw.name} is out of bounds`);
      if (view.getUint32(lo, true) !== SIG_LOCAL) {
        throw new ZipError(`bad local header signature for ${raw.name}`);
      }
      // The local header repeats the name/extra lengths, and its extra field can
      // differ in length from the central one, so we must read it here rather
      // than reuse the central directory's lengths.
      const nameLen = view.getUint16(lo + 26, true);
      const extraLen = view.getUint16(lo + 28, true);
      const start = lo + 30 + nameLen + extraLen;
      const end = start + raw.compressedSize;
      if (end > buf.length) throw new ZipError(`data for ${raw.name} extends past end of file`);
      const compressed = buf.subarray(start, end);

      let out: Uint8Array;
      if (raw.method === METHOD_STORE) {
        out = compressed;
      } else if (raw.method === METHOD_DEFLATE) {
        out = new Uint8Array(inflateRawSync(compressed));
      } else {
        throw new ZipError(`unsupported compression method ${raw.method} for ${raw.name}`);
      }
      if (raw.uncompressedSize !== 0 && out.length !== raw.uncompressedSize) {
        throw new ZipError(
          `size mismatch for ${raw.name}: header says ${raw.uncompressedSize}, got ${out.length}`,
        );
      }
      // A streaming writer that used a data descriptor leaves the central CRC
      // populated but may leave the local one zero; only a non-zero CRC is a
      // meaningful claim to check against.
      if (verifyCrc && raw.crc !== 0 && out.length > 0) {
        const actual = crc32(out);
        if (actual !== raw.crc) {
          throw new ZipError(
            `CRC mismatch for ${raw.name}: expected ${raw.crc.toString(16)}, got ${actual.toString(16)} (corrupt archive)`,
          );
        }
      }
      cached = out;
      return out;
    },
  };
}

/**
 * Locate the end-of-central-directory record by scanning backwards.
 *
 * The record is variable-length because of its trailing comment, so there is no
 * fixed offset. The comment is capped at 65535 bytes, which bounds the scan.
 */
function findEocd(buf: Uint8Array): number {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const min = Math.max(0, buf.length - 22 - 0xffff);
  for (let i = buf.length - 22; i >= min; i--) {
    if (view.getUint32(i, true) === SIG_EOCD) {
      // Guard against a false positive inside file data by checking that the
      // declared comment length lands exactly on the end of the buffer.
      const commentLen = view.getUint16(i + 20, true);
      if (i + 22 + commentLen === buf.length) return i;
    }
  }
  return -1;
}

function dosDateTimeToDate(date: number, time: number): Date {
  const year = 1980 + ((date >> 9) & 0x7f);
  const month = ((date >> 5) & 0x0f) - 1;
  const day = date & 0x1f;
  const hour = (time >> 11) & 0x1f;
  const minute = (time >> 5) & 0x3f;
  const second = (time & 0x1f) * 2;
  return new Date(Date.UTC(year, month, day, hour, minute, second));
}

function dateToDosDateTime(d: Date): { date: number; time: number } {
  // The DOS epoch starts in 1980; anything earlier clamps rather than wrapping.
  const year = Math.max(1980, d.getUTCFullYear());
  const date = ((year - 1980) << 9) | ((d.getUTCMonth() + 1) << 5) | d.getUTCDate();
  const time = (d.getUTCHours() << 11) | (d.getUTCMinutes() << 5) | (d.getUTCSeconds() >> 1);
  return { date, time };
}

export interface ZipWriteEntry {
  name: string;
  data: Uint8Array;
  /** Skip compression, e.g. for already-compressed image parts. */
  store?: boolean;
  modified?: Date;
}

export interface ZipWriteOptions {
  /** zlib level 0-9. 6 is the default trade-off; Excel itself uses roughly this. */
  level?: number;
  /**
   * Fixed timestamp for every entry. Setting this makes writes reproducible,
   * which is what our round-trip tests rely on.
   */
  modified?: Date;
}

/** Build a ZIP archive. Entry order is preserved, which matters for OOXML. */
export function writeZip(entries: ZipWriteEntry[], options: ZipWriteOptions = {}): Uint8Array {
  const level = options.level ?? 6;
  const chunks: Uint8Array[] = [];
  let offset = 0;

  interface CentralRecord {
    name: Uint8Array;
    method: number;
    crc: number;
    compressedSize: number;
    uncompressedSize: number;
    localOffset: number;
    dos: { date: number; time: number };
  }
  const central: CentralRecord[] = [];

  for (const entry of entries) {
    const nameBytes = encodeUtf8(entry.name);
    const dos = dateToDosDateTime(entry.modified ?? options.modified ?? new Date());
    const crc = crc32(entry.data);

    let method = METHOD_STORE;
    let payload = entry.data;
    if (!entry.store && entry.data.length > 0) {
      const deflated = new Uint8Array(deflateRawSync(entry.data, { level }));
      // Only take the compressed form if it actually saved bytes.
      if (deflated.length < entry.data.length) {
        method = METHOD_DEFLATE;
        payload = deflated;
      }
    }

    const local = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, SIG_LOCAL, true);
    lv.setUint16(4, 20, true); // version needed: 2.0 for deflate
    lv.setUint16(6, 0x0800, true); // flag bit 11: names are UTF-8
    lv.setUint16(8, method, true);
    lv.setUint16(10, dos.time, true);
    lv.setUint16(12, dos.date, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, payload.length, true);
    lv.setUint32(22, entry.data.length, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true); // no extra field
    local.set(nameBytes, 30);

    chunks.push(local, payload);
    central.push({
      name: nameBytes,
      method,
      crc,
      compressedSize: payload.length,
      uncompressedSize: entry.data.length,
      localOffset: offset,
      dos,
    });
    offset += local.length + payload.length;
  }

  const centralOffset = offset;
  for (const c of central) {
    const rec = new Uint8Array(46 + c.name.length);
    const cv = new DataView(rec.buffer);
    cv.setUint32(0, SIG_CENTRAL, true);
    cv.setUint16(4, 20, true); // version made by
    cv.setUint16(6, 20, true); // version needed
    cv.setUint16(8, 0x0800, true);
    cv.setUint16(10, c.method, true);
    cv.setUint16(12, c.dos.time, true);
    cv.setUint16(14, c.dos.date, true);
    cv.setUint32(16, c.crc, true);
    cv.setUint32(20, c.compressedSize, true);
    cv.setUint32(24, c.uncompressedSize, true);
    cv.setUint16(28, c.name.length, true);
    cv.setUint16(30, 0, true); // extra length
    cv.setUint16(32, 0, true); // comment length
    cv.setUint16(34, 0, true); // disk number
    cv.setUint16(36, 0, true); // internal attributes
    cv.setUint32(38, 0, true); // external attributes
    cv.setUint32(42, c.localOffset, true);
    rec.set(c.name, 46);
    chunks.push(rec);
    offset += rec.length;
  }
  const centralSize = offset - centralOffset;

  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, SIG_EOCD, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, central.length, true);
  ev.setUint16(10, central.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, centralOffset, true);
  ev.setUint16(20, 0, true);
  chunks.push(eocd);

  return concat(chunks);
}

function concat(chunks: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const c of chunks) total += c.length;
  const out = new Uint8Array(total);
  let p = 0;
  for (const c of chunks) {
    out.set(c, p);
    p += c.length;
  }
  return out;
}

const utf8Decoder = new TextDecoder('utf-8');
const utf8Encoder = new TextEncoder();

export function decodeUtf8(bytes: Uint8Array): string {
  return utf8Decoder.decode(bytes);
}

export function encodeUtf8(text: string): Uint8Array {
  return utf8Encoder.encode(text);
}

/** True when the buffer starts with a local file header - i.e. looks like a ZIP. */
export function looksLikeZip(buf: Uint8Array): boolean {
  return (
    buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b && buf[2] === 0x03 && buf[3] === 0x04
  );
}

export { SIG_DATA_DESCRIPTOR, METHOD_DEFLATE, METHOD_STORE };
