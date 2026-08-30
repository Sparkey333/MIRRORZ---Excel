/**
 * DEFLATE, implemented in portable JavaScript.
 *
 * The zip reader originally called `node:zlib` directly, which meant the whole
 * formats package could only run in Node. That is the wrong constraint for this
 * codebase: the same reader needs to work in an Electron renderer, in a web
 * worker doing background parsing, and in a browser, none of which have
 * `node:zlib`, and all of which need the decompression to be SYNCHRONOUS
 * because the zip entry API hands back bytes rather than a promise. The
 * browser's own `DecompressionStream` is async, so it cannot fill the gap.
 *
 * So the codec is implemented here and injectable: this is the default that
 * works everywhere, and `installNodeCodec()` swaps in zlib where it exists,
 * which is roughly an order of magnitude faster on large parts.
 *
 * Inflate is complete - stored, fixed-Huffman and dynamic-Huffman blocks.
 * Deflate emits stored blocks only: valid DEFLATE that any reader accepts, just
 * larger. That asymmetry is deliberate. Reading has to handle whatever a file
 * contains, so it must be complete; writing only has to produce something
 * correct, and in every environment where output size matters the Node codec is
 * present and takes over. Shipping a half-tuned Huffman encoder to save bytes in
 * a browser nobody writes files from would be the wrong trade.
 */

export class InflateError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'InflateError';
  }
}

/** Reads bits least-significant-first, as DEFLATE specifies. */
class BitReader {
  private pos = 0;
  private bitBuffer = 0;
  private bitCount = 0;

  constructor(private readonly data: Uint8Array) {}

  bits(count: number): number {
    while (this.bitCount < count) {
      if (this.pos >= this.data.length) throw new InflateError('unexpected end of compressed data');
      this.bitBuffer |= this.data[this.pos++]! << this.bitCount;
      this.bitCount += 8;
    }
    const value = this.bitBuffer & ((1 << count) - 1);
    this.bitBuffer >>>= count;
    this.bitCount -= count;
    return value;
  }

  /** Discard partial bits and return to a byte boundary, for stored blocks. */
  alignToByte(): void {
    this.bitBuffer = 0;
    this.bitCount = 0;
  }

  readBytes(count: number): Uint8Array {
    if (this.pos + count > this.data.length) {
      throw new InflateError('stored block extends past end of data');
    }
    const out = this.data.subarray(this.pos, this.pos + count);
    this.pos += count;
    return out;
  }

  get offset(): number {
    return this.pos;
  }
}

/**
 * A canonical Huffman decoding table.
 *
 * Decoding walks one bit at a time through the code-length counts, which is the
 * classic compact formulation: no lookup table to build, and it handles any
 * legal code assignment.
 */
interface HuffmanTable {
  /** Number of codes of each bit length, index 0 unused. */
  counts: Uint16Array;
  /** Symbols ordered by code, shortest codes first. */
  symbols: Uint16Array;
}

function buildHuffman(lengths: Uint8Array): HuffmanTable {
  const counts = new Uint16Array(16);
  for (const length of lengths) counts[length]!++;
  counts[0] = 0;

  const offsets = new Uint16Array(16);
  for (let i = 1; i < 16; i++) offsets[i] = offsets[i - 1]! + counts[i - 1]!;

  const symbols = new Uint16Array(lengths.length);
  for (let symbol = 0; symbol < lengths.length; symbol++) {
    const length = lengths[symbol]!;
    if (length !== 0) symbols[offsets[length]!++] = symbol;
  }
  return { counts, symbols };
}

function decodeSymbol(reader: BitReader, table: HuffmanTable): number {
  let code = 0;
  let first = 0;
  let index = 0;
  for (let length = 1; length < 16; length++) {
    code |= reader.bits(1);
    const count = table.counts[length]!;
    if (code - first < count) return table.symbols[index + (code - first)]!;
    index += count;
    first = (first + count) << 1;
    code <<= 1;
  }
  throw new InflateError('invalid Huffman code');
}

// Length and distance tables from RFC 1951 section 3.2.5.
const LENGTH_BASE = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131,
  163, 195, 227, 258,
];
const LENGTH_EXTRA = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
];
const DISTANCE_BASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049,
  3073, 4097, 6145, 8193, 12_289, 16_385, 24_577,
];
const DISTANCE_EXTRA = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13,
];
/** The order code lengths themselves are transmitted in, for dynamic blocks. */
const CODE_LENGTH_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

let fixedLiteralTable: HuffmanTable | undefined;
let fixedDistanceTable: HuffmanTable | undefined;

function fixedTables(): { literal: HuffmanTable; distance: HuffmanTable } {
  if (!fixedLiteralTable) {
    // The fixed code lengths are defined by the specification, not transmitted.
    const lengths = new Uint8Array(288);
    lengths.fill(8, 0, 144);
    lengths.fill(9, 144, 256);
    lengths.fill(7, 256, 280);
    lengths.fill(8, 280, 288);
    fixedLiteralTable = buildHuffman(lengths);
    fixedDistanceTable = buildHuffman(new Uint8Array(30).fill(5));
  }
  return { literal: fixedLiteralTable, distance: fixedDistanceTable! };
}

/** Growable output buffer; back-references read from what has been written. */
class OutputBuffer {
  private buffer: Uint8Array;
  private length = 0;

  constructor(initial: number, private readonly limit: number) {
    this.buffer = new Uint8Array(Math.max(1024, Math.min(initial, limit)));
  }

  private ensure(extra: number): void {
    if (this.length + extra <= this.buffer.length) return;
    if (this.length + extra > this.limit) {
      throw new InflateError(`decompressed data exceeds the ${this.limit}-byte limit`);
    }
    let size = this.buffer.length * 2;
    while (size < this.length + extra) size *= 2;
    const next = new Uint8Array(Math.min(size, this.limit));
    next.set(this.buffer.subarray(0, this.length));
    this.buffer = next;
  }

  push(byte: number): void {
    this.ensure(1);
    this.buffer[this.length++] = byte;
  }

  append(bytes: Uint8Array): void {
    this.ensure(bytes.length);
    this.buffer.set(bytes, this.length);
    this.length += bytes.length;
  }

  /**
   * Copy a back-reference. The copy may overlap the region being written - that
   * is how DEFLATE encodes runs - so it must proceed byte by byte rather than
   * as a block move.
   */
  copyBack(distance: number, count: number): void {
    if (distance > this.length) throw new InflateError('back-reference before start of output');
    this.ensure(count);
    let from = this.length - distance;
    for (let i = 0; i < count; i++) this.buffer[this.length++] = this.buffer[from++]!;
  }

  toBytes(): Uint8Array {
    return this.buffer.subarray(0, this.length);
  }

  get size(): number {
    return this.length;
  }
}

export interface InflateOptions {
  /** Expected output size, used to size the buffer. A hint, not a constraint. */
  expectedSize?: number;
  /** Hard ceiling, so a malformed or hostile stream cannot exhaust memory. */
  maxSize?: number;
}

/** Decompress a raw DEFLATE stream (no zlib or gzip wrapper). */
export function inflateRaw(data: Uint8Array, options: InflateOptions = {}): Uint8Array {
  const limit = options.maxSize ?? 2 * 1024 * 1024 * 1024;
  const out = new OutputBuffer(options.expectedSize ?? data.length * 4, limit);
  const reader = new BitReader(data);

  for (;;) {
    const isFinal = reader.bits(1);
    const type = reader.bits(2);

    if (type === 0) {
      reader.alignToByte();
      // A stored block repeats its length complemented, which is the only
      // integrity check DEFLATE itself provides.
      const header = reader.readBytes(4);
      const length = header[0]! | (header[1]! << 8);
      const complement = header[2]! | (header[3]! << 8);
      if ((length ^ 0xffff) !== complement) {
        throw new InflateError('stored block length does not match its complement');
      }
      out.append(reader.readBytes(length));
    } else if (type === 1 || type === 2) {
      const tables = type === 1 ? fixedTables() : dynamicTables(reader);
      inflateBlock(reader, tables.literal, tables.distance, out);
    } else {
      throw new InflateError('invalid block type');
    }

    if (isFinal) break;
  }
  return out.toBytes();
}

function dynamicTables(reader: BitReader): { literal: HuffmanTable; distance: HuffmanTable } {
  const literalCount = reader.bits(5) + 257;
  const distanceCount = reader.bits(5) + 1;
  const codeLengthCount = reader.bits(4) + 4;

  const codeLengths = new Uint8Array(19);
  for (let i = 0; i < codeLengthCount; i++) {
    codeLengths[CODE_LENGTH_ORDER[i]!] = reader.bits(3);
  }
  const codeLengthTable = buildHuffman(codeLengths);

  // The literal and distance lengths are themselves Huffman-coded, with three
  // repeat symbols that copy or emit runs of zeros.
  const lengths = new Uint8Array(literalCount + distanceCount);
  let i = 0;
  while (i < lengths.length) {
    const symbol = decodeSymbol(reader, codeLengthTable);
    if (symbol < 16) {
      lengths[i++] = symbol;
    } else if (symbol === 16) {
      if (i === 0) throw new InflateError('repeat code with no previous length');
      const previous = lengths[i - 1]!;
      const repeat = 3 + reader.bits(2);
      for (let r = 0; r < repeat && i < lengths.length; r++) lengths[i++] = previous;
    } else if (symbol === 17) {
      const repeat = 3 + reader.bits(3);
      for (let r = 0; r < repeat && i < lengths.length; r++) lengths[i++] = 0;
    } else {
      const repeat = 11 + reader.bits(7);
      for (let r = 0; r < repeat && i < lengths.length; r++) lengths[i++] = 0;
    }
  }

  return {
    literal: buildHuffman(lengths.subarray(0, literalCount)),
    distance: buildHuffman(lengths.subarray(literalCount)),
  };
}

function inflateBlock(
  reader: BitReader,
  literal: HuffmanTable,
  distance: HuffmanTable,
  out: OutputBuffer,
): void {
  for (;;) {
    const symbol = decodeSymbol(reader, literal);
    if (symbol < 256) {
      out.push(symbol);
      continue;
    }
    if (symbol === 256) return; // end of block
    const lengthIndex = symbol - 257;
    if (lengthIndex >= LENGTH_BASE.length) throw new InflateError('invalid length symbol');
    const length = LENGTH_BASE[lengthIndex]! + reader.bits(LENGTH_EXTRA[lengthIndex]!);

    const distanceSymbol = decodeSymbol(reader, distance);
    if (distanceSymbol >= DISTANCE_BASE.length) throw new InflateError('invalid distance symbol');
    const back = DISTANCE_BASE[distanceSymbol]! + reader.bits(DISTANCE_EXTRA[distanceSymbol]!);

    out.copyBack(back, length);
  }
}

/**
 * Produce a raw DEFLATE stream using stored blocks.
 *
 * Correct but uncompressed. See the file header for why that is the right
 * default: every environment that writes files in anger has `node:zlib`
 * installed over this.
 */
export function deflateRawStored(data: Uint8Array): Uint8Array {
  // A stored block's length field is 16 bits, so long input is split.
  const MAX_BLOCK = 0xffff;
  const blocks = Math.max(1, Math.ceil(data.length / MAX_BLOCK));
  const out = new Uint8Array(data.length + blocks * 5);
  let p = 0;
  let offset = 0;

  for (let b = 0; b < blocks; b++) {
    const size = Math.min(MAX_BLOCK, data.length - offset);
    const isFinal = b === blocks - 1 ? 1 : 0;
    // Block header: final flag plus type 00, then padding to a byte boundary.
    out[p++] = isFinal;
    out[p++] = size & 0xff;
    out[p++] = (size >> 8) & 0xff;
    out[p++] = ~size & 0xff;
    out[p++] = (~size >> 8) & 0xff;
    out.set(data.subarray(offset, offset + size), p);
    p += size;
    offset += size;
  }
  return out.subarray(0, p);
}

/** The codec the zip layer uses. Replaceable, so Node can install a faster one. */
export interface DeflateCodec {
  inflateRaw(data: Uint8Array, options?: InflateOptions): Uint8Array;
  deflateRaw(data: Uint8Array, level?: number): Uint8Array;
}

const portableCodec: DeflateCodec = {
  inflateRaw,
  deflateRaw: deflateRawStored,
};

let activeCodec: DeflateCodec = portableCodec;

export function getCodec(): DeflateCodec {
  return activeCodec;
}

export function setCodec(codec: DeflateCodec): void {
  activeCodec = codec;
}

export function resetCodec(): void {
  activeCodec = portableCodec;
}

export { portableCodec };
