/**
 * MS-OVBA compression, both directions.
 *
 * VBA does not store its streams with deflate. It uses a bespoke LZ77 variant
 * described in [MS-OVBA] section 2.4.1, and nothing else in the toolchain
 * speaks it, so it has to be written out by hand. The format is small but
 * unforgiving: a container is a signature byte followed by chunks, each chunk
 * decodes to at most 4096 bytes, and a copy token is two bytes whose split
 * between offset and length is not fixed. The split is derived from how far the
 * decoder currently is into the *current chunk*, so the same 16 bits mean
 * different things at different positions. That derivation (section 2.4.1.3.19.1,
 * "CopyToken Help") is where naive implementations go wrong: they assume a
 * constant 12/4 split, which happens to be right only for the last quarter of a
 * full chunk, so they decode short streams correctly and mangle long ones.
 *
 * Both directions live here because a round trip is the only cheap oracle we
 * have. We do not currently write vbaProject.bin - macro projects are copied
 * through byte-for-byte - but an encoder that the decoder agrees with, tested
 * against the worked examples in section 3.2 of the specification, is what gives
 * us confidence that the decoder is right rather than merely self-consistent.
 *
 * Security: this file turns bytes into other bytes. It never interprets what it
 * decodes. Everything here treats its input as hostile, because a macro-enabled
 * workbook is precisely the kind of file that arrives with bad intentions: chunk
 * lengths are clamped to the buffer, copy offsets are checked against what has
 * actually been written, and total output is capped so that a small file cannot
 * ask for an unbounded allocation.
 */

/** Raised when a container is not a well-formed MS-OVBA compressed stream. */
export class VbaCompressionError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'VbaCompressionError';
  }
}

/** CompressedContainer.SignatureByte, section 2.4.1.1.1. */
export const SIGNATURE_BYTE = 0x01;

/** CompressedChunkHeader.CompressedChunkSignature, bits 12-14. */
const CHUNK_SIGNATURE = 0b011;

/** A DecompressedChunk is at most this many bytes, section 2.4.1.1.3. */
const MAX_CHUNK_DECOMPRESSED = 4096;

/** CompressedChunk.CompressedData is at most this many bytes, section 2.4.1.1.4. */
const MAX_CHUNK_COMPRESSED = 4096;

/**
 * Ceiling on a single decompression, in bytes. A compressed chunk costs three
 * bytes and yields up to 4096, so without a cap a 12 KiB stream could demand
 * 16 MiB and a 12 MiB one could demand 16 GiB. Real dir streams and module
 * streams are measured in tens of kilobytes.
 */
const DEFAULT_MAX_OUTPUT = 64 * 1024 * 1024;

export interface DecompressOptions {
  /** Refuse to produce more than this many bytes. Default 64 MiB. */
  maxOutput?: number;
}

export interface CompressOptions {
  /**
   * Skip match finding and emit every byte as a literal token. The result is a
   * conforming container that any reader accepts, roughly 12.5% larger than the
   * input. Match finding is quadratic in the chunk length, so this is the escape
   * hatch for bulk data where the size does not matter.
   */
  literalsOnly?: boolean;
}

/**
 * CopyToken Help, section 2.4.1.3.19.1: how many of a copy token's 16 bits
 * carry the offset. The specification says "the smallest integer greater than
 * or equal to log2(difference)", floored at 4, where difference is the distance
 * from the start of the current decompressed chunk. Computed here with clz32 so
 * that no floating-point rounding can put us one bit out at a power of two.
 */
function copyTokenBitCount(difference: number): number {
  if (difference <= 16) return 4;
  return 32 - Math.clz32(difference - 1);
}

/** A growable byte buffer. Decompression needs random access to read its own output back. */
class ByteSink {
  private buf: Uint8Array;
  length = 0;

  constructor(
    capacity: number,
    private readonly limit: number,
  ) {
    this.buf = new Uint8Array(Math.max(64, Math.min(capacity, limit)));
  }

  private reserve(extra: number): void {
    const needed = this.length + extra;
    if (needed > this.limit) {
      throw new VbaCompressionError(`decompressed output would exceed the ${this.limit} byte limit`);
    }
    if (needed <= this.buf.length) return;
    let capacity = this.buf.length;
    while (capacity < needed) capacity *= 2;
    const grown = new Uint8Array(Math.min(capacity, this.limit));
    grown.set(this.buf.subarray(0, this.length));
    this.buf = grown;
  }

  push(byte: number): void {
    this.reserve(1);
    this.buf[this.length++] = byte;
  }

  write(src: Uint8Array): void {
    this.reserve(src.length);
    this.buf.set(src, this.length);
    this.length += src.length;
  }

  /** Copy `length` bytes from an earlier position, one byte at a time. */
  copyFrom(source: number, length: number): void {
    this.reserve(length);
    let from = source;
    for (let i = 0; i < length; i++) this.buf[this.length++] = this.buf[from++]!;
  }

  toUint8Array(): Uint8Array {
    return this.buf.slice(0, this.length);
  }
}

/**
 * Decompress a CompressedContainer, section 2.4.1.3.1.
 *
 * A truncated container is decoded as far as it goes rather than rejected: the
 * point of this code is to show a user their macros, and half a module is more
 * useful than an exception. Structural nonsense - a wrong signature, a copy
 * token pointing before the chunk - still throws, because that means we have
 * lost the frame and everything after it would be invented.
 */
export function decompress(input: Uint8Array, options: DecompressOptions = {}): Uint8Array {
  const limit = options.maxOutput ?? DEFAULT_MAX_OUTPUT;
  if (input.length === 0) throw new VbaCompressionError('empty compressed container');
  if (input[0] !== SIGNATURE_BYTE) {
    throw new VbaCompressionError(
      `bad container signature byte 0x${input[0]!.toString(16).padStart(2, '0')}, expected 0x01`,
    );
  }

  const out = new ByteSink(input.length * 2, limit);
  const end = input.length;
  let pos = 1;

  while (pos < end) {
    // A header needs two bytes. Anything less is a truncated file, not a chunk.
    if (pos + 2 > end) break;
    const header = input[pos]! | (input[pos + 1]! << 8);
    const signature = (header >> 12) & 0b111;
    if (signature !== CHUNK_SIGNATURE) {
      throw new VbaCompressionError(
        `chunk at ${pos} has signature 0b${signature.toString(2)}, expected 0b011`,
      );
    }
    const size = (header & 0x0fff) + 3;
    const isCompressed = (header & 0x8000) !== 0;
    const chunkEnd = Math.min(end, pos + size);
    const chunkStart = out.length;
    let cur = pos + 2;

    if (isCompressed) {
      while (cur < chunkEnd) {
        const flags = input[cur++]!;
        for (let index = 0; index < 8 && cur < chunkEnd; index++) {
          if ((flags & (1 << index)) === 0) {
            out.push(input[cur++]!);
          } else {
            if (cur + 2 > chunkEnd) {
              throw new VbaCompressionError(`copy token at ${cur} runs past the end of its chunk`);
            }
            const token = input[cur]! | (input[cur + 1]! << 8);
            cur += 2;
            const difference = out.length - chunkStart;
            const bitCount = copyTokenBitCount(difference);
            const lengthMask = 0xffff >>> bitCount;
            const length = (token & lengthMask) + 3;
            const offset = ((token & ~lengthMask & 0xffff) >>> (16 - bitCount)) + 1;
            if (offset > difference) {
              throw new VbaCompressionError(
                `copy token at ${cur - 2} reaches ${offset} bytes back, before the start of its chunk`,
              );
            }
            // Overlapping copies are legal and load-bearing: offset 1 with a
            // length of 72 is how a run of one byte is encoded, so this must
            // read back bytes it has only just written.
            out.copyFrom(out.length - offset, length);
          }
          if (out.length - chunkStart > MAX_CHUNK_DECOMPRESSED) {
            throw new VbaCompressionError(`chunk at ${pos} decoded to more than 4096 bytes`);
          }
        }
      }
    } else {
      // A raw chunk is 4096 verbatim bytes, but a truncated file may hold fewer.
      const available = Math.min(MAX_CHUNK_DECOMPRESSED, chunkEnd - cur);
      out.write(input.subarray(cur, cur + available));
    }

    // Resynchronise on the framing rather than on where the tokens left us: the
    // chunk size in the header is authoritative, and trusting it means one
    // malformed chunk cannot shift every chunk after it.
    pos = chunkEnd;
  }

  return out.toUint8Array();
}

/**
 * Matching, section 2.4.1.3.19.4: the longest run at `current` that also occurs
 * earlier in the same chunk. Returns a zero offset when nothing worth encoding
 * was found.
 *
 * Kept deliberately faithful to the specification's search order - candidates
 * are walked backwards from the nearest and only a strictly longer match
 * displaces the incumbent - so that our output matches the reference
 * implementation's byte for byte on the published examples. The two early exits
 * are the only departures and neither can change the result: a candidate whose
 * first byte differs cannot match at all, and no candidate can beat a match
 * that already runs to the end of the chunk.
 */
function findMatch(
  input: Uint8Array,
  chunkStart: number,
  current: number,
  end: number,
): { offset: number; length: number } {
  let bestLength = 0;
  let bestCandidate = 0;
  const maxPossible = end - current;
  for (let candidate = current - 1; candidate >= chunkStart; candidate--) {
    if (input[candidate] !== input[current]) continue;
    let c = candidate;
    let d = current;
    let len = 0;
    while (d < end && input[d] === input[c]) {
      len++;
      c++;
      d++;
    }
    if (len > bestLength) {
      bestLength = len;
      bestCandidate = candidate;
      if (bestLength >= maxPossible) break;
    }
  }
  if (bestLength < 3) return { offset: 0, length: 0 };
  const maximumLength = (0xffff >>> copyTokenBitCount(current - chunkStart)) + 3;
  return { offset: current - bestCandidate, length: Math.min(bestLength, maximumLength) };
}

/**
 * Compressing a DecompressedChunk, section 2.4.1.3.7, minus the header. Returns
 * the chunk body and how many input bytes it swallowed, which is less than
 * asked for when the encoded form would not fit in 4096 bytes.
 */
function encodeChunk(
  input: Uint8Array,
  start: number,
  end: number,
  literalsOnly: boolean,
): { body: Uint8Array; consumed: number } {
  const body = new Uint8Array(MAX_CHUNK_COMPRESSED);
  let p = 0;
  let d = start;

  while (d < end && p < MAX_CHUNK_COMPRESSED) {
    // The flag byte describes the eight tokens that follow, so its slot is
    // reserved now and filled in once they are known.
    const flagIndex = p;
    p++;
    let flags = 0;
    for (let index = 0; index < 8; index++) {
      if (d >= end || p >= MAX_CHUNK_COMPRESSED) break;
      const match = literalsOnly ? { offset: 0, length: 0 } : findMatch(input, start, d, end);
      if (match.offset !== 0) {
        if (p + 1 < MAX_CHUNK_COMPRESSED) {
          const bitCount = copyTokenBitCount(d - start);
          const token = (((match.offset - 1) << (16 - bitCount)) | (match.length - 3)) & 0xffff;
          body[p] = token & 0xff;
          body[p + 1] = (token >>> 8) & 0xff;
          flags |= 1 << index;
          p += 2;
          d += match.length;
        } else {
          p = MAX_CHUNK_COMPRESSED;
        }
      } else {
        body[p++] = input[d++]!;
      }
    }
    body[flagIndex] = flags;
  }

  return { body: body.slice(0, p), consumed: d - start };
}

/** Compression algorithm, section 2.4.1.3.6. */
export function compress(input: Uint8Array, options: CompressOptions = {}): Uint8Array {
  const literalsOnly = options.literalsOnly ?? false;
  const out = new ByteSink(Math.max(64, input.length), Number.POSITIVE_INFINITY);
  out.push(SIGNATURE_BYTE);

  let pos = 0;
  while (pos < input.length) {
    const chunkEnd = Math.min(pos + MAX_CHUNK_DECOMPRESSED, input.length);
    const { body, consumed } = encodeChunk(input, pos, chunkEnd, literalsOnly);

    if (consumed === chunkEnd - pos) {
      writeChunk(out, body, true);
      pos = chunkEnd;
      continue;
    }

    if (chunkEnd - pos === MAX_CHUNK_DECOMPRESSED) {
      // Incompressible full chunk: store it verbatim, which is what the
      // specification's encoder does and costs two bytes rather than 512.
      writeChunk(out, input.subarray(pos, chunkEnd), false);
      pos = chunkEnd;
      continue;
    }

    // An incompressible *short* tail is the one case where the specification's
    // encoder is lossy: a raw chunk must decode to exactly 4096 bytes, so it
    // pads with nulls that the decompressor then hands back as data. We emit
    // however much did fit as a compressed chunk and carry the rest into the
    // next one instead. Chunks may decode to fewer than 4096 bytes, so this is
    // conforming, and unlike the padding it round trips exactly.
    if (consumed <= 0) throw new VbaCompressionError('encoder made no progress');
    writeChunk(out, body, true);
    pos += consumed;
  }

  return out.toUint8Array();
}

function writeChunk(out: ByteSink, data: Uint8Array, compressed: boolean): void {
  // CompressedChunkSize counts the whole chunk, header included, minus three.
  const header = (CHUNK_SIGNATURE << 12) | ((data.length + 2 - 3) & 0x0fff) | (compressed ? 0x8000 : 0);
  out.push(header & 0xff);
  out.push((header >>> 8) & 0xff);
  out.write(data);
}
