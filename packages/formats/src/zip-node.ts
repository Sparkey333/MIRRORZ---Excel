/**
 * Installs Node's zlib as the deflate codec.
 *
 * Imported for its side effect by anything running under Node - the CLI, the
 * Electron main process, the test suite. The portable implementation in
 * inflate.ts is correct but perhaps an order of magnitude slower on a large
 * part, and its deflate does not compress at all, so anywhere zlib exists it
 * should be used.
 *
 * Kept in its own module so that bundling the formats package for a browser or
 * a renderer never pulls `node:zlib` into the graph.
 */

import { deflateRawSync, inflateRawSync } from 'node:zlib';
import { setCodec } from './inflate.js';

setCodec({
  inflateRaw: (data) => new Uint8Array(inflateRawSync(data)),
  deflateRaw: (data, level = 6) => new Uint8Array(deflateRawSync(data, { level })),
});
