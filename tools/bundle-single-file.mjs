/**
 * Inline the built renderer into one HTML file.
 *
 * The renderer already runs without Electron: `resolveHost()` falls back to a
 * browser host that opens files through a file input and saves through a
 * download, and the shell subscriptions no-op when no preload bridge exists. So
 * the same bundle that the desktop app loads is a working spreadsheet on its
 * own, provided nothing it needs sits in a second file - a module script cannot
 * import a sibling over file://, and a stylesheet link would be a second
 * request. Inlining both is what makes it openable straight from a disk with no
 * server, which is the whole point of shipping it this way.
 *
 * Run after `npm run build --workspace @mirrorz/desktop`.
 */

import { readFileSync, writeFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

const dist = new URL('../apps/desktop/dist/renderer/', import.meta.url).pathname;
const assets = join(dist, 'assets');
const pick = (ext) => {
  const name = readdirSync(assets).find((f) => f.endsWith(ext) && !f.endsWith('.map'));
  if (!name) throw new Error(`no ${ext} in ${assets} - build the renderer first`);
  const text = readFileSync(join(assets, name), 'utf8');
  // A literal </script> inside the bundle would close the inline script early
  // and leave the rest of it on the page as text. Vite does not emit one today;
  // this is here so that if some future dependency does, the build stops rather
  // than shipping a file that renders blank.
  if (/<\/script/i.test(text)) throw new Error(`${name} contains </script> and cannot be inlined`);
  return text;
};

const out = process.argv[2] ?? 'mirrorz-sheets.html';
// The replacements go through functions rather than strings on purpose.
// String.replace reads `$&`, `` $` ``, `$'` and `$1` in a *replacement* string
// as insertion patterns, and minified JavaScript is full of `$`. Passing the
// text as a string silently rewrites parts of the bundle - it fails as a syntax
// error a long way from the cause. A function replacement is taken literally.
let html = readFileSync(join(dist, 'index.html'), 'utf8');
html = html
  .replace(/\s*<script type="module"[^>]*src="[^"]+"><\/script>/, '')
  .replace(/\s*<link rel="stylesheet"[^>]*href="[^"]+">/, '')
  .replace('</head>', () => `    <style>\n${pick('.css')}\n    </style>\n  </head>`)
  .replace('</body>', () => `    <script type="module">\n${pick('.js')}\n    </script>\n  </body>`);



writeFileSync(out, html);
console.log(`${out}  ${(Buffer.byteLength(html) / 1024).toFixed(0)} KB`);
