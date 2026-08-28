/**
 * Preload bundle.
 *
 * A sandboxed preload runs in a restricted context: no ESM, no `require` of
 * anything but a short allowlist, and no Node builtins. So this bundle must be
 * CommonJS and must contain everything it uses except `electron` itself - which
 * is why `channels.ts`, its only import, is deliberately free of Node imports.
 */

import { defineConfig } from 'vite';

export default defineConfig({
  resolve: { conditions: ['development'] },
  build: {
    outDir: 'dist/preload',
    emptyOutDir: true,
    target: 'chrome120',
    minify: false,
    sourcemap: true,
    lib: {
      entry: 'src/preload/index.ts',
      formats: ['cjs'],
      fileName: () => 'index.cjs',
    },
    rollupOptions: { external: ['electron'] },
  },
});
