/**
 * Main-process bundle.
 *
 * CommonJS output on purpose. Electron can load an ESM main process, but the
 * package is `"type": "module"`, and a sandboxed preload cannot be ESM at all -
 * so emitting `.cjs` for both halves of the shell keeps one loader story rather
 * than two, and removes a class of "works in dev, fails when packaged" bug.
 *
 * Electron and every Node builtin stay external: they exist in the runtime and
 * bundling them would either fail or produce a second copy.
 */

import { builtinModules } from 'node:module';
import { defineConfig } from 'vite';

const external = [
  'electron',
  ...builtinModules,
  ...builtinModules.map((name) => `node:${name}`),
];

export default defineConfig({
  resolve: { conditions: ['development'] },
  build: {
    outDir: 'dist/main',
    emptyOutDir: true,
    target: 'node22',
    minify: false,
    sourcemap: true,
    lib: {
      entry: 'src/main/index.ts',
      formats: ['cjs'],
      fileName: () => 'index.cjs',
    },
    rollupOptions: { external },
  },
});
