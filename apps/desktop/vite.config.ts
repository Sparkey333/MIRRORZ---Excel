/**
 * Renderer build.
 *
 * `base: './'` matters: an Electron renderer is loaded from a file:// URL, and
 * the default absolute asset paths resolve against the filesystem root there,
 * so a build that works in the dev server shows a blank window when packaged.
 */

import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  resolve: { conditions: ['development'] },
  base: './',
  plugins: [react()],
  build: {
    outDir: 'dist/renderer',
    emptyOutDir: true,
    target: 'chrome120',
    sourcemap: true,
  },
  server: {
    port: 5273,
  },
});
