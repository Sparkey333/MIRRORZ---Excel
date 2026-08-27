/**
 * Renderer tests.
 *
 * The React plugin is deliberately absent: vitest transforms .tsx with esbuild,
 * which reads the jsx setting from tsconfig, and pulling the plugin in would
 * bind the test run to a different Vite major than the dev server uses.
 */

import { defineConfig } from 'vitest/config';

export default defineConfig({
  esbuild: { jsx: 'automatic' },
  test: {
    include: ['test/**/*.test.ts', 'test/**/*.test.tsx'],
    environment: 'jsdom',
    globals: false,
    setupFiles: ['./test/setup.ts'],
  },
});
