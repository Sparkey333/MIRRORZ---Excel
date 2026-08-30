/**
 * Main-process tests.
 *
 * A separate config from the renderer's because these run in Node, with no
 * jsdom and no React testing setup: the code under test is the shell's pure
 * logic - path safety, IPC validation, the recent list, journals, the menu
 * template - and none of it should need a DOM to be exercised.
 *
 * Electron itself is never imported at runtime by anything these tests touch,
 * which is the point: the shell's decisions are testable without a display.
 */

import { defineConfig } from 'vitest/config';

export default defineConfig({
  resolve: { conditions: ['development'] },
  test: {
    include: ['test/main/**/*.test.ts'],
    environment: 'node',
    globals: false,
  },
});
