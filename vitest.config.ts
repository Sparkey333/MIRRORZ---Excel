import { defineConfig } from 'vitest/config';

export default defineConfig({
  /**
   * Resolve workspace packages to their TypeScript sources.
   *
   * Each package exports `development` -> src and `default` -> dist. Node picks
   * `default`, which is what makes the published CLI runnable without a
   * TypeScript loader; asking for `development` here keeps every test and the
   * dev server compiling straight from source, so there is no build step
   * between editing a file and running its tests.
   */
  resolve: { conditions: ['development'] },
  test: {
    include: [
      'packages/*/test/**/*.test.ts',
      'packages/*/src/**/*.test.ts',
      'apps/*/test/**/*.test.ts',
    ],
    environment: 'node',
    setupFiles: ['packages/formats/test/setup-node-codec.ts'],
    testTimeout: 30_000,
    coverage: {
      provider: 'v8',
      reporter: ['text', 'lcov'],
      include: ['packages/*/src/**/*.ts'],
    },
  },
});
