// @ts-check
import js from '@eslint/js';
import tseslint from 'typescript-eslint';

/**
 * Boundaries (see docs/design-log/0002-boundaries-via-restricted-imports.md):
 * 1. Lite code (shared packages, cli, desktop) must never import Pro code.
 * 2. pptx-engine and data-xlsx are pure: no file, network or process I/O,
 *    and no imports from apps.
 */
const PRO_PACKAGES = ['@deckforge/server', '@deckforge/web', '@deckforge/connectors'];
const PRO_PATHS = ['**/apps/server/**', '**/apps/web/**', '**/packages/connectors/**'];
const APP_PACKAGES = ['@deckforge/cli', '@deckforge/desktop'];
const IO_MODULES = [
  'fs', 'fs/promises', 'child_process', 'net', 'http', 'https', 'http2',
  'dgram', 'dns', 'tls', 'worker_threads', 'cluster', 'os',
].flatMap((m) => [m, `node:${m}`]);

const noPro = {
  paths: PRO_PACKAGES.map((name) => ({
    name,
    message: 'Lite code must not import Pro packages (server, web, connectors).',
  })),
  patterns: [
    { group: PRO_PACKAGES.map((p) => `${p}/*`), message: 'Lite code must not import Pro packages.' },
    { group: PRO_PATHS, message: 'Lite code must not import Pro code.' },
  ],
};

export default tseslint.config(
  { ignores: ['**/node_modules/**', '**/dist/**', '**/coverage/**', 'fixtures/**'] },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    rules: {
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_', varsIgnorePattern: '^_' }],
      '@typescript-eslint/consistent-type-imports': 'error',
    },
  },
  {
    files: ['packages/pptx-engine/**', 'packages/data-xlsx/**', 'packages/licensing/**', 'apps/cli/**', 'apps/desktop/**', 'tools/**'],
    rules: { 'no-restricted-imports': ['error', noPro] },
  },
  {
    files: ['packages/pptx-engine/src/**', 'packages/data-xlsx/src/**'],
    rules: {
      'no-restricted-imports': ['error', {
        paths: [
          ...noPro.paths,
          ...IO_MODULES.map((name) => ({ name, message: 'Engine packages are pure: take buffers, return buffers. Do I/O in the app.' })),
          ...APP_PACKAGES.map((name) => ({ name, message: 'Shared packages must not import apps.' })),
        ],
        patterns: noPro.patterns,
      }],
      'no-restricted-globals': ['error',
        { name: 'fetch', message: 'Engine packages make no network calls.' },
        { name: 'process', message: 'Engine packages must not depend on the process environment.' },
        { name: 'Buffer', message: 'Use Uint8Array so the engine also runs outside Node.' },
      ],
    },
  },
);
