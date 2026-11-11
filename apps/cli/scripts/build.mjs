// Bundles the CLI and every workspace package it uses into one file, for npm or a plain copy.
// Node >= 22 runs it; no install step needed.  Output: dist/deckforge.cjs
import { build } from 'esbuild';

await build({
  entryPoints: ['src/main.ts'],
  outfile: 'dist/deckforge.cjs',
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'node22',
  sourcemap: false,
  legalComments: 'external',
  // The source has a shebang already; esbuild keeps it at the top of the bundle.
  // import.meta.url is only used to detect "running from source"; in the bundle it is never a .ts URL.
  define: {
    __DECKFORGE_RELEASE_DATE__: JSON.stringify(process.env.RELEASE_DATE ?? new Date().toISOString().slice(0, 10)),
    'import.meta.url': '"file:///deckforge.cjs"',
  },
  logLevel: 'warning',
});
console.log('built dist/deckforge.cjs');
