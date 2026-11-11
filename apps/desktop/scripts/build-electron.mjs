// Bundles the Electron main and preload processes (TypeScript, plus the workspace packages
// they import) into CommonJS files Electron can load directly.
import { build } from 'esbuild';

const common = {
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'node22',
  sourcemap: true,
  external: ['electron'],
  logLevel: 'warning',
  // Release date decides which licenses cover this build (see docs/design-log/0007-licensing.md).
  define: { __DECKFORGE_RELEASE_DATE__: JSON.stringify(process.env.RELEASE_DATE ?? new Date().toISOString().slice(0, 10)) },
};

await Promise.all([
  build({ ...common, entryPoints: ['electron/main.ts'], outfile: 'dist-electron/main.cjs' }),
  build({ ...common, entryPoints: ['electron/preload.ts'], outfile: 'dist-electron/preload.cjs' }),
]);
console.log('built dist-electron/main.cjs and preload.cjs');
