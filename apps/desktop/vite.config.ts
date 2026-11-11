import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Renderer only. The Electron main/preload processes are bundled by scripts/build-electron.mjs.
export default defineConfig({
  plugins: [react()],
  base: './', // loaded from file:// in the packaged app
  build: { outDir: 'dist', emptyOutDir: true, target: 'chrome140' },
  server: { port: 5199, strictPort: true },
});
