import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';

// The demo is a standalone stand-in for the SharePoint workbench.
export default defineConfig({
  root: fileURLToPath(new URL('.', import.meta.url)),
  plugins: [react()],
  css: {
    modules: { localsConvention: 'camelCase' },
    preprocessorOptions: { scss: { api: 'modern-compiler' } },
  },
  server: { port: 5173, host: true },
  build: { outDir: 'dist', emptyOutDir: true },
});
