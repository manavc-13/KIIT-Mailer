import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import path from 'node:path';

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: { alias: { '@': path.resolve(import.meta.dirname, 'src') } },
  server: {
    port: 5173,
    // Build outputs must not trigger dev reloads.
    watch: { ignored: ['**/.vercel/**', '**/dist/**'] },
    proxy: {
      '/api': { target: 'http://localhost:8787', changeOrigin: false },
      '/u/': { target: 'http://localhost:8787', changeOrigin: false },
    },
  },
  // No source maps in the production build: dist/ is copied to Vercel's public static output.
  build: { outDir: 'dist', sourcemap: false },
});
