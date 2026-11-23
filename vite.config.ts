import { defineConfig } from 'vite';
import { viteSingleFile } from 'vite-plugin-singlefile';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  root: path.resolve(__dirname, 'src/web'),
  plugins: [viteSingleFile({ removeViteModuleLoader: true })],
  build: {
    outDir: path.resolve(__dirname, 'dist/plugin/dist'),
    emptyOutDir: false,
    target: 'es2022',
    assetsInlineLimit: 100_000_000,
    cssCodeSplit: false,
    minify: 'esbuild',
    rollupOptions: {
      input: path.resolve(__dirname, 'src/web/index.html'),
      output: {
        inlineDynamicImports: true,
        entryFileNames: 'app.js',
        assetFileNames: '[name][extname]',
      },
    },
  },
});
