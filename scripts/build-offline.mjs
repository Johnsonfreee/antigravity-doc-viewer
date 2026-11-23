import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build as viteBuild } from 'vite';
import { build as esbuild } from 'esbuild';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, '..');
const pluginOutDir = path.join(rootDir, 'dist', 'plugin');
const distOutDir = path.join(pluginOutDir, 'dist');

async function copyDir(src, dest) {
  await fs.mkdir(dest, { recursive: true });
  const entries = await fs.readdir(src, { withFileTypes: true });
  for (const entry of entries) {
    const s = path.join(src, entry.name);
    const d = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      await copyDir(s, d);
    } else if (entry.isFile()) {
      await fs.copyFile(s, d);
    }
  }
}

async function main() {
  console.log('[build] Cleaning dist/plugin ...');
  await fs.rm(pluginOutDir, { recursive: true, force: true });
  await fs.mkdir(distOutDir, { recursive: true });

  console.log('[build] Bundling single-file frontend app.html with Vite ...');
  await viteBuild({ configFile: path.join(rootDir, 'vite.config.ts') });

  const builtIndex = path.join(distOutDir, 'index.html');
  const appHtmlPath = path.join(distOutDir, 'app.html');
  await fs.rename(builtIndex, appHtmlPath);

  console.log('[build] Bundling standalone CJS MCP server server.cjs with esbuild ...');
  await esbuild({
    entryPoints: [path.join(rootDir, 'src', 'server', 'index.ts')],
    outfile: path.join(distOutDir, 'server.cjs'),
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node20',
    minify: true,
    sourcemap: false,
    logLevel: 'warning',
  });

  console.log('[build] Copying plugin metadata and skills ...');
  await copyDir(path.join(rootDir, 'plugin-meta'), pluginOutDir);

  console.log('[build] Copying offline PDF.js CMaps and standard fonts ...');
  await copyDir(
    path.join(rootDir, 'node_modules', 'pdfjs-dist', 'cmaps'),
    path.join(distOutDir, 'pdfjs-assets', 'cmaps'),
  );
  await copyDir(
    path.join(rootDir, 'node_modules', 'pdfjs-dist', 'standard_fonts'),
    path.join(distOutDir, 'pdfjs-assets', 'standard_fonts'),
  );

  // Verification assertions
  const appStat = await fs.stat(appHtmlPath);
  const serverStat = await fs.stat(path.join(distOutDir, 'server.cjs'));
  const appHtml = await fs.readFile(appHtmlPath, 'utf-8');

  const maxAppBytes = 10 * 1024 * 1024;
  if (appStat.size >= maxAppBytes) {
    throw new Error(`app.html exceeds 10 MiB limit (${appStat.size} bytes)`);
  }

  if (/<(?:script|link)\b[^>]*(?:src|href)\s*=\s*["']https?:\/\//i.test(appHtml)) {
    throw new Error('app.html contains external http(s) <script> or <link> references');
  }

  console.log(
    `[build] Success!\n` +
      `  app.html:   ${(appStat.size / (1024 * 1024)).toFixed(2)} MiB (${appStat.size} bytes)\n` +
      `  server.cjs: ${(serverStat.size / (1024 * 1024)).toFixed(2)} MiB (${serverStat.size} bytes)\n` +
      `  output dir: ${pluginOutDir}`,
  );
}

main().catch((err) => {
  console.error('[build] Failed:', err);
  process.exit(1);
});
