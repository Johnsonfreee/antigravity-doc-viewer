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

  console.log('[build] Copying plugin metadata ...');
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

  // Ensure CSS stylesheet was actually bundled and inlined into app.html (outside <script>)
  const htmlWithoutScripts = appHtml.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '');
  const styleBlocks = [...htmlWithoutScripts.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)].map((m) => m[1]);
  const totalCssBytes = styleBlocks.reduce((acc, s) => acc + Buffer.byteLength(s, 'utf-8'), 0);
  const combinedCss = styleBlocks.join('\n');
  if (styleBlocks.length === 0 || totalCssBytes < 4000 || !combinedCss.includes('.app-shell') || !combinedCss.includes('.hidden')) {
    throw new Error(
      `app.html is missing inlined UI CSS (<style> blocks: ${styleBlocks.length}, css bytes: ${totalCssBytes}). Check import './styles.css' in src/web/main.ts.`,
    );
  }

  console.log(
    `[build] Success!\n` +
      `  app.html:   ${(appStat.size / (1024 * 1024)).toFixed(2)} MiB (${appStat.size} bytes, inline CSS: ${(totalCssBytes / 1024).toFixed(1)} KiB)\n` +
      `  server.cjs: ${(serverStat.size / (1024 * 1024)).toFixed(2)} MiB (${serverStat.size} bytes)\n` +
      `  output dir: ${pluginOutDir}`,
  );
}

main().catch((err) => {
  console.error('[build] Failed:', err);
  process.exit(1);
});
