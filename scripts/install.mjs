import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, '..');
const srcPluginDir = path.join(rootDir, 'dist', 'plugin');
const targetDir = path.join(os.homedir(), '.gemini', 'config', 'plugins', 'doc-viewer');

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
  try {
    await fs.access(path.join(srcPluginDir, 'plugin.json'));
  } catch {
    throw new Error('dist/plugin not found. Run `npm run build` first.');
  }

  console.log(`[install] Installing doc-viewer to ${targetDir} ...`);
  await copyDir(srcPluginDir, targetDir);
  console.log('[install] Done! Antigravity discovers ~/.gemini/config/plugins/doc-viewer automatically.');
}

main().catch((err) => {
  console.error('[install] Failed:', err);
  process.exit(1);
});
