import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, '..');
const srcPluginDir = path.join(rootDir, 'dist', 'plugin');
const zipOutPath = path.join(rootDir, 'dist', 'antigravity-doc-viewer-offline.zip');

async function addFolderToZip(zipFolder, diskDir) {
  const entries = await fs.readdir(diskDir, { withFileTypes: true });
  for (const entry of entries) {
    const full = path.join(diskDir, entry.name);
    if (entry.isDirectory()) {
      await addFolderToZip(zipFolder.folder(entry.name), full);
    } else if (entry.isFile()) {
      const data = await fs.readFile(full);
      zipFolder.file(entry.name, data);
    }
  }
}

async function main() {
  try {
    await fs.access(path.join(srcPluginDir, 'plugin.json'));
  } catch {
    throw new Error('dist/plugin not found. Run `npm run build` first.');
  }

  const zip = new JSZip();
  const pluginRoot = zip.folder('doc-viewer');
  await addFolderToZip(pluginRoot, srcPluginDir);

  const buf = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  await fs.writeFile(zipOutPath, buf);
  console.log(`[pack] Created ${zipOutPath} (${(buf.byteLength / (1024 * 1024)).toFixed(2)} MiB)`);
}

main().catch((err) => {
  console.error('[pack] Failed:', err);
  process.exit(1);
});
