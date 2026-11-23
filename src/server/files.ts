import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export type DocKind =
  | 'pdf'
  | 'markdown'
  | 'docx'
  | 'legacy_doc'
  | 'spreadsheet'
  | 'html'
  | 'image'
  | 'code'
  | 'text'
  | 'binary';

export interface FileEntry {
  name: string;
  path: string;
  isDir: boolean;
  size: number;
  mtimeMs: number;
  ext: string;
  kind: DocKind | 'dir';
}

export interface FileMeta {
  path: string;
  name: string;
  dir: string;
  ext: string;
  kind: DocKind;
  size: number;
  mtimeMs: number;
}

export const MAX_CHUNK_BYTES = 384 * 1024; // 384 KiB raw -> ~512 KiB base64 (safely under 1 MiB MCP limit)

const IGNORED_DIRS = new Set([
  'node_modules',
  '.git',
  '.hg',
  '.svn',
  'dist',
  'build',
  'out',
  '.next',
  '.nuxt',
  '.venv',
  'venv',
  '__pycache__',
  '.cache',
  '.idea',
  '$RECYCLE.BIN',
  'System Volume Information',
]);

const EXT_KIND_MAP: Record<string, DocKind> = {
  '.pdf': 'pdf',
  '.md': 'markdown',
  '.markdown': 'markdown',
  '.mdx': 'markdown',
  '.docx': 'docx',
  '.doc': 'legacy_doc',
  '.xlsx': 'spreadsheet',
  '.xls': 'spreadsheet',
  '.xlsm': 'spreadsheet',
  '.xlsb': 'spreadsheet',
  '.ods': 'spreadsheet',
  '.csv': 'spreadsheet',
  '.tsv': 'spreadsheet',
  '.html': 'html',
  '.htm': 'html',
  '.xhtml': 'html',
  '.png': 'image',
  '.jpg': 'image',
  '.jpeg': 'image',
  '.gif': 'image',
  '.webp': 'image',
  '.svg': 'image',
  '.bmp': 'image',
  '.ico': 'image',
  '.ts': 'code',
  '.tsx': 'code',
  '.js': 'code',
  '.jsx': 'code',
  '.mjs': 'code',
  '.cjs': 'code',
  '.py': 'code',
  '.rs': 'code',
  '.go': 'code',
  '.java': 'code',
  '.kt': 'code',
  '.c': 'code',
  '.cpp': 'code',
  '.h': 'code',
  '.hpp': 'code',
  '.cs': 'code',
  '.rb': 'code',
  '.php': 'code',
  '.sh': 'code',
  '.bash': 'code',
  '.ps1': 'code',
  '.bat': 'code',
  '.cmd': 'code',
  '.sql': 'code',
  '.css': 'code',
  '.scss': 'code',
  '.less': 'code',
  '.xml': 'code',
  '.json': 'code',
  '.jsonc': 'code',
  '.yaml': 'code',
  '.yml': 'code',
  '.toml': 'code',
  '.ini': 'code',
  '.vue': 'code',
  '.svelte': 'code',
  '.txt': 'text',
  '.log': 'text',
  '.conf': 'text',
  '.env': 'text',
  '.gitignore': 'text',
  '.editorconfig': 'text',
};

const BINARY_EXTS = new Set([
  '.exe', '.dll', '.so', '.dylib', '.bin', '.dat', '.zip', '.tar', '.gz',
  '.7z', '.rar', '.woff', '.woff2', '.ttf', '.otf', '.eot', '.mp3', '.mp4',
  '.wav', '.avi', '.mov', '.mkv', '.sqlite', '.db', '.asar', '.pyc',
]);

export const DOCUMENT_EXTS = new Set([
  '.pdf', '.md', '.markdown', '.mdx', '.docx', '.xlsx', '.xls', '.xlsm',
  '.xlsb', '.ods', '.csv', '.tsv', '.html', '.htm', '.txt', '.json', '.yaml', '.yml',
]);

export function classifyFile(filePath: string): DocKind {
  const base = path.basename(filePath).toLowerCase();
  const ext = path.extname(base);
  if (EXT_KIND_MAP[ext]) return EXT_KIND_MAP[ext];
  if (BINARY_EXTS.has(ext)) return 'binary';
  if (!ext && (base === 'readme' || base === 'license' || base === 'makefile' || base === 'dockerfile')) {
    return 'text';
  }
  return ext ? 'text' : 'text';
}

export function resolveInputPath(rawPath: string, baseDir?: string): string {
  let trimmed = rawPath.trim();
  if (trimmed.startsWith('file://')) {
    try {
      const u = new URL(trimmed);
      trimmed = decodeURIComponent(u.pathname);
      if (process.platform === 'win32' && /^\/[A-Za-z]:/.test(trimmed)) {
        trimmed = trimmed.slice(1);
      }
    } catch {
      // Keep original string if URL parsing fails
    }
  }
  if (trimmed === '~' || trimmed.startsWith('~/') || trimmed.startsWith('~\\')) {
    trimmed = path.join(os.homedir(), trimmed.slice(2));
  }
  return path.isAbsolute(trimmed) ? path.normalize(trimmed) : path.resolve(baseDir || process.cwd(), trimmed);
}

export function isPluginInstallDir(dir: string): boolean {
  const normalized = path.normalize(dir).replace(/\\/g, '/').toLowerCase();
  return normalized.includes('/.gemini/config/plugins/');
}

export function getDefaultRootDir(): string {
  const cwd = process.cwd();
  return isPluginInstallDir(cwd) ? os.homedir() : cwd;
}

export function getQuickRoots(): Array<{ label: string; path: string }> {
  const candidates: Array<{ label: string; path: string }> = [];
  const seen = new Set<string>();

  const add = (label: string, dirPath: string) => {
    if (!dirPath) return;
    const norm = path.normalize(dirPath);
    const key = process.platform === 'win32' ? norm.toLowerCase() : norm;
    if (seen.has(key)) return;
    if (existsSync(norm)) {
      seen.add(key);
      candidates.push({ label, path: norm });
    }
  };

  const cwd = process.cwd();
  if (!isPluginInstallDir(cwd)) {
    add('Current Workspace', cwd);
  }
  const home = os.homedir();
  add('Home', home);
  add('Desktop', path.join(home, 'Desktop'));
  add('Documents', path.join(home, 'Documents'));
  add('Downloads', path.join(home, 'Downloads'));

  if (process.platform === 'win32') {
    for (let code = 67; code <= 90; code++) {
      const drive = `${String.fromCharCode(code)}:\\`;
      if (existsSync(drive)) {
        add(drive, drive);
      }
    }
  } else {
    add('/', '/');
  }

  return candidates;
}

export async function listDir(dirPath: string): Promise<{
  dir: string;
  parent: string | null;
  entries: FileEntry[];
}> {
  const resolved = resolveInputPath(dirPath);
  const dirents = await fs.readdir(resolved, { withFileTypes: true });
  const entries: FileEntry[] = [];

  for (const dirent of dirents) {
    if (dirent.name === '.' || dirent.name === '..') continue;
    const fullPath = path.join(resolved, dirent.name);
    try {
      const stat = await fs.stat(fullPath);
      const isDir = stat.isDirectory();
      const ext = isDir ? '' : path.extname(dirent.name).toLowerCase();
      entries.push({
        name: dirent.name,
        path: fullPath,
        isDir,
        size: isDir ? 0 : stat.size,
        mtimeMs: stat.mtimeMs,
        ext,
        kind: isDir ? 'dir' : classifyFile(fullPath),
      });
    } catch {
      // Skip inaccessible system entries gracefully
    }
  }

  entries.sort((a, b) => {
    if (a.isDir !== b.isDir) return a.isDir ? -1 : 1;
    return a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
  });

  const parentDir = path.dirname(resolved);
  const parent = parentDir !== resolved ? parentDir : null;
  return { dir: resolved, parent, entries };
}

export async function searchDocuments(
  rootDir: string,
  options: {
    query?: string;
    extensions?: string[];
    maxDepth?: number;
    limit?: number;
  } = {},
): Promise<{ rootDir: string; items: FileEntry[]; truncated: boolean }> {
  const resolvedRoot = resolveInputPath(rootDir);
  const queryLower = (options.query || '').trim().toLowerCase();
  const maxDepth = Math.min(Math.max(options.maxDepth ?? 5, 1), 10);
  const limit = Math.min(Math.max(options.limit ?? 120, 1), 300);
  const extFilter = options.extensions?.length
    ? new Set(options.extensions.map((e) => (e.startsWith('.') ? e.toLowerCase() : `.${e.toLowerCase()}`)))
    : null;

  const items: FileEntry[] = [];
  let truncated = false;
  const queue: Array<{ dir: string; depth: number }> = [{ dir: resolvedRoot, depth: 0 }];

  while (queue.length > 0) {
    const current = queue.shift()!;
    let dirents;
    try {
      dirents = await fs.readdir(current.dir, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const dirent of dirents) {
      if (items.length >= limit) {
        truncated = true;
        break;
      }
      const fullPath = path.join(current.dir, dirent.name);
      if (dirent.isDirectory()) {
        if (current.depth < maxDepth && !IGNORED_DIRS.has(dirent.name) && !dirent.name.startsWith('.')) {
          queue.push({ dir: fullPath, depth: current.depth + 1 });
        }
        continue;
      }
      if (!dirent.isFile()) continue;

      const ext = path.extname(dirent.name).toLowerCase();
      if (extFilter && !extFilter.has(ext)) continue;

      const relPath = path.relative(resolvedRoot, fullPath);
      if (queryLower && !relPath.toLowerCase().includes(queryLower) && !dirent.name.toLowerCase().includes(queryLower)) {
        continue;
      }

      try {
        const stat = await fs.stat(fullPath);
        items.push({
          name: dirent.name,
          path: fullPath,
          isDir: false,
          size: stat.size,
          mtimeMs: stat.mtimeMs,
          ext,
          kind: classifyFile(fullPath),
        });
      } catch {
        // Ignore stat error
      }
    }
    if (truncated) break;
  }

  return { rootDir: resolvedRoot, items, truncated };
}

export async function getFileMeta(rawPath: string): Promise<FileMeta> {
  const resolved = resolveInputPath(rawPath);
  const stat = await fs.stat(resolved);
  if (!stat.isFile()) {
    throw new Error(`Path is not a file: ${resolved}`);
  }
  return {
    path: resolved,
    name: path.basename(resolved),
    dir: path.dirname(resolved),
    ext: path.extname(resolved).toLowerCase(),
    kind: classifyFile(resolved),
    size: stat.size,
    mtimeMs: stat.mtimeMs,
  };
}

export async function readFileChunk(
  rawPath: string,
  offset = 0,
  length = MAX_CHUNK_BYTES,
): Promise<{
  base64: string;
  offset: number;
  bytesRead: number;
  totalSize: number;
  eof: boolean;
}> {
  const resolved = resolveInputPath(rawPath);
  const safeOffset = Math.max(0, Math.floor(offset));
  const safeLength = Math.min(Math.max(1, Math.floor(length)), MAX_CHUNK_BYTES);

  const handle = await fs.open(resolved, 'r');
  try {
    const stat = await handle.stat();
    const totalSize = stat.size;
    if (safeOffset >= totalSize) {
      return { base64: '', offset: safeOffset, bytesRead: 0, totalSize, eof: true };
    }
    const toRead = Math.min(safeLength, totalSize - safeOffset);
    const buffer = Buffer.allocUnsafe(toRead);
    const { bytesRead } = await handle.read(buffer, 0, toRead, safeOffset);
    const slice = bytesRead === toRead ? buffer : buffer.subarray(0, bytesRead);
    return {
      base64: slice.toString('base64'),
      offset: safeOffset,
      bytesRead,
      totalSize,
      eof: safeOffset + bytesRead >= totalSize,
    };
  } finally {
    await handle.close();
  }
}

export async function getPdfAsset(
  assetsRoot: string,
  kind: 'cmap' | 'standard_font',
  name: string,
): Promise<{ base64: string; byteLength: number; filename: string }> {
  if (!/^[A-Za-z0-9._-]+$/.test(name)) {
    throw new Error(`Invalid PDF asset name: ${name}`);
  }
  const subDir = kind === 'cmap' ? 'cmaps' : 'standard_fonts';
  const baseDir = path.resolve(assetsRoot, subDir);

  const candidates =
    kind === 'cmap' && !name.endsWith('.bcmap') ? [`${name}.bcmap`, name] : [name];

  for (const candidate of candidates) {
    const resolved = path.resolve(baseDir, candidate);
    if (!resolved.startsWith(baseDir + path.sep)) {
      throw new Error(`Path traversal rejected: ${name}`);
    }
    try {
      const buf = await fs.readFile(resolved);
      return {
        base64: buf.toString('base64'),
        byteLength: buf.byteLength,
        filename: candidate,
      };
    } catch {
      // Try next candidate
    }
  }

  throw new Error(`PDF asset not found (${kind}): ${name}`);
}
