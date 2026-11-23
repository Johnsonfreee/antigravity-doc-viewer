import fs from 'node:fs/promises';
import path from 'node:path';

export interface RecentFileEntry {
  path: string;
  name: string;
  kind: string;
  lastOpenedAt: number;
}

export interface BookmarkEntry {
  path: string;
  name: string;
  kind: string;
  addedAt: number;
}

export interface FilePrefs {
  page?: number;
  zoom?: number;
  fitMode?: 'width' | 'custom';
  sheet?: string;
  scrollY?: number;
}

export interface ViewerState {
  recentFiles: RecentFileEntry[];
  bookmarks: BookmarkEntry[];
  lastRootDir?: string;
  filePrefs: Record<string, FilePrefs>;
}

const MAX_RECENT = 30;
const MAX_BOOKMARKS = 100;
const MAX_PREFS = 200;

function emptyState(): ViewerState {
  return {
    recentFiles: [],
    bookmarks: [],
    filePrefs: {},
  };
}

export class StateStore {
  private readonly filePath: string;
  private cached: ViewerState | null = null;

  constructor(dataDir: string) {
    this.filePath = path.join(dataDir, 'state.json');
  }

  async load(): Promise<ViewerState> {
    if (this.cached) return this.cached;
    try {
      const raw = await fs.readFile(this.filePath, 'utf-8');
      const parsed = JSON.parse(raw) as Partial<ViewerState>;
      this.cached = {
        recentFiles: Array.isArray(parsed.recentFiles) ? parsed.recentFiles.slice(0, MAX_RECENT) : [],
        bookmarks: Array.isArray(parsed.bookmarks) ? parsed.bookmarks.slice(0, MAX_BOOKMARKS) : [],
        lastRootDir: typeof parsed.lastRootDir === 'string' ? parsed.lastRootDir : undefined,
        filePrefs: parsed.filePrefs && typeof parsed.filePrefs === 'object' ? parsed.filePrefs : {},
      };
    } catch {
      this.cached = emptyState();
    }
    return this.cached;
  }

  async save(next: Partial<ViewerState>): Promise<ViewerState> {
    const current = await this.load();
    const merged: ViewerState = {
      recentFiles: Array.isArray(next.recentFiles) ? next.recentFiles.slice(0, MAX_RECENT) : current.recentFiles,
      bookmarks: Array.isArray(next.bookmarks) ? next.bookmarks.slice(0, MAX_BOOKMARKS) : current.bookmarks,
      lastRootDir: next.lastRootDir !== undefined ? next.lastRootDir : current.lastRootDir,
      filePrefs: next.filePrefs ? { ...current.filePrefs, ...next.filePrefs } : current.filePrefs,
    };

    const prefEntries = Object.entries(merged.filePrefs);
    if (prefEntries.length > MAX_PREFS) {
      merged.filePrefs = Object.fromEntries(prefEntries.slice(-MAX_PREFS));
    }

    await this.writeAtomic(merged);
    this.cached = merged;
    return merged;
  }

  async recordOpened(file: { path: string; name: string; kind: string }, prefsPatch?: FilePrefs): Promise<ViewerState> {
    const current = await this.load();
    const key = normalizePathKey(file.path);
    const filteredRecent = current.recentFiles.filter((r) => normalizePathKey(r.path) !== key);
    const recentFiles: RecentFileEntry[] = [
      { path: file.path, name: file.name, kind: file.kind, lastOpenedAt: Date.now() },
      ...filteredRecent,
    ].slice(0, MAX_RECENT);

    const filePrefs = { ...current.filePrefs };
    if (prefsPatch) {
      filePrefs[file.path] = { ...filePrefs[file.path], ...prefsPatch };
    }

    return this.save({ recentFiles, filePrefs });
  }

  private async writeAtomic(state: ViewerState): Promise<void> {
    const dir = path.dirname(this.filePath);
    await fs.mkdir(dir, { recursive: true });
    const tmpPath = `${this.filePath}.${process.pid}.${Date.now()}.tmp`;
    await fs.writeFile(tmpPath, JSON.stringify(state, null, 2), 'utf-8');
    await fs.rename(tmpPath, this.filePath);
  }
}

function normalizePathKey(p: string): string {
  const resolved = path.resolve(p);
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved;
}
