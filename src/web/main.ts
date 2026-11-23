import { ViewerMcpClient } from './mcpClient';
import { mountPdfViewer, type OutlineNode, type PdfController } from './viewers/pdfViewer';
import {
  mountHtmlView,
  mountImageView,
  mountRichDocView,
  mountSpreadsheetView,
  mountTextView,
  type SimpleViewerController,
} from './viewers/docViewers';

interface FileEntry {
  name: string;
  path: string;
  isDir: boolean;
  size: number;
  mtimeMs: number;
  ext: string;
  kind: string;
}

interface ViewerState {
  recentFiles: Array<{ path: string; name: string; kind: string; lastOpenedAt: number }>;
  bookmarks: Array<{ path: string; name: string; kind: string; addedAt: number }>;
  lastRootDir?: string;
  filePrefs: Record<
    string,
    { page?: number; zoom?: number; fitMode?: 'width' | 'custom'; sheet?: string; scrollY?: number }
  >;
}

interface OpenTab {
  id: string;
  path?: string;
  name: string;
  kind: string;
  ext: string;
  size: number;
  rawBuffer?: ArrayBuffer;
  page?: number;
  sheet?: string;
}

const FILTER_EXT_MAP: Record<string, string[] | undefined> = {
  all: undefined,
  pdf: ['.pdf'],
  docs: ['.md', '.markdown', '.mdx', '.docx', '.txt'],
  sheet: ['.xlsx', '.xls', '.xlsm', '.xlsb', '.ods', '.csv', '.tsv'],
  html: ['.html', '.htm', '.xhtml'],
  code: ['.ts', '.tsx', '.js', '.jsx', '.py', '.json', '.yaml', '.yml', '.css', '.sql', '.sh', '.ps1'],
};

const mcp = new ViewerMcpClient();

const tabs: OpenTab[] = [];
let activeTabId: string | null = null;
let currentDir = '';
let quickRoots: Array<{ label: string; path: string }> = [];
let currentEntries: FileEntry[] = [];
let activeFilter = 'all';
let searchQuery = '';
let state: ViewerState = { recentFiles: [], bookmarks: [], filePrefs: {} };
let activePdfCtrl: PdfController | null = null;
let activeSimpleCtrl: SimpleViewerController | null = null;
let currentOutline: OutlineNode[] = [];
let selectedSnippet = '';

const $ = <T extends HTMLElement>(sel: string) => document.querySelector<T>(sel)!;

const drawerEl = $('#drawer');
const tabsStripEl = $('#tabs-strip');
const pathInputEl = $<HTMLInputElement>('#path-input');
const quickRootsEl = $('#quick-roots');
const breadcrumbsEl = $('#breadcrumbs');
const fileListEl = $('#file-list');
const outlineListEl = $('#outline-list');
const bookmarksListEl = $('#bookmarks-list');
const recentListEl = $('#recent-list');
const welcomeEl = $('#viewer-welcome');
const viewerContentEl = $('#viewer-content');
const bookmarkBtn = $<HTMLButtonElement>('#btn-bookmark');
const nativeTabBtn = $<HTMLButtonElement>('#btn-native-tab');
const attachCtxBtn = $<HTMLButtonElement>('#btn-attach-ctx');
const askAgentBtn = $<HTMLButtonElement>('#btn-ask-agent');
const selectionBarEl = $('#selection-bar');
const selectionPreviewEl = $('#selection-preview');
const toastEl = $('#toast');

let toastTimer: number | undefined;
function showToast(msg: string) {
  toastEl.textContent = msg;
  toastEl.classList.remove('hidden');
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toastEl.classList.add('hidden'), 2600);
}

function formatSize(bytes: number): string {
  if (!bytes) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function kindBadgeLabel(kind: string, ext: string): string {
  if (kind === 'dir') return 'DIR';
  if (kind === 'pdf') return 'PDF';
  if (kind === 'markdown') return 'MD';
  if (kind === 'docx') return 'DOCX';
  if (kind === 'spreadsheet') return ext === '.csv' ? 'CSV' : 'XLS';
  if (kind === 'html') return 'HTML';
  if (kind === 'image') return 'IMG';
  if (kind === 'code') return ext.replace('.', '').toUpperCase() || 'CODE';
  return 'TXT';
}

function activeTab(): OpenTab | undefined {
  return tabs.find((t) => t.id === activeTabId);
}

function isBookmarked(filePath?: string): boolean {
  if (!filePath) return false;
  return state.bookmarks.some((b) => b.path.toLowerCase() === filePath.toLowerCase());
}

function updateCapabilitiesUi() {
  const canCtx = mcp.canUpdateContext();
  const canMsg = mcp.canMessage();
  attachCtxBtn.classList.toggle('hidden', !canCtx);
  askAgentBtn.classList.toggle('hidden', !canMsg);
  $('#sel-attach').classList.toggle('hidden', !canCtx);
  $('#sel-ask').classList.toggle('hidden', !canMsg);
}

function renderTabs() {
  tabsStripEl.innerHTML = '';
  for (const tab of tabs) {
    const el = document.createElement('div');
    el.className = `doc-tab ${tab.id === activeTabId ? 'active' : ''}`;
    el.title = tab.path || tab.name;

    const span = document.createElement('span');
    span.className = 'doc-tab-title';
    span.textContent = tab.name;
    el.appendChild(span);

    const closeBtn = document.createElement('button');
    closeBtn.className = 'doc-tab-close';
    closeBtn.textContent = '×';
    closeBtn.title = '关闭标签';
    closeBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      closeTab(tab.id);
    });
    el.appendChild(closeBtn);

    el.addEventListener('click', () => void activateTab(tab.id));
    tabsStripEl.appendChild(el);
  }

  const tab = activeTab();
  bookmarkBtn.disabled = !tab?.path;
  nativeTabBtn.disabled = !tab?.path;
  bookmarkBtn.classList.toggle('bookmarked', isBookmarked(tab?.path));
}

function renderOutline(items: OutlineNode[]) {
  currentOutline = items;
  if (!items.length) {
    outlineListEl.innerHTML = `<div class="empty-hint">当前文档无目录大纲</div>`;
    return;
  }
  outlineListEl.innerHTML = '';
  for (const item of items) {
    const btn = document.createElement('button');
    btn.className = 'outline-item';
    btn.style.paddingLeft = `${8 + (item.level - 1) * 12}px`;
    btn.textContent = item.page ? `${item.title} (p.${item.page})` : item.title;
    btn.title = item.title;
    btn.addEventListener('click', () => {
      if (item.page && activePdfCtrl) {
        activePdfCtrl.goToPage(item.page);
      } else if (item.id && activeSimpleCtrl?.scrollToHeading) {
        activeSimpleCtrl.scrollToHeading(item.id);
      }
      if (window.innerWidth <= 600) drawerEl.classList.add('collapsed');
    });
    outlineListEl.appendChild(btn);
  }
}

function renderHistory() {
  const renderSection = (
    target: HTMLElement,
    items: Array<{ path: string; name: string; kind: string }>,
    emptyText: string,
  ) => {
    if (!items.length) {
      target.innerHTML = `<div class="empty-hint">${emptyText}</div>`;
      return;
    }
    target.innerHTML = '';
    for (const item of items) {
      const btn = document.createElement('button');
      btn.className = 'file-item';
      btn.title = item.path;
      const ext = item.name.includes('.') ? item.name.slice(item.name.lastIndexOf('.')).toLowerCase() : '';
      btn.innerHTML = `<span class="file-badge ${item.kind}">${kindBadgeLabel(item.kind, ext)}</span><span class="file-name">${item.name}</span>`;
      btn.addEventListener('click', () => void openDiskFile(item.path));
      target.appendChild(btn);
    }
  };
  renderSection(bookmarksListEl, state.bookmarks, '暂无收藏文档');
  renderSection(recentListEl, state.recentFiles, '暂无最近打开文档');
}

function renderQuickRoots() {
  quickRootsEl.innerHTML = '';
  for (const root of quickRoots.slice(0, 8)) {
    const btn = document.createElement('button');
    btn.className = 'root-pill';
    btn.textContent = root.label;
    btn.title = root.path;
    btn.addEventListener('click', () => void navigateDir(root.path));
    quickRootsEl.appendChild(btn);
  }
}

function renderBreadcrumbs(dir: string) {
  breadcrumbsEl.innerHTML = '';
  if (!dir) return;
  const sep = dir.includes('\\') ? '\\' : '/';
  const parts = dir.split(sep).filter(Boolean);
  let acc = dir.startsWith('/') ? '/' : '';

  parts.forEach((part, idx) => {
    if (idx === 0 && /^[A-Za-z]:$/.test(part)) {
      acc = `${part}\\`;
    } else if (acc === '/' || acc.endsWith('\\')) {
      acc += part;
    } else {
      acc += `${sep}${part}`;
    }
    const targetPath = acc;
    const btn = document.createElement('button');
    btn.className = 'crumb-btn';
    btn.textContent = idx === 0 ? part : `${part}`;
    btn.addEventListener('click', () => void navigateDir(targetPath));
    breadcrumbsEl.appendChild(btn);

    if (idx < parts.length - 1) {
      const slash = document.createElement('span');
      slash.textContent = '/';
      breadcrumbsEl.appendChild(slash);
    }
  });
}

function renderFileList() {
  const allowedExts = FILTER_EXT_MAP[activeFilter];
  const q = searchQuery.trim().toLowerCase();

  const filtered = currentEntries.filter((entry) => {
    if (entry.isDir) return !q || entry.name.toLowerCase().includes(q);
    if (allowedExts && !allowedExts.includes(entry.ext)) return false;
    if (q && !entry.name.toLowerCase().includes(q) && !entry.path.toLowerCase().includes(q)) return false;
    return true;
  });

  if (!filtered.length) {
    fileListEl.innerHTML = `<div class="empty-hint">无匹配文件</div>`;
    return;
  }

  const currentPath = activeTab()?.path?.toLowerCase();
  fileListEl.innerHTML = '';

  for (const entry of filtered) {
    const btn = document.createElement('button');
    const isActive = !entry.isDir && currentPath && entry.path.toLowerCase() === currentPath;
    btn.className = `file-item ${isActive ? 'active' : ''}`;
    btn.title = entry.path;
    btn.innerHTML = `
      <span class="file-badge ${entry.kind}">${kindBadgeLabel(entry.kind, entry.ext)}</span>
      <span class="file-name">${entry.name}</span>
      <span class="file-size">${entry.isDir ? '' : formatSize(entry.size)}</span>
    `;
    btn.addEventListener('click', () => {
      if (entry.isDir) {
        void navigateDir(entry.path);
      } else {
        void openDiskFile(entry.path);
        if (window.innerWidth <= 600) drawerEl.classList.add('collapsed');
      }
    });
    fileListEl.appendChild(btn);
  }
}

let stateSaveTimer: number | undefined;
function scheduleSaveState(patch: Partial<ViewerState>) {
  window.clearTimeout(stateSaveTimer);
  stateSaveTimer = window.setTimeout(async () => {
    try {
      const res = await mcp.callTool<{ state: ViewerState }>('app_save_state', patch as Record<string, unknown>);
      state = res.state;
    } catch {
      // Ignore background save errors
    }
  }, 400);
}

async function navigateDir(dirPath: string) {
  try {
    const res = await mcp.callTool<{
      dir: string;
      parent: string | null;
      entries: FileEntry[];
      quickRoots: Array<{ label: string; path: string }>;
    }>('app_list_dir', { dir: dirPath });
    currentDir = res.dir;
    currentEntries = res.entries;
    if (res.quickRoots?.length) quickRoots = res.quickRoots;
    renderQuickRoots();
    renderBreadcrumbs(currentDir);
    renderFileList();
    scheduleSaveState({ lastRootDir: currentDir });
  } catch (err) {
    showToast(`打开目录失败：${err instanceof Error ? err.message : String(err)}`);
  }
}

async function runRecursiveSearch() {
  const q = searchQuery.trim();
  const exts = FILTER_EXT_MAP[activeFilter];
  if (!q && !exts) {
    if (currentDir) await navigateDir(currentDir);
    return;
  }
  if (!currentDir) return;
  try {
    const res = await mcp.callTool<{ items: FileEntry[] }>('app_search_files', {
      rootDir: currentDir,
      query: q || undefined,
      extensions: exts,
      maxDepth: 5,
      limit: 120,
    });
    currentEntries = res.items;
    renderFileList();
  } catch {
    renderFileList();
  }
}

function unmountActiveViewer() {
  activePdfCtrl?.destroy();
  activePdfCtrl = null;
  activeSimpleCtrl?.destroy?.();
  activeSimpleCtrl = null;
  renderOutline([]);
}

async function activateTab(tabId: string) {
  const tab = tabs.find((t) => t.id === tabId);
  if (!tab) return;
  activeTabId = tabId;
  renderTabs();
  renderFileList();
  pathInputEl.value = tab.path || tab.name;

  welcomeEl.classList.add('hidden');
  viewerContentEl.classList.remove('hidden');
  unmountActiveViewer();

  const prefs = tab.path ? state.filePrefs[tab.path] : undefined;

  try {
    if (tab.kind === 'pdf') {
      activePdfCtrl = await mountPdfViewer({
        container: viewerContentEl,
        mcp,
        filePath: tab.path,
        totalBytes: tab.size,
        rawBuffer: tab.rawBuffer,
        initialPage: tab.page ?? prefs?.page ?? 1,
        initialZoom: prefs?.zoom,
        initialFitMode: prefs?.fitMode,
        onOutline: renderOutline,
        onPageChange: (page) => {
          tab.page = page;
          if (tab.path) {
            state.filePrefs[tab.path] = { ...state.filePrefs[tab.path], page };
            scheduleSaveState({ filePrefs: { [tab.path]: state.filePrefs[tab.path] } });
          }
        },
        onZoomChange: (zoom, fitMode) => {
          if (tab.path) {
            state.filePrefs[tab.path] = { ...state.filePrefs[tab.path], zoom, fitMode };
            scheduleSaveState({ filePrefs: { [tab.path]: state.filePrefs[tab.path] } });
          }
        },
      });
    } else if (tab.kind === 'markdown' || tab.kind === 'docx') {
      if (!tab.path && tab.rawBuffer) {
        const text = new TextDecoder().decode(tab.rawBuffer);
        viewerContentEl.innerHTML = `<div class="viewer-scroll"><pre class="prose" style="white-space:pre-wrap">${text.replace(/</g, '&lt;')}</pre></div>`;
      } else if (tab.path) {
        activeSimpleCtrl = await mountRichDocView({
          container: viewerContentEl,
          mcp,
          kind: tab.kind,
          filePath: tab.path,
          onOutline: renderOutline,
        });
      }
    } else if (tab.kind === 'legacy_doc') {
      viewerContentEl.innerHTML = `<div class="error-box">暂不支持旧版二进制 <code>.doc</code> 格式，请转换为 <code>.docx</code> 或 <code>.pdf</code> 后在侧边栏打开。</div>`;
    } else if (tab.kind === 'spreadsheet') {
      if (tab.path) {
        activeSimpleCtrl = await mountSpreadsheetView({
          container: viewerContentEl,
          mcp,
          filePath: tab.path,
          initialSheet: tab.sheet ?? prefs?.sheet,
          initialPage: tab.page ?? 1,
          onSheetChange: (sheet) => {
            tab.sheet = sheet;
            if (tab.path) {
              state.filePrefs[tab.path] = { ...state.filePrefs[tab.path], sheet };
              scheduleSaveState({ filePrefs: { [tab.path]: state.filePrefs[tab.path] } });
            }
          },
        });
      } else {
        viewerContentEl.innerHTML = `<div class="error-box">拖拽 Excel 需要通过文件路径或右侧目录打开以便服务端分页解析。</div>`;
      }
    } else if (tab.kind === 'html') {
      activeSimpleCtrl = await mountHtmlView({
        container: viewerContentEl,
        mcp,
        filePath: tab.path,
        rawHtmlText: tab.rawBuffer ? new TextDecoder().decode(tab.rawBuffer) : undefined,
      });
    } else if (tab.kind === 'image') {
      activeSimpleCtrl = await mountImageView({
        container: viewerContentEl,
        mcp,
        filePath: tab.path,
        ext: tab.ext,
        rawBuffer: tab.rawBuffer,
      });
    } else if (tab.kind === 'binary') {
      viewerContentEl.innerHTML = `<div class="error-box">二进制文件（${tab.ext || 'binary'}）暂不支持直接文本预览。</div>`;
    } else {
      if (tab.path) {
        activeSimpleCtrl = await mountTextView({
          container: viewerContentEl,
          mcp,
          filePath: tab.path,
        });
      } else if (tab.rawBuffer) {
        const txt = new TextDecoder().decode(tab.rawBuffer);
        viewerContentEl.innerHTML = `<div class="viewer-scroll"><pre style="font-family:var(--font-code);white-space:pre-wrap">${txt.replace(/</g, '&lt;')}</pre></div>`;
      }
    }
  } catch (err) {
    viewerContentEl.innerHTML = `<div class="error-box">渲染失败：${
      err instanceof Error ? err.message : String(err)
    }</div>`;
  }
}

function closeTab(tabId: string) {
  const idx = tabs.findIndex((t) => t.id === tabId);
  if (idx === -1) return;
  tabs.splice(idx, 1);

  if (activeTabId === tabId) {
    if (tabs.length === 0) {
      activeTabId = null;
      unmountActiveViewer();
      viewerContentEl.classList.add('hidden');
      welcomeEl.classList.remove('hidden');
      pathInputEl.value = '';
      renderTabs();
    } else {
      const next = tabs[Math.max(0, idx - 1)];
      void activateTab(next.id);
    }
  } else {
    renderTabs();
  }
}

async function openDiskFile(filePath: string, opts: { page?: number; sheet?: string } = {}) {
  try {
    const res = await mcp.callTool<{
      file: { path: string; name: string; dir: string; ext: string; kind: string; size: number };
      state: ViewerState;
    }>('app_open_file_meta', { path: filePath, recordRecent: true });

    state = res.state;
    renderHistory();

    const existing = tabs.find((t) => t.path && t.path.toLowerCase() === res.file.path.toLowerCase());
    if (existing) {
      if (opts.page) existing.page = opts.page;
      if (opts.sheet) existing.sheet = opts.sheet;
      await activateTab(existing.id);
      if (opts.page && activePdfCtrl) activePdfCtrl.goToPage(opts.page);
      return;
    }

    const tab: OpenTab = {
      id: `disk:${res.file.path}`,
      path: res.file.path,
      name: res.file.name,
      kind: res.file.kind,
      ext: res.file.ext,
      size: res.file.size,
      page: opts.page,
      sheet: opts.sheet,
    };
    tabs.push(tab);
    await activateTab(tab.id);
  } catch (err) {
    showToast(`打开失败：${err instanceof Error ? err.message : String(err)}`);
  }
}

async function openLocalDroppedFile(file: File) {
  const name = file.name;
  const ext = name.includes('.') ? name.slice(name.lastIndexOf('.')).toLowerCase() : '';
  const kind =
    ext === '.pdf'
      ? 'pdf'
      : ext === '.html' || ext === '.htm'
        ? 'html'
        : ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.bmp', '.ico'].includes(ext)
          ? 'image'
          : ext === '.md' || ext === '.markdown'
            ? 'markdown'
            : ext === '.docx'
              ? 'docx'
              : ['.xlsx', '.xls', '.csv', '.tsv'].includes(ext)
                ? 'spreadsheet'
                : 'text';

  if (kind === 'docx' || kind === 'spreadsheet') {
    showToast(`${name} 请通过左侧文件树或路径栏打开以启用服务端完整解析`);
    return;
  }

  const rawBuffer = await file.arrayBuffer();
  const tab: OpenTab = {
    id: `drop:${Date.now()}:${name}`,
    name,
    kind,
    ext,
    size: file.size,
    rawBuffer,
  };
  tabs.push(tab);
  await activateTab(tab.id);
}

async function buildContextPayload(): Promise<string> {
  const tab = activeTab();
  if (!tab) return '';
  if (selectedSnippet) {
    return `[Doc Viewer Selection]\nFile: ${tab.path || tab.name}${activePdfCtrl ? ` (Page ${activePdfCtrl.getCurrentPage()})` : ''}\n\n${selectedSnippet}`;
  }
  if (activePdfCtrl) {
    const text = await activePdfCtrl.getCurrentPageText();
    return `[Doc Viewer PDF Page ${activePdfCtrl.getCurrentPage()}/${activePdfCtrl.getTotalPages()}]\nFile: ${tab.path || tab.name}\n\n${text.slice(0, 6000)}`;
  }
  if (activeSimpleCtrl) {
    return `[Doc Viewer Content]\nFile: ${tab.path || tab.name}\n\n${activeSimpleCtrl.getContextSummary()}`;
  }
  return `File: ${tab.path || tab.name}`;
}

function wireEvents() {
  $('#btn-toggle-drawer').addEventListener('click', () => {
    drawerEl.classList.toggle('collapsed');
  });

  document.querySelectorAll<HTMLButtonElement>('.drawer-tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.drawer!;
      document.querySelectorAll('.drawer-tab').forEach((b) => b.classList.toggle('active', b === btn));
      document.querySelectorAll('.drawer-pane').forEach((p) => p.classList.toggle('active', p.id === `drawer-${target}`));
    });
  });

  document.querySelectorAll<HTMLButtonElement>('#kind-chips .chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      activeFilter = chip.dataset.filter || 'all';
      document.querySelectorAll('#kind-chips .chip').forEach((c) => c.classList.toggle('active', c === chip));
      if (activeFilter === 'all' && !searchQuery.trim() && currentDir) {
        void navigateDir(currentDir);
      } else {
        void runRecursiveSearch();
      }
    });
  });

  let searchDebounce: number | undefined;
  $<HTMLInputElement>('#file-search-input').addEventListener('input', (e) => {
    searchQuery = (e.target as HTMLInputElement).value;
    window.clearTimeout(searchDebounce);
    searchDebounce = window.setTimeout(() => {
      if (!searchQuery.trim() && activeFilter === 'all' && currentDir) {
        void navigateDir(currentDir);
      } else {
        void runRecursiveSearch();
      }
    }, 260);
  });

  const openFromPathInput = () => {
    const val = pathInputEl.value.trim();
    if (val) void openDiskFile(val);
  };
  $('#btn-open-path').addEventListener('click', openFromPathInput);
  pathInputEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') openFromPathInput();
  });

  bookmarkBtn.addEventListener('click', () => {
    const tab = activeTab();
    if (!tab?.path) return;
    if (isBookmarked(tab.path)) {
      state.bookmarks = state.bookmarks.filter((b) => b.path.toLowerCase() !== tab.path!.toLowerCase());
      showToast('已取消收藏');
    } else {
      state.bookmarks = [
        { path: tab.path, name: tab.name, kind: tab.kind, addedAt: Date.now() },
        ...state.bookmarks,
      ];
      showToast('已加入收藏');
    }
    renderTabs();
    renderHistory();
    scheduleSaveState({ bookmarks: state.bookmarks });
  });

  nativeTabBtn.addEventListener('click', async () => {
    const tab = activeTab();
    if (!tab?.path) return;
    const ok = await mcp.openInNativeFileView(tab.path);
    if (!ok) {
      showToast('当前视图不支持打开原生文件标签页');
    }
  });

  attachCtxBtn.addEventListener('click', async () => {
    try {
      const ctx = await buildContextPayload();
      if (!ctx) return;
      await mcp.attachContext(ctx);
      showToast('已挂载为下一次对话上下文');
    } catch (err) {
      showToast(err instanceof Error ? err.message : String(err));
    }
  });

  askAgentBtn.addEventListener('click', async () => {
    const tab = activeTab();
    if (!tab) return;
    try {
      const ctx = await buildContextPayload();
      await mcp.sendMessage(`请总结并解释我当前正在查看的文档内容（${tab.name}）。`, ctx);
    } catch (err) {
      showToast(err instanceof Error ? err.message : String(err));
    }
  });

  $('#sel-attach').addEventListener('click', async () => {
    try {
      const ctx = await buildContextPayload();
      await mcp.attachContext(ctx);
      selectionBarEl.classList.add('hidden');
      showToast('选区已添加到对话上下文');
    } catch (err) {
      showToast(err instanceof Error ? err.message : String(err));
    }
  });

  $('#sel-ask').addEventListener('click', async () => {
    const tab = activeTab();
    try {
      const ctx = await buildContextPayload();
      selectionBarEl.classList.add('hidden');
      await mcp.sendMessage(`请解释我在 ${tab?.name || '文档'} 中选中的这段内容：`, ctx);
    } catch (err) {
      showToast(err instanceof Error ? err.message : String(err));
    }
  });

  document.addEventListener('selectionchange', () => {
    const sel = window.getSelection()?.toString().trim() || '';
    if (sel.length >= 4 && (mcp.canUpdateContext() || mcp.canMessage())) {
      selectedSnippet = sel.slice(0, 6000);
      selectionPreviewEl.textContent = `已选 ${sel.length} 字`;
      selectionBarEl.classList.remove('hidden');
    } else if (!sel) {
      selectedSnippet = '';
      selectionBarEl.classList.add('hidden');
    }
  });

  $<HTMLInputElement>('#file-picker').addEventListener('change', (e) => {
    const files = (e.target as HTMLInputElement).files;
    if (files?.[0]) void openLocalDroppedFile(files[0]);
  });

  const dropOverlay = $('#drop-overlay');
  let dragDepth = 0;
  document.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragDepth += 1;
    dropOverlay.classList.remove('hidden');
  });
  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('dragleave', () => {
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) dropOverlay.classList.add('hidden');
  });
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    dragDepth = 0;
    dropOverlay.classList.add('hidden');
    const file = e.dataTransfer?.files?.[0];
    if (file) void openLocalDroppedFile(file);
  });
}

async function bootstrap() {
  if (window.innerWidth <= 600) {
    drawerEl.classList.add('collapsed');
  }
  wireEvents();

  let handledInitialToolResult = false;

  await mcp.connect({
    onToolInput(args) {
      if (typeof args.path === 'string' && args.path) {
        pathInputEl.value = args.path;
      }
    },
    async onToolResult(payload, isError, errText) {
      if (isError) {
        showToast(errText || '工具执行出错');
        return;
      }
      if (!payload) return;
      handledInitialToolResult = true;

      if (Array.isArray(payload.quickRoots)) {
        quickRoots = payload.quickRoots as Array<{ label: string; path: string }>;
        renderQuickRoots();
      }
      if (payload.state) {
        state = payload.state as ViewerState;
        renderHistory();
      }
      if (typeof payload.rootDir === 'string' && payload.rootDir) {
        await navigateDir(payload.rootDir);
      }
      const file = payload.file as { path?: string } | null | undefined;
      if (file?.path) {
        await openDiskFile(file.path, {
          page: typeof payload.initialPage === 'number' ? payload.initialPage : undefined,
          sheet: typeof payload.initialSheet === 'string' ? payload.initialSheet : undefined,
        });
        if (window.innerWidth <= 600) {
          drawerEl.classList.add('collapsed');
        }
      }
    },
    onToolCancelled(reason) {
      showToast(reason ? `已取消：${reason}` : '操作已取消');
    },
  });

  updateCapabilitiesUi();

  // Fallback if launched cleanly without tool result payload yet
  window.setTimeout(async () => {
    if (handledInitialToolResult) return;
    try {
      const stRes = await mcp.callTool<{ state: ViewerState }>('app_get_state');
      state = stRes.state;
      renderHistory();
      await navigateDir(state.lastRootDir || '.');
    } catch {
      // Ignore standalone preview errors outside MCP host
    }
  }, 250);
}

void bootstrap();
