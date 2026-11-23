import { ViewerMcpClient } from './mcpClient';
import { mountPdfViewer, type OutlineNode, type PdfController } from './viewers/pdfViewer';
import { mountDocxOfficeViewer, type DocxController, type OfficeViewMode } from './viewers/docxOfficeViewer';
import { mountSheetOfficeViewer, type SheetController } from './viewers/sheetOfficeViewer';
import { mountSlideOfficeViewer, type SlideController } from './viewers/slideOfficeViewer';
import {
  mountHtmlView,
  mountImageView,
  mountRichDocView,
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

interface WorktreeMeta {
  originalPath: string;
  fileName: string;
  kind: string;
  status: 'draft' | 'clean';
  draftSize: number;
  summary: string;
  updatedAt: number;
  versions: Array<{ versionId: string; savedAt: number; size: number; summary: string }>;
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
  mode: OfficeViewMode;
  worktree?: WorktreeMeta | null;
}

const FILTER_EXT_MAP: Record<string, string[] | undefined> = {
  all: undefined,
  office: ['.docx', '.xlsx', '.xls', '.xlsm', '.xlsb', '.ods', '.csv', '.tsv', '.pptx'],
  pdf: ['.pdf'],
  docs: ['.md', '.markdown', '.mdx', '.txt'],
  html: ['.html', '.htm', '.xhtml'],
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
let selectedSnippet = '';
let fitModeFlag = true;

let activePdfCtrl: PdfController | null = null;
let activeDocxCtrl: DocxController | null = null;
let activeSheetCtrl: SheetController | null = null;
let activeSlideCtrl: SlideController | null = null;
let activeSimpleCtrl: SimpleViewerController | null = null;

const $ = <T extends HTMLElement>(sel: string) => document.querySelector<T>(sel)!;

const drawerEl = $('#drawer');
const rightDrawerEl = $('#right-drawer');
const tabsStripEl = $('#tabs-strip');
const pathInputEl = $<HTMLInputElement>('#path-input');
const modeAppBadgeEl = $('#mode-app-badge');
const worktreeBannerEl = $('#worktree-banner');
const worktreeBannerTextEl = $('#worktree-banner-text');
const rightDraftDotEl = $('#right-draft-dot');
const quickRootsEl = $('#quick-roots');
const breadcrumbsEl = $('#breadcrumbs');
const fileListEl = $('#file-list');
const outlineListEl = $('#outline-list');
const bookmarksListEl = $('#bookmarks-list');
const recentListEl = $('#recent-list');
const worktreePanelBodyEl = $('#worktree-panel-body');
const welcomeEl = $('#viewer-welcome');
const viewerContentEl = $('#viewer-content');
const plusMenuEl = $('#plus-menu');
const modeMenuEl = $('#mode-menu');
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

function badgeForKind(kind: string, ext: string): { cls: string; label: string } {
  if (kind === 'docx') return { cls: 'w', label: 'W' };
  if (kind === 'spreadsheet') return { cls: 'x', label: 'X' };
  if (kind === 'slides') return { cls: 'p', label: 'P' };
  if (kind === 'pdf') return { cls: 'pdf', label: 'PDF' };
  if (kind === 'markdown') return { cls: 'md', label: 'MD' };
  if (kind === 'html') return { cls: 'html', label: 'HTM' };
  return { cls: 'txt', label: ext ? ext.slice(1, 4).toUpperCase() : 'DOC' };
}

function kindBadgeLabel(kind: string, ext: string): string {
  if (kind === 'dir') return 'DIR';
  if (kind === 'pdf') return 'PDF';
  if (kind === 'markdown') return 'MD';
  if (kind === 'docx') return 'DOCX';
  if (kind === 'spreadsheet') return ext === '.csv' ? 'CSV' : 'XLS';
  if (kind === 'slides') return 'PPTX';
  if (kind === 'html') return 'HTML';
  if (kind === 'image') return 'IMG';
  return ext.replace('.', '').toUpperCase() || 'TXT';
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
  $('#btn-attach-ctx').classList.toggle('hidden', !canCtx);
  $('#btn-ask-agent').classList.toggle('hidden', !canMsg);
  $('#btn-ask-review').classList.toggle('hidden', !canMsg);
  $('#sel-attach').classList.toggle('hidden', !canCtx);
  $('#sel-ask').classList.toggle('hidden', !canMsg);
  $('#ai-cap-hint').textContent =
    canCtx || canMsg
      ? '选中文档文字或点击上方按钮，可直接把当前页面内容发送给对话中的 Agent。'
      : '当前为全局独立视图（无对话上下文），请在侧边栏会话中打开以使用对话联动。';
}

function renderTabs() {
  tabsStripEl.innerHTML = '';
  for (const tab of tabs) {
    const el = document.createElement('div');
    el.className = `doc-tab ${tab.id === activeTabId ? 'active' : ''}`;
    el.title = tab.path || tab.name;

    const badge = badgeForKind(tab.kind, tab.ext);
    const badgeEl = document.createElement('span');
    badgeEl.className = `app-badge ${badge.cls}`;
    badgeEl.textContent = badge.label;
    el.appendChild(badgeEl);

    const span = document.createElement('span');
    span.className = 'doc-tab-title';
    span.textContent = tab.name;
    el.appendChild(span);

    if (tab.worktree?.status === 'draft') {
      const dot = document.createElement('span');
      dot.className = 'draft-dot';
      dot.title = 'Worktree 有未合并草稿';
      el.appendChild(dot);
    }

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
  const b = tab ? badgeForKind(tab.kind, tab.ext) : { cls: 'w', label: 'W' };
  modeAppBadgeEl.className = `app-badge ${b.cls}`;
  modeAppBadgeEl.textContent = b.label;

  const hasDraft = tab?.worktree?.status === 'draft';
  worktreeBannerEl.classList.toggle('hidden', !hasDraft);
  rightDraftDotEl.classList.toggle('hidden', !hasDraft);
  if (hasDraft && tab?.worktree) {
    worktreeBannerTextEl.textContent = `Worktree 草稿待审核 · ${tab.worktree.summary || '有未合并修改'}`;
  }
}

function renderOutline(items: OutlineNode[]) {
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
      if (item.page && activePdfCtrl) activePdfCtrl.goToPage(item.page);
      else if (item.page && activeSlideCtrl) activeSlideCtrl.goToSlide(item.page);
      else if (item.id && activeDocxCtrl) activeDocxCtrl.scrollToHeading(item.id);
      else if (item.id && activeSimpleCtrl?.scrollToHeading) activeSimpleCtrl.scrollToHeading(item.id);
      if (window.innerWidth <= 680) drawerEl.classList.add('collapsed');
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

async function refreshRightWorktreePanel() {
  const tab = activeTab();
  if (!tab?.path) {
    worktreePanelBodyEl.innerHTML = `<div class="empty-hint">打开工作区文档后可管理 Worktree 与历史版本</div>`;
    return;
  }
  try {
    const res = await mcp.callTool<{ items: WorktreeMeta[] }>('app_get_worktree_status', { path: tab.path });
    const wt = res.items[0] ?? null;
    tab.worktree = wt;
    renderTabs();

    const statusLabel = wt?.status === 'draft' ? '有待审核草稿' : '与磁盘同步';
    const versionsHtml = (wt?.versions || []).length
      ? wt!.versions
          .map(
            (v) => `
          <div class="wt-version-row">
            <div style="min-width:0;flex:1">
              <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${v.summary || v.versionId}</div>
              <div style="font-size:10px;color:var(--fg-3)">${new Date(v.savedAt).toLocaleString()} · ${formatSize(v.size)}</div>
            </div>
            <button class="btn-sm" data-restore="${v.versionId}">恢复</button>
          </div>`,
          )
          .join('')
      : `<div class="empty-hint" style="padding:8px">暂无历史备份版本</div>`;

    worktreePanelBodyEl.innerHTML = `
      <div class="wt-card">
        <div class="wt-card-title">${tab.name}</div>
        <div style="font-size:11.5px;color:var(--fg-2)">状态：${statusLabel}</div>
        ${wt?.summary ? `<div style="font-size:11.5px;color:var(--fg-2)">最近说明：${wt.summary}</div>` : ''}
        <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:4px">
          <button id="rp-edit" class="btn-sm">进入编辑</button>
          <button id="rp-compare" class="btn-sm">查看差异</button>
          ${wt?.status === 'draft' ? '<button id="rp-approve" class="btn-sm primary">合并保存</button><button id="rp-discard" class="btn-sm danger">丢弃</button>' : ''}
        </div>
      </div>
      <div class="section-title" style="padding-left:2px">版本历史</div>
      ${versionsHtml}
    `;

    worktreePanelBodyEl.querySelector('#rp-edit')?.addEventListener('click', () => void switchActiveMode('edit'));
    worktreePanelBodyEl.querySelector('#rp-compare')?.addEventListener('click', () => void switchActiveMode('compare'));
    worktreePanelBodyEl.querySelector('#rp-approve')?.addEventListener('click', () => void handleWorktreeApprove());
    worktreePanelBodyEl.querySelector('#rp-discard')?.addEventListener('click', () => void handleWorktreeDiscard());
    worktreePanelBodyEl.querySelectorAll<HTMLButtonElement>('[data-restore]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        await mcp.callTool('app_worktree_action', {
          action: 'restore_version',
          path: tab.path,
          versionId: btn.dataset.restore,
          mode: 'draft',
        });
        showToast('已恢复版本到草稿 Worktree');
        await refreshRightWorktreePanel();
        await activateTab(tab.id);
      });
    });
  } catch {
    worktreePanelBodyEl.innerHTML = `<div class="empty-hint">该文件暂无 Worktree 记录</div>`;
  }
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
    if (idx === 0 && /^[A-Za-z]:$/.test(part)) acc = `${part}\\`;
    else if (acc === '/' || acc.endsWith('\\')) acc += part;
    else acc += `${sep}${part}`;
    const targetPath = acc;
    const btn = document.createElement('button');
    btn.className = 'crumb-btn';
    btn.textContent = part;
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
        if (window.innerWidth <= 680) drawerEl.classList.add('collapsed');
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
  activeDocxCtrl?.destroy();
  activeDocxCtrl = null;
  activeSheetCtrl?.destroy();
  activeSheetCtrl = null;
  activeSlideCtrl?.destroy();
  activeSlideCtrl = null;
  activeSimpleCtrl?.destroy?.();
  activeSimpleCtrl = null;
  renderOutline([]);
}

async function switchActiveMode(mode: OfficeViewMode) {
  const tab = activeTab();
  if (!tab) return;
  tab.mode = mode;
  if (activeDocxCtrl) await activeDocxCtrl.setMode(mode);
  else if (activeSheetCtrl) await activeSheetCtrl.setMode(mode);
  else if (activeSlideCtrl) await activeSlideCtrl.setMode(mode);
  else showToast('当前文件类型仅支持阅读模式');
}

async function handleWorktreeApprove() {
  const tab = activeTab();
  if (!tab?.path) return;
  await mcp.callTool('app_worktree_action', { action: 'approve', path: tab.path });
  showToast('已将 Worktree 草稿合并保存到磁盘');
  await refreshRightWorktreePanel();
  if (tab.mode === 'compare') tab.mode = 'print';
  await activateTab(tab.id);
}

async function handleWorktreeDiscard() {
  const tab = activeTab();
  if (!tab?.path) return;
  await mcp.callTool('app_worktree_action', { action: 'discard', path: tab.path });
  showToast('已丢弃 Worktree 草稿');
  await refreshRightWorktreePanel();
  if (tab.mode === 'compare') tab.mode = 'print';
  await activateTab(tab.id);
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
    if (tab.kind === 'docx') {
      activeDocxCtrl = await mountDocxOfficeViewer({
        container: viewerContentEl,
        mcp,
        filePath: tab.path,
        rawBuffer: tab.rawBuffer,
        mode: tab.mode,
        zoomSetting: fitModeFlag ? 'fit' : 1,
        onOutline: renderOutline,
        onWorktreeChanged: () => void refreshRightWorktreePanel(),
        onToast: showToast,
      });
    } else if (tab.kind === 'spreadsheet') {
      if (tab.path) {
        activeSheetCtrl = await mountSheetOfficeViewer({
          container: viewerContentEl,
          mcp,
          filePath: tab.path,
          initialSheet: tab.sheet ?? prefs?.sheet,
          initialPage: tab.page ?? 1,
          mode: tab.mode,
          onSheetChange: (sheet) => {
            tab.sheet = sheet;
            if (tab.path) {
              state.filePrefs[tab.path] = { ...state.filePrefs[tab.path], sheet };
              scheduleSaveState({ filePrefs: { [tab.path]: state.filePrefs[tab.path] } });
            }
          },
          onWorktreeChanged: () => void refreshRightWorktreePanel(),
          onToast: showToast,
        });
      } else {
        viewerContentEl.innerHTML = `<div class="error-box">请通过文件路径或目录打开电子表格文件。</div>`;
      }
    } else if (tab.kind === 'slides') {
      if (tab.path) {
        activeSlideCtrl = await mountSlideOfficeViewer({
          container: viewerContentEl,
          mcp,
          filePath: tab.path,
          initialSlide: tab.page ?? 1,
          mode: tab.mode,
          onOutline: renderOutline,
          onWorktreeChanged: () => void refreshRightWorktreePanel(),
          onToast: showToast,
        });
      } else {
        viewerContentEl.innerHTML = `<div class="error-box">请通过文件路径或目录打开 PowerPoint (.pptx) 文件。</div>`;
      }
    } else if (tab.kind === 'pdf') {
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
    } else if (tab.kind === 'markdown') {
      if (!tab.path && tab.rawBuffer) {
        const text = new TextDecoder().decode(tab.rawBuffer);
        viewerContentEl.innerHTML = `<div class="viewer-scroll"><pre class="prose" style="white-space:pre-wrap">${text.replace(/</g, '&lt;')}</pre></div>`;
      } else if (tab.path) {
        activeSimpleCtrl = await mountRichDocView({
          container: viewerContentEl,
          mcp,
          kind: 'markdown',
          filePath: tab.path,
          onOutline: renderOutline,
        });
      }
    } else if (tab.kind === 'legacy_doc') {
      viewerContentEl.innerHTML = `<div class="error-box">暂不支持旧版二进制 <code>.doc</code> 格式，请转换为 <code>.docx</code> 或 <code>.pdf</code> 后打开。</div>`;
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
    } else if (tab.path) {
      activeSimpleCtrl = await mountTextView({ container: viewerContentEl, mcp, filePath: tab.path });
    } else if (tab.rawBuffer) {
      const txt = new TextDecoder().decode(tab.rawBuffer);
      viewerContentEl.innerHTML = `<div class="viewer-scroll"><pre style="font-family:var(--font-code);color:var(--chrome-fg);white-space:pre-wrap">${txt.replace(/</g, '&lt;')}</pre></div>`;
    }
  } catch (err) {
    viewerContentEl.innerHTML = `<div class="error-box">渲染失败：${
      err instanceof Error ? err.message : String(err)
    }</div>`;
  }

  void refreshRightWorktreePanel();
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
      void refreshRightWorktreePanel();
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
      worktree?: WorktreeMeta | null;
      state: ViewerState;
    }>('app_open_file_meta', { path: filePath, recordRecent: true });

    state = res.state;
    renderHistory();

    const existing = tabs.find((t) => t.path && t.path.toLowerCase() === res.file.path.toLowerCase());
    if (existing) {
      if (opts.page) existing.page = opts.page;
      if (opts.sheet) existing.sheet = opts.sheet;
      existing.worktree = res.worktree ?? null;
      await activateTab(existing.id);
      if (opts.page && activePdfCtrl) activePdfCtrl.goToPage(opts.page);
      if (opts.page && activeSlideCtrl) activeSlideCtrl.goToSlide(opts.page);
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
      mode: 'print',
      worktree: res.worktree ?? null,
    };
    tabs.push(tab);
    await activateTab(tab.id);
  } catch (err) {
    showToast(`打开失败：${err instanceof Error ? err.message : String(err)}`);
  }
}

async function createOfficeDocument(kind: 'docx' | 'xlsx' | 'pptx') {
  const baseDir = currentDir || quickRoots[0]?.path || '.';
  const sep = baseDir.includes('\\') ? '\\' : '/';
  const ts = new Date().toISOString().slice(11, 19).replace(/:/g, '');
  const targetPath = `${baseDir}${baseDir.endsWith(sep) ? '' : sep}untitled-${ts}.${kind}`;
  try {
    const res = await mcp.callTool<{ file: { path: string } }>('app_create_office_file', {
      path: targetPath,
      kind,
      title: `新建 ${kind.toUpperCase()} 文档`,
    });
    await openDiskFile(res.file.path);
    await switchActiveMode('edit');
    if (currentDir) void navigateDir(currentDir);
  } catch (err) {
    showToast(`新建失败：${err instanceof Error ? err.message : String(err)}`);
  }
}

async function openLocalDroppedFile(file: File) {
  const name = file.name;
  const ext = name.includes('.') ? name.slice(name.lastIndexOf('.')).toLowerCase() : '';
  const kind =
    ext === '.docx'
      ? 'docx'
      : ext === '.pdf'
        ? 'pdf'
        : ext === '.html' || ext === '.htm'
          ? 'html'
          : ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.bmp', '.ico'].includes(ext)
            ? 'image'
            : ext === '.md' || ext === '.markdown'
              ? 'markdown'
              : ['.xlsx', '.xls', '.csv', '.tsv', '.pptx'].includes(ext)
                ? 'spreadsheet'
                : 'text';

  if (ext === '.xlsx' || ext === '.xls' || ext === '.csv' || ext === '.tsv' || ext === '.pptx') {
    showToast(`${name} 请通过文件夹按钮或路径栏打开，以启用完整的服务端编辑与 Worktree 能力`);
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
    mode: 'print',
  };
  tabs.push(tab);
  await activateTab(tab.id);
}

async function buildContextPayload(): Promise<string> {
  const tab = activeTab();
  if (!tab) return '';
  if (selectedSnippet) {
    return `[Office Workbench Selection]\nFile: ${tab.path || tab.name}${activePdfCtrl ? ` (Page ${activePdfCtrl.getCurrentPage()})` : ''}\n\n${selectedSnippet}`;
  }
  if (activePdfCtrl) {
    const text = await activePdfCtrl.getCurrentPageText();
    return `[PDF Page ${activePdfCtrl.getCurrentPage()}/${activePdfCtrl.getTotalPages()}]\nFile: ${tab.path || tab.name}\n\n${text.slice(0, 6000)}`;
  }
  if (activeDocxCtrl) {
    return `[Word Document]\nFile: ${tab.path || tab.name}\n\n${activeDocxCtrl.getContextSummary()}`;
  }
  if (activeSheetCtrl) {
    return `[Spreadsheet]\nFile: ${tab.path || tab.name}\n\n${activeSheetCtrl.getContextSummary()}`;
  }
  if (activeSlideCtrl) {
    return `[Presentation]\nFile: ${tab.path || tab.name}\n\n${activeSlideCtrl.getContextSummary()}`;
  }
  if (activeSimpleCtrl) {
    return `[Document]\nFile: ${tab.path || tab.name}\n\n${activeSimpleCtrl.getContextSummary()}`;
  }
  return `File: ${tab.path || tab.name}`;
}

function positionPopover(menu: HTMLElement, anchor: HTMLElement) {
  const rect = anchor.getBoundingClientRect();
  menu.style.top = `${rect.bottom + 4}px`;
  const maxLeft = Math.max(8, window.innerWidth - 210);
  menu.style.left = `${Math.min(rect.left, maxLeft)}px`;
  menu.classList.remove('hidden');
}

function closePopovers() {
  plusMenuEl.classList.add('hidden');
  modeMenuEl.classList.add('hidden');
}

function wireEvents() {
  const toggleLeftDrawer = () => {
    drawerEl.classList.toggle('collapsed');
    $('#btn-toggle-drawer').classList.toggle('active', !drawerEl.classList.contains('collapsed'));
  };
  $('#btn-toggle-drawer').addEventListener('click', toggleLeftDrawer);
  $('#btn-folder-browse').addEventListener('click', () => {
    drawerEl.classList.remove('collapsed');
    $('#btn-toggle-drawer').classList.add('active');
  });
  $('#welcome-browse').addEventListener('click', () => {
    drawerEl.classList.remove('collapsed');
    $('#btn-toggle-drawer').classList.add('active');
  });

  $('#btn-toggle-right').addEventListener('click', () => {
    rightDrawerEl.classList.toggle('collapsed');
    $('#btn-toggle-right').classList.toggle('active', !rightDrawerEl.classList.contains('collapsed'));
    if (!rightDrawerEl.classList.contains('collapsed')) void refreshRightWorktreePanel();
  });

  $('#btn-toggle-fit').addEventListener('click', () => {
    fitModeFlag = !fitModeFlag;
    $('#btn-toggle-fit').classList.toggle('active', fitModeFlag);
    activeDocxCtrl?.setZoom(fitModeFlag ? 'fit' : 1);
    showToast(fitModeFlag ? '已切换到适应侧栏宽度' : '已切换到 100% 原始宽度');
  });

  $('#btn-refresh').addEventListener('click', () => {
    const tab = activeTab();
    if (tab) void activateTab(tab.id);
    else if (currentDir) void navigateDir(currentDir);
  });

  $('#btn-plus').addEventListener('click', (e) => {
    e.stopPropagation();
    const wasHidden = plusMenuEl.classList.contains('hidden');
    closePopovers();
    if (wasHidden) positionPopover(plusMenuEl, $('#btn-plus'));
  });

  $('#btn-mode-menu').addEventListener('click', (e) => {
    e.stopPropagation();
    const wasHidden = modeMenuEl.classList.contains('hidden');
    closePopovers();
    if (wasHidden) positionPopover(modeMenuEl, $('#btn-mode-menu'));
  });

  document.addEventListener('click', closePopovers);

  plusMenuEl.querySelectorAll<HTMLButtonElement>('[data-plus]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.plus;
      if (action === 'browse') {
        drawerEl.classList.remove('collapsed');
        $('#btn-toggle-drawer').classList.add('active');
      } else if (action === 'new-docx') void createOfficeDocument('docx');
      else if (action === 'new-xlsx') void createOfficeDocument('xlsx');
      else if (action === 'new-pptx') void createOfficeDocument('pptx');
    });
  });

  $('#welcome-new-docx').addEventListener('click', () => void createOfficeDocument('docx'));
  $('#welcome-new-xlsx').addEventListener('click', () => void createOfficeDocument('xlsx'));
  $('#welcome-new-pptx').addEventListener('click', () => void createOfficeDocument('pptx'));

  modeMenuEl.querySelectorAll<HTMLButtonElement>('[data-mode]').forEach((btn) => {
    btn.addEventListener('click', () => void switchActiveMode(btn.dataset.mode as OfficeViewMode));
  });

  modeMenuEl.querySelectorAll<HTMLButtonElement>('[data-zoom]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const z = btn.dataset.zoom!;
      const zoomVal: 'fit' | number = z === 'fit' ? 'fit' : Number(z);
      fitModeFlag = zoomVal === 'fit';
      activeDocxCtrl?.setZoom(zoomVal);
    });
  });

  $('#menu-bookmark').addEventListener('click', () => {
    const tab = activeTab();
    if (!tab?.path) return;
    if (isBookmarked(tab.path)) {
      state.bookmarks = state.bookmarks.filter((b) => b.path.toLowerCase() !== tab.path!.toLowerCase());
      showToast('已取消收藏');
    } else {
      state.bookmarks = [{ path: tab.path, name: tab.name, kind: tab.kind, addedAt: Date.now() }, ...state.bookmarks];
      showToast('已加入收藏');
    }
    renderHistory();
    scheduleSaveState({ bookmarks: state.bookmarks });
  });

  $('#menu-native-tab').addEventListener('click', async () => {
    const tab = activeTab();
    if (!tab?.path) return;
    const ok = await mcp.openInNativeFileView(tab.path);
    if (!ok) showToast('当前视图不支持打开原生文件标签页');
  });

  $('#wt-compare-btn').addEventListener('click', () => void switchActiveMode('compare'));
  $('#wt-approve-btn').addEventListener('click', () => void handleWorktreeApprove());
  $('#wt-discard-btn').addEventListener('click', () => void handleWorktreeDiscard());

  document.querySelectorAll<HTMLButtonElement>('.drawer-tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.drawer!;
      document.querySelectorAll('.drawer-tab').forEach((b) => b.classList.toggle('active', b === btn));
      document.querySelectorAll('.drawer-pane').forEach((p) => {
        if (p.id.startsWith('drawer-')) p.classList.toggle('active', p.id === `drawer-${target}`);
      });
    });
  });

  document.querySelectorAll<HTMLButtonElement>('.right-tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.right!;
      document.querySelectorAll('.right-tab').forEach((b) => b.classList.toggle('active', b === btn));
      document.querySelectorAll('#right-drawer .drawer-pane').forEach((p) => {
        p.classList.toggle('active', p.id === `right-${target}`);
      });
    });
  });

  document.querySelectorAll<HTMLButtonElement>('#kind-chips .chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      activeFilter = chip.dataset.filter || 'all';
      document.querySelectorAll('#kind-chips .chip').forEach((c) => c.classList.toggle('active', c === chip));
      if (activeFilter === 'all' && !searchQuery.trim() && currentDir) void navigateDir(currentDir);
      else void runRecursiveSearch();
    });
  });

  let searchDebounce: number | undefined;
  $<HTMLInputElement>('#file-search-input').addEventListener('input', (e) => {
    searchQuery = (e.target as HTMLInputElement).value;
    window.clearTimeout(searchDebounce);
    searchDebounce = window.setTimeout(() => {
      if (!searchQuery.trim() && activeFilter === 'all' && currentDir) void navigateDir(currentDir);
      else void runRecursiveSearch();
    }, 260);
  });

  pathInputEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      const val = pathInputEl.value.trim();
      if (val) void openDiskFile(val);
    }
  });

  $('#btn-attach-ctx').addEventListener('click', async () => {
    try {
      const ctx = await buildContextPayload();
      if (!ctx) return;
      await mcp.attachContext(ctx);
      showToast('已挂载为下一次对话上下文');
    } catch (err) {
      showToast(err instanceof Error ? err.message : String(err));
    }
  });

  $('#btn-ask-agent').addEventListener('click', async () => {
    const tab = activeTab();
    if (!tab) return;
    try {
      const ctx = await buildContextPayload();
      await mcp.sendMessage(`请总结并分析我当前正在查看的文档（${tab.name}）。`, ctx);
    } catch (err) {
      showToast(err instanceof Error ? err.message : String(err));
    }
  });

  $('#btn-ask-review').addEventListener('click', async () => {
    const tab = activeTab();
    if (!tab) return;
    try {
      const ctx = await buildContextPayload();
      await mcp.sendMessage(
        `请检查当前文档（${tab.name}）的排版结构、内容完整性和错别字，并给出可直接通过 office_edit 工具修改的建议。`,
        ctx,
      );
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
      await mcp.sendMessage(`请解释或润色我在 ${tab?.name || '文档'} 中选中的这段内容：`, ctx);
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
      }
    },
    onToolCancelled(reason) {
      showToast(reason ? `已取消：${reason}` : '操作已取消');
    },
  });

  updateCapabilitiesUi();

  window.setTimeout(async () => {
    if (handledInitialToolResult) return;
    try {
      const stRes = await mcp.callTool<{ state: ViewerState }>('app_get_state');
      state = stRes.state;
      renderHistory();
      await navigateDir(state.lastRootDir || '.');
      if (!activeTabId) {
        drawerEl.classList.remove('collapsed');
        $('#btn-toggle-drawer').classList.add('active');
      }
    } catch {
      // Ignore standalone preview outside MCP host
    }
  }, 250);
}

void bootstrap();
