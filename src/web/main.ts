import './styles.css';
import { extractDroppedPayload, normalizePathCandidate } from './dropParser';
import { ViewerMcpClient } from './mcpClient';
import { mountPdfViewer, type OutlineNode, type PdfController } from './viewers/pdfViewer';
import { mountDocxOfficeViewer, type DocxController, type OfficeViewMode } from './viewers/docxOfficeViewer';
import { mountSheetOfficeViewer, type SheetController } from './viewers/sheetOfficeViewer';
import { mountSlideOfficeViewer, type SlideController } from './viewers/slideOfficeViewer';
import {
  mountHtmlView,
  mountImageView,
  mountLegacyDocView,
  mountRichDocView,
  mountTextView,
  type LegacyDocController,
  type SimpleViewerController,
} from './viewers/docViewers';

interface OpenWithApp {
  id: string;
  name: string;
  exePath: string;
  isDefault: boolean;
  iconDataUri?: string;
}

interface OpenWithResult {
  path: string;
  ext: string;
  platform: string;
  defaultApp: OpenWithApp | null;
  apps: OpenWithApp[];
}

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
  isTempDrop?: boolean;
}

const FILTER_EXT_MAP: Record<string, string[] | undefined> = {
  all: undefined,
  office: [
    '.docx', '.doc', '.wps', '.dot', '.rtf', '.odt',
    '.xlsx', '.xls', '.et', '.xlsm', '.xlsb', '.ods', '.csv', '.tsv',
    '.pptx', '.ppt', '.dps',
  ],
  pdf: ['.pdf'],
  docs: ['.md', '.markdown', '.mdx', '.txt'],
  html: ['.html', '.htm', '.xhtml'],
};

const mcp = new ViewerMcpClient();
const tabs: OpenTab[] = [];
const openWithCache = new Map<string, OpenWithResult>();
let activeTabId: string | null = null;
let currentDir = '';
let quickRoots: Array<{ label: string; path: string }> = [];
let currentEntries: FileEntry[] = [];
let activeFilter = 'all';
let searchQuery = '';
let state: ViewerState = { recentFiles: [], bookmarks: [], filePrefs: {} };
let selectedSnippet = '';
let currentZoom: 'fit' | number = 'fit';
let floatPageCurrent = 1;
let floatPageTotal = 1;
let floatIsDraft = false;
let floatHideTimer: number | undefined;
let floatVisible = false;
let floatHideScheduled = false;

let activePdfCtrl: PdfController | null = null;
let activeDocxCtrl: DocxController | null = null;
let activeLegacyDocCtrl: LegacyDocController | null = null;
let activeSheetCtrl: SheetController | null = null;
let activeSlideCtrl: SlideController | null = null;
let activeSimpleCtrl: SimpleViewerController | null = null;

const $ = <T extends HTMLElement>(sel: string) => document.querySelector<T>(sel)!;

const drawerEl = $('#drawer');
const rightDrawerEl = $('#right-drawer');
const drawerBackdropEl = $('#drawer-backdrop');
const tabsStripEl = $('#tabs-strip');
const pathBoxEl = $('#path-box');
const pathDisplayEl = $('#path-display');
const pathInputEl = $<HTMLInputElement>('#path-input');
const openDefaultBtnEl = $('#btn-open-default');
const openDefaultIconEl = $<HTMLImageElement>('#open-default-icon');
const modeAppBadgeEl = $('#mode-app-badge');
const openWithMenuEl = $('#open-with-menu');
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
const viewerStageEl = $('#viewer-stage');
const welcomeEl = $('#viewer-welcome');
const viewerContentEl = $('#viewer-content');
const stageFloatBarEl = $('#stage-float-bar');
const floatInfoEl = $('#float-info');
const floatZoomGroupEl = $('#float-zoom-group');
const floatZoomLabelEl = $('#float-zoom-label');
const floatSepZoomEl = $('#float-sep-zoom');
const floatModeGroupEl = $('#float-mode-group');
const floatSepModeEl = $('#float-sep-mode');
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
  if (kind === 'docx' || kind === 'legacy_doc') return { cls: 'w', label: 'W' };
  if (kind === 'spreadsheet') return { cls: 'x', label: 'X' };
  if (kind === 'slides' || kind === 'legacy_ppt') return { cls: 'p', label: 'P' };
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
  if (kind === 'legacy_doc') {
    if (ext === '.rtf') return 'RTF';
    if (ext === '.odt') return 'ODT';
    if (ext === '.wps') return 'WPS';
    return 'DOC';
  }
  if (kind === 'spreadsheet') return ext === '.csv' ? 'CSV' : ext === '.et' ? 'ET' : 'XLS';
  if (kind === 'slides') return 'PPTX';
  if (kind === 'legacy_ppt') return ext === '.dps' ? 'DPS' : 'PPT';
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

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function renderPathDisplay(rawPath?: string) {
  const val = (rawPath || '').trim();
  pathInputEl.value = val;
  if (!val) {
    pathDisplayEl.innerHTML = `<span class="addr-placeholder">输入文件路径（支持 .docx / .xlsx / .csv / .pptx / .pdf / .md / .html）按 Enter 打开…</span>`;
    return;
  }
  const idx = Math.max(val.lastIndexOf('\\'), val.lastIndexOf('/'));
  if (idx >= 0) {
    const dir = val.slice(0, idx + 1);
    const file = val.slice(idx + 1);
    pathDisplayEl.innerHTML = `<span class="addr-dir">${escapeHtml(dir)}</span><span class="addr-file">${escapeHtml(file)}</span>`;
  } else {
    pathDisplayEl.innerHTML = `<span class="addr-file">${escapeHtml(val)}</span>`;
  }
}

function isNarrowViewport(): boolean {
  return window.innerWidth <= 680;
}

function syncDrawerBackdrop() {
  const leftOpen = !drawerEl.classList.contains('collapsed');
  const rightOpen = !rightDrawerEl.classList.contains('collapsed');
  $('#btn-toggle-drawer').classList.toggle('active', leftOpen);
  $('#btn-toggle-right').classList.toggle('active', rightOpen);
  drawerBackdropEl.classList.toggle('hidden', !(isNarrowViewport() && (leftOpen || rightOpen)));
}

function selectLeftDrawerTab(target: 'files' | 'outline' | 'history') {
  document.querySelectorAll<HTMLButtonElement>('.drawer-tab').forEach((b) => {
    b.classList.toggle('active', b.dataset.drawer === target);
  });
  document.querySelectorAll('.drawer-pane').forEach((p) => {
    if (p.id.startsWith('drawer-')) p.classList.toggle('active', p.id === `drawer-${target}`);
  });
}

function setLeftDrawerOpen(open: boolean) {
  drawerEl.classList.toggle('collapsed', !open);
  if (open && isNarrowViewport()) {
    rightDrawerEl.classList.add('collapsed');
  }
  syncDrawerBackdrop();
}

function setRightDrawerOpen(open: boolean) {
  rightDrawerEl.classList.toggle('collapsed', !open);
  if (open && isNarrowViewport()) {
    drawerEl.classList.add('collapsed');
  }
  if (open) void refreshRightWorktreePanel();
  syncDrawerBackdrop();
}

function syncModeMenuChecks() {
  const tab = activeTab();
  const mode = tab?.mode ?? 'print';
  modeMenuEl.querySelectorAll<HTMLButtonElement>('[data-mode]').forEach((btn) => {
    btn.classList.toggle('selected', btn.dataset.mode === mode);
  });
  modeMenuEl.querySelectorAll<HTMLButtonElement>('[data-zoom]').forEach((btn) => {
    const z = btn.dataset.zoom!;
    const match =
      (z === 'fit' && currentZoom === 'fit') ||
      (z !== 'fit' && typeof currentZoom === 'number' && Math.abs(currentZoom - Number(z)) < 0.03);
    btn.classList.toggle('selected', match);
  });
  $('#btn-toggle-fit').classList.toggle('active', currentZoom === 'fit');
}

function syncFloatingCapsule() {
  const tab = activeTab();
  const supportsFloat =
    Boolean(tab) &&
    (tab!.kind === 'docx' ||
      tab!.kind === 'legacy_doc' ||
      tab!.kind === 'slides' ||
      tab!.kind === 'legacy_ppt' ||
      tab!.kind === 'pdf');

  if (!supportsFloat || !tab) {
    window.clearTimeout(floatHideTimer);
    floatVisible = false;
    floatHideScheduled = false;
    stageFloatBarEl.classList.add('hidden');
    stageFloatBarEl.classList.remove('visible');
    syncModeMenuChecks();
    return;
  }

  stageFloatBarEl.classList.remove('hidden');

  if (tab.kind === 'slides' || tab.kind === 'legacy_ppt') {
    floatInfoEl.textContent = `幻灯片 ${floatPageCurrent} / ${Math.max(1, floatPageTotal)}`;
  } else {
    const draftSuffix = tab.kind === 'docx' && floatIsDraft ? ' · 草稿' : '';
    floatInfoEl.textContent = `第 ${floatPageCurrent} / ${Math.max(1, floatPageTotal)} 页${draftSuffix}`;
  }

  const showZoom =
    (tab.kind === 'docx' && tab.mode === 'print') ||
    tab.kind === 'legacy_doc' ||
    tab.kind === 'pdf';
  floatZoomGroupEl.classList.toggle('hidden', !showZoom);
  floatSepZoomEl.classList.toggle('hidden', !showZoom);
  floatZoomLabelEl.textContent =
    currentZoom === 'fit' ? '适应' : `${Math.round(Number(currentZoom) * 100)}%`;

  const showModes = tab.kind === 'docx' || tab.kind === 'slides';
  floatModeGroupEl.classList.toggle('hidden', !showModes);
  floatSepModeEl.classList.toggle('hidden', !showModes);

  floatModeGroupEl.querySelectorAll<HTMLButtonElement>('[data-float-mode]').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.floatMode === tab.mode);
  });

  syncModeMenuChecks();
}

function applyZoomToActiveViewer(zoom: 'fit' | number) {
  currentZoom = zoom;
  if (activeDocxCtrl) activeDocxCtrl.setZoom(zoom);
  if (activeLegacyDocCtrl) activeLegacyDocCtrl.setZoom(zoom);
  if (activePdfCtrl) activePdfCtrl.setZoom(zoom);
  syncFloatingCapsule();
}

async function fetchOpenWithInfo(filePath: string, ext: string): Promise<OpenWithResult> {
  const key = ext.toLowerCase() || '<none>';
  const cached = openWithCache.get(key);
  if (cached) {
    return { ...cached, path: filePath };
  }
  const res = await mcp.callTool<OpenWithResult>('app_list_open_with', { path: filePath });
  openWithCache.set(key, res);
  return res;
}

function setFallbackOpenDefaultBadge(tab?: OpenTab) {
  const b = tab ? badgeForKind(tab.kind, tab.ext) : { cls: 'w', label: 'W' };
  openDefaultIconEl.classList.add('hidden');
  openDefaultIconEl.removeAttribute('src');
  modeAppBadgeEl.className = `app-badge ${b.cls}`;
  modeAppBadgeEl.textContent = b.label;
  modeAppBadgeEl.classList.remove('hidden');
  openDefaultBtnEl.title = tab?.path ? '使用系统默认程序打开当前文件' : '使用系统默认程序打开（请先打开磁盘文件）';
}

async function refreshOpenWithForTab(tab?: OpenTab) {
  setFallbackOpenDefaultBadge(tab);
  if (!tab?.path) return;
  try {
    const res = await fetchOpenWithInfo(tab.path, tab.ext);
    if (activeTab()?.id !== tab.id) return;
    if (res.defaultApp?.iconDataUri) {
      openDefaultIconEl.src = res.defaultApp.iconDataUri;
      openDefaultIconEl.classList.remove('hidden');
      modeAppBadgeEl.classList.add('hidden');
    }
    if (res.defaultApp?.name) {
      openDefaultBtnEl.title = `使用 ${res.defaultApp.name}（默认）打开`;
    }
  } catch {
    // Keep clean fallback badge if app discovery is unavailable
  }
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
  if (!tab?.path || openDefaultIconEl.classList.contains('hidden')) {
    const b = tab ? badgeForKind(tab.kind, tab.ext) : { cls: 'w', label: 'W' };
    modeAppBadgeEl.className = `app-badge ${b.cls}`;
    modeAppBadgeEl.textContent = b.label;
  }

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
      else if (item.id && activeLegacyDocCtrl) activeLegacyDocCtrl.scrollToHeading(item.id);
      else if (item.id && activeSimpleCtrl?.scrollToHeading) activeSimpleCtrl.scrollToHeading(item.id);
      if (isNarrowViewport()) setLeftDrawerOpen(false);
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
  if (tab.isTempDrop) {
    worktreePanelBodyEl.innerHTML = `<div class="empty-hint">当前为本地拖入的临时只读预览。如需 Worktree 编辑或版本管理，请通过文件夹按钮或路径栏打开磁盘原文件。</div>`;
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
        if (isNarrowViewport()) setLeftDrawerOpen(false);
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
  activeLegacyDocCtrl?.destroy();
  activeLegacyDocCtrl = null;
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
  syncFloatingCapsule();
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
  floatPageCurrent = 1;
  floatPageTotal = 1;
  floatIsDraft = tab.worktree?.status === 'draft';
  renderTabs();
  renderFileList();
  renderPathDisplay(tab.path || tab.name);
  void refreshOpenWithForTab(tab);

  welcomeEl.classList.add('hidden');
  viewerContentEl.classList.remove('hidden');
  unmountActiveViewer();
  syncFloatingCapsule();

  const prefs = tab.path ? state.filePrefs[tab.path] : undefined;

  try {
    if (tab.kind === 'docx') {
      activeDocxCtrl = await mountDocxOfficeViewer({
        container: viewerContentEl,
        mcp,
        filePath: tab.path,
        rawBuffer: tab.rawBuffer,
        mode: tab.mode,
        zoomSetting: currentZoom,
        onOutline: renderOutline,
        onPageInfo: (info) => {
          floatPageCurrent = info.currentPage;
          floatPageTotal = info.totalPages;
          floatIsDraft = info.isDraft;
          syncFloatingCapsule();
        },
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
          onSlideChange: (current, total) => {
            tab.page = current;
            floatPageCurrent = current;
            floatPageTotal = total;
            syncFloatingCapsule();
          },
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
        onPageChange: (page, totalPages) => {
          tab.page = page;
          floatPageCurrent = page;
          floatPageTotal = totalPages;
          syncFloatingCapsule();
          if (tab.path) {
            state.filePrefs[tab.path] = { ...state.filePrefs[tab.path], page };
            scheduleSaveState({ filePrefs: { [tab.path]: state.filePrefs[tab.path] } });
          }
        },
        onZoomChange: (zoom, fitMode) => {
          currentZoom = fitMode === 'width' ? 'fit' : Number(zoom.toFixed(2));
          syncFloatingCapsule();
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
      if (!tab.path) {
        viewerContentEl.innerHTML = `<div class="error-box">请通过文件路径或 📁 文件夹打开 <code>${escapeHtml(tab.ext || '.doc')}</code> 文档，以启用完整离线解析。</div>`;
      } else {
        const legacyCtrl = await mountLegacyDocView({
          container: viewerContentEl,
          mcp,
          filePath: tab.path,
          zoomSetting: currentZoom,
          onOutline: renderOutline,
          onPageInfo: (info) => {
            floatPageCurrent = info.currentPage;
            floatPageTotal = info.totalPages;
            syncFloatingCapsule();
          },
        });
        if (legacyCtrl.delegatedToDocx) {
          tab.kind = 'docx';
          renderTabs();
          activeDocxCtrl = await mountDocxOfficeViewer({
            container: viewerContentEl,
            mcp,
            filePath: tab.path,
            mode: tab.mode,
            zoomSetting: currentZoom,
            onOutline: renderOutline,
            onPageInfo: (info) => {
              floatPageCurrent = info.currentPage;
              floatPageTotal = info.totalPages;
              floatIsDraft = info.isDraft;
              syncFloatingCapsule();
            },
            onWorktreeChanged: () => void refreshRightWorktreePanel(),
            onToast: showToast,
          });
        } else {
          activeLegacyDocCtrl = legacyCtrl;
        }
      }
    } else if (tab.kind === 'legacy_ppt') {
      if (tab.path) {
        activeSlideCtrl = await mountSlideOfficeViewer({
          container: viewerContentEl,
          mcp,
          filePath: tab.path,
          initialSlide: tab.page ?? 1,
          mode: 'print',
          legacy: true,
          onOutline: renderOutline,
          onSlideChange: (current, total) => {
            tab.page = current;
            floatPageCurrent = current;
            floatPageTotal = total;
            syncFloatingCapsule();
          },
          onWorktreeChanged: () => void refreshRightWorktreePanel(),
          onToast: showToast,
        });
      } else {
        viewerContentEl.innerHTML = `<div class="error-box">请通过文件路径或目录打开旧版 PowerPoint (.ppt / .dps) 文件。</div>`;
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
      renderPathDisplay('');
      renderTabs();
      void refreshOpenWithForTab(undefined);
      syncFloatingCapsule();
      void refreshRightWorktreePanel();
    } else {
      const next = tabs[Math.max(0, idx - 1)];
      void activateTab(next.id);
    }
  } else {
    renderTabs();
  }
}

async function openDiskFile(
  filePath: string,
  opts: {
    page?: number;
    sheet?: string;
    recordRecent?: boolean;
    isTempDrop?: boolean;
    displayName?: string;
  } = {},
): Promise<boolean> {
  try {
    const res = await mcp.callTool<{
      file: { path: string; name: string; dir: string; ext: string; kind: string; size: number };
      worktree?: WorktreeMeta | null;
      state: ViewerState;
    }>('app_open_file_meta', { path: filePath, recordRecent: opts.recordRecent ?? true });

    state = res.state;
    renderHistory();

    const existing = tabs.find((t) => t.path && t.path.toLowerCase() === res.file.path.toLowerCase());
    if (existing) {
      if (opts.page) existing.page = opts.page;
      if (opts.sheet) existing.sheet = opts.sheet;
      if (opts.displayName) existing.name = opts.displayName;
      if (opts.isTempDrop !== undefined) existing.isTempDrop = opts.isTempDrop;
      existing.worktree = existing.isTempDrop ? null : res.worktree ?? null;
      await activateTab(existing.id);
      if (opts.page && activePdfCtrl) activePdfCtrl.goToPage(opts.page);
      if (opts.page && activeSlideCtrl) activeSlideCtrl.goToSlide(opts.page);
      return true;
    }

    const tab: OpenTab = {
      id: `disk:${res.file.path}`,
      path: res.file.path,
      name: opts.displayName || res.file.name,
      kind: res.file.kind,
      ext: res.file.ext,
      size: res.file.size,
      page: opts.page,
      sheet: opts.sheet,
      mode: 'print',
      worktree: opts.isTempDrop ? null : res.worktree ?? null,
      isTempDrop: opts.isTempDrop,
    };
    tabs.push(tab);
    await activateTab(tab.id);
    return true;
  } catch (err) {
    showToast(`打开失败：${err instanceof Error ? err.message : String(err)}`);
    return false;
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

const SERVER_PARSED_DROP_EXTS = new Set([
  '.doc',
  '.wps',
  '.dot',
  '.rtf',
  '.odt',
  '.xlsx',
  '.xls',
  '.et',
  '.xlsm',
  '.xlsb',
  '.ods',
  '.csv',
  '.tsv',
  '.pptx',
  '.ppt',
  '.dps',
  '.md',
  '.markdown',
  '.mdx',
]);

const MAX_DROPPED_STAGE_BYTES = 50 * 1024 * 1024;
const DROP_STAGE_CHUNK_BYTES = 384 * 1024;

function uint8ToBase64(bytes: Uint8Array): string {
  const SLICE = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += SLICE) {
    const sub = bytes.subarray(i, Math.min(i + SLICE, bytes.length));
    binary += String.fromCharCode(...sub);
  }
  return btoa(binary);
}

async function openLocalDroppedFile(file: File) {
  const nativePath = (file as File & { path?: string }).path;
  if (nativePath) {
    const opened = await openDiskFile(nativePath);
    if (opened) return;
  }

  const name = file.name;
  const ext = name.includes('.') ? name.slice(name.lastIndexOf('.')).toLowerCase() : '';

  if (SERVER_PARSED_DROP_EXTS.has(ext)) {
    if (file.size > MAX_DROPPED_STAGE_BYTES) {
      showToast(`${name} 超过 50 MB 临时解析限制，请通过文件夹或路径栏打开`);
      return;
    }
    showToast(`正在解析拖入文件：${name}…`);
    try {
      const rawBuffer = await file.arrayBuffer();
      const total = rawBuffer.byteLength;
      const uploadId = `drop_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
      let stagedPath = '';
      let offset = 0;

      while (true) {
        const end = total === 0 ? 0 : Math.min(offset + DROP_STAGE_CHUNK_BYTES, total);
        const isFinal = end >= total;
        const slice = new Uint8Array(rawBuffer, offset, end - offset);
        const res = await mcp.callTool<{
          staged?: boolean;
          file?: { path: string };
          isTemporary?: boolean;
        }>('app_stage_dropped_file', {
          uploadId,
          name,
          chunkBase64: uint8ToBase64(slice),
          offset,
          final: isFinal,
        });
        if (isFinal) {
          stagedPath = res.file?.path || '';
          break;
        }
        offset = end;
      }

      if (!stagedPath) {
        throw new Error('服务端未返回临时文件路径');
      }
      await openDiskFile(stagedPath, {
        recordRecent: false,
        isTempDrop: true,
        displayName: name,
      });
      return;
    } catch (err) {
      showToast(
        `暂无法离线解析拖入的 ${name}（${err instanceof Error ? err.message : String(err)}）。请在插件管理中重载 Doc Viewer 后重试，或通过文件夹 / 路径栏打开原文件`,
      );
      return;
    }
  }

  const kind =
    ext === '.docx'
      ? 'docx'
      : ext === '.pdf'
        ? 'pdf'
        : ext === '.html' || ext === '.htm' || ext === '.xhtml'
          ? 'html'
          : ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.bmp', '.ico'].includes(ext)
            ? 'image'
            : 'text';

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

type ResolveDropResult =
  | { action: 'open_files'; paths: string[] }
  | { action: 'open_dir'; path: string }
  | { action: 'pick_matches'; query: string; matches: FileEntry[] }
  | { action: 'not_found'; query: string };

async function handleIncomingTransfer(dt: DataTransfer | null, source: 'drop' | 'paste'): Promise<boolean> {
  const payload = extractDroppedPayload(dt);

  if (payload.candidates.length > 0 || payload.fallbackNames.length > 0) {
    try {
      const res = await mcp.callTool<ResolveDropResult>('app_resolve_dropped_target', {
        candidates: payload.candidates.slice(0, 20),
        fallbackNames: payload.fallbackNames.slice(0, 10),
        currentDir: currentDir || undefined,
      });

      if (res.action === 'open_files' && Array.isArray(res.paths) && res.paths.length > 0) {
        let openedAny = false;
        for (const p of res.paths.slice(0, 5)) {
          if (await openDiskFile(p)) openedAny = true;
        }
        return openedAny;
      }

      if (res.action === 'open_dir' && res.path) {
        await navigateDir(res.path);
        selectLeftDrawerTab('files');
        setLeftDrawerOpen(true);
        showToast(`已定位到目录：${res.path}`);
        return true;
      }

      if (res.action === 'pick_matches' && Array.isArray(res.matches) && res.matches.length > 0) {
        searchQuery = res.query;
        $<HTMLInputElement>('#file-search-input').value = res.query;
        currentEntries = res.matches;
        renderFileList();
        selectLeftDrawerTab('files');
        setLeftDrawerOpen(true);
        showToast(`找到 ${res.matches.length} 个匹配文件，请在左侧列表选择`);
        return true;
      }

      if (res.action === 'not_found') {
        if (payload.files.length === 0) {
          showToast(`未找到本地文件：${res.query}（若拖拽对话胶囊无路径，可复制路径后按 Ctrl+V 打开）`);
          return false;
        }
      }
    } catch {
      // Fallback for older running server binary before reload
      for (const rawCandidate of payload.candidates.slice(0, 8)) {
        const normalized = normalizePathCandidate(rawCandidate) || rawCandidate;
        try {
          const meta = await mcp.callTool<{ file: { path: string } }>('app_open_file_meta', {
            path: normalized,
            recordRecent: false,
          });
          if (meta.file?.path) {
            await openDiskFile(meta.file.path);
            return true;
          }
        } catch {
          // Try next candidate
        }
      }

      if (payload.fallbackNames.length > 0) {
        const qLower = payload.fallbackNames[0].toLowerCase();
        const historyHit = [...state.recentFiles, ...state.bookmarks].find(
          (item) =>
            item.name.toLowerCase() === qLower ||
            item.name.toLowerCase().includes(qLower),
        );
        if (historyHit) {
          await openDiskFile(historyHit.path);
          return true;
        }
        if (currentDir) {
          try {
            const searchRes = await mcp.callTool<{ items: FileEntry[] }>('app_search_files', {
              rootDir: currentDir,
              query: payload.fallbackNames[0],
              maxDepth: 5,
              limit: 25,
            });
            if (searchRes.items.length === 1) {
              await openDiskFile(searchRes.items[0].path);
              return true;
            }
            if (searchRes.items.length > 1) {
              searchQuery = payload.fallbackNames[0];
              $<HTMLInputElement>('#file-search-input').value = searchQuery;
              currentEntries = searchRes.items;
              renderFileList();
              selectLeftDrawerTab('files');
              setLeftDrawerOpen(true);
              showToast(`找到 ${searchRes.items.length} 个匹配文件，请在左侧列表选择`);
              return true;
            }
          } catch {
            // Ignore search fallback error
          }
        }
      }
    }
  }

  if (payload.files.length > 0) {
    for (const f of payload.files.slice(0, 5)) {
      await openLocalDroppedFile(f);
    }
    return true;
  }

  if (source === 'drop') {
    if (payload.ignoredWebUrls.length > 0) {
      showToast('Doc Viewer 仅支持本地文档，暂不支持打开网页链接');
    } else {
      showToast('未能识别拖入的文件路径（可拖动对话里的文件链接，或按 Ctrl+V 粘贴路径打开）');
    }
  }
  return false;
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
  if (activeLegacyDocCtrl) {
    return `[Legacy Office Document]\nFile: ${tab.path || tab.name}\n\n${activeLegacyDocCtrl.getContextSummary()}`;
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

function positionPopover(menu: HTMLElement, anchor: HTMLElement, alignRight = false) {
  const rect = anchor.getBoundingClientRect();
  menu.style.top = `${rect.bottom + 4}px`;
  menu.classList.remove('hidden');
  const menuW = menu.offsetWidth || 220;
  const rawLeft = alignRight ? rect.right - menuW : rect.left;
  const maxLeft = Math.max(8, window.innerWidth - menuW - 8);
  menu.style.left = `${Math.max(8, Math.min(rawLeft, maxLeft))}px`;
}

function closePopovers() {
  plusMenuEl.classList.add('hidden');
  modeMenuEl.classList.add('hidden');
  openWithMenuEl.classList.add('hidden');
}

async function runExternalAction(
  filePath: string,
  action: 'default' | 'app' | 'choose' | 'reveal',
  exePath?: string,
) {
  try {
    const res = await mcp.callTool<{ ok: boolean; message: string }>('app_open_external', {
      path: filePath,
      action,
      exePath,
    });
    if (res.message) showToast(res.message);
  } catch (err) {
    showToast(`打开失败：${err instanceof Error ? err.message : String(err)}`);
  }
}

async function openWithMenuShow() {
  const tab = activeTab();
  if (!tab?.path) {
    showToast('请先打开磁盘上的文件，再选择外部打开方式');
    return;
  }

  openWithMenuEl.innerHTML = `<div class="empty-hint" style="padding:10px 12px">正在读取系统打开方式…</div>`;
  positionPopover(openWithMenuEl, $('#open-with-split'), true);

  let info: OpenWithResult = {
    path: tab.path,
    ext: tab.ext,
    platform: 'win32',
    defaultApp: null,
    apps: [],
  };
  try {
    info = await fetchOpenWithInfo(tab.path, tab.ext);
  } catch {
    // Fallback to minimal menu
  }

  openWithMenuEl.innerHTML = '';

  const addActionRow = (
    label: string,
    opts: { iconDataUri?: string; glyph?: string; isDefault?: boolean; onClick: () => void },
  ) => {
    const btn = document.createElement('button');
    btn.className = 'menu-item';
    if (opts.iconDataUri) {
      const img = document.createElement('img');
      img.className = 'open-with-item-icon';
      img.src = opts.iconDataUri;
      img.alt = '';
      btn.appendChild(img);
    } else {
      const sp = document.createElement('span');
      sp.className = 'open-with-item-glyph';
      sp.textContent = opts.glyph || '↗';
      btn.appendChild(sp);
    }
    const textWrap = document.createElement('span');
    textWrap.className = 'open-with-item-label';
    textWrap.textContent = label;
    if (opts.isDefault) {
      const tag = document.createElement('span');
      tag.className = 'open-with-default-tag';
      tag.textContent = '（默认）';
      textWrap.appendChild(tag);
    }
    btn.appendChild(textWrap);
    btn.addEventListener('click', () => {
      closePopovers();
      opts.onClick();
    });
    openWithMenuEl.appendChild(btn);
  };

  if (info.apps.length > 0) {
    for (const app of info.apps) {
      addActionRow(app.name, {
        iconDataUri: app.iconDataUri,
        glyph: '↗',
        isDefault: app.isDefault,
        onClick: () => void runExternalAction(tab.path!, 'app', app.exePath),
      });
    }
  } else {
    addActionRow('使用系统默认程序打开', {
      glyph: '↗',
      isDefault: true,
      onClick: () => void runExternalAction(tab.path!, 'default'),
    });
  }

  const sep1 = document.createElement('div');
  sep1.className = 'menu-sep';
  openWithMenuEl.appendChild(sep1);

  if (info.platform === 'win32') {
    addActionRow('选择其他应用…', {
      glyph: '⌘',
      onClick: () => void runExternalAction(tab.path!, 'choose'),
    });
  }

  addActionRow('在原生文件标签页打开', {
    glyph: '⧉',
    onClick: async () => {
      const ok = await mcp.openInNativeFileView(tab.path!);
      if (!ok) showToast('当前视图不支持打开原生文件标签页');
    },
  });

  const sep2 = document.createElement('div');
  sep2.className = 'menu-sep';
  openWithMenuEl.appendChild(sep2);

  addActionRow('显示文件位置', {
    glyph: '📁',
    onClick: () => void runExternalAction(tab.path!, 'reveal'),
  });

  positionPopover(openWithMenuEl, $('#open-with-split'), true);
}

function wireEvents() {
  const toggleLeftDrawer = () => {
    setLeftDrawerOpen(drawerEl.classList.contains('collapsed'));
  };
  $('#btn-toggle-drawer').addEventListener('click', toggleLeftDrawer);
  $('#btn-folder-browse').addEventListener('click', () => setLeftDrawerOpen(true));
  $('#welcome-browse').addEventListener('click', () => setLeftDrawerOpen(true));

  $('#btn-toggle-right').addEventListener('click', () => {
    setRightDrawerOpen(rightDrawerEl.classList.contains('collapsed'));
  });

  drawerBackdropEl.addEventListener('click', () => {
    setLeftDrawerOpen(false);
    setRightDrawerOpen(false);
  });

  window.addEventListener('resize', syncDrawerBackdrop);

  // Smart path display <-> editable path input
  pathBoxEl.addEventListener('click', () => {
    if (!pathInputEl.classList.contains('hidden')) return;
    pathDisplayEl.classList.add('hidden');
    pathInputEl.classList.remove('hidden');
    pathInputEl.focus();
    pathInputEl.select();
  });

  const exitPathEdit = () => {
    pathInputEl.classList.add('hidden');
    pathDisplayEl.classList.remove('hidden');
    const tab = activeTab();
    renderPathDisplay(tab?.path || tab?.name || '');
  };

  pathInputEl.addEventListener('blur', exitPathEdit);
  pathInputEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      const val = pathInputEl.value.trim();
      pathInputEl.blur();
      if (val) void openDiskFile(val);
    } else if (e.key === 'Escape') {
      exitPathEdit();
    }
  });

  $('#btn-toggle-fit').addEventListener('click', () => {
    const nextZoom: 'fit' | number = currentZoom === 'fit' ? 1 : 'fit';
    applyZoomToActiveViewer(nextZoom);
    showToast(nextZoom === 'fit' ? '已切换到适应侧栏宽度' : '已切换到 100% 原始宽度');
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
    if (wasHidden) {
      syncModeMenuChecks();
      positionPopover(modeMenuEl, $('#btn-mode-menu'), true);
    }
  });

  openDefaultBtnEl.addEventListener('click', () => {
    const tab = activeTab();
    if (!tab?.path) {
      showToast('请先打开磁盘上的文件再使用外部应用打开');
      return;
    }
    void runExternalAction(tab.path, 'default');
  });

  $('#btn-open-with').addEventListener('click', (e) => {
    e.stopPropagation();
    const wasHidden = openWithMenuEl.classList.contains('hidden');
    closePopovers();
    if (wasHidden) void openWithMenuShow();
  });

  document.addEventListener('click', closePopovers);

  plusMenuEl.querySelectorAll<HTMLButtonElement>('[data-plus]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.plus;
      if (action === 'browse') setLeftDrawerOpen(true);
      else if (action === 'new-docx') void createOfficeDocument('docx');
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
      applyZoomToActiveViewer(zoomVal);
    });
  });

  // Hover-revealed bottom floating capsule
  let pointerInsideFloat = false;
  let focusInsideFloat = false;

  const revealFloatBar = () => {
    if (stageFloatBarEl.classList.contains('hidden')) return;
    window.clearTimeout(floatHideTimer);
    floatHideScheduled = false;
    if (!floatVisible) {
      floatVisible = true;
      stageFloatBarEl.classList.add('visible');
    }
  };

  const scheduleHideFloatBar = () => {
    if (!floatVisible || pointerInsideFloat || focusInsideFloat || floatHideScheduled) return;
    floatHideScheduled = true;
    floatHideTimer = window.setTimeout(() => {
      floatHideScheduled = false;
      if (!pointerInsideFloat && !focusInsideFloat) {
        floatVisible = false;
        stageFloatBarEl.classList.remove('visible');
      }
    }, 240);
  };

  viewerStageEl.addEventListener(
    'mousemove',
    (e) => {
      if (stageFloatBarEl.classList.contains('hidden')) return;
      const inZone = window.innerHeight - e.clientY <= 72;
      if (inZone) {
        if (!floatVisible || floatHideScheduled) revealFloatBar();
      } else if (floatVisible && !floatHideScheduled) {
        scheduleHideFloatBar();
      }
    },
    { passive: true },
  );

  viewerStageEl.addEventListener('mouseleave', () => {
    pointerInsideFloat = false;
    scheduleHideFloatBar();
  });

  stageFloatBarEl.addEventListener('mouseenter', () => {
    pointerInsideFloat = true;
    revealFloatBar();
  });

  stageFloatBarEl.addEventListener('mouseleave', () => {
    pointerInsideFloat = false;
    scheduleHideFloatBar();
  });

  stageFloatBarEl.addEventListener('focusin', () => {
    focusInsideFloat = true;
    revealFloatBar();
  });

  stageFloatBarEl.addEventListener('focusout', () => {
    focusInsideFloat = false;
    scheduleHideFloatBar();
  });

  $('#float-prev').addEventListener('click', () => {
    if (activeDocxCtrl) activeDocxCtrl.goToPage(floatPageCurrent - 1);
    else if (activeLegacyDocCtrl) activeLegacyDocCtrl.goToPage(floatPageCurrent - 1);
    else if (activePdfCtrl) activePdfCtrl.goToPage(floatPageCurrent - 1);
    else if (activeSlideCtrl) activeSlideCtrl.goToSlide(floatPageCurrent - 1);
  });

  $('#float-next').addEventListener('click', () => {
    if (activeDocxCtrl) activeDocxCtrl.goToPage(floatPageCurrent + 1);
    else if (activeLegacyDocCtrl) activeLegacyDocCtrl.goToPage(floatPageCurrent + 1);
    else if (activePdfCtrl) activePdfCtrl.goToPage(floatPageCurrent + 1);
    else if (activeSlideCtrl) activeSlideCtrl.goToSlide(floatPageCurrent + 1);
  });

  $('#float-zoom-out').addEventListener('click', () => {
    const base = currentZoom === 'fit' ? 1 : currentZoom;
    applyZoomToActiveViewer(Math.max(0.5, Number((base - 0.15).toFixed(2))));
  });

  $('#float-zoom-in').addEventListener('click', () => {
    const base = currentZoom === 'fit' ? 1 : currentZoom;
    applyZoomToActiveViewer(Math.min(2.2, Number((base + 0.15).toFixed(2))));
  });

  floatZoomLabelEl.addEventListener('click', () => {
    applyZoomToActiveViewer(currentZoom === 'fit' ? 1 : 'fit');
  });

  floatModeGroupEl.querySelectorAll<HTMLButtonElement>('[data-float-mode]').forEach((btn) => {
    btn.addEventListener('click', () => void switchActiveMode(btn.dataset.floatMode as OfficeViewMode));
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

  $('#wt-compare-btn').addEventListener('click', () => void switchActiveMode('compare'));
  $('#wt-approve-btn').addEventListener('click', () => void handleWorktreeApprove());
  $('#wt-discard-btn').addEventListener('click', () => void handleWorktreeDiscard());

  document.querySelectorAll<HTMLButtonElement>('.drawer-tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.drawer as 'files' | 'outline' | 'history';
      selectLeftDrawerTab(target);
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

  let selectionTimer: number | undefined;
  document.addEventListener('selectionchange', () => {
    window.clearTimeout(selectionTimer);
    selectionTimer = window.setTimeout(() => {
      const sel = window.getSelection()?.toString().trim() || '';
      if (sel.length >= 4 && (mcp.canUpdateContext() || mcp.canMessage())) {
        selectedSnippet = sel.slice(0, 6000);
        selectionPreviewEl.textContent = `已选 ${sel.length} 字`;
        selectionBarEl.classList.remove('hidden');
      } else if (!sel && (selectedSnippet || !selectionBarEl.classList.contains('hidden'))) {
        selectedSnippet = '';
        selectionBarEl.classList.add('hidden');
      }
    }, 100);
  });

  $<HTMLInputElement>('#file-picker').addEventListener('change', (e) => {
    const files = (e.target as HTMLInputElement).files;
    if (files?.[0]) void openLocalDroppedFile(files[0]);
  });

  const dropOverlay = $('#drop-overlay');
  let dragDepth = 0;
  let internalDragActive = false;

  const hideDropOverlay = () => {
    dragDepth = 0;
    dropOverlay.classList.add('hidden');
  };

  document.addEventListener('dragstart', () => {
    internalDragActive = true;
  });

  document.addEventListener('dragend', () => {
    internalDragActive = false;
    hideDropOverlay();
  });

  window.addEventListener('blur', hideDropOverlay);

  document.addEventListener('dragenter', (e) => {
    if (internalDragActive) return;
    e.preventDefault();
    dragDepth += 1;
    dropOverlay.classList.remove('hidden');
  });

  document.addEventListener('dragover', (e) => {
    if (internalDragActive) return;
    e.preventDefault();
    if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
  });

  document.addEventListener('dragleave', (e) => {
    if (internalDragActive) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0 || e.relatedTarget === null) {
      hideDropOverlay();
    }
  });

  document.addEventListener('drop', (e) => {
    const wasInternal = internalDragActive;
    internalDragActive = false;
    hideDropOverlay();
    if (wasInternal) return;
    e.preventDefault();
    void handleIncomingTransfer(e.dataTransfer, 'drop');
  });

  document.addEventListener('paste', (e) => {
    const activeEl = document.activeElement as HTMLElement | null;
    if (
      activeEl &&
      (activeEl.tagName === 'INPUT' ||
        activeEl.tagName === 'TEXTAREA' ||
        activeEl.tagName === 'SELECT' ||
        activeEl.isContentEditable ||
        Boolean(activeEl.closest('[contenteditable="true"]')))
    ) {
      return;
    }
    const payload = extractDroppedPayload(e.clipboardData);
    if (payload.files.length > 0 || payload.candidates.length > 0 || payload.fallbackNames.length > 0) {
      e.preventDefault();
      void handleIncomingTransfer(e.clipboardData, 'paste');
    }
  });
}

async function bootstrap() {
  wireEvents();
  syncFloatingCapsule();
  let handledInitialToolResult = false;

  await mcp.connect({
    onToolInput(args) {
      if (typeof args.path === 'string' && args.path) {
        renderPathDisplay(args.path);
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
      if (!activeTabId && window.innerWidth > 760) {
        setLeftDrawerOpen(true);
      }
    } catch {
      // Ignore standalone preview outside MCP host
    }
  }, 250);
}

void bootstrap();
