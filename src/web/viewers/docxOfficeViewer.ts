import { renderAsync } from 'docx-preview';
import type { ViewerMcpClient } from '../mcpClient';
import { base64ToUint8Array, type OutlineNode } from './pdfViewer';

export type OfficeViewMode = 'print' | 'edit' | 'compare';

export interface DocxMountOptions {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  filePath?: string;
  rawBuffer?: ArrayBuffer;
  mode?: OfficeViewMode;
  zoomSetting?: 'fit' | number;
  onOutline?: (items: OutlineNode[]) => void;
  onWorktreeChanged?: () => void;
  onToast?: (msg: string) => void;
}

export interface DocxController {
  setMode: (mode: OfficeViewMode) => Promise<void>;
  setZoom: (zoom: 'fit' | number) => void;
  getContextSummary: () => string;
  scrollToHeading: (id: string) => void;
  destroy: () => void;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

async function loadDocxBytes(mcp: ViewerMcpClient, filePath: string): Promise<{ bytes: Uint8Array; isDraft: boolean }> {
  const chunkSize = 384 * 1024;
  const parts: Uint8Array[] = [];
  let offset = 0;
  let isDraft = false;

  while (true) {
    const chunk = await mcp.callTool<{
      base64: string;
      bytesRead: number;
      eof: boolean;
      isDraft?: boolean;
    }>('app_read_file_chunk', { path: filePath, offset, length: chunkSize, useDraft: true });
    if (chunk.isDraft) isDraft = true;
    if (chunk.bytesRead > 0) {
      parts.push(base64ToUint8Array(chunk.base64));
      offset += chunk.bytesRead;
    }
    if (chunk.eof || chunk.bytesRead === 0) break;
  }

  const combined = new Uint8Array(offset);
  let pos = 0;
  for (const p of parts) {
    combined.set(p, pos);
    pos += p.byteLength;
  }
  return { bytes: combined, isDraft };
}

function clonePageShell(templateSection: HTMLElement, targetHeightPx: number): {
  section: HTMLElement;
  article: HTMLElement;
} {
  const section = document.createElement('section');
  section.className = templateSection.className;
  section.style.cssText = templateSection.style.cssText;
  section.style.minHeight = `${targetHeightPx}px`;
  section.style.boxSizing = 'border-box';

  const article = document.createElement('article');
  const templateArticle = templateSection.querySelector<HTMLElement>(':scope > article');
  if (templateArticle) {
    article.className = templateArticle.className;
    article.style.cssText = templateArticle.style.cssText;
  }

  const header = templateSection.querySelector<HTMLElement>(':scope > header');
  const footer = templateSection.querySelector<HTMLElement>(':scope > footer');
  if (header) section.appendChild(header.cloneNode(true));
  section.appendChild(article);
  if (footer) section.appendChild(footer.cloneNode(true));

  return { section, article };
}

function splitTableForOverflow(
  table: HTMLTableElement,
  currentArticle: HTMLElement,
  maxArticleHeight: number,
): HTMLTableElement | null {
  const rows = Array.from(table.rows);
  if (rows.length < 2) return null;

  const parentSection = currentArticle.closest('section')!;
  const tbody = table.tBodies[0] || table;
  const allRows = Array.from(tbody.children).filter((el): el is HTMLTableRowElement => el.tagName === 'TR');
  if (allRows.length < 2) return null;

  // Binary / incremental fit: remove rows from the bottom until table fits or 1 row remains
  const overflowRows: HTMLTableRowElement[] = [];
  while (allRows.length > 1 && parentSection.scrollHeight > 0 && currentArticle.scrollHeight > maxArticleHeight) {
    const popped = allRows.pop()!;
    tbody.removeChild(popped);
    overflowRows.unshift(popped);
  }

  if (overflowRows.length === 0) return null;
  if (allRows.length === 1 && currentArticle.children.length > 1 && currentArticle.scrollHeight > maxArticleHeight) {
    // Even 1 row didn't fit alongside earlier paragraphs: restore all rows and move entire table to next page
    for (const r of overflowRows) tbody.appendChild(r);
    return null;
  }

  const nextTable = document.createElement('table');
  nextTable.className = table.className;
  nextTable.style.cssText = table.style.cssText;
  const colgroup = table.querySelector(':scope > colgroup');
  if (colgroup) nextTable.appendChild(colgroup.cloneNode(true));
  if (table.tHead) nextTable.appendChild(table.tHead.cloneNode(true));
  const nextTbody = document.createElement('tbody');
  for (const r of overflowRows) nextTbody.appendChild(r);
  nextTable.appendChild(nextTbody);
  return nextTable;
}

/**
 * Splits oversized single-section docx-preview output into discrete A4 paper pages
 * so documents without explicit `<w:lastRenderedPageBreak/>` (common in WPS/Word XML)
 * paginate accurately like a real desktop Office viewer.
 */
function paginateDocxSections(wrapper: HTMLElement): number {
  const sections = Array.from(wrapper.querySelectorAll<HTMLElement>(':scope > section'));
  if (!sections.length) return 0;

  const finalSections: HTMLElement[] = [];

  for (const origSection of sections) {
    const cs = getComputedStyle(origSection);
    const targetHeight = Math.max(700, parseFloat(cs.minHeight) || 1122.5);
    const padTop = parseFloat(cs.paddingTop) || 96;
    const padBottom = parseFloat(cs.paddingBottom) || 96;
    const headerEl = origSection.querySelector<HTMLElement>(':scope > header');
    const footerEl = origSection.querySelector<HTMLElement>(':scope > footer');
    const hfHeight = (headerEl?.offsetHeight || 0) + (footerEl?.offsetHeight || 0);
    const maxArticleHeight = Math.max(240, targetHeight - padTop - padBottom - hfHeight - 6);

    const origArticles = Array.from(origSection.querySelectorAll<HTMLElement>(':scope > article'));
    const allBlocks: HTMLElement[] = [];
    for (const art of origArticles) {
      while (art.firstElementChild) {
        allBlocks.push(art.firstElementChild as HTMLElement);
        art.removeChild(art.firstElementChild);
      }
    }

    if (allBlocks.length === 0) {
      origSection.style.minHeight = `${targetHeight}px`;
      finalSections.push(origSection);
      continue;
    }

    let current = clonePageShell(origSection, targetHeight);
    origSection.replaceWith(current.section);
    finalSections.push(current.section);

    let queue: HTMLElement[] = allBlocks;
    while (queue.length > 0) {
      const block = queue.shift()!;
      current.article.appendChild(block);

      if (current.article.scrollHeight > maxArticleHeight) {
        if (block.tagName === 'TABLE') {
          const splitRest = splitTableForOverflow(block as HTMLTableElement, current.article, maxArticleHeight);
          if (splitRest) {
            queue.unshift(splitRest);
            current = clonePageShell(origSection, targetHeight);
            finalSections[finalSections.length - 1].after(current.section);
            finalSections.push(current.section);
            continue;
          }
        }

        if (current.article.children.length > 1) {
          current.article.removeChild(block);
          current = clonePageShell(origSection, targetHeight);
          finalSections[finalSections.length - 1].after(current.section);
          finalSections.push(current.section);
          current.article.appendChild(block);
        }
      }
    }
  }

  return finalSections.length;
}

export async function mountDocxOfficeViewer(opts: DocxMountOptions): Promise<DocxController> {
  const { container, mcp, filePath } = opts;
  let currentMode: OfficeViewMode = opts.mode ?? 'print';
  let currentZoom: 'fit' | number = opts.zoomSetting ?? 'fit';
  let destroyed = false;
  let resizeObs: ResizeObserver | null = null;
  let pageCount = 1;
  let plainTextCache = '';
  let cachedBytes: Uint8Array | null = null;
  let isDraft = false;

  container.innerHTML = `
    <div id="docx-stage-scroll" class="office-scroll">
      <div class="office-loading">正在加载 Word 文档并排版 A4 页面…</div>
    </div>
  `;

  function applyScale(paperRoot: HTMLElement, scrollHost: HTMLElement) {
    const sections = paperRoot.querySelectorAll<HTMLElement>('section.docx');
    const firstSec = sections[0];
    if (!firstSec) return;
    const naturalWidth = firstSec.offsetWidth || 794;
    const avail = Math.max(240, scrollHost.clientWidth - 28);
    const scale = currentZoom === 'fit' ? Math.min(1.18, avail / naturalWidth) : currentZoom;
    (paperRoot.style as CSSStyleDeclaration & { zoom?: string }).zoom = String(Number(scale.toFixed(3)));
  }

  async function renderPrintLayout() {
    container.innerHTML = `<div id="docx-stage-scroll" class="office-scroll"><div class="office-loading">正在渲染仿真纸张视图…</div></div>`;
    const scrollEl = container.querySelector<HTMLElement>('#docx-stage-scroll')!;

    if (opts.rawBuffer) {
      cachedBytes = new Uint8Array(opts.rawBuffer);
    } else if (filePath) {
      const loaded = await loadDocxBytes(mcp, filePath);
      cachedBytes = loaded.bytes;
      isDraft = loaded.isDraft;
    }
    if (destroyed || !cachedBytes) return;

    scrollEl.innerHTML = `<div id="docx-paper-host" class="docx-paper-host"></div>`;
    const host = scrollEl.querySelector<HTMLElement>('#docx-paper-host')!;

    await renderAsync(cachedBytes.buffer.slice(cachedBytes.byteOffset, cachedBytes.byteOffset + cachedBytes.byteLength), host, undefined, {
      className: 'docx',
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: true,
      experimental: true,
    });
    if (destroyed) return;

    const imgs = Array.from(host.querySelectorAll('img'));
    await Promise.all(
      imgs.map((img) =>
        img.complete
          ? Promise.resolve()
          : new Promise<void>((res) => {
              img.onload = () => res();
              img.onerror = () => res();
            }),
      ),
    );
    if (destroyed) return;

    const wrapper = host.querySelector<HTMLElement>('.docx-wrapper');
    if (wrapper) {
      pageCount = paginateDocxSections(wrapper);
    }

    plainTextCache = host.innerText || '';

    // Build outline TOC from headings inside rendered A4 pages
    const outline: OutlineNode[] = [];
    const headings = host.querySelectorAll<HTMLElement>('h1, h2, h3, p');
    let headingSeq = 0;
    headings.forEach((el) => {
      const tag = el.tagName;
      const text = (el.textContent || '').trim();
      if (!text || text.length > 80) return;
      const isHeadingTag = tag === 'H1' || tag === 'H2' || tag === 'H3';
      const isStyledHeading =
        tag === 'P' &&
        (el.className.toLowerCase().includes('heading') ||
          el.className.toLowerCase().includes('title') ||
          (parseFloat(getComputedStyle(el).fontSize) >= 18 && text.length <= 40));
      if (isHeadingTag || isStyledHeading) {
        headingSeq += 1;
        const id = `docx-h-${headingSeq}`;
        el.id = id;
        const level = tag === 'H1' ? 1 : tag === 'H2' ? 2 : tag === 'H3' ? 3 : 2;
        outline.push({ title: text, level, id });
      }
    });
    opts.onOutline?.(outline.slice(0, 120));

    applyScale(host, scrollEl);
    resizeObs?.disconnect();
    resizeObs = new ResizeObserver(() => {
      if (!destroyed && currentMode === 'print') {
        applyScale(host, scrollEl);
      }
    });
    resizeObs.observe(scrollEl);

    const statusBadge = document.createElement('div');
    statusBadge.className = 'office-page-pill';
    statusBadge.textContent = `${pageCount} 页${isDraft ? ' · 草稿预览' : ''}`;
    scrollEl.appendChild(statusBadge);
  }

  async function renderInteractiveEditor() {
    if (!filePath) {
      container.innerHTML = `<div class="error-box">拖拽本地 .docx 暂不支持保存回磁盘，请通过文件路径打开。</div>`;
      return;
    }
    container.innerHTML = `<div class="office-scroll"><div class="office-loading">正在解析文档段落与表格结构…</div></div>`;
    const inspected = await mcp.callTool<{
      paragraphCount: number;
      tableCount: number;
      paragraphs: Array<{ index: number; text: string; styleId?: string; inTable: boolean }>;
      tables: Array<{ tableIndex: number; rows: string[][] }>;
      hasDraft: boolean;
    }>('app_inspect_docx', { path: filePath, useDraft: true });
    if (destroyed) return;

    const nonEmptyParas = inspected.paragraphs.filter((p) => !p.inTable && p.text.trim().length > 0);
    const paraCards = nonEmptyParas
      .map(
        (p) => `
        <div class="edit-block">
          <div class="edit-block-meta">段落 #${p.index} ${p.styleId ? `<span class="tag">${escapeHtml(p.styleId)}</span>` : ''}</div>
          <textarea class="edit-textarea" data-para-index="${p.index}" rows="${Math.min(6, Math.max(1, Math.ceil(p.text.length / 48)))}">${escapeHtml(p.text)}</textarea>
        </div>
      `,
      )
      .join('');

    const tableCards = inspected.tables
      .map(
        (tbl) => `
        <div class="edit-block">
          <div class="edit-block-meta">表格 #${tbl.tableIndex + 1} (${tbl.rows.length} 行)</div>
          <div style="overflow-x:auto">
            <table class="edit-grid-table">
              ${tbl.rows
                .map(
                  (r, rIdx) =>
                    `<tr>${r
                      .map(
                        (cell, cIdx) =>
                          `<td><input class="edit-cell-input" data-tbl="${tbl.tableIndex}" data-row="${rIdx}" data-col="${cIdx}" value="${escapeHtml(cell)}" /></td>`,
                      )
                      .join('')}</tr>`,
                )
                .join('')}
            </table>
          </div>
        </div>
      `,
      )
      .join('');

    container.innerHTML = `
      <div class="subtoolbar">
        <span style="font-weight:600">交互式 Word 编辑器（保留原字体/样式/图片）</span>
        <span class="spacer"></span>
        <button id="docx-save-draft" class="btn-sm primary">暂存到草稿 Worktree</button>
        <button id="docx-save-direct" class="btn-sm">直接写入磁盘</button>
      </div>
      <div class="editor-canvas">
        <div class="editor-sheet">
          <div class="edit-find-bar">
            <input id="docx-find-input" class="path-input" placeholder="查找原文…" />
            <input id="docx-replace-input" class="path-input" placeholder="替换为…" />
            <button id="docx-apply-replace" class="btn-sm">跨 Run 替换</button>
          </div>
          ${paraCards}
          ${tableCards}
          <div class="edit-block append-box">
            <div class="edit-block-meta">追加新段落</div>
            <div style="display:flex;gap:6px;align-items:center">
              <select id="docx-append-type">
                <option value="paragraph">正文段落</option>
                <option value="heading1">一级标题</option>
                <option value="heading2">二级标题</option>
              </select>
              <input id="docx-append-text" class="path-input" placeholder="输入新增内容…" />
              <button id="docx-append-btn" class="btn-sm">追加</button>
            </div>
          </div>
        </div>
      </div>
    `;

    const dirtyParas = new Map<number, string>();
    const dirtyCells = new Map<string, { tableIndex: number; row: number; col: number; text: string }>();

    container.querySelectorAll<HTMLTextAreaElement>('.edit-textarea').forEach((ta) => {
      ta.addEventListener('input', () => {
        dirtyParas.set(Number(ta.dataset.paraIndex), ta.value);
      });
    });
    container.querySelectorAll<HTMLInputElement>('.edit-cell-input').forEach((inp) => {
      inp.addEventListener('input', () => {
        const tableIndex = Number(inp.dataset.tbl);
        const row = Number(inp.dataset.row);
        const col = Number(inp.dataset.col);
        dirtyCells.set(`${tableIndex}:${row}:${col}`, { tableIndex, row, col, text: inp.value });
      });
    });

    async function saveEdits(mode: 'draft' | 'direct') {
      const updateParagraphs = Array.from(dirtyParas.entries()).map(([index, text]) => ({ index, text }));
      const updateTableCells = Array.from(dirtyCells.values());
      if (!updateParagraphs.length && !updateTableCells.length) {
        opts.onToast?.('没有检测到内容改动');
        return;
      }
      await mcp.callTool('app_save_docx_edits', {
        path: filePath,
        mode,
        summary: `修改 ${updateParagraphs.length} 个段落, ${updateTableCells.length} 个表格单元格`,
        updateParagraphs: updateParagraphs.length ? updateParagraphs : undefined,
        updateTableCells: updateTableCells.length ? updateTableCells : undefined,
      });
      opts.onToast?.(mode === 'draft' ? '已暂存到草稿 Worktree' : '已保存并备份旧版本');
      opts.onWorktreeChanged?.();
      await setMode('print');
    }

    container.querySelector('#docx-save-draft')!.addEventListener('click', () => void saveEdits('draft'));
    container.querySelector('#docx-save-direct')!.addEventListener('click', () => void saveEdits('direct'));

    container.querySelector('#docx-apply-replace')!.addEventListener('click', async () => {
      const find = container.querySelector<HTMLInputElement>('#docx-find-input')!.value;
      const replace = container.querySelector<HTMLInputElement>('#docx-replace-input')!.value;
      if (!find) return;
      await mcp.callTool('app_save_docx_edits', {
        path: filePath,
        mode: 'draft',
        summary: `替换 "${find}" → "${replace}"`,
        replacements: [{ find, replace }],
      });
      opts.onToast?.('已暂存替换结果到草稿');
      opts.onWorktreeChanged?.();
      await renderInteractiveEditor();
    });

    container.querySelector('#docx-append-btn')!.addEventListener('click', async () => {
      const type = container.querySelector<HTMLSelectElement>('#docx-append-type')!.value as
        | 'paragraph'
        | 'heading1'
        | 'heading2';
      const text = container.querySelector<HTMLInputElement>('#docx-append-text')!.value.trim();
      if (!text) return;
      await mcp.callTool('app_save_docx_edits', {
        path: filePath,
        mode: 'draft',
        summary: `追加${type}: ${text.slice(0, 24)}`,
        appendBlocks: [{ type, text }],
      });
      opts.onToast?.('已追加到草稿');
      opts.onWorktreeChanged?.();
      await renderInteractiveEditor();
    });
  }

  async function renderCompareView() {
    if (!filePath) return;
    container.innerHTML = `<div class="office-scroll"><div class="office-loading">正在计算 Worktree 变更差异…</div></div>`;
    const diff = await mcp.callTool<{
      hasDraft: boolean;
      summary?: string;
      changedCount?: number;
      changedParagraphs?: Array<{ index: number; before: string; after: string }>;
    }>('app_worktree_action', { action: 'diff', path: filePath });
    if (destroyed) return;

    if (!diff.hasDraft) {
      container.innerHTML = `<div class="empty-hint" style="padding:36px">当前文档没有待审核的 Worktree 草稿更改。切换到「交互编辑模式」或让 Agent 修改后即可对比差异。</div>`;
      return;
    }

    const rows = (diff.changedParagraphs || [])
      .map(
        (p) => `
        <div class="diff-card">
          <div class="diff-meta">段落 #${p.index}</div>
          <div class="diff-before">- ${escapeHtml(p.before || '(空)')}</div>
          <div class="diff-after">+ ${escapeHtml(p.after || '(空)')}</div>
        </div>
      `,
      )
      .join('');

    container.innerHTML = `
      <div class="subtoolbar">
        <span>Worktree 变更对比：${escapeHtml(diff.summary || '')}（共 ${diff.changedCount ?? 0} 处段落差异）</span>
        <span class="spacer"></span>
        <button id="cmp-back-print" class="btn-sm primary">查看草稿纸张效果</button>
      </div>
      <div class="editor-canvas">
        <div class="editor-sheet">${rows || '<div class="empty-hint">结构或样式已变更（无纯文本差异）</div>'}</div>
      </div>
    `;
    container.querySelector('#cmp-back-print')!.addEventListener('click', () => void setMode('print'));
  }

  async function setMode(mode: OfficeViewMode) {
    currentMode = mode;
    if (mode === 'print') await renderPrintLayout();
    else if (mode === 'edit') await renderInteractiveEditor();
    else await renderCompareView();
  }

  await setMode(currentMode);

  return {
    setMode,
    setZoom(zoom) {
      currentZoom = zoom;
      const host = container.querySelector<HTMLElement>('#docx-paper-host');
      const scrollEl = container.querySelector<HTMLElement>('#docx-stage-scroll');
      if (host && scrollEl) applyScale(host, scrollEl);
    },
    getContextSummary: () => plainTextCache.slice(0, 6000),
    scrollToHeading(id) {
      const target = container.querySelector<HTMLElement>(`#${CSS.escape(id)}`);
      target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    },
    destroy() {
      destroyed = true;
      resizeObs?.disconnect();
    },
  };
}
