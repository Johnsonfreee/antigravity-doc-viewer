import type { ViewerMcpClient } from '../mcpClient';
import { base64ToUint8Array, type OutlineNode } from './pdfViewer';

export interface SimpleViewerController {
  getContextSummary: () => string;
  scrollToHeading?: (id: string) => void;
  destroy?: () => void;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export async function mountRichDocView(opts: {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  kind: 'markdown' | 'docx';
  filePath: string;
  onOutline?: (items: OutlineNode[]) => void;
}): Promise<SimpleViewerController> {
  const { container, mcp, kind, filePath } = opts;
  container.innerHTML = `<div class="empty-hint">正在解析文档…</div>`;

  const toolName = kind === 'markdown' ? 'app_render_markdown' : 'app_render_docx';
  const data = await mcp.callTool<{
    html: string;
    toc: Array<{ id: string; level: number; text: string }>;
    truncated?: boolean;
    warnings?: string[];
    wordCount?: number;
    omittedImages?: number;
  }>(toolName, { path: filePath });

  let fontSize = 13.5;

  const extraMeta = [
    data.truncated ? '已截断前 5 MiB' : '',
    data.wordCount !== undefined ? `约 ${data.wordCount} 词` : '',
    data.omittedImages ? `${data.omittedImages} 张大图已省略` : '',
  ]
    .filter(Boolean)
    .join(' · ');

  container.innerHTML = `
    <div class="subtoolbar">
      <span style="font-weight:600">${kind === 'markdown' ? 'Markdown' : 'Word (.docx)'}</span>
      ${extraMeta ? `<span style="font-size:11px;color:var(--fg-2)">${escapeHtml(extraMeta)}</span>` : ''}
      <span class="spacer"></span>
      <button id="doc-font-down" class="btn-sm">A-</button>
      <button id="doc-font-up" class="btn-sm">A+</button>
      <button id="doc-copy-text" class="btn-sm">复制纯文本</button>
    </div>
    <div class="viewer-scroll">
      <article id="doc-prose" class="prose" style="font-size:${fontSize}px"></article>
    </div>
  `;

  const proseEl = container.querySelector<HTMLElement>('#doc-prose')!;
  proseEl.innerHTML = data.html;

  opts.onOutline?.(
    (data.toc || []).map((t) => ({
      title: t.text,
      level: t.level,
      id: t.id,
    })),
  );

  container.querySelector('#doc-font-down')!.addEventListener('click', () => {
    fontSize = Math.max(11, fontSize - 1);
    proseEl.style.fontSize = `${fontSize}px`;
  });

  container.querySelector('#doc-font-up')!.addEventListener('click', () => {
    fontSize = Math.min(20, fontSize + 1);
    proseEl.style.fontSize = `${fontSize}px`;
  });

  container.querySelector('#doc-copy-text')!.addEventListener('click', async () => {
    await navigator.clipboard?.writeText(proseEl.innerText);
  });

  return {
    getContextSummary: () => proseEl.innerText.slice(0, 6000),
    scrollToHeading(id: string) {
      const target = proseEl.querySelector<HTMLElement>(`#${CSS.escape(id)}`);
      target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    },
  };
}

export async function mountSpreadsheetView(opts: {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  filePath: string;
  initialSheet?: string;
  initialPage?: number;
  onSheetChange?: (sheet: string) => void;
}): Promise<SimpleViewerController> {
  const { container, mcp, filePath } = opts;
  let activeSheet = opts.initialSheet;
  let page = opts.initialPage ?? 1;
  let search = '';
  let firstRowHeader = true;
  let latestRowsCsv = '';

  container.innerHTML = `
    <div class="subtoolbar">
      <input id="ss-search" type="search" placeholder="筛选行…" style="width:140px" />
      <label style="font-size:11.5px;display:inline-flex;align-items:center;gap:4px;cursor:pointer">
        <input id="ss-header-toggle" type="checkbox" checked /> 首行作表头
      </label>
      <span class="spacer"></span>
      <button id="ss-prev" class="btn-sm">上一页</button>
      <span id="ss-page-info" style="font-size:11.5px;color:var(--fg-2)">1/1</span>
      <button id="ss-next" class="btn-sm">下一页</button>
      <button id="ss-copy" class="btn-sm">复制本页 CSV</button>
    </div>
    <div id="ss-sheets" class="sheet-tabs"></div>
    <div id="ss-wrap" class="sheet-wrap">
      <div class="empty-hint">正在解析表格…</div>
    </div>
  `;

  const sheetsEl = container.querySelector<HTMLElement>('#ss-sheets')!;
  const wrapEl = container.querySelector<HTMLElement>('#ss-wrap')!;
  const pageInfoEl = container.querySelector<HTMLElement>('#ss-page-info')!;
  const searchInput = container.querySelector<HTMLInputElement>('#ss-search')!;
  const headerToggle = container.querySelector<HTMLInputElement>('#ss-header-toggle')!;

  let currentTotalPages = 1;

  async function loadSheet() {
    const res = await mcp.callTool<{
      sheetNames: string[];
      activeSheet: string;
      columns: string[];
      rows: Array<{ rowNumber: number; cells: string[] }>;
      totalRows: number;
      filteredRows: number;
      totalCols: number;
      page: number;
      totalPages: number;
    }>('app_render_spreadsheet', {
      path: filePath,
      sheetName: activeSheet,
      page,
      pageSize: 150,
      search,
    });

    activeSheet = res.activeSheet;
    page = res.page;
    currentTotalPages = res.totalPages;
    pageInfoEl.textContent = `${res.page}/${res.totalPages} (${res.filteredRows} 行)`;

    sheetsEl.innerHTML = '';
    for (const name of res.sheetNames) {
      const btn = document.createElement('button');
      btn.className = `sheet-tab ${name === res.activeSheet ? 'active' : ''}`;
      btn.textContent = name;
      btn.addEventListener('click', () => {
        activeSheet = name;
        page = 1;
        opts.onSheetChange?.(name);
        void loadSheet();
      });
      sheetsEl.appendChild(btn);
    }

    if (res.rows.length === 0) {
      wrapEl.innerHTML = `<div class="empty-hint">当前筛选结果为空</div>`;
      latestRowsCsv = '';
      return;
    }

    const useFirstAsHeader = firstRowHeader && res.page === 1 && !search.trim() && res.rows.length > 0;
    const headerCells = useFirstAsHeader ? res.rows[0].cells : res.columns;
    const bodyRows = useFirstAsHeader ? res.rows.slice(1) : res.rows;

    const ths = headerCells.map((h) => `<th title="${escapeHtml(h)}">${escapeHtml(h)}</th>`).join('');
    const trs = bodyRows
      .map(
        (r) =>
          `<tr><td class="row-num">${r.rowNumber}</td>${r.cells
            .map((c) => `<td title="${escapeHtml(c)}">${escapeHtml(c)}</td>`)
            .join('')}</tr>`,
      )
      .join('');

    wrapEl.innerHTML = `<table class="sheet-table"><thead><tr><th class="row-num">#</th>${ths}</tr></thead><tbody>${trs}</tbody></table>`;

    latestRowsCsv = res.rows
      .slice(0, 60)
      .map((r) => r.cells.map((c) => `"${c.replace(/"/g, '""')}"`).join(','))
      .join('\n');
  }

  let timer: number | undefined;
  searchInput.addEventListener('input', () => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => {
      search = searchInput.value;
      page = 1;
      void loadSheet();
    }, 250);
  });

  headerToggle.addEventListener('change', () => {
    firstRowHeader = headerToggle.checked;
    void loadSheet();
  });

  container.querySelector('#ss-prev')!.addEventListener('click', () => {
    if (page > 1) {
      page -= 1;
      void loadSheet();
    }
  });

  container.querySelector('#ss-next')!.addEventListener('click', () => {
    if (page < currentTotalPages) {
      page += 1;
      void loadSheet();
    }
  });

  container.querySelector('#ss-copy')!.addEventListener('click', async () => {
    await navigator.clipboard?.writeText(latestRowsCsv);
  });

  await loadSheet();

  return {
    getContextSummary: () =>
      `Sheet: ${activeSheet || 'default'} (Page ${page}/${currentTotalPages})\n${latestRowsCsv.slice(0, 5000)}`,
  };
}

export async function mountHtmlView(opts: {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  filePath?: string;
  rawHtmlText?: string;
}): Promise<SimpleViewerController> {
  const { container, mcp, filePath } = opts;
  container.innerHTML = `<div class="empty-hint">正在加载 HTML…</div>`;

  let previewHtml: string;
  let sourceHtml: string;

  if (filePath) {
    const res = await mcp.callTool<{
      previewHtml: string;
      sourceHtml: string;
      sourceTruncated: boolean;
    }>('app_render_html', { path: filePath });
    previewHtml = res.previewHtml;
    sourceHtml = res.sourceHtml;
  } else {
    const raw = opts.rawHtmlText || '';
    previewHtml = raw.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '');
    sourceHtml = escapeHtml(raw);
  }

  let mode: 'preview' | 'source' = 'preview';
  let mobileWidth = false;

  container.innerHTML = `
    <div class="subtoolbar">
      <button id="html-mode-preview" class="btn-sm primary">预览（隔离模式）</button>
      <button id="html-mode-source" class="btn-sm">源码</button>
      <button id="html-width-toggle" class="btn-sm">390px 移动端</button>
      <span class="spacer"></span>
      <span style="font-size:11px;color:var(--fg-2)">已禁用脚本与外网请求</span>
    </div>
    <div id="html-stage" class="html-stage"></div>
  `;

  const stageEl = container.querySelector<HTMLElement>('#html-stage')!;
  const btnPreview = container.querySelector<HTMLButtonElement>('#html-mode-preview')!;
  const btnSource = container.querySelector<HTMLButtonElement>('#html-mode-source')!;
  const btnWidth = container.querySelector<HTMLButtonElement>('#html-width-toggle')!;

  function renderMode() {
    btnPreview.className = `btn-sm ${mode === 'preview' ? 'primary' : ''}`;
    btnSource.className = `btn-sm ${mode === 'source' ? 'primary' : ''}`;
    btnWidth.classList.toggle('hidden', mode !== 'preview');

    if (mode === 'source') {
      stageEl.innerHTML = `<div class="viewer-scroll" style="width:100%;background:var(--code-bg)"><pre class="hljs" style="margin:0;font-family:var(--font-code);font-size:12px;white-space:pre-wrap">${sourceHtml}</pre></div>`;
      return;
    }

    // Use Shadow DOM by default so preview works cleanly even when host CSP has frame-src 'none'
    stageEl.innerHTML = '';
    const wrapper = document.createElement('div');
    wrapper.style.width = mobileWidth ? '390px' : '100%';
    wrapper.style.maxWidth = '100%';
    wrapper.style.minHeight = '100%';
    wrapper.style.background = '#ffffff';
    wrapper.style.color = '#111827';
    wrapper.style.border = '1px solid var(--bd-1)';
    wrapper.style.borderRadius = '6px';
    wrapper.style.overflow = 'auto';
    wrapper.style.padding = '12px';
    stageEl.appendChild(wrapper);

    const shadow = wrapper.attachShadow({ mode: 'open' });
    const bodyMatch = /<body\b[^>]*>([\s\S]*?)<\/body>/i.exec(previewHtml);
    const styleMatches = [...previewHtml.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)].map((m) => m[0]).join('\n');
    const bodyContent = bodyMatch ? bodyMatch[1] : previewHtml;
    shadow.innerHTML = `${styleMatches}<div class="html-preview-root">${bodyContent}</div>`;
  }

  btnPreview.addEventListener('click', () => {
    mode = 'preview';
    renderMode();
  });
  btnSource.addEventListener('click', () => {
    mode = 'source';
    renderMode();
  });
  btnWidth.addEventListener('click', () => {
    mobileWidth = !mobileWidth;
    btnWidth.textContent = mobileWidth ? '100% 桌面宽度' : '390px 移动端';
    renderMode();
  });

  renderMode();

  return {
    getContextSummary: () => stageEl.innerText.slice(0, 5000),
  };
}

export async function mountTextView(opts: {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  filePath: string;
}): Promise<SimpleViewerController> {
  const { container, mcp, filePath } = opts;
  let startLine = 1;
  const lineCount = 400;
  let wrap = false;
  let latestText = '';

  container.innerHTML = `
    <div class="subtoolbar">
      <span id="txt-lang" style="font-weight:600">TEXT</span>
      <span id="txt-range" style="font-size:11px;color:var(--fg-2)"></span>
      <span class="spacer"></span>
      <label style="font-size:11.5px;display:inline-flex;align-items:center;gap:4px;cursor:pointer">
        <input id="txt-wrap" type="checkbox" /> 自动换行
      </label>
      <button id="txt-prev" class="btn-sm">上 400 行</button>
      <button id="txt-next" class="btn-sm">下 400 行</button>
      <button id="txt-copy" class="btn-sm">复制</button>
    </div>
    <div id="txt-scroll" class="viewer-scroll" style="padding:0"></div>
  `;

  const langEl = container.querySelector<HTMLElement>('#txt-lang')!;
  const rangeEl = container.querySelector<HTMLElement>('#txt-range')!;
  const scrollEl = container.querySelector<HTMLElement>('#txt-scroll')!;
  const prevBtn = container.querySelector<HTMLButtonElement>('#txt-prev')!;
  const nextBtn = container.querySelector<HTMLButtonElement>('#txt-next')!;
  const wrapInput = container.querySelector<HTMLInputElement>('#txt-wrap')!;

  async function loadChunk() {
    const res = await mcp.callTool<{
      linesHtml: string[];
      rawText: string;
      language: string;
      totalLines: number;
      startLine: number;
      endLine: number;
      hasMore: boolean;
    }>('app_read_text', { path: filePath, startLine, lineCount });

    latestText = res.rawText;
    langEl.textContent = res.language.toUpperCase();
    rangeEl.textContent = `${res.startLine}-${res.endLine} / ${res.totalLines} 行`;
    prevBtn.disabled = res.startLine <= 1;
    nextBtn.disabled = !res.hasMore;

    const rows = res.linesHtml
      .map((lineHtml, idx) => `<tr><td class="ln">${res.startLine + idx}</td><td class="lc">${lineHtml || ' '}</td></tr>`)
      .join('');
    scrollEl.innerHTML = `<table class="code-table ${wrap ? 'wrap' : ''}"><tbody>${rows}</tbody></table>`;
    scrollEl.scrollTop = 0;
  }

  wrapInput.addEventListener('change', () => {
    wrap = wrapInput.checked;
    scrollEl.querySelector('.code-table')?.classList.toggle('wrap', wrap);
  });

  prevBtn.addEventListener('click', () => {
    startLine = Math.max(1, startLine - lineCount);
    void loadChunk();
  });

  nextBtn.addEventListener('click', () => {
    startLine += lineCount;
    void loadChunk();
  });

  container.querySelector('#txt-copy')!.addEventListener('click', async () => {
    await navigator.clipboard?.writeText(latestText);
  });

  await loadChunk();

  return {
    getContextSummary: () => latestText.slice(0, 6000),
  };
}

export async function mountImageView(opts: {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  filePath?: string;
  ext: string;
  rawBuffer?: ArrayBuffer;
}): Promise<SimpleViewerController> {
  const { container, mcp, filePath, ext } = opts;
  const mimeMap: Record<string, string> = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.webp': 'image/webp',
    '.svg': 'image/svg+xml',
    '.bmp': 'image/bmp',
    '.ico': 'image/x-icon',
  };
  const mime = mimeMap[ext.toLowerCase()] || 'image/png';

  container.innerHTML = `<div class="empty-hint">正在加载图片…</div>`;

  let blobUrl = '';
  if (opts.rawBuffer) {
    blobUrl = URL.createObjectURL(new Blob([opts.rawBuffer], { type: mime }));
  } else if (filePath) {
    const parts: Uint8Array[] = [];
    let offset = 0;
    while (true) {
      const chunk = await mcp.callTool<{
        base64: string;
        bytesRead: number;
        eof: boolean;
      }>('app_read_file_chunk', { path: filePath, offset, length: 384 * 1024 });
      if (chunk.bytesRead > 0) {
        parts.push(base64ToUint8Array(chunk.base64));
        offset += chunk.bytesRead;
      }
      if (chunk.eof || chunk.bytesRead === 0) break;
    }
    const merged = new Uint8Array(offset);
    let pos = 0;
    for (const p of parts) {
      merged.set(p, pos);
      pos += p.byteLength;
    }
    blobUrl = URL.createObjectURL(new Blob([merged], { type: mime }));
  }

  let fit = true;
  container.innerHTML = `
    <div class="subtoolbar">
      <span>图片预览</span>
      <span class="spacer"></span>
      <button id="img-fit" class="btn-sm primary">适应窗口</button>
    </div>
    <div class="img-stage"><img id="img-el" src="${blobUrl}" alt="preview" /></div>
  `;

  const imgEl = container.querySelector<HTMLImageElement>('#img-el')!;
  const fitBtn = container.querySelector<HTMLButtonElement>('#img-fit')!;
  fitBtn.addEventListener('click', () => {
    fit = !fit;
    fitBtn.textContent = fit ? '适应窗口' : '原始大小';
    fitBtn.className = `btn-sm ${fit ? 'primary' : ''}`;
    imgEl.style.maxWidth = fit ? '100%' : 'none';
    imgEl.style.maxHeight = fit ? '100%' : 'none';
  });

  return {
    getContextSummary: () => `Image file: ${filePath || 'dropped image'}`,
    destroy: () => {
      if (blobUrl) URL.revokeObjectURL(blobUrl);
    },
  };
}
