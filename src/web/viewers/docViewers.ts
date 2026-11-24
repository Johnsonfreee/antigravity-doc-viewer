import type { ViewerMcpClient } from '../mcpClient';
import { buildShadowHtmlDocument } from '../htmlShadowAdapter';
import { base64ToUint8Array, type OutlineNode } from './pdfViewer';

export interface SimpleViewerController {
  getContextSummary: () => string;
  scrollToHeading?: (id: string) => void;
  destroy?: () => void;
}

export interface LegacyDocController {
  delegatedToDocx?: boolean;
  setZoom: (zoom: 'fit' | number) => void;
  goToPage: (page: number) => void;
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
  let plainText = '';
  let liveUrl: string | null = null;
  let staticUrl: string | null = null;

  if (filePath) {
    const res = await mcp.callTool<{
      previewHtml: string;
      sourceHtml: string;
      sourceTruncated: boolean;
      plainText?: string;
      liveUrl?: string | null;
      staticUrl?: string | null;
    }>('app_render_html', { path: filePath });
    previewHtml = res.previewHtml;
    sourceHtml = res.sourceHtml;
    plainText = res.plainText || '';
    liveUrl = res.liveUrl ?? null;
    staticUrl = res.staticUrl ?? null;
  } else {
    const raw = opts.rawHtmlText || '';
    previewHtml = raw.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '');
    sourceHtml = escapeHtml(raw);
  }

  type HtmlRenderMode = 'live' | 'static' | 'shadow' | 'source';
  type HtmlViewportMode = 'desktop' | 'fit' | 'mobile';
  let mode: HtmlRenderMode = liveUrl ? 'live' : 'shadow';
  let viewport: HtmlViewportMode = 'desktop';
  let resizeObs: ResizeObserver | null = null;
  let resizeTimer: number | undefined;
  let lastStageWidth = 0;
  let lastStageHeight = 0;
  let activeIframe: HTMLIFrameElement | null = null;

  const modeButtonsHtml = liveUrl
    ? `
      <button class="btn-sm primary" data-mode="live">实时浏览器</button>
      <button class="btn-sm" data-mode="static">静态隔离</button>
      <button class="btn-sm" data-mode="shadow">离线内联</button>
      <button class="btn-sm" data-mode="source">源码</button>
    `
    : `
      <button class="btn-sm primary" data-mode="shadow">离线预览</button>
      <button class="btn-sm" data-mode="source">源码</button>
    `;

  container.innerHTML = `
    <div class="subtoolbar">
      <div id="html-mode-group" style="display:inline-flex;gap:4px">
        ${modeButtonsHtml}
      </div>
      <div id="html-viewport-group" style="display:inline-flex;gap:4px;margin-left:6px">
        <button class="btn-sm primary" data-vp="desktop">桌面 1280px</button>
        <button class="btn-sm" data-vp="fit">适应侧栏</button>
        <button class="btn-sm" data-vp="mobile">移动 390px</button>
      </div>
      <span class="spacer"></span>
      <span id="html-status-note" style="font-size:11px;color:var(--fg-2)"></span>
    </div>
    <div id="html-stage" class="html-stage"></div>
  `;

  const stageEl = container.querySelector<HTMLElement>('#html-stage')!;
  const modeGroup = container.querySelector<HTMLElement>('#html-mode-group')!;
  const modeButtons = Array.from(modeGroup.querySelectorAll<HTMLButtonElement>('[data-mode]'));
  const vpGroup = container.querySelector<HTMLElement>('#html-viewport-group')!;
  const vpButtons = Array.from(vpGroup.querySelectorAll<HTMLButtonElement>('[data-vp]'));
  const statusNoteEl = container.querySelector<HTMLElement>('#html-status-note')!;

  function statusLabelForMode(m: HtmlRenderMode): string {
    if (m === 'live') return '完整浏览器引擎 (127.0.0.1)';
    if (m === 'static') return '浏览器静态隔离（禁脚本/外网）';
    if (m === 'shadow') return '离线 Shadow DOM 预览';
    return 'HTML 源码';
  }

  function applyIframeViewport(clipEl: HTMLElement, iframeEl: HTMLIFrameElement, force = false) {
    const pad = viewport === 'mobile' ? 24 : 0;
    const availW = Math.max(240, stageEl.clientWidth - pad);
    const availH = Math.max(240, stageEl.clientHeight - pad);
    if (!force && Math.abs(availW - lastStageWidth) < 4 && Math.abs(availH - lastStageHeight) < 4) return;
    lastStageWidth = availW;
    lastStageHeight = availH;

    if (viewport === 'desktop') {
      if (availW >= 1280) {
        clipEl.style.width = '100%';
        clipEl.style.height = '100%';
        iframeEl.style.width = '100%';
        iframeEl.style.height = '100%';
        iframeEl.style.transform = 'none';
      } else {
        const scale = Math.max(0.2, availW / 1280);
        clipEl.style.width = `${availW}px`;
        clipEl.style.height = `${availH}px`;
        iframeEl.style.width = '1280px';
        iframeEl.style.height = `${Math.ceil(availH / scale)}px`;
        iframeEl.style.transform = `scale(${scale.toFixed(4)})`;
      }
    } else if (viewport === 'fit') {
      clipEl.style.width = '100%';
      clipEl.style.height = '100%';
      iframeEl.style.width = '100%';
      iframeEl.style.height = '100%';
      iframeEl.style.transform = 'none';
    } else {
      clipEl.style.width = `${Math.min(390, availW)}px`;
      clipEl.style.height = `${availH}px`;
      iframeEl.style.width = '100%';
      iframeEl.style.height = '100%';
      iframeEl.style.transform = 'none';
    }
  }

  function applyShadowViewportSizing(zoomShell: HTMLElement, hostEl: HTMLElement, force = false) {
    const avail = Math.max(240, stageEl.clientWidth);
    if (!force && viewport !== 'desktop') return;
    if (!force && Math.abs(avail - lastStageWidth) < 4) return;
    lastStageWidth = avail;

    const zoomStyle = zoomShell.style as CSSStyleDeclaration & { zoom?: string };
    if (viewport === 'desktop') {
      if (avail >= 1280) {
        hostEl.style.width = '100%';
        zoomShell.style.width = '100%';
        zoomStyle.zoom = '1';
      } else {
        const scale = Math.max(0.2, Math.floor((avail / 1280) * 1000) / 1000);
        hostEl.style.width = '1280px';
        zoomShell.style.width = '1280px';
        zoomStyle.zoom = String(scale);
      }
    } else if (viewport === 'fit') {
      zoomShell.style.width = '100%';
      zoomStyle.zoom = '1';
      hostEl.style.width = '100%';
    } else {
      zoomShell.style.width = '390px';
      zoomStyle.zoom = '1';
      hostEl.style.width = '390px';
    }
  }

  function renderShadowMode() {
    stageEl.classList.remove('iframe-mode');
    stageEl.classList.toggle('framed', viewport === 'mobile');
    stageEl.innerHTML = '';

    const zoomShell = document.createElement('div');
    zoomShell.className = 'html-zoom-shell';

    const hostEl = document.createElement('div');
    hostEl.className = 'html-shadow-host';
    zoomShell.appendChild(hostEl);
    stageEl.appendChild(zoomShell);

    const parsed = buildShadowHtmlDocument(previewHtml);
    const shadow = hostEl.attachShadow({ mode: 'open' });

    const resetStyleEl = document.createElement('style');
    resetStyleEl.textContent = parsed.resetCss;
    shadow.appendChild(resetStyleEl);

    if (parsed.stylesCss) {
      const pageStyleEl = document.createElement('style');
      pageStyleEl.textContent = parsed.stylesCss;
      shadow.appendChild(pageStyleEl);
    }

    const viewportBox = document.createElement('div');
    viewportBox.className = 'html-viewport-box';

    const htmlEl = document.createElement('div');
    htmlEl.className = 'html-doc-html';
    for (const [k, v] of parsed.htmlAttrs) {
      if (k.toLowerCase() === 'class') htmlEl.className = `html-doc-html ${v}`.trim();
      else htmlEl.setAttribute(k, v);
    }

    const bodyEl = document.createElement('div');
    bodyEl.className = 'html-doc-body';
    for (const [k, v] of parsed.bodyAttrs) {
      if (k.toLowerCase() === 'class') bodyEl.className = `html-doc-body ${v}`.trim();
      else bodyEl.setAttribute(k, v);
    }
    bodyEl.innerHTML = parsed.bodyHtml;

    htmlEl.appendChild(bodyEl);
    viewportBox.appendChild(htmlEl);
    shadow.appendChild(viewportBox);

    const isTransparent = (bg: string) => !bg || bg === 'transparent' || bg === 'rgba(0, 0, 0, 0)';
    const bodyBg = getComputedStyle(bodyEl).backgroundColor;
    const htmlBg = getComputedStyle(htmlEl).backgroundColor;
    if (!isTransparent(bodyBg)) hostEl.style.backgroundColor = bodyBg;
    else if (!isTransparent(htmlBg)) hostEl.style.backgroundColor = htmlBg;
    else hostEl.style.backgroundColor = '#ffffff';

    shadow.addEventListener('click', (e) => {
      const target = e.target as Element | null;
      const anchor = target?.closest('a[href]') as HTMLAnchorElement | null;
      if (!anchor) return;
      const href = (anchor.getAttribute('href') || '').trim();
      if (href.startsWith('#') && href.length > 1) {
        e.preventDefault();
        let id = href.slice(1);
        try {
          id = decodeURIComponent(id);
        } catch {
          // Keep raw fragment if malformed percent-encoding
        }
        const anchorTarget =
          shadow.getElementById(id) || shadow.querySelector(`[name="${CSS.escape(id)}"]`);
        anchorTarget?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      } else {
        e.preventDefault();
      }
    });

    shadow.addEventListener('submit', (e) => {
      e.preventDefault();
    });

    applyShadowViewportSizing(zoomShell, hostEl, true);

    resizeObs = new ResizeObserver(() => {
      if (mode !== 'shadow' || viewport !== 'desktop') return;
      const w = stageEl.clientWidth;
      if (w <= 0 || Math.abs(w - lastStageWidth) < 4) return;
      window.clearTimeout(resizeTimer);
      resizeTimer = window.setTimeout(() => {
        if (mode === 'shadow' && viewport === 'desktop') {
          applyShadowViewportSizing(zoomShell, hostEl);
        }
      }, 60);
    });
    resizeObs.observe(stageEl);
  }

  function renderMode() {
    window.clearTimeout(resizeTimer);
    resizeObs?.disconnect();
    resizeObs = null;
    lastStageWidth = 0;
    lastStageHeight = 0;
    if (activeIframe) {
      activeIframe.src = 'about:blank';
      activeIframe = null;
    }

    for (const btn of modeButtons) {
      btn.className = `btn-sm ${btn.dataset.mode === mode ? 'primary' : ''}`;
    }
    vpGroup.classList.toggle('hidden', mode === 'source');
    for (const btn of vpButtons) {
      btn.className = `btn-sm ${btn.dataset.vp === viewport ? 'primary' : ''}`;
    }
    statusNoteEl.textContent = statusLabelForMode(mode);

    if (mode === 'source') {
      stageEl.classList.remove('iframe-mode', 'framed');
      stageEl.innerHTML = `<div class="viewer-scroll" style="width:100%;background:var(--code-bg)"><pre class="hljs" style="margin:0;font-family:var(--font-code);font-size:12px;white-space:pre-wrap">${sourceHtml}</pre></div>`;
      return;
    }

    if ((mode === 'live' || mode === 'static') && liveUrl) {
      stageEl.classList.add('iframe-mode');
      stageEl.classList.toggle('framed', viewport === 'mobile');
      stageEl.innerHTML = '';

      const clipEl = document.createElement('div');
      clipEl.className = 'html-iframe-clip';

      const iframeEl = document.createElement('iframe');
      iframeEl.className = 'html-preview-iframe';
      iframeEl.setAttribute(
        'sandbox',
        mode === 'live' ? 'allow-scripts allow-same-origin allow-forms allow-popups' : 'allow-same-origin',
      );
      iframeEl.title = filePath ? filePath.split(/[\\/]/).pop() || 'HTML Preview' : 'HTML Preview';
      iframeEl.src = mode === 'live' ? liveUrl : staticUrl || liveUrl;

      iframeEl.addEventListener('error', () => {
        mode = 'shadow';
        renderMode();
      });

      clipEl.appendChild(iframeEl);
      stageEl.appendChild(clipEl);
      activeIframe = iframeEl;

      applyIframeViewport(clipEl, iframeEl, true);

      resizeObs = new ResizeObserver(() => {
        if (mode !== 'live' && mode !== 'static') return;
        window.clearTimeout(resizeTimer);
        resizeTimer = window.setTimeout(() => {
          if (mode === 'live' || mode === 'static') {
            applyIframeViewport(clipEl, iframeEl);
          }
        }, 60);
      });
      resizeObs.observe(stageEl);
      return;
    }

    renderShadowMode();
  }

  for (const btn of modeButtons) {
    btn.addEventListener('click', () => {
      mode = btn.dataset.mode as HtmlRenderMode;
      renderMode();
    });
  }
  for (const btn of vpButtons) {
    btn.addEventListener('click', () => {
      viewport = btn.dataset.vp as HtmlViewportMode;
      renderMode();
    });
  }

  renderMode();

  return {
    getContextSummary: () => {
      if (plainText) return plainText.slice(0, 5000);
      const shadowBody = stageEl
        .querySelector('.html-shadow-host')
        ?.shadowRoot?.querySelector<HTMLElement>('.html-doc-body');
      return (shadowBody?.innerText || stageEl.innerText).slice(0, 5000);
    },
    destroy: () => {
      window.clearTimeout(resizeTimer);
      resizeObs?.disconnect();
      resizeObs = null;
      if (activeIframe) {
        activeIframe.src = 'about:blank';
        activeIframe = null;
      }
    },
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

export async function mountLegacyDocView(opts: {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  filePath: string;
  zoomSetting?: 'fit' | number;
  onOutline?: (items: OutlineNode[]) => void;
  onPageInfo?: (info: { currentPage: number; totalPages: number }) => void;
}): Promise<LegacyDocController> {
  const { container, mcp, filePath } = opts;
  let currentZoom: 'fit' | number = opts.zoomSetting ?? 'fit';
  let destroyed = false;
  let resizeObs: ResizeObserver | null = null;
  let resizeTimer: number | undefined;
  let lastAvailWidth = 0;
  let lastAppliedScale = 0;
  let currentPage = 1;

  container.innerHTML = `<div class="office-scroll"><div class="office-loading">正在离线解析文档排版…</div></div>`;
  const res = await mcp.callTool<{
    format: string;
    delegateKind?: 'docx';
    pages: Array<{ pageNumber: number; html: string }>;
    toc: Array<{ id: string; level: number; text: string }>;
    plainText: string;
    notices: string[];
  }>('app_render_legacy_doc', { path: filePath });

  if (destroyed) {
    return {
      setZoom() {},
      goToPage() {},
      getContextSummary: () => '',
      scrollToHeading() {},
      destroy() {},
    };
  }

  if (res.delegateKind === 'docx') {
    return {
      delegatedToDocx: true,
      setZoom() {},
      goToPage() {},
      getContextSummary: () => '',
      scrollToHeading() {},
      destroy() {},
    };
  }

  const noticeText = res.notices?.[0] || '';
  const pagesHtml = (res.pages || [])
    .map(
      (p) => `
      <section class="legacy-a4-page" data-page="${p.pageNumber}">
        ${p.html}
      </section>`,
    )
    .join('');

  container.innerHTML = `
    <div id="legacy-stage-scroll" class="office-scroll">
      ${noticeText ? `<div class="legacy-doc-banner"><span>${escapeHtml(noticeText)}</span><span>${res.pages.length} 页</span></div>` : ''}
      <div id="legacy-pages-host" class="legacy-pages-host">${pagesHtml}</div>
    </div>
  `;

  const scrollEl = container.querySelector<HTMLElement>('#legacy-stage-scroll')!;
  const hostEl = container.querySelector<HTMLElement>('#legacy-pages-host')!;
  const pageEls = Array.from(hostEl.querySelectorAll<HTMLElement>('.legacy-a4-page'));
  const totalPages = Math.max(1, res.pages?.length || 1);

  const applyScale = (force = false) => {
    const firstPage = pageEls[0];
    if (!firstPage) return;
    const avail = Math.max(240, scrollEl.clientWidth - 28);
    if (!force && Math.abs(avail - lastAvailWidth) < 4) return;
    const naturalWidth = firstPage.offsetWidth || 794;
    const scale = currentZoom === 'fit' ? Math.min(1.18, avail / naturalWidth) : currentZoom;
    if (!force && Math.abs(scale - lastAppliedScale) < 0.005) {
      lastAvailWidth = avail;
      return;
    }
    lastAvailWidth = avail;
    lastAppliedScale = scale;
    (hostEl.style as CSSStyleDeclaration & { zoom?: string }).zoom = String(Number(scale.toFixed(3)));
  };

  applyScale(true);
  resizeObs = new ResizeObserver(() => {
    if (destroyed) return;
    window.clearTimeout(resizeTimer);
    resizeTimer = window.setTimeout(() => {
      if (!destroyed) applyScale();
    }, 60);
  });
  resizeObs.observe(scrollEl);

  opts.onOutline?.(
    (res.toc || []).map((t) => ({
      title: t.text,
      level: t.level,
      id: t.id,
    })),
  );
  opts.onPageInfo?.({ currentPage: 1, totalPages });

  let scrollRaf = 0;
  scrollEl.addEventListener(
    'scroll',
    () => {
      if (scrollRaf || destroyed) return;
      scrollRaf = window.requestAnimationFrame(() => {
        scrollRaf = 0;
        if (!pageEls.length) return;
        const scrollRect = scrollEl.getBoundingClientRect();
        const probeY = scrollRect.top + Math.min(140, scrollRect.height * 0.25);
        let bestIdx = 0;
        let bestDist = Infinity;
        pageEls.forEach((sec, idx) => {
          const r = sec.getBoundingClientRect();
          const dist = Math.abs(r.top - probeY);
          if (dist < bestDist) {
            bestDist = dist;
            bestIdx = idx;
          }
        });
        const nextPage = bestIdx + 1;
        if (nextPage !== currentPage) {
          currentPage = nextPage;
          opts.onPageInfo?.({ currentPage, totalPages });
        }
      });
    },
    { passive: true },
  );

  return {
    setZoom(zoom) {
      currentZoom = zoom;
      applyScale(true);
    },
    goToPage(page) {
      if (!pageEls.length) return;
      const clamped = Math.max(1, Math.min(pageEls.length, Math.floor(page)));
      currentPage = clamped;
      pageEls[clamped - 1]?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      opts.onPageInfo?.({ currentPage, totalPages });
    },
    getContextSummary: () => (res.plainText || '').slice(0, 6000),
    scrollToHeading(id) {
      const target = hostEl.querySelector<HTMLElement>(`#${CSS.escape(id)}`);
      target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    },
    destroy() {
      destroyed = true;
      window.clearTimeout(resizeTimer);
      resizeObs?.disconnect();
    },
  };
}
