import * as pdfjsLib from 'pdfjs-dist';
import * as pdfjsWorker from 'pdfjs-dist/build/pdf.worker.mjs';
import type { ViewerMcpClient } from '../mcpClient';

(globalThis as unknown as { pdfjsWorker: unknown }).pdfjsWorker = pdfjsWorker;

export interface OutlineNode {
  title: string;
  level: number;
  page?: number;
  id?: string;
}

export interface PdfMountOptions {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  filePath?: string;
  totalBytes?: number;
  rawBuffer?: ArrayBuffer;
  initialPage?: number;
  initialZoom?: number;
  initialFitMode?: 'width' | 'custom';
  onOutline?: (items: OutlineNode[]) => void;
  onPageChange?: (page: number, totalPages: number) => void;
  onZoomChange?: (zoom: number, fitMode: 'width' | 'custom') => void;
  onStatus?: (msg: string) => void;
}

export interface PdfController {
  goToPage: (page: number) => void;
  setZoom: (zoom: 'fit' | number) => void;
  getZoom: () => 'fit' | number;
  getCurrentPageText: () => Promise<string>;
  getCurrentPage: () => number;
  getTotalPages: () => number;
  destroy: () => void;
}

export function base64ToUint8Array(base64: string): Uint8Array {
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) {
    bytes[i] = bin.charCodeAt(i);
  }
  return bytes;
}

const cmapCache = new Map<string, Promise<Uint8Array>>();
const fontCache = new Map<string, Promise<Uint8Array>>();

function createMcpCMapReaderFactory(mcp: ViewerMcpClient) {
  return class McpCMapReaderFactory {
    private readonly isCompressed: boolean;

    constructor({ isCompressed = true }: { baseUrl?: string | null; isCompressed?: boolean } = {}) {
      this.isCompressed = isCompressed;
    }

    async fetch({ name }: { name: string }): Promise<{ cMapData: Uint8Array; isCompressed: boolean }> {
      if (!name) throw new Error('CMap name required');
      let pending = cmapCache.get(name);
      if (!pending) {
        pending = mcp
          .callTool<{ base64: string }>('app_get_pdf_asset', { kind: 'cmap', name })
          .then((res) => base64ToUint8Array(res.base64));
        cmapCache.set(name, pending);
      }
      const cMapData = await pending;
      return { cMapData, isCompressed: this.isCompressed };
    }
  };
}

function createMcpStandardFontDataFactory(mcp: ViewerMcpClient) {
  return class McpStandardFontDataFactory {
    constructor(_opts: { baseUrl?: string | null } = {}) {}

    async fetch({ filename }: { filename: string }): Promise<Uint8Array> {
      if (!filename) throw new Error('Font filename required');
      let pending = fontCache.get(filename);
      if (!pending) {
        pending = mcp
          .callTool<{ base64: string }>('app_get_pdf_asset', { kind: 'standard_font', name: filename })
          .then((res) => base64ToUint8Array(res.base64));
        fontCache.set(filename, pending);
      }
      return pending;
    }
  };
}

async function loadPdfBytes(
  mcp: ViewerMcpClient,
  filePath: string,
  expectedSize: number,
  onProgress: (loaded: number, total: number) => void,
  signal: AbortSignal,
): Promise<Uint8Array> {
  const chunkBytes = 384 * 1024;
  const parts: Uint8Array[] = [];
  let offset = 0;
  let totalSize = expectedSize;

  while (true) {
    if (signal.aborted) {
      throw new Error('Cancelled');
    }
    const chunk = await mcp.callTool<{
      base64: string;
      offset: number;
      bytesRead: number;
      totalSize: number;
      eof: boolean;
    }>('app_read_file_chunk', {
      path: filePath,
      offset,
      length: chunkBytes,
    });
    totalSize = chunk.totalSize;
    if (chunk.bytesRead > 0) {
      parts.push(base64ToUint8Array(chunk.base64));
      offset += chunk.bytesRead;
      onProgress(offset, totalSize);
    }
    if (chunk.eof || chunk.bytesRead === 0) break;
  }

  const combined = new Uint8Array(offset);
  let pos = 0;
  for (const part of parts) {
    combined.set(part, pos);
    pos += part.byteLength;
  }
  return combined;
}

export async function mountPdfViewer(opts: PdfMountOptions): Promise<PdfController> {
  const { container, mcp } = opts;
  const abort = new AbortController();
  let destroyed = false;
  let pdfDoc: pdfjsLib.PDFDocumentProxy | null = null;
  let observer: IntersectionObserver | null = null;

  container.innerHTML = `
    <div class="subtoolbar">
      <button id="pdf-prev" class="btn-sm" title="上一页">◀</button>
      <input id="pdf-page-input" class="page-input" type="number" min="1" value="${opts.initialPage ?? 1}" />
      <span id="pdf-page-total">/ -</span>
      <button id="pdf-next" class="btn-sm" title="下一页">▶</button>
      <button id="pdf-zoom-out" class="btn-sm" title="缩小">－</button>
      <select id="pdf-zoom-select" title="缩放比例">
        <option value="width">适合宽度</option>
        <option value="0.75">75%</option>
        <option value="1">100%</option>
        <option value="1.25">125%</option>
        <option value="1.5">150%</option>
        <option value="2">200%</option>
      </select>
      <button id="pdf-zoom-in" class="btn-sm" title="放大">＋</button>
      <span class="spacer"></span>
      <input id="pdf-search-input" type="search" placeholder="页内搜索…" style="width:130px" />
      <button id="pdf-search-prev" class="btn-sm" title="上一个匹配">↑</button>
      <button id="pdf-search-next" class="btn-sm" title="下一个匹配">↓</button>
      <span id="pdf-search-count" style="font-size:11px;color:var(--fg-2)"></span>
    </div>
    <div id="pdf-scroll" class="pdf-scroll">
      <div class="pdf-progress">
        <div id="pdf-progress-label">正在加载 PDF…</div>
        <div class="progress-bar-track"><div id="pdf-progress-fill" class="progress-bar-fill"></div></div>
      </div>
    </div>
  `;

  const scrollEl = container.querySelector<HTMLElement>('#pdf-scroll')!;
  const progressLabel = container.querySelector<HTMLElement>('#pdf-progress-label')!;
  const progressFill = container.querySelector<HTMLElement>('#pdf-progress-fill')!;
  const pageInput = container.querySelector<HTMLInputElement>('#pdf-page-input')!;
  const pageTotalEl = container.querySelector<HTMLElement>('#pdf-page-total')!;
  const zoomSelect = container.querySelector<HTMLSelectElement>('#pdf-zoom-select')!;
  const searchInput = container.querySelector<HTMLInputElement>('#pdf-search-input')!;
  const searchCountEl = container.querySelector<HTMLElement>('#pdf-search-count')!;

  let currentPage = Math.max(1, opts.initialPage ?? 1);
  let totalPages = 0;
  let fitMode: 'width' | 'custom' = opts.initialFitMode ?? 'width';
  let customZoom = opts.initialZoom ?? 1.15;
  let pageWrappers: HTMLElement[] = [];
  let renderedPages = new Map<number, number>(); // pageNum -> scale rendered
  let renderingSet = new Set<number>();
  let resizeTimer: number | undefined;
  let lastObservedWidth = 0;
  let basePageWidth = 595;
  let basePageHeight = 842;
  let searchHits: number[] = [];
  let searchHitCursor = -1;

  zoomSelect.value = fitMode === 'width' ? 'width' : String(customZoom);

  function computeScale(): number {
    if (fitMode === 'width') {
      const avail = Math.max(240, scrollEl.clientWidth - 28);
      return Math.min(2.5, Math.max(0.5, avail / basePageWidth));
    }
    return customZoom;
  }

  function layoutPlaceholders(scale: number): void {
    for (const wrapper of pageWrappers) {
      const w = Number(wrapper.dataset.baseWidth || basePageWidth) * scale;
      const h = Number(wrapper.dataset.baseHeight || basePageHeight) * scale;
      wrapper.style.width = `${Math.round(w)}px`;
      wrapper.style.height = `${Math.round(h)}px`;
      wrapper.style.setProperty('--scale-factor', String(scale));
    }
  }

  function evictDistantPages(centerPage: number): void {
    for (const [p] of renderedPages) {
      if (Math.abs(p - centerPage) <= 4 || renderingSet.has(p)) continue;
      const wrapper = pageWrappers[p - 1];
      if (wrapper) {
        wrapper.querySelectorAll('canvas').forEach((c) => {
          c.width = 0;
          c.height = 0;
        });
        wrapper.innerHTML = '';
      }
      renderedPages.delete(p);
    }
  }

  async function renderPage(pageNum: number): Promise<void> {
    if (destroyed || !pdfDoc) return;
    const scale = computeScale();
    const prevScale = renderedPages.get(pageNum);
    if (prevScale !== undefined && Math.abs(prevScale - scale) < 0.01) return;
    if (renderingSet.has(pageNum)) return;

    renderingSet.add(pageNum);
    try {
      const page = await pdfDoc.getPage(pageNum);
      if (destroyed) return;

      const viewport = page.getViewport({ scale });
      const wrapper = pageWrappers[pageNum - 1];
      if (!wrapper) return;

      wrapper.dataset.baseWidth = String(viewport.width / scale);
      wrapper.dataset.baseHeight = String(viewport.height / scale);
      wrapper.style.width = `${Math.round(viewport.width)}px`;
      wrapper.style.height = `${Math.round(viewport.height)}px`;
      wrapper.style.setProperty('--scale-factor', String(scale));
      wrapper.querySelectorAll('canvas').forEach((oldCanvas) => {
        oldCanvas.width = 0;
        oldCanvas.height = 0;
      });
      wrapper.innerHTML = '';

      const canvas = document.createElement('canvas');
      const dpr = Math.min(window.devicePixelRatio || 1, 1.5);
      canvas.width = Math.floor(viewport.width * dpr);
      canvas.height = Math.floor(viewport.height * dpr);
      canvas.style.width = `${Math.round(viewport.width)}px`;
      canvas.style.height = `${Math.round(viewport.height)}px`;
      wrapper.appendChild(canvas);

      const textLayerDiv = document.createElement('div');
      textLayerDiv.className = 'textLayer';
      wrapper.appendChild(textLayerDiv);

      const ctx = canvas.getContext('2d');
      if (ctx) {
        const transform = dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined;
        await page.render({ canvasContext: ctx, viewport, transform }).promise;
      }

      if (destroyed) return;

      const textLayer = new pdfjsLib.TextLayer({
        textContentSource: page.streamTextContent(),
        container: textLayerDiv,
        viewport,
      });
      await textLayer.render();

      renderedPages.set(pageNum, scale);
      evictDistantPages(currentPage);
    } catch {
      // Ignore cancelled page renders
    } finally {
      renderingSet.delete(pageNum);
    }
  }

  function setupObserver(): void {
    observer?.disconnect();
    observer = new IntersectionObserver(
      (entries) => {
        if (destroyed) return;
        let topVisiblePage = currentPage;
        let bestTop = Infinity;

        for (const entry of entries) {
          const p = Number((entry.target as HTMLElement).dataset.page);
          if (entry.isIntersecting) {
            void renderPage(p);
            if (p > 1) void renderPage(p - 1);
            if (p < totalPages) void renderPage(p + 1);
            const rectTop = Math.abs(entry.boundingClientRect.top - scrollEl.getBoundingClientRect().top);
            if (rectTop < bestTop) {
              bestTop = rectTop;
              topVisiblePage = p;
            }
          }
        }
        if (topVisiblePage !== currentPage && bestTop < Infinity) {
          currentPage = topVisiblePage;
          pageInput.value = String(currentPage);
          opts.onPageChange?.(currentPage, totalPages);
        }
      },
      { root: scrollEl, rootMargin: '250px 0px', threshold: [0.05, 0.35] },
    );
    for (const wrapper of pageWrappers) {
      observer.observe(wrapper);
    }
  }

  function goToPage(page: number): void {
    if (!totalPages) return;
    const clamped = Math.max(1, Math.min(totalPages, Math.floor(page)));
    currentPage = clamped;
    pageInput.value = String(clamped);
    const target = pageWrappers[clamped - 1];
    if (target) {
      target.scrollIntoView({ behavior: 'auto', block: 'start' });
      void renderPage(clamped);
    }
    opts.onPageChange?.(currentPage, totalPages);
  }

  function reRenderVisible(): void {
    const scale = computeScale();
    renderedPages.clear();
    layoutPlaceholders(scale);
    void renderPage(currentPage);
    if (currentPage > 1) void renderPage(currentPage - 1);
    if (currentPage < totalPages) void renderPage(currentPage + 1);
    opts.onZoomChange?.(scale, fitMode);
  }

  async function extractOutline(doc: pdfjsLib.PDFDocumentProxy): Promise<OutlineNode[]> {
    try {
      const outline = await doc.getOutline();
      if (!outline?.length) return [];
      const out: OutlineNode[] = [];
      const walk = async (items: typeof outline, level: number) => {
        for (const item of items) {
          let pageNum: number | undefined;
          try {
            let dest = item.dest;
            if (typeof dest === 'string') {
              dest = await doc.getDestination(dest);
            }
            if (Array.isArray(dest) && dest[0]) {
              const idx = await doc.getPageIndex(dest[0] as Parameters<typeof doc.getPageIndex>[0]);
              pageNum = idx + 1;
            }
          } catch {
            // Ignore unresolvable dest
          }
          out.push({ title: item.title || 'Untitled', level, page: pageNum });
          if (item.items?.length && out.length < 300) {
            await walk(item.items, Math.min(level + 1, 4));
          }
        }
      };
      await walk(outline, 1);
      return out;
    } catch {
      return [];
    }
  }

  async function runSearch(query: string): Promise<void> {
    for (const w of pageWrappers) w.classList.remove('search-hit');
    searchHits = [];
    searchHitCursor = -1;
    const q = query.trim().toLowerCase();
    if (!q || !pdfDoc) {
      searchCountEl.textContent = '';
      return;
    }
    searchCountEl.textContent = '搜索中…';
    for (let p = 1; p <= totalPages; p++) {
      if (destroyed) return;
      const page = await pdfDoc.getPage(p);
      const tc = await page.getTextContent();
      const text = tc.items
        .map((it) => ('str' in it ? it.str : ''))
        .join(' ')
        .toLowerCase();
      if (text.includes(q)) {
        searchHits.push(p);
        pageWrappers[p - 1]?.classList.add('search-hit');
      }
    }
    if (searchHits.length > 0) {
      searchHitCursor = 0;
      searchCountEl.textContent = `1/${searchHits.length} 页`;
      goToPage(searchHits[0]);
    } else {
      searchCountEl.textContent = '0 匹配';
    }
  }

  // Event wiring
  container.querySelector('#pdf-prev')!.addEventListener('click', () => goToPage(currentPage - 1));
  container.querySelector('#pdf-next')!.addEventListener('click', () => goToPage(currentPage + 1));
  pageInput.addEventListener('change', () => goToPage(Number(pageInput.value) || 1));

  zoomSelect.addEventListener('change', () => {
    if (zoomSelect.value === 'width') {
      fitMode = 'width';
    } else {
      fitMode = 'custom';
      customZoom = Number(zoomSelect.value) || 1;
    }
    reRenderVisible();
  });

  container.querySelector('#pdf-zoom-in')!.addEventListener('click', () => {
    fitMode = 'custom';
    customZoom = Math.min(3, Number((computeScale() + 0.2).toFixed(2)));
    zoomSelect.value = String(customZoom);
    reRenderVisible();
  });

  container.querySelector('#pdf-zoom-out')!.addEventListener('click', () => {
    fitMode = 'custom';
    customZoom = Math.max(0.5, Number((computeScale() - 0.2).toFixed(2)));
    zoomSelect.value = String(customZoom);
    reRenderVisible();
  });

  let searchTimer: number | undefined;
  searchInput.addEventListener('input', () => {
    window.clearTimeout(searchTimer);
    searchTimer = window.setTimeout(() => void runSearch(searchInput.value), 300);
  });

  container.querySelector('#pdf-search-next')!.addEventListener('click', () => {
    if (!searchHits.length) return;
    searchHitCursor = (searchHitCursor + 1) % searchHits.length;
    searchCountEl.textContent = `${searchHitCursor + 1}/${searchHits.length} 页`;
    goToPage(searchHits[searchHitCursor]);
  });

  container.querySelector('#pdf-search-prev')!.addEventListener('click', () => {
    if (!searchHits.length) return;
    searchHitCursor = (searchHitCursor - 1 + searchHits.length) % searchHits.length;
    searchCountEl.textContent = `${searchHitCursor + 1}/${searchHits.length} 页`;
    goToPage(searchHits[searchHitCursor]);
  });

  const resizeObserver = new ResizeObserver(() => {
    if (fitMode !== 'width' || totalPages === 0 || destroyed) return;
    const w = scrollEl.clientWidth;
    if (w <= 0 || Math.abs(w - lastObservedWidth) < 6) return;
    window.clearTimeout(resizeTimer);
    resizeTimer = window.setTimeout(() => {
      if (fitMode === 'width' && totalPages > 0 && !destroyed) {
        lastObservedWidth = scrollEl.clientWidth;
        reRenderVisible();
      }
    }, 180);
  });
  resizeObserver.observe(scrollEl);

  try {
    let bytes: Uint8Array;
    if (opts.rawBuffer) {
      bytes = new Uint8Array(opts.rawBuffer);
      progressFill.style.width = '100%';
    } else if (opts.filePath) {
      bytes = await loadPdfBytes(
        mcp,
        opts.filePath,
        opts.totalBytes ?? 0,
        (loaded, total) => {
          const pct = total > 0 ? Math.min(100, Math.round((loaded / total) * 100)) : 50;
          progressFill.style.width = `${pct}%`;
          progressLabel.textContent = `正在加载 PDF… ${Math.round(loaded / 1024)} KB`;
        },
        abort.signal,
      );
    } else {
      throw new Error('Missing PDF source');
    }

    if (destroyed) {
      throw new Error('Viewer destroyed');
    }

    progressLabel.textContent = '正在解析页面…';

    const loadingTask = pdfjsLib.getDocument({
      data: bytes,
      useWorkerFetch: false,
      cMapUrl: 'mcp://cmaps/',
      cMapPacked: true,
      standardFontDataUrl: 'mcp://standard_fonts/',
      CMapReaderFactory: createMcpCMapReaderFactory(mcp) as unknown as object,
      StandardFontDataFactory: createMcpStandardFontDataFactory(mcp) as unknown as object,
    });

    pdfDoc = await loadingTask.promise;
    if (destroyed) {
      await pdfDoc.destroy();
      throw new Error('Viewer destroyed');
    }

    totalPages = pdfDoc.numPages;
    pageTotalEl.textContent = `/ ${totalPages}`;
    pageInput.max = String(totalPages);

    const firstPage = await pdfDoc.getPage(1);
    const firstVp = firstPage.getViewport({ scale: 1 });
    basePageWidth = firstVp.width;
    basePageHeight = firstVp.height;
    const initialScale = computeScale();

    scrollEl.innerHTML = '';
    pageWrappers = [];
    for (let p = 1; p <= totalPages; p++) {
      const div = document.createElement('div');
      div.className = 'pdf-page';
      div.dataset.page = String(p);
      div.dataset.baseWidth = String(basePageWidth);
      div.dataset.baseHeight = String(basePageHeight);
      div.style.width = `${Math.round(basePageWidth * initialScale)}px`;
      div.style.height = `${Math.round(basePageHeight * initialScale)}px`;
      div.style.setProperty('--scale-factor', String(initialScale));
      scrollEl.appendChild(div);
      pageWrappers.push(div);
    }

    setupObserver();
    goToPage(Math.min(Math.max(1, currentPage), totalPages));
    lastObservedWidth = scrollEl.clientWidth;
    opts.onOutline?.(await extractOutline(pdfDoc));
  } catch (err) {
    if (!destroyed) {
      scrollEl.innerHTML = `<div class="error-box">PDF 加载失败：${
        err instanceof Error ? err.message : String(err)
      }</div>`;
    }
  }

  return {
    goToPage,
    setZoom(zoom: 'fit' | number) {
      if (zoom === 'fit') {
        fitMode = 'width';
        zoomSelect.value = 'width';
      } else {
        fitMode = 'custom';
        customZoom = Math.max(0.5, Math.min(3, Number(zoom.toFixed(2))));
        zoomSelect.value = String(customZoom);
      }
      if (totalPages > 0) reRenderVisible();
    },
    getZoom: () => (fitMode === 'width' ? 'fit' : customZoom),
    getCurrentPage: () => currentPage,
    getTotalPages: () => totalPages,
    async getCurrentPageText(): Promise<string> {
      if (!pdfDoc || !totalPages) return '';
      const page = await pdfDoc.getPage(currentPage);
      const tc = await page.getTextContent();
      return tc.items.map((it) => ('str' in it ? it.str : '')).join(' ');
    },
    destroy(): void {
      destroyed = true;
      window.clearTimeout(resizeTimer);
      abort.abort();
      observer?.disconnect();
      resizeObserver.disconnect();
      for (const wrapper of pageWrappers) {
        wrapper.querySelectorAll('canvas').forEach((c) => {
          c.width = 0;
          c.height = 0;
        });
      }
      void pdfDoc?.destroy();
      pdfDoc = null;
    },
  };
}
