import type { ViewerMcpClient } from '../mcpClient';
import type { OfficeViewMode } from './docxOfficeViewer';
import type { OutlineNode } from './pdfViewer';

interface PptxSlideModel {
  slideNumber: number;
  title: string;
  widthPt: number;
  heightPt: number;
  shapes: Array<{
    id: string;
    kind: 'text' | 'image' | 'table';
    leftPct: number;
    topPct: number;
    widthPct: number;
    heightPct: number;
    paragraphs?: Array<{ text: string; align?: string; fontSizePt?: number; bold?: boolean; colorHex?: string }>;
    imageDataUri?: string;
    tableRows?: string[][];
  }>;
  plainText: string;
}

export interface SlideController {
  setMode: (mode: OfficeViewMode) => Promise<void>;
  goToSlide: (slideNumber: number) => void;
  getSlideInfo: () => { current: number; total: number };
  getContextSummary: () => string;
  destroy: () => void;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export async function mountSlideOfficeViewer(opts: {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  filePath: string;
  initialSlide?: number;
  mode?: OfficeViewMode;
  legacy?: boolean;
  onOutline?: (items: OutlineNode[]) => void;
  onSlideChange?: (current: number, total: number) => void;
  onWorktreeChanged?: () => void;
  onToast?: (msg: string) => void;
}): Promise<SlideController> {
  const { container, mcp, filePath } = opts;
  let isLegacy = Boolean(opts.legacy);
  let currentMode: OfficeViewMode = isLegacy ? 'print' : (opts.mode ?? 'print');
  let activeSlideNum = opts.initialSlide ?? 1;
  let slides: PptxSlideModel[] = [];
  let destroyed = false;

  async function renderSlidesView() {
    container.innerHTML = `<div class="office-scroll"><div class="office-loading">正在解析 PowerPoint 幻灯片…</div></div>`;
    let res: {
      slideCount: number;
      widthPt: number;
      heightPt: number;
      slides: PptxSlideModel[];
      hasDraft?: boolean;
      delegateKind?: 'pptx';
    };
    if (isLegacy) {
      res = await mcp.callTool('app_render_legacy_ppt', { path: filePath });
      if (res.delegateKind === 'pptx') {
        isLegacy = false;
        res = await mcp.callTool('app_render_pptx', { path: filePath, useDraft: true });
      }
    } else {
      res = await mcp.callTool('app_render_pptx', { path: filePath, useDraft: true });
    }
    if (destroyed) return;

    slides = res.slides;
    activeSlideNum = Math.max(1, Math.min(slides.length || 1, activeSlideNum));
    opts.onSlideChange?.(activeSlideNum, slides.length || 1);

    opts.onOutline?.(
      slides.map((s) => ({
        title: `${s.slideNumber}. ${s.title}`,
        level: 1,
        page: s.slideNumber,
      })),
    );

    const activeSlide = slides[activeSlideNum - 1];
    const ratio = res.heightPt > 0 ? (res.heightPt / res.widthPt) * 100 : 56.25;

    const renderSlideDom = (s: PptxSlideModel, isThumb = false) => {
      const shapesHtml = s.shapes
        .map((sh) => {
          const boxStyle = `left:${sh.leftPct}%;top:${sh.topPct}%;width:${sh.widthPct}%;height:${sh.heightPct}%;`;
          if (sh.kind === 'image' && sh.imageDataUri) {
            return `<img class="slide-shape" style="${boxStyle}object-fit:contain;" src="${sh.imageDataUri}" alt="" />`;
          }
          if (sh.kind === 'table' && sh.tableRows) {
            const tbl = sh.tableRows
              .map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(c)}</td>`).join('')}</tr>`)
              .join('');
            return `<div class="slide-shape slide-table-wrap" style="${boxStyle}"><table>${tbl}</table></div>`;
          }
          const paras = (sh.paragraphs || [])
            .map((p) => {
              const fontPx = isThumb
                ? Math.max(4, Math.round((p.fontSizePt || 16) * 0.16))
                : Math.max(10, Math.round((p.fontSizePt || 18) * 0.92));
              return `<div style="text-align:${p.align || 'left'};font-size:${fontPx}px;font-weight:${
                p.bold ? 700 : 400
              };color:${p.colorHex || '#111827'};line-height:1.35;margin-bottom:4px;">${escapeHtml(p.text)}</div>`;
            })
            .join('');
          return `<div class="slide-shape" style="${boxStyle}overflow:hidden;">${paras}</div>`;
        })
        .join('');
      return `<div class="slide-canvas" style="padding-bottom:${ratio}%"><div class="slide-inner">${shapesHtml}</div></div>`;
    };

    const thumbsHtml = slides
      .map(
        (s) => `
        <button class="slide-thumb ${s.slideNumber === activeSlideNum ? 'active' : ''}" data-slide="${s.slideNumber}">
          <span class="slide-thumb-num">${s.slideNumber}</span>
          <div class="slide-thumb-box">${renderSlideDom(s, true)}</div>
        </button>`,
      )
      .join('');

    const editToolbar =
      currentMode === 'edit' && activeSlide
        ? `
        <div class="slide-edit-panel">
          <input id="sl-title" class="path-input" value="${escapeHtml(activeSlide.title)}" placeholder="幻灯片标题…" />
          <textarea id="sl-body" class="edit-textarea" rows="3" placeholder="每行一条要点…">${escapeHtml(
            activeSlide.plainText
              .split('\n')
              .slice(1)
              .join('\n'),
          )}</textarea>
          <div style="display:flex;gap:6px">
            <button id="sl-save-draft" class="btn-sm primary">保存本页到草稿</button>
            <button id="sl-add-slide" class="btn-sm">+ 新建下一页</button>
          </div>
        </div>`
        : '';

    container.innerHTML = `
      <div class="slide-workbench">
        <div class="slide-strip">${thumbsHtml}</div>
        <div class="slide-main">
          ${editToolbar}
          <div class="office-scroll" style="padding:20px">
            <div class="slide-stage-wrap">
              ${activeSlide ? renderSlideDom(activeSlide, false) : '<div class="empty-hint">空演示文稿</div>'}
            </div>
          </div>
        </div>
      </div>
    `;

    container.querySelectorAll<HTMLButtonElement>('.slide-thumb').forEach((btn) => {
      btn.addEventListener('click', () => {
        activeSlideNum = Number(btn.dataset.slide) || 1;
        void renderSlidesView();
      });
    });

    if (currentMode === 'edit' && activeSlide) {
      container.querySelector('#sl-save-draft')!.addEventListener('click', async () => {
        const title = container.querySelector<HTMLInputElement>('#sl-title')!.value;
        const bodyLines = container.querySelector<HTMLTextAreaElement>('#sl-body')!.value.split(/\r?\n/);
        await mcp.callTool('app_save_pptx_edits', {
          path: filePath,
          mode: 'draft',
          summary: `编辑第 ${activeSlideNum} 页幻灯片`,
          updateSlides: [{ slideNumber: activeSlideNum, title, bodyLines }],
        });
        opts.onToast?.('幻灯片改动已暂存到 Worktree');
        opts.onWorktreeChanged?.();
        await renderSlidesView();
      });

      container.querySelector('#sl-add-slide')!.addEventListener('click', async () => {
        await mcp.callTool('app_save_pptx_edits', {
          path: filePath,
          mode: 'draft',
          summary: '新增幻灯片',
          appendSlides: [{ title: `第 ${slides.length + 1} 页`, bullets: ['点击编辑模式修改内容'] }],
        });
        activeSlideNum = slides.length + 1;
        opts.onWorktreeChanged?.();
        await renderSlidesView();
      });
    }
  }

  async function renderCompareView() {
    container.innerHTML = `<div class="office-scroll"><div class="office-loading">正在计算幻灯片差异…</div></div>`;
    const diff = await mcp.callTool<{
      hasDraft: boolean;
      changedCount?: number;
      changedSlides?: Array<{ slideNumber: number; before: string; after: string }>;
    }>('app_worktree_action', { action: 'diff', path: filePath });
    if (destroyed) return;
    if (!diff.hasDraft) {
      container.innerHTML = `<div class="empty-hint" style="padding:36px">当前演示文稿没有待审核的 Worktree 更改。</div>`;
      return;
    }
    const items = (diff.changedSlides || [])
      .map(
        (s) => `
        <div class="diff-card">
          <div class="diff-meta">第 ${s.slideNumber} 页</div>
          <div class="diff-before">- ${escapeHtml(s.before || '(空)')}</div>
          <div class="diff-after">+ ${escapeHtml(s.after || '(空)')}</div>
        </div>`,
      )
      .join('');
    container.innerHTML = `
      <div class="subtoolbar">
        <span>PPT Worktree 变更对比（${diff.changedCount ?? 0} 页变化）</span>
        <span class="spacer"></span>
        <button id="sl-back" class="btn-sm primary">返回幻灯片</button>
      </div>
      <div class="editor-canvas"><div class="editor-sheet">${items}</div></div>
    `;
    container.querySelector('#sl-back')!.addEventListener('click', () => void setMode('print'));
  }

  async function setMode(mode: OfficeViewMode) {
    if (isLegacy && mode !== 'print') {
      currentMode = 'print';
      opts.onToast?.('旧版 .ppt 暂为只读预览，可通过右上角「打开方式」用 PowerPoint / WPS 编辑');
      return;
    }
    currentMode = mode;
    if (mode === 'compare') await renderCompareView();
    else await renderSlidesView();
  }

  await setMode(currentMode);

  return {
    setMode,
    goToSlide(num) {
      activeSlideNum = Math.max(1, Math.min(slides.length || 1, num));
      opts.onSlideChange?.(activeSlideNum, slides.length || 1);
      void renderSlidesView();
    },
    getSlideInfo: () => ({ current: activeSlideNum, total: slides.length || 1 }),
    getContextSummary() {
      return slides.map((s) => `[Slide ${s.slideNumber}] ${s.plainText}`).join('\n\n').slice(0, 6000);
    },
    destroy() {
      destroyed = true;
    },
  };
}
