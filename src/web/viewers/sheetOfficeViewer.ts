import type { ViewerMcpClient } from '../mcpClient';
import type { OfficeViewMode } from './docxOfficeViewer';

export interface SheetMountOptions {
  container: HTMLElement;
  mcp: ViewerMcpClient;
  filePath: string;
  initialSheet?: string;
  initialPage?: number;
  mode?: OfficeViewMode;
  onSheetChange?: (sheet: string) => void;
  onWorktreeChanged?: () => void;
  onToast?: (msg: string) => void;
}

export interface SheetController {
  setMode: (mode: OfficeViewMode) => Promise<void>;
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

function colIndexToName(idx: number): string {
  let n = idx;
  let s = '';
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function colNameToIndex(name: string): number {
  let res = 0;
  for (const ch of name.toUpperCase()) {
    res = res * 26 + (ch.charCodeAt(0) - 64);
  }
  return res - 1;
}

function evalSimpleFormula(
  formula: string,
  cellLookup: (addr: string) => string,
): string {
  const expr = formula.trim().slice(1).trim();
  const fnMatch = /^(SUM|AVERAGE|AVG|COUNT|MIN|MAX)\(([^)]+)\)$/i.exec(expr);
  if (fnMatch) {
    const fn = fnMatch[1].toUpperCase();
    const arg = fnMatch[2].trim();
    const nums: number[] = [];
    for (const part of arg.split(',')) {
      const range = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(part.trim());
      if (range) {
        const c1 = colNameToIndex(range[1]);
        const r1 = Number(range[2]);
        const c2 = colNameToIndex(range[3]);
        const r2 = Number(range[4]);
        for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
          for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
            const v = Number(cellLookup(`${colIndexToName(c)}${r}`));
            if (!Number.isNaN(v)) nums.push(v);
          }
        }
      } else {
        const v = Number(cellLookup(part.trim().toUpperCase()) || part.trim());
        if (!Number.isNaN(v)) nums.push(v);
      }
    }
    if (fn === 'COUNT') return String(nums.length);
    if (!nums.length) return '0';
    if (fn === 'SUM') return String(nums.reduce((a, b) => a + b, 0));
    if (fn === 'AVERAGE' || fn === 'AVG') return String(nums.reduce((a, b) => a + b, 0) / nums.length);
    if (fn === 'MIN') return String(Math.min(...nums));
    if (fn === 'MAX') return String(Math.max(...nums));
  }
  return formula;
}

export async function mountSheetOfficeViewer(opts: SheetMountOptions): Promise<SheetController> {
  const { container, mcp, filePath } = opts;
  let activeSheet = opts.initialSheet;
  let page = opts.initialPage ?? 1;
  let search = '';
  let currentMode: OfficeViewMode = opts.mode ?? 'print';
  let destroyed = false;

  let currentData: {
    sheetNames: string[];
    activeSheet: string;
    columns: string[];
    rows: Array<{ rowNumber: number; cells: string[] }>;
    formulas: Record<string, string>;
    totalRows: number;
    filteredRows: number;
    totalCols: number;
    page: number;
    totalPages: number;
    hasDraft?: boolean;
  } | null = null;

  const pendingEdits = new Map<string, string>(); // "sheet:A1" -> value or =formula
  let selectedAddr = 'A1';

  function currentRawValue(addr: string, fallbackDisplay: string): string {
    const key = `${activeSheet || ''}:${addr}`;
    if (pendingEdits.has(key)) return pendingEdits.get(key)!;
    if (currentData?.formulas?.[addr]) return currentData.formulas[addr];
    return fallbackDisplay;
  }

  async function savePending(mode: 'draft' | 'direct') {
    if (pendingEdits.size === 0) {
      opts.onToast?.('没有待保存的单元格改动');
      return;
    }
    const setCells = Array.from(pendingEdits.entries()).map(([k, val]) => {
      const colon = k.indexOf(':');
      const sheet = k.slice(0, colon);
      const cell = k.slice(colon + 1);
      if (val.startsWith('=')) {
        const computed = evalSimpleFormula(val, (a) => {
          const rMatch = /^([A-Z]+)(\d+)$/.exec(a);
          if (!rMatch || !currentData) return '0';
          const cIdx = colNameToIndex(rMatch[1]);
          const rNum = Number(rMatch[2]);
          const rowObj = currentData.rows.find((r) => r.rowNumber === rNum);
          return rowObj?.cells[cIdx] || '0';
        });
        const num = Number(computed);
        return { sheet, cell, formula: val, value: Number.isNaN(num) ? 0 : num };
      }
      const asNum = val.trim() !== '' && !Number.isNaN(Number(val)) ? Number(val) : val;
      return { sheet, cell, value: asNum };
    });

    await mcp.callTool('app_save_spreadsheet_edits', {
      path: filePath,
      mode,
      summary: `修改 ${setCells.length} 个单元格`,
      setCells,
    });
    pendingEdits.clear();
    opts.onToast?.(mode === 'draft' ? '表格更改已暂存到 Worktree' : '已保存表格并备份旧版本');
    opts.onWorktreeChanged?.();
    await loadGrid();
  }

  async function renderCompare() {
    container.innerHTML = `<div class="office-scroll"><div class="office-loading">正在计算表格变更差异…</div></div>`;
    const diff = await mcp.callTool<{
      hasDraft: boolean;
      summary?: string;
      changedCount?: number;
      changedCells?: Array<{ sheet: string; cell: string; before: string; after: string }>;
    }>('app_worktree_action', { action: 'diff', path: filePath });
    if (destroyed) return;
    if (!diff.hasDraft) {
      container.innerHTML = `<div class="empty-hint" style="padding:36px">当前表格没有待审核的 Worktree 更改。</div>`;
      return;
    }
    const rows = (diff.changedCells || [])
      .map(
        (c) => `
        <tr>
          <td>${escapeHtml(c.sheet)}</td>
          <td><code>${escapeHtml(c.cell)}</code></td>
          <td class="diff-before">${escapeHtml(c.before || '(空)')}</td>
          <td class="diff-after">${escapeHtml(c.after || '(空)')}</td>
        </tr>`,
      )
      .join('');
    container.innerHTML = `
      <div class="subtoolbar">
        <span>表格 Worktree 差异（${diff.changedCount ?? 0} 个单元格）</span>
        <span class="spacer"></span>
        <button id="ss-back-grid" class="btn-sm primary">返回表格视图</button>
      </div>
      <div class="sheet-wrap">
        <table class="sheet-table">
          <thead><tr><th>Sheet</th><th>单元格</th><th>原值</th><th>新值</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    `;
    container.querySelector('#ss-back-grid')!.addEventListener('click', () => void setMode('print'));
  }

  async function loadGrid() {
    container.innerHTML = `
      <div class="subtoolbar">
        <span id="fx-addr" class="fx-addr">${selectedAddr}</span>
        <span class="fx-label">fx</span>
        <input id="fx-input" class="path-input" placeholder="输入值或公式（如 =SUM(B2:B5)）…" />
        <input id="ss-search" type="search" placeholder="筛选行…" style="width:110px" value="${escapeHtml(search)}" />
        <button id="ss-add-row" class="btn-sm">+ 行</button>
        <button id="ss-add-sheet" class="btn-sm">+ Sheet</button>
        <button id="ss-save-draft" class="btn-sm primary">暂存</button>
        <button id="ss-save-direct" class="btn-sm">保存</button>
      </div>
      <div id="ss-wrap" class="sheet-wrap"><div class="empty-hint">正在加载表格…</div></div>
      <div class="sheet-bottom-bar">
        <div id="ss-sheets" class="sheet-tabs-inline"></div>
        <span class="spacer"></span>
        <button id="ss-prev" class="btn-sm">◀</button>
        <span id="ss-page-info" style="font-size:11px;color:var(--fg-2)">1/1</span>
        <button id="ss-next" class="btn-sm">▶</button>
      </div>
    `;

    const res = await mcp.callTool<NonNullable<typeof currentData>>('app_render_spreadsheet', {
      path: filePath,
      sheetName: activeSheet,
      page,
      pageSize: 150,
      search,
      useDraft: true,
    });
    if (destroyed) return;
    currentData = res;
    activeSheet = res.activeSheet;
    page = res.page;

    const wrapEl = container.querySelector<HTMLElement>('#ss-wrap')!;
    const fxAddr = container.querySelector<HTMLElement>('#fx-addr')!;
    const fxInput = container.querySelector<HTMLInputElement>('#fx-input')!;
    const sheetsEl = container.querySelector<HTMLElement>('#ss-sheets')!;
    const pageInfoEl = container.querySelector<HTMLElement>('#ss-page-info')!;
    const searchInput = container.querySelector<HTMLInputElement>('#ss-search')!;

    pageInfoEl.textContent = `${res.page}/${res.totalPages} (${res.filteredRows}行)`;

    sheetsEl.innerHTML = '';
    for (const name of res.sheetNames) {
      const btn = document.createElement('button');
      btn.className = `sheet-tab ${name === res.activeSheet ? 'active' : ''}`;
      btn.textContent = name;
      btn.addEventListener('click', () => {
        activeSheet = name;
        page = 1;
        opts.onSheetChange?.(name);
        void loadGrid();
      });
      sheetsEl.appendChild(btn);
    }

    const cols = res.columns.length ? res.columns : ['A', 'B', 'C', 'D'];
    const ths = cols.map((c) => `<th>${c}</th>`).join('');
    const trs = res.rows
      .map((r) => {
        const tds = cols
          .map((col, cIdx) => {
            const addr = `${col}${r.rowNumber}`;
            const rawDiskVal = r.cells[cIdx] ?? '';
            const edited = pendingEdits.get(`${activeSheet}:${addr}`);
            let displayVal = rawDiskVal;
            if (edited !== undefined) {
              displayVal = edited.startsWith('=')
                ? evalSimpleFormula(edited, (a) => {
                    const m = /^([A-Z]+)(\d+)$/.exec(a);
                    if (!m) return '0';
                    const rr = res.rows.find((x) => x.rowNumber === Number(m[2]));
                    return rr?.cells[colNameToIndex(m[1])] || '0';
                  })
                : edited;
            }
            const dirtyCls = edited !== undefined ? 'cell-dirty' : '';
            const selCls = addr === selectedAddr ? 'cell-selected' : '';
            return `<td class="grid-cell ${dirtyCls} ${selCls}" data-addr="${addr}" data-disp="${escapeHtml(displayVal)}">${escapeHtml(displayVal)}</td>`;
          })
          .join('');
        return `<tr><td class="row-num">${r.rowNumber}</td>${tds}</tr>`;
      })
      .join('');

    wrapEl.innerHTML = `<table class="sheet-table univer-grid"><thead><tr><th class="row-num"></th>${ths}</tr></thead><tbody>${trs}</tbody></table>`;

    const selectCell = (td: HTMLTableCellElement, startEditing = false) => {
      wrapEl.querySelectorAll('.cell-selected').forEach((el) => el.classList.remove('cell-selected'));
      td.classList.add('cell-selected');
      selectedAddr = td.dataset.addr || 'A1';
      fxAddr.textContent = selectedAddr;
      const raw = currentRawValue(selectedAddr, td.dataset.disp || '');
      fxInput.value = raw;

      if (startEditing) {
        const input = document.createElement('input');
        input.className = 'inline-cell-editor';
        input.value = raw;
        td.textContent = '';
        td.appendChild(input);
        input.focus();
        input.select();
        const commit = () => {
          const val = input.value;
          pendingEdits.set(`${activeSheet}:${selectedAddr}`, val);
          fxInput.value = val;
          td.classList.add('cell-dirty');
          const disp = val.startsWith('=')
            ? evalSimpleFormula(val, (a) => {
                const m = /^([A-Z]+)(\d+)$/.exec(a);
                if (!m) return '0';
                const rr = res.rows.find((x) => x.rowNumber === Number(m[2]));
                return rr?.cells[colNameToIndex(m[1])] || '0';
              })
            : val;
          td.dataset.disp = disp;
          td.textContent = disp;
        };
        input.addEventListener('blur', commit);
        input.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') input.blur();
          if (e.key === 'Escape') {
            input.removeEventListener('blur', commit);
            td.textContent = td.dataset.disp || '';
          }
        });
      }
    };

    wrapEl.querySelectorAll<HTMLTableCellElement>('td.grid-cell').forEach((td) => {
      td.addEventListener('click', () => selectCell(td, false));
      td.addEventListener('dblclick', () => selectCell(td, true));
    });

    fxInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        pendingEdits.set(`${activeSheet}:${selectedAddr}`, fxInput.value);
        const td = wrapEl.querySelector<HTMLTableCellElement>(`td.grid-cell[data-addr="${selectedAddr}"]`);
        if (td) {
          td.classList.add('cell-dirty');
          const disp = fxInput.value.startsWith('=')
            ? evalSimpleFormula(fxInput.value, (a) => {
                const m = /^([A-Z]+)(\d+)$/.exec(a);
                if (!m) return '0';
                const rr = res.rows.find((x) => x.rowNumber === Number(m[2]));
                return rr?.cells[colNameToIndex(m[1])] || '0';
              })
            : fxInput.value;
          td.dataset.disp = disp;
          td.textContent = disp;
        }
      }
    });

    let timer: number | undefined;
    searchInput.addEventListener('input', () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => {
        search = searchInput.value;
        page = 1;
        void loadGrid();
      }, 280);
    });

    container.querySelector('#ss-prev')!.addEventListener('click', () => {
      if (page > 1) {
        page -= 1;
        void loadGrid();
      }
    });
    container.querySelector('#ss-next')!.addEventListener('click', () => {
      if (currentData && page < currentData.totalPages) {
        page += 1;
        void loadGrid();
      }
    });

    container.querySelector('#ss-add-row')!.addEventListener('click', async () => {
      const blanks = Array.from({ length: Math.max(1, res.totalCols) }, () => '');
      await mcp.callTool('app_save_spreadsheet_edits', {
        path: filePath,
        mode: 'draft',
        summary: `在 ${activeSheet} 末尾新增空白行`,
        appendRows: [{ sheet: activeSheet, rows: [blanks] }],
      });
      opts.onWorktreeChanged?.();
      await loadGrid();
    });

    container.querySelector('#ss-add-sheet')!.addEventListener('click', async () => {
      const nextName = `Sheet${res.sheetNames.length + 1}`;
      await mcp.callTool('app_save_spreadsheet_edits', {
        path: filePath,
        mode: 'draft',
        summary: `新增工作表 ${nextName}`,
        addSheets: [nextName],
      });
      activeSheet = nextName;
      opts.onWorktreeChanged?.();
      await loadGrid();
    });

    container.querySelector('#ss-save-draft')!.addEventListener('click', () => void savePending('draft'));
    container.querySelector('#ss-save-direct')!.addEventListener('click', () => void savePending('direct'));
  }

  async function setMode(mode: OfficeViewMode) {
    currentMode = mode;
    if (mode === 'compare') await renderCompare();
    else await loadGrid();
  }

  await setMode(currentMode);

  return {
    setMode,
    getContextSummary() {
      if (!currentData) return '';
      const preview = currentData.rows
        .slice(0, 40)
        .map((r) => `${r.rowNumber}: ${r.cells.join(' | ')}`)
        .join('\n');
      return `Sheet: ${currentData.activeSheet} (${currentData.totalRows} rows)\n${preview}`;
    },
    destroy() {
      destroyed = true;
    },
  };
}
