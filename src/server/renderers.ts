import fs from 'node:fs/promises';
import path from 'node:path';
import { Marked, Renderer } from 'marked';
import hljs from 'highlight.js';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { resolveInputPath } from './files.js';

export interface TocItem {
  id: string;
  level: number;
  text: string;
}

const IMAGE_MIME_MAP: Record<string, string> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml',
  '.bmp': 'image/bmp',
  '.ico': 'image/x-icon',
};

const EXT_LANG_MAP: Record<string, string> = {
  '.ts': 'typescript',
  '.tsx': 'typescript',
  '.js': 'javascript',
  '.jsx': 'javascript',
  '.mjs': 'javascript',
  '.cjs': 'javascript',
  '.py': 'python',
  '.rs': 'rust',
  '.go': 'go',
  '.java': 'java',
  '.kt': 'kotlin',
  '.c': 'c',
  '.cpp': 'cpp',
  '.h': 'c',
  '.hpp': 'cpp',
  '.cs': 'csharp',
  '.rb': 'ruby',
  '.php': 'php',
  '.sh': 'bash',
  '.bash': 'bash',
  '.ps1': 'powershell',
  '.bat': 'dos',
  '.cmd': 'dos',
  '.sql': 'sql',
  '.css': 'css',
  '.scss': 'scss',
  '.less': 'less',
  '.xml': 'xml',
  '.html': 'xml',
  '.htm': 'xml',
  '.json': 'json',
  '.jsonc': 'json',
  '.yaml': 'yaml',
  '.yml': 'yaml',
  '.toml': 'ini',
  '.ini': 'ini',
  '.md': 'markdown',
};

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function sanitizeHtml(html: string): string {
  return html
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<noscript\b[^>]*>[\s\S]*?<\/noscript>/gi, '')
    .replace(/<(iframe|object|embed|applet|base)\b[^>]*>[\s\S]*?<\/\1>/gi, '')
    .replace(/<(iframe|object|embed|applet|base)\b[^>]*\/?>/gi, '')
    .replace(/\s+on[a-z]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '')
    .replace(/(href|src)\s*=\s*(["'])\s*(?:javascript|vbscript):[\s\S]*?\2/gi, '$1=$2#$2');
}

function stripTags(html: string): string {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

export function extractTocAndAddIds(html: string): { html: string; toc: TocItem[] } {
  const toc: TocItem[] = [];
  let counter = 0;

  const updated = html.replace(/<h([1-4])\b([^>]*)>([\s\S]*?)<\/h\1>/gi, (_full, levelStr, attrs, inner) => {
    const level = Number(levelStr);
    const text = stripTags(inner);
    if (!text) return `<h${level}${attrs}>${inner}</h${level}>`;
    counter += 1;
    const id = `sec-${counter}`;
    const cleanedAttrs = String(attrs).replace(/\s+id\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/i, '');
    toc.push({ id, level, text });
    return `<h${level} id="${id}"${cleanedAttrs}>${inner}</h${level}>`;
  });

  return { html: updated, toc };
}

function isRemoteOrSpecialUri(url: string): boolean {
  return /^(?:https?:|data:|blob:|mailto:|tel:|#|\/\/)/i.test(url.trim());
}

async function inlineLocalImages(
  html: string,
  baseDir: string,
  maxPerImageBytes = 220 * 1024,
  maxTotalBytes = 450 * 1024,
): Promise<string> {
  let remaining = maxTotalBytes;
  const imgRegex = /<img\b([^>]*?)\bsrc\s*=\s*(?:"([^"]*)"|'([^']*)')([^>]*)>/gi;
  const matches = [...html.matchAll(imgRegex)];
  if (matches.length === 0) return html;

  let out = '';
  let lastIndex = 0;

  for (const m of matches) {
    const full = m[0];
    const start = m.index ?? 0;
    out += html.slice(lastIndex, start);
    lastIndex = start + full.length;

    const before = m[1] || '';
    const src = (m[2] ?? m[3] ?? '').trim();
    const after = m[4] || '';

    if (!src || isRemoteOrSpecialUri(src)) {
      out += full;
      continue;
    }

    const cleanRel = decodeURIComponent(src.split(/[?#]/)[0]);
    const resolved = path.resolve(baseDir, cleanRel);
    const ext = path.extname(resolved).toLowerCase();
    const mime = IMAGE_MIME_MAP[ext];
    if (!mime) {
      out += full;
      continue;
    }

    try {
      const stat = await fs.stat(resolved);
      if (stat.isFile() && stat.size <= maxPerImageBytes && stat.size <= remaining) {
        const buf = await fs.readFile(resolved);
        remaining -= buf.byteLength;
        const dataUri = `data:${mime};base64,${buf.toString('base64')}`;
        out += `<img${before} src="${dataUri}"${after}>`;
        continue;
      }
    } catch {
      // Leave unresolved image tag with alt text fallback
    }
    out += `<span class="omitted-badge">[Image unavailable offline: ${escapeHtml(path.basename(cleanRel))}]</span>`;
  }

  out += html.slice(lastIndex);
  return out;
}

function buildMarkdownInstance(): Marked {
  const renderer = new Renderer();
  renderer.code = ({ text, lang }) => {
    const cleanLang = (lang || '').trim().split(/\s+/)[0].toLowerCase();
    let highlighted: string;
    if (cleanLang && hljs.getLanguage(cleanLang)) {
      highlighted = hljs.highlight(text, { language: cleanLang, ignoreIllegals: true }).value;
    } else {
      highlighted = escapeHtml(text);
    }
    const langLabel = cleanLang ? `<span class="code-lang">${escapeHtml(cleanLang)}</span>` : '';
    return `<pre class="hljs-block">${langLabel}<code class="hljs language-${escapeHtml(cleanLang || 'plaintext')}">${highlighted}</code></pre>`;
  };
  return new Marked({ gfm: true, breaks: false, renderer });
}

const markedInstance = buildMarkdownInstance();

export async function renderMarkdown(rawPath: string): Promise<{
  html: string;
  toc: TocItem[];
  truncated: boolean;
  totalBytes: number;
}> {
  const filePath = resolveInputPath(rawPath);
  const stat = await fs.stat(filePath);
  const maxBytes = 5 * 1024 * 1024;
  let truncated = false;
  let content: string;

  if (stat.size > maxBytes) {
    truncated = true;
    const handle = await fs.open(filePath, 'r');
    try {
      const buf = Buffer.allocUnsafe(maxBytes);
      const { bytesRead } = await handle.read(buf, 0, maxBytes, 0);
      content = buf.subarray(0, bytesRead).toString('utf-8');
    } finally {
      await handle.close();
    }
  } else {
    content = await fs.readFile(filePath, 'utf-8');
  }

  const rawHtml = await markedInstance.parse(content);
  const inlined = await inlineLocalImages(rawHtml, path.dirname(filePath));
  const clean = sanitizeHtml(inlined);
  const { html, toc } = extractTocAndAddIds(clean);
  return { html, toc, truncated, totalBytes: stat.size };
}

function makeOmittedImagePlaceholder(label: string): string {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="320" height="44" viewBox="0 0 320 44"><rect width="320" height="44" rx="6" fill="#f3f4f6" stroke="#d1d5db"/><text x="16" y="27" font-family="system-ui,sans-serif" font-size="12" fill="#4b5563">${escapeHtml(label)}</text></svg>`;
  return `data:image/svg+xml;base64,${Buffer.from(svg, 'utf-8').toString('base64')}`;
}

export async function renderDocx(rawPath: string): Promise<{
  html: string;
  toc: TocItem[];
  warnings: string[];
  wordCount: number;
  omittedImages: number;
}> {
  const filePath = resolveInputPath(rawPath);
  let remainingImageBudget = 450 * 1024;
  const maxSingleImage = 220 * 1024;
  let omittedImages = 0;

  const result = await mammoth.convertToHtml(
    { path: filePath },
    {
      convertImage: mammoth.images.imgElement(async (image) => {
        const buffer = await image.read();
        const byteLen = buffer.byteLength;
        if (byteLen <= maxSingleImage && byteLen <= remainingImageBudget) {
          remainingImageBudget -= byteLen;
          return {
            src: `data:${image.contentType};base64,${buffer.toString('base64')}`,
          };
        }
        omittedImages += 1;
        const sizeKb = Math.round(byteLen / 1024);
        return {
          src: makeOmittedImagePlaceholder(`Image omitted (${sizeKb} KB > payload budget)`),
          alt: `Image omitted (${sizeKb} KB)`,
        };
      }),
    },
  );

  const clean = sanitizeHtml(result.value);
  const { html, toc } = extractTocAndAddIds(clean);
  const plainText = stripTags(html);
  const wordCount = plainText ? plainText.split(/\s+/).length : 0;
  const warnings = result.messages.map((m) => m.message).slice(0, 20);

  return { html, toc, warnings, wordCount, omittedImages };
}

function indexToColName(idx: number): string {
  let n = idx;
  let s = '';
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

export async function renderSpreadsheet(
  rawPath: string,
  options: {
    sheetName?: string;
    page?: number;
    pageSize?: number;
    search?: string;
    bufferOverride?: Buffer;
  } = {},
): Promise<{
  sheetNames: string[];
  activeSheet: string;
  columns: string[];
  rows: Array<{ rowNumber: number; cells: string[] }>;
  formulas: Record<string, string>;
  totalRows: number;
  filteredRows: number;
  totalCols: number;
  page: number;
  pageSize: number;
  totalPages: number;
}> {
  const filePath = resolveInputPath(rawPath);
  const ext = path.extname(filePath).toLowerCase();

  let workbook: XLSX.WorkBook;
  if (ext === '.csv' || ext === '.tsv') {
    let text = options.bufferOverride
      ? options.bufferOverride.toString('utf-8')
      : await fs.readFile(filePath, 'utf-8');
    if (text.charCodeAt(0) === 0xfeff) {
      text = text.slice(1);
    }
    workbook = XLSX.read(text, {
      type: 'string',
      FS: ext === '.tsv' ? '\t' : undefined,
      raw: false,
    });
  } else if (options.bufferOverride) {
    workbook = XLSX.read(options.bufferOverride, {
      type: 'buffer',
      codepage: 65001,
      cellDates: true,
    });
  } else {
    workbook = XLSX.readFile(filePath, {
      codepage: 65001,
      cellDates: true,
    });
  }

  const sheetNames = workbook.SheetNames;
  if (sheetNames.length === 0) {
    return {
      sheetNames: [],
      activeSheet: '',
      columns: [],
      rows: [],
      formulas: {},
      totalRows: 0,
      filteredRows: 0,
      totalCols: 0,
      page: 1,
      pageSize: options.pageSize ?? 150,
      totalPages: 1,
    };
  }

  const activeSheet =
    options.sheetName && sheetNames.includes(options.sheetName) ? options.sheetName : sheetNames[0];
  const sheet = workbook.Sheets[activeSheet];
  const rawMatrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: '',
    raw: false,
  });

  let maxCol = 0;
  for (const row of rawMatrix) {
    if (Array.isArray(row)) {
      let lastNonEmpty = row.length - 1;
      while (lastNonEmpty >= 0 && (row[lastNonEmpty] === '' || row[lastNonEmpty] === null || row[lastNonEmpty] === undefined)) {
        lastNonEmpty -= 1;
      }
      if (lastNonEmpty + 1 > maxCol) maxCol = lastNonEmpty + 1;
    }
  }
  const totalCols = Math.min(Math.max(maxCol, 1), 100);
  const columns = Array.from({ length: totalCols }, (_, i) => indexToColName(i));

  const searchLower = (options.search || '').trim().toLowerCase();
  const indexedRows: Array<{ rowNumber: number; cells: string[] }> = [];

  for (let r = 0; r < rawMatrix.length; r++) {
    const rawRow = Array.isArray(rawMatrix[r]) ? rawMatrix[r] : [];
    const cells = new Array<string>(totalCols);
    let matches = !searchLower;
    for (let c = 0; c < totalCols; c++) {
      const val = rawRow[c] === null || rawRow[c] === undefined ? '' : String(rawRow[c]);
      const clipped = val.length > 600 ? `${val.slice(0, 600)}…` : val;
      cells[c] = clipped;
      if (!matches && clipped.toLowerCase().includes(searchLower)) {
        matches = true;
      }
    }
    if (matches) {
      indexedRows.push({ rowNumber: r + 1, cells });
    }
  }

  const pageSize = Math.min(Math.max(options.pageSize ?? 150, 20), 500);
  const totalPages = Math.max(1, Math.ceil(indexedRows.length / pageSize));
  const page = Math.min(Math.max(options.page ?? 1, 1), totalPages);
  const start = (page - 1) * pageSize;
  const rows = indexedRows.slice(start, start + pageSize);

  const formulas: Record<string, string> = {};
  for (const row of rows) {
    for (let c = 0; c < totalCols; c++) {
      const addr = XLSX.utils.encode_cell({ r: row.rowNumber - 1, c });
      const cellObj = sheet[addr] as XLSX.CellObject | undefined;
      if (cellObj?.f) {
        formulas[addr] = `=${cellObj.f}`;
      }
    }
  }

  return {
    sheetNames,
    activeSheet,
    columns,
    rows,
    formulas,
    totalRows: rawMatrix.length,
    filteredRows: indexedRows.length,
    totalCols,
    page,
    pageSize,
    totalPages,
  };
}

async function inlineLocalStyles(html: string, baseDir: string, maxTotalCssBytes = 200 * 1024): Promise<string> {
  let remaining = maxTotalCssBytes;
  const linkRegex = /<link\b([^>]*?)>/gi;
  const matches = [...html.matchAll(linkRegex)];
  if (matches.length === 0) return html;

  let out = '';
  let lastIndex = 0;
  for (const m of matches) {
    const full = m[0];
    const attrs = m[1] || '';
    const start = m.index ?? 0;
    out += html.slice(lastIndex, start);
    lastIndex = start + full.length;

    const isStylesheet = /\brel\s*=\s*(?:"[^"]*\bstylesheet\b[^"]*"|'[^']*\bstylesheet\b[^']*')/i.test(attrs);
    const hrefMatch = /\bhref\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(attrs);
    const href = (hrefMatch?.[1] ?? hrefMatch?.[2] ?? '').trim();

    if (!isStylesheet || !href || isRemoteOrSpecialUri(href)) {
      out += full;
      continue;
    }

    const cleanRel = decodeURIComponent(href.split(/[?#]/)[0]);
    const resolved = path.resolve(baseDir, cleanRel);
    try {
      const stat = await fs.stat(resolved);
      if (stat.isFile() && stat.size <= remaining) {
        const css = await fs.readFile(resolved, 'utf-8');
        remaining -= Buffer.byteLength(css, 'utf-8');
        out += `<style>\n${css.replace(/<\/style>/gi, '<\\/style>')}\n</style>`;
        continue;
      }
    } catch {
      // Ignore missing stylesheet
    }
    out += '';
  }
  out += html.slice(lastIndex);
  return out;
}

export async function renderHtmlFile(rawPath: string): Promise<{
  previewHtml: string;
  sourceHtml: string;
  sourceTruncated: boolean;
  totalBytes: number;
}> {
  const filePath = resolveInputPath(rawPath);
  const stat = await fs.stat(filePath);
  const maxRead = 1024 * 1024;
  let raw: string;

  if (stat.size > maxRead) {
    const handle = await fs.open(filePath, 'r');
    try {
      const buf = Buffer.allocUnsafe(maxRead);
      const { bytesRead } = await handle.read(buf, 0, maxRead, 0);
      raw = buf.subarray(0, bytesRead).toString('utf-8');
    } finally {
      await handle.close();
    }
  } else {
    raw = await fs.readFile(filePath, 'utf-8');
  }

  const baseDir = path.dirname(filePath);
  const withCss = await inlineLocalStyles(raw, baseDir);
  const withAssets = await inlineLocalImages(withCss, baseDir, 150 * 1024, 300 * 1024);
  const cleaned = sanitizeHtml(withAssets);

  const cspMeta =
    '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src data: blob:; font-src data:;">';

  let previewHtml: string;
  if (/<head\b[^>]*>/i.test(cleaned)) {
    previewHtml = cleaned.replace(/<head\b[^>]*>/i, (tag) => `${tag}\n${cspMeta}`);
  } else if (/<html\b[^>]*>/i.test(cleaned)) {
    previewHtml = cleaned.replace(/<html\b[^>]*>/i, (tag) => `${tag}\n<head>${cspMeta}</head>`);
  } else {
    previewHtml = `<!DOCTYPE html><html><head><meta charset="utf-8">${cspMeta}</head><body>${cleaned}</body></html>`;
  }

  const sourceLimit = 120 * 1024;
  const sourceTruncated = raw.length > sourceLimit;
  const sourceSlice = sourceTruncated ? raw.slice(0, sourceLimit) : raw;
  const sourceHtml = hljs.highlight(sourceSlice, { language: 'xml', ignoreIllegals: true }).value;

  return { previewHtml, sourceHtml, sourceTruncated, totalBytes: stat.size };
}

export async function renderTextFile(
  rawPath: string,
  options: { startLine?: number; lineCount?: number } = {},
): Promise<{
  linesHtml: string[];
  rawText: string;
  language: string;
  totalLines: number;
  startLine: number;
  endLine: number;
  hasMore: boolean;
  truncatedFile: boolean;
}> {
  const filePath = resolveInputPath(rawPath);
  const stat = await fs.stat(filePath);
  const maxBytes = 4 * 1024 * 1024;
  let truncatedFile = false;
  let content: string;

  if (stat.size > maxBytes) {
    truncatedFile = true;
    const handle = await fs.open(filePath, 'r');
    try {
      const buf = Buffer.allocUnsafe(maxBytes);
      const { bytesRead } = await handle.read(buf, 0, maxBytes, 0);
      content = buf.subarray(0, bytesRead).toString('utf-8');
    } finally {
      await handle.close();
    }
  } else {
    content = await fs.readFile(filePath, 'utf-8');
  }

  const allLines = content.split(/\r?\n/);
  const totalLines = allLines.length;
  const startLine = Math.max(1, Math.min(options.startLine ?? 1, Math.max(1, totalLines)));
  const lineCount = Math.min(Math.max(options.lineCount ?? 400, 50), 1000);
  const slice = allLines.slice(startLine - 1, startLine - 1 + lineCount);
  const endLine = startLine + slice.length - 1;
  const hasMore = endLine < totalLines;

  const ext = path.extname(filePath).toLowerCase();
  const mappedLang = EXT_LANG_MAP[ext];
  const rawSlice = slice.join('\n');

  let language = mappedLang || 'plaintext';
  let highlightedText: string;
  if (mappedLang && hljs.getLanguage(mappedLang)) {
    highlightedText = hljs.highlight(rawSlice, { language: mappedLang, ignoreIllegals: true }).value;
  } else if (slice.length <= 400 && rawSlice.length < 100_000) {
    const auto = hljs.highlightAuto(rawSlice);
    language = auto.language || 'plaintext';
    highlightedText = auto.value;
  } else {
    highlightedText = escapeHtml(rawSlice);
  }

  return {
    linesHtml: highlightedText.split('\n'),
    rawText: rawSlice,
    language,
    totalLines,
    startLine,
    endLine,
    hasMore,
    truncatedFile,
  };
}
