import fs from 'node:fs/promises';
import { DOMParser } from '@xmldom/xmldom';
import * as CFB from 'cfb';
import JSZip from 'jszip';
import { resolveInputPath } from './files.js';
import type { PptxRenderResult, PptxShapeModel, PptxSlideModel } from './officeWorktree.js';
import { escapeHtml, extractTocAndAddIds, sanitizeHtml, type TocItem } from './renderers.js';

export interface LegacyPage {
  pageNumber: number;
  html: string;
}

export interface LegacyDocRenderResult {
  format: 'doc' | 'rtf' | 'odt' | 'html_doc' | 'docx_alias';
  delegateKind?: 'docx';
  pages: LegacyPage[];
  toc: TocItem[];
  plainText: string;
  paragraphCount: number;
  notices: string[];
}

export interface LegacyPptRenderResult extends PptxRenderResult {
  format: 'ppt' | 'pptx_alias';
  delegateKind?: 'pptx';
  notices: string[];
}

type DocBlock =
  | { kind: 'heading'; level: 1 | 2 | 3; text: string; center?: boolean }
  | { kind: 'paragraph'; text: string; center?: boolean }
  | { kind: 'table'; rows: string[][] };

function decodeGbkOrLatin1(bytes: Uint8Array): string {
  try {
    return new TextDecoder('gbk', { fatal: false }).decode(bytes);
  } catch {
    return Buffer.from(bytes).toString('latin1');
  }
}

function decodeUtf8OrGbk(buf: Buffer): string {
  const utf8 = buf.toString('utf-8');
  const bad = (utf8.match(/\uFFFD/g) || []).length;
  if (bad > 2 && bad * 40 > utf8.length) {
    return decodeGbkOrLatin1(buf);
  }
  return utf8;
}

function cleanWordBinaryFields(raw: string): string {
  // Handle nested or flat field codes: \x13 instr \x14 display \x15 -> keep display
  let s = raw;
  for (let pass = 0; pass < 4; pass++) {
    const next = s.replace(/\x13[^\x13\x15]*?\x14([^\x13\x15]*?)\x15/g, '$1');
    if (next === s) break;
    s = next;
  }
  s = s.replace(/\x13[^\x13\x15]*?\x15/g, '');
  // Strip inline object/footnote/annotation anchors
  return s.replace(/[\x00-\x06\x08\x0e-\x1f\uFFFC\uFFFF]/g, '');
}

function detectHeadingLevel(text: string, isFirstPage: boolean, blockIdx: number): 1 | 2 | 3 | null {
  const t = text.trim();
  if (!t || t.length > 52) return null;
  if (/^(?:实验|第)[一二三四五六七八九十百0-9]+\s*[章部分节讲]/.test(t)) return 1;
  if (/^[一二三四五六七八九十]+、\s*\S/.test(t)) return 1;
  if (/^（[一二三四五六七八九十]+）\s*\S/.test(t) || /^\([一二三四五六七八九十]+\)\s*\S/.test(t)) return 2;
  if (/^\d+\.\d+\.\d+\s+\S/.test(t)) return 3;
  if (/^\d+\.\d+\s+\S/.test(t)) return 2;
  if (/^\d+[.、]\s+\S/.test(t) && t.length <= 36) return 1;
  if (/^(?:目\s*录|摘\s*要|前\s*言|引\s*言|结\s*论|参\s*考\s*文\s*献|附\s*录)$/.test(t)) return 1;
  if (isFirstPage && blockIdx <= 2 && t.length <= 28 && !/[。；，,：:]$/.test(t)) return 1;
  return null;
}

function buildPagesFromPageBlocks(
  rawPages: DocBlock[][],
  notices: string[],
  format: LegacyDocRenderResult['format'],
): LegacyDocRenderResult {
  // Balance overly dense pages (~34 blocks per A4 sheet) when document lacks explicit page breaks
  const balancedPages: DocBlock[][] = [];
  const maxBlocksPerPage = 34;

  for (const pageBlocks of rawPages) {
    if (pageBlocks.length === 0) continue;
    if (pageBlocks.length <= maxBlocksPerPage) {
      balancedPages.push(pageBlocks);
      continue;
    }
    for (let i = 0; i < pageBlocks.length; i += maxBlocksPerPage) {
      balancedPages.push(pageBlocks.slice(i, i + maxBlocksPerPage));
    }
  }

  if (balancedPages.length === 0) {
    balancedPages.push([{ kind: 'paragraph', text: '（文档内容为空或仅包含图片/图形对象）' }]);
  }

  const toc: TocItem[] = [];
  const plainParts: string[] = [];
  let headingSeq = 0;
  let paragraphCount = 0;

  const pages: LegacyPage[] = balancedPages.map((blocks, idx) => {
    const pageNum = idx + 1;
    const htmlParts: string[] = [];
    for (const blk of blocks) {
      if (blk.kind === 'heading') {
        headingSeq += 1;
        paragraphCount += 1;
        const id = `legacy-h-${headingSeq}`;
        const tag = blk.level === 1 ? 'h2' : blk.level === 2 ? 'h3' : 'h4';
        const alignCls = blk.center ? ' style="text-align:center"' : '';
        htmlParts.push(`<${tag} id="${id}" class="legacy-h${blk.level}"${alignCls}>${escapeHtml(blk.text)}</${tag}>`);
        toc.push({ id, level: blk.level, text: blk.text, });
        plainParts.push(blk.text);
      } else if (blk.kind === 'paragraph') {
        paragraphCount += 1;
        const alignCls = blk.center ? ' style="text-align:center"' : '';
        const formatted = escapeHtml(blk.text).replace(/\n/g, '<br />');
        htmlParts.push(`<p class="legacy-p"${alignCls}>${formatted}</p>`);
        plainParts.push(blk.text);
      } else {
        const rowsHtml = blk.rows
          .map(
            (r) =>
              `<tr>${r
                .map((c) => `<td>${escapeHtml(c).replace(/\n/g, '<br />')}</td>`)
                .join('')}</tr>`,
          )
          .join('');
        htmlParts.push(`<table class="legacy-table"><tbody>${rowsHtml}</tbody></table>`);
        plainParts.push(blk.rows.map((r) => r.join('\t')).join('\n'));
        paragraphCount += blk.rows.length;
      }
    }
    return { pageNumber: pageNum, html: htmlParts.join('\n') };
  });

  return {
    format,
    pages,
    toc,
    plainText: plainParts.join('\n\n'),
    paragraphCount,
    notices,
  };
}

function splitWordTextIntoPages(rawText: string): DocBlock[][] {
  const cleaned = cleanWordBinaryFields(rawText);
  const pageChunks = cleaned.split('\x0c');
  const pages: DocBlock[][] = [];

  for (let pIdx = 0; pIdx < pageChunks.length; pIdx++) {
    const chunk = pageChunks[pIdx];
    const blocks: DocBlock[] = [];
    let pendingTableRows: string[][] = [];

    const flushTable = () => {
      if (pendingTableRows.length > 0) {
        blocks.push({ kind: 'table', rows: pendingTableRows });
        pendingTableRows = [];
      }
    };

    // Normalize \r\n -> \r, split paragraphs on \r
    const lines = chunk.replace(/\r\n/g, '\r').replace(/\n/g, '\r').split('\r');

    for (const rawLine of lines) {
      if (rawLine.includes('\x07')) {
        // Word table rows use \x07 as cell/row terminator; \x07\x07 ends a row
        const rowSegments = rawLine.split(/\x07\x07(?:\r|$)?/);
        for (const seg of rowSegments) {
          if (!seg.trim() && !seg.includes('\x07')) continue;
          const rawCells = seg.split('\x07').map((c) => c.replace(/\x0b/g, '\n').trim());
          // Trim leading/trailing empty cells produced by row-end \x07 markers
          while (rawCells.length > 1 && rawCells[rawCells.length - 1] === '') rawCells.pop();
          while (rawCells.length > 1 && rawCells[0] === '') rawCells.shift();
          if (rawCells.some((c) => c.length > 0)) {
            pendingTableRows.push(rawCells);
          }
        }
        continue;
      }

      flushTable();
      const lineText = rawLine.replace(/\x0b/g, '\n').replace(/\t+/g, '  ').trimEnd();
      const trimmed = lineText.trim();
      if (!trimmed) continue;

      const hLevel = detectHeadingLevel(trimmed, pIdx === 0, blocks.length);
      if (hLevel) {
        blocks.push({
          kind: 'heading',
          level: hLevel,
          text: trimmed,
          center: pIdx === 0 && blocks.length <= 1 && trimmed.length <= 26,
        });
      } else {
        blocks.push({ kind: 'paragraph', text: trimmed });
      }
    }

    flushTable();
    if (blocks.length > 0) {
      pages.push(blocks);
    }
  }

  return pages;
}

function parseWord97BinaryText(buf: Buffer): string {
  const cfb = CFB.read(buf, { type: 'buffer' });
  const wdEntry = CFB.find(cfb, 'WordDocument');
  if (!wdEntry || !wdEntry.content?.length) {
    throw new Error('缺少 WordDocument 数据流');
  }
  const wd = Buffer.from(wdEntry.content);
  if (wd.length < 64) {
    throw new Error('WordDocument 文件头长度不足');
  }

  const flags = wd.readUInt16LE(0x000a);
  const encrypted = (flags & 0x0100) !== 0;
  if (encrypted) {
    throw new Error('该 .doc 文档已加密，暂不支持免密解析');
  }

  const tblName = flags & 0x0200 ? '1Table' : '0Table';
  const tblEntry = CFB.find(cfb, tblName) || CFB.find(cfb, '1Table') || CFB.find(cfb, '0Table');
  if (!tblEntry || !tblEntry.content?.length) {
    throw new Error(`缺少 ${tblName} 表数据流`);
  }
  const tbl = Buffer.from(tblEntry.content);

  let off = 32;
  const csw = wd.readUInt16LE(off);
  off += 2 + csw * 2;
  const cslw = wd.readUInt16LE(off);
  const ccpText = cslw >= 4 ? wd.readInt32LE(off + 2 + 3 * 4) : 0;
  off += 2 + cslw * 4;
  const cbRgFcLcb = wd.readUInt16LE(off);
  off += 2;

  if (cbRgFcLcb < 34 || off + 34 * 8 > wd.length) {
    throw new Error('不支持的旧版 Word FIB 结构');
  }

  const fcClx = wd.readUInt32LE(off + 33 * 8);
  const lcbClx = wd.readUInt32LE(off + 33 * 8 + 4);
  if (lcbClx === 0 || fcClx + lcbClx > tbl.length) {
    throw new Error('Clx 文本索引越界');
  }

  let pos = fcClx;
  const end = fcClx + lcbClx;
  while (pos < end && tbl[pos] === 0x01) {
    if (pos + 3 > end) break;
    const cbGrpprl = tbl.readUInt16LE(pos + 1);
    pos += 3 + cbGrpprl;
  }
  if (pos + 5 > end || tbl[pos] !== 0x02) {
    throw new Error('未找到 Pcdt 片段表');
  }

  const lcb = tbl.readUInt32LE(pos + 1);
  const pcdStart = pos + 5;
  if (lcb < 4 || pcdStart + lcb > end) {
    throw new Error('Pcdt 片段表长度异常');
  }

  const nPieces = Math.floor((lcb - 4) / 12);
  let text = '';
  for (let i = 0; i < nPieces; i++) {
    const cpStart = tbl.readInt32LE(pcdStart + i * 4);
    const cpEnd = tbl.readInt32LE(pcdStart + (i + 1) * 4);
    const charCount = cpEnd - cpStart;
    if (charCount <= 0) continue;

    const pcdOff = pcdStart + (nPieces + 1) * 4 + i * 8;
    const fcRaw = tbl.readUInt32LE(pcdOff + 2);
    const isCompressed = (fcRaw & 0x40000000) !== 0;
    const fc = isCompressed ? (fcRaw & 0x3fffffff) >>> 1 : fcRaw & 0x3fffffff;

    if (isCompressed) {
      if (fc < wd.length) {
        const endByte = Math.min(wd.length, fc + charCount);
        text += wd.subarray(fc, endByte).toString('latin1');
      }
    } else if (fc < wd.length) {
      const endByte = Math.min(wd.length, fc + charCount * 2);
      text += wd.subarray(fc, endByte).toString('utf16le');
    }
  }

  if (ccpText > 0 && text.length > ccpText) {
    text = text.slice(0, ccpText);
  }
  return text;
}

function fallbackScanBinaryStrings(buf: Buffer): string {
  const runs: string[] = [];
  let current = '';
  for (let i = 0; i + 1 < buf.length; i += 2) {
    const code = buf.readUInt16LE(i);
    const isPrintable =
      (code >= 0x4e00 && code <= 0x9fff) ||
      (code >= 0x3000 && code <= 0x303f) ||
      (code >= 0xff00 && code <= 0xffef) ||
      (code >= 0x20 && code <= 0x7e) ||
      code === 0x0d ||
      code === 0x0a;
    if (isPrintable) {
      current += String.fromCharCode(code);
    } else {
      if (current.trim().length >= 4) {
        runs.push(current.trim());
      }
      current = '';
    }
  }
  if (current.trim().length >= 4) runs.push(current.trim());
  return runs.join('\r');
}

function parseRtfText(rawRtf: string): DocBlock[][] {
  const skipDestinations = new Set([
    'fonttbl',
    'colortbl',
    'stylesheet',
    'info',
    'pict',
    'object',
    'header',
    'footer',
    'headerl',
    'headerr',
    'headerf',
    'footerl',
    'footerr',
    'footerf',
    'fldinst',
    'themedata',
    'colorschememapping',
    'latentstyles',
    'datastore',
    'xmlnstbl',
    'listtable',
    'listoverridetable',
    'revtbl',
    'rsidtbl',
  ]);

  interface StackState {
    uc: number;
    skip: boolean;
  }

  const stack: StackState[] = [{ uc: 1, skip: false }];
  let skipNextChars = 0;
  let hexBytes: number[] = [];

  const pages: DocBlock[][] = [[]];
  let curBlocks = pages[0];
  let curParagraph = '';
  let curRowCells: string[] = [];
  let inTableRow = false;
  let pendingTableRows: string[][] = [];

  const flushHex = () => {
    if (hexBytes.length > 0) {
      curParagraph += decodeGbkOrLatin1(new Uint8Array(hexBytes));
      hexBytes = [];
    }
  };

  const flushPendingTable = () => {
    if (pendingTableRows.length > 0) {
      curBlocks.push({ kind: 'table', rows: pendingTableRows });
      pendingTableRows = [];
    }
  };

  const flushParagraph = () => {
    flushHex();
    const trimmed = curParagraph.trim();
    curParagraph = '';
    if (!trimmed) return;
    flushPendingTable();
    const hLevel = detectHeadingLevel(trimmed, pages.length === 1, curBlocks.length);
    if (hLevel) {
      curBlocks.push({ kind: 'heading', level: hLevel, text: trimmed });
    } else {
      curBlocks.push({ kind: 'paragraph', text: trimmed });
    }
  };

  let i = 0;
  const len = rawRtf.length;
  while (i < len) {
    const ch = rawRtf[i];
    const top = stack[stack.length - 1];

    if (ch === '{') {
      flushHex();
      stack.push({ uc: top.uc, skip: top.skip });
      i += 1;
      continue;
    }

    if (ch === '}') {
      flushHex();
      if (stack.length > 1) stack.pop();
      skipNextChars = 0;
      i += 1;
      continue;
    }

    if (ch === '\r' || ch === '\n') {
      i += 1;
      continue;
    }

    if (ch === '\\') {
      i += 1;
      if (i >= len) break;
      const esc = rawRtf[i];

      if (esc === '*') {
        flushHex();
        top.skip = true;
        i += 1;
        continue;
      }

      if (esc === '\'' && i + 2 < len) {
        const hex = rawRtf.slice(i + 1, i + 3);
        const byteVal = Number.parseInt(hex, 16);
        i += 3;
        if (!top.skip && !Number.isNaN(byteVal)) {
          if (skipNextChars > 0) {
            skipNextChars -= 1;
          } else {
            hexBytes.push(byteVal);
          }
        }
        continue;
      }

      flushHex();

      if (esc === '\\' || esc === '{' || esc === '}') {
        i += 1;
        if (!top.skip) {
          if (skipNextChars > 0) skipNextChars -= 1;
          else curParagraph += esc;
        }
        continue;
      }

      if (esc === '~') {
        i += 1;
        if (!top.skip) curParagraph += ' ';
        continue;
      }
      if (esc === '_' || esc === '-') {
        i += 1;
        if (!top.skip && esc === '_') curParagraph += '-';
        continue;
      }
      if (esc === '\r' || esc === '\n') {
        i += 1;
        if (!top.skip) flushParagraph();
        continue;
      }

      // Control word
      const wordMatch = /^[a-zA-Z]+/.exec(rawRtf.slice(i));
      if (!wordMatch) {
        i += 1;
        continue;
      }
      const word = wordMatch[0];
      i += word.length;
      const numMatch = /^-?\d+/.exec(rawRtf.slice(i));
      const param = numMatch ? Number.parseInt(numMatch[0], 10) : null;
      if (numMatch) i += numMatch[0].length;
      if (rawRtf[i] === ' ') i += 1;

      if (skipDestinations.has(word)) {
        top.skip = true;
        continue;
      }
      if (top.skip) continue;

      if (word === 'uc' && param !== null) {
        top.uc = Math.max(0, param);
      } else if (word === 'u' && param !== null) {
        const code = param < 0 ? 65536 + param : param;
        curParagraph += String.fromCharCode(code);
        skipNextChars = top.uc;
      } else if (word === 'par' || word === 'line') {
        if (!inTableRow) flushParagraph();
        else curParagraph += '\n';
      } else if (word === 'tab') {
        curParagraph += '  ';
      } else if (word === 'emdash') {
        curParagraph += '—';
      } else if (word === 'endash') {
        curParagraph += '–';
      } else if (word === 'lquote' || word === 'rquote') {
        curParagraph += "'";
      } else if (word === 'ldblquote' || word === 'rdblquote') {
        curParagraph += '"';
      } else if (word === 'trowd') {
        flushParagraph();
        inTableRow = true;
        curRowCells = [];
      } else if (word === 'cell') {
        curRowCells.push(curParagraph.trim());
        curParagraph = '';
      } else if (word === 'row') {
        if (curParagraph.trim()) curRowCells.push(curParagraph.trim());
        curParagraph = '';
        if (curRowCells.length > 0) pendingTableRows.push(curRowCells);
        curRowCells = [];
        inTableRow = false;
      } else if (word === 'page') {
        flushParagraph();
        flushPendingTable();
        if (curBlocks.length > 0) {
          curBlocks = [];
          pages.push(curBlocks);
        }
      }
      continue;
    }

    if (!top.skip) {
      if (skipNextChars > 0) {
        skipNextChars -= 1;
      } else {
        flushHex();
        curParagraph += ch;
      }
    }
    i += 1;
  }

  flushParagraph();
  flushPendingTable();
  return pages.filter((p) => p.length > 0);
}

async function parseOdtContent(buf: Buffer): Promise<DocBlock[][]> {
  const zip = await JSZip.loadAsync(buf);
  const contentFile = zip.file('content.xml');
  if (!contentFile) throw new Error('ODT 文件缺少 content.xml');
  const xml = await contentFile.async('string');
  const doc = new DOMParser().parseFromString(xml, 'text/xml');

  const body = doc.getElementsByTagName('office:text')[0] || doc.documentElement;
  const blocks: DocBlock[] = [];

  const collectText = (node: Element): string => {
    let out = '';
    for (let i = 0; i < node.childNodes.length; i++) {
      const ch = node.childNodes[i];
      if (ch.nodeType === 3) {
        out += ch.nodeValue || '';
      } else if (ch.nodeType === 1) {
        const el = ch as Element;
        const ln = (el.localName || el.tagName || '').toLowerCase();
        if (ln === 'line-break') out += '\n';
        else if (ln === 'tab') out += '  ';
        else if (ln === 's') out += ' '.repeat(Math.max(1, Number(el.getAttribute('text:c') || 1)));
        else out += collectText(el);
      }
    }
    return out;
  };

  for (let i = 0; i < body.childNodes.length; i++) {
    const ch = body.childNodes[i];
    if (ch.nodeType !== 1) continue;
    const el = ch as Element;
    const ln = (el.localName || el.tagName || '').toLowerCase();
    if (ln === 'h') {
      const text = collectText(el).trim();
      if (!text) continue;
      const lvlRaw = Number(el.getAttribute('text:outline-level') || 1);
      const level: 1 | 2 | 3 = lvlRaw <= 1 ? 1 : lvlRaw === 2 ? 2 : 3;
      blocks.push({ kind: 'heading', level, text });
    } else if (ln === 'p') {
      const text = collectText(el).trim();
      if (text) blocks.push({ kind: 'paragraph', text });
    } else if (ln === 'table') {
      const rows: string[][] = [];
      const rowEls = el.getElementsByTagName('table:table-row');
      for (let r = 0; r < rowEls.length; r++) {
        const rowEl = rowEls[r];
        const cellEls = rowEl.getElementsByTagName('table:table-cell');
        const rowCells: string[] = [];
        for (let c = 0; c < cellEls.length; c++) {
          rowCells.push(collectText(cellEls[c]).trim());
        }
        if (rowCells.some((x) => x.length > 0)) rows.push(rowCells);
      }
      if (rows.length > 0) blocks.push({ kind: 'table', rows });
    }
  }

  return [blocks];
}

function decodeQuotedPrintable(input: string): string {
  const joined = input.replace(/=\r?\n/g, '');
  const bytes: number[] = [];
  for (let i = 0; i < joined.length; i++) {
    if (joined[i] === '=' && /^[0-9A-Fa-f]{2}$/.test(joined.slice(i + 1, i + 3))) {
      bytes.push(Number.parseInt(joined.slice(i + 1, i + 3), 16));
      i += 2;
    } else {
      bytes.push(joined.charCodeAt(i) & 0xff);
    }
  }
  return decodeUtf8OrGbk(Buffer.from(bytes));
}

function parseHtmlOrMhtmlDoc(raw: string): LegacyDocRenderResult {
  let htmlSource = raw;
  if (/MIME-Version:/i.test(raw) && /Content-Type:\s*text\/html/i.test(raw)) {
    const htmlPartMatch = /Content-Type:\s*text\/html[\s\S]*?\r?\n\r?\n([\s\S]*?)(?:\r?\n--[A-Za-z0-9_=-]+|$)/i.exec(
      raw,
    );
    if (htmlPartMatch) {
      htmlSource = /Content-Transfer-Encoding:\s*quoted-printable/i.test(raw)
        ? decodeQuotedPrintable(htmlPartMatch[1])
        : htmlPartMatch[1];
    }
  }

  const bodyMatch = /<body\b[^>]*>([\s\S]*?)<\/body>/i.exec(htmlSource);
  const inner = bodyMatch ? bodyMatch[1] : htmlSource;
  const cleaned = sanitizeHtml(inner);
  const { html, toc } = extractTocAndAddIds(cleaned);
  const plainText = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  return {
    format: 'html_doc',
    pages: [{ pageNumber: 1, html }],
    toc,
    plainText,
    paragraphCount: Math.max(1, (html.match(/<(?:p|h[1-6]|tr)\b/gi) || []).length),
    notices: ['该 .doc 文件实为 HTML/MHTML 网页文档格式，已以安全只读 A4 视图渲染。'],
  };
}

export async function renderLegacyDocument(rawPath: string): Promise<LegacyDocRenderResult> {
  const filePath = resolveInputPath(rawPath);
  const buf = await fs.readFile(filePath);
  if (buf.length < 4) {
    throw new Error('文档内容过短，无法解析');
  }

  // 1. ZIP signature (could be misnamed .docx or .odt)
  if (buf[0] === 0x50 && buf[1] === 0x4b && buf[2] === 0x03 && buf[3] === 0x04) {
    const zip = await JSZip.loadAsync(buf);
    if (zip.file('word/document.xml')) {
      return {
        format: 'docx_alias',
        delegateKind: 'docx',
        pages: [],
        toc: [],
        plainText: '',
        paragraphCount: 0,
        notices: ['检测到 OOXML (.docx) 格式内容，已切换原生 A4 排版渲染。'],
      };
    }
    if (zip.file('content.xml')) {
      const blocks = await parseOdtContent(buf);
      return buildPagesFromPageBlocks(
        blocks,
        ['OpenDocument (.odt) 只读排版视图 · 支持右键选择外部 Office 应用编辑。'],
        'odt',
      );
    }
  }

  const headText = buf.subarray(0, Math.min(buf.length, 512)).toString('latin1');

  // 2. RTF signature
  if (/^\s*\{\\rtf/i.test(headText)) {
    const rtfText = buf.toString('latin1');
    const pages = parseRtfText(rtfText);
    return buildPagesFromPageBlocks(
      pages,
      ['RTF 文档只读排版视图 · 完整样式编辑可通过右上角外部应用打开。'],
      'rtf',
    );
  }

  // 3. HTML / MHTML / Word XML exported as .doc
  if (/^\s*(?:<!doctype\s+html|<html|<head|mime-version:|\s*<\?xml[^>]*>[\s\S]*<w:wordDocument)/i.test(headText)) {
    return parseHtmlOrMhtmlDoc(decodeUtf8OrGbk(buf));
  }

  // 4. OLE2 Compound File (.doc / .wps / .dot)
  const isOle2 =
    buf.length >= 8 &&
    buf[0] === 0xd0 &&
    buf[1] === 0xcf &&
    buf[2] === 0x11 &&
    buf[3] === 0xe0 &&
    buf[4] === 0xa1 &&
    buf[5] === 0xb1 &&
    buf[6] === 0x1a &&
    buf[7] === 0xe1;

  if (!isOle2) {
    // Plain-text misnamed as .doc fallback
    const txt = decodeUtf8OrGbk(buf);
    const pages = splitWordTextIntoPages(txt.replace(/\r?\n/g, '\r'));
    return buildPagesFromPageBlocks(pages, ['按纯文本兼容模式解析。'], 'doc');
  }

  const notices: string[] = [
    '旧版二进制 .doc/.wps 文档只读 A4 视图（已离线解析正文、标题与表格；需编辑时可通过右上角「打开方式」调用 Word / WPS）。',
  ];

  let rawText = '';
  try {
    rawText = parseWord97BinaryText(buf);
  } catch (err) {
    notices.push(`结构化流回退扫描（${err instanceof Error ? err.message : String(err)}）`);
    try {
      const cfb = CFB.read(buf, { type: 'buffer' });
      const wd = CFB.find(cfb, 'WordDocument');
      rawText = fallbackScanBinaryStrings(wd?.content ? Buffer.from(wd.content) : buf);
    } catch {
      rawText = fallbackScanBinaryStrings(buf);
    }
  }

  const pages = splitWordTextIntoPages(rawText);
  return buildPagesFromPageBlocks(pages, notices, 'doc');
}

function buildSlideFromLines(slideNumber: number, lines: string[]): PptxSlideModel {
  const cleaned = lines.map((l) => l.trim()).filter(Boolean);
  const title = cleaned[0] || `第 ${slideNumber} 页`;
  const bodyLines = cleaned.slice(1);
  const widthPt = 960;
  const heightPt = 540;
  const shapes: PptxShapeModel[] = [
    {
      id: `ppt-title-${slideNumber}`,
      kind: 'text',
      leftPct: 7,
      topPct: 9,
      widthPct: 86,
      heightPct: 16,
      paragraphs: [
        {
          text: title,
          align: 'left',
          fontSizePt: 26,
          bold: true,
          colorHex: '#111827',
        },
      ],
    },
  ];

  if (bodyLines.length > 0) {
    shapes.push({
      id: `ppt-body-${slideNumber}`,
      kind: 'text',
      leftPct: 7,
      topPct: 28,
      widthPct: 86,
      heightPct: 62,
      paragraphs: bodyLines.slice(0, 18).map((text) => ({
        text,
        align: 'left',
        fontSizePt: 16,
        bold: false,
        colorHex: '#374151',
      })),
    });
  }

  return {
    slideNumber,
    title,
    widthPt,
    heightPt,
    shapes,
    plainText: cleaned.join('\n'),
  };
}

export async function renderLegacyPpt(rawPath: string): Promise<LegacyPptRenderResult> {
  const filePath = resolveInputPath(rawPath);
  const buf = await fs.readFile(filePath);

  if (buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b && buf[2] === 0x03 && buf[3] === 0x04) {
    const zip = await JSZip.loadAsync(buf);
    if (zip.file('ppt/presentation.xml')) {
      return {
        format: 'pptx_alias',
        delegateKind: 'pptx',
        slideCount: 0,
        widthPt: 960,
        heightPt: 540,
        slides: [],
        notices: ['检测到 OOXML (.pptx) 格式内容，已自动切换到完整幻灯片渲染。'],
      };
    }
  }

  const cfb = CFB.read(buf, { type: 'buffer' });
  const pptEntry = CFB.find(cfb, 'PowerPoint Document');
  if (!pptEntry || !pptEntry.content?.length) {
    throw new Error('缺少 PowerPoint Document 数据流');
  }
  const stream = Buffer.from(pptEntry.content);

  const SLIDE_LIST_WITH_TEXT = 0x0ff0;
  const SLIDE_PERSIST_ATOM = 0x03f3;
  const TEXT_CHARS_ATOM = 0x0fa0;
  const TEXT_BYTES_ATOM = 0x0fa8;
  const SLIDE_CONTAINER = 0x03ee;

  const listGroups: string[][] = [];
  const slideContainerGroups: string[][] = [];

  const readAtomText = (recType: number, start: number, len: number): string => {
    const end = Math.min(stream.length, start + len);
    if (end <= start) return '';
    if (recType === TEXT_CHARS_ATOM) {
      return stream.subarray(start, end).toString('utf16le').replace(/\r/g, '\n');
    }
    if (recType === TEXT_BYTES_ATOM) {
      return decodeGbkOrLatin1(stream.subarray(start, end)).replace(/\r/g, '\n');
    }
    return '';
  };

  const walk = (start: number, end: number, activeGroup: string[] | null, mode: 'root' | 'slwt' | 'slide') => {
    let pos = start;
    let curGroup = activeGroup;
    while (pos + 8 <= end) {
      const verInst = stream.readUInt16LE(pos);
      const recType = stream.readUInt16LE(pos + 2);
      const recLen = stream.readUInt32LE(pos + 4);
      const bodyStart = pos + 8;
      const bodyEnd = Math.min(end, bodyStart + recLen);
      const isContainer = (verInst & 0x0f) === 0x0f;

      if (isContainer) {
        if (recType === SLIDE_LIST_WITH_TEXT) {
          walk(bodyStart, bodyEnd, null, 'slwt');
        } else if (recType === SLIDE_CONTAINER) {
          const grp: string[] = [];
          slideContainerGroups.push(grp);
          walk(bodyStart, bodyEnd, grp, 'slide');
        } else {
          walk(bodyStart, bodyEnd, curGroup, mode);
        }
      } else {
        if (mode === 'slwt' && recType === SLIDE_PERSIST_ATOM) {
          curGroup = [];
          listGroups.push(curGroup);
        } else if (curGroup && (recType === TEXT_CHARS_ATOM || recType === TEXT_BYTES_ATOM)) {
          const txt = readAtomText(recType, bodyStart, recLen);
          for (const line of txt.split(/\r?\n/)) {
            const t = line.trim();
            if (t) curGroup.push(t);
          }
        }
      }

      if (bodyEnd <= pos) break;
      pos = bodyEnd;
    }
  };

  walk(0, stream.length, null, 'root');

  const chosen = listGroups.some((g) => g.length > 0)
    ? listGroups.filter((g) => g.length > 0)
    : slideContainerGroups.filter((g) => g.length > 0);

  const slides: PptxSlideModel[] = chosen.length
    ? chosen.map((lines, idx) => buildSlideFromLines(idx + 1, lines))
    : [buildSlideFromLines(1, ['旧版演示文稿（未检测到纯文本幻灯片内容）'])];

  return {
    format: 'ppt',
    slideCount: slides.length,
    widthPt: 960,
    heightPt: 540,
    slides,
    notices: ['旧版二进制 .ppt/.dps 只读文本视图（需查看复杂动画或编辑可通过右上角「打开方式」调用 PowerPoint / WPS）。'],
  };
}
