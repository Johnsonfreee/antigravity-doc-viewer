import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { MAX_CHUNK_BYTES, resolveInputPath } from './files.js';

type XmlDoc = ReturnType<DOMParser['parseFromString']>;
type XmlEl = ReturnType<XmlDoc['createElementNS']>;

export interface WorktreeVersionEntry {
  versionId: string;
  fileName: string;
  savedAt: number;
  size: number;
  summary: string;
}

export interface WorktreeMeta {
  originalPath: string;
  fileName: string;
  kind: 'docx' | 'spreadsheet' | 'slides';
  ext: string;
  status: 'draft' | 'clean';
  baseSha256: string;
  draftSize: number;
  summary: string;
  createdAt: number;
  updatedAt: number;
  versions: WorktreeVersionEntry[];
}

export interface DocxParagraphInfo {
  index: number;
  text: string;
  styleId?: string;
  inTable: boolean;
  runsCount: number;
}

export interface DocxTableInfo {
  tableIndex: number;
  rowCount: number;
  colCount: number;
  rows: string[][];
}

export interface DocxInspectResult {
  paragraphCount: number;
  tableCount: number;
  paragraphs: DocxParagraphInfo[];
  tables: DocxTableInfo[];
}

export interface DocxEditInput {
  replacements?: Array<{ find: string; replace: string }>;
  updateParagraphs?: Array<{ index: number; text: string }>;
  updateTableCells?: Array<{ tableIndex: number; row: number; col: number; text: string }>;
  appendBlocks?: Array<{
    type: 'heading1' | 'heading2' | 'paragraph' | 'table';
    text?: string;
    rows?: string[][];
  }>;
}

export interface XlsxEditInput {
  addSheets?: string[];
  renameSheets?: Array<{ from: string; to: string }>;
  setCells?: Array<{
    sheet?: string;
    cell?: string;
    row?: number;
    col?: number;
    value?: string | number | boolean;
    formula?: string;
  }>;
  appendRows?: Array<{
    sheet?: string;
    rows: Array<Array<string | number | boolean>>;
  }>;
}

export interface PptxParagraphModel {
  text: string;
  align?: 'left' | 'center' | 'right';
  fontSizePt?: number;
  bold?: boolean;
  colorHex?: string;
}

export interface PptxShapeModel {
  id: string;
  kind: 'text' | 'image' | 'table';
  leftPct: number;
  topPct: number;
  widthPct: number;
  heightPct: number;
  fillHex?: string;
  paragraphs?: PptxParagraphModel[];
  imageDataUri?: string;
  tableRows?: string[][];
}

export interface PptxSlideModel {
  slideNumber: number;
  title: string;
  widthPt: number;
  heightPt: number;
  shapes: PptxShapeModel[];
  plainText: string;
}

export interface PptxRenderResult {
  slideCount: number;
  widthPt: number;
  heightPt: number;
  slides: PptxSlideModel[];
}

export interface PptxEditInput {
  updateSlides?: Array<{
    slideNumber: number;
    title?: string;
    bodyLines?: string[];
    replacements?: Array<{ find: string; replace: string }>;
  }>;
  appendSlides?: Array<{
    title: string;
    bullets?: string[];
  }>;
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

function sha256(buf: Buffer): string {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

function normalizeKeyPath(filePath: string): string {
  const resolved = path.resolve(filePath);
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved;
}

function hashPath(filePath: string): string {
  return crypto.createHash('sha256').update(normalizeKeyPath(filePath)).digest('hex').slice(0, 16);
}

async function writeAtomicBuffer(targetPath: string, buf: Buffer): Promise<void> {
  await fs.mkdir(path.dirname(targetPath), { recursive: true });
  const tmp = `${targetPath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmp, buf);
  await fs.rename(tmp, targetPath);
}

function childElements(node: { childNodes: ArrayLike<unknown> } | null | undefined, localName?: string): XmlEl[] {
  if (!node?.childNodes) return [];
  const out: XmlEl[] = [];
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i] as { nodeType?: number; localName?: string; nodeName?: string };
    if (child && child.nodeType === 1) {
      const name = child.localName || String(child.nodeName || '').split(':').pop();
      if (!localName || name === localName) {
        out.push(child as XmlEl);
      }
    }
  }
  return out;
}

function descendantsByLocalName(root: XmlEl | XmlDoc, localName: string): XmlEl[] {
  const out: XmlEl[] = [];
  const walk = (node: { childNodes?: ArrayLike<unknown> }) => {
    if (!node?.childNodes) return;
    for (let i = 0; i < node.childNodes.length; i++) {
      const ch = node.childNodes[i] as {
        nodeType?: number;
        localName?: string;
        nodeName?: string;
        childNodes?: ArrayLike<unknown>;
      };
      if (ch && ch.nodeType === 1) {
        const name = ch.localName || String(ch.nodeName || '').split(':').pop();
        if (name === localName) out.push(ch as XmlEl);
        walk(ch);
      }
    }
  };
  walk(root);
  return out;
}

function getParagraphTextNodes(pEl: XmlEl): XmlEl[] {
  return descendantsByLocalName(pEl, 't');
}

function getParagraphText(pEl: XmlEl): string {
  return getParagraphTextNodes(pEl)
    .map((t) => t.textContent || '')
    .join('');
}

function setTextPreservingSpace(tEl: XmlEl, text: string): void {
  tEl.textContent = text;
  if (/^\s|\s$/.test(text)) {
    tEl.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  } else {
    tEl.removeAttribute('xml:space');
  }
}

function replaceTextInParagraph(pEl: XmlEl, find: string, replace: string): number {
  if (!find) return 0;
  let count = 0;
  while (true) {
    const tNodes = getParagraphTextNodes(pEl);
    if (tNodes.length === 0) break;
    const segments: Array<{ node: XmlEl; start: number; end: number; text: string }> = [];
    let cursor = 0;
    for (const node of tNodes) {
      const text = node.textContent || '';
      segments.push({ node, start: cursor, end: cursor + text.length, text });
      cursor += text.length;
    }
    const full = segments.map((s) => s.text).join('');
    const idx = full.indexOf(find);
    if (idx === -1) break;

    const matchStart = idx;
    const matchEnd = idx + find.length;
    let firstTouched = true;
    for (const seg of segments) {
      if (seg.end <= matchStart || seg.start >= matchEnd) continue;
      const localStart = Math.max(0, matchStart - seg.start);
      const localEnd = Math.min(seg.text.length, matchEnd - seg.start);
      const prefix = seg.text.slice(0, localStart);
      const suffix = seg.text.slice(localEnd);
      if (firstTouched) {
        setTextPreservingSpace(seg.node, `${prefix}${replace}${suffix}`);
        firstTouched = false;
      } else {
        setTextPreservingSpace(seg.node, `${prefix}${suffix}`);
      }
    }
    count += 1;
  }
  return count;
}

function setParagraphText(doc: XmlDoc, pEl: XmlEl, text: string): void {
  const tNodes = getParagraphTextNodes(pEl);
  if (tNodes.length > 0) {
    setTextPreservingSpace(tNodes[0], text);
    for (let i = 1; i < tNodes.length; i++) {
      setTextPreservingSpace(tNodes[i], '');
    }
    return;
  }
  const rEl = doc.createElementNS(W_NS, 'w:r');
  const tEl = doc.createElementNS(W_NS, 'w:t');
  setTextPreservingSpace(tEl, text);
  rEl.appendChild(tEl);
  pEl.appendChild(rEl);
}

function buildMinimalDocxBuffer(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`,
  );
  zip.folder('_rels')!.file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.folder('word')!.folder('_rels')!.file(
    'document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
  );
  zip.folder('word')!.file(
    'styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="180" w:after="80"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
</w:styles>`,
  );
  zip.folder('word')!.file(
    'document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}">
  <w:body>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export async function inspectDocxBuffer(buf: Buffer): Promise<DocxInspectResult> {
  const zip = await JSZip.loadAsync(buf);
  const xmlFile = zip.file('word/document.xml');
  if (!xmlFile) throw new Error('Invalid .docx: missing word/document.xml');
  const xml = await xmlFile.async('string');
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const body = descendantsByLocalName(doc, 'body')[0];
  if (!body) throw new Error('Invalid .docx: missing w:body');

  const allP = descendantsByLocalName(body, 'p');
  const topTables = childElements(body, 'tbl');
  const topTableParaSet = new Set<XmlEl>();
  for (const tbl of topTables) {
    for (const p of descendantsByLocalName(tbl, 'p')) {
      topTableParaSet.add(p);
    }
  }

  const paragraphs: DocxParagraphInfo[] = allP.slice(0, 300).map((pEl, index) => {
    const pPr = childElements(pEl, 'pPr')[0];
    const pStyle = pPr ? childElements(pPr, 'pStyle')[0] : undefined;
    const styleId = pStyle?.getAttribute('w:val') || pStyle?.getAttribute('val') || undefined;
    const runsCount = childElements(pEl, 'r').length;
    return {
      index,
      text: getParagraphText(pEl),
      styleId,
      inTable: topTableParaSet.has(pEl),
      runsCount,
    };
  });

  const tables: DocxTableInfo[] = topTables.slice(0, 40).map((tblEl, tableIndex) => {
    const trs = childElements(tblEl, 'tr');
    const rows = trs.slice(0, 80).map((tr) => {
      const tcs = childElements(tr, 'tc');
      return tcs.map((tc) =>
        descendantsByLocalName(tc, 'p')
          .map((p) => getParagraphText(p))
          .join('\n')
          .trim(),
      );
    });
    const colCount = rows.reduce((max, r) => Math.max(max, r.length), 0);
    return { tableIndex, rowCount: trs.length, colCount, rows };
  });

  return {
    paragraphCount: allP.length,
    tableCount: topTables.length,
    paragraphs,
    tables,
  };
}

export async function editDocxBuffer(baseBuf: Buffer | null, edits: DocxEditInput): Promise<Buffer> {
  const sourceBuf = baseBuf ?? (await buildMinimalDocxBuffer());
  const zip = await JSZip.loadAsync(sourceBuf);
  const xmlFile = zip.file('word/document.xml');
  if (!xmlFile) throw new Error('Invalid .docx: missing word/document.xml');
  const xml = await xmlFile.async('string');
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const body = descendantsByLocalName(doc, 'body')[0];
  if (!body) throw new Error('Invalid .docx: missing w:body');

  if (edits.replacements?.length) {
    const allP = descendantsByLocalName(body, 'p');
    for (const rep of edits.replacements) {
      for (const pEl of allP) {
        replaceTextInParagraph(pEl, rep.find, rep.replace);
      }
    }
  }

  if (edits.updateParagraphs?.length) {
    const allP = descendantsByLocalName(body, 'p');
    for (const item of edits.updateParagraphs) {
      const pEl = allP[item.index];
      if (!pEl) throw new Error(`Paragraph index ${item.index} out of range (0..${allP.length - 1})`);
      setParagraphText(doc, pEl, item.text);
    }
  }

  if (edits.updateTableCells?.length) {
    const topTables = childElements(body, 'tbl');
    for (const item of edits.updateTableCells) {
      const tbl = topTables[item.tableIndex];
      if (!tbl) throw new Error(`Table index ${item.tableIndex} out of range`);
      const tr = childElements(tbl, 'tr')[item.row];
      if (!tr) throw new Error(`Table row ${item.row} out of range`);
      const tc = childElements(tr, 'tc')[item.col];
      if (!tc) throw new Error(`Table col ${item.col} out of range`);
      let pEl = descendantsByLocalName(tc, 'p')[0];
      if (!pEl) {
        pEl = doc.createElementNS(W_NS, 'w:p');
        tc.appendChild(pEl);
      }
      setParagraphText(doc, pEl, item.text);
    }
  }

  if (edits.appendBlocks?.length) {
    const sectPr = childElements(body, 'sectPr')[0] || null;
    for (const block of edits.appendBlocks) {
      if (block.type === 'table') {
        const rows = block.rows || [['']];
        const tbl = doc.createElementNS(W_NS, 'w:tbl');
        const tblPr = doc.createElementNS(W_NS, 'w:tblPr');
        const borders = doc.createElementNS(W_NS, 'w:tblBorders');
        for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
          const b = doc.createElementNS(W_NS, `w:${side}`);
          b.setAttribute('w:val', 'single');
          b.setAttribute('w:sz', '4');
          b.setAttribute('w:space', '0');
          b.setAttribute('w:color', 'auto');
          borders.appendChild(b);
        }
        tblPr.appendChild(borders);
        tbl.appendChild(tblPr);

        for (const row of rows) {
          const tr = doc.createElementNS(W_NS, 'w:tr');
          for (const cellText of row) {
            const tc = doc.createElementNS(W_NS, 'w:tc');
            const p = doc.createElementNS(W_NS, 'w:p');
            setParagraphText(doc, p, String(cellText ?? ''));
            tc.appendChild(p);
            tr.appendChild(tc);
          }
          tbl.appendChild(tr);
        }
        body.insertBefore(tbl, sectPr);
      } else {
        const p = doc.createElementNS(W_NS, 'w:p');
        if (block.type === 'heading1' || block.type === 'heading2') {
          const pPr = doc.createElementNS(W_NS, 'w:pPr');
          const pStyle = doc.createElementNS(W_NS, 'w:pStyle');
          pStyle.setAttribute('w:val', block.type === 'heading1' ? 'Heading1' : 'Heading2');
          pPr.appendChild(pStyle);
          p.appendChild(pPr);
        }
        setParagraphText(doc, p, block.text || '');
        body.insertBefore(p, sectPr);
      }
    }
  }

  const nextXml = new XMLSerializer().serializeToString(doc);
  zip.file('word/document.xml', nextXml);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export function editSpreadsheetBuffer(
  baseBuf: Buffer | null,
  filePath: string,
  edits: XlsxEditInput,
): Buffer {
  const ext = path.extname(filePath).toLowerCase();
  let wb: XLSX.WorkBook;

  if (!baseBuf) {
    wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['']]);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  } else if (ext === '.csv' || ext === '.tsv') {
    let txt = baseBuf.toString('utf-8');
    if (txt.charCodeAt(0) === 0xfeff) txt = txt.slice(1);
    wb = XLSX.read(txt, { type: 'string', FS: ext === '.tsv' ? '\t' : undefined });
  } else {
    wb = XLSX.read(baseBuf, { type: 'buffer', codepage: 65001, cellDates: true });
  }

  if (edits.addSheets?.length) {
    for (const name of edits.addSheets) {
      if (!wb.SheetNames.includes(name)) {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['']]), name);
      }
    }
  }

  if (edits.renameSheets?.length) {
    for (const { from, to } of edits.renameSheets) {
      const idx = wb.SheetNames.indexOf(from);
      if (idx !== -1 && to && !wb.SheetNames.includes(to)) {
        wb.SheetNames[idx] = to;
        wb.Sheets[to] = wb.Sheets[from];
        delete wb.Sheets[from];
      }
    }
  }

  const ensureSheet = (sheetName?: string): XLSX.WorkSheet => {
    const target = sheetName && wb.Sheets[sheetName] ? sheetName : wb.SheetNames[0];
    if (!target) {
      const ws = XLSX.utils.aoa_to_sheet([['']]);
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      return ws;
    }
    return wb.Sheets[target];
  };

  if (edits.setCells?.length) {
    for (const item of edits.setCells) {
      const ws = ensureSheet(item.sheet);
      const addr =
        item.cell ||
        (item.row !== undefined && item.col !== undefined
          ? XLSX.utils.encode_cell({ r: item.row, c: item.col })
          : undefined);
      if (!addr) continue;

      if (item.formula) {
        const cleanFormula = item.formula.startsWith('=') ? item.formula.slice(1) : item.formula;
        ws[addr] = { t: 'n', f: cleanFormula, v: typeof item.value === 'number' ? item.value : 0 };
      } else if (typeof item.value === 'number') {
        ws[addr] = { t: 'n', v: item.value };
      } else if (typeof item.value === 'boolean') {
        ws[addr] = { t: 'b', v: item.value };
      } else {
        const strVal = item.value === undefined || item.value === null ? '' : String(item.value);
        if (strVal.startsWith('=')) {
          ws[addr] = { t: 'n', f: strVal.slice(1), v: 0 };
        } else {
          ws[addr] = { t: 's', v: strVal };
        }
      }

      const decoded = XLSX.utils.decode_cell(addr);
      const currentRef = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: decoded };
      currentRef.s.r = Math.min(currentRef.s.r, decoded.r);
      currentRef.s.c = Math.min(currentRef.s.c, decoded.c);
      currentRef.e.r = Math.max(currentRef.e.r, decoded.r);
      currentRef.e.c = Math.max(currentRef.e.c, decoded.c);
      ws['!ref'] = XLSX.utils.encode_range(currentRef);
    }
  }

  if (edits.appendRows?.length) {
    for (const group of edits.appendRows) {
      const ws = ensureSheet(group.sheet);
      XLSX.utils.sheet_add_aoa(ws, group.rows, { origin: -1 });
    }
  }

  if (ext === '.csv') {
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const csv = XLSX.utils.sheet_to_csv(sheet);
    return Buffer.from(`\uFEFF${csv}`, 'utf-8');
  }
  if (ext === '.tsv') {
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const tsv = XLSX.utils.sheet_to_csv(sheet, { FS: '\t' });
    return Buffer.from(`\uFEFF${tsv}`, 'utf-8');
  }
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
}

const PPTX_MEDIA_MIME: Record<string, string> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml',
};

function makeSlideXml(title: string, bullets: string[]): string {
  const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const esc = (s: string) =>
    s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  const bulletParas = (bullets.length ? bullets : ['']).map(
    (line) => `<a:p><a:r><a:rPr lang="zh-CN" sz="2000"/><a:t>${esc(line)}</a:t></a:r></a:p>`,
  );

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="457200"/><a:ext cx="10820400" cy="1143000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="3600" b="1"/><a:t>${esc(title)}</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Content 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="1828800"/><a:ext cx="10820400" cy="4343400"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>${bulletParas.join('')}</p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

async function buildMinimalPptxBuffer(slides: Array<{ title: string; bullets: string[] }>): Promise<Buffer> {
  const zip = new JSZip();
  const count = Math.max(1, slides.length);
  const slideOverrides = Array.from(
    { length: count },
    (_, i) =>
      `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
  ).join('\n  ');

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${slideOverrides}
</Types>`,
  );

  zip.folder('_rels')!.file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`,
  );

  const sldIds = Array.from(
    { length: count },
    (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`,
  ).join('');
  const rels = Array.from(
    { length: count },
    (_, i) =>
      `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
  ).join('');

  zip.folder('ppt')!.file(
    'presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst>${sldIds}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`,
  );
  zip.folder('ppt')!.folder('_rels')!.file(
    'presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
  );

  const items = slides.length ? slides : [{ title: 'New Presentation', bullets: [''] }];
  items.forEach((s, idx) => {
    zip.folder('ppt')!.folder('slides')!.file(`slide${idx + 1}.xml`, makeSlideXml(s.title, s.bullets));
  });

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export async function renderPptxBuffer(buf: Buffer): Promise<PptxRenderResult> {
  const zip = await JSZip.loadAsync(buf);
  const parser = new DOMParser();

  let slideWidthEmu = 12192000;
  let slideHeightEmu = 6858000;
  const presFile = zip.file('ppt/presentation.xml');
  if (presFile) {
    const presDoc = parser.parseFromString(await presFile.async('string'), 'text/xml');
    const sldSz = descendantsByLocalName(presDoc, 'sldSz')[0];
    if (sldSz) {
      slideWidthEmu = Number(sldSz.getAttribute('cx')) || slideWidthEmu;
      slideHeightEmu = Number(sldSz.getAttribute('cy')) || slideHeightEmu;
    }
  }

  const slidePaths = Object.keys(zip.files)
    .filter((k) => /^ppt\/slides\/slide\d+\.xml$/i.test(k))
    .sort((a, b) => {
      const na = Number(/slide(\d+)\.xml/i.exec(a)?.[1] || 0);
      const nb = Number(/slide(\d+)\.xml/i.exec(b)?.[1] || 0);
      return na - nb;
    });

  const widthPt = Math.round(slideWidthEmu / 12700);
  const heightPt = Math.round(slideHeightEmu / 12700);
  let remainingImageBudget = 450 * 1024;
  const slides: PptxSlideModel[] = [];

  for (let i = 0; i < slidePaths.length; i++) {
    const slidePath = slidePaths[i];
    const slideXml = await zip.file(slidePath)!.async('string');
    const slideDoc = parser.parseFromString(slideXml, 'text/xml');

    const relPath = slidePath.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels';
    const relMap = new Map<string, string>();
    const relFile = zip.file(relPath);
    if (relFile) {
      const relDoc = parser.parseFromString(await relFile.async('string'), 'text/xml');
      for (const rel of descendantsByLocalName(relDoc, 'Relationship')) {
        const id = rel.getAttribute('Id');
        const target = rel.getAttribute('Target');
        if (id && target) {
          const resolved = path.posix.normalize(path.posix.join('ppt/slides', target));
          relMap.set(id, resolved);
        }
      }
    }

    const shapes: PptxShapeModel[] = [];
    const plainLines: string[] = [];
    let title = '';
    let shapeSeq = 0;

    const extractXfrm = (el: XmlEl) => {
      const xfrm = descendantsByLocalName(el, 'xfrm')[0];
      const off = xfrm ? childElements(xfrm, 'off')[0] : undefined;
      const ext = xfrm ? childElements(xfrm, 'ext')[0] : undefined;
      const x = Number(off?.getAttribute('x') || 0);
      const y = Number(off?.getAttribute('y') || 0);
      const cx = Number(ext?.getAttribute('cx') || slideWidthEmu * 0.8);
      const cy = Number(ext?.getAttribute('cy') || slideHeightEmu * 0.2);
      return {
        leftPct: Math.max(0, Math.min(95, (x / slideWidthEmu) * 100)),
        topPct: Math.max(0, Math.min(95, (y / slideHeightEmu) * 100)),
        widthPct: Math.max(5, Math.min(100, (cx / slideWidthEmu) * 100)),
        heightPct: Math.max(5, Math.min(100, (cy / slideHeightEmu) * 100)),
      };
    };

    for (const sp of descendantsByLocalName(slideDoc, 'sp')) {
      const ph = descendantsByLocalName(sp, 'ph')[0];
      const phType = ph?.getAttribute('type') || '';
      const txBody = childElements(sp, 'txBody')[0];
      if (!txBody) continue;

      const paragraphs: PptxParagraphModel[] = [];
      for (const pEl of childElements(txBody, 'p')) {
        const pPr = childElements(pEl, 'pPr')[0];
        const algn = pPr?.getAttribute('algn');
        const align: PptxParagraphModel['align'] =
          algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : 'left';
        const texts = descendantsByLocalName(pEl, 't')
          .map((t) => t.textContent || '')
          .join('');
        if (!texts.trim()) continue;
        const firstRPr = descendantsByLocalName(pEl, 'rPr')[0];
        const szAttr = firstRPr?.getAttribute('sz');
        const fontSizePt = szAttr ? Math.round(Number(szAttr) / 100) : phType.includes('title') ? 32 : 18;
        const bold = firstRPr?.getAttribute('b') === '1' || phType.includes('title');
        const colorEl = firstRPr ? descendantsByLocalName(firstRPr, 'srgbClr')[0] : undefined;
        const colorHex = colorEl?.getAttribute('val') ? `#${colorEl.getAttribute('val')}` : undefined;
        paragraphs.push({ text: texts, align, fontSizePt, bold, colorHex });
        plainLines.push(texts);
        if (!title && (phType.includes('title') || shapes.length === 0)) {
          title = texts;
        }
      }
      if (!paragraphs.length) continue;
      shapeSeq += 1;
      shapes.push({
        id: `sp-${shapeSeq}`,
        kind: 'text',
        ...extractXfrm(sp),
        paragraphs,
      });
    }

    for (const pic of descendantsByLocalName(slideDoc, 'pic')) {
      const blip = descendantsByLocalName(pic, 'blip')[0];
      const embedId = blip?.getAttribute('r:embed') || blip?.getAttribute('embed');
      const mediaPath = embedId ? relMap.get(embedId) : undefined;
      if (!mediaPath || !zip.file(mediaPath)) continue;
      const mime = PPTX_MEDIA_MIME[path.extname(mediaPath).toLowerCase()];
      if (!mime) continue;
      const imgBuf = await zip.file(mediaPath)!.async('nodebuffer');
      if (imgBuf.byteLength > remainingImageBudget) continue;
      remainingImageBudget -= imgBuf.byteLength;
      shapeSeq += 1;
      shapes.push({
        id: `pic-${shapeSeq}`,
        kind: 'image',
        ...extractXfrm(pic),
        imageDataUri: `data:${mime};base64,${imgBuf.toString('base64')}`,
      });
    }

    for (const gf of descendantsByLocalName(slideDoc, 'graphicFrame')) {
      const tbl = descendantsByLocalName(gf, 'tbl')[0];
      if (!tbl) continue;
      const tableRows = childElements(tbl, 'tr').map((tr) =>
        childElements(tr, 'tc').map((tc) =>
          descendantsByLocalName(tc, 't')
            .map((t) => t.textContent || '')
            .join('')
            .trim(),
        ),
      );
      if (!tableRows.length) continue;
      shapeSeq += 1;
      shapes.push({
        id: `tbl-${shapeSeq}`,
        kind: 'table',
        ...extractXfrm(gf),
        tableRows,
      });
    }

    slides.push({
      slideNumber: i + 1,
      title: title || `Slide ${i + 1}`,
      widthPt,
      heightPt,
      shapes,
      plainText: plainLines.join('\n'),
    });
  }

  return { slideCount: slides.length, widthPt, heightPt, slides };
}

export async function editPptxBuffer(baseBuf: Buffer | null, edits: PptxEditInput): Promise<Buffer> {
  if (!baseBuf) {
    const initialSlides =
      edits.appendSlides?.map((s) => ({ title: s.title, bullets: s.bullets || [] })) ||
      [{ title: 'Untitled Presentation', bullets: ['Created by Antigravity Office Runtime'] }];
    let buf = await buildMinimalPptxBuffer(initialSlides);
    if (edits.updateSlides?.length) {
      buf = await editPptxBuffer(buf, { updateSlides: edits.updateSlides });
    }
    return buf;
  }

  const zip = await JSZip.loadAsync(baseBuf);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  if (edits.updateSlides?.length) {
    for (const upd of edits.updateSlides) {
      const slidePath = `ppt/slides/slide${upd.slideNumber}.xml`;
      const slideFile = zip.file(slidePath);
      if (!slideFile) throw new Error(`Slide ${upd.slideNumber} not found`);
      const doc = parser.parseFromString(await slideFile.async('string'), 'text/xml');

      if (upd.replacements?.length) {
        const paras = descendantsByLocalName(doc, 'p');
        for (const rep of upd.replacements) {
          for (const pEl of paras) {
            replaceTextInParagraph(pEl, rep.find, rep.replace);
          }
        }
      }

      if (upd.title !== undefined || upd.bodyLines !== undefined) {
        const sps = descendantsByLocalName(doc, 'sp');
        let titleSp: XmlEl | undefined;
        let bodySp: XmlEl | undefined;
        for (const sp of sps) {
          const ph = descendantsByLocalName(sp, 'ph')[0];
          const type = ph?.getAttribute('type') || '';
          if (type.includes('title') && !titleSp) titleSp = sp;
          else if (!bodySp) bodySp = sp;
        }
        if (upd.title !== undefined && (titleSp || sps[0])) {
          const targetSp = titleSp || sps[0];
          const firstP = descendantsByLocalName(targetSp, 'p')[0];
          if (firstP) setParagraphText(doc, firstP, upd.title);
        }
        if (upd.bodyLines !== undefined && (bodySp || sps[1])) {
          const targetSp = bodySp || sps[1];
          const txBody = childElements(targetSp, 'txBody')[0];
          if (txBody) {
            const existingP = childElements(txBody, 'p');
            const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
            for (const p of existingP) txBody.removeChild(p);
            for (const line of upd.bodyLines.length ? upd.bodyLines : ['']) {
              const p = doc.createElementNS(A_NS, 'a:p');
              const r = doc.createElementNS(A_NS, 'a:r');
              const t = doc.createElementNS(A_NS, 'a:t');
              t.textContent = line;
              r.appendChild(t);
              p.appendChild(r);
              txBody.appendChild(p);
            }
          }
        }
      }

      zip.file(slidePath, serializer.serializeToString(doc));
    }
  }

  if (edits.appendSlides?.length) {
    const existingCount = Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/i.test(k)).length;
    const presFile = zip.file('ppt/presentation.xml');
    const relsFile = zip.file('ppt/_rels/presentation.xml.rels');
    const ctFile = zip.file('[Content_Types].xml');

    const presDoc = presFile ? parser.parseFromString(await presFile.async('string'), 'text/xml') : null;
    const relsDoc = relsFile ? parser.parseFromString(await relsFile.async('string'), 'text/xml') : null;
    const ctDoc = ctFile ? parser.parseFromString(await ctFile.async('string'), 'text/xml') : null;

    let nextIdx = existingCount;
    for (const s of edits.appendSlides) {
      nextIdx += 1;
      const slidePath = `ppt/slides/slide${nextIdx}.xml`;
      zip.file(slidePath, makeSlideXml(s.title, s.bullets || []));

      const relId = `rId${100 + nextIdx}`;
      if (relsDoc?.documentElement) {
        const relEl = relsDoc.createElement('Relationship');
        relEl.setAttribute('Id', relId);
        relEl.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
        relEl.setAttribute('Target', `slides/slide${nextIdx}.xml`);
        relsDoc.documentElement.appendChild(relEl);
      }
      if (presDoc) {
        const lst = descendantsByLocalName(presDoc, 'sldIdLst')[0];
        if (lst) {
          const sldId = presDoc.createElement('p:sldId');
          sldId.setAttribute('id', String(256 + nextIdx));
          sldId.setAttribute('r:id', relId);
          lst.appendChild(sldId);
        }
      }
      if (ctDoc?.documentElement) {
        const ov = ctDoc.createElement('Override');
        ov.setAttribute('PartName', `/${slidePath}`);
        ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
        ctDoc.documentElement.appendChild(ov);
      }
    }

    if (presDoc) zip.file('ppt/presentation.xml', serializer.serializeToString(presDoc));
    if (relsDoc) zip.file('ppt/_rels/presentation.xml.rels', serializer.serializeToString(relsDoc));
    if (ctDoc) zip.file('[Content_Types].xml', serializer.serializeToString(ctDoc));
  }

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export class WorktreeManager {
  private readonly rootDir: string;

  constructor(dataDir: string) {
    this.rootDir = path.join(dataDir, 'worktrees');
  }

  private workDir(filePath: string): string {
    return path.join(this.rootDir, hashPath(filePath));
  }

  private metaPath(filePath: string): string {
    return path.join(this.workDir(filePath), 'meta.json');
  }

  private draftPath(filePath: string): string {
    const ext = path.extname(filePath).toLowerCase() || '.bin';
    return path.join(this.workDir(filePath), `draft${ext}`);
  }

  private async loadMeta(filePath: string): Promise<WorktreeMeta | null> {
    try {
      const raw = await fs.readFile(this.metaPath(filePath), 'utf-8');
      return JSON.parse(raw) as WorktreeMeta;
    } catch {
      return null;
    }
  }

  private async saveMeta(filePath: string, meta: WorktreeMeta): Promise<void> {
    await writeAtomicBuffer(this.metaPath(filePath), Buffer.from(JSON.stringify(meta, null, 2), 'utf-8'));
  }

  async hasActiveDraft(filePath: string): Promise<boolean> {
    const resolved = resolveInputPath(filePath);
    const meta = await this.loadMeta(resolved);
    return Boolean(meta?.status === 'draft' && existsSync(this.draftPath(resolved)));
  }

  async readActiveBuffer(filePath: string, useDraft = true): Promise<Buffer> {
    const resolved = resolveInputPath(filePath);
    if (useDraft && (await this.hasActiveDraft(resolved))) {
      return fs.readFile(this.draftPath(resolved));
    }
    return fs.readFile(resolved);
  }

  async readActiveChunk(
    filePath: string,
    offset = 0,
    length = MAX_CHUNK_BYTES,
    useDraft = true,
  ): Promise<{ base64: string; offset: number; bytesRead: number; totalSize: number; eof: boolean; isDraft: boolean }> {
    const resolved = resolveInputPath(filePath);
    const isDraft = useDraft && (await this.hasActiveDraft(resolved));
    const targetPath = isDraft ? this.draftPath(resolved) : resolved;

    const safeOffset = Math.max(0, Math.floor(offset));
    const safeLength = Math.min(Math.max(1, Math.floor(length)), MAX_CHUNK_BYTES);
    const handle = await fs.open(targetPath, 'r');
    try {
      const stat = await handle.stat();
      const totalSize = stat.size;
      if (safeOffset >= totalSize) {
        return { base64: '', offset: safeOffset, bytesRead: 0, totalSize, eof: true, isDraft };
      }
      const toRead = Math.min(safeLength, totalSize - safeOffset);
      const buffer = Buffer.allocUnsafe(toRead);
      const { bytesRead } = await handle.read(buffer, 0, toRead, safeOffset);
      const slice = bytesRead === toRead ? buffer : buffer.subarray(0, bytesRead);
      return {
        base64: slice.toString('base64'),
        offset: safeOffset,
        bytesRead,
        totalSize,
        eof: safeOffset + bytesRead >= totalSize,
        isDraft,
      };
    } finally {
      await handle.close();
    }
  }

  private async archiveCurrentDiskVersion(
    resolvedPath: string,
    kind: WorktreeMeta['kind'],
    summary: string,
  ): Promise<WorktreeMeta> {
    const existingMeta = await this.loadMeta(resolvedPath);
    const ext = path.extname(resolvedPath).toLowerCase() || '.bin';
    const now = Date.now();
    const meta: WorktreeMeta = existingMeta ?? {
      originalPath: resolvedPath,
      fileName: path.basename(resolvedPath),
      kind,
      ext,
      status: 'clean',
      baseSha256: '',
      draftSize: 0,
      summary: '',
      createdAt: now,
      updatedAt: now,
      versions: [],
    };

    if (existsSync(resolvedPath)) {
      const diskBuf = await fs.readFile(resolvedPath);
      const versionId = `v_${now}_${sha256(diskBuf).slice(0, 6)}`;
      const versionFileName = `${versionId}${ext}`;
      const versionDest = path.join(this.workDir(resolvedPath), 'versions', versionFileName);
      await writeAtomicBuffer(versionDest, diskBuf);
      meta.versions = [
        { versionId, fileName: versionFileName, savedAt: now, size: diskBuf.byteLength, summary },
        ...meta.versions,
      ].slice(0, 20);
      meta.baseSha256 = sha256(diskBuf);
    }

    return meta;
  }

  async stageOrWrite(
    rawPath: string,
    kind: WorktreeMeta['kind'],
    nextBuffer: Buffer,
    mode: 'draft' | 'direct',
    summary: string,
  ): Promise<WorktreeMeta> {
    const resolved = resolveInputPath(rawPath);
    const ext = path.extname(resolved).toLowerCase() || '.bin';
    const now = Date.now();

    if (mode === 'direct') {
      const meta = await this.archiveCurrentDiskVersion(resolved, kind, `Before direct edit: ${summary}`);
      await writeAtomicBuffer(resolved, nextBuffer);
      await fs.rm(this.draftPath(resolved), { force: true });
      meta.status = 'clean';
      meta.baseSha256 = sha256(nextBuffer);
      meta.draftSize = 0;
      meta.summary = summary;
      meta.updatedAt = now;
      await this.saveMeta(resolved, meta);
      return meta;
    }

    const existing = await this.loadMeta(resolved);
    const baseSha256 = existsSync(resolved) ? sha256(await fs.readFile(resolved)) : '';
    const meta: WorktreeMeta = existing ?? {
      originalPath: resolved,
      fileName: path.basename(resolved),
      kind,
      ext,
      status: 'draft',
      baseSha256,
      draftSize: nextBuffer.byteLength,
      summary,
      createdAt: now,
      updatedAt: now,
      versions: [],
    };
    meta.status = 'draft';
    meta.baseSha256 = baseSha256;
    meta.draftSize = nextBuffer.byteLength;
    meta.summary = summary;
    meta.updatedAt = now;

    await writeAtomicBuffer(this.draftPath(resolved), nextBuffer);
    await this.saveMeta(resolved, meta);
    return meta;
  }

  async getStatus(rawPath?: string): Promise<{ items: WorktreeMeta[] }> {
    if (rawPath) {
      const resolved = resolveInputPath(rawPath);
      const meta = await this.loadMeta(resolved);
      return { items: meta ? [meta] : [] };
    }
    if (!existsSync(this.rootDir)) return { items: [] };
    const dirs = await fs.readdir(this.rootDir, { withFileTypes: true });
    const items: WorktreeMeta[] = [];
    for (const d of dirs) {
      if (!d.isDirectory()) continue;
      try {
        const raw = await fs.readFile(path.join(this.rootDir, d.name, 'meta.json'), 'utf-8');
        items.push(JSON.parse(raw) as WorktreeMeta);
      } catch {
        // Ignore broken metadata
      }
    }
    items.sort((a, b) => b.updatedAt - a.updatedAt);
    return { items };
  }

  async approve(rawPath: string): Promise<WorktreeMeta> {
    const resolved = resolveInputPath(rawPath);
    const meta = await this.loadMeta(resolved);
    const dPath = this.draftPath(resolved);
    if (!meta || meta.status !== 'draft' || !existsSync(dPath)) {
      throw new Error(`No active draft worktree for ${resolved}`);
    }
    const draftBuf = await fs.readFile(dPath);
    const updatedMeta = await this.archiveCurrentDiskVersion(resolved, meta.kind, `Before approve: ${meta.summary}`);
    await writeAtomicBuffer(resolved, draftBuf);
    await fs.rm(dPath, { force: true });
    updatedMeta.status = 'clean';
    updatedMeta.baseSha256 = sha256(draftBuf);
    updatedMeta.draftSize = 0;
    updatedMeta.summary = `Merged: ${meta.summary}`;
    updatedMeta.updatedAt = Date.now();
    await this.saveMeta(resolved, updatedMeta);
    return updatedMeta;
  }

  async discard(rawPath: string): Promise<WorktreeMeta | null> {
    const resolved = resolveInputPath(rawPath);
    const meta = await this.loadMeta(resolved);
    await fs.rm(this.draftPath(resolved), { force: true });
    if (!meta) return null;
    meta.status = 'clean';
    meta.draftSize = 0;
    meta.summary = 'Draft discarded';
    meta.updatedAt = Date.now();
    await this.saveMeta(resolved, meta);
    return meta;
  }

  async restoreVersion(rawPath: string, versionId: string, mode: 'draft' | 'direct' = 'draft'): Promise<WorktreeMeta> {
    const resolved = resolveInputPath(rawPath);
    const meta = await this.loadMeta(resolved);
    const entry = meta?.versions.find((v) => v.versionId === versionId);
    if (!meta || !entry) {
      throw new Error(`Version ${versionId} not found for ${resolved}`);
    }
    const vBuf = await fs.readFile(path.join(this.workDir(resolved), 'versions', entry.fileName));
    return this.stageOrWrite(resolved, meta.kind, vBuf, mode, `Restored ${versionId}`);
  }

  async computeDiff(rawPath: string): Promise<Record<string, unknown>> {
    const resolved = resolveInputPath(rawPath);
    const meta = await this.loadMeta(resolved);
    const dPath = this.draftPath(resolved);
    if (!meta || meta.status !== 'draft' || !existsSync(dPath)) {
      return { hasDraft: false, path: resolved, summary: 'No active draft changes.' };
    }
    const draftBuf = await fs.readFile(dPath);
    const trunkExists = existsSync(resolved);
    const trunkBuf = trunkExists ? await fs.readFile(resolved) : null;

    if (meta.kind === 'docx') {
      const before = trunkBuf ? await inspectDocxBuffer(trunkBuf) : { paragraphs: [], tables: [], paragraphCount: 0, tableCount: 0 };
      const after = await inspectDocxBuffer(draftBuf);
      const maxLen = Math.max(before.paragraphs.length, after.paragraphs.length);
      const changedParagraphs: Array<{ index: number; before: string; after: string }> = [];
      for (let i = 0; i < maxLen; i++) {
        const b = before.paragraphs[i]?.text ?? '';
        const a = after.paragraphs[i]?.text ?? '';
        if (b !== a) {
          changedParagraphs.push({ index: i, before: b, after: a });
        }
      }
      return {
        hasDraft: true,
        kind: 'docx',
        path: resolved,
        isNewFile: !trunkExists,
        summary: meta.summary,
        changedCount: changedParagraphs.length,
        changedParagraphs: changedParagraphs.slice(0, 80),
      };
    }

    if (meta.kind === 'spreadsheet') {
      const readWb = (b: Buffer) => {
        const ext = path.extname(resolved).toLowerCase();
        if (ext === '.csv' || ext === '.tsv') {
          return XLSX.read(b.toString('utf-8'), { type: 'string', FS: ext === '.tsv' ? '\t' : undefined });
        }
        return XLSX.read(b, { type: 'buffer', codepage: 65001 });
      };
      const wbBefore = trunkBuf ? readWb(trunkBuf) : XLSX.utils.book_new();
      const wbAfter = readWb(draftBuf);
      const changedCells: Array<{ sheet: string; cell: string; before: string; after: string }> = [];
      for (const sheetName of wbAfter.SheetNames) {
        const wsA = wbAfter.Sheets[sheetName];
        const wsB = wbBefore.Sheets[sheetName];
        const rowsA = XLSX.utils.sheet_to_json<unknown[]>(wsA, { header: 1, defval: '', raw: false });
        const rowsB = wsB ? XLSX.utils.sheet_to_json<unknown[]>(wsB, { header: 1, defval: '', raw: false }) : [];
        const maxR = Math.max(rowsA.length, rowsB.length);
        for (let r = 0; r < maxR && changedCells.length < 100; r++) {
          const rA = Array.isArray(rowsA[r]) ? rowsA[r] : [];
          const rB = Array.isArray(rowsB[r]) ? rowsB[r] : [];
          const maxC = Math.max(rA.length, rB.length);
          for (let c = 0; c < maxC && changedCells.length < 100; c++) {
            const vA = String(rA[c] ?? '');
            const vB = String(rB[c] ?? '');
            if (vA !== vB) {
              changedCells.push({
                sheet: sheetName,
                cell: XLSX.utils.encode_cell({ r, c }),
                before: vB,
                after: vA,
              });
            }
          }
        }
      }
      return {
        hasDraft: true,
        kind: 'spreadsheet',
        path: resolved,
        isNewFile: !trunkExists,
        summary: meta.summary,
        changedCount: changedCells.length,
        changedCells,
      };
    }

    const slidesBefore = trunkBuf ? await renderPptxBuffer(trunkBuf) : { slides: [] };
    const slidesAfter = await renderPptxBuffer(draftBuf);
    const maxS = Math.max(slidesBefore.slides.length, slidesAfter.slides.length);
    const changedSlides: Array<{ slideNumber: number; before: string; after: string }> = [];
    for (let i = 0; i < maxS; i++) {
      const b = slidesBefore.slides[i]?.plainText ?? '';
      const a = slidesAfter.slides[i]?.plainText ?? '';
      if (b !== a) {
        changedSlides.push({ slideNumber: i + 1, before: b, after: a });
      }
    }
    return {
      hasDraft: true,
      kind: 'slides',
      path: resolved,
      isNewFile: !trunkExists,
      summary: meta.summary,
      changedCount: changedSlides.length,
      changedSlides,
    };
  }
}
