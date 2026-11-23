import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, '..');
const bundledServerPath = path.join(repoRoot, 'dist', 'plugin', 'dist', 'server.cjs');

const ONE_BY_ONE_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';

function buildMinimalTwoPagePdf() {
  const objects = [
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 5 0 R /Resources << /Font << /F1 7 0 R >> >> >>\nendobj\n',
    '4 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 6 0 R /Resources << /Font << /F1 7 0 R >> >> >>\nendobj\n',
    '5 0 obj\n<< /Length 44 >>\nstream\nBT /F1 18 Tf 72 720 Td (Page One Text) Tj ET\nendstream\nendobj\n',
    '6 0 obj\n<< /Length 44 >>\nstream\nBT /F1 18 Tf 72 720 Td (Page Two Text) Tj ET\nendstream\nendobj\n',
    '7 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n',
  ];

  let out = '%PDF-1.4\n';
  const offsets = [0];
  for (const obj of objects) {
    offsets.push(Buffer.byteLength(out, 'utf-8'));
    out += obj;
  }
  const xrefPos = Buffer.byteLength(out, 'utf-8');
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (let i = 1; i <= objects.length; i++) {
    out += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`;
  }
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefPos}\n%%EOF\n`;
  return Buffer.from(out, 'utf-8');
}

async function createMinimalDocx(outPath) {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.folder('_rels').file(
    '.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.folder('word').file(
    'document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>架构设计说明书</w:t></w:r></w:p>
    <w:p><w:r><w:t>专业：</w:t></w:r><w:r><w:t>软件</w:t></w:r><w:r><w:t>工程</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>项目</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>状态</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>模块A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>待测试</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`,
  );
  const buf = await zip.generateAsync({ type: 'nodebuffer' });
  await fs.writeFile(outPath, buf);
}

async function createFixtures(dir) {
  await fs.mkdir(path.join(dir, 'docs'), { recursive: true });
  await fs.mkdir(path.join(dir, 'assets'), { recursive: true });

  await fs.writeFile(path.join(dir, 'assets', 'pixel.png'), Buffer.from(ONE_BY_ONE_PNG_BASE64, 'base64'));
  await fs.writeFile(path.join(dir, 'docs', 'sample.pdf'), buildMinimalTwoPagePdf());
  await fs.writeFile(
    path.join(dir, 'docs', 'guide.md'),
    `# 项目指南\n\n## 安装步骤\n\n![Pixel](../assets/pixel.png)\n\n\`\`\`ts\nconst x: number = 42;\n\`\`\`\n\n<script>alert('xss')</script>\n`,
    'utf-8',
  );
  await createMinimalDocx(path.join(dir, 'docs', 'spec.docx'));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['姓名', '部门', '分数'],
      ['张三', '研发部', 98],
      ['李四', '产品部', 91],
    ]),
    '评分表',
  );
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Key', 'Value'], ['Env', 'Prod']]), '配置');
  XLSX.writeFile(wb, path.join(dir, 'docs', 'report.xlsx'));

  await fs.writeFile(
    path.join(dir, 'docs', 'data.csv'),
    '\uFEFF城市,人数,备注\n北京,120,测试行\n上海,95,第二行\n',
    'utf-8',
  );

  await fs.writeFile(path.join(dir, 'docs', 'page.css'), 'body { color: #123456; }', 'utf-8');
  await fs.writeFile(
    path.join(dir, 'docs', 'preview.html'),
    `<!DOCTYPE html><html><head><link rel="stylesheet" href="./page.css"></head><body><h1>Hello Local HTML</h1><img src="../assets/pixel.png"><script>window.pwned = true;</script></body></html>`,
    'utf-8',
  );
  await fs.writeFile(
    path.join(dir, 'docs', 'hello.ts'),
    `export function greet(name: string): string {\n  return \`Hello \${name}\`;\n}\n`,
    'utf-8',
  );
}

async function main() {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'agy-doc-viewer-test-'));
  const fixtureDir = path.join(tmpRoot, 'workspace');
  const dataDir = path.join(tmpRoot, 'plugin-data');
  await createFixtures(fixtureDir);

  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [bundledServerPath, `--data-dir=${dataDir}`],
    cwd: fixtureDir,
  });

  const client = new Client({ name: 'e2e-harness', version: '1.0.0' }, { capabilities: {} });

  try {
    await client.connect(transport);
    console.log('[e2e] Connected to standalone bundled server.cjs');

    // 1. Check tools/list metadata
    const toolsRes = await client.listTools();
    const openTool = toolsRes.tools.find((t) => t.name === 'open_document_viewer');
    assert.ok(openTool, 'open_document_viewer should be registered');
    assert.equal(openTool._meta?.ui?.resourceUri, 'ui://doc-viewer/app.html');
    assert.deepEqual(openTool._meta?.['com.google.antigravity/ui']?.entrypoints, [
      { type: 'aux_pane', title: 'Doc Viewer' },
      { type: 'global', title: 'Doc Viewer' },
    ]);

    for (const officeToolName of [
      'office_inspect_document',
      'office_edit_docx',
      'office_edit_xlsx',
      'office_edit_pptx',
      'office_manage_worktree',
    ]) {
      assert.ok(toolsRes.tools.some((t) => t.name === officeToolName), `${officeToolName} should be registered`);
    }

    const appOnlyTools = toolsRes.tools.filter((t) => t.name.startsWith('app_'));
    assert.ok(appOnlyTools.length >= 16, 'Expected all app-only tools');
    for (const t of appOnlyTools) {
      assert.deepEqual(t._meta?.ui?.visibility, ['app'], `${t.name} must be app-only`);
    }

    // 2. Check resources/list and resources/read
    const resourcesRes = await client.listResources();
    const uiRes = resourcesRes.resources.find((r) => r.uri === 'ui://doc-viewer/app.html');
    assert.ok(uiRes, 'ui://doc-viewer/app.html should be listed');
    assert.equal(uiRes.mimeType, 'text/html;profile=mcp-app');

    const readUi = await client.readResource({ uri: 'ui://doc-viewer/app.html' });
    assert.equal(readUi.contents[0].mimeType, 'text/html;profile=mcp-app');
    assert.ok(readUi.contents[0].text.includes('<!DOCTYPE html>'));
    assert.ok(Buffer.byteLength(readUi.contents[0].text, 'utf-8') < 10 * 1024 * 1024);

    // 3. open_document_viewer with {} and with PDF path
    const emptyOpen = await client.callTool({ name: 'open_document_viewer', arguments: {} });
    assert.ok(!emptyOpen.isError);
    assert.equal(emptyOpen.structuredContent.rootDir, fixtureDir);

    const pdfPath = path.join(fixtureDir, 'docs', 'sample.pdf');
    const pdfOpen = await client.callTool({
      name: 'open_document_viewer',
      arguments: { path: 'docs/sample.pdf', page: 2 },
    });
    assert.ok(!pdfOpen.isError);
    assert.equal(pdfOpen.structuredContent.file.path, pdfPath);
    assert.equal(pdfOpen.structuredContent.initialPage, 2);
    assert.ok(Buffer.byteLength(JSON.stringify(pdfOpen), 'utf-8') < 1024 * 1024);

    // 4. list_documents & app_list_dir & app_search_files
    const listDocs = await client.callTool({ name: 'list_documents', arguments: { dir: fixtureDir } });
    assert.ok(listDocs.structuredContent.items.some((i) => i.name === 'sample.pdf'));
    assert.ok(listDocs.structuredContent.items.some((i) => i.name === 'spec.docx'));

    const listDir = await client.callTool({ name: 'app_list_dir', arguments: { dir: path.join(fixtureDir, 'docs') } });
    assert.ok(listDir.structuredContent.entries.length >= 6);

    const searchRes = await client.callTool({
      name: 'app_search_files',
      arguments: { rootDir: fixtureDir, query: 'report' },
    });
    assert.equal(searchRes.structuredContent.items.length, 1);
    assert.equal(searchRes.structuredContent.items[0].name, 'report.xlsx');

    // 5. PDF chunking + offline CJK CMap + StandardFont + traversal guard
    const chunk = await client.callTool({
      name: 'app_read_file_chunk',
      arguments: { path: pdfPath, offset: 0, length: 65536 },
    });
    const decodedPdf = Buffer.from(chunk.structuredContent.base64, 'base64').toString('utf-8');
    assert.ok(decodedPdf.startsWith('%PDF-1.4'));
    assert.equal(chunk.structuredContent.eof, true);

    const cmap = await client.callTool({
      name: 'app_get_pdf_asset',
      arguments: { kind: 'cmap', name: 'UniGB-UTF16-H' },
    });
    assert.ok(cmap.structuredContent.byteLength > 0);

    const font = await client.callTool({
      name: 'app_get_pdf_asset',
      arguments: { kind: 'standard_font', name: 'FoxitFixed.pfb' },
    });
    assert.ok(font.structuredContent.byteLength > 0);

    const badAsset = await client.callTool({
      name: 'app_get_pdf_asset',
      arguments: { kind: 'cmap', name: '../etc/passwd' },
    });
    assert.equal(badAsset.isError, true);

    // 6. Markdown rendering
    const md = await client.callTool({
      name: 'app_render_markdown',
      arguments: { path: path.join(fixtureDir, 'docs', 'guide.md') },
    });
    assert.ok(md.structuredContent.html.includes('data:image/png;base64,'));
    assert.ok(!md.structuredContent.html.includes('<script>'));
    assert.equal(md.structuredContent.toc.length, 2);

    // 7. DOCX inspect + cross-run draft edit + worktree diff + approve + version history + restore
    const docxPath = path.join(fixtureDir, 'docs', 'spec.docx');
    const inspectedDocx = await client.callTool({
      name: 'office_inspect_document',
      arguments: { path: docxPath },
    });
    assert.equal(inspectedDocx.structuredContent.kind, 'docx');
    assert.ok(inspectedDocx.structuredContent.paragraphs.some((p) => p.text.includes('软件工程')));

    const editDocxRes = await client.callTool({
      name: 'office_edit_docx',
      arguments: {
        path: docxPath,
        mode: 'draft',
        summary: '更新专业与表格状态',
        replacements: [{ find: '软件工程', replace: '计算机科学与技术' }],
        updateTableCells: [{ tableIndex: 0, row: 1, col: 1, text: '已通过' }],
        appendBlocks: [{ type: 'paragraph', text: '新增结语段落' }],
      },
    });
    assert.equal(editDocxRes.structuredContent.worktree.status, 'draft');

    // Disk file should NOT be modified while in draft mode!
    const diskCheck = await client.callTool({
      name: 'office_inspect_document',
      arguments: { path: docxPath, useDraft: false },
    });
    assert.ok(diskCheck.structuredContent.paragraphs.some((p) => p.text.includes('软件工程')));

    // Draft check SHOULD reflect edits!
    const draftCheck = await client.callTool({
      name: 'office_inspect_document',
      arguments: { path: docxPath, useDraft: true },
    });
    assert.ok(draftCheck.structuredContent.paragraphs.some((p) => p.text.includes('计算机科学与技术')));
    assert.equal(draftCheck.structuredContent.tables[0].rows[1][1], '已通过');
    assert.ok(draftCheck.structuredContent.paragraphs.some((p) => p.text === '新增结语段落'));

    const diffRes = await client.callTool({
      name: 'office_manage_worktree',
      arguments: { action: 'diff', path: docxPath },
    });
    assert.equal(diffRes.structuredContent.hasDraft, true);
    assert.ok(diffRes.structuredContent.changedCount >= 2);

    const approveRes = await client.callTool({
      name: 'office_manage_worktree',
      arguments: { action: 'approve', path: docxPath },
    });
    assert.equal(approveRes.structuredContent.approved, true);
    assert.equal(approveRes.structuredContent.worktree.status, 'clean');
    assert.ok(approveRes.structuredContent.worktree.versions.length >= 1);

    // Confirm disk file now has approved edits
    const postApproveDisk = await client.callTool({
      name: 'office_inspect_document',
      arguments: { path: docxPath, useDraft: false },
    });
    assert.ok(postApproveDisk.structuredContent.paragraphs.some((p) => p.text.includes('计算机科学与技术')));

    // Test restore_version + discard
    const prevVersionId = approveRes.structuredContent.worktree.versions[0].versionId;
    await client.callTool({
      name: 'office_manage_worktree',
      arguments: { action: 'restore_version', path: docxPath, versionId: prevVersionId, mode: 'draft' },
    });
    const restoredDraft = await client.callTool({
      name: 'office_inspect_document',
      arguments: { path: docxPath, useDraft: true },
    });
    assert.ok(restoredDraft.structuredContent.paragraphs.some((p) => p.text.includes('软件工程')));
    await client.callTool({
      name: 'office_manage_worktree',
      arguments: { action: 'discard', path: docxPath },
    });

    // 8. Excel & UTF-8 CSV + office_edit_xlsx
    const xlsxPath = path.join(fixtureDir, 'docs', 'report.xlsx');
    await client.callTool({
      name: 'office_edit_xlsx',
      arguments: {
        path: xlsxPath,
        mode: 'draft',
        summary: '增加公式汇总',
        setCells: [
          { sheet: '评分表', cell: 'A4', value: '平均分' },
          { sheet: '评分表', cell: 'C4', formula: '=AVERAGE(C2:C3)', value: 94.5 },
        ],
      },
    });
    const xlsxRes = await client.callTool({
      name: 'app_render_spreadsheet',
      arguments: { path: xlsxPath, sheetName: '评分表', useDraft: true },
    });
    assert.equal(xlsxRes.structuredContent.hasDraft, true);
    assert.equal(xlsxRes.structuredContent.formulas.C4, '=AVERAGE(C2:C3)');
    assert.equal(xlsxRes.structuredContent.rows[3].cells[0], '平均分');
    await client.callTool({ name: 'office_manage_worktree', arguments: { action: 'approve', path: xlsxPath } });

    const csvRes = await client.callTool({
      name: 'app_render_spreadsheet',
      arguments: { path: path.join(fixtureDir, 'docs', 'data.csv'), search: '上海' },
    });
    assert.equal(csvRes.structuredContent.filteredRows, 1);
    assert.equal(csvRes.structuredContent.rows[0].cells[0], '上海');

    // 9. PowerPoint (.pptx) create, edit, render
    const pptxPath = path.join(fixtureDir, 'docs', 'deck.pptx');
    await client.callTool({
      name: 'office_edit_pptx',
      arguments: {
        path: pptxPath,
        mode: 'direct',
        appendSlides: [
          { title: '人机交互实验总结', bullets: ['支持 A4 分页', '支持 Worktree 审核'] },
          { title: '下一步规划', bullets: ['完善公式引擎'] },
        ],
      },
    });
    const pptRendered = await client.callTool({
      name: 'app_render_pptx',
      arguments: { path: pptxPath },
    });
    assert.equal(pptRendered.structuredContent.slideCount, 2);
    assert.equal(pptRendered.structuredContent.slides[0].title, '人机交互实验总结');

    // 10. Local HTML, Code, State Persistence
    const htmlRes = await client.callTool({
      name: 'app_render_html',
      arguments: { path: path.join(fixtureDir, 'docs', 'preview.html') },
    });
    assert.ok(htmlRes.structuredContent.previewHtml.includes('#123456'));
    assert.ok(!htmlRes.structuredContent.previewHtml.includes('window.pwned'));

    const codeRes = await client.callTool({
      name: 'app_read_text',
      arguments: { path: path.join(fixtureDir, 'docs', 'hello.ts') },
    });
    assert.equal(codeRes.structuredContent.language, 'typescript');

    await client.callTool({
      name: 'app_save_state',
      arguments: {
        lastRootDir: fixtureDir,
        bookmarks: [{ path: pdfPath, name: 'sample.pdf', kind: 'pdf', addedAt: 12345 }],
      },
    });
    const st = await client.callTool({ name: 'app_get_state', arguments: {} });
    assert.equal(st.structuredContent.state.lastRootDir, fixtureDir);

    // 11. Real user experiment report verification (READ-ONLY!)
    const realReportPath = 'D:\\日常任务\\人机交互实验资料\\202405566406-曾凯锋-实验1报告.docx';
    if (existsSync(realReportPath)) {
      const realOpen = await client.callTool({
        name: 'open_document_viewer',
        arguments: { path: realReportPath },
      });
      assert.equal(realOpen.structuredContent.file.kind, 'docx');
      assert.ok(Buffer.byteLength(JSON.stringify(realOpen), 'utf-8') < 1024 * 1024);

      const realInspect = await client.callTool({
        name: 'office_inspect_document',
        arguments: { path: realReportPath, useDraft: false },
      });
      assert.equal(realInspect.structuredContent.kind, 'docx');
      const allText = realInspect.structuredContent.paragraphs.map((p) => p.text).join('\n');
      assert.ok(allText.includes('软件工程'));
      assert.ok(allText.includes('202405566406'));

      const realChunk = await client.callTool({
        name: 'app_read_file_chunk',
        arguments: { path: realReportPath, offset: 0, length: 393216 },
      });
      assert.ok(realChunk.structuredContent.bytesRead > 0);
      console.log(`[e2e] Verified real .docx report (${realOpen.structuredContent.file.size} bytes, ${realInspect.structuredContent.paragraphCount} paragraphs) read-only!`);
    }

    console.log('[e2e] All 11 Office & MCP end-to-end test groups passed!');
  } finally {
    await client.close();
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
}

main().catch((err) => {
  console.error('[e2e] FAILED:', err);
  process.exit(1);
});
