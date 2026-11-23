import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  registerAppResource,
  registerAppTool,
  RESOURCE_MIME_TYPE,
} from '@modelcontextprotocol/ext-apps/server';
import { z } from 'zod';
import {
  classifyFile,
  DOCUMENT_EXTS,
  getDefaultRootDir,
  getFileMeta,
  getPdfAsset,
  getQuickRoots,
  listDir,
  MAX_CHUNK_BYTES,
  resolveInputPath,
  searchDocuments,
} from './files.js';
import {
  renderDocx,
  renderHtmlFile,
  renderMarkdown,
  renderSpreadsheet,
  renderTextFile,
} from './renderers.js';
import {
  editDocxBuffer,
  editPptxBuffer,
  editSpreadsheetBuffer,
  inspectDocxBuffer,
  renderPptxBuffer,
  WorktreeManager,
} from './officeWorktree.js';
import { StateStore } from './state.js';

const APP_URI = 'ui://doc-viewer/app.html';

function parseDataDirArg(): string {
  for (const arg of process.argv.slice(2)) {
    if (arg.startsWith('--data-dir=')) {
      const value = arg.slice('--data-dir='.length).trim();
      if (value) return value;
    }
  }
  return path.join(os.tmpdir(), 'antigravity-doc-viewer-data');
}

function getRuntimeDistDir(): string {
  if (typeof __dirname === 'string') return __dirname;
  return path.dirname(path.resolve(process.argv[1] || process.cwd()));
}

function jsonText(value: unknown): string {
  return JSON.stringify(value);
}

function okResult(structuredContent: Record<string, unknown>, textSummary?: string) {
  return {
    content: [
      {
        type: 'text' as const,
        text: textSummary ?? jsonText(structuredContent),
      },
    ],
    structuredContent,
  };
}

function errorResult(message: string) {
  return {
    isError: true,
    content: [{ type: 'text' as const, text: message }],
  };
}

const docxEditSchema = {
  replacements: z
    .array(z.object({ find: z.string(), replace: z.string() }))
    .optional()
    .describe('Cross-run find & replace pairs preserving original typography and underlines.'),
  updateParagraphs: z
    .array(z.object({ index: z.number().int().min(0), text: z.string() }))
    .optional()
    .describe('Replace text of paragraphs by 0-based index from office_inspect_document.'),
  updateTableCells: z
    .array(
      z.object({
        tableIndex: z.number().int().min(0),
        row: z.number().int().min(0),
        col: z.number().int().min(0),
        text: z.string(),
      }),
    )
    .optional()
    .describe('Update individual table cell text while keeping table borders and styles.'),
  appendBlocks: z
    .array(
      z.object({
        type: z.enum(['heading1', 'heading2', 'paragraph', 'table']),
        text: z.string().optional(),
        rows: z.array(z.array(z.string())).optional(),
      }),
    )
    .optional()
    .describe('Append headings, paragraphs, or tables before the section end.'),
};

const xlsxEditSchema = {
  addSheets: z.array(z.string()).optional(),
  renameSheets: z.array(z.object({ from: z.string(), to: z.string() })).optional(),
  setCells: z
    .array(
      z.object({
        sheet: z.string().optional(),
        cell: z.string().optional().describe('Cell address like A1 or C5'),
        row: z.number().int().min(0).optional(),
        col: z.number().int().min(0).optional(),
        value: z.union([z.string(), z.number(), z.boolean()]).optional(),
        formula: z.string().optional().describe('Excel formula like =SUM(B2:B10)'),
      }),
    )
    .optional(),
  appendRows: z
    .array(
      z.object({
        sheet: z.string().optional(),
        rows: z.array(z.array(z.union([z.string(), z.number(), z.boolean()]))),
      }),
    )
    .optional(),
};

const pptxEditSchema = {
  updateSlides: z
    .array(
      z.object({
        slideNumber: z.number().int().min(1),
        title: z.string().optional(),
        bodyLines: z.array(z.string()).optional(),
        replacements: z.array(z.object({ find: z.string(), replace: z.string() })).optional(),
      }),
    )
    .optional(),
  appendSlides: z
    .array(
      z.object({
        title: z.string(),
        bullets: z.array(z.string()).optional(),
      }),
    )
    .optional(),
};

async function main() {
  const distDir = getRuntimeDistDir();
  const dataDir = parseDataDirArg();
  const appHtmlPath = process.env.DOC_VIEWER_APP_HTML?.trim() || path.join(distDir, 'app.html');
  const pdfAssetsRoot = path.join(distDir, 'pdfjs-assets');
  const store = new StateStore(dataDir);
  const worktrees = new WorktreeManager(dataDir);

  const server = new McpServer({
    name: 'doc-viewer',
    version: '1.1.0',
  });

  registerAppResource(
    server,
    'Doc & PDF Viewer UI',
    APP_URI,
    {
      description:
        'Offline Office & Document Viewer (Word A4 pagination, Excel/CSV grid, PowerPoint slides, PDF CJK, Markdown, HTML, worktree compare).',
      mimeType: RESOURCE_MIME_TYPE,
      _meta: {
        ui: {
          prefersBorder: false,
        },
      },
    },
    async () => {
      const html = await fs.readFile(appHtmlPath, 'utf-8');
      return {
        contents: [
          {
            uri: APP_URI,
            mimeType: RESOURCE_MIME_TYPE,
            text: html,
            _meta: {
              ui: {
                prefersBorder: false,
              },
            },
          },
        ],
      };
    },
  );

  registerAppTool(
    server,
    'open_document_viewer',
    {
      title: 'Doc & PDF Viewer',
      description:
        'Open the interactive Office & Document viewer in Antigravity sidebar (or global view). Supports true A4 paginated .docx, interactive .xlsx/.csv, .pptx slides, PDF (CJK), Markdown, HTML, and worktree review.',
      inputSchema: {
        path: z.string().optional().describe('Optional path to a .docx, .xlsx, .csv, .pptx, .pdf, .md, .html, image, or code file.'),
        rootDir: z.string().optional().describe('Optional directory path to open in the sidebar file browser.'),
        page: z.number().int().min(1).optional().describe('Optional 1-based page or slide number.'),
        sheet: z.string().optional().describe('Optional sheet name for Excel workbooks.'),
        query: z.string().optional().describe('Optional search query inside the document or file browser.'),
      },
      _meta: {
        ui: {
          resourceUri: APP_URI,
        },
        'com.google.antigravity/ui': {
          entrypoints: [
            { type: 'aux_pane', title: 'Doc Viewer' },
            { type: 'global', title: 'Doc Viewer' },
          ],
        },
      },
    },
    async (args) => {
      try {
        const state = await store.load();
        const fallbackRoot = getDefaultRootDir();
        const resolvedRoot = args.rootDir
          ? resolveInputPath(args.rootDir, fallbackRoot)
          : state.lastRootDir || fallbackRoot;

        let fileMeta = null;
        let worktree = null;
        if (args.path) {
          const candidatePath = resolveInputPath(args.path, resolvedRoot);
          fileMeta = await getFileMeta(candidatePath);
          await store.recordOpened(
            { path: fileMeta.path, name: fileMeta.name, kind: fileMeta.kind },
            {
              ...(args.page ? { page: args.page } : {}),
              ...(args.sheet ? { sheet: args.sheet } : {}),
            },
          );
          const wt = await worktrees.getStatus(fileMeta.path);
          worktree = wt.items[0] ?? null;
        }

        const updatedState = await store.load();
        const payload = {
          file: fileMeta,
          worktree,
          rootDir: fileMeta ? fileMeta.dir : resolvedRoot,
          initialPage: args.page ?? (fileMeta ? updatedState.filePrefs[fileMeta.path]?.page : undefined) ?? 1,
          initialSheet: args.sheet ?? (fileMeta ? updatedState.filePrefs[fileMeta.path]?.sheet : undefined),
          initialQuery: args.query ?? '',
          quickRoots: getQuickRoots(),
          state: updatedState,
        };

        const summary = fileMeta
          ? `Opened ${fileMeta.name} (${fileMeta.kind}, ${fileMeta.size} bytes) in Doc & PDF Viewer.`
          : `Opened Doc & PDF Viewer browser rooted at ${payload.rootDir}.`;
        return okResult(payload, summary);
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  server.registerTool(
    'list_documents',
    {
      title: 'List Documents',
      description:
        'Recursively list PDF, Word (.docx), Excel/CSV, PowerPoint (.pptx), Markdown, HTML, and text documents in a directory.',
      inputSchema: {
        dir: z.string().optional().describe('Directory to scan (defaults to home/workspace).'),
        query: z.string().optional().describe('Optional filename/path substring filter.'),
        extensions: z.array(z.string()).optional().describe('Optional extensions filter, e.g. [".docx", ".pdf"].'),
        maxDepth: z.number().int().min(1).max(10).optional().describe('Maximum recursion depth (default 5).'),
        limit: z.number().int().min(1).max(300).optional().describe('Max number of results (default 100).'),
      },
    },
    async (args) => {
      try {
        const root = args.dir ? resolveInputPath(args.dir) : getDefaultRootDir();
        const extensions = args.extensions?.length ? args.extensions : [...DOCUMENT_EXTS];
        const res = await searchDocuments(root, {
          query: args.query,
          extensions,
          maxDepth: args.maxDepth ?? 5,
          limit: args.limit ?? 100,
        });
        return okResult(res);
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  // ---------------- Univer-style Office Runtime Tools ----------------

  server.registerTool(
    'office_inspect_document',
    {
      title: 'Inspect Office Document',
      description:
        'Inspect structured content of Word (.docx paragraphs & tables), Excel/CSV (sheets & rows), or PowerPoint (.pptx slides) from disk or active draft worktree.',
      inputSchema: {
        path: z.string().describe('Path to .docx, .xlsx, .xls, .ods, .csv, .tsv, or .pptx file.'),
        useDraft: z.boolean().optional().describe('Inspect staged draft worktree if active (default true).'),
        sheetName: z.string().optional().describe('Sheet name for spreadsheet files.'),
        page: z.number().int().min(1).optional().describe('Page number for spreadsheet rows.'),
      },
    },
    async ({ path: rawPath, useDraft = true, sheetName, page = 1 }) => {
      try {
        const resolved = resolveInputPath(rawPath);
        const kind = classifyFile(resolved);
        const buf = await worktrees.readActiveBuffer(resolved, useDraft);
        const hasDraft = await worktrees.hasActiveDraft(resolved);

        if (kind === 'docx') {
          const info = await inspectDocxBuffer(buf);
          return okResult({ path: resolved, kind, hasDraft, ...info });
        }
        if (kind === 'spreadsheet') {
          const sheet = await renderSpreadsheet(resolved, {
            sheetName,
            page,
            pageSize: 100,
            bufferOverride: buf,
          });
          return okResult({ path: resolved, kind, hasDraft, ...sheet });
        }
        if (kind === 'slides') {
          const pres = await renderPptxBuffer(buf);
          const compactSlides = pres.slides.map((s) => ({
            slideNumber: s.slideNumber,
            title: s.title,
            plainText: s.plainText,
          }));
          return okResult({ path: resolved, kind, hasDraft, slideCount: pres.slideCount, slides: compactSlides });
        }
        throw new Error(`Unsupported Office inspection type: ${kind} (${resolved})`);
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  server.registerTool(
    'office_edit_docx',
    {
      title: 'Edit or Create Word (.docx)',
      description:
        'Create or edit a Word (.docx) document in-place while preserving original fonts, styles, images, and layout. Defaults to draft worktree mode so changes can be reviewed before merging.',
      inputSchema: {
        path: z.string().describe('Target .docx file path (created if it does not exist).'),
        mode: z.enum(['draft', 'direct']).optional().describe('draft (default, staged for review) or direct (writes to disk with version backup).'),
        summary: z.string().optional().describe('Short description of the edit.'),
        ...docxEditSchema,
      },
    },
    async ({ path: rawPath, mode = 'draft', summary, replacements, updateParagraphs, updateTableCells, appendBlocks }) => {
      try {
        const resolved = resolveInputPath(rawPath);
        const baseBuf = (await worktrees.hasActiveDraft(resolved))
          ? await worktrees.readActiveBuffer(resolved, true)
          : existsSync(resolved)
            ? await fs.readFile(resolved)
            : null;
        const nextBuf = await editDocxBuffer(baseBuf, {
          replacements,
          updateParagraphs,
          updateTableCells,
          appendBlocks,
        });
        const meta = await worktrees.stageOrWrite(
          resolved,
          'docx',
          nextBuf,
          mode,
          summary || 'Edited Word document',
        );
        return okResult({
          path: resolved,
          mode,
          worktree: meta,
          message:
            mode === 'draft'
              ? `Staged .docx edits to draft worktree for ${path.basename(resolved)}. Call open_document_viewer to preview/compare or office_manage_worktree(action="approve") to save.`
              : `Saved .docx changes directly to ${resolved} (previous version archived).`,
        });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  server.registerTool(
    'office_edit_xlsx',
    {
      title: 'Edit or Create Excel/CSV',
      description:
        'Create or edit an Excel (.xlsx) or CSV/TSV spreadsheet (cells, formulas, rows, sheets) in draft worktree or direct mode.',
      inputSchema: {
        path: z.string().describe('Target .xlsx / .csv / .tsv file path.'),
        mode: z.enum(['draft', 'direct']).optional(),
        summary: z.string().optional(),
        ...xlsxEditSchema,
      },
    },
    async ({ path: rawPath, mode = 'draft', summary, addSheets, renameSheets, setCells, appendRows }) => {
      try {
        const resolved = resolveInputPath(rawPath);
        const baseBuf = (await worktrees.hasActiveDraft(resolved))
          ? await worktrees.readActiveBuffer(resolved, true)
          : existsSync(resolved)
            ? await fs.readFile(resolved)
            : null;
        const nextBuf = editSpreadsheetBuffer(baseBuf, resolved, {
          addSheets,
          renameSheets,
          setCells,
          appendRows,
        });
        const meta = await worktrees.stageOrWrite(
          resolved,
          'spreadsheet',
          nextBuf,
          mode,
          summary || 'Edited spreadsheet',
        );
        return okResult({ path: resolved, mode, worktree: meta });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  server.registerTool(
    'office_edit_pptx',
    {
      title: 'Edit or Create PowerPoint (.pptx)',
      description: 'Create or edit PowerPoint (.pptx) slides in draft worktree or direct mode.',
      inputSchema: {
        path: z.string().describe('Target .pptx file path.'),
        mode: z.enum(['draft', 'direct']).optional(),
        summary: z.string().optional(),
        ...pptxEditSchema,
      },
    },
    async ({ path: rawPath, mode = 'draft', summary, updateSlides, appendSlides }) => {
      try {
        const resolved = resolveInputPath(rawPath);
        const baseBuf = (await worktrees.hasActiveDraft(resolved))
          ? await worktrees.readActiveBuffer(resolved, true)
          : existsSync(resolved)
            ? await fs.readFile(resolved)
            : null;
        const nextBuf = await editPptxBuffer(baseBuf, { updateSlides, appendSlides });
        const meta = await worktrees.stageOrWrite(
          resolved,
          'slides',
          nextBuf,
          mode,
          summary || 'Edited presentation',
        );
        return okResult({ path: resolved, mode, worktree: meta });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  server.registerTool(
    'office_manage_worktree',
    {
      title: 'Manage Office Worktree & Versions',
      description:
        'Check draft worktree status, compute semantic diffs, approve/merge draft to disk, discard draft, list version history, or restore a previous version.',
      inputSchema: {
        action: z.enum(['status', 'diff', 'approve', 'discard', 'history', 'restore_version']),
        path: z.string().optional().describe('Document path (optional for action="status").'),
        versionId: z.string().optional().describe('Required for action="restore_version".'),
        mode: z.enum(['draft', 'direct']).optional().describe('Target mode when restoring a version (default "draft").'),
      },
    },
    async ({ action, path: rawPath, versionId, mode = 'draft' }) => {
      try {
        if (action === 'status') {
          return okResult(await worktrees.getStatus(rawPath));
        }
        if (!rawPath) throw new Error(`path is required for action="${action}"`);
        if (action === 'diff') {
          return okResult(await worktrees.computeDiff(rawPath));
        }
        if (action === 'approve') {
          return okResult({ approved: true, worktree: await worktrees.approve(rawPath) });
        }
        if (action === 'discard') {
          return okResult({ discarded: true, worktree: await worktrees.discard(rawPath) });
        }
        if (action === 'history') {
          const st = await worktrees.getStatus(rawPath);
          return okResult({ path: resolveInputPath(rawPath), versions: st.items[0]?.versions ?? [] });
        }
        if (!versionId) throw new Error('versionId is required for action="restore_version"');
        return okResult({ restored: true, worktree: await worktrees.restoreVersion(rawPath, versionId, mode) });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  // ---------------- App-only tools (hidden from model) ----------------

  const appOnlyMeta = { ui: { visibility: ['app'] } };

  registerAppTool(
    server,
    'app_list_dir',
    {
      description: 'List directory contents for the viewer file tree.',
      inputSchema: { dir: z.string().optional() },
      _meta: appOnlyMeta,
    },
    async ({ dir }) => {
      try {
        const target = dir && dir.trim() !== '.' ? resolveInputPath(dir) : getDefaultRootDir();
        const listing = await listDir(target);
        return okResult({ ...listing, quickRoots: getQuickRoots() });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_search_files',
    {
      description: 'Search documents recursively inside a root directory.',
      inputSchema: {
        rootDir: z.string().optional(),
        query: z.string().optional(),
        extensions: z.array(z.string()).optional(),
        maxDepth: z.number().int().min(1).max(10).optional(),
        limit: z.number().int().min(1).max(300).optional(),
      },
      _meta: appOnlyMeta,
    },
    async (args) => {
      try {
        const root = args.rootDir ? resolveInputPath(args.rootDir) : getDefaultRootDir();
        const res = await searchDocuments(root, {
          query: args.query,
          extensions: args.extensions,
          maxDepth: args.maxDepth ?? 5,
          limit: args.limit ?? 150,
        });
        return okResult(res);
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_open_file_meta',
    {
      description: 'Resolve file metadata, worktree status, and update recent file history.',
      inputSchema: {
        path: z.string(),
        recordRecent: z.boolean().optional(),
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, recordRecent = true }) => {
      try {
        const meta = await getFileMeta(rawPath);
        let state = await store.load();
        if (recordRecent) {
          state = await store.recordOpened({ path: meta.path, name: meta.name, kind: meta.kind });
        }
        const wt = await worktrees.getStatus(meta.path);
        return okResult({
          file: meta,
          worktree: wt.items[0] ?? null,
          prefs: state.filePrefs[meta.path] ?? {},
          state,
        });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_read_file_chunk',
    {
      description: 'Read a binary file chunk as base64 (supports draft worktree override).',
      inputSchema: {
        path: z.string(),
        offset: z.number().int().min(0).optional(),
        length: z.number().int().min(1).max(MAX_CHUNK_BYTES).optional(),
        useDraft: z.boolean().optional(),
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, offset = 0, length = MAX_CHUNK_BYTES, useDraft = true }) => {
      try {
        const chunk = await worktrees.readActiveChunk(rawPath, offset, length, useDraft);
        return okResult(chunk);
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_get_pdf_asset',
    {
      description: 'Read a bundled offline PDF.js cmap or standard font file as base64.',
      inputSchema: {
        kind: z.enum(['cmap', 'standard_font']),
        name: z.string(),
      },
      _meta: appOnlyMeta,
    },
    async ({ kind, name }) => {
      try {
        const asset = await getPdfAsset(pdfAssetsRoot, kind, name);
        return okResult(asset);
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_render_markdown',
    {
      description: 'Render Markdown with syntax highlighting, TOC, and inlined local images.',
      inputSchema: { path: z.string() },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath }) => {
      try {
        return okResult(await renderMarkdown(rawPath));
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_render_docx',
    {
      description: 'Fallback HTML rendering of a Word (.docx) document.',
      inputSchema: { path: z.string() },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath }) => {
      try {
        return okResult(await renderDocx(rawPath));
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_inspect_docx',
    {
      description: 'Inspect structured paragraphs and tables of a .docx file for interactive editing.',
      inputSchema: { path: z.string(), useDraft: z.boolean().optional() },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, useDraft = true }) => {
      try {
        const buf = await worktrees.readActiveBuffer(rawPath, useDraft);
        return okResult({ ...(await inspectDocxBuffer(buf)), hasDraft: await worktrees.hasActiveDraft(rawPath) });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_save_docx_edits',
    {
      description: 'Save UI edits to a .docx file (draft worktree or direct).',
      inputSchema: {
        path: z.string(),
        mode: z.enum(['draft', 'direct']).optional(),
        summary: z.string().optional(),
        ...docxEditSchema,
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, mode = 'draft', summary, replacements, updateParagraphs, updateTableCells, appendBlocks }) => {
      try {
        const resolved = resolveInputPath(rawPath);
        const baseBuf = (await worktrees.hasActiveDraft(resolved))
          ? await worktrees.readActiveBuffer(resolved, true)
          : existsSync(resolved)
            ? await fs.readFile(resolved)
            : null;
        const nextBuf = await editDocxBuffer(baseBuf, {
          replacements,
          updateParagraphs,
          updateTableCells,
          appendBlocks,
        });
        const meta = await worktrees.stageOrWrite(resolved, 'docx', nextBuf, mode, summary || 'UI Word edit');
        return okResult({ worktree: meta });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_render_spreadsheet',
    {
      description: 'Parse and paginate Excel (.xlsx/.xls/.ods) or CSV/TSV files (respects draft worktree).',
      inputSchema: {
        path: z.string(),
        sheetName: z.string().optional(),
        page: z.number().int().min(1).optional(),
        pageSize: z.number().int().min(20).max(500).optional(),
        search: z.string().optional(),
        useDraft: z.boolean().optional(),
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, sheetName, page, pageSize, search, useDraft = true }) => {
      try {
        const hasDraft = useDraft && (await worktrees.hasActiveDraft(rawPath));
        const bufferOverride = hasDraft ? await worktrees.readActiveBuffer(rawPath, true) : undefined;
        return okResult({
          ...(await renderSpreadsheet(rawPath, { sheetName, page, pageSize, search, bufferOverride })),
          hasDraft,
        });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_save_spreadsheet_edits',
    {
      description: 'Save UI cell/formula/sheet edits to a spreadsheet (draft worktree or direct).',
      inputSchema: {
        path: z.string(),
        mode: z.enum(['draft', 'direct']).optional(),
        summary: z.string().optional(),
        ...xlsxEditSchema,
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, mode = 'draft', summary, addSheets, renameSheets, setCells, appendRows }) => {
      try {
        const resolved = resolveInputPath(rawPath);
        const baseBuf = (await worktrees.hasActiveDraft(resolved))
          ? await worktrees.readActiveBuffer(resolved, true)
          : existsSync(resolved)
            ? await fs.readFile(resolved)
            : null;
        const nextBuf = editSpreadsheetBuffer(baseBuf, resolved, {
          addSheets,
          renameSheets,
          setCells,
          appendRows,
        });
        const meta = await worktrees.stageOrWrite(
          resolved,
          'spreadsheet',
          nextBuf,
          mode,
          summary || 'UI Spreadsheet edit',
        );
        return okResult({ worktree: meta });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_render_pptx',
    {
      description: 'Parse a PowerPoint (.pptx) presentation into structured slides for UI rendering.',
      inputSchema: { path: z.string(), useDraft: z.boolean().optional() },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, useDraft = true }) => {
      try {
        const hasDraft = useDraft && (await worktrees.hasActiveDraft(rawPath));
        const buf = await worktrees.readActiveBuffer(rawPath, useDraft);
        return okResult({ ...(await renderPptxBuffer(buf)), hasDraft });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_save_pptx_edits',
    {
      description: 'Save UI slide edits to a .pptx file (draft worktree or direct).',
      inputSchema: {
        path: z.string(),
        mode: z.enum(['draft', 'direct']).optional(),
        summary: z.string().optional(),
        ...pptxEditSchema,
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, mode = 'draft', summary, updateSlides, appendSlides }) => {
      try {
        const resolved = resolveInputPath(rawPath);
        const baseBuf = (await worktrees.hasActiveDraft(resolved))
          ? await worktrees.readActiveBuffer(resolved, true)
          : existsSync(resolved)
            ? await fs.readFile(resolved)
            : null;
        const nextBuf = await editPptxBuffer(baseBuf, { updateSlides, appendSlides });
        const meta = await worktrees.stageOrWrite(resolved, 'slides', nextBuf, mode, summary || 'UI Slide edit');
        return okResult({ worktree: meta });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_get_worktree_status',
    {
      description: 'Get draft worktree status & version history for a file or all active drafts.',
      inputSchema: { path: z.string().optional() },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath }) => {
      try {
        return okResult(await worktrees.getStatus(rawPath));
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_worktree_action',
    {
      description: 'Run diff, approve, discard, history, or restore_version from the UI.',
      inputSchema: {
        action: z.enum(['diff', 'approve', 'discard', 'history', 'restore_version']),
        path: z.string(),
        versionId: z.string().optional(),
        mode: z.enum(['draft', 'direct']).optional(),
      },
      _meta: appOnlyMeta,
    },
    async ({ action, path: rawPath, versionId, mode = 'draft' }) => {
      try {
        if (action === 'diff') return okResult(await worktrees.computeDiff(rawPath));
        if (action === 'approve') return okResult({ approved: true, worktree: await worktrees.approve(rawPath) });
        if (action === 'discard') return okResult({ discarded: true, worktree: await worktrees.discard(rawPath) });
        if (action === 'history') {
          const st = await worktrees.getStatus(rawPath);
          return okResult({ versions: st.items[0]?.versions ?? [] });
        }
        if (!versionId) throw new Error('versionId is required');
        return okResult({ restored: true, worktree: await worktrees.restoreVersion(rawPath, versionId, mode) });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_create_office_file',
    {
      description: 'Create a new blank Word (.docx), Excel (.xlsx), or PowerPoint (.pptx) file.',
      inputSchema: {
        path: z.string(),
        kind: z.enum(['docx', 'xlsx', 'pptx']),
        title: z.string().optional(),
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, kind, title }) => {
      try {
        const resolved = resolveInputPath(rawPath);
        let buf: Buffer;
        if (kind === 'docx') {
          buf = await editDocxBuffer(null, {
            appendBlocks: [{ type: 'heading1', text: title || path.basename(resolved, '.docx') }],
          });
        } else if (kind === 'xlsx') {
          buf = editSpreadsheetBuffer(null, resolved, {
            setCells: [{ sheet: 'Sheet1', cell: 'A1', value: title || 'New Sheet' }],
          });
        } else {
          buf = await editPptxBuffer(null, {
            appendSlides: [{ title: title || path.basename(resolved, '.pptx'), bullets: [''] }],
          });
        }
        await fs.mkdir(path.dirname(resolved), { recursive: true });
        await fs.writeFile(resolved, buf);
        const meta = await getFileMeta(resolved);
        return okResult({ file: meta });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_render_html',
    {
      description: 'Prepare sanitized offline preview HTML and highlighted source for local HTML files.',
      inputSchema: { path: z.string() },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath }) => {
      try {
        return okResult(await renderHtmlFile(rawPath));
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_read_text',
    {
      description: 'Read and syntax-highlight a range of lines from a text/code file.',
      inputSchema: {
        path: z.string(),
        startLine: z.number().int().min(1).optional(),
        lineCount: z.number().int().min(50).max(1000).optional(),
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, startLine, lineCount }) => {
      try {
        return okResult(await renderTextFile(rawPath, { startLine, lineCount }));
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_get_state',
    {
      description: 'Load persisted recent files, bookmarks, and per-file viewer preferences.',
      inputSchema: {},
      _meta: appOnlyMeta,
    },
    async () => {
      try {
        return okResult({ state: await store.load() });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_save_state',
    {
      description: 'Persist recent files, bookmarks, root directory, or per-file preferences.',
      inputSchema: {
        lastRootDir: z.string().optional(),
        bookmarks: z
          .array(
            z.object({
              path: z.string(),
              name: z.string(),
              kind: z.string(),
              addedAt: z.number(),
            }),
          )
          .optional(),
        filePrefs: z
          .record(
            z.string(),
            z.object({
              page: z.number().optional(),
              zoom: z.number().optional(),
              fitMode: z.enum(['width', 'custom']).optional(),
              sheet: z.string().optional(),
              scrollY: z.number().optional(),
            }),
          )
          .optional(),
      },
      _meta: appOnlyMeta,
    },
    async (args) => {
      try {
        const state = await store.save(args);
        return okResult({ state });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error('[doc-viewer MCP server fatal]', err);
  process.exit(1);
});
