import fs from 'node:fs/promises';
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
  DOCUMENT_EXTS,
  getDefaultRootDir,
  getFileMeta,
  getPdfAsset,
  getQuickRoots,
  listDir,
  MAX_CHUNK_BYTES,
  readFileChunk,
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

async function main() {
  const distDir = getRuntimeDistDir();
  const dataDir = parseDataDirArg();
  const appHtmlPath = process.env.DOC_VIEWER_APP_HTML?.trim() || path.join(distDir, 'app.html');
  const pdfAssetsRoot = path.join(distDir, 'pdfjs-assets');
  const store = new StateStore(dataDir);

  const server = new McpServer({
    name: 'doc-viewer',
    version: '1.0.0',
  });

  registerAppResource(
    server,
    'Doc & PDF Viewer UI',
    APP_URI,
    {
      description: 'Offline interactive viewer for PDF, Markdown, Word (.docx), Excel/CSV, HTML, images, and code.',
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
        'Open the interactive document & PDF viewer in Antigravity sidebar (or global view). All arguments are optional; calling with {} opens the workspace browser and recent files.',
      inputSchema: {
        path: z.string().optional().describe('Optional path to a document, PDF, spreadsheet, HTML, image, or code file.'),
        rootDir: z.string().optional().describe('Optional directory path to open in the sidebar file browser.'),
        page: z.number().int().min(1).optional().describe('Optional 1-based page number for PDFs or paginated tables.'),
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
        }

        const updatedState = await store.load();
        const payload = {
          file: fileMeta,
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
        'Recursively list PDF, Markdown, Word (.docx), Excel/CSV, HTML, and text documents in a directory.',
      inputSchema: {
        dir: z.string().optional().describe('Directory to scan (defaults to current working directory).'),
        query: z.string().optional().describe('Optional filename/path substring filter.'),
        extensions: z.array(z.string()).optional().describe('Optional extensions filter, e.g. [".pdf", ".docx"].'),
        maxDepth: z.number().int().min(1).max(10).optional().describe('Maximum recursion depth (default 5).'),
        limit: z.number().int().min(1).max(300).optional().describe('Max number of results (default 100).'),
      },
    },
    async (args) => {
      try {
        const root = args.dir ? resolveInputPath(args.dir) : process.cwd();
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
        const root = args.rootDir ? resolveInputPath(args.rootDir) : process.cwd();
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
      description: 'Resolve file metadata and update recent file history.',
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
        return okResult({ file: meta, prefs: state.filePrefs[meta.path] ?? {}, state });
      } catch (err) {
        return errorResult(err instanceof Error ? err.message : String(err));
      }
    },
  );

  registerAppTool(
    server,
    'app_read_file_chunk',
    {
      description: 'Read a binary file chunk as base64 (max 384 KiB raw per call).',
      inputSchema: {
        path: z.string(),
        offset: z.number().int().min(0).optional(),
        length: z.number().int().min(1).max(MAX_CHUNK_BYTES).optional(),
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, offset = 0, length = MAX_CHUNK_BYTES }) => {
      try {
        const chunk = await readFileChunk(rawPath, offset, length);
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
      description: 'Convert a Word (.docx) document to sanitized HTML with TOC.',
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
    'app_render_spreadsheet',
    {
      description: 'Parse and paginate Excel (.xlsx/.xls/.ods) or CSV/TSV files.',
      inputSchema: {
        path: z.string(),
        sheetName: z.string().optional(),
        page: z.number().int().min(1).optional(),
        pageSize: z.number().int().min(20).max(500).optional(),
        search: z.string().optional(),
      },
      _meta: appOnlyMeta,
    },
    async ({ path: rawPath, sheetName, page, pageSize, search }) => {
      try {
        return okResult(await renderSpreadsheet(rawPath, { sheetName, page, pageSize, search }));
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
