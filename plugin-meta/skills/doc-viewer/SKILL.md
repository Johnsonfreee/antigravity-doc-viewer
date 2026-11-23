---
name: doc-viewer
description: Open, browse, and preview PDFs, Markdown, Word (.docx), Excel/CSV, local HTML, images, and code files directly in Antigravity's sidebar via the Doc & PDF Viewer MCP App. Activate when the user asks to view, preview, open, or inspect a document or PDF in the sidebar.
---

# Doc & PDF Viewer

Use the `doc-viewer` MCP server tools whenever the user wants to open or preview documents in Antigravity's sidebar (`aux_pane`) or global left sidebar (`global`).

## Tools

1. **`open_document_viewer`**
   - Opens the interactive sidebar viewer (and updates the already-open pane when called again).
   - Optional arguments:
     - `path`: Absolute or workspace-relative path to a file (`.pdf`, `.md`, `.docx`, `.xlsx`, `.xls`, `.ods`, `.csv`, `.tsv`, `.html`, `.htm`, `.png`, `.jpg`, `.svg`, `.json`, code/text).
     - `rootDir`: Initial directory shown in the sidebar file tree.
     - `page`: Initial 1-based page number (for PDF or paginated spreadsheet).
     - `sheet`: Initial sheet name (for Excel workbooks).
     - `query`: Initial search query inside the document or file browser.
   - Calling with `{}` opens the viewer at the user's recent files / workspace root.

2. **`list_documents`**
   - Searches a directory recursively for supported document files (`pdf`, `md`, `docx`, `xlsx`, `csv`, `html`, etc.) and returns matching file paths and sizes so you can locate the right file before calling `open_document_viewer`.

## Best Practices

- Prefer passing `path` directly to `open_document_viewer` once you know the file path; the tool returns lightweight metadata only while the UI loads PDF bytes and rendered content offline via app-only tools.
- When the user mentions a specific PDF page or Excel sheet, pass `page` or `sheet` so the sidebar jumps straight there.
- Keep `rootDir` set to the active workspace root when opening files without an explicit root so the file browser starts in the project folder.
