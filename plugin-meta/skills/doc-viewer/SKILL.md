---
name: doc-viewer
description: Open, browse, inspect, edit, and review Word (.docx), Excel/CSV (.xlsx/.csv), PowerPoint (.pptx), PDF, Markdown, local HTML, images, and code directly in Antigravity's sidebar with isolated draft worktrees and A4 paginated print layout.
---

# Doc & PDF Viewer · Office Runtime

Use the `doc-viewer` MCP server whenever the user asks to view, preview, create, edit, or review Office documents (`.docx`, `.xlsx`, `.csv`, `.pptx`) or PDF/Markdown/HTML files in the sidebar.

## Tools

1. **`open_document_viewer`**
   - Opens the dark Office & Document workbench in the right auxiliary sidebar (`aux_pane`) or left global panel (`global`).
   - Supports true A4 paginated `.docx` rendering (`docx-preview`), interactive `.xlsx`/`.csv` formula grid, `.pptx` slide deck viewer, PDF (with offline CJK CMaps), Markdown, and local HTML.
   - Calling it again updates the currently open sidebar pane.

2. **`list_documents`**
   - Recursively searches a directory for `.docx`, `.xlsx`, `.csv`, `.pptx`, `.pdf`, `.md`, `.html`, and text files.

3. **Office Inspect & Edit Tools (`dsh-univer-office`-style workflow)**
   - `office_inspect_document`: Inspect paragraphs & tables of `.docx`, sheets & rows of `.xlsx`/`.csv`, or slides of `.pptx`.
   - `office_edit_docx`: Create or edit `.docx` files in-place preserving original fonts, underlines, styles, and images (`replacements`, `updateParagraphs`, `updateTableCells`, `appendBlocks`). Defaults to `mode: "draft"` (staged in an isolated worktree for review).
   - `office_edit_xlsx`: Create or edit `.xlsx`/`.csv` cells, formulas, rows, and sheets.
   - `office_edit_pptx`: Create or edit `.pptx` slides.
   - `office_manage_worktree`: Check `status`, compute semantic `diff`, `approve` (merge draft to disk with automatic version backup), `discard`, view `history`, or `restore_version`.

## Recommended Editing Flow

1. Inspect the target document with `office_inspect_document` first so you know paragraph indices, table coordinates, or sheet names.
2. Apply changes with `office_edit_docx` / `office_edit_xlsx` / `office_edit_pptx` using `mode: "draft"` (unless the user explicitly asked for direct overwrite).
3. Call `open_document_viewer` on that `path` so the user immediately sees the rendered A4/grid/slide draft and can compare changes, **Approve & Save**, or **Discard** right inside the sidebar!
