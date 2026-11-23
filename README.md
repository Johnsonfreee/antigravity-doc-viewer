# Antigravity Doc & PDF Viewer (`doc-viewer`)

一个完全离线、零运行时 `node_modules` 依赖的 **Google Antigravity MCP App** 插件，用于在 Antigravity 右侧辅助侧边栏（`aux_pane`）和左侧全页栏（`global`）直接浏览、预览和检索本地文档与 PDF，并支持划词直接引用到当前对话。

## 支持格式与核心能力

- **PDF (`.pdf`)**：基于本地打包的 `pdfjs-dist@4.10.38` 主线程模式运行，内置完整 **CJK 中文 CMaps 与标准字体**，支持书签大纲导航、页码跳转、自适应宽度/缩放、文字选择（TextLayer）、全文档搜索与翻页进度持久化。
- **Word (`.docx`)**：使用 `mammoth` 在本地服务端转换为带样式 HTML，支持标题目录 TOC 跳转、内嵌图片体积预算控制、字号调节与纯文本复制。
- **Excel / CSV (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`, `.csv`, `.tsv`)**：支持多 Sheet 切换、UTF-8 中文 CSV 无乱码读取、搜索筛选、分页预览、首行作表头切换以及复制本页 CSV。
- **Markdown (`.md`, `.markdown`, `.mdx`)**：支持 GFM 表格、代码块语法高亮（`highlight.js`）、标题目录大纲与本地相对路径图片内联。
- **本地 HTML 预览 (`.html`, `.htm`)**：自动内联同目录局部 CSS 与图片并在禁用脚本的隔离 Shadow DOM 中安全预览，支持切换“隔离预览 / 高亮源码”以及 390px 移动端视口模拟。
- **图片与代码/日志**：支持 `.png/.jpg/.gif/.webp/.svg` 预览，以及代码/JSON/YAML/TXT 分页行号高亮显示。
- **与 Agent 双向联动**：
  - 对话中让 Agent 调用 `open_document_viewer` 直接在侧边栏打开目标文件并定位到指定页码或工作表。
  - 在阅读器中选中一段文本后，点击 **引用到对话**（`updateModelContext`）或 **问 Agent**（`sendMessage`）即可直接把内容喂给 Agent（全局侧边栏模式自动隐藏对话发送按钮）。

---

## 目录结构

```text
antigravity-doc-viewer/
├── plugin-meta/             # 插件清单、MCP 配置、Logo、内置 Skill
├── src/
│   ├── server/              # Node.js MCP Server（文件索引、分片传输、DOCX/XLSX/MD/HTML 解析）
│   └── web/                 # 单文件 HTML 前端（PDF.js + MCP Apps SDK）
├── scripts/
│   ├── build-offline.mjs    # 构建全离线插件包 dist/plugin/
│   ├── install.mjs          # 一键安装到 ~/.gemini/config/plugins/doc-viewer
│   └── pack-zip.mjs         # 打包独立离线压缩包 dist/antigravity-doc-viewer-offline.zip
└── tests/
    └── mcp-e2e.mjs          # Stdio MCP 端到端集成测试
```

---

## 从源码构建与安装

前置要求：Node.js `>= 20`。

```bash
npm install
npm run typecheck
npm run build
npm test
npm run install:plugin
```

安装完成后，插件会被写入 `~/.gemini/config/plugins/doc-viewer/`，Antigravity 会自动发现并加载，无需手动修改配置文件。若希望打包给无网络环境直接解压使用，可运行：

```bash
npm run pack:zip
```

将生成的 `dist/antigravity-doc-viewer-offline.zip` 解压后，把里面的 `doc-viewer` 目录放入 `~/.gemini/config/plugins/` 即可离线使用（运行时只需要本机装有 `node`，无需执行 `npm install`）。

---

## 架构要点与防坑设计

1. **单 HTML 全离线（`< 10 MiB`）**：通过 `vite-plugin-singlefile` 将前端与 `pdfjs-dist`（含主线程 `pdfjsWorker`）打包为单个 `app.html`，不请求任何外网 CDN，符合 MCP App 默认 `default-src 'none'` CSP 策略。
2. **绕过 `1 MiB` 轨迹载荷限制**：`open_document_viewer` 只返回文件元信息；PDF 二进制字节流、CJK CMaps、Markdown/DOCX/表格内容均通过 `visibility: ["app"]` 的内部工具按需加载与分片传输（最大并发控制在 `3` 以内）。
3. **状态后端持久化**：由于隔离 iframe 重载会清空 `localStorage`，最近文件、收藏夹和每份文档的上次页码/缩放比例统一由服务端保存在 `${PLUGIN_DATA}/state.json`。

## License

MIT
