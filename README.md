# Antigravity Doc & Office Workbench (`doc-viewer`)

一个完全离线、零运行时 `node_modules` 依赖的 **Google Antigravity MCP App** 办公插件（参考 `dream-num/dsh-univer-office` 架构），用于在 Antigravity 侧边栏（`aux_pane`）和全局视图（`global`）提供桌面级真实办公阅读、编辑与 Worktree 审阅体验。

## 核心特性

- **仿桌面级暗色办公壳与 A4 纸张画布**：双层顶栏（带 `W`/`X`/`P`/`PDF` 彩色应用徽标的多标签栏 + 完整文件路径地址栏与视图切换器）配 `#141414` 深色画布与阴影纸张页面。
- **Word (`.docx`) 真实分页与原样式编辑**：
  - 前端内嵌离线 `docx-preview`，完整保留字体、下划线、样式、表格与嵌入图片，并在前端自动将单节文档按 A4 纸张高度切分为独立白纸页。
  - 服务端通过 OOXML 精准编辑（支持跨 `<w:r>` run 替换、段落编辑、表格单元格修改与追加内容），不破坏原有排版与图片。
- **Excel / CSV (`.xlsx`, `.xls`, `.ods`, `.csv`, `.tsv`) 交互式表格**：带 `fx` 公式栏、单元格双击编辑、常用公式计算（`SUM`/`AVERAGE`/`COUNT`/`MIN`/`MAX`）、工作表切换与分页筛选。
- **PowerPoint (`.pptx`) 幻灯片工作台**：左侧缩略图胶片条 + 16:9 主画布 + 内嵌图片/表格渲染 + 在线编辑与新增幻灯片。
- **Worktree 草稿隔离与版本历史（类似 `dsh-univer-office`）**：Agent 或用户在 UI 中的修改默认暂存在独立 `worktrees` 草稿区，支持语义 Diff 对比、一键 **Approve 合并保存到磁盘**、**Discard 丢弃** 及自动版本快照回滚。
- **PDF / Markdown / HTML / 代码全能支持**：离线 `pdfjs-dist`（含完整中文 CJK CMaps）、GFM Markdown、禁用脚本的 Shadow DOM 本地 HTML 预览、代码行号高亮以及划选文字直接 **引用到对话 / 问 Agent**。

## 从源码构建与安装

```bash
npm install
npm run typecheck
npm run build
npm test
npm run pack:zip
npm run install:plugin
```

安装目录：`~/.gemini/config/plugins/doc-viewer/`，Antigravity 会自动发现并加载。

## License

MIT
