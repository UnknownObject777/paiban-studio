# Vendored: onlyoffice-web-comp

- **来源**：https://github.com/electroluxcode/onlyoffice-web-comp
- **快照**：commit `99dbe255893b82fba46b82b337c571eb7d41f668`（main，2026-08-10 拉取）
- **许可证**：AGPL-3.0（见 `LICENSE.upstream`）；内置 OnlyOffice Developer Edition 9.4 SDK（静态资产不入本目录，由 `npm run fetch:onlyoffice` 拉取到 `public/packages/onlyoffice/`）
- **拷贝范围**：上游 `src/components/onlyoffice-web-comp/` 全量（含 `docs/` 使用文档）

## 本地适配（与上游的差异）

为纳入本项目 tsc（NodeNext + strict + 浏览器 ESM 直载）与 `paiban://` 自定义协议，做了以下机械修改：

1. 全部相对 import 补显式 `.js` 扩展名（目录导入补 `/index.js`）——`scripts/fix-onlyoffice-imports.mjs` 一次性改写。
2. `internal/editor/x2t.ts`：`new URL("./x2t.worker.ts", …)` → `./x2t.worker.js`（tsc 产物是同目录 .js）。
3. `const/index.ts`：`DEFAULT_ONLYOFFICE_ROOT` 前缀改为 `/public/packages/onlyoffice/…`，适配 `paiban://app/<项目根相对路径>` 的协议映射（见 `src/main.ts`）。
4. `internal/editor/office-format.ts`：`import("exceljs/…")` 改为变量间接持有 specifier（本项目未装 exceljs；仅 CSV→XLSX 路径使用，docx 预览不经过）。
5. `internal/editor/runtime-bridge.ts`：`options.debug ?? false`（strict 空值修复）。
6. `const/index.ts`：`resolveSiteUrl` 的绝对地址豁免从 `https?://` 放宽为任意 scheme（`paiban://` 下主线程已解析的绝对 URL 进 Worker 会被二次拼接）。
7. `internal/vendor/brotli-dec/**`：文件头加 `// @ts-nocheck`（JS 移植代码，无类型标注）。
8. 新增 `type/docs-api.d.ts`：声明运行时由 SDK 注入的 `window.DocsAPI`。

升级上游快照时重跑以上 1–2（脚本）并人工核对 3–5。
