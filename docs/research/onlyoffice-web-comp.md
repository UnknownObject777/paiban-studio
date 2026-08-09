# onlyoffice-web-comp 调研报告（issue #25：能否拿来当文档看板/预览面板）

- 调研对象：<https://github.com/electroluxcode/onlyoffice-web-comp>
- 调研日期：2026-08-09
- 调研方法：GitHub API 读取仓库元数据 / README（中英）/ 组件库 docs 全套 Markdown / package.json / git tree 全量统计资产体积 / issues 与 commits 活跃度。
- 仓库快照：主语言 TypeScript，License **AGPL-3.0**，约 464 stars / 65 forks，创建于 2025-11-13，调研当日仍在活跃维护（最近 push 2026-07-30，约 175 commits，基本为单作者 electroluxcode）；版本 0.1.0，`private: true`，**未发布为 npm 包**。

---

## 1. 它是什么

### 1.1 架构：OnlyOffice 静态 SDK 的浏览器端封装，不是独立组件库

仓库分两部分：

| 部分 | 路径 | 说明 |
|---|---|---|
| 组件封装 | `src/components/onlyoffice-web-comp/` | TS 封装层：`OnlyOfficeManager` 门面 + `EditorManager` + `EventBus` + x2t Worker；配套 11 篇中文接入文档 |
| 演示站 | `src/app/` + `src/features/` | Next.js 15 + React 19 主页、文档站、单/多实例在线示例 |

核心技术路线：把 **OnlyOffice Developer Edition（v9.4.0）从 Docker 镜像里导出的静态资源**（`web-apps/` + `sdkjs/` + `fonts/` + `x2t/`）放到 `public/packages/onlyoffice/9.4.0-develop/`，前端直接加载 `api.js`，由 sdkjs 在浏览器内完成排版渲染（canvas 绘制、真实分页），由 **x2t-wasm**（cryptpad 的 onlyoffice-x2t-wasm）在 Web Worker 内做 docx ↔ OnlyOffice 内部二进制格式的转换。**不需要 OnlyOffice Document Server，也不需要任何在线服务**——这是它相对传统 OnlyOffice 集成（必须自建 Document Server）的核心差异。

### 1.2 集成方式：复制源码 + 复制静态资产

它**不是** `npm install` 就能用的包（`package.json` 里 `private: true`，README 明确说「不是一个 npm install 后直接引入的包，而是一个浏览器端 OnlyOffice 集成模板」）。官方集成路径：

1. 复制 `src/components/onlyoffice-web-comp/` 到自己项目源码目录；
2. 复制 `public/packages/onlyoffice/` 静态资产到自己的静态目录；
3. 参考 `src/features/demo/office-preview-page.tsx` 自己写界面壳：准备容器、维护 `OnlyOfficeManager` 实例、调 `openDocument` / `downloadExport` / `toggleReadOnly`，卸载时 `destroy()`。

封装本身与 React/Next.js 解耦程度一般（依赖 ahooks、lodash-es、nanoid，示例全是 React）；组件 `index.ts` 导出的核心类可脱离 React 用，但需要自行适配。

### 1.3 运行时依赖与 License

- **运行时**：纯浏览器端。sdkjs（排版/编辑引擎）+ x2t.wasm（格式转换）全部在客户端执行，无网络请求。文档数据不出设备，这一点与项目「本地优先」硬约束**方向一致**。
- **License：AGPL-3.0**（本仓库与上游 ONLYOFFICE web-apps/sdkjs 均为 AGPL-3.0）。仓库 issue #36 有用户明确问过闭源商用问题，作者转述 OnlyOffice 商务口径：**「使用商业版的资源不需要开放源码」——即闭源分发必须购买 OnlyOffice 商业授权**。另外，仓库内静态资产来自 Developer Edition Docker 导出，DE 本身是商业产品，把导出的资产再分发到自己的发行包里，授权边界需要与 OnlyOffice 商务确认（社区版 sdkjs/web-apps 源码自编译才是干净的 AGPL 路径）。

---

## 2. 相对 docx-preview 的优劣

| 维度 | docx-preview@0.4（现状） | onlyoffice-web-comp |
|---|---|---|
| 渲染保真度 | DOM 近似渲染；复杂表格、分页、页眉页脚、编号在公文场景下常有偏差 | sdkjs 自研排版引擎，canvas 绘制、真实分页，接近桌面 Word；表格/图片/页眉页脚/编号/分节符全支持 |
| OOXML 特性覆盖 | 子集 | 几乎全量，另含批注、修订（track changes）、目录、公式等；支持 .doc/.odt/.rtf 经 x2t 转换打开 |
| 交互能力 | 只读渲染 | **完整所见即所得编辑器**（可切只读），支持中/英界面、主题、多实例、导出 |
| 体积 | 数百 KB | 仓库内静态资产 **1.14 GB**（详见 §3.2）；裁剪到仅 Word 后约 **220 MB + 字体** |
| 首启/打开成本 | 秒开，`renderAsync` 0.3–0.7s | 首启需加载 sdkjs 数十 MB + x2t.wasm（约 10 MB Brotli）；**每次打开文档要跑一遍 x2t wasm 转换，秒级** |
| 刷新模型 | 无状态重渲：AI 改完 docx 重新 `renderAsync` 即可，0.3–0.7s | 有状态编辑器：外部改 docx 后**没有热刷新通道**，只能重新 x2t 转换 + 重开（秒级，且丢编辑态/滚动位置），或改用 Connector Automation API 直接操作编辑器内文档 |

**关键错配**：当前预览层的用法是「agent 直接改 docx → 预览面板快速重渲」。onlyoffice-web-comp 是有状态编辑器，不是无状态渲染器——把它当「看板」用，每次 AI 修改后的刷新成本从 0.3–0.7s 退化为秒级重开，体验反而倒退。要用好它，正确姿势是让 AI 通过 Connector Automation API 直接操作编辑器内的活文档，但那是**用 OnlyOffice 替换掉自研 OOXML 编辑层**（pizzip + fast-xml-parser）的架构级决策，远超「换一个预览组件」的范围。

已知质量问题（来自其 issues，多数已修但反映成熟度）：大文件 `Array buffer allocation fail`（#34）、PPT 超 10 页渲染失败（#16）、大 XLSX 分块导出回调错误（#25）、移动端滚动困难（#6）等。

---

## 3. 适配性评估

### 3.1 能否嵌进 Electron 渲染进程离线运行

**技术上可行**：

- 资产全静态，可走 Electron 自定义 protocol（如 `app://`）供给；sdkjs 的 `api.js` 内部会创建指向 `web-apps/` 的 iframe，用自定义 protocol + 合理 CSP 即可，无需网络。
- wasm / Web Worker 在 Electron 渲染进程天然可用。
- 数据不出设备，满足政企离线硬约束。

但要做的工程不少：自定义 protocol 注册、CSP 放行 wasm/worker/iframe、把 React 风格的封装适配进现有渲染层、字体注册（`__custom_font_registry__` + ttf-to-catalog 工具链）。其 issue #31（CDN 模式跨域加载失败）说明这类资源路径问题真实存在，只是 Electron 自定义 protocol 下可控。

### 3.2 包体积 / 分发成本（git tree 全量统计）

`public/packages/onlyoffice/9.4.0-develop/` 共 **1.14 GB、约 1.96 万个文件**，构成：

| 子目录 | 体积 |
|---|---|
| web-apps（三编辑器 + 帮助动图等） | 693.8 MB（其中 documenteditor 109.0 MB、公共 25.5 MB、其余为 Excel/PPT 等） |
| sdkjs | 234.5 MB（word 23.1 + common 44.5 + cell/slide/pdf/visio 约 150 MB） |
| fonts | 190.1 MB（可大幅裁剪；公文字体如方正小标宋、楷体_GB2312 需自备且涉及字体授权） |
| x2t + x2t-fonts | 17.6 MB |

裁剪到仅 Word 编辑器（documenteditor + web-apps/common + sdkjs word+common + x2t + 精简字体）≈ **220 MB + 字体**。相比 docx-preview 的数百 KB，安装包体积翻 2–3 倍，分发与更新成本显著上升。

### 3.3 维护健康度

- 活跃：2025-11 创建，约 175 commits，最近 push 2026-07-30（调研前 10 天）；issue 响应快，中文社区氛围。
- 风险：**实质单作者**；版本 0.1.0、未发布 npm、无语义化版本与升级通道；绑死一份 9.4.0-develop 快照资产，跟进 OnlyOffice 上游需自行重新导出验证；测试仅有 playwright e2e。

### 3.4 License 与项目约束核对

paiban-studio 仓库无 LICENSE 文件、package.json 无 license 字段，按闭源/商业意图对待。**AGPL-3.0 组件嵌入闭源 Electron 应用分发构成硬冲突**，出路只有两条：购买 OnlyOffice 商业授权（成本与商务流程未评估），或整个项目转 AGPL。两者都不是「换个预览面板」的决策。

---

## 4. 结论

**不采用**（作为预览/看板面板）。三个独立成立的否决理由，任一都足以否决：

1. **定位错配**：它是有状态全功能编辑器，不是无状态渲染器。当前「agent 改 docx → 快速重渲」链路上，它的每次刷新 = x2t 重转换 + 重开（秒级、丢滚动位置），比现有 docx-preview 的 0.3–0.7s 更差。
2. **License 硬卡点**：仓库与上游均为 AGPL-3.0，闭源分发须购 OnlyOffice 商业授权；且资产导出自商业版 Developer Edition，再分发授权边界需商务确认。
3. **分发成本**：裁剪后仍约 220 MB+，对「本地轻量工作台」定位过重；单作者 0.1.0 项目，无 npm 发布与升级通道。

### 有条件采用的触发条件（未来重新评估的情形）

若产品方向演进为「人在编辑器里直接手改 + AI 协同批注修订」（即**用 OnlyOffice 替换自研 OOXML 编辑层**，而非替换预览层），且商业授权谈妥，则值得重估。届时最小集成路径草案：

1. 复制 `src/components/onlyoffice-web-comp/` 进 `src/`，剥离 React 依赖只留 manager 核心；
2. 从 `9.4.0-develop` 资产裁剪出 Word-only 子集（documenteditor + sdkjs word/common + x2t + 按需字体，约 220 MB），置于自定义 protocol（`app://onlyoffice/`）下；
3. 渲染进程加载 `api.js` 建编辑器 iframe；打开/导出经 x2t Worker；
4. agent 不再走 pizzip 改 XML，改走 **Connector Automation API**（`createConnector()`）下发批注/修订/格式指令，版本链在导出 docx 二进制时挂接现有 S3 抽象；
5. 验收基线：公文 fixtures（红头、页眉页脚、编号、表格）打开保真度 + 编辑-导出往返不丢格式。

### 顺带收获

- 其「字体注册」工具链（`ttf-to-catalog-font.mjs` + `__custom_font_registry__`）与「批注修订 API 封装」（`feature/` 目录）的思路可在未来自研层需要时参考（注意 AGPL 传染性，**只参考思路、不复制代码**）。
- docx-preview 的保真度短板（公文复杂表格/分页）依然存在，本调研未改变 D4 决策；后续如需提升预览保真度，「系统 soffice 转 PDF/图片」兜底路线仍优于引入 OnlyOffice。
