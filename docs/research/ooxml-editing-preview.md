# OOXML 编辑层与 docx 实时预览渲染选型调研

- 调研日期：2026-08-02
- 对应 ticket：wayfinder research #4（UnknownObject777/Word-Formatter-Pro）
- 目标：Node.js 全栈重写的「AI 驱动 Office 排版本地工作台」（Electron 桌面应用，MVP 只做 docx），确定 (1) 对**既有** docx 文档的编辑层技术路线，(2) 秒级刷新的实时预览渲染方案。

---

## 1. docx 编辑层选型

### 1.1 一手数据（2026-08-02 查询）

| 库 | 最新版本 / 发布日期 | 周下载量 | GitHub 活跃度 | License | 解压体积 |
|---|---|---|---|---|---|
| [docx](https://www.npmjs.com/package/docx)（dolanmiu/docx） | 9.7.1 / 2026-05-27 | ~596 万 | 最近 push 2026-06-12，open issues 160，5.9k stars | MIT | ~4.7 MB |
| [docxtemplater](https://www.npmjs.com/package/docxtemplater)（open-xml-templating/docxtemplater） | 3.69.3 / 2026-07 | ~83 万 | 最近 push 2026-07-23，open issues 5，3.6k stars | MIT（核心） | ~1.3 MB |
| [pizzip](https://www.npmjs.com/package/pizzip) | 3.2.0 | ~97 万 | 最近 push 2026-07-24，open issues 0 | MIT 或 GPL-3.0 | ~0.6 MB |
| [jszip](https://www.npmjs.com/package/jszip)（Stuk/jszip） | 3.10.1 | ~3974 万 | 最近 push 2025-03-28（维护放缓），open issues 412 | MIT 或 GPL-3.0 | ~0.8 MB |
| [fast-xml-parser](https://www.npmjs.com/package/fast-xml-parser)（NaturalIntelligence/fast-xml-parser） | 5.10.1 / 2026-07-16 | ~8386 万 | 最近 push 2026-08-01，open issues 20，3.1k stars | MIT | ~1.3 MB |
| [mammoth](https://www.npmjs.com/package/mammoth)（mwilliamson/mammoth.js） | 1.12.0 / 2026-03-12 | ~675 万 | 最近 push 2026-05-24，open issues 65 | BSD-2-Clause | ~2.2 MB |

来源：npm registry API（registry.npmjs.org / api.npmjs.org）、GitHub REST API（gh api repos/...），查询日期 2026-08-02。

### 1.2 路线 (a)：纯 JS 库

**docx (npm, dolanmiu/docx)**
- 本质是**生成器**：声明式 API 从零构造 Document。README 虽写 "generate and modify"，但"modify" 仅指 v8+ 引入的 `patchDocument`——把占位符（placeholders）替换成内容块，**不能**把既有 docx 解析成 Document 对象做自由编辑。
- 维护者在 Discussion [#1265](https://github.com/dolanmiu/docx/discussions/1265)、[#1912](https://github.com/dolanmiu/docx/discussions/1912) 明确回复："no way to read an existing docx file"，且至今无通用 round-trip 读入能力。
- 结论：**不适合做编辑层核心**。可用于「生成新文档/新内容块」的辅助场景（例如 AI 输出整段新内容时用它构造 OOXML 片段）。

**docxtemplater + pizzip**
- 模板填充型：在模板里写 `{tag}`，运行时注入数据。对「用户上传一篇既有公文、AI 对话式调整排版」的场景模型不匹配——它假设你**预先制作模板**。
- run 级编辑（改某段的字体/字号/行距/首行缩进）不是它的抽象；底层要靠付费模块或直接操作 XML。
- 维护极好（open issues 仅 5 个），但用途不符。结论：**不采用为编辑层**，模板批量生成子功能可后续引入。

### 1.3 路线 (b)：直接操作 OOXML（jszip/pizzip + fast-xml-parser）

- docx 即 zip 包：`word/document.xml`（正文）、`word/styles.xml`（样式）、`word/numbering.xml`（编号）、`word/headerN.xml`/`footerN.xml`（页眉页脚）、`word/media/*`（图片）、`[Content_Types].xml`、`word/_rels/document.xml.rels`（关系）。
- 自由度：run（`w:r` / `w:rPr`）级属性——字体（`w:rFonts`，含 `w:eastAsia` 中文字体）、字号（`w:sz`，半磅）、行距（`w:spacing w:line`）、页边距（`w:sectPr/w:pgMar`）、分页（`w:br w:type="page"` / `w:pageBreakBefore`）、编号（numbering.xml）、表格（`w:tbl`）、图片（drawing + rels）、页眉页脚——全部是 XML 属性/节点，**100% 可达**，不受任何库抽象缺失的限制。这正是中文公文排版（仿宋_GB2312、三号字、28 磅行距、首行缩进 2 字符）所需的控制粒度。
- 工作量：需要自建一层薄的「文档模型 + 编辑原语」（如 setRunFont / setParagraphSpacing / setPageMargins / replaceText），预计是本项目最核心的自研代码，但 OOXML WordprocessingML 的子集（段落/run/表格/节/页眉页脚/编号）可控。
- 依赖：`fast-xml-parser`（5.x，极活跃，周下载 8000 万+，preserveOrder 模式可无损 round-trip）+ zip 层选 **pizzip**（同步 API、活跃维护、体积小）或 jszip（生态更大但维护放缓、412 个 open issues；体积都 <1MB）。
- TypeScript：两者均有类型支持，fast-xml-parser 原生 TS。
- 包体积：合计 <3 MB，对 Electron 分发无压力。
- 风险：XML round-trip 保真（空白、属性顺序）需用 `preserveOrder: true` 并配合单元测试锁定；无 schema 校验，错误 XML 会产生 Word 打不开的文档——需要「生成后自检」（重新解析 + 可选 docx-preview 渲染冒烟）。

### 1.4 路线 (c)：桥接 LibreOffice headless / unoconv

- **unoconv 已事实弃维护**，社区共识是直接调 `soffice --headless --convert-to ...`，Node 侧用 `child_process` 或 `libreoffice-convert` 包装。
- 优点：编辑通过 UNO API 可做，且转换保真度高；处理复杂对象（图表、SmartArt、多级编号）强于任何纯 JS 方案。
- 缺点对桌面分发致命：LibreOffice 本体 300–500 MB，打包进 Electron 或要求用户预装都重；UNO 桥接（socket + urp）工程复杂度高；冷启动 1–2 s 不满足对话式秒级刷新。
- 结论：**不作为 MVP 编辑层**。保留为可选的「高保真导出 PDF」插件路径（检测用户机器已装 soffice 时启用）。

### 1.5 编辑层对比矩阵

| 维度 | docx (npm) | docxtemplater+pizzip | jszip/pizzip+fast-xml-parser | LibreOffice headless |
|---|---|---|---|---|
| 编辑既有文档 | 仅 patchDocument 占位符替换 | 仅模板占位符 | **完全自由** | 完全自由（UNO） |
| run 级编辑（字体/字号/颜色） | 仅生成时 | 弱 | **完全** | 完全 |
| 样式/编号 | 生成时定义 | 依赖模板 | 直接改 styles.xml/numbering.xml | 支持 |
| 表格/图片/页眉页脚 | 生成支持 | 部分需付费模块 | 直接 XML 操作 | 支持 |
| 维护活跃度 | 高 | 高 | 高（fxp 极高；jszip 放缓，pizzip 活跃） | 项目活跃，unoconv 弃维护 |
| 包体积 | ~4.7MB | ~2MB | **<3MB** | **300–500MB** |
| TS 支持 | 好 | 好 | 好 | 无（进程桥接） |
| 工程风险 | 能力不匹配 | 模型不匹配 | 自研 XML 层工作量 | 分发/集成重 |

**编辑层结论：路线 (b) 直接操作 OOXML（pizzip + fast-xml-parser）为核心，docx (npm) 作为生成新内容片段的辅助库。**

---

## 2. docx 实时预览渲染选型

### 2.1 候选对比（数据来源同 §1.1，2026-08-02）

| 维度 | docx-preview（docxjs） | mammoth | LibreOffice 转 PDF/HTML |
|---|---|---|---|
| 定位 | 视觉还原渲染 | 语义化转 HTML | 服务端/进程转换 |
| 最新版 | 0.4.0 / 2026-07-07（活跃，push 2026-07-07，2k stars，open issues 63） | 1.12.0 / 2026-03-12 | — |
| 周下载量 | ~123 万 | ~675 万 | — |
| 保真度 | **高**：分页、页眉页脚、节、复杂样式均渲染，社区评价「像素级还原」 | 低：刻意丢弃排版，仅语义结构；多级编号丢失、合并单元格支持有限、无分页 | 最高（字体齐全时），但中文公文字体缺失会导致版式漂移 |
| 渲染速度 | 中（大文档偏慢，典型公文 <1s） | 快 | 慢（冷启动 1–2s + 转换） |
| 分页呈现 | **支持（按节分页渲染）** | 不支持 | PDF 天然分页 |
| Electron 集成 | **渲染进程直接 renderAsync(arrayBuffer, container)，iframe/webview 均可** | 同左（也可 Node 侧） | 需外置进程 + PDF viewer |
| 中文排版（字体/字号/行距/页边距/分页） | 能反映 OOXML 中的实际设置，满足校对级预览 | 基本丢失 | 受系统字体影响 |

（综合来源：[docx-preview npm 页](https://www.npmjs.com/package/docx-preview)、[npm-compare 对比](https://npm-compare.com/docx-preview,docxtemplater,jszip,mammoth,officegen)、[掘金：docx-preview 和 mammoth](https://juejin.cn/post/7596768867231072275)、[Apryse 对比文](https://apryse.com/blog/build-react-file-viewer-pdfs-images-office-docs)，查询日期 2026-08-02。）

### 2.2 预览层结论

- **主方案：docx-preview**。保真度与分页呈现是公文场景的硬需求，它是纯 JS 方案中唯一满足者；周下载 120 万+、2026-07 仍在发版，维护健康；Apache-2.0；解压 <1MB；与编辑层共用 zip 依赖。
- **辅助：mammoth 可选**。用于抽取语义结构（标题层级、段落文本）供 AI 上下文使用，而非给用户看。
- **兜底：LibreOffice→PDF 作为可选高保真导出**（同 §1.4，非 MVP 必需）。
- docx-preview 已知短板：SmartArt/图表/多栏复杂版式渲染不完美——公文场景命中率低，接受。

---

## 3. 推荐组合（MVP）

```
编辑层：  pizzip（解/压 zip） + fast-xml-parser@5（preserveOrder 解析/序列化）
          + 自研薄文档模型层（段落/run/表格/节/页眉页脚/编号 的编辑原语）
辅助生成：docx (npm) —— AI 产出整段新内容时构造 OOXML 片段，再 merge 进 document.xml
预览层：  docx-preview@0.4 —— Electron 渲染进程 iframe 内 renderAsync
语义抽取：mammoth —— 为 AI 提供文档结构上下文（不用于展示）
可选兜底：检测到系统已装 soffice 时，提供「高保真导出 PDF」
```

新增依赖总体积 <8 MB，全部 MIT/Apache/BSD 许可，纯 JS 无原生编译，Electron 打包零障碍。

## 4. 「编辑 → 预览刷新」数据通路预估

```
docx 文件 (Buffer)
  └─ 主进程: pizzip 解压 → fast-xml-parser 解析 document/styles/numbering/header/footer
       → 内存文档模型 (JSON 树)
AI/用户编辑指令
  └─ 文档模型应用编辑原语 → fast-xml-parser 序列化 → pizzip 重新打包
       → ArrayBuffer (IPC, ~几百 KB)
           └─ 渲染进程 iframe: docx-preview.renderAsync(buffer, container)
```

延迟估算（典型中文公文，10–30 页，document.xml 200KB–1MB）：

| 环节 | 预估耗时 |
|---|---|
| XML 序列化 + 重新打包 zip | 20–80 ms |
| IPC 传输（结构化克隆 ArrayBuffer） | <10 ms |
| docx-preview 解析 + 渲染 | 200–600 ms（文档越复杂越慢） |
| **端到端刷新** | **约 0.3–0.7 s**，满足「秒级刷新」目标 |

优化路径（如大文档超标）：防抖合并连续编辑、仅重渲染可视页（docx-preview 渲染结果按节分页，可延迟挂载）、Web Worker 内做 XML 序列化。

## 5. 风险与后续验证

1. **XML round-trip 保真**：fast-xml-parser `preserveOrder` 模式 + 对 Word/WPS 真实产出文档的回归测试集（首要 spike）。
2. **docx-preview 中文公文字体验证**：需实测仿宋/楷体/黑体在系统中的回退行为；必要时在渲染容器注入 CSS `@font-face` 兜底。
3. **编号（numbering.xml）编辑复杂度**：公文多级标题编号是 OOXML 中最绕的部分，建议在文档模型层单独封装并重点测试。
4. **Word/WPS 兼容性**：编辑产出需在 Word 与 WPS 双端打开验证（国产公文环境 WPS 占比高）。

---

*数据来源：npm registry API、npm downloads API、GitHub REST API、GitHub Discussions（dolanmiu/docx #1265/#1912/#1933）、npm-compare、掘金社区评测，查询日期均为 2026-08-02。*
