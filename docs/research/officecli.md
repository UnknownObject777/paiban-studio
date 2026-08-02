# OfficeCLI 调研报告（wayfinder research ticket #2）

- 调研对象：<https://github.com/iOfficeAI/OfficeCLI>
- 调研日期：2026-08-02
- 调研方法：GitHub API 读取仓库元数据 / README / 源码结构 / 关键源码文件，并浅克隆官方 Wiki（`OfficeCLI.wiki.git`）逐页核对 docx 能力文档。
- 仓库快照：主语言 C#，License Apache-2.0，约 24.1k stars，创建于 2026-03-15，调研当日仍在活跃更新；当前版本 v1.0.143（见 `src/officecli/officecli.csproj`）。

---

## 1. 整体架构

### 1.1 进程形态

OfficeCLI 是一个 **C# (.NET 10) 编写的单文件自包含原生二进制**（`PublishSingleFile` + `SelfContained` + 裁剪），跨平台发布 macOS(arm64/x64)、Linux(x64/arm64)、Windows(x64/arm64) 共 6 个目标。无需安装 Office、无需 .NET 运行时、零外部依赖。安装渠道包括 install.sh/ps1、Homebrew、Scoop、npm（`@officecli/officecli`，安装时按平台拉取原生二进制）。

它有四种进程工作模式：

| 模式 | 形态 | 说明 |
|---|---|---|
| 一次性命令 | 每条命令一个进程 | `create` / `add` / `set` / `get` / `view` 等，进程退出即落盘 |
| 驻留模式（resident） | 长驻进程 + 命名管道 | `officecli open doc.docx` 后文档常驻内存，后续命令通过命名管道（`officecli-<sha256路径前16位>`）转发，延迟接近零；空闲约 10s 自动落盘；`save` 落盘保留进程，`close` 落盘并释放。**外部工具（python-docx、Word）读文件前必须先 save/close** |
| 批量模式（batch） | 一次打开/保存周期跑多条命令 | stdin / `--input` / `--commands` 传入 JSON 命令数组；**默认原子化——任一失败整批回滚**；`--best-effort` 保留已成功部分；`--stop-on-error` 遇错即停 |
| 预览模式（watch） | 本地 HTTP 服务 | `officecli watch doc.docx` 起 `http://localhost:26315`，SSE 自动刷新，Excel 支持单元格内联编辑 |

### 1.2 工具暴露方式：可被外部 agent 调用，**不自带 agent 循环**

OfficeCLI 定位是「工具」，不是「智能体」。agent 循环在仓库之外（官方配套的桌面 agent 应用是另一个项目 [AionUi](https://github.com/iOfficeAI/AionUi)）。它向外部 agent 暴露三种接口：

1. **CLI + JSON**：所有命令支持 `--json`，输出 schema 一致；错误是结构化对象（`code` 为 `not_found` / `invalid_value` / `unsupported_property` / `invalid_path` 等 9 种），并附 `suggestion` 与合法取值范围；属性名拼错会返回编辑距离最近的建议。设计上让 agent「无需正则解析 stdout」并能自愈纠错。
2. **内置 MCP 服务器**：`officecli mcp claude|cursor|vscode|lmstudio` 一键注册。实现是 stdio 上的最小 JSON-RPC 2.0（`initialize` / `tools/list` / `tools/call`），**只对外暴露一个工具**（源码注释明确：误路由的调用必须拒绝，避免以错误工具名改动文件）。
3. **SKILL.md 技能文件**：仓库根部的 `SKILL.md`（约 26KB）教 agent 如何安装二进制和使用全部命令；`officecli install` 会自动检测 Claude Code / Cursor / Windsurf / Copilot 等并写入技能文件。

另附官方薄 SDK：`sdk/node`（`@officecli/sdk`）与 `sdk/python`——两者都只是**命名管道客户端**，把与 `batch` 相同的命令对象 `{command, path, props}` 转发给驻留进程，没有第二套词汇表。

### 1.3 支持的文档类型

| 格式 | 读取 | 修改 | 创建 |
|---|---|---|---|
| Word (.docx) | ✅ | ✅ | ✅ |
| Excel (.xlsx) | ✅ | ✅ | ✅ |
| PowerPoint (.pptx) | ✅ | ✅ | ✅ |

创建仅支持这三种 OOXML 格式（按扩展名判断）。不支持旧版二进制格式（.doc/.xls/.ppt）的创建；`Core/CompoundFile.cs` 仅用于 OLE 对象嵌入场景。

### 1.4 三层能力架构（对 agent 的核心设计）

| 层 | 用途 | 命令 |
|---|---|---|
| L1 读取 | 语义视图 | `view`（text / annotated / outline / stats / issues / html / svg / screenshot） |
| L2 DOM | 结构化元素操作 | `get` / `query` / `set` / `add` / `remove` / `move` / `swap` |
| L3 原始 XML | XPath 兜底 | `raw` / `raw-set` / `add-part` / `validate` |

配套机制：

- **路径寻址**：每个元素有稳定路径，如 `/body/p[1]/r[2]`、`/slide[1]/shape[2]`、`/section[1]`、`/header[1]`。自定义语法（1-based、本地元素名），不是 XPath；agent 无需懂 XML 命名空间。
- **CSS 风格查询**：`query doc "p[style=Heading1]"`、`"p:contains(TODO)"`、`"r[bold=true]"`、`:has()` 等。
- **内置帮助**：`officecli help docx set paragraph` 分层列出全部可设属性，agent 不确定属性名时查帮助而不是猜。
- **dump → batch 往返**：`dump` 把整篇文档或任意子树序列化为可重放的 batch JSON，`batch` 重放——agent 读结构化规格「学习」人类范本，而不是反推原始 OOXML。
- **模板合并**：`merge` 把 `{{key}}` 占位符替换为 JSON 数据（段落、表格单元格、页眉页脚、图表标题均支持），实现「设计一次、填充 N 次」。

---

## 2. docx 编辑能力清单与操作粒度

**粒度结论先行：支持 run 级乃至「字符偏移区间」级的精细修改，远不止整段替换。** `find=`/`replace=` 会把命中文本自动切分成独立 run 再精确套用格式；`range=6:15`（0-based 半开区间，支持多段、可跨段落）直接按字符位置设置格式。

### 2.1 能力矩阵（Word / .docx）

| 元素 | 路径 | 能做什 么（add / set / get / remove / move） | 粒度 |
|---|---|---|---|
| 段落 paragraph | `/body/p[N]` | 增删移换；样式、对齐、首行/悬挂缩进、行距、编号列表（numId/ilvl）、分页控制；`find=`/`replace=` 文本查找替换（支持 regex，自动拆 run）；`range=` 字符偏移区间格式化；按样式/对齐/内容查询（`:contains`、`:empty`、`style!=`） | 段落级 + 段内字符区间级 |
| 文本片段 run | `/body/p[N]/r[M]` | 增删；改文本 `text=`；字体三槽位（`font.latin` / `font.ea` / `font.cs`）+ 主题字体绑定 + `font.hint`；字号、粗斜体（含复杂脚本 `bold.cs`/`italic.cs`/`size.cs`）、颜色（hex/主题色名）、17 种下划线枚举 + 下划线颜色、删除线/双删除线、上下标、高亮、全大写/小型大写、隐藏文字、字符间距/字距（kern）、底纹、基线偏移（position，半磅精度）、BCP-47 语言标签三槽位、RTL；w14 文字特效（描边/渐变填充/阴影/发光/映像）；`link=` 超链接；`formula=` LaTeX 转内联 OMML 公式；内嵌图片换图（`src=`）；**修订（track changes）**：`revision.type=ins/del/format/moveFrom/moveTo` 创建修订，`revision.action=accept/reject` 按作者接受/拒绝 | **run 级**（字符格式全量） |
| 样式 style | `/styles/{StyleId}` | 新增自定义样式、修改既有样式、读取；段落/字符/表格/编号四类 | 样式定义级 |
| 表格 table | `/body/tbl[N]`、`/tr[M]`、`/tc[K]` | 增删表/行/单元格；样式、边框（含逐边 `border.*`）、单元格合并（`vMerge` + `hMerge`）；行级 `cantSplit`；单元格级 `nowrap`/`hideMark`；**虚拟列轴**：OOXML 无列元素，v1.0.97 合成列抽象，支持列的 add/remove/move/`--from` 克隆（遇合并单元格槽位拒绝插入）；`--from` 跨文档克隆整表 | 表/行/列/单元格四级 |
| 图片 picture | 段内 inline 或浮动 anchor | 插入 PNG/JPG/GIF/SVG；尺寸、替代文本；7 种文字环绕（inline/square/tight/through/topAndBottom/behindText/inFrontOfText）；浮动定位（位置 + 8 种水平/9 种垂直参照系）；`behindText`、旋转、四边裁剪、亮度/对比度；`image:no-alt` 查询无障碍缺陷 | 图片元素级，属性精细到 OOXML 参照系 |
| 页眉/页脚 | `/header[N]`、`/footer[N]` | default/first/even 三变体；`type=first` 自动写 `titlePage`，even 需先开 `evenAndOddHeaders`；子段落 `/header[N]/p[K]` 支持完整段落属性集，可向内加段落/图片；remove 会清理节引用与孤儿图片部件 | 部件级 + 内部段落/run 级 |
| 节 section | `/section[N]`（别名 `/body/sectPr[N]`） | 分节符类型（continuous/even/odd/nextPage）；页面宽高与方向（landscape 自动互换宽高）；六向页边距 + 装订线；多栏（栏数、栏距、自定义栏宽、分隔线）；页码格式（含中文/阿拉伯/泰/日/韩等数十种本地化格式）与 `pageStart` 起始页码；`titlePage`；页面边框（逐边样式/粗细/颜色/距边、offsetFrom、zOrder、display）；打印机纸盒；`formProt` 节级表单保护；行号（模式/间隔/距离）；节级脚注/尾注编号格式、重启策略、位置；RTL（`bidi` + `rtlGutter`） | 节属性全量（sectPr 几乎全部子元素） |

### 2.2 其他 docx 能力（均有独立 wiki 页）

编号列表（numbering）、目录（TOC，`WordTocBuilder.cs`）、书签、批注、脚注/尾注、水印、超链接、图表（chart）、公式（LaTeX→OMML）、表单域、内容控件（SDT）、域（22 种零参数域 + MERGEFIELD/REF/PAGEREF/SEQ/STYLEREF/DOCPROPERTY/IF）、OLE 对象、文本框/形状、文档属性与设置（`/settings`、`/docProps`）、换行/分页符、**Markdown 导入**（`WordHandler.Add.Markdown.cs`）、查找替换（`Helpers.FindReplace.cs`，约 62KB）、修订接受/拒绝。

### 2.3 Excel / PowerPoint 概览（非本 ticket 重点，仅列要点）

- Excel：单元格（含音标/振假名）、350+ 函数写入即自动求值（含动态数组 `_xlfn.` 前缀、财务/债券/统计族）、工作表管理、表格、排序、条件格式、图表（箱线/帕累托）、**原生 OOXML 数据透视表**（缓存+定义双写）、切片器、命名范围、数据验证、迷你图、CSV/TSV 导入。
- PowerPoint：幻灯片、形状（图案填充/模糊/超链接跳转）、图片（填充模式/亮度/发光）、表格、图表、动画、morph 过渡、3D 模型（.glb，Three.js 渲染）、幻灯片缩放、主题、连接线、音视频、备注。

---

## 3. License 与代码可复用性

- **License：Apache-2.0**（根目录 `LICENSE` 11KB 标准文本 + `NOTICE` + `THIRD-PARTY-NOTICES.txt`）。
- 法律层面：可以直接依赖其二进制（甚至再分发），也可以复制/修改其源码，只需保留版权声明与 NOTICE。专利授权条款完备，对企业使用友好。
- 工程层面要区分两条路：
  1. **直接依赖二进制**：官方 npm 包 `@officecli/officecli`（安装时按平台拉二进制）+ `@officecli/sdk`（命名管道薄客户端）就是为「被 Node.js 程序依赖」设计的。我们的 Node.js 工作台完全可以把 officecli 当作外部进程工具调用——零重写成本，但引入一个 C# 二进制外部依赖（约几十 MB，需随安装器分发或下载）。
  2. **复用代码**：代码是 C#/.NET，无法被 Node.js 直接 require。若坚持「全量 Node.js 重写」，则源码只有**设计参考价值**——但其属性命名表、wiki 的能力清单、错误码设计、batch/dump JSON 协议都是可以直接照搬的「规格」。
- 注意其 OOXML 底座 DocumentFormat.OpenXml 是微软官方库（MIT），即使复用其 C# 代码也不涉及 copyleft 传染。

---

## 4. OOXML 技术路线与对 Node.js 重写的借鉴

### 4.1 技术路线（一手源码证据）

`src/officecli/officecli.csproj` 全部依赖只有两项：

```xml
<PackageReference Include="DocumentFormat.OpenXml" Version="3.4.1" />
<PackageReference Include="System.CommandLine" Version="3.0.0-preview.2..." />
```

即：**不做 XML 字符串拼接，而是站在微软官方 Open XML SDK（DocumentFormat.OpenXml 3.4.1）的强类型 DOM 之上**，自研一层「路径寻址 + 属性词汇表 + 选择器」的抽象。源码组织印证这一点：

- `Handlers/Word/` 下按动词拆分巨型 partial class（`WordHandler.Add.*` / `WordHandler.Set.*` / `WordHandler.Query.cs` / `WordHandler.Navigation.cs`（420KB）/ `WordHandler.Mutations.cs`），是把 Open XML SDK 对象模型映射到路径/属性词汇的适配层。
- `Core/RawXmlHelper.cs` + `raw` / `raw-set` 命令提供 XPath 级逃生舱——强类型层覆盖不到的长尾场景直接改原始 XML。
- `Core/SchemaOrder.cs`、`WordStrictAttributeSanitizer.cs` 等处理 OOXML 元素顺序约束与严格模式属性清洗——这是操作 OOXML 最真实的痛点。
- 渲染引擎完全自研：`WordHandler.HtmlPreview.*`（合计约 60 万字节 C#）把 docx 渲染为 HTML；截图（`Core/HtmlScreenshot.cs`）复用系统浏览器——优先 Playwright CLI，其次 Chrome/Edge/Chromium 无头，再次 Firefox。Excel 公式引擎、透视表缓存生成、OMML→LaTeX→KaTeX 也是自研。

### 4.2 对「全量 Node.js 重写 AI 驱动 Office 排版工作台」的借鉴意义

**坏消息：Node.js 生态没有 Open XML SDK 的等价物。** DocumentFormat.OpenXml 的价值在于 ECMA-376 全 schema 的强类型类库 + 元素顺序校验。Node 侧候选：

- `docx`（npm）：创建导向，读改能力弱，离 OfficeCLI 的能力面差一个数量级；
- `JSZip` + `fast-xml-parser` / `xmlbuilder2` / `@xmldom/xmldom`：需要自己重建「OOXML 语义层」——本质上是把 OfficeCLI 的 Handler 层用 TS 重写，并以 `docx` 库的类型或自建类型兜底。

**可借鉴的设计（按价值排序）：**

1. **三层渐进架构（L1 语义视图 / L2 DOM 操作 / L3 raw XML 兜底）**：不必追求强类型层 100% 覆盖 OOXML——先保证路径寻址 + 属性词汇表统一，长尾场景一律放行 raw XML + XPath。这把「重写 Open XML SDK」这个不可能任务降级为「重写常用 20% 的适配层」。
2. **路径寻址 + CSS 风格选择器作为唯一寻址词汇**：`/body/p[1]/r[2]` 贯穿 CLI、JSON、MCP、SDK 四个接口，agent 不需要懂命名空间。我们的工具层应一开始就定义稳定的 path/selector grammar，而不是暴露内部对象模型。
3. **命令协议即 JSON 文档**：batch 命令数组 `{command, path, props}` 同时是 CLI stdin 格式、驻留管道协议、MCP 工具入参、Node/Python SDK 的唯一词汇。一套协议四处复用，且天然支持原子化回滚（默认整批失败回滚）——我们的 MCP 工具层可直接采用「命令数组 + 原子执行」形态。
4. **结构化错误 + 自愈闭环**：错误码枚举 + `suggestion` + 合法值列表 + 属性名编辑距离纠错，是「agent 无人工介入自纠错」的关键，比功能数量更重要。
5. **操作粒度对齐 OOXML 真相而非用户直觉**：run 级修改 + `find=` 自动拆 run + `range=` 字符区间，三种粒度并存；表格列是「虚拟轴」（OOXML 本无列元素）。我们的 docx 层也应按 run/段落/字符区间三级设计，并诚实暴露合并单元格等 OOXML 约束。
6. **「渲染 → 看 → 改」闭环内置**：排版工作台的核心差异点。OfficeCLI 用 60 万字节 C# 自研 HTML 渲染器 + 系统无头浏览器截图；Node.js 重写时这段可以用更轻的路线（如直接复用浏览器渲染 HTML 预览，或 docx-preview 类库 + Playwright），但「agent 能看见自己的产出」这一闭环必须保留。
7. **dump → batch 往返 + 模板 merge**：「读人类范本 → 结构化规格 → 批量重放」和「设计一次填充 N 次」两个原语，直接对应我们「AI 排版工作台」的模板工作流，建议照搬协议格式。
8. **驻留模式 + 延迟落盘语义**：长驻进程避免每命令一次开包/写包（docx 是 zip，整篇重写昂贵），但要明确「外部读取前必须 flush」的契约（OfficeCLI 用 10s 空闲自动落盘 + 显式 save/close 解决）。Node 重写时建议同样采用长驻服务形态。

**可直接复用而非重写的部分**：若工期优先，最短路径是 Node.js 工作台直接依赖 `@officecli/officecli` 二进制做 OOXML 读写（Apache-2.0 允许），自己只做 agent 编排与排版决策层；把「全量 Node.js 重写」降级为「工具层薄封装 + 自研排版引擎」。

---

## 5. 来源清单（均访问于 2026-08-02）

1. 仓库元数据：<https://github.com/iOfficeAI/OfficeCLI>（gh api repos/iOfficeAI/OfficeCLI）
2. 中文 README：<https://github.com/iOfficeAI/OfficeCLI/blob/main/README_zh.md>
3. 工程文件（依赖、版本、打包方式）：<https://github.com/iOfficeAI/OfficeCLI/blob/main/src/officecli/officecli.csproj>
4. 源码目录结构：<https://github.com/iOfficeAI/OfficeCLI/tree/main/src/officecli>（Core/、Handlers/Word/）
5. MCP 服务器实现：<https://github.com/iOfficeAI/OfficeCLI/blob/main/src/officecli/McpServer.cs>
6. 截图/渲染后端：<https://github.com/iOfficeAI/OfficeCLI/blob/main/src/officecli/Core/HtmlScreenshot.cs>
7. Node SDK（驻留管道协议）：<https://github.com/iOfficeAI/OfficeCLI/blob/main/sdk/node/index.js>
8. License：<https://github.com/iOfficeAI/OfficeCLI/blob/main/LICENSE>（Apache-2.0）、`NOTICE`
9. 官方 Wiki（浅克隆 `OfficeCLI.wiki.git` 逐页核对）：
   - 总览/命令：<https://github.com/iOfficeAI/OfficeCLI/wiki/command-set-word>、<https://github.com/iOfficeAI/OfficeCLI/wiki/command-batch>、<https://github.com/iOfficeAI/OfficeCLI/wiki/command-open>
   - 段落：<https://github.com/iOfficeAI/OfficeCLI/wiki/word-paragraph> / run：<https://github.com/iOfficeAI/OfficeCLI/wiki/word-run>、<https://github.com/iOfficeAI/OfficeCLI/wiki/word-run-set>
   - 表格：<https://github.com/iOfficeAI/OfficeCLI/wiki/word-table>
   - 图片：<https://github.com/iOfficeAI/OfficeCLI/wiki/word-picture>
   - 页眉页脚：<https://github.com/iOfficeAI/OfficeCLI/wiki/word-header-footer>
   - 节：<https://github.com/iOfficeAI/OfficeCLI/wiki/word-section>、<https://github.com/iOfficeAI/OfficeCLI/wiki/word-section-set>
   - 样式：<https://github.com/iOfficeAI/OfficeCLI/wiki/word-style>
   - i18n/RTL：<https://github.com/iOfficeAI/OfficeCLI/wiki/i18n>
10. 配套 agent 应用（证明 agent 循环在仓库外）：<https://github.com/iOfficeAI/AionUi>
