# MVP 实现边界报告 — AI 驱动的 Office 排版本地工作台（Word MVP）

> 状态：**已确认**（R1/R2/R3 已于 2026-08-02 拍板：#6 规则迁移四路分流、#7 运行时骨架、#8 新建独立仓库 paiban-studio；#5 模板规则集方向已拍板，字段级 schema 随原型迭代）
> 日期：2026-08-02
> 来源：wayfinder 调研 #2/#3/#4（docs/research/officecli.md、pi-agent.md、ooxml-editing-preview.md）+ 地图 #1 已锁定的五项地基决策 + workbench/ 原型现状盘点
> 决策档案：UnknownObject777/Word-Formatter-Pro#1（wayfinder 地图，只读归档）；执行票：本仓库 issues

---

## Problem Statement

用户（政企办公人员）需要把排版混乱的 Word 公文快速排成规范格式。现有 v2.7.6 是 Tkinter 桌面**批量排版**工具，一键规范化很强，但不能"对话式精调"——想改某一段字体、行距、页边距，只能整篇重排或手工打开 Word 调整；也看不到"AI 正在改什么"，改坏了没有回滚保障。

用户要的是：上传一篇乱排版的 docx + 选/传模板 → 用自然语言指挥 AI **直接改文档**（"标题改黑体三号居中""正文统一仿宋四号、行距 28 磅"）→ **实时预览**每步修改 → 满意后导出，且**每次改动自动存版本、可随时回滚**。全程本地运行，公文不出内网。

## Solution

一个 Electron 桌面应用（**MVP 只做 Word .docx**）：

- **对话式 AI 编辑**：主进程内嵌 pi agent，通过自研 OOXML 编辑层直接修改既有 docx（run/段落/节/编号/文本替换级粒度）。
- **秒级实时预览**：docx-preview 在渲染进程 iframe 中渲染，编辑 → 刷新端到端预估 0.3–0.7s。
- **模板资产库（排版设计系统）**：上传模板 → 解析为**结构化规则集** + `{{占位符}}` → 一键实例化为新文档；模板库与工作文档分栏管理。
- **版本兜底**：对象存储（sha256 内容寻址）+ 版本链，每次编辑自动快照，可预览/回滚/下载任意历史版本。
- **本地优先与安全**：全本地处理，原稿零改动（只操作安全副本）；模型支持 Anthropic / OpenAI 兼容端点（含 Ollama、本地网关），可完全离线。

## User Stories

1. 作为政企办公人员，我想把一篇排版混乱的 docx 上传到工作台，以便在不离开本机的前提下开始规范化。
2. 作为办公人员，我想把单位已有的公文范本上传为模板，以便把"老范本"变成可复用资产。
3. 作为办公人员，我想看到模板被解析出的结构大纲与占位符清单，以便知道模板里有哪些可填充项。
4. 作为办公人员，我想用自然语言对当前文档说"标题改成黑体三号居中"，以便让 AI 直接完成排版修改。
5. 作为办公人员，我想对文档说"正文统一仿宋_GB2312、四号、行距 28 磅"，以便全文正文快速统一。
6. 作为办公人员，我想让 AI 对某一段（而不是全文）应用格式修改，以便局部精调而不破坏其余内容。
7. 作为办公人员，我想让 AI 把"一、"（一）"1."等层级标题分别套用对应格式，以便多级标题一次排齐。
8. 作为办公人员，我想修改页边距/纸张大小/页码格式，以便满足公文页面规范。
9. 作为办公人员，我想在对话里看到 AI 每次工具调用的摘要与产生的新版本号，以便知道"它改了什么、改到哪个版本"。
10. 作为办公人员，我想在右栏实时看到文档排版预览，以便在对话修改后立即校对效果。
11. 作为办公人员，我想预览历史版本、一键回滚、下载任意版本，以便改错时随时挽回。
12. 作为办公人员，我想从模板实例化一个新文档并预填占位符内容，以便从标准版式快速起草。
13. 作为办公人员，我想让 AI 在改完后再自查一遍（重解析/渲染冒烟），以便避免产出打不开的文档。
14. 作为办公人员，我想在无外网环境下接入本地模型完成全部操作，以便涉密公文不出内网。
15. 作为办公人员，我想下载当前工作文档的最终 .docx，以便交付或继续在 Word/WPS 里微调。
16. 作为办公人员，我想在上传或操作前看到安全提示（只处理副本、原稿不动），以便放心使用。
17. 作为办公人员，我想被提示当前使用的模型与连接状态，以便确认工作台是否就绪。
18. 作为办公人员，我想在对话中引用模板规则（"按《通知》模板的样式排"），以便套用模板版式而不用逐条描述。
19. 作为开发维护者，我想让所有文档编辑都收敛到一个命令接口，以便测试与替换实现。
20. 作为开发维护者，我想让编辑产出能在 Word 与 WPS 双端打开且版式一致，以便覆盖国产公文环境。

## Implementation Decisions

### 已锁定决策（来自 wayfinder 地图 #1 与调研 #2/#3/#4，直接采信）


| #   | 决策                                                                                                                                          | 依据       |
| --- | ------------------------------------------------------------------------------------------------------------------------------------------- | -------- |
| D1  | 全量 Node.js 重写；Python wfp_core **仅作为规则迁移的知识资产，不作运行时依赖**                                                                                      | 地图地基决策 1 |
| D2  | MVP 只做 Word (.docx)；PPT/Excel 进雾区                                                                                                           | 地图地基决策 2 |
| D3  | 编辑层 = **pizzip + fast-xml-parser@5 (preserveOrder) + 自研薄文档模型层**                                                                             | 调研 #4    |
| D4  | 预览 = **docx-preview@0.4**（渲染进程 iframe `renderAsync`）；mammoth 仅作 AI 语义抽取（MVP 后置）；系统 soffice 仅作可选 PDF 导出兜底（MVP 后置） | 调研 #4    |
| D5  | agent = **@earendil-works/pi-coding-agent** SDK 主进程内嵌（`createAgentSession()`，事件流经 IPC 转发渲染层），预计**无需 fork**；Node ≥ 22.19 / Electron ≥ 35     | 调研 #3    |
| D6  | 借鉴 OfficeCLI 的 agent 友好范式：路径寻址（`/body/p[1]/r[2]`）、batch 命令数组原子化、结构化错误 + 自愈建议、`dump → batch` 往返                                              | 调研 #2    |
| D7  | 模板系统按**排版设计系统**哲学：结构化规则集 + 组件化 + 专用预览视角                                                                                                     | 地图地基决策 3 |
| D8  | 版本管理：**S3 兼容抽象接口 + 本地文件实现**（MinIO/云 OSS 为后续实现）                                                                                              | 地图地基决策 4 |
| D9  | agent 自主度：**直接改 + 版本兜底**（每轮修改自动存版，可对话回滚）                                                                                                    | 地图地基决策 5 |


### 拍板决策（R1/R2/R3 已于 2026-08-02 确认）

**R1 · 运行时骨架（回答 #7）**

- **前端框架：无框架的纯 HTML/CSS/JS**（保留 workbench/public 三栏布局），遵循 `AGENTS.md` 规定的 Figma design system（黑白 chrome、pill/圆角、8px 间距体系）。理由：MVP 交互是"对话面板 + 预览 iframe + 列表"，无复杂状态树；少一层构建链，桌面分发更轻。
- **进程结构：主进程跑 agent 编排 + docx 编辑内核 + 存储层；渲染进程跑 UI + iframe 预览；通信用 IPC**（不用本地 HTTP，避免端口与安全模型负担）。
- **agent 关系：主进程内嵌 SDK**（D5），编辑工具串行执行（复用 workbench `toolExecution: "sequential"` 防并发写同一文件）。
- **构建分发：electron-builder**（跨 Win/macOS/Kylin 三平台；信创合规是产品的既定卖点）。

**R2 · wfp_core 规则迁移策略（回答 #6）**

- **→ 模板规则集（声明式，主力去向）**：主副标题识别、四级标题识别与格式、图表标题定位、附件格式化、表格智能对齐——这些是"版式规范"的稳定规则，做成规则集数据而非硬编码。
- **→ 文档预处理器**：Markdown 清理、TXT 空行整理、符号标准化。
- **→ agent 分析工具（可选，MVP 后置）**：标题层级识别/结构概览作为工具暴露给 agent。
- **→ 直接丢弃**：python-docx 专属实现细节、Tkinter GUI 逻辑、COM/LibreOffice 转换路径（MVP 不做旧格式）。

**R3 · 仓库身份（回答 #8，已拍板）**：**新建独立仓库 paiban-studio，从头建库**——新产品与旧 Word-Formatter-Pro 几乎无关，不沿用 fork。旧库转为 wfp_core 知识资产与决策档案的只读存放地；研究报告/spec/设计规范随迁本库；wfp_core 以文档化提取形式进入（不带 git history、不用 submodule），识别规则知识见 `docs/knowledge/wfp-formatting-rules.md`。

**R4 · LLM provider 策略**：MVP 默认支持 Anthropic + OpenAI 兼容端点（覆盖 DeepSeek/Kimi/Qwen/本地 Ollama/vLLM），复用 workbench 的配置优先级（界面配置 > 环境变量 > 默认值）；凭证存本地配置文件。**默认 `deepseek / deepseek-v4-flash`**（走 pi 内置 DeepSeek provider，`DEEPSEEK_API_KEY` 认证）。

### 模块与接口

**核心模块（新建，Node/TS + Electron）**

1. **docx 编辑内核**（`docx-core`）
  - 输入 seam：`applyEdits(docxBuffer, commands)` → `{ buffer, result }`
  - `commands` 协议（借鉴 OfficeCLI batch）：
  `{ command: "set"|"add"|"remove"|"move"|"findReplace"|"normalize", path|parent|selector, props }`
  - 内部：pizzip 解压 → fast-xml-parser@5 `preserveOrder` 解析 `document/styles/numbering/header/footer` → 内存文档模型 → 应用原语 → 序列化 → 重新打包 → **生成后自检**（重解析 + 渲染冒烟）。
  - MVP 编辑原语子集（按公文场景命中率圈定）：
    - 段落：对齐、首行缩进、行距（半磅/倍数）、段前段后、分页控制（`pageBreakBefore`）、大纲级别
    - run：中文字体（`w:eastAsia`）、西文字体、字号（`w:sz` 半磅）、粗/斜/下划线、颜色
    - 节：页边距（`pgMar`）、纸张大小/方向、页码格式（页脚字段）
    - 文本：`findReplace`（自动拆分 run，参考 OfficeCLI `find=` 语义）
    - 编号：numbering.xml 多级编号单独封装模块 + 重点测试（地图 #1 标记首要风险之一）
    - 全文规范化：模板规则集驱动的全文档重排原语（`normalize`，替代 wfp_format 的 Python 引擎）
  - 辅助生成：`docx` (npm) 仅用于构造新内容 OOXML 片段后 merge（D3 辅助位）。
2. **pi agent 接入层**（`agent-core`，重构 workbench/server/agent.js）
  - 自研扩展（`registerTool`）：`doc_edit`（走编辑内核 seam）、`template_read`（模板规则集/占位符读取）、`version_store`（快照/回滚/列表）
  - 写操作工具统一包 `withFileMutationQueue`（与 pi 内置工具共享逐文件串行队列）
  - 权限门：`tool_call` 事件拦截——可写目录白名单 + 危险操作确认
  - 系统提示词：`before_agent_start` 注入排版规范 + 当前文档结构摘要
  - 禁用/裁剪 pi 内置文件工具集（避免绕过编辑内核直接改文件）
3. **存储与版本层**（`storage`，**复用 workbench/server/storage.js**）
  - 内容寻址对象存储（sha256）+ 版本链；`S3` 兼容接口抽象，当前实现为本地文件系统
  - 编辑类工具成功后自动快照；内容无变化不产生空版本（幂等）；回滚也记录为新版本
4. **模板层**（`templates`，重构 workbench/server/templates.js）
  - 上传 → 解析：结构大纲 + `{{占位符}}` 提取 + **规则集抽取**（从样例 docx 反推排版组件样式，MVP 先做"标题/正文/页边距"最常用组件的抽取，完整组件划分随 #5 原型迭代）
  - 实例化：复制模板 →（可选）占位符合并 → 生成新工作文档
  - 规则集 schema（#5 已拍板方向）：**识别与样式两文件分离**——`recognizers.json`（component id → match 规则）+ `styles.json`（component id → style + page），组件 id 绑定；对话微调只动 styles
  - 组件全集（MVP 最小集）：title / subtitle / heading1..4 / body / caption / table / attachment + page
  - 来源混合起步：内置公文默认规则集（翻自 `docs/knowledge/wfp-formatting-rules.md`）+ agent 从样例 docx 反推（dump→batch 范式）→ 预览确认入库
  - 模板预览视角：组件画廊 + 页面线框（原型期静态页造假数据）
5. **预览层**（`preview`，替换 workbench 的 `office.renderHtml`）
  - 渲染进程 iframe 内 `docx-preview.renderAsync(arrayBuffer, container)`
  - 编辑 → IPC 传 ArrayBuffer → 防抖刷新（连续编辑合并）
6. **前端**（`public`，**复用 workbench/public 三栏布局**并切到 Figma design system）
  - 左栏：模板库 / 工作文档；中栏：对话流（token 流式 + 工具调用摘要 + 版本通知）；右栏：预览 iframe + 版本时间线

**明确不新建**：不引入 LibreOffice 子进程、不引入 mammoth 渲染、不引入 React/Vue/Svelte、不保留 OfficeCLI 二进制依赖（原型里 office.js 废弃）。

## Testing Decisions

**原则：只测外部行为，不测实现细节。** 所有编辑语义都从 seam 的 `buffer → buffer` 外部行为断言；不 inspect 内部文档模型结构。

**唯一 seam：`applyEdits(docxBuffer, commands)**`（docx 编辑内核命令接口）。理由：

- 它是全系统最高汇聚点——agent 工具、模板实例化、手动操作最终都收敛到这一条命令。
- 协议稳定，实现可换（pizzip/fxp/文档模型都是内部细节）。
- 一个 seam 同时覆盖两大风险：编辑原语正确性 + round-trip 保真。

**测试模块与先例：**


| 模块        | 测什么                                                                                       | 先例                                                           |
| --------- | ----------------------------------------------------------------------------------------- | ------------------------------------------------------------ |
| docx 编辑内核 | 每个原语的 buffer→buffer 断言；round-trip 保真（解析→序列化→重解析逐字节稳定）；坏输入/不支持命令报错；**真实 Word/WPS 产出文档回归集** | 新增；回归集参考 workbench/data 现有样例与 wfp_core 自带测试                  |
| 存储版本链     | 快照幂等（无变化不产生版本）、回滚语义、内容寻址去重                                                                | workbench/test/smoke.js 第 5 节（版本管理）                          |
| 模板层       | 占位符提取、规则集抽取、实例化产物为合法 docx（走编辑内核自检）                                                        | workbench/test/smoke.js 第 4 节（上传/解析/实例化）                     |
| 端到端冒烟     | 起应用 → 上传文档 → 对话编辑 → 预览 → 版本 → 下载                                                          | workbench/test/smoke.js（REST 冒烟，改造为 IPC/UI 级或保留本地 HTTP 测试模式） |
| 真实 LLM 链路 | agent 用工具完成"改排版"任务并自查                                                                     | workbench/test/e2e-agent.js                                  |


**回归集要求**：内置一组真实 Word 与 WPS 产出的 .docx 样本，每次 round-trip 后重解析校验 + docx-preview 渲染冒烟（不抛错、DOM 非空）；发布前在 Word 与 WPS 双端人工打开验证（地图 #1 首要风险）。

## Out of Scope（MVP 明确不做）

- Excel (.xlsx) / PowerPoint (.pptx) 编辑
- Python wfp_core 作为运行时依赖（仅知识资产）
- LibreOffice 作为编辑层或主预览（仅可选 PDF 导出兜底，后置）
- mammoth 语义抽取给 AI 用（后置）
- 批量处理（继承 WFP 批量能力，MVP 之后）
- 多人协同 / 服务端多租户（本地单用户）
- 移动端 / Web 在线版
- 旧格式 .doc/.wps 转换与修订接受（依赖 COM/LibreOffice，MVP 不做）
- 模板分享 / 导入导出格式
- 会话持久化到磁盘（MVP 内存态即可；文档与版本链不受影响）
- pi 进程级沙箱 / 容器化（仅目录白名单权限门）
- 自定义上下文压缩策略、人机抢话队列（steering/followUp）
- OfficeCLI 二进制依赖（原型已否决，仅借鉴协议范式）
- PDF 导出（soffice 可选路径后置）

## Further Notes

- **首要风险（spike 前置）**：fast-xml-parser `preserveOrder` round-trip 保真。MVP 第一个里程碑必须是"真实 Word/WPS 文档 → 解压 → 解析 → 序列化 → 重新打包 → 双端打开无损"的 spike，通过后才开始编辑原语。
- **workbench 处置**：定位为 wayfinder 期间的可执行原型。其**已验证的资产**（storage.js 版本链、三栏前端、SSE 事件桥接、smoke/e2e 测试骨架、路径寻址/batch 协议范式）直接复用；其**OfficeCLI 编辑引擎**（office.js）与 **pi 旧包名**（`@mariozechner/*`，组织已迁至 `@earendil-works/*`）废弃替换。
- **决策状态**：#6（R2）、#7（R1）、#8（R3）已拍板关闭；#5 模板规则集方向已拍板（两文件分离/混合起步/组件画廊），字段级 schema 随原型迭代。
- **领域词汇**：本文使用"模板规则集 / 排版组件 / 编辑原语 / 版本链 / 对象存储 / 占位符"等词；若与 `CONTEXT.md` 未来定义冲突，以领域文档为准。

&nbsp;