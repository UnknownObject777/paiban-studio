# paiban-studio

**AI 驱动的 Office 排版本地工作台（Word MVP）**——上传排版混乱的 docx + 选/传模板 → 用自然语言指挥 AI 直接修改文档 → 秒级实时预览 → 每次改动自动存版本、可回滚。全程本地运行，公文不出内网。

## 状态

MVP 规格已确认（2026-08-02），处于实现前的原型阶段。当前 frontier：

- 模板规则集原型（识别/样式两文件 schema + 组件画廊预览页）
- round-trip 保真 spike（真实 Word/WPS 文档 → 解析 → 序列化 → 重新打包 → 双端无损，MVP 第一里程碑）

## 文档

| 文档 | 内容 |
|---|---|
| `docs/mvp-spec.md` | MVP 实现边界报告：Problem/Solution/User Stories/全部拍板决策（D1–D9 + R1–R4）/模块接口/测试策略 |
| `docs/research/` | 三份调研报告：OfficeCLI 借鉴范式、pi agent 接入、OOXML 编辑层与预览选型 |
| `docs/knowledge/wfp-formatting-rules.md` | wfp_core 排版识别规则知识提取（模板规则集的知识来源） |
| `docs/legacy-assets.md` | 旧库资产盘点：可复用代码/借鉴范式/知识资产/废弃清单（逐文件级） |
| `reference/OfficeCLI` | OfficeCLI 本地浅克隆副本（一手源码，agent 实现期参考用；`.gitignore` 排除，不进仓库） |
| `docs/design/figma/DESIGN.md` | Figma design system 完整规范（前端审美唯一权威，open-design 快照） |
| `AGENTS.md` | agent 工作约定：卡片实现工作流（worktree 强制）、Figma design system 审美规范 |

## 技术决策速览

- **形态**：Electron ≥ 35 桌面应用（Node ≥ 22.19），无框架纯 HTML/CSS/JS 三栏前端，IPC 通信
- **编辑层**：pizzip + fast-xml-parser@5（preserveOrder）+ 自研薄文档模型层，唯一 seam `applyEdits(buffer, commands)`
- **agent**：@earendil-works/pi-coding-agent SDK 主进程内嵌，自研三工具（doc_edit / template_read / version_store）
- **预览**：docx-preview@0.4 渲染进程 iframe，端到端 0.3–0.7s
- **模板**：排版设计系统哲学——规则集（识别/样式两文件）+ 组件化 + 组件画廊预览
- **版本**：S3 兼容抽象 + 本地文件实现，sha256 内容寻址 + 版本链
- **分发**：electron-builder，跨 Win/macOS/Kylin 三平台（信创合规）

## 渊源

排版识别规则知识源自 [Word-Formatter-Pro](https://github.com/UnknownObject777/Word-Formatter-Pro) 的 wfp_core（85KB Python 排版引擎），经文档化提取迁入；本产品为全新 Node.js 实现，与旧项目无代码依赖。决策过程档案见旧库 wayfinder 地图 #1。
