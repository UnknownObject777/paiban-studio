# paiban-studio — Agent 配置

## 项目身份

AI 驱动的 Office 排版本地工作台（Word MVP）：Electron 桌面应用，对话式 AI 直接修改 docx + 实时预览 + 版本链，本地优先。实现边界与全部拍板决策见 `docs/mvp-spec.md`。

## Agent skills

### Issue tracker

Issues are tracked as GitHub issues on this repo (via the `gh` CLI). See `docs/agents/issue-tracker.md`.

### Triage labels

Default five canonical labels: `needs-triage`, `needs-info`, `ready-for-agent`, `ready-for-human`, `wontfix`. See `docs/agents/triage-labels.md`.

### Domain docs

Single-context layout: one `CONTEXT.md` at the repo root + `docs/adr/`. See `docs/agents/domain.md`.

## 卡片实现工作流（强制规则）

当 agent 选择一张卡片（GitHub issue，含 `ready-for-agent` 等标签的 ticket）来实现时，必须遵守以下规则：

- **必须使用 git worktree 隔离开发**：不得在默认分支（`main`）上直接修改代码。先 `git worktree add` 一个隔离工作目录（例如 `.claude/worktrees/<card-id>`，并创建对应分支）再动手。
- **只有测通的才允许合入**：该卡片的实现必须在隔离 worktree 中跑通相关测试且无回归后，才允许合入到 `main`。未通过测试的变更禁止合并。
- **合入方式**：在 worktree 分支上完成实现与验证 → 切回 `main` 执行合并（fast-forward 或 merge）→ 合入后清理该 worktree。
- 上述规则同样适用于多 agent / 并行实现多张卡片的情形：每张卡片各占一个 worktree，互不干扰，各自测通后再合入。

## 前端审美风格参考（Figma Design System）

本项目前端在生成、重构或新增界面时，遵循 open-design 的 **Figma design system** 审美规范。

权威规范原文：`docs/design/figma/DESIGN.md`（已迁入本库，**动手前先读全文**；源自 open-design `plugins/_official/design-systems/figma/DESIGN.md`，2026-08-02 快照）。

### 核心风格速查

- **界面 chrome 严格黑白**：界面层只允许 `#000000` 与 `#ffffff`，彩色仅出现在内容/产品展示（hero 渐变、截图等）
- **Hero 渐变签名**：电光绿、亮黄、深紫、热粉的多色渐变，作为首页视觉亮点
- **按钮几何**：pill（`border-radius: 50px`）与正圆（`50%`）两种，不用直角；实心黑 / 实心白 / 玻璃黑 `rgba(0,0,0,.08)` / 玻璃白 `rgba(255,255,255,.16)`
- **字重层级**：figmaSans 变体字重 320 / 330 / 340 / 450 / 480 / 540 / 700，靠微差而非「常规 vs 加粗」建立层级；正文不超过 450
- **字距**：正文一律负字距（-0.1px 至 -1.72px，display 级 -0.96/-1.72px）；figmaMono 大写技术标签用正字距 0.54~0.6px；所有文本启用 OpenType `"kern"`
- **focus 指示**：一律 `dashed 2px` 虚线 outline（呼应 Figma 编辑器选择手柄），不用实线
- **阴影克制**：深度主要靠明暗背景对比（白卡片压彩色/深色区），几乎不堆阴影
- **圆角标尺**：2px（小链接）/ 6px（小容器）/ 8px（卡片、对话框）/ 50px（pill 按钮、tab）/ 50%（图标按钮）
- **间距体系**：8px 基数，刻度 1 / 2 / 4 / 4.5 / 8 / 10 / 12 / 16 / 18 / 24 / 32 / 40 / 46 / 48 / 50

### 落地规则

- 修改前端样式或新增前端组件时，先读 `DESIGN.md` 对应小节再动手
- **不要**向界面 chrome 添加 `#4f8cff` 这类彩色强调色 —— 黑白二元是硬约束，彩色只属于内容与渐变
- 按钮、tab、输入框、头像等交互元素使用 pill / 圆形几何
- 保持 8px 间距体系与负字距排版；页面标题级别使用大号紧排（86px / 64px / 48px）
- 断点：<560 / 560-768 / 768-960 / 960-1280 / 1280-1440 / 1440-1920px，窄屏收拢为单列

&nbsp;