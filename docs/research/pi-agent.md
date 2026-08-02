# 调研报告：pi agent 本地接入、工具扩展与自定义改造

- **查询日期**：2026-08-02
- **关联 ticket**：wayfinder research #3（Part of #1）
- **调研方式**：GitHub API（gh CLI）读取仓库 README、官方文档与 package 元数据

---

## 1. 假设验证：pi 是否指 badlogic/pi-mono

**结论：成立，置信度高。**

`badlogic/pi-mono`（文档内链接指向 `earendil-works/pi-mono`，两者均解析到同一仓库，疑似组织迁移）是 Mario Zechner（badlogic）的开源项目：

- 自述定位："AI agent toolkit: unified LLM API, agent loop, TUI, coding agent CLI"，即**编码 agent harness**，含「自我可扩展的编码 agent」
- TypeScript monorepo，**MIT License**，约 82k stars，2026-08-02 当天仍在活跃提交
- 官网 pi.dev，文档 pi.dev/docs/latest
- npm 包：`@earendil-works/pi-coding-agent`（CLI 名 `pi`），当前版本 0.83.0

**其他「pi agent」候选排查：**

| 候选 | 是什么 | 判断 |
|---|---|---|
| Physical Intelligence π0 / openpi | 机器人视觉-语言-动作模型（VLA），开源 |  robotics 领域，与本项目「文档排版 agent」语境不符，排除 |
| Inflection AI 的 Pi | 闭源聊天机器人产品 | 非开源、不可本地嵌入，排除 |
| Raspberry Pi 上的各类 agent 项目 | 泛指树莓派部署 | 无单一知名同名 agent 项目，排除 |

用户语境是「本地接入 + 自定义改造 + 编码/文档能力」，唯一匹配的是 badlogic/pi-mono。

## 2. pi 的形态与嵌入方式

pi 是 monorepo，分层清晰，**同时提供 CLI、SDK、RPC 三种形态**：

| 包 | 作用 |
|---|---|
| `@earendil-works/pi-ai` | 统一多 provider LLM API（OpenAI/Anthropic/Google 等） |
| `@earendil-works/pi-agent-core` | agent 运行时（工具调用、状态管理） |
| `@earendil-works/pi-coding-agent` | 交互式编码 agent CLI（`pi` 命令），**内含 SDK** |
| `@earendil-works/pi-tui` | 终端 UI 库 |

**嵌入 Node.js / Electron 主进程的三条路径（官方文档明确）：**

1. **进程内 SDK（官方对 Node.js 应用的首选）**：`npm install @earendil-works/pi-coding-agent`，调用 `createAgentSession()` 得到 `AgentSession`，用 `session.prompt()` 发指令、`session.subscribe(event => ...)` 订阅事件流（text_delta、工具调用、compaction 等），可 `setModel`、`abort`、`compact`、`steer/followUp`。SDK 文档原话即列举「Build a custom UI (web, desktop, mobile)」为典型用例。
2. **RPC 子进程模式**：`pi --mode rpc`，stdin/stdout 走 JSONL 协议（prompt/steer/abort/new_session 等命令 + 事件流），适合跨语言或需要进程隔离的场景。注意 JSONL 只能按 `\n` 分帧，Node `readline` 会因 U+2028/U+2029 违规拆行，需自实现分帧（官方提供 `rpc-client.ts` 参考）。
3. **JSON 事件流模式**：`pi --mode json "prompt"`，一次性任务输出全部事件为 JSON lines；另有 `-p/--print` 打印模式。

**对 Electron 的建议**：主进程内嵌 SDK（路径 1）。Electron 35+ 内置 Node ≥22，满足引擎要求；事件流经 IPC 转发给渲染进程做 UI。若担心 agent 崩溃拖垮主进程或需要权限隔离，退化为路径 2（spawn `pi --mode rpc` 子进程）。

## 3. 模型 provider 与 API key 配置

**内置 provider 覆盖面极广**（env 变量或 `~/.pi/agent/auth.json` 配置）：

- 国际：Anthropic、OpenAI、Azure OpenAI、Google Gemini、Amazon Bedrock、Mistral、Groq、Cerebras、xAI、OpenRouter、Vercel AI Gateway、Fireworks、Together、Hugging Face、NVIDIA NIM、Cloudflare 等
- **国产模型原生支持**：DeepSeek、Kimi（KIMI_API_KEY）、MiniMax（含中国区）、Qwen Token Plan（含中国区）、Xiaomi MiMo（含中国区/国际多站点）、ZAI Coding Plan、Ant Ling
- **订阅 OAuth**：ChatGPT Plus/Pro（Codex）、Claude Pro/Max、GitHub Copilot、xAI、OpenRouter，`/login` 交互登录，token 存 `auth.json` 自动续期
- **本地模型**：llama.cpp（`/llama` 命令下载/加载本地模型）

**OpenAI 兼容网关：支持，且是官方扩展点。** 两种方式：

1. `pi.registerProvider("my-provider", { baseUrl, apiKey: "$ENV_VAR", api: "openai-completions", models: [...] })`（扩展内注册，可覆盖内置 provider 的 baseUrl 走企业代理）
2. `models.json` 配置文件声明自定义 provider/网关

即自建的 OpenAI 兼容网关（One-API、NewMax、vLLM 等）可直接接入。

## 4. 工具扩展机制（重点）

**Extensions 系统**是核心扩展点：TypeScript 模块（jiti 加载，免编译），放在 `~/.pi/agent/extensions/*.ts`（全局）或 `.pi/extensions/*.ts`（项目级），`/reload` 热重载；也可通过 npm/git 以「pi packages」分发。

**注册自定义工具**（我们要的「文档编辑」「模板读取」「版本存取」正是这条路）：

```typescript
import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import { Type } from "typebox";

export default function (pi: ExtensionAPI) {
  pi.registerTool({
    name: "doc_edit",
    label: "Document Edit",
    description: "对 docx 文档应用排版编辑",
    promptSnippet: "Edit Word documents",
    promptGuidelines: ["Use doc_edit instead of raw file writes when modifying .docx files."],
    parameters: Type.Object({ /* typebox schema */ }),
    async execute(toolCallId, params, signal, onUpdate, ctx) {
      // onUpdate 可流式回报进度；throw 报错（isError）
      return { content: [{ type: "text", text: "..." }], details: {} };
    },
    // 可选 renderCall/renderResult 自定义 UI 渲染
  });
}
```

要点：

- **Schema 用 typebox**（JSON Schema）；字符串枚举必须用 `@earendil-works/pi-ai` 的 `StringEnum`（Google API 兼容性）
- **改文件的工具必须包 `withFileMutationQueue(absPath, fn)`**，与内置 edit/write 共享逐文件串行队列，避免并行工具调用互相覆盖
- **动态工具**：`pi.setActiveTools()` 运行时增删工具；`prepareArguments()` 兼容旧会话参数
- 扩展能力远不止工具：事件钩子（session_start、tool_call、before_agent_start、compaction 自定义等）、自定义命令 `/cmd`、快捷键、`pi.exec` 执行命令、`pi.appendEntry()` 会话持久化、`ctx.ui` 交互组件

**权限模型（重要警告）**：pi **没有内置权限系统**——README 明言 "Pi does not include a built-in permission system"，以启动用户的完整权限运行。权限门需自行实现：

- 官方做法：用 `tool_call` 事件拦截（如检测 `rm -rf` 弹确认，`return { block: true }`）
- 项目级资源（`.pi/extensions` 等）有 trust 机制（`trust.json`、非交互模式 `--approve`）
- 更强隔离需容器化（官方文档给 Gondolin micro-VM / Docker / OpenShell 三方案）
- **对本项目的含义**：「文档编辑/版本存取」工具的权限边界（限制可写目录、版本提交前确认）要我们在扩展里自己把关

**Skills 系统**：实现 Agent Skills 开放标准（agentskills.io），`SKILL.md` 目录式能力包，agent 按需加载；可从 `~/.pi/agent/skills/`、`.pi/skills/` 发现，且兼容复用 Claude Code（`~/.claude/skills`）与 Codex 的 skills。排版工作流知识（如「学位论文排版规范」）适合做成 skill 而非工具。

**Prompt templates**：`.pi/prompts/*.md` 变成 `/name` 斜杠命令，适合预置排版任务模板。

## 5. 可改造点清单

### 官方扩展点（无需 fork）

| 需求 | 扩展点 |
|---|---|
| 注册「文档编辑/模板读取/版本存取」工具 | extension `pi.registerTool()` |
| 工具权限门（写目录白名单、危险操作确认） | `tool_call` 事件拦截 + `block` |
| **系统提示词修改/追加** | `before_agent_start` 事件：`event.systemPrompt` 可改写/追加，`systemPromptOptions` 暴露完整结构化输入（custom prompt、guidelines、工具片段、上下文文件、skills） |
| 注入排版领域上下文 | 上下文文件（AGENTS.md 等）、`before_agent_start` 注入消息、skills |
| 自定义上下文压缩策略 | compaction 自定义事件（`compact(customInstructions)`、compaction 事件钩子） |
| 接入自建 LLM 网关 | `pi.registerProvider()` / models.json（openai-completions API） |
| 自定义 UI（Electron 界面） | **SDK `createAgentSession()` + 事件订阅**，完全绕开 TUI；这是官方列举的头号用例 |
| 排版工作流知识 | skills（SKILL.md）、prompt templates |
| 禁用/裁剪内置工具（如不要 bash） | `createAgentSession({ tools: ["read", ...] })` 或 `pi.setActiveTools()` |

### 需要 fork 改造的点（评估后其实很少）

- **agent 循环内核行为**（`pi-agent-core` 的重试、并行工具调度策略等超出事件钩子可及范围的修改）
- **内置工具的实现替换**（拦截可以，改写源码行为需 fork；但我们可以干脆禁用内置工具、全用自注册工具替代）
- **内置权限沙箱**（官方立场是「不做，去容器化」，若产品化需要进程级沙箱只能自行包装或 fork）
- 值得注意：连系统提示词整体替换都能走 `before_agent_start`，UI 走 SDK 完全自绘——**fork 的必要性很低，建议先纯 extension + SDK 路线**

## 6. 运行要求

- **Node.js ≥ 22.19.0**（根 package.json 与 coding-agent 的 engines 均明确）；对应 Electron ≥ 35 的内置 Node
- TypeScript monorepo（npm workspaces），扩展经 jiti 直接跑 TS 免编译
- 独立二进制用 Bun 编译（`scripts/build-binaries.sh`），可脱离 Node 分发
- 资源占用：官方未给量化指标；形态为单 Node 进程 CLI，常态内存占用为典型 Node 应用量级（数百 MB 以内），主要开销在 LLM 侧网络与 token
- **License：MIT**，可自由商用、修改、嵌入闭源产品

## 7. 接入架构建议

```
┌─────────────────────────── Electron 应用 ───────────────────────────┐
│  渲染进程：排版工作台 UI（文档预览、diff、版本时间线）                  │
│      ▲ IPC（AgentSessionEvent 转发 / 用户指令下发）                   │
│  主进程：                                                            │
│    ┌──────────────────────────────────────────────────────┐        │
│    │ @earendil-works/pi-coding-agent (SDK, 进程内)          │        │
│    │  createAgentSession({                                │        │
│    │    resourceLoader: 加载自研 extensions,               │        │
│    │    tools: 裁剪内置工具集                               │        │
│    │  })                                                   │        │
│    └──────────────────────────────────────────────────────┘        │
│        │                                                           │
│    自研 extension（.pi/extensions/ 或随应用打包路径）：                │
│      ├─ tool: doc_edit      文档编辑（走 docx 服务层，               │
│      │                      withFileMutationQueue 保护）             │
│      ├─ tool: template_read 模板读取                                 │
│      ├─ tool: version_store 版本存取（每次编辑自动快照）               │
│      ├─ hook: tool_call     权限门（可写目录白名单、确认框）           │
│      ├─ hook: before_agent_start 注入排版系统提示词                   │
│      └─ provider: 自建 OpenAI 兼容网关（registerProvider）            │
│        │                                                           │
│    现有 Node 服务层：docx 解析/模板库/版本库（被工具调用复用）          │
└──────────────────────────────────────────────────────────────────────┘
        │ HTTPS（OpenAI 兼容协议）
   LLM 网关（DeepSeek / Kimi / Qwen / Anthropic 等均可热切换）
```

要点：

1. **SDK 进程内嵌入为首选**，RPC 子进程为降级/隔离备选；不要试图复用 pi 的 TUI
2. 「文档编辑/模板读取/版本存取」三工具全部走 `registerTool`，schema 用 typebox + StringEnum，写操作包 `withFileMutationQueue`
3. 权限门自研：pi 无内置权限系统，必须在 `tool_call` 钩子里做目录白名单与危险操作确认
4. 排版规范知识做 skill，任务模板做 prompt template，系统提示词用 `before_agent_start` 注入
5. 预计无需 fork；确需深度改造时 MIT License 无障碍

## 8. 来源 URL 清单（查询日期 2026-08-02）

- https://github.com/badlogic/pi-mono （README、LICENSE、package.json）
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/sdk.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/rpc.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/json.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/extensions.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/providers.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/custom-provider.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/skills.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/prompt-templates.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/compaction.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/containerization.md
- https://github.com/earendil-works/pi-mono/tree/main/packages/coding-agent/docs/usage.md
- https://www.npmjs.com/package/@earendil-works/pi-coding-agent
- https://pi.dev/docs/latest
- 排除项候选：https://github.com/Physical-Intelligence/openpi （π0 机器人模型）
