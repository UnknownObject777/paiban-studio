# 调研报告：pi SDK 的会话持久化与多会话能力

Resolution of #7 (Part of #6)

- **查询日期**：2026-08-02
- **关联 ticket**：wayfinder research #7（Part of #6）
- **调研方式**：npm registry 拉取 `@earendil-works/pi-coding-agent@0.83.0` tarball，直接阅读包内 `docs/`（sdk.md、sessions.md、session-format.md、extensions.md）与 `dist/core/` 编译源码（sdk.js、agent-session.js、session-manager.js、agent-session-runtime.d.ts、resource-loader.js）；一手源码/文档为准
- **结论置信度**：高（SDK 文档与编译源码相互印证；标注「推断」的条目除外）

---

## 结论速览

1. **会话可序列化、可持久化、可完整重建**。SDK 自带 `SessionManager`，会话以 JSONL 树结构文件自动落盘（默认 `~/.pi/agent/sessions/`，可自定义目录）；消息历史、模型/thinking 切换、compaction 检查点、分支摘要、扩展状态全部在文件内，无需自行设计序列化格式。
2. **重启后可恢复并继续对话**（不只是查看历史）：`SessionManager.open(path)` / `SessionManager.continueRecent(cwd)` 打开旧文件后传给 `createAgentSession()`，SDK 内部自动 `buildSessionContext()` 重建 `agent.state.messages`，`session.prompt()` 直接接着聊。这是 SDK 官方示例（examples/sdk/11-sessions.ts）覆盖的标准用法。
3. **多会话并行切换可行**：`AgentSession` 是普通对象，同一进程可同时存活多个实例（每个绑定各自的 `SessionManager` / 会话文件）；另有 `AgentSessionRuntime` 提供「单活跃会话热替换」模型（`switchSession()` / `newSession()` / `fork()`），与 pi 自身 `/resume` 同一层。
4. **已知限制**：会话文件**无文件锁**，两个 SessionManager 同时写同一文件会交错损坏（须应用层保证一个会话文件同一时刻只有一个活跃实例）；共享一个 `ResourceLoader` 时扩展模块级状态可能被多会话共享（jiti 模块缓存，待验证）；`AgentSessionRuntime` 是替换模型而非多活模型，事件订阅与扩展绑定在替换后须重挂。
5. **对 paiban-studio**：D5（主进程内嵌 SDK）不变。建议「每个排版任务一个持久化会话文件 + 应用层 SessionRegistry 管理多实例」，sessionDir 指到应用数据目录而非 `~/.pi`。重启恢复与多会话切换两个目标均落在 SDK 已验证能力内，无需 fork、无需自研序列化。

---

## Q1：`createAgentSession()` 的会话状态/事件流能否序列化、持久化、重建？

### 已验证事实

**能，且是 SDK 内建能力，不需要应用层自己做。**

- 会话状态由 `SessionManager`（`dist/core/session-manager.js`，文档 `docs/session-format.md`）管理，以 **JSONL 文件**持久化，每行一个 entry，`type` 字段区分类型。entry 通过 `id`/`parentId` 组成**树结构**（v3 格式），支持原地分支。
- 文件位置：`~/.pi/agent/sessions/--<cwd 编码>--/<timestamp>_<uuid>.jsonl`；`SessionManager.create(cwd, sessionDir?)` 的第二参数可自定义目录（官方示例 11-sessions.ts 注释中明确）。
- 重建路径（`dist/core/sdk.js` 实证）：

  ```js
  // sdk.js:81  createAgentSession 内部
  const existingSession = sessionManager.buildSessionContext();
  ...
  // sdk.js:233
  agent.state.messages = existingSession.messages;
  ```

  即打开旧会话文件创建 `AgentSession` 时，SDK 自动从 entry 树重建完整 LLM 上下文（含 compaction 折叠逻辑）注入 agent 状态。
- **事件流（`AgentSessionEvent`）本身不落盘**——它是实时流式通知（text_delta、tool_execution_start/end 等），订阅只在会话存活期间有效。但事件流的**结果**（每条完成的 message、toolResult、model_change、compaction）都会作为 entry 追加进 JSONL，因此「事件流 → 可持久化状态」的转换由 SDK 自动完成。UI 重放历史用 `session.messages` / `SessionManager.getEntries()`，而不是回放事件。

### 含义

应用层要存的就是 **一个会话文件路径**（或 session id），其余全在 JSONL 里。事件流不需要也不应该持久化。

## Q2：持久化一个会话需要存什么？SDK 自带什么机制？缺口是什么？

### 已验证事实

会话文件（JSONL）已包含重建所需的全部内容（`docs/session-format.md` 逐类型定义）：

| 内容 | entry 类型 | 说明 |
|---|---|---|
| 会话元数据 | `session`（header，含 version/cwd/id/可选 parentSession） | 首行 |
| 对话历史 | `message`（user / assistant / toolResult / bashExecution / custom / branchSummary / compactionSummary 七种 `AgentMessage`） | 完整含 usage、stopReason、时间戳 |
| 模型切换 | `model_change` | 恢复时还原模型（失败则回退默认并返回 `modelFallbackMessage`） |
| thinking level | `thinking_level_change` | 同上 |
| 上下文压缩检查点 | `compaction`（含 `summary` + `retainedTail` 自包含检查点） | 重启后不必回溯 compaction 之前的条目 |
| 分支摘要 | `branch_summary` | `/tree` 切分支时保留被弃路径的上下文 |
| **扩展自定义状态** | `custom`（`pi.appendEntry(customType, data)` 写入，**不进 LLM 上下文**） | 扩展重启后扫描自己 customType 的 entry 重建内部状态 |
| 扩展注入的上下文消息 | `custom_message`（进 LLM 上下文） | |
| 书签 | `label`；显示名 `session_info` | |

**SDK 自带机制**：自动追加写盘（`appendFileSync`，同步写，落盘即持久，无缓冲丢失窗口）；整树变更时 `_rewriteFile()` 全量重写（`session-manager.js:693-751`）；`SessionManager.list/listAll` 枚举历史会话；`/resume`、`pi -c`、fork/clone/import 全套 CLI/Runtime 流程；旧版本（v1/v2）文件加载时自动迁移到 v3。

**扩展状态持久化**（`docs/extensions.md:1439`）：`pi.appendEntry("my-state", {...})` 是官方给扩展的持久化通道——我们的「文档编辑/版本存取」工具若有跨重启状态（如版本库指针、文档句柄映射），走这条即可。

### 缺口（SDK 不覆盖、需应用层负责）

1. **会话 ↔ 业务实体的映射**：哪个会话对应哪个排版任务/文档，pi 不管，需 paiban-studio 自己维护（可用 `session_info` 名称、自定义 sessionDir 布局或自建索引）。
2. **事件流回放**：重启后 UI 要从 `session.messages` / entries 重建聊天记录，SDK 不提供「事件流重放」。
3. **工具的外部副作用状态**：会话文件只记 toolResult 文本；工具操作过的文档本身、版本库内容的持久化是我们服务层的职责（本来就在 D5 架构的服务层内）。
4. **多实例并发写保护**：见 Q4。

## Q3：应用重启后能否恢复会话并继续对话？具体 API？

### 已验证事实

**可以，且是官方示例的标准用法**（`examples/sdk/11-sessions.ts`、`docs/sdk.md` §Session Management）：

```typescript
import { createAgentSession, SessionManager } from "@earendil-works/pi-coding-agent";

// 方式 A：按路径打开指定会话文件
const { session, modelFallbackMessage } = await createAgentSession({
  sessionManager: SessionManager.open("/path/to/session.jsonl"),
});
await session.prompt("接着上次继续……");   // 直接继续对话

// 方式 B：继续该 cwd 最近的会话（无则新建）
SessionManager.continueRecent(cwd);

// 方式 C：先枚举再选择
const list = await SessionManager.list(cwd);        // 含 id、首条消息摘要、路径
SessionManager.open(list[0].path);
```

恢复语义（`sdk.js:81/233` + `docs/session-format.md` §Context Building 实证）：

- `buildSessionContext()` 从当前 leaf 走到根，应用 compaction 折叠（`retainedTail` 自包含检查点），产出 messages + thinkingLevel + model，注入 `agent.state.messages`。
- 恢复的是**完整 LLM 上下文**（含工具调用与结果、分支摘要），不是只读快照——`prompt()` 之后新 entry 继续追加到同一文件。
- 模型恢复失败（如原 provider 无 key）时回退默认模型并经 `modelFallbackMessage` 告知。
- 还能用树 API 做更花的恢复：`sm.branch(entryId)` 回到历史任意点继续、`session.navigateTree()`、`runtime.fork(entryId)` 分叉。

## Q4：同时存活多个会话实例是什么形态？已知限制？

### 已验证事实

**形态：`AgentSession` 是普通实例对象，天然支持多活。** 每次 `createAgentSession()` 独立构造 Agent、SessionManager、ExtensionRunner（`sdk.js` 全文无模块级可变单例；`agent-session.js` 的扩展回调在实例字段 `_extensionRunner` 上）。两个 session 各自 `subscribe`、各自 `prompt`，互不干扰。事件订阅挂在具体 `AgentSession` 上，返回 unsubscribe 函数，`dispose()` 清理。

另有一层 **`AgentSessionRuntime`**（`createAgentSessionRuntime()`，`agent-session-runtime.d.ts`）：这是 pi 自身 interactive/RPC 模式用的**单活跃会话热替换**模型——`runtime.newSession()` / `switchSession(path)` / `fork(entryId)` / `importFromJsonl()` 替换 `runtime.session` 指向的当前会话。注意它是「替换」而非「多活」：替换后旧订阅失效，须重新 `session.subscribe()` 和 `session.bindExtensions(...)`（sdk.md 明示）。

### 已知限制与风险

1. **会话文件无锁（已验证）**：session-manager 写盘用裸 `appendFileSync`/`writeFileSync`，全文检索无 flock/lockfile 机制。**同一进程或跨进程两个实例同时 `open` 同一个 JSONL 并写入，会交错损坏文件。** 应用层必须保证「一个会话文件同一时刻只有一个活跃 AgentSession」（自建 SessionRegistry 或打开前检查）。
2. **共享 `ResourceLoader` 的扩展状态（推断，待验证）**：`DefaultResourceLoader.getExtensions()` 返回缓存的 `extensionsResult`；扩展经 jiti 加载，jiti 按路径缓存模块——若多个会话共享一个 loader，扩展文件的**模块级顶层状态**可能在会话间共享。我们的工具扩展若持有任务级状态，应放在 `bindExtensions` 后的实例字段或 `pi.appendEntry` 持久化里，避免模块级单例。隔离做法：每会话一个 loader 实例（成本是重复发现/加载）。
3. **共享配置文件的写竞争（已验证）**：`SettingsManager` setter 异步排队写 `settings.json`；`auth.json`/`models.json` 同理。多会话共享同一份全局配置目录时并发改设置可能互相覆盖。规避：多会话共享只读配置，写操作集中串行。
4. **cwd 绑定**：每个 session 绑定一个 cwd（工具相对路径、会话目录命名都以此为准）；`switchSession` 提供 `cwdOverride`。多会话各自不同 cwd 没有问题——SDK 中未发现 `process.chdir` 调用，cwd 是纯参数。
5. **生成并行度**：题目只要求并行存活/切换，不要求并行生成；若日后要并行生成，每个 session 独立发 LLM 请求即可，瓶颈在 API 配额而非 SDK。

## Q5：综合方案——重启可恢复 + 多会话并行切换

### 推荐做法（全部落在 SDK 已验证能力内）

```
Electron 主进程
└─ SessionRegistry（自研，薄层）
   ├─ Map<taskId, AgentSession>          // 活跃会话实例（多活）
   ├─ Map<taskId, sessionFilePath>       // 持久化索引（自己存，如 SQLite/JSON）
   └─ 保证同一 sessionFile 全局只有一个活跃实例（针对限制 1）

每个会话：
  createAgentSession({
    sessionManager: 新建 → SessionManager.create(cwd, appSessionDir)
                    恢复 → SessionManager.open(登记的 sessionFilePath),
    modelRuntime: 共享一个（只读凭证）,
    resourceLoader: 每会话独立实例（规避限制 2，代价可接受）,
    settingsManager: 共享但只读 / 写操作集中,
  })
```

- **新建任务**：`SessionManager.create(taskCwd, appSessionDir)`，`appSessionDir` 指向应用数据目录（不用默认 `~/.pi/agent/sessions`，避免与用户本机 pi CLI 互相污染）。
- **重启恢复**：应用启动时读自建索引，对每个任务 `SessionManager.open(path)` + `createAgentSession()`；UI 从 `session.messages` 渲染历史，用户输入直接 `prompt()` 续聊。也可惰性恢复（切到某任务时才 open）。
- **切换任务**：前台订阅切到目标 session 的事件转发；后台 session 保持实例存活但不生成（除非用户显式触发）。
- **扩展状态**：排版工具的跨重启状态用 `pi.appendEntry(customType, data)` 存进会话文件，扩展 `session_start` 钩子里扫描自家 customType 重建。

### 备选：单活跃模型

若产品上一个窗口只需一个活跃会话，可直接用 `AgentSessionRuntime.switchSession()`（pi `/resume` 同款），省去自研 Registry 的多活管理，代价是切换时要重挂订阅/扩展、且切换有重建开销。

### 风险清单

| 风险 | 等级 | 缓解 |
|---|---|---|
| 同文件双写损坏（无锁） | 中 | Registry 强制单活跃实例；跨进程场景（如同时开 pi CLI 操作同一 sessionDir）用独立 sessionDir 规避 |
| 扩展模块级状态跨会话泄漏 | 低-中（待验证） | 每会话独立 ResourceLoader；扩展状态走实例字段 + appendEntry |
| 事件订阅/扩展绑定在 runtime 替换后失效 | 低 | 用多活模型则不存在；用 runtime 模型则按文档重挂 |
| compaction 后旧 entry 仍在文件中，文件随会话增长 | 低 | JSONL 追加写；必要时 `createBranchedSession` 或导出清理 |
| SDK 版本升级改会话格式 | 低 | 有 version 字段 + 自动迁移（v1→v3 已实证） |

---

## 事实 / 推断标注汇总

**SDK 已验证事实**（出处：包内 docs 与 dist 源码，版本 0.83.0）：JSONL 树格式与全部 entry 类型；`SessionManager.create/open/continueRecent/inMemory/forkFrom`、树导航、append 系列 API；`createAgentSession` 恢复路径（sdk.js:81/233）；同步落盘、无文件锁；`AgentSessionRuntime` 替换语义与重挂要求；`pi.appendEntry` 扩展持久化；SDK 无模块级会话单例；无 `process.chdir`。

**推断 / 待验证**：jiti 模块缓存导致共享 loader 时扩展模块级状态跨会话共享（建议每会话独立 loader 直接规避，不必验证）；多会话并行 prompt 的实际稳定性（架构上无共享态，但未做压测）；`SettingsManager` 异步写排队在崩溃时的丢失窗口（文档自述用 `flush()` 做持久化边界）。

## 对 paiban-studio 的含义

1. **D5 决策不变且更稳**：此前 D5 只验证了「内嵌 SDK + 事件转发」，本调研补齐了持久化与多会话两块，均无需 fork、无需自研序列化层。
2. **R1 运行时骨架可直接落**：会话文件 + SessionRegistry（多活管理 + 单文件写保护）+ 应用侧任务↔会话索引，是 MVP 需要新增的唯一自研薄层。
3. **会话目录选址**：用 `SessionManager.create(cwd, appSessionDir)` 把会话文件收进应用数据目录，与 `agentDir`（auth/models/settings）一并规划，避免依赖 `~/.pi`。
4. **扩展状态 persistence 有官方通道**（`appendEntry`/`custom` entry），版本存取工具的检查点、文档映射等可随会话文件一起走，天然获得重启恢复。

## 来源清单（查询日期 2026-08-02）

一手材料均为 `@earendil-works/pi-coding-agent@0.83.0` npm tarball 包内文件（上游仓库 [earendil-works/pi-mono](https://github.com/earendil-works/pi-mono) `packages/coding-agent`）：

- `docs/session-format.md` — JSONL 格式、entry 类型、SessionManager API 全表
- `docs/sessions.md` — 会话存储、`/resume`、`/tree`、fork/clone 语义
- `docs/sdk.md` — `createAgentSession` / `AgentSessionRuntime` / Session Management 章节
- `docs/extensions.md` — `pi.appendEntry`、custom entry、扩展持久化
- `dist/core/sdk.js`（L67-247）— createAgentSession 恢复路径实证
- `dist/core/session-manager.js`（L590-751）— 落盘机制、无锁实证
- `dist/core/agent-session-runtime.d.ts` — switchSession/newSession/fork/importFromJsonl 签名
- `examples/sdk/11-sessions.ts`、`13-session-runtime.ts` — 官方会话管理示例
