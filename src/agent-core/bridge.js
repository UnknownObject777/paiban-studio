// agent-core/bridge.js — pi agent 接入层（spec 模块 2，调研 #3 路径 1：SDK 主进程内嵌）。
//
// 职责：
//   - createAgentSession() 内嵌，tools 白名单只放自研四工具（裁剪内置文件工具，防绕过编辑内核）
//   - LLM provider（R4）：Anthropic（apiKey 直连）/ OpenAI 兼容端点（models.json 声明，含
//     DeepSeek/Kimi/Qwen/Ollama/vLLM/本地网关）；界面配置 > 环境变量 > 默认值
//   - 事件流翻译为渲染层友好事件（text_delta / tool_start / tool_end / done / error）
//   - 发送 prompt 前注入排版系统提示 + 当前文档结构摘要（before_agent_start 的 MVP 等价物）
//   - SDK / 凭证缺失时优雅降级：status() 报告未就绪，工作台其余功能不受影响

import { join } from 'node:path';
import { mkdirSync, writeFileSync } from 'node:fs';
import { createTools } from './tools.js';

const SYSTEM_PRIMER = `你是「排版工作台」的文档排版助手，专门处理中文公文 Word 文档（.docx）。

工作方式（必须遵守）：
1. 一切文档修改只能调用 doc_edit 工具；不要尝试直接读写文件。
2. 修改前先调用 doc_outline 获取段落路径，按路径精确寻址（/body/p[N]/r[M]）。
3. 每次 doc_edit 成功后系统自动保存新版本；改坏了可用 version_store 回滚。
4. 工具调用失败时阅读错误里的 suggestion 字段并自我修正后重试。
5. 用户说"按某模板排"时：template_read 取 rulesetCommands，原样传给 doc_edit。

中文公文排版常识：
- 字号：三号=16pt、小三=15pt、四号=14pt、小四=12pt、五号=10.5pt
- 正文：仿宋_GB2312、四号或三号、行距 28 磅（lineSpacingPt:28）、首行缩进 2 字符（firstLineChars:200）、两端对齐（align:"justify"）
- 标题：黑体、三号、居中（align:"center"）
- 页面：上 3.7cm 下 3.5cm 左 2.8cm 右 2.6cm（marginsCm）
- 中文字体用 eastAsia 属性设置；西文用 ascii。

回复要求：简要说明做了什么修改、产生的新版本号；不要输出大段无关解释。`;

export class AgentBridge {
  constructor(workspace) {
    this.workspace = workspace;
    this.session = null;
    this.listeners = new Set();
    this.ready = false;
    this.statusInfo = { ready: false };
    this._docPrimed = new Set(); // docId → 已注入结构摘要
  }

  onEvent(fn) {
    this.listeners.add(fn);
  }

  _emit(event) {
    for (const fn of this.listeners) {
      try { fn(event); } catch { /* 渲染层断开不致命 */ }
    }
  }

  async init() {
    const cfg = this.workspace.getFullConfig();
    const agentDir = join(this.workspace.baseDir, 'agent');
    mkdirSync(agentDir, { recursive: true });

    let sdk;
    try {
      sdk = await import('@earendil-works/pi-coding-agent');
    } catch (err) {
      this.statusInfo = { ready: false, reason: 'pi SDK 未安装: ' + err.message };
      return;
    }

    const usingGateway = !!cfg.baseUrl;
    const providerId = usingGateway ? 'paiban-gateway' : (cfg.provider || 'anthropic');

    // OpenAI 兼容端点：写 models.json（apiKey 经环境变量插值，不落明文到 models.json）
    let modelsPath = null;
    if (usingGateway) {
      process.env.PAIBAN_AGENT_KEY = cfg.apiKey || 'paiban-local';
      modelsPath = join(agentDir, 'models.json');
      writeFileSync(modelsPath, JSON.stringify({
        providers: {
          [providerId]: {
            baseUrl: cfg.baseUrl,
            api: 'openai-completions',
            apiKey: '$PAIBAN_AGENT_KEY',
            compat: { supportsDeveloperRole: false, supportsReasoningEffort: false },
            models: [{ id: cfg.model, name: cfg.model }],
          },
        },
      }, null, 2));
    } else if (cfg.apiKey) {
      // 内置 provider：经环境变量供 auth 解析（ANTHROPIC_API_KEY / OPENAI_API_KEY / DEEPSEEK_API_KEY 等）
      const envName = `${providerId.replace(/-/g, '_').toUpperCase()}_API_KEY`;
      if (!process.env[envName]) process.env[envName] = cfg.apiKey;
    }

    try {
      const modelRuntime = await sdk.ModelRuntime.create({
        authPath: join(agentDir, 'auth.json'),
        modelsPath,
      });
      const model = modelRuntime.getModel(providerId, cfg.model);
      if (!model) {
        const available = modelRuntime.getModels(providerId).map((m) => m.id).slice(0, 10);
        this.statusInfo = {
          ready: false,
          reason: `模型 ${providerId}/${cfg.model} 不可用${available.length ? '；可用: ' + available.join(', ') : ''}`,
          provider: providerId, model: cfg.model,
        };
        return;
      }
      const { session } = await sdk.createAgentSession({
        cwd: this.workspace.baseDir,
        agentDir,
        modelRuntime,
        model,
        sessionManager: sdk.SessionManager.inMemory(),
        tools: ['doc_outline', 'doc_edit', 'template_read', 'version_store'], // 白名单：只自研工具
        customTools: createTools(this.workspace),
      });
      this.session = session;
      this._wireEvents(session);
      this.ready = true;
      this.statusInfo = { ready: true, provider: providerId, model: cfg.model, gateway: usingGateway };
    } catch (err) {
      this.statusInfo = { ready: false, reason: err.message, provider: providerId, model: cfg.model };
    }
  }

  _wireEvents(session) {
    session.subscribe((event) => {
      const e = event;
      if (e.type === 'message_update' && e.assistantMessageEvent?.type === 'text_delta') {
        this._emit({ type: 'text_delta', delta: e.assistantMessageEvent.delta });
      } else if (e.type === 'tool_execution_start') {
        this._emit({ type: 'tool_start', name: e.toolName, args: summarizeArgs(e.args) });
      } else if (e.type === 'tool_execution_end') {
        this._emit({ type: 'tool_end', name: e.toolName, isError: !!e.isError, summary: summarizeToolResult(e) });
      } else if (e.type === 'agent_end') {
        this._emit({ type: 'done' });
      } else if (e.type === 'error' || e.type === 'agent_error') {
        this._emit({ type: 'error', message: e.error?.message || String(e.error || e) });
      }
    });
  }

  status() {
    return this.statusInfo;
  }

  async abort() {
    await this.session?.abort();
  }

  /**
   * 发送用户消息。docId 提供时注入文档上下文（首次注入结构摘要）。
   * @returns {Promise<{ ok: boolean, error?: string }>}
   */
  async send(docId, message) {
    if (!this.ready || !this.session) {
      return { ok: false, error: this.statusInfo.reason || 'agent 未就绪（请检查模型配置）' };
    }
    let prompt = message;
    if (docId) {
      let ctx = `\n\n[当前工作文档 docId: ${docId}]`;
      if (!this._docPrimed.has(docId)) {
        try {
          const outline = this.workspace.getOutline(docId, { textPreview: 30 });
          const preview = outline.paragraphs.slice(0, 40)
            .map((p) => `${p.path} ${p.text}`).join('\n');
          ctx += `\n[文档结构摘要（共 ${outline.paragraphCount} 段，前 40 段预览）]\n${preview}`;
        } catch { /* 大纲失败不阻塞对话 */ }
        this._docPrimed.add(docId);
      }
      prompt = message + ctx;
    }
    if (!this._primed) {
      prompt = SYSTEM_PRIMER + '\n\n---\n用户请求：' + prompt;
      this._primed = true;
    }
    this._emit({ type: 'user', message });
    try {
      await this.session.prompt(prompt);
      return { ok: true };
    } catch (err) {
      this._emit({ type: 'error', message: err.message });
      return { ok: false, error: err.message };
    }
  }
}

// ---- 事件摘要 ----

function summarizeArgs(args) {
  if (!args || typeof args !== 'object') return '';
  if (args.commands) return `${args.commands.length} 条命令${args.note ? '：' + args.note : ''}`;
  if (args.action) return args.action + (args.versionId ? ' ' + args.versionId : '');
  if (args.templateId) return args.templateId;
  return Object.keys(args).join(', ');
}

function summarizeToolResult(e) {
  const d = e.result?.details ?? e.details;
  if (!d || typeof d !== 'object') return '';
  if (d.version) return `${d.version.id}${d.versionCreated ? '（新版本）' : ''}，applied ${d.applied?.length ?? 0}，errors ${d.errors?.length ?? 0}`;
  if (d.versions) return `${d.versions.length} 个版本`;
  if (d.templates) return `${d.templates.length} 个模板`;
  if (d.paragraphCount !== undefined) return `${d.paragraphCount} 段`;
  return '';
}
