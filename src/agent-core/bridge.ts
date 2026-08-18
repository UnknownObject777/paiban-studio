// agent-core/bridge.ts — pi agent 接入层（spec 模块 2，调研 #3 路径 1：SDK 主进程内嵌）。
//
// 职责：
//   - createAgentSession() 内嵌，tools 白名单只放自研七工具（裁剪内置文件工具，防绕过编辑内核）
//   - LLM provider（R4）：Anthropic（apiKey 直连）/ OpenAI 兼容端点（models.json 声明，含
//     DeepSeek/Kimi/Qwen/Ollama/vLLM/本地网关）；界面配置 > 环境变量 > 默认值
//   - 事件流翻译为渲染层友好事件（text_delta / tool_start / tool_end / done / error）
//   - 发送 prompt 前注入排版系统提示 + 当前文档结构摘要（before_agent_start 的 MVP 等价物）
//   - SDK / 凭证缺失时优雅降级：status() 报告未就绪，工作台其余功能不受影响

import { join } from 'node:path';
import { mkdirSync, writeFileSync } from 'node:fs';
import { createTools } from './tools.js';
import type { Workspace } from '../server/workspace.js';

const SYSTEM_PRIMER = `你是「排版工作台」的文档排版助手，专门处理中文办公 Word 文档（.docx）：排版已有文档、或从零生成新文档（标书、实验报告、外汇申报单等）。

工作方式（必须遵守）：
1. 修改已有文档只能调用 doc_edit；从零生成新文档调用 doc_generate。不要尝试直接读写文件。
2. 修改前先调用 doc_outline 获取段落路径，按路径精确寻址（/body/p[N]/r[M]）。
3. 每次 doc_edit 成功后系统自动保存新版本；改坏了可用 version_store 回滚。
4. 工具调用失败时阅读错误里的 suggestion 字段并自我修正后重试。
5. 涉及「金额（大写）」栏（投标报价表、外汇申报单）时，必须先调用 amount_words 换算得到大写后填入，禁止心算大写。
6. 用户说"按实验报告排版"→ 用 ruleset_read 取内置规则集 lab-report-default；"按公文排版"→ gongwen-default；"按标书/投标文件排版"→ bid-default；"按外汇申报单/涉外收付款申报表单排版"→ fx-form-default；"按某个上传的模板排"→ template_read。取到 rulesetCommands 后原样传给 doc_edit。内置规则集是手写资产，优先于上传模板反推的规则集。

第二工作模式（写/生成/起草新文档）：
1. 用户要"写/生成/起草"新文档（标书、实验报告、外汇申报单等）时：先用 markdown 起草内容，再调 doc_generate（rulesetId 用 ruleset_read 列出的内置规则集 id，如 bid-default / lab-report-default / fx-form-default / gongwen-default）；生成后用一两句话汇报 docId 与文档名；用户后续要改就走 doc_outline + doc_edit。
2. markdown 约定：# 文档题目（只一个）、## 章节标题、GFM 表格做报价表/字段表、**粗体** 强调。
3. 领域结构（起草时遵循；缺失信息用「×××」占位，不要编造具体公司名/金额）：
   - 投标文件：封面信息（项目名称/编号/投标人）→ 目录 → 投标函 → 法定代表人身份证明/授权委托书 → 商务部分（资质/业绩/财务）→ 技术部分（方案/实施计划/质量保证/服务承诺）→ 报价部分（开标一览表/分项报价表）→ 资格审查资料。
   - 涉外收付款申报单（外汇单）：以字段表格为主——申报号码（银行编制）、日期、汇款人/收款人名称、主体标识码、结算方式（电汇/票汇/信汇）、币种及金额、交易编码、交易附言、申请人签章栏。
   - 实验报告：实验目的 → 实验原理 → 仪器与材料 → 实验步骤 → 数据记录与处理（表格）→ 结果分析 → 结论。

中文公文排版常识：
- 字号：三号=16pt、小三=15pt、四号=14pt、小四=12pt、五号=10.5pt
- 正文：仿宋_GB2312、四号或三号、行距 28 磅（lineSpacingPt:28）、首行缩进 2 字符（firstLineChars:200）、两端对齐（align:"justify"）
- 标题：黑体、三号、居中（align:"center"）
- 页面：上 3.7cm 下 3.5cm 左 2.8cm 右 2.6cm（marginsCm）
- 中文字体用 eastAsia 属性设置；西文用 ascii。

回复要求：简要说明做了什么修改、产生的新版本号；不要输出大段无关解释。`;

/** agent 工具白名单：只放自研工具（裁剪内置文件工具，防绕过编辑内核）。 */
export const TOOL_WHITELIST = ['doc_outline', 'doc_edit', 'doc_generate', 'template_read', 'ruleset_read', 'version_store', 'amount_words'];

export interface AgentStatus {
  ready: boolean;
  reason?: string;
  provider?: string;
  model?: string;
  gateway?: boolean;
}

export type AgentEvent =
  | { type: 'text_delta'; delta: string }
  | { type: 'tool_start'; name: string; args: string }
  | { type: 'tool_end'; name: string; isError: boolean; summary: string; details?: unknown }
  | { type: 'done' }
  | { type: 'error'; message: string }
  | { type: 'user'; message: string };

export class AgentBridge {
  workspace: Workspace;
  session: any;
  listeners: Set<(event: AgentEvent) => void>;
  ready: boolean;
  statusInfo: AgentStatus;
  _docPrimed: Set<string>;
  _primed: boolean;

  constructor(workspace: Workspace) {
    this.workspace = workspace;
    this.session = null;
    this.listeners = new Set();
    this.ready = false;
    this.statusInfo = { ready: false };
    this._docPrimed = new Set(); // docId → 已注入结构摘要
    this._primed = false;
  }

  onEvent(fn: (event: AgentEvent) => void): void {
    this.listeners.add(fn);
  }

  _emit(event: AgentEvent): void {
    for (const fn of this.listeners) {
      try { fn(event); } catch { /* 渲染层断开不致命 */ }
    }
  }

  async init(): Promise<void> {
    const cfg = this.workspace.getFullConfig();
    const agentDir = join(this.workspace.baseDir, 'agent');
    mkdirSync(agentDir, { recursive: true });

    let sdk: any;
    try {
      sdk = await import('@earendil-works/pi-coding-agent');
    } catch (err) {
      this.statusInfo = { ready: false, reason: 'pi SDK 未安装: ' + (err as Error).message };
      return;
    }

    const usingGateway = !!cfg.baseUrl;
    const providerId = usingGateway ? 'paiban-gateway' : (cfg.provider || 'deepseek');

    // OpenAI 兼容端点：写 models.json（apiKey 经环境变量插值，不落明文到 models.json）
    let modelsPath: string | null = null;
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
        const available = modelRuntime.getModels(providerId).map((m: { id: string }) => m.id).slice(0, 10);
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
        tools: TOOL_WHITELIST, // 白名单：只自研工具
        customTools: createTools(this.workspace),
      });
      this.session = session;
      this._wireEvents(session);
      this.ready = true;
      this.statusInfo = { ready: true, provider: providerId, model: cfg.model, gateway: usingGateway };
    } catch (err) {
      this.statusInfo = { ready: false, reason: (err as Error).message, provider: providerId, model: cfg.model };
    }
  }

  _wireEvents(session: any): void {
    session.subscribe((event: any) => {
      const e = event;
      if (e.type === 'message_update' && e.assistantMessageEvent?.type === 'text_delta') {
        this._emit({ type: 'text_delta', delta: e.assistantMessageEvent.delta });
      } else if (e.type === 'tool_execution_start') {
        this._emit({ type: 'tool_start', name: e.toolName, args: summarizeArgs(e.args) });
      } else if (e.type === 'tool_execution_end') {
        this._emit({ type: 'tool_end', name: e.toolName, isError: !!e.isError, summary: summarizeToolResult(e), details: e.result?.details ?? e.details });
      } else if (e.type === 'agent_end') {
        this._emit({ type: 'done' });
      } else if (e.type === 'error' || e.type === 'agent_error') {
        this._emit({ type: 'error', message: e.error?.message || String(e.error || e) });
      }
    });
  }

  status(): AgentStatus {
    return this.statusInfo;
  }

  async abort(): Promise<void> {
    await this.session?.abort();
  }

  /**
   * 发送用户消息。docId 提供时注入文档上下文（首次注入结构摘要）。
   * @returns 发送是否成功（失败时含原因）
   */
  async send(docId: string | undefined, message: string): Promise<{ ok: boolean; error?: string }> {
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
      this._emit({ type: 'error', message: (err as Error).message });
      return { ok: false, error: (err as Error).message };
    }
  }
}

// ---- 事件摘要 ----

export function summarizeArgs(args: any): string {
  if (!args || typeof args !== 'object') return '';
  if (args.commands) return `${args.commands.length} 条命令${args.note ? '：' + args.note : ''}`;
  if (args.action) return args.action + (args.versionId ? ' ' + args.versionId : '');
  if (args.markdown) {
    // doc_generate：markdown 可能很长，只显示行数与文档名/规则集
    const mdLines = typeof args.markdown === 'string' ? args.markdown.trim().split(/\r?\n/).length : 0;
    return `生成《${args.name || '新文档'}》（规则集 ${args.rulesetId ?? '?'}${mdLines ? `，markdown ${mdLines} 行` : ''}）`;
  }
  if (args.templateId) return args.templateId;
  if (args.amount !== undefined) return `金额 ${args.amount}`;
  return Object.keys(args).join(', ');
}

export function summarizeToolResult(e: any): string {
  const d = e.result?.details ?? e.details;
  if (!d || typeof d !== 'object') return '';
  if (d.docId && typeof d.name === 'string') return `《${d.name}》${d.version?.id ? ' · ' + d.version.id : ''}`; // doc_generate：文档名 + v1
  if (d.version) return `${d.version.id}${d.versionCreated ? '（新版本）' : ''}，applied ${d.applied?.length ?? 0}，errors ${d.errors?.length ?? 0}`;
  if (d.versions) return `${d.versions.length} 个版本`;
  if (d.templates) return `${d.templates.length} 个模板`;
  if (d.paragraphCount !== undefined) return `${d.paragraphCount} 段`;
  return '';
}
