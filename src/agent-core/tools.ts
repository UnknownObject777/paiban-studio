// agent-core/tools.ts — 自研 agent 工具（spec 模块 2）。
//
// 工具（+ 结构 dump 工具）全部收敛到 Workspace 服务层：
//   doc_outline    文档结构 dump（dump → batch 往返的 dump 端，D6）
//   doc_edit       编辑命令批（走唯一 seam applyEdits + 自动快照）
//   doc_generate   从零生成新文档（markdown + 内置规则集 → 规范排版 docx）
//   template_read  模板规则集 / 占位符 / 大纲读取
//   ruleset_read   内置规则集（手写资产）列表 / 规则集命令读取（#29）
//   version_store  版本列表 / 回滚
//   amount_words   人民币金额大写换算（确定性纯函数，替代 LLM 心算）
//
// 参数 schema 用纯 JSON Schema 对象（与 typebox 产出的 TSchema 运行时同构）。
// 所有写工具 executionMode: 'sequential'（R1：串行执行，防并发写同一文档）；
// amount_words 虽是纯函数，为与其余工具保持一致同样标记 sequential。
// 内置文件工具（read/bash/edit/write）不入白名单——agent 只能经这些工具改文档。

import type { Workspace } from '../server/workspace.js';
import { amountToWordsCn } from '../docgen/amountWords.js';

/** 工具 schema（纯 JSON Schema 对象，SDK 运行时接受）。 */
export interface AgentToolParameters {
  type?: string;
  properties?: Record<string, unknown>;
  required?: string[];
  description?: string;
  items?: Record<string, unknown>;
  enum?: unknown[];
  additionalProperties?: boolean;
}

/** 本地工具接口（与 SDK ToolDefinition 结构对齐，schema 部分用纯 JSON）。 */
export interface AgentTool {
  name: string;
  label: string;
  description: string;
  promptSnippet?: string;
  promptGuidelines?: string[];
  parameters: AgentToolParameters;
  executionMode?: string;
  execute(id: string, params: Record<string, any>): Promise<AgentToolResult>;
}

export interface AgentToolResult {
  content: Array<{ type: string; text: string }>;
  details: unknown;
  isError: boolean;
}

const COMMANDS_SCHEMA: AgentToolParameters = {
  type: 'array',
  description: '编辑命令数组，按顺序原子应用。协议：{command:"set",path,props}|{command:"set",match:{text},props}|{command:"add",parent,node,position}|{command:"remove",path}|{command:"move",path,parent}|{command:"findReplace",find,replace}|{command:"normalize",ruleset}|{command:"numbering",action,...}|{command:"pageNumber",action:"footer"}',
  items: { type: 'object', additionalProperties: true },
};

export function createTools(workspace: Workspace): AgentTool[] {
  return [
    {
      name: 'doc_outline',
      label: 'Document Outline',
      description: '获取 docx 文档结构大纲：段落路径（/body/p[N]）、文本预览、样式/大纲级/编号标记、节页面设置。编辑前先调用它定位目标路径。',
      promptSnippet: 'Dump the structure of a Word document (paragraph paths for addressing)',
      promptGuidelines: [
        '编辑文档前先调用 doc_outline 获取段落路径，再用 doc_edit 按路径精确修改。',
        '路径形如 /body/p[1]/r[2]（1 起，按同标签兄弟计数）。',
      ],
      parameters: {
        type: 'object',
        properties: {
          docId: { type: 'string', description: '工作文档 ID' },
          textPreview: { type: 'number', description: '段落文本预览长度，默认 60' },
        },
        required: ['docId'],
      },
      executionMode: 'sequential',
      async execute(_id: string, params: Record<string, any>) {
        const outline = workspace.getOutline(params.docId, { textPreview: params.textPreview });
        return jsonResult(outline);
      },
    },
    {
      name: 'doc_edit',
      label: 'Document Edit',
      description: '对 docx 应用一批编辑命令（唯一编辑通道）。每次成功编辑自动生成新版本，可回滚。命令按数组顺序应用；失败命令返回结构化错误（含自愈建议），不中断后续命令。',
      promptSnippet: 'Edit Word documents via batch commands (the ONLY write path)',
      promptGuidelines: [
        '禁止直接读写 .docx 文件；一切文档修改必须走 doc_edit。',
        'set 命令：段落 /body/p[N]（props: align/lineSpacingPt/firstLineChars/spacingBeforePt/pageBreakBefore/outlineLevel）；run /body/p[N]/r[M]（props: eastAsia/ascii/sizePt/bold/italic/underline/color）；节 /body/sectPr（props: marginsCm/pageSize/orientation/pageNumFmt）。',
        '批量统一格式用 {command:"set",match:{text:"正则"},props:{run:{...}}} 命中全部匹配段落。',
        '全文规范化用 normalize + 模板规则集（template_read 可获取规则集并直接复用其命令）。',
        '中文公文字号：三号=16pt 小三=15pt 四号=14pt 小四=12pt 五号=10.5pt；正文行距常用 28 磅（lineSpacingPt:28）；首行缩进 2 字符（firstLineChars:200）。',
      ],
      parameters: {
        type: 'object',
        properties: {
          docId: { type: 'string' },
          commands: COMMANDS_SCHEMA,
          note: { type: 'string', description: '本次修改的人类可读摘要（记入版本链）' },
        },
        required: ['docId', 'commands'],
      },
      executionMode: 'sequential',
      async execute(_id: string, params: Record<string, any>) {
        if (!Array.isArray(params.commands) || !params.commands.length) {
          return jsonResult({ error: 'commands 必须是非空数组' }, true);
        }
        const r = workspace.applyCommands(params.docId, params.commands, {
          source: 'agent', note: params.note,
        });
        return jsonResult({
          applied: r.applied,
          errors: r.errors,
          version: r.version,
          versionCreated: r.versionCreated,
          selfCheck: r.selfCheck,
        }, r.errors.length > 0);
      },
    },
    {
      name: 'doc_generate',
      label: 'Document Generate',
      description: '从零生成新 docx 文档：markdown（# 题目 / ## 章节 / GFM 表格）+ 内置排版规则集 → 规范排版的 Word 文档，自动入库为新工作文档（v1）。用于"一句话生成标书/实验报告"等从零开始的新文档场景。',
      promptSnippet: 'Generate a new Word document from markdown with a builtin ruleset',
      promptGuidelines: [
        'doc_generate 用于从零生成新文档；修改/排版已有文档用 doc_edit。',
        'rulesetId 取值先用 ruleset_read 列出内置规则集（如 lab-report-default / gongwen-default / bid-default / fx-form-default）。',
        'markdown 约定：第一个 # 是文档题目，## 是章节标题；报价表/申报字段表用 GFM 表格（| 表头 | ... | 加分隔行 | --- |）；支持 **粗体** 等行内格式。',
        '生成后可用 doc_outline 查看产物结构，再按需用 doc_edit 微调。',
      ],
      parameters: {
        type: 'object',
        properties: {
          markdown: { type: 'string', description: 'markdown 文本：第一个 # 为文档题目、## 为章节标题、GFM 表格写报价/字段表，支持 **粗体**/*斜体*/`代码` 行内格式' },
          rulesetId: { type: 'string', description: '内置排版规则集 ID（先用 ruleset_read 列出，如 lab-report-default / gongwen-default / bid-default / fx-form-default）' },
          name: { type: 'string', description: '生成文档的文件名（省略时默认"生成文档.docx"）' },
        },
        required: ['markdown', 'rulesetId'],
      },
      executionMode: 'sequential',
      async execute(_id: string, params: Record<string, any>) {
        if (typeof params.markdown !== 'string' || !params.markdown.trim()) {
          return jsonResult({ error: 'markdown 必须是非空字符串' }, true);
        }
        try {
          return jsonResult(workspace.generateDocument(params.markdown, params.rulesetId, params.name));
        } catch (err) {
          return jsonResult({ error: (err as Error).message }, true);
        }
      },
    },
    {
      name: 'template_read',
      label: 'Template Read',
      description: '读取模板资产：不传 templateId 时列出全部模板；传入时返回该模板的规则集（识别/样式）、占位符清单、结构大纲，以及可直接用于 doc_edit 的规则集命令（rulesetCommands）。',
      promptSnippet: 'Read formatting templates and their rulesets',
      promptGuidelines: [
        '用户说"按某模板排"时：template_read 取 rulesetCommands，原样传给 doc_edit 的 commands。',
      ],
      parameters: {
        type: 'object',
        properties: {
          templateId: { type: 'string', description: '模板 ID；省略则列出全部模板' },
        },
      },
      executionMode: 'sequential',
      async execute(_id: string, params: Record<string, any>) {
        if (!params.templateId) {
          return jsonResult({ templates: workspace.listTemplates() });
        }
        const t = workspace.readTemplate(params.templateId);
        return jsonResult({
          meta: t.meta,
          placeholders: t.placeholders,
          outline: t.outline,
          recognizers: t.recognizers,
          styles: t.styles,
          rulesetCommands: workspace.templateRulesetCommands(params.templateId),
        });
      },
    },
    {
      name: 'ruleset_read',
      label: 'Builtin Ruleset Read',
      description: '内置排版规则集（手写资产，质量保底）：不传 rulesetId 时列出全部内置规则集（id + 描述）；传入时返回该规则集可直接用于 doc_edit 的规则集命令（rulesetCommands）。',
      promptSnippet: 'Read builtin formatting rulesets (hand-written, preferred over inferred ones)',
      promptGuidelines: [
        '用户说"按实验报告排版"→ ruleset_read 取 lab-report-default；"按公文排版"→ gongwen-default；"按标书排版"→ bid-default；"按外汇申报表单排版"→ fx-form-default。优先用内置规则集，质量比上传模板反推的更可靠。',
        '用法：ruleset_read 取 rulesetCommands，原样传给 doc_edit 的 commands。',
      ],
      parameters: {
        type: 'object',
        properties: {
          rulesetId: { type: 'string', description: '内置规则集 ID（如 lab-report-default / gongwen-default / bid-default / fx-form-default）；省略则列出全部' },
        },
      },
      executionMode: 'sequential',
      async execute(_id: string, params: Record<string, any>) {
        if (!params.rulesetId) {
          return jsonResult({ rulesets: workspace.listBuiltinRulesets() });
        }
        try {
          return jsonResult({
            rulesetId: params.rulesetId,
            rulesetCommands: workspace.builtinRulesetCommands(params.rulesetId),
          });
        } catch (err) {
          return jsonResult({ error: (err as Error).message }, true);
        }
      },
    },
    {
      name: 'version_store',
      label: 'Version Store',
      description: '文档版本链：列出历史版本（含每版摘要），或回滚到指定版本（回滚本身记录为新版本）。',
      promptSnippet: 'List and roll back document versions',
      parameters: {
        type: 'object',
        properties: {
          action: { type: 'string', enum: ['list', 'rollback'], description: 'list=列版本；rollback=回滚' },
          docId: { type: 'string' },
          versionId: { type: 'string', description: 'rollback 目标版本（如 v3）' },
        },
        required: ['action', 'docId'],
      },
      executionMode: 'sequential',
      async execute(_id: string, params: Record<string, any>) {
        if (params.action === 'list') {
          return jsonResult({ versions: workspace.listVersions(params.docId) });
        }
        if (params.action === 'rollback') {
          if (!params.versionId) return jsonResult({ error: 'rollback 需要 versionId' }, true);
          return jsonResult(workspace.rollback(params.docId, params.versionId));
        }
        return jsonResult({ error: `未知 action: ${params.action}` }, true);
      },
    },
    {
      name: 'amount_words',
      label: 'Amount to Chinese Words',
      description: '人民币金额大写换算（确定性）：金额 → 中文大写（如 12345.67 → 壹万贰仟叁佰肆拾伍元陆角柒分）。金额可含千分位逗号（"1,234.5"）、最多两位小数（超出四舍五入到分）；负数/超千亿级/格式非法返回错误。',
      promptSnippet: 'Convert an amount to Chinese uppercase financial words',
      promptGuidelines: [
        '标书报价表、境外汇款申请书等含「金额（大写）」栏时，必须先调 amount_words 换算后填入，禁止心算。',
      ],
      parameters: {
        type: 'object',
        properties: {
          amount: { type: 'string', description: '金额（可含千分位逗号，最多两位小数，如 "12345.67" / "1,234.5"）' },
        },
        required: ['amount'],
      },
      executionMode: 'sequential',
      async execute(_id: string, params: Record<string, any>) {
        const amount = String(params.amount ?? '').trim();
        if (amount === '') return jsonResult({ error: 'amount 不能为空' }, true);
        try {
          const words = amountToWordsCn(amount);
          return jsonResult({ amount, words });
        } catch (err) {
          return jsonResult({ error: (err as Error).message }, true);
        }
      },
    },
  ];
}

function jsonResult(data: unknown, isError = false): AgentToolResult {
  return {
    content: [{ type: 'text', text: JSON.stringify(data, null, 2) }],
    details: data,
    isError,
  };
}
