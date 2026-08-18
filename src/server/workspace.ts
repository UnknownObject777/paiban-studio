// server/workspace.ts — 工作台服务层（主进程业务核心，headless 可测）。
//
// 汇聚 docx 编辑内核 / 存储版本链 / 模板库，向 IPC 层与 agent 工具层提供统一接口：
//   文档：uploadDocument / generateDocument / getDocumentBuffer / applyCommands（编辑 + 自动快照）/ outline / download
//   版本：listVersions / getVersionBuffer / rollback
//   模板：uploadTemplate / listTemplates / readTemplate / instantiateTemplate / templateRulesetCommands
//   内置规则集：listBuiltinRulesets / builtinRulesetCommands（templates/rulesets/ 手写资产，#29）
//   配置：LLM provider（界面配置 > 环境变量 > 默认值，R4）
//
// 安全约定（user story 16）：只处理上传副本（buffer 入库），原稿路径零改动。

import { join } from 'node:path';
import { mkdirSync, readFileSync, writeFileSync, existsSync, renameSync, rmSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { LocalFsObjectStore } from '../storage/objectStore.js';
import { VersionStore } from '../storage/versionStore.js';
import { TemplateStore } from '../templates/templateStore.js';
import { applyEdits } from '../docx-core/applyEdits.js';
import { dumpOutline } from '../docx-core/outline.js';
import { generateFromMarkdown } from '../docgen/generate.js';
import { loadRuleset } from '../ruleset/load.js';
import { rulesetToCommands } from '../templates/rulesetToCommands.js';
import type { EditCommand } from '../docx-core/applyEdits.js';
import type { VersionEntry } from '../storage/versionStore.js';

// 内置规则集目录：相对本模块定位项目根 templates/rulesets/。
// 源码位于 src/server/（上两级）；编译产物位于 dist/src/server/（上三级）。
// （与 templates/rulesetFromSample.ts 同款双路径 fallback）
const HERE = fileURLToPath(new URL('.', import.meta.url));
const BUILTIN_RULESETS_DIR = [
  join(HERE, '../../templates/rulesets'),
  join(HERE, '../../../templates/rulesets'),
].find((p) => existsSync(p)) || join(HERE, '../../templates/rulesets');

const DEFAULT_CONFIG = {
  provider: 'deepseek',
  model: 'deepseek-v4-flash',
  baseUrl: '',
  apiKey: '',
};

export interface PublicConfig {
  provider: string;
  model: string;
  baseUrl: string;
  hasApiKey: boolean;
  [key: string]: unknown;
}

export interface FullConfig {
  provider: string;
  model: string;
  baseUrl: string;
  apiKey: string;
  hasApiKey?: boolean;
  [key: string]: unknown;
}

export class Workspace {
  baseDir: string;
  objects: LocalFsObjectStore;
  versions: VersionStore;
  templates: TemplateStore;
  configPath: string;

  constructor(baseDir: string) {
    this.baseDir = baseDir;
    mkdirSync(baseDir, { recursive: true });
    this.objects = new LocalFsObjectStore(baseDir);
    this.versions = new VersionStore(baseDir, this.objects);
    this.templates = new TemplateStore(baseDir, this.objects, this.versions);
    this.configPath = join(baseDir, 'config.json');
  }

  // ---- 文档 ----

  /** 上传 docx（bytes + 文件名）→ 新工作文档 v1。 */
  uploadDocument(buffer: Buffer | ArrayBuffer | Uint8Array, name?: string): { docId: string; version: VersionEntry; name: string } {
    const { docId, version } = this.versions.createDocument(buffer, {
      name: name || '未命名.docx', origin: 'upload', note: '上传文档（原稿零改动，仅处理副本）',
    });
    return { docId, version, name: name || '未命名.docx' };
  }

  listDocuments() {
    return this.versions.listDocuments();
  }

  getDocumentBuffer(docId: string, versionId?: string): Buffer {
    return this.versions.getBuffer(docId, versionId);
  }

  getOutline(docId: string, opts?: { textPreview?: number }) {
    return dumpOutline(this.versions.getBuffer(docId), opts);
  }

  /**
   * 应用编辑命令（agent 工具与手动操作统一入口）：
   * applyEdits → 内容有变化才自动快照（幂等）→ 返回编辑结果 + 版本信息。
   */
  applyCommands(
    docId: string,
    commands: EditCommand[],
    { source = 'edit', note = '' }: { source?: string; note?: string } = {},
  ) {
    const before = this.versions.getBuffer(docId);
    const { buffer, result } = applyEdits(before, commands);
    const snap = this.versions.snapshot(docId, buffer, {
      source,
      note: note || summarizeCommands(commands),
    });
    return { ...result, version: snap.version, versionCreated: snap.created };
  }

  // ---- 版本 ----

  listVersions(docId: string): VersionEntry[] {
    return this.versions.list(docId);
  }

  rollback(docId: string, versionId: string) {
    return this.versions.rollback(docId, versionId);
  }

  // ---- 模板 ----

  uploadTemplate(buffer: Buffer | ArrayBuffer | Uint8Array, name?: string) {
    return this.templates.uploadTemplate(buffer, { name });
  }

  listTemplates() {
    return this.templates.listTemplates();
  }

  readTemplate(templateId: string) {
    return this.templates.readTemplate(templateId);
  }

  instantiateTemplate(templateId: string, values: Record<string, unknown>, name?: string) {
    return this.templates.instantiate(templateId, values, { name });
  }

  /** 模板规则集 → 内核命令（"按《通知》模板排"）。 */
  templateRulesetCommands(templateId: string) {
    return this.templates.rulesetCommands(templateId);
  }

  // ---- 内置规则集（templates/rulesets/ 手写资产，反推链路的保底路径，issue #29） ----

  /** 校验 rulesetId 并定位其目录（列表提示逻辑与 builtinRulesetCommands 共用）。 */
  private resolveBuiltinRulesetDir(rulesetId: string): string {
    if (!/^[\w-]+$/.test(rulesetId)) throw new Error(`非法规则集 id：${rulesetId}`);
    const dir = join(BUILTIN_RULESETS_DIR, rulesetId);
    if (!existsSync(dir)) {
      const known = this.listBuiltinRulesets().map((r) => r.id).join(', ') || '（无）';
      throw new Error(`内置规则集不存在：${rulesetId}；可用：${known}`);
    }
    return dir;
  }

  /** 列出内置规则集：id + 描述（取 styles.json 的 description 字段）。 */
  listBuiltinRulesets(): Array<{ id: string; description: string }> {
    if (!existsSync(BUILTIN_RULESETS_DIR)) return [];
    return readdirSync(BUILTIN_RULESETS_DIR, { withFileTypes: true })
      .filter((d) => d.isDirectory() && existsSync(join(BUILTIN_RULESETS_DIR, d.name, 'styles.json')))
      .map((d) => {
        let description = '';
        try {
          description = (JSON.parse(readFileSync(join(BUILTIN_RULESETS_DIR, d.name, 'styles.json'), 'utf8')).description as string) || '';
        } catch { /* 描述缺失不阻塞列出 */ }
        return { id: d.name, description };
      });
  }

  /** 内置规则集 → 内核命令（loadRuleset 校验 + rulesetToCommands 翻译，"按实验报告排版"）。 */
  builtinRulesetCommands(rulesetId: string): EditCommand[] {
    const dir = this.resolveBuiltinRulesetDir(rulesetId);
    const { recognizers, styles } = loadRuleset(dir);
    return rulesetToCommands(recognizers, styles);
  }

  /** markdown + 内置规则集 → 新工作文档（"一句话生成标书/实验报告"的地基）。 */
  generateDocument(markdown: string, rulesetId: string, name?: string): { docId: string; version: VersionEntry; name: string } {
    const dir = this.resolveBuiltinRulesetDir(rulesetId);
    const { recognizers, styles } = loadRuleset(dir);
    const buffer = generateFromMarkdown(markdown, { recognizers, styles });
    const docName = name || '生成文档.docx';
    const { docId, version } = this.versions.createDocument(buffer, {
      name: docName, origin: 'generate', note: `由 markdown 生成（规则集 ${rulesetId}）`,
    });
    return { docId, version, name: docName };
  }

  // ---- 配置（R4：界面配置 > 环境变量 > 默认值） ----

  getConfig(): PublicConfig {
    let fileCfg: Record<string, unknown> = {};
    if (existsSync(this.configPath)) {
      try { fileCfg = JSON.parse(readFileSync(this.configPath, 'utf8')); } catch { /* 忽略坏配置 */ }
    }
    const envCfg: Record<string, unknown> = {
      provider: process.env.PAIBAN_PROVIDER,
      model: process.env.PAIBAN_MODEL,
      baseUrl: process.env.PAIBAN_BASE_URL,
      apiKey: process.env.PAIBAN_API_KEY || process.env.ANTHROPIC_API_KEY || process.env.DEEPSEEK_API_KEY || process.env.OPENAI_API_KEY,
    };
    const cfg: Record<string, unknown> = { ...DEFAULT_CONFIG };
    for (const src of [envCfg, fileCfg]) {
      for (const [k, v] of Object.entries(src)) if (v) cfg[k] = v;
    }
    // 凭证不回传渲染层
    const { apiKey, ...pub } = cfg;
    return { ...pub, hasApiKey: !!apiKey } as PublicConfig;
  }

  /** 供主进程/agent 层使用的完整配置（含凭证，不出主进程）。 */
  getFullConfig(): FullConfig {
    const pub = this.getConfig();
    let fileCfg: Record<string, unknown> = {};
    if (existsSync(this.configPath)) {
      try { fileCfg = JSON.parse(readFileSync(this.configPath, 'utf8')); } catch { /* ignore */ }
    }
    const apiKey = (fileCfg.apiKey as string) || process.env.PAIBAN_API_KEY
      || process.env.ANTHROPIC_API_KEY || process.env.DEEPSEEK_API_KEY || process.env.OPENAI_API_KEY || '';
    return { ...pub, apiKey };
  }

  setConfig(patch: Record<string, unknown>): PublicConfig {
    let fileCfg: Record<string, unknown> = {};
    if (existsSync(this.configPath)) {
      try { fileCfg = JSON.parse(readFileSync(this.configPath, 'utf8')); } catch { /* ignore */ }
    }
    const next = { ...fileCfg };
    for (const [k, v] of Object.entries(patch)) {
      if (v === undefined || v === null || v === '') delete next[k];
      else next[k] = v;
    }
    const tmp = this.configPath + '.tmp-' + process.pid;
    writeFileSync(tmp, JSON.stringify(next, null, 2));
    try { renameSync(tmp, this.configPath); } catch { rmSync(this.configPath, { force: true }); renameSync(tmp, this.configPath); }
    return this.getConfig();
  }
}

// 命令数组 → 人类可读摘要（版本 note / 对话层"它改了什么"）
export function summarizeCommands(commands: EditCommand[]): string {
  const parts: string[] = [];
  for (const c of commands.slice(0, 3)) {
    const where = c.path || (c.match ? `match:${c.match.text}` : c.parent || '');
    parts.push(`${c.command}${where ? ' ' + where : ''}`);
  }
  if (commands.length > 3) parts.push(`…共 ${commands.length} 条`);
  return parts.join('；') || '空编辑';
}
