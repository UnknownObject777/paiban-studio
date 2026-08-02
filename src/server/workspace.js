// server/workspace.js — 工作台服务层（主进程业务核心，headless 可测）。
//
// 汇聚 docx 编辑内核 / 存储版本链 / 模板库，向 IPC 层与 agent 工具层提供统一接口：
//   文档：uploadDocument / getDocumentBuffer / applyCommands（编辑 + 自动快照）/ outline / download
//   版本：listVersions / getVersionBuffer / rollback
//   模板：uploadTemplate / listTemplates / readTemplate / instantiateTemplate / templateRulesetCommands
//   配置：LLM provider（界面配置 > 环境变量 > 默认值，R4）
//
// 安全约定（user story 16）：只处理上传副本（buffer 入库），原稿路径零改动。

import { join } from 'node:path';
import { mkdirSync, readFileSync, writeFileSync, existsSync, renameSync, rmSync } from 'node:fs';
import { LocalFsObjectStore } from '../storage/objectStore.js';
import { VersionStore } from '../storage/versionStore.js';
import { TemplateStore } from '../templates/templateStore.js';
import { applyEdits } from '../docx-core/applyEdits.js';
import { dumpOutline } from '../docx-core/outline.js';

const DEFAULT_CONFIG = {
  provider: 'anthropic',
  model: 'claude-sonnet-4-5',
  baseUrl: '',
  apiKey: '',
};

export class Workspace {
  constructor(baseDir) {
    this.baseDir = baseDir;
    mkdirSync(baseDir, { recursive: true });
    this.objects = new LocalFsObjectStore(baseDir);
    this.versions = new VersionStore(baseDir, this.objects);
    this.templates = new TemplateStore(baseDir, this.objects, this.versions);
    this.configPath = join(baseDir, 'config.json');
  }

  // ---- 文档 ----

  /** 上传 docx（bytes + 文件名）→ 新工作文档 v1。 */
  uploadDocument(buffer, name) {
    const { docId, version } = this.versions.createDocument(buffer, {
      name: name || '未命名.docx', origin: 'upload', note: '上传文档（原稿零改动，仅处理副本）',
    });
    return { docId, version, name: name || '未命名.docx' };
  }

  listDocuments() {
    return this.versions.listDocuments();
  }

  getDocumentBuffer(docId, versionId) {
    return this.versions.getBuffer(docId, versionId);
  }

  getOutline(docId, opts) {
    return dumpOutline(this.versions.getBuffer(docId), opts);
  }

  /**
   * 应用编辑命令（agent 工具与手动操作统一入口）：
   * applyEdits → 内容有变化才自动快照（幂等）→ 返回编辑结果 + 版本信息。
   */
  applyCommands(docId, commands, { source = 'edit', note = '' } = {}) {
    const before = this.versions.getBuffer(docId);
    const { buffer, result } = applyEdits(before, commands);
    const snap = this.versions.snapshot(docId, buffer, {
      source,
      note: note || summarizeCommands(commands),
    });
    return { ...result, version: snap.version, versionCreated: snap.created };
  }

  // ---- 版本 ----

  listVersions(docId) {
    return this.versions.list(docId);
  }

  rollback(docId, versionId) {
    return this.versions.rollback(docId, versionId);
  }

  // ---- 模板 ----

  uploadTemplate(buffer, name) {
    return this.templates.uploadTemplate(buffer, { name });
  }

  listTemplates() {
    return this.templates.listTemplates();
  }

  readTemplate(templateId) {
    return this.templates.readTemplate(templateId);
  }

  instantiateTemplate(templateId, values, name) {
    return this.templates.instantiate(templateId, values, { name });
  }

  /** 模板规则集 → 内核命令（"按《通知》模板排"）。 */
  templateRulesetCommands(templateId) {
    return this.templates.rulesetCommands(templateId);
  }

  // ---- 配置（R4：界面配置 > 环境变量 > 默认值） ----

  getConfig() {
    let fileCfg = {};
    if (existsSync(this.configPath)) {
      try { fileCfg = JSON.parse(readFileSync(this.configPath, 'utf8')); } catch { /* 忽略坏配置 */ }
    }
    const envCfg = {
      provider: process.env.PAIBAN_PROVIDER,
      model: process.env.PAIBAN_MODEL,
      baseUrl: process.env.PAIBAN_BASE_URL,
      apiKey: process.env.PAIBAN_API_KEY || process.env.ANTHROPIC_API_KEY || process.env.OPENAI_API_KEY,
    };
    const cfg = { ...DEFAULT_CONFIG };
    for (const src of [envCfg, fileCfg]) {
      for (const [k, v] of Object.entries(src)) if (v) cfg[k] = v;
    }
    // 凭证不回传渲染层
    const { apiKey, ...pub } = cfg;
    return { ...pub, hasApiKey: !!apiKey };
  }

  /** 供主进程/agent 层使用的完整配置（含凭证，不出主进程）。 */
  getFullConfig() {
    const pub = this.getConfig();
    let fileCfg = {};
    if (existsSync(this.configPath)) {
      try { fileCfg = JSON.parse(readFileSync(this.configPath, 'utf8')); } catch { /* ignore */ }
    }
    const apiKey = fileCfg.apiKey || process.env.PAIBAN_API_KEY
      || process.env.ANTHROPIC_API_KEY || process.env.OPENAI_API_KEY || '';
    return { ...pub, apiKey };
  }

  setConfig(patch) {
    let fileCfg = {};
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
export function summarizeCommands(commands) {
  const parts = [];
  for (const c of commands.slice(0, 3)) {
    const where = c.path || (c.match ? `match:${c.match.text}` : c.parent || '');
    parts.push(`${c.command}${where ? ' ' + where : ''}`);
  }
  if (commands.length > 3) parts.push(`…共 ${commands.length} 条`);
  return parts.join('；') || '空编辑';
}
