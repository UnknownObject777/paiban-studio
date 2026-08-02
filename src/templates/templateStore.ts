// templates/templateStore.ts — 模板资产库（spec 模块 4：上传 → 解析 → 规则集 → 实例化）。
//
// 布局：<baseDir>/templates/<templateId>/
//   meta.json          模板元数据（名称/占位符/大纲摘要/反推组件清单/源文档 hash）
//   recognizers.json   识别规则（两文件分离，#5 已拍板）
//   styles.json        样式与页面
// 模板源 docx 存 ObjectStore（内容寻址）。
//
// 实例化：模板 buffer →（可选）占位符 findReplace 合并 → VersionStore 新建工作文档。

import { mkdirSync, readFileSync, writeFileSync, existsSync, readdirSync, renameSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { randomUUID } from 'node:crypto';
import { extractPlaceholders, placeholderCommands } from './placeholders.js';
import { extractRulesetFromSample } from './rulesetFromSample.js';
import { rulesetToCommands } from './rulesetToCommands.js';
import { dumpOutline } from '../docx-core/outline.js';
import { applyEdits } from '../docx-core/applyEdits.js';
import { validateRuleset } from '../ruleset/schema.js';
import type { LocalFsObjectStore } from '../storage/objectStore.js';
import type { VersionStore, VersionEntry } from '../storage/versionStore.js';
import type { EditCommand } from '../docx-core/applyEdits.js';

export interface TemplateMeta {
  templateId: string;
  name: string;
  hash: string;
  createdAt: string;
  placeholderCount: number;
  paragraphCount: number;
  extractedComponents: string[];
  [key: string]: unknown;
}

export class TemplateStore {
  store: LocalFsObjectStore;
  versions: VersionStore;
  templatesDir: string;

  constructor(baseDir: string, objectStore: LocalFsObjectStore, versionStore: VersionStore) {
    this.store = objectStore;
    this.versions = versionStore;
    this.templatesDir = join(baseDir, 'templates');
    mkdirSync(this.templatesDir, { recursive: true });
  }

  _dir(id: string): string {
    return join(this.templatesDir, id);
  }

  _writeJson(dir: string, name: string, data: unknown): void {
    const p = join(dir, name);
    const tmp = p + '.tmp-' + process.pid;
    writeFileSync(tmp, JSON.stringify(data, null, 2));
    try { renameSync(tmp, p); } catch { rmSync(p, { force: true }); renameSync(tmp, p); }
  }

  /**
   * 上传模板：解析占位符 + 结构大纲 + 反推规则集（title/body/page 实测，其余继承默认集）。
   * @returns 模板 ID + meta + 占位符清单 + 实测组件清单
   */
  uploadTemplate(
    buffer: Buffer | ArrayBuffer | Uint8Array,
    { name, rulesetName }: { name?: string; rulesetName?: string } = {},
  ): { templateId: string; meta: TemplateMeta; placeholders: unknown[]; extracted: string[] } {
    const templateId = randomUUID().slice(0, 8);
    const hash = this.store.put(buffer);
    const placeholders = extractPlaceholders(buffer);
    const { recognizers, styles, extracted } = extractRulesetFromSample(buffer, {
      name: rulesetName || `tpl-${templateId}`,
    });
    // 两文件一致性校验（组件键集必须与 components.js 完全一致）
    const errors = validateRuleset(recognizers, styles);
    if (errors.length) {
      throw new Error(`反推规则集未通过校验：\n- ${errors.join('\n- ')}`);
    }
    const outline = dumpOutline(buffer, { textPreview: 40 });
    const meta: TemplateMeta = {
      templateId,
      name: name || `模板-${templateId}`,
      hash,
      createdAt: new Date().toISOString(),
      placeholderCount: placeholders.length,
      paragraphCount: outline.paragraphCount,
      extractedComponents: extracted,
    };
    const dir = this._dir(templateId);
    mkdirSync(dir, { recursive: true });
    this._writeJson(dir, 'meta.json', meta);
    this._writeJson(dir, 'recognizers.json', recognizers);
    this._writeJson(dir, 'styles.json', styles);
    return { templateId, meta, placeholders, extracted };
  }

  listTemplates(): TemplateMeta[] {
    if (!existsSync(this.templatesDir)) return [];
    return readdirSync(this.templatesDir, { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .map((d) => {
        try {
          return JSON.parse(readFileSync(join(this._dir(d.name), 'meta.json'), 'utf8')) as TemplateMeta;
        } catch {
          return null;
        }
      })
      .filter((x): x is TemplateMeta => x !== null);
  }

  /** 读取模板全量信息（meta + 占位符 + 大纲 + 规则集）。 */
  readTemplate(templateId: string): {
    meta: TemplateMeta;
    recognizers: Record<string, unknown>;
    styles: Record<string, any>;
    placeholders: unknown[];
    outline: unknown;
  } {
    const dir = this._dir(templateId);
    if (!existsSync(dir)) {
      const err = new Error(`模板不存在: ${templateId}`) as Error & { code?: string };
      err.code = 'TEMPLATE_NOT_FOUND';
      throw err;
    }
    const meta = JSON.parse(readFileSync(join(dir, 'meta.json'), 'utf8')) as TemplateMeta;
    const recognizers = JSON.parse(readFileSync(join(dir, 'recognizers.json'), 'utf8')) as Record<string, unknown>;
    const styles = JSON.parse(readFileSync(join(dir, 'styles.json'), 'utf8')) as Record<string, any>;
    const buffer = this.store.get(meta.hash);
    return {
      meta, recognizers, styles,
      placeholders: extractPlaceholders(buffer),
      outline: dumpOutline(buffer, { textPreview: 40 }),
    };
  }

  /** 模板源 docx buffer。 */
  getBuffer(templateId: string): Buffer {
    const meta = JSON.parse(readFileSync(join(this._dir(templateId), 'meta.json'), 'utf8')) as TemplateMeta;
    return this.store.get(meta.hash);
  }

  /**
   * 实例化：复制模板 →（可选）占位符合并 → 生成新工作文档（VersionStore v1）。
   * @returns 新文档 ID + 版本 + 实际替换命令数 + 错误
   */
  instantiate(
    templateId: string,
    values: Record<string, unknown> = {},
    { name }: { name?: string } = {},
  ): { docId: string; version: VersionEntry; replaced: unknown[]; errors: unknown[] } {
    const buffer = this.getBuffer(templateId);
    const commands = placeholderCommands(values);
    let finalBuffer = buffer;
    let applied: unknown[] = [];
    let errors: unknown[] = [];
    if (commands.length) {
      const r = applyEdits(buffer, commands);
      finalBuffer = r.buffer;
      applied = r.result.applied;
      errors = r.result.errors;
    }
    const { docId, version } = this.versions.createDocument(finalBuffer, {
      name: name || `实例化-${templateId}.docx`,
      origin: 'template',
      templateId,
      note: '从模板实例化',
    });
    return { docId, version, replaced: applied, errors };
  }

  /** 模板规则集 → 内核命令（"按《通知》模板的样式排"场景，user story 18）。 */
  rulesetCommands(templateId: string): EditCommand[] {
    const { recognizers, styles } = this.readTemplate(templateId);
    return rulesetToCommands(recognizers, styles);
  }
}
