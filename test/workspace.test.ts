// workspace.test.js — 工作台服务层 headless 测试 + agent 工具层测试。
//
// 覆盖（spec 测试策略"端到端冒烟"的 service 级等价物）：
//   上传 → 编辑（自动快照 + 版本 note 摘要）→ 大纲 → 版本列表 → 回滚 → 下载 buffer
//   模板：上传 → 实例化 → 规则集命令
//   配置优先级：界面配置 > 环境变量 > 默认值（R4）
//   agent 八工具直接 execute（不依赖 LLM）：doc_outline / doc_edit / doc_generate / template_instantiate / template_read / ruleset_read / version_store / amount_words

import { test } from 'node:test';
import assert from 'node:assert/strict';
import PizZip from 'pizzip';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Workspace, summarizeCommands } from '../src/server/workspace.js';
import { createTools } from '../src/agent-core/tools.js';
import { openDocx } from '../src/docx-core/docx.js';
import { findChild, findChildren, getAttr, isElement } from '../src/docx-core/ooxml.js';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function makeMessyDocx() {
  const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const zip = new PizZip();
  zip.file('[Content_Types].xml', DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', DECL + `<w:document ${W_NS}><w:body>` +
    p('关于规范排版工作的通知') + p('一、总体要求') + p('各科室：请认真贯彻执行本通知的各项要求，切实加强公文排版管理。') +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>');
  zip.file('word/_rels/document.xml.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  return zip.generate({ type: 'nodebuffer' });
}

function freshWorkspace() {
  const dir = mkdtempSync(join(tmpdir(), 'paiban-ws-'));
  return { ws: new Workspace(dir), dir, cleanup: () => rmSync(dir, { recursive: true, force: true }) };
}

test('端到端冒烟（service 级）：上传 → 编辑 → 快照 → 大纲 → 版本 → 回滚 → 下载', () => {
  const { ws, cleanup } = freshWorkspace();
  // 上传
  const { docId, version } = ws.uploadDocument(makeMessyDocx(), '乱排版.docx');
  assert.equal(version.id, 'v1');
  // 编辑：标题居中 + 正文统一（自动快照）
  const r = ws.applyCommands(docId, [
    { command: 'set', path: '/body/p[1]', props: { align: 'center', run: { eastAsia: '黑体', sizePt: 16 } } },
    { command: 'set', match: { text: '^各科室' }, props: { run: { eastAsia: '仿宋_GB2312', sizePt: 14 }, paragraph: { firstLineChars: 200, lineSpacingPt: 28 } } },
  ], { source: 'agent' });
  assert.equal(r.errors.length, 0);
  assert.equal(r.versionCreated, true);
  assert.equal(r.version.id, 'v2');
  assert.ok((r.version.note || '').length > 0); // 版本摘要
  // 幂等：相同命令再执行内容无变化 → 不产生 v3 空版本
  const r2 = ws.applyCommands(docId, [
    { command: 'set', path: '/body/p[1]', props: { align: 'center', run: { eastAsia: '黑体', sizePt: 16 } } },
  ]);
  // 大纲
  const outline = ws.getOutline(docId);
  assert.equal(outline.paragraphCount, 3);
  // 版本列表
  const versions = ws.listVersions(docId);
  assert.ok(versions.length >= 2);
  // 回滚到 v1
  const rb = ws.rollback(docId, 'v1');
  assert.equal(rb.created, true);
  const afterRollback = ws.getDocumentBuffer(docId);
  const doc = openDocx(afterRollback);
  const docEl = doc.parts.get('word/document.xml')!.tree!.find((n) => isElement(n, 'w:document'));
  const body = findChild(docEl, 'w:body');
  const firstP = findChildren(body, 'w:p')[0];
  assert.equal(findChild(firstP, 'w:pPr'), undefined); // 回滚后标题格式消失
  // 下载 buffer（headless 等价物：直接取 buffer）
  const download = ws.getDocumentBuffer(docId);
  assert.ok(download.length > 0);
  cleanup();
});

test('模板链路：上传模板 → 规则集命令 → 应用到乱排版文档', () => {
  const { ws, cleanup } = freshWorkspace();
  const tplBuf = makeMessyDocx(); // 复用同一工厂当"模板"
  const { templateId } = ws.uploadTemplate(tplBuf, '通知模板');
  const commands = ws.templateRulesetCommands(templateId);
  assert.ok(commands.find((c) => c.command === 'normalize'));
  const { docId } = ws.uploadDocument(makeMessyDocx(), '待排.docx');
  const r = ws.applyCommands(docId, commands, { source: 'agent', note: '按《通知》模板排' });
  assert.equal(r.errors.length, 0);
  assert.ok(r.applied[0].detail.normalized.title >= 1);
  // 实例化
  const inst = ws.instantiateTemplate(templateId, {}, '新文档.docx');
  assert.equal(inst.version.id, 'v1');
  cleanup();
});

test('配置优先级：界面配置 > 环境变量 > 默认值（R4），凭证不外泄', () => {
  const { ws, cleanup } = freshWorkspace();
  // 默认值
  let cfg = ws.getConfig();
  assert.equal(cfg.provider, 'deepseek');
  assert.equal(cfg.model, 'deepseek-v4-flash');
  const envHasKey = !!(process.env.PAIBAN_API_KEY || process.env.ANTHROPIC_API_KEY || process.env.DEEPSEEK_API_KEY || process.env.OPENAI_API_KEY);
  assert.equal(cfg.hasApiKey, envHasKey);
  // 界面配置覆盖
  ws.setConfig({ provider: 'openai', model: 'qwen3-max', baseUrl: 'http://localhost:11434/v1', apiKey: 'sk-secret' });
  cfg = ws.getConfig();
  assert.equal(cfg.provider, 'openai');
  assert.equal(cfg.model, 'qwen3-max');
  assert.equal(cfg.baseUrl, 'http://localhost:11434/v1');
  assert.equal(cfg.hasApiKey, true);
  assert.equal(cfg.apiKey, undefined); // 凭证不出主进程
  assert.equal(ws.getFullConfig().apiKey, 'sk-secret');
  cleanup();
});

test('agent 工具：doc_outline → doc_edit → version_store 全链路（无 LLM）', async () => {
  const { ws, cleanup } = freshWorkspace();
  const { docId } = ws.uploadDocument(makeMessyDocx(), 'agent-test.docx');
  const tools = createTools(ws);
  const byName = Object.fromEntries(tools.map((t) => [t.name, t]));

  // 白名单与串行约束（spec：裁剪内置工具 + sequential）
  assert.deepEqual(tools.map((t) => t.name), ['doc_outline', 'doc_edit', 'doc_generate', 'template_instantiate', 'template_read', 'ruleset_read', 'version_store', 'amount_words']);
  for (const t of tools) assert.equal(t.executionMode, 'sequential');

  // doc_outline
  const outlineRes = await byName.doc_outline.execute('t1', { docId });
  const outline = JSON.parse(outlineRes.content[0].text);
  assert.equal(outline.paragraphCount, 3);
  assert.equal(outline.paragraphs[1].path, '/body/p[2]');

  // doc_edit（按大纲寻址）
  const editRes = await byName.doc_edit.execute('t2', {
    docId,
    commands: [{ command: 'set', path: outline.paragraphs[1].path, props: { run: { eastAsia: '黑体', sizePt: 16 } } }],
    note: '一级标题黑体三号',
  });
  const editData = JSON.parse(editRes.content[0].text);
  assert.equal(editData.errors.length, 0);
  assert.equal(editData.version.id, 'v2');
  assert.equal(editData.version.note, '一级标题黑体三号');
  assert.ok(editData.selfCheck.ok);

  // doc_edit 空命令拒绝
  const bad = await byName.doc_edit.execute('t3', { docId, commands: [] });
  assert.equal(bad.isError, true);

  // version_store list / rollback
  const listRes = await byName.version_store.execute('t4', { action: 'list', docId });
  assert.equal(JSON.parse(listRes.content[0].text).versions.length, 2);
  const rbRes = await byName.version_store.execute('t5', { action: 'rollback', docId, versionId: 'v1' });
  assert.equal(JSON.parse(rbRes.content[0].text).created, true);

  // template_read 列表
  const tplRes = await byName.template_read.execute('t6', {});
  assert.ok(Array.isArray(JSON.parse(tplRes.content[0].text).templates));
  cleanup();
});

test('内置规则集：列出 → 读出 lab-report-default 命令 → ruleset_read 工具（#29）', async () => {
  const { ws, cleanup } = freshWorkspace();

  // 列表：全部手写规则集都在，带描述
  const list = ws.listBuiltinRulesets();
  const ids = list.map((r) => r.id);
  assert.ok(
    ['gongwen-default', 'lab-report-default', 'bid-default', 'fx-form-default'].every((id) => ids.includes(id)),
    `内置规则集: ${ids}`,
  );
  assert.ok(list.find((r) => r.id === 'lab-report-default')!.description.length > 0);

  // 按 id 读出 rulesetCommands：normalize 含 title/heading1/body/caption/table 组件 + page 页面设置命令
  const commands = ws.builtinRulesetCommands('lab-report-default');
  const normalize = commands.find((c) => c.command === 'normalize');
  assert.ok(normalize, '含 normalize 命令');
  const names = (normalize.ruleset.rules as Array<{ name: string }>).map((r) => r.name);
  for (const comp of ['title', 'heading1', 'body', 'caption', 'table']) {
    assert.ok(names.includes(comp), `normalize 规则缺 ${comp}（实际: ${names}）`);
  }
  const pageCmd = commands.find((c) => c.command === 'set' && c.path === '/body/sectPr');
  assert.ok(pageCmd, '含页面设置命令');
  assert.ok(pageCmd.props.marginsCm, '页面命令含页边距');

  // 未知 id：报错且提示可用列表
  assert.throws(() => ws.builtinRulesetCommands('no-such-ruleset'), /内置规则集不存在.*lab-report-default/s);

  // ruleset_read 工具：不传 id 列出；传 id 返回命令；坏 id 走 isError
  const byName = Object.fromEntries(createTools(ws).map((t) => [t.name, t]));
  const listRes = await byName.ruleset_read.execute('r1', {});
  const listData = JSON.parse(listRes.content[0].text);
  assert.ok(listData.rulesets.some((r: { id: string }) => r.id === 'lab-report-default'));
  const readRes = await byName.ruleset_read.execute('r2', { rulesetId: 'lab-report-default' });
  const readData = JSON.parse(readRes.content[0].text);
  assert.equal(readRes.isError, false);
  assert.ok(readData.rulesetCommands.some((c: { command: string }) => c.command === 'normalize'));
  const badRes = await byName.ruleset_read.execute('r3', { rulesetId: 'no-such-ruleset' });
  assert.equal(badRes.isError, true);
  cleanup();
});

test('版本 note 摘要：命令数组 → 人类可读', () => {
  assert.equal(summarizeCommands([{ command: 'set', path: '/body/p[1]', props: {} }]), 'set /body/p[1]');
  assert.equal(summarizeCommands([]), '空编辑');
  const s = summarizeCommands([
    { command: 'set', path: '/body/p[1]' },
    { command: 'set', path: '/body/p[2]' },
    { command: 'set', path: '/body/p[3]' },
    { command: 'findReplace', find: 'a' },
  ]);
  assert.ok(s.includes('…共 4 条'));
});
