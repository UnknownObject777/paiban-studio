// docgen.test.js — markdown + 内置规则集 → 规范排版 docx 的生成链路测试。
//
// 覆盖（spec 测试策略"端到端冒烟"）：
//   1. 解析器：标题 / 段落合并 / GFM 表格 / 列表 / 行内粗斜体与代码
//   2. 端到端：lab-report-default 规则集生成 → openDocx 重解析 →
//      大纲文本与 outlineLevel、title 黑体、正文首行缩进、表格结构与表头 bold、页面/页码命令
//   3. Workspace.generateDocument 入库为 v1 新工作文档
//   4. 空白 docx 保真：产物再经 applyEdits([]) 纯 round-trip 不自检失败
//   5. doc_generate / template_instantiate 工具 execute（含坏参数/缺模板走 isError）
//   6. docgen 产物含 {{占位符}}：extractPlaceholders 提取 + instantiate 填值替换无残留

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseMarkdown } from '../src/docgen/markdown.js';
import { generateFromMarkdown } from '../src/docgen/generate.js';
import { loadRuleset } from '../src/ruleset/load.js';
import { applyEdits } from '../src/docx-core/applyEdits.js';
import { openDocx } from '../src/docx-core/docx.js';
import { dumpOutline } from '../src/docx-core/outline.js';
import { findChild, findChildren, getAttr, isElement, textOf } from '../src/docx-core/ooxml.js';
import { extractPlaceholders } from '../src/templates/placeholders.js';
import { Workspace } from '../src/server/workspace.js';
import { createTools } from '../src/agent-core/tools.js';
import type { Docx } from '../src/docx-core/docx.js';
import type { XmlNode } from '../src/docx-core/xml.js';

// 源码在 test/；编译产物在 dist/test/（回退路径对齐 lab-report.test.ts 双路径模式）。
const HERE = fileURLToPath(new URL('.', import.meta.url));
const RULESET_DIR = [
  join(HERE, '../templates/rulesets/lab-report-default'),
  join(HERE, '../../templates/rulesets/lab-report-default'),
].find((p) => existsSync(p))!;

function loadLabReport() {
  return loadRuleset(RULESET_DIR);
}

function docBody(docx: Docx): XmlNode {
  const tree = docx.parts.get('word/document.xml')!.tree!;
  const doc = tree.find((n) => isElement(n, 'w:document'))!;
  return findChild(doc, 'w:body')!;
}

function freshWorkspace() {
  const dir = mkdtempSync(join(tmpdir(), 'paiban-docgen-'));
  return { ws: new Workspace(dir), dir, cleanup: () => rmSync(dir, { recursive: true, force: true }) };
}

test('markdown 解析：标题/段落合并/表格/列表/行内粗斜体', () => {
  const blocks = parseMarkdown([
    '# 实验报告标题',
    '',
    '第一段文字。',
    '第二行并入同一段。',
    '',
    '| 列A | 列B |',
    '| --- | --- |',
    '| 甲 | 1 |',
    '| 乙 | 2 |',
    '',
    '- 无序项',
    '3. 有序项',
    '',
    '含 **粗体** 与 *斜体* 与 `code` 的正文',
  ].join('\n'));

  assert.equal(blocks.length, 6);
  // 标题
  assert.deepEqual(blocks[0], { type: 'heading', level: 1, text: '实验报告标题', runs: [{ text: '实验报告标题' }] });
  // 段落：连续两行并入一段（软换行按空格）
  assert.equal(blocks[1].type, 'paragraph');
  assert.equal(blocks[1].runs.length, 1);
  assert.equal(blocks[1].runs[0].text, '第一段文字。 第二行并入同一段。');
  // 表格
  assert.deepEqual(blocks[2], {
    type: 'table',
    header: ['列A', '列B'],
    rows: [['甲', '1'], ['乙', '2']],
  });
  // 列表：无序前缀由生成器加，解析只记标记；有序保留原编号
  assert.deepEqual(blocks[3], { type: 'list-item', ordered: false, index: 1, text: '无序项', runs: [{ text: '无序项' }] });
  assert.deepEqual(blocks[4], { type: 'list-item', ordered: true, index: 3, text: '有序项', runs: [{ text: '有序项' }] });
  // 行内粗斜体与代码
  const last = blocks[5];
  assert.equal(last.type, 'paragraph');
  const texts = last.runs.map((r) => r.text).join('');
  assert.equal(texts, '含 粗体 与 斜体 与 code 的正文');
  assert.ok(last.runs.some((r) => r.props?.bold === true), '缺粗体 run');
  assert.ok(last.runs.some((r) => r.props?.italic === true), '缺斜体 run');
  assert.ok(last.runs.some((r) => r.props?.ascii === 'Courier New'), '缺代码等宽 run');
});

test('端到端：lab-report-default 规则集生成规范排版 docx（标题/章节/正文/表格）', () => {
  const { recognizers, styles } = loadLabReport();
  const md = [
    '# 单摆测重力加速度实验报告',
    '',
    '## 实验目的',
    '',
    '用单摆测量本地重力加速度，**验证**单摆周期公式。',
    '',
    '| 摆长(cm) | 周期(s) |',
    '| --- | --- |',
    '| 50.0 | 1.42 |',
    '| 80.0 | 1.80 |',
  ].join('\n');
  const buffer = generateFromMarkdown(md, { recognizers, styles });

  // 大纲：文本与层级
  const outline = dumpOutline(buffer);
  const texts = outline.paragraphs.map((p) => p.text);
  assert.ok(texts.includes('单摆测重力加速度实验报告'), '缺题目');
  assert.ok(texts.includes('实验目的'), '缺章节标题');
  assert.ok(texts.includes('用单摆测量本地重力加速度，验证单摆周期公式。'), '缺正文');
  assert.ok(texts.includes('摆长(cm)') && texts.includes('50.0'), '缺表格内容');
  const titleEntry = outline.paragraphs.find((p) => p.text === '单摆测重力加速度实验报告');
  const headingEntry = outline.paragraphs.find((p) => p.text === '实验目的');
  assert.equal(titleEntry!.outlineLevel, undefined, 'title 组件不应有大纲级别');
  assert.equal(headingEntry!.outlineLevel, 0, '## → heading1 → OOXML outlineLevel 0');

  // 结构重解析：title 黑体 / 正文首行缩进 2 字符 / 表格
  const doc = openDocx(buffer);
  const body = docBody(doc);
  const ps = findChildren(body, 'w:p');
  const paraByText = (t: string) => ps.find((p) => textOf(p) === t);

  const titleP = paraByText('单摆测重力加速度实验报告');
  const titleRunPr = findChild(findChild(titleP!, 'w:r'), 'w:rPr');
  assert.equal(getAttr(findChild(titleRunPr, 'w:rFonts'), 'w:eastAsia'), '黑体', '题目应黑体');
  assert.equal(getAttr(findChild(titleRunPr, 'w:sz'), 'w:val'), '44', '题目二号 22pt → 半磅 44');

  const bodyP = paraByText('用单摆测量本地重力加速度，验证单摆周期公式。');
  const bodyPPr = findChild(bodyP!, 'w:pPr');
  assert.equal(getAttr(findChild(bodyPPr, 'w:ind'), 'w:firstLineChars'), '200', '正文首行缩进 2 字符');

  // 表格：单线边框 + 表头黑体加粗 + 数据行宋体五号
  const tbl = findChildren(body, 'w:tbl')[0];
  assert.ok(tbl, '表格 w:tbl 存在');
  assert.ok(findChild(findChild(tbl, 'w:tblPr'), 'w:tblBorders'), '表格应有单线边框');
  const trs = findChildren(tbl, 'w:tr');
  assert.equal(trs.length, 3, '表头 + 2 数据行');
  const headerCellPara = findChildren(findChildren(trs[0], 'w:tc')[0], 'w:p')[0];
  const headerRunPr = findChild(findChild(headerCellPara, 'w:r'), 'w:rPr');
  assert.equal(getAttr(findChild(headerRunPr, 'w:b'), 'w:val'), 'true', '表头应加粗');
  assert.equal(getAttr(findChild(headerRunPr, 'w:rFonts'), 'w:eastAsia'), '黑体', '表头应黑体');
  const dataCellPara = findChildren(findChildren(trs[1], 'w:tc')[0], 'w:p')[0];
  const dataRunPr = findChild(findChild(dataCellPara, 'w:r'), 'w:rPr');
  assert.equal(getAttr(findChild(dataRunPr, 'w:rFonts'), 'w:eastAsia'), '宋体', '数据行应宋体');
  assert.equal(getAttr(findChild(dataRunPr, 'w:sz'), 'w:val'), '21', '数据行五号 10.5pt → 半磅 21');

  // 页面 / 页码命令生效：A4 页边距 + 页脚页码部件
  const sect = findChild(body, 'w:sectPr');
  assert.equal(getAttr(findChild(sect, 'w:pgSz'), 'w:w'), '11906', 'A4 宽');
  assert.equal(getAttr(findChild(sect, 'w:pgMar'), 'w:top'), String(Math.round(2.54 * 566.929)), '上边距 2.54cm');
  assert.ok(doc.parts.has('word/footer1.xml'), '页脚页码部件存在');
  assert.ok(JSON.stringify(sect).includes('footerReference'), 'sectPr 引用页脚');
});

test('两个 # 一级标题：首个 → title，后续 → heading1', () => {
  const { recognizers, styles } = loadLabReport();
  const buffer = generateFromMarkdown('# 文档题目\n\n# 第一部分\n\n## 小节', { recognizers, styles });
  const outline = dumpOutline(buffer);
  const entries = outline.paragraphs;
  const first = entries.find((p) => p.text === '文档题目');
  const second = entries.find((p) => p.text === '第一部分');
  const third = entries.find((p) => p.text === '小节');
  assert.equal(first!.outlineLevel, undefined, '首个 # 是 title');
  assert.equal(second!.outlineLevel, 0, '后续 # 是 heading1');
  assert.equal(third!.outlineLevel, 0, '## 是 heading1');
});

test('Workspace：generateDocument 入库为 v1 新工作文档', () => {
  const { ws, cleanup } = freshWorkspace();
  const r = ws.generateDocument('# 新文档\n\n## 第一节\n\n正文内容。', 'lab-report-default', '新报告.docx');
  assert.equal(r.version.id, 'v1');
  assert.equal(r.name, '新报告.docx');
  assert.ok(/^[\w-]{8}$/.test(r.docId));
  // 列表可见 + 版本链 v1
  const docs = ws.listDocuments();
  const mine = docs.find((d) => d.docId === r.docId);
  assert.ok(mine, 'listDocuments 含新文档');
  assert.equal(mine!.name, '新报告.docx');
  assert.equal(mine!.meta.origin, 'generate');
  const versions = ws.listVersions(r.docId);
  assert.equal(versions.length, 1);
  assert.equal(versions[0].id, 'v1');
  assert.equal(versions[0].source, 'generate');
  assert.ok(versions[0].note!.includes('lab-report-default'), '版本 note 带规则集 id');
  // 产物可打开、结构符合预期
  const outline = dumpOutline(ws.getDocumentBuffer(r.docId));
  assert.ok(outline.paragraphs.some((p) => p.text === '新文档'));
  cleanup();
});

test('空白 docx 保真：产物再经 applyEdits([]) 纯 round-trip 不自检失败', () => {
  const { recognizers, styles } = loadLabReport();
  const buffer = generateFromMarkdown('# 标题\n\n正文。', { recognizers, styles });
  const r = applyEdits(buffer, []);
  assert.equal(r.result.errors.length, 0);
  assert.ok(r.result.selfCheck.ok);
  // 再打包一次（toBuffer）仍可重解析，部件不丢失
  const reopened = openDocx(buffer);
  assert.ok(reopened.parts.has('word/document.xml'));
});

test('doc_generate 工具：execute 生成入库；坏参数走 isError', async () => {
  const { ws, cleanup } = freshWorkspace();
  const byName = Object.fromEntries(createTools(ws).map((t) => [t.name, t]));
  assert.ok(byName.doc_generate, '白名单含 doc_generate');
  assert.equal(byName.doc_generate.executionMode, 'sequential');

  const res = await byName.doc_generate.execute('g1', {
    markdown: '# 报价单\n\n| 项目 | 价格 |\n| --- | --- |\n| 咨询费 | 5000 |',
    rulesetId: 'lab-report-default',
    name: '报价单.docx',
  });
  assert.equal(res.isError, false);
  const data = JSON.parse(res.content[0].text);
  assert.equal(data.version.id, 'v1');
  assert.equal(data.name, '报价单.docx');

  // 坏 rulesetId → isError + 错误信息
  const bad = await byName.doc_generate.execute('g2', { markdown: '# x', rulesetId: 'no-such-ruleset' });
  assert.equal(bad.isError, true);
  assert.ok(JSON.parse(bad.content[0].text).error.includes('内置规则集不存在'));
  // 空 markdown → isError
  const empty = await byName.doc_generate.execute('g3', { markdown: '   ', rulesetId: 'lab-report-default' });
  assert.equal(empty.isError, true);
  cleanup();
});

test('docgen 产物含 {{占位符}}：extractPlaceholders 可提取；实例化填值后重解析已替换且无残留', () => {
  const { ws, cleanup } = freshWorkspace();
  const md = [
    '# 投标文件',
    '',
    '{{项目名称}}',
    '',
    '项目编号：{{项目编号}}',
    '',
    '## 一、投标函',
    '',
    '我方愿意以人民币（大写）**{{投标总报价大写}}**（小写：{{投标总报价小写}}元）投标。',
  ].join('\n');
  const { docId } = ws.generateDocument(md, 'bid-default', '投标文件（占位符源）.docx');

  // docgen 产物 → 占位符提取
  const ph = extractPlaceholders(ws.getDocumentBuffer(docId));
  const names = ph.map((p) => p.name).sort();
  assert.deepEqual(names, ['投标总报价大写', '投标总报价小写', '项目名称', '项目编号'].sort());

  // 上传为模板 → 实例化填值
  const { templateId } = ws.uploadTemplate(ws.getDocumentBuffer(docId), '投标占位符模板');
  const { docId: newDocId, version, replaced, errors } = ws.instantiateTemplate(templateId, {
    项目名称: '智慧园区管理系统建设项目',
    项目编号: 'ZB-2026-001',
    投标总报价大写: '壹佰万元整',
    投标总报价小写: '1,000,000.00',
  }, '投标文件（智慧园区）.docx');
  assert.equal(errors.length, 0);
  assert.equal(version.id, 'v1');
  assert.ok(replaced.length >= 4, `应替换 4 处占位符，实际 ${replaced.length}`);

  // 重解析产物文本：含填入值、不含任何 {{占位符}}
  const doc = openDocx(ws.getDocumentBuffer(newDocId));
  const text = JSON.stringify(doc.parts.get('word/document.xml')!.tree!);
  assert.ok(text.includes('智慧园区管理系统建设项目'), '项目名称已填入');
  assert.ok(text.includes('壹佰万元整'), '大写金额已填入');
  assert.ok(text.includes('ZB-2026-001'), '项目编号已填入');
  assert.ok(!text.includes('{{'), '无占位符残留');
  cleanup();
});

test('template_instantiate 工具：execute 填值实例化入库；模板不存在/坏参数走 isError', async () => {
  const { ws, cleanup } = freshWorkspace();
  const byName = Object.fromEntries(createTools(ws).map((t) => [t.name, t]));
  assert.ok(byName.template_instantiate, '白名单含 template_instantiate');
  assert.equal(byName.template_instantiate.executionMode, 'sequential');

  // 先备一个带占位符的模板
  const md = '# 投标文件\n\n{{项目名称}}\n\n项目编号：{{项目编号}}\n';
  const { docId } = ws.generateDocument(md, 'bid-default', '源.docx');
  const { templateId } = ws.uploadTemplate(ws.getDocumentBuffer(docId), '投标占位符模板');

  const res = await byName.template_instantiate.execute('t1', {
    templateId,
    values: { 项目名称: '智慧园区', 项目编号: 'ZB-2026-001' },
    name: '投标文件（智慧园区）.docx',
  });
  assert.equal(res.isError, false);
  const data = JSON.parse(res.content[0].text);
  assert.equal(data.version.id, 'v1');
  assert.equal(data.name, '投标文件（智慧园区）.docx');
  assert.ok(data.replaced.length >= 2, '占位符已被替换');
  assert.equal(data.errors.length, 0);
  assert.ok(ws.listDocuments().some((d) => d.docId === data.docId), '实例化产物已入库');

  // 模板不存在 → isError + 友好错误
  const missing = await byName.template_instantiate.execute('t2', { templateId: 'nope-1234', values: {} });
  assert.equal(missing.isError, true);
  assert.ok(JSON.parse(missing.content[0].text).error.includes('模板不存在'));

  // 坏参数：缺 values / 空 templateId → isError
  const noValues = await byName.template_instantiate.execute('t3', { templateId });
  assert.equal(noValues.isError, true);
  const noTpl = await byName.template_instantiate.execute('t4', { templateId: '  ', values: {} });
  assert.equal(noTpl.isError, true);
  cleanup();
});
