// templates.test.js — 模板层行为测试（spec 测试策略：占位符提取 / 规则集抽取 /
// 实例化产物为合法 docx（走编辑内核自检）/ 规则集 → 内核命令翻译）。

import { test } from 'node:test';
import assert from 'node:assert/strict';
import PizZip from 'pizzip';
import { mkdtempSync, rmSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { LocalFsObjectStore } from '../src/storage/objectStore.js';
import { VersionStore } from '../src/storage/versionStore.js';
import { TemplateStore } from '../src/templates/templateStore.js';
import { extractPlaceholders, placeholderCommands } from '../src/templates/placeholders.js';
import { rulesetToCommands, styleToKernelProps } from '../src/templates/rulesetToCommands.js';
import { loadRuleset } from '../src/ruleset/load.js';
import { applyEdits } from '../src/docx-core/applyEdits.js';
import { openDocx } from '../src/docx-core/docx.js';
import { findChild, findChildren, getAttr, isElement, childrenOf, tagOf } from '../src/docx-core/ooxml.js';
import { fileURLToPath } from 'node:url';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const HERE = fileURLToPath(new URL('.', import.meta.url));
const DEFAULT_RULESET_DIR = [
  join(HERE, '../templates/rulesets/gongwen-default'),
  join(HERE, '../../templates/rulesets/gongwen-default'),
].find((p) => existsSync(p)) || join(HERE, '../templates/rulesets/gongwen-default');

function makeTemplateDocx() {
  const p = (text: string, pPr = '', rPr = '') =>
    `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const zip = new PizZip();
  zip.file('[Content_Types].xml', DECL +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', DECL +
    `<w:document ${W_NS}><w:body>` +
    p('{{发文单位}}关于加强排版工作的通知', '<w:pPr><w:jc w:val="center"/></w:pPr>',
      '<w:rPr><w:rFonts w:eastAsia="方正小标宋简体" w:ascii="Times New Roman"/><w:sz w:val="44"/></w:rPr>') +
    p('各科室、直属各单位：为进一步规范机关公文格式，提高公文处理质量和效率，根据有关规定，现将有关事项通知如下，请结合实际认真贯彻执行。', '', '<w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/><w:sz w:val="32"/></w:rPr>') +
    p('联系人：{{联系人}}　电话：{{联系电话}}') +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2098" w:right="1474" w:bottom="1985" w:left="1588" w:header="851" w:footer="1418" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>');
  zip.file('word/_rels/document.xml.rels', DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  return zip.generate({ type: 'nodebuffer' });
}

function freshStores() {
  const dir = mkdtempSync(join(tmpdir(), 'paiban-tpl-'));
  const objects = new LocalFsObjectStore(dir);
  const versions = new VersionStore(dir, objects);
  return { dir, templates: new TemplateStore(dir, objects, versions), versions, cleanup: () => rmSync(dir, { recursive: true, force: true }) };
}

test('占位符提取：名称 / 次数 / 首现路径', () => {
  const buf = makeTemplateDocx();
  const ph = extractPlaceholders(buf);
  const names = ph.map((x) => x.name).sort();
  assert.deepEqual(names, ['发文单位', '联系人', '联系电话'].sort());
  assert.equal(ph.find((x) => x.name === '发文单位')!.count, 1);
  assert.ok(ph.find((x) => x.name === '发文单位')!.firstPath.startsWith('/body/p['));
});

test('规则集反推：title/body/page 实测并通过两文件 schema 校验', () => {
  const { templates, cleanup } = freshStores();
  const { templateId, meta, extracted } = templates.uploadTemplate(makeTemplateDocx(), { name: '通知模板' });
  assert.ok(extracted.includes('title') && extracted.includes('body') && extracted.includes('page'));
  const { recognizers, styles } = templates.readTemplate(templateId);
  // title 反推：居中 / 方正小标宋 / 22pt
  assert.equal(styles.components.title.fontEastAsia, '方正小标宋简体');
  assert.equal(styles.components.title.sizePt, 22);
  assert.equal(styles.components.title.align, 'center');
  // body 反推：最长段落 → 仿宋 16pt
  assert.equal(styles.components.body.fontEastAsia, '仿宋_GB2312');
  assert.equal(styles.components.body.sizePt, 16);
  // page 反推：A4 + cm 换算
  assert.equal(styles.page.paper, 'A4');
  assert.ok(Math.abs(styles.page.margins.topCm - 3.7) < 0.01);
  assert.ok(Math.abs(styles.page.margins.leftCm - 2.8) < 0.01);
  assert.equal(meta.placeholderCount, 3);
  cleanup();
});

test('实例化：占位符合并 → 新工作文档 v1，产物合法（自检）', () => {
  const { templates, versions, cleanup } = freshStores();
  const { templateId } = templates.uploadTemplate(makeTemplateDocx(), { name: '通知模板' });
  const { docId, version, errors } = templates.instantiate(templateId, {
    发文单位: '某某市人民政府办公室',
    联系人: '张三',
    联系电话: '12345678',
  }, { name: '正式通知.docx' });
  assert.equal(errors.length, 0);
  assert.equal(version.id, 'v1');
  const buf = versions.getBuffer(docId);
  // 走编辑内核自检（重解析）
  const r = applyEdits(buf, []);
  assert.ok(r.result.selfCheck.ok);
  // 占位符已被替换
  const doc = openDocx(buf);
  const text = JSON.stringify(doc.parts.get('word/document.xml')!.tree!);
  assert.ok(text.includes('某某市人民政府办公室'));
  assert.ok(text.includes('张三'));
  assert.ok(!text.includes('{{'));
  cleanup();
});

test('规则集 → 内核命令：normalize 规则顺序与属性映射 + 页面设置', () => {
  const { recognizers, styles } = loadRuleset(DEFAULT_RULESET_DIR);
  const commands = rulesetToCommands(recognizers, styles);
  const normalize = commands.find((c) => c.command === 'normalize');
  assert.ok(normalize);
  const rules = normalize.ruleset.rules as Array<{ name: string; match?: any; set?: any }>;
  const names = rules.map((r) => r.name);
  // 顺序：table（表格内段落优先认领）→ title → subtitle → heading1..4 → caption → attachment → body（兜底最后）
  assert.deepEqual(names, ['table', 'title', 'subtitle', 'heading1', 'heading2', 'heading3', 'heading4', 'caption', 'attachment', 'body']);
  const table = rules[0]!;
  assert.deepEqual(table.match, { element: 'table' }); // heuristic isTableElement → 元素谓词
  assert.equal(table.set.run.eastAsia, '仿宋_GB2312');
  assert.equal(table.set.run.sizePt, 14);
  const title = rules.find((r) => r.name === 'title')!;
  assert.deepEqual(title.match, { position: 'documentStart' });
  assert.equal(title.set.run.eastAsia, '方正小标宋简体');
  assert.equal(title.set.paragraph.align, 'center');
  const h1 = rules.find((r) => r.name === 'heading1')!;
  assert.equal(h1.match.text, '^[一二三四五六七八九十百]+、');
  assert.equal(h1.set.paragraph.firstLineChars, 200);
  assert.equal(h1.set.paragraph.outlineLevel, 0); // 1 起 → 0 起
  const body = rules.find((r) => r.name === 'body')!;
  assert.deepEqual(body.match, { fallback: true });
  // 页面设置命令
  const pageCmd = commands.find((c) => c.command === 'set');
  assert.ok(pageCmd);
  assert.equal(pageCmd.props.marginsCm.top, 3.7);
  assert.equal(pageCmd.props.marginsCm.left, 2.8);
  assert.equal(pageCmd.props.marginsCm.footer, 2.5);
  // 页码命令（奇偶页不同对齐：evenAlign 必须随命令下发）
  const pnCmd = commands.find((c) => c.command === 'pageNumber');
  assert.ok(pnCmd);
  assert.equal(pnCmd.align, 'right');
  assert.equal(pnCmd.evenAlign, 'left');
});

test('端到端：内置公文规则集 normalize 一篇乱排版文档', () => {
  // 乱排版：标题不居中、标题正则命中段落无格式、正文无缩进
  const zip = new PizZip();
  const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  zip.file('[Content_Types].xml', DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  zip.file('word/document.xml', DECL + `<w:document ${W_NS}><w:body>` +
    p('关于加强公文规范化管理工作的通知') +
    p('一、总体要求') +
    p('各科室：为规范公文格式，提高质量，现通知如下，请认真贯彻执行。') +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>');
  zip.file('word/_rels/document.xml.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  const buf = zip.generate({ type: 'nodebuffer' });

  const { recognizers, styles } = loadRuleset(DEFAULT_RULESET_DIR);
  const commands = rulesetToCommands(recognizers, styles);
  const { buffer, result } = applyEdits(buf, commands);
  assert.equal(result.errors.length, 0);
  assert.ok(result.selfCheck.ok);
  const stats = result.applied[0].detail.normalized;
  assert.equal(stats.title, 1);
  assert.equal(stats.heading1, 1);
  assert.equal(stats.body, 1);

  const doc = openDocx(buffer);
  const tree = doc.parts.get('word/document.xml')!.tree!;
  const docEl = tree.find((n) => isElement(n, 'w:document'));
  const body = findChild(docEl, 'w:body');
  const paras = findChildren(body, 'w:p');
  // 标题：居中 + 方正小标宋 22pt
  const tPr = findChild(paras[0], 'w:pPr');
  assert.equal(getAttr(findChild(tPr, 'w:jc'), 'w:val'), 'center');
  const tRPr = findChild(findChild(paras[0], 'w:r'), 'w:rPr');
  assert.equal(getAttr(findChild(tRPr, 'w:rFonts'), 'w:eastAsia'), '方正小标宋简体');
  assert.equal(getAttr(findChild(tRPr, 'w:sz'), 'w:val'), '44');
  // 一级标题：黑体 16pt + 大纲级 0 + 首行缩进 2 字符
  const h1Pr = findChild(paras[1], 'w:pPr');
  assert.equal(getAttr(findChild(h1Pr, 'w:outlineLvl'), 'w:val'), '0');
  assert.equal(getAttr(findChild(h1Pr, 'w:ind'), 'w:firstLineChars'), '200');
  const h1RPr = findChild(findChild(paras[1], 'w:r'), 'w:rPr');
  assert.equal(getAttr(findChild(h1RPr, 'w:rFonts'), 'w:eastAsia'), '黑体');
  // 正文：仿宋 16pt + 28 磅行距 + 两端对齐
  const bPr = findChild(paras[2], 'w:pPr');
  assert.equal(getAttr(findChild(bPr, 'w:spacing'), 'w:line'), '560');
  assert.equal(getAttr(findChild(bPr, 'w:jc'), 'w:val'), 'both');
  // 页面：公文页边距已写入
  const sect = findChild(body, 'w:sectPr');
  assert.equal(getAttr(findChild(sect, 'w:pgMar'), 'w:top'), String(Math.round(3.7 * 566.929)));
});

test('表格组件规则：isTableElement 翻译为 element:table，重排后表格套模板样式', () => {
  // 乱排版文档：题目 + 一张 2×2 表格（含形似三级标题的单元格文本「1. 指标」）+ 表格外正文段
  const zip = new PizZip();
  const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const tc = (t: string) => `<w:tc>${p(t)}</w:tc>`;
  zip.file('[Content_Types].xml', DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  zip.file('word/document.xml', DECL + `<w:document ${W_NS}><w:body>` +
    p('关于加强公文规范化管理工作的通知') +
    '<w:tbl><w:tr>' + tc('指标名称') + tc('1. 指标') + '</w:tr><w:tr>' + tc('完成率') + tc('42') + '</w:tr></w:tbl>' +
    p('各科室：请参照上表数据认真贯彻执行。') +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>');
  zip.file('word/_rels/document.xml.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  const buf = zip.generate({ type: 'nodebuffer' });

  const { recognizers, styles } = loadRuleset(DEFAULT_RULESET_DIR);
  const commands = rulesetToCommands(recognizers, styles);
  const { buffer, result } = applyEdits(buf, commands);
  assert.equal(result.errors.length, 0);
  assert.ok(result.selfCheck.ok);
  const stats = result.applied[0].detail.normalized;
  assert.equal(stats.table, 4); // 4 个单元格段落全部归 table 规则
  assert.equal(stats.title, 1);
  assert.equal(stats.body, 1); // 表格外正文段仍归 body
  assert.equal(stats.heading3, undefined); // 「1. 指标」不得被三级标题正则吞噬

  const doc = openDocx(buffer);
  const tree = doc.parts.get('word/document.xml')!.tree!;
  const docEl = tree.find((n) => isElement(n, 'w:document'));
  const body = findChild(docEl, 'w:body');
  const tbl = findChild(body, 'w:tbl');
  assert.ok(tbl, '表格存在');
  // 每个单元格段落：仿宋_GB2312 + 14pt（sz 半磅 28），且不打大纲级别
  const cellParas: Array<ReturnType<typeof findChild>> = [];
  for (const tr of findChildren(tbl, 'w:tr')) {
    for (const cell of findChildren(tr, 'w:tc')) {
      cellParas.push(...findChildren(cell, 'w:p'));
    }
  }
  assert.equal(cellParas.length, 4);
  for (const cp of cellParas) {
    const rPr = findChild(findChild(cp, 'w:r'), 'w:rPr');
    assert.equal(getAttr(findChild(rPr, 'w:rFonts'), 'w:eastAsia'), '仿宋_GB2312');
    assert.equal(getAttr(findChild(rPr, 'w:sz'), 'w:val'), '28');
    const pPr = findChild(cp, 'w:pPr');
    assert.equal(pPr ? findChild(pPr, 'w:outlineLvl') : undefined, undefined);
  }
});

test('页码奇偶页不同对齐：evenAlign 生成偶数页页脚并声明 evenAndOddHeaders', () => {
  // 最小文档（无 settings.xml 部件，覆盖内核最小创建路径）
  const zip = new PizZip();
  const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  zip.file('[Content_Types].xml', DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  zip.file('word/document.xml', DECL + `<w:document ${W_NS}><w:body>` +
    p('正文。') +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>');
  zip.file('word/_rels/document.xml.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  const buf = zip.generate({ type: 'nodebuffer' });

  const { recognizers, styles } = loadRuleset(DEFAULT_RULESET_DIR);
  const commands = rulesetToCommands(recognizers, styles);
  const { buffer, result } = applyEdits(buf, commands);
  assert.equal(result.errors.length, 0);
  assert.ok(result.selfCheck.ok);

  const doc = openDocx(buffer);
  // settings.xml 已创建并声明奇偶页不同
  const settings = doc.parts.get('word/settings.xml');
  assert.ok(settings, 'settings.xml 部件已创建');
  const sRoot = settings.tree!.find((n) => isElement(n, 'w:settings'))!;
  assert.ok(findChild(sRoot, 'w:evenAndOddHeaders'), 'evenAndOddHeaders 已声明');
  // sectPr 同时挂载 default（奇数页）与 even（偶数页）页脚引用
  const tree = doc.parts.get('word/document.xml')!.tree!;
  const docEl = tree.find((n) => isElement(n, 'w:document'));
  const sect = findChild(findChild(docEl, 'w:body'), 'w:sectPr');
  const refTypes = findChildren(sect, 'w:footerReference').map((f) => getAttr(f, 'w:type'));
  assert.ok(refTypes.includes('default') && refTypes.includes('even'), `footerReference 类型: ${refTypes}`);
  // 奇数页脚右对齐、偶数页脚左对齐，均含 PAGE 字段
  const pnDetail = result.applied.find((a) => a.command === 'pageNumber')!.detail;
  assert.ok(pnDetail.evenFooter, '返回偶数页脚部件名');
  const footerJc = (partName: string) => {
    const fRoot = doc.parts.get(partName)!.tree!.find((n) => isElement(n, 'w:ftr'))!;
    const fp = findChildren(fRoot, 'w:p').find((x) => JSON.stringify(x).includes('PAGE'))!;
    return getAttr(findChild(findChild(fp, 'w:pPr'), 'w:jc'), 'w:val');
  };
  assert.equal(footerJc(pnDetail.footer), 'right');
  assert.equal(footerJc(pnDetail.evenFooter), 'left');
});
