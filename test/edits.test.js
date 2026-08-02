// edits.test.js — docx 编辑内核 MVP 原语的 buffer→buffer 外部行为测试。
//
// 原则（spec 测试策略）：只测 seam 外部行为 —— applyEdits(buffer, commands) 的产物
// 重新打开后断言 OOXML 语义；不 inspect 内部文档模型。覆盖：
//   段落原语 / run 原语 / 节原语 / findReplace（跨 run）/ 结构原语（add/remove/move）
//   numbering 多级编号 / normalize 规则集 / 页脚页码 / 错误处理 / schema 顺序 / 自检

import { test } from 'node:test';
import assert from 'node:assert/strict';
import PizZip from 'pizzip';
import { applyEdits } from '../src/docx-core/applyEdits.js';
import { openDocx } from '../src/docx-core/docx.js';
import { findChild, findChildren, getAttr, childrenOf, isElement, tagOf } from '../src/docx-core/ooxml.js';
import { dumpOutline } from '../src/docx-core/outline.js';

// ---- 最小合法 docx 工厂（精确的初始结构，便于断言） ----

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function para(text, { split = false, pPr = '' } = {}) {
  const runs = split
    ? text.map((t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`).join('')
    : `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  return `<w:p>${pPr}${runs}</w:p>`;
}

function makeDocx(bodyInner, { withRels = true } = {}) {
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
    `<w:document ${W_NS}><w:body>${bodyInner}` +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>');
  if (withRels) {
    zip.file('word/_rels/document.xml.rels', DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  }
  return zip.generate({ type: 'nodebuffer' });
}

// 重新打开产物并取 document.xml 树
function reopen(buffer) {
  return openDocx(buffer);
}
function docBody(docx) {
  const tree = docx.parts.get('word/document.xml').tree;
  const doc = tree.find((n) => isElement(n, 'w:document'));
  return findChild(doc, 'w:body');
}
function pAt(docx, n) {
  return findChildren(docBody(docx), 'w:p')[n - 1];
}

test('set 段落对齐：/body/p[1] jc=center，其余部件原样', () => {
  const src = makeDocx(para('标题') + para('正文'));
  const { buffer, result } = applyEdits(src, [
    { command: 'set', path: '/body/p[1]', props: { align: 'center' } },
  ]);
  assert.equal(result.errors.length, 0);
  assert.ok(result.selfCheck.ok);
  const doc = reopen(buffer);
  const jc = findChild(findChild(pAt(doc, 1), 'w:pPr'), 'w:jc');
  assert.equal(getAttr(jc, 'w:val'), 'center');
  // 第二段不受影响
  assert.equal(findChild(pAt(doc, 2), 'w:pPr'), undefined);
});

test('set run 属性：中文字体 eastAsia + 三号字（16pt→半磅 32）', () => {
  const src = makeDocx(para('标题'));
  const { buffer, result } = applyEdits(src, [
    { command: 'set', path: '/body/p[1]/r[1]', props: { eastAsia: '黑体', ascii: 'Times New Roman', sizePt: 16, bold: true } },
  ]);
  assert.equal(result.errors.length, 0);
  const doc = reopen(buffer);
  const r = findChild(pAt(doc, 1), 'w:r');
  const rPr = findChild(r, 'w:rPr');
  const rFonts = findChild(rPr, 'w:rFonts');
  assert.equal(getAttr(rFonts, 'w:eastAsia'), '黑体');
  assert.equal(getAttr(rFonts, 'w:ascii'), 'Times New Roman');
  assert.equal(getAttr(findChild(rPr, 'w:sz'), 'w:val'), '32');
  assert.equal(getAttr(findChild(rPr, 'w:b'), 'w:val'), 'true');
});

test('set 段落：28 磅固定行距 + 首行缩进 2 字符（公文正文）', () => {
  const src = makeDocx(para('正文第一段。'));
  const { buffer } = applyEdits(src, [
    { command: 'set', path: '/body/p[1]', props: { lineSpacingPt: 28, firstLineChars: 200 } },
  ]);
  const doc = reopen(buffer);
  const pPr = findChild(pAt(doc, 1), 'w:pPr');
  const spacing = findChild(pPr, 'w:spacing');
  assert.equal(getAttr(spacing, 'w:line'), '560');
  assert.equal(getAttr(spacing, 'w:lineRule'), 'exact');
  assert.equal(getAttr(findChild(pPr, 'w:ind'), 'w:firstLineChars'), '200');
});

test('schema 顺序：已有 jc 的段落补 spacing，spacing 必须排在 jc 前', () => {
  const src = makeDocx(para('x', { pPr: '<w:pPr><w:jc w:val="center"/></w:pPr>' }));
  const { buffer } = applyEdits(src, [
    { command: 'set', path: '/body/p[1]', props: { spacingBeforePt: 10 } },
  ]);
  const doc = reopen(buffer);
  const pPr = findChild(pAt(doc, 1), 'w:pPr');
  const tags = childrenOf(pPr).map(tagOf).filter(Boolean);
  assert.ok(tags.indexOf('w:spacing') < tags.indexOf('w:jc'), `顺序错误: ${tags}`);
});

test('set match 批量：正则命中段落统一套用（正文统一仿宋四号）', () => {
  const src = makeDocx(para('第一章 总纲') + para('正文甲。') + para('正文乙。'));
  const { buffer, result } = applyEdits(src, [
    { command: 'set', match: { text: '^正文' }, props: { run: { eastAsia: '仿宋_GB2312', sizePt: 14 } } },
  ]);
  assert.equal(result.errors.length, 0);
  assert.equal(result.applied[0].detail.matched, 2);
  const doc = reopen(buffer);
  for (const n of [2, 3]) {
    const rPr = findChild(findChild(pAt(doc, n), 'w:r'), 'w:rPr');
    assert.equal(getAttr(findChild(rPr, 'w:rFonts'), 'w:eastAsia'), '仿宋_GB2312');
    assert.equal(getAttr(findChild(rPr, 'w:sz'), 'w:val'), '28');
  }
  // 首段未被误伤
  assert.equal(findChild(findChild(pAt(doc, 1), 'w:r'), 'w:rPr'), undefined);
});

test('节属性：公文页边距（cm→twips）+ A4', () => {
  const src = makeDocx(para('x'));
  const { buffer } = applyEdits(src, [
    { command: 'set', path: '/body/sectPr', props: { marginsCm: { top: 3.7, bottom: 3.5, left: 2.8, right: 2.6 }, pageSize: 'a4' } },
  ]);
  const doc = reopen(buffer);
  const sect = findChild(docBody(doc), 'w:sectPr');
  const pgMar = findChild(sect, 'w:pgMar');
  assert.equal(getAttr(pgMar, 'w:top'), String(Math.round(3.7 * 566.929)));
  assert.equal(getAttr(pgMar, 'w:left'), String(Math.round(2.8 * 566.929)));
  assert.equal(getAttr(findChild(sect, 'w:pgSz'), 'w:w'), '11906');
});

test('findReplace 跨 run：拆分 run 也能命中替换', () => {
  const src = makeDocx(para(['正', '文统一', '规范'], { split: true }));
  const { buffer, result } = applyEdits(src, [
    { command: 'findReplace', find: '文统一规', replace: 'X' },
  ]);
  assert.equal(result.errors.length, 0);
  assert.equal(result.applied[0].detail.replaced, 1);
  const doc = reopen(buffer);
  const texts = findChildren(pAt(doc, 1), 'w:r')
    .map((r) => childrenOf(findChild(r, 'w:t')).map((x) => x['#text'] || '').join(''));
  assert.equal(texts.join(''), '正X范');
});

test('add / remove / move 结构原语', () => {
  const src = makeDocx(para('A') + para('B') + para('C'));
  // add 到 body 末尾（sectPr 保持最后）
  const r1 = applyEdits(src, [
    { command: 'add', parent: '/body', node: { kind: 'paragraph', text: 'D', props: { align: 'right' } } },
  ]);
  let doc = reopen(r1.buffer);
  let paras = findChildren(docBody(doc), 'w:p');
  assert.equal(paras.length, 4);
  const bodyChildren = childrenOf(docBody(doc));
  assert.equal(tagOf(bodyChildren[bodyChildren.length - 1]), 'w:sectPr');
  // remove 第 2 段
  const r2 = applyEdits(r1.buffer, [{ command: 'remove', path: '/body/p[2]' }]);
  doc = reopen(r2.buffer);
  paras = findChildren(docBody(doc), 'w:p');
  assert.equal(paras.length, 3);
  // move p[1]（A）到末尾 → [C, D, A]
  const r3 = applyEdits(r2.buffer, [{ command: 'move', path: '/body/p[1]', parent: '/body' }]);
  doc = reopen(r3.buffer);
  paras = findChildren(docBody(doc), 'w:p');
  const textOfP = (p) => childrenOf(findChild(findChild(p, 'w:r'), 'w:t')).map((x) => x['#text']).join('');
  assert.equal(textOfP(paras[0]), 'C');
  assert.equal(textOfP(paras[1]), 'D');
  assert.equal(textOfP(paras[2]), 'A');
});

test('numbering：定义多级编号（幂等）并挂载段落', () => {
  const levels = [
    { ilvl: 0, numFmt: 'chineseCounting', lvlText: '%1、' },
    { ilvl: 1, numFmt: 'chineseCounting', lvlText: '（%2）' },
    { ilvl: 2, numFmt: 'decimal', lvlText: '%3.' },
  ];
  const src = makeDocx(para('总纲') + para('细则'));
  const r1 = applyEdits(src, [
    { command: 'numbering', action: 'define', levels },
    { command: 'numbering', action: 'attach', path: '/body/p[1]', numId: 1, ilvl: 0 },
  ]);
  assert.equal(r1.result.errors.length, 0);
  const numId = r1.result.applied[0].detail.numId;
  const doc = reopen(r1.buffer);
  // numbering.xml 部件存在且结构合法
  const numTree = doc.parts.get('word/numbering.xml').tree;
  const numRoot = numTree.find((n) => isElement(n, 'w:numbering'));
  assert.ok(numRoot);
  assert.equal(findChildren(numRoot, 'w:abstractNum').length, 1);
  assert.equal(findChildren(numRoot, 'w:num').length, 1);
  // 段落挂载 numPr
  const pPr = findChild(pAt(doc, 1), 'w:pPr');
  const numPr = findChild(pPr, 'w:numPr');
  assert.equal(getAttr(findChild(numPr, 'w:numId'), 'w:val'), String(numId));
  // 幂等：相同 levels 再 define 返回同一 numId，不新增 abstractNum
  const r2 = applyEdits(r1.buffer, [{ command: 'numbering', action: 'define', levels }]);
  assert.equal(r2.result.applied[0].detail.numId, numId);
  const doc2 = reopen(r2.buffer);
  const numRoot2 = doc2.parts.get('word/numbering.xml').tree.find((n) => isElement(n, 'w:numbering'));
  assert.equal(findChildren(numRoot2, 'w:abstractNum').length, 1);
  // content-types / rels 已注册
  const ct = doc2.parts.get('[Content_Types].xml').text || '';
  assert.ok(ct.includes('numbering+xml') || reopen(r2.buffer).parts.get('[Content_Types].xml').dirty);
});

test('pageNumber footer：页脚 PAGE 字段 + 部件注册', () => {
  const src = makeDocx(para('正文'));
  const { buffer, result } = applyEdits(src, [
    { command: 'pageNumber', action: 'footer', align: 'center' },
  ]);
  assert.equal(result.errors.length, 0);
  const doc = reopen(buffer);
  const footerName = result.applied[0].detail.footer;
  const footer = doc.parts.get(footerName);
  assert.ok(footer, 'footer 部件存在');
  const fRoot = footer.tree.find((n) => isElement(n, 'w:ftr'));
  const xml = footer.text || '';
  // 树内应有 PAGE 指令
  const hasPage = JSON.stringify(fRoot).includes('PAGE');
  assert.ok(hasPage);
  // sectPr 引用 footer
  const sect = findChild(docBody(doc), 'w:sectPr');
  const ref = findChildren(sect, 'w:footerReference').find((f) => getAttr(f, 'w:type') === 'default');
  assert.ok(ref, 'footerReference 已挂载');
  // rels 与 content-types 已注册
  const rels = doc.parts.get('word/_rels/document.xml.rels');
  assert.ok(JSON.stringify(rels.tree).includes('footer'));
  assert.ok(JSON.stringify(doc.parts.get('[Content_Types].xml').tree).includes('footer'));
});

test('normalize：规则集驱动全文重排（标题/正文两类规则）', () => {
  const src = makeDocx(para('关于加强排版工作的通知') + para('各单位：请遵照执行。'));
  const ruleset = {
    rules: [
      { name: 'title', match: { text: '^关于.*通知$' }, set: { paragraph: { align: 'center' }, run: { eastAsia: '黑体', sizePt: 16 } } },
      { name: 'body', match: { text: '.' }, set: { paragraph: { firstLineChars: 200, lineSpacingPt: 28 }, run: { eastAsia: '仿宋_GB2312', sizePt: 14 } } },
    ],
  };
  const { buffer, result } = applyEdits(src, [{ command: 'normalize', ruleset }]);
  assert.equal(result.errors.length, 0);
  assert.deepEqual(result.applied[0].detail.normalized, { title: 1, body: 1 });
  const doc = reopen(buffer);
  const p1Pr = findChild(pAt(doc, 1), 'w:pPr');
  assert.equal(getAttr(findChild(p1Pr, 'w:jc'), 'w:val'), 'center');
  const p2Pr = findChild(pAt(doc, 2), 'w:pPr');
  assert.equal(getAttr(findChild(p2Pr, 'w:spacing'), 'w:line'), '560');
});

test('错误处理：坏路径收集结构化错误（含建议），后续命令继续执行', () => {
  const src = makeDocx(para('A') + para('B'));
  const { buffer, result } = applyEdits(src, [
    { command: 'set', path: '/body/p[99]', props: { align: 'center' } },
    { command: 'set', path: '/body/p[2]', props: { align: 'right' } },
    { command: 'frobnicate', path: '/body/p[1]' },
  ]);
  assert.equal(result.errors.length, 2);
  assert.equal(result.errors[0].code, 'PATH_NOT_FOUND');
  assert.ok(result.errors[0].suggestion, '带自愈建议');
  assert.equal(result.errors[1].code, 'UNKNOWN_COMMAND');
  assert.equal(result.applied.length, 1); // 第二条成功
  const doc = reopen(buffer);
  assert.equal(getAttr(findChild(findChild(pAt(doc, 2), 'w:pPr'), 'w:jc'), 'w:val'), 'right');
});

test('生成后自检：产物可重解析且 document.xml 结构完整', () => {
  const src = makeDocx(para('x'));
  const { buffer, result } = applyEdits(src, []);
  assert.ok(result.selfCheck.ok);
  assert.ok(result.selfCheck.parts >= 3);
  // 空命令 = 纯 round-trip：未编辑部件原样
  const doc = reopen(buffer);
  assert.equal(doc.parts.get('word/document.xml').text.includes('<w:body>'), true);
});

test('outline dump：段落路径与文本预览可用于寻址', () => {
  const src = makeDocx(para('标题文字') + para('正文内容'));
  const outline = dumpOutline(src);
  assert.equal(outline.paragraphCount, 2);
  assert.equal(outline.paragraphs[0].path, '/body/p[1]');
  assert.equal(outline.paragraphs[0].text, '标题文字');
  assert.equal(outline.sections.length, 1);
  assert.equal(outline.sections[0].pageSize.w, 11906);
});
