// lab-report.test.js — issue #15 实验报告 fixtures 校验。
//
// fixtures 位于 test/fixtures/lab-report/（template.docx 规范模板 / messy-draft.docx 乱排版原稿，
// 由同目录 generate.ts 生成，文字内容与结构一致）。
//
// 断言：
//   1. 两个 fixture 都能经项目 docx-core 解析链路 openDocx 打开，round-trip（open→toBuffer→open）不抛错、部件不丢失
//   2. 全部 XML 部件 build 幂等（round-trip 不动点，同 roundtrip.test.ts 口径）
//   3. 两份文档逐段文本完全一致（同样的文字内容与结构）
//   4. 模板具备样式层级（标题三级 + 题注样式 + 页脚页码域），原稿无任何自定义样式引用、无大纲级别、含多余空行

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openDocx, toBuffer } from '../src/docx-core/docx.js';
import { parseXml, buildXml } from '../src/docx-core/xml.js';

// 源码在 test/fixtures/lab-report；编译产物在 dist/test/ 时回退到 test/fixtures/lab-report。
// 注意：tsc 会把 generate.ts 编译进 dist/test/fixtures/lab-report/，因此按 docx 文件而非目录判断。
const HERE = fileURLToPath(new URL('.', import.meta.url));
const DIR = [
  join(HERE, 'fixtures/lab-report'),
  join(HERE, '../../test/fixtures/lab-report'),
].find((p) => existsSync(join(p, 'template.docx')))!;

const FILES = ['template.docx', 'messy-draft.docx'];

function load(name: string) {
  return openDocx(readFileSync(join(DIR, name)));
}

/** 按 w:p 分段提取 w:t 文本（表内单元格文本归属所在段落）。 */
function extractParagraphTexts(xml: string): string[] {
  const { tree } = parseXml(xml);
  const out: string[] = [];
  let cur: string | null = null;
  const walk = (nodes: Array<Record<string, unknown>>): void => {
    for (const node of nodes) {
      const tag = Object.keys(node).find((k) => k !== ':@' && k !== '#text');
      if (!tag) continue;
      const text = node['#text'];
      if (tag === 'w:t') {
        // fxp preserveOrder 下 #text 是 w:t 的子节点；parseTagValue:false 后文本恒为字符串（issue #27）
        const kids = node[tag];
        const pieces: string[] = [];
        if (typeof text === 'string') pieces.push(text);
        if (Array.isArray(kids)) {
          for (const kid of kids as Array<Record<string, unknown>>) {
            const t = kid['#text'];
            if (typeof t === 'string') pieces.push(t);
          }
        }
        cur = (cur ?? '') + pieces.join('');
      } else if (tag === 'w:p') {
        if (cur !== null) {
          out.push(cur);
          cur = null;
        }
      }
      const kids = node[tag];
      if (Array.isArray(kids)) walk(kids as Array<Record<string, unknown>>);
    }
  };
  walk(tree);
  if (cur !== null) out.push(cur);
  return out;
}

/** 统计 XML 树中某标签出现次数。 */
function countTag(xml: string, tag: string): number {
  const { tree } = parseXml(xml);
  let n = 0;
  const walk = (nodes: Array<Record<string, unknown>>): void => {
    for (const node of nodes) {
      const t = Object.keys(node).find((k) => k !== ':@' && k !== '#text');
      if (t === tag) n++;
      const kids = t ? node[t] : undefined;
      if (Array.isArray(kids)) walk(kids as Array<Record<string, unknown>>);
    }
  };
  walk(tree);
  return n;
}

test('两个 fixture 可经项目 docx-core 打开并 round-trip 不抛错', () => {
  for (const f of FILES) {
    const doc = load(f);
    assert.ok(doc.parts.has('word/document.xml'), `[${f}] 缺 word/document.xml`);
    assert.ok(doc.parts.has('word/media/pendulum.png'), `[${f}] 缺装置图部件`);
    assert.equal(doc.parts.get('word/document.xml')!.kind, 'xml');
    const reopened = openDocx(toBuffer(doc));
    assert.equal(reopened.parts.size, doc.parts.size, `[${f}] round-trip 后部件数量变化`);
    for (const [name] of doc.parts) {
      assert.ok(reopened.parts.has(name), `[${f}] round-trip 后部件 ${name} 丢失`);
    }
  }
});

test('两个 fixture 全部 XML 部件 build 幂等（round-trip 不动点）', () => {
  for (const f of FILES) {
    const doc = load(f);
    for (const [name, part] of doc.parts) {
      if (part.kind !== 'xml') continue;
      const first = buildXml(part.tree, part.meta);
      const { tree, meta } = parseXml(first);
      assert.equal(
        buildXml(tree, meta),
        first,
        `[${f}] ${name} build 非幂等`,
      );
    }
  }
});

test('模板与原稿逐段文本完全一致（同样的文字内容与结构）', () => {
  const tpl = extractParagraphTexts(load('template.docx').parts.get('word/document.xml')!.text as string);
  const messy = extractParagraphTexts(load('messy-draft.docx').parts.get('word/document.xml')!.text as string);
  assert.deepEqual(messy, tpl);
  // 冒烟：确实覆盖了标题/三级标题/题注/表格等关键组件
  assert.ok(tpl.includes('单摆测重力加速度实验报告'), '缺报告标题');
  assert.ok(tpl.some((t) => /^[一二三四五六]、/.test(t)), '缺一级标题');
  assert.ok(tpl.some((t) => /^（[一二三四五六]）/.test(t)), '缺二级标题');
  assert.ok(tpl.some((t) => /^\d+\. /.test(t)), '缺三级标题');
  assert.ok(tpl.includes('图 1 单摆实验装置示意图'), '缺图题');
  assert.ok(tpl.includes('表 1 单摆摆长与周期测量数据'), '缺表题');
  assert.ok(tpl.includes('9.92'), '缺表格数据');
});

test('模板具备规范样式层级与页脚页码，原稿无样式层级且含多余空行', () => {
  const tplDoc = load('template.docx');
  const tplXml = tplDoc.parts.get('word/document.xml')!.text as string;
  const tplStyles = tplDoc.parts.get('word/styles.xml')!.text as string;

  // 模板：标题/正文/题注/三级标题均有命名样式，标题样式带大纲级别
  for (const sid of ['ReportTitle', 'Heading1', 'Heading2', 'Heading3', 'Body', 'Caption']) {
    assert.ok(tplStyles.includes(`w:styleId="${sid}"`), `模板 styles.xml 缺样式 ${sid}`);
  }
  for (const sid of ['ReportTitle', 'Heading1', 'Heading2', 'Heading3', 'Body', 'Caption']) {
    assert.ok(tplXml.includes(`<w:pStyle w:val="${sid}"/>`), `模板 document.xml 缺 pStyle 引用 ${sid}`);
  }
  assert.ok(tplStyles.includes('<w:outlineLvl w:val="0"/>'), 'Heading1 缺大纲级别 1');
  assert.ok(tplStyles.includes('<w:outlineLvl w:val="1"/>'), 'Heading2 缺大纲级别 2');
  assert.ok(tplStyles.includes('<w:outlineLvl w:val="2"/>'), 'Heading3 缺大纲级别 3');

  // 模板：页脚页码域 + 规范页边距（上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm）
  const footer = tplDoc.parts.get('word/footer1.xml')!.text as string;
  assert.ok(footer.includes('PAGE'), '页脚缺 PAGE 页码域');
  assert.ok(tplXml.includes('w:footerReference'), 'document.xml 缺 footerReference');
  assert.ok(tplXml.includes('w:top="2098"'), '模板页边距上边距异常');
  assert.ok(tplXml.includes('w:bottom="1984"'), '模板页边距下边距异常');
  assert.ok(tplXml.includes('w:left="1588"'), '模板页边距左边距异常');
  assert.ok(tplXml.includes('w:right="1474"'), '模板页边距右边距异常');

  // 原稿：无任何自定义样式引用（标题全是普通正文样式）、无大纲级别、页边距随意、含多余空行
  const messyDoc = load('messy-draft.docx');
  const messyXml = messyDoc.parts.get('word/document.xml')!.text as string;
  assert.equal(countTag(messyXml, 'w:pStyle'), 0, '原稿不应引用任何段落样式');
  assert.equal(countTag(messyXml, 'w:outlineLvl'), 0, '原稿不应有大纲级别');
  assert.ok(!messyXml.includes('w:footerReference'), '原稿不应引用页脚');
  const paraCount = (xml: string) => countTag(xml, 'w:p');
  assert.ok(paraCount(messyXml) > paraCount(tplXml), '原稿段落数应多于模板（多余空行）');
  assert.ok(!messyXml.includes('w:top="2098"'), '原稿页边距不应是规范值');
  assert.ok(!messyDoc.parts.has('word/footer1.xml'), '原稿不应有页脚页码');
});
