// roundtrip.test.js — 第一里程碑：round-trip 保真 spike。
//
// 断言（对应 spec 首要风险与测试策略"round-trip 保真（解析→序列化→重解析逐字节稳定）"）：
//   1. 无损：buffer → open → save（未编辑）→ open，所有部件（XML 文本 + 二进制字节）原样
//   2. build 幂等：XML 部件 tree → build → parse → build 两次结果一致（round-trip 不动点，
//      即 dirty 部件序列化不会因多次往返产生漂移）
//   3. dirty 语义：只标记一个部件修改并保存，仅该部件内容被 build 输出替换，其余部件原样，
//      且新 buffer 可正常打开
//
// 回归集：test/fixtures/ 下真实 Word/WPS 产出的 .docx 样本。

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseXml, buildXml } from '../src/docx-core/xml.js';
import { openDocx, toBuffer, markDirty } from '../src/docx-core/docx.js';

const __dirname = fileURLToPath(new URL('.', import.meta.url));
const FIXTURES = join(__dirname, 'fixtures');
const samples = readdirSync(FIXTURES).filter((f) => f.endsWith('.docx'));

test('回归集：全部样本未编辑 round-trip 后部件原样无损', () => {
  for (const sample of samples) {
    const orig = openDocx(readFileSync(join(FIXTURES, sample)));
    const reopened = openDocx(toBuffer(orig));
    assert.equal(
      reopened.parts.size,
      orig.parts.size,
      `[${sample}] round-trip 后部件数量变化`,
    );
    for (const [name, part] of orig.parts) {
      const next = reopened.parts.get(name);
      assert.ok(next, `[${sample}] round-trip 后部件 ${name} 丢失`);
      if (part.kind === 'xml') {
        assert.equal(
          next.text,
          part.text,
          `[${sample}] 未编辑 round-trip 后 XML 部件 ${name} 文本漂移`,
        );
      } else {
        assert.deepEqual(
          Buffer.from(next.bytes),
          Buffer.from(part.bytes),
          `[${sample}] 未编辑 round-trip 后二进制部件 ${name} 字节漂移`,
        );
      }
    }
  }
});

test('回归集：全部 XML 部件 build 幂等（round-trip 不动点）', () => {
  let total = 0;
  let stable = 0;
  for (const sample of samples) {
    const doc = openDocx(readFileSync(join(FIXTURES, sample)));
    for (const [name, part] of doc.parts) {
      if (part.kind !== 'xml') continue;
      total++;
      const first = buildXml(part.tree, part.meta);
      const { tree: tree2, meta: meta2 } = parseXml(first);
      const second = buildXml(tree2, meta2);
      if (first === second) stable++;
      else {
        // 输出第一个 diff 便于定位
        let i = 0;
        while (i < Math.min(first.length, second.length) && first[i] === second[i]) i++;
        assert.equal(
          second,
          first,
          `[${sample}] ${name} build 非幂等，首次 diff @${i}: ${JSON.stringify(first.slice(i, i + 30))} vs ${JSON.stringify(second.slice(i, i + 30))}`,
        );
      }
    }
  }
  assert.equal(stable, total, `${total - stable}/${total} 个 XML 部件 build 非幂等`);
});

test('dirty 部件被 build 输出替换，其余部件原样', () => {
  for (const sample of samples) {
    const doc = openDocx(readFileSync(join(FIXTURES, sample)));
    // 挑第一个 XML 部件标记 dirty
    let target = null;
    for (const [name, part] of doc.parts) {
      if (part.kind === 'xml') {
        target = name;
        break;
      }
    }
    assert.ok(target, `[${sample}] 无 XML 部件`);
    markDirty(doc, target);
    const reopened = openDocx(toBuffer(doc));
    for (const [name, part] of doc.parts) {
      const next = reopened.parts.get(name);
      assert.ok(next, `[${sample}] 部件 ${name} 丢失`);
      if (name === target) {
        // dirty 部件 = build 输出（与原始文本可不同，但必须能正常打开且结构对应）
        assert.equal(next.text, buildXml(part.tree, part.meta));
      } else if (part.kind === 'xml') {
        assert.equal(next.text, part.text, `[${sample}] 非 dirty 部件 ${name} 被改动`);
      } else {
        assert.deepEqual(Buffer.from(next.bytes), Buffer.from(part.bytes));
      }
    }
  }
});

test('roundTripXml 便捷函数保持声明/BOM/self-closing 细节', () => {
  const cases = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document/>',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document xmlns:w="x"><w:p><w:jc w:val="center"/></w:p></w:document>',
    '<?xml version="1.0" encoding="UTF-8"?><a:b><c:d a="1"/><e>&amp;</e></a:b>',
  ];
  for (const xml of cases) {
    const { tree, meta } = parseXml(xml);
    assert.equal(buildXml(tree, meta), xml);
  }
  // 带 BOM
  const bom = '﻿<?xml version="1.0"?><a/>';
  const { tree: t2, meta: m2 } = parseXml(bom);
  assert.equal(buildXml(t2, m2), bom);
  // 文本节点引号还原（TOC 字段场景）
  const toc = String.raw`<?xml version="1.0"?>
<w:instrText xml:space="preserve"> TOC \o "1-3" \h \u </w:instrText>`.replace(/\n/, '\r\n');
  const { tree: t3, meta: m3 } = parseXml(toc);
  assert.equal(buildXml(t3, m3), toc);
});
