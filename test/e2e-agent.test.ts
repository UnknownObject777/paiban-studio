// e2e-agent.test.js — 真实 LLM 链路测试（spec 测试策略：agent 用工具完成"改排版"任务并自查）。
//
// 默认跳过：需要真实模型凭证。启用方式（任选其一）：
//   PAIBAN_E2E=1 + ANTHROPIC_API_KEY=sk-...           （Anthropic 直连）
//   PAIBAN_E2E=1 + PAIBAN_BASE_URL=http://... + PAIBAN_API_KEY + PAIBAN_MODEL（OpenAI 兼容网关）
//
// 断言：agent 收到"标题改黑体三号居中"后，经由 doc_edit 工具完成修改，
// 文档产生新版本且 /body/p[1] 实际变为黑体/16pt/居中（不只听 agent 自述）。

import { test } from 'node:test';
import assert from 'node:assert/strict';
import PizZip from 'pizzip';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const SKIP = !process.env.PAIBAN_E2E;
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function makeDocx() {
  const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const zip = new PizZip();
  zip.file('[Content_Types].xml', DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', DECL + `<w:document ${W_NS}><w:body>` +
    p('测试标题') + p('这是正文段落，用于验证 agent 的排版修改是否真正落到文档里。') +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>');
  zip.file('word/_rels/document.xml.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  return zip.generate({ type: 'nodebuffer' });
}

test('agent 用工具完成"标题改黑体三号居中"并落到文档', { skip: SKIP, timeout: 180000 }, async () => {
  const { Workspace } = await import('../src/server/workspace.js');
  const { AgentBridge } = await import('../src/agent-core/bridge.js');
  const { openDocx } = await import('../src/docx-core/docx.js');
  const { findChild, findChildren, getAttr, isElement } = await import('../src/docx-core/ooxml.js');

  const dir = mkdtempSync(join(tmpdir(), 'paiban-e2e-'));
  try {
    const ws = new Workspace(dir);
    if (process.env.PAIBAN_BASE_URL || process.env.PAIBAN_MODEL) {
      ws.setConfig({
        baseUrl: process.env.PAIBAN_BASE_URL,
        model: process.env.PAIBAN_MODEL,
        apiKey: process.env.PAIBAN_API_KEY,
      });
    }
    const { docId } = ws.uploadDocument(makeDocx(), 'e2e.docx');

    const bridge = new AgentBridge(ws);
    await bridge.init();
    assert.equal(bridge.status().ready, true, `agent 未就绪: ${bridge.status().reason}`);

    const r = await bridge.send(docId, '把第一段标题改成黑体、三号字（16pt）、居中。');
    assert.equal(r.ok, true);

    // 自查：文档实际变化（不听 agent 自述）
    const versions = ws.listVersions(docId);
    assert.ok(versions.length >= 2, '应产生新版本');
    const doc = openDocx(ws.getDocumentBuffer(docId));
    const docEl = doc.parts.get('word/document.xml')!.tree!.find((n) => isElement(n, 'w:document'));
    const body = findChild(docEl, 'w:body');
    const p1 = findChildren(body, 'w:p')[0];
    const pPr = findChild(p1, 'w:pPr');
    assert.equal(getAttr(findChild(pPr, 'w:jc'), 'w:val'), 'center', '标题应居中');
    const rPr = findChild(findChild(p1, 'w:r'), 'w:rPr');
    assert.equal(getAttr(findChild(rPr, 'w:rFonts'), 'w:eastAsia'), '黑体', '标题应黑体');
    assert.equal(getAttr(findChild(rPr, 'w:sz'), 'w:val'), '32', '三号=16pt=32半磅');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
