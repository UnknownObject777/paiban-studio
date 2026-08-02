// e2e-mock-agent.test.js — mock OpenAI 兼容端点的 agent 全链路测试（默认运行，无需真实 LLM）。
//
// 本地 HTTP 服务器扮演 /v1/chat/completions：
//   第 1 次请求 → 返回 doc_edit tool_call（标题黑体三号居中）
//   第 2 次请求 → 返回纯文本完成回复
// 验证（等价真实 LLM 链路的结构行为）：
//   bridge → pi session → 工具分发 → doc_edit 执行 → 自动快照新版本 → 文档实际变化
//   → 事件流（tool_start / tool_end / text_delta / done）

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import PizZip from 'pizzip';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Workspace } from '../src/server/workspace.js';
import { AgentBridge } from '../src/agent-core/bridge.js';
import { openDocx } from '../src/docx-core/docx.js';
import { findChild, findChildren, getAttr, isElement } from '../src/docx-core/ooxml.js';

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

function makeDocx() {
  const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const zip = new PizZip();
  zip.file('[Content_Types].xml', DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', DECL + `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    p('测试标题') + p('这是正文段落。') +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>');
  zip.file('word/_rels/document.xml.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  return zip.generate({ type: 'nodebuffer' });
}

// 判断请求是否为流式
function wantsStream(body: string): boolean {
  try { return JSON.parse(body).stream === true; } catch { return false; }
}

function sse(chunks: unknown[]): string {
  return chunks.map((c) => `data: ${JSON.stringify(c)}\n\n`).join('') + 'data: [DONE]\n\n';
}

async function startMockLlm() {
  let calls = 0;
  const server = createServer((req, res) => {
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => {
      calls++;
      const streaming = wantsStream(body);
      // 从 prompt 注入的上下文提取真实 docId（bridge 会附带 "[当前工作文档 docId: xxx]"）
      const docId = /docId: ([0-9a-f-]+)/.exec(body)?.[1] || 'unknown';
      if (calls === 1) {
        // 第一次：要求调用 doc_edit
        const toolCall = {
          id: 'call_1', type: 'function', index: 0,
          function: {
            name: 'doc_edit',
            arguments: JSON.stringify({
              docId,
              commands: [{ command: 'set', path: '/body/p[1]', props: { align: 'center', run: { eastAsia: '黑体', sizePt: 16 } } }],
              note: '标题黑体三号居中',
            }),
          },
        };
        if (streaming) {
          res.writeHead(200, { 'content-type': 'text/event-stream' });
          res.end(sse([
            { id: 'c1', object: 'chat.completion.chunk', choices: [{ index: 0, delta: { role: 'assistant', tool_calls: [toolCall] }, finish_reason: null }] },
            { id: 'c1', object: 'chat.completion.chunk', choices: [{ index: 0, delta: {}, finish_reason: 'tool_calls' }], usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 } },
          ]));
        } else {
          res.writeHead(200, { 'content-type': 'application/json' });
          res.end(JSON.stringify({
            id: 'c1', object: 'chat.completion',
            choices: [{ index: 0, message: { role: 'assistant', content: null, tool_calls: [toolCall] }, finish_reason: 'tool_calls' }],
            usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
          }));
        }
      } else {
        // 第二次（带工具结果）：文本完成
        if (streaming) {
          res.writeHead(200, { 'content-type': 'text/event-stream' });
          res.end(sse([
            { id: 'c2', object: 'chat.completion.chunk', choices: [{ index: 0, delta: { role: 'assistant', content: '已完成：标题改为黑体三号居中。' }, finish_reason: null }] },
            { id: 'c2', object: 'chat.completion.chunk', choices: [{ index: 0, delta: {}, finish_reason: 'stop' }], usage: { prompt_tokens: 20, completion_tokens: 8, total_tokens: 28 } },
          ]));
        } else {
          res.writeHead(200, { 'content-type': 'application/json' });
          res.end(JSON.stringify({
            id: 'c2', object: 'chat.completion',
            choices: [{ index: 0, message: { role: 'assistant', content: '已完成：标题改为黑体三号居中。' }, finish_reason: 'stop' }],
            usage: { prompt_tokens: 20, completion_tokens: 8, total_tokens: 28 },
          }));
        }
      }
    });
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', () => resolve()));
  const addr = server.address() as AddressInfo;
  return { server, port: addr.port, getCalls: () => calls };
}

test('mock LLM 全链路：doc_edit 工具 → 文档实际变化 + 新版本 + 事件流', { timeout: 120000 }, async () => {
  const mock = await startMockLlm();
  const dir = mkdtempSync(join(tmpdir(), 'paiban-mock-'));
  try {
    const ws = new Workspace(dir);
    const { docId } = ws.uploadDocument(makeDocx(), 'mock.docx');
    ws.setConfig({
      provider: 'openai', model: 'mock-model',
      baseUrl: `http://127.0.0.1:${mock.port}/v1`, apiKey: 'dummy',
    });

    const bridge = new AgentBridge(ws);
    await bridge.init();
    assert.equal(bridge.status().ready, true, `agent 未就绪: ${bridge.status().reason}`);

    // 收集事件
    const events: Array<{ type: string; [k: string]: any }> = [];
    bridge.onEvent((e) => events.push(e as { type: string; [k: string]: any }));

    const r = await bridge.send(docId, '把第一段标题改成黑体三号居中。');
    assert.equal(r.ok, true);

    // 文档实际变化
    const versions = ws.listVersions(docId);
    assert.ok(versions.length >= 2, '应产生新版本');
    assert.equal(versions[versions.length - 1].note, '标题黑体三号居中');
    const doc = openDocx(ws.getDocumentBuffer(docId));
    const docEl = doc.parts.get('word/document.xml')!.tree!.find((n) => isElement(n, 'w:document'));
    const body = findChild(docEl, 'w:body');
    const p1 = findChildren(body, 'w:p')[0];
    assert.equal(getAttr(findChild(findChild(p1, 'w:pPr'), 'w:jc'), 'w:val'), 'center');
    const rPr = findChild(findChild(p1, 'w:r'), 'w:rPr');
    assert.equal(getAttr(findChild(rPr, 'w:rFonts'), 'w:eastAsia'), '黑体');
    assert.equal(getAttr(findChild(rPr, 'w:sz'), 'w:val'), '32');

    // 事件流：tool_start → tool_end → done
    const types = events.map((e) => e.type);
    assert.ok(types.includes('tool_start'), `缺 tool_start: ${types}`);
    assert.ok(types.includes('tool_end'), `缺 tool_end: ${types}`);
    assert.ok(types.includes('done'), `缺 done: ${types}`);
    const toolStart = events.find((e) => e.type === 'tool_start')!;
    assert.equal(toolStart.name, 'doc_edit');
  } finally {
    mock.server.close();
    rmSync(dir, { recursive: true, force: true });
  }
});
