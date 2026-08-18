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
import { findChild, findChildren, getAttr, isElement, textOf } from '../src/docx-core/ooxml.js';

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

// ---- doc_generate 场景（landing 一键生成，无 docId） ----
//
// 独立 mock 端点 + 独立 workspace/bridge 实例：第 1 次请求返回 doc_generate tool_call
// （参数固定，无 docId），第 2 次返回纯文本完成。与上方 doc_edit 场景互不干扰。

async function startMockGenerateLlm() {
  // tool_call.arguments：与任务给定 JSON 逐字节一致（JSON.stringify 转义 \n，键序 markdown/rulesetId/name）
  const generateArgs = JSON.stringify({
    markdown: '# 投标文件\n\n## 一、投标函\n\n我方响应招标文件全部条款。\n\n| 序号 | 分项 | 金额 |\n| --- | --- | --- |\n| 1 | 软件开发费 | ××× |',
    rulesetId: 'bid-default',
    name: 'mock标书.docx',
  });
  let calls = 0;
  const server = createServer((req, res) => {
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => {
      calls++;
      const streaming = wantsStream(body);
      if (calls === 1) {
        // 第一次：要求调用 doc_generate（从零生成，无需 docId）
        const toolCall = {
          id: 'call_g1', type: 'function', index: 0,
          function: { name: 'doc_generate', arguments: generateArgs },
        };
        if (streaming) {
          res.writeHead(200, { 'content-type': 'text/event-stream' });
          res.end(sse([
            { id: 'g1', object: 'chat.completion.chunk', choices: [{ index: 0, delta: { role: 'assistant', tool_calls: [toolCall] }, finish_reason: null }] },
            { id: 'g1', object: 'chat.completion.chunk', choices: [{ index: 0, delta: {}, finish_reason: 'tool_calls' }], usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 } },
          ]));
        } else {
          res.writeHead(200, { 'content-type': 'application/json' });
          res.end(JSON.stringify({
            id: 'g1', object: 'chat.completion',
            choices: [{ index: 0, message: { role: 'assistant', content: null, tool_calls: [toolCall] }, finish_reason: 'tool_calls' }],
            usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
          }));
        }
      } else {
        // 第二次（带工具结果）：纯文本完成回复
        const doneText = '已完成：已生成《mock标书.docx》并入库（v1）。';
        if (streaming) {
          res.writeHead(200, { 'content-type': 'text/event-stream' });
          res.end(sse([
            { id: 'g2', object: 'chat.completion.chunk', choices: [{ index: 0, delta: { role: 'assistant', content: doneText }, finish_reason: null }] },
            { id: 'g2', object: 'chat.completion.chunk', choices: [{ index: 0, delta: {}, finish_reason: 'stop' }], usage: { prompt_tokens: 20, completion_tokens: 8, total_tokens: 28 } },
          ]));
        } else {
          res.writeHead(200, { 'content-type': 'application/json' });
          res.end(JSON.stringify({
            id: 'g2', object: 'chat.completion',
            choices: [{ index: 0, message: { role: 'assistant', content: doneText }, finish_reason: 'stop' }],
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

test('mock LLM 全链路：doc_generate 工具 → 一键生成新文档（无 docId）+ v1 + 事件流', { timeout: 120000 }, async () => {
  const mock = await startMockGenerateLlm();
  const dir = mkdtempSync(join(tmpdir(), 'paiban-mock-gen-'));
  try {
    const ws = new Workspace(dir);
    ws.setConfig({
      provider: 'openai', model: 'mock-model',
      baseUrl: `http://127.0.0.1:${mock.port}/v1`, apiKey: 'dummy',
    });

    const bridge = new AgentBridge(ws);
    await bridge.init();
    assert.equal(bridge.status().ready, true, `agent 未就绪: ${bridge.status().reason}`);

    const events: Array<{ type: string; [k: string]: any }> = [];
    bridge.onEvent((e) => events.push(e as { type: string; [k: string]: any }));

    // landing 一键生成路径：无 docId
    const r = await bridge.send(undefined, '帮我起草一份投标文件');
    assert.equal(r.ok, true);

    // 新工作文档入库（name 与 tool_call 参数一致）
    const docs = ws.listDocuments();
    const doc = docs.find((d) => d.name === 'mock标书.docx');
    assert.ok(doc, `listDocuments 应含 mock标书.docx，实际: ${docs.map((d) => d.name).join(', ')}`);
    const docId = doc!.docId;

    // v1 版本存在
    const versions = ws.listVersions(docId);
    assert.ok(versions.some((v) => v.id === 'v1'), `应含 v1，实际: ${versions.map((v) => v.id).join(', ')}`);
    assert.equal(versions[0].source, 'generate');

    // 产物 buffer 重解析：document.xml 含"投标函"文本 + w:tbl 表格
    const parsed = openDocx(ws.getDocumentBuffer(docId));
    const docEl = parsed.parts.get('word/document.xml')!.tree!.find((n) => isElement(n, 'w:document'));
    const body = findChild(docEl, 'w:body');
    assert.ok(textOf(body).includes('投标函'), 'document.xml 应含"投标函"文本');
    assert.ok(findChildren(body, 'w:tbl').length >= 1, 'document.xml 应含 w:tbl 表格');

    // 事件流：tool_start → tool_end（details.docId 与文档一致）→ done
    const types = events.map((e) => e.type);
    assert.ok(types.includes('tool_start'), `缺 tool_start: ${types}`);
    assert.ok(types.includes('tool_end'), `缺 tool_end: ${types}`);
    assert.ok(types.includes('done'), `缺 done: ${types}`);
    const toolStart = events.find((e) => e.type === 'tool_start')!;
    assert.equal(toolStart.name, 'doc_generate');
    const toolEnd = events.find((e) => e.type === 'tool_end')!;
    assert.equal(toolEnd.name, 'doc_generate');
    assert.equal(toolEnd.isError, false);
    assert.ok(toolEnd.details, 'tool_end 应带 details 字段');
    assert.equal(toolEnd.details.docId, docId, 'tool_end.details.docId 应与实际 docId 一致');
    assert.equal(toolEnd.details.name, 'mock标书.docx');
    assert.equal(toolEnd.details.version.id, 'v1');
  } finally {
    mock.server.close();
    rmSync(dir, { recursive: true, force: true });
  }
});
