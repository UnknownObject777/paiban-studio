// bridge.test.ts — agent 接入层（bridge.ts）白名单与事件摘要单测（本切片：doc_generate 接线）。
//
// 覆盖：
//   TOOL_WHITELIST：含 doc_generate，且与 createTools 六工具一致（防 session 白名单漏接）
//   summarizeArgs / summarizeToolResult 对 doc_generate 的摘要
//   回归：doc_edit / ruleset_read 的摘要不受新分支影响
// 不依赖 LLM / SDK（bridge.ts 仅在 init() 内动态 import SDK）。

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { TOOL_WHITELIST, summarizeArgs, summarizeToolResult } from '../src/agent-core/bridge.js';
import { createTools } from '../src/agent-core/tools.js';

test('工具白名单：包含 doc_generate，且与 createTools 六工具一致', () => {
  assert.ok(TOOL_WHITELIST.includes('doc_generate'), '白名单缺 doc_generate（session 将不加载该工具）');
  assert.deepEqual(TOOL_WHITELIST, createTools({} as never).map((t) => t.name));
});

test('summarizeArgs：doc_generate 摘要为「生成《name》（规则集 xxx，markdown N 行）」，不塞入 markdown 正文', () => {
  const s = summarizeArgs({ markdown: '# 投标文件\n## 投标函\n正文', rulesetId: 'bid-default', name: '投标文件.docx' });
  assert.ok(s.includes('生成《投标文件.docx》'));
  assert.ok(s.includes('bid-default'));
  assert.ok(s.includes('3 行'));
  assert.ok(!s.includes('# 投标文件'), 'markdown 正文不应塞进摘要');
  // 未传 name 的兜底
  const s2 = summarizeArgs({ markdown: '## 实验目的', rulesetId: 'lab-report-default' });
  assert.ok(s2.includes('lab-report-default'));
});

test('summarizeToolResult：doc_generate 结果摘要为文档名 + v1', () => {
  const s = summarizeToolResult({
    result: { details: { docId: 'g1', name: '投标文件.docx', version: { id: 'v1' } } },
  });
  assert.equal(s, '《投标文件.docx》 · v1');
});

test('摘要回归：doc_edit 的 args/result 摘要保持原样；ruleset_read 不误入 doc_generate 分支', () => {
  assert.ok(summarizeArgs({ commands: [{ command: 'set' }, { command: 'set' }], note: '居中' }).includes('2 条命令'));
  const r = summarizeToolResult({
    result: { details: { version: { id: 'v2' }, versionCreated: true, applied: [1], errors: [] } },
  });
  assert.ok(r.includes('v2（新版本）'));
  assert.ok(r.includes('applied 1'));
  // ruleset_read 只传 rulesetId（无 markdown）：不走 doc_generate 分支
  assert.equal(summarizeArgs({ rulesetId: 'gongwen-default' }), 'rulesetId');
});
