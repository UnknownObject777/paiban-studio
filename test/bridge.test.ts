// bridge.test.ts — agent 接入层（bridge.ts）白名单与事件摘要单测（本切片：doc_generate / template_instantiate 接线）。
//
// 覆盖：
//   TOOL_WHITELIST：含 doc_generate / template_instantiate / amount_words，且与 createTools 工具集一致（防 session 白名单漏接）
//   summarizeArgs / summarizeToolResult 对 doc_generate / template_instantiate 的摘要
//   回归：doc_edit / ruleset_read 的摘要不受新分支影响
// 不依赖 LLM / SDK（bridge.ts 仅在 init() 内动态 import SDK）。

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { TOOL_WHITELIST, summarizeArgs, summarizeToolResult } from '../src/agent-core/bridge.js';
import { createTools } from '../src/agent-core/tools.js';

test('工具白名单：包含 doc_generate / template_instantiate / amount_words，且与 createTools 工具集一致', () => {
  assert.ok(TOOL_WHITELIST.includes('doc_generate'), '白名单缺 doc_generate（session 将不加载该工具）');
  assert.ok(TOOL_WHITELIST.includes('template_instantiate'), '白名单缺 template_instantiate（session 将不加载该工具）');
  assert.ok(TOOL_WHITELIST.includes('amount_words'), '白名单缺 amount_words（session 将不加载该工具）');
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

test('summarizeArgs：template_instantiate 摘要为「实例化模板 <templateId>（N 个占位符）」，不塞入 values 正文', () => {
  const s = summarizeArgs({ templateId: 'tpl-abc123', values: { 项目名称: '智慧园区', 项目编号: 'ZB-2026-001' }, name: '投标文件.docx' });
  assert.equal(s, '实例化模板 tpl-abc123（2 个占位符）');
  assert.ok(!s.includes('智慧园区'), 'values 正文不应塞进摘要');
  // 未带 values 的 templateId（template_read）仍走原分支
  assert.equal(summarizeArgs({ templateId: 'tpl-abc123' }), 'tpl-abc123');
});

test('summarizeToolResult：template_instantiate 结果摘要为文档名 + v1 + 替换/未匹配数', () => {
  const s = summarizeToolResult({
    result: { details: { docId: 'd1', name: '投标文件（智慧园区）.docx', version: { id: 'v1' }, replaced: [{}, {}], errors: [] } },
  });
  assert.equal(s, '实例化《投标文件（智慧园区）.docx》 · v1，替换 2 处');
  // 有未匹配占位符时带出错误数
  const s2 = summarizeToolResult({
    result: { details: { docId: 'd2', name: 'x.docx', version: { id: 'v1' }, replaced: [], errors: [{ code: 'x' }] } },
  });
  assert.equal(s2, '实例化《x.docx》 · v1，替换 0 处，1 处未匹配');
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

test('summarizeArgs：amount_words 摘要显示金额本身', () => {
  assert.equal(summarizeArgs({ amount: '1,234.5' }), '金额 1,234.5');
  // 金额为字符串 "0"（falsy）时仍要显示，不能落入 fallback
  assert.equal(summarizeArgs({ amount: '0' }), '金额 0');
});
