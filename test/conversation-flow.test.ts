// conversation-flow.test.ts — 对话流状态机（issue #31）headless 测试。
//
// 覆盖规格 Testing Decisions 的全部点位：
//   landing → editing 转换；瀑布流条目序列（text_delta 累积 / tool_start 截断 / tool_end 回填）；
//   进度模型推导（步骤链顺序与标签映射、状态迁移、done 收尾、error 置失败态、abort 复位）；
//   无文档发送直通（doc_generate 支持从零生成新文档）。
// 只测状态机对事件序列的输出（外部行为），不测 DOM。

import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  createFlow,
  reduce,
  deriveProgress,
  TOOL_LABELS,
  type ChatItem,
} from '../src/ui/conversation-flow.js';

function kinds(items: ChatItem[]): string[] {
  return items.map((i) => i.kind);
}

// ---- 初始态与相位转换 ----

test('初始态：landing、无文档、不忙、无条目', () => {
  const s = createFlow();
  assert.equal(s.phase, 'landing');
  assert.equal(s.docId, null);
  assert.equal(s.busy, false);
  assert.deepEqual(s.items, []);
});

test('doc_opened：进入 editing 并落下版本卡', () => {
  const s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: '通知.docx', versionId: 'v1' });
  assert.equal(s.phase, 'editing');
  assert.equal(s.docId, 'd1');
  assert.equal(s.docName, '通知.docx');
  assert.deepEqual(kinds(s.items), ['version']);
  assert.equal((s.items[0] as any).versionId, 'v1');
});

test('doc_opened：note 覆盖版本卡文案（doc_generate 自动打开新文档用）', () => {
  const s = reduce(createFlow(), { type: 'doc_opened', docId: 'g1', name: '投标文件.docx', versionId: 'v1', note: '已生成《投标文件.docx》' });
  assert.equal((s.items[0] as any).note, '已生成《投标文件.docx》');
});

test('back_to_landing：回到 landing，条目与文档上下文保留', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: '通知.docx', versionId: 'v1' });
  s = reduce(s, { type: 'back_to_landing' });
  assert.equal(s.phase, 'landing');
  assert.equal(s.docId, 'd1'); // 回到 landing 后可继续对同一文档发指令
  assert.equal(s.items.length, 1);
});

// ---- 发送 ----

test('无文档时 send：直接进入发送（doc_generate 支持从零生成），落用户消息并置忙', () => {
  let s = reduce(createFlow(), { type: 'send', text: '帮我起草一份投标文件，项目名称为×××，按 bid-default 规则集生成' });
  assert.equal(s.phase, 'landing');
  assert.equal(s.busy, true);
  const last = s.items[s.items.length - 1];
  assert.equal(last.kind, 'user');
  assert.equal((last as any).text, '帮我起草一份投标文件，项目名称为×××，按 bid-default 规则集生成');
});

test('有文档时 send：落用户消息并置忙；空文本为无操作', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx', versionId: 'v1' });
  s = reduce(s, { type: 'send', text: '标题改成黑体三号居中' });
  assert.equal(s.busy, true);
  const last = s.items[s.items.length - 1];
  assert.equal(last.kind, 'user');
  assert.equal((last as any).text, '标题改成黑体三号居中');
  const same = reduce(s, { type: 'send', text: '   ' });
  assert.equal(same.items.length, s.items.length);
});

// ---- 瀑布流条目序列 ----

test('text_delta 累积进同一条 assistant 消息', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'text_delta', delta: '好的' });
  s = reduce(s, { type: 'text_delta', delta: '，正在处理' });
  const assistants = s.items.filter((i) => i.kind === 'assistant');
  assert.equal(assistants.length, 1);
  assert.equal((assistants[0] as any).text, '好的，正在处理');
  assert.equal((assistants[0] as any).streaming, true);
});

test('tool_start 截断当前 assistant 气泡并生成运行中步骤卡；其后的 text_delta 开新气泡', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'text_delta', delta: '先看一下结构' });
  s = reduce(s, { type: 'tool_start', name: 'doc_outline', args: '' });
  s = reduce(s, { type: 'text_delta', delta: '结构已读取' });
  assert.deepEqual(kinds(s.items), ['version', 'user', 'assistant', 'tool', 'assistant']);
  const [a1, tool, a2] = s.items.slice(2);
  assert.equal((a1 as any).streaming, false); // 被截断收尾
  assert.equal((tool as any).status, 'running');
  assert.equal((a2 as any).text, '结构已读取');
});

test('tool_end 回填最近一个运行中的步骤卡（含失败态）', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'tool_start', name: 'doc_outline', args: '' });
  s = reduce(s, { type: 'tool_end', name: 'doc_outline', isError: false, summary: '42 段' });
  s = reduce(s, { type: 'tool_start', name: 'doc_edit', args: '3 条命令' });
  s = reduce(s, { type: 'tool_end', name: 'doc_edit', isError: true, summary: '路径不存在' });
  const tools = s.items.filter((i) => i.kind === 'tool') as any[];
  assert.equal(tools.length, 2);
  assert.equal(tools[0].status, 'done');
  assert.equal(tools[0].summary, '42 段');
  assert.equal(tools[1].status, 'error');
  assert.equal(tools[1].summary, '路径不存在');
});

test('version_note：UI 动作落版本卡（上传/回滚/实例化/规则集重排）', () => {
  const s = reduce(createFlow(), { type: 'version_note', versionId: 'v3', note: '回滚到 v3' });
  assert.deepEqual(kinds(s.items), ['version']);
  assert.equal((s.items[0] as any).versionId, 'v3');
});

// ---- 收尾与异常 ----

test('done：收尾——闲下来、流式气泡封口、残留运行中步骤关闭', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'text_delta', delta: '完成' });
  s = reduce(s, { type: 'tool_start', name: 'version_store', args: 'list' });
  s = reduce(s, { type: 'done' });
  assert.equal(s.busy, false);
  const a = s.items.find((i) => i.kind === 'assistant') as any;
  assert.equal(a.streaming, false);
  const t = s.items.find((i) => i.kind === 'tool') as any;
  assert.equal(t.status, 'done');
});

test('error 事件：落错误条目、置失败态、停忙', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'tool_start', name: 'doc_edit', args: '' });
  s = reduce(s, { type: 'error', message: '模型超时' });
  assert.equal(s.busy, false);
  assert.deepEqual(kinds(s.items), ['version', 'user', 'tool', 'error']);
  const t = s.items.find((i) => i.kind === 'tool') as any;
  assert.equal(t.status, 'error');
});

test('abort：停忙、封口流式气泡、运行中步骤标记为已停止', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'tool_start', name: 'doc_edit', args: '' });
  s = reduce(s, { type: 'abort' });
  assert.equal(s.busy, false);
  const t = s.items.find((i) => i.kind === 'tool') as any;
  assert.equal(t.status, 'done');
  assert.match(t.summary, /停止/);
});

// ---- 进度模型 ----

test('进度：工具名映射为中文步骤标签，链序保留', () => {
  assert.equal(TOOL_LABELS.doc_outline, '分析文档结构');
  assert.equal(TOOL_LABELS.doc_edit, '应用排版修改');
  assert.equal(TOOL_LABELS.ruleset_read, '读取内置规则集');
  assert.equal(TOOL_LABELS.template_read, '读取模板');
  assert.equal(TOOL_LABELS.version_store, '保存 / 回滚版本');

  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: '按公文排版' });
  s = reduce(s, { type: 'tool_start', name: 'ruleset_read', args: 'gongwen-default' });
  s = reduce(s, { type: 'tool_end', name: 'ruleset_read', isError: false, summary: '' });
  s = reduce(s, { type: 'tool_start', name: 'doc_edit', args: '5 条命令' });
  const p = deriveProgress(s);
  assert.equal(p.active, true);
  assert.deepEqual(p.steps.map((x) => x.label), ['读取内置规则集', '应用排版修改']);
  assert.deepEqual(p.steps.map((x) => x.status), ['done', 'running']);
  assert.equal(p.currentLabel, '应用排版修改');
});

test('进度：无运行中步骤时，流式输出显示「撰写回复」，否则「思考中」', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  assert.equal(deriveProgress(s).currentLabel, '思考中…');
  s = reduce(s, { type: 'text_delta', delta: '好' });
  assert.equal(deriveProgress(s).currentLabel, '撰写回复…');
});

test('进度：done 后不再 active 但步骤链保留可回顾；error 计入失败数', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'tool_start', name: 'doc_edit', args: '' });
  s = reduce(s, { type: 'tool_end', name: 'doc_edit', isError: true, summary: '失败原因' });
  s = reduce(s, { type: 'done' });
  const p = deriveProgress(s);
  assert.equal(p.active, false);
  assert.equal(p.steps.length, 1);
  assert.equal(p.failedCount, 1);
});

test('进度：空闲且无任何步骤时不可见', () => {
  const p = deriveProgress(createFlow());
  assert.equal(p.visible, false);
  assert.equal(p.active, false);
});

// ---- 健壮性 ----

test('轮次外的迟到 tool_end 直接忽略；轮次内无匹配则兜底落成卡', () => {
  let s = createFlow();
  s = reduce(s, { type: 'tool_end', name: 'doc_edit', isError: false, summary: 'v2' });
  assert.equal(s.items.length, 0); // 未在忙：忽略

  s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'tool_end', name: 'doc_edit', isError: false, summary: 'v2' });
  assert.deepEqual(kinds(s.items), ['version', 'user', 'tool']);
  assert.equal((s.items[2] as any).status, 'done');
});

test('abort 后迟到的 tool_end 回填被停止的原卡，不落重复卡', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'tool_start', name: 'doc_edit', args: '' });
  s = reduce(s, { type: 'abort' });
  s = reduce(s, { type: 'tool_end', name: 'doc_edit', isError: false, summary: 'v2（新版本）' });
  const tools = s.items.filter((i) => i.kind === 'tool') as any[];
  assert.equal(tools.length, 1); // 没有重复卡
  assert.equal(tools[0].status, 'done');
  assert.equal(tools[0].summary, 'v2（新版本）');
});

test('reduce 不可变：返回新 state，旧 state 不被改写', () => {
  const s0 = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  const n0 = s0.items.length;
  const s1 = reduce(s0, { type: 'send', text: 'x' });
  assert.equal(s0.items.length, n0);
  assert.ok(s1.items.length > n0);
  assert.notEqual(s0, s1);
});

// ---- 轮次与结果态（code-review 修复） ----

test('新一轮 send 后进度链只统计本轮步骤', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: '第一轮' });
  s = reduce(s, { type: 'tool_start', name: 'doc_outline', args: '' });
  s = reduce(s, { type: 'tool_end', name: 'doc_outline', isError: false, summary: '42 段' });
  s = reduce(s, { type: 'done' });
  assert.equal(deriveProgress(s).steps.length, 1);
  s = reduce(s, { type: 'send', text: '第二轮' });
  const p = deriveProgress(s);
  assert.equal(p.active, true);
  assert.equal(p.steps.length, 0); // 旧轮步骤不再残留
  assert.equal(p.currentLabel, '思考中…');
});

test('结果态文案：done / aborted / error（含无步骤出错仍可见）', () => {
  let s = reduce(createFlow(), { type: 'doc_opened', docId: 'd1', name: 'a.docx' });
  s = reduce(s, { type: 'send', text: 'x' });
  s = reduce(s, { type: 'done' });
  assert.equal(deriveProgress(s).finalLabel, '完成 · 共 0 步');

  s = reduce(s, { type: 'send', text: 'y' });
  s = reduce(s, { type: 'abort' });
  const pa = deriveProgress(s);
  assert.equal(pa.finalLabel, '已停止 · 共 0 步');
  assert.equal(pa.active, false);

  // 无工具步骤直接出错：进度条仍显示失败态
  s = reduce(s, { type: 'send', text: 'z' });
  s = reduce(s, { type: 'error', message: '模型超时' });
  const pe = deriveProgress(s);
  assert.equal(pe.visible, true);
  assert.equal(pe.finalLabel, '出错（详见对话流）');
});

test('note 卡：非版本事件的居中便签（模板解析 / 配置保存等）', () => {
  const s = reduce(createFlow(), { type: 'note', title: '配置', note: '模型配置已保存，重启后生效' });
  assert.deepEqual(kinds(s.items), ['note']);
  assert.equal((s.items[0] as any).title, '配置');
});
