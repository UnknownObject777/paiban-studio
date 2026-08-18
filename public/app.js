// app.js — 两相交互前端（issue #31：landing 居中对话窗 → 编辑态「对话左移 + 大预览」）。
// ES module：交互状态全部收敛到 src/ui/conversation-flow.ts 状态机（经 dist 编译产物 import），
// 本文件只做 DOM 渲染与 IPC 接线。全部后端能力经 window.paiban（preload 白名单）。

import { createFlow, reduce, deriveProgress } from '../dist/src/ui/conversation-flow.js';

const $ = (sel) => document.querySelector(sel);

let flow = createFlow();

const ui = {
  previewReady: false,
  renderTimer: null,
  itemNodes: new Map(), // ChatItem.id → DOM 节点（瀑布流增量渲染）
};

function dispatch(event) {
  flow = reduce(flow, event);
  render();
}

// ---- 渲染入口 ----

function render() {
  renderPhase();
  renderItems();
  renderProgress();
  renderBusy();
}

function renderPhase() {
  const app = $('#app');
  const editing = flow.phase === 'editing';
  if (app.dataset.phase !== flow.phase) {
    app.dataset.phase = flow.phase;
    $('#landing').classList.toggle('hidden', editing);
    $('#workbench').classList.toggle('hidden', !editing);
    // 对话窗随相位迁移：landing 居中 ↔ 编辑态列底
    $(editing ? '#editing-input-slot' : '#landing-input-slot').appendChild($('#chat-form'));
  }
  if (editing) {
    $('#current-doc-name').textContent = flow.docName || '';
    $('#current-doc-name').title = flow.docName || '';
  }
  $('#chat-input').placeholder = editing ? '对当前文档说点什么…' : '对 AI 说想怎么排；没开文档也能直接说「生成一份投标文件」…';
}

// ---- 瀑布流条目（增量：只增不删，按 id 复用节点） ----

function renderItems() {
  const stream = $('#chat-stream');
  for (const item of flow.items) {
    let node = ui.itemNodes.get(item.id);
    if (!node) {
      node = createItemNode(item);
      ui.itemNodes.set(item.id, node);
      stream.appendChild(node);
    }
    updateItemNode(node, item);
  }
  // 只在内容增长时吸底
  const last = flow.items[flow.items.length - 1];
  if (last) ui.itemNodes.get(last.id)?.scrollIntoView({ block: 'end' });
}

function createItemNode(item) {
  const div = document.createElement('div');
  if (item.kind === 'user' || item.kind === 'assistant' || item.kind === 'error') {
    div.className = `msg ${item.kind}`;
  } else if (item.kind === 'tool') {
    div.className = 'tool-chip';
  } else if (item.kind === 'note') {
    div.className = 'note-chip';
  } else if (item.kind === 'version') {
    div.className = 'version-chip';
  }
  return div;
}

function updateItemNode(node, item) {
  if (item.kind === 'user' || item.kind === 'error') {
    node.textContent = item.kind === 'error' ? '出错了：' + item.message : item.text;
  } else if (item.kind === 'assistant') {
    node.textContent = item.text;
    node.classList.toggle('streaming', item.streaming);
  } else if (item.kind === 'tool') {
    node.className = `tool-chip ${item.status}`;
    const tail = item.status === 'running' ? (item.args || '')
      : [item.summary || '', item.status === 'done' ? '✓' : '✗'].filter(Boolean).join(' ');
    node.innerHTML = `<span class="dot"></span><span>${escapeHtml(item.label)}</span><span class="dim">${escapeHtml(tail)}</span>`;
  } else if (item.kind === 'version') {
    node.textContent = `已存版本 ${item.versionId}${item.note ? ' · ' + item.note : ''}`;
  } else if (item.kind === 'note') {
    node.textContent = `${item.title} · ${item.note}`;
  }
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ---- 进度条（用户可感知的工作进度，全部由状态机推导） ----

function renderProgress() {
  const p = deriveProgress(flow);
  const bar = $('#progress-bar');
  bar.classList.toggle('hidden', !p.visible);
  if (!p.visible) return;
  bar.classList.toggle('active', p.active);
  $('#progress-label').textContent = p.active ? p.currentLabel : p.finalLabel;
  const ol = $('#progress-steps');
  ol.innerHTML = '';
  for (const s of p.steps) {
    const li = document.createElement('li');
    li.className = `progress-step ${s.status}`;
    li.textContent = s.label + (s.status === 'error' && s.summary ? `（${s.summary}）` : '');
    ol.appendChild(li);
  }
}

function renderBusy() {
  $('#btn-send').classList.toggle('hidden', flow.busy);
  $('#btn-abort').classList.toggle('hidden', !flow.busy);
  $('#chat-input').disabled = flow.busy;
  // 无文档也可发送：doc_generate 支持一句话从零生成新文档（无需先打开文档）
  $('#btn-send').disabled = false;
}

// ---- 预览（防抖刷新，D4：编辑 → ArrayBuffer → iframe） ----

function refreshPreview() {
  if (!flow.docId) return;
  clearTimeout(ui.renderTimer);
  ui.renderTimer = setTimeout(async () => {
    const buffer = await window.paiban.getBuffer(flow.docId);
    const frame = $('#preview-frame');
    $('#preview-placeholder').classList.add('hidden');
    frame.classList.remove('hidden');
    frame.contentWindow.postMessage({ type: 'render', buffer }, '*', [buffer]);
  }, 250); // 连续编辑合并
}

window.addEventListener('message', (ev) => {
  if (ev.data?.type === 'preview-ready') ui.previewReady = true;
});

// ---- agent 事件流 → 状态机（流式瀑布 + 进度推导） ----

window.paiban.onAgentEvent((event) => {
  dispatch(event);
  if (event.type === 'error') showLandingError(event.message);
  if (event.type === 'tool_end' && (event.name === 'doc_edit' || event.name === 'version_store') && !event.isError) {
    refreshPreview();
    loadVersions();
  }
  if (event.type === 'tool_end' && event.name === 'doc_generate' && !event.isError && event.details?.docId) {
    // 生成成功：立即选中新文档并加载预览（让用户马上看到产物）
    openGeneratedDocument(event.details);
  }
  if (event.type === 'done') {
    refreshPreview();
    loadVersions();
    loadDocuments();
  }
});

// ---- 对话发送 / 中断 ----

// landing 相位没有对话流渲染，agent 错误需就地可见（未就绪/发送失败）
function showLandingError(message) {
  if (flow.phase !== 'landing' || !message) return;
  const el = $('#landing-error');
  el.textContent = message;
  el.classList.remove('hidden');
}
function clearLandingError() {
  $('#landing-error').classList.add('hidden');
}

$('#chat-form').addEventListener('submit', async (ev) => {
  ev.preventDefault();
  const input = $('#chat-input');
  const text = input.value.trim();
  if (!text || flow.busy) return;
  clearLandingError();
  input.value = '';
  dispatch({ type: 'send', text });
  // docId 可为空：doc_generate 支持从零生成新文档（无需先打开文档）
  const r = await window.paiban.agentSend(flow.docId, text);
  if (!r.ok) {
    dispatch({ type: 'error', message: r.error || 'agent 发送失败' });
    showLandingError(r.error || 'agent 发送失败');
  }
});

$('#btn-abort').addEventListener('click', () => {
  window.paiban.agentAbort();
  dispatch({ type: 'abort' });
});

// 空态示例指令 chip：点击填入输入框并聚焦
document.querySelectorAll('.suggest-chip').forEach((chip) => {
  chip.addEventListener('click', () => {
    const input = $('#chat-input');
    input.value = chip.dataset.prompt || '';
    input.focus();
  });
});

// ---- 相位切换 ----

$('#btn-back').addEventListener('click', () => {
  dispatch({ type: 'back_to_landing' });
  loadDocuments(); // 刷新 landing 的最近文档
});

// 窄屏 对话/预览 tab 切换
document.querySelectorAll('.view-tab').forEach((tab) => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.view-tab').forEach((t) => t.classList.toggle('active', t === tab));
    $('#workbench').dataset.tab = tab.dataset.tab;
  });
});

// ---- 文档 ----

async function openDocFlow(openResult) {
  if (!openResult) return;
  dispatch({ type: 'doc_opened', docId: openResult.docId, name: openResult.name ?? openResult.docId, versionId: openResult.version?.id ?? openResult.head, note: openResult.note });
  $('#btn-download').disabled = false;
  await Promise.all([loadDocuments(), loadVersions()]);
  refreshPreview();
}

// doc_generate 成功：刷新文档列表并选中新文档（openDocFlow 内部已刷新列表 + 加载预览）
async function openGeneratedDocument(details) {
  const name = details.name ?? details.docId;
  await openDocFlow({ docId: details.docId, name, versionId: details.version?.id, note: `已生成《${name}》` });
}

async function openDocViaDialog() {
  const r = await window.paiban.openDialog();
  await openDocFlow(r);
}

$('#btn-open-doc').addEventListener('click', async () => {
  await openDocViaDialog();
  $('#drawer-docs').close();
});
$('#btn-open-doc-landing').addEventListener('click', openDocViaDialog);

async function loadDocuments() {
  const docs = await window.paiban.listDocuments();
  const ul = $('#doc-list');
  ul.innerHTML = '';
  for (const d of docs) {
    const li = document.createElement('li');
    li.className = 'item' + (d.docId === flow.docId ? ' active' : '');
    li.innerHTML = `<span class="name">${escapeHtml(d.name)}</span>
      <span class="meta">${d.head} · ${d.versionCount} 个版本</span>`;
    li.addEventListener('click', async () => {
      await openDocFlow({ docId: d.docId, name: d.name, head: d.head });
      $('#drawer-docs').close();
    });
    ul.appendChild(li);
  }
  renderRecents(docs);
}

// landing 最近文档：最多 4 篇，点击进入编辑态
function renderRecents(docs) {
  const wrap = $('#landing-recents');
  wrap.innerHTML = '';
  const list = (docs ?? []).slice(0, 4);
  if (!list.length) return;
  const label = document.createElement('p');
  label.className = 'dim recents-label';
  label.textContent = '继续上次的工作';
  wrap.appendChild(label);
  const row = document.createElement('div');
  row.className = 'recents-row';
  for (const d of list) {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'suggest-chip';
    chip.textContent = d.name;
    chip.addEventListener('click', () => openDocFlow({ docId: d.docId, name: d.name, head: d.head }));
    row.appendChild(chip);
  }
  wrap.appendChild(row);
}

$('#btn-download').addEventListener('click', async () => {
  if (!flow.docId) return;
  await window.paiban.download(flow.docId);
});

// ---- 版本时间线（抽屉） ----

async function loadVersions() {
  if (!flow.docId) return;
  const versions = await window.paiban.listVersions(flow.docId);
  const ul = $('#version-list');
  ul.innerHTML = '';
  for (const v of [...versions].reverse()) {
    const li = document.createElement('li');
    li.className = 'item' + (v.id === versions[versions.length - 1].id ? ' current' : '');
    const when = new Date(v.at).toLocaleString('zh-CN', { hour12: false });
    li.innerHTML = `<span class="name">${v.id} <span class="meta">${escapeHtml(v.source)}</span></span>
      <span class="meta">${escapeHtml(v.note || '')} · ${when}</span>
      <span class="row">
        <button class="btn pill small" data-act="preview">预览</button>
        <button class="btn pill small" data-act="rollback">回滚</button>
        <button class="btn pill small" data-act="download">下载</button>
      </span>`;
    li.querySelector('[data-act="preview"]').addEventListener('click', async () => {
      const buffer = await window.paiban.getBuffer(flow.docId, v.id);
      const frame = $('#preview-frame');
      $('#preview-placeholder').classList.add('hidden');
      frame.classList.remove('hidden');
      frame.contentWindow.postMessage({ type: 'render', buffer }, '*', [buffer]);
    });
    li.querySelector('[data-act="rollback"]').addEventListener('click', async () => {
      await window.paiban.rollback(flow.docId, v.id);
      dispatch({ type: 'version_note', versionId: v.id, note: `回滚到 ${v.id}` });
      await loadVersions();
      await loadDocuments();
      refreshPreview();
    });
    li.querySelector('[data-act="download"]').addEventListener('click', async () => {
      await window.paiban.download(flow.docId, v.id);
    });
    ul.appendChild(li);
  }
}

// ---- 模板（抽屉） ----

async function loadTemplates() {
  const templates = await window.paiban.listTemplates();
  const ul = $('#template-list');
  ul.innerHTML = '';
  for (const t of templates) {
    const li = document.createElement('li');
    li.className = 'item';
    li.innerHTML = `<span class="name">${escapeHtml(t.name)}</span>
      <span class="meta">${t.placeholderCount} 个占位符 · 反推组件: ${(t.extractedComponents || []).join('/')}</span>`;
    li.addEventListener('click', () => openTemplateDialog(t.templateId));
    ul.appendChild(li);
  }
}

$('#btn-upload-template').addEventListener('click', async () => {
  const r = await window.paiban.uploadTemplate();
  if (!r) return;
  dispatch({ type: 'note', title: '模板', note: `已解析：${r.extracted.join('/')}` });
  await loadTemplates();
});

let dialogTemplateId = null;

async function openTemplateDialog(templateId) {
  dialogTemplateId = templateId;
  const t = await window.paiban.readTemplate(templateId);
  $('#template-dialog-title').textContent = t.meta.name;
  const body = $('#template-dialog-body');
  body.innerHTML = '';

  // 占位符填写区
  const phSec = document.createElement('div');
  phSec.className = 'tpl-section';
  phSec.innerHTML = '<h4>占位符（实例化时预填）</h4>';
  const grid = document.createElement('div');
  grid.className = 'placeholder-grid';
  if (!t.placeholders.length) grid.innerHTML = '<span class="dim">无占位符</span>';
  for (const p of t.placeholders) {
    const input = document.createElement('input');
    input.placeholder = `{{${p.name}}}（出现 ${p.count} 次）`;
    input.dataset.ph = p.name;
    grid.appendChild(input);
  }
  phSec.appendChild(grid);
  body.appendChild(phSec);

  // 规则集摘要
  const rsSec = document.createElement('div');
  rsSec.className = 'tpl-section';
  rsSec.innerHTML = `<h4>排版规则集（样式摘要）</h4><pre>${escapeHtml(JSON.stringify(t.styles, null, 1).slice(0, 2000))}</pre>`;
  body.appendChild(rsSec);

  $('#template-dialog').showModal();
}

$('#btn-instantiate').addEventListener('click', async () => {
  if (!dialogTemplateId) return;
  const values = {};
  document.querySelectorAll('#template-dialog-body input[data-ph]').forEach((input) => {
    if (input.value.trim()) values[input.dataset.ph] = input.value.trim();
  });
  const r = await window.paiban.instantiateTemplate(dialogTemplateId, values);
  $('#template-dialog').close();
  $('#drawer-templates').close();
  await openDocFlow({ docId: r.docId, name: r.name ?? r.docId, versionId: r.version?.id });
});

// ---- 内置规则集（模板抽屉内，一键重排当前文档） ----

async function loadBuiltinRulesets() {
  const rulesets = await window.paiban.listBuiltinRulesets();
  const sel = $('#builtin-ruleset-select');
  sel.innerHTML = '';
  if (!rulesets.length) {
    sel.innerHTML = '<option value="">无内置规则集</option>';
    $('#btn-apply-ruleset').disabled = true;
    return;
  }
  for (const r of rulesets) {
    const opt = document.createElement('option');
    opt.value = r.id;
    opt.textContent = r.id;
    opt.title = r.description;
    sel.appendChild(opt);
  }
}

$('#btn-apply-ruleset').addEventListener('click', async () => {
  const rulesetId = $('#builtin-ruleset-select').value;
  if (!rulesetId) return;
  if (!flow.docId) return;
  const r = await window.paiban.applyBuiltinRuleset(flow.docId, rulesetId);
  $('#drawer-templates').close();
  if (r.errors?.length) {
    dispatch({ type: 'error', message: `按 ${rulesetId} 重排完成，但有 ${r.errors.length} 条命令失败：${r.errors[0].error || ''}` });
  } else {
    dispatch({ type: 'version_note', versionId: r.version?.id ?? '—', note: `按内置规则集 ${rulesetId} 重排` });
  }
  await loadVersions();
  refreshPreview();
});

// ---- 抽屉与对话框 ----

const DRAWERS = { docs: '#drawer-docs', templates: '#drawer-templates', versions: '#drawer-versions' };

$('#btn-drawer-docs').addEventListener('click', async () => { await loadDocuments(); $(DRAWERS.docs).showModal(); });
$('#btn-drawer-templates').addEventListener('click', async () => { await loadTemplates(); $(DRAWERS.templates).showModal(); });
$('#btn-drawer-versions').addEventListener('click', async () => { await loadVersions(); $(DRAWERS.versions).showModal(); });

document.querySelectorAll('[data-close]').forEach((btn) => {
  btn.addEventListener('click', () => btn.closest('dialog').close());
});
// 点击抽屉/对话框背板关闭
document.querySelectorAll('dialog').forEach((dlg) => {
  dlg.addEventListener('click', (ev) => { if (ev.target === dlg) dlg.close(); });
});

// ---- 模型设置 / 状态 ----

async function refreshModelStatus() {
  const [cfg, st] = await Promise.all([window.paiban.getConfig(), window.paiban.agentStatus()]);
  const el = $('#model-status');
  if (st.ready) {
    el.textContent = `${st.provider} · ${st.model}`;
    el.className = 'model-status pill ready';
  } else {
    el.textContent = `agent 未就绪：${st.reason || cfg.provider + '/' + cfg.model}`;
    el.className = 'model-status pill down';
    el.title = st.reason || '';
  }
}

$('#btn-settings').addEventListener('click', async () => {
  const cfg = await window.paiban.getConfig();
  $('#cfg-provider').value = cfg.baseUrl ? 'gateway' : cfg.provider;
  $('#cfg-model').value = cfg.model || '';
  $('#cfg-baseurl').value = cfg.baseUrl || '';
  $('#cfg-apikey').value = '';
  $('#settings-dialog').showModal();
});

$('#btn-save-config').addEventListener('click', async () => {
  const provider = $('#cfg-provider').value;
  const patch = {
    provider: provider === 'gateway' ? 'openai' : provider,
    model: $('#cfg-model').value.trim() || undefined,
    baseUrl: provider === 'gateway' ? $('#cfg-baseurl').value.trim() : '',
  };
  const key = $('#cfg-apikey').value.trim();
  if (key) patch.apiKey = key;
  await window.paiban.setConfig(patch);
  $('#settings-dialog').close();
  dispatch({ type: 'note', title: '配置', note: '模型配置已保存，重启应用后生效' });
});

// ---- 启动 ----

(async function boot() {
  await Promise.all([loadDocuments(), loadTemplates(), loadBuiltinRulesets(), refreshModelStatus()]);
  setInterval(refreshModelStatus, 15000);
})();
