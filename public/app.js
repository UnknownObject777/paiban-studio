// app.js — 三栏工作台前端逻辑（纯 JS，无框架，R1）。
// 状态：currentDocId / templates / docs / versions；全部经 window.paiban（preload 白名单）与主进程通信。

const $ = (sel) => document.querySelector(sel);

const state = {
  currentDocId: null,
  previewReady: false,
  renderTimer: null,
  assistantMsg: null, // 流式累积中的 assistant 消息节点
  busy: false,
};

// ---- 预览（防抖刷新，D4：编辑 → ArrayBuffer → iframe） ----

function refreshPreview() {
  if (!state.currentDocId) return;
  clearTimeout(state.renderTimer);
  state.renderTimer = setTimeout(async () => {
    const buffer = await window.paiban.getBuffer(state.currentDocId);
    const frame = $('#preview-frame');
    $('#preview-placeholder').classList.add('hidden');
    frame.classList.remove('hidden');
    frame.contentWindow.postMessage({ type: 'render', buffer }, '*', [buffer]);
  }, 250); // 连续编辑合并
}

window.addEventListener('message', (ev) => {
  if (ev.data?.type === 'preview-ready') {
    state.previewReady = true;
  }
});

// ---- 对话流 ----

function addMsg(role, text) {
  $('#chat-empty')?.remove();
  const div = document.createElement('div');
  div.className = `msg ${role}`;
  div.textContent = text;
  $('#chat-stream').appendChild(div);
  div.scrollIntoView({ block: 'end' });
  return div;
}

function addToolChip(name, summary, isError = false) {
  $('#chat-empty')?.remove();
  const chip = document.createElement('div');
  chip.className = 'tool-chip' + (isError ? ' error' : '');
  chip.innerHTML = `<span class="dot"></span><span>${escapeHtml(name)}</span><span class="dim">${escapeHtml(summary || '')}</span>`;
  $('#chat-stream').appendChild(chip);
  chip.scrollIntoView({ block: 'end' });
  return chip;
}

function addVersionChip(version, note) {
  const chip = document.createElement('div');
  chip.className = 'version-chip';
  chip.textContent = `已存版本 ${version.id}${note ? ' · ' + note : ''}`;
  $('#chat-stream').appendChild(chip);
  chip.scrollIntoView({ block: 'end' });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function setBusy(busy) {
  state.busy = busy;
  $('#btn-send').classList.toggle('hidden', busy);
  $('#btn-abort').classList.toggle('hidden', !busy);
  $('#chat-input').disabled = busy;
}

// agent 事件流
window.paiban.onAgentEvent((event) => {
  if (event.type === 'text_delta') {
    if (!state.assistantMsg) state.assistantMsg = addMsg('assistant', '');
    state.assistantMsg.textContent += event.delta;
    state.assistantMsg.scrollIntoView({ block: 'end' });
  } else if (event.type === 'tool_start') {
    state.assistantMsg = null; // 工具调用切断当前流式消息
    addToolChip(event.name, event.args);
  } else if (event.type === 'tool_end') {
    addToolChip(`${event.name} ✓`, event.summary, event.isError);
    if (event.name === 'doc_edit' && !event.isError) {
      refreshPreview();
      loadVersions();
    }
    if (event.name === 'version_store') {
      refreshPreview();
      loadVersions();
    }
  } else if (event.type === 'done') {
    state.assistantMsg = null;
    setBusy(false);
    refreshPreview();
    loadVersions();
    loadDocuments();
  } else if (event.type === 'error') {
    addMsg('error', '出错了：' + event.message);
    setBusy(false);
  }
});

$('#chat-form').addEventListener('submit', async (ev) => {
  ev.preventDefault();
  const input = $('#chat-input');
  const text = input.value.trim();
  if (!text || state.busy) return;
  if (!state.currentDocId) {
    addMsg('error', '请先在左栏打开或选择一篇工作文档。');
    return;
  }
  input.value = '';
  addMsg('user', text);
  setBusy(true);
  const r = await window.paiban.agentSend(state.currentDocId, text);
  if (!r.ok) {
    addMsg('error', r.error || 'agent 发送失败');
    setBusy(false);
  }
});

$('#btn-abort').addEventListener('click', () => window.paiban.agentAbort());

// 空态示例指令 chip：点击填入输入框并聚焦
document.querySelectorAll('.suggest-chip').forEach((chip) => {
  chip.addEventListener('click', () => {
    const input = $('#chat-input');
    input.value = chip.dataset.prompt || '';
    input.focus();
  });
});

// ---- 文档 ----

async function loadDocuments() {
  const docs = await window.paiban.listDocuments();
  const ul = $('#doc-list');
  ul.innerHTML = '';
  for (const d of docs) {
    const li = document.createElement('li');
    li.className = 'item' + (d.docId === state.currentDocId ? ' active' : '');
    li.innerHTML = `<span class="name">${escapeHtml(d.name)}</span>
      <span class="meta">${d.head} · ${d.versionCount} 个版本</span>`;
    li.addEventListener('click', () => {
      state.currentDocId = d.docId;
      loadDocuments();
      loadVersions();
      refreshPreview();
      $('#btn-download').disabled = false;
    });
    ul.appendChild(li);
  }
}

$('#btn-open-doc').addEventListener('click', async () => {
  const r = await window.paiban.openDialog();
  if (!r) return;
  state.currentDocId = r.docId;
  addVersionChip(r.version, '已上传（原稿未动）');
  $('#btn-download').disabled = false;
  await loadDocuments();
  await loadVersions();
  refreshPreview();
});

$('#btn-download').addEventListener('click', async () => {
  if (!state.currentDocId) return;
  await window.paiban.download(state.currentDocId);
});

// ---- 版本时间线 ----

async function loadVersions() {
  if (!state.currentDocId) return;
  const versions = await window.paiban.listVersions(state.currentDocId);
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
      const buffer = await window.paiban.getBuffer(state.currentDocId, v.id);
      const frame = $('#preview-frame');
      $('#preview-placeholder').classList.add('hidden');
      frame.classList.remove('hidden');
      frame.contentWindow.postMessage({ type: 'render', buffer }, '*', [buffer]);
    });
    li.querySelector('[data-act="rollback"]').addEventListener('click', async () => {
      await window.paiban.rollback(state.currentDocId, v.id);
      addVersionChip({ id: v.id }, `回滚到 ${v.id}`);
      await loadVersions();
      await loadDocuments();
      refreshPreview();
    });
    li.querySelector('[data-act="download"]').addEventListener('click', async () => {
      await window.paiban.download(state.currentDocId, v.id);
    });
    ul.appendChild(li);
  }
}

// ---- 模板 ----

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
  addVersionChip({ id: '模板' }, `已解析：${r.extracted.join('/')}`);
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
  state.currentDocId = r.docId;
  $('#btn-download').disabled = false;
  addVersionChip(r.version, '从模板实例化');
  await loadDocuments();
  await loadVersions();
  refreshPreview();
});

// ---- 内置规则集（手写资产，一键重排当前文档） ----

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
  if (!state.currentDocId) {
    addMsg('assistant', '请先打开一篇文档，再按内置规则集重排。');
    return;
  }
  const r = await window.paiban.applyBuiltinRuleset(state.currentDocId, rulesetId);
  if (r.errors?.length) {
    addMsg('assistant', `按 ${rulesetId} 重排完成，但有 ${r.errors.length} 条命令失败：${r.errors[0].error || ''}`);
  } else {
    addVersionChip(r.version, `按内置规则集 ${rulesetId} 重排`);
  }
  await loadVersions();
  refreshPreview();
});

document.querySelectorAll('[data-close]').forEach((btn) => {
  btn.addEventListener('click', () => btn.closest('dialog').close());
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
  addMsg('assistant', '配置已保存。重启应用后 agent 以新配置初始化。');
});

// ---- 启动 ----

(async function boot() {
  await Promise.all([loadDocuments(), loadTemplates(), loadBuiltinRulesets(), refreshModelStatus()]);
  setInterval(refreshModelStatus, 15000);
})();
