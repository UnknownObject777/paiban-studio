// ui/conversation-flow.ts — 对话流状态机（issue #31：landing → 编辑态交互重写）。
//
// 纯函数、无 DOM 依赖，是前端交互的唯一新 seam：
//   - 两相视图：landing（居中对话窗）/ editing（对话左移 + 右侧大预览）
//   - 瀑布流条目：user / assistant（流式累积）/ tool（步骤卡）/ version / note / error
//   - 进度模型：由 agent 事件流推导（deriveProgress），后端零改动
//   - 轮次：每次 send 开新一轮，进度链只呈现本轮步骤（lastOutcome 记录收尾方式）
//
// 渲染层（public/app.js，ES module）持有一份 FlowState，每个事件 reduce 后按 items 增量渲染。

// ---- 事件 ----

/** agent 事件（与 agent-core/bridge.ts 的 AgentEvent 对齐，渲染层只消费这五种） */
export type AgentEvent =
  | { type: 'text_delta'; delta: string }
  | { type: 'tool_start'; name: string; args: string }
  | { type: 'tool_end'; name: string; isError: boolean; summary: string }
  | { type: 'done' }
  | { type: 'error'; message: string };

/** UI 动作（打开文档、发送、停止、返回首页、版本卡 / 便签卡补记） */
export type UiAction =
  | { type: 'doc_opened'; docId: string; name: string; versionId?: string }
  | { type: 'send'; text: string }
  | { type: 'abort' }
  | { type: 'back_to_landing' }
  | { type: 'version_note'; versionId: string; note: string }
  | { type: 'note'; title: string; note: string };

export type FlowEvent = AgentEvent | UiAction;

// ---- 状态 ----

export type Phase = 'landing' | 'editing';

export type ChatItem =
  | { kind: 'user'; id: number; text: string }
  | { kind: 'assistant'; id: number; text: string; streaming: boolean }
  | { kind: 'tool'; id: number; name: string; label: string; args: string; status: 'running' | 'done' | 'error'; summary: string }
  | { kind: 'version'; id: number; versionId: string; note: string }
  | { kind: 'note'; id: number; title: string; note: string }
  | { kind: 'error'; id: number; message: string };

export interface FlowState {
  phase: Phase;
  docId: string | null;
  docName: string | null;
  busy: boolean;
  items: ChatItem[];
  nextId: number;
  /** 本轮第一个条目的 id：进度链只统计 id >= roundStartId 的步骤卡 */
  roundStartId: number;
  /** 最近一轮的收尾方式；进行中 / 从未开始为 null */
  lastOutcome: 'done' | 'error' | 'aborted' | null;
}

/** 工具名 → 用户可感知的步骤标签（进度链与步骤卡共用） */
export const TOOL_LABELS: Record<string, string> = {
  doc_outline: '分析文档结构',
  doc_edit: '应用排版修改',
  ruleset_read: '读取内置规则集',
  template_read: '读取模板',
  version_store: '保存 / 回滚版本',
};

/** abort 时给运行中步骤卡的占位 summary，迟到 tool_end 凭它找回原卡而不是落成重复卡 */
const ABORTED_SUMMARY = '已被用户停止';

export function createFlow(): FlowState {
  return {
    phase: 'landing',
    docId: null,
    docName: null,
    busy: false,
    items: [],
    nextId: 1,
    roundStartId: 1,
    lastOutcome: null,
  };
}

// ---- reducer ----

export function reduce(state: FlowState, event: FlowEvent): FlowState {
  switch (event.type) {
    case 'doc_opened':
      return push(state, {
        phase: 'editing',
        docId: event.docId,
        docName: event.name,
      }, { kind: 'version', versionId: event.versionId ?? 'v1', note: `已打开《${event.name}》（原稿未动）` });

    case 'back_to_landing':
      // 文档上下文与对话保留：回到首页后可直接继续对同一文档发指令
      return { ...state, phase: 'landing' };

    case 'send': {
      const text = event.text.trim();
      if (!text) return state;
      if (!state.docId) {
        return push(state, {}, { kind: 'error', message: '请先打开或选择一篇工作文档。' });
      }
      // 新一轮：进度链从此处重新计数
      return push(state, { busy: true, roundStartId: state.nextId, lastOutcome: null }, { kind: 'user', text });
    }

    case 'text_delta': {
      const last = state.items[state.items.length - 1];
      if (last?.kind === 'assistant' && last.streaming) {
        return replaceLast(state, { ...last, text: last.text + event.delta });
      }
      return push(state, {}, { kind: 'assistant', text: event.delta, streaming: true });
    }

    case 'tool_start': {
      // 工具调用截断当前流式气泡，落一张运行中步骤卡
      return push(sealAssistant(state), {}, toolCard(event.name, event.args, 'running', ''));
    }

    case 'tool_end': {
      const status = event.isError ? 'error' : 'done';
      // 1) 优先回填最近一个运行中的同名步骤卡
      let idx = state.items.findLastIndex((i) => i.kind === 'tool' && i.status === 'running' && i.name === event.name);
      // 2) 退到任意运行中步骤卡
      if (idx < 0) idx = state.items.findLastIndex((i) => i.kind === 'tool' && i.status === 'running');
      // 3) abort 后迟到的 tool_end：回填被停止的原卡，不落成重复卡
      if (idx < 0) idx = state.items.findLastIndex((i) => i.kind === 'tool' && i.name === event.name && i.summary === ABORTED_SUMMARY);
      if (idx >= 0) {
        const item = state.items[idx];
        if (item.kind !== 'tool') return state;
        const items = state.items.slice();
        items[idx] = { ...item, status, summary: event.summary || item.summary };
        return { ...state, items };
      }
      // 4) 兜底：轮次外的迟到事件直接忽略，轮次内落成卡
      if (!state.busy) return state;
      return push(state, {}, toolCard(event.name, '', status, event.summary));
    }

    case 'done':
      return { ...closeRound(sealAssistant({ ...state, busy: false }), 'done'), lastOutcome: 'done' };

    case 'error':
      return {
        ...closeRound(sealAssistant(push({ ...state, busy: false }, {}, { kind: 'error', message: event.message })), 'error'),
        lastOutcome: 'error',
      };

    case 'abort':
      return { ...closeRound(sealAssistant({ ...state, busy: false }), 'done', ABORTED_SUMMARY), lastOutcome: 'aborted' };

    case 'version_note':
      return push(state, {}, { kind: 'version', versionId: event.versionId, note: event.note });

    case 'note':
      return push(state, {}, { kind: 'note', title: event.title, note: event.note });

    default:
      // 未知事件（如 bridge 回显的 user 事件）忽略，状态不变
      return state;
  }
}

// ---- 进度模型（派生，单一事实源 = 本轮 items + busy + lastOutcome） ----

export interface ProgressStep {
  id: number;
  label: string;
  status: 'running' | 'done' | 'error';
  summary: string;
}

export interface ProgressView {
  /** 本轮有步骤、正在工作、或上一轮以出错收尾时显示进度条 */
  visible: boolean;
  /** 一轮工作进行中 */
  active: boolean;
  /** 当前进行中步骤的标签；无运行中步骤时为「撰写回复…」/「思考中…」/空 */
  currentLabel: string;
  /** 非进行中的一句话总结（完成 / 出错 / 已停止） */
  finalLabel: string;
  steps: ProgressStep[];
  failedCount: number;
}

export function deriveProgress(state: FlowState): ProgressView {
  const steps: ProgressStep[] = state.items
    .filter((i): i is Extract<ChatItem, { kind: 'tool' }> => i.kind === 'tool' && i.id >= state.roundStartId)
    .map((t) => ({ id: t.id, label: t.label, status: t.status, summary: t.summary }));
  const running = steps.findLast((s) => s.status === 'running');
  const last = state.items[state.items.length - 1];
  const streaming = last?.kind === 'assistant' && last.streaming;
  const failedCount = steps.filter((s) => s.status === 'error').length;

  let currentLabel = '';
  if (state.busy) currentLabel = running ? running.label : streaming ? '撰写回复…' : '思考中…';

  let finalLabel = '';
  if (!state.busy && state.lastOutcome) {
    if (state.lastOutcome === 'done') finalLabel = `完成 · 共 ${steps.length} 步`;
    else if (state.lastOutcome === 'aborted') finalLabel = `已停止 · 共 ${steps.length} 步`;
    else finalLabel = failedCount > 0 ? `出错 · ${failedCount} 步失败` : '出错（详见对话流）';
  }

  return {
    visible: state.busy || steps.length > 0 || state.lastOutcome === 'error',
    active: state.busy,
    currentLabel,
    finalLabel,
    steps,
    failedCount,
  };
}

// ---- 内部辅助 ----

type DistributiveOmit<T, K extends keyof any> = T extends any ? Omit<T, K> : never;

function push(state: FlowState, patch: Partial<FlowState>, item: DistributiveOmit<ChatItem, 'id'>): FlowState {
  const withId = { ...item, id: state.nextId } as ChatItem;
  return { ...state, ...patch, items: [...state.items, withId], nextId: state.nextId + 1 };
}

function toolCard(name: string, args: string, status: 'running' | 'done' | 'error', summary: string): DistributiveOmit<ChatItem, 'id'> {
  return { kind: 'tool', name, label: TOOL_LABELS[name] ?? name, args, status, summary };
}

function replaceLast(state: FlowState, item: ChatItem): FlowState {
  const items = state.items.slice();
  items[items.length - 1] = item;
  return { ...state, items };
}

/** 流式气泡封口（tool_start / done / error / abort 时调用） */
function sealAssistant(state: FlowState): FlowState {
  const last = state.items[state.items.length - 1];
  if (last?.kind === 'assistant' && last.streaming) {
    return replaceLast(state, { ...last, streaming: false });
  }
  return state;
}

/** 一轮结束：残留的运行中步骤卡统一收尾 */
function closeRound(state: FlowState, status: 'done' | 'error', summary?: string): FlowState {
  if (!state.items.some((i) => i.kind === 'tool' && i.status === 'running')) return state;
  return {
    ...state,
    items: state.items.map((i) =>
      i.kind === 'tool' && i.status === 'running'
        ? { ...i, status, summary: summary ?? i.summary }
        : i,
    ),
  };
}
