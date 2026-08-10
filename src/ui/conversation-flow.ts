// ui/conversation-flow.ts — 对话流状态机（issue #31：landing → 编辑态交互重写）。
//
// 纯函数、无 DOM 依赖，是前端交互的唯一新 seam：
//   - 两相视图：landing（居中对话窗）/ editing（对话左移 + 右侧大预览）
//   - 瀑布流条目：user / assistant（流式累积）/ tool（步骤卡）/ version / error
//   - 进度模型：由 agent 事件流推导（deriveProgress），后端零改动
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

/** UI 动作（打开文档、发送、停止、返回首页、版本卡补记） */
export type UiAction =
  | { type: 'doc_opened'; docId: string; name: string; versionId?: string }
  | { type: 'send'; text: string }
  | { type: 'abort' }
  | { type: 'back_to_landing' }
  | { type: 'version_note'; versionId: string; note: string };

export type FlowEvent = AgentEvent | UiAction;

// ---- 状态 ----

export type Phase = 'landing' | 'editing';

export type ChatItem =
  | { kind: 'user'; id: number; text: string }
  | { kind: 'assistant'; id: number; text: string; streaming: boolean }
  | { kind: 'tool'; id: number; name: string; label: string; args: string; status: 'running' | 'done' | 'error'; summary: string }
  | { kind: 'version'; id: number; versionId: string; note: string }
  | { kind: 'error'; id: number; message: string };

export interface FlowState {
  phase: Phase;
  docId: string | null;
  docName: string | null;
  busy: boolean;
  items: ChatItem[];
  nextId: number;
}

/** 工具名 → 用户可感知的步骤标签（进度链与步骤卡共用） */
export const TOOL_LABELS: Record<string, string> = {
  doc_outline: '分析文档结构',
  doc_edit: '应用排版修改',
  ruleset_read: '读取内置规则集',
  template_read: '读取模板',
  version_store: '保存 / 回滚版本',
};

export function createFlow(): FlowState {
  return { phase: 'landing', docId: null, docName: null, busy: false, items: [], nextId: 1 };
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
      return push(state, { busy: true }, { kind: 'user', text });
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
      const s = sealAssistant(state);
      return push(s, {}, {
        kind: 'tool',
        name: event.name,
        label: TOOL_LABELS[event.name] ?? event.name,
        args: event.args,
        status: 'running',
        summary: '',
      });
    }

    case 'tool_end': {
      // 回填最近一个运行中的同名步骤卡；找不到运行中步骤则兜底落成卡
      const idx = findLastIndex(state.items, (i) => i.kind === 'tool' && i.status === 'running' && i.name === event.name)
        ?? findLastIndex(state.items, (i) => i.kind === 'tool' && i.status === 'running');
      if (idx === null || idx === undefined) {
        return push(state, {}, {
          kind: 'tool',
          name: event.name,
          label: TOOL_LABELS[event.name] ?? event.name,
          args: '',
          status: event.isError ? 'error' : 'done',
          summary: event.summary,
        });
      }
      const item = state.items[idx];
      if (item.kind !== 'tool') return state;
      const items = state.items.slice();
      items[idx] = { ...item, status: event.isError ? 'error' : 'done', summary: event.summary };
      return { ...state, items };
    }

    case 'done':
      return closeRound(sealAssistant({ ...state, busy: false }), 'done');

    case 'error':
      return closeRound(sealAssistant(push({ ...state, busy: false }, {}, { kind: 'error', message: event.message })), 'error');

    case 'abort':
      return closeRound(sealAssistant({ ...state, busy: false }), 'done', '已被用户停止');

    case 'version_note':
      return push(state, {}, { kind: 'version', versionId: event.versionId, note: event.note });

    default:
      // 未知/不关心的事件（如 bridge 回显的 user 事件）原样透传
      return state;
  }
}

// ---- 进度模型（派生，单一事实源 = items + busy） ----

export interface ProgressStep {
  id: number;
  label: string;
  status: 'running' | 'done' | 'error';
  summary: string;
}

export interface ProgressView {
  /** 有步骤或正在工作时显示进度条 */
  visible: boolean;
  /** 一轮工作进行中 */
  active: boolean;
  /** 当前进行中步骤的标签；无运行中步骤时为「撰写回复…」/「思考中…」/空 */
  currentLabel: string;
  steps: ProgressStep[];
  failedCount: number;
}

export function deriveProgress(state: FlowState): ProgressView {
  const steps: ProgressStep[] = state.items
    .filter((i): i is Extract<ChatItem, { kind: 'tool' }> => i.kind === 'tool')
    .map((t) => ({ id: t.id, label: t.label, status: t.status, summary: t.summary }));
  const running = [...steps].reverse().find((s) => s.status === 'running');
  const last = state.items[state.items.length - 1];
  const streaming = last?.kind === 'assistant' && last.streaming;
  let currentLabel = '';
  if (state.busy) currentLabel = running ? running.label : streaming ? '撰写回复…' : '思考中…';
  return {
    visible: state.busy || steps.length > 0,
    active: state.busy,
    currentLabel,
    steps,
    failedCount: steps.filter((s) => s.status === 'error').length,
  };
}

// ---- 内部辅助 ----

type DistributiveOmit<T, K extends keyof any> = T extends any ? Omit<T, K> : never;

function push(state: FlowState, patch: Partial<FlowState>, item: DistributiveOmit<ChatItem, 'id'>): FlowState {
  const withId = { ...item, id: state.nextId } as ChatItem;
  return { ...state, ...patch, items: [...state.items, withId], nextId: state.nextId + 1 };
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

function findLastIndex<T>(arr: T[], pred: (x: T) => boolean): number | null {
  for (let i = arr.length - 1; i >= 0; i--) {
    if (pred(arr[i])) return i;
  }
  return null;
}
