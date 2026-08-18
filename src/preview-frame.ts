// preview-frame.ts — 预览 iframe 脚本（D4：OnlyOffice 静态 SDK 只读渲染，wrapper 见 src/onlyoffice-comp/）。
// 契约（与 docx-preview 时代一致，父页面 app.js 无需改动）：
//   收  {type:'render', buffer}        — docx ArrayBuffer（transfer），防抖在父页面做
//   回  {type:'preview-ready'}         — 脚本就绪
//   回  {type:'rendered', ok, error?}  — 渲染冒烟信号
// 只读：编辑仍走 agent / docx-core / 版本链，这里只是显示器。

import {
  FILE_TYPE,
  ONLYOFFICE_ID,
  OnlyOfficeManager,
} from './onlyoffice-comp/index.js';

const status = document.getElementById('status')!;

let manager: OnlyOfficeManager | null = null;
let renderSeq = 0; // 丢弃过期的异步渲染结果（连续编辑时旧的 openFile 后返回）

async function render(buffer: ArrayBuffer): Promise<void> {
  const seq = ++renderSeq;
  status.style.display = 'block';
  status.textContent = '渲染中…';
  const file = new File([buffer], 'preview.docx', {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  if (manager) {
    await manager.openFile(file, true);
  } else {
    manager = await OnlyOfficeManager.createWithFile(
      {
        containerId: ONLYOFFICE_ID,
        fileType: FILE_TYPE.DOCX,
        defaultFileName: 'preview.docx',
        readOnly: true,
        lang: 'zh',
      },
      file,
    );
  }
  if (seq !== renderSeq) return;
  status.style.display = 'none';
  parent.postMessage({ type: 'rendered', ok: true }, '*');
}

window.addEventListener('message', (ev) => {
  const data = ev.data || {};
  if (data.type !== 'render' || !data.buffer) return;
  const handler = new URLSearchParams(location.search).has('x2t-only') ? renderX2tOnly : render;
  handler(data.buffer as ArrayBuffer).catch((err) => {
    const msg = err instanceof Error ? err.message : String(err);
    status.style.display = 'block';
    status.textContent = '预览渲染失败：' + msg;
    parent.postMessage({ type: 'rendered', ok: false, error: msg }, '*');
  });
});

parent.postMessage({ type: 'preview-ready' }, '*');

// x2t 隔离测试（?x2t-only 调试入口）：跳过 OnlyOffice 编辑器，只跑 docx → Editor.bin 转换。
// 2026-08-18：完整渲染路径已验证可用，本路径仅留作转换层问题排查开关。
import { converter } from './onlyoffice-comp/internal/editor/x2t.js';
import { getX2tConvertFormats } from './onlyoffice-comp/internal/editor/utils.js';

async function renderX2tOnly(buffer: ArrayBuffer): Promise<void> {
  const { formatFrom, formatTo } = getX2tConvertFormats('docx');
  console.log('[diag] x2t convert start, docx bytes=', buffer.byteLength);
  const t0 = performance.now();
  const result = await converter.convert({
    data: buffer,
    fileFrom: 'doc.docx',
    fileTo: 'Editor.bin',
    formatFrom,
    formatTo,
  });
  console.log('[diag] x2t convert OK, bin bytes=', result.output?.length, 'ms=', Math.round(performance.now() - t0));
  status.style.display = 'block';
  status.textContent = `x2t 转换成功（${result.output?.length} 字节）`;
  parent.postMessage({ type: 'rendered', ok: true }, '*');
}
