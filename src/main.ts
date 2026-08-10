// main.ts — Electron 主进程（R1：主进程跑 agent 编排 + docx 编辑内核 + 存储层；IPC 通信）。
//
// 职责：
//   1. 创建 BrowserWindow（三栏工作台 UI，public/index.html）
//   2. 注册 IPC handlers（workbench:* 全部走 Workspace 服务层；agent:* 走 agent-core）
//   3. 数据目录：<userData>/paiban-studio（对象存储 + 版本链 + 模板库 + 配置）

import { app, BrowserWindow, ipcMain, dialog, shell } from 'electron';
import { join, normalize, sep, extname } from 'node:path';
import { readFileSync, writeFileSync, appendFileSync, mkdirSync, createReadStream, statSync } from 'node:fs';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { fileURLToPath } from 'node:url';
import { Workspace } from './server/workspace.js';

const __dirname = fileURLToPath(new URL('.', import.meta.url));
// 编译产物位于 dist/src/，项目根在其上两级（public/index.html、templates/ 等均位于项目根）。
const ROOT = join(__dirname, '../..');
const SMOKE = process.env.PAIBAN_SMOKE === '1';

// ---- 回环静态服务器（D4 预览替换：OnlyOffice 静态 SDK）----
// OnlyOffice SDK 运行时要动态加载脚本分块 / Web Worker / x2t wasm（fetch/XHR），
// file:// 下 Chromium 禁止这些请求；paiban:// 自定义 scheme 实测在 module worker 内
// 出现安全上下文缺失（DecompressionStream 不可用）与 wasm 内存分配 OOM，因此渲染层
// 整体改走 127.0.0.1 回环 HTTP（随机端口、仅监听本机、仅放行 public/ 与 dist/）。
// 缓存策略：SDK 资产（public/packages/，版本目录内容不变）长缓存；其余 no-cache，
// 避免升级/重建后磁盘缓存里的旧模块脚本被重用（调试中实测踩到）。
const STATIC_MIME: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.wasm': 'application/wasm',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.ico': 'image/x-icon',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.txt': 'text/plain; charset=utf-8',
};

function startStaticServer(): Promise<number> {
  const allowedRoots = [normalize(join(ROOT, 'public')) + sep, normalize(join(ROOT, 'dist')) + sep];
  const immutablePrefix = normalize(join(ROOT, 'public', 'packages')) + sep;
  const server = createServer((req, res) => {
    const rel = decodeURIComponent(new URL(req.url ?? '/', 'http://127.0.0.1').pathname).replace(/^\/+/, '');
    const filePath = normalize(join(ROOT, rel));
    if (!allowedRoots.some((root) => filePath.startsWith(root))) {
      res.writeHead(403).end('forbidden');
      return;
    }
    let size: number;
    try {
      size = statSync(filePath).size;
    } catch {
      res.writeHead(404).end('not found');
      return;
    }
    res.writeHead(200, {
      'Content-Type': STATIC_MIME[extname(filePath).toLowerCase()] ?? 'application/octet-stream',
      'Content-Length': size,
      'Cache-Control': filePath.startsWith(immutablePrefix) ? 'public, max-age=31536000, immutable' : 'no-cache',
    });
    createReadStream(filePath).pipe(res);
  });
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => resolve((server.address() as AddressInfo).port));
  });
}

let staticPort = 0;

// ---- 崩溃可观测性（issue #30）----
// 背景：#19 演示中 GUI 三次静默退出，stdout/stderr 无任何日志，headless smoke 正常，
// 疑似渲染进程崩溃或主进程未捕获异常。这里捕获关键异常/生命周期事件，统一写
// stderr（终端启动时可见）与 <userData>/paiban-studio/crash.log（append + 时间戳）。

function crashLog(line: string): void {
  const msg = `[${new Date().toISOString()}] ${line}`;
  try {
    console.error(msg); // stderr 镜像
    const dir = join(app.getPath('userData'), 'paiban-studio');
    mkdirSync(dir, { recursive: true });
    appendFileSync(join(dir, 'crash.log'), msg + '\n');
  } catch {
    /* 日志通道自身故障不致命 */
  }
}

// 1) 主进程未捕获异常 / 未处理 Promise 拒绝。
//    默认 Node 行为是打 stderr 后退出；这里先留证据再继续运行，避免演示中途直接消失。
process.on('uncaughtException', (err) => {
  crashLog(`[uncaughtException] ${err?.stack ?? String(err)}`);
});
process.on('unhandledRejection', (reason) => {
  const detail = reason instanceof Error ? (reason.stack ?? String(reason)) : String(reason);
  crashLog(`[unhandledRejection] ${detail}`);
});

// 2) 渲染进程 / 子进程消失（崩溃、OOM、被 kill 是「窗口无声消失」的头号嫌疑）。
//    Electron ≥22 在 app 上也提供 render-process-gone（附带所属 webContents）。
app.on('render-process-gone', (_e, _wc, details) => {
  crashLog(`[app:render-process-gone] reason=${details.reason} exitCode=${details.exitCode}`);
});
// 兜底：任何 webContents（含未来新建窗口/DevTools）的渲染进程消失都记录。
app.on('web-contents-created', (_e, contents) => {
  contents.on('render-process-gone', (_ev, details) => {
    crashLog(`[webContents:render-process-gone] reason=${details.reason} exitCode=${details.exitCode}`);
  });
});
// 子进程（GPU / Utility 等，不含渲染进程）异常消失。
app.on('child-process-gone', (_e, details) => {
  crashLog(`[child-process-gone] type=${details.type} reason=${details.reason} exitCode=${details.exitCode} name=${details.name ?? '-'}`);
});

// 3) 退出生命周期：正常退出会走 before-quit → will-quit → quit 链。
//    若 crash.log 里没有这条链而进程消失，即非优雅退出（崩溃/被杀），与 #19 现象吻合。
app.on('before-quit', () => crashLog('[app:before-quit]'));
app.on('will-quit', () => crashLog('[app:will-quit]'));
app.on('quit', () => crashLog('[app:quit]'));

crashLog(`[app:start] pid=${process.pid} electron=${process.versions.electron} node=${process.versions.node} smoke=${SMOKE}`);

let workspace: Workspace;
let agentBridge: any = null;

function dataDir(): string {
  return join(app.getPath('userData'), 'paiban-studio');
}

// ---- IPC：工作台 ----

function registerWorkbenchIpc(): void {
  const W = () => workspace;

  ipcMain.handle('workbench:uploadDocument', async (_e, { name, bytes }) => {
    const r = W().uploadDocument(Buffer.from(bytes), name);
    return r;
  });

  // 经系统文件对话框打开 docx（副本入库，原稿不动）
  ipcMain.handle('workbench:openDialog', async () => {
    const win = BrowserWindow.getAllWindows()[0];
    const r = await dialog.showOpenDialog(win, {
      filters: [{ name: 'Word 文档', extensions: ['docx'] }],
      properties: ['openFile'],
    });
    if (r.canceled || !r.filePaths.length) return null;
    const path = r.filePaths[0];
    const buffer = readFileSync(path);
    const name = path.split(/[\\/]/).pop() ?? 'document.docx';
    return { ...W().uploadDocument(buffer, name), sourcePath: path };
  });

  ipcMain.handle('workbench:listDocuments', () => W().listDocuments());
  ipcMain.handle('workbench:getBuffer', (_e, { docId, versionId }) => {
    const buf = W().getDocumentBuffer(docId, versionId);
    return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength); // ArrayBuffer for IPC
  });
  ipcMain.handle('workbench:outline', (_e, { docId }) => W().getOutline(docId));
  ipcMain.handle('workbench:applyCommands', (_e, { docId, commands, note }) =>
    W().applyCommands(docId, commands, { source: 'manual', note }));
  ipcMain.handle('workbench:listVersions', (_e, { docId }) => W().listVersions(docId));
  ipcMain.handle('workbench:rollback', (_e, { docId, versionId }) => W().rollback(docId, versionId));

  ipcMain.handle('workbench:download', async (_e, { docId, versionId }) => {
    const win = BrowserWindow.getAllWindows()[0];
    const doc = W().listDocuments().find((d) => d.docId === docId);
    const r = await dialog.showSaveDialog(win, {
      defaultPath: doc?.name || 'document.docx',
      filters: [{ name: 'Word 文档', extensions: ['docx'] }],
    });
    if (r.canceled || !r.filePath) return null;
    writeFileSync(r.filePath, W().getDocumentBuffer(docId, versionId));
    return { path: r.filePath };
  });

  // 模板
  ipcMain.handle('workbench:uploadTemplate', async () => {
    const win = BrowserWindow.getAllWindows()[0];
    const r = await dialog.showOpenDialog(win, {
      filters: [{ name: 'Word 模板', extensions: ['docx'] }],
      properties: ['openFile'],
    });
    if (r.canceled || !r.filePaths.length) return null;
    const path = r.filePaths[0];
    const name = (path.split(/[\\/]/).pop() ?? '').replace(/\.docx$/i, '');
    return W().uploadTemplate(readFileSync(path), name);
  });
  ipcMain.handle('workbench:listTemplates', () => W().listTemplates());
  ipcMain.handle('workbench:readTemplate', (_e, { templateId }) => W().readTemplate(templateId));
  ipcMain.handle('workbench:instantiateTemplate', (_e, { templateId, values, name }) =>
    W().instantiateTemplate(templateId, values, name));

  // 内置规则集（#29：手写规则集 → 一键重排）
  ipcMain.handle('workbench:listBuiltinRulesets', () => W().listBuiltinRulesets());
  ipcMain.handle('workbench:builtinRulesetCommands', (_e, { rulesetId }) => W().builtinRulesetCommands(rulesetId));
  ipcMain.handle('workbench:applyBuiltinRuleset', (_e, { docId, rulesetId }) =>
    W().applyCommands(docId, W().builtinRulesetCommands(rulesetId), {
      source: 'ruleset', note: `按内置规则集 ${rulesetId} 重排`,
    }));

  // 配置
  ipcMain.handle('workbench:getConfig', () => W().getConfig());
  ipcMain.handle('workbench:setConfig', (_e, patch) => W().setConfig(patch));

  // agent
  ipcMain.handle('agent:send', async (_e, { docId, message }) => {
    if (!agentBridge) throw new Error('agent 未初始化');
    return agentBridge.send(docId, message);
  });
  ipcMain.handle('agent:status', () => agentBridge?.status() ?? { ready: false, reason: 'agent 未初始化' });
  ipcMain.handle('agent:abort', () => agentBridge?.abort());
}

// agent 事件 → 渲染层广播
function wireAgentEvents(win: BrowserWindow): void {
  if (!agentBridge) return;
  agentBridge.onEvent((event: unknown) => {
    if (!win.isDestroyed()) win.webContents.send('agent:event', event);
  });
}

async function createWindow(): Promise<void> {
  crashLog('[window] createWindow');
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 960,
    minHeight: 600,
    backgroundColor: '#ffffff',
    title: '排版工作台 paiban-studio',
    webPreferences: {
      preload: join(__dirname, 'preload.mjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false, // preload 需要 ESM 与 Buffer 桥接
    },
    show: !SMOKE,
  });
  win.removeMenu();
  if (process.env.PAIBAN_DEVTOOLS === '1') {
    win.webContents.openDevTools({ mode: 'detach' });
    // TEMP 调试：渲染层（含 iframe / worker）console 全量写入 crash.log
    win.webContents.on('console-message', (_e, level, message, line, sourceId) => {
      crashLog(`[renderer:${level}] ${message} (${sourceId}:${line})`);
    });
  }
  await win.loadURL(`http://127.0.0.1:${staticPort}/public/index.html`);
  crashLog('[window] renderer 加载完成');
  wireAgentEvents(win);
}

// headless 冒烟：PAIBAN_SMOKE=1 时执行 service 级全链路并退出（CI/无显示环境用）
async function runSmoke(): Promise<void> {
  const results: { steps: Array<Record<string, unknown>>; ok: boolean } = { steps: [], ok: true };
  const step = (name: string, fn: () => unknown): void => {
    try {
      const r = fn();
      results.steps.push({ name, ok: true, detail: r });
    } catch (err) {
      results.steps.push({ name, ok: false, error: (err as Error).message });
      results.ok = false;
    }
  };
  // 合成最小 docx 走全链路
  const PizZip = (await import('pizzip')).default;
  const zip = new PizZip();
  const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
  zip.file('[Content_Types].xml', DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', DECL + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>冒烟标题</w:t></w:r></w:p><w:p><w:r><w:t>冒烟正文。</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>');
  zip.file('word/_rels/document.xml.rels', DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');
  const buf = zip.generate({ type: 'nodebuffer' });

  let docId: string;
  step('uploadDocument', () => { docId = workspace.uploadDocument(buf, 'smoke.docx').docId; return docId; });
  step('applyCommands', () => {
    const r = workspace.applyCommands(docId, [{ command: 'set', path: '/body/p[1]', props: { align: 'center', run: { eastAsia: '黑体', sizePt: 16 } } }]);
    if (r.errors.length) throw new Error(JSON.stringify(r.errors));
    return r.version.id;
  });
  step('outline', () => workspace.getOutline(docId).paragraphCount + ' 段');
  step('versions', () => workspace.listVersions(docId).length + ' 个版本');
  step('rollback', () => workspace.rollback(docId, 'v1').version.id);
  step('templates', () => workspace.uploadTemplate(buf, '冒烟模板').extracted.join('/'));
  step('agentStatus', () => JSON.stringify(workspace.getConfig()));
  console.log('[SMOKE]', JSON.stringify(results));
  app.exit(results.ok ? 0 : 1);
}

app.whenReady().then(async () => {
  workspace = new Workspace(dataDir());
  crashLog(`[app:ready] workspace=${dataDir()}`);
  if (!SMOKE) {
    staticPort = await startStaticServer();
    crashLog(`[static] 回环静态服务器 http://127.0.0.1:${staticPort}`);
  }
  // agent-core 动态加载：SDK 缺失/无凭证时降级为不可用状态，工作台其余功能不受影响
  try {
    const { AgentBridge } = await import('./agent-core/bridge.js');
    agentBridge = new AgentBridge(workspace);
    await agentBridge.init();
  } catch (err) {
    console.error('[agent] 初始化失败（降级为无 agent 模式）:', (err as Error).message);
    agentBridge = null;
  }
  registerWorkbenchIpc();
  if (SMOKE) {
    await runSmoke();
    return;
  }
  await createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  crashLog('[window-all-closed] 所有窗口已关闭');
  if (process.platform !== 'darwin') app.quit();
});
