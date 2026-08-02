// main.ts — Electron 主进程（R1：主进程跑 agent 编排 + docx 编辑内核 + 存储层；IPC 通信）。
//
// 职责：
//   1. 创建 BrowserWindow（三栏工作台 UI，public/index.html）
//   2. 注册 IPC handlers（workbench:* 全部走 Workspace 服务层；agent:* 走 agent-core）
//   3. 数据目录：<userData>/paiban-studio（对象存储 + 版本链 + 模板库 + 配置）

import { app, BrowserWindow, ipcMain, dialog, shell } from 'electron';
import { join } from 'node:path';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { Workspace } from './server/workspace.js';

const __dirname = fileURLToPath(new URL('.', import.meta.url));
// 编译产物位于 dist/src/，项目根在其上两级（public/index.html、templates/ 等均位于项目根）。
const ROOT = join(__dirname, '../..');
const SMOKE = process.env.PAIBAN_SMOKE === '1';

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
  await win.loadFile(join(ROOT, 'public/index.html'));
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
  if (process.platform !== 'darwin') app.quit();
});
