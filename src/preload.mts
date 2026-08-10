// preload.mts — contextBridge 桥接（contextIsolation 开启，渲染层只见到白名单 API）。
// 编译产物为 preload.mjs（Electron ESM preload 需要 .mjs 扩展名）。

import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('paiban', {
  // 文档
  openDialog: () => ipcRenderer.invoke('workbench:openDialog'),
  uploadDocument: (name: string, bytes: ArrayBuffer) => ipcRenderer.invoke('workbench:uploadDocument', { name, bytes }),
  listDocuments: () => ipcRenderer.invoke('workbench:listDocuments'),
  getBuffer: (docId: string, versionId?: string) => ipcRenderer.invoke('workbench:getBuffer', { docId, versionId }),
  outline: (docId: string) => ipcRenderer.invoke('workbench:outline', { docId }),
  applyCommands: (docId: string, commands: unknown[], note?: string) => ipcRenderer.invoke('workbench:applyCommands', { docId, commands, note }),
  download: (docId: string, versionId?: string) => ipcRenderer.invoke('workbench:download', { docId, versionId }),

  // 版本
  listVersions: (docId: string) => ipcRenderer.invoke('workbench:listVersions', { docId }),
  rollback: (docId: string, versionId: string) => ipcRenderer.invoke('workbench:rollback', { docId, versionId }),

  // 模板
  uploadTemplate: () => ipcRenderer.invoke('workbench:uploadTemplate'),
  listTemplates: () => ipcRenderer.invoke('workbench:listTemplates'),
  readTemplate: (templateId: string) => ipcRenderer.invoke('workbench:readTemplate', { templateId }),
  instantiateTemplate: (templateId: string, values: Record<string, unknown>, name?: string) =>
    ipcRenderer.invoke('workbench:instantiateTemplate', { templateId, values, name }),

  // 内置规则集
  listBuiltinRulesets: () => ipcRenderer.invoke('workbench:listBuiltinRulesets'),
  applyBuiltinRuleset: (docId: string, rulesetId: string) =>
    ipcRenderer.invoke('workbench:applyBuiltinRuleset', { docId, rulesetId }),

  // 配置
  getConfig: () => ipcRenderer.invoke('workbench:getConfig'),
  setConfig: (patch: Record<string, unknown>) => ipcRenderer.invoke('workbench:setConfig', patch),

  // agent
  agentSend: (docId: string, message: string) => ipcRenderer.invoke('agent:send', { docId, message }),
  agentStatus: () => ipcRenderer.invoke('agent:status'),
  agentAbort: () => ipcRenderer.invoke('agent:abort'),
  onAgentEvent: (fn: (event: unknown) => void) => {
    const listener = (_e: Electron.IpcRendererEvent, event: unknown) => fn(event);
    ipcRenderer.on('agent:event', listener);
    return () => ipcRenderer.removeListener('agent:event', listener);
  },
});
