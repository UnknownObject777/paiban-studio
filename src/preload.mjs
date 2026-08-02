// preload.js — contextBridge 桥接（contextIsolation 开启，渲染层只见到白名单 API）。

import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('paiban', {
  // 文档
  openDialog: () => ipcRenderer.invoke('workbench:openDialog'),
  uploadDocument: (name, bytes) => ipcRenderer.invoke('workbench:uploadDocument', { name, bytes }),
  listDocuments: () => ipcRenderer.invoke('workbench:listDocuments'),
  getBuffer: (docId, versionId) => ipcRenderer.invoke('workbench:getBuffer', { docId, versionId }),
  outline: (docId) => ipcRenderer.invoke('workbench:outline', { docId }),
  applyCommands: (docId, commands, note) => ipcRenderer.invoke('workbench:applyCommands', { docId, commands, note }),
  download: (docId, versionId) => ipcRenderer.invoke('workbench:download', { docId, versionId }),

  // 版本
  listVersions: (docId) => ipcRenderer.invoke('workbench:listVersions', { docId }),
  rollback: (docId, versionId) => ipcRenderer.invoke('workbench:rollback', { docId, versionId }),

  // 模板
  uploadTemplate: () => ipcRenderer.invoke('workbench:uploadTemplate'),
  listTemplates: () => ipcRenderer.invoke('workbench:listTemplates'),
  readTemplate: (templateId) => ipcRenderer.invoke('workbench:readTemplate', { templateId }),
  instantiateTemplate: (templateId, values, name) =>
    ipcRenderer.invoke('workbench:instantiateTemplate', { templateId, values, name }),

  // 配置
  getConfig: () => ipcRenderer.invoke('workbench:getConfig'),
  setConfig: (patch) => ipcRenderer.invoke('workbench:setConfig', patch),

  // agent
  agentSend: (docId, message) => ipcRenderer.invoke('agent:send', { docId, message }),
  agentStatus: () => ipcRenderer.invoke('agent:status'),
  agentAbort: () => ipcRenderer.invoke('agent:abort'),
  onAgentEvent: (fn) => {
    const listener = (_e, event) => fn(event);
    ipcRenderer.on('agent:event', listener);
    return () => ipcRenderer.removeListener('agent:event', listener);
  },
});
