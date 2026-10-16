const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('akhtar', {
  openAny: () => ipcRenderer.invoke('file:openAny'),
  readExternal: (filePath) => ipcRenderer.invoke('file:readExternal',filePath),
  saveCode: (payload) => ipcRenderer.invoke('code:save',payload),
  onExternalFiles: (callback) => {const fn=(_event,paths)=>callback(paths);ipcRenderer.on('external:files',fn);return ()=>ipcRenderer.removeListener('external:files',fn)},
  openProject: () => ipcRenderer.invoke('file:openProject'),
  saveProject: (payload) => ipcRenderer.invoke('file:saveProject', payload),
  saveText: (payload) => ipcRenderer.invoke('file:saveText', payload),
  pickImage: () => ipcRenderer.invoke('file:pickImage'),
  printPdf: (suggestedName) => ipcRenderer.invoke('window:printPdf', suggestedName),
  exportPptx: (payload) => ipcRenderer.invoke('slides:exportPptx', payload),
  getVersion: () => ipcRenderer.invoke('app:getVersion'),
  writeRecovery: (data) => ipcRenderer.invoke('file:writeRecovery',data),
  readRecovery: () => ipcRenderer.invoke('file:readRecovery'),
  exportDocx: (data) => ipcRenderer.invoke('docs:exportDocx',data)
});
