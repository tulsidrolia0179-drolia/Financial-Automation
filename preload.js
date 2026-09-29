const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
    selectFile: () => ipcRenderer.invoke('dialog:selectFile'),
    processFile: (filePath, taskType) => ipcRenderer.invoke('process:file', filePath, taskType)
});
