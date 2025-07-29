const { contextBridge, ipcRenderer } = require('electron');

// API que se expondrá al proceso de renderizado
const api = {
    openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),
    processExcelAndExport: (filePath) => ipcRenderer.invoke('process-excel-and-export', filePath),
    readCleanedExcel: (filePath) => ipcRenderer.invoke('read-cleaned-excel', filePath),
    sendDataToDB: (data) => ipcRenderer.invoke('send-data-to-db', data),
    getSystemInfo: () => ipcRenderer.invoke('get-system-info')
};

// Exponer la API al proceso de renderizado
contextBridge.exposeInMainWorld('api', api);

module.exports = api;