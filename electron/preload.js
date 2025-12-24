const { contextBridge, ipcRenderer } = require('electron');

// 暴露安全的API到渲染进程
contextBridge.exposeInMainWorld('electronAPI', {
  // 批量转换PDF，支持options参数
  convertBatchToPdf: (files, options) =>
    ipcRenderer.invoke('batch-convert-pdf', files, options),
  // 检查Office是否安装
  checkOfficeInstallation: () =>
    ipcRenderer.invoke('check-office-installation'),
  // 下载Excel模板
  downloadExcelTemplate: () => ipcRenderer.invoke('download-excel-template'),
  // 下载Word模板
  downloadWordTemplate: () => ipcRenderer.invoke('download-word-template'),
});
