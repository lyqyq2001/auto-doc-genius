const { contextBridge, ipcRenderer } = require('electron');

// 暴露安全的API到渲染进程
contextBridge.exposeInMainWorld('electronAPI', {
  // 批量转换PDF
  convertBatchToPdf: files => ipcRenderer.invoke('batch-convert-pdf', files),
  // 检查Office是否安装
  checkOfficeInstallation: () =>
    ipcRenderer.invoke('check-office-installation'),
  // 下载Excel模板
  downloadExcelTemplate: () => ipcRenderer.invoke('download-excel-template'),
  // 下载Word模板
  downloadWordTemplate: () => ipcRenderer.invoke('download-word-template'),
  // 监听PDF转换进度
  onPdfProgress: callback => {
    const listener = (_event, data) => callback(data);
    ipcRenderer.on('pdf-conversion-progress', listener);
    return () =>
      ipcRenderer.removeListener('pdf-conversion-progress', listener);
  },
});
