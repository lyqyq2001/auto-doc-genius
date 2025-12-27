const { contextBridge, ipcRenderer } = require('electron');

// 暴露安全的API到渲染进程
contextBridge.exposeInMainWorld('electronAPI', {
  // 批量转换PDF
  convertBatchToPdf: files => ipcRenderer.invoke('batch-convert-pdf', files),
  // 检查Office是否安装
  checkOfficeInstallation: () =>
    ipcRenderer.invoke('check-office-installation'),
  // 监听PDF转换进度
  onPdfProgress: callback => {
    const listener = (_event, data) => callback(data);
    ipcRenderer.on('pdf-conversion-progress', listener);
    return () =>
      ipcRenderer.removeListener('pdf-conversion-progress', listener);
  },
});
