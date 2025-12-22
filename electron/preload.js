const { contextBridge, ipcRenderer } = require('electron');

// 暴露安全的API到渲染进程
contextBridge.exposeInMainWorld('electronAPI', {
  // 批量转换PDF，支持options参数
  convertBatchToPdf: (files, options) => ipcRenderer.invoke('batch-convert-pdf', files, options),
  
  // 检查Office是否安装
  checkOfficeInstallation: () => ipcRenderer.invoke('check-office-installation'),

  // 隐藏窗口使用：监听渲染任务
  onRenderDocx: (callback) =>
    ipcRenderer.on('render-docx', (event, buffer) => callback(buffer)),

  // 隐藏窗口使用：通知渲染完成
  sendRenderDone: () => ipcRenderer.send('render-done'),
  sendRenderError: (msg) => ipcRenderer.send('render-error', msg),
});