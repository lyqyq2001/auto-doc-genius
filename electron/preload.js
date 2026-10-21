const { contextBridge, ipcRenderer, webUtils } = require('electron');

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
  // 获取文件在磁盘上的真实路径（用于解析Excel中的相对图片路径）
  getFilePath: file => webUtils.getPathForFile(file),
  // 批量读取图片文件，返回 { 路径: { ok, base64, width, height, ext } }
  readImageFiles: filePaths =>
    ipcRenderer.invoke('read-image-files', filePaths),
});
