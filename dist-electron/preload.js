"use strict";
const { contextBridge, ipcRenderer } = require("electron");
contextBridge.exposeInMainWorld("electronAPI", {
  // 批量转换PDF，支持options参数
  convertBatchToPdf: (files, options) => ipcRenderer.invoke("batch-convert-pdf", files, options),
  // 检查Office是否安装
  checkOfficeInstallation: () => ipcRenderer.invoke("check-office-installation")
});
