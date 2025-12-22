"use strict";
const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const temp = require("temp");
const {
  convertWordToPdfWithOffice,
  checkWordInstallation
} = require("./officeConverter");
process.env.ELECTRON_DISABLE_SECURITY_WARNINGS = "true";
let mainWindow = null;
const createWindow = () => {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, "preload.js")
    }
  });
  if (process.env.VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    mainWindow.loadFile(path.join(__dirname, "../dist/index.html"));
  }
};
async function convertPdfByOffice(docxFiles) {
  const pdfResults = [];
  try {
    if (!checkWordInstallation()) {
      return { success: false, error: "未检测到Microsoft Word安装" };
    }
    for (const [_index, file] of docxFiles.entries()) {
      const tempDir = temp.mkdirSync("autodocgenius");
      const docxPath = path.join(tempDir, file.name);
      const pdfPath = path.join(tempDir, file.name.replace(".docx", ".pdf"));
      try {
        fs.writeFileSync(docxPath, Buffer.from(file.buffer));
        const success = convertWordToPdfWithOffice(docxPath, pdfPath);
        if (success && fs.existsSync(pdfPath)) {
          const pdfBuffer = fs.readFileSync(pdfPath);
          pdfResults.push({
            name: file.name.replace(".docx", ".pdf"),
            data: pdfBuffer
          });
        } else {
          return { success: false, error: `转换失败: ${file.name}` };
        }
      } finally {
        if (fs.existsSync(docxPath)) fs.unlinkSync(docxPath);
        if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
        if (fs.existsSync(tempDir))
          fs.rmSync(tempDir, { recursive: true, force: true });
      }
    }
    return { success: true, results: pdfResults };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
ipcMain.handle("batch-convert-pdf", async (event, docxFiles, options = {}) => {
  return convertPdfByOffice(docxFiles);
});
ipcMain.handle("check-office-installation", () => {
  return checkWordInstallation();
});
app.whenReady().then(createWindow);
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
