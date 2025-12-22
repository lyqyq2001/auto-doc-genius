const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const temp = require('temp');
const {
  convertWordToPdfWithOffice,
  checkWordInstallation,
} = require('./officeConverter');

// 屏蔽安全警告
process.env.ELECTRON_DISABLE_SECURITY_WARNINGS = 'true';

// 主窗口引用
let mainWindow = null;

// 创建主窗口
const createWindow = () => {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  // 开发环境：加载 Vite 提供的 URL
  if (process.env.VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    // 生产环境：加载打包后的 index.html
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
  }
};

// 3. Office转换PDF
async function convertPdfByOffice(docxFiles) {
  const pdfResults = [];

  try {
    // 检查Word是否安装
    if (!checkWordInstallation()) {
      return { success: false, error: '未检测到Microsoft Word安装' };
    }

    // 串行处理循环
    for (const [_index, file] of docxFiles.entries()) {
      // 创建临时文件
      const tempDir = temp.mkdirSync('autodocgenius');
      const docxPath = path.join(tempDir, file.name);
      const pdfPath = path.join(tempDir, file.name.replace('.docx', '.pdf'));

      try {
        // 写入临时Word文件
        fs.writeFileSync(docxPath, Buffer.from(file.buffer));

        // 使用Office转换
        const success = convertWordToPdfWithOffice(docxPath, pdfPath);

        if (success && fs.existsSync(pdfPath)) {
          // 读取转换后的PDF
          const pdfBuffer = fs.readFileSync(pdfPath);
          pdfResults.push({
            name: file.name.replace('.docx', '.pdf'),
            data: pdfBuffer,
          });
        } else {
          return { success: false, error: `转换失败: ${file.name}` };
        }
      } finally {
        // 清理临时文件
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

// 4. 注册 IPC 处理函数
ipcMain.handle('batch-convert-pdf', async (event, docxFiles, options = {}) => {
  return convertPdfByOffice(docxFiles);
});

// 5. 检查Office是否安装
ipcMain.handle('check-office-installation', () => {
  return checkWordInstallation();
});

// 创建主窗口
app.whenReady().then(createWindow);

// 当所有窗口都关闭时退出应用
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
