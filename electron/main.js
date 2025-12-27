const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const {
  checkWordInstallation,
  convertBatchWordToPdf,
} = require('./officeConverter');

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

async function convertPdfByOffice(docxFiles, sendProgress) {
  try {
    const fileCount = docxFiles.length;

    sendProgress?.({
      stage: 'init',
      message: `准备转换 ${fileCount} 个文件...`,
    });

    const tempDir = path.join(os.tmpdir(), `autodocgenius_batch_${Date.now()}`);
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    const inputOutputPairs = [];

    for (let i = 0; i < fileCount; i++) {
      const file = docxFiles[i];
      const docxPath = path.join(tempDir, file.name);
      const pdfPath = path.join(tempDir, file.name.replace('.docx', '.pdf'));
      fs.writeFileSync(docxPath, Buffer.from(file.buffer));
      inputOutputPairs.push({ input: docxPath, output: pdfPath });
    }

    sendProgress?.({
      stage: 'converting',
      message: `开始转换 ${fileCount} 个文件...`,
    });

    try {
      const batchResult = await convertBatchWordToPdf(inputOutputPairs);

      if (batchResult.results && batchResult.results.length > 0) {
        console.log(
          `[PDF转换] 转换完成，成功 ${batchResult.results.length}/${fileCount} 个文件`
        );
        sendProgress?.({
          stage: 'completed',
          progress: 100,
          message: '转换完成',
        });
        return { success: true, results: batchResult.results };
      }
    } finally {
      if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      }
    }
  } catch (error) {
    sendProgress?.({ stage: 'error', message: `转换失败: ${error.message}` });
    console.error('[PDF转换] 总处理失败:', error);
    return { success: false, error: error.message };
  }
}

//  注册 IPC 处理函数
ipcMain.handle('batch-convert-pdf', async (_event, docxFiles) => {
  const sendProgress = data => {
    _event.sender.send('pdf-conversion-progress', data);
  };
  return convertPdfByOffice(docxFiles, sendProgress);
});

//  检查Office是否安装
ipcMain.handle('check-office-installation', () => {
  return checkWordInstallation();
});

// 下载Excel模板
ipcMain.handle('download-excel-template', async () => {
  try {
    // 尝试从多个位置读取模板文件
    let templatePath;

    // 开发模式下的路径
    const devPaths = [
      path.join(__dirname, '../Excel Template.xls'),
      path.join(process.cwd(), 'Excel Template.xls'),
      path.join(__dirname, '../../Excel Template.xls'),
    ];

    // 查找存在的路径
    for (const p of devPaths) {
      if (fs.existsSync(p)) {
        templatePath = p;
        break;
      }
    }

    if (!templatePath) {
      throw new Error('Excel模板文件未找到');
    }

    const buffer = fs.readFileSync(templatePath);
    return {
      success: true,
      buffer: buffer.toString('base64'),
      filename: 'Excel Template.xls',
    };
  } catch (error) {
    console.error('下载Excel模板失败:', error);
    return {
      success: false,
      error: error.message,
    };
  }
});

//  下载Word模板
ipcMain.handle('download-word-template', async () => {
  try {
    // 尝试从多个位置读取模板文件
    let templatePath;

    // 开发模式下的路径
    const devPaths = [
      path.join(__dirname, '../Word Template.docx'),
      path.join(process.cwd(), 'Word Template.docx'),
      path.join(__dirname, '../../Word Template.docx'),
    ];

    // 查找存在的路径
    for (const p of devPaths) {
      if (fs.existsSync(p)) {
        templatePath = p;
        break;
      }
    }

    if (!templatePath) {
      throw new Error('Word模板文件未找到');
    }

    const buffer = fs.readFileSync(templatePath);
    return {
      success: true,
      buffer: buffer.toString('base64'),
      filename: 'Word Template.docx',
    };
  } catch (error) {
    console.error('下载Word模板失败:', error);
    return {
      success: false,
      error: error.message,
    };
  }
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
