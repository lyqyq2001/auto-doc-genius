const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const temp = require('temp');
const {
  checkWordInstallation,
  convertBatchWordToPdf,
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

// Office转换PDF - 优化版本（多实例并行）
async function convertPdfByOffice(docxFiles) {
  try {
    // 根据文件数量确定并行度
    const fileCount = docxFiles.length;
    const parallelCount = Math.min(Math.max(2, Math.ceil(fileCount / 3)), 4);

    // 将文件分成多个批次
    const batches = [];
    const batchSize = Math.ceil(fileCount / parallelCount);
    for (let i = 0; i < fileCount; i += batchSize) {
      batches.push(docxFiles.slice(i, i + batchSize));
    }

    const pdfResults = [];
    // 并行处理每个批次
    const batchPromises = batches.map(async (batch, batchIndex) => {
      const tempDir = temp.mkdirSync(`autodocgenius_batch_${batchIndex}`);
      const inputOutputPairs = [];
      for (const file of batch) {
        const docxPath = path.join(tempDir, file.name);
        const pdfPath = path.join(tempDir, file.name.replace('.docx', '.pdf'));
        fs.writeFileSync(docxPath, Buffer.from(file.buffer));

        inputOutputPairs.push({ input: docxPath, output: pdfPath });
      }

      // 使用批量转换函数
      const r = await convertBatchWordToPdf(inputOutputPairs);

      if (r.success) {
        for (const pair of inputOutputPairs) {
          const pdfPath = pair.output;

          const pdfBuffer = fs.readFileSync(pdfPath);
          pdfResults.push({
            name: path.basename(pdfPath),
            data: pdfBuffer,
          });
        }
      }

      // 清理临时文件
      try {
        if (fs.existsSync(tempDir)) {
          fs.rmSync(tempDir, { recursive: true, force: true });
        }
      } catch (e) {
        console.warn(`清理临时目录失败 [批次${batchIndex}]:`, e);
      }
    });

    await Promise.all(batchPromises);

    return { success: true, results: pdfResults };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

//  注册 IPC 处理函数
ipcMain.handle('batch-convert-pdf', async (_event, docxFiles) => {
  return convertPdfByOffice(docxFiles);
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
