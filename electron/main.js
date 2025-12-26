const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
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

async function convertPdfByOffice(docxFiles, sendProgress) {
  try {
    const fileCount = docxFiles.length;
    const cpuCount = os.cpus().length;

    sendProgress?.({
      stage: 'init',
      message: `准备转换 ${fileCount} 个文件...`,
    });

    let parallelCount;
    if (fileCount <= 5) {
      parallelCount = 1;
    } else if (fileCount <= 20) {
      parallelCount = Math.min(2, cpuCount);
    } else if (fileCount <= 50) {
      parallelCount = Math.min(3, cpuCount);
    } else {
      parallelCount = Math.min(4, cpuCount);
    }

    const batches = [];
    const batchSize = Math.ceil(fileCount / parallelCount);
    for (let i = 0; i < fileCount; i += batchSize) {
      batches.push(docxFiles.slice(i, i + batchSize));
    }

    console.log(
      `[PDF转换] 文件数: ${fileCount}, 并行批次: ${batches.length}, 每批次: ${batchSize}`
    );
    sendProgress?.({
      stage: 'converting',
      message: `开始转换，使用 ${batches.length} 个并行批次...`,
    });

    let completedBatches = 0;
    const batchPromises = batches.map(async (batch, batchIndex) => {
      const tempDir = temp.mkdirSync(`autodocgenius_batch_${batchIndex}`);
      const inputOutputPairs = [];

      for (const file of batch) {
        const docxPath = path.join(tempDir, file.name);
        const pdfPath = path.join(tempDir, file.name.replace('.docx', '.pdf'));
        fs.writeFileSync(docxPath, Buffer.from(file.buffer));
        inputOutputPairs.push({ input: docxPath, output: pdfPath });
      }

      const batchResult = await convertBatchWordToPdf(
        inputOutputPairs,
        progress => {
          sendProgress?.({
            stage: 'converting',
            batchIndex: batchIndex + 1,
            totalBatches: batches.length,
            progress: progress,
            message: `批次 ${batchIndex + 1}/${batches.length}: ${progress}%`,
          });
        }
      );

      try {
        if (fs.existsSync(tempDir)) {
          fs.rmSync(tempDir, { recursive: true, force: true });
        }
      } catch (e) {
        console.warn(`清理临时目录失败 [批次${batchIndex}]:`, e);
      }

      completedBatches++;
      sendProgress?.({
        stage: 'converting',
        progress: Math.round((completedBatches / batches.length) * 100),
        message: `已完成 ${completedBatches}/${batches.length} 个批次`,
      });

      return { results: batchResult.results || [], error: batchResult.error };
    });

    const batchResults = await Promise.all(batchPromises);

    const pdfResults = [];
    const errors = [];

    batchResults.forEach((result, index) => {
      console.log(`[批次${index + 1}] 结果:`, result);
      if (result.results && result.results.length > 0) {
        console.log(
          `[批次${index + 1}] 成功生成 ${result.results.length} 个PDF文件`
        );
        pdfResults.push(...result.results);
      } else {
        console.warn(`[批次${index + 1}] 没有生成PDF文件`);
      }
      if (result.error) {
        errors.push(`批次${index + 1}: ${result.error}`);
      }
    });

    console.log(
      `[PDF转换] 总共生成 ${pdfResults.length} 个PDF文件，期望 ${fileCount} 个`
    );

    if (errors.length > 0) {
      console.warn('[PDF] 部分批次转换失败:', errors.join('; '));
    }

    sendProgress?.({ stage: 'completed', progress: 100, message: '转换完成' });

    return { success: true, results: pdfResults };
  } catch (error) {
    sendProgress?.({ stage: 'error', message: `转换失败: ${error.message}` });
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
