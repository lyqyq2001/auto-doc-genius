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

    sendProgress?.({
      stage: 'init',
      message: `准备转换 ${fileCount} 个文件...`,
    });

    const cpuCount = Math.max(1, os.cpus().length - 1);
    const maxParallelBatches = Math.min(2, cpuCount); // 最多2个并行批次
    const batchSize = Math.ceil(fileCount / maxParallelBatches);
    const batches = [];

    for (let i = 0; i < fileCount; i += batchSize) {
      batches.push(docxFiles.slice(i, i + batchSize));
    }

    sendProgress?.({
      stage: 'converting',
      message: `开始转换，使用 ${batches.length} 个并行批次...`,
    });

    const pdfResults = [];
    const errors = [];

    // 并行处理每个批次
    const batchPromises = batches.map(async (batch, batchIndex) => {
      try {
        const tempDir = temp.mkdirSync(`autodocgenius_batch_${batchIndex}`);
        const inputOutputPairs = [];

        for (const file of batch) {
          const docxPath = path.join(tempDir, file.name);
          const pdfPath = path.join(
            tempDir,
            file.name.replace('.docx', '.pdf')
          );
          fs.writeFileSync(docxPath, Buffer.from(file.buffer));
          inputOutputPairs.push({ input: docxPath, output: pdfPath });
        }

        await new Promise(resolve => setTimeout(resolve, 500));

        const batchResult = await convertBatchWordToPdf(inputOutputPairs);

        if (fs.existsSync(tempDir)) {
          fs.rmSync(tempDir, { recursive: true, force: true });
        }

        return {
          results: batchResult.results || [],
          error: batchResult.error,
          batchIndex: batchIndex + 1,
        };
      } catch (error) {
        console.error(`[批次${batchIndex + 1}] 处理失败:`, error.message);
        return {
          results: [],
          error: error.message,
          batchIndex: batchIndex + 1,
        };
      }
    });

    // 等待所有批次完成
    const batchResults = await Promise.all(batchPromises);

    // 处理结果
    let completedBatches = 0;
    let completedFiles = 0;

    for (const result of batchResults) {
      completedBatches++;
      completedFiles += result.results.length;

      if (result.results.length > 0) {
        pdfResults.push(...result.results);
      }

      if (result.error) {
        errors.push(`批次${result.batchIndex}: ${result.error}`);
      }

      // 更新进度
      const progress = Math.round((completedBatches / batches.length) * 100);
      sendProgress?.({
        stage: 'converting',
        progress: progress,
        message: `已完成 ${completedBatches}/${batches.length} 个批次，成功 ${completedFiles}/${fileCount} 个文件`,
      });
    }

    // 单独重试转换失败的文件
    if (completedFiles < fileCount) {
      // 找出失败的文件
      const successfulFileNames = new Set(
        pdfResults.map(pdf => pdf.name.replace('.pdf', '.docx'))
      );
      const allFiles = docxFiles;

      const failedFiles = allFiles.filter(
        file => !successfulFileNames.has(file.name)
      );

      if (failedFiles.length > 0) {
        // 逐个重试失败的文件
        for (let i = 0; i < failedFiles.length; i++) {
          const file = failedFiles[i];

          try {
            // 为每个重试文件创建独立的临时目录
            const retryTempDir = temp.mkdirSync(`autodocgenius_retry_${i}`);
            const docxPath = path.join(retryTempDir, file.name);
            const pdfPath = path.join(
              retryTempDir,
              file.name.replace('.docx', '.pdf')
            );

            fs.writeFileSync(docxPath, Buffer.from(file.buffer));

            await new Promise(resolve => setTimeout(resolve, 1500));

            // 单独转换这个文件
            const retryResult = await convertBatchWordToPdf([
              { input: docxPath, output: pdfPath },
            ]);

            if (retryResult.results && retryResult.results.length > 0) {
              pdfResults.push(...retryResult.results);
              completedFiles++;
            }

            if (fs.existsSync(retryTempDir)) {
              fs.rmSync(retryTempDir, { recursive: true, force: true });
            }

            // 更新进度
            sendProgress?.({
              stage: 'converting',
              progress: Math.round(
                ((completedBatches + (i + 1) / failedFiles.length) /
                  batches.length) *
                  100
              ),
              message: `已完成 ${completedBatches}/${batches.length} 个批次，成功 ${completedFiles}/${fileCount} 个文件`,
            });
          } catch (retryError) {
            console.error(`[PDF转换] 重试文件 ${file.name} 失败`);
          }
        }
      }
    }

    console.log(
      `[PDF转换] 转换完成，成功 ${pdfResults.length}/${fileCount} 个文件`
    );

    if (errors.length > 0) {
      console.warn('[PDF] 部分批次转换失败:', errors.join('; '));
    }

    sendProgress?.({ stage: 'completed', progress: 100, message: '转换完成' });

    return { success: true, results: pdfResults };
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
