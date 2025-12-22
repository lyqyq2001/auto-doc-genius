const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const temp = require('temp');

// 屏蔽安全警告
process.env.ELECTRON_DISABLE_SECURITY_WARNINGS = 'true';

// 主窗口引用
let mainWindow = null;

// 隐藏打印窗口引用
let workerWindow = null;

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

// === 核心：批量转换逻辑 ===

// 导入Office转换器
const { convertWordToPdfWithOffice, checkWordInstallation } = require('./officeConverter');
const os = require('os');


// 1. 初始化/获取隐藏窗口
async function getWorkerWindow() {
  if (workerWindow && !workerWindow.isDestroyed()) {
    return workerWindow;
  }

  workerWindow = new BrowserWindow({
    show: false, // 关键：隐藏
    width: 1200,  // 增加宽度，提高渲染精度
    height: 1700, // 增加高度，提高渲染精度
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  // 加载隐藏路由
  const loadUrl = process.env.VITE_DEV_SERVER_URL
    ? `${process.env.VITE_DEV_SERVER_URL}#/hidden-print`
    : `file://${path.join(__dirname, '../dist/index.html')}#/hidden-print`;

  await workerWindow.loadURL(loadUrl);
  return workerWindow;
}

// 2. 渲染引擎转换PDF
async function convertPdfByRendering(docxFiles) {
  const pdfResults = [];
  const worker = await getWorkerWindow();

  try {
    // 串行处理循环
    for (const [index, file] of docxFiles.entries()) {
      console.log(`正在转换第 ${index + 1}/${docxFiles.length} 个文件: ${file.name}`);

      // A. 发送 Buffer 给渲染进程
      worker.webContents.send('render-docx', file.buffer);

      // B. 等待渲染完成信号 (Promise 包装一次性事件监听)
      await new Promise((resolve, reject) => {
        const cleanUp = () => {
          ipcMain.removeListener('render-done', onDone);
          ipcMain.removeListener('render-error', onError);
        };

        const onDone = () => {
          cleanUp();
          resolve();
        };

        const onError = (e, msg) => {
          cleanUp();
          reject(new Error(msg));
        };

        ipcMain.on('render-done', onDone);
        ipcMain.on('render-error', onError);
        
        // 5秒超时防止卡死
        setTimeout(() => {
            cleanUp();
            // 超时也 resolve，只是记录错误，不要打断整个队列
            console.error(`File ${file.name} timeout`);
            resolve();
        }, 5000);
      });

      // C. 生成 PDF
      const pdfData = await worker.webContents.printToPDF({
        printBackground: true, // 确保背景内容和图片正确渲染
        pageSize: 'A4', // A4页面大小
        margins: { top: 0, bottom: 0, left: 0, right: 0 }, // 使用0边距，docx-preview自带边距
        scale: 1.0, // 1:1缩放，保持原始大小
        preferCSSPageSize: true, // 优先使用CSS定义的页面大小
        printSelectionOnly: false, // 打印整个页面
        landscape: false, // 纵向打印
        pageRanges: '', // 打印所有页面
        ignoreInvalidPageRanges: true, // 忽略无效页面范围
        displayHeaderFooter: false, // 不显示默认页眉页脚
        headerTemplate: '', // 自定义页眉模板（空）
        footerTemplate: '', // 自定义页脚模板（空）
        generateTaggedPDF: true, // 生成带标签的PDF，提高可访问性
        optimizeForSpeed: false // 优先质量而非速度
      });

      // D. 收集结果 (返回 Buffer 给前端打包 ZIP)
      pdfResults.push({
        name: file.name.replace('.docx', '.pdf'),
        data: pdfData
      });
    }

    return { success: true, results: pdfResults };

  } catch (error) {
    console.error('批量转换错误:', error);
    return { success: false, error: error.message };
  } finally {
    // 任务全部结束后销毁窗口释放内存
    if (worker) worker.destroy();
    workerWindow = null;
  }
}

// 3. Office转换PDF
async function convertPdfByOffice(docxFiles) {
  const pdfResults = [];
  
  try {
    // 检查Word是否安装
    if (!checkWordInstallation()) {
      return { success: false, error: '未检测到Microsoft Word安装' };
    }
    
    // 串行处理循环
    for (const [index, file] of docxFiles.entries()) {
      console.log(`正在使用Office转换第 ${index + 1}/${docxFiles.length} 个文件: ${file.name}`);
      
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
            data: pdfBuffer
          });
        } else {
          console.error(`转换失败: ${file.name}`);
          return { success: false, error: `转换失败: ${file.name}` };
        }
      } finally {
        // 清理临时文件
        if (fs.existsSync(docxPath)) fs.unlinkSync(docxPath);
        if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
       if (fs.existsSync(tempDir)) fs.rmSync(tempDir, { recursive: true, force: true });
      }
    }
    
    return { success: true, results: pdfResults };
  } catch (error) {
    console.error('Office批量转换错误:', error);
    return { success: false, error: error.message };
  }
}

// 4. 注册 IPC 处理函数
ipcMain.handle('batch-convert-pdf', async (event, docxFiles, options = {}) => {
  // docxFiles 结构: [{ name: 'doc1.docx', buffer: ArrayBuffer }, ...]
  // options 结构: { method: 'render' | 'office' }
  
  const { method = 'render' } = options;
  
  if (method === 'office') {
    // 使用Office转换
    return convertPdfByOffice(docxFiles);
  } else {
    // 默认使用渲染引擎转换
    return convertPdfByRendering(docxFiles);
  }
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
