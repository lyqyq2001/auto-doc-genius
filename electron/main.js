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

// Office转换PDF
async function convertPdfByOffice(docxFiles) {
  try {
    // 检查Word是否安装
    if (!(await checkWordInstallation())) {
      return { success: false, error: '未检测到Microsoft Word安装' };
    }

    // 并行处理文件转换，每次处理2个文件，避免Office应用资源占用过高
    const batchSize = 2;
    const pdfResults = [];

    // 创建转换单个文件的函数
    const convertSingleFile = async file => {
      const tempDir = temp.mkdirSync('autodocgenius');
      const docxPath = path.join(tempDir, file.name);
      const pdfPath = path.join(tempDir, file.name.replace('.docx', '.pdf'));

      try {
        // 写入临时Word文件
        fs.writeFileSync(docxPath, Buffer.from(file.buffer));

        // 使用Office转换
        const success = await convertWordToPdfWithOffice(docxPath, pdfPath);

        if (success && fs.existsSync(pdfPath)) {
          // 读取转换后的PDF
          const pdfBuffer = fs.readFileSync(pdfPath);
          return {
            name: file.name.replace('.docx', '.pdf'),
            data: pdfBuffer,
          };
        } else {
          throw new Error(`转换失败: ${file.name}`);
        }
      } finally {
        // 清理临时文件
        if (fs.existsSync(docxPath)) fs.unlinkSync(docxPath);
        if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
        if (fs.existsSync(tempDir))
          fs.rmSync(tempDir, { recursive: true, force: true });
      }
    };

    // 分批处理文件
    for (let i = 0; i < docxFiles.length; i += batchSize) {
      const batch = docxFiles.slice(i, i + batchSize);
      const batchResults = await Promise.all(batch.map(convertSingleFile));
      pdfResults.push(...batchResults);
    }

    return { success: true, results: pdfResults };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

//  注册 IPC 处理函数
ipcMain.handle('batch-convert-pdf', async (event, docxFiles, options = {}) => {
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
