# AutoDoc Genius

## 项目概述

**AutoDoc Genius** 是一个基于 Electron + Vue 3 的桌面应用程序，用于批量生成 Word/PDF 报告。该工具能够读取 Excel 数据表中的多行数据，结合 Word 模板文件，自动生成对应数量的 Word 或 PDF 报告文件。

## 技术栈

### 前端技术
- **Vue 3** - 现代化前端框架
- **Element Plus** - Vue 3 组件库，提供美观的 UI 组件
- **Vite** - 现代化的前端构建工具
- **XLSX.js** - Excel 文件解析库
- **Docxtemplater** - Word 模板渲染引擎
- **Pizzip** - ZIP 文件处理库
- **JSZip** - 创建和读取 ZIP 文件
- **File-saver** - 文件下载工具

### 桌面应用技术
- **Electron** - 跨平台桌面应用框架
- **Vite-plugin-electron** - Vite 与 Electron 集成插件

### 文档转换技术
- **PowerShell 脚本** - 调用 Office 应用进行 Word 转 PDF
- **Microsoft Word COM 接口** - 文档格式转换
- **WPS Office 兼容** - 支持 WPS 作为 Office 替代品

## 功能特性

### 核心功能
1. **Excel 数据读取**
   - 支持 .xlsx 和 .xls 格式
   - 自动解析表头和数据行
   - 日期和时间格式自动处理

2. **Word 模板处理**
   - 支持 .docx 模板文件
   - 占位符替换机制（使用 XXX01、XXX02 等标记）
   - 动态生成个性化文档

3. **批量文档生成**
   - 支持导出为 Word (.docx) 格式
   - 支持导出为 PDF 格式（需要 Office 环境）
   - 进度条显示生成进度
   - 自动生成 ZIP 压缩包

4. **Office 集成**
   - 检测 Microsoft Word 安装状态
   - 支持 WPS Office 作为替代方案
   - 通过 PowerShell 脚本调用 COM 接口


## 项目结构

```
vite-project/
├── electron/                    # Electron 主进程代码
│   ├── main.js                 # 主进程入口文件
│   ├── preload.js              # 预加载脚本
│   └── officeConverter.js      # Office 文档转换模块
├── src/
│   ├── App.vue                 # 主应用组件
│   └── main.js                 # Vue 应用入口
├── public/                     # 静态资源
├── package.json                # 项目配置
├── vite.config.js              # Vite 配置文件
├── auto-imports.d.ts          # 自动导入类型定义
└── components.d.ts            # 组件类型定义

```



## 使用方法

1. **准备 Excel 数据表**
   - 第一行：描述信息（可选）
   - 第二行：字段名（用于模板占位符）
   - 第三行起：数据行
2. **准备 Word 模板**
   - 使用 XXX01、XXX02 等格式标记占位符
   - 占位符需与 Excel 字段名对应
3. **上传文件并生成**
   - 上传 Excel 文件和 Word 模板
   - 选择导出格式（Word 或 PDF）
   - 点击生成按钮
   - 下载生成的 ZIP 文件

## 其它说明

### Word 转 PDF 性能测试工具

项目提供了一个 `WordToPdf.bat` 批处理脚本，用于测试原生 Word 转 PDF 的转换速度，方便对比和验证应用性能，由于该应用需要先将excel转换为word，再转成pdf，速度上会比原生的慢一些。

#### 使用方法

1. 将需要转换的 `.docx` 文件放在与 `WordToPdf.bat` 相同的目录下
2. 双击运行 `WordToPdf.bat` 脚本
3. 脚本会自动将当前目录下所有 `.docx` 文件转换为 `.pdf` 文件
4. 转换完成后按任意键退出

#### 注意事项

- 需要安装 Microsoft Word 或 WPS Office
- 转换过程中 Word 应用会在后台运行
- 转换速度取决于文档复杂度和计算机性能