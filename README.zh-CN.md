# AutoDoc Genius

[English](./README.md) | **简体中文**

## 项目概述

**AutoDoc Genius** 是一个基于 Electron + Vue 3 的桌面应用程序，用于批量生成 Word/PDF 报告。该工具能够读取 Excel 数据表中的多行数据，结合 Word 模板文件，自动生成对应数量的 Word 或 PDF 报告文件。

### 核心功能

1. **Excel 数据读取**
   - 支持 .xlsx 和 .xls 格式
   - 自动解析表头和数据行
   - 日期和时间格式自动处理

2. **Word 模板处理**
   - 支持 .docx 模板文件
   - 占位符替换机制（使用 XXX01、XXX02 等标记）
   - 动态生成个性化文档

3. **多选项支持**
   - 支持复选框模板功能
   - 使用 `{{check01XXX}}{{check02XXX}}`等 格式标记多选项占位符
   - Excel 中使用逗号、中文逗号或顿号分隔多个选项
   - 自动将选中的复选框标记为 `☑`，未选中的保持 `□`

4. **图片替换支持**
   - Excel 中以 `img` 开头（不区分大小写）的列识别为图片列（如 `img01`）
   - 单元格中填写图片文件路径，支持绝对路径与相对路径（相对于 Excel 文件所在目录）
   - Word 模板中直接书写图片占位符（如 `img01`），生成时替换为真实图片
   - 支持 png / jpg / jpeg / gif / bmp 格式，按原始宽高比嵌入，超过 15cm 宽自动等比缩放
   - 图片缺失时对应位置留空并给出警告

5. **批量文档生成**
   - 支持导出为 Word (.docx) 格式
   - 支持导出为 PDF 格式（需要 Office 环境）
   - 进度条显示生成进度
   - 自动生成 ZIP 压缩包
6. **Office 集成**
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
│   ├── utils/                  # 工具函数
│   │   ├── documentUtils.js    # 文档处理工具
│   │   ├── docxGenerator.js    # Word 文档生成器
│   │   ├── imageUtils.js       # 图片替换工具
│   │   └── fileUtils.js        # 文件操作工具
│   ├── App.vue                 # 主应用组件
│   └── main.js                 # Vue 应用入口
├── public/                     # 静态资源
│   ├── Excel Template.xls      # Excel 模板示例
│   ├── Word Template.docx      # Word 模板示例
│   └── vite.svg                # Vite 图标
├── .vscode/                    # VS Code 配置
├── package.json                # 项目配置
├── vite.config.js              # Vite 配置文件
├── auto-imports.d.ts           # 自动导入类型定义
└── components.d.ts             # 组件类型定义
```

## 安装与运行

### 安装依赖

```bash
npm install
```

### 开发模式

```bash
npm run dev
```

### 构建应用

```bash
npm run build
```

构建完成后，安装包将生成在 `release` 目录中。

## 使用方法

1. **准备 Excel 数据表**
   - 第一行：描述信息（可选）
   - 第二行：字段名（用于模板占位符）
   - 第三行起：数据行

2. **准备 Word 模板**
   - 使用 XXX01、XXX02 等格式标记普通占位符
   - 占位符需与 Excel 字段名对应
   - 使用 `{{checkXXX}}` 格式标记多选项占位符
   - 在 `{{checkXXX}}` 中使用 `□ 选项1 □ 选项2 □ 选项3` 格式定义选项

3. **多选项使用示例**
   - **Excel 表头**: `check01`
   - **Excel 数据**: `游泳,跑步,阅读` （使用逗号、中文逗号或顿号分隔）
   - **Word 模板**: `{{check01□ 游泳 □ 跑步 □ 阅读 □ 编程}}`
   - **生成结果**: `☑ 游泳 ☑ 跑步 ☑ 阅读 □ 编程`

4. **图片替换使用示例**
   - **Excel 表头**: `img01`
   - **Excel 数据**: `photos/photo1.jpg`（相对路径，相对于 Excel 文件所在目录）或 `D:\pics\photo1.jpg`（绝对路径）
   - **Word 模板**: 在需要插入图片的位置直接书写 `img01`（建议单独占用一个段落）
   - **生成结果**: 占位符被替换为对应图片

5. **上传文件并生成**
   - 上传 Excel 文件和 Word 模板
   - 选择导出格式（Word 或 PDF）
   - 点击生成按钮
   - 下载生成的 ZIP 文件

## 注意事项

- PDF 导出功能需要安装 Microsoft Word 或 WPS Office
- 确保模板中的占位符与 Excel 表头完全匹配
- 大批量文档生成可能需要较长时间，请耐心等待
- 生成的 ZIP 文件包含所有生成的文档
- 图片替换要求 Excel 中填写的图片路径真实存在，支持 png / jpg / jpeg / gif / bmp 格式
- GIF 动图在 Word 文档中仅显示第一帧

## 常见问题

**Q: PDF 导出失败怎么办？**

A: 请检查是否安装了 Microsoft Word 或 WPS Office，并确保可以正常打开 Word 文档。

**Q: 生成的文档格式不正确？**

A: 请检查 Word 模板文件格式是否正确，占位符是否与 Excel 表头匹配。

**Q: 如何自定义模板？**

A: 在 Word 文档中使用 XXX01、XXX02 等标记作为占位符，这些标记会被 Excel 中对应列的数据替换。

**Q: 多选项功能如何使用？**

A:

1. 在 Excel 中使用 `check` 开头的字段名（如 `check01`）
2. 在单元格中使用逗号、中文逗号或顿号分隔多个选项（如 `游泳,跑步,阅读`）
3. 在 Word 模板中使用 `{{check01}}` 包裹复选框选项（如 `{{check01□ 游泳 □ 跑步 □ 阅读}}`）
4. 系统会自动将选中的选项标记为 `☑`

**Q: 多选项的分隔符支持哪些？**

A: 支持英文逗号 `,`、中文逗号 `，` 和顿号 `、` 作为选项分隔符。

**Q: 复选框没有正确勾选怎么办？**

A: 请检查：

1. Excel 中的选项名称是否与 Word 模板中的选项文本一致
2. 是否使用了正确的分隔符
3. Word 模板中的复选框是否使用 `□` 符号

**Q: 图片替换功能如何使用？**

A:

1. 在 Excel 中使用 `img` 开头的字段名（如 `img01`）
2. 单元格中填写图片文件路径，支持绝对路径（如 `D:\pics\a.jpg`）和相对路径（如 `photos/a.jpg`，相对于 Excel 文件所在目录解析）
3. 在 Word 模板中需要插入图片的位置书写相同名称（如 `img01`），建议占位符单独占用一个段落
4. 生成时占位符会被替换为对应图片；若图片文件不存在或格式不支持，该位置留空并给出警告

**Q: 图片没有被替换（位置留空）怎么办？**

A: 请检查：

1. Excel 单元格中的路径是否真实存在（相对路径以 Excel 文件所在目录为基准）
2. 图片格式是否为 png / jpg / jpeg / gif / bmp
3. Word 模板中的占位符是否与 Excel 表头一致
