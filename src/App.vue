<!-- Excel转Word批量生成器 -->
<!-- 所需依赖：npm install xlsx docxtemplater pizzip jszip file-saver element-plus -->

<template>
  <div class="app-container">
    <h1 class="title">Excel转Word批量生成器</h1>

    <div class="card-container">
      <!-- Excel上传区域 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>上传Excel数据表</span>
          </div>
        </template>

        <el-upload
          v-model:file-list="excelFileList"
          class="upload-demo"
          drag
          action="#"
          :auto-upload="false"
          :on-change="handleExcelChange"
          :before-upload="beforeExcelUpload"
          accept=".xlsx,.xls"
          :show-file-list="true"
        >
          <el-icon class="el-icon--upload"><upload-filled /></el-icon>
          <div class="el-upload__text">拖拽文件到此处或<em>点击上传</em></div>
          <template #tip>
            <div class="el-upload__tip">请上传.xlsx或.xls格式的Excel文件</div>
          </template>
        </el-upload>
      </el-card>

      <!-- Word模板上传区域 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>上传Word模板</span>
          </div>
        </template>

        <el-upload
          v-model:file-list="wordFileList"
          class="upload-demo"
          drag
          action="#"
          :auto-upload="false"
          :on-change="handleWordChange"
          :before-upload="beforeWordUpload"
          accept=".docx"
          :show-file-list="true"
        >
          <el-icon class="el-icon--upload"><upload-filled /></el-icon>
          <div class="el-upload__text">拖拽文件到此处或<em>点击上传</em></div>
          <template #tip>
            <div class="el-upload__tip">请上传.docx格式的Word模板文件</div>
          </template>
        </el-upload>
      </el-card>

      <!-- 操作区域 -->
      <el-card class="operation-card">
        <el-button
          type="primary"
          :loading="generating"
          :disabled="!excelFile || !wordFile || generating"
          @click="startGenerate"
          size="large"
        >
          {{ generating ? '生成中...' : '开始生成Word文件' }}
        </el-button>

        <!-- 进度条 -->
        <div v-if="generating" class="progress-container">
          <el-progress
            :percentage="progress"
            :status="progress === 100 ? 'success' : 'active'"
            stroke-width="2"
          />
          <div class="progress-text">{{ progressText }}</div>
        </div>
      </el-card>
    </div>
  </div>
</template>

<script setup>
  import { ref, reactive } from 'vue';
  import { ElMessage, ElNotification } from 'element-plus';
  import * as XLSX from 'xlsx';
  import PizZip from 'pizzip';
  import Docxtemplater from 'docxtemplater';
  import JSZip from 'jszip';
  import { saveAs } from 'file-saver';
  import { UploadFilled } from '@element-plus/icons-vue';

  // 文件状态
  const excelFile = ref(null);
  const wordFile = ref(null);
  const excelFileList = ref([]);
  const wordFileList = ref([]);
  const generating = ref(false);
  const progress = ref(0);
  const progressText = ref('');
  // 导出格式，只保留docx
  const exportFormat = ref('docx');

  // Excel文件上传前校验
  const beforeExcelUpload = (file) => {
    const isExcel =
      file.type ===
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.type === 'application/vnd.ms-excel' ||
      file.name.endsWith('.xlsx') ||
      file.name.endsWith('.xls');
    if (!isExcel) {
      ElMessage.error('请上传.xlsx或.xls格式的Excel文件');
      return false;
    }
    return true;
  };

  // Word文件上传前校验
  const beforeWordUpload = (file) => {
    const isDOCX =
      file.type ===
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      file.name.endsWith('.docx');
    if (!isDOCX) {
      ElMessage.error('请上传.docx格式的Word模板文件');
      return false;
    }
    return true;
  };

  // Excel文件上传处理
  const handleExcelChange = (file, fileList) => {
    excelFile.value = file.raw;
    excelFileList.value = fileList;
  };

  // Word文件上传处理
  const handleWordChange = (file, fileList) => {
    wordFile.value = file.raw;
    wordFileList.value = fileList;
  };

  // 测试Word文件上传处理
  const handleTestDocxChange = (file, fileList) => {
    testDocxFile.value = file.raw;
    testDocxFileList.value = fileList;
  };

  // LibreOffice转换测试
  const testLibreOfficeConvert = async () => {
    if (!testDocxFile.value) {
      ElMessage.warning('请先上传Word文件');
      return;
    }

    converting.value = true;

    try {
      if (window.electronAPI) {
        // 直接使用文件路径进行转换（这里简化处理，实际应用中需要先保存文件）
        // 注意：实际应用中需要先将上传的文件保存到本地，然后再调用转换API
        const result = await window.electronAPI.convertDocxToPdf(
          testDocxFile.value.path || testDocxFile.value.name
        );

        ElNotification.success({
          title: '成功',
          message: `PDF转换成功！文件路径：${result.pdfPath}`,
          duration: 5000,
        });
      } else {
        ElMessage.error('Electron API未加载，请在Electron环境中运行');
      }
    } catch (error) {
      console.error('转换失败:', error);
      ElMessage.error(`转换失败: ${error.message}`);
    } finally {
      converting.value = false;
    }
  };

  // 解析Excel文件
  const parseExcel = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];

          // 解析为JSON格式，header: 1 表示使用数组索引作为默认表头
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          // 第一行是描述，第二行是表头，所以至少需要3行数据
          if (jsonData.length < 3) {
            reject(
              new Error(
                'Excel表格格式不正确，需要包含描述行、表头行和至少一行数据'
              )
            );
            return;
          }

          // 获取表头（第二行）和数据行（第三行及以后）
          const headers = jsonData[1];
          // 打印提取的表头，用于调试
          console.log('提取的Excel表头:', headers);

          // 跳过第一行（描述）和第二行（表头），从第三行开始处理数据
          const rows = jsonData
            .slice(2)
            .filter((row) =>
              row.some(
                (cell) => cell !== undefined && cell !== null && cell !== ''
              )
            );

          if (rows.length === 0) {
            reject(new Error('Excel表格中未检测到有效数据行'));
            return;
          }

          // 将每行数据转换为对象，键为表头
          const formattedRows = rows.map((row) => {
            const obj = {};
            headers.forEach((header, index) => {
              let value = row[index] || '';

              // 修复Excel日期和时间格式问题
              if (typeof value === 'number') {
                // 检查是否为日期数字（大于25569表示1970年以后的日期）
                if (value > 25569) {
                  // 转换为JavaScript日期
                  const date = new Date((value - 25569) * 86400000);
                  // 格式化日期为YYYY-MM-DD或YYYY年MM月DD日
                  value = `${date.getFullYear()}年${(date.getMonth() + 1)
                    .toString()
                    .padStart(2, '0')}月${date
                    .getDate()
                    .toString()
                    .padStart(2, '0')}日`;
                }
                // 检查是否为时间数字（小于1表示时间）
                else if (value > 0 && value < 1) {
                  // 转换为小时数（1天=24小时）
                  const hours = Math.floor(value * 24);
                  // 转换为分钟数
                  const minutes = Math.floor((value * 24 * 60) % 60);
                  // 格式化时间为HH:MM
                  value = `${hours.toString().padStart(2, '0')}:${minutes
                    .toString()
                    .padStart(2, '0')}`;
                }
              }

              obj[header] = value;

              // 添加别名处理，解决XXX01和XXX11的匹配问题
              // 如果表头是XXX01，同时添加XXX11作为别名
              if (header === 'XXX01') {
                console.log('为XXX01添加别名XXX11，值为:', value);
                obj['XXX11'] = value;
              }
              // 如果表头是XXX11，同时添加XXX01作为别名
              else if (header === 'XXX11') {
                console.log('为XXX11添加别名XXX01，值为:', value);
                obj['XXX01'] = value;
              }
            });

            // 额外添加所有可能的XXX01-XXX11别名，确保所有情况都能匹配
            // 检查是否有XXX01或XXX11字段，互相添加别名
            if (obj['XXX01'] && !obj['XXX11']) {
              obj['XXX11'] = obj['XXX01'];
            }
            if (obj['XXX11'] && !obj['XXX01']) {
              obj['XXX01'] = obj['XXX11'];
            }

            return obj;
          });

          resolve(formattedRows);
        } catch (error) {
          reject(new Error('Excel文件解析失败'));
        }
      };
      reader.onerror = () => reject(new Error('文件读取失败'));
      reader.readAsArrayBuffer(file);
    });
  };

  // 读取Word模板文件
  const readWordTemplate = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const content = e.target.result;
          resolve(content);
        } catch (error) {
          reject(new Error('Word模板读取失败'));
        }
      };
      reader.onerror = () => reject(new Error('文件读取失败'));
      reader.readAsArrayBuffer(file);
    });
  };

  // 开始生成
  const startGenerate = async () => {
    if (!excelFile.value || !wordFile.value) {
      ElMessage.warning('请先上传Excel文件和Word模板文件');
      return;
    }

    generating.value = true;
    progress.value = 0;
    progressText.value = '准备中...';

    try {
      // 解析Excel数据
      const excelData = await parseExcel(excelFile.value);
      const totalRows = excelData.length;

      // 读取Word模板
      const wordTemplate = await readWordTemplate(wordFile.value);

      // 创建JSZip实例
      const zip = new JSZip();

      // 遍历Excel数据行
      for (let i = 0; i < totalRows; i++) {
        progress.value = Math.round((i / totalRows) * 100);
        progressText.value = `正在处理第 ${i + 1} 行，共 ${totalRows} 行`;

        const rowData = excelData[i];
        console.log(`第${i + 1}行数据:`, rowData);

        // 使用pizzip加载Word模板
        const templateZip = new PizZip(wordTemplate);

        // 遍历所有XML文件，将XXX01格式替换为{XXX01}格式
        Object.keys(templateZip.files).forEach((filename) => {
          if (filename.endsWith('.xml')) {
            const fileContent = templateZip.file(filename).asText();
            let updatedContent = fileContent;

            // 检查文件内容中是否包含任何XXX字段
            const hasXXX = fileContent.includes('XXX');
            console.log(`文件 ${filename} 中包含XXX: ${hasXXX}`);

            if (hasXXX) {
              // 打印所有包含XXX的行，用于调试
              const lines = fileContent.split('\n');
              lines.forEach((line, index) => {
                if (line.includes('XXX')) {
                  console.log(
                    `文件 ${filename} 第${index + 1}行: ${line.trim()}`
                  );
                }
              });
            }

            // 遍历所有数据键，替换模板中的占位符
            Object.keys(rowData).forEach((key) => {
              // 直接替换所有出现的key，使用全局替换
              const regex = new RegExp(key, 'g');
              const updated = updatedContent.replace(regex, `{${key}}`);

              // 计算替换数量
              const originalMatches = updatedContent.match(regex) || [];
              const newMatches =
                updated.match(new RegExp(`\{${key}\}`, 'g')) || [];

              if (originalMatches.length > 0) {
                console.log(
                  `文件 ${filename} 中: ${originalMatches.length}个${key} → ${newMatches.length}个{${key}}`
                );
              }

              updatedContent = updated;
            });

            // 更新文件内容
            templateZip.file(filename, updatedContent);
          }
        });

        // 查看当前模板中包含的占位符
        console.log(
          '模板中的XML文件列表:',
          Object.keys(templateZip.files).filter((f) => f.endsWith('.xml'))
        );

        // 创建docxtemplater实例，配置支持{字段名}格式的占位符
        const doc = new Docxtemplater(templateZip, {
          paragraphLoop: true,
          linebreaks: true,
          // 确保分隔符配置正确
          delimiters: {
            start: '{',
            end: '}',
          },
        });

        // 查看要填充的数据
        console.log('要填充的数据:', rowData);

        // 检查数据中是否包含XXX01字段
        if (rowData.XXX01) {
          console.log('数据中包含XXX01字段，值为:', rowData.XXX01);
        } else {
          console.log('数据中不包含XXX01字段');
          // 检查是否有大小写问题或其他命名问题
          console.log('数据中的所有字段:', Object.keys(rowData));
        }

        // 填充数据
        console.log('最终要填充的数据:', rowData);

        // 确保XXX01字段被正确处理
        if (rowData.XXX01) {
          console.log('正在填充XXX01字段，值为:', rowData.XXX01);
        }

        doc.render(rowData);

        // 生成Word文件
        const docxBlob = doc.getZip().generate({
          type: 'blob',
          mimeType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        // 生成文件名，使用第一列数据或默认名称
        const firstKey = Object.keys(rowData)[0];
        const baseFileName = rowData[firstKey] || `doc_${i + 1}`;

        // 直接添加Word文件到ZIP中
        zip.file(`${baseFileName}.docx`, docxBlob);
      }

      // 完成进度
      progress.value = 100;
      progressText.value = '正在打包文件...';

      // 生成ZIP文件
      const zipBlob = await zip.generateAsync({ type: 'blob' });

      // 生成ZIP文件名
      const zipFileName = '生成的Word文件.zip';

      // 下载ZIP文件
      saveAs(zipBlob, zipFileName);

      ElNotification.success({
        title: '成功',
        message: `已生成 ${totalRows} 个Word文件，已打包下载`,
        duration: 3000,
      });
    } catch (error) {
      console.error('生成失败:', error);
      ElMessage.error(`生成失败: ${error.message}`);
    } finally {
      generating.value = false;
      progress.value = 0;
      progressText.value = '';
    }
  };
</script>

<style scoped>
  .app-container {
    max-width: 1000px;
    margin: 0 auto;
    padding: 20px;
    font-family: Arial, sans-serif;
    background-color: #f5f7fa;
    min-height: 100vh;
  }

  .title {
    text-align: center;
    color: #333;
    margin-bottom: 30px;
    font-size: 28px;
    font-weight: 600;
  }

  .card-container {
    display: flex;
    flex-direction: column;
    gap: 20px;
  }

  .upload-card {
    box-shadow: 0 2px 12px 0 rgba(0, 0, 0, 0.1);
  }

  .card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-weight: bold;
    color: #333;
  }

  .operation-card {
    display: flex;
    flex-direction: column;
    gap: 20px;
    align-items: center;
    padding: 30px;
    box-shadow: 0 2px 12px 0 rgba(0, 0, 0, 0.1);
  }

  .export-options {
    margin-bottom: 10px;
  }

  .export-options .el-radio-group {
    display: flex;
    gap: 10px;
  }

  .progress-container {
    width: 100%;
    max-width: 500px;
  }

  .progress-text {
    text-align: center;
    margin-top: 10px;
    color: #606266;
    font-size: 14px;
  }

  :deep(.el-upload-dragger) {
    width: 100%;
    height: 200px;
  }

  :deep(.el-upload__tip) {
    text-align: center;
  }
</style>
