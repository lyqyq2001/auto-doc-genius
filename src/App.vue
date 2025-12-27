<template>
  <div class="app-container">
    <h1 class="title">Excel转报告批量生成器</h1>

    <div class="card-container">
      <!-- Excel上传区域 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>上传Excel数据表</span>
          </div>
        </template>

        <el-upload
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
        <!-- 模板下载区域 -->
        <div class="template-download">
          <el-button
            type="info"
            @click="downloadExcelTemplate"
            size="medium"
            style="margin-right: 10px"
          >
            下载excel使用示例模板（可修改）
          </el-button>
          <el-button type="info" @click="downloadWordTemplate" size="medium">
            下载word使用示例模板（可修改）
          </el-button>
        </div>

        <!-- 导出格式选择 -->
        <div class="export-options">
          <el-radio-group v-model="exportFormat" size="large">
            <el-radio-button label="docx">导出Word</el-radio-button>
            <el-radio-button label="pdf">导出PDF</el-radio-button>
          </el-radio-group>
        </div>

        <el-button
          type="primary"
          style="width: 100%"
          :loading="generating"
          :disabled="disabled"
          @click="startGenerate"
          size="large"
        >
          开始生成报告
          <span
            v-if="exportFormat === 'pdf' && !officeInstalled"
            style="color: #f56c6c; font-size: 14px"
          >
            (未安装Office，无法生成PDF)</span
          >
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
  import { ref, computed } from 'vue';
  import { ElMessage, ElNotification } from 'element-plus';
  import * as XLSX from 'xlsx';
  import { saveAs } from 'file-saver';
  import PizZip from 'pizzip';
  import Docxtemplater from 'docxtemplater';
  import JSZip from 'jszip';
  import { UploadFilled } from '@element-plus/icons-vue';

  // 文件状态
  const excelFile = ref(null);
  const wordFile = ref(null);
  const generating = ref(false);
  const progress = ref(0);
  const progressText = ref('');
  // 导出格式，默认docx
  const exportFormat = ref('docx');
  // Office是否安装
  const officeInstalled = ref(false);
  const disabled = computed(() => {
    return (
      !excelFile.value ||
      !wordFile.value ||
      generating.value ||
      (exportFormat.value === 'pdf' && !officeInstalled.value)
    );
  });
  // 下载Excel模板
  const downloadExcelTemplate = () => {
    try {
      // 直接从public目录下载模板文件
      const link = document.createElement('a');
      link.href = '/Excel Template.xls';
      link.download = 'Excel Template.xls';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      ElMessage.success('Excel模板下载成功');
    } catch (error) {
      console.error('下载Excel模板出错:', error);
      ElMessage.error('Excel模板下载出错');
    }
  };

  // 下载Word模板
  const downloadWordTemplate = () => {
    try {
      // 直接从public目录下载模板文件
      const link = document.createElement('a');
      link.href = '/Word Template.docx';
      link.download = 'Word Template.docx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      ElMessage.success('Word模板下载成功');
    } catch (error) {
      console.error('下载Word模板出错:', error);
      ElMessage.error('Word模板下载出错');
    }
  };

  // 检查Office是否安装
  const checkOfficeInstallation = async () => {
    try {
      if (window.electronAPI && window.electronAPI.checkOfficeInstallation) {
        officeInstalled.value =
          await window.electronAPI.checkOfficeInstallation();
      }
    } catch (error) {
      officeInstalled.value = false;
    }
  };

  // 组件挂载时检查Office安装状态
  checkOfficeInstallation();

  // Excel文件上传前校验
  const beforeExcelUpload = file => {
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
  const beforeWordUpload = file => {
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
  const handleExcelChange = file => {
    excelFile.value = file.raw;
  };

  // Word文件上传处理
  const handleWordChange = file => {
    wordFile.value = file.raw;
  };

  // 解析Excel文件
  const parseExcel = file => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          // 转换为JSON数据，header:1表示返回一个二维数组，jsonData中每个元素都是一个数组，数组中的每个元素都是单元格的值
          // 默认情况下，即无header参数时，返回的是一个对象数组，每个对象中的属性名是第一行单元格的内容，属性值是其余行单元格的值
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          if (jsonData.length < 3) {
            reject(
              new Error(
                'Excel表格格式不正确，需要包含描述行、表头行和至少一行数据'
              )
            );
            return;
          }

          const headers = jsonData[1];
          const rows = jsonData
            .slice(2)
            .filter(row =>
              row.some(
                cell => cell !== undefined && cell !== null && cell !== ''
              )
            );

          if (rows.length === 0) {
            reject(new Error('Excel表格中未检测到有效数据行'));
            return;
          }

          const formattedRows = rows.map(row => {
            const obj = {};
            headers.forEach((header, index) => {
              let value = row[index] || '';

              if (typeof value === 'number') {
                if (value > 25569) {
                  const date = new Date((value - 25569) * 86400000);
                  value = `${date.getFullYear()}年${(date.getMonth() + 1)
                    .toString()
                    .padStart(2, '0')}月${date
                    .getDate()
                    .toString()
                    .padStart(2, '0')}日`;
                } else if (value > 0 && value < 1) {
                  const hours = Math.floor(value * 24);
                  const minutes = Math.floor((value * 24 * 60) % 60);
                  value = `${hours.toString().padStart(2, '0')}:${minutes
                    .toString()
                    .padStart(2, '0')}`;
                }
              }

              obj[header] = value;
            });

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

  // 处理多选项
  const processCheckboxOptions = (content, rowData) => {
    let updatedContent = content;

    const checkboxPattern = /\{\{(check\d+)(.*?)\}\}/g;

    updatedContent = updatedContent.replace(
      checkboxPattern,
      (_, key, optionsPart) => {
        const excelValue = rowData[key];

        const selectedValues = excelValue
          ? excelValue
              .toString()
              .split(/[，,、]/)
              .map(v => v.trim())
          : [];

        let result = optionsPart;

        const textTagPattern = /<w:t[^>]*>([^<]*)<\/w:t>/g;
        let matchResult;
        const allMatches = [];

        while ((matchResult = textTagPattern.exec(optionsPart)) !== null) {
          allMatches.push({
            fullMatch: matchResult[0],
            text: matchResult[1],
            index: matchResult.index,
          });
        }

        for (let i = 0; i < allMatches.length; i++) {
          const current = allMatches[i];

          if (current.text.trim() === '□' && i + 1 < allMatches.length) {
            const next = allMatches[i + 1];
            const nextText = next.text.trim();

            const isSelected = selectedValues.some(
              selectedValue =>
                nextText.includes(selectedValue) ||
                selectedValue.includes(nextText)
            );

            if (isSelected) {
              const newFullMatch = current.fullMatch.replace('□', '☑');
              result =
                result.substring(0, current.index) +
                newFullMatch +
                result.substring(current.index + current.fullMatch.length);
            }
          }
        }

        return result;
      }
    );

    return updatedContent;
  };

  // 读取Word模板文件
  const readWordTemplate = file => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => {
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

    const startTime = Date.now();

    let progressListener = null;

    try {
      // [{XXX01:'xxx',XXX02:'xxx',......} , ....]
      const excelData = await parseExcel(excelFile.value);
      // arraybuffer
      const wordTemplate = await readWordTemplate(wordFile.value);
      const zip = new JSZip();

      // 如果是 PDF 模式，使用流式处理
      if (exportFormat.value === 'pdf') {
        progressText.value = '正在生成Word文档并转换PDF...';

        // 监听PDF转换进度
        progressListener = window.electronAPI.onPdfProgress(data => {
          if (data.stage === 'converting') {
            progress.value = Math.round(50 + (data.progress || 0) * 0.4);
            progressText.value = data.message || '正在转换PDF...';
          }
        });

        const BATCH_SIZE = 10;
        const pdfResults = [];

        for (let i = 0; i < excelData.length; i += BATCH_SIZE) {
          const batch = excelData.slice(i, i + BATCH_SIZE);
          const batchDocxList = [];

          for (let j = 0; j < batch.length; j++) {
            const rowData = batch[j];
            const globalIndex = i + j;

            progress.value = Math.round((globalIndex / excelData.length) * 50);
            progressText.value = `正在生成Word文档 (${globalIndex + 1}/${
              excelData.length
            })...`;

            const templateZip = new PizZip(wordTemplate);
            Object.keys(templateZip.files).forEach(filename => {
              if (filename.endsWith('.xml')) {
                const fileContent = templateZip.file(filename).asText();
                let updatedContent = fileContent;

                updatedContent = processCheckboxOptions(
                  updatedContent,
                  rowData
                );

                Object.keys(rowData).forEach(key => {
                  if (key.startsWith('check')) return;
                  const regex = new RegExp(key, 'g');
                  if (!updatedContent.includes(`{${key}}`)) {
                    updatedContent = updatedContent.replace(regex, `{${key}}`);
                  }
                });

                templateZip.file(filename, updatedContent);
              }
            });
            const doc = new Docxtemplater(templateZip, {
              paragraphLoop: true,
              linebreaks: true,
            });

            doc.render(rowData);

            const docxArrayBuffer = doc.getZip().generate({
              type: 'arraybuffer',
              mimeType:
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            });

            const firstKey = Object.keys(rowData)[0];
            const baseFileName = rowData[firstKey] || `doc_${globalIndex + 1}`;

            batchDocxList.push({
              name: `${baseFileName}.docx`,
              buffer: docxArrayBuffer,
            });
          }

          progressText.value = `正在转换PDF (${i + 1}-${Math.min(
            i + BATCH_SIZE,
            excelData.length
          )}/${excelData.length})...`;

          const result = await window.electronAPI.convertBatchToPdf(
            batchDocxList
          );

          if (result.success) {
            pdfResults.push(...result.results);
            progress.value = Math.round(
              ((i + BATCH_SIZE) / excelData.length) * 90
            );
          } else {
            throw new Error(result.error);
          }
        }

        progressText.value = '正在打包PDF文件...';

        pdfResults.forEach(pdfFile => {
          zip.file(pdfFile.name, pdfFile.data);
        });
      } else {
        // Word 模式：直接生成
        for (let i = 0; i < excelData.length; i++) {
          progress.value = Math.round((i / excelData.length) * 90);
          progressText.value = `正在处理第 ${i + 1} 行，共 ${
            excelData.length
          } 行`;

          const rowData = excelData[i];

          const templateZip = new PizZip(wordTemplate);
          Object.keys(templateZip.files).forEach(filename => {
            if (filename.endsWith('.xml')) {
              const fileContent = templateZip.file(filename).asText();
              let updatedContent = fileContent;

              updatedContent = processCheckboxOptions(updatedContent, rowData);

              Object.keys(rowData).forEach(key => {
                if (key.startsWith('check')) return;
                const regex = new RegExp(key, 'g');
                if (!updatedContent.includes(`{${key}}`)) {
                  updatedContent = updatedContent.replace(regex, `{${key}}`);
                }
              });

              templateZip.file(filename, updatedContent);
            }
          });

          const doc = new Docxtemplater(templateZip, {
            paragraphLoop: true,
            linebreaks: true,
          });

          doc.render(rowData);
          const docxArrayBuffer = doc.getZip().generate({
            type: 'arraybuffer',
            mimeType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          });
          const firstKey = Object.keys(rowData)[0];
          const baseFileName = rowData[firstKey] || `doc_${i + 1}`;

          zip.file(`${baseFileName}.docx`, docxArrayBuffer);
        }
      }

      progressText.value = '正在打包文件...';

      // 生成ZIP文件
      const zipBlob = await zip.generateAsync({ type: 'blob' });

      // 生成ZIP文件名
      const zipFileName =
        exportFormat.value === 'docx'
          ? '生成的Word文件.zip'
          : '生成的PDF文件.zip';

      // 下载ZIP文件
      saveAs(zipBlob, zipFileName);

      const endTime = Date.now();
      const elapsedSeconds = Math.round((endTime - startTime) / 1000);

      ElNotification.success({
        title: '成功',
        message: `已生成 ${excelData.length} 个${
          exportFormat.value === 'docx' ? 'Word' : 'PDF'
        }文件，已打包下载。耗时: ${elapsedSeconds}秒`,
        duration: 3000,
      });
    } catch (error) {
      console.error('生成失败:', error);
      ElMessage.error(`生成失败: ${error.message}`);
    } finally {
      if (progressListener) {
        progressListener();
      }
      generating.value = false;
      progress.value = 0;
      progressText.value = '';
    }
  };
</script>

<style>
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

  .template-download {
    display: flex;
    gap: 10px;
    margin-bottom: 10px;
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
