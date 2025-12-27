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
  import {
    downloadExcelTemplate,
    downloadWordTemplate,
    beforeExcelUpload,
    beforeWordUpload,
  } from './utils/fileUtils';
  import {
    parseExcel,
    processCheckboxOptions,
    readWordTemplate,
  } from './utils/documentUtils';

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

  checkOfficeInstallation();

  const handleExcelChange = file => {
    excelFile.value = file.raw;
  };

  const handleWordChange = file => {
    wordFile.value = file.raw;
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
