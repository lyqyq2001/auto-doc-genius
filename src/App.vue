<template>
  <div class="app-container">
    <h1 class="title">{{ t('title') }}</h1>

    <div class="card-container">
      <!-- Excel上传区域 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>{{ t('upload.excelTitle') }}</span>
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
          <div class="el-upload__text">
            {{ t('upload.dragOrClick')
            }}<em>{{ t('upload.clickToUpload') }}</em>
          </div>
          <template #tip>
            <div class="el-upload__tip">{{ t('upload.excelTip') }}</div>
          </template>
        </el-upload>
      </el-card>

      <!-- Word模板上传区域 -->
      <el-card class="upload-card">
        <template #header>
          <div class="card-header">
            <span>{{ t('upload.wordTitle') }}</span>
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
          <div class="el-upload__text">
            {{ t('upload.dragOrClick')
            }}<em>{{ t('upload.clickToUpload') }}</em>
          </div>
          <template #tip>
            <div class="el-upload__tip">{{ t('upload.wordTip') }}</div>
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
            {{ t('download.excelButton') }}
          </el-button>
          <el-button type="info" @click="downloadWordTemplate" size="medium">
            {{ t('download.wordButton') }}
          </el-button>
        </div>

        <!-- 导出格式选择 -->
        <div class="export-options">
          <el-radio-group v-model="exportFormat" size="large">
            <el-radio-button label="docx">{{
              t('export.word')
            }}</el-radio-button>
            <el-radio-button label="pdf">{{ t('export.pdf') }}</el-radio-button>
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
          {{ t('generate.start') }}
          <span
            v-if="exportFormat === 'pdf' && !officeInstalled"
            style="color: #f56c6c; font-size: 14px"
          >
            {{ t('generate.noOffice') }}</span
          >
        </el-button>

        <!-- 进度条 -->
        <div v-if="generating" class="progress-container">
          <el-progress
            :percentage="progress"
            :status="progress === 100 ? 'success' : ''"
            :stroke-width="2"
          />
          <div class="progress-text">{{ progressText }}</div>
        </div>
      </el-card>
    </div>
  </div>
</template>

<script setup>
import { ref, computed } from 'vue';
import { useI18n } from 'vue-i18n';
import { ElMessage, ElNotification } from 'element-plus';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { UploadFilled } from '@element-plus/icons-vue';
import {
  downloadExcelTemplate,
  downloadWordTemplate,
  beforeExcelUpload,
  beforeWordUpload,
} from './utils/fileUtils';
import { parseExcel, readWordTemplate } from './utils/documentUtils';
import { generateDocx, resolveDuplicateName } from './utils/docxGenerator';
import { isImageColumn, collectImagePaths } from './utils/imageUtils';

const { t } = useI18n();

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

// 预读取 Excel 中图片列引用的所有图片，返回 { imageAssets, excelDir }
const preloadImages = async excelData => {
  const headers = excelData.length > 0 ? Object.keys(excelData[0]) : [];
  if (!headers.some(isImageColumn)) {
    return { imageAssets: null, excelDir: '' };
  }

  if (
    !window.electronAPI ||
    !window.electronAPI.getFilePath ||
    !window.electronAPI.readImageFiles
  ) {
    ElMessage.warning(t('image.noElectron'));
    return { imageAssets: null, excelDir: '' };
  }

  let excelDir = '';
  try {
    const excelPath = window.electronAPI.getFilePath(excelFile.value);
    if (excelPath) {
      excelDir = excelPath.replace(/[\\/]+[^\\/]*$/, '');
    }
  } catch (error) {
    console.error('获取Excel路径失败:', error);
  }

  const paths = collectImagePaths(excelData, excelDir);
  if (paths.length === 0) {
    return { imageAssets: {}, excelDir };
  }

  progressText.value = t('progress.loadingImages');
  const imageAssets = (await window.electronAPI.readImageFiles(paths)) || {};

  const failedPaths = paths.filter(p => !imageAssets[p] || !imageAssets[p].ok);
  if (failedPaths.length > 0) {
    console.warn('以下图片读取失败:', failedPaths);
    ElNotification.warning({
      title: t('image.missingTitle'),
      message: t('image.missingMessage', { count: failedPaths.length }),
      duration: 5000,
    });
  }

  return { imageAssets, excelDir };
};

// 开始生成
const startGenerate = async () => {
  if (!excelFile.value || !wordFile.value) {
    ElMessage.warning(t('generate.uploadFirst'));
    return;
  }

  generating.value = true;
  progress.value = 0;
  progressText.value = t('progress.preparing');

  const startTime = Date.now();

  let progressListener = null;

  try {
    // [{XXX01:'xxx',XXX02:'xxx',......} , ....]
    const excelData = await parseExcel(excelFile.value);
    // arraybuffer
    const wordTemplate = await readWordTemplate(wordFile.value);

    // 图片替换：预读取所有被引用的图片
    const { imageAssets, excelDir } = await preloadImages(excelData);

    const zip = new JSZip();

    // 输出文件名去重：重名文件自动追加序号，避免在ZIP中互相覆盖
    const usedFileNames = new Set();
    let renamedCount = 0;

    // 如果是 PDF 模式，使用流式处理
    if (exportFormat.value === 'pdf') {
      progressText.value = t('progress.generatingAndConverting');

      // 监听PDF转换进度
      progressListener = window.electronAPI.onPdfProgress(data => {
        if (data.stage === 'converting') {
          progress.value = Math.round(50 + (data.progress || 0) * 0.4);
          progressText.value = t('progress.convertingPdf');
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
          progressText.value = t('progress.generatingWord', {
            current: globalIndex + 1,
            total: excelData.length,
          });

          const docx = generateDocx(
            wordTemplate,
            rowData,
            globalIndex,
            imageAssets,
            excelDir,
          );
          const uniqueName = resolveDuplicateName(docx.name, usedFileNames);
          if (uniqueName !== docx.name) renamedCount += 1;
          batchDocxList.push({ ...docx, name: uniqueName });
        }

        progressText.value = t('progress.convertingPdfBatch', {
          start: i + 1,
          end: Math.min(i + BATCH_SIZE, excelData.length),
          total: excelData.length,
        });

        const result =
          await window.electronAPI.convertBatchToPdf(batchDocxList);

        if (result?.success) {
          pdfResults.push(...result.results);
          progress.value = Math.round(
            ((i + BATCH_SIZE) / excelData.length) * 90,
          );
        } else {
          const errorKey =
            result?.errorCode === 'officeUnavailable'
              ? 'generate.officeUnavailable'
              : 'generate.pdfConversionFailed';
          throw new Error(t(errorKey));
        }
      }

      progressText.value = t('progress.packagingPdf');

      pdfResults.forEach(pdfFile => {
        zip.file(pdfFile.name, pdfFile.data);
      });
    } else {
      // Word 模式：直接生成
      for (let i = 0; i < excelData.length; i++) {
        progress.value = Math.round((i / excelData.length) * 90);
        progressText.value = t('progress.processingRow', {
          current: i + 1,
          total: excelData.length,
        });

        const rowData = excelData[i];
        const docx = generateDocx(
          wordTemplate,
          rowData,
          i,
          imageAssets,
          excelDir,
        );

        const uniqueName = resolveDuplicateName(docx.name, usedFileNames);
        if (uniqueName !== docx.name) renamedCount += 1;
        zip.file(uniqueName, docx.buffer);
      }
    }

    progressText.value = t('progress.packaging');

    // 生成ZIP文件
    const zipBlob = await zip.generateAsync({ type: 'blob' });

    // 生成ZIP文件名
    const zipFileName =
      exportFormat.value === 'docx'
        ? t('result.wordZipName')
        : t('result.pdfZipName');

    // 下载ZIP文件
    saveAs(zipBlob, zipFileName);

    if (renamedCount > 0) {
      ElNotification.warning({
        title: t('result.duplicateTitle'),
        message: t('result.duplicateFiles', { count: renamedCount }),
        duration: 5000,
      });
    }

    const endTime = Date.now();
    const elapsedSeconds = Math.round((endTime - startTime) / 1000);

    ElNotification.success({
      title: t('result.successTitle'),
      message: t('result.successMessage', {
        count: excelData.length,
        format: exportFormat.value === 'docx' ? 'Word' : 'PDF',
        seconds: elapsedSeconds,
      }),
      duration: 3000,
    });
  } catch (error) {
    console.error('生成失败:', error);
    const knownErrorKeys = [
      'error.excelFormat',
      'error.excelNoData',
      'error.excelParseFailed',
      'error.fileReadFailed',
      'error.wordReadFailed',
      'generate.officeUnavailable',
      'generate.pdfConversionFailed',
    ];
    const knownMessage = knownErrorKeys
      .map(key => t(key))
      .find(message => message === error?.message);
    ElMessage.error(
      t('generate.failed', {
        message: knownMessage || t('generate.unexpectedFailed'),
      }),
    );
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
