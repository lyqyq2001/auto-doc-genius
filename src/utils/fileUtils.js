// 根据 i18n 当前语言选择对应的模板文件，非中文一律按英文处理
import { isUserLanguageChinese, t } from '../i18n';

const triggerDownload = fileName => {
  const link = document.createElement('a');
  link.href = `/${encodeURIComponent(fileName)}`;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

export const downloadExcelTemplate = () => {
  try {
    const fileName = isUserLanguageChinese()
      ? 'Excel Template.xls'
      : 'Excel Template - (English).xls';
    triggerDownload(fileName);
    ElMessage.success(t('message.excelDownloadSuccess'));
  } catch (error) {
    console.error('下载Excel模板出错:', error);
    ElMessage.error(t('message.excelDownloadError'));
  }
};

export const downloadWordTemplate = () => {
  try {
    const fileName = isUserLanguageChinese()
      ? 'Word Template.docx'
      : 'Word Template (English).docx';
    triggerDownload(fileName);
    ElMessage.success(t('message.wordDownloadSuccess'));
  } catch (error) {
    console.error('下载Word模板出错:', error);
    ElMessage.error(t('message.wordDownloadError'));
  }
};

export const beforeExcelUpload = file => {
  const isExcel =
    file.type ===
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    file.type === 'application/vnd.ms-excel' ||
    file.name.endsWith('.xlsx') ||
    file.name.endsWith('.xls');
  if (!isExcel) {
    ElMessage.error(t('upload.excelTip'));
    return false;
  }
  return true;
};

export const beforeWordUpload = file => {
  const isDOCX =
    file.type ===
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
    file.name.endsWith('.docx');
  if (!isDOCX) {
    ElMessage.error(t('upload.wordTip'));
    return false;
  }
  return true;
};
