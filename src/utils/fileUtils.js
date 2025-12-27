export const downloadExcelTemplate = () => {
  try {
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

export const downloadWordTemplate = () => {
  try {
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

export const beforeExcelUpload = file => {
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

export const beforeWordUpload = file => {
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
