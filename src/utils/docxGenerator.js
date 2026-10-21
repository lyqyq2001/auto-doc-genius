import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { processCheckboxOptions } from './documentUtils';
import { injectImages, isImageColumn } from './imageUtils';

export const generateDocx = (
  wordTemplate,
  rowData,
  index,
  imageAssets = null,
  excelDir = '',
) => {
  const templateZip = new PizZip(wordTemplate);

  // 图片占位符注入（需在复选框处理和文本占位符包裹之前执行）
  if (imageAssets) {
    injectImages(templateZip, rowData, imageAssets, excelDir);
  }

  Object.keys(templateZip.files).forEach(filename => {
    if (filename.endsWith('.xml')) {
      const fileContent = templateZip.file(filename).asText();
      let updatedContent = fileContent;

      updatedContent = processCheckboxOptions(updatedContent, rowData);

      Object.keys(rowData).forEach(key => {
        if (key.startsWith('check')) return;
        if (isImageColumn(key)) return;
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
  const baseFileName = rowData[firstKey] || `doc_${index + 1}`;

  return {
    buffer: docxArrayBuffer,
    name: `${baseFileName}.docx`,
  };
};
