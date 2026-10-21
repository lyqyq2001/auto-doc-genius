import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { processCheckboxOptions } from './documentUtils';
import { injectImages, isImageColumn } from './imageUtils';

// 转义正则特殊字符，避免 Excel 表头被当作正则解析
const escapeRegExp = str => str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// run 之间允许忽略的间隔内容（空白、拼写检查标记）
const IGNORABLE_GAP = /^(?:\s|<w:proofErr\b[^>]*\/>)*$/;
// 简单 run：仅包含可选 rPr 和若干 w:t（含图片、换行等其他内容的 run 不参与合并）
const SIMPLE_RUN_PATTERN =
  /<w:r(?:\s[^>]*)?>((?:<w:rPr>[\s\S]*?<\/w:rPr>)?)((?:<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>)+)<\/w:r>/g;
const WT_PATTERN = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;

// 提取 run 的纯文本内容
const extractRunText = textXml => {
  let text = '';
  let match;
  WT_PATTERN.lastIndex = 0;
  while ((match = WT_PATTERN.exec(textXml)) !== null) {
    text += match[1];
  }
  return text;
};

// 合并相邻的同格式简单 run。
// Word 经常把一个占位符拆分到多个 run 中（编辑历史、拼写标记等），
// 导致基于文本匹配的占位符替换失败；合并后即可正常识别。
// 注意：文本含 {、}、□ 的 run 不参与合并，避免破坏复选框占位符的结构。
export const mergeAdjacentRuns = xml => {
  const runs = [];
  let match;
  SIMPLE_RUN_PATTERN.lastIndex = 0;

  while ((match = SIMPLE_RUN_PATTERN.exec(xml)) !== null) {
    runs.push({
      start: match.index,
      end: match.index + match[0].length,
      rPr: match[1] || '',
      textXml: match[2],
      text: extractRunText(match[2]),
    });
  }

  if (runs.length < 2) return xml;

  const mergeable = run => !/[{}□]/.test(run.text);

  let result = '';
  let cursor = 0;
  let i = 0;

  while (i < runs.length) {
    const first = runs[i];

    if (!mergeable(first)) {
      result += xml.slice(cursor, first.end);
      cursor = first.end;
      i += 1;
      continue;
    }

    // 向后收集可合并的连续 run
    const group = [first];
    let j = i + 1;
    while (j < runs.length) {
      const prev = runs[j - 1];
      const cur = runs[j];
      const gap = xml.slice(prev.end, cur.start);

      if (IGNORABLE_GAP.test(gap) && cur.rPr === first.rPr && mergeable(cur)) {
        group.push(cur);
        j += 1;
      } else {
        break;
      }
    }

    result += xml.slice(cursor, first.start);
    if (group.length > 1) {
      const mergedText = group.map(run => run.text).join('');
      result += `<w:r>${first.rPr}<w:t xml:space="preserve">${mergedText}</w:t></w:r>`;
    } else {
      result += xml.slice(first.start, first.end);
    }

    cursor = group[group.length - 1].end;
    i = j;
  }

  result += xml.slice(cursor);
  return result;
};

// 清洗文件名中的非法字符（Windows 文件名限制）
export const sanitizeFileName = name =>
  String(name ?? '')
    // eslint-disable-next-line no-control-regex
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .trim();

// 重名时追加 _2、_3 等序号，返回不重复的名称并登记到 usedNames
export const resolveDuplicateName = (name, usedNames) => {
  const dotIndex = name.lastIndexOf('.');
  const base = dotIndex === -1 ? name : name.slice(0, dotIndex);
  const ext = dotIndex === -1 ? '' : name.slice(dotIndex);

  let uniqueName = name;
  let suffix = 2;
  while (usedNames.has(uniqueName)) {
    uniqueName = `${base}_${suffix}${ext}`;
    suffix += 1;
  }

  usedNames.add(uniqueName);
  return uniqueName;
};

export const generateDocx = (
  wordTemplate,
  rowData,
  index,
  imageAssets = null,
  excelDir = '',
) => {
  const templateZip = new PizZip(wordTemplate);

  // 合并相邻同格式 run，防止 Word 拆分占位符导致替换失败
  // （需在图片注入、复选框处理和占位符包裹之前执行）
  Object.keys(templateZip.files).forEach(filename => {
    if (filename.endsWith('.xml')) {
      templateZip.file(
        filename,
        mergeAdjacentRuns(templateZip.file(filename).asText()),
      );
    }
  });

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
        const regex = new RegExp(escapeRegExp(key), 'g');
        if (!updatedContent.includes(`{${key}}`)) {
          // 只修复 Word 文本节点中的裸字段名，避免改坏 XML 属性/元数据。
          updatedContent = updatedContent.replace(
            /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g,
            (_match, openTag, text, closeTag) =>
              `${openTag}${text.replace(regex, () => `{${key}}`)}${closeTag}`,
          );
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
  const baseFileName =
    sanitizeFileName(rowData[firstKey]) || `doc_${index + 1}`;

  return {
    buffer: docxArrayBuffer,
    name: `${baseFileName}.docx`,
  };
};
