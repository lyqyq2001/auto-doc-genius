// 图片替换核心逻辑：
// 1. 识别 Excel 中 img 开头的图片列
// 2. 解析单元格中的图片路径（绝对路径 / 相对 Excel 目录）
// 3. 将 Word 模板 document.xml 中的图片占位符替换为 <w:drawing>，
//    并把图片写入 docx 包的 word/media/，注册 rels 与 Content Types

// 1 像素(96dpi) 对应的 EMU 值
const PX_TO_EMU = 9525;
// 图片最大宽度 15cm（EMU），超出等比缩放
const MAX_WIDTH_EMU = 5400000;

// 扩展名 -> ContentType 映射
const CONTENT_TYPE_MAP = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
};

// 判断表头是否为图片列（img 开头，不区分大小写）
export const isImageColumn = header =>
  typeof header === 'string' && header.trim().toLowerCase().startsWith('img');

// 计算 EMU 尺寸，超宽时等比缩放
export const calcEmuSize = (widthPx, heightPx) => {
  let cx = Math.round(widthPx * PX_TO_EMU);
  let cy = Math.round(heightPx * PX_TO_EMU);

  if (cx > MAX_WIDTH_EMU) {
    cy = Math.round((cy * MAX_WIDTH_EMU) / cx);
    cx = MAX_WIDTH_EMU;
  }

  return { cx, cy };
};

// 解析单元格中的图片路径
// - 绝对路径（盘符/UNC/POSIX）原样返回
// - 相对路径基于 Excel 所在目录解析，解析结果必须仍位于该目录下，否则视为非法
export const resolveImagePath = (cellValue, excelDir) => {
  const value = String(cellValue ?? '').trim();
  if (!value) return '';

  // Windows 盘符绝对路径、UNC 路径、POSIX 绝对路径
  if (/^([a-zA-Z]:[\\/]|\\\\|\/)/.test(value)) {
    return value.replace(/\//g, '\\');
  }

  if (!excelDir) return '';

  const sep = excelDir.includes('\\') ? '\\' : '/';
  const normalizedBase = excelDir.replace(/[\\/]+$/, '');
  const baseParts = normalizedBase.split(/[\\/]+/).filter(Boolean);
  const segments = value.split(/[\\/]+/).filter(Boolean);

  for (const segment of segments) {
    if (segment === '.') continue;
    if (segment === '..') {
      baseParts.pop();
      continue;
    }
    baseParts.push(segment);
  }

  const prefix = normalizedBase.startsWith('\\\\') ? '\\\\' : '';
  const resolved = prefix + baseParts.join(sep);

  // 防止 .. 越权：解析结果必须仍在 Excel 目录之下
  if (!resolved.toLowerCase().startsWith(normalizedBase.toLowerCase())) {
    return '';
  }

  return resolved;
};

// 收集所有数据行中的图片路径（去重）
export const collectImagePaths = (rows, excelDir) => {
  const paths = new Set();

  rows.forEach(row => {
    Object.keys(row).forEach(key => {
      if (!isImageColumn(key)) return;
      const resolved = resolveImagePath(row[key], excelDir);
      if (resolved) paths.add(resolved);
    });
  });

  return [...paths];
};

// 生成 <w:drawing> 所在 run 的 XML
const buildDrawingRun = ({ cx, cy, rId, docPrId, fileName }) =>
  '<w:r>' +
  '<w:drawing>' +
  '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
  `<wp:extent cx="${cx}" cy="${cy}"/>` +
  '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  `<wp:docPr id="${docPrId}" name="Picture ${docPrId}"/>` +
  '<wp:cNvGraphicFramePr>' +
  '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>' +
  '</wp:cNvGraphicFramePr>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:nvPicPr>' +
  `<pic:cNvPr id="${docPrId}" name="${fileName}"/>` +
  '<pic:cNvPicPr/>' +
  '</pic:nvPicPr>' +
  '<pic:blipFill>' +
  `<a:blip r:embed="${rId}"/>` +
  '<a:stretch><a:fillRect/></a:stretch>' +
  '</pic:blipFill>' +
  '<pic:spPr>' +
  `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '</pic:spPr>' +
  '</pic:pic>' +
  '</a:graphicData>' +
  '</a:graphic>' +
  '</wp:inline>' +
  '</w:drawing>' +
  '</w:r>';

// 用文本内容构造一个保留原 run 属性的文本 run
const buildTextRun = (rPr, text) =>
  `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;

// 提取 run 中的 <w:rPr>...</w:rPr>（含自闭合），找不到返回空串
const extractRunPr = runXml => {
  const match = runXml.match(/<w:rPr(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/w:rPr>)/);
  return match ? match[0] : '';
};

// 找到 key 所在位置外层包裹的 <w:r>...</w:r> 边界，找不到返回 null
const findEnclosingRun = (xml, keyIndex) => {
  const before = xml.slice(0, keyIndex);
  const runStartRegex = /<w:r(?:\s[^>]*)?>/g;
  let runStart = -1;
  let match;

  while ((match = runStartRegex.exec(before)) !== null) {
    runStart = match.index;
  }

  if (runStart === -1) return null;

  // run 起点与 key 之间不能已有 </w:r>，否则 key 不在该 run 内
  const closingBetween = xml.indexOf('</w:r>', runStart);
  if (closingBetween !== -1 && closingBetween < keyIndex) return null;

  const runEnd = xml.indexOf('</w:r>', keyIndex);
  if (runEnd === -1) return null;

  return { start: runStart, end: runEnd + '</w:r>'.length };
};

// 将 xml 中所有包含 key 的 run 替换为 replacementBuilder 生成的 XML。
// 若占位符与其他文本同处一个 run，则拆分 run：前后文本保留，仅占位符替换为图片。
const replacePlaceholderRuns = (xml, key, replacementBuilder) => {
  let result = '';
  let rest = xml;
  let guard = 0;

  while (guard++ < 200) {
    const keyIndex = rest.indexOf(key);
    if (keyIndex === -1) {
      result += rest;
      break;
    }

    const runRange = findEnclosingRun(rest, keyIndex);
    if (!runRange) {
      // 异常结构：跳过该处文本，避免死循环
      result += rest.slice(0, keyIndex + key.length);
      rest = rest.slice(keyIndex + key.length);
      continue;
    }

    const runXml = rest.slice(runRange.start, runRange.end);
    const rPr = extractRunPr(runXml);

    // 在 run 内定位包含 key 的 <w:t>，拆分前后文本
    const keyPosInRun = runXml.indexOf(key);
    const wtRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    let beforeText = '';
    let afterText = '';
    let wtMatch;

    while ((wtMatch = wtRegex.exec(runXml)) !== null) {
      const text = wtMatch[1];
      const innerIndex = text.indexOf(key);

      if (innerIndex === -1) {
        if (wtMatch.index < keyPosInRun) {
          beforeText += text;
        } else {
          afterText += text;
        }
        continue;
      }

      beforeText += text.slice(0, innerIndex);
      afterText += text.slice(innerIndex + key.length);
    }

    let replacement = '';
    if (beforeText) replacement += buildTextRun(rPr, beforeText);
    replacement += replacementBuilder();
    if (afterText) replacement += buildTextRun(rPr, afterText);

    result += rest.slice(0, runRange.start);
    result += replacement;
    rest = rest.slice(runRange.end);
  }

  return result;
};

// rels 中已有 rId 的最大编号
const getMaxRId = relsXml => {
  let max = 0;
  const regex = /Id="rId(\d+)"/g;
  let match;

  while ((match = regex.exec(relsXml)) !== null) {
    max = Math.max(max, Number(match[1]));
  }

  return max;
};

// 追加图片关系到 rels XML
const addRelationship = (relsXml, rId, mediaName) =>
  relsXml.replace(
    '</Relationships>',
    `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${mediaName}"/></Relationships>`,
  );

// 按需补齐 [Content_Types].xml 中的扩展名声明
const ensureContentType = (xml, ext) => {
  const exists = new RegExp(`<Default\\s+Extension="${ext}"`, 'i');
  if (exists.test(xml)) return xml;

  const contentType = CONTENT_TYPE_MAP[ext] || `image/${ext}`;
  return xml.replace(
    /<Types[^>]*>/,
    match =>
      `${match}<Default Extension="${ext}" ContentType="${contentType}"/>`,
  );
};

const DEFAULT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '</Relationships>';

// 对模板 zip 执行图片注入。
// imageAssets: { 绝对路径: { ok, base64, width, height, ext } }
// 返回注入成功的图片数量
export const injectImages = (templateZip, rowData, imageAssets, excelDir) => {
  const documentFile = templateZip.file('word/document.xml');
  if (!documentFile) return 0;

  let documentXml = documentFile.asText();
  let relsFile = templateZip.file('word/_rels/document.xml.rels');
  let relsXml = relsFile ? relsFile.asText() : DEFAULT_RELS;
  let contentTypesXml = templateZip.file('[Content_Types].xml').asText();

  let nextRId = getMaxRId(relsXml) + 1;
  let docPrId = 1000;
  let seq = 0;
  let injected = 0;
  let relsChanged = false;
  let contentTypesChanged = false;

  Object.keys(rowData).forEach(key => {
    if (!isImageColumn(key)) return;
    if (!documentXml.includes(key)) return;

    const resolved = resolveImagePath(rowData[key], excelDir);
    const asset = resolved ? imageAssets?.[resolved] : null;

    if (!asset || !asset.ok) {
      // 缺图/读取失败：移除占位符，位置留空
      documentXml = replacePlaceholderRuns(documentXml, key, () => '');
      return;
    }

    seq += 1;
    const mediaName = `imggen_${seq}.${asset.ext}`;
    const rId = `rId${nextRId++}`;
    const { cx, cy } = calcEmuSize(asset.width, asset.height);
    const currentDocPrId = docPrId++;

    templateZip.file(`word/media/${mediaName}`, asset.base64, {
      base64: true,
    });
    relsXml = addRelationship(relsXml, rId, mediaName);
    relsChanged = true;

    const newContentTypes = ensureContentType(contentTypesXml, asset.ext);
    if (newContentTypes !== contentTypesXml) {
      contentTypesXml = newContentTypes;
      contentTypesChanged = true;
    }

    documentXml = replacePlaceholderRuns(documentXml, key, () =>
      buildDrawingRun({
        cx,
        cy,
        rId,
        docPrId: currentDocPrId,
        fileName: mediaName,
      }),
    );
    injected += 1;
  });

  templateZip.file('word/document.xml', documentXml);
  if (relsChanged) {
    templateZip.file('word/_rels/document.xml.rels', relsXml);
  }
  if (contentTypesChanged) {
    templateZip.file('[Content_Types].xml', contentTypesXml);
  }

  return injected;
};
