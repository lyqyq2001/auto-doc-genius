import { describe, expect, it } from 'vitest';
import PizZip from 'pizzip';
import {
  generateDocx,
  mergeAdjacentRuns,
  resolveDuplicateName,
  sanitizeFileName,
} from '../../src/utils/docxGenerator';

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';

const ROOT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

const DOCUMENT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '</Relationships>';

// 构造一个最小可用的 docx 模板（ArrayBuffer）
const buildTemplate = bodyXml => {
  const zip = new PizZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS);
  zip.file(
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `<w:body>${bodyXml}</w:body></w:document>`,
  );
  return zip.generate({ type: 'arraybuffer' });
};

const paragraph = text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const readDocumentXml = buffer =>
  new PizZip(buffer).file('word/document.xml').asText();

describe('mergeAdjacentRuns', () => {
  it('合并相邻的同格式简单 run', () => {
    const xml = '<w:r><w:t>Hello</w:t></w:r><w:r><w:t>World</w:t></w:r>';
    expect(mergeAdjacentRuns(xml)).toBe(
      '<w:r><w:t xml:space="preserve">HelloWorld</w:t></w:r>',
    );
  });

  it('合并时保留 run 属性', () => {
    const rPr = '<w:rPr><w:b/></w:rPr>';
    const xml = `<w:r>${rPr}<w:t>A</w:t></w:r><w:r>${rPr}<w:t>B</w:t></w:r>`;
    expect(mergeAdjacentRuns(xml)).toBe(
      `<w:r>${rPr}<w:t xml:space="preserve">AB</w:t></w:r>`,
    );
  });

  it('run 属性不同时不合并', () => {
    const xml =
      '<w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r>';
    expect(mergeAdjacentRuns(xml)).toBe(xml);
  });

  it('文本含 { } □ 的 run 不参与合并，避免破坏占位符结构', () => {
    expect(
      mergeAdjacentRuns('<w:r><w:t>{na</w:t></w:r><w:r><w:t>me}</w:t></w:r>'),
    ).toBe('<w:r><w:t>{na</w:t></w:r><w:r><w:t>me}</w:t></w:r>');
    expect(
      mergeAdjacentRuns('<w:r><w:t>□</w:t></w:r><w:r><w:t>男</w:t></w:r>'),
    ).toBe('<w:r><w:t>□</w:t></w:r><w:r><w:t>男</w:t></w:r>');
  });

  it('run 之间的空白与拼写检查标记可被忽略', () => {
    const xml =
      '<w:r><w:t>A</w:t></w:r>  <w:proofErr w:type="spellStart"/><w:r><w:t>B</w:t></w:r>';
    expect(mergeAdjacentRuns(xml)).toBe(
      '<w:r><w:t xml:space="preserve">AB</w:t></w:r>',
    );
  });

  it('run 之间存在其他内容时不合并', () => {
    const xml = '<w:r><w:t>A</w:t></w:r><w:br/><w:r><w:t>B</w:t></w:r>';
    expect(mergeAdjacentRuns(xml)).toBe(xml);
  });

  it('含图片等非文本内容的 run 不参与合并', () => {
    const xml =
      '<w:r><w:t>A</w:t></w:r><w:r><w:drawing/></w:r><w:r><w:t>B</w:t></w:r>';
    expect(mergeAdjacentRuns(xml)).toBe(xml);
  });

  it('run 数量不足 2 时原样返回', () => {
    expect(mergeAdjacentRuns('<w:r><w:t>A</w:t></w:r>')).toBe(
      '<w:r><w:t>A</w:t></w:r>',
    );
    expect(mergeAdjacentRuns('<w:p/>')).toBe('<w:p/>');
  });

  it('多个 w:t 片段的 run 合并后文本按顺序拼接', () => {
    const xml = '<w:r><w:t>a</w:t><w:t>b</w:t></w:r><w:r><w:t>c</w:t></w:r>';
    expect(mergeAdjacentRuns(xml)).toBe(
      '<w:r><w:t xml:space="preserve">abc</w:t></w:r>',
    );
  });
});

describe('sanitizeFileName', () => {
  it('替换 Windows 文件名非法字符', () => {
    expect(sanitizeFileName('a\\b/c:d*e?f"g<h>i|j')).toBe(
      'a_b_c_d_e_f_g_h_i_j',
    );
  });

  it('替换控制字符并去掉首尾空白', () => {
    expect(sanitizeFileName('  报告\u0001v1  ')).toBe(
      '报告\u0001v1'.replace(/\u0001/, '_'),
    );
  });

  it('空值返回空串', () => {
    expect(sanitizeFileName('')).toBe('');
    expect(sanitizeFileName(null)).toBe('');
    expect(sanitizeFileName(undefined)).toBe('');
  });

  it('非字符串输入被转为字符串', () => {
    expect(sanitizeFileName(2023)).toBe('2023');
  });
});

describe('resolveDuplicateName', () => {
  it('首次出现的名称原样返回并登记', () => {
    const used = new Set();
    expect(resolveDuplicateName('a.docx', used)).toBe('a.docx');
    expect(used.has('a.docx')).toBe(true);
  });

  it('重名时依次追加序号', () => {
    const used = new Set();
    expect(resolveDuplicateName('a.docx', used)).toBe('a.docx');
    expect(resolveDuplicateName('a.docx', used)).toBe('a_2.docx');
    expect(resolveDuplicateName('a.docx', used)).toBe('a_3.docx');
  });

  it('无扩展名时直接在末尾追加序号', () => {
    const used = new Set();
    expect(resolveDuplicateName('report', used)).toBe('report');
    expect(resolveDuplicateName('report', used)).toBe('report_2');
  });

  it('已存在的序号名也会被跳过', () => {
    const used = new Set(['a.docx', 'a_2.docx']);
    expect(resolveDuplicateName('a.docx', used)).toBe('a_3.docx');
  });
});

describe('generateDocx', () => {
  it('渲染显式 {占位符} 并以首列值命名文件', () => {
    const template = buildTemplate(
      paragraph('姓名：{name}') + paragraph('城市：{city}'),
    );

    const result = generateDocx(
      template,
      { name: 'Alice', city: 'Beijing' },
      0,
    );

    expect(result.name).toBe('Alice.docx');
    const xml = readDocumentXml(result.buffer);
    expect(xml).toContain('姓名：Alice');
    expect(xml).toContain('城市：Beijing');
  });

  it('模板中的裸表头会被自动包裹为占位符后渲染', () => {
    const template = buildTemplate(paragraph('项目：project'));

    const result = generateDocx(template, { project: 'Alpha' }, 0);
    expect(readDocumentXml(result.buffer)).toContain('项目：Alpha');
  });

  it('Word 拆分到多个 run 的占位符在合并后仍可渲染', () => {
    const template = buildTemplate(
      '<w:p><w:r><w:t>姓名：{na</w:t></w:r><w:r><w:t>me}</w:t></w:r></w:p>',
    );

    const result = generateDocx(template, { name: 'Alice' }, 0);
    expect(readDocumentXml(result.buffer)).toContain('姓名：Alice');
  });

  it('首列为空时使用 doc_序号 兜底命名', () => {
    const template = buildTemplate(paragraph('{name}'));

    expect(generateDocx(template, { name: '' }, 0).name).toBe('doc_1.docx');
    expect(generateDocx(template, { name: '' }, 4).name).toBe('doc_5.docx');
  });

  it('首列含非法字符时清洗后再命名', () => {
    const template = buildTemplate(paragraph('{name}'));

    expect(generateDocx(template, { name: 'a/b:c' }, 0).name).toBe(
      'a_b_c.docx',
    );
  });

  it('处理复选框占位符', () => {
    // 复选框占位符在 Word 中被拆成多个 run，形如 {{check1 □男 □女 }}
    const template = buildTemplate(
      '<w:p>' +
        '<w:r><w:t>性别：{{check1</w:t></w:r>' +
        '<w:r><w:t>□</w:t></w:r>' +
        '<w:r><w:t>男</w:t></w:r>' +
        '<w:r><w:t>□</w:t></w:r>' +
        '<w:r><w:t>女</w:t></w:r>' +
        '<w:r><w:t>}}</w:t></w:r>' +
        '</w:p>',
    );

    const result = generateDocx(template, { check1: '女' }, 0);
    const xml = readDocumentXml(result.buffer);
    expect(xml).toContain('☑');
    expect(xml).not.toContain('check1');
  });

  it('注入图片占位符对应的图片', () => {
    const template = buildTemplate(
      paragraph('姓名：{name}') + paragraph('imgPhoto'),
    );
    const imageAssets = {
      'C:\\imgs\\a.png': {
        ok: true,
        base64: PNG_BASE64,
        width: 100,
        height: 50,
        ext: 'png',
      },
    };

    const result = generateDocx(
      template,
      { name: 'Alice', imgPhoto: 'C:\\imgs\\a.png' },
      0,
      imageAssets,
      '',
    );

    const zip = new PizZip(result.buffer);
    expect(zip.file('word/document.xml').asText()).toContain('<w:drawing>');
    expect(zip.file('word/media/imggen_1.png')).not.toBeNull();
  });

  it('未传 imageAssets 时图片列不注入图片，占位符原样保留', () => {
    const template = buildTemplate(paragraph('imgPhoto'));

    const result = generateDocx(template, { imgPhoto: 'C:\\imgs\\a.png' }, 0);
    const xml = readDocumentXml(result.buffer);
    expect(xml).not.toContain('<w:drawing>');
    // 图片列不会被当作普通占位符渲染成 Excel 中的路径文本
    expect(xml).toContain('imgPhoto');
    expect(xml).not.toContain('a.png');
  });

  it('返回的 buffer 是可被重新解析的 docx 包', () => {
    const template = buildTemplate(paragraph('{name}'));
    const result = generateDocx(template, { name: 'Alice' }, 0);

    expect(result.buffer).toBeInstanceOf(ArrayBuffer);
    const zip = new PizZip(result.buffer);
    expect(zip.file('[Content_Types].xml')).not.toBeNull();
    expect(zip.file('word/document.xml')).not.toBeNull();
  });
});
