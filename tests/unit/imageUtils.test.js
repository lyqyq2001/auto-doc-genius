import { describe, expect, it } from 'vitest';
import PizZip from 'pizzip';
import {
  calcEmuSize,
  collectImagePaths,
  injectImages,
  isImageColumn,
  resolveImagePath,
} from '../../src/utils/imageUtils';

// 1x1 透明 PNG 的 base64
const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '</Types>';

const buildTemplateZip = (bodyXml, relsXml) => {
  const zip = new PizZip();
  zip.file(
    'word/document.xml',
    `<w:document><w:body>${bodyXml}</w:body></w:document>`,
  );
  if (relsXml !== null) {
    zip.file('word/_rels/document.xml.rels', relsXml);
  }
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  return zip;
};

const run = (text, rPr = '') =>
  `<w:p><w:r>${rPr}<w:t>${text}</w:t></w:r></w:p>`;

describe('isImageColumn', () => {
  it('识别 img 开头的表头，且不区分大小写与首尾空格', () => {
    expect(isImageColumn('imgPhoto')).toBe(true);
    expect(isImageColumn('IMG_1')).toBe(true);
    expect(isImageColumn('  img2  ')).toBe(true);
  });

  it('非字符串或非 img 开头返回 false', () => {
    expect(isImageColumn('photo')).toBe(false);
    expect(isImageColumn('图片img')).toBe(false);
    expect(isImageColumn(undefined)).toBe(false);
    expect(isImageColumn(123)).toBe(false);
  });
});

describe('calcEmuSize', () => {
  it('按 96dpi 将像素换算为 EMU', () => {
    expect(calcEmuSize(100, 50)).toEqual({ cx: 952500, cy: 476250 });
  });

  it('超过 15cm 最大宽度时等比缩放', () => {
    const { cx, cy } = calcEmuSize(1000, 500);
    expect(cx).toBe(5400000);
    expect(cy).toBe(Math.round((500 * 9525 * 5400000) / (1000 * 9525)));
    expect(cy / cx).toBeCloseTo(500 / 1000, 6);
  });

  it('未超过最大宽度时不缩放', () => {
    expect(calcEmuSize(500, 200)).toEqual({ cx: 4762500, cy: 1905000 });
  });
});

describe('resolveImagePath', () => {
  it('空值返回空串', () => {
    expect(resolveImagePath('', 'C:\\data')).toBe('');
    expect(resolveImagePath('   ', 'C:\\data')).toBe('');
    expect(resolveImagePath(null, 'C:\\data')).toBe('');
    expect(resolveImagePath(undefined, 'C:\\data')).toBe('');
  });

  it('Windows 盘符绝对路径原样返回并统一分隔符', () => {
    expect(resolveImagePath('C:/imgs/a.png', '')).toBe('C:\\imgs\\a.png');
    expect(resolveImagePath('D:\\imgs\\a.png', '')).toBe('D:\\imgs\\a.png');
  });

  it('UNC 与 POSIX 绝对路径原样返回', () => {
    expect(resolveImagePath('\\\\server\\share\\a.png', '')).toBe(
      '\\\\server\\share\\a.png',
    );
    expect(resolveImagePath('/home/user/a.png', '')).toBe(
      '\\home\\user\\a.png',
    );
  });

  it('相对路径基于 Excel 目录解析', () => {
    expect(resolveImagePath('imgs/a.png', 'C:\\data\\excel')).toBe(
      'C:\\data\\excel\\imgs\\a.png',
    );
    expect(resolveImagePath('./imgs/a.png', 'C:\\data\\excel')).toBe(
      'C:\\data\\excel\\imgs\\a.png',
    );
    expect(resolveImagePath('sub/../a.png', 'C:\\data\\excel')).toBe(
      'C:\\data\\excel\\a.png',
    );
    // 目录以分隔符结尾时同样可解析
    expect(resolveImagePath('a.png', 'C:\\data\\excel\\')).toBe(
      'C:\\data\\excel\\a.png',
    );
  });

  it('相对路径缺少 Excel 目录时返回空串', () => {
    expect(resolveImagePath('imgs/a.png', '')).toBe('');
  });

  it('阻止 .. 越出 Excel 目录', () => {
    expect(resolveImagePath('../../secret.png', 'C:\\data\\excel')).toBe('');
    // 解析结果必须仍在 Excel 目录之下，上一级目录同样被拒绝
    expect(resolveImagePath('../a.png', 'C:\\data\\excel\\sub')).toBe('');
  });

  it('POSIX 风格目录下的相对路径当前无法解析', () => {
    // 已知限制：excelDir 以 "/" 开头时，前缀被 filter(Boolean) 丢弃导致越权校验失败
    expect(resolveImagePath('imgs/a.png', '/data/excel')).toBe('');
  });
});

describe('collectImagePaths', () => {
  it('只收集图片列且去重', () => {
    const rows = [
      { name: 'A', imgPhoto: 'C:\\imgs\\a.png', imgSign: 'imgs/b.png' },
      { name: 'B', imgPhoto: 'C:\\imgs\\a.png' },
    ];
    expect(collectImagePaths(rows, 'C:\\data').sort()).toEqual(
      ['C:\\imgs\\a.png', 'C:\\data\\imgs\\b.png'].sort(),
    );
  });

  it('无图片列或路径非法时返回空数组', () => {
    expect(collectImagePaths([{ name: 'A' }], 'C:\\data')).toEqual([]);
    expect(collectImagePaths([{ imgPhoto: '  ' }], 'C:\\data')).toEqual([]);
  });
});

describe('injectImages', () => {
  it('缺少 word/document.xml 时直接返回 0', () => {
    const zip = new PizZip();
    zip.file('[Content_Types].xml', CONTENT_TYPES);
    expect(injectImages(zip, { imgPhoto: 'C:\\a.png' }, {}, '')).toBe(0);
  });

  it('注入图片并写入 media、rels 与 Content Types', () => {
    const zip = buildTemplateZip(run('imgPhoto'), null);
    const assets = {
      'C:\\imgs\\a.png': {
        ok: true,
        base64: PNG_BASE64,
        width: 100,
        height: 50,
        ext: 'png',
      },
    };

    const injected = injectImages(
      zip,
      { imgPhoto: 'C:\\imgs\\a.png' },
      assets,
      '',
    );

    expect(injected).toBe(1);
    const xml = zip.file('word/document.xml').asText();
    expect(xml).toContain('<w:drawing>');
    expect(xml).toContain('<wp:extent cx="952500" cy="476250"/>');
    expect(xml).toContain('r:embed="rId1"');
    expect(xml).not.toContain('imgPhoto');

    expect(zip.file('word/media/imggen_1.png')).not.toBeNull();
    expect(zip.file('word/_rels/document.xml.rels').asText()).toContain(
      'Target="media/imggen_1.png"',
    );
    expect(zip.file('[Content_Types].xml').asText()).toContain(
      '<Default Extension="png" ContentType="image/png"/>',
    );
  });

  it('沿用已有 rels 并递增 rId 编号', () => {
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId5" Type="http://x" Target="styles.xml"/>' +
      '</Relationships>';
    const zip = buildTemplateZip(run('imgPhoto'), relsXml);

    injectImages(
      zip,
      { imgPhoto: 'C:\\imgs\\a.jpg' },
      {
        'C:\\imgs\\a.jpg': {
          ok: true,
          base64: PNG_BASE64,
          width: 10,
          height: 10,
          ext: 'jpg',
        },
      },
      '',
    );

    const rels = zip.file('word/_rels/document.xml.rels').asText();
    expect(rels).toContain('Id="rId5"');
    expect(rels).toContain('Id="rId6"');
    expect(rels).toContain('Target="media/imggen_1.jpg"');
  });

  it('已声明的扩展名不重复写入 Content Types', () => {
    const zip = new PizZip();
    zip.file(
      'word/document.xml',
      `<w:document><w:body>${run('imgPhoto')}</w:body></w:document>`,
    );
    zip.file(
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '</Types>',
        '<Default Extension="png" ContentType="image/png"/></Types>',
      ),
    );

    injectImages(
      zip,
      { imgPhoto: 'C:\\imgs\\a.png' },
      {
        'C:\\imgs\\a.png': {
          ok: true,
          base64: PNG_BASE64,
          width: 10,
          height: 10,
          ext: 'png',
        },
      },
      '',
    );

    const contentTypes = zip.file('[Content_Types].xml').asText();
    expect(contentTypes.match(/Extension="png"/g)).toHaveLength(1);
  });

  it('图片缺失或读取失败时移除占位符并留空', () => {
    const zip = buildTemplateZip(run('imgPhoto'), null);
    const injected = injectImages(
      zip,
      { imgPhoto: 'C:\\imgs\\missing.png' },
      { 'C:\\imgs\\missing.png': { ok: false } },
      '',
    );

    expect(injected).toBe(0);
    const xml = zip.file('word/document.xml').asText();
    expect(xml).not.toContain('imgPhoto');
    expect(xml).not.toContain('<w:drawing>');
  });

  it('非图片列与模板中不存在的列都跳过', () => {
    const zip = buildTemplateZip(run('imgPhoto'), null);
    const injected = injectImages(
      zip,
      { name: 'C:\\imgs\\a.png', imgOther: 'C:\\imgs\\a.png' },
      {},
      '',
    );

    expect(injected).toBe(0);
    expect(zip.file('word/document.xml').asText()).toContain('imgPhoto');
  });

  it('占位符与文本同处一个 run 时拆分并保留 run 属性', () => {
    const zip = buildTemplateZip(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>照片: imgPhoto 结束</w:t></w:r></w:p>',
      null,
    );

    const injected = injectImages(
      zip,
      { imgPhoto: 'C:\\imgs\\a.png' },
      {
        'C:\\imgs\\a.png': {
          ok: true,
          base64: PNG_BASE64,
          width: 10,
          height: 10,
          ext: 'png',
        },
      },
      '',
    );

    expect(injected).toBe(1);
    const xml = zip.file('word/document.xml').asText();
    expect(xml).toContain('<w:t xml:space="preserve">照片: </w:t>');
    expect(xml).toContain('<w:t xml:space="preserve"> 结束</w:t>');
    expect(xml).toContain('<w:rPr><w:b/></w:rPr>');
    expect(xml).toContain('<w:drawing>');
  });

  it('同一占位符出现多次时全部替换', () => {
    const zip = buildTemplateZip(run('imgPhoto') + run('imgPhoto'), null);

    const injected = injectImages(
      zip,
      { imgPhoto: 'C:\\imgs\\a.png' },
      {
        'C:\\imgs\\a.png': {
          ok: true,
          base64: PNG_BASE64,
          width: 10,
          height: 10,
          ext: 'png',
        },
      },
      '',
    );

    expect(injected).toBe(1);
    const xml = zip.file('word/document.xml').asText();
    expect(xml.match(/<w:drawing>/g)).toHaveLength(2);
  });

  it('占位符不在任何 run 内时保持原样且不卡死', () => {
    const zip = new PizZip();
    zip.file('word/document.xml', '<w:document>imgPhoto</w:document>');
    zip.file('[Content_Types].xml', CONTENT_TYPES);

    injectImages(
      zip,
      { imgPhoto: 'C:\\imgs\\a.png' },
      {
        'C:\\imgs\\a.png': {
          ok: true,
          base64: PNG_BASE64,
          width: 10,
          height: 10,
          ext: 'png',
        },
      },
      '',
    );

    expect(zip.file('word/document.xml').asText()).toContain('imgPhoto');
  });
});
