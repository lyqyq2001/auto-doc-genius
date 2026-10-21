import { describe, expect, it } from 'vitest';
import {
  calcEmuSize,
  collectImagePaths,
  injectImages,
  isImageColumn,
  resolveImagePath,
} from '../../src/utils/imageUtils';

describe('image column and path helpers', () => {
  it('recognizes trimmed img headers without case sensitivity', () => {
    expect(isImageColumn('img')).toBe(true);
    expect(isImageColumn('  ImgPortrait ')).toBe(true);
    expect(isImageColumn('photo')).toBe(false);
    expect(isImageColumn(null)).toBe(false);
  });

  it('resolves relative paths beneath the Excel folder and normalizes separators', () => {
    expect(resolveImagePath('photos/a.png', 'C:\\docs')).toBe('C:\\docs\\photos\\a.png');
    expect(resolveImagePath('photos\\a.png', '/docs')).toBe('/docs/photos/a.png');
    expect(resolveImagePath('C:/assets/a.png', '/docs')).toBe('C:\\assets\\a.png');
  });

  it('rejects paths that escape the Excel folder and empty relative bases', () => {
    expect(resolveImagePath('../../secret.png', '/docs/files')).toBe('');
    expect(resolveImagePath('../secret.png', 'C:\\docs')).toBe('');
    expect(resolveImagePath('image.png', '')).toBe('');
    expect(resolveImagePath('   ', '/docs')).toBe('');
  });

  it('collects unique paths from image columns only', () => {
    expect(
      collectImagePaths(
        [
          { img1: 'a.png', Name: 'b.png' },
          { IMG2: 'a.png', img3: 'nested/c.png' },
        ],
        '/excel',
      ),
    ).toEqual(['/excel/a.png', '/excel/nested/c.png']);
  });

  it('converts pixels to EMU and scales images wider than 15 cm', () => {
    expect(calcEmuSize(100, 50)).toEqual({ cx: 952500, cy: 476250 });
    expect(calcEmuSize(1000, 500)).toEqual({ cx: 5400000, cy: 2700000 });
  });
});

const mockZip = entries => {
  const files = { ...entries };
  return {
    file(path, data, options) {
      if (arguments.length === 1) {
        return files[path] == null ? null : { asText: () => files[path] };
      }
      files[path] = options?.base64 ? `BASE64:${data}` : data;
      return this;
    },
    files,
  };
};

describe('injectImages', () => {
  const contentTypes = '<Types></Types>';
  const rels = '<Relationships><Relationship Id="rId3"/></Relationships>';

  it('injects image content, relationship, size, and content type', () => {
    const zip = mockZip({
      'word/document.xml': '<w:document><w:r><w:t>imgPhoto</w:t></w:r></w:document>',
      'word/_rels/document.xml.rels': rels,
      '[Content_Types].xml': contentTypes,
    });
    const count = injectImages(
      zip,
      { imgPhoto: 'photo.png' },
      { '/excel/photo.png': { ok: true, base64: 'abc', width: 100, height: 50, ext: 'png' } },
      '/excel',
    );

    expect(count).toBe(1);
    expect(zip.files['word/document.xml']).toContain('<wp:extent cx="952500" cy="476250"/>');
    expect(zip.files['word/document.xml']).toContain('r:embed="rId4"');
    expect(zip.files['word/media/imggen_1.png']).toBe('BASE64:abc');
    expect(zip.files['word/_rels/document.xml.rels']).toContain('Target="media/imggen_1.png"');
    expect(zip.files['[Content_Types].xml']).toContain('Extension="png" ContentType="image/png"');
  });

  it('removes missing image placeholders without creating relationships', () => {
    const zip = mockZip({
      'word/document.xml': '<w:r><w:t>before imgPhoto after</w:t></w:r>',
      'word/_rels/document.xml.rels': rels,
      '[Content_Types].xml': contentTypes,
    });

    expect(injectImages(zip, { imgPhoto: 'missing.png' }, {}, '/excel')).toBe(0);
    expect(zip.files['word/document.xml']).not.toContain('imgPhoto');
    expect(zip.files['word/document.xml']).toContain('before ');
    expect(zip.files['word/document.xml']).toContain(' after');
    expect(zip.files['word/_rels/document.xml.rels']).toBe(rels);
  });

  it('returns zero if the document XML is absent', () => {
    expect(injectImages(mockZip({}), {}, {}, '')).toBe(0);
  });
});
