import { describe, expect, it } from 'vitest';
import PizZip from 'pizzip';
import {
  generateDocx,
  mergeAdjacentRuns,
  resolveDuplicateName,
  sanitizeFileName,
} from '../../src/utils/docxGenerator';

describe('mergeAdjacentRuns', () => {
  it('joins adjacent simple runs with matching formatting across proofing markers', () => {
    const xml = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Fir</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:b/></w:rPr><w:t>st</w:t></w:r></w:p>';
    const merged = mergeAdjacentRuns(xml);
    expect(merged).toContain('<w:t xml:space="preserve">First</w:t>');
    expect(merged).not.toContain('spellStart');
  });

  it('does not merge runs that contain template braces or checkbox symbols', () => {
    const xml = '<w:r><w:t>{Name}</w:t></w:r><w:r><w:t>□</w:t></w:r>';
    expect(mergeAdjacentRuns(xml)).toBe(xml);
  });
});

describe('file naming', () => {
  it('replaces invalid Windows filename characters and trims whitespace', () => {
    expect(sanitizeFileName('  a/b:c*?"<>|  ')).toBe('a_b_c______');
    expect(sanitizeFileName(null)).toBe('');
  });

  it('adds increasing suffixes before the file extension', () => {
    const names = new Set(['report.docx', 'report_2.docx']);
    expect(resolveDuplicateName('report.docx', names)).toBe('report_3.docx');
    expect(resolveDuplicateName('README', names)).toBe('README');
    expect(names.has('README')).toBe(true);
  });
});

describe('generateDocx', () => {
  const makeTemplate = () => {
    const zip = new PizZip();
    zip.file(
      'word/document.xml',
      '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello {Name}!</w:t></w:r></w:p></w:body></w:document>',
    );
    zip.file(
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    );
    zip.file(
      '_rels/.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    );
    return zip.generate({ type: 'uint8array' });
  };

  it('renders a DOCX buffer and derives a safe filename from the first field', () => {
    const result = generateDocx(makeTemplate(), { Name: 'Ada Lovelace' }, 0);
    expect(result.name).toBe('Ada Lovelace.docx');
    expect(result.buffer).toBeInstanceOf(ArrayBuffer);

    const generated = new PizZip(result.buffer);
    const xml = generated.file('word/document.xml').asText();
    expect(xml).toContain('Hello Ada Lovelace!');
    expect(xml).not.toContain('{Name}');
  });

  it('uses a row-based fallback filename when the first field is blank', () => {
    const result = generateDocx(makeTemplate(), { Name: '  ' }, 2);
    expect(result.name).toBe('doc_3.docx');
  });
});
