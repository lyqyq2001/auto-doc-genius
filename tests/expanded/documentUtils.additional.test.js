import { afterEach, describe, expect, it, vi } from 'vitest';
import { utils, write } from 'xlsx';
import { i18n } from '../../src/i18n';
import {
  parseExcel,
  processCheckboxOptions,
  readWordTemplate,
} from '../../src/utils/documentUtils';

const makeExcelFile = rows => {
  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, utils.aoa_to_sheet(rows), 'Data');
  return new File([write(workbook, { type: 'array', bookType: 'xlsx' })], 'data.xlsx');
};

afterEach(() => {
  vi.restoreAllMocks();
  i18n.global.locale.value = 'en';
});

describe('parseExcel', () => {
  it('parses headers from row two and non-empty records from row three', async () => {
    const data = await parseExcel(
      makeExcelFile([
        ['Title row'],
        ['Name', 'Count'],
        ['Alice', 3],
        [],
        ['Bob', 4],
      ]),
    );

    expect(data).toEqual([
      { Name: 'Alice', Count: 3 },
      { Name: 'Bob', Count: 4 },
    ]);
  });

  it('preserves zero and boolean values in cells', async () => {
    const data = await parseExcel(
      makeExcelFile([['Title'], ['Score', 'Active'], [0, false]]),
    );

    expect(data).toEqual([{ Score: 0, Active: false }]);
  });

  it('rejects a workbook without header and data rows', async () => {
    await expect(parseExcel(makeExcelFile([['Title'], ['Header']]))).rejects.toThrow();
  });

  it('rejects a workbook with headers but no records', async () => {
    await expect(parseExcel(makeExcelFile([['Title'], ['Header'], []]))).rejects.toThrow();
  });

  it('reports file read errors', async () => {
    const file = new File(['broken'], 'data.xlsx');
    const readSpy = vi.spyOn(FileReader.prototype, 'readAsArrayBuffer').mockImplementation(function () {
      this.onerror?.(new Event('error'));
    });

    await expect(parseExcel(file)).rejects.toThrow();
    expect(readSpy).toHaveBeenCalledOnce();
  });
});

describe('processCheckboxOptions', () => {
  const template = '{{check1}}';

  it('marks selected options for comma, Chinese comma, or enumeration comma input', () => {
    const xml = '{{check1}}';
    const options = '<w:t>□</w:t><w:t>Apple</w:t><w:t>□</w:t><w:t>Banana</w:t><w:t>□</w:t><w:t>Cherry</w:t>';

    expect(processCheckboxOptions(xml.replace('check1', `check1${options}`), { check1: 'Apple, Cherry' })).toContain('☑');
    expect(processCheckboxOptions('{{check1' + options + '}}', { check1: 'Apple，Banana' })).toContain('☑');
    expect(processCheckboxOptions('{{check1' + options + '}}', { check1: 'Banana、Cherry' })).toContain('☑');
  });

  it('leaves options unchanged when there is no selected value or no checkbox XML', () => {
    const content = '{{check1<w:t>□</w:t><w:t>Apple</w:t>}}';
    expect(processCheckboxOptions(content, { check1: '' })).toBe('<w:t>□</w:t><w:t>Apple</w:t>');
    expect(processCheckboxOptions('{{check2 plain}}', {})).toBe(' plain');
  });
});

describe('readWordTemplate', () => {
  it('reads a DOCX file as an ArrayBuffer', async () => {
    const file = new File(['docx bytes'], 'template.docx');
    const content = await readWordTemplate(file);
    expect(content).toBeInstanceOf(ArrayBuffer);
    expect(new TextDecoder().decode(content)).toBe('docx bytes');
  });

  it('rejects when the file cannot be read', async () => {
    vi.spyOn(FileReader.prototype, 'readAsArrayBuffer').mockImplementation(function () {
      this.onerror?.(new Event('error'));
    });
    await expect(readWordTemplate(new File(['x'], 'template.docx'))).rejects.toThrow();
  });
});
