import { beforeEach, describe, expect, it } from 'vitest';
import { utils as XLSXUtils, write as XLSXWrite } from 'xlsx';
import {
  parseExcel,
  processCheckboxOptions,
  readWordTemplate,
} from '../../src/utils/documentUtils';
import { i18n } from '../../src/i18n';
import en from '../../src/locales/en.json';

// 用单元格对象直接构造工作表，便于精确控制数字格式（cell.z）
const buildSheet = cells => {
  const sheet = { ...cells };
  const end = Object.keys(cells)
    .map(addr => XLSXUtils.decode_cell(addr))
    .reduce(
      (acc, cell) => ({
        r: Math.max(acc.r, cell.r),
        c: Math.max(acc.c, cell.c),
      }),
      { r: 0, c: 0 },
    );
  sheet['!ref'] = XLSXUtils.encode_range({ s: { r: 0, c: 0 }, e: end });
  return sheet;
};

const buildExcelFile = (cells, name = 'data.xlsx') => {
  const workbook = XLSXUtils.book_new();
  XLSXUtils.book_append_sheet(workbook, buildSheet(cells), 'Sheet1');
  const buffer = XLSXWrite(workbook, { bookType: 'xlsx', type: 'array' });
  return new File([buffer], name, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
};

beforeEach(() => {
  i18n.global.locale.value = 'en';
});

describe('parseExcel', () => {
  it('跳过描述行，以第二行为表头解析数据行', async () => {
    const file = buildExcelFile({
      A1: { t: 's', v: '这是描述行' },
      A2: { t: 's', v: 'name' },
      B2: { t: 's', v: 'city' },
      A3: { t: 's', v: 'Alice' },
      B3: { t: 's', v: 'Beijing' },
      A4: { t: 's', v: 'Bob' },
      B4: { t: 's', v: 'Shanghai' },
    });

    await expect(parseExcel(file)).resolves.toEqual([
      { name: 'Alice', city: 'Beijing' },
      { name: 'Bob', city: 'Shanghai' },
    ]);
  });

  it('忽略空行，缺失单元格补空串', async () => {
    const file = buildExcelFile({
      A2: { t: 's', v: 'name' },
      B2: { t: 's', v: 'city' },
      A3: { t: 's', v: 'Alice' },
      A4: { t: 's', v: '' },
      A5: { t: 's', v: 'Bob' },
      B5: { t: 's', v: 'Shanghai' },
    });

    await expect(parseExcel(file)).resolves.toEqual([
      { name: 'Alice', city: '' },
      { name: 'Bob', city: 'Shanghai' },
    ]);
  });

  it('日期与时间格式的数值单元格按当前语言格式化', async () => {
    const file = buildExcelFile({
      A2: { t: 's', v: 'name' },
      B2: { t: 's', v: 'date' },
      C2: { t: 's', v: 'time' },
      D2: { t: 's', v: 'weight' },
      A3: { t: 's', v: 'Alice' },
      // 45000 -> 2023-03-15
      B3: { t: 'n', v: 45000, z: 'yyyy-mm-dd' },
      // 0.5 -> 12:00
      C3: { t: 'n', v: 0.5, z: 'hh:mm' },
      // 普通大数字不应被误判为日期
      D3: { t: 'n', v: 45000 },
    });

    await expect(parseExcel(file)).resolves.toEqual([
      { name: 'Alice', date: '15-Mar-2023', time: '12:00', weight: 45000 },
    ]);
  });

  it('中文环境下日期使用中文格式', async () => {
    i18n.global.locale.value = 'zh';
    const file = buildExcelFile({
      A2: { t: 's', v: 'date' },
      A3: { t: 'n', v: 45000, z: 'yyyy-mm-dd' },
    });

    await expect(parseExcel(file)).resolves.toEqual([
      { date: '2023年03月15日' },
    ]);
  });

  it('行数不足 3 行时拒绝并提示格式错误', async () => {
    const file = buildExcelFile({
      A1: { t: 's', v: '描述行' },
      A2: { t: 's', v: 'name' },
    });

    await expect(parseExcel(file)).rejects.toThrow(en.error.excelFormat);
  });

  it('没有有效数据行时拒绝并提示无数据', async () => {
    const file = buildExcelFile({
      A1: { t: 's', v: '描述行' },
      A2: { t: 's', v: 'name' },
      A3: { t: 's', v: '' },
    });

    await expect(parseExcel(file)).rejects.toThrow(en.error.excelNoData);
  });

  it('文件损坏无法解析时拒绝并提示解析失败', async () => {
    // 带 ZIP 魔数但内容损坏，xlsx 解析时会抛错
    const corrupt = new Uint8Array([
      0x50, 0x4b, 0x03, 0x04, 0xff, 0xfe, 0x01, 0x02, 0x03,
    ]);
    const file = new File([corrupt], 'bad.xlsx');

    await expect(parseExcel(file)).rejects.toThrow(en.error.excelParseFailed);
  });
});

describe('readWordTemplate', () => {
  it('以 ArrayBuffer 形式读取文件内容', async () => {
    const bytes = new Uint8Array([1, 2, 3, 4]);
    const file = new File([bytes], 'template.docx');

    const result = await readWordTemplate(file);
    expect(result).toBeInstanceOf(ArrayBuffer);
    expect(new Uint8Array(result)).toEqual(bytes);
  });
});

describe('processCheckboxOptions', () => {
  const options = '<w:t>□</w:t><w:t>男</w:t><w:t>□</w:t><w:t>女</w:t>';

  it('命中选项时把对应 □ 替换为 ☑', () => {
    const content = `性别：{{check1${options}}}`;
    expect(processCheckboxOptions(content, { check1: '男' })).toBe(
      '性别：<w:t>☑</w:t><w:t>男</w:t><w:t>□</w:t><w:t>女</w:t>',
    );
  });

  it('支持中英文逗号与顿号分隔的多选', () => {
    const multi =
      '<w:t>□</w:t><w:t>A</w:t><w:t>□</w:t><w:t>B</w:t><w:t>□</w:t><w:t>C</w:t>';
    const result = processCheckboxOptions(`{{check2${multi}}}`, {
      check2: 'A、C',
    });
    expect(result).toBe(
      '<w:t>☑</w:t><w:t>A</w:t><w:t>□</w:t><w:t>B</w:t><w:t>☑</w:t><w:t>C</w:t>',
    );
  });

  it('Excel 值为空时保留原样并去掉占位符包裹', () => {
    expect(processCheckboxOptions(`{{check1${options}}}`, { check1: '' })).toBe(
      options,
    );
    expect(processCheckboxOptions(`{{check1${options}}}`, {})).toBe(options);
  });

  it('选项内没有 w:t 文本时原样返回', () => {
    expect(processCheckboxOptions('{{check1}}', { check1: '男' })).toBe('');
  });

  it('未命中任何选项时全部保持未选中', () => {
    expect(
      processCheckboxOptions(`{{check1${options}}}`, { check1: '未知' }),
    ).toBe(options);
  });

  it('不影响非 check 占位符', () => {
    const content = '<w:t>{name}</w:t>';
    expect(processCheckboxOptions(content, { name: 'Alice' })).toBe(content);
  });

  it('支持多个复选框占位符', () => {
    const content = `{{check1${options}}}{{check3<w:t>□</w:t><w:t>是</w:t>}}`;
    expect(
      processCheckboxOptions(content, { check1: '女', check3: '是' }),
    ).toBe(
      '<w:t>□</w:t><w:t>男</w:t><w:t>☑</w:t><w:t>女</w:t><w:t>☑</w:t><w:t>是</w:t>',
    );
  });
});
