import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { ElMessage } from 'element-plus';
import { i18n } from '../../src/i18n';
import {
  beforeExcelUpload,
  beforeWordUpload,
  downloadExcelTemplate,
  downloadWordTemplate,
} from '../../src/utils/fileUtils';

vi.mock('element-plus', () => ({
  ElMessage: {
    success: vi.fn(),
    error: vi.fn(),
    warning: vi.fn(),
  },
}));

const setNavigatorLanguage = language => {
  Object.defineProperty(navigator, 'language', {
    value: language,
    configurable: true,
  });
};

// 同时设置系统语言（决定模板文件名）与 i18n locale（决定提示文案）
const setLanguage = language => {
  setNavigatorLanguage(language);
  i18n.global.locale.value = language.toLowerCase().startsWith('zh')
    ? 'zh'
    : 'en';
};

const originalLanguage = navigator.language;

let clickSpy;

beforeEach(() => {
  vi.clearAllMocks();
  // jsdom 未实现导航下载，替换为空实现避免 Not implemented 报错
  clickSpy = vi
    .spyOn(HTMLAnchorElement.prototype, 'click')
    .mockImplementation(() => {});
});

afterEach(() => {
  clickSpy.mockRestore();
  setNavigatorLanguage(originalLanguage);
  i18n.global.locale.value = 'en';
});

describe('downloadExcelTemplate', () => {
  it('中文环境下载中文模板并提示成功', () => {
    setLanguage('zh-CN');

    downloadExcelTemplate();

    expect(clickSpy).toHaveBeenCalledTimes(1);
    expect(clickSpy.mock.instances[0].download).toBe('Excel Template.xls');
    expect(clickSpy.mock.instances[0].getAttribute('href')).toBe(
      '/Excel%20Template.xls',
    );
    expect(ElMessage.success).toHaveBeenCalledWith('Excel模板下载成功');
  });

  it('英文环境下载英文模板', () => {
    setLanguage('en-US');

    downloadExcelTemplate();

    expect(clickSpy.mock.instances[0].download).toBe(
      'Excel Template - (English).xls',
    );
    expect(ElMessage.success).toHaveBeenCalledWith(
      'Excel template downloaded successfully',
    );
  });

  it('下载链接在点击后从 DOM 中移除', () => {
    const appendSpy = vi.spyOn(document.body, 'appendChild');
    const removeSpy = vi.spyOn(document.body, 'removeChild');

    downloadExcelTemplate();

    expect(appendSpy).toHaveBeenCalledTimes(1);
    expect(removeSpy).toHaveBeenCalledTimes(1);
    expect(removeSpy.mock.calls[0][0]).toBe(appendSpy.mock.calls[0][0]);

    appendSpy.mockRestore();
    removeSpy.mockRestore();
  });

  it('下载过程出错时提示失败', () => {
    setLanguage('zh-CN');
    const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const createSpy = vi
      .spyOn(document, 'createElement')
      .mockImplementation(() => {
        throw new Error('boom');
      });

    downloadExcelTemplate();

    expect(ElMessage.error).toHaveBeenCalledWith('Excel模板下载出错');
    expect(errorSpy).toHaveBeenCalled();

    errorSpy.mockRestore();
    createSpy.mockRestore();
  });
});

describe('downloadWordTemplate', () => {
  it('中文环境下载中文模板', () => {
    setLanguage('zh-CN');

    downloadWordTemplate();

    expect(clickSpy.mock.instances[0].download).toBe('Word Template.docx');
    expect(ElMessage.success).toHaveBeenCalledWith('Word模板下载成功');
  });

  it('英文环境下载英文模板', () => {
    setLanguage('en-US');

    downloadWordTemplate();

    expect(clickSpy.mock.instances[0].download).toBe(
      'Word Template (English).docx',
    );
    expect(ElMessage.success).toHaveBeenCalledWith(
      'Word template downloaded successfully',
    );
  });

  it('下载过程出错时提示失败', () => {
    setLanguage('zh-CN');
    const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const createSpy = vi
      .spyOn(document, 'createElement')
      .mockImplementation(() => {
        throw new Error('boom');
      });

    downloadWordTemplate();

    expect(ElMessage.error).toHaveBeenCalledWith('Word模板下载出错');

    errorSpy.mockRestore();
    createSpy.mockRestore();
  });
});

describe('beforeExcelUpload', () => {
  it('接受 xlsx 与 xls', () => {
    expect(
      beforeExcelUpload({
        name: 'data.xlsx',
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
    ).toBe(true);
    expect(
      beforeExcelUpload({ name: 'data.xls', type: 'application/vnd.ms-excel' }),
    ).toBe(true);
    expect(ElMessage.error).not.toHaveBeenCalled();
  });

  it('MIME 缺失时按扩展名兜底放行', () => {
    expect(beforeExcelUpload({ name: 'data.xlsx', type: '' })).toBe(true);
  });

  it('非 Excel 文件被拒绝并提示', () => {
    expect(
      beforeExcelUpload({ name: 'doc.docx', type: 'application/pdf' }),
    ).toBe(false);
    expect(ElMessage.error).toHaveBeenCalledWith(
      'Please upload an Excel file in .xlsx or .xls format',
    );
  });
});

describe('beforeWordUpload', () => {
  it('接受 docx', () => {
    expect(
      beforeWordUpload({
        name: 'template.docx',
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    ).toBe(true);
    expect(ElMessage.error).not.toHaveBeenCalled();
  });

  it('MIME 缺失时按扩展名兜底放行', () => {
    expect(beforeWordUpload({ name: 'template.docx', type: '' })).toBe(true);
  });

  it('doc 与 xlsx 均被拒绝并提示', () => {
    expect(
      beforeWordUpload({ name: 'template.doc', type: 'application/msword' }),
    ).toBe(false);
    expect(beforeWordUpload({ name: 'data.xlsx', type: '' })).toBe(false);
    expect(ElMessage.error).toHaveBeenCalledWith(
      'Please upload a Word template file in .docx format',
    );
  });
});
