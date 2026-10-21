import { afterEach, describe, expect, it, vi } from 'vitest';
const { message } = vi.hoisted(() => ({
  message: { success: vi.fn(), error: vi.fn() },
}));
vi.mock('element-plus', () => ({ ElMessage: message }));
import { i18n } from '../../src/i18n';
import {
  beforeExcelUpload,
  beforeWordUpload,
  downloadExcelTemplate,
  downloadWordTemplate,
} from '../../src/utils/fileUtils';

const originalLanguage = navigator.language;

beforeAll(() => {
  globalThis.ElMessage = message;
});

afterEach(() => {
  vi.clearAllMocks();
  vi.restoreAllMocks();
  Object.defineProperty(navigator, 'language', { value: originalLanguage, configurable: true });
  i18n.global.locale.value = 'en';
  document.body.innerHTML = '';
});

describe('upload validation', () => {
  it('accepts Excel MIME types and extensions, and rejects unrelated files', () => {
    expect(beforeExcelUpload({ type: 'application/vnd.ms-excel', name: 'data.bin' })).toBe(true);
    expect(beforeExcelUpload({ type: '', name: 'data.xlsx' })).toBe(true);
    expect(beforeExcelUpload({ type: '', name: 'data.xls' })).toBe(true);
    expect(beforeExcelUpload({ type: 'text/plain', name: 'data.txt' })).toBe(false);
    expect(message.error).toHaveBeenCalledOnce();
  });

  it('accepts DOCX MIME type or extension, and rejects other documents', () => {
    expect(beforeWordUpload({ type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', name: 'data.bin' })).toBe(true);
    expect(beforeWordUpload({ type: '', name: 'template.docx' })).toBe(true);
    expect(beforeWordUpload({ type: 'application/pdf', name: 'template.pdf' })).toBe(false);
    expect(message.error).toHaveBeenCalledOnce();
  });
});

describe('template downloads', () => {
  const clickSpy = () => vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});

  it('downloads the Chinese Excel template for Chinese system language', () => {
    Object.defineProperty(navigator, 'language', { value: 'zh-CN', configurable: true });
    const click = clickSpy();
    downloadExcelTemplate();
    const anchor = document.querySelector('a');
    expect(anchor).toBeNull();
    expect(click).toHaveBeenCalledOnce();
    expect(message.success).toHaveBeenCalledOnce();
  });

  it('downloads the English Word template for non-Chinese system language', () => {
    Object.defineProperty(navigator, 'language', { value: 'fr-FR', configurable: true });
    const click = clickSpy();
    downloadWordTemplate();
    expect(click).toHaveBeenCalledOnce();
    expect(message.success).toHaveBeenCalledOnce();
  });

  it('reports download errors when browser download setup throws', () => {
    vi.spyOn(document, 'createElement').mockImplementation(() => { throw new Error('blocked'); });
    downloadExcelTemplate();
    expect(message.error).toHaveBeenCalledOnce();
  });
});
