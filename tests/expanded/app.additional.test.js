import { beforeEach, describe, expect, it, vi } from 'vitest';
import { defineComponent } from 'vue';
import { shallowMount } from '@vue/test-utils';
import { i18n } from '../../src/i18n';

const mocks = vi.hoisted(() => ({
  message: { success: vi.fn(), error: vi.fn(), warning: vi.fn() },
  notification: { success: vi.fn(), warning: vi.fn() },
  saveAs: vi.fn(),
  zipEntries: [],
  parseExcel: vi.fn(),
  readWordTemplate: vi.fn(),
  generateDocx: vi.fn(),
  resolveDuplicateName: vi.fn((name, used) => {
    const extIndex = name.lastIndexOf('.');
    const base = extIndex < 0 ? name : name.slice(0, extIndex);
    const ext = extIndex < 0 ? '' : name.slice(extIndex);
    let candidate = name;
    let suffix = 2;
    while (used.has(candidate)) candidate = `${base}_${suffix++}${ext}`;
    used.add(candidate);
    return candidate;
  }),
  downloadExcelTemplate: vi.fn(),
  downloadWordTemplate: vi.fn(),
  beforeExcelUpload: vi.fn(() => true),
  beforeWordUpload: vi.fn(() => true),
  isImageColumn: vi.fn(key => key.toLowerCase().startsWith('img')),
  collectImagePaths: vi.fn(() => []),
}));

vi.mock('element-plus', async importOriginal => ({
  ...(await importOriginal()),
  ElMessage: mocks.message,
  ElNotification: mocks.notification,
}));
vi.mock('file-saver', () => ({ saveAs: mocks.saveAs }));
vi.mock('jszip', () => ({
  default: class MockZip {
    file = vi.fn((name, data) => mocks.zipEntries.push({ name, data }));
    generateAsync = vi.fn().mockResolvedValue(new Blob(['zip']));
  },
}));
vi.mock('../../src/utils/fileUtils', () => ({
  downloadExcelTemplate: mocks.downloadExcelTemplate,
  downloadWordTemplate: mocks.downloadWordTemplate,
  beforeExcelUpload: mocks.beforeExcelUpload,
  beforeWordUpload: mocks.beforeWordUpload,
}));
vi.mock('../../src/utils/documentUtils', () => ({
  parseExcel: mocks.parseExcel,
  readWordTemplate: mocks.readWordTemplate,
}));
vi.mock('../../src/utils/docxGenerator', () => ({
  generateDocx: mocks.generateDocx,
  resolveDuplicateName: mocks.resolveDuplicateName,
}));
vi.mock('../../src/utils/imageUtils', () => ({
  isImageColumn: mocks.isImageColumn,
  collectImagePaths: mocks.collectImagePaths,
}));

const UploadStub = defineComponent({
  name: 'ElUpload',
  props: { onChange: Function },
  template: '<div class="upload-stub"></div>',
});

const CardStub = defineComponent({
  name: 'ElCard',
  template: '<section><slot name="header" /><slot /><slot name="footer" /></section>',
});

const ButtonStub = defineComponent({
  name: 'ElButton',
  props: { disabled: Boolean, loading: Boolean },
  emits: ['click'],
  template: '<button :disabled="disabled" @click="$emit(\'click\')"><slot /></button>',
});

const mountApp = async () => {
  const { default: App } = await import('../../src/App.vue');
  return shallowMount(App, {
    global: {
      plugins: [i18n],
      stubs: {
        ElCard: CardStub,
        ElUpload: UploadStub,
        ElIcon: true,
        ElButton: ButtonStub,
        ElRadioGroup: true,
        ElRadioButton: true,
        ElProgress: true,
        UploadFilled: true,
      },
    },
  });
};

const uploadBothFiles = wrapper => {
  const state = wrapper.vm.$.setupState;
  const excelFile = new File(['excel'], 'data.xlsx');
  state.handleExcelChange({ raw: excelFile });
  state.handleWordChange({ raw: new File(['word'], 'template.docx') });
  return excelFile;
};

const appState = wrapper => wrapper.vm.$.setupState;

beforeEach(() => {
  vi.clearAllMocks();
  mocks.zipEntries.length = 0;
  i18n.global.locale.value = 'en';
  delete window.electronAPI;
  mocks.parseExcel.mockResolvedValue([{ Name: 'Alex' }, { Name: 'Alex' }]);
  mocks.readWordTemplate.mockResolvedValue(new ArrayBuffer(4));
  mocks.generateDocx
    .mockReturnValueOnce({ name: 'Alex.docx', buffer: new ArrayBuffer(1) })
    .mockReturnValueOnce({ name: 'Alex.docx', buffer: new ArrayBuffer(2) });
});

describe('App document generation flow', () => {
  it('keeps generation disabled until both files are selected', async () => {
    const wrapper = await mountApp();
    expect(appState(wrapper).disabled).toBe(true);

    uploadBothFiles(wrapper);
    await wrapper.vm.$nextTick();
    expect(appState(wrapper).disabled).toBe(false);
    wrapper.unmount();
  }, 20000);

  it('generates all DOCX files, resolves duplicate names, and downloads a ZIP', async () => {
    const wrapper = await mountApp();
    uploadBothFiles(wrapper);
    await wrapper.vm.$nextTick();
    await appState(wrapper).startGenerate();
    await vi.waitFor(() => expect(mocks.saveAs).toHaveBeenCalledOnce());

    expect(mocks.parseExcel).toHaveBeenCalledOnce();
    expect(mocks.readWordTemplate).toHaveBeenCalledOnce();
    expect(mocks.generateDocx).toHaveBeenCalledTimes(2);
    expect(mocks.resolveDuplicateName).toHaveBeenCalledTimes(2);
    expect(mocks.saveAs.mock.calls[0][1]).toBe('Generated Word Files.zip');
    expect(mocks.notification.warning).toHaveBeenCalledOnce();
    expect(mocks.notification.success).toHaveBeenCalledOnce();
    expect(appState(wrapper).disabled).toBe(false);
    wrapper.unmount();
  }, 20000);

  it('shows the failure message and restores the generate button after an error', async () => {
    mocks.parseExcel.mockRejectedValueOnce(new Error('invalid workbook'));
    const wrapper = await mountApp();
    uploadBothFiles(wrapper);
    await wrapper.vm.$nextTick();
    await appState(wrapper).startGenerate();

    await vi.waitFor(() => expect(mocks.message.error).toHaveBeenCalledOnce());
    expect(mocks.message.error.mock.calls[0][0]).toContain(
      i18n.global.t('generate.unexpectedFailed'),
    );
    expect(appState(wrapper).disabled).toBe(false);
    wrapper.unmount();
  }, 20000);

  it('converts generated Word files to PDF, packages the results, and removes the progress listener', async () => {
    const removeProgressListener = vi.fn();
    const convertBatchToPdf = vi.fn().mockResolvedValue({
      success: true,
      results: [{ name: 'Alex.pdf', data: new Uint8Array([1, 2, 3]) }],
    });
    const onPdfProgress = vi.fn(() => removeProgressListener);
    window.electronAPI = { convertBatchToPdf, onPdfProgress };

    const wrapper = await mountApp();
    appState(wrapper).exportFormat = 'pdf';
    appState(wrapper).officeInstalled = true;
    uploadBothFiles(wrapper);
    await appState(wrapper).startGenerate();

    expect(convertBatchToPdf).toHaveBeenCalledOnce();
    expect(convertBatchToPdf.mock.calls[0][0]).toHaveLength(2);
    expect(mocks.zipEntries).toEqual([
      { name: 'Alex.pdf', data: new Uint8Array([1, 2, 3]) },
    ]);
    expect(mocks.saveAs.mock.calls[0][1]).toBe('Generated PDF Files.zip');
    expect(removeProgressListener).toHaveBeenCalledOnce();
    expect(mocks.notification.success).toHaveBeenCalledOnce();
    wrapper.unmount();
  }, 20000);

  it('reports a PDF conversion failure and unregisters progress notifications', async () => {
    i18n.global.locale.value = 'zh';
    const removeProgressListener = vi.fn();
    window.electronAPI = {
      onPdfProgress: vi.fn(() => removeProgressListener),
      convertBatchToPdf: vi.fn().mockResolvedValue({
        success: false,
        errorCode: 'officeUnavailable',
      }),
    };
    const wrapper = await mountApp();
    appState(wrapper).exportFormat = 'pdf';
    appState(wrapper).officeInstalled = true;
    uploadBothFiles(wrapper);
    await appState(wrapper).startGenerate();

    expect(mocks.message.error.mock.calls[0][0]).toBe(
      i18n.global.t('generate.failed', {
        message: i18n.global.t('generate.officeUnavailable'),
      }),
    );
    expect(removeProgressListener).toHaveBeenCalledOnce();
    wrapper.unmount();
  }, 20000);

  it('splits 11 PDF records into batches of 10 and 1 and applies conversion progress', async () => {
    const rows = Array.from({ length: 11 }, (_, index) => ({
      Name: `Person${index + 1}`,
    }));
    mocks.parseExcel.mockResolvedValueOnce(rows);
    mocks.generateDocx.mockReset().mockImplementation((_template, row, index) => ({
      name: `${row.Name}.docx`,
      buffer: new Uint8Array([index]),
    }));

    let progressCallback;
    let progressDuringConversion;
    let textDuringConversion;
    let wrapper;
    const convertBatchToPdf = vi.fn(async batch => {
      progressCallback({
        stage: 'converting',
        progress: 25,
        message: 'Converting batch',
      });
      progressDuringConversion = appState(wrapper).progress;
      textDuringConversion = appState(wrapper).progressText;
      return {
        success: true,
        results: batch.map(file => ({
          name: file.name.replace(/\.docx$/i, '.pdf'),
          data: new Uint8Array([1]),
        })),
      };
    });
    window.electronAPI = {
      onPdfProgress: vi.fn(callback => {
        progressCallback = callback;
        return vi.fn();
      }),
      convertBatchToPdf,
    };

    wrapper = await mountApp();
    appState(wrapper).exportFormat = 'pdf';
    appState(wrapper).officeInstalled = true;
    uploadBothFiles(wrapper);
    await appState(wrapper).startGenerate();

    expect(convertBatchToPdf.mock.calls.map(([batch]) => batch.length)).toEqual([10, 1]);
    expect(convertBatchToPdf.mock.calls[0][0][0].name).toBe('Person1.docx');
    expect(convertBatchToPdf.mock.calls[1][0][0].name).toBe('Person11.docx');
    expect(progressDuringConversion).toBe(60);
    expect(textDuringConversion).toBe(i18n.global.t('progress.convertingPdf'));
    expect(mocks.zipEntries).toHaveLength(11);
    expect(mocks.saveAs.mock.calls[0][1]).toBe('Generated PDF Files.zip');
    wrapper.unmount();
  }, 20000);

  it('warns when image columns are present outside Electron and still generates documents', async () => {
    mocks.parseExcel.mockResolvedValueOnce([{ Name: 'Alex', imgPhoto: 'photo.png' }]);
    mocks.isImageColumn.mockReturnValue(true);
    const wrapper = await mountApp();
    uploadBothFiles(wrapper);
    await appState(wrapper).startGenerate();

    expect(mocks.message.warning).toHaveBeenCalledWith(
      i18n.global.t('image.noElectron'),
    );
    expect(mocks.generateDocx.mock.calls[0][3]).toBeNull();
    expect(mocks.saveAs).toHaveBeenCalledOnce();
    wrapper.unmount();
  }, 20000);

  it('reads referenced images, warns about missing assets, and continues DOCX generation', async () => {
    const rows = [{ Name: 'Alex', imgPhoto: 'photo.png' }];
    const imagePaths = ['C:\\data\\photo.png', 'C:\\data\\missing.png'];
    const imageAssets = {
      [imagePaths[0]]: { ok: true, base64: 'encoded', width: 20, height: 10, ext: 'png' },
      [imagePaths[1]]: { ok: false, error: 'not found' },
    };
    mocks.parseExcel.mockResolvedValueOnce(rows);
    mocks.isImageColumn.mockReturnValue(true);
    mocks.collectImagePaths.mockReturnValueOnce(imagePaths);
    const getFilePath = vi.fn(() => 'C:\\data\\records.xlsx');
    const readImageFiles = vi.fn().mockResolvedValue(imageAssets);
    window.electronAPI = { getFilePath, readImageFiles };

    const wrapper = await mountApp();
    const excelFile = uploadBothFiles(wrapper);
    await appState(wrapper).startGenerate();

    expect(getFilePath).toHaveBeenCalledWith(excelFile);
    expect(mocks.collectImagePaths).toHaveBeenCalledWith(rows, 'C:\\data');
    expect(readImageFiles).toHaveBeenCalledWith(imagePaths);
    expect(mocks.notification.warning).toHaveBeenCalledWith(
      expect.objectContaining({
        title: i18n.global.t('image.missingTitle'),
        message: i18n.global.t('image.missingMessage', { count: 1 }),
      }),
    );
    expect(mocks.generateDocx.mock.calls[0][3]).toEqual(imageAssets);
    expect(mocks.generateDocx.mock.calls[0][4]).toBe('C:\\data');
    expect(mocks.saveAs).toHaveBeenCalledOnce();
    wrapper.unmount();
  }, 20000);

  it('reports image read failures and resets the generating state', async () => {
    mocks.parseExcel.mockResolvedValueOnce([{ Name: 'Alex', imgPhoto: 'photo.png' }]);
    mocks.isImageColumn.mockReturnValue(true);
    mocks.collectImagePaths.mockReturnValueOnce(['C:\\data\\photo.png']);
    window.electronAPI = {
      getFilePath: vi.fn(() => 'C:\\data\\records.xlsx'),
      readImageFiles: vi.fn().mockRejectedValue(new Error('disk read failed')),
    };
    const wrapper = await mountApp();
    uploadBothFiles(wrapper);
    await appState(wrapper).startGenerate();

    expect(mocks.message.error.mock.calls[0][0]).toContain(
      i18n.global.t('generate.unexpectedFailed'),
    );
    expect(mocks.message.error.mock.calls[0][0]).not.toContain('disk read failed');
    expect(mocks.generateDocx).not.toHaveBeenCalled();
    expect(appState(wrapper).generating).toBe(false);
    wrapper.unmount();
  }, 20000);
});
