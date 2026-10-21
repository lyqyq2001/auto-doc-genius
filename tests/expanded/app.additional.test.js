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
  state.handleExcelChange({ raw: new File(['excel'], 'data.xlsx') });
  state.handleWordChange({ raw: new File(['word'], 'template.docx') });
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
    expect(mocks.message.error.mock.calls[0][0]).toContain('invalid workbook');
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
    const removeProgressListener = vi.fn();
    window.electronAPI = {
      onPdfProgress: vi.fn(() => removeProgressListener),
      convertBatchToPdf: vi.fn().mockResolvedValue({
        success: false,
        error: 'Office conversion failed',
      }),
    };
    const wrapper = await mountApp();
    appState(wrapper).exportFormat = 'pdf';
    appState(wrapper).officeInstalled = true;
    uploadBothFiles(wrapper);
    await appState(wrapper).startGenerate();

    expect(mocks.message.error.mock.calls[0][0]).toContain('Office conversion failed');
    expect(removeProgressListener).toHaveBeenCalledOnce();
    wrapper.unmount();
  }, 20000);
});
