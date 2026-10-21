import { mount } from '@vue/test-utils';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import App from '../../src/App.vue';
import { i18n } from '../../src/i18n';

// App.vue 中的 ElMessage/ElNotification 为命名导入，这里直接 mock 掉 element-plus，
// 同时避免 element-plus 入口 import .css 在 Node 端加载失败
vi.mock('element-plus', () => {
  const stub = (name = 'div') => ({ template: `<${name}><slot /></${name}>` });
  const messageMock = {
    success: vi.fn(),
    error: vi.fn(),
    warning: vi.fn(),
  };
  const notifyMock = {
    success: vi.fn(),
    error: vi.fn(),
    warning: vi.fn(),
  };
  return {
    ElMessage: messageMock,
    ElNotification: notifyMock,
    // 模板编译产物引用这些命名导出，提供最小 stub
    ElIcon: stub('i'),
    ElCard: {
      template: '<div><slot name="header" /><slot /><slot name="footer" /></div>',
    },
    ElUpload: stub(),
    ElButton: stub('button'),
    ElRadioGroup: stub(),
    ElRadioButton: stub(),
    ElProgress: stub(),
  };
});

// JSZip 体积较大且与断言无关，仅提供生成链路所需的最小实现
vi.mock('jszip', () => ({
  default: class {
    constructor() {
      this.files = new Map();
    }
    file(name, data) {
      this.files.set(name, data);
    }
    async generateAsync() {
      return new Blob([]);
    }
  },
}));

import { ElMessage } from 'element-plus';

const uploadStub = {
  template: '<div class="upload-stub"><slot /><slot name="tip" /></div>',
};

const mountApp = () =>
  mount(App, {
    global: {
      plugins: [i18n],
      stubs: {
        // 覆盖 mock 中提供的占位实现，使用可交互的轻量 stub
        ElCard: {
          template: '<div><slot name="header" /><slot /><slot name="footer" /></div>',
        },
        ElUpload: uploadStub,
        ElIcon: { template: '<i><slot /></i>' },
        ElButton: {
          template:
            '<button class="btn" :disabled="disabled" @click="$emit(\'click\')"><slot /></button>',
          props: ['disabled', 'loading', 'type', 'size'],
        },
        ElRadioGroup: { template: '<div><slot /></div>' },
        ElRadioButton: {
          template: '<span class="radio-btn"><slot /></span>',
          props: ['label'],
        },
        ElProgress: {
          template: '<div />',
          props: ['percentage', 'status', 'strokeWidth'],
        },
        UploadFilled: { template: '<i />' },
      },
    },
  });

const appState = wrapper => wrapper.vm.$.setupState;

beforeEach(() => {
  vi.clearAllMocks();
  i18n.global.locale.value = 'en';
  delete window.electronAPI;
});

describe('App 组件渲染', () => {
  it('正常挂载并渲染标题与上传区域', () => {
    const wrapper = mountApp();

    expect(wrapper.text()).toContain('Excel to Report Batch Generator');
    expect(wrapper.text()).toContain('Upload Excel Data Sheet');
    expect(wrapper.text()).toContain('Upload Word Template');
    expect(wrapper.text()).toContain('Start Generating Reports');
  });

  it('初始状态生成按钮不可用（未上传文件）', () => {
    const wrapper = mountApp();
    expect(appState(wrapper).disabled).toBe(true);
  });

  it('上传 Excel 与 Word 后按钮恢复可用', async () => {
    const wrapper = mountApp();

    // 模拟 el-upload 的 on-change 回调
    appState(wrapper).handleExcelChange({ raw: new File(['x'], 'data.xlsx') });
    await wrapper.vm.$nextTick();
    expect(appState(wrapper).disabled).toBe(true);

    appState(wrapper).handleWordChange({ raw: new File(['x'], 'template.docx') });
    await wrapper.vm.$nextTick();
    expect(appState(wrapper).disabled).toBe(false);
  });

  it('未上传文件就点击生成时给出警告且不进入生成态', async () => {
    const wrapper = mountApp();

    await appState(wrapper).startGenerate();

    expect(ElMessage.warning).toHaveBeenCalledWith(
      'Please upload the Excel file and Word template file first',
    );
    expect(appState(wrapper).generating).toBe(false);
  });

  it('选择 PDF 导出但未检测到 Office 时禁用生成按钮', async () => {
    const wrapper = mountApp();

    appState(wrapper).handleExcelChange({ raw: new File(['x'], 'data.xlsx') });
    appState(wrapper).handleWordChange({ raw: new File(['x'], 'template.docx') });
    appState(wrapper).exportFormat = 'pdf';
    appState(wrapper).officeInstalled = false;
    await wrapper.vm.$nextTick();

    expect(appState(wrapper).disabled).toBe(true);
  });

  it('Office 已安装时 PDF 模式按钮可用', async () => {
    const wrapper = mountApp();

    appState(wrapper).handleExcelChange({ raw: new File(['x'], 'data.xlsx') });
    appState(wrapper).handleWordChange({ raw: new File(['x'], 'template.docx') });
    appState(wrapper).exportFormat = 'pdf';
    appState(wrapper).officeInstalled = true;
    await wrapper.vm.$nextTick();

    expect(appState(wrapper).disabled).toBe(false);
  });

  it('存在 electronAPI 时检查 Office 安装状态', async () => {
    window.electronAPI = {
      checkOfficeInstallation: vi.fn().mockResolvedValue(true),
    };
    const wrapper = mountApp();
    // 等待组件 setup 中触发的异步检查完成
    await vi.waitFor(() => {
      expect(appState(wrapper).officeInstalled).toBe(true);
    });
  });
});
