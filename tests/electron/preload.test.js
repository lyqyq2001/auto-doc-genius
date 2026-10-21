// @vitest-environment node
import { describe, expect, it, vi } from 'vitest';
import path from 'node:path';
import { loadCommonJs } from './loadCommonJs';

describe('Electron preload bridge', () => {
  it('exposes only the supported IPC methods and removes progress listeners', async () => {
    const exposed = {};
    const listeners = new Map();
    const ipcRenderer = {
      invoke: vi.fn((...args) => Promise.resolve(args)),
      on: vi.fn((channel, listener) => listeners.set(channel, listener)),
      removeListener: vi.fn((channel, listener) => listeners.delete(channel)),
    };
    const webUtils = { getPathForFile: vi.fn(file => `C:\\files\\${file.name}`) };
    loadCommonJs(path.resolve('electron/preload.js'), {
      electron: {
        contextBridge: {
          exposeInMainWorld: (name, api) => {
            exposed[name] = api;
          },
        },
        ipcRenderer,
        webUtils,
      },
    });

    const api = exposed.electronAPI;
    expect(Object.keys(api).sort()).toEqual([
      'checkOfficeInstallation',
      'convertBatchToPdf',
      'getFilePath',
      'onPdfProgress',
      'readImageFiles',
    ]);
    await expect(api.convertBatchToPdf([{ name: 'a.docx' }])).resolves.toEqual([
      'batch-convert-pdf',
      [{ name: 'a.docx' }],
    ]);
    await expect(api.checkOfficeInstallation()).resolves.toEqual(['check-office-installation']);
    await expect(api.readImageFiles(['a.png'])).resolves.toEqual([
      'read-image-files',
      ['a.png'],
    ]);
    expect(api.getFilePath({ name: 'data.xlsx' })).toBe('C:\\files\\data.xlsx');

    const callback = vi.fn();
    const unsubscribe = api.onPdfProgress(callback);
    const listener = listeners.get('pdf-conversion-progress');
    const payload = { stage: 'converting', progress: 30 };
    listener({ sender: 'ignored' }, payload);
    expect(callback).toHaveBeenCalledWith(payload);
    unsubscribe();
    expect(ipcRenderer.removeListener).toHaveBeenCalledWith(
      'pdf-conversion-progress',
      listener,
    );
    expect(listeners.has('pdf-conversion-progress')).toBe(false);
  });
});
