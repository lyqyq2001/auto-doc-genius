// @vitest-environment node
import { afterEach, describe, expect, it, vi } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { loadCommonJs } from './loadCommonJs';

const mainPath = path.resolve('electron/main.js');
const tempDirs = new Set();

const createMainHarness = ({
  checkWordInstallation = vi.fn().mockResolvedValue(true),
  convertBatchWordToPdf = vi.fn().mockResolvedValue({ success: true, results: [] }),
  imageSize = vi.fn(() => ({ width: 4, height: 3 })),
} = {}) => {
  const handlers = new Map();
  const electron = {
    app: {
      whenReady: () => ({ then: vi.fn() }),
      on: vi.fn(),
      quit: vi.fn(),
    },
    BrowserWindow: class {
      loadURL = vi.fn();
      loadFile = vi.fn();
      static getAllWindows = () => [];
    },
    ipcMain: {
      handle: vi.fn((channel, handler) => handlers.set(channel, handler)),
    },
  };
  const module = loadCommonJs(mainPath, {
    electron,
    'image-size': { imageSize },
    './officeConverter': { checkWordInstallation, convertBatchWordToPdf },
  }, {
    process: { env: {}, platform: 'win32' },
    console: { log: vi.fn(), error: vi.fn(), warn: vi.fn() },
  });

  return { handlers, electron, checkWordInstallation, convertBatchWordToPdf, imageSize, module };
};

const makeTempDir = () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'autodoc-test-'));
  tempDirs.add(dir);
  return dir;
};

afterEach(() => {
  for (const dir of tempDirs) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
  tempDirs.clear();
  vi.restoreAllMocks();
});

describe('Electron main IPC handlers', () => {
  it('registers PDF conversion, Office detection, and image reading IPC handlers', () => {
    const { handlers } = createMainHarness();
    expect([...handlers.keys()]).toEqual([
      'batch-convert-pdf',
      'check-office-installation',
      'read-image-files',
    ]);
  });

  it('delegates Office installation checks to the converter module', async () => {
    const { handlers, checkWordInstallation } = createMainHarness();
    await expect(handlers.get('check-office-installation')()).resolves.toBe(true);
    expect(checkWordInstallation).toHaveBeenCalledOnce();
  });

  it('reads allowed image files and returns failures for unsupported or missing files', () => {
    const dir = makeTempDir();
    const pngPath = path.join(dir, 'photo.png');
    const txtPath = path.join(dir, 'notes.txt');
    const missingPath = path.join(dir, 'missing.jpg');
    fs.writeFileSync(pngPath, Buffer.from([1, 2, 3]));
    fs.writeFileSync(txtPath, 'text');
    const imageSize = vi.fn(() => ({ width: 640, height: 480 }));
    const { handlers } = createMainHarness({ imageSize });

    const result = handlers.get('read-image-files')(null, [pngPath, txtPath, missingPath]);

    expect(result[pngPath]).toEqual({
      ok: true,
      base64: Buffer.from([1, 2, 3]).toString('base64'),
      width: 640,
      height: 480,
      ext: 'png',
    });
    expect(result[txtPath]).toMatchObject({ ok: false });
    expect(result[txtPath].error).toContain('不支持的图片格式');
    expect(result[missingPath]).toMatchObject({ ok: false });
    expect(result[missingPath].error).toContain('ENOENT');
    expect(imageSize).toHaveBeenCalledOnce();
  });

  it('uses fallback dimensions when image metadata cannot be decoded', () => {
    const dir = makeTempDir();
    const imagePath = path.join(dir, 'broken.png');
    fs.writeFileSync(imagePath, Buffer.from([1]));
    const imageSize = vi.fn(() => {
      throw new Error('invalid image');
    });
    const { handlers } = createMainHarness({ imageSize });
    const result = handlers.get('read-image-files')(null, [imagePath]);

    expect(result[imagePath]).toMatchObject({ ok: true, width: 200, height: 200 });
  });

  it('writes DOCX inputs, reports conversion progress, and removes its temporary folder', async () => {
    let inputPairs;
    const converter = vi.fn(async pairs => {
      inputPairs = pairs;
      expect(fs.readFileSync(pairs[0].input)).toEqual(Buffer.from([5, 6]));
      expect(path.dirname(pairs[0].input)).toBe(path.dirname(pairs[0].output));
      return { success: true, results: [{ name: 'result.pdf', data: Buffer.from('pdf') }] };
    });
    const { handlers } = createMainHarness({ convertBatchWordToPdf: converter });
    const sent = [];
    const event = { sender: { send: (channel, data) => sent.push({ channel, data }) } };

    const result = await handlers.get('batch-convert-pdf')(event, [
      { name: 'report.docx', buffer: new Uint8Array([5, 6]) },
    ]);

    expect(result).toMatchObject({ success: true });
    expect(result.results[0].name).toBe('result.pdf');
    expect(converter).toHaveBeenCalledOnce();
    expect(sent.map(item => item.data.stage)).toEqual(['init', 'converting', 'completed']);
    expect(sent.every(item => item.channel === 'pdf-conversion-progress')).toBe(true);
    expect(fs.existsSync(path.dirname(inputPairs[0].input))).toBe(false);
  });

  it('returns conversion errors and removes temporary DOCX files after a thrown failure', async () => {
    let inputPairs;
    const converter = vi.fn(async pairs => {
      inputPairs = pairs;
      throw new Error('converter unavailable');
    });
    const { handlers } = createMainHarness({ convertBatchWordToPdf: converter });
    const sent = [];
    const event = { sender: { send: (_channel, data) => sent.push(data) } };

    await expect(
      handlers.get('batch-convert-pdf')(event, [
        { name: 'report.docx', buffer: new Uint8Array([1]) },
      ]),
    ).resolves.toEqual({ success: false, error: 'converter unavailable' });

    expect(fs.existsSync(path.dirname(inputPairs[0].input))).toBe(false);
    expect(sent.at(-1)).toMatchObject({ stage: 'error' });
  });

  it('returns a converter-reported failure and cleans up temporary inputs', async () => {
    let inputPairs;
    const converter = vi.fn(async pairs => {
      inputPairs = pairs;
      return { success: false, error: 'Office is unavailable' };
    });
    const { handlers } = createMainHarness({ convertBatchWordToPdf: converter });
    const sent = [];
    const event = { sender: { send: (_channel, data) => sent.push(data) } };

    await expect(
      handlers.get('batch-convert-pdf')(event, [
        { name: 'report.docx', buffer: new Uint8Array([1]) },
      ]),
    ).resolves.toEqual({ success: false, error: 'Office is unavailable' });

    expect(fs.existsSync(path.dirname(inputPairs[0].input))).toBe(false);
    expect(sent.at(-1)).toMatchObject({ stage: 'error' });
  });
});
