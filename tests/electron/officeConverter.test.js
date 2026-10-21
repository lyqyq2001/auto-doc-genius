// @vitest-environment node
import { afterEach, describe, expect, it, vi } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { loadCommonJs } from './loadCommonJs';

const officeConverterPath = path.resolve('electron/officeConverter.js');
const tempDirs = new Set();

const makeTempDir = () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'office-converter-test-'));
  tempDirs.add(dir);
  return dir;
};

const loadOfficeConverter = exec =>
  loadCommonJs(officeConverterPath, {
    'child_process': { exec },
    path,
    fs,
    os,
  }, {
    console: {
      log: vi.fn(),
      error: vi.fn(),
    },
  });

afterEach(() => {
  for (const dir of tempDirs) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
  tempDirs.clear();
  vi.restoreAllMocks();
});

describe('checkWordInstallation', () => {
  it('returns true when Word or WPS automation launches successfully', async () => {
    const exec = vi.fn((_command, _options, callback) => callback(null, '', ''));
    const { checkWordInstallation } = loadOfficeConverter(exec);

    await expect(checkWordInstallation()).resolves.toBe(true);
    expect(exec).toHaveBeenCalledWith(
      expect.stringContaining('powershell -Command'),
      expect.objectContaining({ timeout: 10000 }),
      expect.any(Function),
    );
  });

  it('returns false when Office automation cannot launch', async () => {
    const exec = vi.fn((_command, _options, callback) =>
      callback(new Error('not installed'), '', ''),
    );
    const { checkWordInstallation } = loadOfficeConverter(exec);

    await expect(checkWordInstallation()).resolves.toBe(false);
  });
});

describe('convertBatchWordToPdf', () => {
  it('returns an empty success without invoking PowerShell for an empty batch', async () => {
    const exec = vi.fn();
    const { convertBatchWordToPdf } = loadOfficeConverter(exec);

    await expect(convertBatchWordToPdf([])).resolves.toEqual({
      success: true,
      results: [],
    });
    expect(exec).not.toHaveBeenCalled();
  });

  it('writes a conversion script, reads produced PDFs, and cleans up its script directory', async () => {
    const outputDir = makeTempDir();
    const pairs = [
      { input: 'C:\\source\\first.docx', output: path.join(outputDir, 'first.pdf') },
    ];
    let scriptPath;
    const exec = vi.fn((command, options, callback) => {
      scriptPath = command.match(/-File "([^"]+)"/)[1];
      expect(fs.existsSync(scriptPath)).toBe(true);
      expect(fs.readFileSync(scriptPath, 'utf8')).toContain('C:\\source\\first.docx');
      expect(options).toMatchObject({ timeout: 300000, windowsHide: true });
      fs.writeFileSync(pairs[0].output, Buffer.from('pdf data'));
      callback(null, 'BATCH_CONVERSION_SUCCESS', '');
    });
    const { convertBatchWordToPdf } = loadOfficeConverter(exec);

    await expect(convertBatchWordToPdf(pairs)).resolves.toEqual({
      success: true,
      results: [{ name: 'first.pdf', data: Buffer.from('pdf data') }],
    });
    expect(exec).toHaveBeenCalledOnce();
    expect(fs.existsSync(path.dirname(scriptPath))).toBe(false);
  });

  it('returns a failed result and cleans up after the PowerShell process errors', async () => {
    const pairs = [{ input: 'input.docx', output: path.join(makeTempDir(), 'output.pdf') }];
    let scriptPath;
    const exec = vi.fn((command, _options, callback) => {
      scriptPath = command.match(/-File "([^"]+)"/)[1];
      callback(new Error('PowerShell failed'), '', '');
    });
    const { convertBatchWordToPdf } = loadOfficeConverter(exec);

    await expect(convertBatchWordToPdf(pairs)).resolves.toMatchObject({
      success: false,
      errorCode: 'pdfConversionFailed',
      error: 'PowerShell failed',
    });
    expect(fs.existsSync(path.dirname(scriptPath))).toBe(false);
  });

  it('returns only the PDFs that exist after a partial conversion', async () => {
    const outputDir = makeTempDir();
    const pairs = [
      { input: 'first.docx', output: path.join(outputDir, 'first.pdf') },
      { input: 'second.docx', output: path.join(outputDir, 'second.pdf') },
    ];
    const exec = vi.fn((_command, _options, callback) => {
      fs.writeFileSync(pairs[0].output, Buffer.from('first pdf'));
      callback(null, 'BATCH_CONVERSION_PARTIAL_SUCCESS', '');
    });
    const { convertBatchWordToPdf } = loadOfficeConverter(exec);

    await expect(convertBatchWordToPdf(pairs)).resolves.toEqual({
      success: true,
      results: [{ name: 'first.pdf', data: Buffer.from('first pdf') }],
    });
  });

  it('returns the stable officeUnavailable code when PowerShell reports no Office installation', async () => {
    const pair = {
      input: 'input.docx',
      output: path.join(makeTempDir(), 'output.pdf'),
    };
    const exec = vi.fn((_command, _options, callback) =>
      callback(null, 'BATCH_CONVERSION_PARTIAL_SUCCESS\nOFFICE_UNAVAILABLE', ''),
    );
    const { convertBatchWordToPdf } = loadOfficeConverter(exec);

    await expect(convertBatchWordToPdf([pair])).resolves.toMatchObject({
      success: false,
      errorCode: 'officeUnavailable',
    });
  });
});
