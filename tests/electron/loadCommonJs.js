import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import vm from 'node:vm';

const nodeRequire = createRequire(import.meta.url);

export const loadCommonJs = (filename, mocks = {}, globals = {}) => {
  const absolutePath = path.resolve(filename);
  const module = { exports: {} };
  const localRequire = id =>
    Object.hasOwn(mocks, id) ? mocks[id] : nodeRequire(id);
  const context = {
    require: localRequire,
    module,
    exports: module.exports,
    __filename: absolutePath,
    __dirname: path.dirname(absolutePath),
    process: globals.process || process,
    Buffer,
    console: globals.console || console,
    setTimeout,
    clearTimeout,
    ...globals,
  };

  vm.runInNewContext(readFileSync(absolutePath, 'utf8'), context, {
    filename: absolutePath,
  });
  return module.exports;
};
