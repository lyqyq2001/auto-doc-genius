import { afterEach, describe, expect, it } from 'vitest';
import { i18n, isUserLanguageChinese, t } from '../../src/i18n';
import zh from '../../src/locales/zh.json';
import en from '../../src/locales/en.json';

const setNavigatorLanguage = language => {
  Object.defineProperty(navigator, 'language', {
    value: language,
    configurable: true,
  });
};

const originalLanguage = navigator.language;

afterEach(() => {
  setNavigatorLanguage(originalLanguage);
  i18n.global.locale.value = 'en';
});

// 递归收集对象的键路径，用于校验中英文语言包结构一致
const flattenKeys = (obj, prefix = '') =>
  Object.entries(obj).flatMap(([key, value]) => {
    const path = prefix ? `${prefix}.${key}` : key;
    return value && typeof value === 'object'
      ? flattenKeys(value, path)
      : [path];
  });

describe('isUserLanguageChinese', () => {
  it('中文环境返回 true', () => {
    setNavigatorLanguage('zh-CN');
    expect(isUserLanguageChinese()).toBe(true);
  });

  it('繁体中文环境返回 true', () => {
    setNavigatorLanguage('zh-TW');
    expect(isUserLanguageChinese()).toBe(true);
  });

  it('英文环境返回 false', () => {
    setNavigatorLanguage('en-US');
    expect(isUserLanguageChinese()).toBe(false);
  });
});

describe('t 翻译函数', () => {
  it('按当前 locale 返回对应文案', () => {
    i18n.global.locale.value = 'zh';
    expect(t('title')).toBe(zh.title);

    i18n.global.locale.value = 'en';
    expect(t('title')).toBe(en.title);
  });

  it('支持参数插值', () => {
    i18n.global.locale.value = 'en';
    expect(t('format.date', { day: '15', monthShort: 'Mar', year: 2023 })).toBe(
      '15-Mar-2023',
    );

    i18n.global.locale.value = 'zh';
    expect(t('format.date', { year: 2023, month: '03', day: '15' })).toBe(
      '2023年03月15日',
    );
  });
});

describe('语言包', () => {
  it('中英文语言包的键结构一致', () => {
    expect(flattenKeys(zh).sort()).toEqual(flattenKeys(en).sort());
  });
});
