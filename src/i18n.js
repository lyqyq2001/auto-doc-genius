import { createI18n } from 'vue-i18n';
import zh from './locales/zh.json';
import en from './locales/en.json';

// 判断用户电脑的首选语言是否为中文，非中文一律按英文处理
export const isUserLanguageChinese = () => {
  const language = (
    navigator.language ||
    navigator.userLanguage ||
    'en'
  ).toLowerCase();
  return language.startsWith('zh');
};

export const i18n = createI18n({
  legacy: false,
  locale: isUserLanguageChinese() ? 'zh' : 'en',
  fallbackLocale: 'en',
  messages: { zh, en },
});

// 供非组件的 JS 模块使用
export const t = (key, params) => i18n.global.t(key, params);
