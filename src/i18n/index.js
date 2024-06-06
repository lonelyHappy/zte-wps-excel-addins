import zh from './zh';
import en from './en';

export default function getI18n(lang, id) {
  switch (lang) {
    case 'zh-ui':
      return zh[id] || '';
    case 'en-ui':
      return en[id] || '';
    default:
      return zh[id] || '';
  }
}