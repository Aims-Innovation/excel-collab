import { TranslationKeys } from './locales';
import { LanguageType } from '../types';
declare const i18n: {
    changeLanguage: (lang: LanguageType) => void;
    init: () => void;
    t: (key: TranslationKeys, options?: Record<string, string | number>) => string;
    subscribe: (listener: () => void) => () => void;
    readonly current: "en-US" | "zh-CN" | "fr-FR" | "zh-HK" | "ja-JP" | "ko-KR" | "es-ES" | "de-DE" | "it-IT" | "pt-BR" | "pt-PT" | "nl-NL" | "ar-EG" | "ar-SA";
};
export default i18n;
