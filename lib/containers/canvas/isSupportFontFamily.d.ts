import { OptionItem } from '../../types';
declare const isSupportFontFamily: (font: string) => boolean;
export { isSupportFontFamily };
export declare function initFontFamilyList(check?: (font: string) => boolean, fontList?: string[]): OptionItem[];
