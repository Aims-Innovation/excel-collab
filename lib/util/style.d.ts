import { StyleType, ResultType } from '../types';
export declare const FONT_SIZE_LIST: number[];
export declare function makeFont(fontStyle?: 'none' | 'normal' | 'italic' | 'oblique', fontWeight?: string, fontSize?: number, fontFamily?: string): string;
export declare function convertToCssString(style: Partial<StyleType>): string;
export declare function convertToPx(value: string): number;
export declare function convertPxToPt(px: number, other?: string): string;
export declare function parseText(text: string, splitter?: string): ResultType[][];
export declare function parseHTML(html: string): {
    textList: ResultType[][];
    styleList: Partial<StyleType>[][];
    rowMap: Map<number, number>;
    colMap: Map<number, number>;
};
