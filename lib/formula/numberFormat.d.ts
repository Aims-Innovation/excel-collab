import { NumberFormatValue } from '../types';
import * as numberFormatUtil from 'numfmt';
type Options = Parameters<typeof numberFormatUtil.format>[2];
export declare function numberFormat(value: NumberFormatValue, pattern?: string, options?: Options): string;
export declare function isDateFormat(format?: string): boolean;
export declare const convertDateToNumber: typeof numberFormatUtil.dateToSerial;
export {};
