import { WorksheetType, Coordinate, ModelJSON, ResultType, IRange, ChangeEventType } from '../types';
import { Transaction } from 'yjs';
export declare function parseNumber(value: any): [boolean, number];
export declare function getListMaxNum(list?: string[]): number;
export declare function getDefaultSheetInfo(list?: WorksheetType[]): Pick<WorksheetType, 'name' | 'sheetId' | 'sort'>;
export declare function convertStringToResultType(value: any): ResultType;
export declare function stringToCoordinate(key: string): Coordinate;
export declare function getWorksheetKey(sheetId: string, row: number, col: number): string;
export declare function convertWorksheetKey(key: string): Pick<IRange, 'sheetId' | 'row' | 'col'> | null;
export declare function getCustomWidthOrHeightKey(sheetId: string, rowOrCol: number): `${string}_${number}`;
export declare function widthOrHeightKeyToData(key: string): {
    sheetId: string;
    rowOrCol: number;
};
export declare function isMobile(): boolean;
export declare const KEY_LIST: Array<keyof ModelJSON>;
export declare function sleep(ms: number): Promise<unknown>;
export declare function isTestEnv(): boolean;
export declare function isMac(): boolean;
export declare function isFormula(value: ResultType): boolean;
export declare function isMergeContent(isMergeCell: boolean, text: string): boolean;
export declare function getRandomColor(): string;
export declare function stringToUint8Array(str: string): Uint8Array<ArrayBuffer>;
export declare function uint8ArrayToString(bytes: Uint8Array): string;
export declare function modelToChangeSet(list: Transaction): Set<ChangeEventType>;
/**
 * Adjust the decimal-places in an Excel-style number format string.
 *
 * Examples:
 *   adjustDecimalFormat('General', 1)   -> '0.0'
 *   adjustDecimalFormat('0',       1)   -> '0.0'
 *   adjustDecimalFormat('0.00',    1)   -> '0.000'
 *   adjustDecimalFormat('0.00',   -1)   -> '0.0'
 *   adjustDecimalFormat('0.0',    -1)   -> '0'
 *   adjustDecimalFormat('#,##0.00', 1)  -> '#,##0.000'
 *   adjustDecimalFormat('0%',      1)   -> '0.0%'
 *   adjustDecimalFormat('0.00%',  -1)   -> '0.0%'
 *
 * Formats we don't recognise (scientific, fractions, dates, custom
 * currency with escaped text) are returned unchanged — users can still
 * use the explicit "More formats" dropdown for those.
 */
export declare function adjustDecimalFormat(format: string, delta: 1 | -1): string;
