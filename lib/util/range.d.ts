import { IRange } from '../types';
export declare function isSheet(range: IRange): boolean;
export declare function isRow(range: IRange): boolean;
export declare function isCol(range: IRange): boolean;
export declare function isSameRange(oldRange: IRange, newRange: IRange): boolean;
export declare function containRange(sourceRange: IRange, targetRange: IRange): boolean;
export declare function toIRange(range: IRange): IRange;
export declare class SheetRange implements IRange {
    row: number;
    col: number;
    colCount: number;
    rowCount: number;
    sheetId: string;
    constructor(row: number, col: number, rowCount: number, colCount: number, sheetId: string);
    static makeRange(range: IRange): SheetRange;
    toIRange(): IRange;
}
export declare function iterateRange(range: IRange, sheetRowCount: number, sheetColCount: number, fn: (row: number, col: number) => boolean): void;
