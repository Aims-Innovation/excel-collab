import { ModelJSON, IRangeMap, IRange, IModel } from '../types';
export declare class RangeMap implements IRangeMap {
    private model;
    private rangeMap;
    constructor(model: IModel);
    validateRange(range: IRange): boolean;
    fromJSON(json: ModelJSON): void;
    toJSON(): {
        [x: string]: {
            row: number;
            col: number;
            rowCount: number;
            colCount: number;
            sheetId: string;
        };
    };
    getActiveRange(): {
        range: {
            row: number;
            col: number;
            rowCount: number;
            colCount: number;
            sheetId: string;
        };
        isMerged: boolean;
    };
    setActiveRange(newRange: IRange): void;
    deleteAll(sheetId?: string): void;
}
