import { ModelJSON, IRange, IMergeCell, IModel } from '../types';
export declare class MergeCell implements IMergeCell {
    private readonly model;
    constructor(model: IModel);
    private get mergeCells();
    fromJSON(json: ModelJSON): void;
    getMergeCellList(sheetId?: string): IRange[];
    addMergeCell(range: IRange): void;
    deleteMergeCell(range: IRange): void;
    deleteAll(sheetId?: string): void;
    private convertSheetIdToName;
}
