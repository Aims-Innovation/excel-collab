import { ModelJSON, ICol, IModel, CustomItem } from '../types';
export declare class ColManager implements ICol {
    private model;
    constructor(model: IModel);
    private get customWidth();
    private getModel;
    fromJSON(json: ModelJSON): void;
    hideCol(colIndex: number, count: number): void;
    private toggleHideCol;
    unhideCol(colIndex: number, count: number): void;
    getCol(col: number, sheetId?: string): CustomItem;
    setColWidth(col: number, width: number, sheetId?: string): void;
    deleteAll(sheetId?: string): void;
}
