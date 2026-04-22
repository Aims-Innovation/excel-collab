import { ModelJSON, IRow, CustomItem, IModel } from '../types';
export declare class RowManager implements IRow {
    private model;
    constructor(model: IModel);
    private get customHeight();
    private getModel;
    fromJSON(json: ModelJSON): void;
    hideRow(rowIndex: number, count: number): void;
    private toggleHideRow;
    unhideRow(rowIndex: number, count: number): void;
    getRow(row: number, sheetId?: string): CustomItem;
    setRowHeight(row: number, height: number, sheetId?: string): void;
    deleteAll(sheetId?: string): void;
}
