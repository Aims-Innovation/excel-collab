import { ModelJSON, IModel, AutoFilterItem, IFilter, IRange } from '../types';
export declare class FilterManger implements IFilter {
    private model;
    constructor(model: IModel);
    private get autoFilter();
    fromJSON(json: ModelJSON): void;
    deleteAll(sheetId?: string): void;
    addFilter(range: IRange): void;
    deleteFilter(sheetId?: string): void;
    updateFilter(sheetId: string, value: Partial<AutoFilterItem>): void;
    private clearFilter;
    private applyFilter;
    getFilter(sheetId?: string): {
        range: {
            row: number;
            col: number;
            rowCount: number;
            colCount: number;
            sheetId: string;
        };
        col?: number | undefined;
        value?: {
            type: "number";
            value: {
                type: string;
                value: number;
            }[];
        } | {
            type: "text";
            value: {
                type: string;
                value: string;
            }[];
        } | {
            type: "color";
            value: string;
            colorType: "fillColor" | "fontColor";
        } | {
            type: "normal";
            value: import('..').ResultType[];
        } | undefined;
    } | undefined;
}
