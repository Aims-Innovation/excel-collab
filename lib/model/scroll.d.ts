import { ModelJSON, IScroll, IModel, ModelScroll } from '../types';
export declare class ScrollManager implements IScroll {
    private model;
    private scroll;
    constructor(model: IModel);
    fromJSON(json: ModelJSON): void;
    toJSON(): {
        [x: string]: {
            row: number;
            col: number;
        };
    };
    deleteAll(sheetId?: string): void;
    getScroll(sheetId?: string): ModelScroll;
    setScroll(value: ModelScroll, sheetId?: string): boolean;
}
