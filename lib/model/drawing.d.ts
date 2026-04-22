import { ModelJSON, IDrawings, DrawingElement, IModel } from '../types';
export declare class Drawing implements IDrawings {
    private readonly model;
    constructor(model: IModel);
    private get drawings();
    validateDrawing(data: DrawingElement): boolean;
    fromJSON(json: ModelJSON): void;
    getDrawingList(sheetId?: string): DrawingElement[];
    addDrawing(data: DrawingElement): void;
    updateDrawing(uuid: string, value: Partial<DrawingElement>): void;
    deleteDrawing(uuid: string): void;
    addCol(colIndex: number, count: number, isRight?: boolean): void;
    addRow(rowIndex: number, count: number, isAbove?: boolean): void;
    deleteCol(colIndex: number, count: number): void;
    deleteRow(rowIndex: number, count: number): void;
    deleteAll(sheetId?: string): void;
}
