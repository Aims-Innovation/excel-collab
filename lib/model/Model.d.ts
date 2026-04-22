import { StyleType, ModelJSON, WorksheetType, IModel, ResultType, IRange, CustomItem, DrawingElement, DefinedNameItem, WorksheetData, EMergeCellType, IHooks, AutoFilterItem, ModelRoot, ModelScroll, ModelEventEmitterType } from '../types';
import { EventEmitter } from '../util';
export declare class Model extends EventEmitter<ModelEventEmitterType> implements IModel {
    private readonly workbookManager;
    private readonly rangeMapManager;
    private readonly drawingsManager;
    private readonly definedNameManager;
    private readonly worksheetManager;
    private readonly mergeCellManager;
    private readonly rowManager;
    private readonly colManager;
    private readonly filterManager;
    private readonly scrollManager;
    private readonly doc;
    private readonly undoManager;
    constructor(hooks: Pick<IHooks, 'doc' | 'worker'>);
    transaction: <T>(fn: () => T, origin?: any) => T;
    clearHistory(): void;
    getRoot(): ModelRoot;
    computeFormulas(): Promise<boolean>;
    getSheetList(): WorksheetType[];
    getActiveRange(r?: IRange): {
        range: {
            sheetId: string;
            row: number;
            col: number;
            rowCount: number;
            colCount: number;
        };
        isMerged: boolean;
    };
    setActiveRange(range: IRange): void;
    addSheet(): WorksheetType;
    deleteSheet(sheetId?: string): void;
    updateSheetInfo(data: Partial<WorksheetType>, sheetId?: string): void;
    hideSheet(sheetId?: string): void;
    unhideSheet(sheetId: string): void;
    renameSheet(sheetName: string, sheetId?: string): void;
    getSheetInfo(sheetId?: string): WorksheetType | undefined;
    setCurrentSheetId(sheetId: string): void;
    getCurrentSheetId(): string;
    fromJSON: (json: ModelJSON) => void;
    toJSON: () => ModelJSON;
    deleteCell(range: IRange): void;
    setCell(value: ResultType[][], style: Array<Array<Partial<StyleType>>>, range: IRange): Promise<void>;
    setCellValue(value: ResultType, range: IRange): void;
    updateCellStyle(style: Partial<StyleType>, range: IRange): void;
    getCell: (range: IRange) => import('..').ModelCellType | undefined;
    getWorksheet(sheetId?: string): WorksheetData;
    setWorksheet(data: WorksheetData): void;
    addCol(colIndex: number, count: number, isRight?: boolean): void;
    deleteCol(colIndex: number, count: number): void;
    hideCol(colIndex: number, count: number): void;
    getCol(col: number, sheetId?: string): CustomItem;
    setColWidth(col: number, width: number, sheetId?: string): void;
    addRow(rowIndex: number, count: number, isAbove?: boolean): void;
    deleteRow(rowIndex: number, count: number): void;
    hideRow(rowIndex: number, count: number): void;
    unhideRow(rowIndex: number, count: number): void;
    unhideCol(colIndex: number, count: number): void;
    getRow(row: number, sheetId?: string): CustomItem;
    setRowHeight(row: number, height: number, sheetId?: string): void;
    canRedo(): boolean;
    canUndo(): boolean;
    undo(): void;
    redo(): void;
    pasteRange(fromRange: IRange, isCut: boolean): IRange;
    deleteAll(sheetId?: string): void;
    getDefineNameList(): DefinedNameItem[];
    getDefineName(range: IRange): string;
    setDefineName(range: IRange, name: string): boolean;
    checkDefineName(name: string): IRange | undefined;
    getDrawingList(sheetId?: string): DrawingElement[];
    addDrawing(...list: DrawingElement[]): void;
    updateDrawing(uuid: string, value: Partial<DrawingElement>): void;
    deleteDrawing(uuid: string): void;
    getMergeCellList(sheetId?: string): IRange[];
    addMergeCell(range: IRange, type?: EMergeCellType): void;
    deleteMergeCell(range: IRange): void;
    validateRange(range: IRange): boolean;
    validateDrawing(data: DrawingElement): boolean;
    validateDefinedName(name: string): boolean;
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
            value: ResultType[];
        } | undefined;
    } | undefined;
    addFilter(range: IRange): void;
    deleteFilter(sheetId?: string): void;
    updateFilter(sheetId: string, value: Partial<AutoFilterItem>): void;
    getScroll(sheetId?: string): ModelScroll;
    setScroll(value: ModelScroll, sheetId?: string): boolean;
    private getNextSheetId;
}
