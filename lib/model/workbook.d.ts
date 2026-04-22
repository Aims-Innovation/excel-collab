import { ModelJSON, WorksheetType, IWorkbook, IModel } from '../types';
export declare class Workbook implements IWorkbook {
    private readonly model;
    private currentSheetId;
    constructor(model: IModel);
    private get workbook();
    validateSheet(data: WorksheetType): boolean;
    fromJSON(json: ModelJSON): void;
    updateSheetInfo(data: Partial<WorksheetType>, sheetId?: string): void;
    getSheetList(): WorksheetType[];
    addSheet(): WorksheetType;
    deleteSheet(sheetId?: string): void;
    hideSheet(sheetId?: string): void;
    unhideSheet(sheetId: string): void;
    renameSheet(sheetName: string, sheetId?: string): void;
    getSheetInfo(id?: string): WorksheetType | undefined;
    setCurrentSheetId(newSheetId: string): void;
    getCurrentSheetId(): string;
    private getSheetId;
    deleteAll(sheetId?: string): void;
    private checkSheetSize;
}
