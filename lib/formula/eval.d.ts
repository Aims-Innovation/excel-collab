import { CellDataMap, InterpreterResult, IRange, ResultType, Coordinate, WorksheetType, ModelCellType, RequestFormulas, ResponseFormulas } from '../types';
import { SheetRange } from '../util';
export declare function parseFormula(formula: string, currentCoord?: Pick<IRange, 'sheetId' | 'row' | 'col'>, cellData?: CellDataMap, cache?: Map<string, InterpreterResult>): InterpreterResult;
export declare class CellDataMapImpl implements CellDataMap {
    private readonly map;
    private readonly definedNameMap;
    private readonly currentSheetId;
    private sheetList;
    private getKey;
    getFunction: (name: string) => any;
    setSheetList(list: WorksheetType[]): void;
    set(range: IRange, value: ResultType[][]): void;
    getCell(range: IRange): ModelCellType | undefined;
    getSheetInfo(sheetId?: string, sheetName?: string): WorksheetType | undefined;
    handleCell: (value: ModelCellType | undefined, _coord: Coordinate) => ResultType[];
    setDefinedName(name: string, value: IRange): void;
    getDefinedName(name: string): SheetRange | undefined;
}
export declare function computeFormulas(eventData: RequestFormulas, cb: (data: ResponseFormulas) => boolean): boolean;
