import { EditorStatus, OptionItem, WorksheetType, DrawingElement } from '../../types';
export type SheetItem = Required<Pick<WorksheetType, 'sheetId' | 'isHide' | 'name' | 'tabColor'>>;
export type FloatElementItem = DrawingElement & {
    labels: string[];
    datasets: Array<{
        label: string;
        data: number[];
    }>;
    top: number;
    left: number;
};
export interface CoreStore {
    editorStatus: EditorStatus;
    canRedo: boolean;
    canUndo: boolean;
    activeUuid: string;
    currentSheetId: string;
    isFilter: boolean;
    defineNames: string[];
    fontFamilies: OptionItem[];
    sheetList: SheetItem[];
    drawings: FloatElementItem[];
}
type Action = {
    setActiveUuid(uuid: string): void;
    setEditorStatus(status: EditorStatus): void;
    setDefineNames(list: string[]): void;
    setFontFamilies(list: OptionItem[]): void;
    setSheetList(list: SheetItem[]): void;
    setDrawings(drawings: FloatElementItem[]): void;
};
export declare const useCoreStore: import('zustand').UseBoundStore<import('zustand').StoreApi<CoreStore & Action>>;
export {};
