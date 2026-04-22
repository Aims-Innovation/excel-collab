import { KeyboardEventItem, IController, IRange } from '../types';
export declare function handleTabClick(controller: IController): void;
export declare function handleEnterClick(controller: IController): void;
export declare function computeScrollRowAndCol(controller: IController, left: number, top: number): {
    row: number;
    col: number;
};
export declare function scrollSheetToView(sheetId: string): void;
export declare function scrollToView(controller: IController, range: IRange): void;
export declare function computeScrollPosition(controller: IController): {
    maxHeight: number;
    maxWidth: number;
    maxScrollHeight: number;
    maxScrollWidth: number;
};
export declare function scrollBar(controller: IController, scrollX: number, scrollY: number): void;
export declare function recalculateScroll(controller: IController): void;
export declare function checkFocus(): boolean;
export declare function setActiveCellValue(controller: IController): void;
export declare const keyboardEventList: KeyboardEventItem[];
