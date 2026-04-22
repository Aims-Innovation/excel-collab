import { ModelJSON, EHorizontalAlign, EVerticalAlign, EUnderLine, IWindowSize, BorderType } from '../../types';
export declare function convertToPt(num: number): number;
export declare const chartTypeList: readonly ["bar", "line", "pie"];
export interface XfItem {
    fontId: string;
    fillId: string;
    numFmtId: string;
    borderId: string;
    applyFill?: string;
    applyFont?: string;
    applyNumberFormat?: string;
    applyAlignment?: string;
    applyBorder?: string;
    alignment?: {
        vertical?: 'top' | 'center' | 'bottom';
        horizontal?: 'left' | 'center' | 'right';
        wrapText?: string;
    };
}
type ObjectItem = Record<string, any>;
export declare function convertXMLToJSON(xmlStr: string): ObjectItem;
export declare function convertXMLDataToModel(xmlData: ObjectItem, imageSizeMap: Record<string, IWindowSize>): ModelJSON;
export declare function importXLSX(file: File): Promise<{
    worksheets: {
        [x: string]: {
            value: import('../..').ResultType;
            formula?: string | undefined;
            fontColor?: string | undefined;
            fillColor?: string | undefined;
            fontSize?: number | undefined;
            fontFamily?: string | undefined;
            verticalAlign?: EVerticalAlign | undefined;
            horizontalAlign?: EHorizontalAlign | undefined;
            isWrapText?: boolean | undefined;
            underline?: EUnderLine | undefined;
            isItalic?: boolean | undefined;
            isBold?: boolean | undefined;
            isStrike?: boolean | undefined;
            numberFormat?: string | undefined;
            borderLeft?: {
                color: string;
                type: BorderType;
            } | undefined;
            borderRight?: {
                color: string;
                type: BorderType;
            } | undefined;
            borderTop?: {
                color: string;
                type: BorderType;
            } | undefined;
            borderBottom?: {
                color: string;
                type: BorderType;
            } | undefined;
        };
    };
    workbook: {
        [x: string]: {
            sheetId: string;
            name: string;
            isHide: boolean;
            rowCount: number;
            colCount: number;
            sort: number;
            tabColor?: string | undefined;
        };
    };
    mergeCells: {
        [x: string]: {
            row: number;
            col: number;
            rowCount: number;
            colCount: number;
            sheetId: string;
        };
    };
    currentSheetId: string;
    customHeight: {
        [x: string]: {
            len: number;
            isHide: boolean;
        };
    };
    customWidth: {
        [x: string]: {
            len: number;
            isHide: boolean;
        };
    };
    definedNames: {
        [x: string]: {
            row: number;
            col: number;
            rowCount: number;
            colCount: number;
            sheetId: string;
        };
    };
    drawings: {
        [x: string]: {
            title: string;
            type: "floating-picture" | "chart";
            uuid: string;
            width: number;
            height: number;
            originWidth: number;
            originHeight: number;
            fromCol: number;
            fromRow: number;
            sheetId: string;
            marginX: number;
            marginY: number;
            imageAngle?: number | undefined;
            imageSrc?: string | undefined;
            chartType?: import('../..').ChartType | undefined;
            chartRange?: {
                row: number;
                col: number;
                rowCount: number;
                colCount: number;
                sheetId: string;
            } | undefined;
        };
    };
    rangeMap: {
        [x: string]: {
            row: number;
            col: number;
            rowCount: number;
            colCount: number;
            sheetId: string;
        };
    };
    autoFilter: {
        [x: string]: {
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
                value: import('../..').ResultType[];
            } | undefined;
        };
    };
    scroll: {
        [x: string]: {
            row: number;
            col: number;
        };
    };
}>;
export {};
