import { ModelCellType, Coordinate } from './model';
import { Doc } from 'yjs';
import { Remote } from 'comlink';
import { IRange } from './range';
import { WorkerMethod } from './event';
import { LANGUAGE_LIST } from '../util';
export type LanguageType = (typeof LANGUAGE_LIST)[number];
export interface OptionItem {
    value: string | number;
    label: string;
    disabled: boolean;
}
export interface CanvasOverlayPosition {
    top: number;
    left: number;
    width: number;
    height: number;
}
export type ActiveCellType = ModelCellType & Coordinate & CanvasOverlayPosition;
export type Point = [x: number, y: number];
export declare enum ScrollStatus {
    NONE = 0,
    VERTICAL = 1,
    HORIZONTAL = 2
}
export declare enum EditorStatus {
    NONE = 0,
    EDIT_CELL = 1,
    EDIT_FORMULA_BAR = 2
}
export type ThemeType = 'dark' | 'light';
export type ClipboardData = {
    'text/plain': string;
    'text/html': string;
    'custom/model': CustomClipboardData | null;
    'image/png': Blob | null;
};
/**
 * Interface representing the hooks for clipboard operations and document handling.
 */
export interface IHooks {
    /**
     * The worker method for handling remote operations.
     */
    worker: Remote<WorkerMethod>;
    /**
     * The Yjs Doc object.
     */
    doc: Doc;
}
export type CustomClipboardData = {
    type: 'cut' | 'copy';
    range: IRange;
    floatElementUuid: string;
};
export type DocumentItem = {
    id: string;
    name?: string;
    create_time: string;
    content?: string;
};
