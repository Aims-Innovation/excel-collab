import { CanvasOverlayPosition } from '../../types';
export type CellStoreType = CanvasOverlayPosition & {
    defineName: string;
    value: string;
    displayValue: string;
    row: number;
    col: number;
    rowCount: number;
    colCount: number;
};
export declare const useActiveCell: import('zustand').UseBoundStore<import('zustand').StoreApi<CellStoreType>>;
