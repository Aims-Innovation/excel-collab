import { default as React } from 'react';
import { SheetItem } from '../store';
interface Props {
    position: number;
    sheetList: SheetItem[];
    currentSheetId: string;
    hideMenu: () => void;
    editSheetName: () => void;
}
export declare const SheetBarContextMenu: React.FunctionComponent<Props>;
export {};
