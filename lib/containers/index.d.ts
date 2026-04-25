import { default as FormulaBarContainer } from './FormulaBar';
import { default as ToolbarContainer } from './ToolBar';
import { default as CanvasContainer } from './canvas';
import { default as SheetBarContainer } from './SheetBar';
declare function useCollaboration(): {
    isLoading: boolean;
};
export type EditorProps = {
    style?: React.CSSProperties;
    menubarLeftChildren?: React.ReactNode;
    menubarRightChildren?: React.ReactNode;
    toolbarChildren?: React.ReactNode;
    sheetBarChildren?: React.ReactNode;
    /** Hide the "New File" entry in the File menu (host app owns document creation). */
    hideNewFile?: boolean;
    /** Hide the "Rename File" menu entry and the clickable filename widget (host app owns document naming). */
    hideRenameFile?: boolean;
    /** Hide the user / avatar widget on the right of the menubar. */
    hideUser?: boolean;
    /** Hide the language switcher on the right of the menubar. */
    hideI18n?: boolean;
    /** Hide the light / dark theme toggle on the right of the menubar. */
    hideTheme?: boolean;
    /** When false, hide all Export menu items (XLSX / CSV / JSON). Defaults to true. */
    canExport?: boolean;
};
declare const ExcelEditor: React.FunctionComponent<EditorProps>;
export * from './store';
export * from './MenuBar';
export { ExcelEditor, useCollaboration, ToolbarContainer, FormulaBarContainer, CanvasContainer, SheetBarContainer, };
