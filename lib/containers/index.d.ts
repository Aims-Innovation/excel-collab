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
};
declare const ExcelEditor: React.FunctionComponent<EditorProps>;
export * from './store';
export * from './MenuBar';
export { ExcelEditor, useCollaboration, ToolbarContainer, FormulaBarContainer, CanvasContainer, SheetBarContainer, };
