import { default as React, CSSProperties } from 'react';
import { EditorStatus } from '../../types';
import { CellStoreType, StyleStoreType } from '../store';
interface MultipleLineEditorProps {
    isMergeCell: boolean;
    initValue: string;
    style: CSSProperties | undefined;
    testId?: string;
    className?: string;
}
export declare function getDisplayStyle(style: StyleStoreType, isFormulaBar?: boolean): CSSProperties;
export declare function getEditorStyle(style: CellStoreType, editorStatus: EditorStatus, cellStyle: StyleStoreType): CSSProperties | undefined;
export declare const MultipleLineEditor: React.FunctionComponent<MultipleLineEditorProps>;
export {};
