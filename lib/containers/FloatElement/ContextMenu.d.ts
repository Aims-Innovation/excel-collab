import { default as React } from 'react';
import { FloatElementItem } from '../../containers/store';
import { IWindowSize } from '../../types';
type Props = FloatElementItem & {
    menuLeft: number;
    menuTop: number;
    resetResize: (size: IWindowSize) => void;
    hideContextMenu: () => void;
};
export declare const FloatElementContextMenu: React.FunctionComponent<Props>;
export {};
