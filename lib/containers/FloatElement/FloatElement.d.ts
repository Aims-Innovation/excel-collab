import { default as React } from 'react';
import { FloatElementItem } from '../../containers/store';
import { IWindowSize } from '../../types';
type FloatElementProps = FloatElementItem & {
    active: boolean;
    resetResize: (size: IWindowSize) => void;
    pointerDown: (event: React.PointerEvent<HTMLDivElement>) => void;
    resizeDown: (event: React.PointerEvent<HTMLDivElement>) => void;
};
export declare const FloatElement: React.FunctionComponent<FloatElementProps>;
export {};
