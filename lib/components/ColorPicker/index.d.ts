import { default as React, FunctionComponent } from 'react';
export interface ColorPickerProps {
    color: string;
    onChange: (value: string) => void;
    className?: string;
    position?: 'top' | 'bottom' | 'right';
    testId?: string;
}
export declare const ColorPicker: FunctionComponent<React.PropsWithChildren<ColorPickerProps>>;
