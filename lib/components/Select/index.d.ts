import { default as React, CSSProperties, FunctionComponent } from 'react';
import { OptionItem } from '../../types';
export interface SelectProps {
    value?: string | number;
    defaultValue?: string | number;
    data: Array<string | number | OptionItem>;
    getItemStyle?: (value: string | number) => CSSProperties;
    onChange: (value: string | number) => void;
    title?: string;
    className?: string;
    testId?: string;
}
export declare const Select: FunctionComponent<SelectProps>;
export interface SelectPopupProps {
    active: boolean;
    value: string;
    data: Array<OptionItem>;
    onChange: (value: string) => void;
    position?: 'top' | 'bottom';
    testId?: string;
    className?: string;
}
export declare const SelectPopup: FunctionComponent<SelectPopupProps>;
export type SelectListProps = Omit<SelectPopupProps, 'active'>;
export declare const SelectList: FunctionComponent<React.PropsWithChildren<SelectListProps>>;
