import { FunctionComponent } from 'react';
interface PathItem {
    d: string;
    'fill-opacity'?: string;
}
export interface BaseIconProps {
    className?: string;
    paths: PathItem[];
    testId?: string;
}
export declare const BaseIcon: FunctionComponent<BaseIconProps>;
export {};
