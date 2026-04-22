import { default as React } from 'react';
import { HsvaColor } from './types';
interface Props {
    hsva: HsvaColor;
    testId?: string;
    onChange: (newColor: {
        s: number;
        v: number;
    }) => void;
}
export declare const Saturation: React.MemoExoticComponent<({ hsva, testId, onChange }: Props) => import("react/jsx-runtime").JSX.Element>;
export {};
