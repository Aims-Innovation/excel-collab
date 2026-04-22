import { default as React } from 'react';
interface Props {
    className?: string;
    hue: number;
    testId?: string;
    onChange: (newHue: {
        h: number;
    }) => void;
}
export declare const Hue: React.MemoExoticComponent<({ className, hue, testId, onChange }: Props) => import("react/jsx-runtime").JSX.Element>;
export {};
