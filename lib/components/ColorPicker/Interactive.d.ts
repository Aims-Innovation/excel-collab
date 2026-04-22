import { default as React } from 'react';
export interface Interaction {
    left: number;
    top: number;
}
interface Props {
    onMove: (interaction: Interaction) => void;
    children: React.ReactNode;
    testId?: string;
}
export declare const Interactive: React.MemoExoticComponent<({ onMove, testId, ...rest }: Props) => import("react/jsx-runtime").JSX.Element>;
export {};
