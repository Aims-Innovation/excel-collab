import { default as React, FunctionComponent } from 'react';
interface DialogProps {
    title: string;
    visible: boolean;
    className?: string;
    children?: React.ReactNode;
    testId?: string;
    getContainer?: () => HTMLElement;
    onOk?: React.MouseEventHandler<HTMLButtonElement>;
    onCancel?: React.MouseEventHandler<HTMLButtonElement>;
}
export declare const Dialog: FunctionComponent<DialogProps>;
export declare function info(props: DialogProps): {
    close: () => void;
    update: (modalProps: DialogProps) => void;
};
export {};
