import { default as React } from 'react';
import { MessageType } from '../../types';
type Props = {
    message: string;
    type: MessageType;
    duration?: number;
    testId: string;
};
export declare const Toast: React.FunctionComponent<Omit<Props, 'duration'>>;
export declare function toast(props: Props): () => void;
export declare namespace toast {
    var error: (message: string, testId?: string) => () => void;
    var info: (message: string, testId?: string) => () => void;
    var warning: (message: string, testId?: string) => () => void;
    var success: (message: string, testId?: string) => () => void;
}
export {};
