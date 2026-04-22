export declare function type(content: string, isEnter?: boolean): void;
export declare function extractDataFromTransform(transform: string, type: 'translateX' | 'translateY' | 'rotate'): number;
export declare function renderComponent(): Promise<{
    controller: import('..').IController;
}>;
