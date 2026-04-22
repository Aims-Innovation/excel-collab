type NameSpaceType = 'react' | 'model' | 'test' | 'collaboration' | 'controller';
export declare class Debug {
    namespace: NameSpaceType;
    static readonly colorMap: Map<NameSpaceType, string>;
    constructor(namespace: NameSpaceType);
    init: () => (...rest: unknown[]) => void;
    log: (...rest: unknown[]) => void;
    enable(): string | null;
    setColor(): void;
}
export declare const modelLog: (...rest: unknown[]) => void;
export declare const controllerLog: (...rest: unknown[]) => void;
export declare const reactLog: (...rest: unknown[]) => void;
export declare const collaborationLog: (...rest: unknown[]) => void;
export {};
