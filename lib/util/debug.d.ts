type NameSpaceType = 'react' | 'model' | 'test' | 'collaboration' | 'controller' | 'perf' | 'render' | 'input';
export declare function isDebugEnabled(namespace: NameSpaceType): boolean;
export declare class Debug {
    namespace: NameSpaceType;
    static readonly colorMap: Map<NameSpaceType, string>;
    constructor(namespace: NameSpaceType);
    init: () => (...rest: unknown[]) => void;
    log: (...rest: unknown[]) => void;
    enable(): boolean;
    setColor(): void;
}
export declare const modelLog: (...rest: unknown[]) => void;
export declare const controllerLog: (...rest: unknown[]) => void;
export declare const reactLog: (...rest: unknown[]) => void;
export declare const collaborationLog: (...rest: unknown[]) => void;
export declare const perfLog: (...rest: unknown[]) => void;
export declare const renderLog: (...rest: unknown[]) => void;
export declare const inputLog: (...rest: unknown[]) => void;
/**
 * Measure a synchronous or async block of work. No-ops when 'perf' is
 * not enabled, so it's safe to leave in hot paths.
 *
 *   await perfMeasure('canvas.render', () => mainCanvas.render(data))
 */
export declare function perfMeasure<T>(label: string, fn: () => T | Promise<T>): Promise<T>;
export {};
