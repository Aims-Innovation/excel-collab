type Params = any[];
export declare function throttle(fn: (...params: Params) => void, wait: number): (...args: Params) => void;
export declare function get<T>(obj: Record<string, any> | null | undefined, path: string, defaultValue?: T): T;
export declare function isEmpty(value: unknown): boolean;
export declare function deepEqual(x: any, y: any): boolean;
export declare function noop(): void;
export declare function camelCase(str: string): string;
export declare function omit<T, K extends keyof T>(obj: T, keys: K[]): Omit<T, K>;
export {};
