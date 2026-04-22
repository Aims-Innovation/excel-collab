import { ErrorTypes } from './constant';
export declare const getFunction: (name: string) => any;
export declare const getCustomError: (value: ErrorTypes) => Error;
export declare const anyError: (err: unknown) => "#GETTING_DATA" | "#DIV/0!" | "#NULL!" | "#NUM!" | "#REF!" | "#VALUE!" | "#N/A" | "#NAME?" | undefined;
