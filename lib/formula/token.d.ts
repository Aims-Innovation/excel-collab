import { TokenType } from '../types';
export declare class Token {
    type: TokenType;
    value: string;
    constructor(type: TokenType, value: string);
    toString(): string;
}
