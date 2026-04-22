import { Token } from '../../token';
import { TokenType } from '../../../types';
export type BlockType = [string, Token[]];
export declare function itBlock(list: BlockType[]): void;
export declare function getToken(type: TokenType, value: string): Token;
