import { IRange, ResultType } from '../../../types';
import { CellDataMapImpl } from '../../eval';
export declare function expectFormula(str: string, expected: ResultType[], coord?: Pick<IRange, 'sheetId' | 'row' | 'col'>, cellDataMap?: CellDataMapImpl): void;
