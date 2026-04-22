import { CellDataMap, ResultType, IRange } from '../types';
import { Visitor, Expression, CellRangeExpression, PostUnaryExpression, BinaryExpression, UnaryExpression, CellExpression, CallExpression, LiteralExpression, GroupExpression, ArrayExpression } from './expression';
export declare class Interpreter implements Visitor {
    private readonly expressions;
    private readonly cellDataMap;
    private currentCoord;
    constructor(expressions: Expression[], currentCoord: Pick<IRange, 'sheetId' | 'row' | 'col'>, cellDataMap: CellDataMap);
    interpret(): ResultType[];
    visitBinaryExpression(data: BinaryExpression): any;
    visitCallExpression(expr: CallExpression): any;
    visitCellExpression(data: CellExpression): IRange;
    visitLiteralExpression(expr: LiteralExpression): string | number | boolean;
    visitUnaryExpression(data: UnaryExpression): any;
    visitCellRangeExpression(expr: CellRangeExpression): any;
    visitGroupExpression(expr: GroupExpression): any;
    visitArrayExpression(expr: ArrayExpression): ResultType[];
    visitPostUnaryExpression(expr: PostUnaryExpression): any;
    private evaluate;
    private getRangeCellValue;
    private getOneValue;
    private parseNumber;
    private getCellValue;
}
