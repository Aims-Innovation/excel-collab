import { IRange, WorksheetType, ModelCellType } from '../types';
import { SheetRange } from '../util/range';
export type ResultType = boolean | string | number;
/**
 * Enum representing various token types used in the parser.
 */
export declare enum TokenType {
    /**
     * Equal sign `=`.
     */
    EQUAL = 0,
    /**
     * Not equal sign `<>`.
     */
    NOT_EQUAL = 1,
    /**
     * Plus sign `+`.
     */
    PLUS = 2,
    /**
     * Minus sign `-`.
     */
    MINUS = 3,
    /**
     * Multiplication sign `*`.
     */
    STAR = 4,
    /**
     * Division sign `/`.
     */
    SLASH = 5,
    /**
     * Exponentiation sign `^`.
     */
    EXPONENT = 6,
    /**
     * Greater than sign `>`.
     */
    GREATER = 7,
    /**
     * Greater than or equal to sign `>=`.
     */
    GREATER_EQUAL = 8,
    /**
     * Concatenation sign `&`.
     */
    CONCATENATE = 9,
    /**
     * Colon `:` used as a range operator.
     */
    COLON = 10,
    /**
     * Comma `,` used as a union operator.
     */
    COMMA = 11,
    /**
     * Empty character `' '` used as an intersection operator.
     */
    EMPTY_CHAR = 12,
    /**
     * Percent sign `%`.
     */
    PERCENT = 13,
    /**
     * Less than sign `<`.
     */
    LESS = 14,
    /**
     * Less than or equal to sign `<=`.
     */
    LESS_EQUAL = 15,
    /**
     * String literal.
     */
    STRING = 16,
    /**
     * Number literal (integer or float).
     */
    NUMBER = 17,
    /**
     * Boolean literal (`TRUE` or `FALSE`).
     */
    BOOL = 18,
    /**
     * Error value (e.g., `#DIV/0!`, `#NAME?`, `#N/A`, `#NULL!`, `#NUM!`, `#VALUE!`, `#GETTING_DATA`).
     */
    ERROR = 19,
    /**
     * Reference error `#REF!`.
     */
    ERROR_REF = 20,
    /**
     * Left bracket `(`.
     */
    LEFT_BRACKET = 21,
    /**
     * Right bracket `)`.
     */
    RIGHT_BRACKET = 22,
    /**
     * Left brace `{`.
     */
    lEFT_BRACE = 23,
    /**
     * Right brace `}`.
     */
    RIGHT_BRACE = 24,
    /**
     * Semicolon `;`.
     */
    SEMICOLON = 25,
    /**
     * Exclamation mark `!`.
     */
    EXCLAMATION = 26,
    /**
     * R1C1 reference style.
     */
    R1C1 = 27,
    /**
     * Sheet name followed by an exclamation mark `!`.
     */
    SHEET_NAME = 28,
    /**
     * Cell reference (e.g., `A1`, `$A$1`).
     */
    CELL = 29,
    /**
     * Column reference (e.g., `A`, `$A`).
     */
    COLUMN = 30,
    /**
     * Row reference (e.g., `1`, `$1`).
     */
    ROW = 31,
    /**
     * Defined name (e.g., named ranges or variables).
     */
    DEFINED_NAME = 32,
    /**
     * Built-in Excel function followed by a left parenthesis `(`.
     */
    EXCEL_FUNCTION = 33,
    /**
     * Reference function (e.g., `INDEX`, `OFFSET`, `INDIRECT`) followed by a left parenthesis `(`.
     */
    REF_FUNCTION = 34,
    /**
     * Conditional reference function (e.g., `IF`, `CHOOSE`) followed by a left parenthesis `(`.
     */
    REF_FUNCTION_COND = 35,
    /**
     * End of file.
     */
    EOF = 36
}
/**
 * Interface representing a map of cell data operations.
 */
export interface CellDataMap {
    /**
     * Sets the value of a specified range of cells.
     * @param range - The range of cells to set the value for.
     * @param value - A 2D array of values to set in the specified range.
     */
    set: (range: IRange, value: ResultType[][]) => void;
    /**
     * Retrieves the data of a specific cell within a range.
     * @param range - The range containing the cell to retrieve.
     * @returns The data of the specified cell, or undefined if the cell does not exist.
     */
    getCell: (range: IRange) => ModelCellType | undefined;
    /**
     * Retrieves information about a worksheet.
     * @param sheetId - The ID of the sheet to retrieve information for (optional).
     * @param sheetName - The name of the sheet to retrieve information for (optional).
     * @returns The information of the specified worksheet, or undefined if the sheet does not exist.
     */
    getSheetInfo: (sheetId?: string, sheetName?: string) => WorksheetType | undefined;
    /**
     * Sets a defined name for a specific range.
     * @param name - The name to define.
     * @param value - The range to associate with the defined name.
     */
    setDefinedName: (name: string, value: IRange) => void;
    /**
     * Retrieves the range associated with a defined name.
     * @param name - The defined name to look up.
     * @returns The range associated with the defined name, or undefined if the name does not exist.
     */
    getDefinedName: (name: string) => SheetRange | undefined;
    /**
     * Handles the data of a specific cell.
     * @param value - The data of the cell to handle.
     * @param coord - The coordinates of the cell to handle.
     * @returns An array of results from handling the cell data.
     */
    handleCell: (value: ModelCellType | undefined, coord: Pick<IRange, 'sheetId' | 'row' | 'col'>) => ResultType[];
    /**
     * Retrieves a function by its name.
     * @param name - The name of the function to retrieve.
     * @returns The function associated with the specified name, or undefined if the function does not exist.
     */
    getFunction: (name: string) => any;
}
/**
 * Represents the result of an interpretation process.
 */
export interface InterpreterResult {
    /**
     * An array of results produced by the interpreter.
     */
    result: ResultType[];
    /**
     * An optional string representation of the expression that was interpreted.
     */
    expressionStr?: string;
}
export type ReferenceType = 'absolute' | 'mixed' | 'relative';
export type ConvertSheetName = (value: string) => string;
