import { ModelJSON, DefinedNameItem, IRange, IModel, IDefinedName } from '../types';
export declare class DefinedName implements IDefinedName {
    private model;
    constructor(model: IModel);
    private get definedNames();
    validateDefinedName(name: string): boolean;
    fromJSON(json: ModelJSON): void;
    getDefineNameList(): DefinedNameItem[];
    getDefineName(range: IRange): string;
    setDefineName(range: IRange, name: string): boolean;
    checkDefineName(name: string): IRange | undefined;
    deleteAll(sheetId?: string): void;
}
