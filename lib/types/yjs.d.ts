import { Map, Array } from 'yjs';
import { WorksheetType, ModelCellType, CustomItem, DrawingElement, AutoFilterItem, ModelScroll } from './model';
import { IRange } from './range';
export type StringKeyOf<T> = Extract<keyof T, string>;
export interface TypedMap<Data extends Record<string, any>> extends Map<any> {
    /**
     * Constructor type is unusable with interface declaration, but this is how I would implement it:
     * new <Key extends StringKeyOf<Data>>(entries?: Iterable<[Key, Data[Key]]>): TypedMap<Data>;
     * */
    clone(): TypedMap<Data>;
    keys<Key extends StringKeyOf<Data>>(): IterableIterator<Key>;
    values<Key extends StringKeyOf<Data>>(): IterableIterator<Data[Key]>;
    entries<Key extends StringKeyOf<Data>>(): IterableIterator<[Key, Data[Key]]>;
    forEach<Key extends StringKeyOf<Data>>(f: (arg0: Data[Key], arg1: Key, arg2: TypedMap<Data>) => void): void;
    delete<Key extends StringKeyOf<Data>>(key: Key): void;
    set<Key extends StringKeyOf<Data>>(key: Key, value: Data[Key]): Data[Key];
    get<Key extends StringKeyOf<Data>>(key: Key): Data[Key] | undefined;
    has<Key extends StringKeyOf<Data>>(key: Key): boolean;
    [Symbol.iterator]<Key extends StringKeyOf<Data>>(): IterableIterator<[
        Key,
        Data[Key]
    ]>;
    toJSON(): Convert<Data>;
}
export interface TypedArray<T> extends Array<any> {
    get(index: number): T | undefined;
}
export type YWorksheet = TypedMap<Record<string, TypedMap<ModelCellType>>>;
export type YWorkbook = TypedMap<Record<string, TypedMap<WorksheetType>>>;
export type YRange = TypedMap<Record<string, IRange>>;
export type YCustomSize = TypedMap<Record<string, CustomItem>>;
export type YDrawings = TypedMap<Record<string, TypedMap<DrawingElement>>>;
export type YAutoFilter = TypedMap<Record<string, TypedMap<AutoFilterItem>>>;
export type YScroll = TypedMap<Record<string, TypedMap<ModelScroll>>>;
export type YjsModelJson = {
    worksheets: YWorksheet;
    workbook: YWorkbook;
    mergeCells: YRange;
    currentSheetId: string;
    customHeight: YCustomSize;
    customWidth: YCustomSize;
    definedNames: YRange;
    drawings: YDrawings;
    rangeMap: YRange;
    autoFilter: YAutoFilter;
    scroll: YScroll;
};
type Convert<T> = T extends Record<string, any> ? T extends TypedMap<infer U> ? Convert<U> : {
    [K in keyof T]: Convert<T[K]>;
} : NonNullable<T>;
export type ModelJSON = Convert<YjsModelJson>;
export type ModelRoot = TypedMap<YjsModelJson>;
export declare enum ProviderStatus {
    ONLINE = "connection",
    LOCAL = "disconnection",
    SYNCING = "syncing"
}
export {};
