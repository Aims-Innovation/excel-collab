import { StyleType } from '../../types';
export type StyleStoreType = Omit<StyleType, 'border'> & {
    isMergeCell: boolean;
    mergeType: string;
};
export declare const useStyleStore: import('zustand').UseBoundStore<import('zustand').StoreApi<StyleStoreType>>;
