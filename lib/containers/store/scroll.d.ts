export interface ScrollStore {
    scrollTop: number;
    scrollLeft: number;
    top: number;
    left: number;
    row: number;
    col: number;
    canvasHeight: number;
    canvasWidth: number;
}
export declare const useScrollStore: import('zustand').UseBoundStore<import('zustand').StoreApi<ScrollStore>>;
