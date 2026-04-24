import { EventType, MainView, IController } from '../types';
export declare class MainCanvas implements MainView {
    static instance: MainCanvas;
    private readonly controller;
    private readonly canvas;
    /**
     * Last measurements reported by the render worker. Used by the
     * double-click auto-fit handler: the worker measures actual content
     * width/height per rendered cell and returns rows/cols whose content
     * exceeds the default size. For auto-fit we read these back.
     */
    private lastMeasuredRowMap;
    private lastMeasuredColMap;
    constructor(controller: IController, canvas: HTMLCanvasElement);
    getMeasuredRowHeight(row: number): number | undefined;
    getMeasuredColWidth(col: number): number | undefined;
    private readonly renderCallback;
    render(data: EventType): Promise<void>;
    private renderInner;
    resize(): void;
}
