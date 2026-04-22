import { RequestRender, IWindowSize, WorkerMainView, ResponseRender } from '../types';
/**
 * run OffScreenWorker in Web Worker env
 */
export default class OffScreenWorker implements WorkerMainView {
    private readonly canvas;
    private readonly ctx;
    private width;
    private height;
    private isRendering;
    private rowMap;
    private colMap;
    private eventData;
    constructor(canvas: OffscreenCanvas);
    render(data: RequestRender): ResponseRender | undefined;
    resize(data: IWindowSize): void;
    private clear;
    private renderRowsHeader;
    private renderColsHeader;
    private renderFilter;
    private renderGrid;
    private renderTriangle;
    private getRowHeight;
    private getColWidth;
    private isHighlightRow;
    private isHighlightCol;
    private renderAntLine;
    private renderMergeCell;
    private getCellSize;
    private computeCellPosition;
    private getActiveRange;
    private clearRect;
    private renderContent;
    private renderCell;
    private renderSelection;
    private renderSelectRange;
    private renderSelectAll;
    private renderSelectCol;
    private renderSelectRow;
    private renderActiveCell;
}
