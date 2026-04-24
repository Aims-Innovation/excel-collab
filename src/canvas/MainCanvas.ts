import type {
  EventType,
  MainView,
  IController,
  RequestInit,
  ResponseRender,
  RequestRender,
  IWindowSize,
} from '../types';
import { dpr, renderLog, perfMeasure, isDebugEnabled } from '../util';
import { getTheme } from '../theme';
import { transfer, proxy } from 'comlink';

export class MainCanvas implements MainView {
  static instance: MainCanvas;
  private readonly controller: IController;
  private readonly canvas: HTMLCanvasElement;
  /**
   * Last measurements reported by the render worker. Used by the
   * double-click auto-fit handler: the worker measures actual content
   * width/height per rendered cell and returns rows/cols whose content
   * exceeds the default size. For auto-fit we read these back.
   */
  private lastMeasuredRowMap: Record<string, number> = {};
  private lastMeasuredColMap: Record<string, number> = {};
  constructor(controller: IController, canvas: HTMLCanvasElement) {
    this.controller = controller;
    this.canvas = canvas;
    const offscreen = canvas?.transferControlToOffscreen?.();
    const worker = this.controller.getHooks().worker;
    if (offscreen) {
      const data: RequestInit = {
        canvas: offscreen,
        dpr: dpr(),
      };
      worker.init(transfer(data, [data.canvas]));
    }
  }
  getMeasuredRowHeight(row: number): number | undefined {
    return this.lastMeasuredRowMap[row];
  }
  getMeasuredColWidth(col: number): number | undefined {
    return this.lastMeasuredColMap[col];
  }
  private readonly renderCallback = (result: ResponseRender) => {
    const { rowMap, colMap } = result;
    this.lastMeasuredRowMap = rowMap;
    this.lastMeasuredColMap = colMap;
    const rowKeys = Object.keys(rowMap);
    const colKeys = Object.keys(colMap);
    if (colKeys.length === 0 && rowKeys.length === 0) {
      return;
    }
    // IMPORTANT: write measurements via the model layer, NOT through
    // controller.setRowHeight / setColWidth. The controller methods
    // call emitChange() which re-emits 'renderChange' -- but the
    // canvas is already painted with these exact sizes (that's where
    // the measurements came from). Re-emitting queues a redundant
    // render that paints the same thing, and for N rows + M cols of
    // measurement updates that's N+M redundant renders, each of which
    // runs its own renderCallback and can produce further updates.
    // During a column drag that compounded into an unresponsive
    // render queue that never settled. Writing through the model
    // layer persists to Y.Doc (remote peers get the update via
    // doc.on('update'); useCollaboration filters SYNC_FLAG.MODEL-
    // origin transactions so the main thread doesn't double-render).
    this.controller.transaction(() => {
      for (const [row, h] of Object.entries(rowMap)) {
        const r = parseInt(row, 10);
        if (this.controller.getRowHeight(r) === h) {
          continue;
        }
        this.controller.model.setRowHeight(r, h);
      }
      for (const [col, w] of Object.entries(colMap)) {
        const c = parseInt(col, 10);
        if (this.controller.getColWidth(c) === w) {
          continue;
        }
        this.controller.model.setColWidth(c, w);
      }
    });
  };
  async render(data: EventType) {
    return perfMeasure('MainCanvas.render', () => this.renderInner(data));
  }
  private async renderInner(data: EventType) {
    const { controller } = this;
    const currentId = controller.getCurrentSheetId();
    const sheetInfo = controller.getSheetInfo(currentId);
    if (!sheetInfo) {
      if (isDebugEnabled('render')) {
        renderLog('dispatch skipped (no sheetInfo)', {
          changeSet: Array.from(data.changeSet),
          sheetId: currentId,
        });
      }
      return;
    }
    const copyRange = controller.getCopyRange();
    const jsonData = await perfMeasure('controller.toJSON', () =>
      controller.toJSON(),
    );
    if (isDebugEnabled('render')) {
      // Canvas element introspection. After transferControlToOffscreen()
      // the main-thread element still exposes getBoundingClientRect()
      // reporting the CSS display size (NOT the bitmap dimensions). If
      // width/height here are 0, the element was not laid out at paint
      // time — the worker's bitmap was also sized 0 (via MainCanvas
      // .resize), so any rendering no-ops against empty canvas. If
      // storedCanvasSize is 0 too, the "canvas-sizing race" is the
      // cause; if only cssRect is 0 but stored is non-zero, layout
      // changed after the first measure and no one re-measured.
      const canvasEl = this.canvas as HTMLCanvasElement;
      const rect = canvasEl?.getBoundingClientRect?.();
      const sheetData = jsonData.worksheets ?? {};
      const allKeys = Object.keys(sheetData);
      const cellCountForSheet = allKeys.filter((k) =>
        k.startsWith(`${currentId}_`),
      ).length;
      renderLog('dispatch', {
        changeSet: Array.from(data.changeSet),
        sheetId: currentId,
        canvasCssRect: rect
          ? { width: rect.width, height: rect.height }
          : 'no getBoundingClientRect',
        storedCanvasSize: controller.getCanvasSize(),
        headerSize: controller.getHeaderSize(),
        scroll: controller.getScroll(currentId),
        totalCellCount: allKeys.length,
        cellCountForSheet,
      });
    }
    // NOTE: v0.1.13.4 added a defensive JSON.parse(JSON.stringify(...))
    // around sheetData / customHeight / customWidth / autoFilter to
    // strip any Y.Doc reference that leaked through Y.Map#toJSON's
    // non-Y-value passthrough. That was pure overhead on every render
    // (5-20 ms for non-trivial sheets) and doubled-cloned the payload
    // since postMessage's structuredClone runs right after. The guard
    // in controller/workerGuard.ts now rethrows DataCloneError with
    // the non-cloneable key path, so leaks surface loudly without us
    // paying the clone cost on every keystroke / cursor move / remote
    // update. If a leak resurfaces, fix it at the model source, not
    // here.
    const eventData: RequestRender = {
      changeSet: data.changeSet,
      theme: getTheme(),
      canvasSize: controller.getCanvasSize(),
      headerSize: controller.getHeaderSize(),
      currentSheetInfo: sheetInfo,
      scroll: controller.getScroll(currentId),
      range: controller.getActiveRange().range,
      copyRange,
      currentMergeCells: controller.getMergeCellList(currentId),
      customHeight: jsonData.customHeight,
      customWidth: jsonData.customWidth,
      sheetData: jsonData.worksheets,
      autoFilter: jsonData.autoFilter[currentId],
    };

    return perfMeasure('worker.render', () =>
      this.controller
        .getHooks()
        .worker.render(eventData, proxy(this.renderCallback)),
    );
  }
  resize() {
    const { canvas } = this;
    const { width, height } = this.controller.getCanvasSize();
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;
    const eventData: IWindowSize = {
      width,
      height,
    };
    this.controller.getHooks().worker.resize(eventData);
  }
}
