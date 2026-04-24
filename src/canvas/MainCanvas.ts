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
    // Overwrite rather than accumulate. Previous revisions merged into
    // a long-lived map that grew with every render, turning these two
    // object spreads into O(all-measurements) per paint and leaking
    // stale measurements forever. Auto-fit only reads measurements for
    // the currently-visible column/row (you can't double-click a
    // header that isn't in view), so the most-recent render's maps are
    // the only ones we need.
    this.lastMeasuredRowMap = rowMap;
    this.lastMeasuredColMap = colMap;
    const rowKeys = Object.keys(rowMap);
    const colKeys = Object.keys(colMap);
    if (colKeys.length === 0 && rowKeys.length === 0) {
      return;
    }
    this.controller.transaction(() => {
      for (const [row, h] of Object.entries(rowMap)) {
        const r = parseInt(row, 10);

        const old = this.controller.getRowHeight(r);
        if (old === h) {
          continue;
        }

        this.controller.setRowHeight(r, h);
      }
      for (const [col, w] of Object.entries(colMap)) {
        const c = parseInt(col, 10);
        const old = this.controller.getColWidth(c);
        if (old === w) {
          continue;
        }

        this.controller.setColWidth(c, w);
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
      return;
    }
    if (isDebugEnabled('render')) {
      renderLog('dispatch', {
        changeSet: Array.from(data.changeSet),
        sheetId: currentId,
      });
    }
    const copyRange = controller.getCopyRange();
    const jsonData = await perfMeasure('controller.toJSON', () =>
      controller.toJSON(),
    );
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
