import type { IController } from '../types';
import { HIDE_CELL } from '../util';

/**
 * Pixel tolerance around a header edge where a pointer counts as
 * "over the resize handle".
 */
export const RESIZE_HIT_TOLERANCE = 4;

export type ResizeHit = {
  axis: 'row' | 'col';
  /** The row or column whose size this handle controls (the one before the boundary). */
  index: number;
  /** The pixel offset (in canvas space) of the edge being dragged. */
  edgeOffset: number;
};

/**
 * Hit-test the resize handles near column/row header edges.
 *
 * Returns a ResizeHit if (x, y) is inside the header strip AND within
 * RESIZE_HIT_TOLERANCE pixels of a row boundary (for row resize) or a
 * column boundary (for column resize). Returns null otherwise.
 *
 * Column handles live on the right edge of the column header (y is in
 * the column-header strip, x near the boundary of two columns). Row
 * handles live on the bottom edge of the row header.
 */
export function getResizeHit(
  controller: IController,
  x: number,
  y: number,
): ResizeHit | null {
  const header = controller.getHeaderSize();
  const scroll = controller.getScroll();
  const sheetInfo = controller.getSheetInfo(controller.getCurrentSheetId());
  if (!sheetInfo) return null;

  const inColHeader = y >= 0 && y <= header.height && x > header.width;
  const inRowHeader = x >= 0 && x <= header.width && y > header.height;

  if (inColHeader) {
    let cx = header.width;
    for (let c = scroll.col; c < sheetInfo.colCount; c++) {
      const w = controller.getColWidth(c);
      if (w === HIDE_CELL) continue;
      cx += w;
      if (Math.abs(x - cx) <= RESIZE_HIT_TOLERANCE) {
        return { axis: 'col', index: c, edgeOffset: cx };
      }
      if (cx - RESIZE_HIT_TOLERANCE > x + RESIZE_HIT_TOLERANCE) break;
    }
    return null;
  }

  if (inRowHeader) {
    let cy = header.height;
    for (let r = scroll.row; r < sheetInfo.rowCount; r++) {
      const h = controller.getRowHeight(r);
      if (h === HIDE_CELL) continue;
      cy += h;
      if (Math.abs(y - cy) <= RESIZE_HIT_TOLERANCE) {
        return { axis: 'row', index: r, edgeOffset: cy };
      }
      if (cy - RESIZE_HIT_TOLERANCE > y + RESIZE_HIT_TOLERANCE) break;
    }
    return null;
  }

  return null;
}
