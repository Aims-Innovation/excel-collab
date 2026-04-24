import { IController } from '../types';
/**
 * Pixel tolerance around a header edge where a pointer counts as
 * "over the resize handle".
 */
export declare const RESIZE_HIT_TOLERANCE = 4;
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
export declare function getResizeHit(controller: IController, x: number, y: number): ResizeHit | null;
