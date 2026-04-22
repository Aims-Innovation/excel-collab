import { IController } from '../../types';
export declare enum ResizePosition {
    top = "top",
    topRight = "top-right",
    topLeft = "top-left",
    bottom = "bottom",
    bottomLeft = "bottom-left",
    bottomRight = "bottom-right",
    left = "left",
    right = "right",
    rotate = "rotate",
    active = "active"
}
export type FloatElementPosition = {
    width: number;
    top: number;
    left: number;
    height: number;
    imageAngle: number;
};
export type State = {
    moveStartX: number;
    moveStartY: number;
    resizePosition: string;
    activeUuid: string;
    position: FloatElementPosition;
};
export declare const INITIAL_STATE: State;
export declare function roundPosition(top: number, left: number, controller: IController): {
    top: number;
    left: number;
};
