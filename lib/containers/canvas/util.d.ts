import { IController } from '../../types';
import { MainCanvas } from '../../canvas';
export declare function initCanvas(controller: IController, canvas: HTMLCanvasElement): {
    mainCanvas: MainCanvas;
    dispose: () => void;
};
