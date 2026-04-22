import { IController, ThemeType, RequestRender } from '../../types';
export declare function getRenderData(controller: IController, theme: ThemeType): RequestRender;
export declare function compareScreenShot(controller: IController, options?: {
    theme?: ThemeType;
    maxThreshold?: number;
}): Promise<void>;
