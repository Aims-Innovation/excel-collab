import { EventHandler, EventData } from '../../../types';
import { PointerEvent } from 'react';
export declare class MainHandler implements EventHandler {
    private lastTimeStamp;
    pointerMove(data: EventData): boolean;
    pointerDown(data: EventData, event: PointerEvent<HTMLCanvasElement>): boolean;
}
