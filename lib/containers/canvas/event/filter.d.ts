import { EventHandler, EventData, ModalValue } from '../../../types';
import { PointerEvent } from 'react';
export declare class FilterHandler implements EventHandler {
    pointerDown(data: EventData, event: PointerEvent<HTMLCanvasElement>): false | ModalValue;
    pointerMove(): boolean;
}
