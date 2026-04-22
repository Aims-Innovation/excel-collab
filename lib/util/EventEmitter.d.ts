import { IEventEmitter } from '../types';
export declare class EventEmitter<EventType extends Record<string, unknown> = Record<string, unknown>> implements IEventEmitter<EventType> {
    private event;
    constructor();
    getEventLength<T extends keyof EventType>(name: T): number;
    on: <T extends keyof EventType>(name: T, callback: (data: EventType[T]) => void) => VoidFunction;
    emit: <T extends keyof EventType>(name: T, data: EventType[T]) => void;
    off: <T extends keyof EventType>(name: T, callback?: (data: EventType[T]) => void) => void;
    once: <T extends keyof EventType>(name: T, callback: (data: EventType[T]) => void) => VoidFunction;
}
