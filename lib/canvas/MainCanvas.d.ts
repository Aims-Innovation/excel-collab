import { EventType, MainView, IController } from '../types';
export declare class MainCanvas implements MainView {
    static instance: MainCanvas;
    private readonly controller;
    private readonly canvas;
    constructor(controller: IController, canvas: HTMLCanvasElement);
    private readonly renderCallback;
    render(data: EventType): Promise<void>;
    resize(): void;
}
