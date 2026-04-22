import { IController } from '../types';
import { Doc } from 'yjs';
export declare function initController(options?: {
    worker: Worker;
    doc: Doc;
}): IController;
