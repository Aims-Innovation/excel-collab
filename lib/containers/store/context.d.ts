import { IController, ICollaborationProvider } from '../../types';
import { Awareness } from 'y-protocols/awareness';
export type StateContextValue = {
    controller: IController;
    provider?: Partial<ICollaborationProvider>;
    awareness?: Awareness;
};
export declare const StateContext: import('react').Context<StateContextValue | undefined>;
export declare function useExcel(): StateContextValue;
