import { IController } from '../../types';
export declare function convertToXMLData(controller: IController): Record<string, string>;
export declare function exportXLSX(fileName: string, controller: IController): Promise<void>;
