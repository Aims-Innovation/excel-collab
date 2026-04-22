import { EditorProps, StateContextValue } from './containers';
import { Doc } from 'yjs';
export type DocConfig = ConstructorParameters<typeof Doc>[0];
export type ExcelProps = EditorProps & {
    doc?: Doc;
    docConfig?: DocConfig;
} & Pick<StateContextValue, 'provider' | 'awareness'>;
export declare const Excel: React.FunctionComponent<ExcelProps>;
