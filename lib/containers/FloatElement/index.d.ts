import { default as React } from 'react';
import { FloatElementItem } from '../../containers/store';
interface Props {
    floatElementList: FloatElementItem[];
    activeUuid: string;
}
declare const FloatElementContainer: React.FunctionComponent<Props>;
export default FloatElementContainer;
