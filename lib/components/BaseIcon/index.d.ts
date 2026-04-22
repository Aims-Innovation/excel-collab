import { FunctionComponent } from 'react';
import { BaseIconName } from './icon';
interface IconProps {
    name: BaseIconName;
    className?: string;
    testId?: string;
}
export declare const Icon: FunctionComponent<IconProps>;
export { FillColorIcon } from './FillColorIcon';
