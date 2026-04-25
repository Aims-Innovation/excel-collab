import { default as React } from 'react';
type Props = {
    leftChildren?: React.ReactNode;
    rightChildren?: React.ReactNode;
    hideNewFile?: boolean;
    hideRenameFile?: boolean;
    hideUser?: boolean;
    hideI18n?: boolean;
    hideTheme?: boolean;
    /** When false, hide all Export menu items (XLSX / CSV / JSON). Defaults to true. */
    canExport?: boolean;
};
export declare const MenuBarContainer: React.FunctionComponent<Props>;
export default MenuBarContainer;
