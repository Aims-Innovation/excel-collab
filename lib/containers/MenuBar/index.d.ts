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
    /** Hide the entire File menu (and the clickable filename widget on the
     *  left). Wins over hideNewFile / hideRenameFile / canExport when set --
     *  intended for view-only sessions where the user lacks edit permission. */
    hideFileMenu?: boolean;
};
export declare const MenuBarContainer: React.FunctionComponent<Props>;
export default MenuBarContainer;
