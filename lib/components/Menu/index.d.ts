import { default as React, FunctionComponent } from 'react';
type MenuItemProps = {
    onClick?: () => void;
    testId?: string;
    active?: boolean;
};
type MenuProps = {
    label: React.ReactNode;
    isPlain?: boolean;
    testId?: string;
    className?: string;
    position?: 'right' | 'bottom';
    size?: 'normal' | 'small';
    portalClassName?: string;
};
type SubMenuProps = Pick<MenuProps, 'label' | 'testId' | 'className' | 'portalClassName'>;
export declare const MenuItem: FunctionComponent<React.PropsWithChildren<MenuItemProps>>;
export declare const SubMenu: FunctionComponent<React.PropsWithChildren<SubMenuProps>>;
export declare const Menu: FunctionComponent<React.PropsWithChildren<MenuProps>>;
export {};
