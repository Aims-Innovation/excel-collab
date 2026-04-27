import React, {
  FunctionComponent,
  useState,
  useRef,
  useEffect,
  useLayoutEffect,
  useCallback,
  memo,
  CSSProperties,
} from 'react';
import { createPortal } from 'react-dom';
import styles from './index.module.css';
import { classnames } from '../../util';
import { Button } from '../Button';

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

type SubMenuProps = Pick<
  MenuProps,
  'label' | 'testId' | 'className' | 'portalClassName'
>;

// Rendering popups via React Portal into document.body so they
// escape any ancestor's overflow / transform / contain context. The
// previous absolute-positioned popup was clipped when the toolbar's
// `overflow-x: auto` (added for narrow-viewport horizontal scroll)
// implicitly promoted overflow-y to auto -- per the CSS spec, when
// one axis is non-visible the other axis's `visible` value computes
// to `auto`. Result: dropdowns extending below the toolbar were cut
// off and a phantom vertical scrollbar appeared on the toolbar.
const useIsoLayoutEffect =
  typeof window === 'undefined' ? useEffect : useLayoutEffect;

export const MenuItem: FunctionComponent<
  React.PropsWithChildren<MenuItemProps>
> = ({ onClick, children, testId, active = false }) => {
  return (
    <li
      className={classnames(styles.menuItem, { [styles.active]: active })}
      onClick={onClick}
      data-testid={testId}
    >
      {children}
    </li>
  );
};

export const SubMenu: FunctionComponent<
  React.PropsWithChildren<SubMenuProps>
> = ({ label, children, testId, className, portalClassName }) => {
  const [open, setOpen] = useState(false);
  const handleClick = useCallback(() => {
    setOpen((v) => !v);
  }, []);
  return (
    <li
      className={classnames(styles.menuItem, className)}
      onClick={handleClick}
      data-testid={testId}
    >
      <div>{label}</div>
      <div
        className={classnames(
          styles.subMenuContainer,
          styles.portal,
          portalClassName,
        )}
        hidden={!open}
      >
        <ul className={styles.menu}>{children}</ul>
      </div>
    </li>
  );
};

export const Menu: FunctionComponent<React.PropsWithChildren<MenuProps>> = memo(
  ({
    label,
    children,
    testId,
    className,
    isPlain = false,
    position = 'right',
    size = 'normal',
    portalClassName,
  }) => {
    const buttonType = isPlain ? 'plain' : undefined;
    const [open, setOpen] = useState(false);
    const [popupStyle, setPopupStyle] = useState<CSSProperties>({});
    const triggerRef = useRef<HTMLDivElement>(null);
    const popupRef = useRef<HTMLDivElement>(null);

    const handleClick = useCallback(() => {
      setOpen((v) => !v);
    }, []);

    // Compute popup position from the trigger's bounding rect when
    // the menu opens. Use layoutEffect so the popup paints in the
    // right place on the same frame as the open transition.
    useIsoLayoutEffect(() => {
      if (!open || !triggerRef.current) return;
      const rect = triggerRef.current.getBoundingClientRect();
      if (position === 'bottom') {
        // Below the trigger, right-aligned to the trigger's right
        // edge (matches the previous .menuContainer.bottom rule).
        setPopupStyle({
          position: 'fixed',
          top: rect.bottom,
          right: window.innerWidth - rect.right,
        });
      } else {
        // To the right of the trigger, top-aligned (default).
        setPopupStyle({
          position: 'fixed',
          top: rect.top,
          left: rect.right,
        });
      }
    }, [open, position]);

    // Close on scroll / resize. Excel and Google Sheets behave the
    // same way -- popups don't track their trigger after open. The
    // capture phase catches scrolls in any ancestor (including the
    // toolbar's own overflow-x scroller).
    useEffect(() => {
      if (!open) return;
      const close = () => setOpen(false);
      window.addEventListener('scroll', close, true);
      window.addEventListener('resize', close);
      return () => {
        window.removeEventListener('scroll', close, true);
        window.removeEventListener('resize', close);
      };
    }, [open]);

    // Click-outside that respects both the trigger AND the portal'd
    // popup. The previous useClickOutside hook only knew about a
    // single ref, which doesn't work when the popup lives under
    // document.body (clicking inside it would otherwise be detected
    // as "outside" and close before the click handler ran).
    useEffect(() => {
      if (!open) return;
      const onPointerDown = (e: Event) => {
        const target = e.target as Node | null;
        if (!target) return;
        if (
          triggerRef.current?.contains(target) ||
          popupRef.current?.contains(target)
        ) {
          return;
        }
        setOpen(false);
      };
      const onBlur = () => setOpen(false);
      document.addEventListener('pointerdown', onPointerDown);
      window.addEventListener('blur', onBlur);
      return () => {
        document.removeEventListener('pointerdown', onPointerDown);
        window.removeEventListener('blur', onBlur);
      };
    }, [open]);

    const popup = (
      <div
        ref={popupRef}
        className={classnames(
          styles.menuContainer,
          styles.portal,
          portalClassName,
          {
            [styles.bottom]: position === 'bottom',
          },
        )}
        style={popupStyle}
        data-testid={`${testId}-portal`}
        hidden={!open}
      >
        <ul className={styles.menu}>{children}</ul>
      </div>
    );

    return (
      <div
        ref={triggerRef}
        className={classnames(styles.container, className, {
          [styles.small]: size === 'small',
        })}
        data-testid={testId}
      >
        <Button
          onClick={handleClick}
          testId={`${testId}-trigger`}
          type={buttonType}
          className={styles.trigger}
        >
          {label}
        </Button>
        {typeof document !== 'undefined' &&
          createPortal(popup, document.body)}
      </div>
    );
  },
);

Menu.displayName = 'Menu';
