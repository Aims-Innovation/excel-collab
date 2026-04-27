import React, {
  FunctionComponent,
  useState,
  useRef,
  useEffect,
  useLayoutEffect,
  CSSProperties,
  memo,
  useCallback,
} from 'react';
import { createPortal } from 'react-dom';
import { classnames, COLOR_PICKER_COLOR_LIST } from '../../util';
import styles from './index.module.css';
import { ColorPickerPanel } from './ColorPickerPanel';
import { Button } from '../Button';
import i18n from '../../i18n';

const useIsoLayoutEffect =
  typeof window === 'undefined' ? useEffect : useLayoutEffect;

export interface ColorPickerProps {
  color: string;
  onChange: (value: string) => void;
  className?: string;
  position?: 'top' | 'bottom' | 'right';
  testId?: string;
}

export const ColorPicker: FunctionComponent<
  React.PropsWithChildren<ColorPickerProps>
> = memo((props) => {
  const {
    onChange,
    children,
    color,
    position = 'bottom',
    testId,
    className,
  } = props;
  const [visible, setVisible] = useState(false);
  const [popupStyle, setPopupStyle] = useState<CSSProperties>({});
  const triggerRef = useRef<HTMLDivElement>(null);
  const popupRef = useRef<HTMLDivElement>(null);

  // Compute popup position from the trigger's bounding rect when
  // the picker opens. Same pattern used in Menu and SelectList:
  // render via Portal with `position: fixed` so the popup escapes
  // the toolbar's overflow-x scroll context. Without this the
  // ColorPicker swatches were either clipped at the toolbar's
  // bottom edge or pushed the toolbar to grow vertically.
  useIsoLayoutEffect(() => {
    if (!visible || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    if (position === 'top') {
      // Above the trigger, growing upward via translate(-100%).
      setPopupStyle({
        position: 'fixed',
        top: rect.top,
        left: rect.left,
        transform: 'translate(0, -100%)',
      });
    } else if (position === 'right') {
      // To the right of the trigger.
      setPopupStyle({
        position: 'fixed',
        top: rect.top,
        left: rect.right + 16,
      });
    } else {
      // 'bottom' (default): below the trigger.
      setPopupStyle({
        position: 'fixed',
        top: rect.bottom,
        left: rect.left,
      });
    }
  }, [visible, position]);

  // Close on scroll / resize -- standard floating-popup behavior.
  useEffect(() => {
    if (!visible) return;
    const close = () => setVisible(false);
    window.addEventListener('scroll', close, true);
    window.addEventListener('resize', close);
    return () => {
      window.removeEventListener('scroll', close, true);
      window.removeEventListener('resize', close);
    };
  }, [visible]);

  // Click-outside that respects both the trigger ref AND the
  // portal'd popup ref.
  useEffect(() => {
    if (!visible) return;
    const onPointerDown = (e: Event) => {
      const target = e.target as Node | null;
      if (!target) return;
      if (
        triggerRef.current?.contains(target) ||
        popupRef.current?.contains(target)
      ) {
        return;
      }
      setVisible(false);
    };
    const onBlur = () => setVisible(false);
    document.addEventListener('pointerdown', onPointerDown);
    window.addEventListener('blur', onBlur);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown);
      window.removeEventListener('blur', onBlur);
    };
  }, [visible]);

  const openColorPicker = useCallback(() => {
    setVisible(true);
  }, []);
  const reset = useCallback(() => {
    onChange('');
  }, []);
  const handleClick = useCallback(
    (event: React.MouseEvent<HTMLDivElement>) => {
      const color = (event.target as { dataset?: { value?: string } })?.dataset
        ?.value;
      if (color) {
        onChange(color);
      }
    },
    [],
  );

  const popup = (
    <div
      ref={popupRef}
      className={classnames(styles['color-picker-wrapper'])}
      style={popupStyle}
    >
      <div
        className={styles['color-picker-list']}
        onClick={handleClick}
        data-testid={`${testId}-list`}
      >
        {COLOR_PICKER_COLOR_LIST.map((item) => {
          return (
            <div
              key={item}
              className={styles['color-picker-item']}
              style={{ backgroundColor: item }}
              data-value={item}
            />
          );
        })}
      </div>
      <div>
        <ColorPickerPanel
          color={color}
          onChange={onChange}
          testId={testId}
        />
      </div>
      <div>
        <Button
          type="normal"
          className={styles.reset}
          onClick={reset}
          testId={`${testId}-reset`}
        >
          {i18n.t('reset')}
        </Button>
      </div>
    </div>
  );

  return (
    <div
      ref={triggerRef}
      className={classnames(styles['color-picker'], className, {
        [styles.top]: position === 'top',
        [styles.right]: position === 'right',
      })}
    >
      <div
        className={styles['color-picker-trigger']}
        onClick={openColorPicker}
        data-testid={`${testId}-trigger`}
      >
        {children}
      </div>
      {visible &&
        typeof document !== 'undefined' &&
        createPortal(popup, document.body)}
    </div>
  );
});
ColorPicker.displayName = 'ColorPicker';
