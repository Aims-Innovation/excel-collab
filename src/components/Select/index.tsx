import React, {
  CSSProperties,
  FunctionComponent,
  memo,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import { createPortal } from 'react-dom';
import { classnames } from '../../util';
import { OptionItem } from '../../types';
import styles from './index.module.css';
import { Icon } from '../BaseIcon';
import { Button } from '../Button';

const useIsoLayoutEffect =
  typeof window === 'undefined' ? useEffect : useLayoutEffect;

export interface SelectProps {
  value?: string | number;
  defaultValue?: string | number;
  data: Array<string | number | OptionItem>;
  getItemStyle?: (value: string | number) => CSSProperties;
  onChange: (value: string | number) => void;
  title?: string;
  className?: string;
  testId?: string;
}

export const Select: FunctionComponent<SelectProps> = memo((props) => {
  const {
    data,
    value: activeValue,
    className,
    onChange,
    getItemStyle,
    title,
    defaultValue,
    testId,
  } = props;
  const handleChange = useCallback(
    (event: React.ChangeEvent<HTMLSelectElement>) => {
      onChange(event.target.value);
    },
    [],
  );

  return (
    <select
      onChange={handleChange}
      value={activeValue}
      defaultValue={defaultValue}
      name="select"
      className={classnames(styles.selectList, className)}
      title={title}
      data-testid={testId}
    >
      {data.map((item) => {
        const value = typeof item === 'object' ? item.value : item;
        const label = typeof item === 'object' ? item.label : item;
        const disabled = typeof item === 'object' ? item.disabled : false;
        let itemStyle = undefined;
        if (typeof getItemStyle === 'function') {
          itemStyle = getItemStyle(value);
        }
        return (
          <option
            key={value}
            value={value}
            disabled={!!disabled}
            className={classnames(styles.selectItem, {
              [styles['disabled']]: disabled,
            })}
            style={itemStyle}
          >
            {label}
          </option>
        );
      })}
    </select>
  );
});
Select.displayName = 'Select';

export interface SelectPopupProps {
  active: boolean;
  value: string;
  data: Array<OptionItem>;
  onChange: (value: string) => void;
  position?: 'top' | 'bottom';
  testId?: string;
  className?: string;
}

export const SelectPopup: FunctionComponent<SelectPopupProps> = memo(
  ({
    onChange,
    value,
    active,
    data,
    className,
    position = 'bottom',
    testId,
  }) => {
    const handleSelect = useCallback(
      (event: React.MouseEvent<HTMLDivElement>) => {
        const v = (event.target as { dataset?: { value?: string } }).dataset
          ?.value;
        if (!v) {
          return;
        }
        onChange(v);
      },
      [onChange],
    );
    return (
      <div
        className={classnames(
          styles['popup-container'],
          className,
          position === 'top' ? styles.top : '',
          {
            [styles.active]: active,
          },
        )}
        onClick={handleSelect}
        data-testid={testId}
      >
        {data.map((v) => (
          <div key={v.value} className={styles['popup-item']}>
            <span className={styles['popup-item-content']} data-value={v.value}>
              {v.label}
            </span>
            {v.value == value && (
              <span className={styles['popup-item-icon']}>
                <Icon name="confirm" />
              </span>
            )}
          </div>
        ))}
      </div>
    );
  },
);

SelectPopup.displayName = 'SelectPopup';

export type SelectListProps = Omit<SelectPopupProps, 'active'>;

export const SelectList: FunctionComponent<
  React.PropsWithChildren<SelectListProps>
> = memo(({ children, value, data, onChange, position, testId, className }) => {
  const [active, setActive] = useState(false);
  const [popupStyle, setPopupStyle] = useState<CSSProperties>({});
  const triggerRef = useRef<HTMLDivElement>(null);
  const popupRef = useRef<HTMLDivElement>(null);
  const handleClick = useCallback(() => {
    setActive((v) => !v);
  }, []);
  const handleChange = useCallback(
    (value: string) => {
      if (value) {
        onChange(value);
        setActive(false);
      }
    },
    [onChange],
  );

  // Compute popup position from the trigger's bounding rect when
  // the popup opens. position='top' anchors at the trigger's top
  // edge (popup grows upward via the .top transform); 'bottom'
  // (default) anchors at the trigger's bottom edge.
  useIsoLayoutEffect(() => {
    if (!active || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    const placement = position === 'top' ? 'top' : 'bottom';
    setPopupStyle({
      position: 'fixed',
      top: placement === 'top' ? rect.top : rect.bottom,
      left: rect.left,
      width: rect.width,
    });
  }, [active, position]);

  // Close on scroll / resize -- standard floating-popup behavior.
  useEffect(() => {
    if (!active) return;
    const close = () => setActive(false);
    window.addEventListener('scroll', close, true);
    window.addEventListener('resize', close);
    return () => {
      window.removeEventListener('scroll', close, true);
      window.removeEventListener('resize', close);
    };
  }, [active]);

  // Click-outside that respects both the trigger and the portal'd
  // popup. Identical pattern to the Menu component.
  useEffect(() => {
    if (!active) return;
    const onPointerDown = (e: Event) => {
      const target = e.target as Node | null;
      if (!target) return;
      if (
        triggerRef.current?.contains(target) ||
        popupRef.current?.contains(target)
      ) {
        return;
      }
      setActive(false);
    };
    const onBlur = () => setActive(false);
    document.addEventListener('pointerdown', onPointerDown);
    window.addEventListener('blur', onBlur);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown);
      window.removeEventListener('blur', onBlur);
    };
  }, [active]);

  return (
    <div
      ref={triggerRef}
      className={classnames(styles['select-list-container'], className, {
        [styles.active]: active,
      })}
      data-testid={testId}
    >
      {children}
      <Button
        className={styles['select-list-trigger']}
        onClick={handleClick}
        testId={`${testId}-trigger`}
        type="plain"
      >
        <Icon name="down"></Icon>
      </Button>
      {active &&
        data.length > 0 &&
        typeof document !== 'undefined' &&
        createPortal(
          <div ref={popupRef} style={popupStyle}>
            <SelectPopup
              active
              value={value}
              data={data}
              onChange={handleChange}
              position={position}
              testId={`${testId}-popup`}
            />
          </div>,
          document.body,
        )}
    </div>
  );
});

SelectList.displayName = 'SelectList';
