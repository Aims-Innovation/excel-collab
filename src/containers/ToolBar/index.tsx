import React, { useMemo, memo, useCallback } from 'react';
import {
  MdContentCopy,
  MdContentCut,
  MdContentPaste,
  MdFormatClear,
  MdAttachMoney,
  MdPercent,
  MdFormatBold,
  MdFormatItalic,
  MdStrikethroughS,
  MdFormatColorText,
  MdFormatColorFill,
  MdFormatAlignLeft,
  MdFormatAlignCenter,
  MdFormatAlignRight,
  MdVerticalAlignTop,
  MdVerticalAlignCenter,
  MdVerticalAlignBottom,
  MdWrapText,
  MdCallMerge,
  MdFunctions,
  MdFilterAlt,
  MdTextIncrease,
  MdTextDecrease,
} from 'react-icons/md';
import {
  Button,
  Select,
  ColorPicker,
  SelectList,
} from '../../components';
import {
  FONT_SIZE_LIST,
  QUERY_ALL_LOCAL_FONT,
  LOCAL_FONT_KEY,
  DEFAULT_FORMAT_CODE,
  DEFAULT_FONT_SIZE,
  adjustDecimalFormat,
} from '../../util';
import {
  EUnderLine,
  OptionItem,
  EHorizontalAlign,
  EVerticalAlign,
  EditorStatus,
} from '../../types';
import styles from './index.module.css';
import { useStyleStore, useCoreStore, useExcel } from '../../containers/store';
import { InsertFloatingPicture, InsertChart } from '../FloatElement/Toolbar';
import i18n from '../../i18n';
import { BorderToolBar } from './Border';
import { isSupportFontFamily } from '../canvas/isSupportFontFamily';
import {
  numberFormatOptionList,
  underlineOptionList,
  mergeOptionList,
} from './constant';

const CURRENCY_FORMAT = '"$"#,##0.00_);[Red]("$"#,##0.00)';
const PERCENT_FORMAT = '0.00%';

type SectionProps = React.PropsWithChildren<{
  label: string;
  className?: string;
}>;
const Section: React.FunctionComponent<SectionProps> = ({
  label,
  className,
  children,
}) => (
  <div
    className={`${styles.section}${className ? ' ' + className : ''}`}
  >
    <div className={styles['section-body']}>{children}</div>
    <span className={styles['section-label']}>{label}</span>
  </div>
);

const Stack: React.FunctionComponent<React.PropsWithChildren> = ({
  children,
}) => <div className={styles['section-rows']}>{children}</div>;

const Row: React.FunctionComponent<React.PropsWithChildren> = ({
  children,
}) => <div className={styles['section-row']}>{children}</div>;

const iconProps = { className: styles['icon-svg'] };
const smallIconProps = { className: styles['icon-svg-sm'] };

export const ToolbarContainer: React.FunctionComponent<React.PropsWithChildren> =
  memo(({ children }) => {
    const { controller } = useExcel();
    const isFilter = useCoreStore((s) => s.isFilter);
    const fontFamilies = useCoreStore((s) => s.fontFamilies);
    const setFontFamilies = useCoreStore((s) => s.setFontFamilies);
    const setEditorStatus = useCoreStore((s) => s.setEditorStatus);
    const cellStyle = useStyleStore();

    const fontFamilyList = useCoreStore((s) => s.fontFamilies);

    const fillStyle = useMemo(
      () => ({ color: cellStyle.fillColor }),
      [cellStyle.fillColor],
    );
    const fontStyle = useMemo(
      () => ({ color: cellStyle.fontColor }),
      [cellStyle.fontColor],
    );
    const [numberFormatLabel, numberFormatValue] = useMemo(() => {
      let item: OptionItem = numberFormatOptionList[0];
      if (cellStyle.numberFormat) {
        const t = numberFormatOptionList.find(
          (v) => v.value === cellStyle.numberFormat,
        );
        if (t) {
          item = t;
        } else {
          item = numberFormatOptionList[numberFormatOptionList.length - 1];
        }
      }
      return [item.label, String(item.value)];
    }, [cellStyle.numberFormat]);
    const getItemStyle = useCallback(
      (value: string | number): React.CSSProperties => ({
        fontFamily: String(value),
      }),
      [],
    );
    const handleFontFamilyChange = useCallback((value: string | number) => {
      if (
        String(value) === QUERY_ALL_LOCAL_FONT &&
        typeof window.queryLocalFonts === 'function'
      ) {
        window.queryLocalFonts().then((list) => {
          let fontList = list.map((v) => v.fullName);
          fontList = Array.from(new Set(fontList)).filter((v) =>
            isSupportFontFamily(v),
          );
          fontList.sort((a, b) => a.localeCompare(b));
          const l = fontList.map((v) => ({
            label: v,
            value: v,
            disabled: false,
          }));
          if (fontList.length > 0) {
            setFontFamilies(l);
            localStorage.setItem(LOCAL_FONT_KEY, JSON.stringify(fontList));
          } else {
            setFontFamilies(
              fontFamilies.filter((v) => v.value !== QUERY_ALL_LOCAL_FONT),
            );
          }
        });
      } else {
        controller.updateCellStyle(
          { fontFamily: String(value) },
          controller.getActiveRange().range,
        );
      }
    }, []);
    const copy = useCallback(() => controller.copy(), []);
    const cut = useCallback(() => controller.cut(), []);
    const paste = useCallback(() => controller.paste(), []);
    const setFontSize = useCallback((value: string | number) => {
      controller.updateCellStyle(
        { fontSize: Number(value) },
        controller.getActiveRange().range,
      );
    }, []);
    const stepFontSize = useCallback(
      (delta: 1 | -1) => {
        const current = Number(cellStyle.fontSize) || DEFAULT_FONT_SIZE;
        const sizes = FONT_SIZE_LIST;
        let next = current;
        if (delta > 0) {
          const found = sizes.find((s) => s > current);
          next = found ?? sizes[sizes.length - 1];
        } else {
          for (let i = sizes.length - 1; i >= 0; i--) {
            if (sizes[i] < current) {
              next = sizes[i];
              break;
            }
          }
        }
        if (next !== current) {
          controller.updateCellStyle(
            { fontSize: next },
            controller.getActiveRange().range,
          );
        }
      },
      [cellStyle.fontSize],
    );
    const increaseFontSize = useCallback(() => stepFontSize(1), [stepFontSize]);
    const decreaseFontSize = useCallback(
      () => stepFontSize(-1),
      [stepFontSize],
    );
    const toggleBold = useCallback(() => {
      controller.updateCellStyle(
        { isBold: !cellStyle.isBold },
        controller.getActiveRange().range,
      );
    }, [cellStyle.isBold]);
    const toggleItalic = useCallback(() => {
      controller.updateCellStyle(
        { isItalic: !cellStyle.isItalic },
        controller.getActiveRange().range,
      );
    }, [cellStyle.isItalic]);
    const toggleStrike = useCallback(() => {
      controller.updateCellStyle(
        { isStrike: !cellStyle.isStrike },
        controller.getActiveRange().range,
      );
    }, [cellStyle.isStrike]);
    const setUnderline = useCallback((value: string | number) => {
      const t = Number(value);
      let underline = EUnderLine.NONE;
      if (t === EUnderLine.SINGLE) {
        underline = EUnderLine.SINGLE;
      } else if (t === EUnderLine.DOUBLE) {
        underline = EUnderLine.DOUBLE;
      }
      controller.updateCellStyle(
        { underline },
        controller.getActiveRange().range,
      );
    }, []);
    const setFillColor = useCallback((value: string) => {
      controller.updateCellStyle(
        { fillColor: value },
        controller.getActiveRange().range,
      );
    }, []);
    const setFontColor = useCallback((value: string) => {
      controller.updateCellStyle(
        { fontColor: value },
        controller.getActiveRange().range,
      );
    }, []);
    const toggleWrapText = useCallback(() => {
      controller.updateCellStyle(
        { isWrapText: !cellStyle.isWrapText },
        controller.getActiveRange().range,
      );
    }, [cellStyle.isWrapText]);
    const toggleMergeCell = useCallback(() => {
      const { range, isMerged } = controller.getActiveRange();
      if (isMerged) {
        controller.deleteMergeCell(range);
      } else {
        controller.addMergeCell(range);
      }
    }, []);
    const handleMergeCell = useCallback((value: string) => {
      if (!value) return;
      const { range, isMerged } = controller.getActiveRange();
      if (isMerged) {
        controller.deleteMergeCell(range);
      } else {
        controller.addMergeCell(range, Number(value));
      }
    }, []);
    const handleNumberFormat = useCallback((value: string) => {
      if (!value) return;
      controller.updateCellStyle(
        { numberFormat: value },
        controller.getActiveRange().range,
      );
    }, []);
    const applyCurrency = useCallback(() => {
      controller.updateCellStyle(
        { numberFormat: CURRENCY_FORMAT },
        controller.getActiveRange().range,
      );
    }, []);
    const applyPercent = useCallback(() => {
      controller.updateCellStyle(
        { numberFormat: PERCENT_FORMAT },
        controller.getActiveRange().range,
      );
    }, []);
    const decreaseDecimal = useCallback(() => {
      const next = adjustDecimalFormat(
        cellStyle.numberFormat || DEFAULT_FORMAT_CODE,
        -1,
      );
      controller.updateCellStyle(
        { numberFormat: next },
        controller.getActiveRange().range,
      );
    }, [cellStyle.numberFormat]);
    const increaseDecimal = useCallback(() => {
      const next = adjustDecimalFormat(
        cellStyle.numberFormat || DEFAULT_FORMAT_CODE,
        1,
      );
      controller.updateCellStyle(
        { numberFormat: next },
        controller.getActiveRange().range,
      );
    }, [cellStyle.numberFormat]);
    const clearFormatting = useCallback(() => {
      controller.updateCellStyle(
        {
          fontColor: '',
          fillColor: '',
          fontSize: DEFAULT_FONT_SIZE,
          fontFamily: '',
          verticalAlign: undefined,
          horizontalAlign: undefined,
          isWrapText: false,
          underline: EUnderLine.NONE,
          isItalic: false,
          isBold: false,
          isStrike: false,
          numberFormat: DEFAULT_FORMAT_CODE,
          borderLeft: undefined,
          borderRight: undefined,
          borderTop: undefined,
          borderBottom: undefined,
        },
        controller.getActiveRange().range,
      );
    }, []);
    const insertFunction = useCallback(() => {
      setEditorStatus(EditorStatus.EDIT_FORMULA_BAR);
    }, []);
    const horizontalLeft = useCallback(() => {
      controller.updateCellStyle(
        { horizontalAlign: EHorizontalAlign.LEFT },
        controller.getActiveRange().range,
      );
    }, []);
    const horizontalCenter = useCallback(() => {
      controller.updateCellStyle(
        { horizontalAlign: EHorizontalAlign.CENTER },
        controller.getActiveRange().range,
      );
    }, []);
    const horizontalRight = useCallback(() => {
      controller.updateCellStyle(
        { horizontalAlign: EHorizontalAlign.RIGHT },
        controller.getActiveRange().range,
      );
    }, []);
    const verticalTop = useCallback(() => {
      controller.updateCellStyle(
        { verticalAlign: EVerticalAlign.TOP },
        controller.getActiveRange().range,
      );
    }, []);
    const verticalMiddle = useCallback(() => {
      controller.updateCellStyle(
        { verticalAlign: EVerticalAlign.MIDDLE },
        controller.getActiveRange().range,
      );
    }, []);
    const verticalBottom = useCallback(() => {
      controller.updateCellStyle(
        { verticalAlign: EVerticalAlign.BOTTOM },
        controller.getActiveRange().range,
      );
    }, []);
    const handleFilter = useCallback(() => {
      const filter = controller.getFilter();
      if (filter) {
        controller.deleteFilter();
      } else {
        controller.addFilter(controller.getActiveRange().range);
      }
    }, []);
    return (
      <div className={styles['toolbar-wrapper']} data-testid="toolbar">
        <Section label={i18n.t('section-clipboard')}>
          <div className={styles['clipboard-split']}>
            <Button
              type="toolbar"
              onClick={paste}
              testId="toolbar-paste"
              title={`${i18n.t('paste')} (Ctrl+V)`}
              className={styles['big-button']}
            >
              <span className={styles['big-button-inner']}>
                <MdContentPaste {...iconProps} />
                <span className={styles['big-button-label']}>
                  {i18n.t('paste')}
                </span>
              </span>
            </Button>
            <div className={styles['stack-col']}>
              <Button
                type="toolbar"
                onClick={cut}
                testId="toolbar-cut"
                title={`${i18n.t('cut')} (Ctrl+X)`}
                className={styles['small-button']}
              >
                <MdContentCut {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                onClick={copy}
                testId="toolbar-copy"
                title={`${i18n.t('copy')} (Ctrl+C)`}
                className={styles['small-button']}
              >
                <MdContentCopy {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                onClick={clearFormatting}
                testId="toolbar-clear-format"
                title={i18n.t('clear-format')}
                className={styles['small-button']}
              >
                <MdFormatClear {...smallIconProps} />
              </Button>
            </div>
          </div>
        </Section>

        <Section label={i18n.t('section-font')}>
          <Stack>
            <Row>
              <Select
                data={fontFamilyList}
                value={cellStyle.fontFamily}
                getItemStyle={getItemStyle}
                onChange={handleFontFamilyChange}
                testId="toolbar-font-family"
                className={styles.fontFamily}
              />
              <Select
                data={FONT_SIZE_LIST}
                value={cellStyle.fontSize}
                onChange={setFontSize}
                testId="toolbar-font-size"
                className={styles.fontSize}
              />
              <Button
                type="toolbar"
                onClick={increaseFontSize}
                testId="toolbar-font-size-increase"
                title={i18n.t('increase-font-size')}
                className={styles['small-button']}
              >
                <MdTextIncrease {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                onClick={decreaseFontSize}
                testId="toolbar-font-size-decrease"
                title={i18n.t('decrease-font-size')}
                className={styles['small-button']}
              >
                <MdTextDecrease {...smallIconProps} />
              </Button>
            </Row>
            <Row>
              <Button
                type="toolbar"
                active={cellStyle.isBold}
                onClick={toggleBold}
                testId="toolbar-bold"
                title={i18n.t('toolbar-bold')}
                className={styles['small-button']}
              >
                <MdFormatBold {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.isItalic}
                onClick={toggleItalic}
                testId="toolbar-italic"
                title={i18n.t('toolbar-italic')}
                className={styles['small-button']}
              >
                <MdFormatItalic {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.isStrike}
                onClick={toggleStrike}
                testId="toolbar-strike"
                title={i18n.t('toolbar-strike')}
                className={styles['small-button']}
              >
                <MdStrikethroughS {...smallIconProps} />
              </Button>
              <Select
                data={underlineOptionList}
                value={cellStyle.underline}
                title={i18n.t('toolbar-underline')}
                onChange={setUnderline}
                testId="toolbar-underline"
                className={styles['underline-select']}
              />
              <span className={styles['mini-divider']} />
              <BorderToolBar />
              <ColorPicker
                key="fill-color"
                color={cellStyle.fillColor}
                onChange={setFillColor}
                testId="toolbar-fill-color"
              >
                <Button
                  type="toolbar"
                  style={fillStyle}
                  testId="toolbar-fill-color"
                  title={i18n.t('toolbar-fill-color')}
                  className={styles['small-button']}
                >
                  <MdFormatColorFill {...smallIconProps} />
                </Button>
              </ColorPicker>
              <ColorPicker
                key="font-color"
                color={cellStyle.fontColor}
                onChange={setFontColor}
                testId="toolbar-font-color"
              >
                <Button
                  type="toolbar"
                  style={fontStyle}
                  testId="toolbar-font-color"
                  title={i18n.t('toolbar-text-color')}
                  className={styles['small-button']}
                >
                  <MdFormatColorText {...smallIconProps} />
                </Button>
              </ColorPicker>
            </Row>
          </Stack>
        </Section>

        <Section label={i18n.t('section-align')}>
          <Stack>
            <Row>
              <Button
                type="toolbar"
                active={cellStyle.verticalAlign === EVerticalAlign.TOP}
                onClick={verticalTop}
                testId="toolbar-vertical-top"
                title={i18n.t('toolbar-align-top')}
                className={styles['small-button']}
              >
                <MdVerticalAlignTop {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.verticalAlign === EVerticalAlign.MIDDLE}
                onClick={verticalMiddle}
                testId="toolbar-vertical-middle"
                title={i18n.t('toolbar-align-middle')}
                className={styles['small-button']}
              >
                <MdVerticalAlignCenter {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.verticalAlign === EVerticalAlign.BOTTOM}
                onClick={verticalBottom}
                testId="toolbar-vertical-bottom"
                title={i18n.t('toolbar-align-bottom')}
                className={styles['small-button']}
              >
                <MdVerticalAlignBottom {...smallIconProps} />
              </Button>
              <span className={styles['mini-divider']} />
              <Button
                type="toolbar"
                active={cellStyle.isWrapText}
                onClick={toggleWrapText}
                testId="toolbar-wrap-text"
                title={i18n.t('wrap-text')}
                className={styles['small-button']}
              >
                <MdWrapText {...smallIconProps} />
              </Button>
            </Row>
            <Row>
              <Button
                type="toolbar"
                active={cellStyle.horizontalAlign === EHorizontalAlign.LEFT}
                onClick={horizontalLeft}
                testId="toolbar-horizontal-left"
                title={i18n.t('toolbar-align-left')}
                className={styles['small-button']}
              >
                <MdFormatAlignLeft {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.horizontalAlign === EHorizontalAlign.CENTER}
                onClick={horizontalCenter}
                testId="toolbar-horizontal-center"
                title={i18n.t('toolbar-align-center')}
                className={styles['small-button']}
              >
                <MdFormatAlignCenter {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.horizontalAlign === EHorizontalAlign.RIGHT}
                onClick={horizontalRight}
                testId="toolbar-horizontal-right"
                title={i18n.t('toolbar-align-right')}
                className={styles['small-button']}
              >
                <MdFormatAlignRight {...smallIconProps} />
              </Button>
              <span className={styles['mini-divider']} />
              <SelectList
                data={mergeOptionList}
                value={cellStyle.mergeType}
                onChange={handleMergeCell}
                className={styles['merge-cell']}
                testId="toolbar-merge-cell-select"
              >
                <Button
                  type="toolbar"
                  active={cellStyle.isMergeCell}
                  onClick={toggleMergeCell}
                  testId="toolbar-merge-cell"
                  title={i18n.t('merge-and-center')}
                  className={styles['small-button']}
                >
                  <MdCallMerge {...smallIconProps} />
                </Button>
              </SelectList>
            </Row>
          </Stack>
        </Section>

        <Section label={i18n.t('section-number')}>
          <Stack>
            <Row>
              <SelectList
                data={numberFormatOptionList}
                value={numberFormatValue}
                onChange={handleNumberFormat}
                className={styles['number-format']}
                testId="toolbar-number-format"
              >
                <div
                  className={styles['number-format-value']}
                  data-testid="toolbar-number-format-value"
                  title={numberFormatLabel}
                >
                  {numberFormatLabel}
                </div>
              </SelectList>
            </Row>
            <Row>
              <Button
                type="toolbar"
                onClick={applyCurrency}
                testId="toolbar-currency"
                title={i18n.t('format-as-currency')}
                className={styles['small-button']}
              >
                <MdAttachMoney {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                onClick={applyPercent}
                testId="toolbar-percent"
                title={i18n.t('format-as-percent')}
                className={styles['small-button']}
              >
                <MdPercent {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                onClick={decreaseDecimal}
                testId="toolbar-decimal-decrease"
                title={i18n.t('decrease-decimal')}
                className={styles['small-button']}
              >
                <span className={styles.glyph}>.0&larr;</span>
              </Button>
              <Button
                type="toolbar"
                onClick={increaseDecimal}
                testId="toolbar-decimal-increase"
                title={i18n.t('increase-decimal')}
                className={styles['small-button']}
              >
                <span className={styles.glyph}>.00&rarr;</span>
              </Button>
            </Row>
          </Stack>
        </Section>

        <Section label={i18n.t('section-insert')}>
          <Row>
            <Button
              type="toolbar"
              onClick={insertFunction}
              testId="toolbar-fx"
              title={i18n.t('insert-function')}
            >
              <MdFunctions {...iconProps} />
            </Button>
            <Button
              type="toolbar"
              active={isFilter}
              onClick={handleFilter}
              testId="toolbar-filter"
              title={i18n.t('filter')}
            >
              <MdFilterAlt {...iconProps} />
            </Button>
            <InsertFloatingPicture />
            <InsertChart />
            {children}
          </Row>
        </Section>
      </div>
    );
  });

ToolbarContainer.displayName = 'ToolbarContainer';

export default ToolbarContainer;
