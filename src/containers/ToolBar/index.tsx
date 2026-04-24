import React, { useMemo, memo, useCallback } from 'react';
import {
  MdUndo,
  MdRedo,
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

type SectionProps = React.PropsWithChildren<{ label: string }>;
const Section: React.FunctionComponent<SectionProps> = ({ label, children }) => (
  <div className={styles.section}>
    <div className={styles['section-row']}>{children}</div>
    <span className={styles['section-label']}>{label}</span>
  </div>
);

const iconProps = { className: styles['icon-svg'] };

export const ToolbarContainer: React.FunctionComponent<React.PropsWithChildren> =
  memo(({ children }) => {
    const { controller } = useExcel();
    const canRedo = useCoreStore((s) => s.canRedo);
    const canUndo = useCoreStore((s) => s.canUndo);
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
    const undo = useCallback(() => controller.undo(), []);
    const redo = useCallback(() => controller.redo(), []);
    const copy = useCallback(() => controller.copy(), []);
    const cut = useCallback(() => controller.cut(), []);
    const paste = useCallback(() => controller.paste(), []);
    const setFontSize = useCallback((value: string | number) => {
      controller.updateCellStyle(
        { fontSize: Number(value) },
        controller.getActiveRange().range,
      );
    }, []);
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
        <Section label={i18n.t('section-history')}>
          <Button
            type="toolbar"
            disabled={!canUndo}
            onClick={undo}
            testId="toolbar-undo"
            title={i18n.t('toolbar-undo')}
          >
            <MdUndo {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            disabled={!canRedo}
            onClick={redo}
            testId="toolbar-redo"
            title={i18n.t('toolbar-redo')}
          >
            <MdRedo {...iconProps} />
          </Button>
        </Section>

        <Section label={i18n.t('section-clipboard')}>
          <Button
            type="toolbar"
            onClick={copy}
            testId="toolbar-copy"
            title={`${i18n.t('copy')} (Ctrl+C)`}
          >
            <MdContentCopy {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            onClick={cut}
            testId="toolbar-cut"
            title={`${i18n.t('cut')} (Ctrl+X)`}
          >
            <MdContentCut {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            onClick={paste}
            testId="toolbar-paste"
            title={`${i18n.t('paste')} (Ctrl+V)`}
          >
            <MdContentPaste {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            onClick={clearFormatting}
            testId="toolbar-clear-format"
            title={i18n.t('clear-format')}
          >
            <MdFormatClear {...iconProps} />
          </Button>
        </Section>

        <Section label={i18n.t('section-number')}>
          <Button
            type="toolbar"
            onClick={applyCurrency}
            testId="toolbar-currency"
            title={i18n.t('format-as-currency')}
          >
            <MdAttachMoney {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            onClick={applyPercent}
            testId="toolbar-percent"
            title={i18n.t('format-as-percent')}
          >
            <MdPercent {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            onClick={decreaseDecimal}
            testId="toolbar-decimal-decrease"
            title={i18n.t('decrease-decimal')}
          >
            <span className={styles.glyph}>.0&larr;</span>
          </Button>
          <Button
            type="toolbar"
            onClick={increaseDecimal}
            testId="toolbar-decimal-increase"
            title={i18n.t('increase-decimal')}
          >
            <span className={styles.glyph}>.00&rarr;</span>
          </Button>
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
        </Section>

        <Section label={i18n.t('section-font')}>
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
        </Section>

        <Section label={i18n.t('section-style')}>
          <Button
            type="toolbar"
            active={cellStyle.isBold}
            onClick={toggleBold}
            testId="toolbar-bold"
            title={i18n.t('toolbar-bold')}
          >
            <MdFormatBold {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            active={cellStyle.isItalic}
            onClick={toggleItalic}
            testId="toolbar-italic"
            title={i18n.t('toolbar-italic')}
          >
            <MdFormatItalic {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            active={cellStyle.isStrike}
            onClick={toggleStrike}
            testId="toolbar-strike"
            title={i18n.t('toolbar-strike')}
          >
            <MdStrikethroughS {...iconProps} />
          </Button>
          <Select
            data={underlineOptionList}
            value={cellStyle.underline}
            title={i18n.t('toolbar-underline')}
            onChange={setUnderline}
            testId="toolbar-underline"
          />
        </Section>

        <Section label={i18n.t('section-format')}>
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
            >
              <MdFormatColorText {...iconProps} />
            </Button>
          </ColorPicker>
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
            >
              <MdFormatColorFill {...iconProps} />
            </Button>
          </ColorPicker>
          <BorderToolBar />
        </Section>

        <Section label={i18n.t('section-align')}>
          <Button
            type="toolbar"
            active={cellStyle.horizontalAlign === EHorizontalAlign.LEFT}
            onClick={horizontalLeft}
            testId="toolbar-horizontal-left"
            title={i18n.t('toolbar-align-left')}
          >
            <MdFormatAlignLeft {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            active={cellStyle.horizontalAlign === EHorizontalAlign.CENTER}
            onClick={horizontalCenter}
            testId="toolbar-horizontal-center"
            title={i18n.t('toolbar-align-center')}
          >
            <MdFormatAlignCenter {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            active={cellStyle.horizontalAlign === EHorizontalAlign.RIGHT}
            onClick={horizontalRight}
            testId="toolbar-horizontal-right"
            title={i18n.t('toolbar-align-right')}
          >
            <MdFormatAlignRight {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            active={cellStyle.verticalAlign === EVerticalAlign.TOP}
            onClick={verticalTop}
            testId="toolbar-vertical-top"
            title={i18n.t('toolbar-align-top')}
          >
            <MdVerticalAlignTop {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            active={cellStyle.verticalAlign === EVerticalAlign.MIDDLE}
            onClick={verticalMiddle}
            testId="toolbar-vertical-middle"
            title={i18n.t('toolbar-align-middle')}
          >
            <MdVerticalAlignCenter {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            active={cellStyle.verticalAlign === EVerticalAlign.BOTTOM}
            onClick={verticalBottom}
            testId="toolbar-vertical-bottom"
            title={i18n.t('toolbar-align-bottom')}
          >
            <MdVerticalAlignBottom {...iconProps} />
          </Button>
          <Button
            type="toolbar"
            active={cellStyle.isWrapText}
            onClick={toggleWrapText}
            testId="toolbar-wrap-text"
            title={i18n.t('wrap-text')}
          >
            <MdWrapText {...iconProps} />
          </Button>
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
            >
              <MdCallMerge {...iconProps} />
            </Button>
          </SelectList>
        </Section>

        <Section label={i18n.t('section-insert')}>
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
        </Section>
      </div>
    );
  });

ToolbarContainer.displayName = 'ToolbarContainer';

export default ToolbarContainer;
