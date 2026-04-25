import React, { useMemo, memo, useCallback } from 'react';
/* Icon strategy: Lucide (react-icons/lu) is the primary set -- its
 * outlined glyphs match Excel's modern ribbon style closely. Material
 * (react-icons/md) is kept only where Lucide lacks a direct match
 * (text-color A bar, text-size A^/Av step buttons). */
import {
  LuScissors,
  LuCopy,
  LuClipboard,
  LuEraser,
  LuDollarSign,
  LuPercent,
  LuBold,
  LuItalic,
  LuStrikethrough,
  LuPaintBucket,
  LuAlignLeft,
  LuAlignCenter,
  LuAlignRight,
  LuAlignStartHorizontal,
  LuAlignCenterHorizontal,
  LuAlignEndHorizontal,
  LuWrapText,
  LuTableCellsMerge,
  LuSigma,
  LuFilter,
  LuSquarePlus,
  LuSquareMinus,
  LuSettings,
  LuChevronDown,
} from 'react-icons/lu';
import {
  MdFormatColorText,
  MdTextIncrease,
  MdTextDecrease,
} from 'react-icons/md';
import {
  Button,
  Select,
  ColorPicker,
  SelectList,
  Menu,
  MenuItem,
  info,
  toast,
} from '../../components';
import {
  FONT_SIZE_LIST,
  QUERY_ALL_LOCAL_FONT,
  LOCAL_FONT_KEY,
  DEFAULT_FORMAT_CODE,
  DEFAULT_FONT_SIZE,
  adjustDecimalFormat,
  intToColumnName,
} from '../../util';
import {
  EUnderLine,
  OptionItem,
  EHorizontalAlign,
  EVerticalAlign,
  EditorStatus,
  IRange,
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
    const sheetList = useCoreStore((s) => s.sheetList);
    const currentSheetId = useCoreStore((s) => s.currentSheetId);
    const cellStyle = useStyleStore();
    const currentTabColor = useMemo(
      () =>
        sheetList.find((v) => v.sheetId === currentSheetId)?.tabColor || '',
      [sheetList, currentSheetId],
    );

    const fontFamilyList = useCoreStore((s) => s.fontFamilies);

    const fillStyle = useMemo(
      () => ({ backgroundColor: cellStyle.fillColor || '#ffff00' }),
      [cellStyle.fillColor],
    );
    const fontStyle = useMemo(
      () => ({ backgroundColor: cellStyle.fontColor || '#c00000' }),
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
    const autoSum = useCallback(() => {
      const { range } = controller.getActiveRange();
      const { row, col, rowCount, colCount, sheetId } = range;
      const sId = sheetId ?? '';

      // Single-cell case (no multi-cell selection): walk up the column
      // from row-1 collecting the contiguous run of numeric cells, then
      // write =SUM(start:end) into the active cell. Mirrors Excel's
      // single-cell AutoSum heuristic.
      if (rowCount <= 1 && colCount <= 1) {
        let endRow = row - 1;
        let startRow = endRow;
        let foundNumeric = false;
        for (let r = endRow; r >= 0; r--) {
          const cell = controller.getCell({
            row: r,
            col,
            rowCount: 1,
            colCount: 1,
            sheetId: sId,
          });
          const v = cell?.value;
          if (typeof v === 'number') {
            if (!foundNumeric) {
              endRow = r;
            }
            foundNumeric = true;
            startRow = r;
          } else if (foundNumeric) {
            break;
          } else if (v !== undefined && v !== '') {
            break;
          }
        }
        const colLetter = intToColumnName(col);
        const formula = foundNumeric
          ? `=SUM(${colLetter}${startRow + 1}:${colLetter}${endRow + 1})`
          : '=SUM()';
        const target: IRange = {
          row,
          col,
          rowCount: 1,
          colCount: 1,
          sheetId: sId,
        };
        controller.setCellValue(formula, target);
        return;
      }

      // Multi-cell case. The last row of the selection is treated as
      // the totals row -- each column gets =SUM(<col>top:<col>bottom-1)
      // written into <col>bottom. For a single-row, multi-col selection,
      // the analogue is "rightmost cell = SUM of the rest of the row".
      // Wrapped in transaction() so the whole fill is one undo step.
      controller.transaction(() => {
        if (rowCount === 1) {
          const totalsCol = col + colCount - 1;
          const startLetter = intToColumnName(col);
          const endLetter = intToColumnName(totalsCol - 1);
          const formula = `=SUM(${startLetter}${row + 1}:${endLetter}${row + 1})`;
          controller.setCellValue(formula, {
            row,
            col: totalsCol,
            rowCount: 1,
            colCount: 1,
            sheetId: sId,
          });
          return;
        }
        const totalsRow = row + rowCount - 1;
        const sourceEnd = totalsRow - 1;
        for (let c = col; c < col + colCount; c++) {
          const letter = intToColumnName(c);
          const formula = `=SUM(${letter}${row + 1}:${letter}${sourceEnd + 1})`;
          controller.setCellValue(formula, {
            row: totalsRow,
            col: c,
            rowCount: 1,
            colCount: 1,
            sheetId: sId,
          });
        }
      });
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

    const insertSheetRows = useCallback(() => {
      const { row, rowCount } = controller.getActiveRange().range;
      controller.addRow(row, rowCount, true);
    }, []);
    const insertSheetCols = useCallback(() => {
      const { col, colCount } = controller.getActiveRange().range;
      controller.addCol(col, colCount);
    }, []);
    const insertSheet = useCallback(() => {
      controller.addSheet();
    }, []);
    const deleteCells = useCallback(() => {
      controller.deleteCell(controller.getActiveRange().range);
    }, []);
    const deleteSheetRows = useCallback(() => {
      const { row, rowCount } = controller.getActiveRange().range;
      controller.deleteRow(row, rowCount);
    }, []);
    const deleteSheetCols = useCallback(() => {
      const { col, colCount } = controller.getActiveRange().range;
      controller.deleteCol(col, colCount);
    }, []);
    const deleteCurrentSheet = useCallback(() => {
      controller.deleteSheet();
    }, []);
    const hideRows = useCallback(() => {
      const { row, rowCount } = controller.getActiveRange().range;
      controller.hideRow(row, rowCount);
    }, []);
    const unhideRows = useCallback(() => {
      const { row, rowCount } = controller.getActiveRange().range;
      controller.unhideRow(row, rowCount);
    }, []);
    const hideCols = useCallback(() => {
      const { col, colCount } = controller.getActiveRange().range;
      controller.hideCol(col, colCount);
    }, []);
    const unhideCols = useCallback(() => {
      const { col, colCount } = controller.getActiveRange().range;
      controller.unhideCol(col, colCount);
    }, []);
    const hideCurrentSheet = useCallback(() => {
      controller.hideSheet();
    }, []);
    const promptDimension = useCallback(
      (
        title: string,
        initial: number,
        apply: (value: number) => void,
      ) => {
        let value = initial;
        info({
          visible: true,
          title,
          testId: 'toolbar-cells-dimension-dialog',
          children: (
            <input
              type="number"
              min={0}
              defaultValue={initial}
              style={{ width: '200px' }}
              onChange={(e) => {
                const t = parseInt(e.target.value, 10);
                if (!isNaN(t) && t >= 0) {
                  value = t;
                }
                e.stopPropagation();
              }}
              data-testid="toolbar-cells-dimension-input"
            />
          ),
          onOk: () => {
            if (value < 0) {
              toast.error(i18n.t('greater-than-zero'));
              return;
            }
            apply(value);
          },
        });
      },
      [],
    );
    const promptRowHeight = useCallback(() => {
      const { row, rowCount } = controller.getActiveRange().range;
      const current = controller.getRowHeight(row);
      promptDimension(i18n.t('row-height'), current, (value) => {
        controller.transaction(() => {
          for (let i = 0; i < rowCount; i++) {
            controller.setRowHeight(row + i, value);
          }
        });
      });
    }, [promptDimension]);
    const promptColumnWidth = useCallback(() => {
      const { col, colCount } = controller.getActiveRange().range;
      const current = controller.getColWidth(col);
      promptDimension(i18n.t('column-width'), current, (value) => {
        controller.transaction(() => {
          for (let i = 0; i < colCount; i++) {
            controller.setColWidth(col + i, value);
          }
        });
      });
    }, [promptDimension]);
    const promptRenameSheet = useCallback(() => {
      let value = '';
      info({
        visible: true,
        title: i18n.t('rename-sheet'),
        testId: 'toolbar-cells-rename-dialog',
        children: (
          <input
            type="text"
            style={{ width: '200px' }}
            onChange={(e) => {
              value = e.target.value;
              e.stopPropagation();
            }}
            data-testid="toolbar-cells-rename-input"
          />
        ),
        onOk: () => {
          if (!value.trim()) {
            toast.error(i18n.t('the-value-cannot-be-empty'));
            return;
          }
          controller.renameSheet(value.trim());
        },
      });
    }, []);
    const setSheetTabColor = useCallback((color: string) => {
      controller.updateSheetInfo({ tabColor: color });
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
                <LuClipboard {...iconProps} />
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
                className={styles['row-button']}
              >
                <LuScissors {...smallIconProps} />
                <span className={styles['row-label']}>{i18n.t('cut')}</span>
              </Button>
              <Button
                type="toolbar"
                onClick={copy}
                testId="toolbar-copy"
                title={`${i18n.t('copy')} (Ctrl+C)`}
                className={styles['row-button']}
              >
                <LuCopy {...smallIconProps} />
                <span className={styles['row-label']}>{i18n.t('copy')}</span>
              </Button>
              <Button
                type="toolbar"
                onClick={clearFormatting}
                testId="toolbar-clear-format"
                title={i18n.t('clear-format')}
                className={styles['row-button']}
              >
                <LuEraser {...smallIconProps} />
                <span className={styles['row-label']}>{i18n.t('clear-format')}</span>
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
                <LuBold {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.isItalic}
                onClick={toggleItalic}
                testId="toolbar-italic"
                title={i18n.t('toolbar-italic')}
                className={styles['small-button']}
              >
                <LuItalic {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.isStrike}
                onClick={toggleStrike}
                testId="toolbar-strike"
                title={i18n.t('toolbar-strike')}
                className={styles['small-button']}
              >
                <LuStrikethrough {...smallIconProps} />
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
                  testId="toolbar-fill-color"
                  title={i18n.t('toolbar-fill-color')}
                  className={styles['color-button']}
                >
                  <LuPaintBucket {...smallIconProps} />
                  <span
                    className={styles['color-strip']}
                    style={fillStyle}
                  />
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
                  testId="toolbar-font-color"
                  title={i18n.t('toolbar-text-color')}
                  className={styles['color-button']}
                >
                  <MdFormatColorText {...smallIconProps} />
                  <span
                    className={styles['color-strip']}
                    style={fontStyle}
                  />
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
                <LuAlignStartHorizontal {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.verticalAlign === EVerticalAlign.MIDDLE}
                onClick={verticalMiddle}
                testId="toolbar-vertical-middle"
                title={i18n.t('toolbar-align-middle')}
                className={styles['small-button']}
              >
                <LuAlignCenterHorizontal {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.verticalAlign === EVerticalAlign.BOTTOM}
                onClick={verticalBottom}
                testId="toolbar-vertical-bottom"
                title={i18n.t('toolbar-align-bottom')}
                className={styles['small-button']}
              >
                <LuAlignEndHorizontal {...smallIconProps} />
              </Button>
              <span className={styles['mini-divider']} />
              <Button
                type="toolbar"
                active={cellStyle.isWrapText}
                onClick={toggleWrapText}
                testId="toolbar-wrap-text"
                title={i18n.t('wrap-text')}
                className={styles['wide-button']}
              >
                <LuWrapText {...smallIconProps} />
                <span className={styles['wide-label']}>
                  {i18n.t('wrap-text')}
                </span>
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
                <LuAlignLeft {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.horizontalAlign === EHorizontalAlign.CENTER}
                onClick={horizontalCenter}
                testId="toolbar-horizontal-center"
                title={i18n.t('toolbar-align-center')}
                className={styles['small-button']}
              >
                <LuAlignCenter {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                active={cellStyle.horizontalAlign === EHorizontalAlign.RIGHT}
                onClick={horizontalRight}
                testId="toolbar-horizontal-right"
                title={i18n.t('toolbar-align-right')}
                className={styles['small-button']}
              >
                <LuAlignRight {...smallIconProps} />
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
                  className={styles['wide-button']}
                >
                  <LuTableCellsMerge {...smallIconProps} />
                  <span className={styles['wide-label']}>
                    {i18n.t('merge-and-center')}
                  </span>
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
                <LuDollarSign {...smallIconProps} />
              </Button>
              <Button
                type="toolbar"
                onClick={applyPercent}
                testId="toolbar-percent"
                title={i18n.t('format-as-percent')}
                className={styles['small-button']}
              >
                <LuPercent {...smallIconProps} />
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

        <Section label={i18n.t('section-cells')}>
          <Row>
            <Menu
              position="bottom"
              testId="toolbar-cells-insert"
              isPlain
              className={styles['cells-menu']}
              label={
                <span className={styles['tall-menu-trigger']}>
                  <LuSquarePlus className={styles['tall-menu-icon']} />
                  <span className={styles['tall-menu-label']}>
                    {i18n.t('insert')}
                  </span>
                  <LuChevronDown className={styles['tall-menu-chevron']} />
                </span>
              }
            >
              <MenuItem onClick={insertSheetRows} testId="toolbar-cells-insert-rows">
                {i18n.t('insert-rows')}
              </MenuItem>
              <MenuItem onClick={insertSheetCols} testId="toolbar-cells-insert-cols">
                {i18n.t('insert-columns')}
              </MenuItem>
              <MenuItem onClick={insertSheet} testId="toolbar-cells-insert-sheet">
                {i18n.t('insert-sheet')}
              </MenuItem>
            </Menu>
            <Menu
              position="bottom"
              testId="toolbar-cells-delete"
              isPlain
              className={styles['cells-menu']}
              label={
                <span className={styles['tall-menu-trigger']}>
                  <LuSquareMinus className={styles['tall-menu-icon']} />
                  <span className={styles['tall-menu-label']}>
                    {i18n.t('delete')}
                  </span>
                  <LuChevronDown className={styles['tall-menu-chevron']} />
                </span>
              }
            >
              <MenuItem onClick={deleteCells} testId="toolbar-cells-delete-cells">
                {i18n.t('delete-cells')}
              </MenuItem>
              <MenuItem onClick={deleteSheetRows} testId="toolbar-cells-delete-rows">
                {i18n.t('delete-rows')}
              </MenuItem>
              <MenuItem onClick={deleteSheetCols} testId="toolbar-cells-delete-cols">
                {i18n.t('delete-columns')}
              </MenuItem>
              <MenuItem onClick={deleteCurrentSheet} testId="toolbar-cells-delete-sheet">
                {i18n.t('delete-sheet')}
              </MenuItem>
            </Menu>
            <Menu
              position="bottom"
              testId="toolbar-cells-format"
              isPlain
              className={styles['cells-menu']}
              label={
                <span className={styles['tall-menu-trigger']}>
                  <LuSettings className={styles['tall-menu-icon']} />
                  <span className={styles['tall-menu-label']}>
                    {i18n.t('format')}
                  </span>
                  <LuChevronDown className={styles['tall-menu-chevron']} />
                </span>
              }
            >
              <MenuItem onClick={promptRowHeight} testId="toolbar-cells-row-height">
                {i18n.t('row-height')}
              </MenuItem>
              <MenuItem onClick={promptColumnWidth} testId="toolbar-cells-col-width">
                {i18n.t('column-width')}
              </MenuItem>
              <MenuItem onClick={hideRows} testId="toolbar-cells-hide-rows">
                {i18n.t('hide-rows')}
              </MenuItem>
              <MenuItem onClick={unhideRows} testId="toolbar-cells-unhide-rows">
                {i18n.t('unhide-rows')}
              </MenuItem>
              <MenuItem onClick={hideCols} testId="toolbar-cells-hide-cols">
                {i18n.t('hide-columns')}
              </MenuItem>
              <MenuItem onClick={unhideCols} testId="toolbar-cells-unhide-cols">
                {i18n.t('unhide-columns')}
              </MenuItem>
              <MenuItem onClick={hideCurrentSheet} testId="toolbar-cells-hide-sheet">
                {i18n.t('hide-sheet')}
              </MenuItem>
              <MenuItem onClick={promptRenameSheet} testId="toolbar-cells-rename-sheet">
                {i18n.t('rename-sheet')}
              </MenuItem>
              <MenuItem testId="toolbar-cells-tab-color">
                <ColorPicker
                  color={currentTabColor}
                  onChange={setSheetTabColor}
                  testId="toolbar-cells-tab-color-picker"
                >
                  <span>{i18n.t('tab-color')}</span>
                </ColorPicker>
              </MenuItem>
            </Menu>
          </Row>
        </Section>

        <Section label={i18n.t('section-insert')}>
          <Row>
            <Button
              type="toolbar"
              onClick={autoSum}
              testId="toolbar-autosum"
              title={i18n.t('autosum')}
            >
              <LuSigma {...iconProps} />
            </Button>
            <Button
              type="toolbar"
              onClick={insertFunction}
              testId="toolbar-fx"
              title={i18n.t('insert-function')}
            >
              <span className={styles.fx}>fx</span>
            </Button>
            <Button
              type="toolbar"
              active={isFilter}
              onClick={handleFilter}
              testId="toolbar-filter"
              title={i18n.t('filter')}
            >
              <LuFilter {...iconProps} />
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
