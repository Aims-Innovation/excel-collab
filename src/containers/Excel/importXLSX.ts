import {
  ModelJSON,
  WorksheetType,
  ModelCellType,
  StyleType,
  EHorizontalAlign,
  EVerticalAlign,
  EUnderLine,
  DrawingElement,
  IWindowSize,
  BorderItem,
  BorderType,
} from '../../types';
import {
  get,
  parseReference,
  CELL_HEIGHT,
  CELL_WIDTH,
  XLSX_MAX_ROW_COUNT,
  XLSX_MAX_COL_COUNT,
  getCustomWidthOrHeightKey,
  IMAGE_TYPE_MAP,
  SheetRange,
  mergeRange,
  FORMULA_PREFIX,
  getImageSize,
  NUMBER_FORMAT_LIST,
  BORDER_TYPE_MAP,
  getWorksheetKey,
} from '../../util';

const COMMON_PREFIX = 'xl';
const STYLE_PATH = 'xl/styles.xml';
const WORKBOOK_PATH = 'xl/workbook.xml';
const WORKBOOK_RELATION_PATH = 'xl/_rels/workbook.xml.rels';
const THEME_PATH = 'xl/theme/theme1.xml';
const SHARED_STRINGS = 'xl/sharedStrings.xml';
const textKey = '#text';
const DRAWING_PREFIX_KEY = 'xl/drawings/';
const DRAWING_FLAG = '../drawings/';

const CUSTOM_WIdTH_RADIO = 8;

function convertFromPt(value: string) {
  const val = parseFloat(value);
  if (isNaN(val)) {
    return 0;
  }
  return Math.floor(val * CUSTOM_WIdTH_RADIO);
}

export function convertToPt(num: number) {
  return Math.floor(num / CUSTOM_WIdTH_RADIO);
}

export const chartTypeList = [
  // 'area',
  'bar',
  // 'bubble',
  // 'doughnut',
  'line',
  'pie',
  // 'ofPie',
  // 'radar',
  // 'scatter',
  // 'surface',
] as const;

type RichTextRun = {
  rPr?: ObjectItem;
  t?: {
    [textKey]: string;
  };
};
type SharedStringItem = {
  t?: {
    [textKey]: string;
  };
  // Rich-text runs. Excel writes these for cells with mixed font /
  // color / weight inside a single string. We can't represent
  // per-run formatting in the lib's flat ModelCellType.value
  // (boolean | string | number), so we concat run text into a single
  // plain string -- losing the formatting but keeping the text.
  // Without this, rich-text cells were silently dropped.
  r?: RichTextRun | RichTextRun[];
};
type ThemeData = Record<
  string,
  {
    'a:srgbClr': { val: string };
    'a:sysClr': { lastClr: string; val: string };
  }
>;
type SheetItem = Pick<WorksheetType, 'name' | 'sheetId'> & {
  'r:id': string;
  state?: string;
};
interface RelationItem {
  Id: string;
  Target: string;
  Type?: string;
}

interface ColItem {
  r: string;
  s: string;
  t?: string;
  f?: {
    [textKey]: string;
  };
  v?: {
    [textKey]: string;
  };
  // Inline string. Used when the cell has t="inlineStr". May contain
  // plain <t> or rich-text <r> runs same as a shared string.
  is?: {
    t?: { [textKey]: string };
    r?: RichTextRun | RichTextRun[];
  };
}
interface SheetDataRowItem {
  customHeight?: string;
  hidden?: string;
  ht?: string;
  r: string;
  c: ColItem[];
}

interface CustomColItem {
  min: string;
  max: string;
  width: string;
  customWidth: string;
  hidden?: string;
}

type ChartData = {
  'xdr:from': {
    'xdr:col': {
      [textKey]: string;
    };
    'xdr:row': {
      [textKey]: string;
    };
  };
  'xdr:graphicFrame'?: {
    'xdr:nvGraphicFramePr': {
      'xdr:cNvPr': {
        id: string;
        name: string;
        descr?: string;
        title?: string;
      };
    };
    'a:graphic': {
      'a:graphicData': {
        'c:chart': {
          'r:id': string;
        };
      };
    };
  };
  'xdr:pic'?: {
    'xdr:nvPicPr': {
      'xdr:cNvPr': {
        id: string;
        name: string;
        descr?: string;
        title?: string;
      };
    };
    'xdr:blipFill': {
      'a:blip': {
        'r:embed': string;
      };
    };
  };
};

export interface XfItem {
  fontId: string;
  fillId: string;
  numFmtId: string;
  borderId: string;
  applyFill?: string;
  applyFont?: string;
  applyNumberFormat?: string;
  applyAlignment?: string;
  applyBorder?: string;
  alignment?: {
    vertical?: 'top' | 'center' | 'bottom';
    horizontal?: 'left' | 'center' | 'right';
    wrapText?: string;
  };
}

type ObjectItem = Record<string, any>;
interface ColorItem {
  rgb?: string;
  theme?: string;
  indexed?: string;
  tint?: string;
  auto?: string;
}
interface FontItem {
  b?: ObjectItem;
  i?: ObjectItem;
  u?: ObjectItem;
  strike?: ObjectItem;
  color: ColorItem;
  name?: {
    val: string;
  };
  sz?: {
    val: string;
  };
}

interface NumFmtItem {
  numFmtId: string;
  formatCode: string;
}
interface FillItem {
  patternFill?: {
    patternType: string;
    fgColor: ColorItem;
    bgColor: ColorItem;
  };
  gradientFill?: {
    stop: Array<{ color: ColorItem }>;
  };
}

type BorderDataItem = {
  style?: string;
  color?: ColorItem;
};
interface StyleBorderItem {
  left: BorderDataItem;
  right: BorderDataItem;
  top: BorderDataItem;
  bottom: BorderDataItem;
  diagonal: {};
}

interface DefineNameItem {
  name: string;
  [textKey]: string;
}

function getArray<T>(
  obj: Record<string, any> | null | undefined,
  path: string,
  defaultValue?: T,
): T {
  const value = get<T>(obj, path, defaultValue);
  if (Array.isArray(value)) {
    return value;
  }
  // @ts-ignore
  return [value];
}

function xmlToJson(xml: any) {
  // Create the return object
  let obj: ObjectItem = {};

  if (xml.nodeType === Node.ELEMENT_NODE) {
    // element
    // do attributes
    if (xml.attributes.length > 0) {
      for (const attribute of xml.attributes) {
        obj[attribute.nodeName] = attribute.nodeValue;
      }
    }
  } else if (xml.nodeType === Node.TEXT_NODE) {
    // text
    obj = xml.nodeValue;
  }

  // do children
  if (xml.childNodes.length > 0) {
    for (const item of xml.childNodes) {
      const n = item.nodeName;
      // clear empty text
      if (typeof item.nodeValue === 'string' && !item.nodeValue.trim()) {
        continue;
      }
      if (typeof obj[n] === 'undefined') {
        obj[n] = xmlToJson(item);
      } else {
        if (typeof obj[n].push === 'undefined') {
          obj[n] = [obj[n]];
        }
        obj[n].push(xmlToJson(item));
      }
    }
  }
  return obj;
}

export function convertXMLToJSON(xmlStr: string) {
  const parser = new DOMParser();
  const xml = parser.parseFromString(xmlStr, 'text/xml');
  const json = xmlToJson(xml);
  return json;
}

// Concat the text content of a rich-text or plain string item into
// a single string. XLSX writes cells with mixed formatting as one
// or more <r> runs, each containing a <t> child. The lib's model
// has no rich-text representation (ModelCellType.value is
// string | boolean | number), so we keep the text and drop the
// per-run formatting. Without this fallback, cells whose only text
// lived inside <r> runs (no top-level <t>) ended up with value=''
// and visually disappeared on import.
function extractStringValue(
  data: { t?: { [k: string]: string }; r?: RichTextRun | RichTextRun[] } | undefined,
): string {
  if (!data) return '';
  if (data.t?.[textKey]) return data.t[textKey];
  if (data.r) {
    const runs = Array.isArray(data.r) ? data.r : [data.r];
    return runs
      .map((run) => run?.t?.[textKey] ?? '')
      .join('');
  }
  return '';
}

// OOXML standard 64-color indexed palette. Indices 64 and 65 are
// "system foreground" / "system background" (the auto colors) --
// returned as empty so the consuming style logic falls through to
// its own auto-color handling.
// Source: ECMA-376 Part 1, §18.8.27 (indexedColors).
const INDEXED_PALETTE: string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

// OOXML tint via HSL (the spec-correct method). Tint is in [-1, 1]:
// negative darkens toward black, positive lightens toward white.
// Without this, theme colors with non-zero tint rendered as the
// raw theme color -- e.g. a "dark blue, lighter 40%" cell would
// come back fully saturated dark blue instead of the lighter shade
// Excel actually paints.
function applyTint(rgbHex: string, tint: number): string {
  if (!rgbHex || rgbHex.length !== 6 || !Number.isFinite(tint) || tint === 0) {
    return rgbHex;
  }
  const r = parseInt(rgbHex.slice(0, 2), 16) / 255;
  const g = parseInt(rgbHex.slice(2, 4), 16) / 255;
  const b = parseInt(rgbHex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  let newL: number;
  if (tint < 0) newL = l * (1 + tint);
  else newL = l + (1 - l) * tint;
  newL = Math.max(0, Math.min(1, newL));
  const hueToRgb = (p: number, q: number, t: number): number => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r2: number;
  let g2: number;
  let b2: number;
  if (s === 0) {
    r2 = g2 = b2 = newL;
  } else {
    const q = newL < 0.5 ? newL * (1 + s) : newL + s - newL * s;
    const p = 2 * newL - q;
    r2 = hueToRgb(p, q, h + 1 / 3);
    g2 = hueToRgb(p, q, h);
    b2 = hueToRgb(p, q, h - 1 / 3);
  }
  const toHex = (n: number) =>
    Math.round(n * 255)
      .toString(16)
      .padStart(2, '0')
      .toUpperCase();
  return `${toHex(r2)}${toHex(g2)}${toHex(b2)}`;
}

// Resolve a XLSX color reference to a "#RRGGBB" string. Handles all
// four variants (rgb / theme / indexed / auto) plus the OOXML tint
// modifier, which is critical for matching Excel's actual rendered
// colors -- "Dark Blue, lighter 40%" cells reference theme=4 with
// tint=0.4 and need both lookup AND tint applied.
function convertColor(themeData: ThemeData, color?: ColorItem) {
  if (!color) {
    return '';
  }

  // Auto colors: foreground/background. We don't know the host's
  // current theme so return empty -- the consuming style logic falls
  // through to its own default (typically black-on-white).
  if (color.auto === '1' || color.auto === 'true') {
    return '';
  }

  // Resolve to a 6-char uppercase hex (no '#' prefix) so we can
  // optionally apply tint before formatting the final string.
  let hex = '';

  if (color.theme !== undefined) {
    // ECMA-376 indexed theme color order (§18.8.43):
    //   0 -> lt1, 1 -> dk1, 2 -> lt2, 3 -> dk2,
    //   4..9 -> accent1..accent6,
    //   10 -> hlink, 11 -> folHlink
    const themeIndex = [
      'a:lt1',
      'a:dk1',
      'a:lt2',
      'a:dk2',
      'a:accent1',
      'a:accent2',
      'a:accent3',
      'a:accent4',
      'a:accent5',
      'a:accent6',
      'a:hlink',
      'a:folHlink',
    ];
    const i = parseInt(color.theme, 10);
    const key = themeIndex[i];
    if (key) {
      // lt1/dk1 are stored as <a:sysClr val="..." lastClr="HEX"/>;
      // the rest as <a:srgbClr val="HEX"/>. Try both shapes.
      hex =
        themeData[key]?.['a:sysClr']?.lastClr ||
        themeData[key]?.['a:srgbClr']?.val ||
        '';
    }
  } else if (color.indexed !== undefined) {
    const i = parseInt(color.indexed, 10);
    // Indices 64 (system foreground) and 65 (system background) are
    // auto -- return empty as for color.auto.
    if (i >= 0 && i < INDEXED_PALETTE.length) {
      hex = INDEXED_PALETTE[i];
    }
  } else if (color.rgb) {
    // Strip leading 'FF' alpha if present (XLSX writes ARGB).
    if (color.rgb.length === 8) {
      hex = color.rgb.slice(2);
    } else if (color.rgb.length === 6) {
      hex = color.rgb;
    }
  }

  if (!hex) return '';

  // Apply tint on the resolved RGB.
  if (color.tint) {
    const tint = parseFloat(color.tint);
    if (Number.isFinite(tint) && tint !== 0) {
      hex = applyTint(hex, tint);
    }
  }

  return `#${hex.toUpperCase()}`;
}

function getBorder(
  data: BorderDataItem,
  themeData: ThemeData,
): BorderItem | undefined {
  if (!data?.style) {
    return undefined;
  }
  const type = data?.style as BorderType;
  if (typeof BORDER_TYPE_MAP[type] !== 'number') {
    return undefined;
  }
  const color = convertColor(themeData, data?.color);
  return {
    type,
    color,
  };
}

function getCellStyle(
  xml: ObjectItem,
  styleId: number,
  themeData: ThemeData,
): Partial<StyleType> {
  const result: Partial<StyleType> = {};
  const xfList = getArray<XfItem[]>(xml, 'styleSheet.cellXfs.xf', []);
  const cellStyles = getArray<XfItem[]>(xml, 'styleSheet.cellStyles.xf', []);
  const list = xfList.length > 0 ? xfList : cellStyles;
  if (!styleId || list.length === 0 || !list[styleId]) {
    return result;
  }
  const xf = list[styleId];
  if (xf.applyAlignment && xf.alignment) {
    if (xf.alignment.horizontal) {
      const alignMap = {
        left: EHorizontalAlign.LEFT,
        center: EHorizontalAlign.CENTER,
        right: EHorizontalAlign.RIGHT,
      };
      result.horizontalAlign = alignMap[xf.alignment.horizontal];
    }
    if (xf.alignment.vertical) {
      const alignMap = {
        top: EVerticalAlign.TOP,
        center: EVerticalAlign.MIDDLE,
        bottom: EVerticalAlign.BOTTOM,
      };
      result.verticalAlign = alignMap[xf.alignment.vertical];
    } else {
      result.verticalAlign = EVerticalAlign.BOTTOM;
    }
    result.isWrapText = Boolean(xf.alignment.wrapText);
  }
  if (xf.applyFont && xf.fontId) {
    const fontList = getArray<FontItem[]>(xml, 'styleSheet.fonts.font', []);
    const fontId = parseInt(xf.fontId, 10);
    if (fontList[fontId]) {
      const font = fontList[fontId];
      const fz = font?.sz?.val ? parseInt(font?.sz?.val, 10) : undefined;
      result.fontSize = fz ? fz : undefined;
      result.isBold = Boolean(font?.b);
      result.isItalic = Boolean(font?.i);
      result.isStrike = Boolean(font?.strike);
      result.underline = font.u ? EUnderLine.SINGLE : EUnderLine.NONE;
      result.fontFamily = font?.name?.val;
      const color = convertColor(themeData, font.color);
      if (color) {
        result.fontColor = color;
      }
    }
  }
  if (xf.applyNumberFormat && xf.numFmtId) {
    const list = getArray<NumFmtItem[]>(xml, 'styleSheet.numFmts.numFmt', []);
    const item = list.find((v) => v.numFmtId === xf.numFmtId);
    if (item) {
      result.numberFormat = item.formatCode;
    } else {
      const id = parseInt(xf.numFmtId, 10);
      const item = NUMBER_FORMAT_LIST.find((v) => v.id === id);
      if (item) {
        result.numberFormat = item.formatCode;
      }
    }
  }
  if (xf.applyFill && xf.fillId) {
    const list = getArray<FillItem[]>(xml, 'styleSheet.fills.fill', []);
    const i = parseInt(xf.fillId, 10);
    if (list[i]) {
      const g = list[i].gradientFill;
      const p = list[i].patternFill;
      if (g && g.stop[0]) {
        // TODO: handle multiple color
        const color = convertColor(themeData, g.stop[0].color);
        if (color) {
          result.fillColor = color;
        }
      } else if (p) {
        const color = convertColor(themeData, p.fgColor);
        if (color) {
          result.fillColor = color;
        }
      }
    }
  }
  if (xf.applyBorder && xf.borderId) {
    const list = getArray<StyleBorderItem[]>(
      xml,
      'styleSheet.borders.border',
      [],
    );
    const i = parseInt(xf.borderId, 10);
    if (list[i]) {
      const item = list[i];
      result.borderLeft = getBorder(item.left, themeData);
      result.borderRight = getBorder(item.right, themeData);
      result.borderTop = getBorder(item.top, themeData);
      result.borderBottom = getBorder(item.bottom, themeData);
    }
  }
  return result;
}

export function convertXMLDataToModel(
  xmlData: ObjectItem,
  imageSizeMap: Record<string, IWindowSize>,
): ModelJSON {
  const workbook = xmlData[WORKBOOK_PATH];

  const sharedStrings: SharedStringItem[] = getArray(
    xmlData[SHARED_STRINGS],
    'sst.si',
    [],
  );
  const themeData = get<ThemeData>(
    xmlData[THEME_PATH],
    'a:theme.a:themeElements.a:clrScheme',
    {},
  );

  const result: ModelJSON = {
    workbook: {},
    mergeCells: {},
    customHeight: {},
    customWidth: {},
    definedNames: {},
    currentSheetId: '',
    drawings: {},
    rangeMap: {},
    worksheets: {},
    autoFilter: {},
    scroll: {},
  };
  const relationList = getArray<RelationItem[]>(
    xmlData[WORKBOOK_RELATION_PATH],
    'Relationships.Relationship',
    [],
  );
  const sheetPathMap: Record<string, string> = {};
  const drawingMap: Record<string, string[]> = {};
  let drawingCount = 0;
  const sheetList: SheetItem[] = getArray(
    workbook,
    'workbook.sheets.sheet',
    [],
  );
  let sheetSort = 0;
  for (const item of sheetList) {
    if (!item) {
      continue;
    }
    const sheetTarget =
      relationList.find((v) => v.Id === item['r:id'])?.Target || '';
    const worksheetPrefix = 'worksheets/';
    const baseSheet = sheetTarget.slice(worksheetPrefix.length);
    const refPath = `${COMMON_PREFIX}/${worksheetPrefix}_rels/${baseSheet}.rels`;
    if (xmlData[refPath]) {
      const list: RelationItem[] = getArray(
        xmlData[refPath],
        'Relationships.Relationship',
        [],
      );
      for (const v of list) {
        if (!drawingMap[item.sheetId]) {
          drawingMap[item.sheetId] = [];
        }
        drawingMap[item.sheetId].push(v.Target);
        if (v.Target.startsWith(DRAWING_FLAG)) {
          drawingCount++;
        }
      }
    }
    const sheetPath = `${COMMON_PREFIX}/${sheetTarget}`;
    sheetPathMap[item.sheetId] = sheetPath;
    const range = parseReference(
      get<string>(
        xmlData[sheetPath],
        'worksheet.sheetViews.sheetView.selection.sqref',
        '',
      ),
    )!;
    const tabSelected = get<string>(
      xmlData[sheetPath],
      'worksheet.sheetViews.sheetView.tabSelected',
      '',
    );
    const tabColor = convertColor(
      themeData,
      get<ColorItem>(xmlData[sheetPath], 'worksheet.sheetPr.tabColor', {
        rgb: '',
      }),
    );
    if (tabSelected === '1') {
      result.currentSheetId = item.sheetId;
    }
    const sheetData: WorksheetType = {
      sheetId: item.sheetId,
      name: item.name,
      isHide: item.state === 'hidden',
      rowCount: 200,
      colCount: 200,
      sort: sheetSort++,
      tabColor: tabColor === '' ? undefined : tabColor,
    };
    result.workbook[sheetData.sheetId] = sheetData;
    if (range) {
      range.sheetId = sheetData.sheetId;
      result.rangeMap[item.sheetId] = range.toIRange();
    }
  }
  const sheets = Object.values(result.workbook);
  sheets.sort((a, b) => a.sort - b.sort);
  result.currentSheetId = result.currentSheetId || sheets?.[0]?.sheetId;
  const convertSheetName = (sheetName: string) => {
    const list = Object.values(result.workbook);
    return list.find((v) => v.name === sheetName)?.sheetId || '';
  };
  for (const item of sheets) {
    const sheetPath = sheetPathMap[item.sheetId];
    const sheetData: SheetDataRowItem[] = getArray(
      xmlData[sheetPath],
      'worksheet.sheetData.row',
      [],
    );
    const mergeCellData: Array<{ ref: string }> = getArray(
      xmlData[sheetPath],
      'worksheet.mergeCells.mergeCell',
      [],
    );
    for (const mergeCell of mergeCellData) {
      const range = parseReference(mergeCell.ref, convertSheetName);
      if (range) {
        range.sheetId = item.sheetId;
        const ref = `${item.name}!${mergeCell.ref}`;
        result.mergeCells[ref] = range.toIRange();
      }
    }
    const customWidth: CustomColItem[] = getArray(
      xmlData[sheetPath],
      'worksheet.cols.col',
      [],
    );
    const defaultWOrH = get(xmlData[sheetPath], 'worksheet.sheetFormatPr', {
      defaultColWidth: '',
      defaultRowHeight: '',
      outlineLevelRow: '',
    });
    if (customWidth.length > 0) {
      for (const col of customWidth) {
        if (col && col.customWidth && col.width && col.min && col.max) {
          const isDefault = defaultWOrH.defaultColWidth === col.width;
          const w = isDefault ? CELL_WIDTH : convertFromPt(col.width);
          const isHide = Boolean(col.hidden);
          for (
            let start = parseInt(col.min, 10) - 1, end = parseInt(col.max, 10);
            start < end;
            start++
          ) {
            result.customWidth[getCustomWidthOrHeightKey(item.sheetId, start)] =
              {
                len: w,
                isHide,
              };
          }
        }
      }
    }

    if (sheetData.length === 0) {
      continue;
    }

    let { colCount } = item;
    let { rowCount } = item;
    for (const row of sheetData) {
      if (!row) {
        continue;
      }
      const realRow = parseInt(row.r, 10) - 1;
      rowCount = Math.max(rowCount, realRow + 1);
      if (rowCount > XLSX_MAX_ROW_COUNT) {
        continue;
      }
      if (row.customHeight && row.ht) {
        const isDefault = defaultWOrH.defaultRowHeight === row.ht;
        result.customHeight[getCustomWidthOrHeightKey(item.sheetId, realRow)] =
          {
            len: isDefault ? CELL_HEIGHT : Math.floor(parseInt(row.ht, 10)),
            isHide: Boolean(row.hidden),
          };
      }
      const colList = Array.isArray(row.c) ? row.c : [row.c];
      for (const col of colList) {
        if (!col) {
          continue;
        }
        const range = parseReference(col.r);
        if (!range) {
          continue;
        }
        colCount = Math.max(colCount, range.col + 1);
        if (colCount > XLSX_MAX_COL_COUNT) {
          continue;
        }
        const val = col?.v?.[textKey] ?? '';
        const styleId = parseInt(col.s, 10);
        const style = getCellStyle(xmlData[STYLE_PATH], styleId, themeData);
        const t: ModelCellType = {
          ...style,
          value: val,
        };
        const formula = col?.f?.[textKey] || '';
        if (formula) {
          if (formula.startsWith(FORMULA_PREFIX)) {
            t.formula = formula;
          } else {
            t.formula = FORMULA_PREFIX + formula;
          }
        }
        if (col.t === 's') {
          const i = parseInt(val, 10);
          if (!isNaN(i)) {
            // extractStringValue handles both <t> (plain) and <r><t>
            // (rich-text runs). Without the rich-text fallback, cells
            // referencing a shared string that consisted ONLY of
            // formatted runs were silently set to value='' and
            // disappeared from the imported sheet.
            t.value = extractStringValue(sharedStrings[i]);
          }
        }
        if (col.t === 'b') {
          t.value = val === '1';
        }
        // Inline string (<c t="inlineStr"><is>...</is></c>). May be
        // plain <t> or rich-text <r> runs. Same extractStringValue
        // path as shared strings.
        if (col.t === 'inlineStr' || col.is) {
          const inline = extractStringValue(col.is);
          if (inline) {
            t.value = inline;
          }
        }
        result.worksheets[getWorksheetKey(item.sheetId, realRow, range.col)] =
          t;
        colCount = Math.max(colCount, range.col + 1);
      }
    }
    item.rowCount = Math.max(item.rowCount, rowCount);
    item.colCount = Math.max(item.colCount, colCount);
  }

  const definedNames: DefineNameItem[] = getArray(
    workbook,
    'workbook.definedNames.definedName',
    [],
  );

  for (const item of definedNames) {
    const range = parseReference(item[textKey], convertSheetName);
    if (range && item?.name) {
      result.definedNames[item.name.toLowerCase()] = range.toIRange();
    }
  }
  for (let drawingId = 1; drawingId <= drawingCount; drawingId++) {
    const basePath = `drawing${drawingId}.xml`;
    const key = `${DRAWING_PREFIX_KEY}${basePath}`;
    const ref = `${DRAWING_PREFIX_KEY}_rels/${basePath}.rels`;
    if (!xmlData[key] || !xmlData[ref]) {
      break;
    }

    const relations: RelationItem[] = getArray(
      xmlData[ref],
      'Relationships.Relationship',
      [],
    );
    let sheetId = '';
    for (const [id, list] of Object.entries(drawingMap)) {
      if (list.some((v) => v === DRAWING_FLAG + basePath)) {
        sheetId = id;
        break;
      }
    }

    const floatElementList: ChartData[] = getArray(
      xmlData[key],
      'xdr:wsDr.xdr:twoCellAnchor',
      [],
    );
    for (const float of floatElementList) {
      const chartId = get(
        float,
        'xdr:graphicFrame.a:graphic.a:graphicData.c:chart.r:id',
        '',
      );
      const imageId = get(float, 'xdr:pic.xdr:blipFill.a:blip.r:embed', '');
      const isImage = !!imageId;
      const rid = isImage ? imageId : chartId;
      const target = relations.find((v) => v.Id === rid)?.Target || '';
      if (!target) {
        continue;
      }
      const filePath = COMMON_PREFIX + target.slice(2);
      const fromRow = float['xdr:from']['xdr:col'][textKey];
      const fromCol = float['xdr:from']['xdr:row'][textKey];
      if (!fromRow || !fromCol) {
        continue;
      }
      const uuid = String(drawingId);
      const floatElementData: DrawingElement = {
        title: '',
        type: isImage ? 'floating-picture' : 'chart',
        uuid,
        width: 300,
        height: 300,
        originHeight: 300,
        originWidth: 300,
        fromCol: parseInt(fromCol, 10),
        fromRow: parseInt(fromRow, 10),
        sheetId,
        marginX: 0,
        marginY: 0,
      };

      if (isImage) {
        if (!xmlData[filePath]) {
          continue;
        }
        const name = get(float, 'xdr:pic.xdr:nvPicPr.xdr:cNvPr.name', '');
        const title = get(float, 'xdr:pic.xdr:nvPicPr.xdr:cNvPr.title', '');
        floatElementData.title = title || name;
        floatElementData.imageAngle = 0;
        floatElementData.imageSrc = xmlData[filePath] || '';
        const size = imageSizeMap[filePath];
        if (size && size.width > 0 && size.height > 0) {
          floatElementData.width = size.width;
          floatElementData.height = size.height;
          floatElementData.originWidth = size.width;
          floatElementData.originHeight = size.height;
        }
      } else {
        floatElementData.title = get(
          xmlData[filePath],
          'c:chartSpace.c:chart.c:title.c:tx.c:rich.a:p.0.a:r.a:t.' + textKey,
          '',
        );
        let refList: string[] = [];
        for (const chartType of chartTypeList) {
          const chartData = get(
            xmlData[filePath],
            `c:chartSpace.c:chart.c:plotArea.c:${chartType}Chart`,
            '',
          );
          if (chartData) {
            floatElementData.chartType = chartType;
            const list = getArray(chartData, 'c:ser', []);
            refList = list
              .map((v) => get(v, `c:val.c:numRef.c:f.${textKey}`, ''))
              .filter((v) => v);
            break;
          }
        }

        if (refList.length === 0) {
          continue;
        }
        const convertSheetNameToSheetId = (sheetName: string) => {
          const data = Object.values(result.workbook).find(
            (v) => v.name === sheetName,
          );
          return data?.sheetId || '';
        };
        let chartRange: SheetRange | undefined = undefined;
        for (const item of refList) {
          const t = parseReference(item, convertSheetNameToSheetId);
          if (t) {
            if (!chartRange) {
              chartRange = t;
            } else {
              chartRange = mergeRange(chartRange, t);
            }
          }
        }
        if (!chartRange) {
          continue;
        }
        chartRange.sheetId = chartRange.sheetId || sheetId;
        floatElementData.chartRange = chartRange.toIRange();
      }

      result.drawings[uuid] = floatElementData;
    }
  }

  return result;
}

export async function importXLSX(file: File) {
  const jszip = await import('jszip');
  const zip = await jszip.default.loadAsync(file);
  const { files } = zip;
  const result: ObjectItem = {};
  const imageSizeMap: Record<string, IWindowSize> = {};
  for (const key of Object.keys(files)) {
    if (files[key].dir) {
      continue;
    }
    if (key.includes('.xml')) {
      const data = await files[key].async('string');
      if (data) {
        result[key] = convertXMLToJSON(data);
      }
    } else {
      let imageType = '';
      for (const [type, list] of Object.entries(IMAGE_TYPE_MAP)) {
        if (list.some((v) => key.endsWith(v))) {
          imageType = type;
          break;
        }
      }
      if (!imageType) {
        continue;
      }
      const data = await files[key].async('base64');
      if (data) {
        result[key] = `data:${imageType};base64,${data}`;
        imageSizeMap[key] = await getImageSize(result[key]);
      }
    }
  }

  const model = convertXMLDataToModel(result, imageSizeMap);
  return model;
}
