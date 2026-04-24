import { jsxDEV as I } from "react/jsx-dev-runtime";
import je, { memo as ue, useRef as Xe, useEffect as Re, useCallback as z, useState as he, useMemo as We, createContext as Kc, useContext as Jc, Fragment as rn, Suspense as Qc, useSyncExternalStore as jc } from "react";
import { createPortal as Zc } from "react-dom";
import { createRoot as la } from "react-dom/client";
import { Map as we, UndoManager as eu, Doc as $o } from "yjs";
var tn = /* @__PURE__ */ ((e) => (e[e.NONE = 0] = "NONE", e[e.VERTICAL = 1] = "VERTICAL", e[e.HORIZONTAL = 2] = "HORIZONTAL", e))(tn || {}), dt = /* @__PURE__ */ ((e) => (e[e.NONE = 0] = "NONE", e[e.EDIT_CELL = 1] = "EDIT_CELL", e[e.EDIT_FORMULA_BAR = 2] = "EDIT_FORMULA_BAR", e))(dt || {}), Te = /* @__PURE__ */ ((e) => (e[e.TOP = 0] = "TOP", e[e.MIDDLE = 1] = "MIDDLE", e[e.BOTTOM = 2] = "BOTTOM", e))(Te || {}), xe = /* @__PURE__ */ ((e) => (e[e.LEFT = 0] = "LEFT", e[e.CENTER = 1] = "CENTER", e[e.RIGHT = 2] = "RIGHT", e))(xe || {}), Ce = /* @__PURE__ */ ((e) => (e[e.NONE = 0] = "NONE", e[e.SINGLE = 1] = "SINGLE", e[e.DOUBLE = 2] = "DOUBLE", e))(Ce || {}), Ct = /* @__PURE__ */ ((e) => (e[e.MERGE_CENTER = 0] = "MERGE_CENTER", e[e.MERGE_CELL = 1] = "MERGE_CELL", e[e.MERGE_CONTENT = 2] = "MERGE_CONTENT", e))(Ct || {}), sn = /* @__PURE__ */ ((e) => (e.MODEL = "model", e.SKIP_UPDATE = "skip-update", e.SKIP_UNDO_REDO = "skip-undo-redo", e.SKIP_UNDO_REDO_UPDATE = "skip-undo-redo-update", e))(sn || {}), F = /* @__PURE__ */ ((e) => (e[e.EQUAL = 0] = "EQUAL", e[e.NOT_EQUAL = 1] = "NOT_EQUAL", e[e.PLUS = 2] = "PLUS", e[e.MINUS = 3] = "MINUS", e[e.STAR = 4] = "STAR", e[e.SLASH = 5] = "SLASH", e[e.EXPONENT = 6] = "EXPONENT", e[e.GREATER = 7] = "GREATER", e[e.GREATER_EQUAL = 8] = "GREATER_EQUAL", e[e.CONCATENATE = 9] = "CONCATENATE", e[e.COLON = 10] = "COLON", e[e.COMMA = 11] = "COMMA", e[e.EMPTY_CHAR = 12] = "EMPTY_CHAR", e[e.PERCENT = 13] = "PERCENT", e[e.LESS = 14] = "LESS", e[e.LESS_EQUAL = 15] = "LESS_EQUAL", e[e.STRING = 16] = "STRING", e[e.NUMBER = 17] = "NUMBER", e[e.BOOL = 18] = "BOOL", e[e.ERROR = 19] = "ERROR", e[e.ERROR_REF = 20] = "ERROR_REF", e[e.LEFT_BRACKET = 21] = "LEFT_BRACKET", e[e.RIGHT_BRACKET = 22] = "RIGHT_BRACKET", e[e.lEFT_BRACE = 23] = "lEFT_BRACE", e[e.RIGHT_BRACE = 24] = "RIGHT_BRACE", e[e.SEMICOLON = 25] = "SEMICOLON", e[e.EXCLAMATION = 26] = "EXCLAMATION", e[e.R1C1 = 27] = "R1C1", e[e.SHEET_NAME = 28] = "SHEET_NAME", e[e.CELL = 29] = "CELL", e[e.COLUMN = 30] = "COLUMN", e[e.ROW = 31] = "ROW", e[e.DEFINED_NAME = 32] = "DEFINED_NAME", e[e.EXCEL_FUNCTION = 33] = "EXCEL_FUNCTION", e[e.REF_FUNCTION = 34] = "REF_FUNCTION", e[e.REF_FUNCTION_COND = 35] = "REF_FUNCTION_COND", e[e.EOF = 36] = "EOF", e))(F || {}), qo = /* @__PURE__ */ ((e) => (e.ONLINE = "connection", e.LOCAL = "disconnection", e.SYNCING = "syncing", e))(qo || {});
const Cn = 12, ca = "sans-serif", tu = "Sheet", nr = 22, rr = 76, fs = nr, ds = rr / 2, Wo = 0, nu = 10, mt = "&#10;", dr = `
`, ua = ",", hr = 16384, mr = 1048576, ru = 200, Eo = 200, ou = 30, ht = -999, Yo = 16, fa = "__sheet_item_", Fn = "=", Pn = "_", or = 35, da = 256, sN = 9, ha = "__FORMULA_EDITOR_ROLE__", cn = "General", ma = "@", ga = ["ar-EG", "ar-SA"], Xo = [
  "en-US",
  "zh-CN",
  "fr-FR",
  "zh-HK",
  "ja-JP",
  "ko-KR",
  "es-ES",
  "de-DE",
  "it-IT",
  "pt-BR",
  "pt-PT",
  "nl-NL",
  ...ga
], pa = "#GETTING_DATA", hs = /* @__PURE__ */ new Set([
  "#DIV/0!",
  "#NULL!",
  "#NUM!",
  "#REF!",
  "#VALUE!",
  "#N/A",
  "#NAME?",
  pa
]), Z = 1, oo = {
  thin: Z,
  hair: Z,
  dotted: Z,
  dashed: Z,
  dashDot: Z,
  dashDotDot: Z,
  double: Z,
  medium: Z * 2,
  mediumDashed: Z * 2,
  mediumDashDot: Z * 2,
  mediumDashDotDot: Z * 2,
  thick: Z * 4
}, iu = [
  "#35322B",
  "#505050",
  "#606060",
  "#6F6F6F",
  "#8B8B8B",
  "#B2B2B2",
  "#CCCCCC",
  "#E5E5E5",
  "#F5F5F5",
  "#FFFFFF",
  "#9D0000",
  "#B20000",
  "#CD0F0F",
  "#E30909",
  "#F30B0B",
  "#FE4B4B",
  "#FA7979",
  "#FB9D9D",
  "#FDCECE",
  "#FEE7E7",
  "#B24000",
  "#CC4F10",
  "#DF5D00",
  "#F96800",
  "#FB8937",
  "#FF8C51",
  "#FCA669",
  "#FDC49B",
  "#FEE1CD",
  "#FEF0E6",
  "#B19401",
  "#C5A300",
  "#D8B300",
  "#EBC301",
  "#F9D700",
  "#FBE137",
  "#FCE869",
  "#FDF09B",
  "#FEF7CD",
  "#FEFBE6",
  "#58770A",
  "#688C0D",
  "#7AA017",
  "#8BBB11",
  "#A4DC16",
  "#BEEE44",
  "#CEF273",
  "#DEF6A2",
  "#EFFBD0",
  "#F7FDE8",
  "#007676",
  "#008A8A",
  "#009E9E",
  "#00BBBB",
  "#1CD8D8",
  "#2AEAEA",
  "#76EFEF",
  "#A3F5F5",
  "#D1FAFA",
  "#E8FCFC",
  "#001F9C",
  "#0025B7",
  "#012BD2",
  "#133DE3",
  "#2F55EB",
  "#4567ED",
  "#738DF2",
  "#A2B3F6",
  "#D0D9FB",
  "#E8ECFD",
  "#3F0198",
  "#510EB0",
  "#6721CB",
  "#7735D4",
  "#894EDE",
  "#9E6DE3",
  "#AA82E3",
  "#C7ABED",
  "#E3D5F6",
  "#F1EAFA",
  "#8F0550",
  "#A1095C",
  "#C1026B",
  "#D4157E",
  "#E7258F",
  "#F248A6",
  "#F273B9",
  "#F6A2D0",
  "#FBD0E8",
  "#FDE8F3"
], su = /^[$]?[A-Za-z]{1,3}[$]?[1-9][0-9]*$/, au = /^[$]?[A-Za-z]{1,3}$/, lu = /^\$[1-9][0-9]*$/, Ci = /^[a-zA-Z_][a-zA-Z0-9_.?]*$/, cu = /* @__PURE__ */ new Set([
  "ABS",
  "ACCRINT",
  "ACCRINTM",
  "ACOS",
  "ACOSH",
  "ACOT",
  "ACOTH",
  "ADDRESS",
  "AGGREGATE",
  "AMORDEGRC",
  "AMORLINC",
  "AND",
  "ARABIC",
  "AREAS",
  "ARRAYTOTEXT",
  "ASC",
  "ASIN",
  "ASINH",
  "ATAN",
  "ATAN2",
  "ATANH",
  "AVEDEV",
  "AVERAGE",
  "AVERAGEA",
  "AVERAGEIF",
  "AVERAGEIFS",
  "BAHTTEXT",
  "BASE",
  "BESSELI",
  "BESSELJ",
  "BESSELK",
  "BESSELY",
  "BETA.DIST",
  "BETA.INV",
  "BETADIST",
  "BETAINV",
  "BIN2DEC",
  "BIN2HEX",
  "BIN2OCT",
  "BINOM.DIST",
  "BINOM.DIST.RANGE",
  "BINOM.INV",
  "BINOMDIST",
  "BITAND",
  "BITLSHIFT",
  "BITOR",
  "BITRSHIFT",
  "BITXOR",
  "CALL",
  "CEILING",
  "CEILING.MATH",
  "CEILING.PRECISE",
  "CELL",
  "CHAR",
  "CHIDIST",
  "CHIINV",
  "CHISQ.DIST",
  "CHISQ.DIST.RT",
  "CHISQ.INV",
  "CHISQ.INV.RT",
  "CHISQ.TEST",
  "CHITEST",
  "CLEAN",
  "CODE",
  "COLUMN",
  "COLUMNS",
  "COMBIN",
  "COMBINA",
  "COMPLEX",
  "CONCAT",
  "CONCATENATE",
  "CONFIDENCE",
  "CONFIDENCE.NORM",
  "CONFIDENCE.T",
  "CONVERT",
  "CORREL",
  "COS",
  "COSH",
  "COT",
  "COTH",
  "COUNT",
  "COUNTA",
  "COUNTBLANK",
  "COUNTIF",
  "COUNTIFS",
  "COUPDAYBS",
  "COUPDAYS",
  "COUPDAYSNC",
  "COUPNCD",
  "COUPNUM",
  "COUPPCD",
  "COVAR",
  "COVARIANCE.P",
  "COVARIANCE.S",
  "CRITBINOM",
  "CSC",
  "CSCH",
  "CUBEKPIMEMBER",
  "CUBEMEMBER",
  "CUBEMEMBERPROPERTY",
  "CUBERANKEDMEMBER",
  "CUBESET",
  "CUBESETCOUNT",
  "CUBEVALUE",
  "CUMIPMT",
  "CUMPRINC",
  "DATE",
  "DATEDIF",
  "DATEVALUE",
  "DAVERAGE",
  "DAY",
  "DAYS",
  "DAYS360",
  "DB",
  "DBCS",
  "DCOUNT",
  "DCOUNTA",
  "DDB",
  "DEC2BIN",
  "DEC2HEX",
  "DEC2OCT",
  "DECIMAL",
  "DEGREES",
  "DELTA",
  "DEVSQ",
  "DGET",
  "DISC",
  "DMAX",
  "DMIN",
  "DOLLAR",
  "DOLLARDE",
  "DOLLARFR",
  "DPRODUCT",
  "DSTDEV",
  "DSTDEVP",
  "DSUM",
  "DURATION",
  "DVAR",
  "DVARP",
  "EDATE",
  "EFFECT",
  "ENCODEURL",
  "EOMONTH",
  "ERF",
  "ERF.PRECISE",
  "ERFC",
  "ERFC.PRECISE",
  "ERROR.TYPE",
  "EUROCONVERT",
  "EVEN",
  "EXACT",
  "EXP",
  "EXPON.DIST",
  "EXPONDIST",
  "F.DIST",
  "F.DIST.RT",
  "F.INV",
  "F.INV.RT",
  "F.TEST",
  "FACT",
  "FACTDOUBLE",
  "FALSE",
  "FDIST",
  "FILTER",
  "FILTERXML",
  "FIND",
  "FINDB",
  "FINV",
  "FISHER",
  "FISHERINV",
  "FIXED",
  "FLOOR",
  "FLOOR.MATH",
  "FLOOR.PRECISE",
  "FORECAST",
  "FORECAST.ETS",
  "FORECAST.ETS.CONFINT",
  "FORECAST.ETS.SEASONALITY",
  "FORECAST.ETS.STAT",
  "FORECAST.LINEAR",
  "FORMULATEXT",
  "FREQUENCY",
  "FTEST",
  "FV",
  "FVSCHEDULE",
  "GAMMA",
  "GAMMA.DIST",
  "GAMMA.INV",
  "GAMMADIST",
  "GAMMAINV",
  "GAMMALN",
  "GAMMALN.PRECISE",
  "GAUSS",
  "GCD",
  "GEOMEAN",
  "GESTEP",
  "GETPIVOTDATA",
  "GROWTH",
  "HARMEAN",
  "HEX2BIN",
  "HEX2DEC",
  "HEX2OCT",
  "HLOOKUP",
  "HOUR",
  "HYPERLINK",
  "HYPGEOM.DIST",
  "HYPGEOMDIST",
  "IFERROR",
  "IFNA",
  "IFS",
  "IMABS",
  "IMAGINARY",
  "IMARGUMENT",
  "IMCONJUGATE",
  "IMCOS",
  "IMCOSH",
  "IMCOT",
  "IMCSC",
  "IMCSCH",
  "IMDIV",
  "IMEXP",
  "IMLN",
  "IMLOG10",
  "IMLOG2",
  "IMPOWER",
  "IMPRODUCT",
  "IMREAL",
  "IMSEC",
  "IMSECH",
  "IMSIN",
  "IMSINH",
  "IMSQRT",
  "IMSUB",
  "IMSUM",
  "IMTAN",
  "INFO",
  "INT",
  "INTERCEPT",
  "INTRATE",
  "IPMT",
  "IRR",
  "ISBLANK",
  "ISERR",
  "ISERROR",
  "ISEVEN",
  "ISFORMULA",
  "ISLOGICAL",
  "ISNA",
  "ISNONTEXT",
  "ISNUMBER",
  "ISO.CEILING",
  "ISODD",
  "ISOWEEKNUM",
  "ISPMT",
  "ISREF",
  "ISTEXT",
  "JIS",
  "KURT",
  "LARGE",
  "LCM",
  "LEFT",
  "LEFTB",
  "LEN",
  "LENB",
  "LET",
  "LINEST",
  "LN",
  "LOG",
  "LOG10",
  "LOGEST",
  "LOGINV",
  "LOGNORM.DIST",
  "LOGNORM.INV",
  "LOGNORMDIST",
  "LOOKUP",
  "LOWER",
  "MATCH",
  "MAX",
  "MAXA",
  "MAXIFS",
  "MDETERM",
  "MDURATION",
  "MEDIAN",
  "MID",
  "MIDB",
  "MIN",
  "MINA",
  "MINIFS",
  "MINUTE",
  "MINVERSE",
  "MIRR",
  "MMULT",
  "MOD",
  "MODE",
  "MODE.MULT",
  "MODE.SNGL",
  "MONTH",
  "MROUND",
  "MULTINOMIAL",
  "MUNIT",
  "N",
  "NA",
  "NEGBINOM.DIST",
  "NEGBINOMDIST",
  "NETWORKDAYS",
  "NETWORKDAYS.INTL",
  "NOMINAL",
  "NORM.DIST",
  "NORM.INV",
  "NORM.S.DIST",
  "NORM.S.INV",
  "NORMDIST",
  "NORMINV",
  "NORMSDIST",
  "NORMSINV",
  "NOT",
  "NOW",
  "NPER",
  "NPV",
  "NUMBERVALUE",
  "OCT2BIN",
  "OCT2DEC",
  "OCT2HEX",
  "ODD",
  "ODDFPRICE",
  "ODDFYIELD",
  "ODDLPRICE",
  "ODDLYIELD",
  "OR",
  "PDURATION",
  "PEARSON",
  "PERCENTILE",
  "PERCENTILE.EXC",
  "PERCENTILE.INC",
  "PERCENTRANK",
  "PERCENTRANK.EXC",
  "PERCENTRANK.INC",
  "PERMUT",
  "PERMUTATIONA",
  "PHI",
  "PHONETIC",
  "PI",
  "PMT",
  "POISSON",
  "POISSON.DIST",
  "POWER",
  "PPMT",
  "PRICE",
  "PRICEDISC",
  "PRICEMAT",
  "PROB",
  "PRODUCT",
  "PROPER",
  "PV",
  "QUARTILE",
  "QUARTILE.EXC",
  "QUARTILE.INC",
  "QUOTIENT",
  "RADIANS",
  "RAND",
  "RANDARRAY",
  "RANDBETWEEN",
  "RANK",
  "RANK.AVG",
  "RANK.EQ",
  "RATE",
  "RECEIVED",
  "REGISTER.ID",
  "REPLACE",
  "REPLACEB",
  "REPT",
  "RIGHT",
  "RIGHTB",
  "ROMAN",
  "ROUND",
  "ROUNDDOWN",
  "ROUNDUP",
  "ROW",
  "ROWS",
  "RRI",
  "RSQ",
  "RTD",
  "SEARCH",
  "SEARCHB",
  "SEC",
  "SECH",
  "SECOND",
  "SEQUENCE",
  "SERIESSUM",
  "SHEET",
  "SHEETS",
  "SIGN",
  "SIN",
  "SINH",
  "SKEW",
  "SKEW.P",
  "SLN",
  "SLOPE",
  "SMALL",
  "SORT",
  "SORTBY",
  "SQRT",
  "SQRTPI",
  "STANDARDIZE",
  "STDEV",
  "STDEV.P",
  "STDEV.S",
  "STDEVA",
  "STDEVP",
  "STDEVPA",
  "STEYX",
  "SUBSTITUTE",
  "SUBTOTAL",
  "SUM",
  "SUMIF",
  "SUMIFS",
  "SUMPRODUCT",
  "SUMSQ",
  "SUMX2MY2",
  "SUMX2PY2",
  "SUMXMY2",
  "SWITCH",
  "SYD",
  "T",
  "T.DIST",
  "T.DIST.2T",
  "T.DIST.RT",
  "T.INV",
  "T.INV.2T",
  "T.TEST",
  "TAN",
  "TANH",
  "TBILLEQ",
  "TBILLPRICE",
  "TBILLYIELD",
  "TDIST",
  "TEXT",
  "TEXTJOIN",
  "TIME",
  "TIMEVALUE",
  "TINV",
  "TODAY",
  "TRANSPOSE",
  "TREND",
  "TRIM",
  "TRIMMEAN",
  "TRUE",
  "TRUNC",
  "TTEST",
  "TYPE",
  "UNICHAR",
  "UNICODE",
  "UNIQUE",
  "UPPER",
  "VALUE",
  "VALUETOTEXT",
  "VAR",
  "VAR.P",
  "VAR.S",
  "VARA",
  "VARP",
  "VARPA",
  "VDB",
  "VLOOKUP",
  "WEBSERVICE",
  "WEEKDAY",
  "WEEKNUM",
  "WEIBULL",
  "WEIBULL.DIST",
  "WORKDAY",
  "WORKDAY.INTL",
  "XIRR",
  "XLOOKUP",
  "XMATCH",
  "XNPV",
  "XOR",
  "YEAR",
  "YEARFRAC",
  "YIELD",
  "YIELDDISC",
  "YIELDMAT",
  "Z.TEST",
  "ZTEST"
]), Ko = [
  {
    value: "line",
    label: "line-chart"
  },
  {
    value: "bar",
    label: "bar-chart"
  },
  {
    value: "pie",
    label: "pie-chart"
  },
  {
    value: "scatter",
    label: "scatter-chart"
  },
  {
    value: "radar",
    label: "radar-chart"
  },
  {
    value: "polarArea",
    label: "polar-area-chart"
  }
], Ei = [
  { formatCode: cn, id: 0 },
  { formatCode: "0", id: 1 },
  { formatCode: "0.00", id: 2 },
  { formatCode: "#,##0", id: 3 },
  { formatCode: "#,##0.00", id: 4 },
  { formatCode: '"$"#,##0.00_);[Red]("$"#,##0.00)', id: 8 },
  { formatCode: "0%", id: 9 },
  { formatCode: "0.00%", id: 10 },
  { formatCode: "0.00E+00", id: 11 },
  { formatCode: "# ?/?", id: 12 },
  { formatCode: "# ??/??", id: 13 },
  { formatCode: "mm-dd-yy", id: 14 },
  { formatCode: "d-mmm-yy", id: 15 },
  { formatCode: "d-mmm", id: 16 },
  { formatCode: "mmm-yy", id: 17 },
  { formatCode: "h:mm AM/PM", id: 18 },
  { formatCode: "h:mm:ss AM/PM", id: 19 },
  { formatCode: "h:mm", id: 20 },
  { formatCode: "h:mm:ss", id: 21 },
  { formatCode: "m/d/yy h:mm", id: 22 },
  {
    formatCode: 'yyyy"年"m"月"',
    id: 27
  },
  { formatCode: "#,##0 ;(#,##0)", id: 37 },
  { formatCode: "#,##0 ;[Red](#,##0)", id: 38 },
  { formatCode: "#,##0.00 ;(#,##0.00)", id: 39 },
  { formatCode: "#,##0.00 ;[Red](#,##0.00)", id: 40 },
  {
    formatCode: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    id: 44
  },
  { formatCode: "mm:ss", id: 45 },
  { formatCode: "[h]:mm:ss", id: 46 },
  { formatCode: "mmss.0", id: 47 },
  { formatCode: "##0.0E+0", id: 48 },
  { formatCode: ma, id: 49 },
  {
    formatCode: '上午/下午 h"时"mm"分"ss"秒"',
    id: 56
  }
], xt = (e) => Ei.find((t) => t.id === e).formatCode;
function uu(e, t) {
  let n = !1, r;
  return function(...o) {
    n ? r = o : (e.apply(this, o), setTimeout(() => {
      r && e.apply(this, r), n = !1;
    }, t));
  };
}
function tt(e, t, n) {
  const r = e == null ? void 0 : t.replace(/\[/g, ".").replace(/\]/g, "").split(".").reduce((o, s) => o == null ? o : o[s], e);
  return r === void 0 ? n : r;
}
function On(e) {
  const t = e || {};
  return [Object, Array].includes(t.constructor) && !Object.entries(t).length;
}
function io(e, t) {
  if (e === t)
    return !0;
  if (e && typeof e == "object" && t && typeof t == "object") {
    if (Object.keys(e).length !== Object.keys(t).length)
      return !1;
    const n = Object.keys(e), r = Object.keys(t);
    for (let o = 0; o < n.length; o++) {
      const s = n[o], l = r[o];
      if (s !== l || !io(e[s], t[l]))
        return !1;
    }
    return !0;
  }
  return !1;
}
function fu() {
}
function du(e) {
  const t = e.toLowerCase().replace(/[-_\s.]+(.)?/g, (n, r) => r ? r.toUpperCase() : "");
  return t.substring(0, 1).toLowerCase() + t.substring(1);
}
function $n(e, t) {
  const n = { ...e };
  return t.forEach((r) => {
    delete n[r];
  }), n;
}
function Un(e) {
  if (typeof e == "boolean")
    return [!0, Number(e)];
  if (typeof e == "number" && !isNaN(e))
    return [!0, e];
  if (typeof e != "string")
    return [!1, 0];
  if (e.length > 12)
    return [!1, 0];
  const t = Number(e);
  return isNaN(t) ? [!1, 0] : [!0, t];
}
function hu(e = []) {
  const t = e.map((n) => Un(n)[1]).filter((n) => !isNaN(n));
  return Math.max(Math.max(...t), 0);
}
function mu(e = []) {
  const t = Math.ceil(hu(e.map((n) => n.sheetId))) + 1;
  return {
    name: `${tu}${t}`,
    sheetId: String(t),
    sort: t
  };
}
function ms(e) {
  if (typeof e == "string") {
    const r = e.toUpperCase();
    if (["TRUE", "FALSE"].includes(r))
      return r === "TRUE";
  }
  if (typeof e == "boolean")
    return e;
  if (e === "")
    return "";
  const [t, n] = Un(e);
  return t ? n : e;
}
function aN(e) {
  const [t, n] = e.split(Pn), r = parseInt(t, 10), o = parseInt(n, 10);
  return {
    row: isNaN(r) ? -1 : r,
    col: isNaN(o) ? -1 : o
  };
}
function nt(e, t, n) {
  return `${e}${Pn}${t}${Pn}${n}`;
}
function jn(e) {
  const [t, n, r] = e.split(Pn), o = parseInt(n, 10), s = parseInt(r, 10), l = {
    row: isNaN(o) ? -1 : o,
    col: isNaN(s) ? -1 : s,
    sheetId: t
  };
  return l.row < 0 || l.col < 0 || !t ? null : l;
}
function bt(e, t) {
  return `${e}${Pn}${t}`;
}
function ir(e) {
  const [t, n] = e.split(Pn), r = parseInt(n, 10);
  return {
    sheetId: t,
    rowOrCol: isNaN(r) ? -1 : r
  };
}
function lN() {
  const e = [
    /Android/i,
    /webOS/i,
    /iPhone/i,
    /iPad/i,
    /iPod/i,
    /BlackBerry/i,
    /Windows Phone/i
  ], t = navigator.userAgent;
  return e.some((n) => t.match(n));
}
const bi = [
  "autoFilter",
  "currentSheetId",
  "customHeight",
  "customWidth",
  "workbook",
  "worksheets",
  "definedNames",
  "drawings",
  "mergeCells",
  "rangeMap",
  "scroll"
];
function cN(e) {
  return new Promise((t) => setTimeout(t, e));
}
function va() {
  return process.env.NODE_ENV === "test";
}
function gu() {
  return navigator?.userAgent?.indexOf("Mac OS X") >= 0;
}
function Na(e) {
  return !!(typeof e == "string" && e && e.startsWith(Fn) && e.length > 1);
}
function wi(e, t) {
  return e && t.includes(mt);
}
function pu() {
  const e = Math.floor(Math.random() * 256).toString(16).padStart(2, "0"), t = Math.floor(Math.random() * 256).toString(16).padStart(2, "0"), n = Math.floor(Math.random() * 256).toString(16).padStart(2, "0");
  return `#${e}${t}${n}`;
}
function uN(e) {
  const t = atob(e);
  return Uint8Array.from(t, (n) => n.codePointAt(0));
}
function fN(e) {
  const t = Array.from(
    e,
    (n) => String.fromCodePoint(n)
  ).join("");
  return btoa(t);
}
function vu(e) {
  const t = /* @__PURE__ */ new Set(), n = new Set(bi);
  for (const r of e.changed.keys()) {
    const o = r._item?.parentSub;
    o && n.has(o) && t.add(o);
  }
  for (const r of e.changedParentTypes.keys()) {
    const o = r._item?.parentSub;
    o && n.has(o) && t.add(o);
  }
  return t;
}
const Nu = [
  6,
  8,
  9,
  10,
  Cn,
  14,
  16,
  18,
  20,
  22,
  24,
  26,
  28,
  36,
  48,
  72
];
function Ca(e = "normal", t = "normal", n = Cn, r = "") {
  const o = `${e} ${t} ${n}px `, s = [
    "Source Code Pro",
    r,
    ca
  ].filter((l) => l.trim());
  return o + s.join(",");
}
function Cu(e) {
  let t = "";
  return e.fontColor && (t += `color:${e.fontColor};`), e.fillColor && (t += `background-color:${e.fillColor};`), e.fontSize && (t += `font-size:${e.fontSize}pt;`), e.fontFamily && (t += `font-family:${e.fontFamily};`), e.isItalic && (t += "font-style:italic;"), e.isBold && (t += "font-weight:700;"), e.isWrapText && (t += "white-space:normal;"), e.underline && e.isStrike ? t += "text-decoration-line:underline line-through;" : e.underline ? t += "text-decoration-line:underline;" : e.isStrike && (t += "text-decoration-line:line-through;"), e.underline === Ce.DOUBLE && (t += "text-decoration-style: double;"), t;
}
function bo(e) {
  const {
    color: t,
    backgroundColor: n,
    fontSize: r,
    fontFamily: o,
    fontStyle: s,
    fontWeight: l,
    whiteSpace: f,
    textDecorationLine: d,
    textDecorationStyle: c
  } = e, i = {};
  if (t && (i.fontColor = t), n && (i.fillColor = n), r) {
    const a = parseInt(r, 10);
    isNaN(a) || (i.fontSize = a);
  }
  return o && (i.fontFamily = o), s === "italic" && (i.isItalic = !0), l && [700, 800, 900, "bold"].includes(l) && (i.isBold = !0), f && [
    "normal",
    "pre-wrap",
    "pre-line",
    "break-spaces",
    "revert",
    "unset"
  ].includes(f) && (i.isWrapText = !0), d?.includes("underline") && (i.underline = Ce.SINGLE, c === "double" && (i.underline = Ce.DOUBLE)), d?.includes("line-through") && (i.isStrike = !0), i;
}
function Eu(e, t, n, r = "") {
  let o = {};
  const s = r.toLowerCase(), l = e[r] || e[s];
  return l && (o = bo(l)), e[n] && (o = Object.assign(o, bo(e[n]))), o = Object.assign(o, bo(t)), s === "s" || s === "strike" ? o.isStrike = !0 : s === "i" ? o.isItalic = !0 : s === "b" || s === "strong" ? o.isBold = !0 : s === "u" && (o.underline = Ce.SINGLE), o;
}
function bu(e) {
  const t = e.replace(/\s+/g, "").replace("<!--", ""), n = /([^{}]+)\s*\{([^}]*)\}/g, r = {};
  let o;
  for (; (o = n.exec(t)) !== null; ) {
    const s = (o[1] || "").trim(), l = (o[2] || "").trim();
    if (!s || !l)
      continue;
    r[s] = {};
    const f = l.split(";");
    for (const d of f) {
      const [c, i] = d.split(":").map((m) => m.trim());
      if (!i || !c)
        continue;
      const a = du(c), [u, h] = Un(i), g = u ? h : i;
      r[s][a] = g;
    }
  }
  return r;
}
function gs(e) {
  const t = parseInt(e, 10);
  return isNaN(t) ? -1 : e.endsWith("pt") ? Math.ceil(t * 96 / 72) : t;
}
function ps(e, t = "pt") {
  const n = Math.floor(e * 72 / 96);
  return String(n) + t;
}
function Ea(e, t = "	") {
  let n;
  e.indexOf(`\r
`) >= 0 ? n = e.split(`\r
`) : n = e.split(`
`);
  const r = n.map((s) => s.trim()).filter((s) => s).map(
    (s) => s.split(t).map((l) => l.trim()).filter((l) => l)
  ), o = [];
  for (const s of r)
    s.length !== 0 && o.push(s);
  return o;
}
function wu(e) {
  let t = document.createElement("template");
  t.innerHTML = e;
  const n = t.content, r = {};
  for (const i of n.querySelectorAll("style")) {
    const a = bu(i.textContent || "");
    for (const [u, h] of Object.entries(a))
      r[u] = Object.assign(r[u] || {}, h);
  }
  const o = [], s = [], l = n.querySelectorAll("tr"), f = /* @__PURE__ */ new Map(), d = /* @__PURE__ */ new Map();
  let c = 0;
  for (const i of l) {
    const a = gs(i.style.height);
    a > 0 && d.set(c++, a);
    const u = [], h = [];
    let g = 0;
    for (const m of i.children) {
      if (m.tagName !== "TD")
        continue;
      let p = m;
      const E = gs(p.style.width);
      E > 0 && !f.has(g) && f.set(g++, E);
      let b = {};
      for (; p.nodeType !== Node.TEXT_NODE; ) {
        const N = Eu(
          r,
          p.style || {},
          p.className ? `.${p.className}` : "",
          p.tagName || ""
        );
        if (b = Object.assign(b, N), p.firstChild)
          p = p.firstChild;
        else
          break;
      }
      h.push(b), u.push((p.textContent ?? "").trim());
    }
    o.push(u), s.push(h);
  }
  return t = void 0, {
    textList: o,
    styleList: s,
    rowMap: d,
    colMap: f
  };
}
let Jo;
const Iu = 2, ba = (e) => Math.max(Math.floor(e || 1), Iu);
function sr(e = typeof devicePixelRatio > "u" ? void 0 : devicePixelRatio) {
  return typeof Jo == "number" ? Jo : ba(e);
}
function ce(e) {
  return Math.floor(e * sr());
}
function Mu(e) {
  Jo = ba(e);
}
function dN(e, t = "assert error") {
  if (!e)
    throw new Error(t);
}
function yu(e) {
  const t = e.toUpperCase();
  let n = 0;
  for (let r = 0; r < t.length; r++)
    n = t.charCodeAt(r) - 64 + n * 26;
  return n - 1;
}
function wa(e) {
  const t = e + 1;
  let n = "", r = Math.floor(Math.abs(t)), o;
  for (; r > 0; )
    o = (r - 1) % 26, n = String.fromCharCode(65 + o) + n, r = Math.floor((r - o) / 26);
  return n.toUpperCase();
}
function Su(e) {
  let t = parseInt(e, 10);
  return isNaN(t) ? t = -1 : t = Math.max(t - 1, -1), t;
}
function ae(...e) {
  let t = "";
  for (const n of e)
    if (n !== void 0 && (typeof n == "string" && n && (t += `${n} `), typeof n == "object"))
      for (const r of Object.keys(n))
        n[r] && (t += `${r} `);
  return t.trim();
}
function An(e) {
  return Wt(e) && Yt(e);
}
function Wt(e) {
  return e.colCount === 0;
}
function Yt(e) {
  return e.rowCount === 0;
}
function Ia(e, t) {
  return e.col === t.col && e.row === t.row && e.colCount === t.colCount && e.rowCount === t.rowCount && e.sheetId === t.sheetId;
}
function Ma(e, t) {
  if (e.colCount === 0 && t.colCount !== 0 || e.rowCount === 0 && t.rowCount !== 0)
    return !1;
  if (t.colCount === 0 && t.row === e.row || t.rowCount === 0 && t.col === e.col)
    return !0;
  const { row: n, col: r } = e;
  return n >= t.row && n < t.row + t.rowCount && r >= t.col && r < t.col + t.colCount;
}
function un(e) {
  return wt.makeRange(e).toIRange();
}
class wt {
  row = 0;
  col = 0;
  colCount = 0;
  rowCount = 0;
  sheetId = "";
  constructor(t, n, r, o, s) {
    this.row = t, this.col = n, this.colCount = o, this.rowCount = r, this.sheetId = s;
  }
  static makeRange(t) {
    return new wt(
      t.row,
      t.col,
      t.rowCount,
      t.colCount,
      t.sheetId
    );
  }
  toIRange() {
    return {
      row: this.row,
      col: this.col,
      rowCount: this.rowCount,
      colCount: this.colCount,
      sheetId: this.sheetId
    };
  }
}
function Tn(e, t, n, r) {
  const { row: o, col: s, rowCount: l, colCount: f } = e;
  if (An(e)) {
    for (let i = 0; i < t; i++)
      for (let a = 0; a < n; a++)
        if (r(i, a))
          return;
    return;
  }
  if (Wt(e)) {
    for (let i = 0; i < n; i++)
      if (r(o, i))
        return;
    return;
  }
  if (Yt(e)) {
    for (let i = 0; i < t; i++)
      if (r(i, s))
        return;
    return;
  }
  const d = o + l, c = s + f;
  for (let i = o; i < d; i++)
    for (let a = s; a < c; a++)
      if (r(i, a))
        return;
}
const Qo = (e) => e >= "a" && e <= "z" || e >= "A" && e <= "Z", Gt = (e) => e >= "0" && e <= "9";
function ya(e) {
  return e;
}
function vs(e, t) {
  if (!e)
    return;
  const n = e.trim();
  let [r, o = ""] = n.split("!");
  n.includes("!") || (r = "", o = n), r && r[0] === "'" && r[r.length - 1] === "'" && (r = r.slice(1, -1));
  let s = 0, l = "", f = "";
  for (o[s] === "$" && s++; s < o.length && Qo(o[s]); )
    f += o[s++];
  for (o[s] === "$" && s++; s < o.length && Gt(o[s]); )
    l += o[s++];
  if (s !== o.length || !l && !f)
    return;
  let d = 1, c = 1, i = -1, a = -1;
  return l === "" ? (d = 0, i = 0) : i = Su(l), f === "" ? (c = 0, a = 0) : a = yu(f), i < 0 || a < 0 || a >= hr || i >= mr ? void 0 : new wt(
    i,
    a,
    d,
    c,
    t(r)
  );
}
function Ot(e, t = ya) {
  const [n = "", r = ""] = e?.split(":") || [], o = vs(n, t);
  if (!o)
    return;
  const s = vs(r, t);
  return s ? (s.sheetId = o.sheetId, Ii(o, s)) : o;
}
function Ii(e, t) {
  if (e.sheetId !== t.sheetId)
    return;
  if (e.row === t.row && e.col === t.col && e.rowCount === t.rowCount && e.colCount === t.colCount)
    return wt.makeRange(e);
  if (e.rowCount === 0 && t.rowCount !== 0 || e.rowCount !== 0 && t.rowCount === 0 || e.colCount === 0 && t.colCount !== 0 || e.colCount !== 0 && t.colCount === 0)
    return;
  const n = Math.abs(e.row - t.row) + 1, r = Math.abs(e.col - t.col) + 1, o = e.row < t.row ? e.row : t.row, s = e.col < t.col ? e.col : t.col;
  return new wt(o, s, n, r, e.sheetId);
}
function Ns(e, t, n) {
  const r = n === "absolute" || n === "mixed", o = n === "absolute";
  return `${r ? "$" : ""}${wa(t)}${o ? "$" : ""}${e + 1}`;
}
function ut(e, t = "relative", n = ya) {
  let r = Ns(e.row, e.col, t), o = n(e.sheetId);
  if (o = o ? `${o}!` : "", e.colCount > 1 && e.rowCount > 1) {
    const s = Ns(
      e.row + e.rowCount - 1,
      e.col + e.colCount - 1,
      t
    );
    r = `${r}:${s}`;
  }
  return o + r;
}
const hN = /^R(\[-\d+\]|\[\d+\]|\d+)?C(\[-\d+\]|\[\d+\]|\d+)?$/i;
function Cs(e, t) {
  let n = -1;
  if (e.startsWith("[")) {
    const r = parseInt(e.slice(1, -1), 10);
    isNaN(r) || (n = t + r);
  } else {
    const r = parseInt(e, 10);
    isNaN(r) || (n = r - 1);
  }
  return isNaN(n) ? -1 : n;
}
function Ru(e, t = {
  row: -1,
  col: -1,
  sheetId: ""
}) {
  const n = e.toUpperCase();
  if (n[0] !== "R")
    return;
  const r = n.slice(1).split("C");
  if (r.length !== 2)
    return;
  const [o, s] = r;
  let l = -1, f = -1;
  return o ? l = Cs(o, t.row) : l = t.row, s ? f = Cs(s, t.col) : f = t.col, f >= hr || l >= mr || l < 0 || f < 0 ? void 0 : new wt(l, f, 1, 1, t.sheetId);
}
class Xt {
  namespace;
  static colorMap = /* @__PURE__ */ new Map();
  constructor(t) {
    this.namespace = t;
  }
  init = () => (this.setColor(), this.log);
  log = (...t) => {
    if (!this.enable())
      return;
    const { namespace: n } = this, r = Xt.colorMap.get(n), o = [`%c ${n}:`, `color:${r};`, ...t];
    console.log(...o);
  };
  enable() {
    return window.localStorage.getItem("debug");
  }
  setColor() {
    Xt.colorMap.has(this.namespace) || Xt.colorMap.set(this.namespace, pu());
  }
}
const mN = new Xt("model").init(), jo = new Xt("controller").init(), Au = new Xt("react").init(), gN = new Xt("collaboration").init(), Be = "text/plain", ze = "text/html", Ge = "custom/model", Qe = "image/png", Zo = "<caption>", Sa = "</caption>";
function Tu(e) {
  const t = e.hasAttribute("readonly");
  return t || e.setAttribute("readonly", ""), e.select(), e.setSelectionRange(0, e.value.length), t || e.removeAttribute("readonly"), e.value;
}
function xu(e) {
  const t = document.documentElement.getAttribute("dir") === "rtl", n = document.createElement("textarea");
  n.style.fontSize = "12pt", n.style.border = "0", n.style.padding = "0", n.style.margin = "0", n.style.position = "absolute", n.style[t ? "right" : "left"] = "-9999px";
  const r = window.pageYOffset || document.documentElement.scrollTop;
  return n.style.top = `${r}px`, n.setAttribute("readonly", ""), n.value = e, n;
}
function Du(e, t) {
  let n = xu(e);
  document.body.appendChild(n);
  const r = Tu(n);
  return typeof document.execCommand == "function" && document.execCommand(t), n.remove(), n = void 0, r;
}
async function Ou(e) {
  const t = new Blob([e[Be]], { type: Be }), n = new Blob([e[ze]], { type: ze }), r = {
    [Be]: t,
    [ze]: n
  };
  e[Qe] && (r[Qe] = e[Qe]), await navigator?.clipboard?.write([new ClipboardItem(r)]);
}
async function Es(e, t) {
  try {
    await Ou(e);
    return;
  } catch (n) {
    console.log(n);
  }
  try {
    Du(e[Be], t);
  } catch (n) {
    console.log(n);
  }
}
async function ku() {
  const e = {
    [ze]: "",
    [Be]: "",
    [Ge]: null,
    [Qe]: null
  }, t = await navigator?.clipboard?.read() || [];
  for (const n of t) {
    if (n.types.includes(Be)) {
      const r = await n.getType(Be);
      e[Be] = await r?.text?.() || "";
    }
    if (n.types.includes(ze)) {
      const r = await n.getType(ze);
      e[ze] = await r?.text?.() || "";
    }
    if (n.types.includes(Qe)) {
      const r = await n.getType(Qe);
      e[Qe] = r;
    }
  }
  return e[Ge] = e[Ge] || Pu(e[ze]), e;
}
async function Ra() {
  try {
    return await ku();
  } catch (e) {
    console.log(e);
  }
  return {
    [ze]: "",
    [Be]: "",
    [Ge]: null,
    [Qe]: null
  };
}
function Lu(e, t, n = "") {
  return `
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="ProgId" content="Excel.Sheet" />
    <meta name="Generator" content="Microsoft Excel 15" />
    <style>${e}</style>
  </head>
    <body>
      <table>${n}${t}</table>
    </body>
  </html>`.replaceAll(`
`, "");
}
function Fu(e) {
  return e ? `${Zo}${e}${Sa}` : "";
}
function Pu(e) {
  const t = e.indexOf(Zo), n = e.indexOf(Sa);
  if (t >= 0 && n >= 0) {
    const r = e.slice(t + Zo.length, n).trim();
    if (!r || r[0] !== "{" || r[r.length - 1] !== "}")
      return null;
    try {
      return JSON.parse(r);
    } catch (o) {
      return console.log(o), null;
    }
  }
  return null;
}
const ei = "QUERY_ALL_LOCAL_FONT", Aa = "LOCAL_FONT_KEY", Uu = [
  "PingFang SC",
  "Microsoft Yahei",
  "SimSun",
  "Times New Roman",
  "Source Code Pro",
  "等线",
  "方正舒体",
  "仿宋",
  "黑体",
  "楷体",
  "隶书",
  "宋体",
  "微软雅黑",
  "新宋体",
  "Arial",
  "Calibri"
];
function Ta(e, t, n) {
  if (t < 0 || n < 0)
    return;
  const r = e.getSheetInfo(e.getCurrentSheetId());
  if (!r)
    return;
  const o = e.getScroll(), s = e.getHeaderSize();
  let l = s.width, f = s.height, d = t - s.width, c = n - s.height, { row: i, col: a } = o;
  for (; l + e.getColWidth(a) <= t; ) {
    const u = e.getColWidth(a);
    l += u, d -= u, a++;
  }
  for (; f + e.getRowHeight(i) <= n; ) {
    const u = e.getRowHeight(i);
    f += u, c -= u, i++;
  }
  if (!(i >= r.rowCount || a >= r.colCount))
    return { row: i, col: a, marginX: d, marginY: c };
}
function ar(e, t) {
  let n = document.createElement("a");
  n.download = t || "download", n.rel = "noopener", n.target = "_blank", n.setAttribute("data-testid", "save-as-link"), typeof e == "string" ? n.href = e : n.href = URL.createObjectURL(e), n.dispatchEvent(new MouseEvent("click")), n = void 0;
}
const bs = {
  maroon: "#800000",
  red: "#ff0000",
  orange: "#ffA500",
  yellow: "#ffff00",
  olive: "#808000",
  purple: "#800080",
  fuchsia: "#ff00ff",
  white: "#ffffff",
  lime: "#00ff00",
  green: "#008000",
  navy: "#000080",
  blue: "#0000ff",
  aqua: "#00ffff",
  teal: "#008080",
  black: "#000000",
  silver: "#c0c0c0",
  gray: "#808080",
  transparent: "#0000",
  cyan: "#F00",
  magenta: "#00F"
}, _u = "(?:#([a-f0-9]{3,8}))", wo = "\\s*([.\\d%]+)\\s*", Bu = "(?:,\\s*([.\\d]+)\\s*)?", xa = `\\(${[wo, wo, wo]}${Bu}\\)`, zu = "(?:rgb)a?", Hu = "(?:hsl)a?", Da = RegExp(_u, "i"), ws = RegExp(zu + xa, "i"), Is = RegExp(Hu + xa, "i");
function Vu(e) {
  return e.length === 1 ? `0${e}` : e;
}
function Oa(e, t, n, r) {
  return `#${[e, t, n, r * 255].map((s) => Vu(Math.round(s).toString(16))).join("")}`;
}
function Gu(e, t, n, r) {
  let o = 1;
  r && (o = parseFloat(r));
  const s = parseFloat(e), l = parseFloat(t), f = parseFloat(n);
  return s <= 255 && s >= 0 && l <= 255 && l >= 0 && f <= 255 && f >= 0 && o <= 1 && o >= 0 ? Oa(s, l, f, o) : "";
}
function Io(e, t, n) {
  return n < 0 && (n += 1), n > 1 && (n -= 1), n < 1 / 6 ? e + (t - e) * 6 * n : n < 1 / 2 ? t : n < 2 / 3 ? e + (t - e) * (2 / 3 - n) * 6 : e;
}
function $u(e, t, n, r) {
  let o = 1;
  r && (o = parseFloat(r) / 1);
  const s = parseFloat(e) / 360, l = parseFloat(t) / 360, f = parseFloat(n) / 360;
  if (s > 1 || s < 0 || l > 1 || l < 0 || f > 1 || f < 0 || o > 1 || o < 0)
    return "";
  let d, c, i;
  if (l === 0)
    d = f, c = f, i = f;
  else {
    const a = f < 0.5 ? f * (1 + l) : f + l - f * l, u = 2 * f - a;
    d = Io(u, a, s + 1 / 3), c = Io(u, a, s), i = Io(u, a, s - 1 / 3);
  }
  return Oa(d * 255, c * 255, i * 255, o);
}
function wr(e) {
  if (!Da.test(e))
    return "";
  if (e.length === 9)
    return e.toUpperCase();
  if (e.length === 7)
    return `${e}ff`.toUpperCase();
  if (e.length === 4) {
    const [, t, n, r] = e;
    return `#${t}${t}${n}${n}${r}${r}ff`.toUpperCase();
  }
  return "";
}
function ti(e) {
  if (bs[e])
    return wr(bs[e]);
  if (Da.test(e))
    return wr(e);
  if (ws.test(e)) {
    const [, t, n, r, o] = e.match(ws);
    return wr(Gu(t, n, r, o));
  }
  if (Is.test(e)) {
    const [, t, n, r, o] = e.match(Is);
    return wr($u(t, n, r, o));
  }
  return "";
}
class ka {
  event;
  constructor() {
    this.event = {};
  }
  getEventLength(t) {
    return this.event[t]?.length || 0;
  }
  on = (t, n) => (this.event[t] = this.event[t] || [], this.event[t].push(n), () => this.off(t, n));
  emit = (t, n) => {
    const r = this.event[t];
    if (!(!r || r.length <= 0))
      for (const o of r)
        o(n);
  };
  off = (t, n) => {
    const r = [], o = this.event[t];
    if (o && n)
      for (const s of o)
        s !== n && s._ !== n && r.push(s);
    r.length ? this.event[t] = r : delete this.event[t];
  };
  once = (t, n) => {
    const r = (o) => {
      this.off(t, r), n(o);
    };
    return r._ = n, this.on(t, r);
  };
}
const La = {
  "image/apng": [".apng"],
  "image/bmp": [".bmp"],
  "image/x-icon": [".ico", ".cur"],
  // 'image/tiff': ['.tif', '.tiff'], // only Safari
  "image/png": [".png"],
  "image/webp": [".webp"],
  "image/svg+xml": [".svg"],
  "image/avif": [".avif"],
  "image/gif": [".gif"],
  "image/jpeg": [".jpeg", ".jpg", ".jfif", ".pjpeg", ".pjp"]
};
function Fa(e) {
  if (!e)
    return { ext: "", base64: "", type: "" };
  const t = ";base64,", n = e.indexOf(t), o = e.slice("data:".length, n), s = La[o] || [];
  return s.length > 0 ? { ext: s[0], type: o, base64: e.slice(n + t.length) } : { ext: "", base64: "", type: "" };
}
function Mi(e, t = !1) {
  return new Promise((n, r) => {
    const o = new FileReader();
    o.onload = function(s) {
      const l = s.target?.result;
      n(l && typeof l == "string" ? l : "");
    }, o.onerror = function(s) {
      r(s);
    }, t ? o.readAsDataURL(e) : o.readAsText(e);
  });
}
function yi(e) {
  return new Promise((t, n) => {
    const r = new Image();
    r.onload = () => {
      const o = r.width, s = r.height;
      t({ width: o, height: s });
    }, r.onerror = (o) => {
      n(o);
    }, r.src = e;
  });
}
function qu(e, t = "image/png") {
  const r = atob(e), o = r.length, s = Math.ceil(o / 1024), l = Array.from({ length: s });
  for (let f = 0; f < s; ++f) {
    let d = f * 1024, c = Math.min(d + 1024, o), i = Array.from({ length: c - d });
    for (let a = d, u = 0; a < c; ++u, ++a)
      i[u] = r[a].charCodeAt(0);
    l[f] = new Uint8Array(i);
  }
  return new Blob(l, { type: t });
}
function Ms(e) {
  return e ? typeof Intl > "u" || typeof Intl.Segmenter != "function" ? [...e] : [...new Intl.Segmenter([], { granularity: "word" }).segment(e)].map((r) => r.segment) : [];
}
var pN = typeof globalThis < "u" ? globalThis : typeof window < "u" ? window : typeof global < "u" ? global : typeof self < "u" ? self : {};
function Pa(e) {
  return e && e.__esModule && Object.prototype.hasOwnProperty.call(e, "default") ? e.default : e;
}
var _r = { exports: {} }, Wu = _r.exports, ys;
function Yu() {
  return ys || (ys = 1, (function(e, t) {
    (function(n, r) {
      e.exports = r();
    })(Wu, function() {
      var n = (function(r, o) {
        var s = Array.prototype.concat, l = Array.prototype.slice, f = Object.prototype.toString;
        function d(b, N) {
          var v = b > N ? b : N;
          return r.pow(
            10,
            17 - ~~(r.log(v > 0 ? v : -v) * r.LOG10E)
          );
        }
        var c = Array.isArray || function(N) {
          return f.call(N) === "[object Array]";
        };
        function i(b) {
          return f.call(b) === "[object Function]";
        }
        function a(b) {
          return typeof b == "number" ? b - b === 0 : !1;
        }
        function u(b) {
          return s.apply([], b);
        }
        function h() {
          return new h._init(arguments);
        }
        h.fn = h.prototype, h._init = function(N) {
          if (c(N[0]))
            if (c(N[0][0])) {
              i(N[1]) && (N[0] = h.map(N[0], N[1]));
              for (var v = 0; v < N[0].length; v++)
                this[v] = N[0][v];
              this.length = N[0].length;
            } else
              this[0] = i(N[1]) ? h.map(N[0], N[1]) : N[0], this.length = 1;
          else if (a(N[0]))
            this[0] = h.seq.apply(null, N), this.length = 1;
          else {
            if (N[0] instanceof h)
              return h(N[0].toArray());
            this[0] = [], this.length = 1;
          }
          return this;
        }, h._init.prototype = h.prototype, h._init.constructor = h, h.utils = {
          calcRdx: d,
          isArray: c,
          isFunction: i,
          isNumber: a,
          toVector: u
        }, h._random_fn = r.random, h.setRandom = function(N) {
          if (typeof N != "function")
            throw new TypeError("fn is not a function");
          h._random_fn = N;
        }, h.extend = function(N) {
          var v, C;
          if (arguments.length === 1) {
            for (C in N)
              h[C] = N[C];
            return this;
          }
          for (v = 1; v < arguments.length; v++)
            for (C in arguments[v])
              N[C] = arguments[v][C];
          return N;
        }, h.rows = function(N) {
          return N.length || 1;
        }, h.cols = function(N) {
          return N[0].length || 1;
        }, h.dimensions = function(N) {
          return {
            rows: h.rows(N),
            cols: h.cols(N)
          };
        }, h.row = function(N, v) {
          return c(v) ? v.map(function(C) {
            return h.row(N, C);
          }) : N[v];
        }, h.rowa = function(N, v) {
          return h.row(N, v);
        }, h.col = function(N, v) {
          if (c(v)) {
            var C = h.arange(N.length).map(function() {
              return new Array(v.length);
            });
            return v.forEach(function(A, x) {
              h.arange(N.length).forEach(function(L) {
                C[L][x] = N[L][A];
              });
            }), C;
          }
          for (var M = new Array(N.length), R = 0; R < N.length; R++)
            M[R] = [N[R][v]];
          return M;
        }, h.cola = function(N, v) {
          return h.col(N, v).map(function(C) {
            return C[0];
          });
        }, h.diag = function(N) {
          for (var v = h.rows(N), C = new Array(v), M = 0; M < v; M++)
            C[M] = [N[M][M]];
          return C;
        }, h.antidiag = function(N) {
          for (var v = h.rows(N) - 1, C = new Array(v), M = 0; v >= 0; v--, M++)
            C[M] = [N[M][v]];
          return C;
        }, h.transpose = function(N) {
          var v = [], C, M, R, A, x;
          for (c(N[0]) || (N = [N]), M = N.length, R = N[0].length, x = 0; x < R; x++) {
            for (C = new Array(M), A = 0; A < M; A++)
              C[A] = N[A][x];
            v.push(C);
          }
          return v.length === 1 ? v[0] : v;
        }, h.map = function(N, v, C) {
          var M, R, A, x, L;
          for (c(N[0]) || (N = [N]), R = N.length, A = N[0].length, x = C ? N : new Array(R), M = 0; M < R; M++)
            for (x[M] || (x[M] = new Array(A)), L = 0; L < A; L++)
              x[M][L] = v(N[M][L], M, L);
          return x.length === 1 ? x[0] : x;
        }, h.cumreduce = function(N, v, C) {
          var M, R, A, x, L;
          for (c(N[0]) || (N = [N]), R = N.length, A = N[0].length, x = C ? N : new Array(R), M = 0; M < R; M++)
            for (x[M] || (x[M] = new Array(A)), A > 0 && (x[M][0] = N[M][0]), L = 1; L < A; L++)
              x[M][L] = v(x[M][L - 1], N[M][L]);
          return x.length === 1 ? x[0] : x;
        }, h.alter = function(N, v) {
          return h.map(N, v, !0);
        }, h.create = function(N, v, C) {
          var M = new Array(N), R, A;
          for (i(v) && (C = v, v = N), R = 0; R < N; R++)
            for (M[R] = new Array(v), A = 0; A < v; A++)
              M[R][A] = C(R, A);
          return M;
        };
        function g() {
          return 0;
        }
        h.zeros = function(N, v) {
          return a(v) || (v = N), h.create(N, v, g);
        };
        function m() {
          return 1;
        }
        h.ones = function(N, v) {
          return a(v) || (v = N), h.create(N, v, m);
        }, h.rand = function(N, v) {
          return a(v) || (v = N), h.create(N, v, h._random_fn);
        };
        function p(b, N) {
          return b === N ? 1 : 0;
        }
        h.identity = function(N, v) {
          return a(v) || (v = N), h.create(N, v, p);
        }, h.symmetric = function(N) {
          var v = N.length, C, M;
          if (N.length !== N[0].length)
            return !1;
          for (C = 0; C < v; C++)
            for (M = 0; M < v; M++)
              if (N[M][C] !== N[C][M])
                return !1;
          return !0;
        }, h.clear = function(N) {
          return h.alter(N, g);
        }, h.seq = function(N, v, C, M) {
          i(M) || (M = !1);
          var R = [], A = d(N, v), x = (v * A - N * A) / ((C - 1) * A), L = N, D;
          for (D = 0; L <= v && D < C; D++, L = (N * A + x * A * D) / A)
            R.push(M ? M(L, D) : L);
          return R;
        }, h.arange = function(N, v, C) {
          var M = [], R;
          if (C = C || 1, v === o && (v = N, N = 0), N === v || C === 0)
            return [];
          if (N < v && C < 0)
            return [];
          if (N > v && C > 0)
            return [];
          if (C > 0)
            for (R = N; R < v; R += C)
              M.push(R);
          else
            for (R = N; R > v; R += C)
              M.push(R);
          return M;
        }, h.slice = /* @__PURE__ */ (function() {
          function b(v, C, M, R) {
            var A, x = [], L = v.length;
            if (C === o && M === o && R === o)
              return h.copy(v);
            if (C = C || 0, M = M || v.length, C = C >= 0 ? C : L + C, M = M >= 0 ? M : L + M, R = R || 1, C === M || R === 0)
              return [];
            if (C < M && R < 0)
              return [];
            if (C > M && R > 0)
              return [];
            if (R > 0)
              for (A = C; A < M; A += R)
                x.push(v[A]);
            else
              for (A = C; A > M; A += R)
                x.push(v[A]);
            return x;
          }
          function N(v, C) {
            var M, R;
            if (C = C || {}, a(C.row)) {
              if (a(C.col))
                return v[C.row][C.col];
              var A = h.rowa(v, C.row);
              return M = C.col || {}, b(A, M.start, M.end, M.step);
            }
            if (a(C.col)) {
              var x = h.cola(v, C.col);
              return R = C.row || {}, b(x, R.start, R.end, R.step);
            }
            R = C.row || {}, M = C.col || {};
            var L = b(v, R.start, R.end, R.step);
            return L.map(function(D) {
              return b(D, M.start, M.end, M.step);
            });
          }
          return N;
        })(), h.sliceAssign = function(N, v, C) {
          var M, R;
          if (a(v.row)) {
            if (a(v.col))
              return N[v.row][v.col] = C;
            v.col = v.col || {}, v.col.start = v.col.start || 0, v.col.end = v.col.end || N[0].length, v.col.step = v.col.step || 1, M = h.arange(
              v.col.start,
              r.min(N.length, v.col.end),
              v.col.step
            );
            var A = v.row;
            return M.forEach(function(L, D) {
              N[A][L] = C[D];
            }), N;
          }
          if (a(v.col)) {
            v.row = v.row || {}, v.row.start = v.row.start || 0, v.row.end = v.row.end || N.length, v.row.step = v.row.step || 1, R = h.arange(
              v.row.start,
              r.min(N[0].length, v.row.end),
              v.row.step
            );
            var x = v.col;
            return R.forEach(function(L, D) {
              N[L][x] = C[D];
            }), N;
          }
          return C[0].length === o && (C = [C]), v.row.start = v.row.start || 0, v.row.end = v.row.end || N.length, v.row.step = v.row.step || 1, v.col.start = v.col.start || 0, v.col.end = v.col.end || N[0].length, v.col.step = v.col.step || 1, R = h.arange(
            v.row.start,
            r.min(N.length, v.row.end),
            v.row.step
          ), M = h.arange(
            v.col.start,
            r.min(N[0].length, v.col.end),
            v.col.step
          ), R.forEach(function(L, D) {
            M.forEach(function(W, J) {
              N[L][W] = C[D][J];
            });
          }), N;
        }, h.diagonal = function(N) {
          var v = h.zeros(N.length, N.length);
          return N.forEach(function(C, M) {
            v[M][M] = C;
          }), v;
        }, h.copy = function(N) {
          return N.map(function(v) {
            return a(v) ? v : v.map(function(C) {
              return C;
            });
          });
        };
        var E = h.prototype;
        return E.length = 0, E.push = Array.prototype.push, E.sort = Array.prototype.sort, E.splice = Array.prototype.splice, E.slice = Array.prototype.slice, E.toArray = function() {
          return this.length > 1 ? l.call(this) : l.call(this)[0];
        }, E.map = function(N, v) {
          return h(h.map(this, N, v));
        }, E.cumreduce = function(N, v) {
          return h(h.cumreduce(this, N, v));
        }, E.alter = function(N) {
          return h.alter(this, N), this;
        }, (function(b) {
          for (var N = 0; N < b.length; N++) (function(v) {
            E[v] = function(C) {
              var M = this, R;
              return C ? (setTimeout(function() {
                C.call(M, E[v].call(M));
              }), this) : (R = h[v](this), c(R) ? h(R) : R);
            };
          })(b[N]);
        })("transpose clear symmetric rows cols dimensions diag antidiag".split(" ")), (function(b) {
          for (var N = 0; N < b.length; N++) (function(v) {
            E[v] = function(C, M) {
              var R = this;
              return M ? (setTimeout(function() {
                M.call(R, E[v].call(R, C));
              }), this) : h(h[v](this, C));
            };
          })(b[N]);
        })("row col".split(" ")), (function(b) {
          for (var N = 0; N < b.length; N++) (function(v) {
            E[v] = function() {
              return h(h[v].apply(null, arguments));
            };
          })(b[N]);
        })("create zeros ones rand identity".split(" ")), h;
      })(Math);
      return (function(r, o) {
        var s = r.utils.isFunction;
        function l(c, i) {
          return c - i;
        }
        function f(c, i, a) {
          return o.max(i, o.min(c, a));
        }
        r.sum = function(i) {
          for (var a = 0, u = i.length; --u >= 0; )
            a += i[u];
          return a;
        }, r.sumsqrd = function(i) {
          for (var a = 0, u = i.length; --u >= 0; )
            a += i[u] * i[u];
          return a;
        }, r.sumsqerr = function(i) {
          for (var a = r.mean(i), u = 0, h = i.length, g; --h >= 0; )
            g = i[h] - a, u += g * g;
          return u;
        }, r.sumrow = function(i) {
          for (var a = 0, u = i.length; --u >= 0; )
            a += i[u];
          return a;
        }, r.product = function(i) {
          for (var a = 1, u = i.length; --u >= 0; )
            a *= i[u];
          return a;
        }, r.min = function(i) {
          for (var a = i[0], u = 0; ++u < i.length; )
            i[u] < a && (a = i[u]);
          return a;
        }, r.max = function(i) {
          for (var a = i[0], u = 0; ++u < i.length; )
            i[u] > a && (a = i[u]);
          return a;
        }, r.unique = function(i) {
          for (var a = {}, u = [], h = 0; h < i.length; h++)
            a[i[h]] || (a[i[h]] = !0, u.push(i[h]));
          return u;
        }, r.mean = function(i) {
          return r.sum(i) / i.length;
        }, r.meansqerr = function(i) {
          return r.sumsqerr(i) / i.length;
        }, r.geomean = function(i) {
          var a = i.map(o.log), u = r.mean(a);
          return o.exp(u);
        }, r.median = function(i) {
          var a = i.length, u = i.slice().sort(l);
          return a & 1 ? u[a / 2 | 0] : (u[a / 2 - 1] + u[a / 2]) / 2;
        }, r.cumsum = function(i) {
          return r.cumreduce(i, function(a, u) {
            return a + u;
          });
        }, r.cumprod = function(i) {
          return r.cumreduce(i, function(a, u) {
            return a * u;
          });
        }, r.diff = function(i) {
          var a = [], u = i.length, h;
          for (h = 1; h < u; h++)
            a.push(i[h] - i[h - 1]);
          return a;
        }, r.rank = function(c) {
          var i, a = [], u = {};
          for (i = 0; i < c.length; i++) {
            var h = c[i];
            u[h] ? u[h]++ : (u[h] = 1, a.push(h));
          }
          var g = a.sort(l), m = {}, p = 1;
          for (i = 0; i < g.length; i++) {
            var h = g[i], E = u[h], b = p, N = p + E - 1, v = (b + N) / 2;
            m[h] = v, p += E;
          }
          return c.map(function(C) {
            return m[C];
          });
        }, r.mode = function(i) {
          var a = i.length, u = i.slice().sort(l), h = 1, g = 0, m = 0, p = [], E;
          for (E = 0; E < a; E++)
            u[E] === u[E + 1] ? h++ : (h > g ? (p = [u[E]], g = h, m = 0) : h === g && (p.push(u[E]), m++), h = 1);
          return m === 0 ? p[0] : p;
        }, r.range = function(i) {
          return r.max(i) - r.min(i);
        }, r.variance = function(i, a) {
          return r.sumsqerr(i) / (i.length - (a ? 1 : 0));
        }, r.pooledvariance = function(i) {
          var a = i.reduce(function(h, g) {
            return h + r.sumsqerr(g);
          }, 0), u = i.reduce(function(h, g) {
            return h + g.length;
          }, 0);
          return a / (u - i.length);
        }, r.deviation = function(c) {
          for (var i = r.mean(c), a = c.length, u = new Array(a), h = 0; h < a; h++)
            u[h] = c[h] - i;
          return u;
        }, r.stdev = function(i, a) {
          return o.sqrt(r.variance(i, a));
        }, r.pooledstdev = function(i) {
          return o.sqrt(r.pooledvariance(i));
        }, r.meandev = function(i) {
          for (var a = r.mean(i), u = [], h = i.length - 1; h >= 0; h--)
            u.push(o.abs(i[h] - a));
          return r.mean(u);
        }, r.meddev = function(i) {
          for (var a = r.median(i), u = [], h = i.length - 1; h >= 0; h--)
            u.push(o.abs(i[h] - a));
          return r.median(u);
        }, r.coeffvar = function(i) {
          return r.stdev(i) / r.mean(i);
        }, r.quartiles = function(i) {
          var a = i.length, u = i.slice().sort(l);
          return [
            u[o.round(a / 4) - 1],
            u[o.round(a / 2) - 1],
            u[o.round(a * 3 / 4) - 1]
          ];
        }, r.quantiles = function(i, a, u, h) {
          var g = i.slice().sort(l), m = [a.length], p = i.length, E, b, N, v, C, M;
          for (typeof u > "u" && (u = 3 / 8), typeof h > "u" && (h = 3 / 8), E = 0; E < a.length; E++)
            b = a[E], N = u + b * (1 - u - h), v = p * b + N, C = o.floor(f(v, 1, p - 1)), M = f(v - C, 0, 1), m[E] = (1 - M) * g[C - 1] + M * g[C];
          return m;
        }, r.percentile = function(i, a, u) {
          var h = i.slice().sort(l), g = a * (h.length + (u ? 1 : -1)) + (u ? 0 : 1), m = parseInt(g), p = g - m;
          return m + 1 < h.length ? h[m - 1] + p * (h[m] - h[m - 1]) : h[m - 1];
        }, r.percentileOfScore = function(i, a, u) {
          var h = 0, g = i.length, m = !1, p, E;
          for (u === "strict" && (m = !0), E = 0; E < g; E++)
            p = i[E], (m && p < a || !m && p <= a) && h++;
          return h / g;
        }, r.histogram = function(i, a) {
          a = a || 4;
          var u = r.min(i), h = (r.max(i) - u) / a, g = i.length, m = [], p;
          for (p = 0; p < a; p++)
            m[p] = 0;
          for (p = 0; p < g; p++)
            m[o.min(o.floor((i[p] - u) / h), a - 1)] += 1;
          return m;
        }, r.covariance = function(i, a) {
          var u = r.mean(i), h = r.mean(a), g = i.length, m = new Array(g), p;
          for (p = 0; p < g; p++)
            m[p] = (i[p] - u) * (a[p] - h);
          return r.sum(m) / (g - 1);
        }, r.corrcoeff = function(i, a) {
          return r.covariance(i, a) / r.stdev(i, 1) / r.stdev(a, 1);
        }, r.spearmancoeff = function(c, i) {
          return c = r.rank(c), i = r.rank(i), r.corrcoeff(c, i);
        }, r.stanMoment = function(i, a) {
          for (var u = r.mean(i), h = r.stdev(i), g = i.length, m = 0, p = 0; p < g; p++)
            m += o.pow((i[p] - u) / h, a);
          return m / i.length;
        }, r.skewness = function(i) {
          return r.stanMoment(i, 3);
        }, r.kurtosis = function(i) {
          return r.stanMoment(i, 4) - 3;
        };
        var d = r.prototype;
        (function(c) {
          for (var i = 0; i < c.length; i++) (function(a) {
            d[a] = function(u, h) {
              var g = [], m = 0, p = this;
              if (s(u) && (h = u, u = !1), h)
                return setTimeout(function() {
                  h.call(p, d[a].call(p, u));
                }), this;
              if (this.length > 1) {
                for (p = u === !0 ? this : this.transpose(); m < p.length; m++)
                  g[m] = r[a](p[m]);
                return g;
              }
              return r[a](this[0], u);
            };
          })(c[i]);
        })("cumsum cumprod".split(" ")), (function(c) {
          for (var i = 0; i < c.length; i++) (function(a) {
            d[a] = function(u, h) {
              var g = [], m = 0, p = this;
              if (s(u) && (h = u, u = !1), h)
                return setTimeout(function() {
                  h.call(p, d[a].call(p, u));
                }), this;
              if (this.length > 1) {
                for (a !== "sumrow" && (p = u === !0 ? this : this.transpose()); m < p.length; m++)
                  g[m] = r[a](p[m]);
                return u === !0 ? r[a](r.utils.toVector(g)) : g;
              }
              return r[a](this[0], u);
            };
          })(c[i]);
        })("sum sumsqrd sumsqerr sumrow product min max unique mean meansqerr geomean median diff rank mode range variance deviation stdev meandev meddev coeffvar quartiles histogram skewness kurtosis".split(" ")), (function(c) {
          for (var i = 0; i < c.length; i++) (function(a) {
            d[a] = function() {
              var u = [], h = 0, g = this, m = Array.prototype.slice.call(arguments), p;
              if (s(m[m.length - 1])) {
                p = m[m.length - 1];
                var E = m.slice(0, m.length - 1);
                return setTimeout(function() {
                  p.call(
                    g,
                    d[a].apply(g, E)
                  );
                }), this;
              } else {
                p = void 0;
                var b = function(v) {
                  return r[a].apply(g, [v].concat(m));
                };
              }
              if (this.length > 1) {
                for (g = g.transpose(); h < g.length; h++)
                  u[h] = b(g[h]);
                return u;
              }
              return b(this[0]);
            };
          })(c[i]);
        })("quantiles percentileOfScore".split(" "));
      })(n, Math), (function(r, o) {
        r.gammaln = function(l) {
          var f = 0, d = [
            76.18009172947146,
            -86.50532032941678,
            24.01409824083091,
            -1.231739572450155,
            0.001208650973866179,
            -5395239384953e-18
          ], c = 1.000000000190015, i, a, u;
          for (u = (a = i = l) + 5.5, u -= (i + 0.5) * o.log(u); f < 6; f++)
            c += d[f] / ++a;
          return o.log(2.5066282746310007 * c / i) - u;
        }, r.loggam = function(l) {
          var f, d, c, i, a, u, h, g = [
            0.08333333333333333,
            -0.002777777777777778,
            7936507936507937e-19,
            -5952380952380952e-19,
            8417508417508418e-19,
            -0.001917526917526918,
            0.00641025641025641,
            -0.02955065359477124,
            0.1796443723688307,
            -1.3924322169059
          ];
          if (f = l, h = 0, l == 1 || l == 2)
            return 0;
          for (l <= 7 && (h = o.floor(7 - l), f = l + h), d = 1 / (f * f), c = 2 * o.PI, a = g[9], u = 8; u >= 0; u--)
            a *= d, a += g[u];
          if (i = a / f + 0.5 * o.log(c) + (f - 0.5) * o.log(f) - f, l <= 7)
            for (u = 1; u <= h; u++)
              i -= o.log(f - 1), f -= 1;
          return i;
        }, r.gammafn = function(l) {
          var f = [
            -1.716185138865495,
            24.76565080557592,
            -379.80425647094563,
            629.3311553128184,
            866.9662027904133,
            -31451.272968848367,
            -36144.413418691176,
            66456.14382024054
          ], d = [
            -30.8402300119739,
            315.35062697960416,
            -1015.1563674902192,
            -3107.771671572311,
            22538.11842098015,
            4755.846277527881,
            -134659.9598649693,
            -115132.2596755535
          ], c = !1, i = 0, a = 0, u = 0, h = l, g, m, p, E;
          if (l > 171.6243769536076)
            return 1 / 0;
          if (h <= 0)
            if (E = h % 1 + 36e-17, E)
              c = (h & 1 ? -1 : 1) * o.PI / o.sin(o.PI * E), h = 1 - h;
            else
              return 1 / 0;
          for (p = h, h < 1 ? m = h++ : m = (h -= i = (h | 0) - 1) - 1, g = 0; g < 8; ++g)
            u = (u + f[g]) * m, a = a * m + d[g];
          if (E = u / a + 1, p < h)
            E /= p;
          else if (p > h)
            for (g = 0; g < i; ++g)
              E *= h, h++;
          return c && (E = c / E), E;
        }, r.gammap = function(l, f) {
          return r.lowRegGamma(l, f) * r.gammafn(l);
        }, r.lowRegGamma = function(l, f) {
          var d = r.gammaln(l), c = l, i = 1 / l, a = i, u = f + 1 - l, h = 1 / 1e-30, g = 1 / u, m = g, p = 1, E = -~(o.log(l >= 1 ? l : 1 / l) * 8.5 + l * 0.4 + 17), b;
          if (f < 0 || l <= 0)
            return NaN;
          if (f < l + 1) {
            for (; p <= E; p++)
              i += a *= f / ++c;
            return i * o.exp(-f + l * o.log(f) - d);
          }
          for (; p <= E; p++)
            b = -p * (p - l), u += 2, g = b * g + u, h = u + b / h, g = 1 / g, m *= g * h;
          return 1 - m * o.exp(-f + l * o.log(f) - d);
        }, r.factorialln = function(l) {
          return l < 0 ? NaN : r.gammaln(l + 1);
        }, r.factorial = function(l) {
          return l < 0 ? NaN : r.gammafn(l + 1);
        }, r.combination = function(l, f) {
          return l > 170 || f > 170 ? o.exp(r.combinationln(l, f)) : r.factorial(l) / r.factorial(f) / r.factorial(l - f);
        }, r.combinationln = function(l, f) {
          return r.factorialln(l) - r.factorialln(f) - r.factorialln(l - f);
        }, r.permutation = function(l, f) {
          return r.factorial(l) / r.factorial(l - f);
        }, r.betafn = function(l, f) {
          if (!(l <= 0 || f <= 0))
            return l + f > 170 ? o.exp(r.betaln(l, f)) : r.gammafn(l) * r.gammafn(f) / r.gammafn(l + f);
        }, r.betaln = function(l, f) {
          return r.gammaln(l) + r.gammaln(f) - r.gammaln(l + f);
        }, r.betacf = function(l, f, d) {
          var c = 1e-30, i = 1, a = f + d, u = f + 1, h = f - 1, g = 1, m = 1 - a * l / u, p, E, b, N;
          for (o.abs(m) < c && (m = c), m = 1 / m, N = m; i <= 100 && (p = 2 * i, E = i * (d - i) * l / ((h + p) * (f + p)), m = 1 + E * m, o.abs(m) < c && (m = c), g = 1 + E / g, o.abs(g) < c && (g = c), m = 1 / m, N *= m * g, E = -(f + i) * (a + i) * l / ((f + p) * (u + p)), m = 1 + E * m, o.abs(m) < c && (m = c), g = 1 + E / g, o.abs(g) < c && (g = c), m = 1 / m, b = m * g, N *= b, !(o.abs(b - 1) < 3e-7)); i++)
            ;
          return N;
        }, r.gammapinv = function(l, f) {
          var d = 0, c = f - 1, i = 1e-8, a = r.gammaln(f), u, h, g, m, p, E, b;
          if (l >= 1)
            return o.max(100, f + 100 * o.sqrt(f));
          if (l <= 0)
            return 0;
          for (f > 1 ? (E = o.log(c), b = o.exp(c * (E - 1) - a), p = l < 0.5 ? l : 1 - l, g = o.sqrt(-2 * o.log(p)), u = (2.30753 + g * 0.27061) / (1 + g * (0.99229 + g * 0.04481)) - g, l < 0.5 && (u = -u), u = o.max(
            1e-3,
            f * o.pow(1 - 1 / (9 * f) - u / (3 * o.sqrt(f)), 3)
          )) : (g = 1 - f * (0.253 + f * 0.12), l < g ? u = o.pow(l / g, 1 / f) : u = 1 - o.log(1 - (l - g) / (1 - g))); d < 12; d++) {
            if (u <= 0)
              return 0;
            if (h = r.lowRegGamma(f, u) - l, f > 1 ? g = b * o.exp(-(u - c) + c * (o.log(u) - E)) : g = o.exp(-u + c * o.log(u) - a), m = h / g, u -= g = m / (1 - 0.5 * o.min(1, m * ((f - 1) / u - 1))), u <= 0 && (u = 0.5 * (u + g)), o.abs(g) < i * u)
              break;
          }
          return u;
        }, r.erf = function(l) {
          var f = [
            -1.3026537197817094,
            0.6419697923564902,
            0.019476473204185836,
            -0.00956151478680863,
            -946595344482036e-18,
            366839497852761e-18,
            42523324806907e-18,
            -20278578112534e-18,
            -1624290004647e-18,
            130365583558e-17,
            15626441722e-18,
            -85238095915e-18,
            6529054439e-18,
            5059343495e-18,
            -991364156e-18,
            -227365122e-18,
            96467911e-18,
            2394038e-18,
            -6886027e-18,
            894487e-18,
            313092e-18,
            -112708e-18,
            381e-18,
            7106e-18,
            -1523e-18,
            -94e-18,
            121e-18,
            -28e-18
          ], d = f.length - 1, c = !1, i = 0, a = 0, u, h, g, m;
          for (l < 0 && (l = -l, c = !0), u = 2 / (2 + l), h = 4 * u - 2; d > 0; d--)
            g = i, i = h * i - a + f[d], a = g;
          return m = u * o.exp(-l * l + 0.5 * (f[0] + h * i) - a), c ? m - 1 : 1 - m;
        }, r.erfc = function(l) {
          return 1 - r.erf(l);
        }, r.erfcinv = function(l) {
          var f = 0, d, c, i, a;
          if (l >= 2)
            return -100;
          if (l <= 0)
            return 100;
          for (a = l < 1 ? l : 2 - l, i = o.sqrt(-2 * o.log(a / 2)), d = -0.70711 * ((2.30753 + i * 0.27061) / (1 + i * (0.99229 + i * 0.04481)) - i); f < 2; f++)
            c = r.erfc(d) - a, d += c / (1.1283791670955126 * o.exp(-d * d) - d * c);
          return l < 1 ? d : -d;
        }, r.ibetainv = function(l, f, d) {
          var c = 1e-8, i = f - 1, a = d - 1, u = 0, h, g, m, p, E, b, N, v, C, M, R;
          if (l <= 0)
            return 0;
          if (l >= 1)
            return 1;
          for (f >= 1 && d >= 1 ? (m = l < 0.5 ? l : 1 - l, p = o.sqrt(-2 * o.log(m)), N = (2.30753 + p * 0.27061) / (1 + p * (0.99229 + p * 0.04481)) - p, l < 0.5 && (N = -N), v = (N * N - 3) / 6, C = 2 / (1 / (2 * f - 1) + 1 / (2 * d - 1)), M = N * o.sqrt(v + C) / C - (1 / (2 * d - 1) - 1 / (2 * f - 1)) * (v + 5 / 6 - 2 / (3 * C)), N = f / (f + d * o.exp(2 * M))) : (h = o.log(f / (f + d)), g = o.log(d / (f + d)), p = o.exp(f * h) / f, E = o.exp(d * g) / d, M = p + E, l < p / M ? N = o.pow(f * M * l, 1 / f) : N = 1 - o.pow(d * M * (1 - l), 1 / d)), R = -r.gammaln(f) - r.gammaln(d) + r.gammaln(f + d); u < 10; u++) {
            if (N === 0 || N === 1)
              return N;
            if (b = r.ibeta(N, f, d) - l, p = o.exp(i * o.log(N) + a * o.log(1 - N) + R), E = b / p, N -= p = E / (1 - 0.5 * o.min(1, E * (i / N - a / (1 - N)))), N <= 0 && (N = 0.5 * (N + p)), N >= 1 && (N = 0.5 * (N + p + 1)), o.abs(p) < c * N && u > 0)
              break;
          }
          return N;
        }, r.ibeta = function(l, f, d) {
          var c = l === 0 || l === 1 ? 0 : o.exp(r.gammaln(f + d) - r.gammaln(f) - r.gammaln(d) + f * o.log(l) + d * o.log(1 - l));
          return l < 0 || l > 1 ? !1 : l < (f + 1) / (f + d + 2) ? c * r.betacf(l, f, d) / f : 1 - c * r.betacf(1 - l, d, f) / d;
        }, r.randn = function(l, f) {
          var d, c, i, a, u;
          if (f || (f = l), l)
            return r.create(l, f, function() {
              return r.randn();
            });
          do
            d = r._random_fn(), c = 1.7156 * (r._random_fn() - 0.5), i = d - 0.449871, a = o.abs(c) + 0.386595, u = i * i + a * (0.196 * a - 0.25472 * i);
          while (u > 0.27597 && (u > 0.27846 || c * c > -4 * o.log(d) * d * d));
          return c / d;
        }, r.randg = function(l, f, d) {
          var c = l, i, a, u, h, g, m;
          if (d || (d = f), l || (l = 1), f)
            return m = r.zeros(f, d), m.alter(function() {
              return r.randg(l);
            }), m;
          l < 1 && (l += 1), i = l - 1 / 3, a = 1 / o.sqrt(9 * i);
          do {
            do
              g = r.randn(), h = 1 + a * g;
            while (h <= 0);
            h = h * h * h, u = r._random_fn();
          } while (u > 1 - 0.331 * o.pow(g, 4) && o.log(u) > 0.5 * g * g + i * (1 - h + o.log(h)));
          if (l == c)
            return i * h;
          do
            u = r._random_fn();
          while (u === 0);
          return o.pow(u, 1 / c) * i * h;
        }, (function(s) {
          for (var l = 0; l < s.length; l++) (function(f) {
            r.fn[f] = function() {
              return r(
                r.map(this, function(d) {
                  return r[f](d);
                })
              );
            };
          })(s[l]);
        })("gammaln gammafn factorial factorialln".split(" ")), (function(s) {
          for (var l = 0; l < s.length; l++) (function(f) {
            r.fn[f] = function() {
              return r(r[f].apply(null, arguments));
            };
          })(s[l]);
        })("randn".split(" "));
      })(n, Math), (function(r, o) {
        (function(c) {
          for (var i = 0; i < c.length; i++) (function(a) {
            r[a] = function u(h, g, m) {
              return this instanceof u ? (this._a = h, this._b = g, this._c = m, this) : new u(h, g, m);
            }, r.fn[a] = function(u, h, g) {
              var m = r[a](u, h, g);
              return m.data = this, m;
            }, r[a].prototype.sample = function(u) {
              var h = this._a, g = this._b, m = this._c;
              return u ? r.alter(u, function() {
                return r[a].sample(h, g, m);
              }) : r[a].sample(h, g, m);
            }, (function(u) {
              for (var h = 0; h < u.length; h++) (function(g) {
                r[a].prototype[g] = function(m) {
                  var p = this._a, E = this._b, b = this._c;
                  return !m && m !== 0 && (m = this.data), typeof m != "number" ? r.fn.map.call(m, function(N) {
                    return r[a][g](N, p, E, b);
                  }) : r[a][g](m, p, E, b);
                };
              })(u[h]);
            })("pdf cdf inv".split(" ")), (function(u) {
              for (var h = 0; h < u.length; h++) (function(g) {
                r[a].prototype[g] = function() {
                  return r[a][g](this._a, this._b, this._c);
                };
              })(u[h]);
            })("mean median mode variance".split(" "));
          })(c[i]);
        })("beta centralF cauchy chisquare exponential gamma invgamma kumaraswamy laplace lognormal noncentralt normal pareto studentt weibull uniform binomial negbin hypgeom poisson triangular tukey arcsine".split(" ")), r.extend(r.beta, {
          pdf: function(i, a, u) {
            return i > 1 || i < 0 ? 0 : a == 1 && u == 1 ? 1 : a < 512 && u < 512 ? o.pow(i, a - 1) * o.pow(1 - i, u - 1) / r.betafn(a, u) : o.exp((a - 1) * o.log(i) + (u - 1) * o.log(1 - i) - r.betaln(a, u));
          },
          cdf: function(i, a, u) {
            return i > 1 || i < 0 ? (i > 1) * 1 : r.ibeta(i, a, u);
          },
          inv: function(i, a, u) {
            return r.ibetainv(i, a, u);
          },
          mean: function(i, a) {
            return i / (i + a);
          },
          median: function(i, a) {
            return r.ibetainv(0.5, i, a);
          },
          mode: function(i, a) {
            return (i - 1) / (i + a - 2);
          },
          // return a random sample
          sample: function(i, a) {
            var u = r.randg(i);
            return u / (u + r.randg(a));
          },
          variance: function(i, a) {
            return i * a / (o.pow(i + a, 2) * (i + a + 1));
          }
        }), r.extend(r.centralF, {
          // This implementation of the pdf function avoids float overflow
          // See the way that R calculates this value:
          // https://svn.r-project.org/R/trunk/src/nmath/df.c
          pdf: function(i, a, u) {
            var h, g, m;
            return i < 0 ? 0 : a <= 2 ? i === 0 && a < 2 ? 1 / 0 : i === 0 && a === 2 ? 1 : 1 / r.betafn(a / 2, u / 2) * o.pow(a / u, a / 2) * o.pow(i, a / 2 - 1) * o.pow(1 + a / u * i, -(a + u) / 2) : (h = a * i / (u + i * a), g = u / (u + i * a), m = a * g / 2, m * r.binomial.pdf((a - 2) / 2, (a + u - 2) / 2, h));
          },
          cdf: function(i, a, u) {
            return i < 0 ? 0 : r.ibeta(a * i / (a * i + u), a / 2, u / 2);
          },
          inv: function(i, a, u) {
            return u / (a * (1 / r.ibetainv(i, a / 2, u / 2) - 1));
          },
          mean: function(i, a) {
            return a > 2 ? a / (a - 2) : void 0;
          },
          mode: function(i, a) {
            return i > 2 ? a * (i - 2) / (i * (a + 2)) : void 0;
          },
          // return a random sample
          sample: function(i, a) {
            var u = r.randg(i / 2) * 2, h = r.randg(a / 2) * 2;
            return u / i / (h / a);
          },
          variance: function(i, a) {
            if (!(a <= 4))
              return 2 * a * a * (i + a - 2) / (i * (a - 2) * (a - 2) * (a - 4));
          }
        }), r.extend(r.cauchy, {
          pdf: function(i, a, u) {
            return u < 0 ? 0 : u / (o.pow(i - a, 2) + o.pow(u, 2)) / o.PI;
          },
          cdf: function(i, a, u) {
            return o.atan((i - a) / u) / o.PI + 0.5;
          },
          inv: function(c, i, a) {
            return i + a * o.tan(o.PI * (c - 0.5));
          },
          median: function(i) {
            return i;
          },
          mode: function(i) {
            return i;
          },
          sample: function(i, a) {
            return r.randn() * o.sqrt(1 / (2 * r.randg(0.5))) * a + i;
          }
        }), r.extend(r.chisquare, {
          pdf: function(i, a) {
            return i < 0 ? 0 : i === 0 && a === 2 ? 0.5 : o.exp((a / 2 - 1) * o.log(i) - i / 2 - a / 2 * o.log(2) - r.gammaln(a / 2));
          },
          cdf: function(i, a) {
            return i < 0 ? 0 : r.lowRegGamma(a / 2, i / 2);
          },
          inv: function(c, i) {
            return 2 * r.gammapinv(c, 0.5 * i);
          },
          mean: function(c) {
            return c;
          },
          // TODO: this is an approximation (is there a better way?)
          median: function(i) {
            return i * o.pow(1 - 2 / (9 * i), 3);
          },
          mode: function(i) {
            return i - 2 > 0 ? i - 2 : 0;
          },
          sample: function(i) {
            return r.randg(i / 2) * 2;
          },
          variance: function(i) {
            return 2 * i;
          }
        }), r.extend(r.exponential, {
          pdf: function(i, a) {
            return i < 0 ? 0 : a * o.exp(-a * i);
          },
          cdf: function(i, a) {
            return i < 0 ? 0 : 1 - o.exp(-a * i);
          },
          inv: function(c, i) {
            return -o.log(1 - c) / i;
          },
          mean: function(c) {
            return 1 / c;
          },
          median: function(c) {
            return 1 / c * o.log(2);
          },
          mode: function() {
            return 0;
          },
          sample: function(i) {
            return -1 / i * o.log(r._random_fn());
          },
          variance: function(c) {
            return o.pow(c, -2);
          }
        }), r.extend(r.gamma, {
          pdf: function(i, a, u) {
            return i < 0 ? 0 : i === 0 && a === 1 ? 1 / u : o.exp((a - 1) * o.log(i) - i / u - r.gammaln(a) - a * o.log(u));
          },
          cdf: function(i, a, u) {
            return i < 0 ? 0 : r.lowRegGamma(a, i / u);
          },
          inv: function(c, i, a) {
            return r.gammapinv(c, i) * a;
          },
          mean: function(c, i) {
            return c * i;
          },
          mode: function(i, a) {
            if (i > 1) return (i - 1) * a;
          },
          sample: function(i, a) {
            return r.randg(i) * a;
          },
          variance: function(i, a) {
            return i * a * a;
          }
        }), r.extend(r.invgamma, {
          pdf: function(i, a, u) {
            return i <= 0 ? 0 : o.exp(-(a + 1) * o.log(i) - u / i - r.gammaln(a) + a * o.log(u));
          },
          cdf: function(i, a, u) {
            return i <= 0 ? 0 : 1 - r.lowRegGamma(a, u / i);
          },
          inv: function(c, i, a) {
            return a / r.gammapinv(1 - c, i);
          },
          mean: function(c, i) {
            return c > 1 ? i / (c - 1) : void 0;
          },
          mode: function(i, a) {
            return a / (i + 1);
          },
          sample: function(i, a) {
            return a / r.randg(i);
          },
          variance: function(i, a) {
            if (!(i <= 2))
              return a * a / ((i - 1) * (i - 1) * (i - 2));
          }
        }), r.extend(r.kumaraswamy, {
          pdf: function(i, a, u) {
            return i === 0 && a === 1 ? u : i === 1 && u === 1 ? a : o.exp(o.log(a) + o.log(u) + (a - 1) * o.log(i) + (u - 1) * o.log(1 - o.pow(i, a)));
          },
          cdf: function(i, a, u) {
            return i < 0 ? 0 : i > 1 ? 1 : 1 - o.pow(1 - o.pow(i, a), u);
          },
          inv: function(i, a, u) {
            return o.pow(1 - o.pow(1 - i, 1 / u), 1 / a);
          },
          mean: function(c, i) {
            return i * r.gammafn(1 + 1 / c) * r.gammafn(i) / r.gammafn(1 + 1 / c + i);
          },
          median: function(i, a) {
            return o.pow(1 - o.pow(2, -1 / a), 1 / i);
          },
          mode: function(i, a) {
            if (i >= 1 && a >= 1 && i !== 1 && a !== 1)
              return o.pow((i - 1) / (i * a - 1), 1 / i);
          },
          variance: function() {
            throw new Error("variance not yet implemented");
          }
        }), r.extend(r.lognormal, {
          pdf: function(i, a, u) {
            return i <= 0 ? 0 : o.exp(-o.log(i) - 0.5 * o.log(2 * o.PI) - o.log(u) - o.pow(o.log(i) - a, 2) / (2 * u * u));
          },
          cdf: function(i, a, u) {
            return i < 0 ? 0 : 0.5 + 0.5 * r.erf((o.log(i) - a) / o.sqrt(2 * u * u));
          },
          inv: function(c, i, a) {
            return o.exp(-1.4142135623730951 * a * r.erfcinv(2 * c) + i);
          },
          mean: function(i, a) {
            return o.exp(i + a * a / 2);
          },
          median: function(i) {
            return o.exp(i);
          },
          mode: function(i, a) {
            return o.exp(i - a * a);
          },
          sample: function(i, a) {
            return o.exp(r.randn() * a + i);
          },
          variance: function(i, a) {
            return (o.exp(a * a) - 1) * o.exp(2 * i + a * a);
          }
        }), r.extend(r.noncentralt, {
          pdf: function(i, a, u) {
            var h = 1e-14;
            return o.abs(u) < h ? r.studentt.pdf(i, a) : o.abs(i) < h ? o.exp(r.gammaln((a + 1) / 2) - u * u / 2 - 0.5 * o.log(o.PI * a) - r.gammaln(a / 2)) : a / i * (r.noncentralt.cdf(i * o.sqrt(1 + 2 / a), a + 2, u) - r.noncentralt.cdf(i, a, u));
          },
          cdf: function(i, a, u) {
            var h = 1e-14, g = 200;
            if (o.abs(u) < h)
              return r.studentt.cdf(i, a);
            var m = !1;
            i < 0 && (m = !0, u = -u);
            for (var p = r.normal.cdf(-u, 0, 1), E = h + 1, b = E, N = i * i / (i * i + a), v = 0, C = o.exp(-u * u / 2), M = o.exp(-u * u / 2 - 0.5 * o.log(2) - r.gammaln(3 / 2)) * u; v < g || b > h || E > h; )
              b = E, v > 0 && (C *= u * u / (2 * v), M *= u * u / (2 * (v + 1 / 2))), E = C * r.beta.cdf(N, v + 0.5, a / 2) + M * r.beta.cdf(N, v + 1, a / 2), p += 0.5 * E, v++;
            return m ? 1 - p : p;
          }
        }), r.extend(r.normal, {
          pdf: function(i, a, u) {
            return o.exp(-0.5 * o.log(2 * o.PI) - o.log(u) - o.pow(i - a, 2) / (2 * u * u));
          },
          cdf: function(i, a, u) {
            return 0.5 * (1 + r.erf((i - a) / o.sqrt(2 * u * u)));
          },
          inv: function(c, i, a) {
            return -1.4142135623730951 * a * r.erfcinv(2 * c) + i;
          },
          mean: function(c) {
            return c;
          },
          median: function(i) {
            return i;
          },
          mode: function(c) {
            return c;
          },
          sample: function(i, a) {
            return r.randn() * a + i;
          },
          variance: function(c, i) {
            return i * i;
          }
        }), r.extend(r.pareto, {
          pdf: function(i, a, u) {
            return i < a ? 0 : u * o.pow(a, u) / o.pow(i, u + 1);
          },
          cdf: function(i, a, u) {
            return i < a ? 0 : 1 - o.pow(a / i, u);
          },
          inv: function(i, a, u) {
            return a / o.pow(1 - i, 1 / u);
          },
          mean: function(i, a) {
            if (!(a <= 1))
              return a * o.pow(i, a) / (a - 1);
          },
          median: function(i, a) {
            return i * (a * o.SQRT2);
          },
          mode: function(i) {
            return i;
          },
          variance: function(c, i) {
            if (!(i <= 2))
              return c * c * i / (o.pow(i - 1, 2) * (i - 2));
          }
        }), r.extend(r.studentt, {
          pdf: function(i, a) {
            return a = a > 1e100 ? 1e100 : a, 1 / (o.sqrt(a) * r.betafn(0.5, a / 2)) * o.pow(1 + i * i / a, -((a + 1) / 2));
          },
          cdf: function(i, a) {
            var u = a / 2;
            return r.ibeta((i + o.sqrt(i * i + a)) / (2 * o.sqrt(i * i + a)), u, u);
          },
          inv: function(c, i) {
            var a = r.ibetainv(2 * o.min(c, 1 - c), 0.5 * i, 0.5);
            return a = o.sqrt(i * (1 - a) / a), c > 0.5 ? a : -a;
          },
          mean: function(i) {
            return i > 1 ? 0 : void 0;
          },
          median: function() {
            return 0;
          },
          mode: function() {
            return 0;
          },
          sample: function(i) {
            return r.randn() * o.sqrt(i / (2 * r.randg(i / 2)));
          },
          variance: function(i) {
            return i > 2 ? i / (i - 2) : i > 1 ? 1 / 0 : void 0;
          }
        }), r.extend(r.weibull, {
          pdf: function(i, a, u) {
            return i < 0 || a < 0 || u < 0 ? 0 : u / a * o.pow(i / a, u - 1) * o.exp(-o.pow(i / a, u));
          },
          cdf: function(i, a, u) {
            return i < 0 ? 0 : 1 - o.exp(-o.pow(i / a, u));
          },
          inv: function(c, i, a) {
            return i * o.pow(-o.log(1 - c), 1 / a);
          },
          mean: function(c, i) {
            return c * r.gammafn(1 + 1 / i);
          },
          median: function(i, a) {
            return i * o.pow(o.log(2), 1 / a);
          },
          mode: function(i, a) {
            return a <= 1 ? 0 : i * o.pow((a - 1) / a, 1 / a);
          },
          sample: function(i, a) {
            return i * o.pow(-o.log(r._random_fn()), 1 / a);
          },
          variance: function(i, a) {
            return i * i * r.gammafn(1 + 2 / a) - o.pow(r.weibull.mean(i, a), 2);
          }
        }), r.extend(r.uniform, {
          pdf: function(i, a, u) {
            return i < a || i > u ? 0 : 1 / (u - a);
          },
          cdf: function(i, a, u) {
            return i < a ? 0 : i < u ? (i - a) / (u - a) : 1;
          },
          inv: function(c, i, a) {
            return i + c * (a - i);
          },
          mean: function(i, a) {
            return 0.5 * (i + a);
          },
          median: function(i, a) {
            return r.mean(i, a);
          },
          mode: function() {
            throw new Error("mode is not yet implemented");
          },
          sample: function(i, a) {
            return i / 2 + a / 2 + (a / 2 - i / 2) * (2 * r._random_fn() - 1);
          },
          variance: function(i, a) {
            return o.pow(a - i, 2) / 12;
          }
        });
        function s(c, i, a, u) {
          for (var h = 0, g = 1, m = 1, p = 1, E = 0, b = 0, N; o.abs((m - b) / m) > u; )
            b = m, N = -(i + E) * (i + a + E) * c / (i + 2 * E) / (i + 2 * E + 1), h = m + N * h, g = p + N * g, E = E + 1, N = E * (a - E) * c / (i + 2 * E - 1) / (i + 2 * E), m = h + N * m, p = g + N * p, h = h / p, g = g / p, m = m / p, p = 1;
          return m / i;
        }
        r.extend(r.binomial, {
          pdf: function(i, a, u) {
            return u === 0 || u === 1 ? a * u === i ? 1 : 0 : r.combination(a, i) * o.pow(u, i) * o.pow(1 - u, a - i);
          },
          cdf: function(i, a, u) {
            var h, g = 1e-10;
            if (i < 0)
              return 0;
            if (i >= a)
              return 1;
            if (u < 0 || u > 1 || a <= 0)
              return NaN;
            i = o.floor(i);
            var m = u, p = i + 1, E = a - i, b = p + E, N = o.exp(r.gammaln(b) - r.gammaln(E) - r.gammaln(p) + p * o.log(m) + E * o.log(1 - m));
            return m < (p + 1) / (b + 2) ? h = N * s(m, p, E, g) : h = 1 - N * s(1 - m, E, p, g), o.round((1 - h) * (1 / g)) / (1 / g);
          }
        }), r.extend(r.negbin, {
          pdf: function(i, a, u) {
            return i !== i >>> 0 ? !1 : i < 0 ? 0 : r.combination(i + a - 1, a - 1) * o.pow(1 - u, i) * o.pow(u, a);
          },
          cdf: function(i, a, u) {
            var h = 0, g = 0;
            if (i < 0) return 0;
            for (; g <= i; g++)
              h += r.negbin.pdf(g, a, u);
            return h;
          }
        }), r.extend(r.hypgeom, {
          pdf: function(i, a, u, h) {
            if (i !== i | 0)
              return !1;
            if (i < 0 || i < u - (a - h))
              return 0;
            if (i > h || i > u)
              return 0;
            if (u * 2 > a)
              return h * 2 > a ? r.hypgeom.pdf(a - u - h + i, a, a - u, a - h) : r.hypgeom.pdf(h - i, a, a - u, h);
            if (h * 2 > a)
              return r.hypgeom.pdf(u - i, a, u, a - h);
            if (u < h)
              return r.hypgeom.pdf(i, a, h, u);
            for (var g = 1, m = 0, p = 0; p < i; p++) {
              for (; g > 1 && m < h; )
                g *= 1 - u / (a - m), m++;
              g *= (h - p) * (u - p) / ((p + 1) * (a - u - h + p + 1));
            }
            for (; m < h; m++)
              g *= 1 - u / (a - m);
            return o.min(1, o.max(0, g));
          },
          cdf: function(i, a, u, h) {
            if (i < 0 || i < u - (a - h))
              return 0;
            if (i >= h || i >= u)
              return 1;
            if (u * 2 > a)
              return h * 2 > a ? r.hypgeom.cdf(a - u - h + i, a, a - u, a - h) : 1 - r.hypgeom.cdf(h - i - 1, a, a - u, h);
            if (h * 2 > a)
              return 1 - r.hypgeom.cdf(u - i - 1, a, u, a - h);
            if (u < h)
              return r.hypgeom.cdf(i, a, h, u);
            for (var g = 1, m = 1, p = 0, E = 0; E < i; E++) {
              for (; g > 1 && p < h; ) {
                var b = 1 - u / (a - p);
                m *= b, g *= b, p++;
              }
              m *= (h - E) * (u - E) / ((E + 1) * (a - u - h + E + 1)), g += m;
            }
            for (; p < h; p++)
              g *= 1 - u / (a - p);
            return o.min(1, o.max(0, g));
          }
        }), r.extend(r.poisson, {
          pdf: function(i, a) {
            return a < 0 || i % 1 !== 0 || i < 0 ? 0 : o.pow(a, i) * o.exp(-a) / r.factorial(i);
          },
          cdf: function(i, a) {
            var u = [], h = 0;
            if (i < 0) return 0;
            for (; h <= i; h++)
              u.push(r.poisson.pdf(h, a));
            return r.sum(u);
          },
          mean: function(c) {
            return c;
          },
          variance: function(c) {
            return c;
          },
          sampleSmall: function(i) {
            var a = 1, u = 0, h = o.exp(-i);
            do
              u++, a *= r._random_fn();
            while (a > h);
            return u - 1;
          },
          sampleLarge: function(i) {
            var a = i, u, h, g, m, p, E, b, N, v, C;
            for (m = o.sqrt(a), p = o.log(a), b = 0.931 + 2.53 * m, E = -0.059 + 0.02483 * b, N = 1.1239 + 1.1328 / (b - 3.4), v = 0.9277 - 3.6224 / (b - 2); ; ) {
              if (h = o.random() - 0.5, g = o.random(), C = 0.5 - o.abs(h), u = o.floor((2 * E / C + b) * h + a + 0.43), C >= 0.07 && g <= v)
                return u;
              if (!(u < 0 || C < 0.013 && g > C) && o.log(g) + o.log(N) - o.log(E / (C * C) + b) <= -a + u * p - r.loggam(u + 1))
                return u;
            }
          },
          sample: function(i) {
            return i < 10 ? this.sampleSmall(i) : this.sampleLarge(i);
          }
        }), r.extend(r.triangular, {
          pdf: function(i, a, u, h) {
            return u <= a || h < a || h > u ? NaN : i < a || i > u ? 0 : i < h ? 2 * (i - a) / ((u - a) * (h - a)) : i === h ? 2 / (u - a) : 2 * (u - i) / ((u - a) * (u - h));
          },
          cdf: function(i, a, u, h) {
            return u <= a || h < a || h > u ? NaN : i <= a ? 0 : i >= u ? 1 : i <= h ? o.pow(i - a, 2) / ((u - a) * (h - a)) : 1 - o.pow(u - i, 2) / ((u - a) * (u - h));
          },
          inv: function(i, a, u, h) {
            return u <= a || h < a || h > u ? NaN : i <= (h - a) / (u - a) ? a + (u - a) * o.sqrt(i * ((h - a) / (u - a))) : a + (u - a) * (1 - o.sqrt((1 - i) * (1 - (h - a) / (u - a))));
          },
          mean: function(i, a, u) {
            return (i + a + u) / 3;
          },
          median: function(i, a, u) {
            if (u <= (i + a) / 2)
              return a - o.sqrt((a - i) * (a - u)) / o.sqrt(2);
            if (u > (i + a) / 2)
              return i + o.sqrt((a - i) * (u - i)) / o.sqrt(2);
          },
          mode: function(i, a, u) {
            return u;
          },
          sample: function(i, a, u) {
            var h = r._random_fn();
            return h < (u - i) / (a - i) ? i + o.sqrt(h * (a - i) * (u - i)) : a - o.sqrt((1 - h) * (a - i) * (a - u));
          },
          variance: function(i, a, u) {
            return (i * i + a * a + u * u - i * a - i * u - a * u) / 18;
          }
        }), r.extend(r.arcsine, {
          pdf: function(i, a, u) {
            return u <= a ? NaN : i <= a || i >= u ? 0 : 2 / o.PI * o.pow(o.pow(u - a, 2) - o.pow(2 * i - a - u, 2), -0.5);
          },
          cdf: function(i, a, u) {
            return i < a ? 0 : i < u ? 2 / o.PI * o.asin(o.sqrt((i - a) / (u - a))) : 1;
          },
          inv: function(c, i, a) {
            return i + (0.5 - 0.5 * o.cos(o.PI * c)) * (a - i);
          },
          mean: function(i, a) {
            return a <= i ? NaN : (i + a) / 2;
          },
          median: function(i, a) {
            return a <= i ? NaN : (i + a) / 2;
          },
          mode: function() {
            throw new Error("mode is not yet implemented");
          },
          sample: function(i, a) {
            return (i + a) / 2 + (a - i) / 2 * o.sin(2 * o.PI * r.uniform.sample(0, 1));
          },
          variance: function(i, a) {
            return a <= i ? NaN : o.pow(a - i, 2) / 8;
          }
        });
        function l(c) {
          return c / o.abs(c);
        }
        r.extend(r.laplace, {
          pdf: function(i, a, u) {
            return u <= 0 ? 0 : o.exp(-o.abs(i - a) / u) / (2 * u);
          },
          cdf: function(i, a, u) {
            return u <= 0 ? 0 : i < a ? 0.5 * o.exp((i - a) / u) : 1 - 0.5 * o.exp(-(i - a) / u);
          },
          mean: function(c) {
            return c;
          },
          median: function(c) {
            return c;
          },
          mode: function(c) {
            return c;
          },
          variance: function(c, i) {
            return 2 * i * i;
          },
          sample: function(i, a) {
            var u = r._random_fn() - 0.5;
            return i - a * l(u) * o.log(1 - 2 * o.abs(u));
          }
        });
        function f(c, i, a) {
          var u = 12, h = 6, g = -30, m = -50, p = 60, E = 8, b = 3, N = 2, v = 3, C = [
            0.9815606342467192,
            0.9041172563704749,
            0.7699026741943047,
            0.5873179542866175,
            0.3678314989981802,
            0.1252334085114689
          ], M = [
            0.04717533638651183,
            0.10693932599531843,
            0.16007832854334622,
            0.20316742672306592,
            0.2334925365383548,
            0.24914704581340277
          ], R = c * 0.5;
          if (R >= E)
            return 1;
          var A = 2 * r.normal.cdf(R, 0, 1, 1, 0) - 1;
          A >= o.exp(m / a) ? A = o.pow(A, a) : A = 0;
          var x;
          c > b ? x = N : x = v;
          for (var L = R, D = (E - R) / x, W = L + D, J = 0, ne = a - 1, K = 1; K <= x; K++) {
            for (var q = 0, H = 0.5 * (W + L), Y = 0.5 * (W - L), te = 1; te <= u; te++) {
              var X, le;
              h < te ? (X = u - te + 1, le = C[X - 1]) : (X = te, le = -C[X - 1]);
              var $ = Y * le, j = H + $, pe = j * j;
              if (pe > p)
                break;
              var St = 2 * r.normal.cdf(j, 0, 1, 1, 0), De = 2 * r.normal.cdf(j, c, 1, 1, 0), vt = St * 0.5 - De * 0.5;
              vt >= o.exp(g / ne) && (vt = M[X - 1] * o.exp(-(0.5 * pe)) * o.pow(vt, ne), q += vt);
            }
            q *= 2 * Y * a / o.sqrt(2 * o.PI), J += q, L = W, W += D;
          }
          return A += J, A <= o.exp(g / i) ? 0 : (A = o.pow(A, i), A >= 1 ? 1 : A);
        }
        function d(c, i, a) {
          var u = 0.322232421088, h = 0.099348462606, g = -1, m = 0.588581570495, p = -0.342242088547, E = 0.531103462366, b = -0.204231210125, N = 0.10353775285, v = -453642210148e-16, C = 0.0038560700634, M = 0.8832, R = 0.2368, A = 1.208, x = 1.4142, L = 120, D = 0.5 - 0.5 * c, W = o.sqrt(o.log(1 / (D * D))), J = W + ((((W * v + b) * W + p) * W + g) * W + u) / ((((W * C + N) * W + E) * W + m) * W + h);
          a < L && (J += (J * J * J + J) / a / 4);
          var ne = M - R * J;
          return a < L && (ne += -1.214 / a + A * J / a), J * (ne * o.log(i - 1) + x);
        }
        r.extend(r.tukey, {
          cdf: function(i, a, u) {
            var h = 1, g = a, m = 16, p = 8, E = -30, b = 1e-14, N = 100, v = 800, C = 5e3, M = 25e3, R = 1, A = 0.5, x = 0.25, L = 0.125, D = [
              0.9894009349916499,
              0.9445750230732326,
              0.8656312023878318,
              0.755404408355003,
              0.6178762444026438,
              0.45801677765722737,
              0.2816035507792589,
              0.09501250983763744
            ], W = [
              0.027152459411754096,
              0.062253523938647894,
              0.09515851168249279,
              0.12462897125553388,
              0.14959598881657674,
              0.16915651939500254,
              0.18260341504492358,
              0.1894506104550685
            ];
            if (i <= 0)
              return 0;
            if (u < 2 || h < 1 || g < 2) return NaN;
            if (!Number.isFinite(i))
              return 1;
            if (u > M)
              return f(i, h, g);
            var J = u * 0.5, ne = J * o.log(u) - u * o.log(2) - r.gammaln(J), K = J - 1, q = u * 0.25, H;
            u <= N ? H = R : u <= v ? H = A : u <= C ? H = x : H = L, ne += o.log(H);
            for (var Y = 0, te = 1; te <= 50; te++) {
              for (var X = 0, le = (2 * te - 1) * H, $ = 1; $ <= m; $++) {
                var j, pe;
                p < $ ? (j = $ - p - 1, pe = ne + K * o.log(le + D[j] * H) - (D[j] * H + le) * q) : (j = $ - 1, pe = ne + K * o.log(le - D[j] * H) + (D[j] * H - le) * q);
                var St;
                if (pe >= E) {
                  p < $ ? St = i * o.sqrt((D[j] * H + le) * 0.5) : St = i * o.sqrt((-(D[j] * H) + le) * 0.5);
                  var De = f(St, h, g), vt = De * W[j] * o.exp(pe);
                  X += vt;
                }
              }
              if (te * H >= 1 && X <= b)
                break;
              Y += X;
            }
            if (X > b)
              throw new Error("tukey.cdf failed to converge");
            return Y > 1 && (Y = 1), Y;
          },
          inv: function(c, i, a) {
            var u = 1, h = i, g = 1e-4, m = 50;
            if (a < 2 || u < 1 || h < 2) return NaN;
            if (c < 0 || c > 1) return NaN;
            if (c === 0) return 0;
            if (c === 1) return 1 / 0;
            var p = d(c, h, a), E = r.tukey.cdf(p, i, a) - c, b;
            E > 0 ? b = o.max(0, p - 1) : b = p + 1;
            for (var N = r.tukey.cdf(b, i, a) - c, v, C = 1; C < m; C++) {
              v = b - N * (b - p) / (N - E), E = N, p = b, v < 0 && (v = 0, N = -c), N = r.tukey.cdf(v, i, a) - c, b = v;
              var M = o.abs(b - p);
              if (M < g)
                return v;
            }
            throw new Error("tukey.inv failed to converge");
          }
        });
      })(n, Math), (function(r, o) {
        var s = Array.prototype.push, l = r.utils.isArray;
        function f(d) {
          return l(d) || d instanceof r;
        }
        r.extend({
          // add a vector/matrix to a vector/matrix or scalar
          add: function(c, i) {
            return f(i) ? (f(i[0]) || (i = [i]), r.map(c, function(a, u, h) {
              return a + i[u][h];
            })) : r.map(c, function(a) {
              return a + i;
            });
          },
          // subtract a vector or scalar from the vector
          subtract: function(c, i) {
            return f(i) ? (f(i[0]) || (i = [i]), r.map(c, function(a, u, h) {
              return a - i[u][h] || 0;
            })) : r.map(c, function(a) {
              return a - i;
            });
          },
          // matrix division
          divide: function(c, i) {
            return f(i) ? (f(i[0]) || (i = [i]), r.multiply(c, r.inv(i))) : r.map(c, function(a) {
              return a / i;
            });
          },
          // matrix multiplication
          multiply: function(c, i) {
            var a, u, h, g, m, p, E, b;
            if (c.length === void 0 && i.length === void 0)
              return c * i;
            if (m = c.length, p = c[0].length, E = r.zeros(m, h = f(i) ? i[0].length : p), b = 0, f(i)) {
              for (; b < h; b++)
                for (a = 0; a < m; a++) {
                  for (g = 0, u = 0; u < p; u++)
                    g += c[a][u] * i[u][b];
                  E[a][b] = g;
                }
              return m === 1 && b === 1 ? E[0][0] : E;
            }
            return r.map(c, function(N) {
              return N * i;
            });
          },
          // outer([1,2,3],[4,5,6])
          // ===
          // [[1],[2],[3]] times [[4,5,6]]
          // ->
          // [[4,5,6],[8,10,12],[12,15,18]]
          outer: function(c, i) {
            return r.multiply(c.map(function(a) {
              return [a];
            }), [i]);
          },
          // Returns the dot product of two matricies
          dot: function(c, i) {
            f(c[0]) || (c = [c]), f(i[0]) || (i = [i]);
            for (var a = c[0].length === 1 && c.length !== 1 ? r.transpose(c) : c, u = i[0].length === 1 && i.length !== 1 ? r.transpose(i) : i, h = [], g = 0, m = a.length, p = a[0].length, E, b; g < m; g++) {
              for (h[g] = [], E = 0, b = 0; b < p; b++)
                E += a[g][b] * u[g][b];
              h[g] = E;
            }
            return h.length === 1 ? h[0] : h;
          },
          // raise every element by a scalar
          pow: function(c, i) {
            return r.map(c, function(a) {
              return o.pow(a, i);
            });
          },
          // exponentiate every element
          exp: function(c) {
            return r.map(c, function(i) {
              return o.exp(i);
            });
          },
          // generate the natural log of every element
          log: function(c) {
            return r.map(c, function(i) {
              return o.log(i);
            });
          },
          // generate the absolute values of the vector
          abs: function(c) {
            return r.map(c, function(i) {
              return o.abs(i);
            });
          },
          // computes the p-norm of the vector
          // In the case that a matrix is passed, uses the first row as the vector
          norm: function(c, i) {
            var a = 0, u = 0;
            for (isNaN(i) && (i = 2), f(c[0]) && (c = c[0]); u < c.length; u++)
              a += o.pow(o.abs(c[u]), i);
            return o.pow(a, 1 / i);
          },
          // computes the angle between two vectors in rads
          // In case a matrix is passed, this uses the first row as the vector
          angle: function(c, i) {
            return o.acos(r.dot(c, i) / (r.norm(c) * r.norm(i)));
          },
          // augment one matrix by another
          // Note: this function returns a matrix, not a jStat object
          aug: function(c, i) {
            var a = [], u;
            for (u = 0; u < c.length; u++)
              a.push(c[u].slice());
            for (u = 0; u < a.length; u++)
              s.apply(a[u], i[u]);
            return a;
          },
          // The inv() function calculates the inverse of a matrix
          // Create the inverse by augmenting the matrix by the identity matrix of the
          // appropriate size, and then use G-J elimination on the augmented matrix.
          inv: function(c) {
            for (var i = c.length, a = c[0].length, u = r.identity(i, a), h = r.gauss_jordan(c, u), g = [], m = 0, p; m < i; m++)
              for (g[m] = [], p = a; p < h[0].length; p++)
                g[m][p - a] = h[m][p];
            return g;
          },
          // calculate the determinant of a matrix
          det: function d(c) {
            if (c.length === 2)
              return c[0][0] * c[1][1] - c[0][1] * c[1][0];
            for (var i = 0, a = 0; a < c.length; a++) {
              for (var u = [], h = 1; h < c.length; h++) {
                u[h - 1] = [];
                for (var g = 0; g < c.length; g++)
                  g < a ? u[h - 1][g] = c[h][g] : g > a && (u[h - 1][g - 1] = c[h][g]);
              }
              var m = a % 2 ? -1 : 1;
              i += d(u) * c[0][a] * m;
            }
            return i;
          },
          gauss_elimination: function(c, i) {
            var a = 0, u = 0, h = c.length, g = c[0].length, m = 1, p = 0, E = [], b, N, v, C;
            for (c = r.aug(c, i), b = c[0].length, a = 0; a < h; a++) {
              for (N = c[a][a], u = a, C = a + 1; C < g; C++)
                N < o.abs(c[C][a]) && (N = c[C][a], u = C);
              if (u != a)
                for (C = 0; C < b; C++)
                  v = c[a][C], c[a][C] = c[u][C], c[u][C] = v;
              for (u = a + 1; u < h; u++)
                for (m = c[u][a] / c[a][a], C = a; C < b; C++)
                  c[u][C] = c[u][C] - m * c[a][C];
            }
            for (a = h - 1; a >= 0; a--) {
              for (p = 0, u = a + 1; u <= h - 1; u++)
                p = p + E[u] * c[a][u];
              E[a] = (c[a][b - 1] - p) / c[a][a];
            }
            return E;
          },
          gauss_jordan: function(c, i) {
            var a = r.aug(c, i), u = a.length, h = a[0].length, g = 0, m, p, E;
            for (p = 0; p < u; p++) {
              var b = p;
              for (E = p + 1; E < u; E++)
                o.abs(a[E][p]) > o.abs(a[b][p]) && (b = E);
              var N = a[p];
              for (a[p] = a[b], a[b] = N, E = p + 1; E < u; E++)
                for (g = a[E][p] / a[p][p], m = p; m < h; m++)
                  a[E][m] -= a[p][m] * g;
            }
            for (p = u - 1; p >= 0; p--) {
              for (g = a[p][p], E = 0; E < p; E++)
                for (m = h - 1; m > p - 1; m--)
                  a[E][m] -= a[p][m] * a[E][p] / g;
              for (a[p][p] /= g, m = u; m < h; m++)
                a[p][m] /= g;
            }
            return a;
          },
          // solve equation
          // Ax=b
          // A is upper triangular matrix
          // A=[[1,2,3],[0,4,5],[0,6,7]]
          // b=[1,2,3]
          // triaUpSolve(A,b) // -> [2.666,0.1666,1.666]
          // if you use matrix style
          // A=[[1,2,3],[0,4,5],[0,6,7]]
          // b=[[1],[2],[3]]
          // will return [[2.666],[0.1666],[1.666]]
          triaUpSolve: function(c, i) {
            var a = c[0].length, u = r.zeros(1, a)[0], h, g = !1;
            return i[0].length != null && (i = i.map(function(m) {
              return m[0];
            }), g = !0), r.arange(a - 1, -1, -1).forEach(function(m) {
              h = r.arange(m + 1, a).map(function(p) {
                return u[p] * c[m][p];
              }), u[m] = (i[m] - r.sum(h)) / c[m][m];
            }), g ? u.map(function(m) {
              return [m];
            }) : u;
          },
          triaLowSolve: function(c, i) {
            var a = c[0].length, u = r.zeros(1, a)[0], h, g = !1;
            return i[0].length != null && (i = i.map(function(m) {
              return m[0];
            }), g = !0), r.arange(a).forEach(function(m) {
              h = r.arange(m).map(function(p) {
                return c[m][p] * u[p];
              }), u[m] = (i[m] - r.sum(h)) / c[m][m];
            }), g ? u.map(function(m) {
              return [m];
            }) : u;
          },
          // A -> [L,U]
          // A=LU
          // L is lower triangular matrix
          // U is upper triangular matrix
          lu: function(c) {
            var i = c.length, a = r.identity(i), u = r.zeros(c.length, c[0].length), h;
            return r.arange(i).forEach(function(g) {
              u[0][g] = c[0][g];
            }), r.arange(1, i).forEach(function(g) {
              r.arange(g).forEach(function(m) {
                h = r.arange(m).map(function(p) {
                  return a[g][p] * u[p][m];
                }), a[g][m] = (c[g][m] - r.sum(h)) / u[m][m];
              }), r.arange(g, i).forEach(function(m) {
                h = r.arange(g).map(function(p) {
                  return a[g][p] * u[p][m];
                }), u[g][m] = c[h.length][m] - r.sum(h);
              });
            }), [a, u];
          },
          // A -> T
          // A=TT'
          // T is lower triangular matrix
          cholesky: function(c) {
            var i = c.length, a = r.zeros(c.length, c[0].length), u;
            return r.arange(i).forEach(function(h) {
              u = r.arange(h).map(function(g) {
                return o.pow(a[h][g], 2);
              }), a[h][h] = o.sqrt(c[h][h] - r.sum(u)), r.arange(h + 1, i).forEach(function(g) {
                u = r.arange(h).map(function(m) {
                  return a[h][m] * a[g][m];
                }), a[g][h] = (c[h][g] - r.sum(u)) / a[h][h];
              });
            }), a;
          },
          gauss_jacobi: function(c, i, a, u) {
            for (var h = 0, g = 0, m = c.length, p = [], E = [], b = [], N, v, C, M; h < m; h++)
              for (p[h] = [], E[h] = [], b[h] = [], g = 0; g < m; g++)
                h > g ? (p[h][g] = c[h][g], E[h][g] = b[h][g] = 0) : h < g ? (E[h][g] = c[h][g], p[h][g] = b[h][g] = 0) : (b[h][g] = c[h][g], p[h][g] = E[h][g] = 0);
            for (C = r.multiply(r.multiply(r.inv(b), r.add(p, E)), -1), v = r.multiply(r.inv(b), i), N = a, M = r.add(r.multiply(C, a), v), h = 2; o.abs(r.norm(r.subtract(M, N))) > u; )
              N = M, M = r.add(r.multiply(C, N), v), h++;
            return M;
          },
          gauss_seidel: function(c, i, a, u) {
            for (var h = 0, g = c.length, m = [], p = [], E = [], b, N, v, C, M; h < g; h++)
              for (m[h] = [], p[h] = [], E[h] = [], b = 0; b < g; b++)
                h > b ? (m[h][b] = c[h][b], p[h][b] = E[h][b] = 0) : h < b ? (p[h][b] = c[h][b], m[h][b] = E[h][b] = 0) : (E[h][b] = c[h][b], m[h][b] = p[h][b] = 0);
            for (C = r.multiply(r.multiply(r.inv(r.add(E, m)), p), -1), v = r.multiply(r.inv(r.add(E, m)), i), N = a, M = r.add(r.multiply(C, a), v), h = 2; o.abs(r.norm(r.subtract(M, N))) > u; )
              N = M, M = r.add(r.multiply(C, N), v), h = h + 1;
            return M;
          },
          SOR: function(c, i, a, u, h) {
            for (var g = 0, m = c.length, p = [], E = [], b = [], N, v, C, M, R; g < m; g++)
              for (p[g] = [], E[g] = [], b[g] = [], N = 0; N < m; N++)
                g > N ? (p[g][N] = c[g][N], E[g][N] = b[g][N] = 0) : g < N ? (E[g][N] = c[g][N], p[g][N] = b[g][N] = 0) : (b[g][N] = c[g][N], p[g][N] = E[g][N] = 0);
            for (M = r.multiply(
              r.inv(r.add(b, r.multiply(p, h))),
              r.subtract(
                r.multiply(b, 1 - h),
                r.multiply(E, h)
              )
            ), C = r.multiply(r.multiply(r.inv(r.add(
              b,
              r.multiply(p, h)
            )), i), h), v = a, R = r.add(r.multiply(M, a), C), g = 2; o.abs(r.norm(r.subtract(R, v))) > u; )
              v = R, R = r.add(r.multiply(M, v), C), g++;
            return R;
          },
          householder: function(c) {
            for (var i = c.length, a = c[0].length, u = 0, h = [], g = [], m, p, E, b, N; u < i - 1; u++) {
              for (m = 0, b = u + 1; b < a; b++)
                m += c[b][u] * c[b][u];
              for (N = c[u + 1][u] > 0 ? -1 : 1, m = N * o.sqrt(m), p = o.sqrt((m * m - c[u + 1][u] * m) / 2), h = r.zeros(i, 1), h[u + 1][0] = (c[u + 1][u] - m) / (2 * p), E = u + 2; E < i; E++) h[E][0] = c[E][u] / (2 * p);
              g = r.subtract(
                r.identity(i, a),
                r.multiply(r.multiply(h, r.transpose(h)), 2)
              ), c = r.multiply(g, r.multiply(c, g));
            }
            return c;
          },
          // A -> [Q,R]
          // Q is orthogonal matrix
          // R is upper triangular
          QR: (function() {
            var d = r.sum, c = r.arange;
            function i(a) {
              var u = a.length, h = a[0].length, g = r.zeros(h, h);
              a = r.copy(a);
              var m, p, E;
              for (p = 0; p < h; p++) {
                for (g[p][p] = o.sqrt(d(c(u).map(function(b) {
                  return a[b][p] * a[b][p];
                }))), m = 0; m < u; m++)
                  a[m][p] = a[m][p] / g[p][p];
                for (E = p + 1; E < h; E++)
                  for (g[p][E] = d(c(u).map(function(b) {
                    return a[b][p] * a[b][E];
                  })), m = 0; m < u; m++)
                    a[m][E] = a[m][E] - a[m][p] * g[p][E];
              }
              return [a, g];
            }
            return i;
          })(),
          lstsq: /* @__PURE__ */ (function() {
            function d(i) {
              i = r.copy(i);
              var a = i.length, u = r.identity(a);
              return r.arange(a - 1, -1, -1).forEach(function(h) {
                r.sliceAssign(
                  u,
                  { row: h },
                  r.divide(r.slice(u, { row: h }), i[h][h])
                ), r.sliceAssign(
                  i,
                  { row: h },
                  r.divide(r.slice(i, { row: h }), i[h][h])
                ), r.arange(h).forEach(function(g) {
                  var m = r.multiply(i[g][h], -1), p = r.slice(i, { row: g }), E = r.multiply(r.slice(i, { row: h }), m);
                  r.sliceAssign(i, { row: g }, r.add(p, E));
                  var b = r.slice(u, { row: g }), N = r.multiply(r.slice(u, { row: h }), m);
                  r.sliceAssign(u, { row: g }, r.add(b, N));
                });
              }), u;
            }
            function c(i, a) {
              var u = !1;
              a[0].length === void 0 && (a = a.map(function(M) {
                return [M];
              }), u = !0);
              var h = r.QR(i), g = h[0], m = h[1], p = i[0].length, E = r.slice(g, { col: { end: p } }), b = r.slice(m, { row: { end: p } }), N = d(b), v = r.transpose(E);
              v[0].length === void 0 && (v = [v]);
              var C = r.multiply(r.multiply(N, v), a);
              return C.length === void 0 && (C = [[C]]), u ? C.map(function(M) {
                return M[0];
              }) : C;
            }
            return c;
          })(),
          jacobi: function(c) {
            for (var i = 1, a = c.length, u = r.identity(a, a), h = [], g, m, p, E, b, N, v, C; i === 1; ) {
              for (N = c[0][1], E = 0, b = 1, m = 0; m < a; m++)
                for (p = 0; p < a; p++)
                  m != p && N < o.abs(c[m][p]) && (N = o.abs(c[m][p]), E = m, b = p);
              for (c[E][E] === c[b][b] ? v = c[E][b] > 0 ? o.PI / 4 : -o.PI / 4 : v = o.atan(2 * c[E][b] / (c[E][E] - c[b][b])) / 2, C = r.identity(a, a), C[E][E] = o.cos(v), C[E][b] = -o.sin(v), C[b][E] = o.sin(v), C[b][b] = o.cos(v), u = r.multiply(u, C), g = r.multiply(r.multiply(r.inv(C), c), C), c = g, i = 0, m = 1; m < a; m++)
                for (p = 1; p < a; p++)
                  m != p && o.abs(c[m][p]) > 1e-3 && (i = 1);
            }
            for (m = 0; m < a; m++) h.push(c[m][m]);
            return [u, h];
          },
          rungekutta: function(c, i, a, u, h, g) {
            var m, p, E, b, N;
            if (g === 2)
              for (; u <= a; )
                m = i * c(u, h), p = i * c(u + i, h + m), E = h + (m + p) / 2, h = E, u = u + i;
            if (g === 4)
              for (; u <= a; )
                m = i * c(u, h), p = i * c(u + i / 2, h + m / 2), b = i * c(u + i / 2, h + p / 2), N = i * c(u + i, h + b), E = h + (m + 2 * p + 2 * b + N) / 6, h = E, u = u + i;
            return h;
          },
          romberg: function(c, i, a, u) {
            for (var h = 0, g = (a - i) / 2, m = [], p = [], E = [], b, N, v, C, M; h < u / 2; ) {
              for (M = c(i), v = i, C = 0; v <= a; v = v + g, C++) m[C] = v;
              for (b = m.length, v = 1; v < b - 1; v++)
                M += (v % 2 !== 0 ? 4 : 2) * c(m[v]);
              M = g / 3 * (M + c(a)), E[h] = M, g /= 2, h++;
            }
            for (N = E.length, b = 1; N !== 1; ) {
              for (v = 0; v < N - 1; v++)
                p[v] = (o.pow(4, b) * E[v + 1] - E[v]) / (o.pow(4, b) - 1);
              N = p.length, E = p, p = [], b++;
            }
            return E;
          },
          richardson: function(c, i, a, u) {
            function h(R, A) {
              for (var x = 0, L = R.length, D; x < L; x++)
                R[x] === A && (D = x);
              return D;
            }
            for (var g = o.abs(a - c[h(c, a) + 1]), m = 0, p = [], E = [], b, N, v, C, M; u >= g; )
              b = h(c, a + u), N = h(c, a), p[m] = (i[b] - 2 * i[N] + i[2 * N - b]) / (u * u), u /= 2, m++;
            for (C = p.length, v = 1; C != 1; ) {
              for (M = 0; M < C - 1; M++)
                E[M] = (o.pow(4, v) * p[M + 1] - p[M]) / (o.pow(4, v) - 1);
              C = E.length, p = E, E = [], v++;
            }
            return p;
          },
          simpson: function(c, i, a, u) {
            for (var h = (a - i) / u, g = c(i), m = [], p = i, E = 0, b = 1, N; p <= a; p = p + h, E++)
              m[E] = p;
            for (N = m.length; b < N - 1; b++)
              g += (b % 2 !== 0 ? 4 : 2) * c(m[b]);
            return h / 3 * (g + c(a));
          },
          hermite: function(c, i, a, u) {
            for (var h = c.length, g = 0, m = 0, p = [], E = [], b = [], N = [], v; m < h; m++) {
              for (p[m] = 1, v = 0; v < h; v++)
                m != v && (p[m] *= (u - c[v]) / (c[m] - c[v]));
              for (E[m] = 0, v = 0; v < h; v++)
                m != v && (E[m] += 1 / (c[m] - c[v]));
              b[m] = (1 - 2 * (u - c[m]) * E[m]) * (p[m] * p[m]), N[m] = (u - c[m]) * (p[m] * p[m]), g += b[m] * i[m] + N[m] * a[m];
            }
            return g;
          },
          lagrange: function(c, i, a) {
            for (var u = 0, h = 0, g, m, p = c.length; h < p; h++) {
              for (m = i[h], g = 0; g < p; g++)
                h != g && (m *= (a - c[g]) / (c[h] - c[g]));
              u += m;
            }
            return u;
          },
          cubic_spline: function(c, i, a) {
            for (var u = c.length, h = 0, g, m = [], p = [], E = [], b = [], N = [], v = [], C = []; h < u - 1; h++)
              N[h] = c[h + 1] - c[h];
            for (E[0] = 0, h = 1; h < u - 1; h++)
              E[h] = 3 / N[h] * (i[h + 1] - i[h]) - 3 / N[h - 1] * (i[h] - i[h - 1]);
            for (h = 1; h < u - 1; h++)
              m[h] = [], p[h] = [], m[h][h - 1] = N[h - 1], m[h][h] = 2 * (N[h - 1] + N[h]), m[h][h + 1] = N[h], p[h][0] = E[h];
            for (b = r.multiply(r.inv(m), p), g = 0; g < u - 1; g++)
              v[g] = (i[g + 1] - i[g]) / N[g] - N[g] * (b[g + 1][0] + 2 * b[g][0]) / 3, C[g] = (b[g + 1][0] - b[g][0]) / (3 * N[g]);
            for (g = 0; g < u && !(c[g] > a); g++)
              ;
            return g -= 1, i[g] + (a - c[g]) * v[g] + r.sq(a - c[g]) * b[g] + (a - c[g]) * r.sq(a - c[g]) * C[g];
          },
          gauss_quadrature: function() {
            throw new Error("gauss_quadrature not yet implemented");
          },
          PCA: function(c) {
            var i = c.length, a = c[0].length, u = 0, h, g, m = [], p = [], E = [], b = [], N = [], v = [], C = [], M = [], R = [], A = [];
            for (u = 0; u < i; u++)
              m[u] = r.sum(c[u]) / a;
            for (u = 0; u < a; u++)
              for (C[u] = [], h = 0; h < i; h++)
                C[u][h] = c[h][u] - m[h];
            for (C = r.transpose(C), u = 0; u < i; u++)
              for (M[u] = [], h = 0; h < i; h++)
                M[u][h] = r.dot([C[u]], [C[h]]) / (a - 1);
            for (E = r.jacobi(M), R = E[0], p = E[1], A = r.transpose(R), u = 0; u < p.length; u++)
              for (h = u; h < p.length; h++)
                p[u] < p[h] && (g = p[u], p[u] = p[h], p[h] = g, b = A[u], A[u] = A[h], A[h] = b);
            for (v = r.transpose(C), u = 0; u < i; u++)
              for (N[u] = [], h = 0; h < v.length; h++)
                N[u][h] = r.dot([A[u]], [v[h]]);
            return [c, p, A, N];
          }
        }), (function(d) {
          for (var c = 0; c < d.length; c++) (function(i) {
            r.fn[i] = function(a, u) {
              var h = this;
              return u ? (setTimeout(function() {
                u.call(h, r.fn[i].call(h, a));
              }, 15), this) : typeof r[i](this, a) == "number" ? r[i](this, a) : r(r[i](this, a));
            };
          })(d[c]);
        })("add divide multiply subtract dot pow exp log abs norm angle".split(" "));
      })(n, Math), (function(r, o) {
        var s = [].slice, l = r.utils.isNumber, f = r.utils.isArray;
        r.extend({
          // 2 different parameter lists:
          // (value, mean, sd)
          // (value, array, flag)
          zscore: function() {
            var i = s.call(arguments);
            return l(i[1]) ? (i[0] - i[1]) / i[2] : (i[0] - r.mean(i[1])) / r.stdev(i[1], i[2]);
          },
          // 3 different paramter lists:
          // (value, mean, sd, sides)
          // (zscore, sides)
          // (value, array, sides, flag)
          ztest: function() {
            var i = s.call(arguments), a;
            return f(i[1]) ? (a = r.zscore(i[0], i[1], i[3]), i[2] === 1 ? r.normal.cdf(-o.abs(a), 0, 1) : r.normal.cdf(-o.abs(a), 0, 1) * 2) : i.length > 2 ? (a = r.zscore(i[0], i[1], i[2]), i[3] === 1 ? r.normal.cdf(-o.abs(a), 0, 1) : r.normal.cdf(-o.abs(a), 0, 1) * 2) : (a = i[0], i[1] === 1 ? r.normal.cdf(-o.abs(a), 0, 1) : r.normal.cdf(-o.abs(a), 0, 1) * 2);
          }
        }), r.extend(r.fn, {
          zscore: function(i, a) {
            return (i - this.mean()) / this.stdev(a);
          },
          ztest: function(i, a, u) {
            var h = o.abs(this.zscore(i, u));
            return a === 1 ? r.normal.cdf(-h, 0, 1) : r.normal.cdf(-h, 0, 1) * 2;
          }
        }), r.extend({
          // 2 parameter lists
          // (value, mean, sd, n)
          // (value, array)
          tscore: function() {
            var i = s.call(arguments);
            return i.length === 4 ? (i[0] - i[1]) / (i[2] / o.sqrt(i[3])) : (i[0] - r.mean(i[1])) / (r.stdev(i[1], !0) / o.sqrt(i[1].length));
          },
          // 3 different paramter lists:
          // (value, mean, sd, n, sides)
          // (tscore, n, sides)
          // (value, array, sides)
          ttest: function() {
            var i = s.call(arguments), a;
            return i.length === 5 ? (a = o.abs(r.tscore(i[0], i[1], i[2], i[3])), i[4] === 1 ? r.studentt.cdf(-a, i[3] - 1) : r.studentt.cdf(-a, i[3] - 1) * 2) : l(i[1]) ? (a = o.abs(i[0]), i[2] == 1 ? r.studentt.cdf(-a, i[1] - 1) : r.studentt.cdf(-a, i[1] - 1) * 2) : (a = o.abs(r.tscore(i[0], i[1])), i[2] == 1 ? r.studentt.cdf(-a, i[1].length - 1) : r.studentt.cdf(-a, i[1].length - 1) * 2);
          }
        }), r.extend(r.fn, {
          tscore: function(i) {
            return (i - this.mean()) / (this.stdev(!0) / o.sqrt(this.cols()));
          },
          ttest: function(i, a) {
            return a === 1 ? 1 - r.studentt.cdf(o.abs(this.tscore(i)), this.cols() - 1) : r.studentt.cdf(-o.abs(this.tscore(i)), this.cols() - 1) * 2;
          }
        }), r.extend({
          // Paramter list is as follows:
          // (array1, array2, array3, ...)
          // or it is an array of arrays
          // array of arrays conversion
          anovafscore: function() {
            var i = s.call(arguments), a, u, h, g, m, p, E, b;
            if (i.length === 1) {
              for (m = new Array(i[0].length), E = 0; E < i[0].length; E++)
                m[E] = i[0][E];
              i = m;
            }
            for (u = new Array(), E = 0; E < i.length; E++)
              u = u.concat(i[E]);
            for (h = r.mean(u), a = 0, E = 0; E < i.length; E++)
              a = a + i[E].length * o.pow(r.mean(i[E]) - h, 2);
            for (a /= i.length - 1, p = 0, E = 0; E < i.length; E++)
              for (g = r.mean(i[E]), b = 0; b < i[E].length; b++)
                p += o.pow(i[E][b] - g, 2);
            return p /= u.length - i.length, a / p;
          },
          // 2 different paramter setups
          // (array1, array2, array3, ...)
          // (anovafscore, df1, df2)
          anovaftest: function() {
            var i = s.call(arguments), a, u, h, g;
            if (l(i[0]))
              return 1 - r.centralF.cdf(i[0], i[1], i[2]);
            var m = r.anovafscore(i);
            for (a = i.length - 1, h = 0, g = 0; g < i.length; g++)
              h = h + i[g].length;
            return u = h - a - 1, 1 - r.centralF.cdf(m, a, u);
          },
          ftest: function(i, a, u) {
            return 1 - r.centralF.cdf(i, a, u);
          }
        }), r.extend(r.fn, {
          anovafscore: function() {
            return r.anovafscore(this.toArray());
          },
          anovaftes: function() {
            var i = 0, a;
            for (a = 0; a < this.length; a++)
              i = i + this[a].length;
            return r.ftest(this.anovafscore(), this.length - 1, i - this.length);
          }
        }), r.extend({
          // 2 parameter lists
          // (mean1, mean2, n1, n2, sd)
          // (array1, array2, sd)
          qscore: function() {
            var i = s.call(arguments), a, u, h, g, m;
            return l(i[0]) ? (a = i[0], u = i[1], h = i[2], g = i[3], m = i[4]) : (a = r.mean(i[0]), u = r.mean(i[1]), h = i[0].length, g = i[1].length, m = i[2]), o.abs(a - u) / (m * o.sqrt((1 / h + 1 / g) / 2));
          },
          // 3 different parameter lists:
          // (qscore, n, k)
          // (mean1, mean2, n1, n2, sd, n, k)
          // (array1, array2, sd, n, k)
          qtest: function() {
            var i = s.call(arguments), a;
            i.length === 3 ? (a = i[0], i = i.slice(1)) : i.length === 7 ? (a = r.qscore(i[0], i[1], i[2], i[3], i[4]), i = i.slice(5)) : (a = r.qscore(i[0], i[1], i[2]), i = i.slice(3));
            var u = i[0], h = i[1];
            return 1 - r.tukey.cdf(a, h, u - h);
          },
          tukeyhsd: function(i) {
            for (var a = r.pooledstdev(i), u = i.map(function(b) {
              return r.mean(b);
            }), h = i.reduce(function(b, N) {
              return b + N.length;
            }, 0), g = [], m = 0; m < i.length; ++m)
              for (var p = m + 1; p < i.length; ++p) {
                var E = r.qtest(u[m], u[p], i[m].length, i[p].length, a, h, i.length);
                g.push([[m, p], E]);
              }
            return g;
          }
        }), r.extend({
          // 2 different parameter setups
          // (value, alpha, sd, n)
          // (value, alpha, array)
          normalci: function() {
            var i = s.call(arguments), a = new Array(2), u;
            return i.length === 4 ? u = o.abs(r.normal.inv(i[1] / 2, 0, 1) * i[2] / o.sqrt(i[3])) : u = o.abs(r.normal.inv(i[1] / 2, 0, 1) * r.stdev(i[2]) / o.sqrt(i[2].length)), a[0] = i[0] - u, a[1] = i[0] + u, a;
          },
          // 2 different parameter setups
          // (value, alpha, sd, n)
          // (value, alpha, array)
          tci: function() {
            var i = s.call(arguments), a = new Array(2), u;
            return i.length === 4 ? u = o.abs(r.studentt.inv(i[1] / 2, i[3] - 1) * i[2] / o.sqrt(i[3])) : u = o.abs(r.studentt.inv(i[1] / 2, i[2].length - 1) * r.stdev(i[2], !0) / o.sqrt(i[2].length)), a[0] = i[0] - u, a[1] = i[0] + u, a;
          },
          significant: function(i, a) {
            return i < a;
          }
        }), r.extend(r.fn, {
          normalci: function(i, a) {
            return r.normalci(i, a, this.toArray());
          },
          tci: function(i, a) {
            return r.tci(i, a, this.toArray());
          }
        });
        function d(c, i, a, u) {
          if (c > 1 || a > 1 || c <= 0 || a <= 0)
            throw new Error("Proportions should be greater than 0 and less than 1");
          var h = (c * i + a * u) / (i + u), g = o.sqrt(h * (1 - h) * (1 / i + 1 / u));
          return (c - a) / g;
        }
        r.extend(r.fn, {
          oneSidedDifferenceOfProportions: function(i, a, u, h) {
            var g = d(i, a, u, h);
            return r.ztest(g, 1);
          },
          twoSidedDifferenceOfProportions: function(i, a, u, h) {
            var g = d(i, a, u, h);
            return r.ztest(g, 2);
          }
        });
      })(n, Math), n.models = /* @__PURE__ */ (function() {
        function r(d) {
          var c = d[0].length, i = n.arange(c).map(function(a) {
            var u = n.arange(c).filter(function(h) {
              return h !== a;
            });
            return o(
              n.col(d, a).map(function(h) {
                return h[0];
              }),
              n.col(d, u)
            );
          });
          return i;
        }
        function o(d, c) {
          var i = d.length, a = c[0].length - 1, u = i - a - 1, h = n.lstsq(c, d), g = n.multiply(c, h.map(function(C) {
            return [C];
          })).map(function(C) {
            return C[0];
          }), m = n.subtract(d, g), p = n.mean(d), E = n.sum(g.map(function(C) {
            return Math.pow(C - p, 2);
          })), b = n.sum(d.map(function(C, M) {
            return Math.pow(C - g[M], 2);
          })), N = E + b, v = E / N;
          return {
            exog: c,
            endog: d,
            nobs: i,
            df_model: a,
            df_resid: u,
            coef: h,
            predict: g,
            resid: m,
            ybar: p,
            SST: N,
            SSE: E,
            SSR: b,
            R2: v
          };
        }
        function s(d) {
          var c = r(d.exog), i = Math.sqrt(d.SSR / d.df_resid), a = c.map(function(p) {
            var E = p.SST, b = p.R2;
            return i / Math.sqrt(E * (1 - b));
          }), u = d.coef.map(function(p, E) {
            return (p - 0) / a[E];
          }), h = u.map(function(p) {
            var E = n.studentt.cdf(p, d.df_resid);
            return (E > 0.5 ? 1 - E : E) * 2;
          }), g = n.studentt.inv(0.975, d.df_resid), m = d.coef.map(function(p, E) {
            var b = g * a[E];
            return [p - b, p + b];
          });
          return {
            se: a,
            t: u,
            p: h,
            sigmaHat: i,
            interval95: m
          };
        }
        function l(d) {
          var c = d.R2 / d.df_model / ((1 - d.R2) / d.df_resid), i = function(u, h, g) {
            return n.beta.cdf(u / (g / h + u), h / 2, g / 2);
          }, a = 1 - i(c, d.df_model, d.df_resid);
          return { F_statistic: c, pvalue: a };
        }
        function f(d, c) {
          var i = o(d, c), a = s(i), u = l(i), h = 1 - (1 - i.R2) * ((i.nobs - 1) / i.df_resid);
          return i.t = a, i.f = u, i.adjust_R2 = h, i;
        }
        return { ols: f };
      })(), n.extend({
        buildxmatrix: function() {
          for (var o = new Array(arguments.length), s = 0; s < arguments.length; s++) {
            var l = [1];
            o[s] = l.concat(arguments[s]);
          }
          return n(o);
        },
        builddxmatrix: function() {
          for (var o = new Array(arguments[0].length), s = 0; s < arguments[0].length; s++) {
            var l = [1];
            o[s] = l.concat(arguments[0][s]);
          }
          return n(o);
        },
        buildjxmatrix: function(o) {
          for (var s = new Array(o.length), l = 0; l < o.length; l++)
            s[l] = o[l];
          return n.builddxmatrix(s);
        },
        buildymatrix: function(o) {
          return n(o).transpose();
        },
        buildjymatrix: function(o) {
          return o.transpose();
        },
        matrixmult: function(o, s) {
          var l, f, d, c, i;
          if (o.cols() == s.rows()) {
            if (s.rows() > 1) {
              for (c = [], l = 0; l < o.rows(); l++)
                for (c[l] = [], f = 0; f < s.cols(); f++) {
                  for (i = 0, d = 0; d < o.cols(); d++)
                    i += o.toArray()[l][d] * s.toArray()[d][f];
                  c[l][f] = i;
                }
              return n(c);
            }
            for (c = [], l = 0; l < o.rows(); l++)
              for (c[l] = [], f = 0; f < s.cols(); f++) {
                for (i = 0, d = 0; d < o.cols(); d++)
                  i += o.toArray()[l][d] * s.toArray()[f];
                c[l][f] = i;
              }
            return n(c);
          }
        },
        //regress and regresst to be fixed
        regress: function(o, s) {
          var l = n.xtranspxinv(o), f = o.transpose(), d = n.matrixmult(n(l), f);
          return n.matrixmult(d, s);
        },
        regresst: function(o, s, l) {
          var f = n.regress(o, s), d = {};
          d.anova = {};
          var c = n.jMatYBar(o, f);
          d.yBar = c;
          var i = s.mean();
          d.anova.residuals = n.residuals(s, c), d.anova.ssr = n.ssr(c, i), d.anova.msr = d.anova.ssr / (o[0].length - 1), d.anova.sse = n.sse(s, c), d.anova.mse = d.anova.sse / (s.length - (o[0].length - 1) - 1), d.anova.sst = n.sst(s, i), d.anova.mst = d.anova.sst / (s.length - 1), d.anova.r2 = 1 - d.anova.sse / d.anova.sst, d.anova.r2 < 0 && (d.anova.r2 = 0), d.anova.fratio = d.anova.msr / d.anova.mse, d.anova.pvalue = n.anovaftest(
            d.anova.fratio,
            o[0].length - 1,
            s.length - (o[0].length - 1) - 1
          ), d.anova.rmse = Math.sqrt(d.anova.mse), d.anova.r2adj = 1 - d.anova.mse / d.anova.mst, d.anova.r2adj < 0 && (d.anova.r2adj = 0), d.stats = new Array(o[0].length);
          for (var a = n.xtranspxinv(o), u, h, g, m = 0; m < f.length; m++)
            u = Math.sqrt(d.anova.mse * Math.abs(a[m][m])), h = Math.abs(f[m] / u), g = n.ttest(h, s.length - o[0].length - 1, l), d.stats[m] = [f[m], u, h, g];
          return d.regress = f, d;
        },
        xtranspx: function(o) {
          return n.matrixmult(o.transpose(), o);
        },
        xtranspxinv: function(o) {
          var s = n.matrixmult(o.transpose(), o), l = n.inv(s);
          return l;
        },
        jMatYBar: function(o, s) {
          var l = n.matrixmult(o, s);
          return new n(l);
        },
        residuals: function(o, s) {
          return n.matrixsubtract(o, s);
        },
        ssr: function(o, s) {
          for (var l = 0, f = 0; f < o.length; f++)
            l += Math.pow(o[f] - s, 2);
          return l;
        },
        sse: function(o, s) {
          for (var l = 0, f = 0; f < o.length; f++)
            l += Math.pow(o[f] - s[f], 2);
          return l;
        },
        sst: function(o, s) {
          for (var l = 0, f = 0; f < o.length; f++)
            l += Math.pow(o[f] - s, 2);
          return l;
        },
        matrixsubtract: function(o, s) {
          for (var l = new Array(o.length), f = 0; f < o.length; f++) {
            l[f] = new Array(o[f].length);
            for (var d = 0; d < o[f].length; d++)
              l[f][d] = o[f][d] - s[f][d];
          }
          return n(l);
        }
      }), n.jStat = n, n;
    });
  })(_r)), _r.exports;
}
var Xu = Yu();
const U = /* @__PURE__ */ Pa(Xu);
var Mo = {}, Ss;
function Ku() {
  return Ss || (Ss = 1, (function(e) {
    (function(t) {
      t(typeof DO_NOT_EXPORT_BESSEL > "u" ? e : {});
    })(function(t) {
      t.version = "1.0.2";
      var n = Math;
      function r(i, a) {
        for (var u = 0, h = 0; u < i.length; ++u) h = a * h + i[u];
        return h;
      }
      function o(i, a, u, h, g) {
        if (a === 0) return u;
        if (a === 1) return h;
        for (var m = 2 / i, p = h, E = 1; E < a; ++E)
          p = h * E * m + g * u, u = h, h = p;
        return p;
      }
      function s(i, a, u, h, g) {
        return function(p, E) {
          if (h) {
            if (p === 0) return h == 1 ? -1 / 0 : 1 / 0;
            if (p < 0) return NaN;
          }
          if (E === 0) return i(p);
          if (E === 1) return a(p);
          if (E < 0) return NaN;
          E |= 0;
          var b = i(p), N = a(p);
          return o(p, E, b, N, g);
        };
      }
      var l = (function() {
        var i = 0.636619772, a = [57568490574, -13362590354, 6516196407e-1, -1121442418e-2, 77392.33017, -184.9052456].reverse(), u = [57568490411, 1029532985, 9494680718e-3, 59272.64853, 267.8532712, 1].reverse(), h = [1, -0.001098628627, 2734510407e-14, -2073370639e-15, 2093887211e-16].reverse(), g = [-0.01562499995, 1430488765e-13, -6911147651e-15, 7621095161e-16, -934935152e-16].reverse();
        function m(C) {
          var M = 0, R = 0, A = 0, x = C * C;
          if (C < 8)
            R = r(a, x), A = r(u, x), M = R / A;
          else {
            var L = C - 0.785398164;
            x = 64 / x, R = r(h, x), A = r(g, x), M = n.sqrt(i / C) * (n.cos(L) * R - n.sin(L) * A * 8 / C);
          }
          return M;
        }
        var p = [72362614232, -7895059235, 2423968531e-1, -2972611439e-3, 15704.4826, -30.16036606].reverse(), E = [144725228442, 2300535178, 1858330474e-2, 99447.43394, 376.9991397, 1].reverse(), b = [1, 183105e-8, -3516396496e-14, 2457520174e-15, -240337019e-15].reverse(), N = [0.04687499995, -2002690873e-13, 8449199096e-15, -88228987e-14, 105787412e-15].reverse();
        function v(C) {
          var M = 0, R = 0, A = 0, x = C * C, L = n.abs(C) - 2.356194491;
          return Math.abs(C) < 8 ? (R = C * r(p, x), A = r(E, x), M = R / A) : (x = 64 / x, R = r(b, x), A = r(N, x), M = n.sqrt(i / n.abs(C)) * (n.cos(L) * R - n.sin(L) * A * 8 / n.abs(C)), C < 0 && (M = -M)), M;
        }
        return function C(M, R) {
          if (R = Math.round(R), !isFinite(M)) return isNaN(M) ? M : 0;
          if (R < 0) return (R % 2 ? -1 : 1) * C(M, -R);
          if (M < 0) return (R % 2 ? -1 : 1) * C(-M, R);
          if (R === 0) return m(M);
          if (R === 1) return v(M);
          if (M === 0) return 0;
          var A = 0;
          if (M > R)
            A = o(M, R, m(M), v(M), -1);
          else {
            for (var x = 2 * n.floor((R + n.floor(n.sqrt(40 * R))) / 2), L = !1, D = 0, W = 0, J = 1, ne = 0, K = 2 / M, q = x; q > 0; q--)
              ne = q * K * J - D, D = J, J = ne, n.abs(J) > 1e10 && (J *= 1e-10, D *= 1e-10, A *= 1e-10, W *= 1e-10), L && (W += J), L = !L, q == R && (A = D);
            W = 2 * W - J, A /= W;
          }
          return A;
        };
      })(), f = (function() {
        var i = 0.636619772, a = [-2957821389, 7062834065, -5123598036e-1, 1087988129e-2, -86327.92757, 228.4622733].reverse(), u = [40076544269, 7452499648e-1, 7189466438e-3, 47447.2647, 226.1030244, 1].reverse(), h = [1, -0.001098628627, 2734510407e-14, -2073370639e-15, 2093887211e-16].reverse(), g = [-0.01562499995, 1430488765e-13, -6911147651e-15, 7621095161e-16, -934945152e-16].reverse();
        function m(C) {
          var M = 0, R = 0, A = 0, x = C * C, L = C - 0.785398164;
          return C < 8 ? (R = r(a, x), A = r(u, x), M = R / A + i * l(C, 0) * n.log(C)) : (x = 64 / x, R = r(h, x), A = r(g, x), M = n.sqrt(i / C) * (n.sin(L) * R + n.cos(L) * A * 8 / C)), M;
        }
        var p = [-4900604943e3, 127527439e4, -51534381390, 7349264551e-1, -4237922726e-3, 8511.937935].reverse(), E = [249958057e5, 424441966400, 3733650367, 2245904002e-2, 102042.605, 354.9632885, 1].reverse(), b = [1, 183105e-8, -3516396496e-14, 2457520174e-15, -240337019e-15].reverse(), N = [0.04687499995, -2002690873e-13, 8449199096e-15, -88228987e-14, 105787412e-15].reverse();
        function v(C) {
          var M = 0, R = 0, A = 0, x = C * C, L = C - 2.356194491;
          return C < 8 ? (R = C * r(p, x), A = r(E, x), M = R / A + i * (l(C, 1) * n.log(C) - 1 / C)) : (x = 64 / x, R = r(b, x), A = r(N, x), M = n.sqrt(i / C) * (n.sin(L) * R + n.cos(L) * A * 8 / C)), M;
        }
        return s(m, v, "BESSELY", 1, -1);
      })(), d = (function() {
        var i = [1, 3.5156229, 3.0899424, 1.2067492, 0.2659732, 0.0360768, 45813e-7].reverse(), a = [0.39894228, 0.01328592, 225319e-8, -157565e-8, 916281e-8, -0.02057706, 0.02635537, -0.01647633, 392377e-8].reverse();
        function u(p) {
          return p <= 3.75 ? r(i, p * p / (3.75 * 3.75)) : n.exp(n.abs(p)) / n.sqrt(n.abs(p)) * r(a, 3.75 / n.abs(p));
        }
        var h = [0.5, 0.87890594, 0.51498869, 0.15084934, 0.02658733, 301532e-8, 32411e-8].reverse(), g = [0.39894228, -0.03988024, -362018e-8, 163801e-8, -0.01031555, 0.02282967, -0.02895312, 0.01787654, -420059e-8].reverse();
        function m(p) {
          return p < 3.75 ? p * r(h, p * p / (3.75 * 3.75)) : (p < 0 ? -1 : 1) * n.exp(n.abs(p)) / n.sqrt(n.abs(p)) * r(g, 3.75 / n.abs(p));
        }
        return function p(E, b) {
          if (b = Math.round(b), b === 0) return u(E);
          if (b === 1) return m(E);
          if (b < 0) return NaN;
          if (n.abs(E) === 0) return 0;
          if (E == 1 / 0) return 1 / 0;
          var N = 0, v, C = 2 / n.abs(E), M = 0, R = 1, A = 0, x = 2 * n.round((b + n.round(n.sqrt(40 * b))) / 2);
          for (v = x; v > 0; v--)
            A = v * C * R + M, M = R, R = A, n.abs(R) > 1e10 && (R *= 1e-10, M *= 1e-10, N *= 1e-10), v == b && (N = M);
          return N *= p(E, 0) / R, E < 0 && b % 2 ? -N : N;
        };
      })(), c = (function() {
        var i = [-0.57721566, 0.4227842, 0.23069756, 0.0348859, 262698e-8, 1075e-7, 74e-7].reverse(), a = [1.25331414, -0.07832358, 0.02189568, -0.01062446, 587872e-8, -25154e-7, 53208e-8].reverse();
        function u(p) {
          return p <= 2 ? -n.log(p / 2) * d(p, 0) + r(i, p * p / 4) : n.exp(-p) / n.sqrt(p) * r(a, 2 / p);
        }
        var h = [1, 0.15443144, -0.67278579, -0.18156897, -0.01919402, -110404e-8, -4686e-8].reverse(), g = [1.25331414, 0.23498619, -0.0365562, 0.01504268, -780353e-8, 325614e-8, -68245e-8].reverse();
        function m(p) {
          return p <= 2 ? n.log(p / 2) * d(p, 1) + 1 / p * r(h, p * p / 4) : n.exp(-p) / n.sqrt(p) * r(g, 2 / p);
        }
        return s(u, m, "BESSELK", 2, 1);
      })();
      t.besselj = l, t.bessely = f, t.besseli = d, t.besselk = c;
    });
  })(Mo)), Mo;
}
var Ju = Ku();
const so = /* @__PURE__ */ Pa(Ju), Si = new Error("#NULL!"), Ue = new Error("#DIV/0!"), S = new Error("#VALUE!"), _n = new Error("#REF!"), ao = new Error("#NAME?"), T = new Error("#NUM!"), G = new Error("#N/A"), Ri = new Error("#ERROR!"), Ua = new Error("#GETTING_DATA"), ni = new Error("#CALC!");
var Qu = /* @__PURE__ */ Object.freeze({
  __proto__: null,
  calc: ni,
  data: Ua,
  div0: Ue,
  error: Ri,
  na: G,
  name: ao,
  nil: Si,
  num: T,
  ref: _n,
  value: S
});
let _t = !1;
function ju() {
  _t = !0;
}
function Zu() {
  _t = !1;
}
function _a(e) {
  e < 60 && (e += 1);
  const n = Math.floor(e - 25569) * 86400, r = new Date(n * 1e3), o = e - Math.floor(e) + 1e-7;
  let s = Math.floor(86400 * o);
  const l = s % 60;
  s -= l;
  const f = Math.floor(s / 3600), d = Math.floor(s / 60) % 60;
  let c = r.getUTCDate(), i = r.getUTCMonth();
  return e >= 60 && e < 61 && (c = 29, i = 1), new Date(r.getUTCFullYear(), i, c, f, d, l);
}
function Ft(e) {
  const t = new Date(1900, 0, 1), n = e > -22038912e5 ? 2 : 1;
  return Math.ceil((e - t) / 864e5) + n;
}
var ef = /* @__PURE__ */ Object.freeze({
  __proto__: null,
  dateToSerial: Ft,
  get returnSerial() {
    return _t;
  },
  serialToDate: _a,
  useDate: Zu,
  useSerial: ju
});
const tf = "=", nf = [">", ">=", "<", "<=", "=", "<>"], Ba = "operator", za = "literal", rf = [Ba, za], Br = Ba, jt = za;
function kt(e, t) {
  if (rf.indexOf(t) === -1)
    throw new Error("Unsupported token type: " + t);
  return {
    value: e,
    type: t
  };
}
function of(e) {
  return typeof e != "string" || /^\d+(\.\d+)?$/.test(e) && (e = e.indexOf(".") === -1 ? parseInt(e, 10) : parseFloat(e)), e;
}
function sf(e) {
  const t = e.length, n = [];
  let r = 0, o = "", s = "";
  for (; r < t; ) {
    const l = e.charAt(r);
    switch (l) {
      case ">":
      case "<":
      case "=":
        s = s + l, o.length > 0 && (n.push(o), o = "");
        break;
      default:
        s.length > 0 && (n.push(s), s = ""), o = o + l;
        break;
    }
    r++;
  }
  return o.length > 0 && n.push(o), s.length > 0 && n.push(s), n;
}
function af(e) {
  let t = "";
  const n = [];
  for (let r = 0; r < e.length; r++) {
    const o = e[r];
    r === 0 && nf.indexOf(o) >= 0 ? n.push(kt(o, Br)) : t += o;
  }
  return t.length > 0 && n.push(kt(of(t), jt)), n.length > 0 && n[0].type !== Br && n.unshift(kt(tf, Br)), n;
}
function lf(e) {
  const t = [];
  let n;
  for (let r = 0; r < e.length; r++) {
    const o = e[r];
    switch (o.type) {
      case Br:
        n = o.value;
        break;
      case jt:
        t.push(o.value);
        break;
    }
  }
  return cf(t, n);
}
function cf(e, t) {
  let n = !1;
  switch (t) {
    case ">":
      n = e[0] > e[1];
      break;
    case ">=":
      n = e[0] >= e[1];
      break;
    case "<":
      n = e[0] < e[1];
      break;
    case "<=":
      n = e[0] <= e[1];
      break;
    case "=":
      n = e[0] == e[1];
      break;
    case "<>":
      n = e[0] != e[1];
      break;
  }
  return n;
}
function Bn(e) {
  return af(sf(e));
}
const zn = lf;
function lo(e) {
  const t = [];
  return He(e, (n) => {
    t.push(n);
  }), t;
}
function He(e, t) {
  let n = -1;
  const r = e.length;
  for (; ++n < r && t(e[n], n, e) !== !1; )
    ;
  return e;
}
function Ai(e) {
  let t = e.length, n;
  for (; t--; )
    if (n = e[t], typeof n != "number") {
      if (n === !0) {
        e[t] = 1;
        continue;
      }
      if (n === !1) {
        e[t] = 0;
        continue;
      }
      if (typeof n == "string") {
        const r = w(n);
        e[t] = r instanceof Error ? 0 : r;
      }
    }
  return e;
}
function gr(e) {
  const t = e.length, n = e.reduce((r, o) => Math.max(r, o.length), 0);
  return [t, n];
}
function uf(e, t) {
  const n = [e, t];
  return n.some((r) => !Ye(r)) ? G : n.some((r) => r <= 0) ? S : Array.from({ length: e }, () => Array.from({ length: t }, () => {
  }));
}
function Ha(e, t) {
  if (!e)
    return S;
  (!e.every((o) => Array.isArray(o)) || e.length === 0) && (e = [[...e]]), e.map((o, s) => {
    o.map((l, f) => {
      l || (e[s][f] = 0);
    });
  });
  const n = e.reduce((o, s, l) => s.length > e[o].length ? l : o, 0), r = e[n].length;
  return e.map((o) => [...o, ...Array(r - o.length).fill(0)]);
}
function P() {
  let e;
  if (arguments.length === 1) {
    const t = arguments[0];
    e = df(t) ? lo.apply(null, arguments) : [t];
  } else
    e = Array.from(arguments);
  for (; !hf(e); )
    e = ri(e);
  return e;
}
function ri(e) {
  return !e || !e.reduce ? [e] : e.reduce((t, n) => {
    const r = Array.isArray(t), o = Array.isArray(n);
    return r && o ? t.concat(n) : r ? (t.push(n), t) : o ? [t].concat(n) : [t, n];
  });
}
function ff(e, t) {
  return t = t || 1, !e || typeof e.slice != "function" ? e : e.slice(0, e.length - t);
}
function df(e) {
  return e != null && typeof e.length == "number" && typeof e != "string";
}
function hf(e) {
  if (!e)
    return !1;
  for (let t = 0; t < e.length; ++t)
    if (Array.isArray(e[t]))
      return !1;
  return !0;
}
function Ne(e, t) {
  return t = t || 1, !e || typeof e.slice != "function" ? e : e.slice(t);
}
function Vr(e) {
  return e ? e[0].map((t, n) => e.map((r) => r[n])) : S;
}
function gt(e, t) {
  let n = null;
  return He(e, (r, o) => {
    if (r[0] === t)
      return n = o, !1;
  }), n ?? S;
}
function Q() {
  for (let e = 0; e < arguments.length; e++)
    if (arguments[e] instanceof Error)
      return arguments[e];
}
function O() {
  let e = arguments.length;
  for (; e--; )
    if (arguments[e] instanceof Error)
      return !0;
  return !1;
}
function Va(e) {
  return Math.round(e * 1e14) / 1e14;
}
function Bt() {
  return P.apply(null, arguments).filter((t) => typeof t == "number");
}
function Ti(e) {
  if (typeof e == "boolean" || e instanceof Error)
    return e;
  if (typeof e == "number")
    return e !== 0;
  if (typeof e == "string") {
    const t = e.toUpperCase();
    if (t === "TRUE")
      return !0;
    if (t === "FALSE")
      return !1;
  }
  return e instanceof Date && !isNaN(e) ? !0 : S;
}
function ee(e) {
  if (!isNaN(e)) {
    if (e instanceof Date)
      return new Date(e);
    const t = parseFloat(e);
    return t < 0 || t >= 2958466 ? T : _a(t);
  }
  return typeof e == "string" && (e = /(\d{4})-(\d\d?)-(\d\d?)$/.test(e) ? /* @__PURE__ */ new Date(e + "T00:00:00.000") : new Date(e), !isNaN(e)) ? e : S;
}
function Ga(e) {
  let t = e.length, n;
  for (; t--; ) {
    if (n = ee(e[t]), n === S)
      return n;
    e[t] = n;
  }
  return e;
}
function w(e) {
  return e instanceof Error ? e : e == null ? 0 : (typeof e == "boolean" && (e = +e), !isNaN(e) && e !== "" ? parseFloat(e) : S);
}
function _(e) {
  let t;
  if (!e || (t = e.length) === 0)
    return S;
  let n;
  for (; t--; ) {
    if (e[t] instanceof Error)
      return e[t];
    if (n = w(e[t]), n instanceof Error)
      return n;
    e[t] = n;
  }
  return e;
}
function Me(e) {
  return e instanceof Error ? e : e == null ? "" : e.toString();
}
function co() {
  let e = arguments.length;
  for (; e--; )
    if (typeof arguments[e] == "string")
      return !0;
  return !1;
}
function uo() {
  const e = lo(arguments), t = _(P(e.shift()));
  if (t instanceof Error)
    return t;
  const n = e, r = n.length / 2;
  for (let s = 0; s < r; s++)
    n[s * 2] = P(n[s * 2]);
  let o = [];
  for (let s = 0; s < t.length; s++) {
    let l = !1;
    for (let f = 0; f < r; f++) {
      const d = n[f * 2][s], c = n[f * 2 + 1], i = c === void 0 || c === "*";
      let a = !1;
      if (i)
        a = !0;
      else {
        const u = Bn(c + ""), h = [kt(d, jt)].concat(
          u
        );
        a = zn(h);
      }
      if (!a) {
        l = !1;
        break;
      }
      l = !0;
    }
    l && o.push(t[s]);
  }
  return o;
}
function Ye(e) {
  return e != null;
}
const $a = {};
$a.TYPE = (e) => {
  switch (e) {
    case Si:
      return 1;
    case Ue:
      return 2;
    case S:
      return 3;
    case _n:
      return 4;
    case ao:
      return 5;
    case T:
      return 6;
    case G:
      return 7;
    case Ua:
      return 8;
  }
  return G;
};
function mf(e) {
  return e === null;
}
function qa(e) {
  return [S, _n, Ue, T, ao, Si].indexOf(e) >= 0 || typeof e == "number" && (isNaN(e) || !isFinite(e));
}
function fo(e) {
  return qa(e) || e === G;
}
function gf(e) {
  return !(Math.floor(Math.abs(e)) & 1);
}
function Wa(e) {
  return e === !0 || e === !1;
}
function pf(e) {
  return e === G;
}
function vf(e) {
  return typeof e != "string";
}
function ho(e) {
  return typeof e == "number" && !isNaN(e) && isFinite(e);
}
function Nf(e) {
  return !!(Math.floor(Math.abs(e)) & 1);
}
function Ya(e) {
  return typeof e == "string";
}
function Cf(e) {
  return ho(e) ? e : e instanceof Date ? e.getTime() : e === !0 ? 1 : e === !1 ? 0 : fo(e) ? e : 0;
}
function Ef() {
  return G;
}
function bf(e) {
  if (ho(e))
    return 1;
  if (Ya(e))
    return 2;
  if (Wa(e))
    return 4;
  if (fo(e))
    return 16;
  if (Array.isArray(e))
    return 64;
}
function wf() {
  if (arguments.length < 2)
    return G;
  const e = arguments[0];
  return e < 1 || e > 254 || arguments.length < e + 1 ? S : arguments[e];
}
function If(e, t, ...n) {
  if (!Array.isArray(e)) return S;
  const r = P([t, ...n]), [o, s] = gr(e);
  if (r.some((c) => {
    if (!Ye(c)) return !0;
    const i = Math.abs(c);
    return !(0 < i && i <= s);
  })) return S;
  const f = r.map((c) => c < 0 ? c + s + 1 : c), d = new Array(o);
  for (let c = 0; c < o; c++) {
    d[c] = new Array(f.length);
    for (let i = 0; i < f.length; i++)
      d[c][i] = e[c][f[i] - 1] ?? 0;
  }
  return d;
}
function Mf(e, t, ...n) {
  if (!Array.isArray(e)) return S;
  const r = P([t, ...n]);
  return r.some((l) => {
    if (!Ye(l)) return !0;
    const f = Math.abs(l);
    return !(0 < f && f <= e.length);
  }) ? S : r.map((l) => l < 0 ? l + e.length + 1 : l).reduce((l, f) => {
    const d = [...e[--f]].map((c) => Ye(c) ? c : 0);
    return l.push(d), l;
  }, []);
}
function yf(e, t) {
  if (arguments.length !== 2)
    return G;
  if (t < 0)
    return T;
  if (!(e instanceof Array) || typeof t != "number")
    return S;
  if (e.length !== 0)
    return U.col(e, t);
}
function Sf(e) {
  return arguments.length !== 1 ? G : e instanceof Array ? e.length === 0 ? 0 : U.cols(e) : S;
}
function Rf(e, t, n) {
  if (O(e, t, n))
    return Q(e, t, n);
  if (!(e instanceof Array))
    return S;
  const r = [], [o, s] = gr(e);
  let l = 0, f = o;
  if (typeof t == "number") {
    if (Math.abs(t) > o)
      return ni;
    t > 0 ? l = t : t < 0 && (f += t);
  }
  let d = 0, c = s;
  if (typeof n == "number") {
    if (Math.abs(n) > s)
      return ni;
    n > 0 ? d = n : n < 0 && (c += n);
  }
  for (let i = l; i < f; i++) {
    const a = [...e[i].slice(d, c).map((u) => u ?? 0)];
    r.push(a);
  }
  return r;
}
function Af(e, t, n, r) {
  const [o, s] = gr(e);
  if (Ye(t) || (t = o), Ye(n) || (n = s), t < o || n < s) return S;
  const l = uf(t, n);
  for (let f = 0; f < t; f++)
    for (let d = 0; d < n; d++) {
      let c = e[f]?.[d];
      Ye(c) || (c = f < o && d < s ? 0 : r ?? G), l[f][d] = c;
    }
  return l;
}
function Tf(e, t, n, r) {
  return Xa(e, Vr(t), n, r);
}
function xf(e, t, n) {
  const r = Q(e, t, n);
  if (r)
    return r;
  if (!Array.isArray(e))
    return S;
  const o = e.length > 0 && !Array.isArray(e[0]);
  return o && !n ? (n = t, t = 1) : (n = n || 1, t = t || 1), n < 0 || t < 0 ? S : o && t === 1 && n <= e.length ? e[n - 1] : t <= e.length && n <= e[t - 1].length ? e[t - 1][n - 1] : _n;
}
function Df(e, t, n) {
  t = P(t), n = n ? P(n) : t;
  const r = typeof e == "number";
  let o = G;
  for (let s = 0; s < t.length; s++) {
    if (t[s] === e)
      return n[s];
    if (r && t[s] <= e || typeof t[s] == "string" && t[s].localeCompare(e) < 0)
      o = n[s];
    else if (r && t[s] > e)
      return o;
  }
  return o;
}
function Of(e, t, n) {
  if (!e && e !== 0 || !t || (arguments.length === 2 && (n = 1), t = P(t), !(t instanceof Array)) || n !== -1 && n !== 0 && n !== 1)
    return G;
  let r, o;
  for (let s = 0; s < t.length; s++)
    if (n === 1) {
      if (t[s] === e)
        return s + 1;
      t[s] < e && (o ? t[s] > o && (r = s + 1, o = t[s]) : (r = s + 1, o = t[s]));
    } else if (n === 0) {
      if (typeof e == "string" && typeof t[s] == "string") {
        const l = e.toLowerCase().replace(/\?/g, ".").replace(/\*/g, ".*").replace(/~/g, "\\").replace(/\+/g, "\\+").replace(/\(/g, "\\(").replace(/\)/g, "\\)").replace(/\[/g, "\\[").replace(/\]/g, "\\]");
        if (new RegExp("^" + l + "$").test(t[s].toLowerCase()))
          return s + 1;
      } else if (t[s] === e)
        return s + 1;
    } else if (n === -1) {
      if (t[s] === e)
        return s + 1;
      t[s] > e && (o ? t[s] < o && (r = s + 1, o = t[s]) : (r = s + 1, o = t[s]));
    }
  return r || G;
}
function kf(e) {
  return arguments.length !== 1 ? G : e instanceof Array ? e.length === 0 ? 0 : U.rows(e) : S;
}
function Lf(e, t = 1, n = 1, r = !1) {
  if (!e || !Array.isArray(e))
    return G;
  if (e.length === 0)
    return 0;
  if (t = w(t), !t || t < 1 || (n = w(n), n !== 1 && n !== -1))
    return S;
  if (r = Ti(r), typeof r != "boolean")
    return ao;
  const o = (f) => f.sort((d, c) => (d = Me(d[t - 1]), c = Me(c[t - 1]), n === 1 ? d < c ? n * -1 : n : d > c ? n : n * -1)), s = Ha(e), l = r ? Vr(s) : s;
  return t >= 1 && t <= l[0].length ? r ? Vr(o(l)) : o(l) : S;
}
function Ff(e) {
  if (!e)
    return G;
  const t = Ha(e);
  return Vr(t);
}
function xi() {
  const e = [];
  for (let t = 0; t < arguments.length; ++t) {
    let n = !1;
    const r = arguments[t];
    for (let o = 0; o < e.length && (n = e[o] === r, !n); ++o)
      ;
    n || e.push(r);
  }
  return e;
}
function Xa(e, t, n, r) {
  if (!t || !n)
    return G;
  r = !(r === 0 || r === !1);
  let o = G, s = !1;
  const l = typeof e == "number", f = typeof e == "string" ? e.toLowerCase() : e;
  for (let d = 0; d < t.length; d++) {
    const c = t[d], i = typeof c[0] == "string" ? c[0].toLowerCase() : c[0];
    if (i === f) {
      o = n < c.length + 1 ? c[n - 1] : _n;
      break;
    } else !s && (l && r && i <= e || r && typeof i == "string" && i.localeCompare(e) < 0) && (o = n < c.length + 1 ? c[n - 1] : _n);
    l && i > e && (s = !0);
  }
  return o;
}
function Pf(e, ...t) {
  const n = [e, ...t];
  let r = 0, o = 0;
  const s = [];
  for (const d of n) {
    const [c, i] = gr(d);
    r = Math.max(r, c), o += i, s.push(i);
  }
  const l = Array.from({ length: r }, () => new Array(o));
  let f = 0;
  for (let d = 0; d < n.length; d++) {
    const c = n[d], i = s[d];
    for (let a = 0; a < i; a++) {
      for (let u = 0; u < r; u++)
        l[u][f] = c[u]?.[a] ?? (u < c.length ? 0 : G);
      f++;
    }
  }
  return l;
}
function Uf(e, ...t) {
  const n = [e, ...t];
  let r = 0, o = 0;
  for (const d of n) {
    const [c, i] = gr(d);
    r = Math.max(r, i), o += c;
  }
  const s = new Array(o);
  let l, f = -1;
  for (; (l = n.shift()) !== void 0; )
    for (let d = 0; d < l.length; d++) {
      s[++f] = new Array(r);
      for (let c = 0; c < r; c++)
        s[f][c] = l[d][c] ?? (c < l[d].length ? 0 : G);
    }
  return s;
}
function Ka(e) {
  return e = w(e), e === 0 ? S : e instanceof Error ? e : String.fromCharCode(e);
}
function _f(e) {
  if (O(e))
    return e;
  e = e || "";
  const t = /[\0-\x1F]/g;
  return e.replace(t, "");
}
function Ja(e) {
  if (O(e))
    return e;
  e = e || "";
  let t = e.charCodeAt(0);
  return isNaN(t) && (t = S), t;
}
function Qa() {
  const e = P(arguments), t = Q.apply(void 0, e);
  if (t)
    return t;
  let n = 0;
  for (; (n = e.indexOf(!0)) > -1; )
    e[n] = "TRUE";
  let r = 0;
  for (; (r = e.indexOf(!1)) > -1; )
    e[r] = "FALSE";
  return e.join("");
}
const Bf = Qa;
function zf(e, t = 2) {
  if (e = w(e), isNaN(e))
    return S;
  e = cl(e, t);
  const n = {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: t >= 0 ? t : 0,
    maximumFractionDigits: t >= 0 ? t : 0
  }, r = e.toLocaleString("en-US", n);
  return e < 0 ? "$(" + r.slice(2) + ")" : r;
}
function Hf(e, t) {
  if (arguments.length !== 2)
    return G;
  const n = Q(e, t);
  return n || (e = Me(e), t = Me(t), e === t);
}
function Vf(e, t, n) {
  if (arguments.length < 2)
    return G;
  e = Me(e), t = Me(t), n = n === void 0 ? 0 : n;
  const r = t.indexOf(e, n - 1);
  return r === -1 ? S : r + 1;
}
function ja(e, t = 2, n = !1) {
  if (e = w(e), isNaN(e) || (t = w(t), isNaN(t)))
    return S;
  if (t < 0) {
    const r = Math.pow(10, -t);
    e = Math.round(e / r) * r;
  } else
    e = e.toFixed(t);
  if (n)
    e = e.toString().replace(/,/g, "");
  else {
    const r = e.toString().split(".");
    r[0] = r[0].replace(/\B(?=(\d{3})+$)/g, ","), e = r.join(".");
  }
  return e;
}
function Gf(e, t) {
  const n = Q(e, t);
  return n || (e = Me(e), t = t === void 0 ? 1 : t, t = w(t), t instanceof Error || typeof e != "string" ? S : e.substring(0, t));
}
function $f(e) {
  return arguments.length === 0 ? Ri : e instanceof Error ? e : Array.isArray(e) ? S : Me(e).length;
}
function qf(e) {
  return arguments.length !== 1 ? S : (e = Me(e), O(e) ? e : e.toLowerCase());
}
function Wf(e, t, n) {
  if (O(e, t, n))
    return Q(e, t, n);
  const r = [t, n].map((f) => w(f));
  if (!r.every((f) => !(f instanceof Error) && f > 0)) return S;
  Ye(e) ? typeof e != "string" && (e = String(e)) : e = "", [t, n] = r;
  const s = t - 1, l = s + n;
  return e.substring(s, l);
}
function Yf(e, t, n) {
  return e = Ye(e) ? e : "", typeof e == "number" ? e : typeof e != "string" ? G : (t = typeof t > "u" ? "." : t, n = typeof n > "u" ? "," : n, Number(e.replace(t, ".").replace(n, "")));
}
function Xf(e) {
  return O(e) ? e : isNaN(e) && typeof e == "number" ? S : (e = Me(e), e.replace(/\w\S*/g, (t) => t.charAt(0).toUpperCase() + t.substr(1).toLowerCase()));
}
function Kf(e, t, n, r) {
  return t = w(t), n = w(n), O(t, n) || typeof e != "string" || typeof r != "string" ? S : e.substr(0, t - 1) + r + e.substr(t - 1 + n);
}
function at(e, t) {
  const n = Q(e, t);
  return n || (e = Me(e), t = w(t), t instanceof Error ? t : new Array(t + 1).join(e));
}
function Jf(e, t) {
  const n = Q(e, t);
  return n || (e = Me(e), t = t === void 0 ? 1 : t, t = w(t), t instanceof Error ? t : e.substring(e.length - t));
}
function Qf(e, t, n) {
  let r;
  return typeof e != "string" || typeof t != "string" ? S : (n = n === void 0 ? 0 : n, r = t.toLowerCase().indexOf(e.toLowerCase(), n - 1) + 1, r === 0 ? S : r);
}
function jf(e, t, n, r) {
  if (arguments.length < 3)
    return G;
  if (!e || !t)
    return e;
  if (r === void 0)
    return e.split(t).join(n);
  {
    if (r = Math.floor(Number(r)), Number.isNaN(r) || r <= 0)
      return S;
    let o = 0, s = 0;
    for (; o > -1 && e.indexOf(t, o) > -1; )
      if (o = e.indexOf(t, o + 1), s++, o > -1 && s === r)
        return e.substring(0, o) + n + e.substring(o + t.length);
    return e;
  }
}
function et(e) {
  return e instanceof Error || typeof e == "string" ? e : "";
}
function Zf(e, t) {
  if (e === void 0 || e instanceof Error || t instanceof Error)
    return G;
  if (e instanceof Date)
    return e.toISOString().slice(0, 10);
  if (t == null)
    return "";
  if (typeof t == "number")
    return String(t);
  if (typeof t != "string")
    return S;
  const n = t.startsWith("$") ? "$" : "", r = t.endsWith("%");
  t = t.replace(/%/g, "").replace(/\$/g, "");
  const o = t.includes(".") ? t.split(".")[1].match(/0/g).length : 0, s = !t.includes(",");
  return r && (e = e * 100), e = ja(e, o, s), e.startsWith("-") ? (e = e.replace("-", ""), e = "-" + n + e) : e = n + e, r && (e = e + "%"), e;
}
function ed(e, t, ...n) {
  if (typeof t != "boolean" && (t = Ti(t)), arguments.length < 3)
    return G;
  e = e ?? "";
  let r = P(n), o = t ? r.filter((s) => s) : r;
  if (Array.isArray(e)) {
    e = P(e);
    let s = o.map((f) => [f]), l = 0;
    for (let f = 0; f < s.length - 1; f++)
      s[f].push(e[l]), l++, l === e.length && (l = 0);
    return o = P(s), o.join("");
  }
  return o.join(e);
}
function td(e) {
  return e = Me(e), e instanceof Error ? e : e.replace(/\s+/g, " ").trim();
}
const nd = Ka, rd = Ja;
function od(e) {
  return e = Me(e), e instanceof Error ? e : e.toUpperCase();
}
function id(e) {
  const t = Q(e);
  if (t)
    return t;
  if (typeof e == "number")
    return e;
  if (Ye(e) || (e = ""), typeof e != "string")
    return S;
  const n = /(%)$/.test(e) || /^(%)/.test(e);
  if (e = e.replace(/^[^0-9-]{0,3}/, ""), e = e.replace(/[^0-9]{0,3}$/, ""), e = e.replace(/[ ,]/g, ""), e === "")
    return 0;
  let r = Number(e);
  return isNaN(r) ? S : (r = r || 0, n && (r = r * 0.01), r);
}
const sd = 2.5066282746310002;
function ad() {
  const t = P(arguments).filter(Ye);
  if (t.length === 0)
    return T;
  const n = _(t);
  return n instanceof Error ? n : U.sum(U(n).subtract(U.mean(n)).abs()[0]) / n.length;
}
function fn() {
  const t = P(arguments).filter(Ye);
  if (t.length === 0)
    return Ue;
  const n = Q.apply(void 0, t);
  if (n)
    return n;
  const r = Bt(t), o = r.length;
  let s = 0, l = 0, f;
  for (let d = 0; d < o; d++)
    s += r[d], l += 1;
  return f = s / l, isNaN(f) && (f = T), f;
}
function Di() {
  const t = P(arguments).filter(Ye);
  if (t.length === 0)
    return Ue;
  const n = Q.apply(void 0, t);
  if (n)
    return n;
  const r = t, o = r.length;
  let s = 0, l = 0, f;
  for (let d = 0; d < o; d++) {
    const c = r[d];
    typeof c == "number" && (s += c), c === !0 && s++, c !== null && l++;
  }
  return f = s / l, isNaN(f) && (f = T), f;
}
function ld(e, t, n) {
  if (arguments.length <= 1)
    return G;
  n = n || e;
  const o = P(n).filter(Ye);
  if (n = _(o), e = P(e), n instanceof Error)
    return n;
  let s = 0, l = 0;
  const f = t === void 0 || t === "*", d = f ? null : Bn(t + "");
  for (let c = 0; c < e.length; c++) {
    const i = e[c];
    if (f)
      l += n[c], s++;
    else {
      const a = [kt(i, jt)].concat(d);
      zn(a) && (l += n[c], s++);
    }
  }
  return l / s;
}
function cd() {
  const e = uo(...arguments), n = e.reduce((r, o) => r + o, 0) / e.length;
  return isNaN(n) ? 0 : n;
}
const pr = {};
pr.DIST = function(e, t, n, r, o, s) {
  return arguments.length < 4 || (o = o === void 0 ? 0 : o, s = s === void 0 ? 1 : s, e = w(e), t = w(t), n = w(n), o = w(o), s = w(s), O(e, t, n, o, s)) ? S : (e = (e - o) / (s - o), r ? U.beta.cdf(e, t, n) : U.beta.pdf(e, t, n));
};
pr.INV = (e, t, n, r, o) => (r = r === void 0 ? 0 : r, o = o === void 0 ? 1 : o, e = w(e), t = w(t), n = w(n), r = w(r), o = w(o), O(e, t, n, r, o) ? S : U.beta.inv(e, t, n) * (o - r) + r);
const Hn = {};
Hn.DIST = (e, t, n, r) => (e = w(e), t = w(t), n = w(n), r = w(r), O(e, t, n, r) ? S : r ? U.binomial.cdf(e, t, n) : U.binomial.pdf(e, t, n));
Hn.DIST.RANGE = (e, t, n, r) => {
  if (r = r === void 0 ? n : r, e = w(e), t = w(t), n = w(n), r = w(r), O(e, t, n, r))
    return S;
  let o = 0;
  for (let s = n; s <= r; s++)
    o += kn(e, s) * Math.pow(t, s) * Math.pow(1 - t, e - s);
  return o;
};
Hn.INV = (e, t, n) => {
  if (e = w(e), t = w(t), n = w(n), O(e, t, n))
    return S;
  let r = 0;
  for (; r <= e; ) {
    if (U.binomial.cdf(r, e, t) >= n)
      return r;
    r++;
  }
};
const It = {};
It.DIST = (e, t, n) => (e = w(e), t = w(t), O(e, t) ? S : n ? U.chisquare.cdf(e, t) : U.chisquare.pdf(e, t));
It.DIST.RT = (e, t) => !e | !t ? G : e < 1 || t > Math.pow(10, 10) ? T : typeof e != "number" || typeof t != "number" ? S : 1 - U.chisquare.cdf(e, t);
It.INV = (e, t) => (e = w(e), t = w(t), O(e, t) ? S : U.chisquare.inv(e, t));
It.INV.RT = (e, t) => !e | !t ? G : e < 0 || e > 1 || t < 1 || t > Math.pow(10, 10) ? T : typeof e != "number" || typeof t != "number" ? S : U.chisquare.inv(1 - e, t);
It.TEST = function(e, t) {
  if (arguments.length !== 2)
    return G;
  if (!(e instanceof Array) || !(t instanceof Array) || e.length !== t.length || e[0] && t[0] && e[0].length !== t[0].length)
    return S;
  const n = e.length;
  let r, o, s;
  for (o = 0; o < n; o++)
    e[o] instanceof Array || (r = e[o], e[o] = [], e[o].push(r)), t[o] instanceof Array || (r = t[o], t[o] = [], t[o].push(r));
  const l = e[0].length, f = l === 1 ? n - 1 : (n - 1) * (l - 1);
  let d = 0;
  const c = Math.PI;
  for (o = 0; o < n; o++)
    for (s = 0; s < l; s++)
      d += Math.pow(e[o][s] - t[o][s], 2) / t[o][s];
  function i(a, u) {
    let h = Math.exp(-0.5 * a);
    u % 2 === 1 && (h = h * Math.sqrt(2 * a / c));
    let g = u;
    for (; g >= 2; )
      h = h * a / g, g = g - 2;
    let m = h, p = u;
    for (; m > 1e-10 * h; )
      p = p + 2, m = m * a / p, h = h + m;
    return 1 - h;
  }
  return Math.round(i(d, f) * 1e6) / 1e6;
};
const Oi = {};
Oi.NORM = (e, t, n) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : U.normalci(1, e, t, n)[1] - 1);
Oi.T = (e, t, n) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : U.tci(1, e, t, n)[1] - 1);
function ud(e, t) {
  return e = _(P(e)), t = _(P(t)), O(e, t) ? S : U.corrcoeff(e, t);
}
function lr() {
  const e = P(arguments);
  return Bt(e).length;
}
function cr() {
  const e = P(arguments);
  return e.length - Za(e);
}
function Za() {
  const e = P(arguments);
  let t = 0, n;
  for (let r = 0; r < e.length; r++)
    n = e[r], (n == null || n === "") && t++;
  return t;
}
function fd(e, t) {
  if (e = P(e), t === void 0 || t === "*")
    return e.length;
  let r = 0;
  const o = Bn(t + "");
  for (let s = 0; s < e.length; s++) {
    const l = e[s], f = [kt(l, jt)].concat(o);
    zn(f) && r++;
  }
  return r;
}
function dd() {
  const e = lo(arguments), t = new Array(P(e[0]).length);
  for (let r = 0; r < t.length; r++)
    t[r] = !0;
  for (let r = 0; r < e.length; r += 2) {
    const o = P(e[r]), s = e[r + 1];
    if (!(s === void 0 || s === "*")) {
      const f = Bn(s + "");
      for (let d = 0; d < o.length; d++) {
        const c = o[d], i = [kt(c, jt)].concat(f);
        t[d] = t[d] && zn(i);
      }
    }
  }
  let n = 0;
  for (let r = 0; r < t.length; r++)
    t[r] && n++;
  return n;
}
const Vn = {};
Vn.P = (e, t) => {
  if (e = _(P(e)), t = _(P(t)), O(e, t))
    return S;
  const n = U.mean(e), r = U.mean(t);
  let o = 0;
  const s = e.length;
  for (let l = 0; l < s; l++)
    o += (e[l] - n) * (t[l] - r);
  return o / s;
};
Vn.S = (e, t) => (e = _(P(e)), t = _(P(t)), O(e, t) ? S : U.covariance(e, t));
function hd() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  const t = U.mean(e);
  let n = 0;
  for (let r = 0; r < e.length; r++)
    n += Math.pow(e[r] - t, 2);
  return n;
}
const ki = {};
ki.DIST = (e, t, n) => (e = w(e), t = w(t), O(e, t) ? S : n ? U.exponential.cdf(e, t) : U.exponential.pdf(e, t));
const Mt = {};
Mt.DIST = (e, t, n, r) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : r ? U.centralF.cdf(e, t, n) : U.centralF.pdf(e, t, n));
Mt.DIST.RT = function(e, t, n) {
  return arguments.length !== 3 ? G : e < 0 || t < 1 || n < 1 ? T : typeof e != "number" || typeof t != "number" || typeof n != "number" ? S : 1 - U.centralF.cdf(e, t, n);
};
Mt.INV = (e, t, n) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : e <= 0 || e > 1 ? T : U.centralF.inv(e, t, n));
Mt.INV.RT = function(e, t, n) {
  return arguments.length !== 3 ? G : e < 0 || e > 1 || t < 1 || t > Math.pow(10, 10) || n < 1 || n > Math.pow(10, 10) ? T : typeof e != "number" || typeof t != "number" || typeof n != "number" ? S : U.centralF.inv(1 - e, t, n);
};
Mt.TEST = (e, t) => {
  if (!e || !t || !(e instanceof Array) || !(t instanceof Array))
    return G;
  if (e.length < 2 || t.length < 2)
    return Ue;
  const n = (f, d) => {
    let c = 0;
    for (let i = 0; i < f.length; i++)
      c += Math.pow(f[i] - d, 2);
    return c;
  }, r = Pt(e) / e.length, o = Pt(t) / t.length, s = n(e, r) / (e.length - 1), l = n(t, o) / (t.length - 1);
  return s / l;
};
function md(e) {
  return e = w(e), e instanceof Error ? e : Math.log((1 + e) / (1 - e)) / 2;
}
function gd(e) {
  if (e = w(e), e instanceof Error)
    return e;
  const t = Math.exp(2 * e);
  return (t - 1) / (t + 1);
}
function el(e, t, n) {
  if (e = w(e), t = _(P(t)), n = _(P(n)), O(e, t, n))
    return S;
  const r = U.mean(n), o = U.mean(t), s = n.length;
  let l = 0, f = 0;
  for (let i = 0; i < s; i++)
    l += (n[i] - r) * (t[i] - o), f += Math.pow(n[i] - r, 2);
  const d = l / f;
  return o - d * r + d * e;
}
function pd(e, t) {
  if (e = _(P(e)), t = _(P(t)), O(e, t))
    return S;
  const n = e.length, r = t.length, o = [];
  for (let s = 0; s <= r; s++) {
    o[s] = 0;
    for (let l = 0; l < n; l++)
      s === 0 ? e[l] <= t[0] && (o[0] += 1) : s < r ? e[l] > t[s - 1] && e[l] <= t[s] && (o[s] += 1) : s === r && e[l] > t[r - 1] && (o[r] += 1);
  }
  return o;
}
function vr(e) {
  return e = w(e), e instanceof Error ? e : e === 0 || parseInt(e, 10) === e && e < 0 ? T : U.gammafn(e);
}
vr.DIST = function(e, t, n, r) {
  return arguments.length !== 4 ? G : e < 0 || t <= 0 || n <= 0 || typeof e != "number" || typeof t != "number" || typeof n != "number" ? S : r ? U.gamma.cdf(e, t, n, !0) : U.gamma.pdf(e, t, n, !1);
};
vr.INV = function(e, t, n) {
  return arguments.length !== 3 ? G : e < 0 || e > 1 || t <= 0 || n <= 0 ? T : typeof e != "number" || typeof t != "number" || typeof n != "number" ? S : U.gamma.inv(e, t, n);
};
function Li(e) {
  return e = w(e), e instanceof Error ? e : U.gammaln(e);
}
Li.PRECISE = function(e) {
  return arguments.length !== 1 ? G : e <= 0 ? T : typeof e != "number" ? S : U.gammaln(e);
};
function vd(e) {
  return e = w(e), e instanceof Error ? e : U.normal.cdf(e, 0, 1) - 0.5;
}
function Nd() {
  const e = _(P(arguments));
  return e instanceof Error ? e : U.geomean(e);
}
function Cd(e, t, n, r) {
  if (e = _(P(e)), e instanceof Error)
    return e;
  let o;
  if (t === void 0)
    for (t = [], o = 1; o <= e.length; o++)
      t.push(o);
  if (n === void 0 && (n = t), t = _(P(t)), n = _(P(n)), O(t, n))
    return S;
  r === void 0 && (r = !0);
  const s = e.length;
  let l = 0, f = 0, d = 0, c = 0;
  for (o = 0; o < s; o++) {
    const h = t[o], g = Math.log(e[o]);
    l += h, f += g, d += h * g, c += h * h;
  }
  l /= s, f /= s, d /= s, c /= s;
  let i, a;
  r ? (i = (d - l * f) / (c - l * l), a = f - i * l) : (i = d / c, a = 0);
  const u = [];
  for (o = 0; o < n.length; o++)
    u.push(Math.exp(a + i * n[o]));
  return u;
}
function Ed() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  const t = e.length;
  let n = 0;
  for (let r = 0; r < t; r++)
    n += 1 / e[r];
  return t / n;
}
const Fi = {};
Fi.DIST = (e, t, n, r, o) => {
  if (e = w(e), t = w(t), n = w(n), r = w(r), O(e, t, n, r))
    return S;
  function s(f, d, c, i) {
    return kn(c, f) * kn(i - c, d - f) / kn(i, d);
  }
  function l(f, d, c, i) {
    let a = 0;
    for (let u = 0; u <= f; u++)
      a += s(u, d, c, i);
    return a;
  }
  return o ? l(e, t, n, r) : s(e, t, n, r);
};
function bd(e, t) {
  return e = _(e), t = _(t), O(e, t) ? S : e.length !== t.length ? G : el(0, e, t);
}
function wd() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  const t = U.mean(e), n = e.length;
  let r = 0;
  for (let o = 0; o < n; o++)
    r += Math.pow(e[o] - t, 4);
  return r = r / Math.pow(U.stdev(e, !0), 4), n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * r - 3 * (n - 1) * (n - 1) / ((n - 2) * (n - 3));
}
function tl(e, t) {
  const n = Q.apply(void 0, e);
  return n || (O(t) ? t : (e = Bt(P(e)), t = w(t), t < 0 || e.length < t ? S : e.sort((r, o) => o - r)[t - 1]));
}
function Pi(e, t) {
  if (e = _(P(e)), t = _(P(t)), O(e, t))
    return S;
  const n = U.mean(e), r = U.mean(t), o = t.length;
  let s = 0, l = 0;
  for (let c = 0; c < o; c++)
    s += (t[c] - r) * (e[c] - n), l += Math.pow(t[c] - r, 2);
  const f = s / l, d = n - f * r;
  return [f, d];
}
function Id(e, t) {
  if (e = _(P(e)), t = _(P(t)), O(e, t) || e.length !== t.length)
    return S;
  for (let r = 0; r < e.length; r++)
    e[r] = Math.log(e[r]);
  const n = Pi(e, t);
  return n[0] = Math.round(Math.exp(n[0]) * 1e6) / 1e6, n[1] = Math.round(Math.exp(n[1]) * 1e6) / 1e6, n;
}
const Gn = {};
Gn.DIST = (e, t, n, r) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : r ? U.lognormal.cdf(e, t, n) : U.lognormal.pdf(e, t, n));
Gn.INV = (e, t, n) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : U.lognormal.inv(e, t, n));
function Gr() {
  const e = P(arguments), t = Q.apply(void 0, e);
  if (t)
    return t;
  const n = Bt(e);
  return n.length === 0 ? 0 : Math.max.apply(Math, n);
}
function Md() {
  const e = P(arguments), t = Q.apply(void 0, e);
  if (t)
    return t;
  let n = Ai(e);
  return n = n.map((r) => r ?? 0), n.length === 0 ? 0 : Math.max.apply(Math, n);
}
function yd() {
  const e = uo(...arguments);
  return e.length === 0 ? 0 : Math.max.apply(Math, e);
}
function nl() {
  const e = P(arguments), t = Q.apply(void 0, e);
  if (t)
    return t;
  const n = Ai(e);
  let r = U.median(n);
  return isNaN(r) && (r = T), r;
}
function $r() {
  const e = P(arguments), t = Q.apply(void 0, e);
  if (t)
    return t;
  const n = Bt(e);
  return n.length === 0 ? 0 : Math.min.apply(Math, n);
}
function Sd() {
  const e = P(arguments), t = Q.apply(void 0, e);
  if (t)
    return t;
  let n = Ai(e);
  return n = n.map((r) => r ?? 0), n.length === 0 ? 0 : Math.min.apply(Math, n);
}
function Rd() {
  const e = uo(...arguments);
  return e.length === 0 ? 0 : Math.min.apply(Math, e);
}
const dn = {};
dn.MULT = function() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  const t = e.length, n = {};
  let r = [], o = 0, s;
  for (let l = 0; l < t; l++)
    s = e[l], n[s] = n[s] ? n[s] + 1 : 1, n[s] > o && (o = n[s], r = []), n[s] === o && (r[r.length] = s);
  return r;
};
dn.SNGL = function() {
  const e = _(P(arguments));
  return e instanceof Error ? e : dn.MULT(e).sort((t, n) => t - n)[0];
};
const Ui = {};
Ui.DIST = (e, t, n, r) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : r ? U.negbin.cdf(e, t, n) : U.negbin.pdf(e, t, n));
const yt = {};
yt.DIST = (e, t, n, r) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : n <= 0 ? T : r ? U.normal.cdf(e, t, n) : U.normal.pdf(e, t, n));
yt.INV = (e, t, n) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : U.normal.inv(e, t, n));
yt.S = {};
yt.S.DIST = (e, t) => (e = w(e), e instanceof Error ? S : t ? U.normal.cdf(e, 0, 1) : U.normal.pdf(e, 0, 1));
yt.S.INV = (e) => (e = w(e), e instanceof Error ? S : U.normal.inv(e, 0, 1));
function rl(e, t) {
  if (t = _(P(t)), e = _(P(e)), O(t, e))
    return S;
  const n = U.mean(e), r = U.mean(t), o = e.length;
  let s = 0, l = 0, f = 0;
  for (let d = 0; d < o; d++)
    s += (e[d] - n) * (t[d] - r), l += Math.pow(e[d] - n, 2), f += Math.pow(t[d] - r, 2);
  return s / Math.sqrt(l * f);
}
const st = {};
st.EXC = (e, t) => {
  if (e = _(P(e)), t = w(t), O(e, t))
    return S;
  e = e.sort((s, l) => s - l);
  const n = e.length;
  if (t < 1 / (n + 1) || t > 1 - 1 / (n + 1))
    return T;
  const r = t * (n + 1) - 1, o = Math.floor(r);
  return Va(r === o ? e[r] : e[o] + (r - o) * (e[o + 1] - e[o]));
};
st.INC = (e, t) => {
  if (e = _(P(e)), t = w(t), O(e, t))
    return S;
  e = e.sort((s, l) => s - l);
  const n = e.length, r = t * (n - 1), o = Math.floor(r);
  return Va(r === o ? e[r] : e[o] + (r - o) * (e[o + 1] - e[o]));
};
const Nr = {};
Nr.EXC = (e, t, n) => {
  if (n = n === void 0 ? 3 : n, e = _(P(e)), t = w(t), n = w(n), O(e, t, n))
    return S;
  e = e.sort((i, a) => i - a);
  const r = xi.apply(null, e), o = e.length, s = r.length, l = Math.pow(10, n);
  let f = 0, d = !1, c = 0;
  for (; !d && c < s; )
    t === r[c] ? (f = (e.indexOf(r[c]) + 1) / (o + 1), d = !0) : t >= r[c] && (t < r[c + 1] || c === s - 1) && (f = (e.indexOf(r[c]) + 1 + (t - r[c]) / (r[c + 1] - r[c])) / (o + 1), d = !0), c++;
  return Math.floor(f * l) / l;
};
Nr.INC = (e, t, n) => {
  if (n = n === void 0 ? 3 : n, e = _(P(e)), t = w(t), n = w(n), O(e, t, n))
    return S;
  e = e.sort((i, a) => i - a);
  const r = xi.apply(null, e), o = e.length, s = r.length, l = Math.pow(10, n);
  let f = 0, d = !1, c = 0;
  for (; !d && c < s; )
    t === r[c] ? (f = e.indexOf(r[c]) / (o - 1), d = !0) : t >= r[c] && (t < r[c + 1] || c === s - 1) && (f = (e.indexOf(r[c]) + (t - r[c]) / (r[c + 1] - r[c])) / (o - 1), d = !0), c++;
  return Math.floor(f * l) / l;
};
function Ad(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : Lt(e) / Lt(e - t);
}
function Td(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : Math.pow(e, t);
}
function xd(e) {
  return e = w(e), e instanceof Error ? S : Math.exp(-0.5 * e * e) / sd;
}
const _i = {};
_i.DIST = (e, t, n) => (e = w(e), t = w(t), O(e, t) ? S : n ? U.poisson.cdf(e, t) : U.poisson.pdf(e, t));
function Dd(e, t, n, r) {
  if (n === void 0)
    return 0;
  if (r = r === void 0 ? n : r, e = _(P(e)), t = _(P(t)), n = w(n), r = w(r), O(e, t, n, r))
    return S;
  if (n === r)
    return e.indexOf(n) >= 0 ? t[e.indexOf(n)] : 0;
  const o = e.sort((f, d) => f - d), s = o.length;
  let l = 0;
  for (let f = 0; f < s; f++)
    o[f] >= n && o[f] <= r && (l += t[e.indexOf(o[f])]);
  return l;
}
const hn = {};
hn.EXC = (e, t) => {
  if (e = _(Bt(P(e))), t = w(t), O(e, t))
    return S;
  switch (t) {
    case 1:
      return st.EXC(e, 0.25);
    case 2:
      return st.EXC(e, 0.5);
    case 3:
      return st.EXC(e, 0.75);
    default:
      return T;
  }
};
hn.INC = (e, t) => {
  if (e = _(Bt(P(e))), t = w(t), O(e, t))
    return S;
  switch (t) {
    case 1:
      return st.INC(e, 0.25);
    case 2:
      return st.INC(e, 0.5);
    case 3:
      return st.INC(e, 0.75);
    default:
      return T;
  }
};
const Cr = {};
Cr.AVG = (e, t, n) => {
  if (e = w(e), t = _(P(t)), O(e, t))
    return S;
  t = P(t), n = n || !1;
  const r = n ? (l, f) => l - f : (l, f) => f - l;
  t = t.sort(r);
  const o = t.length;
  let s = 0;
  for (let l = 0; l < o; l++)
    t[l] === e && s++;
  return s > 1 ? (2 * t.indexOf(e) + s + 1) / 2 : t.indexOf(e) + 1;
};
Cr.EQ = (e, t, n) => {
  if (e = w(e), t = _(P(t)), O(e, t))
    return S;
  n = n || !1;
  const r = n ? (o, s) => o - s : (o, s) => s - o;
  return t = t.sort(r), t.indexOf(e) + 1;
};
function Od(e, t) {
  if (arguments.length !== 2)
    return G;
  if (t < 0)
    return T;
  if (!(e instanceof Array) || typeof t != "number")
    return S;
  if (e.length !== 0)
    return U.row(e, t);
}
function kd(e, t) {
  return e = _(P(e)), t = _(P(t)), O(e, t) ? S : Math.pow(rl(e, t), 2);
}
function Bi() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  const t = U.mean(e), n = e.length;
  let r = 0;
  for (let o = 0; o < n; o++)
    r += Math.pow(e[o] - t, 3);
  return n * r / ((n - 1) * (n - 2) * Math.pow(U.stdev(e, !0), 3));
}
Bi.P = function() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  const t = U.mean(e), n = e.length;
  let r = 0, o = 0;
  for (let s = 0; s < n; s++)
    o += Math.pow(e[s] - t, 3), r += Math.pow(e[s] - t, 2);
  return o = o / n, r = r / n, o / Math.pow(r, 3 / 2);
};
function Ld(e, t) {
  if (e = _(P(e)), t = _(P(t)), O(e, t))
    return S;
  const n = U.mean(t), r = U.mean(e), o = t.length;
  let s = 0, l = 0;
  for (let f = 0; f < o; f++)
    s += (t[f] - n) * (e[f] - r), l += Math.pow(t[f] - n, 2);
  return s / l;
}
function ol(e, t) {
  return e = _(P(e)), t = w(t), O(e, t) ? e : e.sort((n, r) => n - r)[t - 1];
}
function Fd(e, t, n) {
  return e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : (e - t) / n;
}
const Ze = {};
Ze.P = function() {
  const e = qe.P.apply(this, arguments);
  let t = Math.sqrt(e);
  return isNaN(t) && (t = T), t;
};
Ze.S = function() {
  const e = qe.S.apply(this, arguments);
  return Math.sqrt(e);
};
function Pd() {
  const e = il.apply(this, arguments);
  return Math.sqrt(e);
}
function Ud() {
  const e = sl.apply(this, arguments);
  let t = Math.sqrt(e);
  return isNaN(t) && (t = T), t;
}
function _d(e, t) {
  if (e = _(P(e)), t = _(P(t)), O(e, t))
    return S;
  const n = U.mean(t), r = U.mean(e), o = t.length;
  let s = 0, l = 0, f = 0;
  for (let d = 0; d < o; d++)
    s += Math.pow(e[d] - r, 2), l += (t[d] - n) * (e[d] - r), f += Math.pow(t[d] - n, 2);
  return Math.sqrt((s - l * l / f) / (o - 2));
}
et.DIST = (e, t, n) => n !== 1 && n !== 2 ? T : n === 1 ? et.DIST.RT(e, t) : et.DIST["2T"](e, t);
et.DIST["2T"] = function(e, t) {
  return arguments.length !== 2 ? G : e < 0 || t < 1 ? T : typeof e != "number" || typeof t != "number" ? S : (1 - U.studentt.cdf(e, t)) * 2;
};
et.DIST.RT = function(e, t) {
  return arguments.length !== 2 ? G : e < 0 || t < 1 ? T : typeof e != "number" || typeof t != "number" ? S : 1 - U.studentt.cdf(e, t);
};
et.INV = (e, t) => (e = w(e), t = w(t), O(e, t) ? S : U.studentt.inv(e, t));
et.INV["2T"] = (e, t) => (e = w(e), t = w(t), e <= 0 || e > 1 || t < 1 ? T : O(e, t) ? S : Math.abs(U.studentt.inv(e / 2, t)));
et.TEST = (e, t) => {
  if (e = _(P(e)), t = _(P(t)), O(e, t))
    return S;
  const n = U.mean(e), r = U.mean(t);
  let o = 0, s = 0, l;
  for (l = 0; l < e.length; l++)
    o += Math.pow(e[l] - n, 2);
  for (l = 0; l < t.length; l++)
    s += Math.pow(t[l] - r, 2);
  o = o / (e.length - 1), s = s / (t.length - 1);
  const f = Math.abs(n - r) / Math.sqrt(o / e.length + s / t.length);
  return et.DIST["2T"](f, e.length + t.length - 2);
};
function Bd(e, t, n) {
  if (e = _(P(e)), t = _(P(t)), n = _(P(n)), O(e, t, n))
    return S;
  const r = Pi(e, t), o = r[0], s = r[1], l = [];
  return n.forEach((f) => {
    l.push(o * f + s);
  }), l;
}
function zd(e, t) {
  if (e = _(P(e)), t = w(t), O(e, t))
    return S;
  const n = mn(e.length * t, 2) / 2;
  return U.mean(
    ff(
      Ne(
        e.sort((r, o) => r - o),
        n
      ),
      n
    )
  );
}
const qe = {};
qe.P = function() {
  const e = Bt(P(arguments)), t = e.length;
  let n = 0;
  const r = fn(e);
  let o;
  for (let s = 0; s < t; s++)
    n += Math.pow(e[s] - r, 2);
  return o = n / t, isNaN(o) && (o = T), o;
};
qe.S = function() {
  const e = Bt(P(arguments)), t = e.length;
  let n = 0;
  const r = fn(e);
  for (let o = 0; o < t; o++)
    n += Math.pow(e[o] - r, 2);
  return n / (t - 1);
};
function il() {
  const e = P(arguments), t = e.length;
  let n = 0, r = 0;
  const o = Di(e);
  for (let s = 0; s < t; s++) {
    const l = e[s];
    typeof l == "number" ? n += Math.pow(l - o, 2) : l === !0 ? n += Math.pow(1 - o, 2) : n += Math.pow(0 - o, 2), l !== null && r++;
  }
  return n / (r - 1);
}
function sl() {
  const e = P(arguments), t = e.length;
  let n = 0, r = 0;
  const o = Di(e);
  let s;
  for (let l = 0; l < t; l++) {
    const f = e[l];
    typeof f == "number" ? n += Math.pow(f - o, 2) : f === !0 ? n += Math.pow(1 - o, 2) : n += Math.pow(0 - o, 2), f !== null && r++;
  }
  return s = n / r, isNaN(s) && (s = T), s;
}
const zi = {};
zi.DIST = (e, t, n, r) => (e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : r ? 1 - Math.exp(-Math.pow(e / n, t)) : Math.pow(e, t - 1) * Math.exp(-Math.pow(e / n, t)) * t / Math.pow(n, t));
const Hi = {};
Hi.TEST = (e, t, n) => {
  if (e = _(P(e)), t = w(t), O(e, t))
    return S;
  n = n || Ze.S(e);
  const r = e.length;
  return 1 - yt.S.DIST((fn(e) - t) / (n / Math.sqrt(r)), !0);
};
function Hd(e) {
  return e = w(e), e instanceof Error ? e : Math.abs(e);
}
function Vd(e) {
  if (e = w(e), e instanceof Error)
    return e;
  let t = Math.acos(e);
  return isNaN(t) && (t = T), t;
}
function Gd(e) {
  if (e = w(e), e instanceof Error)
    return e;
  let t = Math.log(e + Math.sqrt(e * e - 1));
  return isNaN(t) && (t = T), t;
}
function $d(e) {
  return e = w(e), e instanceof Error ? e : Math.atan(1 / e);
}
function qd(e) {
  if (e = w(e), e instanceof Error)
    return e;
  let t = 0.5 * Math.log((e + 1) / (e - 1));
  return isNaN(t) && (t = T), t;
}
function Wd(e, t, n, r) {
  if (e = w(e), t = w(e), O(e, t))
    return S;
  switch (e) {
    case 1:
      return fn(n);
    case 2:
      return lr(n);
    case 3:
      return cr(n);
    case 4:
      return Gr(n);
    case 5:
      return $r(n);
    case 6:
      return qr(n);
    case 7:
      return Ze.S(n);
    case 8:
      return Ze.P(n);
    case 9:
      return Pt(n);
    case 10:
      return qe.S(n);
    case 11:
      return qe.P(n);
    case 12:
      return nl(n);
    case 13:
      return dn.SNGL(n);
    case 14:
      return tl(n, r);
    case 15:
      return ol(n, r);
    case 16:
      return st.INC(n, r);
    case 17:
      return hn.INC(n, r);
    case 18:
      return st.EXC(n, r);
    case 19:
      return hn.EXC(n, r);
  }
}
function Yd(e) {
  if (e == null)
    return 0;
  if (e instanceof Error)
    return e;
  if (!/^M*(?:D?C{0,3}|C[MD])(?:L?X{0,3}|X[CL])(?:V?I{0,3}|I[XV])$/.test(e))
    return S;
  let t = 0;
  return e.replace(/[MDLV]|C[MD]?|X[CL]?|I[XV]?/g, (n) => {
    t += {
      M: 1e3,
      CM: 900,
      D: 500,
      CD: 400,
      C: 100,
      XC: 90,
      L: 50,
      XL: 40,
      X: 10,
      IX: 9,
      V: 5,
      IV: 4,
      I: 1
    }[n];
  }), t;
}
function Xd(e) {
  if (e = w(e), e instanceof Error)
    return e;
  let t = Math.asin(e);
  return isNaN(t) && (t = T), t;
}
function Kd(e) {
  return e = w(e), e instanceof Error ? e : Math.log(e + Math.sqrt(e * e + 1));
}
function Jd(e) {
  return e = w(e), e instanceof Error ? e : Math.atan(e);
}
function Qd(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || Math.atan2(e, t);
}
function jd(e) {
  if (e = w(e), e instanceof Error)
    return e;
  let t = Math.log((1 + e) / (1 - e)) / 2;
  return isNaN(t) && (t = T), t;
}
function Zd(e, t, n) {
  e = w(e), t = w(t), n = w(n);
  const r = Q(e, t, n);
  if (r)
    return r;
  if (t === 0)
    return T;
  const o = e.toString(t);
  return new Array(Math.max(n + 1 - o.length, 0)).join("0") + o;
}
function Jt(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || (t === 0 ? 0 : e > 0 && t < 0 ? T : Math.ceil(e / t) * t);
}
Jt.MATH = (e, t, n = 0) => {
  t === void 0 && (t = e > 0 ? 1 : -1), e = w(e), t = w(t), n = w(n);
  const r = Q(e, t, n);
  return r || (t === 0 ? 0 : (t = Math.abs(t), n === 0 || e > 0 ? Math.ceil(e / t) * t : Math.floor(e / t) * t));
};
Jt.PRECISE = (e, t) => Jt.MATH(e, t);
function kn(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || (e < t ? T : Lt(e) / (Lt(t) * Lt(e - t)));
}
function eh(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || (e < t ? T : e === 0 && t === 0 ? 1 : kn(e + t - 1, e - 1));
}
function th(e) {
  return e = w(e), e instanceof Error ? e : Math.cos(e);
}
function nh(e) {
  return e = w(e), e instanceof Error ? e : (Math.exp(e) + Math.exp(-e)) / 2;
}
function rh(e) {
  return e = w(e), e instanceof Error ? e : e === 0 ? Ue : 1 / Math.tan(e);
}
function oh(e) {
  if (e = w(e), e instanceof Error)
    return e;
  if (e === 0)
    return Ue;
  const t = Math.exp(2 * e);
  return (t + 1) / (t - 1);
}
function ih(e) {
  return e = w(e), e instanceof Error ? e : e === 0 ? Ue : 1 / Math.sin(e);
}
function sh(e) {
  return e = w(e), e instanceof Error ? e : e === 0 ? Ue : 2 / (Math.exp(e) - Math.exp(-e));
}
function ah(e, t) {
  if (arguments.length < 2)
    return G;
  e = e || "0", t = w(t);
  const n = Q(e, t);
  if (n)
    return n;
  if (t === 0)
    return T;
  const r = parseInt(e, t);
  return isNaN(r) ? T : r;
}
function lh(e) {
  return e = w(e), e instanceof Error ? e : e * 180 / Math.PI;
}
function ch(e) {
  return e = w(e), e instanceof Error ? e : Jt.MATH(e, -2, -1);
}
function uh(e) {
  return arguments.length < 1 ? G : arguments.length > 1 ? Ri : (e = w(e), e instanceof Error || (e = Math.exp(e)), e);
}
const Ir = [];
function Lt(e) {
  if (e = w(e), e instanceof Error)
    return e;
  const t = Math.floor(e);
  return t === 0 || t === 1 ? 1 : (Ir[t] > 0 || (Ir[t] = Lt(t - 1) * t), Ir[t]);
}
function al(e) {
  if (e = w(e), e instanceof Error)
    return e;
  const t = Math.floor(e);
  return t <= 0 ? 1 : t * al(t - 2);
}
function mn(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || (t ? e > 0 && t < 0 ? T : Math.floor(e / t) * t : Ue);
}
mn.MATH = (e, t = 1, n = 0) => {
  e = w(e), t = w(t), n = w(n);
  const r = Q(e, t, n);
  return r || (t === 0 ? 0 : (t = Math.abs(t), n === 0 || e > 0 ? Math.floor(e / t) * t : Math.ceil(e / t) * t));
};
mn.PRECISE = (e, t) => mn.MATH(e, t);
function fh() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  const t = e.length, n = e[0];
  let r = n < 0 ? -n : n;
  for (let o = 1; o < t; o++) {
    const s = e[o];
    let l = s < 0 ? -s : s;
    for (; r && l; )
      r > l ? r %= l : l %= r;
    r += l;
  }
  return r;
}
function dh(e) {
  return e = w(e), e instanceof Error ? e : Math.floor(e);
}
const hh = {
  CEILING: Jt
};
function mh() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  for (var t, n, r, o, s = 1; (r = e.pop()) !== void 0; ) {
    if (r === 0)
      return 0;
    for (; r > 1; ) {
      if (r % 2) {
        for (t = 3, n = Math.floor(Math.sqrt(r)); t <= n && r % t; t += 2)
          ;
        o = t <= n ? t : r;
      } else
        o = 2;
      for (r /= o, s *= o, t = e.length; t; e[--t] % o === 0 && (e[t] /= o) === 1 && e.splice(t, 1))
        ;
    }
  }
  return s;
}
function gh(e) {
  return e = w(e), e instanceof Error ? e : e === 0 ? T : Math.log(e);
}
function ph(e, t) {
  e = w(e), t = t ? w(t) : 10;
  const n = Q(e, t);
  return n || (e === 0 || t === 0 ? T : Math.log(e) / Math.log(t));
}
function vh(e) {
  return e = w(e), e instanceof Error ? e : e === 0 ? T : Math.log(e) / Math.log(10);
}
function Nh(e, t) {
  return (
    //Arguments are not arrays
    !Array.isArray(e) || !Array.isArray(t) || // There are empty arrays
    e.some((r) => !r.length) || t.some((r) => !r.length) || // Not all array elements are numbers
    ri(e).some((r) => typeof r != "number") || ri(t).some((r) => typeof r != "number") || // Number of columns in array1 is different from the number of rows in array2
    e[0].length !== t.length ? S : Array(e.length).fill(0).map(() => Array(t[0].length).fill(0)).map((r, o) => r.map((s, l) => e[o].reduce((f, d, c) => f + d * t[c][l], 0)))
  );
}
function Ch(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  if (n)
    return n;
  if (t === 0)
    return Ue;
  let r = Math.abs(e % t);
  return r = e < 0 ? t - r : r, t > 0 ? r : -r;
}
function Eh(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || (e * t === 0 ? 0 : e * t < 0 ? T : Math.round(e / t) * t);
}
function bh() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  let t = 0, n = 1;
  for (let r = 0; r < e.length; r++)
    t += e[r], n *= Lt(e[r]);
  return Lt(t) / n;
}
function wh(e) {
  return arguments.length > 1 ? G : (e = parseInt(e), !e || e <= 0 ? S : Array(e).fill(0).map(() => Array(e).fill(0)).map((t, n) => (t[n] = 1, t)));
}
function Ih(e) {
  if (e = w(e), e instanceof Error)
    return e;
  let t = Math.ceil(Math.abs(e));
  return t = t & 1 ? t : t + 1, e >= 0 ? t : -t;
}
function Mh() {
  return Math.PI;
}
function ll(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  if (n)
    return n;
  if (e === 0 && t === 0)
    return T;
  const r = Math.pow(e, t);
  return isNaN(r) ? T : r;
}
function qr() {
  const t = P(arguments).filter((o) => o != null);
  if (t.length === 0)
    return 0;
  const n = _(t);
  if (n instanceof Error)
    return n;
  let r = 1;
  for (let o = 0; o < n.length; o++)
    r *= n[o];
  return r;
}
function yh(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || parseInt(e / t, 10);
}
function Sh(e) {
  return e = w(e), e instanceof Error ? e : e * Math.PI / 180;
}
function Rh() {
  return Math.random();
}
function Ah(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || e + Math.ceil((t - e + 1) * Math.random()) - 1;
}
function Th(e) {
  if (e = w(e), e instanceof Error)
    return e;
  const t = String(e).split(""), n = [
    "",
    "C",
    "CC",
    "CCC",
    "CD",
    "D",
    "DC",
    "DCC",
    "DCCC",
    "CM",
    "",
    "X",
    "XX",
    "XXX",
    "XL",
    "L",
    "LX",
    "LXX",
    "LXXX",
    "XC",
    "",
    "I",
    "II",
    "III",
    "IV",
    "V",
    "VI",
    "VII",
    "VIII",
    "IX"
  ];
  let r = "", o = 3;
  for (; o--; )
    r = (n[+t.pop() + o * 10] || "") + r;
  return new Array(+t.join("") + 1).join("M") + r;
}
function Vi(e, t, n) {
  e = w(e), t = w(t);
  const r = Q(e, t);
  if (r)
    return r;
  const o = e >= 0 ? 1 : -1;
  let l = (Math.abs(e) + "e" + t).split("e");
  return l = (n(l[0] + "e" + l[1]) + "e" + -t).split("e"), +(l[0] + "e" + l[1]) * o;
}
function cl(e, t) {
  return Vi(e, t, Math.round);
}
function xh(e, t) {
  return Vi(e, t, Math.floor);
}
function Dh(e, t) {
  return Vi(e, t, Math.ceil);
}
function Oh(e) {
  return e = w(e), e instanceof Error ? e : 1 / Math.cos(e);
}
function kh(e) {
  return e = w(e), e instanceof Error ? e : 2 / (Math.exp(e) + Math.exp(-e));
}
function Lh(e, t, n, r) {
  if (e = w(e), t = w(t), n = w(n), r = _(r), O(e, t, n, r))
    return S;
  let o = r[0] * Math.pow(e, t);
  for (let s = 1; s < r.length; s++)
    o += r[s] * Math.pow(e, t + s * n);
  return o;
}
function Fh(e) {
  return e = w(e), e instanceof Error ? e : e < 0 ? -1 : e === 0 ? 0 : 1;
}
function Ph(e) {
  return e = w(e), e instanceof Error ? e : Math.sin(e);
}
function Uh(e) {
  return e = w(e), e instanceof Error ? e : (Math.exp(e) - Math.exp(-e)) / 2;
}
function _h(e) {
  return e = w(e), e instanceof Error ? e : e < 0 ? T : Math.sqrt(e);
}
function Bh(e) {
  return e = w(e), e instanceof Error ? e : Math.sqrt(e * Math.PI);
}
function zh(e, t) {
  if (e = w(e), e instanceof Error)
    return e;
  switch (e) {
    case 1:
      return fn(t);
    case 2:
      return lr(t);
    case 3:
      return cr(t);
    case 4:
      return Gr(t);
    case 5:
      return $r(t);
    case 6:
      return qr(t);
    case 7:
      return Ze.S(t);
    case 8:
      return Ze.P(t);
    case 9:
      return Pt(t);
    case 10:
      return qe.S(t);
    case 11:
      return qe.P(t);
    // no hidden values for us
    case 101:
      return fn(t);
    case 102:
      return lr(t);
    case 103:
      return cr(t);
    case 104:
      return Gr(t);
    case 105:
      return $r(t);
    case 106:
      return qr(t);
    case 107:
      return Ze.S(t);
    case 108:
      return Ze.P(t);
    case 109:
      return Pt(t);
    case 110:
      return qe.S(t);
    case 111:
      return qe.P(t);
  }
}
function Pt() {
  let e = 0;
  return He(lo(arguments), (t) => {
    if (e instanceof Error)
      return !1;
    if (t instanceof Error)
      e = t;
    else if (typeof t == "number")
      e += t;
    else if (typeof t == "string") {
      const n = parseFloat(t);
      !isNaN(n) && (e += n);
    } else if (Array.isArray(t)) {
      const n = Pt.apply(null, t);
      n instanceof Error ? e = n : e += n;
    }
  }), e;
}
function Hh(e, t, n) {
  if (e = P(e), n = n ? P(n) : e, e instanceof Error)
    return e;
  if (t == null || t instanceof Error)
    return 0;
  let r = 0;
  const o = t === "*", s = o ? null : Bn(t + "");
  for (let l = 0; l < e.length; l++) {
    const f = e[l], d = n[l];
    if (o)
      r += f;
    else {
      const c = [kt(f, jt)].concat(s);
      r += zn(c) ? d : 0;
    }
  }
  return r;
}
function Vh() {
  const e = uo(...arguments);
  return Pt(e);
}
function Gh() {
  if (!arguments || arguments.length === 0)
    return S;
  const e = arguments.length + 1;
  let t = 0, n, r, o, s;
  for (let l = 0; l < arguments[0].length; l++)
    if (arguments[0][l] instanceof Array)
      for (let f = 0; f < arguments[0][l].length; f++) {
        for (n = 1, r = 1; r < e; r++) {
          const d = arguments[r - 1][l][f];
          if (d instanceof Error)
            return d;
          if (s = w(d), s instanceof Error)
            return s;
          n *= s;
        }
        t += n;
      }
    else {
      for (n = 1, r = 1; r < e; r++) {
        const f = arguments[r - 1][l];
        if (f instanceof Error)
          return f;
        if (o = w(f), o instanceof Error)
          return o;
        n *= o;
      }
      t += n;
    }
  return t;
}
function $h() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  let t = 0;
  const n = e.length;
  for (let r = 0; r < n; r++)
    t += ho(e[r]) ? e[r] * e[r] : 0;
  return t;
}
function qh(e, t) {
  if (e = _(P(e)), t = _(P(t)), O(e, t))
    return S;
  let n = 0;
  for (let r = 0; r < e.length; r++)
    n += e[r] * e[r] - t[r] * t[r];
  return n;
}
function Wh(e, t) {
  if (e = _(P(e)), t = _(P(t)), O(e, t))
    return S;
  let n = 0;
  e = _(P(e)), t = _(P(t));
  for (let r = 0; r < e.length; r++)
    n += e[r] * e[r] + t[r] * t[r];
  return n;
}
function Yh(e, t) {
  if (e = _(P(e)), t = _(P(t)), O(e, t))
    return S;
  let n = 0;
  e = P(e), t = P(t);
  for (let r = 0; r < e.length; r++)
    n += Math.pow(e[r] - t[r], 2);
  return n;
}
function Xh(e) {
  return e = w(e), e instanceof Error ? e : Math.tan(e);
}
function Kh(e) {
  if (e = w(e), e instanceof Error)
    return e;
  const t = Math.exp(2 * e);
  return (t - 1) / (t + 1);
}
function Jh(e, t) {
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || (e > 0 ? 1 : -1) * Math.floor(Math.abs(e) * Math.pow(10, t)) / Math.pow(10, t);
}
function Qh(e, t) {
  if (arguments.length !== 2)
    return G;
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || e + t;
}
function jh(e, t) {
  if (arguments.length !== 2)
    return G;
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || (t === 0 ? Ue : e / t);
}
function Zh(e, t) {
  return arguments.length !== 2 ? G : e instanceof Error ? e : t instanceof Error ? t : (e === null && (e = void 0), t === null && (t = void 0), e === t);
}
function em(e, t) {
  if (arguments.length !== 2)
    return G;
  if (e instanceof Error)
    return e;
  if (t instanceof Error)
    return t;
  co(e, t) ? (e = Me(e), t = Me(t)) : (e = w(e), t = w(t));
  const n = Q(e, t);
  return n || e > t;
}
function tm(e, t) {
  if (arguments.length !== 2)
    return G;
  co(e, t) ? (e = Me(e), t = Me(t)) : (e = w(e), t = w(t));
  const n = Q(e, t);
  return n || e >= t;
}
function nm(e, t) {
  if (arguments.length !== 2)
    return G;
  co(e, t) ? (e = Me(e), t = Me(t)) : (e = w(e), t = w(t));
  const n = Q(e, t);
  return n || e < t;
}
function rm(e, t) {
  if (arguments.length !== 2)
    return G;
  co(e, t) ? (e = Me(e), t = Me(t)) : (e = w(e), t = w(t));
  const n = Q(e, t);
  return n || e <= t;
}
function om(e, t) {
  if (arguments.length !== 2)
    return G;
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || e - t;
}
function im(e, t) {
  if (arguments.length !== 2)
    return G;
  e = w(e), t = w(t);
  const n = Q(e, t);
  return n || e * t;
}
function sm(e, t) {
  return arguments.length !== 2 ? G : e instanceof Error ? e : t instanceof Error ? t : (e === null && (e = void 0), t === null && (t = void 0), e !== t);
}
function am(e, t) {
  return arguments.length !== 2 ? G : ll(e, t);
}
var lm = /* @__PURE__ */ Object.freeze({
  __proto__: null,
  ADD: Qh,
  DIVIDE: jh,
  EQ: Zh,
  GT: em,
  GTE: tm,
  LT: nm,
  LTE: rm,
  MINUS: om,
  MULTIPLY: im,
  NE: sm,
  POW: am
});
const cm = [
  void 0,
  0,
  1,
  void 0,
  void 0,
  void 0,
  void 0,
  void 0,
  void 0,
  void 0,
  void 0,
  void 0,
  1,
  2,
  3,
  4,
  5,
  6,
  0
], um = [
  [],
  [1, 2, 3, 4, 5, 6, 7],
  [7, 1, 2, 3, 4, 5, 6],
  [6, 0, 1, 2, 3, 4, 5],
  [],
  [],
  [],
  [],
  [],
  [],
  [],
  [7, 1, 2, 3, 4, 5, 6],
  [6, 7, 1, 2, 3, 4, 5],
  [5, 6, 7, 1, 2, 3, 4],
  [4, 5, 6, 7, 1, 2, 3],
  [3, 4, 5, 6, 7, 1, 2],
  [2, 3, 4, 5, 6, 7, 1],
  [1, 2, 3, 4, 5, 6, 7]
], Wr = [
  [],
  [6, 0],
  [0, 1],
  [1, 2],
  [2, 3],
  [3, 4],
  [4, 5],
  [5, 6],
  void 0,
  void 0,
  void 0,
  [0, 0],
  [1, 1],
  [2, 2],
  [3, 3],
  [4, 4],
  [5, 5],
  [6, 6]
];
function fm(e, t, n) {
  let r;
  return e = w(e), t = w(t), n = w(n), O(e, t, n) ? r = S : (r = new Date(e, t - 1, n), r.getFullYear() < 0 && (r = T)), _t ? Ft(r) : r;
}
function Kt(e, t, n) {
  n = n.toUpperCase(), e = ee(e), t = ee(t);
  const r = e.getFullYear(), o = e.getMonth(), s = e.getDate(), l = t.getFullYear(), f = t.getMonth(), d = t.getDate();
  let c;
  switch (n) {
    case "Y":
      c = Math.floor(Gi(e, t));
      break;
    case "D":
      c = an(t, e);
      break;
    case "M":
      c = f - o + 12 * (l - r), d < s && c--;
      break;
    case "MD":
      s <= d ? c = d - s : (f === 0 ? (e.setFullYear(l - 1), e.setMonth(12)) : (e.setFullYear(l), e.setMonth(f - 1)), c = an(t, e));
      break;
    case "YM":
      c = f - o + 12 * (l - r), d < s && c--, c = c % 12;
      break;
    case "YD":
      f > o || f === o && d < s ? e.setFullYear(l) : e.setFullYear(l - 1), c = an(t, e);
      break;
  }
  return c;
}
function dm(e) {
  if (typeof e != "string")
    return S;
  const t = Date.parse(e);
  if (isNaN(t))
    return S;
  const n = new Date(e);
  return _t ? Ft(n) : n;
}
function hm(e) {
  const t = ee(e);
  return t instanceof Error ? t : t.getDate();
}
function Yr(e) {
  const t = new Date(e);
  return t.setHours(0, 0, 0, 0), t;
}
function an(e, t) {
  return e = ee(e), t = ee(t), e instanceof Error ? e : t instanceof Error ? t : Ft(Yr(e)) - Ft(Yr(t));
}
function Qt(e, t, n) {
  if (n = Ti(n || "false"), e = ee(e), t = ee(t), e instanceof Error)
    return e;
  if (t instanceof Error)
    return t;
  if (n instanceof Error)
    return n;
  const r = e.getMonth();
  let o = t.getMonth(), s, l;
  if (n)
    s = e.getDate() === 31 ? 30 : e.getDate(), l = t.getDate() === 31 ? 30 : t.getDate();
  else {
    const f = new Date(e.getFullYear(), r + 1, 0).getDate(), d = new Date(t.getFullYear(), o + 1, 0).getDate();
    s = e.getDate() === f ? 30 : e.getDate(), t.getDate() === d ? s < 30 ? (o++, l = 1) : l = 30 : l = t.getDate();
  }
  return 360 * (t.getFullYear() - e.getFullYear()) + 30 * (o - r) + (l - s);
}
function mm(e, t) {
  if (e = ee(e), e instanceof Error)
    return e;
  if (isNaN(t))
    return S;
  let n = e.getDate();
  e.setDate(1), t = parseInt(t, 10), e.setMonth(e.getMonth() + t);
  let r = e.getMonth();
  if (n > 28) {
    let o = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][r], s = e.getFullYear();
    r === 1 && (s % 4 === 0 && s % 100 !== 0 || s % 400 === 0) && (o = 29), n = Math.min(n, o);
  }
  return e.setDate(n), _t ? Ft(e) : e;
}
function gm(e, t) {
  if (e = ee(e), e instanceof Error)
    return e;
  if (isNaN(t))
    return S;
  t = parseInt(t, 10);
  const n = new Date(e.getFullYear(), e.getMonth() + t + 1, 0);
  return _t ? Ft(n) : n;
}
function pm(e) {
  return e = ee(e), e instanceof Error ? e : e.getHours();
}
function ul(e) {
  if (e = ee(e), e instanceof Error)
    return e;
  e = Yr(e), e.setDate(e.getDate() + 4 - (e.getDay() || 7));
  const t = new Date(e.getFullYear(), 0, 1);
  return Math.ceil(((e - t) / 864e5 + 1) / 7);
}
function vm(e) {
  return e = ee(e), e instanceof Error ? e : e.getMinutes();
}
function Nm(e) {
  return e = ee(e), e instanceof Error ? e : e.getMonth() + 1;
}
function mo(e, t, n) {
  return mo.INTL(e, t, 1, n);
}
mo.INTL = (e, t, n, r) => {
  if (e = ee(e), e instanceof Error)
    return e;
  if (t = ee(t), t instanceof Error)
    return t;
  let o = !1;
  const s = [], l = [1, 2, 3, 4, 5, 6, 0], f = new RegExp("^[0|1]{7}$");
  if (n === void 0)
    n = Wr[1];
  else if (typeof n == "string" && f.test(n)) {
    o = !0, n = n.split("");
    for (let a = 0; a < n.length; a++)
      n[a] === "1" && s.push(l[a]);
  } else
    n = Wr[n];
  if (!(n instanceof Array))
    return S;
  r === void 0 ? r = [] : r instanceof Array || (r = [r]);
  for (let a = 0; a < r.length; a++) {
    const u = ee(r[a]);
    if (u instanceof Error)
      return u;
    r[a] = u;
  }
  const d = Math.round((t - e) / (1e3 * 60 * 60 * 24)) + 1;
  let c = d;
  const i = e;
  for (let a = 0; a < d; a++) {
    const u = (/* @__PURE__ */ new Date()).getTimezoneOffset() > 0 ? i.getUTCDay() : i.getDay();
    let h = o ? s.includes(u) : u === n[0] || u === n[1];
    for (let g = 0; g < r.length; g++) {
      const m = r[g];
      if (m.getDate() === i.getDate() && m.getMonth() === i.getMonth() && m.getFullYear() === i.getFullYear()) {
        h = !0;
        break;
      }
    }
    h && c--, i.setDate(i.getDate() + 1);
  }
  return c;
};
function Cm() {
  return _t ? Ft(/* @__PURE__ */ new Date()) : /* @__PURE__ */ new Date();
}
function Em(e) {
  return e = ee(e), e instanceof Error ? e : e.getSeconds();
}
function bm(e, t, n) {
  return e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : e < 0 || t < 0 || n < 0 ? T : (3600 * e + 60 * t + n) / 86400;
}
function wm(e) {
  return e = ee(e), e instanceof Error ? e : (3600 * e.getHours() + 60 * e.getMinutes() + e.getSeconds()) / 86400;
}
function Im() {
  const e = Yr(/* @__PURE__ */ new Date());
  return _t ? Ft(e) : e;
}
function Mm(e, t) {
  if (e = ee(e), e instanceof Error)
    return e;
  t === void 0 && (t = 1);
  const n = e.getDay();
  return um[t][n];
}
function ym(e, t) {
  if (e = ee(e), e instanceof Error)
    return e;
  if (t === void 0 && (t = 1), t === 21)
    return ul(e);
  const n = cm[t];
  let r = new Date(e.getFullYear(), 0, 1);
  const o = r.getDay() < n ? 1 : 0;
  return r -= Math.abs(r.getDay() - n) * 24 * 60 * 60 * 1e3, Math.floor((e - r) / (1e3 * 60 * 60 * 24) / 7 + 1) + o;
}
function go(e, t, n) {
  return go.INTL(e, t, 1, n);
}
go.INTL = (e, t, n, r) => {
  if (e = ee(e), e instanceof Error)
    return e;
  if (t = w(t), t instanceof Error)
    return t;
  if (n === void 0 ? n = Wr[1] : n = Wr[n], !(n instanceof Array))
    return S;
  r === void 0 ? r = [] : r instanceof Array || (r = [r]);
  for (let l = 0; l < r.length; l++) {
    const f = ee(r[l]);
    if (f instanceof Error)
      return f;
    r[l] = f;
  }
  let o = 0;
  const s = Math.sign(t);
  for (; o < t * s; ) {
    e.setDate(e.getDate() + s);
    const l = e.getDay();
    if (!(l === n[0] || l === n[1])) {
      for (let f = 0; f < r.length; f++) {
        const d = r[f];
        if (d.getDate() === e.getDate() && d.getMonth() === e.getMonth() && d.getFullYear() === e.getFullYear()) {
          o--;
          break;
        }
      }
      o++;
    }
  }
  return e.getFullYear() < 1900 ? S : e;
};
function Sm(e) {
  return e = ee(e), e instanceof Error ? e : e.getFullYear();
}
function yo(e) {
  return new Date(e, 1, 29).getMonth() === 1;
}
function Mr(e, t) {
  return Math.ceil((t - e) / 1e3 / 60 / 60 / 24);
}
function Gi(e, t, n) {
  if (e = ee(e), e instanceof Error)
    return e;
  if (t = ee(t), t instanceof Error)
    return t;
  n = n || 0;
  let r = e.getDate();
  const o = e.getMonth() + 1, s = e.getFullYear();
  let l = t.getDate();
  const f = t.getMonth() + 1, d = t.getFullYear();
  switch (n) {
    case 0:
      return r === 31 && l === 31 ? (r = 30, l = 30) : r === 31 ? r = 30 : r === 30 && l === 31 && (l = 30), (l + f * 30 + d * 360 - (r + o * 30 + s * 360)) / 360;
    case 1: {
      const c = (g, m) => {
        const p = g.getFullYear(), E = new Date(p, 2, 1);
        if (yo(p) && g < E && m >= E)
          return !0;
        const b = m.getFullYear(), N = new Date(b, 2, 1);
        return yo(b) && m >= N && g < N;
      };
      let i = 365;
      if (s === d || s + 1 === d && (o > f || o === f && r >= l))
        return (s === d && yo(s) || c(e, t) || f === 1 && l === 29) && (i = 366), Mr(e, t) / i;
      const a = d - s + 1, h = (new Date(d + 1, 0, 1) - new Date(s, 0, 1)) / 1e3 / 60 / 60 / 24 / a;
      return Mr(e, t) / h;
    }
    case 2:
      return Mr(e, t) / 360;
    case 3:
      return Mr(e, t) / 365;
    case 4:
      return (l + f * 30 + d * 360 - (r + o * 30 + s * 360)) / 360;
  }
}
function $i(e) {
  return /^[01]{1,10}$/.test(e);
}
function Rm(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : so.besseli(e, t);
}
function Am(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : so.besselj(e, t);
}
function Tm(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : so.besselk(e, t);
}
function xm(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : so.bessely(e, t);
}
function Dm(e) {
  if (!$i(e))
    return T;
  const t = parseInt(e, 2), n = e.toString();
  return n.length === 10 && n.substring(0, 1) === "1" ? parseInt(n.substring(1), 2) - 512 : t;
}
function Om(e, t) {
  if (!$i(e))
    return T;
  const n = e.toString();
  if (n.length === 10 && n.substring(0, 1) === "1")
    return (1099511627264 + parseInt(n.substring(1), 2)).toString(16);
  const r = parseInt(e, 2).toString(16);
  return t === void 0 ? r : isNaN(t) ? S : t < 0 ? T : (t = Math.floor(t), t >= r.length ? at("0", t - r.length) + r : T);
}
function km(e, t) {
  if (!$i(e))
    return T;
  const n = e.toString();
  if (n.length === 10 && n.substring(0, 1) === "1")
    return (1073741312 + parseInt(n.substring(1), 2)).toString(8);
  const r = parseInt(e, 2).toString(8);
  return t === void 0 ? r : isNaN(t) ? S : t < 0 ? T : (t = Math.floor(t), t >= r.length ? at("0", t - r.length) + r : T);
}
function Lm(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : e < 0 || t < 0 || Math.floor(e) !== e || Math.floor(t) !== t || e > 281474976710655 || t > 281474976710655 ? T : e & t;
}
function Fm(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : e < 0 || Math.floor(e) !== e || e > 281474976710655 || Math.abs(t) > 53 ? T : t >= 0 ? e << t : e >> -t;
}
function Pm(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : e < 0 || t < 0 || Math.floor(e) !== e || Math.floor(t) !== t || e > 281474976710655 || t > 281474976710655 ? T : e | t;
}
function Um(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : e < 0 || Math.floor(e) !== e || e > 281474976710655 || Math.abs(t) > 53 ? T : t >= 0 ? e >> t : e << -t;
}
function _m(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : e < 0 || t < 0 || Math.floor(e) !== e || Math.floor(t) !== t || e > 281474976710655 || t > 281474976710655 ? T : e ^ t;
}
function Ve(e, t, n) {
  if (e = w(e), t = w(t), O(e, t))
    return e;
  if (n = n === void 0 ? "i" : n, n !== "i" && n !== "j")
    return S;
  if (e === 0 && t === 0)
    return 0;
  if (e === 0)
    return t === 1 ? n : t.toString() + n;
  if (t === 0)
    return e.toString();
  {
    const r = t > 0 ? "+" : "";
    return e.toString() + r + (t === 1 ? n : t.toString() + n);
  }
}
function Bm(e, t, n) {
  if (e = w(e), e instanceof Error)
    return e;
  const r = [
    ["a.u. of action", "?", null, "action", !1, !1, 105457168181818e-48],
    ["a.u. of charge", "e", null, "electric_charge", !1, !1, 160217653141414e-33],
    ["a.u. of energy", "Eh", null, "energy", !1, !1, 435974417757576e-32],
    ["a.u. of length", "a?", null, "length", !1, !1, 529177210818182e-25],
    ["a.u. of mass", "m?", null, "mass", !1, !1, 910938261616162e-45],
    ["a.u. of time", "?/Eh", null, "time", !1, !1, 241888432650516e-31],
    ["admiralty knot", "admkn", null, "speed", !1, !0, 0.514773333],
    ["ampere", "A", null, "electric_current", !0, !1, 1],
    ["ampere per meter", "A/m", null, "magnetic_field_intensity", !0, !1, 1],
    ["ångström", "Å", ["ang"], "length", !1, !0, 1e-10],
    ["are", "ar", null, "area", !1, !0, 100],
    ["astronomical unit", "ua", null, "length", !1, !1, 149597870691667e-25],
    ["bar", "bar", null, "pressure", !1, !1, 1e5],
    ["barn", "b", null, "area", !1, !1, 1e-28],
    ["becquerel", "Bq", null, "radioactivity", !0, !1, 1],
    ["bit", "bit", ["b"], "information", !1, !0, 1],
    ["btu", "BTU", ["btu"], "energy", !1, !0, 1055.05585262],
    ["byte", "byte", null, "information", !1, !0, 8],
    ["candela", "cd", null, "luminous_intensity", !0, !1, 1],
    ["candela per square metre", "cd/m?", null, "luminance", !0, !1, 1],
    ["coulomb", "C", null, "electric_charge", !0, !1, 1],
    ["cubic ångström", "ang3", ["ang^3"], "volume", !1, !0, 1e-30],
    ["cubic foot", "ft3", ["ft^3"], "volume", !1, !0, 0.028316846592],
    ["cubic inch", "in3", ["in^3"], "volume", !1, !0, 16387064e-12],
    ["cubic light-year", "ly3", ["ly^3"], "volume", !1, !0, 846786664623715e-61],
    ["cubic metre", "m3", ["m^3"], "volume", !0, !0, 1],
    ["cubic mile", "mi3", ["mi^3"], "volume", !1, !0, 416818182544058e-5],
    ["cubic nautical mile", "Nmi3", ["Nmi^3"], "volume", !1, !0, 6352182208],
    ["cubic Pica", "Pica3", ["Picapt3", "Pica^3", "Picapt^3"], "volume", !1, !0, 758660370370369e-22],
    ["cubic yard", "yd3", ["yd^3"], "volume", !1, !0, 0.764554857984],
    ["cup", "cup", null, "volume", !1, !0, 2365882365e-13],
    ["dalton", "Da", ["u"], "mass", !1, !1, 166053886282828e-41],
    ["day", "d", ["day"], "time", !1, !0, 86400],
    ["degree", "°", null, "angle", !1, !1, 0.0174532925199433],
    ["degrees Rankine", "Rank", null, "temperature", !1, !0, 0.555555555555556],
    ["dyne", "dyn", ["dy"], "force", !1, !0, 1e-5],
    ["electronvolt", "eV", ["ev"], "energy", !1, !0, 1.60217656514141],
    ["ell", "ell", null, "length", !1, !0, 1.143],
    ["erg", "erg", ["e"], "energy", !1, !0, 1e-7],
    ["farad", "F", null, "electric_capacitance", !0, !1, 1],
    ["fluid ounce", "oz", null, "volume", !1, !0, 295735295625e-16],
    ["foot", "ft", null, "length", !1, !0, 0.3048],
    ["foot-pound", "flb", null, "energy", !1, !0, 1.3558179483314],
    ["gal", "Gal", null, "acceleration", !1, !1, 0.01],
    ["gallon", "gal", null, "volume", !1, !0, 0.003785411784],
    ["gauss", "G", ["ga"], "magnetic_flux_density", !1, !0, 1],
    ["grain", "grain", null, "mass", !1, !0, 647989e-10],
    ["gram", "g", null, "mass", !1, !0, 1e-3],
    ["gray", "Gy", null, "absorbed_dose", !0, !1, 1],
    ["gross registered ton", "GRT", ["regton"], "volume", !1, !0, 2.8316846592],
    ["hectare", "ha", null, "area", !1, !0, 1e4],
    ["henry", "H", null, "inductance", !0, !1, 1],
    ["hertz", "Hz", null, "frequency", !0, !1, 1],
    ["horsepower", "HP", ["h"], "power", !1, !0, 745.69987158227],
    ["horsepower-hour", "HPh", ["hh", "hph"], "energy", !1, !0, 2684519538e-3],
    ["hour", "h", ["hr"], "time", !1, !0, 3600],
    ["imperial gallon (U.K.)", "uk_gal", null, "volume", !1, !0, 454609e-8],
    ["imperial hundredweight", "lcwt", ["uk_cwt", "hweight"], "mass", !1, !0, 50.802345],
    ["imperial quart (U.K)", "uk_qt", null, "volume", !1, !0, 0.0011365225],
    ["imperial ton", "brton", ["uk_ton", "LTON"], "mass", !1, !0, 1016.046909],
    ["inch", "in", null, "length", !1, !0, 0.0254],
    ["international acre", "uk_acre", null, "area", !1, !0, 4046.8564224],
    ["IT calorie", "cal", null, "energy", !1, !0, 4.1868],
    ["joule", "J", null, "energy", !0, !0, 1],
    ["katal", "kat", null, "catalytic_activity", !0, !1, 1],
    ["kelvin", "K", ["kel"], "temperature", !0, !0, 1],
    ["kilogram", "kg", null, "mass", !0, !0, 1],
    ["knot", "kn", null, "speed", !1, !0, 0.514444444444444],
    ["light-year", "ly", null, "length", !1, !0, 9460730472580800],
    ["litre", "L", ["l", "lt"], "volume", !1, !0, 1e-3],
    ["lumen", "lm", null, "luminous_flux", !0, !1, 1],
    ["lux", "lx", null, "illuminance", !0, !1, 1],
    ["maxwell", "Mx", null, "magnetic_flux", !1, !1, 1e-18],
    ["measurement ton", "MTON", null, "volume", !1, !0, 1.13267386368],
    ["meter per hour", "m/h", ["m/hr"], "speed", !1, !0, 27777777777778e-17],
    ["meter per second", "m/s", ["m/sec"], "speed", !0, !0, 1],
    ["meter per second squared", "m?s??", null, "acceleration", !0, !1, 1],
    ["parsec", "pc", ["parsec"], "length", !1, !0, 30856775814671900],
    ["meter squared per second", "m?/s", null, "kinematic_viscosity", !0, !1, 1],
    ["metre", "m", null, "length", !0, !0, 1],
    ["miles per hour", "mph", null, "speed", !1, !0, 0.44704],
    ["millimetre of mercury", "mmHg", null, "pressure", !1, !1, 133.322],
    ["minute", "?", null, "angle", !1, !1, 290888208665722e-18],
    ["minute", "min", ["mn"], "time", !1, !0, 60],
    ["modern teaspoon", "tspm", null, "volume", !1, !0, 5e-6],
    ["mole", "mol", null, "amount_of_substance", !0, !1, 1],
    ["morgen", "Morgen", null, "area", !1, !0, 2500],
    ["n.u. of action", "?", null, "action", !1, !1, 105457168181818e-48],
    ["n.u. of mass", "m?", null, "mass", !1, !1, 910938261616162e-45],
    ["n.u. of speed", "c?", null, "speed", !1, !1, 299792458],
    ["n.u. of time", "?/(me?c??)", null, "time", !1, !1, 128808866778687e-35],
    ["nautical mile", "M", ["Nmi"], "length", !1, !0, 1852],
    ["newton", "N", null, "force", !0, !0, 1],
    ["œrsted", "Oe ", null, "magnetic_field_intensity", !1, !1, 79.5774715459477],
    ["ohm", "Ω", null, "electric_resistance", !0, !1, 1],
    ["ounce mass", "ozm", null, "mass", !1, !0, 0.028349523125],
    ["pascal", "Pa", null, "pressure", !0, !1, 1],
    ["pascal second", "Pa?s", null, "dynamic_viscosity", !0, !1, 1],
    ["pferdestärke", "PS", null, "power", !1, !0, 735.49875],
    ["phot", "ph", null, "illuminance", !1, !1, 1e-4],
    ["pica (1/6 inch)", "pica", null, "length", !1, !0, 35277777777778e-17],
    ["pica (1/72 inch)", "Pica", ["Picapt"], "length", !1, !0, 0.00423333333333333],
    ["poise", "P", null, "dynamic_viscosity", !1, !1, 0.1],
    ["pond", "pond", null, "force", !1, !0, 980665e-8],
    ["pound force", "lbf", null, "force", !1, !0, 4.4482216152605],
    ["pound mass", "lbm", null, "mass", !1, !0, 0.45359237],
    ["quart", "qt", null, "volume", !1, !0, 946352946e-12],
    ["radian", "rad", null, "angle", !0, !1, 1],
    ["second", "?", null, "angle", !1, !1, 484813681109536e-20],
    ["second", "s", ["sec"], "time", !0, !0, 1],
    ["short hundredweight", "cwt", ["shweight"], "mass", !1, !0, 45.359237],
    ["siemens", "S", null, "electrical_conductance", !0, !1, 1],
    ["sievert", "Sv", null, "equivalent_dose", !0, !1, 1],
    ["slug", "sg", null, "mass", !1, !0, 14.59390294],
    ["square ångström", "ang2", ["ang^2"], "area", !1, !0, 1e-20],
    ["square foot", "ft2", ["ft^2"], "area", !1, !0, 0.09290304],
    ["square inch", "in2", ["in^2"], "area", !1, !0, 64516e-8],
    ["square light-year", "ly2", ["ly^2"], "area", !1, !0, 895054210748189e17],
    ["square meter", "m?", null, "area", !0, !0, 1],
    ["square mile", "mi2", ["mi^2"], "area", !1, !0, 2589988110336e-6],
    ["square nautical mile", "Nmi2", ["Nmi^2"], "area", !1, !0, 3429904],
    ["square Pica", "Pica2", ["Picapt2", "Pica^2", "Picapt^2"], "area", !1, !0, 1792111111111e-17],
    ["square yard", "yd2", ["yd^2"], "area", !1, !0, 0.83612736],
    ["statute mile", "mi", null, "length", !1, !0, 1609.344],
    ["steradian", "sr", null, "solid_angle", !0, !1, 1],
    ["stilb", "sb", null, "luminance", !1, !1, 1e-4],
    ["stokes", "St", null, "kinematic_viscosity", !1, !1, 1e-4],
    ["stone", "stone", null, "mass", !1, !0, 6.35029318],
    ["tablespoon", "tbs", null, "volume", !1, !0, 147868e-10],
    ["teaspoon", "tsp", null, "volume", !1, !0, 492892e-11],
    ["tesla", "T", null, "magnetic_flux_density", !0, !0, 1],
    ["thermodynamic calorie", "c", null, "energy", !1, !0, 4.184],
    ["ton", "ton", null, "mass", !1, !0, 907.18474],
    ["tonne", "t", null, "mass", !1, !1, 1e3],
    ["U.K. pint", "uk_pt", null, "volume", !1, !0, 56826125e-11],
    ["U.S. bushel", "bushel", null, "volume", !1, !0, 0.03523907],
    ["U.S. oil barrel", "barrel", null, "volume", !1, !0, 0.158987295],
    ["U.S. pint", "pt", ["us_pt"], "volume", !1, !0, 473176473e-12],
    ["U.S. survey mile", "survey_mi", null, "length", !1, !0, 1609.347219],
    ["U.S. survey/statute acre", "us_acre", null, "area", !1, !0, 4046.87261],
    ["volt", "V", null, "voltage", !0, !1, 1],
    ["watt", "W", null, "power", !0, !0, 1],
    ["watt-hour", "Wh", ["wh"], "energy", !1, !0, 3600],
    ["weber", "Wb", null, "magnetic_flux", !0, !1, 1],
    ["yard", "yd", null, "length", !1, !0, 0.9144],
    ["year", "yr", null, "time", !1, !0, 31557600]
  ], o = {
    Yi: ["yobi", 80, 12089258196146292e8, "Yi", "yotta"],
    Zi: ["zebi", 70, 11805916207174113e5, "Zi", "zetta"],
    Ei: ["exbi", 60, 1152921504606847e3, "Ei", "exa"],
    Pi: ["pebi", 50, 1125899906842624, "Pi", "peta"],
    Ti: ["tebi", 40, 1099511627776, "Ti", "tera"],
    Gi: ["gibi", 30, 1073741824, "Gi", "giga"],
    Mi: ["mebi", 20, 1048576, "Mi", "mega"],
    ki: ["kibi", 10, 1024, "ki", "kilo"]
  }, s = {
    Y: ["yotta", 1e24, "Y"],
    Z: ["zetta", 1e21, "Z"],
    E: ["exa", 1e18, "E"],
    P: ["peta", 1e15, "P"],
    T: ["tera", 1e12, "T"],
    G: ["giga", 1e9, "G"],
    M: ["mega", 1e6, "M"],
    k: ["kilo", 1e3, "k"],
    h: ["hecto", 100, "h"],
    e: ["dekao", 10, "e"],
    d: ["deci", 0.1, "d"],
    c: ["centi", 0.01, "c"],
    m: ["milli", 1e-3, "m"],
    u: ["micro", 1e-6, "u"],
    n: ["nano", 1e-9, "n"],
    p: ["pico", 1e-12, "p"],
    f: ["femto", 1e-15, "f"],
    a: ["atto", 1e-18, "a"],
    z: ["zepto", 1e-21, "z"],
    y: ["yocto", 1e-24, "y"]
  };
  let l = null, f = null, d = t, c = n, i = 1, a = 1, u;
  for (let h = 0; h < r.length; h++)
    u = r[h][2] === null ? [] : r[h][2], (r[h][1] === d || u.indexOf(d) >= 0) && (l = r[h]), (r[h][1] === c || u.indexOf(c) >= 0) && (f = r[h]);
  if (l === null) {
    const h = o[t.substring(0, 2)];
    let g = s[t.substring(0, 1)];
    t.substring(0, 2) === "da" && (g = ["dekao", 10, "da"]), h ? (i = h[2], d = t.substring(2)) : g && (i = g[1], d = t.substring(g[2].length));
    for (let m = 0; m < r.length; m++)
      u = r[m][2] === null ? [] : r[m][2], (r[m][1] === d || u.indexOf(d) >= 0) && (l = r[m]);
  }
  if (f === null) {
    const h = o[n.substring(0, 2)];
    let g = s[n.substring(0, 1)];
    n.substring(0, 2) === "da" && (g = ["dekao", 10, "da"]), h ? (a = h[2], c = n.substring(2)) : g && (a = g[1], c = n.substring(g[2].length));
    for (let m = 0; m < r.length; m++)
      u = r[m][2] === null ? [] : r[m][2], (r[m][1] === c || u.indexOf(c) >= 0) && (f = r[m]);
  }
  return l === null || f === null || l[3] !== f[3] ? G : e * l[6] * i / (f[6] * a);
}
function zm(e, t) {
  if (e = w(e), e instanceof Error)
    return e;
  if (!/^-?[0-9]{1,3}$/.test(e) || e < -512 || e > 511)
    return T;
  if (e < 0)
    return "1" + at("0", 9 - (512 + e).toString(2).length) + (512 + e).toString(2);
  const n = parseInt(e, 10).toString(2);
  return typeof t > "u" ? n : isNaN(t) ? S : t < 0 ? T : (t = Math.floor(t), t >= n.length ? at("0", t - n.length) + n : T);
}
function Hm(e, t) {
  if (e = w(e), e instanceof Error)
    return e;
  if (!/^-?[0-9]{1,12}$/.test(e) || e < -549755813888 || e > 549755813887)
    return T;
  if (e < 0)
    return (1099511627776 + e).toString(16);
  const n = parseInt(e, 10).toString(16);
  return typeof t > "u" ? n : isNaN(t) ? S : t < 0 ? T : (t = Math.floor(t), t >= n.length ? at("0", t - n.length) + n : T);
}
function Vm(e, t) {
  if (e = w(e), e instanceof Error)
    return e;
  if (!/^-?[0-9]{1,9}$/.test(e) || e < -536870912 || e > 536870911)
    return T;
  if (e < 0)
    return (1073741824 + e).toString(8);
  const n = parseInt(e, 10).toString(8);
  return typeof t > "u" ? n : isNaN(t) ? S : t < 0 ? T : (t = Math.floor(t), t >= n.length ? at("0", t - n.length) + n : T);
}
function Gm(e, t) {
  return t = t === void 0 ? 0 : t, e = w(e), t = w(t), O(e, t) ? S : e === t ? 1 : 0;
}
function fl(e, t) {
  return t = t === void 0 ? 0 : t, e = w(e), t = w(t), O(e, t) ? S : U.erf(e);
}
function dl(e) {
  return isNaN(e) ? S : U.erfc(e);
}
function $m(e, t) {
  return t = t || 0, e = w(e), O(t, e) ? e : e >= t ? 1 : 0;
}
function qm(e, t) {
  if (!/^[0-9A-Fa-f]{1,10}$/.test(e))
    return T;
  const n = e.length === 10 && e.substring(0, 1).toLowerCase() === "f", r = n ? parseInt(e, 16) - 1099511627776 : parseInt(e, 16);
  if (r < -512 || r > 511)
    return T;
  if (n)
    return "1" + at("0", 9 - (512 + r).toString(2).length) + (512 + r).toString(2);
  const o = r.toString(2);
  return t === void 0 ? o : isNaN(t) ? S : t < 0 ? T : (t = Math.floor(t), t >= o.length ? at("0", t - o.length) + o : T);
}
function Wm(e) {
  if (!/^[0-9A-Fa-f]{1,10}$/.test(e))
    return T;
  const t = parseInt(e, 16);
  return t >= 549755813888 ? t - 1099511627776 : t;
}
function Ym(e, t) {
  if (!/^[0-9A-Fa-f]{1,10}$/.test(e))
    return T;
  const n = parseInt(e, 16);
  if (n > 536870911 && n < 1098974756864)
    return T;
  if (n >= 1098974756864)
    return (n - 1098437885952).toString(8);
  const r = n.toString(8);
  return t === void 0 ? r : isNaN(t) ? S : t < 0 ? T : (t = Math.floor(t), t >= r.length ? at("0", t - r.length) + r : T);
}
function qi(e) {
  const t = ge(e), n = me(e);
  return O(t, n) ? S : Math.sqrt(Math.pow(t, 2) + Math.pow(n, 2));
}
function me(e) {
  if (e === void 0 || e === !0 || e === !1)
    return S;
  if (e === 0 || e === "0")
    return 0;
  if (["i", "j"].indexOf(e) >= 0)
    return 1;
  e = e + "", e = e.replace("+i", "+1i").replace("-i", "-1i").replace("+j", "+1j").replace("-j", "-1j");
  let t = e.indexOf("+"), n = e.indexOf("-");
  t === 0 && (t = e.indexOf("+", 1)), n === 0 && (n = e.indexOf("-", 1));
  const r = e.substring(e.length - 1, e.length), o = r === "i" || r === "j";
  return t >= 0 || n >= 0 ? o ? t >= 0 ? isNaN(e.substring(0, t)) || isNaN(e.substring(t + 1, e.length - 1)) ? T : Number(e.substring(t + 1, e.length - 1)) : isNaN(e.substring(0, n)) || isNaN(e.substring(n + 1, e.length - 1)) ? T : -Number(e.substring(n + 1, e.length - 1)) : T : o ? isNaN(e.substring(0, e.length - 1)) ? T : e.substring(0, e.length - 1) : isNaN(e) ? T : 0;
}
function Wi(e) {
  const t = ge(e), n = me(e);
  return O(t, n) ? S : t === 0 && n === 0 ? Ue : t === 0 && n > 0 ? Math.PI / 2 : t === 0 && n < 0 ? -Math.PI / 2 : n === 0 && t > 0 ? 0 : n === 0 && t < 0 ? -Math.PI : t > 0 ? Math.atan(n / t) : t < 0 && n >= 0 ? Math.atan(n / t) + Math.PI : Math.atan(n / t) - Math.PI;
}
function Xm(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  return r = r === "i" || r === "j" ? r : "i", n !== 0 ? Ve(t, -n, r) : e;
}
function po(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  return r = r === "i" || r === "j" ? r : "i", Ve(
    Math.cos(t) * (Math.exp(n) + Math.exp(-n)) / 2,
    -Math.sin(t) * (Math.exp(n) - Math.exp(-n)) / 2,
    r
  );
}
function hl(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  return r = r === "i" || r === "j" ? r : "i", Ve(
    Math.cos(n) * (Math.exp(t) + Math.exp(-t)) / 2,
    Math.sin(n) * (Math.exp(t) - Math.exp(-t)) / 2,
    r
  );
}
function Km(e) {
  const t = ge(e), n = me(e);
  return O(t, n) ? S : En(po(e), vo(e));
}
function En(e, t) {
  const n = ge(e), r = me(e), o = ge(t), s = me(t);
  if (O(n, r, o, s))
    return S;
  const l = e.substring(e.length - 1), f = t.substring(t.length - 1);
  let d = "i";
  if ((l === "j" || f === "j") && (d = "j"), o === 0 && s === 0)
    return T;
  const c = o * o + s * s;
  return Ve((n * o + r * s) / c, (r * o - n * s) / c, d);
}
function Jm(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  r = r === "i" || r === "j" ? r : "i";
  const o = Math.exp(t);
  return Ve(o * Math.cos(n), o * Math.sin(n), r);
}
function Qm(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  return r = r === "i" || r === "j" ? r : "i", Ve(Math.log(Math.sqrt(t * t + n * n)), Math.atan(n / t), r);
}
function jm(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  return r = r === "i" || r === "j" ? r : "i", Ve(Math.log(Math.sqrt(t * t + n * n)) / Math.log(10), Math.atan(n / t) / Math.log(10), r);
}
function Zm(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  return r = r === "i" || r === "j" ? r : "i", Ve(Math.log(Math.sqrt(t * t + n * n)) / Math.log(2), Math.atan(n / t) / Math.log(2), r);
}
function e0(e, t) {
  t = w(t);
  const n = ge(e), r = me(e);
  if (O(t, n, r))
    return S;
  let o = e.substring(e.length - 1);
  o = o === "i" || o === "j" ? o : "i";
  const s = Math.pow(qi(e), t), l = Wi(e);
  return Ve(s * Math.cos(t * l), s * Math.sin(t * l), o);
}
function t0() {
  let e = arguments[0];
  if (!arguments.length)
    return S;
  for (let t = 1; t < arguments.length; t++) {
    const n = ge(e), r = me(e), o = ge(arguments[t]), s = me(arguments[t]);
    if (O(n, r, o, s))
      return S;
    e = Ve(n * o - r * s, n * s + r * o);
  }
  return e;
}
function ge(e) {
  if (e === void 0 || e === !0 || e === !1)
    return S;
  if (e === 0 || e === "0" || ["i", "+i", "1i", "+1i", "-i", "-1i", "j", "+j", "1j", "+1j", "-j", "-1j"].indexOf(e) >= 0)
    return 0;
  e = e + "";
  let t = e.indexOf("+"), n = e.indexOf("-");
  t === 0 && (t = e.indexOf("+", 1)), n === 0 && (n = e.indexOf("-", 1));
  const r = e.substring(e.length - 1, e.length), o = r === "i" || r === "j";
  return t >= 0 || n >= 0 ? o ? t >= 0 ? isNaN(e.substring(0, t)) || isNaN(e.substring(t + 1, e.length - 1)) ? T : Number(e.substring(0, t)) : isNaN(e.substring(0, n)) || isNaN(e.substring(n + 1, e.length - 1)) ? T : Number(e.substring(0, n)) : T : o ? isNaN(e.substring(0, e.length - 1)) ? T : 0 : isNaN(e) ? T : e;
}
function n0(e) {
  if (e === !0 || e === !1)
    return S;
  const t = ge(e), n = me(e);
  return O(t, n) ? S : En("1", po(e));
}
function r0(e) {
  const t = ge(e), n = me(e);
  return O(t, n) ? S : En("1", hl(e));
}
function vo(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  return r = r === "i" || r === "j" ? r : "i", Ve(
    Math.sin(t) * (Math.exp(n) + Math.exp(-n)) / 2,
    Math.cos(t) * (Math.exp(n) - Math.exp(-n)) / 2,
    r
  );
}
function ml(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  return r = r === "i" || r === "j" ? r : "i", Ve(
    Math.cos(n) * (Math.exp(t) - Math.exp(-t)) / 2,
    Math.sin(n) * (Math.exp(t) + Math.exp(-t)) / 2,
    r
  );
}
function o0(e) {
  const t = ge(e), n = me(e);
  if (O(t, n))
    return S;
  let r = e.substring(e.length - 1);
  r = r === "i" || r === "j" ? r : "i";
  const o = Math.sqrt(qi(e)), s = Wi(e);
  return Ve(o * Math.cos(s / 2), o * Math.sin(s / 2), r);
}
function i0(e) {
  if (e === !0 || e === !1)
    return S;
  const t = ge(e), n = me(e);
  return O(t, n) ? T : En("1", vo(e));
}
function s0(e) {
  if (e === !0 || e === !1)
    return S;
  const t = ge(e), n = me(e);
  return O(t, n) ? T : En("1", ml(e));
}
function a0(e, t) {
  const n = ge(e), r = me(e), o = ge(t), s = me(t);
  if (O(n, r, o, s))
    return S;
  const l = e.substring(e.length - 1), f = t.substring(t.length - 1);
  let d = "i";
  return (l === "j" || f === "j") && (d = "j"), Ve(n - o, r - s, d);
}
function l0() {
  if (!arguments.length)
    return S;
  const e = P(arguments);
  let t = 0, n = 0;
  for (const r of e) {
    const o = +ge(r), s = +me(r);
    if (O(o, s))
      return S;
    t += o, n += s;
  }
  return Ve(t, n, "i");
}
function c0(e) {
  if (e === !0 || e === !1)
    return S;
  const t = ge(e), n = me(e);
  return O(t, n) ? S : En(vo(e), po(e));
}
function u0(e, t) {
  if (!/^[0-7]{1,10}$/.test(e))
    return T;
  const n = e.length === 10 && e.substring(0, 1) === "7", r = n ? parseInt(e, 8) - 1073741824 : parseInt(e, 8);
  if (r < -512 || r > 511)
    return T;
  if (n)
    return "1" + at("0", 9 - (512 + r).toString(2).length) + (512 + r).toString(2);
  const o = r.toString(2);
  return typeof t > "u" ? o : isNaN(t) ? S : t < 0 ? T : (t = Math.floor(t), t >= o.length ? at("0", t - o.length) + o : T);
}
function f0(e) {
  if (!/^[0-7]{1,10}$/.test(e))
    return T;
  const t = parseInt(e, 8);
  return t >= 536870912 ? t - 1073741824 : t;
}
function d0(e, t) {
  if (!/^[0-7]{1,10}$/.test(e))
    return T;
  const n = parseInt(e, 8);
  if (n >= 536870912)
    return "ff" + (n + 3221225472).toString(16);
  const r = n.toString(16);
  return t === void 0 ? r : isNaN(t) ? S : t < 0 ? T : (t = Math.floor(t), t >= r.length ? at("0", t - r.length) + r : T);
}
const h0 = pr.DIST, m0 = pr.INV, g0 = Hn.DIST, p0 = Jt.MATH, v0 = Jt.PRECISE, N0 = It.DIST, C0 = It.DIST.RT, E0 = It.INV, b0 = It.INV.RT, w0 = It.TEST, I0 = Vn.P, M0 = Vn.P, y0 = Vn.S, S0 = Hn.INV, R0 = dl.PRECISE, A0 = fl.PRECISE, T0 = ki.DIST, x0 = Mt.DIST, D0 = Mt.DIST.RT, O0 = Mt.INV, k0 = Mt.INV.RT, L0 = mn.MATH, F0 = mn.PRECISE, P0 = Mt.TEST, U0 = vr.DIST, _0 = vr.INV, B0 = Li.PRECISE, z0 = Fi.DIST, H0 = Gn.INV, V0 = Gn.DIST, G0 = Gn.INV, $0 = dn.MULT, q0 = dn.SNGL, W0 = Ui.DIST, Y0 = mo.INTL, X0 = yt.DIST, K0 = yt.INV, J0 = yt.S.DIST, Q0 = yt.S.INV, j0 = st.EXC, Z0 = st.INC, e1 = Nr.EXC, t1 = Nr.INC, n1 = _i.DIST, r1 = hn.EXC, o1 = hn.INC, i1 = Cr.AVG, s1 = Cr.EQ, a1 = Bi.P, l1 = Ze.P, c1 = Ze.S, u1 = et.DIST, f1 = et.DIST.RT, d1 = et.INV, h1 = et.TEST, m1 = qe.P, g1 = qe.S, p1 = zi.DIST, v1 = go.INTL, N1 = Hi.TEST;
function Yi(e) {
  const t = [];
  return He(e, (n) => {
    n && t.push(n);
  }), t;
}
function pt(e, t) {
  const n = {};
  for (let s = 1; s < e[0].length; ++s)
    n[s] = !0;
  let r = t[0].length;
  for (let s = 1; s < t.length; ++s)
    t[s].length > r && (r = t[s].length);
  for (let s = 1; s < e.length; ++s)
    for (let l = 1; l < e[s].length; ++l) {
      let f = !1, d = !1;
      for (let c = 0; c < t.length; ++c) {
        const i = t[c];
        if (i.length < r)
          continue;
        const a = i[0];
        if (e[s][0] === a) {
          d = !0;
          for (let u = 1; u < i.length; ++u)
            if (!f)
              if (i[u] === void 0 || i[u] === "*")
                f = !0;
              else {
                const g = Bn(i[u] + ""), m = [kt(e[s][l], jt)].concat(
                  g
                );
                f = zn(m);
              }
        }
      }
      d && (n[l] = n[l] && f);
    }
  const o = [];
  for (let s = 0; s < e[0].length; ++s)
    n[s] && o.push(s - 1);
  return o;
}
function C1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  let s = 0;
  return He(r, (l) => {
    s += o[l];
  }), r.length === 0 ? Ue : s / r.length;
}
function E1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  const s = [];
  return He(r, (l) => {
    s.push(o[l]);
  }), lr(s);
}
function b1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  const s = [];
  return He(r, (l) => {
    s.push(o[l]);
  }), cr(s);
}
function w1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const s = gt(e, t);
    o = Ne(e[s]);
  } else
    o = Ne(e[t]);
  return r.length === 0 ? S : r.length > 1 ? T : o[r[0]];
}
function I1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  let s = o[r[0]];
  return He(r, (l) => {
    s < o[l] && (s = o[l]);
  }), s;
}
function M1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  let s = o[r[0]];
  return He(r, (l) => {
    s > o[l] && (s = o[l]);
  }), s;
}
function y1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const f = gt(e, t);
    o = Ne(e[f]);
  } else
    o = Ne(e[t]);
  let s = [];
  He(r, (f) => {
    s.push(o[f]);
  }), s = Yi(s);
  let l = 1;
  return He(s, (f) => {
    l *= f;
  }), l;
}
function S1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  let s = [];
  return He(r, (l) => {
    s.push(o[l]);
  }), s = Yi(s), Ze.S(s);
}
function R1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  let s = [];
  return He(r, (l) => {
    s.push(o[l]);
  }), s = Yi(s), Ze.P(s);
}
function A1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  const s = [];
  return He(r, (l) => {
    s.push(o[l]);
  }), Pt(s);
}
function T1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  const s = [];
  return He(r, (l) => {
    s.push(o[l]);
  }), qe.S(s);
}
function x1(e, t, n) {
  if (isNaN(t) && typeof t != "string")
    return S;
  const r = pt(e, n);
  let o = [];
  if (typeof t == "string") {
    const l = gt(e, t);
    o = Ne(e[l]);
  } else
    o = Ne(e[t]);
  const s = [];
  return He(r, (l) => {
    s.push(o[l]);
  }), qe.P(s);
}
function So(e) {
  return e && e.getTime && !isNaN(e.getTime());
}
function Ro(e) {
  return e instanceof Date ? e : new Date(e);
}
function D1(e, t, n) {
  let r = ee(t);
  for (r.setFullYear(e.getFullYear()), r < e && r.setFullYear(r.getFullYear() + 1); r > e; )
    r.setMonth(r.getMonth() + -12 / n);
  return r;
}
function gl(e) {
  return e = w(e), [1, 2, 4].indexOf(e) === -1 ? T : e;
}
function pl(e) {
  return e = w(e), [0, 1, 2, 3, 4].indexOf(e) === -1 ? T : e;
}
function O1(e, t, n, r, o, s, l) {
  return e = Ro(e), t = Ro(t), n = Ro(n), s = gl(s), l = pl(l), Q(s, l) ? T : !So(e) || !So(t) || !So(n) ? S : r <= 0 || o <= 0 || n <= e ? T : (o = o || 0, l = l || 0, o * r * Gi(e, n, l));
}
function k1(e, t, n, r) {
  if (r = pl(r), n = gl(n), e = ee(e), t = ee(t), Q(e, t))
    return S;
  if (Q(n, r) || e >= t)
    return T;
  if (r === 1) {
    let s = D1(e, t, n), l = ee(s);
    return l.setMonth(l.getMonth() + 12 / n), Kt(s, l, "D");
  }
  let o;
  switch (r) {
    case 0:
    case 2:
    case 4:
      o = 360;
      break;
    case 3:
      o = 365;
      break;
    default:
      return T;
  }
  return o / n;
}
function L1(e, t, n, r, o, s) {
  if (e = w(e), t = w(t), n = w(n), O(e, t, n))
    return S;
  if (e <= 0 || t <= 0 || n <= 0 || r < 1 || o < 1 || r > o || s !== 0 && s !== 1)
    return T;
  const l = Er(e, t, n, 0, s);
  let f = 0;
  r === 1 && (s === 0 && (f = -n), r++);
  for (let d = r; d <= o; d++)
    f += s === 1 ? gn(e, d - 2, l, n, 1) - l : gn(e, d - 1, l, n, 0);
  return f *= e, f;
}
function F1(e, t, n, r, o, s) {
  if (e = w(e), t = w(t), n = w(n), O(e, t, n))
    return S;
  if (e <= 0 || t <= 0 || n <= 0 || r < 1 || o < 1 || r > o || s !== 0 && s !== 1)
    return T;
  const l = Er(e, t, n, 0, s);
  let f = 0;
  r === 1 && (f = s === 0 ? l + n * e : l, r++);
  for (let d = r; d <= o; d++)
    f += s > 0 ? l - (gn(e, d - 2, l, n, 1) - l) * e : l - gn(e, d - 1, l, n, 0) * e;
  return f;
}
function P1(e, t, n, r, o) {
  if (o = o === void 0 ? 12 : o, e = w(e), t = w(t), n = w(n), r = w(r), o = w(o), O(e, t, n, r, o))
    return S;
  if (e < 0 || t < 0 || n < 0 || r < 0 || [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].indexOf(o) === -1 || r > n)
    return T;
  if (t >= e)
    return 0;
  const s = (1 - Math.pow(t / e, 1 / n)).toFixed(3), l = e * s * o / 12;
  let f = l, d = 0;
  const c = r === n ? n - 1 : r;
  for (let i = 2; i <= c; i++)
    d = (e - f) * s, f += d;
  return r === 1 ? l : r === n ? (e - f) * s : d;
}
function U1(e, t, n, r, o) {
  if (o = o === void 0 ? 2 : o, e = w(e), t = w(t), n = w(n), r = w(r), o = w(o), O(e, t, n, r, o))
    return S;
  if (e < 0 || t < 0 || n < 0 || r < 0 || o <= 0 || r > n)
    return T;
  if (t >= e)
    return 0;
  let s = 0, l = 0;
  for (let f = 1; f <= r; f++)
    l = Math.min((e - s) * (o / n), e - t - s), s += l;
  return l;
}
function _1(e, t, n, r, o) {
  if (e = ee(e), t = ee(t), n = w(n), r = w(r), o = w(o), o = o || 0, O(e, t, n, r, o))
    return S;
  if (n <= 0 || r <= 0)
    return T;
  if (e >= t)
    return S;
  let s, l;
  switch (o) {
    case 0:
      s = 360, l = Qt(e, t, !1);
      break;
    case 1:
      s = 365, l = Kt(e, t, "D");
      break;
    case 2:
      s = 360, l = Kt(e, t, "D");
      break;
    case 3:
      s = 365, l = Kt(e, t, "D");
      break;
    case 4:
      s = 360, l = Qt(e, t, !0);
      break;
    default:
      return T;
  }
  return (r - n) / r * s / l;
}
function B1(e, t) {
  if (e = w(e), t = w(t), O(e, t))
    return S;
  if (t < 0)
    return T;
  if (t >= 0 && t < 1)
    return Ue;
  t = parseInt(t, 10);
  let n = parseInt(e, 10);
  n += e % 1 * Math.pow(10, Math.ceil(Math.log(t) / Math.LN10)) / t;
  const r = Math.pow(10, Math.ceil(Math.log(t) / Math.LN2) + 1);
  return n = Math.round(n * r) / r, n;
}
function z1(e, t) {
  if (e = w(e), t = w(t), O(e, t))
    return S;
  if (t < 0)
    return T;
  if (t >= 0 && t < 1)
    return Ue;
  t = parseInt(t, 10);
  let n = parseInt(e, 10);
  return n += e % 1 * Math.pow(10, -Math.ceil(Math.log(t) / Math.LN10)) * t, n;
}
function H1(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : e <= 0 || t < 1 ? T : (t = parseInt(t, 10), Math.pow(1 + e / t, t) - 1);
}
function gn(e, t, n, r, o) {
  if (r = r || 0, o = o || 0, e = w(e), t = w(t), n = w(n), r = w(r), o = w(o), O(e, t, n, r, o))
    return S;
  let s;
  if (e === 0)
    s = r + n * t;
  else {
    const l = Math.pow(1 + e, t);
    s = o === 1 ? r * l + n * (1 + e) * (l - 1) / e : r * l + n * (l - 1) / e;
  }
  return -s;
}
function V1(e, t) {
  if (e = w(e), t = _(P(t)), O(e, t))
    return S;
  const n = t.length;
  let r = e;
  for (let o = 0; o < n; o++)
    r *= 1 + t[o];
  return r;
}
function vl(e, t, n, r, o, s) {
  if (o = o || 0, s = s || 0, e = w(e), t = w(t), n = w(n), r = w(r), o = w(o), s = w(s), O(e, t, n, r, o, s))
    return S;
  const l = Er(e, n, r, o, s);
  return (t === 1 ? s === 1 ? 0 : -r : s === 1 ? gn(e, t - 2, l, r, 1) - l : gn(e, t - 1, l, r, 0)) * e;
}
function G1(e, t) {
  if (t = typeof t == "number" ? t : typeof t > "u" ? 0.1 : w(t), e = P(e).filter(Ye), e = _(e), O(e, t))
    return S;
  const n = new Float64Array(e.length);
  let r = !1, o = !1;
  for (let c = 0; c < e.length; c++)
    n[c] = e[c], n[c] > 0 && (r = !0), n[c] < 0 && (o = !0);
  if (!r || !o)
    return T;
  const s = (c) => {
    c <= -1 && (c = -0.999999999);
    let i = n[0];
    const a = 1 + c;
    let u = 1;
    for (let h = 1; h < n.length; h++)
      u *= a, i += n[h] / u;
    return i;
  }, l = /* @__PURE__ */ new Map(), f = function(c) {
    const i = Math.round(c * 1e10) / 1e10;
    if (l.has(i))
      return l.get(i);
    const a = s(i);
    return l.set(i, a), a;
  };
  return function() {
    let a = t, u = a, h = 0;
    for (; h < 1e3; ) {
      const b = f(a);
      if (Math.abs(b) < 1e-10)
        return a;
      if (h > 0 && Math.abs(a - u) < 1e-10 * 10)
        break;
      const N = Math.max(1e-4, Math.abs(a * 1e-4)), v = (f(a + N) - b) / N;
      if (Math.abs(v) < 1e-10)
        break;
      u = a;
      const C = b / v, M = Math.max(0.1, Math.abs(a) * 0.5);
      Math.abs(C) > M ? a -= Math.sign(C) * M : a -= C, a <= -1 && (a = -0.99999999), a > 1e3 && (a = 1e3), h++;
    }
    let g = f(a);
    if (Math.abs(g) < 1e-10)
      return a;
    let m, p;
    if (g > 0) {
      for (m = a, p = a + 0.1; f(p) > 0 && p < 1e3; )
        p = p * 2 + 0.1;
      if (p >= 1e3) return a;
    } else {
      for (p = a, m = Math.max(-0.99999999, a - 0.1); f(m) < 0 && m > -0.99999999; )
        m = Math.max(-0.99999999, m - 0.1);
      if (m <= -0.99999999) return a;
    }
    let E;
    for (let b = 0; b < 1e3; b++) {
      E = (m + p) / 2;
      const N = f(E);
      if (Math.abs(N) < 1e-10 || Math.abs(p - m) < 1e-10)
        return E;
      N * f(m) < 0 ? p = E : m = E;
    }
    return E;
  }();
}
function $1(e, t, n, r) {
  return e = w(e), t = w(t), n = w(n), r = w(r), O(e, t, n, r) ? S : r * e * (t / n - 1);
}
function q1(e, t, n) {
  if (e = _(P(e)), t = w(t), n = w(n), O(e, t, n))
    return S;
  const r = e.length, o = [], s = [];
  for (let d = 0; d < r; d++)
    e[d] < 0 ? o.push(e[d]) : s.push(e[d]);
  const l = -oi(n, s) * Math.pow(1 + n, r - 1), f = oi(t, o) * (1 + t);
  return Math.pow(l / f, 1 / (r - 1)) - 1;
}
function W1(e, t) {
  return e = w(e), t = w(t), O(e, t) ? S : e <= 0 || t < 1 ? T : (t = parseInt(t, 10), (Math.pow(e + 1, 1 / t) - 1) * t);
}
function Y1(e, t, n, r, o) {
  if (o = o === void 0 ? 0 : o, r = r === void 0 ? 0 : r, e = w(e), t = w(t), n = w(n), r = w(r), o = w(o), O(e, t, n, r, o))
    return S;
  if (e === 0)
    return -(n + r) / t;
  {
    const s = t * (1 + e * o) - r * e, l = n * e + t * (1 + e * o);
    return Math.log(s / l) / Math.log(1 + e);
  }
}
function oi() {
  const e = _(P(arguments));
  if (e instanceof Error)
    return e;
  const t = e[0];
  let n = 0;
  for (let r = 1; r < e.length; r++)
    n += e[r] / Math.pow(1 + t, r);
  return n;
}
function X1(e, t, n) {
  return e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : e <= 0 ? T : (Math.log(n) - Math.log(t)) / Math.log(1 + e);
}
function Er(e, t, n, r, o) {
  if (r = r || 0, o = o || 0, e = w(e), t = w(t), n = w(n), r = w(r), o = w(o), O(e, t, n, r, o))
    return S;
  let s;
  if (e === 0)
    s = (n + r) / t;
  else {
    const l = Math.pow(1 + e, t);
    s = o === 1 ? (r * e / (l - 1) + n * e / (1 - 1 / l)) / (1 + e) : r * e / (l - 1) + n * e / (1 - 1 / l);
  }
  return -s;
}
function K1(e, t, n, r, o, s) {
  return o = o || 0, s = s || 0, e = w(e), n = w(n), r = w(r), o = w(o), s = w(s), O(e, n, r, o, s) ? S : Er(e, n, r, o, s) - vl(e, t, n, r, o, s);
}
function J1(e, t, n, r, o) {
  if (e = ee(e), t = ee(t), n = w(n), r = w(r), o = w(o), o = o || 0, O(e, t, n, r, o))
    return S;
  if (n <= 0 || r <= 0)
    return T;
  if (e >= t)
    return S;
  let s, l;
  switch (o) {
    case 0:
      s = 360, l = Qt(e, t, !1);
      break;
    case 1:
      s = 365, l = Kt(e, t, "D");
      break;
    case 2:
      s = 360, l = Kt(e, t, "D");
      break;
    case 3:
      s = 365, l = Kt(e, t, "D");
      break;
    case 4:
      s = 360, l = Qt(e, t, !0);
      break;
    default:
      return T;
  }
  return r - n * r * l / s;
}
function Q1(e, t, n, r, o) {
  return r = r || 0, o = o || 0, e = w(e), t = w(t), n = w(n), r = w(r), o = w(o), O(e, t, n, r, o) ? S : e === 0 ? -n * t - r : ((1 - Math.pow(1 + e, t)) / e * n * (1 + e * o) - r) / Math.pow(1 + e, t);
}
function j1(e, t, n, r, o, s) {
  if (s = s === void 0 ? 0.1 : s, r = r === void 0 ? 0 : r, o = o === void 0 ? 0 : o, e = w(e), t = w(t), n = w(n), r = w(r), o = w(o), s = w(s), O(e, t, n, r, o, s))
    return S;
  const l = 1e-10, f = 100;
  let d = s;
  o = o ? 1 : 0;
  for (let c = 0; c < f; c++) {
    if (d <= -1)
      return T;
    let i, a;
    if (Math.abs(d) < l ? i = n * (1 + e * d) + t * (1 + d * o) * e + r : (a = Math.pow(1 + d, e), i = n * a + t * (1 / d + o) * (a - 1) + r), Math.abs(i) < l)
      return d;
    let u;
    if (Math.abs(d) < l)
      u = n * e + t * o * e;
    else {
      a = Math.pow(1 + d, e);
      const h = e * Math.pow(1 + d, e - 1);
      u = n * h + t * (1 / d + o) * h + t * (-1 / (d * d)) * (a - 1);
    }
    d -= i / u;
  }
  return d;
}
function Z1(e, t, n) {
  return e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : e === 0 || t === 0 ? T : Math.pow(n / t, 1 / e) - 1;
}
function eg(e, t, n) {
  return e = w(e), t = w(t), n = w(n), O(e, t, n) ? S : n === 0 ? T : (e - t) / n;
}
function tg(e, t, n, r) {
  return e = w(e), t = w(t), n = w(n), r = w(r), O(e, t, n, r) ? S : n === 0 || r < 1 || r > n ? T : (r = parseInt(r, 10), (e - t) * (n - r + 1) * 2 / (n * (n + 1)));
}
function ng(e, t, n) {
  return e = ee(e), t = ee(t), n = w(n), O(e, t, n) ? S : n <= 0 || e > t || t - e > 365 * 24 * 60 * 60 * 1e3 ? T : 365 * n / (360 - n * Qt(e, t, !1));
}
function rg(e, t, n) {
  return e = ee(e), t = ee(t), n = w(n), O(e, t, n) ? S : n <= 0 || e > t || t - e > 365 * 24 * 60 * 60 * 1e3 ? T : 100 * (1 - n * Qt(e, t, !1) / 360);
}
function og(e, t, n) {
  return e = ee(e), t = ee(t), n = w(n), O(e, t, n) ? S : n <= 0 || e > t || t - e > 365 * 24 * 60 * 60 * 1e3 ? T : (100 - n) * 360 / (n * Qt(e, t, !1));
}
function ig(e, t, n) {
  if (e = _(P(e)), t = Ga(P(t)), n = w(n), O(e, t, n))
    return S;
  const r = (h, g, m) => {
    const p = m + 1;
    let E = h[0];
    for (let b = 1; b < h.length; b++)
      E += h[b] / Math.pow(p, an(g[b], g[0]) / 365);
    return E;
  }, o = (h, g, m) => {
    const p = m + 1;
    let E = 0;
    for (let b = 1; b < h.length; b++) {
      const N = an(g[b], g[0]) / 365;
      E -= N * h[b] / Math.pow(p, N + 1);
    }
    return E;
  };
  let s = !1, l = !1;
  for (let h = 0; h < e.length; h++)
    e[h] > 0 && (s = !0), e[h] < 0 && (l = !0);
  if (!s || !l)
    return T;
  n = n || 0.1;
  let f = n;
  const d = 1e-10;
  let c, i, a, u = !0;
  do
    a = r(e, t, f), c = f - a / o(e, t, f), i = Math.abs(c - f), f = c, u = i > d && Math.abs(a) > d;
  while (u);
  return f;
}
function sg(e, t, n) {
  if (e = w(e), t = _(P(t)), n = Ga(P(n)), O(e, t, n))
    return S;
  let r = 0;
  for (let o = 0; o < t.length; o++)
    r += t[o] / Math.pow(1 + e, an(n[o], n[0]) / 365);
  return r;
}
function ag() {
  const e = P(arguments);
  let t = S;
  for (let n = 0; n < e.length; n++) {
    if (e[n] instanceof Error)
      return e[n];
    e[n] === void 0 || e[n] === null || typeof e[n] == "string" || (t === S && (t = !0), e[n] || (t = !1));
  }
  return t;
}
function lg() {
  return !1;
}
function cg(e, t, n) {
  return e instanceof Error ? e : (t = arguments.length >= 2 ? t : !0, t == null && (t = 0), n = arguments.length === 3 ? n : !1, n == null && (n = 0), e ? t : n);
}
function ug() {
  for (let e = 0; e < arguments.length / 2; e++)
    if (arguments[e * 2])
      return arguments[e * 2 + 1];
  return G;
}
function fg(e, t) {
  return fo(e) ? t : e;
}
function dg(e, t) {
  return e === G ? t : e;
}
function hg(e) {
  return typeof e == "string" ? S : e instanceof Error ? e : !e;
}
function mg() {
  const e = P(arguments);
  let t = S;
  for (let n = 0; n < e.length; n++) {
    if (e[n] instanceof Error)
      return e[n];
    e[n] === void 0 || e[n] === null || typeof e[n] == "string" || (t === S && (t = !1), e[n] && (t = !0));
  }
  return t;
}
function gg() {
  return !0;
}
function pg() {
  const e = P(arguments);
  let t = S;
  for (let n = 0; n < e.length; n++) {
    if (e[n] instanceof Error)
      return e[n];
    e[n] === void 0 || e[n] === null || typeof e[n] == "string" || (t === S && (t = 0), e[n] && t++);
  }
  return t === S ? t : !!(Math.floor(Math.abs(t)) & 1);
}
function vg() {
  let e;
  if (arguments.length > 0) {
    const t = arguments[0], n = arguments.length - 1, r = Math.floor(n / 2);
    let o = !1;
    const s = n % 2 !== 0, l = n % 2 === 0 ? null : arguments[arguments.length - 1];
    if (r) {
      for (let f = 0; f < r; f++)
        if (t === arguments[f * 2 + 1]) {
          e = arguments[f * 2 + 2], o = !0;
          break;
        }
    }
    o || (e = s ? l : G);
  } else
    e = S;
  return e;
}
const Nl = { errors: Qu, symbols: lm, date: ef }, Ng = /* @__PURE__ */ Object.freeze(/* @__PURE__ */ Object.defineProperty({
  __proto__: null,
  ABS: Hd,
  ACCRINT: O1,
  ACOS: Vd,
  ACOSH: Gd,
  ACOT: $d,
  ACOTH: qd,
  AGGREGATE: Wd,
  AND: ag,
  ARABIC: Yd,
  ASIN: Xd,
  ASINH: Kd,
  ATAN: Jd,
  ATAN2: Qd,
  ATANH: jd,
  AVEDEV: ad,
  AVERAGE: fn,
  AVERAGEA: Di,
  AVERAGEIF: ld,
  AVERAGEIFS: cd,
  BASE: Zd,
  BESSELI: Rm,
  BESSELJ: Am,
  BESSELK: Tm,
  BESSELY: xm,
  BETA: pr,
  BETADIST: h0,
  BETAINV: m0,
  BIN2DEC: Dm,
  BIN2HEX: Om,
  BIN2OCT: km,
  BINOM: Hn,
  BINOMDIST: g0,
  BITAND: Lm,
  BITLSHIFT: Fm,
  BITOR: Pm,
  BITRSHIFT: Um,
  BITXOR: _m,
  CEILING: Jt,
  CEILINGMATH: p0,
  CEILINGPRECISE: v0,
  CHAR: Ka,
  CHIDIST: N0,
  CHIDISTRT: C0,
  CHIINV: E0,
  CHIINVRT: b0,
  CHISQ: It,
  CHITEST: w0,
  CHOOSE: wf,
  CHOOSECOLS: If,
  CHOOSEROWS: Mf,
  CLEAN: _f,
  CODE: Ja,
  COLUMN: yf,
  COLUMNS: Sf,
  COMBIN: kn,
  COMBINA: eh,
  COMPLEX: Ve,
  CONCAT: Bf,
  CONCATENATE: Qa,
  CONFIDENCE: Oi,
  CONVERT: Bm,
  CORREL: ud,
  COS: th,
  COSH: nh,
  COT: rh,
  COTH: oh,
  COUNT: lr,
  COUNTA: cr,
  COUNTBLANK: Za,
  COUNTIF: fd,
  COUNTIFS: dd,
  COUPDAYS: k1,
  COVAR: I0,
  COVARIANCE: Vn,
  COVARIANCEP: M0,
  COVARIANCES: y0,
  CRITBINOM: S0,
  CSC: ih,
  CSCH: sh,
  CUMIPMT: L1,
  CUMPRINC: F1,
  DATE: fm,
  DATEDIF: Kt,
  DATEVALUE: dm,
  DAVERAGE: C1,
  DAY: hm,
  DAYS: an,
  DAYS360: Qt,
  DB: P1,
  DCOUNT: E1,
  DCOUNTA: b1,
  DDB: U1,
  DEC2BIN: zm,
  DEC2HEX: Hm,
  DEC2OCT: Vm,
  DECIMAL: ah,
  DEGREES: lh,
  DELTA: Gm,
  DEVSQ: hd,
  DGET: w1,
  DISC: _1,
  DMAX: I1,
  DMIN: M1,
  DOLLAR: zf,
  DOLLARDE: B1,
  DOLLARFR: z1,
  DPRODUCT: y1,
  DROP: Rf,
  DSTDEV: S1,
  DSTDEVP: R1,
  DSUM: A1,
  DVAR: T1,
  DVARP: x1,
  EDATE: mm,
  EFFECT: H1,
  EOMONTH: gm,
  ERF: fl,
  ERFC: dl,
  ERFCPRECISE: R0,
  ERFPRECISE: A0,
  ERROR: $a,
  EVEN: ch,
  EXACT: Hf,
  EXP: uh,
  EXPAND: Af,
  EXPON: ki,
  EXPONDIST: T0,
  F: Mt,
  FACT: Lt,
  FACTDOUBLE: al,
  FALSE: lg,
  FDIST: x0,
  FDISTRT: D0,
  FIND: Vf,
  FINV: O0,
  FINVRT: k0,
  FISHER: md,
  FISHERINV: gd,
  FIXED: ja,
  FLOOR: mn,
  FLOORMATH: L0,
  FLOORPRECISE: F0,
  FORECAST: el,
  FREQUENCY: pd,
  FTEST: P0,
  FV: gn,
  FVSCHEDULE: V1,
  GAMMA: vr,
  GAMMADIST: U0,
  GAMMAINV: _0,
  GAMMALN: Li,
  GAMMALNPRECISE: B0,
  GAUSS: vd,
  GCD: fh,
  GEOMEAN: Nd,
  GESTEP: $m,
  GROWTH: Cd,
  HARMEAN: Ed,
  HEX2BIN: qm,
  HEX2DEC: Wm,
  HEX2OCT: Ym,
  HLOOKUP: Tf,
  HOUR: pm,
  HSTACK: Pf,
  HYPGEOM: Fi,
  HYPGEOMDIST: z0,
  IF: cg,
  IFERROR: fg,
  IFNA: dg,
  IFS: ug,
  IMABS: qi,
  IMAGINARY: me,
  IMARGUMENT: Wi,
  IMCONJUGATE: Xm,
  IMCOS: po,
  IMCOSH: hl,
  IMCOT: Km,
  IMCSC: i0,
  IMCSCH: s0,
  IMDIV: En,
  IMEXP: Jm,
  IMLN: Qm,
  IMLOG10: jm,
  IMLOG2: Zm,
  IMPOWER: e0,
  IMPRODUCT: t0,
  IMREAL: ge,
  IMSEC: n0,
  IMSECH: r0,
  IMSIN: vo,
  IMSINH: ml,
  IMSQRT: o0,
  IMSUB: a0,
  IMSUM: l0,
  IMTAN: c0,
  INDEX: xf,
  INT: dh,
  INTERCEPT: bd,
  IPMT: vl,
  IRR: G1,
  ISBLANK: mf,
  ISERR: qa,
  ISERROR: fo,
  ISEVEN: gf,
  ISLOGICAL: Wa,
  ISNA: pf,
  ISNONTEXT: vf,
  ISNUMBER: ho,
  ISO: hh,
  ISODD: Nf,
  ISOWEEKNUM: ul,
  ISPMT: $1,
  ISTEXT: Ya,
  KURT: wd,
  LARGE: tl,
  LCM: mh,
  LEFT: Gf,
  LEN: $f,
  LINEST: Pi,
  LN: gh,
  LOG: ph,
  LOG10: vh,
  LOGEST: Id,
  LOGINV: H0,
  LOGNORM: Gn,
  LOGNORMDIST: V0,
  LOGNORMINV: G0,
  LOOKUP: Df,
  LOWER: qf,
  MATCH: Of,
  MAX: Gr,
  MAXA: Md,
  MAXIFS: yd,
  MEDIAN: nl,
  MID: Wf,
  MIN: $r,
  MINA: Sd,
  MINIFS: Rd,
  MINUTE: vm,
  MIRR: q1,
  MMULT: Nh,
  MOD: Ch,
  MODE: dn,
  MODEMULT: $0,
  MODESNGL: q0,
  MONTH: Nm,
  MROUND: Eh,
  MULTINOMIAL: bh,
  MUNIT: wh,
  N: Cf,
  NA: Ef,
  NEGBINOM: Ui,
  NEGBINOMDIST: W0,
  NETWORKDAYS: mo,
  NETWORKDAYSINTL: Y0,
  NOMINAL: W1,
  NORM: yt,
  NORMDIST: X0,
  NORMINV: K0,
  NORMSDIST: J0,
  NORMSINV: Q0,
  NOT: hg,
  NOW: Cm,
  NPER: Y1,
  NPV: oi,
  NUMBERVALUE: Yf,
  OCT2BIN: u0,
  OCT2DEC: f0,
  OCT2HEX: d0,
  ODD: Ih,
  OR: mg,
  PDURATION: X1,
  PEARSON: rl,
  PERCENTILE: st,
  PERCENTILEEXC: j0,
  PERCENTILEINC: Z0,
  PERCENTRANK: Nr,
  PERCENTRANKEXC: e1,
  PERCENTRANKINC: t1,
  PERMUT: Ad,
  PERMUTATIONA: Td,
  PHI: xd,
  PI: Mh,
  PMT: Er,
  POISSON: _i,
  POISSONDIST: n1,
  POWER: ll,
  PPMT: K1,
  PRICEDISC: J1,
  PROB: Dd,
  PRODUCT: qr,
  PROPER: Xf,
  PV: Q1,
  QUARTILE: hn,
  QUARTILEEXC: r1,
  QUARTILEINC: o1,
  QUOTIENT: yh,
  RADIANS: Sh,
  RAND: Rh,
  RANDBETWEEN: Ah,
  RANK: Cr,
  RANKAVG: i1,
  RANKEQ: s1,
  RATE: j1,
  REPLACE: Kf,
  REPT: at,
  RIGHT: Jf,
  ROMAN: Th,
  ROUND: cl,
  ROUNDDOWN: xh,
  ROUNDUP: Dh,
  ROW: Od,
  ROWS: kf,
  RRI: Z1,
  RSQ: kd,
  SEARCH: Qf,
  SEC: Oh,
  SECH: kh,
  SECOND: Em,
  SERIESSUM: Lh,
  SIGN: Fh,
  SIN: Ph,
  SINH: Uh,
  SKEW: Bi,
  SKEWP: a1,
  SLN: eg,
  SLOPE: Ld,
  SMALL: ol,
  SORT: Lf,
  SQRT: _h,
  SQRTPI: Bh,
  STANDARDIZE: Fd,
  STDEV: Ze,
  STDEVA: Pd,
  STDEVP: l1,
  STDEVPA: Ud,
  STDEVS: c1,
  STEYX: _d,
  SUBSTITUTE: jf,
  SUBTOTAL: zh,
  SUM: Pt,
  SUMIF: Hh,
  SUMIFS: Vh,
  SUMPRODUCT: Gh,
  SUMSQ: $h,
  SUMX2MY2: qh,
  SUMX2PY2: Wh,
  SUMXMY2: Yh,
  SWITCH: vg,
  SYD: tg,
  T: et,
  TAN: Xh,
  TANH: Kh,
  TBILLEQ: ng,
  TBILLPRICE: rg,
  TBILLYIELD: og,
  TDIST: u1,
  TDISTRT: f1,
  TEXT: Zf,
  TEXTJOIN: ed,
  TIME: bm,
  TIMEVALUE: wm,
  TINV: d1,
  TODAY: Im,
  TRANSPOSE: Ff,
  TREND: Bd,
  TRIM: td,
  TRIMMEAN: zd,
  TRUE: gg,
  TRUNC: Jh,
  TTEST: h1,
  TYPE: bf,
  UNICHAR: nd,
  UNICODE: rd,
  UNIQUE: xi,
  UPPER: od,
  VALUE: id,
  VAR: qe,
  VARA: il,
  VARP: m1,
  VARPA: sl,
  VARS: g1,
  VLOOKUP: Xa,
  VSTACK: Uf,
  WEEKDAY: Mm,
  WEEKNUM: ym,
  WEIBULL: zi,
  WEIBULLDIST: p1,
  WORKDAY: go,
  WORKDAYINTL: v1,
  XIRR: ig,
  XNPV: sg,
  XOR: pg,
  YEAR: Sm,
  YEARFRAC: Gi,
  Z: Hi,
  ZTEST: N1,
  utils: Nl
}, Symbol.toStringTag, { value: "Module" })), Cl = (e) => {
  const t = Ng[e];
  if (!t)
    throw fe("#NAME?");
  return t;
}, fe = (e) => {
  const { errors: t } = Nl;
  return {
    "#GETTING_DATA": t.data,
    "#DIV/0!": t.div0,
    "#NULL!": t.nil,
    "#NUM!": t.num,
    "#REF!": t.ref,
    "#VALUE!": t.value,
    "#N/A": t.na,
    "#NAME?": t.name
  }[e];
}, No = (e) => {
  let t;
  return e instanceof Error && hs.has(e.message) && (t = e.message), e instanceof String && hs.has(e) && (t = e), t;
}, Cg = "_buttonWrapper_16zfq_1", Eg = "_circle_16zfq_15", bg = "_disabled_16zfq_18", wg = "_active_16zfq_23", Ig = "_plain_16zfq_26", Mg = "_primary_16zfq_30", wn = {
  buttonWrapper: Cg,
  circle: Eg,
  disabled: bg,
  active: wg,
  plain: Ig,
  primary: Mg
}, V = ue((e) => {
  const {
    className: t = "",
    onClick: n = fu,
    disabled: r = !1,
    active: o = !1,
    type: s = "normal",
    style: l,
    testId: f,
    title: d,
    dataType: c,
    buttonType: i,
    children: a
  } = e, u = ae(wn.buttonWrapper, t, {
    [wn.disabled]: r,
    [wn.active]: o,
    [wn.circle]: s === "circle",
    [wn.plain]: s === "plain",
    [wn.primary]: s === "primary"
  });
  return /* @__PURE__ */ I(
    "button",
    {
      onClick: n,
      style: l,
      title: d,
      disabled: r,
      className: u,
      "data-testid": f,
      "data-type": c,
      type: i,
      children: a
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/components/Button/index.tsx",
      lineNumber: 41,
      columnNumber: 7
    },
    void 0
  );
});
V.displayName = "Button";
const yg = {
  alignCenter: [
    "M142.2 227.6h739.6v56.9H142.2zM142.2 568.9h739.6v56.9H142.2zM256 398.2h512v56.9H256zM256 739.6h512v56.9H256z"
  ],
  alignLeft: [
    "M627.712 788.48c11.305 0 20.48-9.155 20.48-20.48s-9.175-20.48-20.48-20.48H218.46c-11.305 0-20.48 9.155-20.48 20.48s9.175 20.48 20.48 20.48h409.252zM832.86 583.68h-614.4c-11.305 0-20.48 9.155-20.48 20.48s9.175 20.48 20.48 20.48h614.4c11.305 0 20.48-9.155 20.48-20.48s-9.175-20.48-20.48-20.48zM832.86 256h-614.4c-11.305 0-20.48 9.155-20.48 20.48s9.175 20.48 20.48 20.48h614.4c11.305 0 20.48-9.155 20.48-20.48S844.165 256 832.86 256zM218.46 460.8h409.252c11.305 0 20.48-9.155 20.48-20.48s-9.175-20.48-20.48-20.48H218.46c-11.305 0-20.48 9.155-20.48 20.48s9.155 20.48 20.48 20.48z"
  ],
  alignRight: [
    "M832.86 747.52H423.588a20.48 20.48 0 1 0 0 40.96H832.84a20.48 20.48 0 0 0 0.02-40.96z m20.48-143.36a20.48 20.48 0 0 0-20.48-20.48h-614.4a20.48 20.48 0 1 0 0 40.96h614.4a20.48 20.48 0 0 0 20.48-20.48zM832.86 256h-614.4a20.48 20.48 0 1 0 0 40.96h614.4a20.48 20.48 0 1 0 0-40.96z m0 163.84H423.588a20.48 20.48 0 1 0 0 40.96H832.84a20.48 20.48 0 0 0 0.02-40.96z"
  ],
  bold: [
    "M724.342857 477.028571c38.4-40 61.942857-94.057143 61.942857-153.485714v-11.657143C786.285714 188.914286 685.6 89.142857 561.485714 89.142857H223.314286c-17.257143 0-31.314286 14.057143-31.314286 31.314286v776.114286c0 18.628571 15.085714 33.714286 33.714286 33.714285h364.228571c133.714286 0 242.057143-107.657143 242.057143-240.571428v-12.571429c0-83.428571-42.742857-156.914286-107.657143-200.114286zM301.714286 198.857143h256.8c65.257143 0 118.057143 50.742857 118.057143 113.485714v10.857143c0 62.628571-52.914286 113.485714-118.057143 113.485714H301.714286V198.857143z m418.971428 490.742857c0 71.885714-59.085714 130.171429-132 130.171429H301.714286V547.085714h286.971428c72.914286 0 132 58.285714 132 130.171429v12.342857z"
  ],
  fontColor: [
    "M650.496 597.333333H373.504l-68.266667 170.666667H213.333333l256-640h85.333334l256 640h-91.904l-68.266667-170.666667z m-34.133333-85.333333L512 251.093333 407.637333 512h208.725334zM128 853.333333h768v85.333334H128v-85.333334z"
  ],
  italic: [
    "M219.428571 949.714286l9.714286-48.571429q3.428571-1.142857 46.571429-12.285714t63.714286-21.428571q16-20 23.428571-57.714286 0.571429-4 35.428571-165.142857t65.142857-310.571429 29.714286-169.428571l0-14.285714q-13.714286-7.428571-31.142857-10.571429t-39.714286-4.571429-33.142857-3.142857l10.857143-58.857143q18.857143 1.142857 68.571429 3.714286t85.428571 4 68.857143 1.428571q27.428571 0 56.285714-1.428571t69.142857-4 56.285714-3.714286q-2.857143 22.285714-10.857143 50.857143-17.142857 5.714286-58 16.285714t-62 19.142857q-4.571429 10.857143-8 24.285714t-5.142857 22.857143-4.285714 26-3.714286 24q-15.428571 84.571429-50 239.714286t-44.285714 203.142857q-1.142857 5.142857-7.428571 33.142857t-11.428571 51.428571-9.142857 47.714286-3.428571 32.857143l0.571429 10.285714q9.714286 2.285714 105.714286 17.714286-1.714286 25.142857-9.142857 56.571429-6.285714 0-18.571429 0.857143t-18.571429 0.857143q-16.571429 0-49.714286-5.714286t-49.142857-5.714286q-78.857143-1.142857-117.714286-1.142857-29.142857 0-81.714286 5.142857t-69.142857 6.285714z"
  ],
  middleAlign: [
    "M740.43392 788.48c11.30496 0 20.48-9.15456 20.48-20.48s-9.17504-20.48-20.48-20.48L331.18208 747.52c-11.30496 0-20.48 9.15456-20.48 20.48s9.17504 20.48 20.48 20.48L740.43392 788.48zM863.49824 604.16c0-11.32544-9.17504-20.48-20.48-20.48l-614.4 0c-11.30496 0-20.48 9.15456-20.48 20.48s9.17504 20.48 20.48 20.48l614.4 0C854.3232 624.64 863.49824 615.48544 863.49824 604.16zM208.13824 276.48c0 11.32544 9.17504 20.48 20.48 20.48l614.4 0c11.30496 0 20.48-9.15456 20.48-20.48s-9.17504-20.48-20.48-20.48l-614.4 0C217.2928 256 208.13824 265.15456 208.13824 276.48zM740.43392 460.8c11.30496 0 20.48-9.15456 20.48-20.48s-9.17504-20.48-20.48-20.48L331.18208 419.84c-11.30496 0-20.48 9.15456-20.48 20.48s9.17504 20.48 20.48 20.48L740.43392 460.8z"
  ],
  plus: [
    "M896 468.571429H555.428571V100.571429h-86.857142v368H128c-5.028571 0-9.142857 4.114286-9.142857 9.142857v68.571428c0 5.028571 4.114286 9.142857 9.142857 9.142857h340.571429v368h86.857142V555.428571h340.571429c5.028571 0 9.142857-4.114286 9.142857-9.142857v-68.571428c0-5.028571-4.114286-9.142857-9.142857-9.142857z"
  ],
  redo: [
    "M611.783111 569.344L549.622519 644.740741h284.444444l-65.498074-265.481482-59.922963 72.666074c-35.422815-28.48237-108.278519-68.342519-238.667852-68.342518-202.827852 0-280.651852 206.01363-280.651852 206.013629s116.318815-132.778667 246.215111-132.778666c97.204148-0.037926 153.865481 74.827852 176.241778 112.526222z"
  ],
  underline: [
    "M512 725.333333c166.4 0 298.666667-132.266667 298.666667-298.666666V128c0-25.6-17.066667-42.666667-42.666667-42.666667s-42.666667 17.066667-42.666667 42.666667v298.666667c0 119.466667-93.866667 213.333333-213.333333 213.333333s-213.333333-93.866667-213.333333-213.333333V128c0-25.6-17.066667-42.666667-42.666667-42.666667s-42.666667 17.066667-42.666667 42.666667v298.666667c0 166.4 132.266667 298.666667 298.666667 298.666666zM853.333333 853.333333H170.666667c-25.6 0-42.666667 17.066667-42.666667 42.666667s17.066667 42.666667 42.666667 42.666667h682.666666c25.6 0 42.666667-17.066667 42.666667-42.666667s-17.066667-42.666667-42.666667-42.666667z"
  ],
  undo: [
    "M489.244444 568.888889l60.681482 75.851852H265.481481l64.474075-265.481482 60.681481 72.05926c34.133333-30.340741 109.985185-68.266667 238.933333-68.266667 201.007407 0 280.651852 204.8 280.651852 204.8S792.651852 455.111111 663.703704 455.111111c-98.607407 0-155.496296 75.851852-174.45926 113.777778z"
  ],
  success: [
    "M666.272 472.288l-175.616 192a31.904 31.904 0 0 1-23.616 10.4h-0.192a32 32 0 0 1-23.68-10.688l-85.728-96a32 32 0 1 1 47.744-42.624l62.144 69.6 151.712-165.888a32 32 0 1 1 47.232 43.2m-154.24-344.32C300.224 128 128 300.32 128 512c0 211.776 172.224 384 384 384 211.68 0 384-172.224 384-384 0-211.68-172.32-384-384-384"
  ],
  info: [
    "M512 97.52381c228.912762 0 414.47619 185.563429 414.47619 414.47619s-185.563429 414.47619-414.47619 414.47619S97.52381 740.912762 97.52381 512 283.087238 97.52381 512 97.52381z m36.571429 341.333333h-73.142858v292.571428h73.142858V438.857143z m0-121.904762h-73.142858v73.142857h73.142858v-73.142857z"
  ],
  warning: [
    "M512 64q190.016 4.992 316.512 131.488T960 512q-4.992 190.016-131.488 316.512T512 960q-190.016-4.992-316.512-131.488T64 512q4.992-190.016 131.488-316.512T512 64z m0 192q-26.016 0-43.008 19.008T453.984 320l23.008 256q2.016 14.016 11.488 22.496t23.488 8.512 23.488-8.512 11.488-22.496l23.008-256q2.016-26.016-15.008-44.992T511.936 256z m0 512q22.016-0.992 36.512-15.008t14.496-36-14.496-36.512T512 665.984t-36.512 14.496-14.496 36.512 14.496 36T512 768z"
  ],
  error: [
    "M512 97.52381c228.912762 0 414.47619 185.563429 414.47619 414.47619s-185.563429 414.47619-414.47619 414.47619S97.52381 740.912762 97.52381 512 283.087238 97.52381 512 97.52381z m129.29219 233.447619l-129.267809 129.29219-129.316571-129.29219-51.736381 51.736381 129.316571 129.267809-129.316571 129.316571 51.736381 51.736381L512 563.687619l129.29219 129.316571 51.736381-51.73638L563.687619 512l129.316571-129.29219-51.73638-51.736381z"
  ],
  rotate: [
    "M934.08 416.448h-220.352a46.4 46.4 0 0 1-47.232-47.232c0-26.752 20.48-47.232 47.232-47.232h173.184V148.8c0-26.752 20.48-47.168 47.168-47.168 26.816 0 47.232 20.48 47.232 47.168V369.28a47.36 47.36 0 0 1-47.232 47.232z",
    "M509.056 978.432A470.976 470.976 0 0 1 38.4 507.712 470.976 470.976 0 0 1 509.056 37.12a472.32 472.32 0 0 1 434.56 288.064c9.408 23.68-1.6 51.968-25.216 61.44-23.68 9.408-51.968-1.6-61.44-25.216a377.92 377.92 0 0 0-347.904-229.824 376.192 376.192 0 1 0 0 752.448c190.528 0 351.104-141.696 373.12-328.96a47.36 47.36 0 0 1 51.904-40.96 47.36 47.36 0 0 1 40.96 51.904 469.504 469.504 0 0 1-465.92 412.48z"
  ],
  moon: [
    "M735.996 244.2C693.136 221.106 644.1 208 592 208c-167.894 0-304 136.106-304 304s136.106 304 304 304c52.1 0 101.136-13.106 143.996-36.2C621.368 744.742 538 638.108 538 512s83.368-232.742 197.996-267.8z m26.32 86.064C683.17 354.472 628 427.888 628 512s55.168 157.528 134.318 181.736c76.442 23.378 86.74 127.376 16.37 165.294C721.764 889.702 657.958 906 592 906c-217.6 0-394-176.4-394-394S374.4 118 592 118c65.96 0 129.764 16.298 186.688 46.97 70.372 37.918 60.072 141.916-16.37 165.294z"
  ],
  sun: [
    "M512 768c-140.8 0-256-115.2-256-256s115.2-256 256-256 256 115.2 256 256-115.2 256-256 256z m0-426.666667c-93.866667 0-170.666667 76.8-170.666667 170.666667s76.8 170.666667 170.666667 170.666667 170.666667-76.8 170.666667-170.666667-76.8-170.666667-170.666667-170.666667zM512 170.666667c-25.6 0-42.666667-17.066667-42.666667-42.666667V42.666667c0-25.6 17.066667-42.666667 42.666667-42.666667s42.666667 17.066667 42.666667 42.666667v85.333333c0 25.6-17.066667 42.666667-42.666667 42.666667zM512 1024c-25.6 0-42.666667-17.066667-42.666667-42.666667v-85.333333c0-25.6 17.066667-42.666667 42.666667-42.666667s42.666667 17.066667 42.666667 42.666667v85.333333c0 25.6-17.066667 42.666667-42.666667 42.666667zM238.933333 281.6c-12.8 0-21.333333-4.266667-29.866666-12.8L149.333333 209.066667c-17.066667-17.066667-17.066667-42.666667 0-59.733334s42.666667-17.066667 59.733334 0l59.733333 59.733334c17.066667 17.066667 17.066667 42.666667 0 59.733333-4.266667 8.533333-17.066667 12.8-29.866667 12.8zM844.8 887.466667c-12.8 0-21.333333-4.266667-29.866667-12.8l-59.733333-59.733334c-17.066667-17.066667-17.066667-42.666667 0-59.733333s42.666667-17.066667 59.733333 0l59.733334 59.733333c17.066667 17.066667 17.066667 42.666667 0 59.733334-8.533333 8.533333-21.333333 12.8-29.866667 12.8zM128 554.666667H42.666667c-25.6 0-42.666667-17.066667-42.666667-42.666667s17.066667-42.666667 42.666667-42.666667h85.333333c25.6 0 42.666667 17.066667 42.666667 42.666667s-17.066667 42.666667-42.666667 42.666667zM981.333333 554.666667h-85.333333c-25.6 0-42.666667-17.066667-42.666667-42.666667s17.066667-42.666667 42.666667-42.666667h85.333333c25.6 0 42.666667 17.066667 42.666667 42.666667s-17.066667 42.666667-42.666667 42.666667zM179.2 887.466667c-12.8 0-21.333333-4.266667-29.866667-12.8-17.066667-17.066667-17.066667-42.666667 0-59.733334l59.733334-59.733333c17.066667-17.066667 42.666667-17.066667 59.733333 0s17.066667 42.666667 0 59.733333l-59.733333 59.733334c-8.533333 8.533333-17.066667 12.8-29.866667 12.8zM785.066667 281.6c-12.8 0-21.333333-4.266667-29.866667-12.8-17.066667-17.066667-17.066667-42.666667 0-59.733333l59.733333-59.733334c17.066667-17.066667 42.666667-17.066667 59.733334 0s17.066667 42.666667 0 59.733334l-59.733334 59.733333c-8.533333 8.533333-21.333333 12.8-29.866666 12.8z"
  ],
  down: [
    "M512.146286 619.52L245.296762 352.792381 193.584762 404.48l318.585905 318.415238 318.268952-318.415238-51.736381-51.687619z"
  ],
  menu: [
    "M904 160H120c-4.4 0-8 3.6-8 8v64c0 4.4 3.6 8 8 8h784c4.4 0 8-3.6 8-8v-64c0-4.4-3.6-8-8-8zM904 784H120c-4.4 0-8 3.6-8 8v64c0 4.4 3.6 8 8 8h784c4.4 0 8-3.6 8-8v-64c0-4.4-3.6-8-8-8zM904 472H120c-4.4 0-8 3.6-8 8v64c0 4.4 3.6 8 8 8h784c4.4 0 8-3.6 8-8v-64c0-4.4-3.6-8-8-8z"
  ],
  confirm: [
    "M394.792 797.429c-9.71 0-19.443-3.513-27.137-10.622L87.488 527.917c-16.225-14.993-17.224-40.3-2.231-56.525 14.993-16.226 40.3-17.225 56.524-2.231l280.167 258.89c16.225 14.993 17.224 40.3 2.231 56.524-7.883 8.532-18.62 12.854-29.387 12.854z",
    "M394.808 797.429c-10.555 0-21.088-4.15-28.949-12.391-15.249-15.984-14.652-41.304 1.333-56.553l514.564-490.858c15.983-15.249 41.303-14.652 56.553 1.333 15.248 15.985 14.651 41.305-1.334 56.553l-514.563 490.86c-7.745 7.387-17.684 11.056-27.604 11.056z"
  ],
  horizontalLeft: [
    "M484.693 477.013h-371.2a32 32 0 0 0-32 32v32a32 32 0 0 0 32 32h371.2a32 32 0 0 0 32-32v-32a32 32 0 0 0-32-32z m435.2-338.133h-806.4a32 32 0 0 0-32 31.787v32a32 32 0 0 0 32 32h806.4a32 32 0 0 0 32-32v-32a32 32 0 0 0-32-32z m0 674.133h-806.4a32 32 0 0 0-32 32v32a32 32 0 0 0 32 32h806.4a32 32 0 0 0 32-32v-32a32 32 0 0 0-32-32z"
  ],
  horizontalRight: [
    "M922.453 472.107H551.04a32 32 0 0 0-32 32v32a32 32 0 0 0 32 32h371.413a32 32 0 0 0 32-32v-32a32 32 0 0 0-32-32z m0 337.066H115.84a32 32 0 0 0-32 32v32a32 32 0 0 0 32 32h806.613a32 32 0 0 0 32-32v-32a32 32 0 0 0-32-32z m0-674.133H115.84a32 32 0 0 0-32 32v32a32 32 0 0 0 32 32h806.613a32 32 0 0 0 32-32v-32a32 32 0 0 0-32-32z"
  ],
  horizontalCenter: [
    "M697.6 472.107H326.4a32 32 0 0 0-32 32v32a32 32 0 0 0 32 32h371.2a32 32 0 0 0 32-32v-32a32 32 0 0 0-32-32zM915.2 135.04H108.8a32 32 0 0 0-32 32v32a32 32 0 0 0 32 32h806.4a32 32 0 0 0 32-32v-32a32 32 0 0 0-32-32z m0 674.133H108.8a32 32 0 0 0-32 32v32a32 32 0 0 0 32 32h806.4a32 32 0 0 0 32-32v-32a32 32 0 0 0-32-32z"
  ],
  verticalTop: [
    "M902.890667 125.482667H121.109333a9.045333 9.045333 0 0 0-9.109333 8.981333v67.413333c0 4.949333 4.053333 9.002667 9.109333 9.002667h781.781334c5.056 0 9.109333-4.053333 9.109333-9.002667v-67.413333a9.045333 9.045333 0 0 0-9.109333-8.981333z m-383.808 210.112a8.96 8.96 0 0 0-14.165334 0l-125.845333 159.210666a8.96 8.96 0 0 0 7.082667 14.506667h83.029333v380.224c0 4.928 4.053333 8.981333 9.002667 8.981333h67.413333c4.949333 0 8.981333-4.053333 8.981333-8.981333V509.418667H637.866667a8.96 8.96 0 0 0 7.082666-14.506667l-125.866666-159.317333z"
  ],
  verticalMiddle: [
    "M512 614.4l117.028571 153.6h-73.142857V950.857143H468.114286v-182.857143h-73.142857L512 614.4z m438.857143-153.6v102.4H73.142857V460.8h877.714286zM555.885714 73.142857v182.857143h73.142857L512 409.6 394.971429 256h73.142857V73.142857h87.771428z"
  ],
  verticalBottom: [
    "M902.890667 813.12H121.109333a9.045333 9.045333 0 0 0-9.109333 9.002667v67.413333c0 4.928 4.053333 8.981333 9.109333 8.981333h781.781334c5.056 0 9.109333-4.053333 9.109333-8.981333v-67.413333a9.045333 9.045333 0 0 0-9.109333-9.002667zM504.917333 688.426667a8.96 8.96 0 0 0 14.165334 0l125.845333-159.210667a8.96 8.96 0 0 0-7.082667-14.506667h-83.264V134.464a9.024 9.024 0 0 0-8.981333-8.981333h-67.413333a9.024 9.024 0 0 0-9.002667 8.981333v380.117333h-83.029333a8.96 8.96 0 0 0-7.082667 14.506667l125.866667 159.317333z"
  ],
  connection: [
    "M640 576c105.856 0 201.856 43.072 271.392 112.608l-90.464 90.464c-46.368-46.368-110.368-75.072-180.928-75.072s-134.592 28.704-180.928 75.072l-90.464-90.464c69.536-69.536 165.536-112.608 271.392-112.608zM187.456 507.456c120.864-120.864 281.6-187.456 452.544-187.456s331.68 66.56 452.544 187.456l-90.496 90.496c-96.704-96.704-225.28-149.952-362.048-149.952s-265.344 53.248-362.048 149.952l-90.496-90.496zM988.768 134.432c106.688 45.12 202.528 109.728 284.768 192l0 0-90.496 90.496c-145.056-145.056-337.92-224.928-543.072-224.928s-398.016 79.872-543.072 224.928l-90.496-90.496c82.272-82.272 178.08-146.848 284.768-192 110.496-46.752 227.84-70.432 348.768-70.432s238.272 23.712 348.768 70.432zM576 896c0-35.36 28.64-64 64-64s64 28.64 64 64c0 35.36-28.64 64-64 64s-64-28.64-64-64z"
  ],
  disconnection: [
    "M296.192 426.1376A440.8832 440.8832 0 0 0 199.68 499.2a40.96 40.96 0 1 0 51.2 63.744 43.5712 43.5712 0 0 0 6.144-6.144 359.0656 359.0656 0 0 1 95.8464-67.2256zM825.6 499.2A441.1904 441.1904 0 0 0 655.36 393.0112l155.1872 173.6192a39.3728 39.3728 0 0 0 15.36-9.8304 40.2944 40.2944 0 0 0 0.5632-57.0368zM164.1984 278.528a637.696 637.696 0 0 0-96.9728 78.592 40.96 40.96 0 0 0 28.7744 69.7856A39.1168 39.1168 0 0 0 128 414.72a554.3936 554.3936 0 0 1 91.2384-74.6496zM433.152 579.3792a205.1584 205.1584 0 0 0-65.792 43.3664 40.2432 40.2432 0 0 0-0.5632 56.9856l0.5632 0.6144a41.6768 41.6768 0 0 0 57.6 0A122.88 122.88 0 0 1 491.9296 645.12zM958.1056 359.0656A629.76 629.76 0 0 0 512 172.8c-17.2544 0-34.5088 0.7168-51.6608 2.1504l72.8064 81.4592a545.5872 545.5872 0 0 1 367.36 160.256 40.96 40.96 0 0 0 57.6-57.6zM512.6656 735.3856H512a54.4768 54.4768 0 1 0 56.9856 57.6 56.9344 56.9344 0 0 0-56.32-57.6z",
    "M195.047336 166.99934m25.07662-22.421086l0 0q25.07662-22.421086 47.497705 2.655534l534.181517 597.449514q22.421086 25.07662-2.655534 47.497706l0 0q-25.07662 22.421086-47.497705-2.655534l-534.181517-597.449514q-22.421086-25.07662 2.655534-47.497706Z"
  ],
  syncing: [
    "M148.935111 380.131556l2.56 2.616888 120.092445 168.163556a11.377778 11.377778 0 0 1-6.257778 17.578667L262.371556 568.888889H175.388444a341.447111 341.447111 0 0 0 558.478223 202.524444 28.444444 28.444444 0 0 1 36.977777 43.235556A398.336 398.336 0 0 1 117.76 568.888889H22.186667a11.377778 11.377778 0 0 1-10.922667-14.734222l1.649778-3.242667 120.092444-168.163556a11.377778 11.377778 0 0 1 15.928889-2.616888zM512 113.777778a398.336 398.336 0 0 1 394.24 341.333333h95.630222a11.377778 11.377778 0 0 1 10.922667 14.734222l-1.649778 3.242667-120.092444 168.163556a11.377778 11.377778 0 0 1-15.928889 2.616888l-2.616889-2.616888-120.092445-168.163556a11.377778 11.377778 0 0 1 6.257778-17.578667L761.628444 455.111111h86.926223A341.447111 341.447111 0 0 0 290.133333 252.586667a28.444444 28.444444 0 1 1-36.977777-43.235556A396.686222 396.686222 0 0 1 512 113.777778z"
  ],
  github: [
    "M512 42.666667A464.64 464.64 0 0 0 42.666667 502.186667 460.373333 460.373333 0 0 0 363.52 938.666667c23.466667 4.266667 32-9.813333 32-22.186667v-78.08c-130.56 27.733333-158.293333-61.44-158.293333-61.44a122.026667 122.026667 0 0 0-52.053334-67.413333c-42.666667-28.16 3.413333-27.733333 3.413334-27.733334a98.56 98.56 0 0 1 71.68 47.36 101.12 101.12 0 0 0 136.533333 37.973334 99.413333 99.413333 0 0 1 29.866667-61.44c-104.106667-11.52-213.333333-50.773333-213.333334-226.986667a177.066667 177.066667 0 0 1 47.36-124.16 161.28 161.28 0 0 1 4.693334-121.173333s39.68-12.373333 128 46.933333a455.68 455.68 0 0 1 234.666666 0c89.6-59.306667 128-46.933333 128-46.933333a161.28 161.28 0 0 1 4.693334 121.173333A177.066667 177.066667 0 0 1 810.666667 477.866667c0 176.64-110.08 215.466667-213.333334 226.986666a106.666667 106.666667 0 0 1 32 85.333334v125.866666c0 14.933333 8.533333 26.88 32 22.186667A460.8 460.8 0 0 0 981.333333 502.186667 464.64 464.64 0 0 0 512 42.666667"
  ]
}, Sg = "_baseIcon_27qkj_1", Rg = {
  baseIcon: Sg
}, Xi = ({
  className: e = "",
  paths: t,
  testId: n
}) => /* @__PURE__ */ I(
  "svg",
  {
    className: ae(Rg.baseIcon, e),
    viewBox: "0 0 1137 1024",
    "aria-hidden": !0,
    "data-testid": n,
    children: t.map((r, o) => /* @__PURE__ */ I("path", { d: r.d, fillOpacity: r["fill-opacity"] }, o, !1, {
      fileName: "/home/user/excel-collab/src/components/BaseIcon/BaseIcon.tsx",
      lineNumber: 29,
      columnNumber: 9
    }, void 0))
  },
  void 0,
  !1,
  {
    fileName: "/home/user/excel-collab/src/components/BaseIcon/BaseIcon.tsx",
    lineNumber: 22,
    columnNumber: 5
  },
  void 0
);
Xi.displayName = "BaseIcon";
const El = ue(() => Xi({
  paths: [
    {
      d: "M0 0h1024v1024H0z",
      "fill-opacity": ".01"
    },
    {
      d: "M496.512 32a128 128 0 0 1 127.84 121.6l0.16 6.4-0.032 113.504 264.256 264.256a32 32 0 0 1-8.16 51.2l-144.064 72.96-269.12 269.12a64 64 0 0 1-90.496 0l-294.144-294.176a64 64 0 0 1 0-90.496l286.016-286.08-0.192-2.048-0.064-2.08V160a128 128 0 0 1 128-128z m-6.464 197.568L128 591.616l294.144 294.144 276.32-276.32 113.792-57.632-187.776-187.776V416a32 32 0 0 1-28.256 31.776l-3.712 0.224a32 32 0 0 1-31.808-28.256L560.512 416l-0.032-115.968-70.432-70.464z m402.016 395.936l1.792 2.24 5.472 8.416c30.112 46.752 45.184 80.032 45.184 99.84a64 64 0 1 1-128 0c0-20.96 16.864-57.024 50.624-108.224a16 16 0 0 1 24.928-2.24zM496.512 96a64 64 0 0 0-63.84 59.2l-0.16 4.8-0.032 36.576 34.944-34.88a32 32 0 0 1 45.248 0l47.808 47.808V160a64 64 0 0 0-59.2-63.84L496.48 96z"
    }
  ]
}));
El.displayName = "FillColorIcon";
const Ae = ue(
  ({ name: e, className: t = "", testId: n }) => {
    const r = yg[e].map((o) => ({ d: o }));
    return Xi({ className: t, paths: r, testId: n });
  }
);
Ae.displayName = "Icon";
const Ag = "_selectList_nyspg_1", Tg = "_selectItem_nyspg_6", xg = "_disabled_nyspg_9", Dg = "_top_nyspg_32", Og = "_active_nyspg_36", ft = {
  selectList: Ag,
  selectItem: Tg,
  disabled: xg,
  "popup-container": "_popup-container_nyspg_18",
  top: Dg,
  active: Og,
  "popup-item": "_popup-item_nyspg_40",
  "popup-item-icon": "_popup-item-icon_nyspg_50",
  "popup-item-content": "_popup-item-content_nyspg_54",
  "select-list-container": "_select-list-container_nyspg_59",
  "select-list-trigger": "_select-list-trigger_nyspg_71"
};
function bn(e = !1, t) {
  const n = Xe(null);
  function r(o) {
    const s = o.target;
    n?.current && (n.current === s || s instanceof Node && n.current.contains(s)) || t();
  }
  return Re(() => {
    if (e)
      return document.addEventListener("pointerdown", r), window.addEventListener("blur", r), () => {
        document.removeEventListener("pointerdown", r), window.removeEventListener("blur", r);
      };
  }, [e]), n;
}
function bl(e) {
  const t = Xe(e), n = Xe((r) => {
    t.current?.(r);
  });
  return t.current = e, n.current;
}
const ln = ue((e) => {
  const {
    data: t,
    value: n,
    className: r,
    onChange: o,
    getItemStyle: s,
    title: l,
    defaultValue: f,
    testId: d
  } = e, c = z(
    (i) => {
      o(i.target.value);
    },
    []
  );
  return /* @__PURE__ */ I(
    "select",
    {
      onChange: c,
      value: n,
      defaultValue: f,
      name: "select",
      className: ae(ft.selectList, r),
      title: l,
      "data-testid": d,
      children: t.map((i) => {
        const a = typeof i == "object" ? i.value : i, u = typeof i == "object" ? i.label : i, h = typeof i == "object" ? i.disabled : !1;
        let g;
        return typeof s == "function" && (g = s(a)), /* @__PURE__ */ I(
          "option",
          {
            value: a,
            disabled: !!h,
            className: ae(ft.selectItem, {
              [ft.disabled]: h
            }),
            style: g,
            children: u
          },
          a,
          !1,
          {
            fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
            lineNumber: 63,
            columnNumber: 11
          },
          void 0
        );
      })
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
      lineNumber: 45,
      columnNumber: 5
    },
    void 0
  );
});
ln.displayName = "Select";
const Ki = ue(
  ({
    onChange: e,
    value: t,
    active: n,
    data: r,
    className: o,
    position: s = "bottom",
    testId: l
  }) => {
    const f = z(
      (d) => {
        const c = d.target.dataset?.value;
        c && e(c);
      },
      [e]
    );
    return /* @__PURE__ */ I(
      "div",
      {
        className: ae(
          ft["popup-container"],
          o,
          s === "top" ? ft.top : "",
          {
            [ft.active]: n
          }
        ),
        onClick: f,
        "data-testid": l,
        children: r.map((d) => /* @__PURE__ */ I("div", { className: ft["popup-item"], children: [
          /* @__PURE__ */ I("span", { className: ft["popup-item-content"], "data-value": d.value, children: d.label }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
            lineNumber: 127,
            columnNumber: 13
          }, void 0),
          d.value == t && /* @__PURE__ */ I("span", { className: ft["popup-item-icon"], children: /* @__PURE__ */ I(Ae, { name: "confirm" }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
            lineNumber: 132,
            columnNumber: 17
          }, void 0) }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
            lineNumber: 131,
            columnNumber: 15
          }, void 0)
        ] }, d.value, !0, {
          fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
          lineNumber: 126,
          columnNumber: 11
        }, void 0))
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
        lineNumber: 113,
        columnNumber: 7
      },
      void 0
    );
  }
);
Ki.displayName = "SelectPopup";
const Xr = ue(({ children: e, value: t, data: n, onChange: r, position: o, testId: s, className: l }) => {
  const [f, d] = he(!1), c = z(() => {
    d((u) => !u);
  }, []), i = z(
    (u) => {
      u && (r(u), d(!1));
    },
    [r]
  ), a = bn(!0, () => {
    d(!1);
  });
  return /* @__PURE__ */ I(
    "div",
    {
      ref: a,
      className: ae(ft["select-list-container"], l, {
        [ft.active]: f
      }),
      "data-testid": s,
      children: [
        e,
        /* @__PURE__ */ I(
          V,
          {
            className: ft["select-list-trigger"],
            onClick: c,
            testId: `${s}-trigger`,
            type: "plain",
            children: /* @__PURE__ */ I(Ae, { name: "down" }, void 0, !1, {
              fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
              lineNumber: 180,
              columnNumber: 9
            }, void 0)
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
            lineNumber: 174,
            columnNumber: 7
          },
          void 0
        ),
        f && n.length > 0 && /* @__PURE__ */ I(
          Ki,
          {
            active: !0,
            value: t,
            data: n,
            onChange: i,
            position: o,
            testId: `${s}-popup`
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
            lineNumber: 183,
            columnNumber: 9
          },
          void 0
        )
      ]
    },
    void 0,
    !0,
    {
      fileName: "/home/user/excel-collab/src/components/Select/index.tsx",
      lineNumber: 166,
      columnNumber: 5
    },
    void 0
  );
});
Xr.displayName = "SelectList";
const In = {
  "dialog-modal": "_dialog-modal_1hyk9_1",
  "dialog-container": "_dialog-container_1hyk9_13",
  "dialog-title": "_dialog-title_1hyk9_20",
  "dialog-content": "_dialog-content_1hyk9_25",
  "dialog-button": "_dialog-button_1hyk9_29",
  "dialog-cancel": "_dialog-cancel_1hyk9_36"
}, kg = {
  copy: "Copy",
  cut: "Cut",
  paste: "Paste",
  delete: "Delete",
  add: "Add",
  rows: "Rows",
  menu: "Menu",
  file: "File",
  insert: "Insert",
  rename: "Rename",
  reset: "Reset",
  hide: "Hide",
  unhide: "Unhide",
  cancel: "Cancel",
  "line-chart": "Line Chart",
  "bar-chart": "Bar Chart",
  "pie-chart": "Pie Chart",
  "scatter-chart": "Scatter Chart",
  "radar-chart": "Radar Chart",
  "polar-area-chart": "PolarArea Chart",
  confirm: "Confirm",
  "tab-color": "Tab Color",
  "unhide-sheet": "Unhide sheet:",
  "wrap-text": "Wrap Text",
  import: "Import {{format}}",
  export: "Export {{format}}",
  "floating-picture": "Floating Picture",
  chart: "Chart",
  "chart-title": "Chart Title",
  duplicate: "Duplicate",
  "select-data": "Select Data",
  "change-chart-title": "Change Chart Title",
  "change-chart-type": "Change Chart Type",
  "save-as-picture": "Save as Picture",
  "reset-size": "Reset Size",
  "add-at-the-bottom": "Add at the bottom",
  "row-height": "Row Height",
  "column-width": "Column Width",
  "greater-than-zero": "The value must be greater than or equal to 0",
  "insert-columns": "Insert Columns",
  "delete-columns": "Delete Columns",
  "hide-columns": "Hide Columns",
  "unhide-columns": "Unhide Columns",
  "insert-rows": "Insert Rows",
  "delete-rows": "Delete Rows",
  "hide-rows": "Hide rows",
  "unhide-rows": "Unhide rows",
  "reference-is-empty": "The reference is empty",
  "reference-is-not-valid": "The reference is not valid",
  "the-value-cannot-be-empty": "The value cannot be empty",
  "edit-data-source": "Edit Data Source",
  "sheet-is-not-exist": "The sheet is not exist",
  "sheet-id-is-duplicate": "The sheet id is duplicate",
  "a-workbook-must-contains-at-least-one-visible-worksheet": "A workbook must contains at least one visible worksheet",
  "sheet-name-is-duplicate": "The sheet name is duplicate",
  "uuid-is-duplicate": "The uuid is duplicate",
  "cells-must-contain-data": "The selected cells must contain the data",
  "image-source-is-empty": "The image source is empty",
  "merging-cell-is-duplicate": "The merging cell is duplicate",
  "insert-column-left": "Insert Column Left",
  "insert-column-right": "Insert Column Right",
  "insert-row-above": "Insert Row Above",
  "insert-row-below": "Insert Row Below",
  "get-all-installed-fonts": "Get all the fonts installed locally",
  "unsupported-chart-types": "Unsupported chart types",
  "merge-cell": "Merge",
  "merge-and-center": "Merge and Center",
  "merge-cells": "Merge Cells",
  "merge-content": "Merge Content",
  none: "none",
  "single-underline": "single underline",
  "double-underline": "double underline",
  general: "General",
  number: "Number",
  currency: "Currency",
  accounting: "Accounting",
  "short-date": "Short Date",
  "long-date": "Long Date",
  time: "Time",
  percentage: "Percentage",
  fraction: "Fraction",
  scientific: "Scientific",
  text: "Text",
  "more-number-formats": "More Number Formats",
  "short-date-format": "m/d/yyyy;@",
  "long-date-format": "[$-409]dddd,mmmm d,yyyy;@",
  "time-format": "h:mm:ss AM/PM",
  "sheet-id-can-not-be-empty": "The sheetId can not be empty",
  "no-border": "No Border",
  "all-borders": "All Borders",
  "outside-borders": "OutSide Borders",
  "thick-box-border": "Thick Box Border",
  "bottom-border": "Bottom Border",
  "top-border": "Top Border",
  "left-border": "Left Border",
  "right-border": "Right Border",
  "line-color": "Line Color",
  "line-style": "Line Style",
  filter: "Filter",
  "filter-empty": "(empty)",
  "filter-all": "All",
  "new-file": "New File",
  "rename-file": "Rename",
  "default-name": "default name",
  "change-file-name": "Change File Name",
  "user-name": "User",
  "no-login-editing": "You can't edit a document until you're logged in",
  login: "Login",
  "log-out": "Log Out"
}, Lg = {
  "en-US": kg,
  "zh-CN": {
    copy: "复制",
    cut: "剪切",
    paste: "粘贴",
    delete: "删除",
    add: "添加",
    rows: "行",
    menu: "菜单",
    file: "文件",
    insert: "插入",
    rename: "重命名",
    reset: "重置",
    hide: "隐藏",
    unhide: "显示",
    cancel: "取消",
    "line-chart": "折线图",
    "bar-chart": "柱状图",
    "pie-chart": "饼图",
    "scatter-chart": "散点图",
    "radar-chart": "雷达图",
    "polar-area-chart": "极区图",
    confirm: "确认",
    "tab-color": "选项卡颜色",
    "unhide-sheet": "显示工作表:",
    "wrap-text": "自动换行",
    import: "导入 {{format}}",
    export: "导出 {{format}}",
    "floating-picture": "浮动图片",
    chart: "图表",
    "chart-title": "图表标题",
    duplicate: "复制",
    "select-data": "选择数据",
    "change-chart-title": "更改图表标题",
    "change-chart-type": "更改图表类型",
    "save-as-picture": "保存为图片",
    "reset-size": "重置大小",
    "add-at-the-bottom": "在底部添加",
    "row-height": "行高",
    "column-width": "列宽",
    "greater-than-zero": "值必须大于或等于 0",
    "insert-columns": "插入列",
    "delete-columns": "删除列",
    "hide-columns": "隐藏列",
    "unhide-columns": "取消隐藏列",
    "insert-rows": "插入行",
    "delete-rows": "删除行",
    "hide-rows": "隐藏行",
    "unhide-rows": "取消隐藏行",
    "reference-is-empty": "引用为空",
    "reference-is-not-valid": "引用无效",
    "the-value-cannot-be-empty": "值不能为空",
    "edit-data-source": "编辑数据源",
    "sheet-id-is-duplicate": "工作表 ID 重复",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "工作簿必须至少包含一个可见工作表",
    "sheet-name-is-duplicate": "工作表名称重复",
    "uuid-is-duplicate": "UUID 重复",
    "cells-must-contain-data": "所选单元格必须包含数据",
    "image-source-is-empty": "图像源为空",
    "merging-cell-is-duplicate": "合并单元格重复",
    "insert-column-left": "向左插入列",
    "insert-column-right": "向右插入列",
    "insert-row-above": "在上面插入行",
    "insert-row-below": "在下面插入行",
    "get-all-installed-fonts": "获取本地安装的所有字体",
    "unsupported-chart-types": "不支持的图表类型",
    "merge-cell": "合并单元格",
    "merge-and-center": "合并并且居中",
    "merge-cells": "合并单元格",
    "merge-content": "合并内容",
    none: "无",
    "single-underline": "单层下划线",
    "double-underline": "双层下划线",
    general: "常规",
    number: "数值",
    currency: "货币",
    accounting: "会计专用",
    "short-date": "短日期",
    "long-date": "长日期",
    time: "时间",
    percentage: "百分比",
    fraction: "分数",
    scientific: "科学计数",
    text: "文本",
    "more-number-formats": "其他数字格式",
    "short-date-format": "yyyy/m/d;@",
    "long-date-format": 'yyyy"年"m"月"d"日"',
    "time-format": "h:mm:ss",
    "sheet-id-can-not-be-empty": "SheetId 不能为空",
    "no-border": "无边框",
    "all-borders": "全边框",
    "outside-borders": "外边框",
    "thick-box-border": "厚框边框",
    "bottom-border": "底边框",
    "top-border": "顶边框",
    "left-border": "左边框",
    "right-border": "右边框",
    "line-color": "线条颜色",
    "line-style": "线条样式",
    filter: "筛选",
    "filter-empty": "(空)",
    "filter-all": "全部",
    "sheet-is-not-exist": "工作表不存在",
    "new-file": "新建文件",
    "rename-file": "重命名",
    "default-name": "默认名称",
    "change-file-name": "修改文件名称",
    "user-name": "用户",
    "no-login-editing": "你只有登录后，才可以编辑文档",
    login: "登录",
    "log-out": "退出"
  },
  "fr-FR": {
    copy: "Copier",
    cut: "Couper",
    paste: "Coller",
    delete: "Supprimer",
    add: "Ajouter",
    rows: "Lignes",
    menu: "Menu",
    file: "Fichier",
    insert: "Insérer",
    rename: "Renommer",
    reset: "Réinitialiser",
    hide: "Masquer",
    unhide: "Afficher",
    cancel: "Annuler",
    "line-chart": "Graphique linéaire",
    "bar-chart": "Graphique à barres",
    "pie-chart": "Graphique circulaire",
    "scatter-chart": "Graphique de dispersion",
    "radar-chart": "Graphique radar",
    "polar-area-chart": "Graphique polaire",
    confirm: "Confirmer",
    "tab-color": "Couleur de l'onglet",
    "unhide-sheet": "Afficher la feuille:",
    "wrap-text": "Ajustement automatique",
    import: "Importer {{format}}",
    export: "Exporter {{format}}",
    "floating-picture": "Image flottante",
    chart: "Graphique",
    "chart-title": "Titre du graphique",
    duplicate: "Dupliquer",
    "select-data": "Sélectionner les données",
    "change-chart-title": "Modifier le titre du graphique",
    "change-chart-type": "Modifier le type de graphique",
    "save-as-picture": "Enregistrer en tant qu'image",
    "reset-size": "Réinitialiser la taille",
    "add-at-the-bottom": "Ajouter en bas",
    "row-height": "Hauteur de ligne",
    "column-width": "Largeur de colonne",
    "greater-than-zero": "La valeur doit être supérieure ou égale à 0",
    "insert-columns": "Insérer des colonnes",
    "delete-columns": "Supprimer des colonnes",
    "hide-columns": "Masquer les colonnes",
    "unhide-columns": "Afficher les colonnes",
    "insert-rows": "Insérer des lignes",
    "delete-rows": "Supprimer des lignes",
    "hide-rows": "Masquer les lignes",
    "unhide-rows": "Afficher les lignes",
    "reference-is-empty": "La référence est vide",
    "reference-is-not-valid": "La référence n'est pas valide",
    "the-value-cannot-be-empty": "La valeur ne peut pas être vide",
    "edit-data-source": "Modifier la source de données",
    "sheet-is-not-exist": "La feuille n'existe pas",
    "sheet-id-is-duplicate": "L'ID de la feuille est dupliqué",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "Un classeur doit contenir au moins une feuille de calcul visible",
    "sheet-name-is-duplicate": "Le nom de la feuille est dupliqué",
    "uuid-is-duplicate": "L'UUID est dupliqué",
    "cells-must-contain-data": "Les cellules sélectionnées doivent contenir des données",
    "image-source-is-empty": "La source de l'image est vide",
    "merging-cell-is-duplicate": "La cellule de fusion est dupliquée",
    "insert-column-left": "Insérer une colonne à gauche",
    "insert-column-right": "Insérer une colonne à droite",
    "insert-row-above": "Insérer une ligne au-dessus",
    "insert-row-below": "Insérer une ligne en dessous",
    "get-all-installed-fonts": "Obtenir toutes les polices installées localement",
    "unsupported-chart-types": "Types de graphiques non pris en charge",
    "merge-cell": "Fusionner les cellules",
    "merge-and-center": "Fusionner et centrer",
    "merge-cells": "Fusionner les cellules",
    "merge-content": "Fusionner le contenu",
    none: "aucun",
    "single-underline": "soulignement simple",
    "double-underline": "double soulignement",
    general: "Général",
    number: "Nombre",
    currency: "Devise",
    accounting: "Comptabilité",
    "short-date": "Date courte",
    "long-date": "Date longue",
    time: "Heure",
    percentage: "Pourcentage",
    fraction: "Fraction",
    scientific: "Scientifique",
    text: "Texte",
    "more-number-formats": "Plus de formats de nombre",
    "short-date-format": "dd/mm/yyyy;@",
    "long-date-format": "dddd d mmmm yyyy;@",
    "time-format": "hh:mm:ss",
    "sheet-id-can-not-be-empty": "L'ID de la feuille ne peut pas être vide",
    "no-border": "Aucune bordure",
    "all-borders": "Toutes les bordures",
    "outside-borders": "Bordures extérieures",
    "thick-box-border": "Bordure de boîte épaisse",
    "bottom-border": "Bordure inférieure",
    "top-border": "Bordure supérieure",
    "left-border": "Bordure gauche",
    "right-border": "Bordure droite",
    "line-color": "Couleur de ligne",
    "line-style": "Style de ligne",
    filter: "Filtre",
    "filter-empty": "(vide)",
    "filter-all": "Tout",
    "new-file": "Nouveau fichier",
    "rename-file": "Renommer",
    "default-name": "nom par défaut",
    "change-file-name": "Modifier le nom du fichier",
    "user-name": "Utilisateur",
    "no-login-editing": "Vous ne pouvez pas modifier un document tant que vous n'êtes pas connecté",
    login: "Se connecter",
    "log-out": "Se déconnecter"
  },
  "zh-HK": {
    copy: "複製",
    cut: "剪下",
    paste: "貼上",
    delete: "刪除",
    add: "新增",
    rows: "行",
    menu: "選單",
    file: "檔案",
    insert: "插入",
    rename: "重新命名",
    reset: "重設",
    hide: "隱藏",
    unhide: "顯示",
    cancel: "取消",
    "line-chart": "折線圖",
    "bar-chart": "長條圖",
    "pie-chart": "圓餅圖",
    "scatter-chart": "散佈圖",
    "radar-chart": "雷達圖",
    "polar-area-chart": "極區圖",
    confirm: "確認",
    "tab-color": "頁籤顏色",
    "unhide-sheet": "顯示工作表:",
    "wrap-text": "自動換行",
    import: "匯入 {{format}}",
    export: "匯出 {{format}}",
    "floating-picture": "浮動圖片",
    chart: "圖表",
    "chart-title": "圖表標題",
    duplicate: "複製",
    "select-data": "選擇資料",
    "change-chart-title": "變更圖表標題",
    "change-chart-type": "變更圖表類型",
    "save-as-picture": "另存為圖片",
    "reset-size": "重設大小",
    "add-at-the-bottom": "在底部新增",
    "row-height": "行高",
    "column-width": "欄寬",
    "greater-than-zero": "值必須大於或等於 0",
    "insert-columns": "插入欄",
    "delete-columns": "刪除欄",
    "hide-columns": "隱藏欄",
    "unhide-columns": "取消隱藏欄",
    "insert-rows": "插入行",
    "delete-rows": "刪除行",
    "hide-rows": "隱藏行",
    "unhide-rows": "取消隱藏行",
    "reference-is-empty": "參考為空",
    "reference-is-not-valid": "參考無效",
    "the-value-cannot-be-empty": "值不能為空",
    "edit-data-source": "編輯資料來源",
    "sheet-is-not-exist": "工作表不存在",
    "sheet-id-is-duplicate": "工作表 ID 重複",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "活頁簿必須至少包含一個可見的工作表",
    "sheet-name-is-duplicate": "工作表名稱重複",
    "uuid-is-duplicate": "UUID 重複",
    "cells-must-contain-data": "選擇的儲存格必須包含資料",
    "image-source-is-empty": "圖片來源為空",
    "merging-cell-is-duplicate": "合併儲存格重複",
    "insert-column-left": "向左插入欄",
    "insert-column-right": "向右插入欄",
    "insert-row-above": "在上面插入行",
    "insert-row-below": "在下面插入行",
    "get-all-installed-fonts": "取得本機安裝的所有字型",
    "unsupported-chart-types": "不支援的圖表類型",
    "merge-cell": "合併儲存格",
    "merge-and-center": "合併並置中",
    "merge-cells": "合併儲存格",
    "merge-content": "合併內容",
    none: "無",
    "single-underline": "單底線",
    "double-underline": "雙底線",
    general: "一般",
    number: "數字",
    currency: "貨幣",
    accounting: "會計",
    "short-date": "短日期",
    "long-date": "長日期",
    time: "時間",
    percentage: "百分比",
    fraction: "分數",
    scientific: "科學記號",
    text: "文字",
    "more-number-formats": "其他數字格式",
    "short-date-format": "yyyy/m/d;@",
    "long-date-format": 'yyyy"年"m"月"d"日"',
    "time-format": "h:mm:ss",
    "sheet-id-can-not-be-empty": "SheetId 不能為空",
    "no-border": "無框線",
    "all-borders": "所有框線",
    "outside-borders": "外框線",
    "thick-box-border": "粗外框線",
    "bottom-border": "下框線",
    "top-border": "上框線",
    "left-border": "左框線",
    "right-border": "右框線",
    "line-color": "線條顏色",
    "line-style": "線條樣式",
    filter: "篩選",
    "filter-empty": "(空白)",
    "filter-all": "全部",
    "new-file": "新增檔案",
    "rename-file": "重新命名",
    "default-name": "預設名稱",
    "change-file-name": "修改檔案名稱",
    "user-name": "使用者",
    "no-login-editing": "您必須登入後才能編輯文件",
    login: "登入",
    "log-out": "登出"
  },
  "ja-JP": {
    copy: "コピー",
    cut: "切り取り",
    paste: "貼り付け",
    delete: "削除",
    add: "追加",
    rows: "行",
    menu: "メニュー",
    file: "ファイル",
    insert: "挿入",
    rename: "名前の変更",
    reset: "リセット",
    hide: "非表示",
    unhide: "表示",
    cancel: "キャンセル",
    "line-chart": "折れ線グラフ",
    "bar-chart": "棒グラフ",
    "pie-chart": "円グラフ",
    "scatter-chart": "散布図",
    "radar-chart": "レーダーチャート",
    "polar-area-chart": "極域グラフ",
    confirm: "確認",
    "tab-color": "タブの色",
    "unhide-sheet": "シートの表示:",
    "wrap-text": "文字列の折り返し",
    import: "{{format}}をインポート",
    export: "{{format}}をエクスポート",
    "floating-picture": "浮動画像",
    chart: "グラフ",
    "chart-title": "グラフタイトル",
    duplicate: "複製",
    "select-data": "データの選択",
    "change-chart-title": "グラフタイトルの変更",
    "change-chart-type": "グラフの種類の変更",
    "save-as-picture": "画像として保存",
    "reset-size": "サイズのリセット",
    "add-at-the-bottom": "下部に追加",
    "row-height": "行の高さ",
    "column-width": "列の幅",
    "greater-than-zero": "値は0以上である必要があります",
    "insert-columns": "列の挿入",
    "delete-columns": "列の削除",
    "hide-columns": "列の非表示",
    "unhide-columns": "列の表示",
    "insert-rows": "行の挿入",
    "delete-rows": "行の削除",
    "hide-rows": "行の非表示",
    "unhide-rows": "行の表示",
    "reference-is-empty": "参照が空です",
    "reference-is-not-valid": "参照が無効です",
    "the-value-cannot-be-empty": "値を空にできません",
    "edit-data-source": "データソースの編集",
    "sheet-is-not-exist": "シートが存在しません",
    "sheet-id-is-duplicate": "シートIDが重複しています",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "ワークブックには少なくとも1つの表示可能なワークシートが必要です",
    "sheet-name-is-duplicate": "シート名が重複しています",
    "uuid-is-duplicate": "UUIDが重複しています",
    "cells-must-contain-data": "選択したセルにはデータが含まれている必要があります",
    "image-source-is-empty": "画像ソースが空です",
    "merging-cell-is-duplicate": "結合セルが重複しています",
    "insert-column-left": "左に列を挿入",
    "insert-column-right": "右に列を挿入",
    "insert-row-above": "上に行を挿入",
    "insert-row-below": "下に行を挿入",
    "get-all-installed-fonts": "ローカルにインストールされたすべてのフォントを取得",
    "unsupported-chart-types": "サポートされていないグラフの種類",
    "merge-cell": "セルの結合",
    "merge-and-center": "結合して中央揃え",
    "merge-cells": "セルの結合",
    "merge-content": "コンテンツの結合",
    none: "なし",
    "single-underline": "下線",
    "double-underline": "二重下線",
    general: "標準",
    number: "数値",
    currency: "通貨",
    accounting: "会計",
    "short-date": "短い日付",
    "long-date": "長い日付",
    time: "時刻",
    percentage: "パーセンテージ",
    fraction: "分数",
    scientific: "指数",
    text: "テキスト",
    "more-number-formats": "その他の数値形式",
    "short-date-format": "yyyy/m/d;@",
    "long-date-format": 'yyyy"年"m"月"d"日"',
    "time-format": "h:mm:ss",
    "sheet-id-can-not-be-empty": "SheetIDを空にできません",
    "no-border": "枠線なし",
    "all-borders": "すべての枠線",
    "outside-borders": "外枠",
    "thick-box-border": "太い外枠",
    "bottom-border": "下枠線",
    "top-border": "上枠線",
    "left-border": "左枠線",
    "right-border": "右枠線",
    "line-color": "線の色",
    "line-style": "線のスタイル",
    filter: "フィルター",
    "filter-empty": "(空白)",
    "filter-all": "すべて",
    "new-file": "新しいファイル",
    "rename-file": "名前の変更",
    "default-name": "デフォルト名",
    "change-file-name": "ファイル名の変更",
    "user-name": "ユーザー",
    "no-login-editing": "ログインするまでドキュメントを編集できません",
    login: "ログイン",
    "log-out": "ログアウト"
  },
  "ko-KR": {
    copy: "복사",
    cut: "잘라내기",
    paste: "붙여넣기",
    delete: "삭제",
    add: "추가",
    rows: "행",
    menu: "메뉴",
    file: "파일",
    insert: "삽입",
    rename: "이름 바꾸기",
    reset: "재설정",
    hide: "숨기기",
    unhide: "숨기기 취소",
    cancel: "취소",
    "line-chart": "꺾은선형 차트",
    "bar-chart": "막대형 차트",
    "pie-chart": "원형 차트",
    "scatter-chart": "분산형 차트",
    "radar-chart": "방사형 차트",
    "polar-area-chart": "극지역 차트",
    confirm: "확인",
    "tab-color": "탭 색상",
    "unhide-sheet": "시트 숨기기 취소:",
    "wrap-text": "텍스트 줄 바꿈",
    import: "{{format}} 가져오기",
    export: "{{format}} 내보내기",
    "floating-picture": "부동 그림",
    chart: "차트",
    "chart-title": "차트 제목",
    duplicate: "복제",
    "select-data": "데이터 선택",
    "change-chart-title": "차트 제목 변경",
    "change-chart-type": "차트 종류 변경",
    "save-as-picture": "그림으로 저장",
    "reset-size": "크기 재설정",
    "add-at-the-bottom": "맨 아래에 추가",
    "row-height": "행 높이",
    "column-width": "열 너비",
    "greater-than-zero": "값은 0보다 크거나 같아야 합니다",
    "insert-columns": "열 삽입",
    "delete-columns": "열 삭제",
    "hide-columns": "열 숨기기",
    "unhide-columns": "열 숨기기 취소",
    "insert-rows": "행 삽입",
    "delete-rows": "행 삭제",
    "hide-rows": "행 숨기기",
    "unhide-rows": "행 숨기기 취소",
    "reference-is-empty": "참조가 비어 있습니다",
    "reference-is-not-valid": "참조가 유효하지 않습니다",
    "the-value-cannot-be-empty": "값은 비워둘 수 없습니다",
    "edit-data-source": "데이터 소스 편집",
    "sheet-is-not-exist": "시트가 존재하지 않습니다",
    "sheet-id-is-duplicate": "시트 ID가 중복됩니다",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "통합 문서에는 최소한 하나의 표시되는 워크시트가 있어야 합니다",
    "sheet-name-is-duplicate": "시트 이름이 중복됩니다",
    "uuid-is-duplicate": "UUID가 중복됩니다",
    "cells-must-contain-data": "선택한 셀에는 데이터가 포함되어야 합니다",
    "image-source-is-empty": "이미지 소스가 비어 있습니다",
    "merging-cell-is-duplicate": "병합 셀이 중복됩니다",
    "insert-column-left": "왼쪽에 열 삽입",
    "insert-column-right": "오른쪽에 열 삽입",
    "insert-row-above": "위에 행 삽입",
    "insert-row-below": "아래에 행 삽입",
    "get-all-installed-fonts": "로컬에 설치된 모든 글꼴 가져오기",
    "unsupported-chart-types": "지원되지 않는 차트 유형",
    "merge-cell": "셀 병합",
    "merge-and-center": "병합하고 가운데 맞춤",
    "merge-cells": "셀 병합",
    "merge-content": "내용 병합",
    none: "없음",
    "single-underline": "밑줄",
    "double-underline": "이중 밑줄",
    general: "일반",
    number: "숫자",
    currency: "통화",
    accounting: "회계",
    "short-date": "간단한 날짜",
    "long-date": "긴 날짜",
    time: "시간",
    percentage: "백분율",
    fraction: "분수",
    scientific: "지수",
    text: "텍스트",
    "more-number-formats": "기타 숫자 형식",
    "short-date-format": "yyyy/m/d;@",
    "long-date-format": 'yyyy"년" m"월" d"일"',
    "time-format": "h:mm:ss",
    "sheet-id-can-not-be-empty": "SheetId는 비워둘 수 없습니다",
    "no-border": "테두리 없음",
    "all-borders": "모든 테두리",
    "outside-borders": "바깥쪽 테두리",
    "thick-box-border": "굵은 상자 테두리",
    "bottom-border": "아래쪽 테두리",
    "top-border": "위쪽 테두리",
    "left-border": "왼쪽 테두리",
    "right-border": "오른쪽 테두리",
    "line-color": "선 색상",
    "line-style": "선 스타일",
    filter: "필터",
    "filter-empty": "(빈 값)",
    "filter-all": "모두",
    "new-file": "새 파일",
    "rename-file": "이름 바꾸기",
    "default-name": "기본 이름",
    "change-file-name": "파일 이름 변경",
    "user-name": "사용자",
    "no-login-editing": "로그인해야 문서를 편집할 수 있습니다",
    login: "로그인",
    "log-out": "로그아웃"
  },
  "es-ES": {
    copy: "Copiar",
    cut: "Cortar",
    paste: "Pegar",
    delete: "Eliminar",
    add: "Agregar",
    rows: "Filas",
    menu: "Menú",
    file: "Archivo",
    insert: "Insertar",
    rename: "Renombrar",
    reset: "Restablecer",
    hide: "Ocultar",
    unhide: "Mostrar",
    cancel: "Cancelar",
    "line-chart": "Gráfico de líneas",
    "bar-chart": "Gráfico de barras",
    "pie-chart": "Gráfico circular",
    "scatter-chart": "Gráfico de dispersión",
    "radar-chart": "Gráfico de radar",
    "polar-area-chart": "Gráfico polar",
    confirm: "Confirmar",
    "tab-color": "Color de pestaña",
    "unhide-sheet": "Mostrar hoja:",
    "wrap-text": "Ajustar texto",
    import: "Importar {{format}}",
    export: "Exportar {{format}}",
    "floating-picture": "Imagen flotante",
    chart: "Gráfico",
    "chart-title": "Título del gráfico",
    duplicate: "Duplicar",
    "select-data": "Seleccionar datos",
    "change-chart-title": "Cambiar título del gráfico",
    "change-chart-type": "Cambiar tipo de gráfico",
    "save-as-picture": "Guardar como imagen",
    "reset-size": "Restablecer tamaño",
    "add-at-the-bottom": "Agregar al final",
    "row-height": "Alto de fila",
    "column-width": "Ancho de columna",
    "greater-than-zero": "El valor debe ser mayor o igual a 0",
    "insert-columns": "Insertar columnas",
    "delete-columns": "Eliminar columnas",
    "hide-columns": "Ocultar columnas",
    "unhide-columns": "Mostrar columnas",
    "insert-rows": "Insertar filas",
    "delete-rows": "Eliminar filas",
    "hide-rows": "Ocultar filas",
    "unhide-rows": "Mostrar filas",
    "reference-is-empty": "La referencia está vacía",
    "reference-is-not-valid": "La referencia no es válida",
    "the-value-cannot-be-empty": "El valor no puede estar vacío",
    "edit-data-source": "Editar fuente de datos",
    "sheet-is-not-exist": "La hoja no existe",
    "sheet-id-is-duplicate": "El ID de la hoja está duplicado",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "Un libro debe contener al menos una hoja de cálculo visible",
    "sheet-name-is-duplicate": "El nombre de la hoja está duplicado",
    "uuid-is-duplicate": "El UUID está duplicado",
    "cells-must-contain-data": "Las celdas seleccionadas deben contener datos",
    "image-source-is-empty": "La fuente de la imagen está vacía",
    "merging-cell-is-duplicate": "La celda combinada está duplicada",
    "insert-column-left": "Insertar columna a la izquierda",
    "insert-column-right": "Insertar columna a la derecha",
    "insert-row-above": "Insertar fila arriba",
    "insert-row-below": "Insertar fila abajo",
    "get-all-installed-fonts": "Obtener todas las fuentes instaladas localmente",
    "unsupported-chart-types": "Tipos de gráficos no compatibles",
    "merge-cell": "Combinar celdas",
    "merge-and-center": "Combinar y centrar",
    "merge-cells": "Combinar celdas",
    "merge-content": "Combinar contenido",
    none: "ninguno",
    "single-underline": "subrayado simple",
    "double-underline": "subrayado doble",
    general: "General",
    number: "Número",
    currency: "Moneda",
    accounting: "Contabilidad",
    "short-date": "Fecha corta",
    "long-date": "Fecha larga",
    time: "Hora",
    percentage: "Porcentaje",
    fraction: "Fracción",
    scientific: "Científico",
    text: "Texto",
    "more-number-formats": "Más formatos de número",
    "short-date-format": "dd/mm/yyyy;@",
    "long-date-format": "dddd, d de mmmm de yyyy;@",
    "time-format": "h:mm:ss",
    "sheet-id-can-not-be-empty": "El SheetId no puede estar vacío",
    "no-border": "Sin borde",
    "all-borders": "Todos los bordes",
    "outside-borders": "Bordes exteriores",
    "thick-box-border": "Borde de caja grueso",
    "bottom-border": "Borde inferior",
    "top-border": "Borde superior",
    "left-border": "Borde izquierdo",
    "right-border": "Borde derecho",
    "line-color": "Color de línea",
    "line-style": "Estilo de línea",
    filter: "Filtro",
    "filter-empty": "(vacío)",
    "filter-all": "Todo",
    "new-file": "Nuevo archivo",
    "rename-file": "Renombrar",
    "default-name": "nombre predeterminado",
    "change-file-name": "Cambiar nombre del archivo",
    "user-name": "Usuario",
    "no-login-editing": "No puedes editar un documento hasta que inicies sesión",
    login: "Iniciar sesión",
    "log-out": "Cerrar sesión"
  },
  "de-DE": {
    copy: "Kopieren",
    cut: "Ausschneiden",
    paste: "Einfügen",
    delete: "Löschen",
    add: "Hinzufügen",
    rows: "Zeilen",
    menu: "Menü",
    file: "Datei",
    insert: "Einfügen",
    rename: "Umbenennen",
    reset: "Zurücksetzen",
    hide: "Ausblenden",
    unhide: "Einblenden",
    cancel: "Abbrechen",
    "line-chart": "Liniendiagramm",
    "bar-chart": "Balkendiagramm",
    "pie-chart": "Kreisdiagramm",
    "scatter-chart": "Punktdiagramm",
    "radar-chart": "Netzdiagramm",
    "polar-area-chart": "Polar-Flächendiagramm",
    confirm: "Bestätigen",
    "tab-color": "Tab-Farbe",
    "unhide-sheet": "Blatt einblenden:",
    "wrap-text": "Text umbrechen",
    import: "{{format}} importieren",
    export: "{{format}} exportieren",
    "floating-picture": "Schwebendes Bild",
    chart: "Diagramm",
    "chart-title": "Diagrammtitel",
    duplicate: "Duplizieren",
    "select-data": "Daten auswählen",
    "change-chart-title": "Diagrammtitel ändern",
    "change-chart-type": "Diagrammtyp ändern",
    "save-as-picture": "Als Bild speichern",
    "reset-size": "Größe zurücksetzen",
    "add-at-the-bottom": "Am Ende hinzufügen",
    "row-height": "Zeilenhöhe",
    "column-width": "Spaltenbreite",
    "greater-than-zero": "Der Wert muss größer oder gleich 0 sein",
    "insert-columns": "Spalten einfügen",
    "delete-columns": "Spalten löschen",
    "hide-columns": "Spalten ausblenden",
    "unhide-columns": "Spalten einblenden",
    "insert-rows": "Zeilen einfügen",
    "delete-rows": "Zeilen löschen",
    "hide-rows": "Zeilen ausblenden",
    "unhide-rows": "Zeilen einblenden",
    "reference-is-empty": "Die Referenz ist leer",
    "reference-is-not-valid": "Die Referenz ist ungültig",
    "the-value-cannot-be-empty": "Der Wert darf nicht leer sein",
    "edit-data-source": "Datenquelle bearbeiten",
    "sheet-is-not-exist": "Das Blatt existiert nicht",
    "sheet-id-is-duplicate": "Die Blatt-ID ist doppelt vorhanden",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "Eine Arbeitsmappe muss mindestens ein sichtbares Arbeitsblatt enthalten",
    "sheet-name-is-duplicate": "Der Blattname ist doppelt vorhanden",
    "uuid-is-duplicate": "Die UUID ist doppelt vorhanden",
    "cells-must-contain-data": "Die ausgewählten Zellen müssen Daten enthalten",
    "image-source-is-empty": "Die Bildquelle ist leer",
    "merging-cell-is-duplicate": "Die zusammengeführte Zelle ist doppelt vorhanden",
    "insert-column-left": "Spalte links einfügen",
    "insert-column-right": "Spalte rechts einfügen",
    "insert-row-above": "Zeile oberhalb einfügen",
    "insert-row-below": "Zeile unterhalb einfügen",
    "get-all-installed-fonts": "Alle lokal installierten Schriftarten abrufen",
    "unsupported-chart-types": "Nicht unterstützte Diagrammtypen",
    "merge-cell": "Zellen verbinden",
    "merge-and-center": "Verbinden und zentrieren",
    "merge-cells": "Zellen verbinden",
    "merge-content": "Inhalt verbinden",
    none: "keine",
    "single-underline": "einfache Unterstreichung",
    "double-underline": "doppelte Unterstreichung",
    general: "Standard",
    number: "Zahl",
    currency: "Währung",
    accounting: "Buchhaltung",
    "short-date": "Kurzes Datum",
    "long-date": "Langes Datum",
    time: "Zeit",
    percentage: "Prozent",
    fraction: "Bruch",
    scientific: "Wissenschaftlich",
    text: "Text",
    "more-number-formats": "Weitere Zahlenformate",
    "short-date-format": "dd.mm.yyyy;@",
    "long-date-format": "dddd, d. mmmm yyyy;@",
    "time-format": "hh:mm:ss",
    "sheet-id-can-not-be-empty": "Die SheetId darf nicht leer sein",
    "no-border": "Kein Rahmen",
    "all-borders": "Alle Rahmen",
    "outside-borders": "Außenrahmen",
    "thick-box-border": "Dicker Kastenrahmen",
    "bottom-border": "Unterer Rahmen",
    "top-border": "Oberer Rahmen",
    "left-border": "Linker Rahmen",
    "right-border": "Rechter Rahmen",
    "line-color": "Linienfarbe",
    "line-style": "Linienstil",
    filter: "Filter",
    "filter-empty": "(leer)",
    "filter-all": "Alle",
    "new-file": "Neue Datei",
    "rename-file": "Umbenennen",
    "default-name": "Standardname",
    "change-file-name": "Dateiname ändern",
    "user-name": "Benutzer",
    "no-login-editing": "Sie können ein Dokument erst bearbeiten, wenn Sie angemeldet sind",
    login: "Anmelden",
    "log-out": "Abmelden"
  },
  "it-IT": {
    copy: "Copia",
    cut: "Taglia",
    paste: "Incolla",
    delete: "Elimina",
    add: "Aggiungi",
    rows: "Righe",
    menu: "Menu",
    file: "File",
    insert: "Inserisci",
    rename: "Rinomina",
    reset: "Reimposta",
    hide: "Nascondi",
    unhide: "Mostra",
    cancel: "Annulla",
    "line-chart": "Grafico a linee",
    "bar-chart": "Grafico a barre",
    "pie-chart": "Grafico a torta",
    "scatter-chart": "Grafico a dispersione",
    "radar-chart": "Grafico radar",
    "polar-area-chart": "Grafico polare",
    confirm: "Conferma",
    "tab-color": "Colore scheda",
    "unhide-sheet": "Mostra foglio:",
    "wrap-text": "Testo a capo",
    import: "Importa {{format}}",
    export: "Esporta {{format}}",
    "floating-picture": "Immagine mobile",
    chart: "Grafico",
    "chart-title": "Titolo del grafico",
    duplicate: "Duplica",
    "select-data": "Seleziona dati",
    "change-chart-title": "Cambia titolo del grafico",
    "change-chart-type": "Cambia tipo di grafico",
    "save-as-picture": "Salva come immagine",
    "reset-size": "Reimposta dimensioni",
    "add-at-the-bottom": "Aggiungi in fondo",
    "row-height": "Altezza riga",
    "column-width": "Larghezza colonna",
    "greater-than-zero": "Il valore deve essere maggiore o uguale a 0",
    "insert-columns": "Inserisci colonne",
    "delete-columns": "Elimina colonne",
    "hide-columns": "Nascondi colonne",
    "unhide-columns": "Mostra colonne",
    "insert-rows": "Inserisci righe",
    "delete-rows": "Elimina righe",
    "hide-rows": "Nascondi righe",
    "unhide-rows": "Mostra righe",
    "reference-is-empty": "Il riferimento è vuoto",
    "reference-is-not-valid": "Il riferimento non è valido",
    "the-value-cannot-be-empty": "Il valore non può essere vuoto",
    "edit-data-source": "Modifica origine dati",
    "sheet-is-not-exist": "Il foglio non esiste",
    "sheet-id-is-duplicate": "L'ID del foglio è duplicato",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "Una cartella di lavoro deve contenere almeno un foglio di lavoro visibile",
    "sheet-name-is-duplicate": "Il nome del foglio è duplicato",
    "uuid-is-duplicate": "L'UUID è duplicato",
    "cells-must-contain-data": "Le celle selezionate devono contenere dati",
    "image-source-is-empty": "L'origine dell'immagine è vuota",
    "merging-cell-is-duplicate": "La cella unita è duplicata",
    "insert-column-left": "Inserisci colonna a sinistra",
    "insert-column-right": "Inserisci colonna a destra",
    "insert-row-above": "Inserisci riga sopra",
    "insert-row-below": "Inserisci riga sotto",
    "get-all-installed-fonts": "Ottieni tutti i caratteri installati localmente",
    "unsupported-chart-types": "Tipi di grafico non supportati",
    "merge-cell": "Unisci celle",
    "merge-and-center": "Unisci e centra",
    "merge-cells": "Unisci celle",
    "merge-content": "Unisci contenuto",
    none: "nessuno",
    "single-underline": "sottolineatura singola",
    "double-underline": "sottolineatura doppia",
    general: "Generale",
    number: "Numero",
    currency: "Valuta",
    accounting: "Contabilità",
    "short-date": "Data breve",
    "long-date": "Data lunga",
    time: "Ora",
    percentage: "Percentuale",
    fraction: "Frazione",
    scientific: "Scientifico",
    text: "Testo",
    "more-number-formats": "Altri formati numerici",
    "short-date-format": "dd/mm/yyyy;@",
    "long-date-format": "dddd d mmmm yyyy;@",
    "time-format": "hh:mm:ss",
    "sheet-id-can-not-be-empty": "Il SheetId non può essere vuoto",
    "no-border": "Nessun bordo",
    "all-borders": "Tutti i bordi",
    "outside-borders": "Bordi esterni",
    "thick-box-border": "Bordo riquadro spesso",
    "bottom-border": "Bordo inferiore",
    "top-border": "Bordo superiore",
    "left-border": "Bordo sinistro",
    "right-border": "Bordo destro",
    "line-color": "Colore linea",
    "line-style": "Stile linea",
    filter: "Filtro",
    "filter-empty": "(vuoto)",
    "filter-all": "Tutto",
    "new-file": "Nuovo file",
    "rename-file": "Rinomina",
    "default-name": "nome predefinito",
    "change-file-name": "Cambia nome file",
    "user-name": "Utente",
    "no-login-editing": "Non puoi modificare un documento finché non effettui l'accesso",
    login: "Accedi",
    "log-out": "Esci"
  },
  "pt-BR": {
    copy: "Copiar",
    cut: "Recortar",
    paste: "Colar",
    delete: "Excluir",
    add: "Adicionar",
    rows: "Linhas",
    menu: "Menu",
    file: "Arquivo",
    insert: "Inserir",
    rename: "Renomear",
    reset: "Redefinir",
    hide: "Ocultar",
    unhide: "Exibir",
    cancel: "Cancelar",
    "line-chart": "Gráfico de linhas",
    "bar-chart": "Gráfico de barras",
    "pie-chart": "Gráfico de pizza",
    "scatter-chart": "Gráfico de dispersão",
    "radar-chart": "Gráfico de radar",
    "polar-area-chart": "Gráfico polar",
    confirm: "Confirmar",
    "tab-color": "Cor da aba",
    "unhide-sheet": "Exibir planilha:",
    "wrap-text": "Quebrar texto",
    import: "Importar {{format}}",
    export: "Exportar {{format}}",
    "floating-picture": "Imagem flutuante",
    chart: "Gráfico",
    "chart-title": "Título do gráfico",
    duplicate: "Duplicar",
    "select-data": "Selecionar dados",
    "change-chart-title": "Alterar título do gráfico",
    "change-chart-type": "Alterar tipo de gráfico",
    "save-as-picture": "Salvar como imagem",
    "reset-size": "Redefinir tamanho",
    "add-at-the-bottom": "Adicionar na parte inferior",
    "row-height": "Altura da linha",
    "column-width": "Largura da coluna",
    "greater-than-zero": "O valor deve ser maior ou igual a 0",
    "insert-columns": "Inserir colunas",
    "delete-columns": "Excluir colunas",
    "hide-columns": "Ocultar colunas",
    "unhide-columns": "Exibir colunas",
    "insert-rows": "Inserir linhas",
    "delete-rows": "Excluir linhas",
    "hide-rows": "Ocultar linhas",
    "unhide-rows": "Exibir linhas",
    "reference-is-empty": "A referência está vazia",
    "reference-is-not-valid": "A referência não é válida",
    "the-value-cannot-be-empty": "O valor não pode estar vazio",
    "edit-data-source": "Editar fonte de dados",
    "sheet-is-not-exist": "A planilha não existe",
    "sheet-id-is-duplicate": "O ID da planilha está duplicado",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "Uma pasta de trabalho deve conter pelo menos uma planilha visível",
    "sheet-name-is-duplicate": "O nome da planilha está duplicado",
    "uuid-is-duplicate": "O UUID está duplicado",
    "cells-must-contain-data": "As células selecionadas devem conter dados",
    "image-source-is-empty": "A origem da imagem está vazia",
    "merging-cell-is-duplicate": "A célula mesclada está duplicada",
    "insert-column-left": "Inserir coluna à esquerda",
    "insert-column-right": "Inserir coluna à direita",
    "insert-row-above": "Inserir linha acima",
    "insert-row-below": "Inserir linha abaixo",
    "get-all-installed-fonts": "Obter todas as fontes instaladas localmente",
    "unsupported-chart-types": "Tipos de gráfico não suportados",
    "merge-cell": "Mesclar células",
    "merge-and-center": "Mesclar e centralizar",
    "merge-cells": "Mesclar células",
    "merge-content": "Mesclar conteúdo",
    none: "nenhum",
    "single-underline": "sublinhado simples",
    "double-underline": "sublinhado duplo",
    general: "Geral",
    number: "Número",
    currency: "Moeda",
    accounting: "Contabilidade",
    "short-date": "Data abreviada",
    "long-date": "Data por extenso",
    time: "Hora",
    percentage: "Porcentagem",
    fraction: "Fração",
    scientific: "Científico",
    text: "Texto",
    "more-number-formats": "Mais formatos de número",
    "short-date-format": "dd/mm/yyyy;@",
    "long-date-format": "dddd, d de mmmm de yyyy;@",
    "time-format": "hh:mm:ss",
    "sheet-id-can-not-be-empty": "O SheetId não pode estar vazio",
    "no-border": "Sem borda",
    "all-borders": "Todas as bordas",
    "outside-borders": "Bordas externas",
    "thick-box-border": "Borda de caixa espessa",
    "bottom-border": "Borda inferior",
    "top-border": "Borda superior",
    "left-border": "Borda esquerda",
    "right-border": "Borda direita",
    "line-color": "Cor da linha",
    "line-style": "Estilo da linha",
    filter: "Filtro",
    "filter-empty": "(vazio)",
    "filter-all": "Tudo",
    "new-file": "Novo arquivo",
    "rename-file": "Renomear",
    "default-name": "nome padrão",
    "change-file-name": "Alterar nome do arquivo",
    "user-name": "Usuário",
    "no-login-editing": "Você não pode editar um documento até fazer login",
    login: "Entrar",
    "log-out": "Sair"
  },
  "pt-PT": {
    copy: "Copiar",
    cut: "Cortar",
    paste: "Colar",
    delete: "Eliminar",
    add: "Adicionar",
    rows: "Linhas",
    menu: "Menu",
    file: "Ficheiro",
    insert: "Inserir",
    rename: "Mudar nome",
    reset: "Repor",
    hide: "Ocultar",
    unhide: "Mostrar",
    cancel: "Cancelar",
    "line-chart": "Gráfico de linhas",
    "bar-chart": "Gráfico de barras",
    "pie-chart": "Gráfico circular",
    "scatter-chart": "Gráfico de dispersão",
    "radar-chart": "Gráfico de radar",
    "polar-area-chart": "Gráfico polar",
    confirm: "Confirmar",
    "tab-color": "Cor do separador",
    "unhide-sheet": "Mostrar folha:",
    "wrap-text": "Quebrar texto",
    import: "Importar {{format}}",
    export: "Exportar {{format}}",
    "floating-picture": "Imagem flutuante",
    chart: "Gráfico",
    "chart-title": "Título do gráfico",
    duplicate: "Duplicar",
    "select-data": "Seleccionar dados",
    "change-chart-title": "Alterar título do gráfico",
    "change-chart-type": "Alterar tipo de gráfico",
    "save-as-picture": "Guardar como imagem",
    "reset-size": "Repor tamanho",
    "add-at-the-bottom": "Adicionar na parte inferior",
    "row-height": "Altura da linha",
    "column-width": "Largura da coluna",
    "greater-than-zero": "O valor deve ser maior ou igual a 0",
    "insert-columns": "Inserir colunas",
    "delete-columns": "Eliminar colunas",
    "hide-columns": "Ocultar colunas",
    "unhide-columns": "Mostrar colunas",
    "insert-rows": "Inserir linhas",
    "delete-rows": "Eliminar linhas",
    "hide-rows": "Ocultar linhas",
    "unhide-rows": "Mostrar linhas",
    "reference-is-empty": "A referência está vazia",
    "reference-is-not-valid": "A referência não é válida",
    "the-value-cannot-be-empty": "O valor não pode estar vazio",
    "edit-data-source": "Editar origem de dados",
    "sheet-is-not-exist": "A folha não existe",
    "sheet-id-is-duplicate": "O ID da folha está duplicado",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "Um livro deve conter pelo menos uma folha de cálculo visível",
    "sheet-name-is-duplicate": "O nome da folha está duplicado",
    "uuid-is-duplicate": "O UUID está duplicado",
    "cells-must-contain-data": "As células seleccionadas devem conter dados",
    "image-source-is-empty": "A origem da imagem está vazia",
    "merging-cell-is-duplicate": "A célula intercalada está duplicada",
    "insert-column-left": "Inserir coluna à esquerda",
    "insert-column-right": "Inserir coluna à direita",
    "insert-row-above": "Inserir linha acima",
    "insert-row-below": "Inserir linha abaixo",
    "get-all-installed-fonts": "Obter todos os tipos de letra instalados localmente",
    "unsupported-chart-types": "Tipos de gráfico não suportados",
    "merge-cell": "Intercalar células",
    "merge-and-center": "Intercalar e centrar",
    "merge-cells": "Intercalar células",
    "merge-content": "Intercalar conteúdo",
    none: "nenhum",
    "single-underline": "sublinhado simples",
    "double-underline": "sublinhado duplo",
    general: "Geral",
    number: "Número",
    currency: "Moeda",
    accounting: "Contabilidade",
    "short-date": "Data abreviada",
    "long-date": "Data por extenso",
    time: "Hora",
    percentage: "Percentagem",
    fraction: "Fracção",
    scientific: "Científico",
    text: "Texto",
    "more-number-formats": "Mais formatos de número",
    "short-date-format": "dd/mm/yyyy;@",
    "long-date-format": "dddd, d de mmmm de yyyy;@",
    "time-format": "hh:mm:ss",
    "sheet-id-can-not-be-empty": "O SheetId não pode estar vazio",
    "no-border": "Sem limite",
    "all-borders": "Todos os limites",
    "outside-borders": "Limites exteriores",
    "thick-box-border": "Limite de caixa espesso",
    "bottom-border": "Limite inferior",
    "top-border": "Limite superior",
    "left-border": "Limite esquerdo",
    "right-border": "Limite direito",
    "line-color": "Cor da linha",
    "line-style": "Estilo da linha",
    filter: "Filtro",
    "filter-empty": "(vazio)",
    "filter-all": "Tudo",
    "new-file": "Novo ficheiro",
    "rename-file": "Mudar nome",
    "default-name": "nome predefinido",
    "change-file-name": "Alterar nome do ficheiro",
    "user-name": "Utilizador",
    "no-login-editing": "Não pode editar um documento até iniciar sessão",
    login: "Iniciar sessão",
    "log-out": "Terminar sessão"
  },
  "nl-NL": {
    copy: "Kopiëren",
    cut: "Knippen",
    paste: "Plakken",
    delete: "Verwijderen",
    add: "Toevoegen",
    rows: "Rijen",
    menu: "Menu",
    file: "Bestand",
    insert: "Invoegen",
    rename: "Hernoemen",
    reset: "Herstellen",
    hide: "Verbergen",
    unhide: "Zichtbaar maken",
    cancel: "Annuleren",
    "line-chart": "Lijndiagram",
    "bar-chart": "Staafdiagram",
    "pie-chart": "Cirkeldiagram",
    "scatter-chart": "Spreidingsdiagram",
    "radar-chart": "Radardiagram",
    "polar-area-chart": "Polair vlakdiagram",
    confirm: "Bevestigen",
    "tab-color": "Tabkleur",
    "unhide-sheet": "Werkblad zichtbaar maken:",
    "wrap-text": "Tekst terugloop",
    import: "{{format}} importeren",
    export: "{{format}} exporteren",
    "floating-picture": "Zwevende afbeelding",
    chart: "Diagram",
    "chart-title": "Diagramtitel",
    duplicate: "Dupliceren",
    "select-data": "Gegevens selecteren",
    "change-chart-title": "Diagramtitel wijzigen",
    "change-chart-type": "Diagramtype wijzigen",
    "save-as-picture": "Opslaan als afbeelding",
    "reset-size": "Grootte herstellen",
    "add-at-the-bottom": "Onderaan toevoegen",
    "row-height": "Rijhoogte",
    "column-width": "Kolombreedte",
    "greater-than-zero": "De waarde moet groter dan of gelijk aan 0 zijn",
    "insert-columns": "Kolommen invoegen",
    "delete-columns": "Kolommen verwijderen",
    "hide-columns": "Kolommen verbergen",
    "unhide-columns": "Kolommen zichtbaar maken",
    "insert-rows": "Rijen invoegen",
    "delete-rows": "Rijen verwijderen",
    "hide-rows": "Rijen verbergen",
    "unhide-rows": "Rijen zichtbaar maken",
    "reference-is-empty": "De verwijzing is leeg",
    "reference-is-not-valid": "De verwijzing is niet geldig",
    "the-value-cannot-be-empty": "De waarde mag niet leeg zijn",
    "edit-data-source": "Gegevensbron bewerken",
    "sheet-is-not-exist": "Het werkblad bestaat niet",
    "sheet-id-is-duplicate": "Het werkblad-ID is gedupliceerd",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "Een werkmap moet minstens één zichtbaar werkblad bevatten",
    "sheet-name-is-duplicate": "De werkbladnaam is gedupliceerd",
    "uuid-is-duplicate": "De UUID is gedupliceerd",
    "cells-must-contain-data": "De geselecteerde cellen moeten gegevens bevatten",
    "image-source-is-empty": "De afbeeldingsbron is leeg",
    "merging-cell-is-duplicate": "De samengevoegde cel is gedupliceerd",
    "insert-column-left": "Kolom links invoegen",
    "insert-column-right": "Kolom rechts invoegen",
    "insert-row-above": "Rij hierboven invoegen",
    "insert-row-below": "Rij hieronder invoegen",
    "get-all-installed-fonts": "Alle lokaal geïnstalleerde lettertypen ophalen",
    "unsupported-chart-types": "Niet-ondersteunde diagramtypen",
    "merge-cell": "Cellen samenvoegen",
    "merge-and-center": "Samenvoegen en centreren",
    "merge-cells": "Cellen samenvoegen",
    "merge-content": "Inhoud samenvoegen",
    none: "geen",
    "single-underline": "enkele onderstreping",
    "double-underline": "dubbele onderstreping",
    general: "Algemeen",
    number: "Getal",
    currency: "Valuta",
    accounting: "Boekhouding",
    "short-date": "Korte datum",
    "long-date": "Lange datum",
    time: "Tijd",
    percentage: "Percentage",
    fraction: "Breuk",
    scientific: "Wetenschappelijk",
    text: "Tekst",
    "more-number-formats": "Meer getalformaten",
    "short-date-format": "dd-mm-yyyy;@",
    "long-date-format": "dddd d mmmm yyyy;@",
    "time-format": "uu:mm:ss",
    "sheet-id-can-not-be-empty": "Het SheetId mag niet leeg zijn",
    "no-border": "Geen rand",
    "all-borders": "Alle randen",
    "outside-borders": "Buitenranden",
    "thick-box-border": "Dikke kastkader",
    "bottom-border": "Onderrand",
    "top-border": "Bovenrand",
    "left-border": "Linkerrand",
    "right-border": "Rechterrand",
    "line-color": "Lijnkleur",
    "line-style": "Lijnstijl",
    filter: "Filter",
    "filter-empty": "(leeg)",
    "filter-all": "Alles",
    "new-file": "Nieuw bestand",
    "rename-file": "Hernoemen",
    "default-name": "standaardnaam",
    "change-file-name": "Bestandsnaam wijzigen",
    "user-name": "Gebruiker",
    "no-login-editing": "U kunt geen document bewerken totdat u bent ingelogd",
    login: "Inloggen",
    "log-out": "Uitloggen"
  },
  "ar-EG": {
    copy: "نسخ",
    cut: "قص",
    paste: "لصق",
    delete: "حذف",
    add: "إضافة",
    rows: "صفوف",
    menu: "قائمة",
    file: "ملف",
    insert: "إدراج",
    rename: "إعادة تسمية",
    reset: "إعادة تعيين",
    hide: "إخفاء",
    unhide: "إظهار",
    cancel: "إلغاء",
    "line-chart": "مخطط خطي",
    "bar-chart": "مخطط أعمدة",
    "pie-chart": "مخطط دائري",
    "scatter-chart": "مخطط مبعثر",
    "radar-chart": "مخطط راداري",
    "polar-area-chart": "مخطط منطقة قطبية",
    confirm: "تأكيد",
    "tab-color": "لون التبويب",
    "unhide-sheet": "إظهار الورقة:",
    "wrap-text": "التفاف النص",
    import: "استيراد {{format}}",
    export: "تصدير {{format}}",
    "floating-picture": "صورة عائمة",
    chart: "مخطط",
    "chart-title": "عنوان المخطط",
    duplicate: "تكرار",
    "select-data": "تحديد البيانات",
    "change-chart-title": "تغيير عنوان المخطط",
    "change-chart-type": "تغيير نوع المخطط",
    "save-as-picture": "حفظ كصورة",
    "reset-size": "إعادة تعيين الحجم",
    "add-at-the-bottom": "إضافة في الأسفل",
    "row-height": "ارتفاع الصف",
    "column-width": "عرض العمود",
    "greater-than-zero": "يجب أن تكون القيمة أكبر من أو تساوي 0",
    "insert-columns": "إدراج أعمدة",
    "delete-columns": "حذف أعمدة",
    "hide-columns": "إخفاء أعمدة",
    "unhide-columns": "إظهار أعمدة",
    "insert-rows": "إدراج صفوف",
    "delete-rows": "حذف صفوف",
    "hide-rows": "إخفاء صفوف",
    "unhide-rows": "إظهار صفوف",
    "reference-is-empty": "المرجع فارغ",
    "reference-is-not-valid": "المرجع غير صالح",
    "the-value-cannot-be-empty": "لا يمكن أن تكون القيمة فارغة",
    "edit-data-source": "تحرير مصدر البيانات",
    "sheet-is-not-exist": "الورقة غير موجودة",
    "sheet-id-is-duplicate": "معرف الورقة مكرر",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "يجب أن يحتوي المصنف على ورقة عمل مرئية واحدة على الأقل",
    "sheet-name-is-duplicate": "اسم الورقة مكرر",
    "uuid-is-duplicate": "المعرف الفريد مكرر",
    "cells-must-contain-data": "يجب أن تحتوي الخلايا المحددة على بيانات",
    "image-source-is-empty": "مصدر الصورة فارغ",
    "merging-cell-is-duplicate": "الخلية المدمجة مكررة",
    "insert-column-left": "إدراج عمود على اليسار",
    "insert-column-right": "إدراج عمود على اليمين",
    "insert-row-above": "إدراج صف أعلى",
    "insert-row-below": "إدراج صف أسفل",
    "get-all-installed-fonts": "الحصول على جميع الخطوط المثبتة محلياً",
    "unsupported-chart-types": "أنواع مخططات غير مدعومة",
    "merge-cell": "دمج الخلايا",
    "merge-and-center": "دمج وتوسيط",
    "merge-cells": "دمج الخلايا",
    "merge-content": "دمج المحتوى",
    none: "لا شيء",
    "single-underline": "تسطير مفرد",
    "double-underline": "تسطير مزدوج",
    general: "عام",
    number: "رقم",
    currency: "عملة",
    accounting: "محاسبة",
    "short-date": "تاريخ مختصر",
    "long-date": "تاريخ طويل",
    time: "وقت",
    percentage: "نسبة مئوية",
    fraction: "كسر",
    scientific: "علمي",
    text: "نص",
    "more-number-formats": "المزيد من تنسيقات الأرقام",
    "short-date-format": "dd/mm/yyyy;@",
    "long-date-format": "dddd، d mmmm yyyy;@",
    "time-format": "hh:mm:ss",
    "sheet-id-can-not-be-empty": "لا يمكن أن يكون معرف الورقة فارغاً",
    "no-border": "بدون حدود",
    "all-borders": "جميع الحدود",
    "outside-borders": "الحدود الخارجية",
    "thick-box-border": "حدود صندوق سميك",
    "bottom-border": "الحد السفلي",
    "top-border": "الحد العلوي",
    "left-border": "الحد الأيسر",
    "right-border": "الحد الأيمن",
    "line-color": "لون الخط",
    "line-style": "نمط الخط",
    filter: "تصفية",
    "filter-empty": "(فارغ)",
    "filter-all": "الكل",
    "new-file": "ملف جديد",
    "rename-file": "إعادة تسمية",
    "default-name": "الاسم الافتراضي",
    "change-file-name": "تغيير اسم الملف",
    "user-name": "المستخدم",
    "no-login-editing": "لا يمكنك تحرير مستند حتى تسجل الدخول",
    login: "تسجيل الدخول",
    "log-out": "تسجيل الخروج"
  },
  "ar-SA": {
    copy: "نسخ",
    cut: "قص",
    paste: "لصق",
    delete: "حذف",
    add: "إضافة",
    rows: "صفوف",
    menu: "قائمة",
    file: "ملف",
    insert: "إدراج",
    rename: "إعادة تسمية",
    reset: "إعادة تعيين",
    hide: "إخفاء",
    unhide: "إظهار",
    cancel: "إلغاء",
    "line-chart": "مخطط خطي",
    "bar-chart": "مخطط أعمدة",
    "pie-chart": "مخطط دائري",
    "scatter-chart": "مخطط مبعثر",
    "radar-chart": "مخطط راداري",
    "polar-area-chart": "مخطط منطقة قطبية",
    confirm: "تأكيد",
    "tab-color": "لون التبويب",
    "unhide-sheet": "إظهار الورقة:",
    "wrap-text": "التفاف النص",
    import: "استيراد {{format}}",
    export: "تصدير {{format}}",
    "floating-picture": "صورة عائمة",
    chart: "مخطط",
    "chart-title": "عنوان المخطط",
    duplicate: "تكرار",
    "select-data": "تحديد البيانات",
    "change-chart-title": "تغيير عنوان المخطط",
    "change-chart-type": "تغيير نوع المخطط",
    "save-as-picture": "حفظ كصورة",
    "reset-size": "إعادة تعيين الحجم",
    "add-at-the-bottom": "إضافة في الأسفل",
    "row-height": "ارتفاع الصف",
    "column-width": "عرض العمود",
    "greater-than-zero": "يجب أن تكون القيمة أكبر من أو تساوي 0",
    "insert-columns": "إدراج أعمدة",
    "delete-columns": "حذف أعمدة",
    "hide-columns": "إخفاء أعمدة",
    "unhide-columns": "إظهار أعمدة",
    "insert-rows": "إدراج صفوف",
    "delete-rows": "حذف صفوف",
    "hide-rows": "إخفاء صفوف",
    "unhide-rows": "إظهار صفوف",
    "reference-is-empty": "المرجع فارغ",
    "reference-is-not-valid": "المرجع غير صالح",
    "the-value-cannot-be-empty": "لا يمكن أن تكون القيمة فارغة",
    "edit-data-source": "تحرير مصدر البيانات",
    "sheet-is-not-exist": "الورقة غير موجودة",
    "sheet-id-is-duplicate": "معرف الورقة مكرر",
    "a-workbook-must-contains-at-least-one-visible-worksheet": "يجب أن يحتوي المصنف على ورقة عمل مرئية واحدة على الأقل",
    "sheet-name-is-duplicate": "اسم الورقة مكرر",
    "uuid-is-duplicate": "المعرف الفريد مكرر",
    "cells-must-contain-data": "يجب أن تحتوي الخلايا المحددة على بيانات",
    "image-source-is-empty": "مصدر الصورة فارغ",
    "merging-cell-is-duplicate": "الخلية المدمجة مكررة",
    "insert-column-left": "إدراج عمود على اليسار",
    "insert-column-right": "إدراج عمود على اليمين",
    "insert-row-above": "إدراج صف أعلى",
    "insert-row-below": "إدراج صف أسفل",
    "get-all-installed-fonts": "الحصول على جميع الخطوط المثبتة محلياً",
    "unsupported-chart-types": "أنواع مخططات غير مدعومة",
    "merge-cell": "دمج الخلايا",
    "merge-and-center": "دمج وتوسيط",
    "merge-cells": "دمج الخلايا",
    "merge-content": "دمج المحتوى",
    none: "لا شيء",
    "single-underline": "تسطير مفرد",
    "double-underline": "تسطير مزدوج",
    general: "عام",
    number: "رقم",
    currency: "عملة",
    accounting: "محاسبة",
    "short-date": "تاريخ مختصر",
    "long-date": "تاريخ طويل",
    time: "وقت",
    percentage: "نسبة مئوية",
    fraction: "كسر",
    scientific: "علمي",
    text: "نص",
    "more-number-formats": "المزيد من تنسيقات الأرقام",
    "short-date-format": "dd/mm/yyyy;@",
    "long-date-format": "dddd، d mmmm yyyy;@",
    "time-format": "hh:mm:ss",
    "sheet-id-can-not-be-empty": "لا يمكن أن يكون معرف الورقة فارغاً",
    "no-border": "بدون حدود",
    "all-borders": "جميع الحدود",
    "outside-borders": "الحدود الخارجية",
    "thick-box-border": "حدود صندوق سميك",
    "bottom-border": "الحد السفلي",
    "top-border": "الحد العلوي",
    "left-border": "الحد الأيسر",
    "right-border": "الحد الأيمن",
    "line-color": "لون الخط",
    "line-style": "نمط الخط",
    filter: "تصفية",
    "filter-empty": "(فارغ)",
    "filter-all": "الكل",
    "new-file": "ملف جديد",
    "rename-file": "إعادة تسمية",
    "default-name": "الاسم الافتراضي",
    "change-file-name": "تغيير اسم الملف",
    "user-name": "المستخدم",
    "no-login-editing": "لا يمكنك تحرير مستند حتى تسجل الدخول",
    login: "تسجيل الدخول",
    "log-out": "تسجيل الخروج"
  }
};
let Rs = "excel-language";
function Fg() {
  const e = "en-US";
  let t = e;
  const n = /* @__PURE__ */ new Set();
  function r(s) {
    const l = s || navigator?.language || e;
    if (Xo.includes(l))
      return l;
    const f = Xo.find((d) => d.includes(l));
    return f || e;
  }
  function o() {
    for (const s of n) s();
  }
  return {
    changeLanguage: (s) => {
      if (s === t)
        return;
      const l = r(s);
      t = l, localStorage.setItem(Rs, l), o();
    },
    init: () => {
      t = r(localStorage.getItem(Rs));
    },
    t: (s, l = {}) => Lg[t][s].replace(/{([a-z]+)}/gi, (d, c) => {
      if (c in l)
        return l[c];
      throw new Error(`i18n.t not found key: "${c}"`);
    }),
    subscribe: (s) => (n.add(s), () => {
      n.delete(s);
    }),
    get current() {
      return t;
    }
  };
}
const k = Fg();
typeof window < "u" && (window.__yallyExcelI18n = k);
const Ji = ue((e) => {
  const {
    children: t,
    title: n,
    className: r,
    onCancel: o,
    onOk: s,
    visible: l,
    getContainer: f = () => document.body,
    testId: d
  } = e;
  if (!l)
    return;
  const c = d ? `${d}-cancel` : void 0, i = d ? `${d}-confirm` : void 0;
  return Zc(
    /* @__PURE__ */ I("div", { className: In["dialog-modal"], "data-testid": d, children: /* @__PURE__ */ I("div", { className: ae(In["dialog-container"], r), children: [
      /* @__PURE__ */ I("div", { className: In["dialog-title"], children: n }, void 0, !1, {
        fileName: "/home/user/excel-collab/src/components/Dialog/index.tsx",
        lineNumber: 39,
        columnNumber: 9
      }, void 0),
      /* @__PURE__ */ I("div", { className: In["dialog-content"], children: t }, void 0, !1, {
        fileName: "/home/user/excel-collab/src/components/Dialog/index.tsx",
        lineNumber: 40,
        columnNumber: 9
      }, void 0),
      /* @__PURE__ */ I("div", { className: In["dialog-button"], children: [
        /* @__PURE__ */ I(V, { onClick: o, testId: c, children: k.t("cancel") }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/components/Dialog/index.tsx",
          lineNumber: 42,
          columnNumber: 11
        }, void 0),
        /* @__PURE__ */ I(
          V,
          {
            onClick: s,
            className: In["dialog-cancel"],
            type: "primary",
            testId: i,
            children: k.t("confirm")
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/components/Dialog/index.tsx",
            lineNumber: 45,
            columnNumber: 11
          },
          void 0
        )
      ] }, void 0, !0, {
        fileName: "/home/user/excel-collab/src/components/Dialog/index.tsx",
        lineNumber: 41,
        columnNumber: 9
      }, void 0)
    ] }, void 0, !0, {
      fileName: "/home/user/excel-collab/src/components/Dialog/index.tsx",
      lineNumber: 38,
      columnNumber: 7
    }, void 0) }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/components/Dialog/index.tsx",
      lineNumber: 37,
      columnNumber: 5
    }, void 0),
    f()
  );
});
Ji.displayName = "Dialog";
function Zn(e) {
  const t = document.createDocumentFragment();
  let n;
  function r() {
    n?.unmount(), n = void 0;
  }
  function o(l) {
    n = n || la(t), n.render(
      /* @__PURE__ */ I(
        Ji,
        {
          visible: l.visible,
          title: l.title,
          className: l.className,
          onCancel: (f) => {
            f.stopPropagation(), l.onCancel && l.onCancel(f), r();
          },
          onOk: (f) => {
            f.stopPropagation(), l.onOk && l.onOk(f), r();
          },
          testId: l.testId,
          children: l.children
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/components/Dialog/index.tsx",
          lineNumber: 72,
          columnNumber: 7
        },
        this
      )
    );
  }
  function s(l) {
    o(l);
  }
  return o(e), {
    close: r,
    update: s
  };
}
const Pg = "_container_uqw8i_1", Ug = "_portal_uqw8i_5", _g = "_menuContainer_uqw8i_13", Bg = "_bottom_uqw8i_16", zg = "_subMenuContainer_uqw8i_21", Hg = "_menu_uqw8i_13", Vg = "_small_uqw8i_35", Gg = "_menuItem_uqw8i_35", $g = "_active_uqw8i_53", qg = "_trigger_uqw8i_57", ot = {
  container: Pg,
  portal: Ug,
  menuContainer: _g,
  bottom: Bg,
  subMenuContainer: zg,
  menu: Hg,
  small: Vg,
  menuItem: Gg,
  active: $g,
  trigger: qg
}, Le = ({ onClick: e, children: t, testId: n, active: r = !1 }) => /* @__PURE__ */ I(
  "li",
  {
    className: ae(ot.menuItem, { [ot.active]: r }),
    onClick: e,
    "data-testid": n,
    children: t
  },
  void 0,
  !1,
  {
    fileName: "/home/user/excel-collab/src/components/Menu/index.tsx",
    lineNumber: 31,
    columnNumber: 5
  },
  void 0
), Wg = ({ label: e, children: t, testId: n, className: r, portalClassName: o }) => {
  const [s, l] = he(!1), f = z(() => {
    l((d) => !d);
  }, []);
  return /* @__PURE__ */ I(
    "li",
    {
      className: ae(ot.menuItem, r),
      onClick: f,
      "data-testid": n,
      children: [
        /* @__PURE__ */ I("div", { children: e }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/components/Menu/index.tsx",
          lineNumber: 54,
          columnNumber: 7
        }, void 0),
        /* @__PURE__ */ I(
          "div",
          {
            className: ae(
              ot.subMenuContainer,
              ot.portal,
              o
            ),
            hidden: !s,
            children: /* @__PURE__ */ I("ul", { className: ot.menu, children: t }, void 0, !1, {
              fileName: "/home/user/excel-collab/src/components/Menu/index.tsx",
              lineNumber: 63,
              columnNumber: 9
            }, void 0)
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/components/Menu/index.tsx",
            lineNumber: 55,
            columnNumber: 7
          },
          void 0
        )
      ]
    },
    void 0,
    !0,
    {
      fileName: "/home/user/excel-collab/src/components/Menu/index.tsx",
      lineNumber: 49,
      columnNumber: 5
    },
    void 0
  );
}, Qi = ue(
  ({
    label: e,
    children: t,
    testId: n,
    className: r,
    isPlain: o = !1,
    position: s = "right",
    size: l = "normal",
    portalClassName: f
  }) => {
    const d = o ? "plain" : void 0, [c, i] = he(!1), a = z(() => {
      i((g) => !g);
    }, []), u = z(() => {
      i(!1);
    }, []), h = bn(c, u);
    return /* @__PURE__ */ I(
      "div",
      {
        className: ae(ot.container, r, {
          [ot.small]: l === "small"
        }),
        ref: h,
        "data-testid": n,
        children: [
          /* @__PURE__ */ I(
            V,
            {
              onClick: a,
              testId: `${n}-trigger`,
              type: d,
              className: ot.trigger,
              children: e
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/components/Menu/index.tsx",
              lineNumber: 97,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            "div",
            {
              className: ae(
                ot.menuContainer,
                ot.portal,
                f,
                {
                  [ot.bottom]: s === "bottom"
                }
              ),
              "data-testid": `${n}-portal`,
              hidden: !c,
              children: /* @__PURE__ */ I("ul", { className: ot.menu, children: t }, void 0, !1, {
                fileName: "/home/user/excel-collab/src/components/Menu/index.tsx",
                lineNumber: 118,
                columnNumber: 11
              }, void 0)
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/components/Menu/index.tsx",
              lineNumber: 106,
              columnNumber: 9
            },
            void 0
          )
        ]
      },
      void 0,
      !0,
      {
        fileName: "/home/user/excel-collab/src/components/Menu/index.tsx",
        lineNumber: 90,
        columnNumber: 7
      },
      void 0
    );
  }
);
Qi.displayName = "Menu";
const Yg = "_container_ivcy9_1", Xg = "_content_ivcy9_7", Kg = "_toast_ivcy9_10", Jg = "_icon_ivcy9_19", Qg = "_success_ivcy9_22", jg = "_success_icon_ivcy9_26", Zg = "_error_ivcy9_29", ep = "_error_icon_ivcy9_33", tp = "_info_ivcy9_36", np = "_info_icon_ivcy9_40", rp = "_warning_ivcy9_43", op = "_warning_icon_ivcy9_47", Mn = {
  container: Yg,
  content: Xg,
  toast: Kg,
  icon: Jg,
  success: Qg,
  success_icon: jg,
  error: Zg,
  error_icon: ep,
  info: tp,
  info_icon: np,
  warning: rp,
  warning_icon: op
}, ip = ({
  message: e,
  type: t,
  testId: n
}) => /* @__PURE__ */ I(
  "div",
  {
    className: ae(Mn.toast, Mn[t]),
    "data-testid": n,
    children: [
      /* @__PURE__ */ I(
        Ae,
        {
          name: t,
          className: ae(Mn[`${t}_icon`], Mn.icon)
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/components/Toast/index.tsx",
          lineNumber: 25,
          columnNumber: 7
        },
        void 0
      ),
      /* @__PURE__ */ I("div", { className: Mn.content, children: e }, void 0, !1, {
        fileName: "/home/user/excel-collab/src/components/Toast/index.tsx",
        lineNumber: 29,
        columnNumber: 7
      }, void 0)
    ]
  },
  void 0,
  !0,
  {
    fileName: "/home/user/excel-collab/src/components/Toast/index.tsx",
    lineNumber: 21,
    columnNumber: 5
  },
  void 0
);
function Pe(e) {
  const { duration: t = 3, ...n } = e;
  let r = document.createElement("div");
  r.className = Mn.container, document.body.appendChild(r);
  const o = la(r);
  o.render(/* @__PURE__ */ I(ip, { ...n }, void 0, !1, {
    fileName: "/home/user/excel-collab/src/components/Toast/index.tsx",
    lineNumber: 40,
    columnNumber: 15
  }, this));
  function s() {
    r && (o.unmount(), document.body.removeChild(r), r = void 0);
  }
  return setTimeout(s, t * 1e3), s;
}
Pe.error = function(e, t = "error-toast") {
  return Pe({ message: e, type: "error", testId: t });
};
Pe.info = function(e, t = "info-toast") {
  return Pe({ message: e, type: "info", testId: t });
};
Pe.warning = function(e, t = "warning-toast") {
  return Pe({ message: e, type: "warning", testId: t });
};
Pe.success = function(e, t = "success-toast") {
  return Pe({ message: e, type: "success", testId: t });
};
const sp = "_container_926r9_1", ap = "_loading_926r9_12", As = {
  container: sp,
  loading: ap
}, Co = ue(() => /* @__PURE__ */ I("div", { className: As.container, children: /* @__PURE__ */ I("div", { className: As.loading, "data-test": "loading" }, void 0, !1, {
  fileName: "/home/user/excel-collab/src/components/Loading/index.tsx",
  lineNumber: 7,
  columnNumber: 7
}, void 0) }, void 0, !1, {
  fileName: "/home/user/excel-collab/src/components/Loading/index.tsx",
  lineNumber: 6,
  columnNumber: 5
}, void 0));
Co.displayName = "Loading";
const lp = "_top_18907_16", cp = "_right_18907_20", up = "_reset_18907_46", zt = {
  "color-picker": "_color-picker_18907_1",
  "color-picker-wrapper": "_color-picker-wrapper_18907_5",
  top: lp,
  right: cp,
  "color-picker-list": "_color-picker-list_18907_26",
  "color-picker-item": "_color-picker-item_18907_32",
  reset: up
}, Ts = (e, t = 0, n = 1) => e > n ? n : e < t ? t : e, Ut = {
  "color-picker-panel": "_color-picker-panel_mnh2i_1",
  "color-picker-panel__saturation": "_color-picker-panel__saturation_mnh2i_11",
  "color-picker-panel__pointer-fill": "_color-picker-panel__pointer-fill_mnh2i_22",
  "color-picker-panel__hue": "_color-picker-panel__hue_mnh2i_38",
  "color-picker-panel__last-control": "_color-picker-panel__last-control_mnh2i_56",
  "color-picker-panel__interactive": "_color-picker-panel__interactive_mnh2i_60",
  "color-picker-panel__pointer": "_color-picker-panel__pointer_mnh2i_22",
  "color-picker-panel__saturation-pointer": "_color-picker-panel__saturation-pointer_mnh2i_88",
  "color-picker-panel__hue-pointer": "_color-picker-panel__hue-pointer_mnh2i_92"
}, fp = 200, dp = 234, xs = (e, t, n) => {
  const r = e.getBoundingClientRect();
  return {
    left: Ts((t - r.left) / (r.width || dp)),
    top: Ts((n - r.top) / (r.height || fp))
  };
}, hp = ({ onMove: e, testId: t, ...n }) => {
  const r = Xe(null), o = bl(e), [s, l] = We(() => {
    const f = (a) => {
      r.current && (a.preventDefault(), r.current.focus(), o(
        xs(r.current, a.clientX, a.clientY)
      ), i(!0));
    }, d = (a) => {
      a.buttons <= 0 || r.current && (a.preventDefault(), o(
        xs(r.current, a.clientX, a.clientY)
      ));
    }, c = () => i(!1);
    function i(a) {
      const u = a ? document.body.addEventListener : document.body.removeEventListener;
      u("pointermove", d), u("pointerup", c);
    }
    return [f, i];
  }, [o]);
  return Re(() => l, [l]), /* @__PURE__ */ I(
    "div",
    {
      ...n,
      onPointerDown: s,
      className: Ut["color-picker-panel__interactive"],
      ref: r,
      tabIndex: 0,
      role: "slider",
      "data-testid": t
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/components/ColorPicker/Interactive.tsx",
      lineNumber: 80,
      columnNumber: 5
    },
    void 0
  );
}, wl = je.memo(hp), Il = ({
  className: e,
  color: t,
  left: n,
  top: r = 0.5
}) => {
  const o = {
    top: `${r * 100}%`,
    left: `${n * 100}%`
  };
  return /* @__PURE__ */ I(
    "div",
    {
      className: ae(Ut["color-picker-panel__pointer"], e),
      style: o,
      children: /* @__PURE__ */ I(
        "div",
        {
          className: Ut["color-picker-panel__pointer-fill"],
          style: { backgroundColor: t }
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/components/ColorPicker/Pointer.tsx",
          lineNumber: 27,
          columnNumber: 7
        },
        void 0
      )
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/components/ColorPicker/Pointer.tsx",
      lineNumber: 23,
      columnNumber: 5
    },
    void 0
  );
}, Fe = (e, t = 0, n = Math.pow(10, t)) => Math.round(n * e) / n, mp = (e) => {
  const t = gp(e);
  return Ep(t);
}, gp = (e) => (e[0] === "#" && (e = e.substring(1)), e.length < 6 ? {
  r: parseInt(e[0] + e[0], 16),
  g: parseInt(e[1] + e[1], 16),
  b: parseInt(e[2] + e[2], 16),
  a: e.length === 4 ? Fe(parseInt(e[3] + e[3], 16) / 255, 2) : 1
} : {
  r: parseInt(e.substring(0, 2), 16),
  g: parseInt(e.substring(2, 4), 16),
  b: parseInt(e.substring(4, 6), 16),
  a: e.length === 8 ? Fe(parseInt(e.substring(6, 8), 16) / 255, 2) : 1
}), pp = (e) => Cp(Np(e)), vp = ({ h: e, s: t, v: n, a: r }) => {
  const o = (200 - t) * n / 100;
  return {
    h: Fe(e),
    s: Fe(
      o > 0 && o < 200 ? t * n / 100 / (o <= 100 ? o : 200 - o) * 100 : 0
    ),
    l: Fe(o / 2),
    a: Fe(r, 2)
  };
}, ii = (e) => {
  const { h: t, s: n, l: r } = vp(e);
  return `hsl(${t}, ${n}%, ${r}%)`;
}, Np = ({ h: e, s: t, v: n, a: r }) => {
  e = e / 360 * 6, t = t / 100, n = n / 100;
  const o = Math.floor(e), s = n * (1 - t), l = n * (1 - (e - o) * t), f = n * (1 - (1 - e + o) * t), d = o % 6;
  return {
    r: Fe([n, l, s, s, f, n][d] * 255),
    g: Fe([f, n, n, l, s, s][d] * 255),
    b: Fe([s, s, f, n, n, l][d] * 255),
    a: Fe(r, 2)
  };
}, yr = (e) => {
  const t = e.toString(16);
  return t.length < 2 ? "0" + t : t;
}, Cp = ({ r: e, g: t, b: n, a: r }) => {
  const o = r < 1 ? yr(Fe(r * 255)) : "";
  return "#" + yr(e) + yr(t) + yr(n) + o;
}, Ep = ({ r: e, g: t, b: n, a: r }) => {
  const o = Math.max(e, t, n), s = o - Math.min(e, t, n), l = s ? o === e ? (t - n) / s : o === t ? 2 + (n - e) / s : 4 + (e - t) / s : 0;
  return {
    h: Fe(60 * (l < 0 ? l + 6 : l)),
    s: Fe(o ? s / o * 100 : 0),
    v: Fe(o / 255 * 100),
    a: r
  };
}, bp = ({ className: e, hue: t, testId: n, onChange: r }) => {
  const o = (s) => {
    r({ h: 360 * s.left });
  };
  return /* @__PURE__ */ I("div", { className: ae(Ut["color-picker-panel__hue"], e), children: /* @__PURE__ */ I(
    wl,
    {
      onMove: o,
      "aria-label": "Hue",
      "aria-valuenow": Fe(t),
      "aria-valuemax": "360",
      "aria-valuemin": "0",
      testId: `${n}-hue`,
      children: /* @__PURE__ */ I(
        Il,
        {
          className: Ut["color-picker-panel__hue-pointer"],
          left: t / 360,
          color: ii({ h: t, s: 100, v: 100, a: 1 })
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/components/ColorPicker/Hue.tsx",
          lineNumber: 31,
          columnNumber: 9
        },
        void 0
      )
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/components/ColorPicker/Hue.tsx",
      lineNumber: 23,
      columnNumber: 7
    },
    void 0
  ) }, void 0, !1, {
    fileName: "/home/user/excel-collab/src/components/ColorPicker/Hue.tsx",
    lineNumber: 22,
    columnNumber: 5
  }, void 0);
}, wp = je.memo(bp), Ip = ({ hsva: e, testId: t, onChange: n }) => {
  const r = z((s) => {
    n({
      s: s.left * 100,
      v: 100 - s.top * 100
    });
  }, []), o = {
    backgroundColor: ii({ h: e.h, s: 100, v: 100, a: 1 })
  };
  return /* @__PURE__ */ I(
    "div",
    {
      className: Ut["color-picker-panel__saturation"],
      style: o,
      children: /* @__PURE__ */ I(
        wl,
        {
          onMove: r,
          "aria-label": "Color",
          "aria-valuetext": `Saturation ${Fe(e.s)}%, Brightness ${Fe(
            e.v
          )}%`,
          testId: `${t}-saturation`,
          children: /* @__PURE__ */ I(
            Il,
            {
              className: Ut["color-picker-panel__saturation-pointer"],
              top: 1 - e.v / 100,
              left: e.s / 100,
              color: ii(e)
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/components/ColorPicker/Saturation.tsx",
              lineNumber: 40,
              columnNumber: 9
            },
            void 0
          )
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/components/ColorPicker/Saturation.tsx",
          lineNumber: 32,
          columnNumber: 7
        },
        void 0
      )
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/components/ColorPicker/Saturation.tsx",
      lineNumber: 28,
      columnNumber: 5
    },
    void 0
  );
}, Mp = je.memo(Ip);
function yp(e, t, n) {
  const r = bl(n), [o, s] = he(
    () => e.toHsva(t)
  ), l = Xe({ color: t, hsva: o });
  Re(() => {
    if (!e.equal(t, l.current.color)) {
      const d = e.toHsva(t);
      l.current = { hsva: d, color: t }, s(d);
    }
  }, [t, e]), Re(() => {
    let d;
    !io(o, l.current.hsva) && !e.equal(
      d = e.fromHsva(o),
      l.current.color
    ) && (l.current = { hsva: o, color: d }, r(d));
  }, [o, e, r]);
  const f = z((d) => {
    s((c) => Object.assign({}, c, d));
  }, []);
  return [o, f];
}
const Sp = ({
  className: e,
  colorModel: t,
  color: n,
  onChange: r,
  testId: o
}) => {
  const [s, l] = yp(
    t,
    n || t.defaultColor,
    r
  );
  return /* @__PURE__ */ I("div", { className: ae(Ut["color-picker-panel"], e), children: [
    /* @__PURE__ */ I(Mp, { hsva: s, onChange: l, testId: o }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/components/ColorPicker/ColorPickerPanel.tsx",
      lineNumber: 31,
      columnNumber: 7
    }, void 0),
    /* @__PURE__ */ I(
      wp,
      {
        hue: s.h,
        onChange: l,
        className: Ut["color-picker-panel__last-control"],
        testId: o
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/components/ColorPicker/ColorPickerPanel.tsx",
        lineNumber: 32,
        columnNumber: 7
      },
      void 0
    )
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/components/ColorPicker/ColorPickerPanel.tsx",
    lineNumber: 30,
    columnNumber: 5
  }, void 0);
}, Rp = {
  defaultColor: "#000",
  toHsva: mp,
  fromHsva: ({ h: e, s: t, v: n }) => pp({ h: e, s: t, v: n, a: 1 }),
  equal: (e, t) => {
    const n = ti(e), r = ti(t);
    return n === r;
  }
}, Ap = ue(
  (e) => /* @__PURE__ */ I(
    Sp,
    {
      className: e.className,
      color: e.color,
      onChange: e.onChange,
      testId: e.testId,
      colorModel: Rp
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/components/ColorPicker/ColorPickerPanel.tsx",
      lineNumber: 55,
      columnNumber: 5
    },
    void 0
  )
), ur = ue((e) => {
  const {
    onChange: t,
    children: n,
    color: r,
    position: o = "bottom",
    testId: s,
    className: l
  } = e, [f, d] = he(!1), c = bn(f, () => {
    d(!1);
  }), i = z(() => {
    d(!0);
  }, []), a = z(() => {
    t("");
  }, []), u = z((h) => {
    const g = h.target?.dataset?.value;
    g && t(g);
  }, []);
  return /* @__PURE__ */ I(
    "div",
    {
      className: ae(zt["color-picker"], l, {
        [zt.top]: o === "top",
        [zt.right]: o === "right"
      }),
      ref: c,
      children: [
        /* @__PURE__ */ I(
          "div",
          {
            className: zt["color-picker-trigger"],
            onClick: i,
            "data-testid": `${s}-trigger`,
            children: n
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/components/ColorPicker/index.tsx",
            lineNumber: 53,
            columnNumber: 7
          },
          void 0
        ),
        f && /* @__PURE__ */ I("div", { className: ae(zt["color-picker-wrapper"]), children: [
          /* @__PURE__ */ I(
            "div",
            {
              className: zt["color-picker-list"],
              onClick: u,
              "data-testid": `${s}-list`,
              children: iu.map((h) => /* @__PURE__ */ I(
                "div",
                {
                  className: zt["color-picker-item"],
                  style: { backgroundColor: h },
                  "data-value": h
                },
                h,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/components/ColorPicker/index.tsx",
                  lineNumber: 69,
                  columnNumber: 17
                },
                void 0
              ))
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/components/ColorPicker/index.tsx",
              lineNumber: 62,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I("div", { children: /* @__PURE__ */ I(
            Ap,
            {
              color: r,
              onChange: t,
              testId: s
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/components/ColorPicker/index.tsx",
              lineNumber: 79,
              columnNumber: 13
            },
            void 0
          ) }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/components/ColorPicker/index.tsx",
            lineNumber: 78,
            columnNumber: 11
          }, void 0),
          /* @__PURE__ */ I("div", { children: /* @__PURE__ */ I(
            V,
            {
              type: "normal",
              className: zt.reset,
              onClick: a,
              testId: `${s}-reset`,
              children: k.t("reset")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/components/ColorPicker/index.tsx",
              lineNumber: 86,
              columnNumber: 13
            },
            void 0
          ) }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/components/ColorPicker/index.tsx",
            lineNumber: 85,
            columnNumber: 11
          }, void 0)
        ] }, void 0, !0, {
          fileName: "/home/user/excel-collab/src/components/ColorPicker/index.tsx",
          lineNumber: 61,
          columnNumber: 9
        }, void 0)
      ]
    },
    void 0,
    !0,
    {
      fileName: "/home/user/excel-collab/src/components/ColorPicker/index.tsx",
      lineNumber: 46,
      columnNumber: 5
    },
    void 0
  );
});
ur.displayName = "ColorPicker";
const Tp = {
  "app-container": "_app-container_1v3r1_1"
}, xp = "_wrap_1flh9_49", Dp = "_show_1flh9_53", Dt = {
  "formula-bar-wrapper": "_formula-bar-wrapper_1flh9_1",
  "defined-name": "_defined-name_1flh9_11",
  "defined-name-editor": "_defined-name-editor_1flh9_16",
  "formula-bar-editor-wrapper": "_formula-bar-editor-wrapper_1flh9_28",
  "formula-bar-value": "_formula-bar-value_1flh9_38",
  wrap: xp,
  show: Dp,
  "formula-editor": "_formula-editor_1flh9_58",
  "edit-cell": "_edit-cell_1flh9_73"
}, Ds = (e) => {
  let t;
  const n = /* @__PURE__ */ new Set(), r = (c, i) => {
    const a = typeof c == "function" ? c(t) : c;
    if (!Object.is(a, t)) {
      const u = t;
      t = i ?? (typeof a != "object" || a === null) ? a : Object.assign({}, t, a), n.forEach((h) => h(t, u));
    }
  }, o = () => t, f = { setState: r, getState: o, getInitialState: () => d, subscribe: (c) => (n.add(c), () => n.delete(c)) }, d = t = e(r, o, f);
  return f;
}, Op = ((e) => e ? Ds(e) : Ds), kp = (e) => e;
function Lp(e, t = kp) {
  const n = je.useSyncExternalStore(
    e.subscribe,
    je.useCallback(() => t(e.getState()), [e, t]),
    je.useCallback(() => t(e.getInitialState()), [e, t])
  );
  return je.useDebugValue(n), n;
}
const Os = (e) => {
  const t = Op(e), n = (r) => Lp(t, r);
  return Object.assign(n, t), n;
}, br = ((e) => e ? Os(e) : Os), Fp = {
  value: "",
  displayValue: "",
  row: 0,
  col: 0,
  left: ht,
  top: ht,
  width: 0,
  height: 0,
  defineName: "",
  rowCount: 1,
  colCount: 1
}, xn = br(() => ({
  ...Fp
})), Pp = {
  editorStatus: dt.NONE,
  canRedo: !1,
  canUndo: !1,
  activeUuid: "",
  currentSheetId: "",
  isFilter: !1,
  defineNames: [],
  sheetList: [],
  fontFamilies: [],
  drawings: []
}, be = br((e) => ({
  ...Pp,
  setActiveUuid(t) {
    e({ activeUuid: t });
  },
  setEditorStatus(t) {
    e({ editorStatus: t });
  },
  setDefineNames(t) {
    e({ defineNames: t });
  },
  setFontFamilies(t) {
    e({ fontFamilies: t });
  },
  setSheetList(t) {
    e({ sheetList: t });
  },
  setDrawings(t) {
    e({ drawings: t });
  }
})), pn = br(() => ({
  scrollTop: 0,
  scrollLeft: 0,
  top: 0,
  left: 0,
  row: 0,
  col: 0,
  canvasHeight: 0,
  canvasWidth: 0
})), Up = {
  smallFont: "10px",
  font: "12px",
  largeFont: "14px",
  padding: "12px",
  lineHeight: 1.5,
  mediumPadding: "8px",
  largePadding: "20px",
  borderRadius: "4px",
  borderLargeRadius: "8px",
  tinyPadding: "4px",
  scrollBarContent: "30px",
  formulaBarHeight: "25px"
}, _p = {
  lowZIndex: 2,
  middleZIndex: 3,
  highZIndex: 4,
  highestZIndex: 9
};
function Bp(e, t, n, r) {
  t = t / 100, n = n / 100;
  const o = (1 - Math.abs(2 * n - 1)) * t, s = o * (1 - Math.abs(e / 60 % 2 - 1)), l = n - o / 2;
  let f, d, c;
  if (e >= 0 && e < 60 ? (f = o, d = s, c = 0) : e >= 60 && e < 120 ? (f = s, d = o, c = 0) : e >= 120 && e < 180 ? (f = 0, d = o, c = s) : e >= 180 && e < 240 ? (f = 0, d = s, c = o) : e >= 240 && e < 300 ? (f = s, d = 0, c = o) : (f = o, d = 0, c = s), r !== 1)
    return `rgba(${f},${d},${c},${r})`;
  const i = (g) => Math.round(g * 255).toString(16).padStart(2, "0"), a = i(f + l), u = i(d + l), h = i(c + l);
  return `#${a}${u}${h}`;
}
const zp = (e, t, n) => {
  const r = e / 255, o = t / 255, s = n / 255, l = Math.min(r, o, s), f = Math.max(r, o, s), d = f - l, c = Hp(r, o, s), i = Gp(f, l), a = Vp(d, i);
  return [c, a * 100, i * 100];
}, Hp = (e, t, n) => {
  const r = Math.max(e, t, n), o = Math.min(e, t, n), s = r - o;
  let l = 0;
  if (s == 0)
    l = 0;
  else
    switch (r) {
      case e: {
        let c = (t - n) / s, i = 0;
        c < 0 && (i = 360 / 60), l = c + i;
        break;
      }
      case t: {
        let c = (n - e) / s, i = 120 / 60;
        l = c + i;
        break;
      }
      case n:
        let f = (e - t) / s, d = 240 / 60;
        l = f + d;
        break;
    }
  return l * 60;
}, Vp = (e, t) => e == 0 ? 0 : e / (1 - Math.abs(2 * t - 1)), Gp = (e, t) => (e + t) / 2, ks = ([e, t, n, r]) => {
  const [o, s, l] = zp(e, t, n);
  return Bp(o, s * 0.9, 100 - l, r);
};
function $p(e) {
  e.startsWith("#") && (e = e.slice(1));
  let t = parseInt(e.substring(0, 2), 16), n = parseInt(e.substring(2, 4), 16), r = parseInt(e.substring(4, 6), 16), o = parseInt(e.substring(6, 8) || "ff", 16);
  return [t, n, r, o / 255];
}
const qp = (e) => {
  if (e.startsWith("rgb")) {
    const n = e.replace("rgb", "").replace("a", "").replace("(", "").replace(")", "").split(",").map(Number);
    return n.length <= 3 && n.push(1), ks(n);
  } else
    return ks($p(e));
}, ji = Object.freeze({
  primaryColor: "#217346",
  buttonActiveColor: "#c6c6c6",
  selectionColor: "rgba(198,198,198,0.3)",
  backgroundColor: "#e6e6e6",
  hoverColor: "rgba(0, 0, 0, 0.04)",
  scrollbarColor: "rgba(0, 0, 0, 0.1)",
  scrollbarHoveColor: "rgba(0, 0, 0, 0.2)",
  dialogBackground: "rgba(0, 0, 0, 0.6)",
  white: "#ffffff",
  black: "#000000",
  triangleFillColor: "#dddddd",
  contentColor: "#333333",
  borderColor: "#cccccc",
  activeBorderColor: "#808080",
  errorFormulaColor: "#ff0000"
}), Ml = Object.freeze(
  Object.fromEntries(
    Object.entries(ji).map(([e, t]) => [e, qp(t)])
  )
), Kr = {
  ...Up,
  ..._p
}, si = "data-theme";
function Wp(e) {
  sessionStorage.setItem(si, e), document.documentElement.setAttribute(si, e);
}
function Zi() {
  if (typeof sessionStorage < "u") {
    const e = sessionStorage.getItem(si);
    if (e && (e === "dark" || e === "light"))
      return e;
  }
  return typeof matchMedia == "function" && matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}
function Ee(e, t) {
  return t === "dark" || Zi() === "dark" ? Ml[e] : ji[e];
}
const Yp = {
  isBold: !1,
  isItalic: !1,
  isStrike: !1,
  fontColor: Ee("contentColor"),
  fontSize: Cn,
  fontFamily: "",
  fillColor: "",
  isWrapText: !1,
  underline: Ce.NONE,
  verticalAlign: Te.TOP,
  horizontalAlign: xe.LEFT,
  numberFormat: cn,
  isMergeCell: !1,
  mergeType: ""
}, es = br(() => ({
  ...Yp
})), yl = Kc(
  void 0
);
function Se() {
  const e = Jc(yl);
  if (e === void 0)
    throw new Error("must use StateContext.Provider before");
  return e;
}
const vn = br((e) => ({
  clientId: 0,
  users: [],
  fileId: "",
  fileName: "",
  setUsers(t) {
    e({ users: t });
  },
  setClientId(t) {
    e({ clientId: t });
  },
  setFileName(t) {
    e({ fileName: t });
  },
  setFileInfo(t, n) {
    e({ fileId: t, fileName: n });
  }
}));
function Sl(e, t = !0) {
  const n = {};
  return e?.isItalic && (n.fontStyle = "italic"), e?.isBold && (n.fontWeight = "bold"), e?.fontFamily && (n.fontFamily = e?.fontFamily), e?.fontSize && !t && (n.fontSize = e?.fontSize), e?.fillColor && !t && (n.backgroundColor = e.fillColor), e?.fontColor && !t && (n.color = e?.fontColor), e?.underline && e?.isStrike ? n.textDecorationLine = "underline line-through" : e?.underline ? n.textDecorationLine = "underline" : e?.isStrike && (n.textDecorationLine = "line-through"), n;
}
function Xp(e, t, n) {
  if (t === dt.NONE)
    return;
  const r = t === dt.EDIT_FORMULA_BAR, o = Sl(n, r);
  return r ? o : {
    ...o,
    top: e.top,
    left: e.left,
    width: e.width,
    height: e.height,
    borderRadius: 0,
    border: "2px solid var(--primaryColor)"
  };
}
const Rl = 1;
function Ao(e) {
  return Math.max(Math.min(nu, e), Rl);
}
const Al = ue(({ initValue: e, style: t, testId: n, isMergeCell: r, className: o }) => {
  const { controller: s } = Se(), l = be((a) => a.setEditorStatus), f = Xe(null), [d, c] = he(Rl);
  Re(() => {
    if (r) {
      const a = e.split(mt).length;
      c(Ao(a));
    } else {
      const a = Math.ceil((f.current?.scrollHeight || 20) / 20);
      c(Ao(a));
    }
  }, [r, e]);
  const i = z(
    (a) => {
      if (a.stopPropagation(), a.key === "Enter" || a.key === "Tab") {
        let u = a.currentTarget.value;
        const { range: h, isMerged: g } = s.getActiveRange(), m = s.getCell(h);
        typeof m?.value == "string" && wi(g, m?.value) && (u = u.replaceAll(dr, mt)), s.transaction(() => {
          s.setCellValue(u, h), a.key === "Enter" ? s.setNextActiveCell("down") : s.setNextActiveCell("right");
        }), l(dt.NONE), a.currentTarget.value = "", a.currentTarget.blur();
      } else {
        const u = Math.ceil(a.currentTarget.scrollHeight / 20);
        c(Ao(u));
      }
    },
    []
  );
  return /* @__PURE__ */ I(
    "textarea",
    {
      spellCheck: !0,
      autoFocus: !0,
      ref: f,
      style: t,
      maxLength: or * 100,
      "data-testid": n,
      "data-role": ha,
      onKeyDown: i,
      className: ae(Dt["formula-editor"], o),
      defaultValue: e,
      rows: d
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/containers/FormulaBar/FormulaEditor.tsx",
      lineNumber: 145,
      columnNumber: 7
    },
    void 0
  );
});
Al.displayName = "MultipleLineEditor";
const Tl = /* @__PURE__ */ Symbol("Comlink.proxy"), Kp = /* @__PURE__ */ Symbol("Comlink.endpoint"), Jp = /* @__PURE__ */ Symbol("Comlink.releaseProxy"), To = /* @__PURE__ */ Symbol("Comlink.finalizer"), zr = /* @__PURE__ */ Symbol("Comlink.thrown"), xl = (e) => typeof e == "object" && e !== null || typeof e == "function", Qp = {
  canHandle: (e) => xl(e) && e[Tl],
  serialize(e) {
    const { port1: t, port2: n } = new MessageChannel();
    return Ol(e, t), [n, [n]];
  },
  deserialize(e) {
    return e.start(), Ll(e);
  }
}, jp = {
  canHandle: (e) => xl(e) && zr in e,
  serialize({ value: e }) {
    let t;
    return e instanceof Error ? t = {
      isError: !0,
      value: {
        message: e.message,
        name: e.name,
        stack: e.stack
      }
    } : t = { isError: !1, value: e }, [t, []];
  },
  deserialize(e) {
    throw e.isError ? Object.assign(new Error(e.value.message), e.value) : e.value;
  }
}, Dl = /* @__PURE__ */ new Map([
  ["proxy", Qp],
  ["throw", jp]
]);
function Zp(e, t) {
  for (const n of e)
    if (t === n || n === "*" || n instanceof RegExp && n.test(t))
      return !0;
  return !1;
}
function Ol(e, t = globalThis, n = ["*"]) {
  t.addEventListener("message", function r(o) {
    if (!o || !o.data)
      return;
    if (!Zp(n, o.origin)) {
      console.warn(`Invalid origin '${o.origin}' for comlink proxy`);
      return;
    }
    const { id: s, type: l, path: f } = Object.assign({ path: [] }, o.data), d = (o.data.argumentList || []).map(nn);
    let c;
    try {
      const i = f.slice(0, -1).reduce((u, h) => u[h], e), a = f.reduce((u, h) => u[h], e);
      switch (l) {
        case "GET":
          c = a;
          break;
        case "SET":
          i[f.slice(-1)[0]] = nn(o.data.value), c = !0;
          break;
        case "APPLY":
          c = a.apply(i, d);
          break;
        case "CONSTRUCT":
          {
            const u = new a(...d);
            c = ts(u);
          }
          break;
        case "ENDPOINT":
          {
            const { port1: u, port2: h } = new MessageChannel();
            Ol(e, h), c = Ul(u, [u]);
          }
          break;
        case "RELEASE":
          c = void 0;
          break;
        default:
          return;
      }
    } catch (i) {
      c = { value: i, [zr]: 0 };
    }
    Promise.resolve(c).catch((i) => ({ value: i, [zr]: 0 })).then((i) => {
      const [a, u] = jr(i);
      t.postMessage(Object.assign(Object.assign({}, a), { id: s }), u), l === "RELEASE" && (t.removeEventListener("message", r), kl(t), To in e && typeof e[To] == "function" && e[To]());
    }).catch((i) => {
      const [a, u] = jr({
        value: new TypeError("Unserializable return value"),
        [zr]: 0
      });
      t.postMessage(Object.assign(Object.assign({}, a), { id: s }), u);
    });
  }), t.start && t.start();
}
function e2(e) {
  return e.constructor.name === "MessagePort";
}
function kl(e) {
  e2(e) && e.close();
}
function Ll(e, t) {
  const n = /* @__PURE__ */ new Map();
  return e.addEventListener("message", function(o) {
    const { data: s } = o;
    if (!s || !s.id)
      return;
    const l = n.get(s.id);
    if (l)
      try {
        l(s);
      } finally {
        n.delete(s.id);
      }
  }), ai(e, n, [], t);
}
function Sr(e) {
  if (e)
    throw new Error("Proxy has been released and is not useable");
}
function Fl(e) {
  return yn(e, /* @__PURE__ */ new Map(), {
    type: "RELEASE"
  }).then(() => {
    kl(e);
  });
}
const Jr = /* @__PURE__ */ new WeakMap(), Qr = "FinalizationRegistry" in globalThis && new FinalizationRegistry((e) => {
  const t = (Jr.get(e) || 0) - 1;
  Jr.set(e, t), t === 0 && Fl(e);
});
function t2(e, t) {
  const n = (Jr.get(t) || 0) + 1;
  Jr.set(t, n), Qr && Qr.register(e, t, e);
}
function n2(e) {
  Qr && Qr.unregister(e);
}
function ai(e, t, n = [], r = function() {
}) {
  let o = !1;
  const s = new Proxy(r, {
    get(l, f) {
      if (Sr(o), f === Jp)
        return () => {
          n2(s), Fl(e), t.clear(), o = !0;
        };
      if (f === "then") {
        if (n.length === 0)
          return { then: () => s };
        const d = yn(e, t, {
          type: "GET",
          path: n.map((c) => c.toString())
        }).then(nn);
        return d.then.bind(d);
      }
      return ai(e, t, [...n, f]);
    },
    set(l, f, d) {
      Sr(o);
      const [c, i] = jr(d);
      return yn(e, t, {
        type: "SET",
        path: [...n, f].map((a) => a.toString()),
        value: c
      }, i).then(nn);
    },
    apply(l, f, d) {
      Sr(o);
      const c = n[n.length - 1];
      if (c === Kp)
        return yn(e, t, {
          type: "ENDPOINT"
        }).then(nn);
      if (c === "bind")
        return ai(e, t, n.slice(0, -1));
      const [i, a] = Ls(d);
      return yn(e, t, {
        type: "APPLY",
        path: n.map((u) => u.toString()),
        argumentList: i
      }, a).then(nn);
    },
    construct(l, f) {
      Sr(o);
      const [d, c] = Ls(f);
      return yn(e, t, {
        type: "CONSTRUCT",
        path: n.map((i) => i.toString()),
        argumentList: d
      }, c).then(nn);
    }
  });
  return t2(s, e), s;
}
function r2(e) {
  return Array.prototype.concat.apply([], e);
}
function Ls(e) {
  const t = e.map(jr);
  return [t.map((n) => n[0]), r2(t.map((n) => n[1]))];
}
const Pl = /* @__PURE__ */ new WeakMap();
function Ul(e, t) {
  return Pl.set(e, t), e;
}
function ts(e) {
  return Object.assign(e, { [Tl]: !0 });
}
function jr(e) {
  for (const [t, n] of Dl)
    if (n.canHandle(e)) {
      const [r, o] = n.serialize(e);
      return [
        {
          type: "HANDLER",
          name: t,
          value: r
        },
        o
      ];
    }
  return [
    {
      type: "RAW",
      value: e
    },
    Pl.get(e) || []
  ];
}
function nn(e) {
  switch (e.type) {
    case "HANDLER":
      return Dl.get(e.name).deserialize(e.value);
    case "RAW":
      return e.value;
  }
}
function yn(e, t, n, r) {
  return new Promise((o) => {
    const s = o2();
    t.set(s, o), e.start && e.start(), e.postMessage(Object.assign({ id: s }, n), r);
  });
}
function o2() {
  return new Array(4).fill(0).map(() => Math.floor(Math.random() * Number.MAX_SAFE_INTEGER).toString(16)).join("-");
}
class xo {
  static instance;
  controller;
  canvas;
  constructor(t, n) {
    this.controller = t, this.canvas = n;
    const r = n?.transferControlToOffscreen?.(), o = this.controller.getHooks().worker;
    if (r) {
      const s = {
        canvas: r,
        dpr: sr()
      };
      o.init(Ul(s, [s.canvas]));
    }
  }
  renderCallback = (t) => {
    const { rowMap: n, colMap: r } = t, o = Object.keys(n);
    Object.keys(r).length === 0 && o.length === 0 || this.controller.transaction(() => {
      for (const [l, f] of Object.entries(n)) {
        const d = parseInt(l, 10);
        this.controller.getRowHeight(d) !== f && this.controller.setRowHeight(d, f);
      }
      for (const [l, f] of Object.entries(r)) {
        const d = parseInt(l, 10);
        this.controller.getColWidth(d) !== f && this.controller.setColWidth(d, f);
      }
    });
  };
  async render(t) {
    const { controller: n } = this, r = n.getCurrentSheetId(), o = n.getSheetInfo(r);
    if (!o)
      return;
    const s = n.getCopyRange(), l = n.toJSON(), f = JSON.parse(JSON.stringify(l.worksheets)), d = JSON.parse(JSON.stringify(l.customHeight)), c = JSON.parse(JSON.stringify(l.customWidth)), i = l.autoFilter[r] === void 0 ? void 0 : JSON.parse(JSON.stringify(l.autoFilter[r])), a = {
      changeSet: t.changeSet,
      theme: Zi(),
      canvasSize: n.getCanvasSize(),
      headerSize: n.getHeaderSize(),
      currentSheetInfo: o,
      scroll: n.getScroll(r),
      range: n.getActiveRange().range,
      copyRange: s,
      currentMergeCells: n.getMergeCellList(r),
      customHeight: d,
      customWidth: c,
      sheetData: f,
      autoFilter: i
    };
    return this.controller.getHooks().worker.render(a, ts(this.renderCallback));
  }
  resize() {
    const { canvas: t } = this, { width: n, height: r } = this.controller.getCanvasSize();
    t.style.width = `${n}px`, t.style.height = `${r}px`;
    const o = {
      width: n,
      height: r
    };
    this.controller.getHooks().worker.resize(o);
  }
}
function Fs(e) {
  e.transaction(() => {
    $t(e), e.setNextActiveCell("right"), Zr(e);
  });
}
function Ps(e) {
  e.transaction(() => {
    $t(e), e.setNextActiveCell("down"), Zr(e);
  });
}
function _l(e, t, n) {
  const r = e.getScroll();
  let { row: o, col: s } = r;
  if (r.top !== n) {
    o = 0;
    let l = n;
    for (; l > 0; ) {
      const f = e.getRowHeight(o);
      if (f > l)
        break;
      l -= f, o++;
    }
  }
  if (r.left !== t) {
    s = 0;
    let l = t;
    for (; l > 0; ) {
      const f = e.getColWidth(s);
      if (f > l)
        break;
      l -= f, s++;
    }
  }
  return {
    row: o,
    col: s
  };
}
function i2(e) {
  const t = `div[data-testid="${fa}${e}"]`, n = document.querySelector(t);
  n && typeof n.scrollIntoView == "function" && n.scrollIntoView();
}
function Do(e, t) {
  e.transaction(() => {
    const n = t.sheetId || e.getCurrentSheetId();
    n !== e.getCurrentSheetId() && e.setCurrentSheetId(n);
    const r = e.getSheetInfo(n);
    if (!r || t.row < 0 || t.col < 0 || t.row >= r.rowCount || t.col >= r.colCount)
      return !0;
    const o = e.getScroll(n), s = e.computeCellPosition({
      row: o.row,
      col: o.col,
      colCount: 1,
      rowCount: 1,
      sheetId: n
    }), l = e.getCanvasSize(), f = e.getHeaderSize(), { top: d, left: c } = e.computeCellPosition(t), i = s.top, a = s.left, u = s.top + l.height - f.height, h = s.left + l.width - f.width;
    if (d >= i && d < u && c >= a && c <= h)
      return e.setActiveRange(t), !0;
    const g = e.computeCellPosition(
      e.getActiveRange().range
    );
    return on(e, c - g.left, d - g.top), e.setActiveRange(t), !0;
  });
}
function ns(e) {
  const r = e.getSheetViewSize(), o = e.getCanvasSize(), s = r.height - o.height + 200, l = r.width - o.width + 200, f = o.height - 30, d = o.width - 30;
  return {
    maxHeight: s,
    maxWidth: l,
    maxScrollHeight: f,
    maxScrollWidth: d
  };
}
function on(e, t, n) {
  const r = e.getScroll(), o = ns(e), { maxHeight: s, maxWidth: l, maxScrollHeight: f, maxScrollWidth: d } = o;
  let c = r.top + Math.ceil(n), i = r.left + Math.ceil(t);
  c < 0 ? c = 0 : c > s && (c = s), i < 0 ? i = 0 : i > l && (i = l);
  const a = Math.floor(c * f / s), u = Math.floor(i * d / l), { row: h, col: g } = _l(e, i, c), m = {
    left: i,
    top: c,
    scrollLeft: u,
    scrollTop: a,
    row: h,
    col: g
  };
  e.setScroll(m);
}
function Zr(e) {
  const t = e.getActiveRange().range, n = {
    row: t.row,
    col: t.col,
    colCount: 1,
    rowCount: 1,
    sheetId: ""
  }, r = e.computeCellPosition(n), o = e.getCellSize(n), s = e.getCanvasSize(), l = e.getScroll(), f = e.getSheetInfo(e.getCurrentSheetId());
  if (!f)
    return;
  const d = e.getHeaderSize(), c = 5, i = ns(e), { maxHeight: a, maxWidth: u, maxScrollHeight: h, maxScrollWidth: g } = i;
  if (r.left + o.width + c > s.width && l.col <= f.colCount - 2) {
    const m = l.left + e.getCol(l.col).len, p = Math.floor(m * g / u);
    e.setScroll({
      ...l,
      col: l.col + 1,
      left: m,
      scrollLeft: p
    });
  }
  if (r.left - d.width < s.left + c && l.col >= 1) {
    const m = l.left - e.getCol(l.col).len, p = Math.floor(m * g / u);
    e.setScroll({
      ...l,
      col: l.col - 1,
      left: m,
      scrollLeft: p
    });
  }
  if (r.top + o.height + c > s.height && l.row <= f.rowCount - 2) {
    const m = l.top + e.getRowHeight(l.row), p = Math.floor(m * h / a);
    e.setScroll({
      ...l,
      row: l.row + 1,
      top: m,
      scrollTop: p
    });
  }
  if (r.top - d.height < s.top + c && l.row >= 1) {
    const m = l.top - e.getRowHeight(l.row), p = Math.floor(m * h / a);
    e.setScroll({
      ...l,
      row: l.row - 1,
      top: m,
      scrollTop: p
    });
  }
}
function Ke() {
  const e = document.activeElement;
  return !(!e || e.getAttribute("data-role") !== ha);
}
function Bl(e) {
  const t = document.activeElement, { range: n, isMerged: r } = e.getActiveRange(), o = e.getCell(n);
  let s = t.value;
  typeof o?.value == "string" && wi(r, o?.value) && (s = s.replaceAll(dr, mt)), e.setCellValue(s, n), t.value = "", t.blur(), be.getState().setEditorStatus(dt.NONE);
}
function $t(e) {
  Ke() && Bl(e);
}
const Rt = [
  gu() ? "meta" : "ctrl"
], s2 = [
  {
    key: "Enter",
    modifierKey: [],
    handler: Ps
  },
  {
    key: "Tab",
    modifierKey: [],
    handler: Fs
  },
  {
    key: "ArrowDown",
    modifierKey: Rt,
    handler: (e) => {
      Ke() || e.transaction(() => {
        $t(e);
        const t = e.getSheetViewSize();
        return on(e, 0, t.height), !0;
      });
    }
  },
  {
    key: "ArrowUp",
    modifierKey: Rt,
    handler: (e) => {
      Ke() || e.transaction(() => {
        $t(e);
        const t = e.getSheetViewSize();
        return on(e, 0, -t.height), !0;
      });
    }
  },
  {
    key: "ArrowRight",
    modifierKey: Rt,
    handler: (e) => {
      Ke() || e.transaction(() => {
        $t(e);
        const t = e.getSheetViewSize();
        return on(e, t.width, 0), !0;
      });
    }
  },
  {
    key: "ArrowLeft",
    modifierKey: Rt,
    handler: (e) => {
      Ke() || e.transaction(() => {
        $t(e);
        const t = e.getSheetViewSize();
        return on(e, -t.width, 0), !0;
      });
    }
  },
  {
    key: "ArrowDown",
    modifierKey: [],
    handler: (e) => {
      Ke() || Ps(e);
    }
  },
  {
    key: "ArrowUp",
    modifierKey: [],
    handler: (e) => {
      Ke() || e.transaction(() => ($t(e), e.setNextActiveCell("up"), Zr(e), !0));
    }
  },
  {
    key: "ArrowRight",
    modifierKey: [],
    handler: (e) => {
      Ke() || Fs(e);
    }
  },
  {
    key: "ArrowLeft",
    modifierKey: [],
    handler: (e) => {
      Ke() || e.transaction(() => ($t(e), e.setNextActiveCell("left"), Zr(e), !0));
    }
  },
  {
    key: "b",
    modifierKey: Rt,
    handler: (e) => {
      if (Ke())
        return;
      const t = e.getCell(e.getActiveRange().range);
      e.updateCellStyle(
        { isBold: !t?.isBold },
        e.getActiveRange().range
      );
    }
  },
  {
    key: "i",
    modifierKey: Rt,
    handler: (e) => {
      if (Ke())
        return;
      const t = e.getCell(e.getActiveRange().range);
      e.updateCellStyle(
        { isItalic: !t?.isItalic },
        e.getActiveRange().range
      );
    }
  },
  {
    key: "5",
    modifierKey: Rt,
    handler: (e) => {
      if (Ke())
        return;
      const t = e.getCell(e.getActiveRange().range);
      e.updateCellStyle(
        { isStrike: !t?.isStrike },
        e.getActiveRange().range
      );
    }
  },
  {
    key: "u",
    modifierKey: Rt,
    handler: (e) => {
      if (Ke())
        return;
      const n = e.getCell(e.getActiveRange().range)?.underline;
      let r = Ce.NONE;
      n === void 0 || n === Ce.NONE ? r = Ce.SINGLE : r = Ce.NONE, e.updateCellStyle(
        { underline: r },
        e.getActiveRange().range
      );
    }
  },
  {
    key: "z",
    modifierKey: Rt,
    handler: (e) => {
      e.undo();
    }
  },
  {
    key: "y",
    modifierKey: Rt,
    handler: (e) => {
      e.redo();
    }
  },
  {
    key: "Backspace",
    modifierKey: [],
    handler: (e) => {
      e.deleteCell(e.getActiveRange().range);
    }
  },
  {
    key: "Delete",
    modifierKey: [],
    handler: (e) => {
      e.deleteCell(e.getActiveRange().range);
    }
  }
];
function qn(e) {
  const t = e?.target?.tagName || "", n = e?.currentTarget?.tagName || "";
  return ["INPUT", "TEXTAREA"].includes(t) || ["INPUT", "TEXTAREA"].includes(n);
}
function a2(e, t) {
  function n(d) {
    if (qn(d))
      return;
    Au("keydown:", d);
    const c = s2.filter((a) => a.key === d.key);
    c.sort((a, u) => u.modifierKey.length - a.modifierKey.length);
    let i;
    for (const a of c)
      if (a.modifierKey.length > 0) {
        if (a.modifierKey.some((u) => d[`${u}Key`])) {
          i = a;
          break;
        }
      } else {
        i = a;
        break;
      }
    if (i) {
      d.preventDefault(), i.handler(e);
      return;
    }
    d.metaKey || d.ctrlKey || be.getState().setEditorStatus(dt.EDIT_CELL);
  }
  const r = uu((d) => {
    (d?.target?.tagName?.toLowerCase() === "canvas" || va()) && on(e, d.deltaX, d.deltaY);
  }, 1e3 / 60);
  function o(d) {
    qn(d) || (d.preventDefault(), e.paste(d));
  }
  function s(d) {
    qn(d) || (d.preventDefault(), e.copy(d));
  }
  function l(d) {
    qn(d) || (d.preventDefault(), e.cut(d));
  }
  function f(d) {
    qn(d) || Ra().then((c) => {
      const i = e.getCopyRange();
      let a;
      if (c[Ge]) {
        const u = c[Ge];
        a = !u.floatElementUuid && u.range ? u.range : void 0, e.setFloatElementUuid(u.floatElementUuid);
      } else
        e.setFloatElementUuid("");
      io(a, i) || (e.setCopyRange(a), e.emit("renderChange", {
        changeSet: /* @__PURE__ */ new Set(["cellStyle"])
      }));
    });
  }
  return window.addEventListener("resize", t), document.body.addEventListener("keydown", n), document.body.addEventListener("wheel", r), document.body.addEventListener("paste", o), document.body.addEventListener("copy", s), document.body.addEventListener("cut", l), window.addEventListener("focus", f), () => {
    window.removeEventListener("resize", t), document.body.removeEventListener("keydown", n), document.body.removeEventListener("wheel", r), document.body.removeEventListener("paste", o), document.body.removeEventListener("copy", s), document.body.removeEventListener("cut", l), window.removeEventListener("focus", f);
  };
}
const zl = ue(
  ({ displayName: e, defineName: t }) => {
    const { controller: n } = Se(), r = Xe(null), [o, s] = he(e), l = be((a) => a.defineNames), f = We(() => l.map((a) => ({
      disabled: !1,
      value: a,
      label: a
    })), [l]);
    Re(() => {
      s(e);
    }, [e]);
    const d = z(
      (a) => {
        if (a.stopPropagation(), a.key === "Enter") {
          const u = a.currentTarget.value.trim().toLowerCase();
          if (r.current?.blur(), !u) {
            s(e);
            return;
          }
          const h = n.checkDefineName(u);
          if (h) {
            s(e), Do(n, h);
            return;
          }
          const g = Ot(u, (p) => n.getSheetList().find((N) => N.name === p)?.sheetId || ""), m = n.getSheetInfo(
            g?.sheetId || n.getCurrentSheetId()
          );
          if (!m)
            return;
          if (g && g.col < m.colCount && g.row < m.rowCount) {
            g.sheetId = g.sheetId || n.getCurrentSheetId(), s(e), Do(n, g);
            return;
          }
          Ci.test(u) && u.length <= da ? n.setDefineName(n.getActiveRange().range, u) : s(e);
        }
      },
      [e]
    ), c = z(
      (a) => {
        s(a.target.value);
      },
      []
    ), i = z((a) => {
      const u = n.checkDefineName(a);
      u && Do(n, u);
    }, []);
    return /* @__PURE__ */ I(
      Xr,
      {
        testId: "formula-bar-name",
        value: t,
        data: f,
        onChange: i,
        className: Dt["defined-name"],
        children: /* @__PURE__ */ I(
          "input",
          {
            value: o,
            ref: r,
            spellCheck: !0,
            type: "text",
            onChange: c,
            className: Dt["defined-name-editor"],
            onKeyDown: d,
            maxLength: or * 8,
            "data-testid": "formula-bar-name-input"
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/FormulaBar/DefineName.tsx",
            lineNumber: 106,
            columnNumber: 9
          },
          void 0
        )
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/FormulaBar/DefineName.tsx",
        lineNumber: 99,
        columnNumber: 7
      },
      void 0
    );
  }
);
zl.displayName = "DefineName";
const Hl = ue(() => {
  const e = xn(), t = es(), n = be((f) => f.editorStatus), r = be((f) => f.setEditorStatus), o = We(() => e.defineName || ut({
    row: e.row,
    col: e.col,
    rowCount: 1,
    colCount: 1,
    sheetId: ""
  }), [e.defineName, e.col, e.row]), s = z(() => {
    r(dt.EDIT_FORMULA_BAR);
  }, []), l = We(() => Sl(t), [t]);
  return /* @__PURE__ */ I("div", { className: Dt["formula-bar-wrapper"], "data-testid": "formula-bar", children: [
    /* @__PURE__ */ I(
      zl,
      {
        displayName: o,
        defineName: e.defineName
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/FormulaBar/index.tsx",
        lineNumber: 44,
        columnNumber: 7
      },
      void 0
    ),
    /* @__PURE__ */ I("div", { className: Dt["formula-bar-editor-wrapper"], children: [
      n !== dt.NONE && /* @__PURE__ */ I(
        Al,
        {
          initValue: e.value,
          style: Xp(e, n, t),
          testId: "formula-editor",
          isMergeCell: t.isMergeCell,
          className: n === dt.EDIT_CELL ? Dt["edit-cell"] : ""
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/containers/FormulaBar/index.tsx",
          lineNumber: 50,
          columnNumber: 11
        },
        void 0
      ),
      /* @__PURE__ */ I(
        "div",
        {
          className: ae(Dt["formula-bar-value"], {
            [Dt.show]: n !== dt.EDIT_FORMULA_BAR,
            [Dt.wrap]: t.isMergeCell && e.displayValue.includes(dr)
          }),
          style: l,
          onClick: s,
          "data-testid": "formula-editor-trigger",
          children: e.displayValue
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/containers/FormulaBar/index.tsx",
          lineNumber: 60,
          columnNumber: 9
        },
        void 0
      )
    ] }, void 0, !0, {
      fileName: "/home/user/excel-collab/src/containers/FormulaBar/index.tsx",
      lineNumber: 48,
      columnNumber: 7
    }, void 0)
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/FormulaBar/index.tsx",
    lineNumber: 43,
    columnNumber: 5
  }, void 0);
});
Hl.displayName = "FormulaBarContainer";
const l2 = "_fontFamily_196ac_35", c2 = "_bold_196ac_38", u2 = "_italic_196ac_41", f2 = "_strike_196ac_44", ye = {
  "toolbar-wrapper": "_toolbar-wrapper_196ac_1",
  "merge-cell": "_merge-cell_196ac_13",
  "merge-cell-button": "_merge-cell-button_196ac_21",
  "number-format": "_number-format_196ac_26",
  "number-format-value": "_number-format-value_196ac_30",
  fontFamily: l2,
  bold: c2,
  italic: u2,
  strike: f2,
  "wrap-text": "_wrap-text_196ac_48",
  "icon-center": "_icon-center_196ac_52"
}, Oe = [];
for (let e = 0; e < 256; ++e)
  Oe.push((e + 256).toString(16).slice(1));
function d2(e, t = 0) {
  return (Oe[e[t + 0]] + Oe[e[t + 1]] + Oe[e[t + 2]] + Oe[e[t + 3]] + "-" + Oe[e[t + 4]] + Oe[e[t + 5]] + "-" + Oe[e[t + 6]] + Oe[e[t + 7]] + "-" + Oe[e[t + 8]] + Oe[e[t + 9]] + "-" + Oe[e[t + 10]] + Oe[e[t + 11]] + Oe[e[t + 12]] + Oe[e[t + 13]] + Oe[e[t + 14]] + Oe[e[t + 15]]).toLowerCase();
}
let Oo;
const h2 = new Uint8Array(16);
function m2() {
  if (!Oo) {
    if (typeof crypto > "u" || !crypto.getRandomValues)
      throw new Error("crypto.getRandomValues() not supported. See https://github.com/uuidjs/uuid#getrandomvalues-not-supported");
    Oo = crypto.getRandomValues.bind(crypto);
  }
  return Oo(h2);
}
const g2 = typeof crypto < "u" && crypto.randomUUID && crypto.randomUUID.bind(crypto), Us = { randomUUID: g2 };
function p2(e, t, n) {
  e = e || {};
  const r = e.random ?? e.rng?.() ?? m2();
  if (r.length < 16)
    throw new Error("Random bytes length must be >= 16");
  return r[6] = r[6] & 15 | 64, r[8] = r[8] & 63 | 128, d2(r);
}
function fr(e, t, n) {
  return Us.randomUUID && !e ? Us.randomUUID() : p2(e);
}
const Vl = ue(() => {
  const { controller: e, provider: t } = Se(), n = Xe(null), r = z(
    async (o) => {
      const s = () => {
        n.current && (n.current.value = "", n.current.blur());
      }, l = o.target.files?.[0];
      if (!l)
        return;
      const f = 25 * 1024 * 1024;
      if (l.size > f) {
        Pe.error("max image file size 25MB"), s();
        return;
      }
      let d = l.name;
      const c = l.type.slice(6);
      d = d.slice(0, -(c.length + 1));
      const i = await Mi(l, !0);
      if (!i) {
        s();
        return;
      }
      const a = await yi(i);
      let u = i;
      if (t?.uploadFile && (u = await t.uploadFile(
        e.getHooks().doc.guid,
        l,
        i
      )), !u) {
        Pe.warning("choose image file"), s();
        return;
      }
      const h = e.getActiveRange().range;
      e.addDrawing({
        width: a.width,
        height: a.height,
        originHeight: a.height,
        originWidth: a.width,
        title: d,
        type: "floating-picture",
        uuid: fr(),
        imageSrc: u,
        sheetId: h.sheetId,
        fromRow: h.row,
        fromCol: h.col,
        marginX: 0,
        marginY: 0
      }), s();
    },
    [t]
  );
  return /* @__PURE__ */ I(V, { testId: "toolbar-floating-picture", title: "Floating Picture", children: [
    /* @__PURE__ */ I(
      "input",
      {
        type: "file",
        hidden: !0,
        onChange: r,
        accept: "image/*",
        ref: n,
        id: "upload_float_image",
        "data-testid": "toolbar-floating-picture-input"
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/FloatElement/Toolbar.tsx",
        lineNumber: 76,
        columnNumber: 7
      },
      void 0
    ),
    /* @__PURE__ */ I("label", { htmlFor: "upload_float_image", children: k.t("floating-picture") }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/FloatElement/Toolbar.tsx",
      lineNumber: 85,
      columnNumber: 7
    }, void 0)
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/FloatElement/Toolbar.tsx",
    lineNumber: 75,
    columnNumber: 5
  }, void 0);
});
Vl.displayName = "InsertFloatingPicture";
const Gl = ue(() => {
  const { controller: e } = Se(), t = z(async () => {
    const n = e.getActiveRange().range;
    e.addDrawing({
      width: 400,
      height: 300,
      originHeight: 300,
      originWidth: 400,
      title: k.t("chart-title"),
      type: "chart",
      uuid: fr(),
      sheetId: n.sheetId,
      fromRow: n.row,
      fromCol: n.col,
      chartRange: un(n),
      chartType: "line",
      marginX: 0,
      marginY: 0
    });
  }, []);
  return /* @__PURE__ */ I(V, { testId: "toolbar-chart", onClick: t, title: "Chart", children: k.t("chart") }, void 0, !1, {
    fileName: "/home/user/excel-collab/src/containers/FloatElement/Toolbar.tsx",
    lineNumber: 114,
    columnNumber: 5
  }, void 0);
});
Gl.displayName = "InsertChart";
const v2 = "_container_11gw9_1", N2 = "_main_11gw9_14", C2 = "_menu_11gw9_19", E2 = "_portal_11gw9_24", Rr = {
  container: v2,
  main: N2,
  menu: C2,
  portal: E2
}, $l = ue(() => {
  const { controller: e } = Se(), [t, n] = he(""), [r, o] = he("thin"), [s, l] = he("all-borders"), f = Xe({ color: t, borderType: r, type: s }), d = () => ({
    color: f.current.color,
    type: f.current.borderType
  }), c = () => {
    l("all-borders");
    const v = d();
    e.updateCellStyle(
      {
        borderLeft: v,
        borderRight: v,
        borderTop: v,
        borderBottom: v
      },
      e.getActiveRange().range
    );
  }, i = (v) => {
    f.current.color = v, n(v), c();
  }, a = (v) => {
    f.current.borderType = v, o(v), c();
  }, u = () => {
    l("no-border"), e.updateCellStyle(
      {
        borderLeft: void 0,
        borderRight: void 0,
        borderTop: void 0,
        borderBottom: void 0
      },
      e.getActiveRange().range
    );
  }, h = () => {
    l("bottom-border");
    const v = e.getActiveRange().range, C = d(), { row: M, col: R, colCount: A, rowCount: x } = v;
    e.updateCellStyle(
      {
        borderBottom: C
      },
      {
        row: M + x - 1,
        rowCount: 1,
        colCount: A,
        col: Wt(v) ? 0 : R,
        sheetId: ""
      }
    );
  }, g = () => {
    l("top-border");
    const v = e.getActiveRange().range, C = d(), { row: M, col: R, colCount: A } = v, x = {
      row: M,
      rowCount: 1,
      colCount: A,
      col: Wt(v) ? 0 : R,
      sheetId: ""
    };
    e.updateCellStyle(
      {
        borderTop: C
      },
      x
    );
  }, m = () => {
    l("left-border");
    const v = e.getActiveRange().range, C = d(), { row: M, col: R, rowCount: A } = v;
    e.updateCellStyle(
      {
        borderLeft: C
      },
      {
        row: Yt(v) ? 0 : M,
        rowCount: A,
        colCount: 1,
        col: R,
        sheetId: ""
      }
    );
  }, p = () => {
    l("right-border");
    const v = e.getActiveRange().range, C = d(), { row: M, col: R, rowCount: A, colCount: x } = v;
    e.updateCellStyle(
      {
        borderRight: C
      },
      {
        row: Yt(v) ? 0 : M,
        rowCount: A,
        colCount: 1,
        col: R + x - 1,
        sheetId: ""
      }
    );
  }, E = () => {
    g(), p(), h(), m(), l("outside-borders");
  }, b = () => {
    const v = f.current.borderType;
    f.current.borderType = "medium", E(), f.current.borderType = v, l("thick-box-border");
  }, N = () => {
    ({
      "all-borders": c,
      "no-border": u,
      "bottom-border": h,
      "top-border": g,
      "left-border": m,
      "right-border": p,
      "thick-box-border": b,
      "outside-borders": E
    })[s]();
  };
  return /* @__PURE__ */ I("div", { className: Rr.container, children: [
    /* @__PURE__ */ I(
      V,
      {
        onClick: N,
        type: "plain",
        className: Rr.main,
        testId: "toolbar-border-shortcut",
        title: k.t(s),
        children: k.t(s)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
        lineNumber: 178,
        columnNumber: 7
      },
      void 0
    ),
    /* @__PURE__ */ I(
      Qi,
      {
        className: Rr.menu,
        label: /* @__PURE__ */ I(Ae, { name: "down" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
          lineNumber: 189,
          columnNumber: 16
        }, void 0),
        isPlain: !0,
        testId: "toolbar-border",
        position: "bottom",
        size: "small",
        portalClassName: Rr.portal,
        children: [
          /* @__PURE__ */ I(Le, { onClick: u, testId: "toolbar-no-border", children: k.t("no-border") }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
            lineNumber: 196,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(Le, { onClick: c, testId: "toolbar-all-borders", children: k.t("all-borders") }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
            lineNumber: 199,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(
            Le,
            {
              onClick: E,
              testId: "toolbar-outside-borders",
              children: k.t("outside-borders")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
              lineNumber: 202,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            Le,
            {
              onClick: b,
              testId: "toolbar-thick-box-border",
              children: k.t("thick-box-border")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
              lineNumber: 208,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(Le, { onClick: h, testId: "toolbar-bottom-border", children: k.t("bottom-border") }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
            lineNumber: 214,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(Le, { onClick: g, testId: "toolbar-top-border", children: k.t("top-border") }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
            lineNumber: 217,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(Le, { onClick: m, testId: "toolbar-left-border", children: k.t("left-border") }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
            lineNumber: 220,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(Le, { onClick: p, testId: "toolbar-right-border", children: k.t("right-border") }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
            lineNumber: 223,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(Le, { children: /* @__PURE__ */ I(
            ur,
            {
              color: t,
              onChange: i,
              position: "right",
              testId: "toolbar-border-color",
              children: /* @__PURE__ */ I("span", { style: { color: t }, children: [
                k.t("line-color"),
                " >"
              ] }, void 0, !0, {
                fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
                lineNumber: 233,
                columnNumber: 13
              }, void 0)
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
              lineNumber: 227,
              columnNumber: 11
            },
            void 0
          ) }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
            lineNumber: 226,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(Wg, { label: `${k.t("line-style")} >`, testId: "toolbar-border-style", children: Object.keys(oo).map((v) => /* @__PURE__ */ I(
            Le,
            {
              onClick: () => a(v),
              testId: `toolbar-border-style-${v}`,
              active: r === v,
              children: v
            },
            v,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
              lineNumber: 238,
              columnNumber: 13
            },
            void 0
          )) }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
            lineNumber: 236,
            columnNumber: 9
          }, void 0)
        ]
      },
      void 0,
      !0,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
        lineNumber: 187,
        columnNumber: 7
      },
      void 0
    )
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/ToolBar/Border.tsx",
    lineNumber: 177,
    columnNumber: 5
  }, void 0);
});
$l.displayName = "BorderToolBar";
function _s(e) {
  if (typeof document > "u")
    return { offsetHeight: 0, offsetWidth: 0 };
  const t = document.createElement("span");
  t.style.fontSize = "72px", t.innerHTML = "mmmmmmmmmmlli", t.style.fontFamily = e, document.body.appendChild(t);
  const { offsetWidth: n, offsetHeight: r } = t;
  return document.body.removeChild(t), {
    offsetHeight: r,
    offsetWidth: n
  };
}
function b2() {
  const e = [ca, "serif"], t = {}, n = {};
  for (const o of e) {
    const { offsetWidth: s, offsetHeight: l } = _s(o);
    t[o] = s, n[o] = l;
  }
  function r(o) {
    for (const s of e) {
      const { offsetWidth: l, offsetHeight: f } = _s(o + "," + s);
      if (l !== t[s] || f !== n[s])
        return !0;
    }
    return !1;
  }
  return r;
}
const ql = b2();
function w2(e = ql, t = Uu) {
  const n = localStorage.getItem(Aa);
  if (n) {
    const o = JSON.parse(n);
    if (o.length > 0)
      return o.map((s) => ({ value: s, label: s, disabled: !1 }));
  }
  const r = [];
  for (const o of t)
    e(o) && r.push({
      label: o,
      value: o,
      disabled: !1
    });
  return typeof window.queryLocalFonts == "function" && r.push({
    value: ei,
    label: k.t("get-all-installed-fonts"),
    disabled: !1
  }), r;
}
const I2 = [
  {
    value: Ce.NONE,
    label: k.t("none"),
    disabled: !1
  },
  {
    value: Ce.SINGLE,
    label: k.t("single-underline"),
    disabled: !1
  },
  {
    value: Ce.DOUBLE,
    label: k.t("double-underline"),
    disabled: !1
  }
], M2 = [
  {
    value: Ct.MERGE_CENTER,
    label: k.t("merge-and-center"),
    disabled: !1
  },
  {
    value: Ct.MERGE_CELL,
    label: k.t("merge-cells"),
    disabled: !1
  },
  {
    value: Ct.MERGE_CONTENT,
    label: k.t("merge-content"),
    disabled: !1
  }
], Wn = [
  {
    value: xt(0),
    label: k.t("general"),
    disabled: !1
  },
  {
    value: xt(2),
    label: k.t("number"),
    disabled: !1
  },
  {
    value: xt(8),
    label: k.t("currency"),
    disabled: !1
  },
  {
    value: xt(44),
    label: k.t("accounting"),
    disabled: !1
  },
  {
    value: k.t("short-date-format"),
    label: k.t("short-date"),
    disabled: !1
  },
  {
    value: k.t("long-date-format"),
    label: k.t("long-date"),
    disabled: !1
  },
  {
    value: k.t("time-format"),
    label: k.t("time"),
    disabled: !1
  },
  {
    value: xt(10),
    label: k.t("percentage"),
    disabled: !1
  },
  {
    value: xt(12),
    label: k.t("fraction"),
    disabled: !1
  },
  {
    value: xt(11),
    label: k.t("scientific"),
    disabled: !1
  },
  {
    value: xt(49),
    label: k.t("text"),
    disabled: !1
  }
  // {
  //   value: '',
  //   label: ('more-number-formats'),
  //   disabled: false,
  // },
], Wl = ue(({ children: e }) => {
  const { controller: t } = Se(), n = be(($) => $.canRedo), r = be(($) => $.canUndo), o = be(($) => $.isFilter), s = be(($) => $.fontFamilies), l = be(($) => $.setFontFamilies), f = es(), d = be(($) => $.fontFamilies), c = We(() => ({ color: f.fillColor }), [f.fillColor]), i = We(() => ({ color: f.fontColor }), [f.fontColor]), [a, u] = We(() => {
    let $ = Wn[0];
    if (f.numberFormat) {
      const j = Wn.find(
        (pe) => pe.value === f.numberFormat
      );
      j ? $ = j : $ = Wn[Wn.length - 1];
    }
    return [$.label, String($.value)];
  }, [f.numberFormat]), h = z(
    ($) => ({
      fontFamily: String($)
    }),
    []
  ), g = z(($) => {
    String($) === ei && typeof window.queryLocalFonts == "function" ? window.queryLocalFonts().then((j) => {
      let pe = j.map((De) => De.fullName);
      pe = Array.from(new Set(pe)).filter(
        (De) => ql(De)
      ), pe.sort((De, vt) => De.localeCompare(vt));
      const St = pe.map((De) => ({
        label: De,
        value: De,
        disabled: !1
      }));
      pe.length > 0 ? (l(St), localStorage.setItem(Aa, JSON.stringify(pe))) : l(
        s.filter((De) => De.value !== ei)
      );
    }) : t.updateCellStyle(
      { fontFamily: String($) },
      t.getActiveRange().range
    );
  }, []), m = z(() => {
    t.undo();
  }, []), p = z(() => {
    t.redo();
  }, []), E = z(() => {
    t.copy();
  }, []), b = z(() => {
    t.cut();
  }, []), N = z(() => {
    t.paste();
  }, []), v = z(($) => {
    t.updateCellStyle(
      { fontSize: Number($) },
      t.getActiveRange().range
    );
  }, []), C = z(() => {
    t.updateCellStyle(
      { isBold: !f.isBold },
      t.getActiveRange().range
    );
  }, [f.isBold]), M = z(() => {
    t.updateCellStyle(
      { isItalic: !f.isItalic },
      t.getActiveRange().range
    );
  }, [f.isItalic]), R = z(() => {
    t.updateCellStyle(
      { isStrike: !f.isStrike },
      t.getActiveRange().range
    );
  }, [f.isStrike]), A = z(($) => {
    const j = Number($);
    let pe = Ce.NONE;
    j === Ce.SINGLE ? pe = Ce.SINGLE : j === Ce.DOUBLE && (pe = Ce.DOUBLE), t.updateCellStyle(
      { underline: pe },
      t.getActiveRange().range
    );
  }, []), x = z(($) => {
    t.updateCellStyle(
      { fillColor: $ },
      t.getActiveRange().range
    );
  }, []), L = z(($) => {
    t.updateCellStyle(
      { fontColor: $ },
      t.getActiveRange().range
    );
  }, []), D = z(() => {
    t.updateCellStyle(
      { isWrapText: !f.isWrapText },
      t.getActiveRange().range
    );
  }, [f.isWrapText]), W = z(() => {
    const { range: $, isMerged: j } = t.getActiveRange();
    j ? t.deleteMergeCell($) : t.addMergeCell($);
  }, []), J = z(($) => {
    if (!$)
      return;
    const { range: j, isMerged: pe } = t.getActiveRange();
    pe ? t.deleteMergeCell(j) : t.addMergeCell(j, Number($));
  }, []), ne = z(($) => {
    $ && t.updateCellStyle(
      { numberFormat: $ },
      t.getActiveRange().range
    );
  }, []), K = z(() => {
    t.updateCellStyle(
      { horizontalAlign: xe.LEFT },
      t.getActiveRange().range
    );
  }, []), q = z(() => {
    t.updateCellStyle(
      { horizontalAlign: xe.CENTER },
      t.getActiveRange().range
    );
  }, []), H = z(() => {
    t.updateCellStyle(
      { horizontalAlign: xe.RIGHT },
      t.getActiveRange().range
    );
  }, []), Y = z(() => {
    t.updateCellStyle(
      { verticalAlign: Te.TOP },
      t.getActiveRange().range
    );
  }, []), te = z(() => {
    t.updateCellStyle(
      { verticalAlign: Te.MIDDLE },
      t.getActiveRange().range
    );
  }, []), X = z(() => {
    t.updateCellStyle(
      { verticalAlign: Te.BOTTOM },
      t.getActiveRange().range
    );
  }, []), le = z(() => {
    t.getFilter() ? t.deleteFilter() : t.addFilter(t.getActiveRange().range);
  }, []);
  return /* @__PURE__ */ I("div", { className: ye["toolbar-wrapper"], "data-testid": "toolbar", children: [
    /* @__PURE__ */ I(
      V,
      {
        disabled: !r,
        onClick: m,
        testId: "toolbar-undo",
        title: "Undo",
        className: ye["icon-center"],
        children: /* @__PURE__ */ I(Ae, { name: "undo" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 257,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 250,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        disabled: !n,
        onClick: p,
        testId: "toolbar-redo",
        title: "Redo",
        className: ye["icon-center"],
        children: /* @__PURE__ */ I(Ae, { name: "redo" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 266,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 259,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(V, { onClick: E, testId: "toolbar-copy", title: "Copy", children: k.t("copy") }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
      lineNumber: 268,
      columnNumber: 9
    }, void 0),
    /* @__PURE__ */ I(V, { onClick: b, testId: "toolbar-cut", title: "Cut", children: k.t("cut") }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
      lineNumber: 271,
      columnNumber: 9
    }, void 0),
    /* @__PURE__ */ I(V, { onClick: N, testId: "toolbar-paste", title: "Paste", children: k.t("paste") }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
      lineNumber: 274,
      columnNumber: 9
    }, void 0),
    /* @__PURE__ */ I(
      ln,
      {
        data: d,
        value: f.fontFamily,
        getItemStyle: h,
        onChange: g,
        testId: "toolbar-font-family",
        className: ye.fontFamily
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 278,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      ln,
      {
        data: Nu,
        value: f.fontSize,
        onChange: v,
        testId: "toolbar-font-size"
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 286,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.isBold,
        onClick: C,
        testId: "toolbar-bold",
        title: "Bold",
        children: /* @__PURE__ */ I("span", { className: ye.bold, children: "B" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 298,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 292,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.isItalic,
        onClick: M,
        testId: "toolbar-italic",
        title: "Italic",
        children: /* @__PURE__ */ I("span", { className: ye.italic, children: "I" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 306,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 300,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.isStrike,
        onClick: R,
        testId: "toolbar-strike",
        title: "Strike",
        children: /* @__PURE__ */ I("span", { className: ye.strike, children: "A" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 314,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 308,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      ln,
      {
        data: I2,
        value: f.underline,
        title: "Underline",
        onChange: A,
        testId: "toolbar-underline"
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 316,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I($l, {}, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
      lineNumber: 323,
      columnNumber: 9
    }, void 0),
    /* @__PURE__ */ I(
      ur,
      {
        color: f.fillColor,
        onChange: x,
        testId: "toolbar-fill-color",
        children: /* @__PURE__ */ I(
          V,
          {
            style: c,
            testId: "toolbar-fill-color",
            className: ye["icon-center"],
            title: "Fill Color",
            children: /* @__PURE__ */ I(El, {}, void 0, !1, {
              fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
              lineNumber: 336,
              columnNumber: 13
            }, void 0)
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
            lineNumber: 330,
            columnNumber: 11
          },
          void 0
        )
      },
      "fill-color",
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 324,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      ur,
      {
        color: f.fontColor,
        onChange: L,
        testId: "toolbar-font-color",
        children: /* @__PURE__ */ I(
          V,
          {
            style: i,
            testId: "toolbar-font-color",
            className: ye["icon-center"],
            title: "Font Color",
            children: /* @__PURE__ */ I(Ae, { name: "fontColor" }, void 0, !1, {
              fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
              lineNumber: 352,
              columnNumber: 13
            }, void 0)
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
            lineNumber: 346,
            columnNumber: 11
          },
          void 0
        )
      },
      "font-color",
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 340,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.verticalAlign === Te.TOP,
        onClick: Y,
        testId: "toolbar-vertical-top",
        className: ye["icon-center"],
        title: "Top Align",
        children: /* @__PURE__ */ I(Ae, { name: "verticalTop" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 362,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 355,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.verticalAlign === Te.MIDDLE,
        onClick: te,
        testId: "toolbar-vertical-middle",
        className: ye["icon-center"],
        title: "Middle Align",
        children: /* @__PURE__ */ I(Ae, { name: "verticalMiddle" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 371,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 364,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.verticalAlign === Te.BOTTOM,
        onClick: X,
        testId: "toolbar-vertical-bottom",
        className: ye["icon-center"],
        title: "Bottom Align",
        children: /* @__PURE__ */ I(Ae, { name: "verticalBottom" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 380,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 373,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.horizontalAlign === xe.LEFT,
        onClick: K,
        testId: "toolbar-horizontal-left",
        className: ye["icon-center"],
        title: "Align Text Left",
        children: /* @__PURE__ */ I(Ae, { name: "horizontalLeft" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 389,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 382,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.horizontalAlign === xe.CENTER,
        onClick: q,
        testId: "toolbar-horizontal-center",
        className: ye["icon-center"],
        title: "Align Text Center",
        children: /* @__PURE__ */ I(Ae, { name: "horizontalCenter" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 398,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 391,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.horizontalAlign === xe.RIGHT,
        onClick: H,
        testId: "toolbar-horizontal-right",
        className: ye["icon-center"],
        title: "Align Text Right",
        children: /* @__PURE__ */ I(Ae, { name: "horizontalRight" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
          lineNumber: 407,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 400,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: f.isWrapText,
        onClick: D,
        testId: "toolbar-wrap-text",
        className: ye["wrap-text"],
        title: "Wrap Text",
        children: k.t("wrap-text")
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 409,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      Xr,
      {
        data: M2,
        value: f.mergeType,
        onChange: J,
        className: ye["merge-cell"],
        testId: "toolbar-merge-cell-select",
        children: /* @__PURE__ */ I(
          V,
          {
            active: f.isMergeCell,
            onClick: W,
            testId: "toolbar-merge-cell",
            className: ye["merge-cell-button"],
            type: "plain",
            title: "Merge And Center",
            children: k.t("merge-and-center")
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
            lineNumber: 425,
            columnNumber: 11
          },
          void 0
        )
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 418,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      Xr,
      {
        data: Wn,
        value: u,
        onChange: ne,
        className: ye["number-format"],
        testId: "toolbar-number-format",
        children: /* @__PURE__ */ I(
          "div",
          {
            className: ye["number-format-value"],
            "data-testid": "toolbar-number-format-value",
            children: a
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
            lineNumber: 443,
            columnNumber: 11
          },
          void 0
        )
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 436,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(
      V,
      {
        active: o,
        onClick: le,
        testId: "toolbar-filter",
        className: ye["wrap-text"],
        title: "Filter",
        children: k.t("filter")
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
        lineNumber: 450,
        columnNumber: 9
      },
      void 0
    ),
    /* @__PURE__ */ I(Vl, {}, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
      lineNumber: 459,
      columnNumber: 9
    }, void 0),
    /* @__PURE__ */ I(Gl, {}, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
      lineNumber: 460,
      columnNumber: 9
    }, void 0),
    e
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/ToolBar/index.tsx",
    lineNumber: 249,
    columnNumber: 7
  }, void 0);
});
Wl.displayName = "ToolbarContainer";
const y2 = "_active_1sd55_81", S2 = "_user_1sd55_99", R2 = "_userContent_1sd55_103", $e = {
  "canvas-container": "_canvas-container_1sd55_1",
  "vertical-scroll-bar": "_vertical-scroll-bar_1sd55_9",
  "vertical-scroll-bar-content": "_vertical-scroll-bar-content_1sd55_22",
  "horizontal-scroll-bar": "_horizontal-scroll-bar_1sd55_30",
  "horizontal-scroll-bar-content": "_horizontal-scroll-bar-content_1sd55_43",
  "context-menu": "_context-menu_1sd55_57",
  "bottom-bar": "_bottom-bar_1sd55_70",
  active: y2,
  "bottom-bar-text": "_bottom-bar-text_1sd55_85",
  "bottom-bar-input": "_bottom-bar-input_1sd55_89",
  "add-button": "_add-button_1sd55_93",
  user: S2,
  userContent: R2
};
function Bs(e, t, n) {
  const r = e.getScroll(), o = ns(e), { maxHeight: s, maxScrollHeight: l, maxScrollWidth: f, maxWidth: d } = o;
  let c = r.scrollTop + Math.ceil(n), i = r.scrollLeft + Math.ceil(t);
  c < 0 ? c = 0 : c > l && (c = l), i < 0 ? i = 0 : i > f && (i = f);
  const a = Math.floor(s * c / l), u = Math.floor(d * i / f), { row: h, col: g } = _l(e, u, a);
  e.setScroll({ row: h, col: g, top: a, left: u, scrollLeft: i, scrollTop: c });
}
const zs = {
  prevPageX: 0,
  prevPageY: 0,
  scrollStatus: tn.NONE
}, Yl = ue(() => {
  const { controller: e } = Se(), t = Xe({ ...zs }), n = pn((d) => d.scrollLeft), r = pn((d) => d.scrollTop), [o] = We(() => {
    function d() {
      t.current = { ...zs }, i(!1);
    }
    function c(a) {
      if (a.buttons <= 0) {
        i(!1);
        return;
      }
      t.current.scrollStatus !== tn.NONE && (t.current.scrollStatus === tn.VERTICAL ? (Bs(e, 0, a.clientY - t.current.prevPageY), t.current.prevPageY = a.clientY) : t.current.scrollStatus === tn.HORIZONTAL && (Bs(e, a.clientX - t.current.prevPageX, 0), t.current.prevPageX = a.clientX));
    }
    function i(a) {
      const u = a ? document.addEventListener : document.removeEventListener;
      u("pointermove", c), u("pointerup", d);
    }
    return [i];
  }, []);
  Re(() => o, [o]);
  const s = z(
    (d, c) => {
      d.buttons <= 0 || (o(!0), t.current.prevPageX = d.clientX, t.current.prevPageY = d.clientY, t.current.scrollStatus = c);
    },
    [o]
  ), l = z(
    (d) => {
      s(d, tn.VERTICAL);
    },
    [s]
  ), f = z(
    (d) => {
      s(d, tn.HORIZONTAL);
    },
    [o]
  );
  return /* @__PURE__ */ I(rn, { children: [
    /* @__PURE__ */ I(
      "div",
      {
        className: $e["vertical-scroll-bar"],
        "data-testid": "vertical-scroll-bar",
        onPointerDown: l,
        children: /* @__PURE__ */ I(
          "div",
          {
            className: $e["vertical-scroll-bar-content"],
            style: {
              transform: `translateY(${r}px)`
            },
            "data-testid": "vertical-scroll-bar-content"
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/canvas/ScrollBar.tsx",
            lineNumber: 116,
            columnNumber: 9
          },
          void 0
        )
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/canvas/ScrollBar.tsx",
        lineNumber: 111,
        columnNumber: 7
      },
      void 0
    ),
    /* @__PURE__ */ I(
      "div",
      {
        className: $e["horizontal-scroll-bar"],
        "data-testid": "horizontal-scroll-bar",
        onPointerDown: f,
        children: /* @__PURE__ */ I(
          "div",
          {
            className: $e["horizontal-scroll-bar-content"],
            style: {
              transform: `translateX(${n}px)`
            },
            "data-testid": "horizontal-scroll-bar-content"
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/canvas/ScrollBar.tsx",
            lineNumber: 129,
            columnNumber: 9
          },
          void 0
        )
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/canvas/ScrollBar.tsx",
        lineNumber: 124,
        columnNumber: 7
      },
      void 0
    )
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/canvas/ScrollBar.tsx",
    lineNumber: 110,
    columnNumber: 5
  }, void 0);
});
Yl.displayName = "ScrollBar";
const Hs = 110, ko = 20;
function A2(e, t, n) {
  const r = n.getCanvasSize(), o = n.getHeaderSize();
  let s = 3, l = ko * 3;
  const f = e - r.top, d = t - r.left;
  f < o.height && d < o.width ? s = 2 : f < o.height ? (s = 0, l = ko * 6) : d < o.width && (s = 1, l = ko * 6);
  let c = e, i = t;
  const a = 18, u = r.height + r.top;
  return c + l > u && (c = u - l - a), i + Hs > r.width && (i = r.width - Hs - a), {
    style: {
      top: c,
      left: i
    },
    position: s
  };
}
const Lo = 1e4, Xl = ue((e) => {
  const { controller: t } = Se(), { top: n, left: r, hideContextMenu: o } = e, s = xn((h) => h.row), l = xn((h) => h.col), f = xn((h) => h.colCount), d = xn((h) => h.rowCount), c = bn(!0, o), { style: i, position: a } = We(() => A2(n, r, t), [n, r]), u = (h) => {
    let g = h ? t.getRow(s).len : t.getCol(l).len;
    const m = (p) => {
      const E = parseInt(p.target.value, 10);
      isNaN(E) || (E < 0 ? g = 0 : g > Lo ? g = Lo : g = E), p.stopPropagation();
    };
    Zn({
      visible: !0,
      title: h ? k.t("row-height") : k.t("column-width"),
      testId: "context-menu-width-height-dialog",
      children: /* @__PURE__ */ I(
        "input",
        {
          type: "number",
          min: 0,
          max: Lo,
          style: { width: "200px" },
          defaultValue: g,
          onChange: m,
          "data-testid": "context-menu-width-height-dialog-input"
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
          lineNumber: 97,
          columnNumber: 9
        },
        void 0
      ),
      onOk: () => {
        if (g < 0)
          return Pe.error(k.t("greater-than-zero"));
        h ? t.transaction(() => {
          for (let p = 0; p < d; p++)
            t.setRowHeight(s + p, g);
        }) : t.transaction(() => {
          for (let p = 0; p < f; p++)
            t.setColWidth(l + p, g);
        }), o();
      },
      onCancel: () => {
        o();
      }
    });
  };
  return /* @__PURE__ */ I(
    "div",
    {
      className: $e["context-menu"],
      "data-testid": "context-menu",
      style: i,
      ref: c,
      children: [
        /* @__PURE__ */ I(
          V,
          {
            onClick: () => {
              o(), t.setFloatElementUuid(""), t.copy();
            },
            testId: "context-menu-copy",
            children: k.t("copy")
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
            lineNumber: 138,
            columnNumber: 7
          },
          void 0
        ),
        /* @__PURE__ */ I(
          V,
          {
            onClick: () => {
              o(), t.setFloatElementUuid(""), t.cut();
            },
            testId: "context-menu-cut",
            children: k.t("cut")
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
            lineNumber: 148,
            columnNumber: 7
          },
          void 0
        ),
        /* @__PURE__ */ I(
          V,
          {
            testId: "context-menu-paste",
            onClick: () => {
              o(), t.paste();
            },
            children: k.t("paste")
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
            lineNumber: 158,
            columnNumber: 7
          },
          void 0
        ),
        (a === 1 || a === 3) && /* @__PURE__ */ I(rn, { children: [
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-insert-row-above",
              onClick: () => {
                o(), t.addRow(s, d, !0);
              },
              children: k.t("insert-row-above")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 170,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-insert-row-below",
              onClick: () => {
                o(), t.addRow(s, d);
              },
              children: k.t("insert-row-below")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 179,
              columnNumber: 11
            },
            void 0
          )
        ] }, void 0, !0, {
          fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
          lineNumber: 169,
          columnNumber: 9
        }, void 0),
        (a === 0 || a === 3) && /* @__PURE__ */ I(rn, { children: [
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-insert-column-left",
              onClick: () => {
                o(), t.addCol(l, f);
              },
              children: k.t("insert-column-left")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 193,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-insert-column-right",
              onClick: () => {
                o(), t.addCol(l, f, !0);
              },
              children: k.t("insert-column-right")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 202,
              columnNumber: 11
            },
            void 0
          )
        ] }, void 0, !0, {
          fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
          lineNumber: 192,
          columnNumber: 9
        }, void 0),
        a === 2 && /* @__PURE__ */ I(
          V,
          {
            testId: "context-menu-delete",
            onClick: () => {
              o(), t.deleteAll(t.getCurrentSheetId());
            },
            children: k.t("delete")
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
            lineNumber: 214,
            columnNumber: 9
          },
          void 0
        ),
        a === 0 && /* @__PURE__ */ I(rn, { children: [
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-delete-column",
              onClick: () => {
                o(), t.deleteCol(l, f);
              },
              children: k.t("delete-columns")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 226,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-hide-column",
              onClick: () => {
                o(), t.hideCol(l, f);
              },
              children: k.t("hide-columns")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 235,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-unhide-column",
              onClick: () => {
                o(), t.unhideCol(l, f);
              },
              children: k.t("unhide-columns")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 244,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-column-width",
              onClick: () => {
                u(!1);
              },
              children: k.t("column-width")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 253,
              columnNumber: 11
            },
            void 0
          )
        ] }, void 0, !0, {
          fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
          lineNumber: 225,
          columnNumber: 9
        }, void 0),
        a === 1 && /* @__PURE__ */ I(rn, { children: [
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-delete-row",
              onClick: () => {
                o(), t.deleteRow(s, d);
              },
              children: k.t("delete-rows")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 265,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-hide-row",
              onClick: () => {
                o(), t.hideRow(s, d);
              },
              children: k.t("hide-rows")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 274,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-unhide-row",
              onClick: () => {
                o(), t.unhideRow(s, d);
              },
              children: k.t("unhide-rows")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 283,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "context-menu-row-height",
              onClick: () => {
                u(!0);
              },
              children: k.t("row-height")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
              lineNumber: 292,
              columnNumber: 11
            },
            void 0
          )
        ] }, void 0, !0, {
          fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
          lineNumber: 264,
          columnNumber: 9
        }, void 0)
      ]
    },
    void 0,
    !0,
    {
      fileName: "/home/user/excel-collab/src/containers/canvas/ContextMenu.tsx",
      lineNumber: 132,
      columnNumber: 5
    },
    void 0
  );
});
Xl.displayName = "CanvasContextMenu";
class Nt {
  type;
  value;
  constructor(t, n) {
    this.type = t, this.value = n;
  }
  toString() {
    return this.value;
  }
}
const Vs = "";
class T2 {
  list;
  current = 0;
  start = 0;
  tokens = [];
  constructor(t) {
    this.list = [...t];
  }
  scan() {
    for (; !this.isAtEnd(); )
      this.start = this.current, this.scanToken();
    return this.tokens.push(new Nt(F.EOF, "")), this.tokens.length > 0 && this.tokens[0].type === F.EQUAL && this.tokens.shift(), this.tokens;
  }
  peek() {
    return this.isAtEnd() ? Vs : this.list[this.current];
  }
  match(t) {
    return this.peek() !== t ? !1 : (this.next(), !0);
  }
  next() {
    return this.isAtEnd() ? Vs : this.list[this.current++];
  }
  isAtEnd() {
    return this.current >= this.list.length;
  }
  addToken(t) {
    const n = this.list.slice(this.start, this.current).join("");
    this.tokens.push(new Nt(t, n));
  }
  string(t) {
    for (; !this.isAtEnd() && this.peek() !== t; )
      this.next();
    if (this.peek() !== t)
      throw fe("#VALUE!");
    this.next();
    const n = this.list.slice(this.start + 1, this.current - 1).join("");
    if (t === "'" && this.match("!")) {
      this.tokens.push(new Nt(F.SHEET_NAME, n));
      return;
    }
    this.tokens.push(new Nt(F.STRING, n));
  }
  getDigits() {
    for (; !this.isAtEnd() && Gt(this.peek()); )
      this.next();
  }
  getAlphas() {
    for (; !this.isAtEnd() && Qo(this.peek()); )
      this.next();
  }
  matchR1C1() {
    if (this.match("[")) {
      if (this.match("-"), Gt(this.peek()))
        this.getDigits();
      else
        throw fe("#VALUE!");
      if (this.peek() !== "]")
        throw fe("#VALUE!");
      this.next();
    } else
      this.getDigits();
  }
  matchScientificCounting() {
    if (this.match("E") || this.match("e")) {
      if (this.match("+") || this.match("-"))
        return this.getDigits(), this.addToken(F.NUMBER), !0;
      if (Gt(this.peek()))
        return this.getDigits(), this.addToken(F.NUMBER), !0;
      throw fe("#VALUE!");
    }
    return !1;
  }
  number() {
    this.getDigits(), !(this.matchScientificCounting() || (this.match(".") && this.getDigits(), this.matchScientificCounting())) && this.addToken(F.NUMBER);
  }
  addIdentifier() {
    for (; !this.isAtEnd() && this.anyChar(this.peek()); )
      this.next();
    let t = this.list.slice(this.start, this.current).join("");
    const n = t.toUpperCase();
    if (cu.has(n) && this.match("(")) {
      this.tokens.push(new Nt(F.EXCEL_FUNCTION, n));
      return;
    }
    if (["IF", "CHOOSE"].includes(n) && this.match("(")) {
      this.tokens.push(new Nt(F.REF_FUNCTION_COND, n));
      return;
    }
    if (["INDEX", "OFFSET", "INDIRECT"].includes(n) && this.match("(")) {
      this.tokens.push(new Nt(F.REF_FUNCTION, n));
      return;
    }
    if (this.match("!")) {
      this.tokens.push(new Nt(F.SHEET_NAME, t));
      return;
    }
    let r = F.BOOL;
    if (n === "TRUE" || n === "FALSE")
      t = n, r = F.BOOL;
    else if (No(n))
      t = n, r = n === "#REF!" ? F.ERROR_REF : F.ERROR;
    else if (su.test(n))
      r = F.CELL;
    else if (au.test(n))
      r = F.COLUMN;
    else if (lu.test(n))
      r = F.ROW;
    else if (Ci.test(n))
      r = F.DEFINED_NAME;
    else
      throw fe("#NAME?");
    this.tokens.push(new Nt(r, t));
  }
  scanToken() {
    const t = this.next();
    switch (t) {
      case "$": {
        Qo(this.peek()) ? (this.getAlphas(), this.match("$") ? Gt(this.peek()) ? (this.getDigits(), this.addToken(F.CELL)) : this.addIdentifier() : Gt(this.peek()) ? (this.getDigits(), this.addToken(F.CELL)) : this.addToken(F.COLUMN)) : Gt(this.peek()) ? (this.getDigits(), this.addToken(F.ROW)) : this.addIdentifier();
        break;
      }
      case "r":
      case "R": {
        if (this.matchR1C1(), this.match("C") || this.match("c")) {
          this.matchR1C1();
          const n = this.list.slice(this.start, this.current).join("").toUpperCase();
          this.tokens.push(new Nt(F.R1C1, n));
        } else
          this.addIdentifier();
        break;
      }
      case "(":
        this.addToken(F.LEFT_BRACKET);
        break;
      case ")":
        this.addToken(F.RIGHT_BRACKET);
        break;
      case ",":
        this.addToken(F.COMMA);
        break;
      case ":":
        this.addToken(F.COLON);
        break;
      case Fn:
        this.addToken(F.EQUAL);
        break;
      case "<":
        this.match(">") ? this.addToken(F.NOT_EQUAL) : this.match("=") ? this.addToken(F.LESS_EQUAL) : this.addToken(F.LESS);
        break;
      case ">":
        this.match("=") ? this.addToken(F.GREATER_EQUAL) : this.addToken(F.GREATER);
        break;
      case "+":
        this.addToken(F.PLUS);
        break;
      case "-":
        this.addToken(F.MINUS);
        break;
      case "*":
        this.addToken(F.STAR);
        break;
      case "/":
        this.addToken(F.SLASH);
        break;
      case "^":
        this.addToken(F.EXPONENT);
        break;
      case "&":
        this.addToken(F.CONCATENATE);
        break;
      case "%":
        this.addToken(F.PERCENT);
        break;
      case '"':
        this.string(t);
        break;
      case "'":
        this.string(t);
        break;
      case "!":
        this.addToken(F.EXCLAMATION);
        break;
      case ";":
        this.addToken(F.SEMICOLON);
        break;
      case "{":
        this.addToken(F.lEFT_BRACE);
        break;
      case "}":
        this.addToken(F.RIGHT_BRACE);
        break;
      case " ":
        break;
      case "\r":
      case "	":
      case `
`:
        break;
      default:
        Gt(t) ? this.number() : this.addIdentifier();
        break;
    }
  }
  anyChar(t) {
    return !'(),:=<>+-*/^&%"{}!'.includes(t) && !this.isWhiteSpace(t);
  }
  isWhiteSpace(t) {
    return t === " " || t === "\r" || t === `
` || t === "	";
  }
}
class Yn {
  left;
  right;
  operator;
  constructor(t, n, r) {
    this.left = t, this.operator = n, this.right = r;
  }
  accept(t) {
    return t.visitBinaryExpression(this);
  }
  handleConcatenate(t) {
    const n = t.toString();
    return this.operator.type === F.CONCATENATE && t instanceof li && t.value.type === F.STRING ? JSON.stringify(n) : n;
  }
  toString() {
    const t = this.handleConcatenate(this.left), n = this.handleConcatenate(this.right);
    return `${t}${this.operator.toString()}${n}`;
  }
}
class x2 {
  right;
  operator;
  constructor(t, n) {
    this.operator = t, this.right = n;
  }
  accept(t) {
    return t.visitUnaryExpression(this);
  }
  toString() {
    return this.operator.toString() + this.right.toString();
  }
}
class D2 {
  left;
  operator;
  constructor(t, n) {
    this.operator = t, this.left = n;
  }
  accept(t) {
    return t.visitPostUnaryExpression(this);
  }
  toString() {
    return this.left.toString() + this.operator.toString();
  }
}
class li {
  value;
  constructor(t) {
    this.value = t;
  }
  accept(t) {
    return t.visitLiteralExpression(this);
  }
  toString() {
    return this.value.toString();
  }
}
class Xn {
  value;
  sheetName;
  constructor(t, n) {
    this.value = t, this.sheetName = n;
  }
  accept(t) {
    return t.visitCellExpression(this);
  }
  toString() {
    const t = this.value.toString().toUpperCase();
    if (this.sheetName) {
      let n = this.sheetName.toString();
      return n.includes(" ") && (n = `'${n}'`), `${n}!${t}`;
    } else
      return t;
  }
}
class O2 {
  name;
  params;
  constructor(t, n) {
    this.name = t, this.params = n;
  }
  accept(t) {
    return t.visitCallExpression(this);
  }
  toString() {
    return `${this.name.toString()}(${this.params.map((t) => t.toString()).join(",")})`;
  }
}
class k2 {
  right;
  operator;
  left;
  constructor(t, n, r) {
    this.left = t, this.operator = n, this.right = r;
  }
  accept(t) {
    return t.visitCellRangeExpression(this);
  }
  toString() {
    return this.left.toString() + this.operator.toString() + this.right.toString();
  }
}
class L2 {
  value;
  constructor(t) {
    this.value = t;
  }
  accept(t) {
    return t.visitGroupExpression(this);
  }
  toString() {
    return `(${this.value.toString()})`;
  }
}
class F2 {
  value;
  constructor(t) {
    this.value = t;
  }
  accept(t) {
    return t.visitArrayExpression(this);
  }
  toString() {
    return `{${this.value.map((n) => n.toString()).join(",")}}`;
  }
}
class P2 {
  tokens;
  current = 0;
  constructor(t) {
    this.tokens = t;
  }
  parse() {
    const t = [];
    for (; !this.isAtEnd(); )
      t.push(this.expression());
    return t;
  }
  expression() {
    return this.comparison();
  }
  comparison() {
    let t = this.concatenate();
    for (; this.match(
      F.EQUAL,
      F.NOT_EQUAL,
      F.GREATER,
      F.GREATER_EQUAL,
      F.LESS,
      F.LESS_EQUAL
    ); ) {
      const n = this.previous(), r = this.concatenate();
      t = new Yn(t, n, r);
    }
    return t;
  }
  concatenate() {
    let t = this.term();
    for (; this.match(F.CONCATENATE); ) {
      const n = this.previous(), r = this.term();
      t = new Yn(t, n, r);
    }
    return t;
  }
  term() {
    let t = this.factor();
    for (; this.match(F.PLUS, F.MINUS); ) {
      const n = this.previous(), r = this.factor();
      t = new Yn(t, n, r);
    }
    return t;
  }
  factor() {
    let t = this.expo();
    for (; this.match(F.SLASH, F.STAR); ) {
      const n = this.previous(), r = this.expo();
      t = new Yn(t, n, r);
    }
    return t;
  }
  expo() {
    let t = this.unary();
    for (; this.match(F.EXPONENT); ) {
      const n = this.previous(), r = this.unary();
      t = new Yn(t, n, r);
    }
    return t;
  }
  unary() {
    if (this.match(F.PLUS, F.MINUS)) {
      const t = this.previous(), n = this.unary();
      return new x2(t, n);
    }
    return this.postUnary();
  }
  postUnary() {
    let t = this.cellRange();
    if (this.match(F.PERCENT)) {
      const n = this.previous();
      t = new D2(n, t);
    }
    return t;
  }
  cellRange() {
    let t = this.sheetRange();
    for (; this.match(F.COLON); ) {
      const n = this.previous(), r = this.sheetRange(), o = this.convertToCellExpression(r), s = this.convertToCellExpression(t);
      t = new k2(s, n, o);
    }
    return t;
  }
  convertToCellExpression(t) {
    if (t instanceof Xn)
      return t;
    if (t instanceof li && t.value.type === F.NUMBER)
      return new Xn(t.value, void 0);
    throw fe("#NAME?");
  }
  sheetRange() {
    if (this.match(F.SHEET_NAME)) {
      const t = this.previous();
      let n = this.call();
      if (n instanceof Xn)
        return new Xn(n.value, t);
      throw fe("#REF!");
    }
    return this.call();
  }
  call() {
    for (; this.match(
      F.EXCEL_FUNCTION,
      F.REF_FUNCTION,
      F.REF_FUNCTION_COND
    ); ) {
      const t = this.previous(), n = [];
      if (!this.check(F.RIGHT_BRACKET))
        do {
          if (this.peek().type == F.RIGHT_BRACKET)
            break;
          n.push(this.expression());
        } while (this.match(F.COMMA));
      return this.expect(F.RIGHT_BRACKET), new O2(t, n);
    }
    return this.primary();
  }
  primary() {
    if (this.match(F.lEFT_BRACE)) {
      const t = [];
      do {
        if (this.peek().type == F.RIGHT_BRACE)
          break;
        t.push(this.expression());
      } while (this.match(F.COMMA));
      return this.expect(F.RIGHT_BRACE), new F2(t);
    }
    if (this.match(F.LEFT_BRACKET)) {
      const t = this.expression();
      return this.expect(F.RIGHT_BRACKET), new L2(t);
    }
    if (this.match(F.NUMBER, F.STRING, F.BOOL))
      return new li(this.previous());
    if (this.match(
      F.CELL,
      F.COLUMN,
      F.ROW,
      F.DEFINED_NAME,
      F.R1C1
    )) {
      const t = this.previous();
      return new Xn(t, void 0);
    }
    throw this.match(F.ERROR, F.ERROR_REF) ? fe(this.previous().value) : fe("#VALUE!");
  }
  match(...t) {
    const { type: n } = this.peek();
    return t.includes(n) ? (this.next(), !0) : !1;
  }
  previous() {
    return this.tokens[this.current - 1];
  }
  check(t) {
    return this.peek().type === t;
  }
  expect(t) {
    if (this.check(t))
      return this.next(), this.previous();
    throw fe("#VALUE!");
  }
  next() {
    this.current++;
  }
  isAtEnd() {
    return this.peek().type === F.EOF;
  }
  peek() {
    return this.current < this.tokens.length ? this.tokens[this.current] : new Nt(F.EOF, "");
  }
}
class U2 {
  expressions;
  cellDataMap;
  currentCoord;
  constructor(t, n, r) {
    this.expressions = t, this.cellDataMap = r, this.currentCoord = n;
  }
  interpret() {
    let t = [];
    for (const n of this.expressions) {
      const r = this.evaluate(n), o = this.getRangeCellValue(r);
      Array.isArray(o) ? t = t.concat(o) : t.push(o);
    }
    return t;
  }
  visitBinaryExpression(t) {
    let n = this.evaluate(t.left), r = this.evaluate(t.right);
    switch (n = this.getOneValue(n), r = this.getOneValue(r), [
      F.MINUS,
      F.PLUS,
      F.SLASH,
      F.STAR,
      F.EXPONENT
    ].includes(t.operator.type) && (n = this.parseNumber(n), r = this.parseNumber(r)), t.operator.type) {
      case F.MINUS:
        return n - r;
      case F.PLUS:
        return n + r;
      case F.SLASH: {
        if (r === 0)
          throw fe("#DIV/0!");
        return n / r;
      }
      case F.STAR:
        return n * r;
      case F.EXPONENT:
        return Math.pow(n, r);
      case F.EQUAL:
        return n === r;
      case F.NOT_EQUAL:
        return n !== r;
      case F.GREATER:
        return n > r;
      case F.GREATER_EQUAL:
        return n >= r;
      case F.LESS:
        return n < r;
      case F.LESS_EQUAL:
        return n <= r;
      case F.CONCATENATE:
        return `${n}${r}`;
      default:
        throw fe("#GETTING_DATA");
    }
  }
  visitCallExpression(t) {
    const n = this.cellDataMap.getFunction(t.name.value);
    if (n && typeof n == "function") {
      let r = [];
      for (const l of t.params) {
        const f = this.evaluate(l);
        if (f instanceof wt) {
          const d = this.getCellValue(f, !1);
          r = r.concat(d);
        } else
          r.push(f);
      }
      const o = n(...r), s = No(o);
      if (s)
        throw fe(s);
      return o;
    }
    throw fe("#NAME?");
  }
  visitCellExpression(t) {
    let n = "";
    if (t.sheetName) {
      const o = this.cellDataMap.getSheetInfo("", t.sheetName.value);
      if (!o?.sheetId)
        throw fe("#REF!");
      n = o.sheetId;
    }
    let r;
    if ([
      F.DEFINED_NAME,
      F.CELL,
      F.COLUMN,
      F.R1C1
    ].includes(t.value.type) && (r = this.cellDataMap.getDefinedName(t.value.value)), r || (t.value.type === F.R1C1 ? r = Ru(t.value.value, this.currentCoord) : r = Ot(t.value.value)), !r)
      throw fe("#REF!");
    return n && (r.sheetId = n), r;
  }
  visitLiteralExpression(t) {
    const { type: n, value: r } = t.value;
    switch (n) {
      case F.STRING:
        return r;
      case F.NUMBER: {
        const [o, s] = Un(r);
        if (o)
          return s;
        throw fe("#VALUE!");
      }
      case F.BOOL:
        return r === "TRUE";
      default:
        throw fe("#VALUE!");
    }
  }
  visitUnaryExpression(t) {
    const n = this.evaluate(t.right), r = this.parseNumber(this.getOneValue(n));
    switch (t.operator.type) {
      case F.MINUS:
        return -r;
      case F.PLUS:
        return r;
      default:
        throw fe("#VALUE!");
    }
  }
  visitCellRangeExpression(t) {
    switch (t.operator.type) {
      case F.COLON: {
        const n = this.visitCellExpression(t.left), r = this.visitCellExpression(t.right), o = Ii(n, r);
        if (!o)
          throw fe("#NAME?");
        return o;
      }
      default:
        throw fe("#NAME?");
    }
  }
  visitGroupExpression(t) {
    return this.evaluate(t.value);
  }
  visitArrayExpression(t) {
    const n = [];
    for (const r of t.value)
      n.push(this.evaluate(r));
    return n;
  }
  visitPostUnaryExpression(t) {
    const n = this.evaluate(t.left), r = this.parseNumber(this.getOneValue(n));
    if (t.operator.type === F.PERCENT)
      return r * 0.01;
    throw fe("#VALUE!");
  }
  evaluate(t) {
    return t.accept(this);
  }
  getRangeCellValue(t) {
    if (t instanceof wt) {
      if (t.colCount === t.rowCount && t.colCount === 1)
        return this.getCellValue(t, !0);
      throw fe("#REF!");
    }
    return t;
  }
  getOneValue(t) {
    const n = this.getRangeCellValue(t);
    return Array.isArray(n) ? n[0] : n;
  }
  parseNumber(t) {
    const [n, r] = Un(t);
    if (!n)
      throw fe("#VALUE!");
    return r;
  }
  getCellValue(t, n) {
    let r = [];
    if (n) {
      const d = this.cellDataMap.getCell(t), c = this.cellDataMap.handleCell(d, {
        row: t.row,
        col: t.col,
        sheetId: t.sheetId
      });
      return r = r.concat(c), r;
    }
    const { row: o, col: s, sheetId: l } = t, f = this.cellDataMap.getSheetInfo(l);
    if (!f || o >= f.rowCount || s >= f.colCount)
      throw fe("#REF!");
    return Tn(t, f.rowCount, f.colCount, (d, c) => {
      const i = this.cellDataMap.getCell({
        row: d,
        col: c,
        rowCount: 1,
        colCount: 1,
        sheetId: l
      }), a = this.cellDataMap.handleCell(i, { row: d, col: c, sheetId: l });
      return r = r.concat(a), !1;
    }), r;
  }
}
function Kl(e, t = {
  row: 0,
  col: 0,
  sheetId: ""
}, n = new _2(), r = /* @__PURE__ */ new Map()) {
  if (r.has(e)) {
    const o = r.get(e);
    return o.result.length === 0 ? { result: [0] } : o;
  }
  try {
    n.handleCell = (c, i) => c ? c.formula ? Kl(c.formula, i, n, r).result : [c.value] : [];
    const o = {
      result: [],
      expressionStr: ""
    };
    r.set(e, o);
    const s = new T2(e).scan(), l = new P2(s).parse(), f = new U2(
      l,
      t,
      n
    ).interpret(), d = [];
    for (const c of l)
      d.push(c.toString());
    return o.expressionStr = d.join(""), o.result = f, o;
  } catch (o) {
    const s = No(o);
    if (s) {
      const l = {
        result: [s]
      };
      return r.set(e, l), l;
    }
    throw o;
  }
}
const Fo = "_test_";
class _2 {
  map = /* @__PURE__ */ new Map();
  definedNameMap = /* @__PURE__ */ new Map();
  currentSheetId = Fo;
  sheetList = [
    {
      sheetId: Fo,
      name: "_test_",
      rowCount: 200,
      colCount: 30,
      isHide: !1,
      sort: 1
    }
  ];
  getKey(t, n, r) {
    return `${t}_${n}_${r || Fo}`;
  }
  getFunction = Cl;
  setSheetList(t) {
    this.sheetList = t;
  }
  set(t, n) {
    const { row: r, col: o, sheetId: s } = t;
    for (let l = 0; l < n.length; l++)
      for (let f = 0; f < n[l].length; f++) {
        const d = this.getKey(r + l, o + f, s), c = n[l][f];
        typeof c == "string" && Na(c) ? this.map.set(d, { formula: c, value: "" }) : this.map.set(d, { value: c });
      }
  }
  getCell(t) {
    const { row: n, col: r, sheetId: o } = t, s = this.getKey(n, r, o);
    return this.map.get(s);
  }
  getSheetInfo(t, n) {
    if (n)
      return this.sheetList.find((o) => o.name === n);
    const r = t || this.currentSheetId;
    return this.sheetList.find((o) => o.sheetId === r);
  }
  handleCell = (t, n) => t ? [t.value] : [];
  setDefinedName(t, n) {
    const r = wt.makeRange(n);
    this.definedNameMap.set(t, r);
  }
  getDefinedName(t) {
    return this.definedNameMap.get(t);
  }
}
function B2(e, t) {
  const { currentSheetId: n, worksheets: r, workbook: o, definedNames: s } = e, l = /* @__PURE__ */ new Map(), f = {
    list: []
  }, d = {
    handleCell: () => [],
    getFunction: Cl,
    getCell: (c) => {
      const { row: i, col: a, sheetId: u } = c, h = nt(u || n, i, a);
      return r[h];
    },
    set: () => {
      throw fe("#REF!");
    },
    getSheetInfo: (c, i) => {
      if (i)
        return o.find((u) => u.name === i);
      const a = c || n;
      return o.find((u) => u.sheetId === a);
    },
    setDefinedName: () => {
      throw fe("#REF!");
    },
    getDefinedName: (c) => {
      const i = s[c];
      return i && wt.makeRange(i);
    }
  };
  for (const [c, i] of Object.entries(r)) {
    const a = jn(c);
    if (!a || a.sheetId !== n || !i?.formula)
      continue;
    const u = Kl(i?.formula, a, d, l);
    if (!u)
      continue;
    const h = u.result[0];
    h !== i.value && f.list.push({
      key: c,
      newValue: h,
      sheetId: n
    });
  }
  return t(f);
}
const Ar = 2, Tr = 2 ** 2, Po = 2 ** 3, xr = 2 ** 4, Dr = 2 ** 5, Or = 2 ** 6, ci = 2 ** 7, ui = 2 ** 8, fi = 2 ** 9, z2 = 0, H2 = 2958466, V2 = -694324, G2 = 35830291, $2 = -1, Gs = 1, Ln = 6, Jl = "general", Ql = "hash", er = "zero", jl = "qmark", Zl = "slash", eo = "group", di = "scale", ec = "comma", tc = "break", rs = "text", nc = "plus", rc = "minus", hi = "point", oc = "space", ic = "percent", os = "digit", sc = "calendar", to = "error", ac = "datetime", lc = "duration", cc = "condition", uc = "dbnum", fc = "natnum", dc = "locale", hc = "color", q2 = "modifier", mc = "ampm", gc = "escaped", pc = "string", vc = "skip", Nc = "exp", Cc = "fill", Ec = "paren", is = "char", Uo = Object.freeze({
  1078: "af",
  // Afrikaans
  1052: "sq",
  // Albanian
  1118: "am",
  // Amharic
  5121: "ar_DZ",
  // Arabic - Algeria
  15361: "ar_BH",
  // Arabic - Bahrain
  3073: "ar_EG",
  // Arabic - Egypt
  2049: "ar_IQ",
  // Arabic - Iraq
  11265: "ar_JO",
  // Arabic - Jordan
  13313: "ar_KW",
  // Arabic - Kuwait
  12289: "ar_LB",
  // Arabic - Lebanon
  4097: "ar_LY",
  // Arabic - Libya
  6145: "ar_MA",
  // Arabic - Morocco
  8193: "ar_OM",
  // Arabic - Oman
  16385: "ar_QA",
  // Arabic - Qatar
  1025: "ar_SA",
  // Arabic - Saudi Arabia
  10241: "ar_SY",
  // Arabic - Syria
  7169: "ar_TN",
  // Arabic - Tunisia
  14337: "ar_AE",
  // Arabic - United Arab Emirates
  9217: "ar_YE",
  // Arabic - Yemen
  1067: "hy",
  // Armenian
  1101: "as",
  // Assamese
  2092: "az_AZ",
  // Azeri - Cyrillic
  1068: "az_AZ",
  // Azeri - Latin
  1069: "eu",
  // Basque
  1059: "be",
  // Belarusian
  2117: "bn",
  // Bengali - Bangladesh
  1093: "bn_IN",
  // Bengali - India
  5146: "bs",
  // Bosnian
  1026: "bg",
  // Bulgarian
  1109: "my",
  // Burmese
  1027: "ca",
  // Catalan
  2052: "zh_CN",
  // Chinese - China
  3076: "zh_HK",
  // Chinese - Hong Kong SAR
  5124: "zh_MO",
  // Chinese - Macau SAR
  4100: "zh_SG",
  // Chinese - Singapore
  1028: "zh_TW",
  // Chinese - Taiwan
  1050: "hr",
  // Croatian
  1029: "cs",
  // Czech
  1030: "da",
  // Danish
  1125: "dv",
  // Divehi; Dhivehi; Maldivian
  2067: "nl_BE",
  // Dutch - Belgium
  1043: "nl_NL",
  // Dutch - Netherlands
  1126: "bin",
  // Edo
  3081: "en_AU",
  // English - Australia
  10249: "en_BZ",
  // English - Belize
  4105: "en_CA",
  // English - Canada
  9225: "en_CB",
  // English - Caribbean
  2057: "en_GB",
  // English - Great Britain
  16393: "en_IN",
  // English - India
  6153: "en_IE",
  // English - Ireland
  8201: "en_JM",
  // English - Jamaica
  5129: "en_NZ",
  // English - New Zealand
  13321: "en_PH",
  // English - Phillippines
  7177: "en_ZA",
  // English - Southern Africa
  11273: "en_TT",
  // English - Trinidad
  1033: "en_US",
  // English - United States
  12297: "en_ZW",
  // English - Zimbabwe
  1061: "et",
  // Estonian
  1071: "mk",
  // FYRO Macedonia
  1080: "fo",
  // Faroese
  1065: "fa",
  // Farsi - Persian
  1124: "fil",
  // Filipino
  1035: "fi",
  // Finnish
  2060: "fr_BE",
  // French - Belgium
  11276: "fr_CM",
  // French - Cameroon
  3084: "fr_CA",
  // French - Canada
  9228: "fr_CG",
  // French - Congo
  12300: "fr_CI",
  // French - Cote d'Ivoire
  1036: "fr_FR",
  // French - France
  5132: "fr_LU",
  // French - Luxembourg
  13324: "fr_ML",
  // French - Mali
  6156: "fr_MC",
  // French - Monaco
  14348: "fr_MA",
  // French - Morocco
  10252: "fr_SN",
  // French - Senegal
  4108: "fr_CH",
  // French - Switzerland
  7180: "fr",
  // French - West Indies
  1122: "fy_NL",
  // Frisian - Netherlands
  2108: "gd_IE",
  // Gaelic - Ireland
  1084: "gd",
  // Gaelic - Scotland
  1110: "gl",
  // Galician
  1079: "ka",
  // Georgian
  3079: "de_AT",
  // German - Austria
  1031: "de_DE",
  // German - Germany
  5127: "de_LI",
  // German - Liechtenstein
  4103: "de_LU",
  // German - Luxembourg
  2055: "de_CH",
  // German - Switzerland
  1032: "el",
  // Greek
  1140: "gn",
  // Guarani - Paraguay
  1095: "gu",
  // Gujarati
  1279: "en",
  // HID (Human Interface Device)
  1037: "he",
  // Hebrew
  1081: "hi",
  // Hindi
  1038: "hu",
  // Hungarian
  1039: "is",
  // Icelandic
  1136: "ig_NG",
  // Igbo - Nigeria
  1057: "id",
  // Indonesian
  1040: "it_IT",
  // Italian - Italy
  2064: "it_CH",
  // Italian - Switzerland
  1041: "ja",
  // Japanese
  1099: "kn",
  // Kannada
  1120: "ks",
  // Kashmiri
  1087: "kk",
  // Kazakh
  1107: "km",
  // Khmer
  1111: "kok",
  // Konkani
  1042: "ko",
  // Korean
  1088: "ky",
  // Kyrgyz - Cyrillic
  1108: "lo",
  // Lao
  1142: "la",
  // Latin
  1062: "lv",
  // Latvian
  1063: "lt",
  // Lithuanian
  2110: "ms_BN",
  // Malay - Brunei
  1086: "ms_MY",
  // Malay - Malaysia
  1100: "ml",
  // Malayalam
  1082: "mt",
  // Maltese
  1112: "mni",
  // Manipuri
  1153: "mi",
  // Maori
  1102: "mr",
  // Marathi
  1104: "mn",
  // Mongolian
  2128: "mn",
  // Mongolian
  1121: "ne",
  // Nepali
  1044: "no_NO",
  // Norwegian - Bokml
  2068: "no_NO",
  // Norwegian - Nynorsk
  1096: "or",
  // Oriya
  1045: "pl",
  // Polish
  1046: "pt_BR",
  // Portuguese - Brazil
  2070: "pt_PT",
  // Portuguese - Portugal
  1094: "pa",
  // Punjabi
  1047: "rm",
  // Raeto-Romance
  2072: "ro_MO",
  // Romanian - Moldova
  1048: "ro_RO",
  // Romanian - Romania
  1049: "ru",
  // Russian
  2073: "ru_MO",
  // Russian - Moldova
  1083: "se",
  // Sami Lappish
  1103: "sa",
  // Sanskrit
  3098: "sr_SP",
  // Serbian - Cyrillic
  2074: "sr_SP",
  // Serbian - Latin
  1072: "st",
  // Sesotho (Sutu)
  1074: "tn",
  // Setsuana
  1113: "sd",
  // Sindhi
  1115: "si",
  // Sinhala; Sinhalese
  1051: "sk",
  // Slovak
  1060: "sl",
  // Slovenian
  1143: "so",
  // Somali
  1070: "sb",
  // Sorbian
  11274: "es_AR",
  // Spanish - Argentina
  16394: "es_BO",
  // Spanish - Bolivia
  13322: "es_CL",
  // Spanish - Chile
  9226: "es_CO",
  // Spanish - Colombia
  5130: "es_CR",
  // Spanish - Costa Rica
  7178: "es_DO",
  // Spanish - Dominican Republic
  12298: "es_EC",
  // Spanish - Ecuador
  17418: "es_SV",
  // Spanish - El Salvador
  4106: "es_GT",
  // Spanish - Guatemala
  18442: "es_HN",
  // Spanish - Honduras
  2058: "es_MX",
  // Spanish - Mexico
  19466: "es_NI",
  // Spanish - Nicaragua
  6154: "es_PA",
  // Spanish - Panama
  15370: "es_PY",
  // Spanish - Paraguay
  10250: "es_PE",
  // Spanish - Peru
  20490: "es_PR",
  // Spanish - Puerto Rico
  1034: "es_ES",
  // Spanish - Spain (Traditional)
  14346: "es_UY",
  // Spanish - Uruguay
  8202: "es_VE",
  // Spanish - Venezuela
  1089: "sw",
  // Swahili
  2077: "sv_FI",
  // Swedish - Finland
  1053: "sv_SE",
  // Swedish - Sweden
  1114: "syc",
  // Syriac
  1064: "tg",
  // Tajik
  1097: "ta",
  // Tamil
  1092: "tt",
  // Tatar
  1098: "te",
  // Telugu
  1054: "th",
  // Thai
  1105: "bo",
  // Tibetan
  1073: "ts",
  // Tsonga
  1055: "tr",
  // Turkish
  1090: "tk",
  // Turkmen
  1058: "uk",
  // Ukrainian
  1056: "ur",
  // Urdu
  2115: "uz_UZ",
  // Uzbek - Cyrillic
  1091: "uz_UZ",
  // Uzbek - Latin
  1075: "ve",
  // Venda
  1066: "vi",
  // Vietnamese
  1106: "cy",
  // Welsh
  1076: "xh",
  // Xhosa
  1085: "yi",
  // Yiddish
  1077: "zu"
  // Zulu
}), bc = /^([a-z\d]+)(?:[_-]([a-z\d]+))?(?:\.([a-z\d]+))?(?:@([a-z\d]+))?$/i, Dn = {}, y = (e, t = "") => e.replace(/~/g, t).split(";"), ve = (e, t = 0, n = 0) => (e.mmm || (e.mmm = t < 1 ? e.mmmm.concat() : e.mmmm.map((r) => {
  const o = r.slice(0, t % 10);
  return o + (t < 10 || r === o ? "" : ".");
})), e.ddd || (e.ddd = n < 1 ? e.dddd.concat() : e.dddd.map((r) => {
  const o = r.slice(0, n % 10);
  return o + (n < 10 || r === o ? "" : ".");
})), !e.mmm6 && e.mmmm6 && (e.mmm6 = e.mmmm6), e), W2 = {
  group: " ",
  decimal: ".",
  positive: "+",
  negative: "-",
  percent: "%",
  exponent: "E",
  nan: "NaN",
  infinity: "∞",
  ampm: y("AM;PM"),
  mmmm6: y("Muharram;Safar;Rabiʻ I;Rabiʻ II;Jumada I;Jumada II;Rajab;Shaʻban;Ramadan;Shawwal;Dhuʻl-Qiʻdah;Dhuʻl-Hijjah"),
  mmm6: y("Muh.;Saf.;Rab. I;Rab. II;Jum. I;Jum. II;Raj.;Sha.;Ram.;Shaw.;Dhuʻl-Q.;Dhuʻl-H."),
  mmmm: y("January;February;March;April;May;June;July;August;September;October;November;December"),
  mmm: y("Jan;Feb;Mar;Apr;May;Jun;Jul;Aug;Sep;Oct;Nov;Dec"),
  dddd: y("Sunday;Monday;Tuesday;Wednesday;Thursday;Friday;Saturday"),
  ddd: y("Sun;Mon;Tue;Wed;Thu;Fri;Sat"),
  bool: y("TRUE;FALSE"),
  preferMDY: !1
};
function wc(e) {
  const t = bc.exec(e);
  if (!t)
    throw new SyntaxError(`Malformed locale: ${e}`);
  return {
    lang: t[1] + (t[2] ? "_" + t[2] : ""),
    language: t[1],
    territory: t[2] || ""
  };
}
function ss(e) {
  if (typeof e == "number")
    return Uo[e & 65535] || null;
  const t = parseInt(e, 16);
  return isFinite(t) && Uo[t & 65535] ? Uo[t & 65535] || null : bc.test(e) ? e : null;
}
function Y2(e) {
  const t = ss(e);
  let n = null;
  if (t) {
    const r = wc(t);
    n = Dn[r.lang] || Dn[r.language] || null;
  }
  return n;
}
function mi(e) {
  return Object.assign({}, W2, e);
}
function B(e, t) {
  const n = typeof t == "object" ? t : wc(t);
  return Dn[n.lang] = mi(e), n.language !== n.lang && !Dn[n.language] && (Dn[n.language] = mi(e)), Dn[n.lang];
}
const no = mi({ group: ",", preferMDY: !0 });
no.isDefault = !0;
B({
  group: ",",
  ampm: y("上午;下午"),
  mmmm: y("一月;二月;三月;四月;五月;六月;七月;八月;九月;十月;十一月;十二月"),
  mmm: y("1月;2月;3月;4月;5月;6月;7月;8月;9月;10月;11月;12月"),
  dddd: y("~日;~一;~二;~三;~四;~五;~六", "星期"),
  ddd: y("周日;周一;周二;周三;周四;周五;周六")
}, "zh_CN");
const as = {
  group: ",",
  ampm: y("上午;下午"),
  mmmm: y("1月;2月;3月;4月;5月;6月;7月;8月;9月;10月;11月;12月"),
  mmm: y("1月;2月;3月;4月;5月;6月;7月;8月;9月;10月;11月;12月"),
  dddd: y("~日;~一;~二;~三;~四;~五;~六", "星期"),
  ddd: y("周日;周一;周二;周三;周四;周五;周六")
};
B({
  ...as,
  nan: "非數值",
  dddd: y("~日;~一;~二;~三;~四;~五;~六", "星期")
}, "zh_TW");
B({
  ...as,
  dddd: y("~日;~一;~二;~三;~四;~五;~六", "星期")
}, "zh_HK");
B({
  ...as,
  ampm: y("午前;午後"),
  dddd: y("日~;月~;火~;水~;木~;金~;土~", "曜日"),
  ddd: y("日;月;火;水;木;金;土")
}, "ja");
B({
  group: ",",
  ampm: y("오전;오후"),
  mmmm: y("1월;2월;3월;4월;5월;6월;7월;8월;9월;10월;11월;12월"),
  mmm: y("1월;2월;3월;4월;5월;6월;7월;8월;9월;10월;11월;12월"),
  dddd: y("일요일;월요일;화요일;수요일;목요일;금요일;토요일"),
  ddd: y("일;월;화;수;목;금;토")
}, "ko");
B({
  group: ",",
  ampm: y("ก่อนเที่ยง;หลังเที่ยง"),
  mmmm: y("มกร~;กุมภาพันธ์;มีน~;เมษายน;พฤษภ~;มิถุนายน;กรกฎ~;สิงห~;กันยายน;ตุล~;พฤศจิกายน;ธันว~", "าคม"),
  mmm: y("ม.ค.;ก.พ.;มี.ค.;เม.ย.;พ.ค.;มิ.ย.;ก.ค.;ส.ค.;ก.ย.;ต.ค.;พ.ย.;ธ.ค."),
  dddd: y("วันอาทิตย์;วันจันทร์;วันอังคาร;วันพุธ;วันพฤหัสบดี;วันศุกร์;วันเสาร์"),
  ddd: y("อา.;จ.;อ.;พ.;พฤ.;ศ.;ส.")
}, "th");
B(ve({
  decimal: ",",
  ampm: y("dop.;odp."),
  mmmm: y("ledna;února;března;dubna;května;června;července;srpna;září;října;listopadu;prosince"),
  mmm: y("I;II;III;IV;V;VI;VII;VIII;IX;X;XI;XII"),
  dddd: y("neděle;pondělí;úterý;středa;čtvrtek;pátek;sobota"),
  bool: y("PRAVDA;NEPRAVDA")
}, -1, 2), "cs");
B(ve({
  group: ".",
  decimal: ",",
  mmmm: y("januar;februar;marts;april;maj;juni;juli;august;september;oktober;november;december"),
  dddd: y("søn~;man~;tirs~;ons~;tors~;fre~;lør~", "dag"),
  bool: y("SAND;FALSK")
}, 13, 13), "da");
B(ve({
  group: ".",
  decimal: ",",
  ampm: y("a.m.;p.m."),
  mmmm: y("januari;februari;maart;april;mei;juni;juli;augustus;september;oktober;november;december"),
  mmm: y("jan.;feb.;mrt.;apr.;mei;jun.;jul.;aug.;sep.;okt.;nov.;dec."),
  dddd: y("zondag;maandag;dinsdag;woensdag;donderdag;vrijdag;zaterdag"),
  bool: y("WAAR;ONWAAR")
}, -1, 2), "nl");
B({ group: ",", preferMDY: !0 }, "en");
B({ group: ",", preferMDY: !0 }, "en_US");
B({ group: "," }, "en_AU");
B({ group: "," }, "en_CA");
B({ group: "," }, "en_GB");
B({ group: ",", mmm: y("Jan;Feb;Mar;Apr;May;Jun;Jul;Aug;Sept;Oct;Nov;Dec") }, "en_IE");
B(ve({
  decimal: ",",
  nan: "epäluku",
  ampm: y("ap.;ip."),
  mmmm: y("tammi~;helmi~;maalis~;huhti~;touko~;kesä~;heinä~;elo~;syys~;loka~;marras~;joulu~", "kuuta"),
  mmm: y("tammik.;helmik.;maalisk.;huhtik.;toukok.;kesäk.;heinäk.;elok.;syysk.;lokak.;marrask.;jouluk."),
  dddd: y("sunnun~;maanan~;tiis~;keskiviikkona;tors~;perjan~;lauan~", "taina"),
  bool: y("TOSI;EPÄTOSI")
}, -1, 2), "fi");
const ls = ve({
  group: " ",
  decimal: ",",
  mmmm: y("janvier;février;mars;avril;mai;juin;juillet;août;septembre;octobre;novembre;décembre"),
  mmm: y("janv.;févr.;mars;avr.;mai;juin;juil.;août;sept.;oct.;nov.;déc."),
  dddd: y("~manche;lun~;mar~;mercre~;jeu~;vendre~;same~", "di"),
  bool: y("VRAI;FAUX")
}, -1, 13);
B({ ...ls }, "fr");
B({ ...ls, mmm: y("janv.;févr.;mars;avr.;mai;juin;juill.;août;sept.;oct.;nov.;déc.") }, "fr_CA");
B({ group: "'", decimal: ".", ...ls }, "fr_CH");
const Ic = ve({
  mmmm: y("Januar;Februar;März;April;Mai;Juni;Juli;August;September;Oktober;November;Dezember"),
  mmm: y("Jan.;Feb.;März;Apr.;Mai;Juni;Juli;Aug.;Sept.;Okt.;Nov.;Dez."),
  dddd: y("Sonn~;Mon~;Diens~;Mittwoch;Donners~;Frei~;Sams~", "tag"),
  bool: y("WAHR;FALSCH")
}, -1, 12);
B({ group: ".", decimal: ",", ...Ic }, "de");
B({ group: "'", decimal: ".", ...Ic }, "de_CH");
B(ve({
  group: ".",
  decimal: ",",
  ampm: y("π.μ.;μ.μ."),
  mmmm: y("Ιανουαρ~;Φεβρουαρ~;Μαρτ~;Απριλ~;Μαΐου;Ιουν~;Ιουλ~;Αυγούστου;Σεπτεμβρ~;Οκτωβρ~;Νοεμβρ~;Δεκεμβρ~", "ίου"),
  mmm: y("Ιαν;Φεβ;Μαρ;Απρ;Μαΐ;Ιουν;Ιουλ;Αυγ;Σεπ;Οκτ;Νοε;Δεκ"),
  dddd: y("Κυριακή;Δευτέρα;Τρίτη;Τετάρτη;Πέμπτη;Παρασκευή;Σάββατο")
}, -1, 3), "el");
B({
  decimal: ",",
  ampm: y("de.;du."),
  mmmm: y("január;február;március;április;május;június;július;augusztus;szeptember;október;november;december"),
  mmm: y("jan.;febr.;márc.;ápr.;máj.;jún.;júl.;aug.;szept.;okt.;nov.;dec."),
  dddd: y("vasárnap;hétfő;kedd;szerda;csütörtök;péntek;szombat"),
  ddd: y("V;H;K;Sze;Cs;P;Szo"),
  bool: y("IGAZ;HAMIS")
}, "hu");
B(ve({
  group: ".",
  decimal: ",",
  ampm: y("f.h.;e.h."),
  mmmm: y("janúar;febrúar;mars;apríl;maí;júní;júlí;ágúst;september;október;nóvember;desember"),
  dddd: y("sunnu~;mánu~;þriðju~;miðviku~;fimmtu~;föstu~;laugar~", "dagur")
}, 13, 13), "is");
B(ve({
  group: ".",
  decimal: ",",
  mmmm: y("Januari;Februari;Maret;April;Mei;Juni;Juli;Agustus;September;Oktober;November;Desember"),
  dddd: y("Minggu;Senin;Selasa;Rabu;Kamis;Jumat;Sabtu")
}, 3, 3), "id");
const Mc = ve({
  mmmm: y("gennaio;febbraio;marzo;aprile;maggio;giugno;luglio;agosto;settembre;ottobre;novembre;dicembre"),
  dddd: y("domenica;lunedì;martedì;mercoledì;giovedì;venerdì;sabato"),
  bool: y("VERO;FALSO")
}, 3, 3);
B({ group: ".", decimal: ",", ...Mc }, "it");
B({ group: "'", decimal: ".", ...Mc }, "it_CH");
const yc = {
  decimal: ",",
  ampm: y("a.m.;p.m."),
  mmmm: y("januar;februar;mars;april;mai;juni;juli;august;september;oktober;november;desember"),
  mmm: y("jan.;feb.;mar.;apr.;mai;jun.;jul.;aug.;sep.;okt.;nov.;des."),
  dddd: y("søn~;man~;tirs~;ons~;tors~;fre~;lør~", "dag"),
  bool: y("SANN;USANN")
};
B(ve({ ...yc }, -1, 13), "nb");
B(ve({ ...yc }, -1, 13), "no");
B(ve({
  decimal: ",",
  mmmm: y("stycznia;lutego;marca;kwietnia;maja;czerwca;lipca;sierpnia;września;października;listopada;grudnia"),
  dddd: y("niedziela;poniedziałek;wtorek;środa;czwartek;piątek;sobota"),
  ddd: y("niedz.;pon.;wt.;śr.;czw.;pt.;sob."),
  bool: y("PRAWDA;FAŁSZ")
}, 3, -1), "pl");
const Sc = {
  group: ".",
  decimal: ",",
  mmmm: y("janeiro;fevereiro;março;abril;maio;junho;julho;agosto;setembro;outubro;novembro;dezembro"),
  dddd: y("domingo;segunda-feira;terça-feira;quarta-feira;quinta-feira;sexta-feira;sábado"),
  bool: y("VERDADEIRO;FALSO")
};
B(ve(Sc, 13, 13), "pt");
B(ve(Sc, 13, 13), "pt_BR");
B({
  decimal: ",",
  nan: "не число",
  mmmm: y("января;февраля;марта;апреля;мая;июня;июля;августа;сентября;октября;ноября;декабря"),
  mmm: y("янв.;февр.;мар.;апр.;мая;июн.;июл.;авг.;сент.;окт.;нояб.;дек."),
  dddd: y("воскресенье;понедельник;вторник;среда;четверг;пятница;суббота"),
  ddd: y("вс;пн;вт;ср;чт;пт;сб"),
  mmmm6: y("рамадан;шавваль;зуль-каада;зуль-хиджжа;мухаррам;раби-уль-авваль;раби-уль-ахир;джумад-уль-авваль;джумад-уль-ахир;раджаб;шаабан;рамадан"),
  mmm6: y("рам.;шав.;зуль-к.;зуль-х.;мух.;раб. I;раб. II;джум. I;джум. II;радж.;шааб.;рам."),
  bool: y("ИСТИНА;ЛОЖЬ")
}, "ru");
B(ve({
  decimal: ",",
  mmmm: y("januára;februára;marca;apríla;mája;júna;júla;augusta;septembra;októbra;novembra;decembra"),
  dddd: y("nedeľa;pondelok;utorok;streda;štvrtok;piatok;sobota")
}, 3, 2), "sk");
const Tt = {
  group: ".",
  decimal: ",",
  ampm: y("a. m.;p. m."),
  mmmm: y("enero;febrero;marzo;abril;mayo;junio;julio;agosto;septiem~;octu~;noviem~;diciem~", "bre"),
  mmm: y("ene;feb;mar;abr;may;jun;jul;ago;sept;oct;nov;dic"),
  dddd: y("domingo;lunes;martes;miércoles;jueves;viernes;sábado"),
  ddd: y("dom;lun;mar;mié;jue;vie;sáb"),
  bool: y("VERDADERO;FALSO")
}, X2 = y("ene;feb;mar;abr;may;jun;jul;ago;sep;oct;nov;dic"), cs = y("ene.;feb.;mar.;abr.;may.;jun.;jul.;ago.;sept.;oct.;nov.;dic.");
B({ ...Tt }, "es");
B({ ...Tt }, "es_AR");
B({ ...Tt }, "es_BO");
B({ ...Tt }, "es_CL");
B({ ...Tt }, "es_CO");
B({ ...Tt }, "es_EC");
B({ ...Tt, mmm: X2, ampm: y("a.m.;p.m.") }, "es_MX");
B({ ...Tt, mmm: cs }, "es_PY");
B({ ...Tt, mmm: cs }, "es_UY");
B({ ...Tt, mmm: cs, mmmm: y("enero;febrero;marzo;abril;mayo;junio;julio;agosto;setiembre;octubre;noviembre;diciembre") }, "es_VE");
B({
  decimal: ",",
  ampm: y("fm;em"),
  mmmm: y("januari;februari;mars;april;maj;juni;juli;augusti;september;oktober;november;december"),
  mmm: y("jan.;feb.;mars;apr.;maj;juni;juli;aug.;sep.;okt.;nov.;dec."),
  dddd: y("sön~;mån~;tis~;ons~;tors~;fre~;lör~", "dag"),
  ddd: y("sön;mån;tis;ons;tors;fre;lör")
}, "sv");
B(ve({
  group: ".",
  decimal: ",",
  ampm: y("ÖÖ;ÖS"),
  mmmm: y("Ocak;Şubat;Mart;Nisan;Mayıs;Haziran;Temmuz;Ağustos;Eylül;Ekim;Kasım;Aralık"),
  mmm: y("Oca;Şub;Mar;Nis;May;Haz;Tem;Ağu;Eyl;Eki;Kas;Ara"),
  dddd: y("Pazar;Pazartesi;Salı;Çarşamba;Perşembe;Cuma;Cumartesi"),
  ddd: y("Paz;Pzt;Sal;Çar;Per;Cum;Cmt"),
  bool: y("DOĞRU;YANLIŞ")
}, 3, -1), "tr");
B({
  group: ",",
  ampm: y("yb;yh"),
  mmmm: y("Ionawr;Chwefror;Mawrth;Ebrill;Mai;Mehefin;Gorffennaf;Awst;Medi;Hydref;Tachwedd;Rhagfyr"),
  mmm: y("Ion;Chwef;Maw;Ebr;Mai;Meh;Gorff;Awst;Medi;Hyd;Tach;Rhag"),
  dddd: y("Dydd Sul;Dydd Llun;Dydd Mawrth;Dydd Mercher;Dydd Iau;Dydd Gwener;Dydd Sadwrn"),
  ddd: y("Sul;Llun;Maw;Mer;Iau;Gwen;Sad")
}, "cy");
B({
  group: ".",
  decimal: ",",
  mmmm: y("yanvar;fevral;mart;aprel;may;iyun;iyul;avqust;sentyabr;oktyabr;noyabr;dekabr"),
  mmm: y("yan;fev;mar;apr;may;iyn;iyl;avq;sen;okt;noy;dek"),
  dddd: y("bazar;bazar ertəsi;çərşənbə axşamı;çərşənbə;cümə axşamı;cümə;şənbə"),
  ddd: y("B.;B.e.;Ç.a.;Ç.;C.a.;C.;Ş.")
}, "az");
B(ve({
  decimal: ",",
  mmmm: y("студзеня;лютага;сакавіка;красавіка;мая;чэрвеня;ліпеня;жніўня;верасня;кастрычніка;лістапада;снежня"),
  dddd: y("нядзеля;панядзелак;аўторак;серада;чацвер;пятніца;субота"),
  ddd: y("нд;пн;аў;ср;чц;пт;сб")
}, 3, -1), "be");
B({
  decimal: ",",
  ampm: y("пр.об.;сл.об."),
  mmmm: y("януари;февруари;март;април;май;юни;юли;август;септември;октомври;ноември;декември"),
  mmm: y("яну;фев;март;апр;май;юни;юли;авг;сеп;окт;ное;дек"),
  dddd: y("неделя;понеделник;вторник;сряда;четвъртък;петък;събота"),
  ddd: y("нд;пн;вт;ср;чт;пт;сб"),
  bool: y("ИСТИНА;ЛОЖЬ")
}, "bg");
B({
  group: ".",
  decimal: ",",
  mmmm: y("de gener;de febrer;de març;d’abril;de maig;de juny;de juliol;d’agost;de setembre;d’octubre;de novembre;de desembre"),
  mmm: y("de gen.;de febr.;de març;d’abr.;de maig;de juny;de jul.;d’ag.;de set.;d’oct.;de nov.;de des."),
  dddd: y("diumenge;dilluns;dimarts;dimecres;dijous;divendres;dissabte"),
  ddd: y("dg.;dl.;dt.;dc.;dj.;dv.;ds."),
  ampm: y("a. m.;p. m.")
}, "ca");
B(ve({
  group: ",",
  decimal: ".",
  mmmm: y("Enero;Pebrero;Marso;Abril;Mayo;Hunyo;Hulyo;Agosto;Setyembre;Oktubre;Nobyembre;Disyembre"),
  dddd: y("Linggo;Lunes;Martes;Miyerkules;Huwebes;Biyernes;Sabado")
}, 3, 3), "fil");
B({
  group: ",",
  decimal: ".",
  mmmm: y("જાન્યુઆરી;ફેબ્રુઆરી;માર્ચ;એપ્રિલ;મે;જૂન;જુલાઈ;ઑગસ્ટ;સપ્ટેમ્બર;ઑક્ટોબર;નવેમ્બર;ડિસેમ્બર"),
  mmm: y("જાન્યુ;ફેબ્રુ;માર્ચ;એપ્રિલ;મે;જૂન;જુલાઈ;ઑગસ્ટ;સપ્ટે;ઑક્ટો;નવે;ડિસે"),
  dddd: y("રવિ~;સોમ~;મંગળ~;બુધ~;ગુરુ~;શુક્ર~;શનિ~", "વાર"),
  ddd: y("રવિ;સોમ;મંગળ;બુધ;ગુરુ;શુક્ર;શનિ")
}, "gu");
B({
  group: ",",
  decimal: ".",
  ampm: y("לפנה״צ;אחה״צ"),
  dddd: y("~ראשון;~שני;~שלישי;~רביעי;~חמישי;~שישי;~שבת", "יום "),
  ddd: y("~א׳;~ב׳;~ג׳;~ד׳;~ה׳;~ו׳;שבת", "יום "),
  mmmm: y("ינואר;פברואר;מרץ;אפריל;מאי;יוני;יולי;אוגוסט;ספטמבר;אוקטובר;נובמבר;דצמבר"),
  mmm: y("ינו׳;פבר׳;מרץ;אפר׳;מאי;יוני;יולי;אוג׳;ספט׳;אוק׳;נוב׳;דצמ׳"),
  mmmm6: y("רמדאן;שוואל;ד׳ו אל־קעדה;ד׳ו אל־חיג׳ה;מוחרם;רביע אל־אוול;רביע א־ת׳אני;ג׳ומאדא אל־אולא;ג׳ומאדא א־ת׳אניה;רג׳ב;שעבאן;רמדאן"),
  mmm6: y("רמדאן;שוואל;ד׳ו אל־קעדה;ד׳ו אל־חיג׳ה;מוחרם;רביע א׳;רביע ב׳;ג׳ומאדא א׳;ג׳ומאדא ב׳;רג׳ב;שעבאן;רמדאן")
}, "he");
B(ve({
  group: ".",
  decimal: ",",
  mmmm: y("siječnja;veljače;ožujka;travnja;svibnja;lipnja;srpnja;kolovoza;rujna;listopada;studenoga;prosinca"),
  mmm: y("sij;velj;ožu;tra;svi;lip;srp;kol;ruj;lis;stu;pro"),
  dddd: y("nedjelja;ponedjeljak;utorak;srijeda;četvrtak;petak;subota")
}, -1, 3), "hr");
B({
  decimal: ",",
  mmmm: y("հունվարի;փետրվարի;մարտի;ապրիլի;մայիսի;հունիսի;հուլիսի;օգոստոսի;սեպտեմբերի;հոկտեմբերի;նոյեմբերի;դեկտեմբերի"),
  mmm: y("հնվ;փտվ;մրտ;ապր;մյս;հնս;հլս;օգս;սեպ;հոկ;նոյ;դեկ"),
  dddd: y("կիրակի;երկուշաբթի;երեքշաբթի;չորեքշաբթի;հինգշաբթի;ուրբաթ;շաբաթ"),
  ddd: y("կիր;երկ;երք;չրք;հնգ;ուր;շբթ")
}, "hy");
B(ve({
  decimal: ",",
  mmmm: y("იანვარი;თებერვალი;მარტი;აპრილი;მაისი;ივნისი;ივლისი;აგვისტო;სექტემბერი;ოქტომბერი;ნოემბერი;დეკემბერი"),
  dddd: y("კვირა;ორშაბათი;სამშაბათი;ოთხშაბათი;ხუთშაბათი;პარასკევი;შაბათი")
}, 3, 3), "ka");
B(ve({
  decimal: ",",
  mmmm: y("қаңтар;ақпан;наурыз;сәуір;мамыр;маусым;шілде;тамыз;қыркүйек;қазан;қараша;желтоқсан"),
  dddd: y("жексенбі;дүйсенбі;сейсенбі;сәрсенбі;бейсенбі;жұма;сенбі"),
  ddd: y("жс;дс;сс;ср;бс;жм;сб")
}, 13, -1), "kk");
B({
  group: ",",
  mmmm: y("ಜನವರಿ;ಫೆಬ್ರವರಿ;ಮಾರ್ಚ್;ಏಪ್ರಿಲ್;ಮೇ;ಜೂನ್;ಜುಲೈ;ಆಗಸ್ಟ್;ಸೆಪ್ಟೆಂಬರ್;ಅಕ್ಟೋಬರ್;ನವೆಂಬರ್;ಡಿಸೆಂಬರ್"),
  mmm: y("ಜನವರಿ;ಫೆಬ್ರವರಿ;ಮಾರ್ಚ್;ಏಪ್ರಿ;ಮೇ;ಜೂನ್;ಜುಲೈ;ಆಗಸ್ಟ್;ಸೆಪ್ಟೆಂ;ಅಕ್ಟೋ;ನವೆಂ;ಡಿಸೆಂ"),
  dddd: y("ಭಾನು~;ಸೋಮ~;ಮಂಗಳ~;ಬುಧ~;ಗುರು~;ಶುಕ್ರ~;ಶನಿ~", "ವಾರ"),
  ddd: y("ಭಾನು;ಸೋಮ;ಮಂಗಳ;ಬುಧ;ಗುರು;ಶುಕ್ರ;ಶನಿ"),
  ampm: y("ಪೂರ್ವಾಹ್ನ;ಅಪರಾಹ್ನ")
}, "kn");
B({
  decimal: ",",
  mmmm: y("sausio;vasario;kovo;balandžio;gegužės;birželio;liepos;rugpjūčio;rugsėjo;spalio;lapkričio;gruodžio"),
  mmm: y("saus.;vas.;kov.;bal.;geg.;birž.;liep.;rugp.;rugs.;spal.;lapkr.;gruod."),
  dddd: y("sekmadienis;pirmadienis;antradienis;trečiadienis;ketvirtadienis;penktadienis;šeštadienis"),
  ddd: y("sk;pr;an;tr;kt;pn;št"),
  ampm: y("priešpiet;popiet")
}, "lt");
B({
  decimal: ",",
  mmmm: y("janvāris;februāris;marts;aprīlis;maijs;jūnijs;jūlijs;augusts;septembris;oktobris;novembris;decembris"),
  mmm: y("janv.;febr.;marts;apr.;maijs;jūn.;jūl.;aug.;sept.;okt.;nov.;dec."),
  dddd: y("svētdiena;pirmdiena;otrdiena;trešdiena;ceturtdiena;piektdiena;sestdiena"),
  ddd: y("svētd.;pirmd.;otrd.;trešd.;ceturtd.;piektd.;sestd."),
  ampm: y("priekšpusdienā;pēcpusdienā")
}, "lv");
B({
  group: ",",
  decimal: ".",
  mmmm: y("ജനുവരി;ഫെബ്രുവരി;മാർച്ച്;ഏപ്രിൽ;മേയ്;ജൂൺ;ജൂലൈ;ഓഗസ്റ്റ്;സെപ്റ്റംബർ;ഒക്‌ടോബർ;നവംബർ;ഡിസംബർ"),
  mmm: y("ജനു;ഫെബ്രു;മാർ;ഏപ്രി;മേയ്;ജൂൺ;ജൂലൈ;ഓഗ;സെപ്റ്റം;ഒക്ടോ;നവം;ഡിസം"),
  dddd: y("ഞായറാഴ്‌ച;തിങ്കളാഴ്‌ച;ചൊവ്വാഴ്ച;ബുധനാഴ്‌ച;വ്യാഴാഴ്‌ച;വെള്ളിയാഴ്‌ച;ശനിയാഴ്‌ച"),
  ddd: y("ഞായർ;തിങ്കൾ;ചൊവ്വ;ബുധൻ;വ്യാഴം;വെള്ളി;ശനി")
}, "ml");
B({
  group: ",",
  decimal: ".",
  mmmm: y("нэгдүгээ~;хоёрдугаа~;гуравдугаа~;дөрөвдүгээ~;тавдугаа~;зургаадугаа~;долоодугаа~;наймдугаа~;есдүгээ~;аравдугаа~;арван нэгдүгээ~;арван хоёрдугаа~", "р сар"),
  mmm: y("1~;2~;3~;4~;5~;6~;7~;8~;9~;10~;11~;12~", "-р сар"),
  dddd: y("ням;даваа;мягмар;лхагва;пүрэв;баасан;бямба"),
  ddd: y("Ня;Да;Мя;Лх;Пү;Ба;Бя"),
  ampm: y("ү.ө.;ү.х.")
}, "mn");
B({
  group: ",",
  decimal: ".",
  mmmm: y("जानेवारी;फेब्रुवारी;मार्च;एप्रिल;मे;जून;जुलै;ऑगस्ट;सप्टेंबर;ऑक्टोबर;नोव्हेंबर;डिसेंबर"),
  mmm: y("जाने;फेब्रु;मार्च;एप्रि;मे;जून;जुलै;ऑग;सप्टें;ऑक्टो;नोव्हें;डिसें"),
  dddd: y("रविवार;सोमवार;मंगळवार;बुधवार;गुरुवार;शुक्रवार;शनिवार"),
  ddd: y("रवि;सोम;मंगळ;बुध;गुरु;शुक्र;शनि")
}, "mr");
B(ve({
  group: ",",
  decimal: ".",
  mmmm: y("ဇန်နဝါရီ;ဖေဖော်ဝါရီ;မတ်;ဧပြီ;မေ;ဇွန်;ဇူလိုင်;ဩဂုတ်;စက်တင်ဘာ;အောက်တိုဘာ;နိုဝင်ဘာ;ဒီဇင်ဘာ"),
  mmm: y("ဇန်;ဖေ;မတ်;ဧ;မေ;ဇွန်;ဇူ;ဩ;စက်;အောက်;နို;ဒီ"),
  dddd: y("တနင်္ဂနွေ;တနင်္လာ;အင်္ဂါ;ဗုဒ္ဓဟူး;ကြာသပတေး;သောကြာ;စနေ"),
  ampm: y("နံနက်;ညနေ")
}, -1, 0), "my");
B({
  group: ",",
  decimal: ".",
  mmmm: y("ਜਨਵਰੀ;ਫ਼ਰਵਰੀ;ਮਾਰਚ;ਅਪ੍ਰੈਲ;ਮਈ;ਜੂਨ;ਜੁਲਾਈ;ਅਗਸਤ;ਸਤੰਬਰ;ਅਕਤੂਬਰ;ਨਵੰਬਰ;ਦਸੰਬਰ"),
  mmm: y("ਜਨ;ਫ਼ਰ;ਮਾਰਚ;ਅਪ੍ਰੈ;ਮਈ;ਜੂਨ;ਜੁਲਾ;ਅਗ;ਸਤੰ;ਅਕਤੂ;ਨਵੰ;ਦਸੰ"),
  dddd: y("ਐਤਵਾਰ;ਸੋਮਵਾਰ;ਮੰਗਲਵਾਰ;ਬੁੱਧਵਾਰ;ਵੀਰਵਾਰ;ਸ਼ੁੱਕਰਵਾਰ;ਸ਼ਨਿੱਚਰਵਾਰ"),
  ddd: y("ਐਤ;ਸੋਮ;ਮੰਗਲ;ਬੁੱਧ;ਵੀਰ;ਸ਼ੁੱਕਰ;ਸ਼ਨਿੱਚਰ"),
  ampm: y("ਪੂ.ਦੁ.;ਬਾ.ਦੁ.")
}, "pa");
B({
  group: ".",
  decimal: ",",
  mmmm: y("ianuarie;februarie;martie;aprilie;mai;iunie;iulie;august;septem~;octom~;noiem~;decem~", "brie"),
  mmm: y("ian.;feb.;mar.;apr.;mai;iun.;iul.;aug.;sept.;oct.;nov.;dec."),
  dddd: y("duminică;luni;marți;miercuri;joi;vineri;sâmbătă"),
  ddd: y("dum.;lun.;mar.;mie.;joi;vin.;sâm."),
  ampm: y("a.m.;p.m.")
}, "ro");
B(ve({
  group: ".",
  decimal: ",",
  mmmm: y("januar;februar;marec;april;maj;junij;julij;avgust;september;oktober;november;december"),
  mmm: y("jan.;feb.;mar.;apr.;maj;jun.;jul.;avg.;sep.;okt.;nov.;dec."),
  dddd: y("nedelja;ponedeljek;torek;sreda;četrtek;petek;sobota"),
  ampm: y("dop.;pop.")
}, -1, 13), "sl");
B(ve({
  group: ".",
  decimal: ",",
  mmmm: y("јануар;фебруар;март;април;мај;јун;јул;август;септембар;октобар;новембар;децембар"),
  dddd: y("недеља;понедељак;уторак;среда;четвртак;петак;субота")
}, 3, 3), "sr");
B({
  group: ",",
  decimal: ".",
  mmmm: y("ஜனவரி;பிப்ரவரி;மார்ச்;ஏப்ரல்;மே;ஜூன்;ஜூலை;ஆகஸ்ட்;செப்டம்பர்;அக்டோபர்;நவம்பர்;டிசம்பர்"),
  mmm: y("ஜன.;பிப்.;மார்.;ஏப்.;மே;ஜூன்;ஜூலை;ஆக.;செப்.;அக்.;நவ.;டிச."),
  dddd: y("ஞாயிறு;திங்கள்;செவ்வாய்;புதன்;வியாழன்;வெள்ளி;சனி"),
  ddd: y("ஞாயி.;திங்.;செவ்.;புத.;வியா.;வெள்.;சனி")
}, "ta");
B({
  group: ",",
  decimal: ".",
  mmmm: y("జనవరి;ఫిబ్రవరి;మార్చి;ఏప్రిల్;మే;జూన్;జులై;ఆగస్టు;సెప్టెంబర్;అక్టోబర్;నవంబర్;డిసెంబర్"),
  mmm: y("జన;ఫిబ్ర;మార్చి;ఏప్రి;మే;జూన్;జులై;ఆగ;సెప్టెం;అక్టో;నవం;డిసెం"),
  dddd: y("ఆదివారం;సోమవారం;మంగళవారం;బుధవారం;గురువారం;శుక్రవారం;శనివారం"),
  ddd: y("ఆది;సోమ;మంగళ;బుధ;గురు;శుక్ర;శని")
}, "te");
B({
  decimal: ",",
  mmmm: y("січня;лютого;березня;квітня;травня;червня;липня;серпня;вересня;жовтня;листопада;грудня"),
  mmm: y("січ.;лют.;бер.;квіт.;трав.;черв.;лип.;серп.;вер.;жовт.;лист.;груд."),
  dddd: y("неділю;понеділок;вівторок;середу;четвер;пʼятницю;суботу"),
  ddd: y("нд;пн;вт;ср;чт;пт;сб"),
  ampm: y("дп;пп")
}, "uk");
B({
  group: ".",
  decimal: ",",
  mmmm: y("~1;~2;~3;~4;~5;~6;~7;~8;~9;~10;~11;~12", "tháng "),
  mmm: y("~1;~2;~3;~4;~5;~6;~7;~8;~9;~10;~11;~12", "thg "),
  dddd: y("Chủ Nhật;Thứ Hai;Thứ Ba;Thứ Tư;Thứ Năm;Thứ Sáu;Thứ Bảy"),
  ddd: y("CN;Th 2;Th 3;Th 4;Th 5;Th 6;Th 7"),
  ampm: y("SA;CH")
}, "vi");
B(ve({
  group: "٬",
  decimal: "٫",
  ampm: y("ص;م"),
  mmmm: y("يناير;فبراير;مارس;أبريل;مايو;يونيو;يوليو;أغسطس;سبتمبر;أكتوبر;نوفمبر;ديسمبر"),
  dddd: y("الأحد;الاثنين;الثلاثاء;الأربعاء;الخميس;الجمعة;السبت"),
  mmmm6: y("رمضان;شوال;ذو القعدة;ذو الحجة;محرم;ربيع الأول;ربيع الآخرة;جمادى الأولى;جمادى الآخرة;رجب;شعبان;رمضان")
}, 0, 0), "ar");
B({
  group: ",",
  decimal: ".",
  mmmm: y("জানুয়ারী;ফেব্রুয়ারী;মার্চ;এপ্রিল;মে;জুন;জুলাই;আগস্ট;সেপ্টেম্বর;অক্টোবর;নভেম্বর;ডিসেম্বর"),
  mmm: y("জানু;ফেব;মার্চ;এপ্রি;মে;জুন;জুল;আগ;সেপ্টেঃ;অক্টোঃ;নভেঃ;ডিসেঃ"),
  dddd: y("রবিবার;সোমবার;মঙ্গলবার;বুধবার;বৃহস্পতিবার;শুক্রবার;শনিবার"),
  ddd: y("রবি;সোম;মঙ্গল;বুধ;বৃহস্পতি;শুক্র;শনি")
}, "bn");
B({
  group: ",",
  decimal: ".",
  mmmm: y("जनवरी;फ़रवरी;मार्च;अप्रैल;मई;जून;जुलाई;अगस्त;सितंबर;अक्तूबर;नवंबर;दिसंबर"),
  mmm: y("जन॰;फ़र॰;मार्च;अप्रैल;मई;जून;जुल॰;अग॰;सित॰;अक्तू॰;नव॰;दिस॰"),
  dddd: y("रविवार;सोमवार;मंगलवार;बुधवार;गुरुवार;शुक्रवार;शनिवार"),
  ddd: y("रवि;सोम;मंगल;बुध;गुरु;शुक्र;शनि"),
  ampm: y("am;pm")
}, "hi");
const K2 = {
  // Overflow error string
  overflow: "######",
  // dateErrorThrow needs to be off! [prev in locale]
  // Should it throw when there is an overflow error?
  dateErrorThrows: !1,
  // Should it emit a number when date has an overflow error? (Sheets does this)
  dateErrorNumber: !0,
  // dateErrorThrow needs to be off!
  // Should it emit a number when bigint has an is an overflow error?
  bigintErrorNumber: !1,
  // Sheets mode (see #3)
  dateSpanLarge: !0,
  // Simulate the Lotus 1-2-3 leap year bug
  leap1900: !0,
  // Emit regular vs. non-breaking spaces
  nbsp: !1,
  // Robust/throw mode
  throws: !0,
  // What is emitted when robust mode fails to parse (###### currently)
  invalid: "######",
  // Locale
  locale: "",
  // Don't adjust dates to UTC when converting them to serial time
  ignoreTimezone: !1,
  // Integer digit grouping
  grouping: [3, 3],
  // resolve indexed colors to hex
  indexColors: !0,
  // Skip-next signifier character
  skipChar: "",
  // Repear-next signifier character
  repeatChar: ""
};
function Et(e, t = 0) {
  if (typeof e != "number")
    return e;
  if (e < 0)
    return -Et(-e, t);
  if (t) {
    const n = 10 ** t || 1;
    return Et(e * n, 0) / n;
  }
  return Math.round(e);
}
const J2 = 1e-13;
function Q2(e, t = 2, n = 2) {
  const r = e < 0 ? -1 : 1, o = 10 ** (t || 2), s = 10 ** (n || 2);
  let l = Math.abs(e), f = 0, d = 0, c = 0, i = 1, a, u;
  if (e = l, e % 1 === 0)
    u = [e * r, 1];
  else if (e < 1e-19)
    u = [r, 1e19];
  else if (e > 1e19)
    u = [1e19 * r, 1];
  else {
    do
      if (l = 1 / (l - Math.floor(l)), a = i, i = i * Math.floor(l) + f, f = a, d = c, c = Math.floor(e * i + 0.5), c >= o || i >= s)
        return [r * d, f];
    while (Math.abs(e - c / i) >= J2 && l !== Math.floor(l));
    u = [r * c, i];
  }
  return u;
}
const it = Math.floor;
function Rc(e, t = !0) {
  if (t && e >= 0) {
    if (e === 0)
      return [1900, 1, 0];
    if (e === 60)
      return [1900, 2, 29];
    if (e < 60)
      return [1900, e < 32 ? 1 : 2, (e - 1) % 31 + 1];
  }
  let n = e + 68569 + 2415019;
  const r = it(4 * n / 146097);
  n = n - it((146097 * r + 3) / 4);
  const o = it(4e3 * (n + 1) / 1461001);
  n = n - it(1461 * o / 4) + 31;
  const s = it(80 * n / 2447), l = n - it(2447 * s / 80);
  n = it(s / 11);
  const f = s + 2 - 12 * n;
  return [100 * (r - 49) + o + n | 0, f | 0, l | 0];
}
function j2(e) {
  return Rc(e + 1462);
}
function Z2(e) {
  if (e === 60)
    throw new Error("#VALUE!");
  if (e <= 1)
    return [1317, 8, 29];
  if (e < 60)
    return [1317, e < 32 ? 9 : 10, 1 + (e - 2) % 30];
  const t = 10631 / 30, n = 8.01 / 60;
  let r = e + 466935;
  const o = it(r / 10631);
  r = r - 10631 * o;
  const s = it((r - n) / t);
  r = r - it(s * t + n);
  const l = it((r + 28.5001) / 29.5);
  return l === 13 ? [30 * o + s, 12, 30] : [30 * o + s, l, r - it(29.5001 * l - 29)];
}
function ev(e, t = 0, n = !0) {
  const r = it(e);
  return t === Ln ? Z2(r) : t === $2 ? j2(r) : Rc(r, n);
}
function Ac(e, t) {
  let n = null;
  if (Array.isArray(e)) {
    const [r, o, s, l, f, d] = e;
    n = Date.UTC(r, o == null ? 0 : o - 1, s ?? 1, l || 0, f || 0, d || 0);
  } else if (e instanceof Date && (n = e * 1, !t?.ignoreTimezone)) {
    const r = /* @__PURE__ */ new Date();
    r.setUTCFullYear(
      e.getFullYear(),
      e.getMonth(),
      e.getDate()
    ), r.setUTCHours(
      e.getHours(),
      e.getMinutes(),
      e.getSeconds(),
      e.getMilliseconds()
    ), n = r * 1;
  }
  if (n != null && isFinite(n)) {
    const r = n / 864e5;
    return r - (r <= -25509 ? -25568 : -25569);
  }
  return null;
}
const tv = [
  // day-month-year
  "!d-m-y",
  "!d-m-Y",
  "!j-m-y",
  "!j-m-Y",
  "!d-n-y",
  "!d-n-Y",
  "!j-n-y",
  "!j-n-Y",
  // month-day-year
  "?m-d-y",
  "?m-d-Y",
  "?m-j-y",
  "?m-j-Y",
  "?n-d-y",
  "?n-d-Y",
  "?n-j-y",
  "?n-j-Y",
  // unab
  "d-M-y",
  "d-M-Y",
  "j-M-y",
  "j-M-Y",
  "M-d-y",
  "M-d-Y",
  "M-j-y",
  "M-j-Y",
  "d-F-y",
  "d-F-Y",
  "F-d-y",
  "F-d-Y",
  "F-j-y",
  "F-j-Y",
  "j-F-y",
  "j-F-Y",
  "y-F-d",
  "y-F-j",
  "y-M-d",
  "y-M-j",
  "Y-F-d",
  "Y-F-j",
  "Y-M-d",
  "Y-m-d",
  "Y-M-j",
  "Y-m-j",
  "Y-n-d",
  "Y-n-j",
  "j-F",
  // 2-April
  "j-M",
  // 2-Apr
  "d-F",
  // 02-April
  "d-M",
  // 02-Apr
  "n-d",
  // 4-02
  "n-j",
  // 4-2
  "n-Y",
  // 4-1908
  "m-d",
  // 04-02
  "m-j",
  // 04-2
  "m-Y",
  // 04-1908
  "M-Y",
  // Apr-1908
  "M-y",
  // Apr-08
  "F-y",
  // April-08
  "F-Y",
  // April-1908
  "Y-M",
  // 1908-Apr
  "Y-n",
  // 1908-4
  "Y-m",
  // 1908-04
  "Y-F",
  // 1908-April
  "Y-M"
  // 1908-Apr
], nv = {}, rv = {};
function rt(e, t, n = 1) {
  if (e) {
    const r = e[0], o = e.slice(1);
    r === "!" ? rt(o, t, 4) : r === "?" ? rt(o, t, 2) : (t[r] = t[r] || {}, rt(o, t[r], n));
  } else
    t.$ = n;
}
function $s(e, t) {
  rt(e, t), rt(e + " x", t), rt(e + " l", t), rt(e + " l x", t), rt("l " + e, t), rt("l " + e + " x", t), rt(e + " D", t), rt(e + " D x", t), rt("D " + e, t), rt("D " + e + " x", t);
}
tv.forEach((e) => {
  e.startsWith("?") || $s(e, nv), e.startsWith("!") || $s(e, rv);
});
(/* @__PURE__ */ new Date()).getUTCFullYear();
function Zt(e, t, n) {
  return t[n + "_max"] = e.length, t[n + "_min"] = e.replace(/#/g, "").length, t;
}
function Ie(e, t) {
  typeof e == "string" ? t.push({ type: "string", value: e }) : t.push(e);
}
function qs(e, t) {
  const n = e?.type;
  return n === Ql || n === er || n === jl || n === os && t === "den";
}
function Jn(e) {
  const t = [], n = {
    scale: 1,
    percent: !1,
    text: !1,
    date: 0,
    date_eval: !1,
    date_system: Gs,
    sec_decimals: 0,
    general: !1,
    clock: 24,
    int_pattern: [],
    frac_pattern: [],
    man_pattern: [],
    den_pattern: [],
    num_pattern: [],
    tokens: t
  };
  let r = "int", o = null;
  const s = [];
  let l, f = !1, d = -1, c = !1, i = "", a = !1;
  for (; ++d < e.length && !c; ) {
    const N = e[d], v = N.type || to;
    if (i += N.raw, v === Jl)
      n.general = !0, Ie(N, t);
    else if (qs(N, r)) {
      const C = n[r + "_pattern"];
      qs(l, r) || l?.type === eo ? (C.push((C.pop() || "") + N.value), o.num += N.value) : (C.push(N.value), o = { type: r, num: N.value }, Ie(o, t));
    } else if (v === Ec)
      N.value === "(" && (n.parens = !0), Ie(N.value, t);
    else if (v === os)
      Ie(N.value, t);
    else if (v === Zl)
      if (a = !0, n[r + "_pattern"].length) {
        if (!o)
          throw new SyntaxError("Format pattern is missing a numerator");
        n.fractions = !0, n.num_pattern.push(n[r + "_pattern"].pop()), o.type = "num", r = "den", Ie({ type: "div" }, t);
      } else
        Ie(N.value, t);
    else if (v === ec)
      Ie(",", t);
    else if (v === di)
      n.scale = 1e-3 ** N.raw.length;
    else if (v === eo) {
      if (r === "int" && (n.grouping = !0), r === "den")
        throw new SyntaxError("Cannot group denominator digits");
    } else if (v === oc)
      Ie(N, t);
    else if (v === tc) {
      c = !0;
      break;
    } else if (v === rs)
      n.text = !0, Ie(N, t);
    else if (v === nc || v === rc)
      Ie(N, t);
    else if (v === lc) {
      const C = N.value.toLowerCase(), M = C[0], R = { type: "", size: 0, date: 1, pad: C.length };
      M === "h" ? (R.size = xr, R.type = "hour-elap") : M === "m" ? (R.size = Dr, R.type = "min-elap") : (R.size = Or, R.type = "sec-elap"), n.date = n.date | R.size, s.push(R), Ie(R, t);
    } else if (n.date && v === hi && e[d + 1]?.type === er) {
      let C = 1;
      d++;
      let M = "0";
      e[d + 1]?.type === er && (M += "0", C = 2, d++), e[d + 1]?.type === er && (M += "0", C = 3, d++), i += M;
      const R = [Or, ci, ui, fi][C];
      n.date = n.date | R, n.date_eval = !0, n.sec_decimals = Math.max(n.sec_decimals, C), Ie({
        type: "subsec",
        size: R,
        decimals: C,
        date: 1
      }, t);
    } else if (v === sc)
      f || (N.value === "B2" || N.value === "b2" ? n.date_system = Ln : n.date_system = Gs);
    else if (v === ac) {
      const C = { type: "", size: 0, date: 1 }, M = N.value.toLowerCase(), R = M[0];
      if (M === "y" || M === "yy")
        C.size = Ar, C.type = "year-short";
      else if (R === "y" || R === "e")
        C.size = Ar, C.type = "year";
      else if (M === "b" || M === "bb")
        C.size = Ar, C.type = "b-year-short";
      else if (R === "b")
        C.size = Ar, C.type = "b-year";
      else if (M === "d" || M === "dd")
        C.size = Po, C.type = "day", C.pad = /dd/.test(M);
      else if (M === "ddd" || M === "aaa")
        C.size = Po, C.type = "weekday-short";
      else if (R === "d" || R === "a")
        C.size = Po, C.type = "weekday";
      else if (R === "h")
        C.size = xr, C.type = "hour", C.pad = /hh/i.test(M);
      else if (R === "m") {
        M.length === 3 ? (C.size = Tr, C.type = "monthname-short") : M.length === 5 ? (C.size = Tr, C.type = "monthname-single") : M.length >= 4 && (C.size = Tr, C.type = "monthname");
        const A = s[s.length - 1];
        !C.type && A && !A.used && A.size & (xr | Or) && (A.used = !0, C.size = Dr, C.type = "min", C.pad = /mm/.test(M)), C.type || (C.size = Tr, C.type = "month", C.pad = /mm/.test(M), C.indeterminate = !0);
      } else if (R === "s") {
        C.size = Or, C.type = "sec", C.pad = /ss/.test(M);
        const A = s[s.length - 1];
        A && A.size & Dr ? C.used = !0 : A?.indeterminate && (delete A.indeterminate, A.size = Dr, A.type = "min", C.used = !0);
      }
      n.date = n.date | C.size, n.date_eval = !0, s.push(C), Ie(C, t);
    } else if (v === mc)
      n.clock = 12, n.date = n.date | xr, n.date_eval = !0, N.short = N.value === "A/P", Ie(N, t);
    else if (v === pc || v === gc || v === is)
      Ie(N.value, t);
    else if (v === cc)
      n.condition = [
        N.value[0],
        // operator
        parseFloat(N.value[1])
        // operand
      ];
    else if (v === dc) {
      const C = N.value.split("-"), M = C.length < 2 ? "" : C.slice(1).join("-"), R = C[0];
      R && Ie(R, t);
      const A = ss(M);
      A && (n.locale = A);
      const x = parseInt(M, 16);
      isFinite(x) && x & 16711680 && (x >> 16 & 255) === 6 && (n.date_system = Ln), f = !0;
    } else if (v === hc) {
      let C, M = N.value.toLowerCase();
      (C = /^color\s*(\d+)$/i.exec(M)) && (M = parseInt(C[1], 10)), n.color = M;
    } else if (v === ic)
      n.scale = 100, n.percent = !0, Ie("%", t);
    else if (v === hi)
      Ie(N, t), n.date || (n.dec_fractions = !0, r = "frac");
    else if (v === Nc)
      n.exponential = !0, n.exp_plus = N.value.includes("+"), r = "man", Ie({ type: "exp", plus: n.exp_plus }, t);
    else if (v === vc)
      Ie(N, t);
    else if (v === Cc)
      Ie(N, t);
    else if (!(v === uc || v === fc)) throw v === to ? new SyntaxError(`Illegal character: ${i}`) : new SyntaxError(`Unknown token ${v} in ${i}`);
    l = N;
  }
  if (n.tokensUsed = d, n.pattern = i, /^((?:\[[^\]]+\])+)(;|$)/.test(n.pattern) && !/^\[(?:h+|m+|s+)\]/.test(n.pattern) && Ie({ type: "text" }, t), n.fractions && n.dec_fractions || n.grouping && !n.int_pattern.length || n.fractions && n.exponential || n.fractions && n.den_pattern.length * n.num_pattern.length === 0 || a && !n.fractions && !n.date || n.exponential && (n.int_pattern.length || n.frac_pattern.length) * n.man_pattern.length === 0)
    throw new SyntaxError(`Invalid pattern: ${i}`);
  const u = n.int_pattern.join(""), h = n.man_pattern.join(""), g = n.frac_pattern.join("");
  Zt(u, n, "int");
  let m = 0;
  for (let N = 0; N < u.length; N++) {
    const v = u[u.length - 1 - N];
    /^[0-9?]/.test(v) && (m = N + 1);
  }
  n.int_min = m, Zt(g, n, "frac"), Zt(h, n, "man");
  let p = n.num_pattern.join(""), E = n.den_pattern[0] || "";
  if (E.includes("?") || p.includes("?") ? (E = E.replace(/\d/g, "?"), E = E.replace(/#$/g, "?"), Zt(p, n, "num"), Zt(E, n, "den"), p = p.replace(/#$/g, "?")) : (Zt(p, n, "num"), Zt(E, n, "den")), n.int_p = u, n.man_p = h, n.num_p = p, n.den_p = E, n.den_pattern.length && (n.denominator = parseInt(n.den_pattern.join("").replace(/\D/g, ""), 10)), n.integer = !!u.length, !n.integer && !n.exponential && g.length) {
    const N = n.tokens.findIndex((v) => v.type === "point");
    n.tokens.splice(N, 0, { type: "int", value: "#" }), n.integer = !0, n.int_pattern = ["#"], n.int_p = "#";
  }
  if (n.fractions)
    for (let N = 0; N < t.length - 1; N++) {
      const v = t[N];
      if (v.type !== "string" && v.type !== "space")
        continue;
      const C = t[N + 1].type;
      C === "num" ? v.rule = "num+int" : C === "div" ? v.rule = "num" : C === "den" && (v.rule = "den");
    }
  return n.grouping && n.int_pattern.length > 1 && (n.grouping = !1), n;
}
function ov(e) {
  if (e === 0)
    return e;
  const t = Math.ceil(Math.log10(e < 0 ? -e : e)), n = 10 ** (16 - Math.floor(t));
  return isFinite(n) ? Math.round(e * n) / n : 0;
}
function Tc(e, t = 0) {
  const n = Math.floor(Math.log10(e));
  return t > 1 ? Math.floor(n / t) * t : n;
}
function xc(e, t = 1) {
  return t < -300 ? parseFloat(e.toExponential().split("e")[0]) : e * 10 ** -t;
}
const iv = {
  total: 1,
  sign: 0,
  period: 0,
  int: 1,
  frac: 0
};
function sv(e, t = !0) {
  const n = Math.abs(e);
  if (!n)
    return iv;
  const r = t && e < 0 ? 1 : 0, o = Math.floor(n), s = Math.floor(Math.log10(n) + 1);
  let l = 0, f = 0;
  if (o !== n) {
    l = 1;
    const d = String(
      Et(n * 10 ** -s, 15)
    );
    let c = d.length, i = !0, a = 0;
    for (; a <= d.length; ) {
      if (d[a] === ".") {
        c--;
        break;
      } else d[a] === "0" && i ? c-- : i = !1;
      a++;
    }
    f = c - s, f < 0 && (f = 0, l = 0);
  }
  return {
    total: r + Math.max(s, 1) + l + f,
    digits: Math.max(s, 0) + f,
    sign: r,
    period: l,
    int: Math.max(s, 1),
    frac: f
  };
}
const Sn = (e, t) => e.replace(/\./, t.decimal), Ws = (e, t, n) => {
  const r = Math.abs(t);
  let o;
  return e === 1 ? o = e : o = Et(e, 5), [
    Sn(o + "", n),
    n.exponent,
    t < 0 ? n.negative : n.positive,
    r < 10 ? "0" : "",
    r
  ];
};
function Ys(e, t, n, r) {
  const o = n | 0;
  if (typeof n == "string")
    e.push(n);
  else if (n === o)
    e.push(Math.abs(o));
  else {
    const s = Math.abs(n);
    let l = Tc(s), f = xc(s, l);
    f === 10 && (f = 1, l++);
    const d = sv(s);
    if (l >= -4 && l <= -1) {
      const c = s.toPrecision(10 + l).replace(/\.?0+$/, "");
      e.push(Sn(c, r));
    } else if (l === 10) {
      const c = s.toFixed(10).slice(0, 12).replace(/\.$/, "");
      e.push(Sn(c, r));
    } else if (Math.abs(l) <= 9)
      if (d.total <= 11) {
        const c = Et(s, 9).toFixed(d.frac);
        e.push(Sn(c, r));
      } else l === 9 ? e.push(Math.floor(s)) : l >= 0 && l < 9 ? e.push(Sn(String(Et(s, 9 - l)), r)) : e.push(...Ws(f, l, r));
    else d.total >= 12 ? e.push(...Ws(f, l, r)) : e.push(Sn(Et(s, 9).toFixed(d.frac), r));
  }
  return e;
}
function en(e, t = !1) {
  return e === "0" ? "0" : e === "?" ? t ? " " : " " : "";
}
const Kn = 86400, av = (e, t, n) => n ? e < V2 || t >= G2 : e < z2 || t >= H2;
function Xs(e, t, n, r) {
  let o = "", s = "", l = "", f = "", d = "", c = "", i = 0, a = 0;
  if (typeof e == "bigint") {
    if (e <= Number.MAX_SAFE_INTEGER && e >= Number.MIN_SAFE_INTEGER)
      e = Number(e);
    else
      return n.bigintErrorNumber ? String(e) : n.overflow;
    a = e;
  } else
    a = Math.trunc(e);
  let u = 0, h = 0, g = 1, m = 0, p = 0, E = 0, b = 0, N = 0, v = 0;
  const C = r || no;
  if (!t.text && isFinite(t.scale) && t.scale !== 1 && (e = ov(e * t.scale)), t.exponential) {
    let K = Math.abs(e);
    K && (i = Tc(K, t.int_max)), e && !t.integer && i++, K = xc(K, i), t.int_max === 1 && Et(K, t.frac_max) === 10 && (K = 1, i++), e = e < 0 ? -K : K, o += Math.abs(i);
  }
  if (t.integer) {
    const K = Math.abs(Et(e, t.fractions ? 1 : t.frac_max));
    c += K < 1 ? "" : Math.floor(K);
  }
  const M = n.grouping[0] ?? 3, R = n.grouping[1] ?? M;
  t.dec_fractions && (d = String(Et(e, t.frac_max)).split(".")[1] || "");
  const A = !t.error && (t.num_p.includes("0") || t.den_p.includes("0"));
  let x = A;
  if (t.fractions) {
    x = A || !!(e % 1);
    const K = Math.abs(t.integer ? e % 1 : e);
    if (K)
      if (x = !0, t.denominator && isFinite(t.denominator))
        f += t.denominator, l += Et(K * t.denominator), l === "0" && (l = "", f = "", x = A);
      else {
        const q = Q2(K, 1 / 0, t.den_max);
        l += q[0], f += q[1], t.integer && l === "0" && (l = "", f = "", x = A);
      }
    else !e && !t.integer && (x = !0, l = "0", f = "1");
    t.integer && !x && !Math.trunc(e) && (c = "0");
  }
  if (t.date) {
    a = Math.trunc(e);
    const K = Kn * (e - a);
    if (u = Math.floor(K), v = K - u, Math.abs(v) < 1e-6 ? v = 0 : v > 0.9999 && (v = 0, u += 1, u === Kn && (u = 0, a += 1)), v) {
      const q = t.date & fi || t.date & ui || t.date & ci;
      (q === fi && v > 0.9995 || q === ui && v > 0.995 || q === ci && v > 0.95 || !q && v >= 0.5) && (u++, v = 0);
    }
    if (a || t.date_system) {
      const q = ev(e, t.date_system, n.leap1900);
      h = q[0], g = q[1], m = q[2];
    }
    if (u) {
      const q = u < 0 ? Kn + u : u;
      N = Math.floor(q) % 60, b = Math.floor(q / 60) % 60, E = Math.floor(q / 60 / 60) % 60;
    }
    if (p = (6 + a) % 7, t.date_eval && av(e, a + u / Kn, n.dateSpanLarge)) {
      if (n.dateErrorThrows)
        throw new Error("Date out of bounds");
      if (n.dateErrorNumber) {
        const q = e < 0 ? [C.negative] : [];
        return Ys(q, {}, e, C).join("");
      }
      return n.overflow;
    }
  }
  const L = en("?", n.nbsp);
  i < 0 ? s = "-" : t.exp_plus && (s = "+");
  const D = [], W = (K, q, H, Y) => {
    const te = !Y && K.length > q.length ? H.length + K.length - q.length : H.length;
    K.length < q.length && (Y += K.length - q.length);
    for (let X = 0; X < te; X++)
      D.push(K[X + Y] || en(H[X], n.nbsp));
    return te;
  };
  let J = !1;
  const ne = { int: 0, frac: 0, man: 0, num: 0, den: 0 };
  for (let K = 0, q = t.tokens.length; K < q; K++) {
    const H = t.tokens[K], Y = H.type, te = H.num ? H.num.length : 0;
    if (Y === "string")
      H.rule ? H.rule === "num" ? x ? D.push(H.value.replace(/ /g, L)) : (t.num_min > 0 || t.den_min > 0) && D.push(H.value.replace(/./g, L)) : H.rule === "num+int" ? x && c ? D.push(H.value.replace(/ /g, L)) : t.den_min > 0 && (c || t.num_min) && D.push(H.value.replace(/./g, L)) : H.rule === "den" && (x ? D.push(H.value.replace(/ /g, L)) : (t.den_min > 0 || t.den_min > 0) && D.push(H.value.replace(/./g, L))) : D.push(H.value.replace(/ /g, L));
    else if (Y === "space")
      H.rule === "num+int" ? (x || t.num_min || t.den_min) && (c || t.num_min) && D.push(L) : D.push(L);
    else if (Y === "error")
      D.push(n.invalid);
    else if (Y === "point")
      D.push(t.date ? H.value : C.decimal);
    else if (Y === "general")
      Ys(D, t, e, C);
    else if (Y === "exp")
      D.push(C.exponent);
    else if (Y === "minus")
      H.volatile && t.date || H.volatile && (e >= 0 || typeof e != "number") || (H.volatile && !t.fractions && (t.integer || t.dec_fractions) ? (e < 0 && c && c !== "0" || d) && D.push(C.negative) : D.push(C.negative));
    else if (Y === "plus")
      D.push(C.positive);
    else if (Y === "text")
      D.push(e);
    else if (Y === "fill")
      n.fillChar && D.push(n.fillChar, H.value);
    else if (Y === "skip")
      n.skipChar ? D.push(n.skipChar, H.value) : D.push(n.nbsp ? " " : " ");
    else if (Y === "div")
      x ? D.push("/") : t.num_min > 0 || t.den_min > 0 ? D.push(L) : D.push(en("#", n.nbsp));
    else if (Y === "int")
      if (t.int_pattern.length === 1) {
        const X = t.int_p, le = Math.max(t.int_min, c.length);
        let $ = "";
        for (let j = le; j > 0; j--) {
          const pe = c.charAt(c.length - j), St = pe ? "" : X.charAt(X.length - j) || X[0];
          let De = "";
          if (t.grouping) {
            const vt = j - 1 - M;
            vt >= 0 && !(vt % R) && (De = pe || St === "0" ? C.group : en("?", n.nbsp));
          }
          $ += (pe || en(St, n.nbsp)) + De;
        }
        D.push($);
      } else
        ne.int += W(c, t.int_p, H.num, ne.int);
    else if (Y === "frac") {
      const X = ne.frac;
      for (let le = 0; le < te; le++)
        D.push(d[le + X] || en(H.num[le], n.nbsp));
      ne.frac += te;
    } else if (Y === "man")
      !ne[Y] && !ne.man && D.push(s), ne.man += W(o, t.man_p, H.num, ne.man);
    else if (Y === "num")
      ne.num += W(l, t.num_p, H.num, ne.num);
    else if (Y === "den") {
      const X = ne.den;
      for (let le = 0; le < te; le++) {
        let $ = f[le + X];
        if (!$) {
          const j = H.num[le];
          "123456789".includes(j) || J && j === "0" ? (J = !0, $ = n.nbsp ? " " : " ") : !J && le === te - 1 && j === "0" && !f ? $ = "1" : $ = en(j, n.nbsp);
        }
        D.push($);
      }
      ne.den += te;
    } else if (Y === "year")
      h < 0 && D.push(C.negative), D.push(String(Math.abs(h)).padStart(4, "0"));
    else if (Y === "year-short") {
      const X = h % 100;
      D.push(X < 10 ? "0" : "", X);
    } else if (Y === "month")
      D.push(H.pad && g < 10 ? "0" : "", g);
    else if (Y === "monthname-single")
      t.date_system === Ln ? D.push(C.mmmm6[g - 1].charAt(0)) : D.push(C.mmmm[g - 1].charAt(0));
    else if (Y === "monthname-short")
      t.date_system === Ln ? D.push(C.mmm6[g - 1]) : D.push(C.mmm[g - 1]);
    else if (Y === "monthname")
      t.date_system === Ln ? D.push(C.mmmm6[g - 1]) : D.push(C.mmmm[g - 1]);
    else if (H.type === "weekday-short")
      D.push(C.ddd[p]);
    else if (Y === "weekday")
      D.push(C.dddd[p]);
    else if (Y === "day")
      D.push(H.pad && m < 10 ? "0" : "", m);
    else if (Y === "hour") {
      const X = E % t.clock || (t.clock < 24 ? t.clock : 0);
      D.push(H.pad && X < 10 ? "0" : "", X);
    } else if (Y === "min")
      D.push(H.pad && b < 10 ? "0" : "", b);
    else if (Y === "sec")
      D.push(H.pad && N < 10 ? "0" : "", N);
    else if (Y === "subsec") {
      D.push(C.decimal);
      const X = v.toFixed(t.sec_decimals);
      D.push(X.slice(2, 2 + H.decimals));
    } else if (Y === "ampm") {
      const X = E < 12 ? 0 : 1;
      H.short && !r ? D.push("AP"[X]) : D.push(C.ampm[X]);
    } else if (Y === "hour-elap") {
      e < 0 && D.push(C.negative);
      const X = a * 24 + Math.floor(Math.abs(u) / 3600);
      D.push(String(Math.abs(X)).padStart(H.pad, "0"));
    } else if (Y === "min-elap") {
      e < 0 && D.push(C.negative);
      const X = a * 1440 + Math.floor(Math.abs(u) / 60);
      D.push(String(Math.abs(X)).padStart(H.pad, "0"));
    } else if (Y === "sec-elap") {
      e < 0 && D.push(C.negative);
      const X = a * Kn + Math.abs(u);
      D.push(String(Math.abs(X)).padStart(H.pad, "0"));
    } else if (Y === "b-year")
      D.push(h + 543);
    else if (Y === "b-year-short") {
      const X = (h + 543) % 100;
      D.push(X < 10 ? "0" : "", X);
    }
  }
  return D.join("");
}
const lv = Jn([
  { type: rs, value: "@", raw: "@" }
]);
function cv(e, t) {
  for (let n = 0; n < 3; n++) {
    const r = t[n];
    if (r) {
      let o;
      if (r.condition) {
        const s = r.condition[0], l = r.condition[1];
        s === "=" ? o = e === l : s === ">" ? o = e > l : s === "<" ? o = e < l : s === ">=" ? o = e >= l : s === "<=" ? o = e <= l : s === "<>" && (o = e !== l);
      } else
        o = !0;
      if (o)
        return r;
    }
  }
}
function uv(e, t, n) {
  const r = t.partitions, o = Y2(t.locale || n.locale), s = r[3] ? r[3] : lv;
  if (typeof e == "boolean" && (e = (o || no).bool[e ? 0 : 1]), e == null)
    return "";
  const l = typeof e == "bigint";
  if (typeof e != "number" && !l)
    return Xs(e, s, n, o);
  if (!l && !isFinite(e)) {
    const d = o || no;
    return isNaN(e) ? d.nan : (e < 0 ? d.negative : "") + d.infinity;
  }
  const f = cv(e, r);
  return f ? Xs(e, f, n, o) : n.overflow;
}
function fv(e) {
  return !!(e[0]?.date || e[1]?.date || e[2]?.date || e[3]?.date);
}
const dv = [
  [Jl, /^General/i, 0],
  [Ql, /^#/, 0],
  [er, /^0/, 0],
  [jl, /^\?/, 0],
  [Zl, /^\//, 0],
  // Commas are dealt with as a special case in the tokenizer but will end up
  // as one of these:
  // [ TOKEN_GROUP, /^(,),*/, 1 ],
  // [ TOKEN_SCALE, /^(,),*/, 1 ],
  // [ TOKEN_COMMA, /^(,),*/, 1 ],
  [tc, /^;/, 0],
  [rs, /^@/, 0],
  [nc, /^\+/, 0],
  [rc, /^-/, 0],
  [hi, /^\./, 0],
  [oc, /^ /, 0],
  [ic, /^%/, 0],
  [os, /^[1-9]/, 0],
  [sc, /^(?:B[12])/i, 0],
  [to, /^B$/, 0],
  // pattern must not end in a "B"
  [ac, /^(?:[hH]+|[mM]+|[sS]+|[yY]+|[bB]+|[dD]+|[gG]+|[aA]{3,}|e+)/, 0],
  [lc, /^(?:\[(h+|m+|s+)\])/i, 1],
  [cc, /^\[(<[=>]?|>=?|=)\s*(-?[.\d]+)\]/, [1, 2]],
  [uc, /^\[(DBNum[0-4]?\d)\]/i, 1],
  [fc, /^\[(NatNum[0-4]?\d)\]/i, 1],
  [dc, /^\[\$([^\]]+)\]/, 1],
  [hc, /^\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)\]/i, 1],
  // conditionally allow these open ended directions?
  [q2, /^\[([^\]]+)\]/, 1],
  [mc, /^(?:AM\/PM|am\/pm|A\/P)/, 0],
  [gc, /^\\(.)/, 1],
  [pc, /^"([^"]*?)"/, 1],
  [vc, /^_(\\.|.)/, 1],
  // Google Sheets and Excel diverge on "e": Excel only accepts E.
  [Nc, /^[Ee]([+-])/, 1],
  [Cc, /^\*(\\.|.)/, 1],
  [Ec, /^[()]/, 0],
  [to, /^[EÈÉÊËèéêëĒēĔĕĖėĘęĚěȄȅȆȇȨȩNnÑñŃńŅņŇňǸǹ["*/\\_]/, 0],
  [is, /^./, 0]
], hv = 63, mv = 35, gv = 48, pv = 57, _o = (e) => {
  const t = (e || "\0").charCodeAt(0);
  return t === hv || t === mv || t >= gv && t <= pv;
};
function kr(e) {
  let t = 0;
  const n = [], r = [];
  for (; t < e.length; ) {
    const o = e.slice(t);
    let s = 0;
    const l = /^(,+)(.)?/.exec(o);
    if (l) {
      const f = l[1];
      s = f.length;
      const d = e[t - 1] || "";
      let c = !1, i = !1;
      _o(d) ? (c = !0, i = !0) : d === "." && (i = !0);
      const a = l[2] || "";
      if (c && (!a || a === ";") && (c = !1), i && _o(a) && (i = !1), c && !i)
        n.push({ type: eo, value: ",", raw: f });
      else if (!c && i)
        n.push({ type: di, value: ",", raw: f });
      else if (c && i) {
        const u = { type: di, value: ",", raw: f };
        n.push(u), r.push(u);
      } else
        n.push({ type: ec, value: ",", raw: f });
    } else {
      let f;
      for (const [d, c, i] of dv) {
        const a = c.exec(o);
        if (a) {
          const u = Array.isArray(i) ? i.map((h) => a[h]) : a[i || 0];
          f = { type: d, value: u, raw: a[0] }, n.push(f), s = a[0].length;
          break;
        }
      }
      r.length && f.raw === ";" && (r.length = 0), r.length && _o(f.raw) && (r.forEach((d) => d.type = eo), r.length = 0);
    }
    if (!s) {
      const f = o[0];
      s = 1, n.push({ type: is, value: f, raw: f });
    }
    t += s;
  }
  return n;
}
const Bo = (e) => {
  const [t, n] = e.condition ?? [];
  n < 0 && (t === "<" || t === "<=" || t === "=") || n === 0 && t === "<" || e.tokens.unshift({
    type: "minus",
    volatile: !0
  });
}, Ks = (e, t = null) => {
  const n = {};
  for (const r in e)
    Array.isArray(e[r]) ? n[r] = [...e[r]] : n[r] = e[r];
  return t && n.tokens.unshift(t), n.generated = !0, n;
};
function vv(e) {
  const t = [];
  let n = !1, r, o = null, s = 0, l = !1, f = 0, d = 0, c = kr(e);
  do {
    if (l = Jn(c), (l.date || l.general) && (l.int_pattern.length || l.frac_pattern.length || l.scale !== 1 || l.text))
      throw new Error("Illegal format");
    if (l.condition && (d++, n = !0), l.text) {
      if (o)
        throw new Error("Unexpected partition");
      o = l;
    }
    l.locale && (r = ss(l.locale)), t.push(l), s = c[l.tokensUsed]?.type === "break" ? 1 : 0, c = c.slice(l.tokensUsed + s), f++;
  } while (s && f < 4 && d < 3);
  if (s)
    throw new Error("Unexpected partition");
  if (d > 2)
    throw new Error("Unexpected condition");
  const i = t[3];
  if (i && (i.int_pattern.length || i.frac_pattern.length || i.date))
    throw new Error("Unexpected partition");
  if (n) {
    t[0].condition || (t[0].condition = [">", 0]);
    const a = t.length;
    if (a === 1 && (t[1] = Jn(kr("General")), t[1].generated = !0), a < 3) {
      const u = t[0], h = t[1];
      if (Bo(u), h.condition)
        Bo(h);
      else {
        const g = u.condition;
        (g[0] === "=" || g[1] >= 0 && (g[0] === ">" || g[0] === ">=")) && h.tokens.unshift({
          type: "minus",
          volatile: !0
        });
      }
    } else
      t.forEach(Bo);
  } else {
    if (t.length < 4 && o)
      for (let a = 0, u = t.length; a < u; a++)
        t[a] === o && t.splice(a, 1);
    if (t.length < 1 && o && (t[0] = Jn(kr("General")), t[0].generated = !0), t.length < 2) {
      const a = { type: "minus", volatile: !0 };
      t.push(Ks(t[0], a));
    }
    if (t.length < 3 && t.push(Ks(t[0])), t.length < 4)
      if (o)
        t.push(o);
      else {
        const a = Jn(kr("@"));
        a.generated = !0, t.push(a);
      }
    t[0].condition = [">", 0], t[1].condition = ["<", 0], t[2].condition = null;
  }
  return {
    pattern: e,
    partitions: t,
    locale: r
  };
}
const Js = /* @__PURE__ */ Object.create({});
function Dc(e, t = !1) {
  e || (e = "General");
  let n = Js[e];
  if (!n)
    try {
      n = vv(e), Js[e] = n;
    } catch (r) {
      if (t)
        throw r;
      const o = {
        tokens: [{ type: "error" }],
        error: r.message
      };
      n = {
        pattern: e,
        partitions: [o, o, o, o],
        error: r.message,
        locale: null
      };
    }
  return n;
}
function Nv(e, t, n = {}) {
  const r = Object.assign({}, K2, n), o = Dc(e, r.throws), s = Ac(t, r) ?? t;
  return uv(s, o, r);
}
function Cv(e) {
  const t = Dc(e, !1);
  return fv(t.partitions);
}
function Nn(e, t, n = {}) {
  return Nv(
    t ?? cn,
    e,
    n
  );
}
function us(e) {
  return e ? Cv(e) : !1;
}
const Ev = Ac;
function bv(e, t) {
  const { row: n, col: r, rowCount: o, colCount: s, sheetId: l } = e, f = {
    labels: [],
    datasets: []
  };
  for (let d = n, c = 1, i = n + o; d < i; d++, c++) {
    if (t.getRowHeight(d) === Wo)
      continue;
    const u = [];
    for (let h = r, g = r + s; h < g; h++) {
      const m = t.getCell({
        row: d,
        col: h,
        rowCount: 1,
        colCount: 1,
        sheetId: l
      });
      m && u.push(Un(m.value)[1]);
    }
    u.length > 0 && f.datasets.push({ label: `Series${c}`, data: u });
  }
  return f.datasets[0] && f.datasets[0].data.length > 0 && (f.labels = Array.from({ length: f.datasets[0].data.length }).fill("").map((d, c) => String(c + 1))), f;
}
function wv(e, t) {
  const n = Oc(t), { range: r, isMerged: o } = e.getActiveRange(), s = r.sheetId || e.getCurrentSheetId(), l = e.getCell(r), f = e.getDefineName({
    row: r.row,
    col: r.col,
    rowCount: 1,
    colCount: 1,
    sheetId: s
  }), d = e.getCellSize(r), c = e.computeCellPosition(r);
  c.top += n?.top ?? 0;
  let i = l?.fontFamily ?? "";
  if (!i) {
    let D = "";
    const W = be.getState().fontFamilies;
    for (const J of W)
      if (!J.disabled) {
        D = String(J.value);
        break;
      }
    i = D;
  }
  const {
    isBold: a = !1,
    isItalic: u = !1,
    isStrike: h = !1,
    fontSize: g = Cn,
    fontColor: m = Ee("contentColor"),
    fillColor: p = "",
    isWrapText: E = !1,
    underline: b = Ce.NONE,
    horizontalAlign: N,
    verticalAlign: v
  } = l || {}, C = l?.numberFormat || cn, M = C === cn && typeof l?.value == "number";
  let R = N;
  N === void 0 && M && (R = xe.RIGHT);
  let A = Nn(l?.value, C), x = "";
  if (us(C) || C === xt(10))
    x = A;
  else {
    const D = l?.value ?? "";
    x = typeof D == "boolean" ? D.toString().toUpperCase() : String(D), A = x;
  }
  let L = "";
  o && (A.includes(mt) ? (L = String(Ct.MERGE_CONTENT), A = A.replaceAll(mt, dr)) : L = String(Ct.MERGE_CENTER)), xn.setState({
    top: c.top,
    left: c.left,
    width: d.width,
    height: d.height,
    row: r.row,
    col: r.col,
    rowCount: r.rowCount,
    colCount: r.colCount,
    value: l?.formula || x,
    defineName: f,
    displayValue: l?.formula || A
  }), es.setState({
    isBold: a,
    isItalic: u,
    isStrike: h,
    fontColor: m,
    fontSize: g,
    fontFamily: i,
    fillColor: p,
    isWrapText: E,
    underline: b,
    numberFormat: C,
    isMergeCell: o,
    mergeType: L,
    horizontalAlign: R,
    verticalAlign: v
  });
}
const zo = (e, t, n) => {
  (e.has("rangeMap") || e.has("cellStyle") || e.has("worksheets") || e.has("currentSheetId") || e.has("mergeCells")) && wv(t, n);
  const r = {
    canRedo: t.canRedo(),
    canUndo: t.canUndo()
  };
  if (e.has("workbook")) {
    const o = t.getSheetList().map((s) => ({
      sheetId: s.sheetId,
      name: s.name,
      isHide: s.isHide,
      tabColor: s.tabColor || ""
    }));
    r.sheetList = o;
  }
  if (e.has("currentSheetId") && (r.activeUuid = "", r.currentSheetId = t.getCurrentSheetId(), r.isFilter = !!t.getFilter()), e.has("autoFilter") && (r.isFilter = !!t.getFilter()), e.has("definedNames")) {
    const o = t.getDefineNameList().map((s) => s.name);
    r.defineNames = o;
  }
  if (e.has("drawings") || e.has("worksheets") || e.has("customHeight") || e.has("customWidth") || e.has("currentSheetId") || e.has("scroll")) {
    const o = t.getScroll(), s = t.getCanvasSize(), l = o.left, f = o.top, d = s.width + l, c = s.height + f, i = t.getDrawingList(t.getCurrentSheetId()), a = [];
    for (const u of i) {
      const h = t.computeCellPosition({
        row: u.fromRow,
        col: u.fromCol,
        colCount: 1,
        rowCount: 1,
        sheetId: ""
      }), g = h.top + u.marginY, m = h.left + u.marginX;
      if (g > f && g < c || m > l && m < d || g + u.height > f && g + u.height < c || m + u.width > l && m + u.width < d) {
        const E = {
          ...u,
          top: g,
          left: m,
          labels: [],
          datasets: []
        };
        if (u.type === "chart") {
          const b = bv(u.chartRange, t);
          E.labels = b.labels, E.datasets = b.datasets;
        }
        a.push(E);
      }
    }
    r.drawings = a;
  }
  if (be.setState(r), e.has("scroll")) {
    const o = t.getScroll(), s = t.getCanvasSize();
    pn.setState({
      scrollLeft: o.scrollLeft,
      scrollTop: o.scrollTop,
      row: o.row,
      col: o.col,
      canvasHeight: s.height,
      canvasWidth: s.width
    });
  }
  e.has("currentSheetId") && be.getState().sheetList.length > 5 && setTimeout(() => {
    i2(t.getCurrentSheetId());
  }, 0);
};
function Oc(e) {
  const t = e.parentElement;
  if (!t)
    return null;
  const n = 20, r = t.getBoundingClientRect(), o = t.clientWidth - n, s = t.clientHeight - n;
  return {
    top: r.top,
    left: r.left,
    width: o,
    height: s
  };
}
function Iv(e, t) {
  be.getState().setFontFamilies(w2()), vn.getState().setClientId(e.getHooks().doc.clientID);
  const n = xo.instance || (xo.instance = new xo(e, t)), r = (i) => {
    const a = Oc(t);
    return a && e.setCanvasSize(a), n.resize(), n.render({ changeSet: i });
  }, o = () => {
    r(/* @__PURE__ */ new Set(["customWidth"]));
  }, s = e.on("renderChange", ({ changeSet: i }) => {
    zo(i, e, t), n.render({ changeSet: i });
  }), l = a2(e, o), f = /* @__PURE__ */ new Set([
    ...bi,
    "scroll",
    "cellStyle",
    "antLine",
    "undo",
    "redo"
  ]);
  zo(f, e, t), r(f), va() || setTimeout(() => {
    zo(f, e, t), r(f);
  }, 50);
  const d = e.on(
    "toastMessage",
    ({ type: i, message: a, duration: u = 5, testId: h }) => {
      Pe({ type: i, message: a, duration: u, testId: h ?? `${i}-toast` });
    }
  ), c = e.model.on(
    "toastMessage",
    ({ type: i, message: a, duration: u = 5, testId: h }) => {
      Pe({ type: i, message: a, duration: u, testId: h ?? `${i}-toast` });
    }
  );
  return () => {
    l(), s(), d(), c();
  };
}
const Mv = 10, kc = ue(() => {
  const { controller: e } = Se(), [t, n] = he(Mv), r = pn((d) => d.scrollTop), o = pn((d) => d.canvasHeight), s = r / o >= 0.91, l = (d) => {
    d.stopPropagation();
    const c = parseInt(d.target.value, 10);
    isNaN(c) || (c < 1 ? n(1) : c > Eo ? n(Eo) : n(c));
  }, f = () => {
    const d = e.getSheetInfo(e.getCurrentSheetId());
    if (!d)
      return;
    e.addRow(d.rowCount - 1, t);
    const c = e.getSheetViewSize();
    on(e, 0, c.height);
  };
  return /* @__PURE__ */ I(
    "div",
    {
      className: ae($e["bottom-bar"], {
        [$e.active]: s
      }),
      "data-testid": "canvas-bottom-bar",
      children: [
        /* @__PURE__ */ I("div", { className: $e["bottom-bar-text"], children: k.t("add-at-the-bottom") }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/canvas/BottomBar.tsx",
          lineNumber: 46,
          columnNumber: 7
        }, void 0),
        /* @__PURE__ */ I(
          "input",
          {
            value: t,
            onChange: l,
            type: "number",
            min: 1,
            max: Eo,
            "data-testid": "canvas-bottom-bar-input",
            className: $e["bottom-bar-input"]
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/canvas/BottomBar.tsx",
            lineNumber: 47,
            columnNumber: 7
          },
          void 0
        ),
        /* @__PURE__ */ I("div", { className: $e["bottom-bar-text"], children: k.t("rows") }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/canvas/BottomBar.tsx",
          lineNumber: 56,
          columnNumber: 7
        }, void 0),
        /* @__PURE__ */ I(
          V,
          {
            testId: "canvas-bottom-bar-add",
            className: $e["add-button"],
            onClick: f,
            children: k.t("add")
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/canvas/BottomBar.tsx",
            lineNumber: 57,
            columnNumber: 7
          },
          void 0
        )
      ]
    },
    void 0,
    !0,
    {
      fileName: "/home/user/excel-collab/src/containers/canvas/BottomBar.tsx",
      lineNumber: 40,
      columnNumber: 5
    },
    void 0
  );
});
kc.displayName = "BottomBar";
const yv = "_active_k5rmq_11", Sv = "_image_k5rmq_52", Rv = "_top_k5rmq_89", Av = "_bottom_k5rmq_95", Tv = "_right_k5rmq_121", xv = "_left_k5rmq_127", Dv = "_rotate_k5rmq_78", se = {
  "float-element": "_float-element_k5rmq_1",
  active: yv,
  "float-element-mask": "_float-element-mask_k5rmq_18",
  "context-menu": "_context-menu_k5rmq_30",
  image: Sv,
  "resize-handler": "_resize-handler_k5rmq_57",
  "scale-dot": "_scale-dot_k5rmq_70",
  "rotate-icon": "_rotate-icon_k5rmq_78",
  top: Rv,
  bottom: Av,
  "top-right": "_top-right_k5rmq_101",
  "top-left": "_top-left_k5rmq_106",
  "bottom-right": "_bottom-right_k5rmq_111",
  "bottom-left": "_bottom-left_k5rmq_116",
  right: Tv,
  left: xv,
  rotate: Dv,
  "chart-type-select": "_chart-type-select_k5rmq_148"
}, Lc = ue(
  (e) => {
    const {
      menuLeft: t,
      menuTop: n,
      uuid: r,
      type: o,
      chartType: s,
      title: l,
      resetResize: f,
      hideContextMenu: d,
      originHeight: c,
      originWidth: i,
      width: a,
      height: u
    } = e, { controller: h } = Se(), g = bn(!0, d), m = () => {
      let N = ut(
        e.chartRange,
        "absolute",
        (v) => h.getSheetInfo(v)?.name || ""
      );
      Zn({
        visible: !0,
        title: k.t("edit-data-source"),
        testId: "dialog-select-data",
        children: /* @__PURE__ */ I(
          "input",
          {
            type: "text",
            spellCheck: !0,
            style: { width: "400px" },
            defaultValue: N,
            onChange: (v) => {
              N = v.target.value.trim(), v.stopPropagation();
            },
            maxLength: or * 2,
            "data-testid": "dialog-select-data-input"
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
            lineNumber: 57,
            columnNumber: 11
          },
          void 0
        ),
        onOk: () => {
          if (!N)
            return Pe.error(
              k.t("reference-is-empty"),
              "select-data-empty-toast"
            );
          const v = h.getSheetList(), C = Ot(N, (M) => v.find((R) => R.name === M)?.sheetId || "");
          if (!C || !h.validateRange(C) || e.chartRange && Ia(C, e.chartRange))
            return Pe.error(
              k.t("reference-is-not-valid"),
              "select-data-invalid-toast"
            );
          C.sheetId = C.sheetId || h.getCurrentSheetId(), h.updateDrawing(r, { chartRange: C }), d();
        },
        onCancel: () => {
          d();
        }
      });
    }, p = () => {
      let N = l.trim();
      Zn({
        visible: !0,
        title: k.t("change-chart-title"),
        testId: "dialog-change-chart-title",
        children: /* @__PURE__ */ I(
          "input",
          {
            type: "text",
            spellCheck: !0,
            style: { width: "200px" },
            defaultValue: N,
            onChange: (v) => {
              N = v.target.value.trim(), v.stopPropagation();
            },
            maxLength: or,
            "data-testid": "dialog-change-chart-title-input"
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
            lineNumber: 107,
            columnNumber: 11
          },
          void 0
        ),
        onOk: () => {
          if (!N)
            return Pe.error(
              k.t("the-value-cannot-be-empty"),
              "change-chart-title-toast"
            );
          h.updateDrawing(r, { title: N }), d();
        },
        onCancel: () => {
          d();
        }
      });
    }, E = () => {
      let N = s;
      Zn({
        title: k.t("change-chart-type"),
        testId: "dialog-change-chart-type",
        visible: !0,
        children: /* @__PURE__ */ I(
          ln,
          {
            className: se["chart-type-select"],
            defaultValue: N,
            data: Ko.map((v) => ({ ...v, disabled: !1 })),
            onChange: (v) => N = String(v),
            testId: "dialog-change-chart-type-select"
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
            lineNumber: 142,
            columnNumber: 11
          },
          void 0
        ),
        onCancel() {
          d();
        },
        onOk() {
          h.updateDrawing(r, { chartType: N }), d();
        }
      });
    }, b = () => {
      d();
      const v = h.getDrawingList(h.getCurrentSheetId()).find((C) => C.uuid === r);
      if (v) {
        if (o === "floating-picture" && v.imageSrc) {
          const C = Fa(v.imageSrc);
          ar(v.imageSrc, v.title + C.ext);
        }
        if (o === "chart") {
          const C = document.querySelector(
            `canvas[data-uuid="${r}"]`
          );
          if (!C)
            return;
          const M = C.toDataURL();
          ar(M, v.title + ".png");
        }
      }
    };
    return /* @__PURE__ */ I(
      "div",
      {
        className: se["context-menu"],
        "data-testid": "float-element-context-menu",
        ref: g,
        style: { top: n, left: t },
        children: [
          /* @__PURE__ */ I(
            V,
            {
              testId: "float-element-context-menu-copy",
              onClick: () => {
                d(), h.setFloatElementUuid(r), h.copy();
              },
              children: k.t("copy")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
              lineNumber: 188,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "float-element-context-menu-cut",
              onClick: () => {
                d(), h.setFloatElementUuid(r), h.cut();
              },
              children: k.t("cut")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
              lineNumber: 199,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "float-element-context-menu-paste",
              onClick: () => {
                d(), h.paste();
              },
              children: k.t("paste")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
              lineNumber: 209,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "float-element-context-menu-duplicate",
              onClick: () => {
                d(), h.setFloatElementUuid(r), h.copy(), h.paste(), h.setFloatElementUuid("");
              },
              children: k.t("duplicate")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
              lineNumber: 218,
              columnNumber: 9
            },
            void 0
          ),
          o === "chart" && /* @__PURE__ */ I(je.Fragment, { children: [
            /* @__PURE__ */ I(
              V,
              {
                onClick: m,
                testId: "float-element-context-menu-select-data",
                children: k.t("select-data")
              },
              void 0,
              !1,
              {
                fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
                lineNumber: 232,
                columnNumber: 13
              },
              void 0
            ),
            /* @__PURE__ */ I(
              V,
              {
                onClick: p,
                testId: "float-element-context-menu-change-chart-title",
                children: k.t("change-chart-title")
              },
              void 0,
              !1,
              {
                fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
                lineNumber: 238,
                columnNumber: 13
              },
              void 0
            ),
            /* @__PURE__ */ I(
              V,
              {
                onClick: E,
                testId: "float-element-context-menu-change-chart-type",
                children: k.t("change-chart-type")
              },
              void 0,
              !1,
              {
                fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
                lineNumber: 244,
                columnNumber: 13
              },
              void 0
            )
          ] }, void 0, !0, {
            fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
            lineNumber: 231,
            columnNumber: 11
          }, void 0),
          /* @__PURE__ */ I(
            V,
            {
              testId: "float-element-context-menu-save-as-picture",
              onClick: b,
              children: k.t("save-as-picture")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
              lineNumber: 252,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              disabled: a === i && u === c,
              testId: "float-element-context-reset-size",
              onClick: () => {
                d(), h.updateDrawing(r, {
                  height: c,
                  width: i
                }), f({ width: i, height: c });
              },
              children: k.t("reset-size")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
              lineNumber: 258,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "float-element-context-menu-delete",
              onClick: () => {
                d(), h.deleteDrawing(r);
              },
              children: k.t("delete")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
              lineNumber: 272,
              columnNumber: 9
            },
            void 0
          )
        ]
      },
      void 0,
      !0,
      {
        fileName: "/home/user/excel-collab/src/containers/FloatElement/ContextMenu.tsx",
        lineNumber: 182,
        columnNumber: 7
      },
      void 0
    );
  }
);
Lc.displayName = "FloatElementContextMenu";
var de = /* @__PURE__ */ ((e) => (e.top = "top", e.topRight = "top-right", e.topLeft = "top-left", e.bottom = "bottom", e.bottomLeft = "bottom-left", e.bottomRight = "bottom-right", e.left = "left", e.right = "right", e.rotate = "rotate", e.active = "active", e))(de || {});
const Ht = {
  resizePosition: "",
  moveStartX: 0,
  moveStartY: 0,
  activeUuid: "",
  position: {
    width: -1,
    height: -1,
    imageAngle: 0,
    top: -1,
    left: -1
  }
};
function Qs(e, t, n) {
  const r = n.getCanvasSize(), o = n.getHeaderSize(), s = o.height, l = o.width;
  return e < s && (e = s), e > r.height && (e = r.height), t < l && (t = l), t > r.width && (t = r.width), {
    top: e,
    left: t
  };
}
const Ov = ue(
  ({ title: e, imageSrc: t = "", uuid: n }) => {
    const [r, o] = he(""), [s, l] = he(!1), { provider: f, controller: d } = Se();
    return Re(() => {
      if (!f?.downloadFile) {
        o(t);
        return;
      }
      l(!0), f.downloadFile(d.getHooks().doc.guid, t).then((c) => {
        o(c);
      }).finally(() => {
        l(!1);
      });
    }, [t, f]), s || !r ? /* @__PURE__ */ I(Co, {}, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/FloatElement/Image.tsx",
      lineNumber: 34,
      columnNumber: 14
    }, void 0) : /* @__PURE__ */ I(
      "img",
      {
        title: e,
        alt: e,
        src: r,
        className: se.image,
        "data-uuid": n,
        "data-testid": "float-element-image"
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/FloatElement/Image.tsx",
        lineNumber: 37,
        columnNumber: 7
      },
      void 0
    );
  }
), kv = je.lazy(() => import("./Chart-Cstm8qF_.mjs")), Lv = ue(
  (e) => {
    const {
      top: t,
      left: n,
      active: r,
      width: o,
      height: s,
      type: l,
      imageAngle: f = 0,
      resetResize: d,
      pointerDown: c,
      resizeDown: i
    } = e, [a, u] = he({
      top: ht,
      left: ht
    }), h = z(
      (p) => {
        p.preventDefault(), p.stopPropagation(), u({ top: p.clientY, left: p.clientX });
      },
      []
    ), g = z(() => {
      u({
        top: ht,
        left: ht
      });
    }, []);
    let m;
    return l === "floating-picture" ? m = /* @__PURE__ */ I(
      Ov,
      {
        title: e.title,
        imageSrc: e.imageSrc,
        uuid: e.uuid
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
        lineNumber: 57,
        columnNumber: 9
      },
      void 0
    ) : l === "chart" && (m = /* @__PURE__ */ I(Qc, { fallback: /* @__PURE__ */ I(Co, {}, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
      lineNumber: 65,
      columnNumber: 29
    }, void 0), children: /* @__PURE__ */ I(kv, { ...e }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
      lineNumber: 66,
      columnNumber: 11
    }, void 0) }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
      lineNumber: 65,
      columnNumber: 9
    }, void 0)), m && /* @__PURE__ */ I(je.Fragment, { children: [
      /* @__PURE__ */ I(
        "div",
        {
          onPointerDown: c,
          onContextMenu: h,
          className: ae(se["float-element"], {
            [se.active]: r
          }),
          style: {
            transform: `translateX(${n}px) translateY(${t}px) ${l === "floating-picture" ? `rotate(${f}deg)` : ""} `,
            width: o,
            height: s
          },
          "data-testid": "float-element",
          children: [
            m,
            r && /* @__PURE__ */ I(je.Fragment, { children: [
              /* @__PURE__ */ I(
                "div",
                {
                  className: ae(se["resize-handler"], se.top),
                  "data-position": de.top,
                  onPointerDown: i,
                  "data-testid": "float-element-resize-top",
                  children: /* @__PURE__ */ I("div", { className: se["scale-dot"] }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 99,
                    columnNumber: 17
                  }, void 0)
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                  lineNumber: 93,
                  columnNumber: 15
                },
                void 0
              ),
              /* @__PURE__ */ I(
                "div",
                {
                  className: ae(
                    se["resize-handler"],
                    se["top-right"]
                  ),
                  "data-position": de.topRight,
                  onPointerDown: i,
                  "data-testid": "float-element-resize-top-right",
                  children: /* @__PURE__ */ I("div", { className: se["scale-dot"] }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 110,
                    columnNumber: 17
                  }, void 0)
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                  lineNumber: 101,
                  columnNumber: 15
                },
                void 0
              ),
              /* @__PURE__ */ I(
                "div",
                {
                  className: ae(
                    se["resize-handler"],
                    se["top-left"]
                  ),
                  "data-position": de.topLeft,
                  onPointerDown: i,
                  "data-testid": "float-element-resize-top-left",
                  children: /* @__PURE__ */ I("div", { className: se["scale-dot"] }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 121,
                    columnNumber: 17
                  }, void 0)
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                  lineNumber: 112,
                  columnNumber: 15
                },
                void 0
              ),
              /* @__PURE__ */ I(
                "div",
                {
                  className: ae(se["resize-handler"], se.left),
                  "data-position": de.left,
                  onPointerDown: i,
                  "data-testid": "float-element-resize-left",
                  children: /* @__PURE__ */ I("div", { className: se["scale-dot"] }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 129,
                    columnNumber: 17
                  }, void 0)
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                  lineNumber: 123,
                  columnNumber: 15
                },
                void 0
              ),
              /* @__PURE__ */ I(
                "div",
                {
                  className: ae(
                    se["resize-handler"],
                    se.right
                  ),
                  "data-position": de.right,
                  onPointerDown: i,
                  "data-testid": "float-element-resize-right",
                  children: /* @__PURE__ */ I("div", { className: se["scale-dot"] }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 140,
                    columnNumber: 17
                  }, void 0)
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                  lineNumber: 131,
                  columnNumber: 15
                },
                void 0
              ),
              /* @__PURE__ */ I(
                "div",
                {
                  className: ae(
                    se["resize-handler"],
                    se["bottom-right"]
                  ),
                  "data-position": de.bottomRight,
                  onPointerDown: i,
                  "data-testid": "float-element-resize-bottom-right",
                  children: /* @__PURE__ */ I("div", { className: se["scale-dot"] }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 151,
                    columnNumber: 17
                  }, void 0)
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                  lineNumber: 142,
                  columnNumber: 15
                },
                void 0
              ),
              /* @__PURE__ */ I(
                "div",
                {
                  className: ae(
                    se["resize-handler"],
                    se["bottom-left"]
                  ),
                  "data-position": de.bottomLeft,
                  onPointerDown: i,
                  "data-testid": "float-element-resize-bottom-left",
                  children: /* @__PURE__ */ I("div", { className: se["scale-dot"] }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 162,
                    columnNumber: 17
                  }, void 0)
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                  lineNumber: 153,
                  columnNumber: 15
                },
                void 0
              ),
              /* @__PURE__ */ I(
                "div",
                {
                  className: ae(
                    se["resize-handler"],
                    se.bottom
                  ),
                  "data-position": de.bottom,
                  onPointerDown: i,
                  "data-testid": "float-element-resize-bottom",
                  children: /* @__PURE__ */ I("div", { className: se["scale-dot"] }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 173,
                    columnNumber: 17
                  }, void 0)
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                  lineNumber: 164,
                  columnNumber: 15
                },
                void 0
              ),
              l === "floating-picture" && /* @__PURE__ */ I(
                "div",
                {
                  className: ae(
                    se["resize-handler"],
                    se.rotate
                  ),
                  "data-position": de.rotate,
                  onPointerDown: i,
                  "data-testid": "float-element-rotate",
                  children: /* @__PURE__ */ I("div", { className: se["rotate-icon"], children: /* @__PURE__ */ I(Ae, { name: "rotate" }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 186,
                    columnNumber: 21
                  }, void 0) }, void 0, !1, {
                    fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                    lineNumber: 185,
                    columnNumber: 19
                  }, void 0)
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
                  lineNumber: 176,
                  columnNumber: 17
                },
                void 0
              )
            ] }, void 0, !0, {
              fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
              lineNumber: 92,
              columnNumber: 13
            }, void 0)
          ]
        },
        void 0,
        !0,
        {
          fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
          lineNumber: 75,
          columnNumber: 9
        },
        void 0
      ),
      a.top >= 0 && a.left >= 0 && /* @__PURE__ */ I(
        Lc,
        {
          ...e,
          resetResize: d,
          menuLeft: a.left,
          menuTop: a.top,
          hideContextMenu: g
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
          lineNumber: 194,
          columnNumber: 11
        },
        void 0
      )
    ] }, void 0, !0, {
      fileName: "/home/user/excel-collab/src/containers/FloatElement/FloatElement.tsx",
      lineNumber: 74,
      columnNumber: 7
    }, void 0);
  }
), Fc = ue(
  ({ floatElementList: e, activeUuid: t }) => {
    const { controller: n } = Se(), r = be((a) => a.setActiveUuid), o = Xe({
      ...Ht,
      position: { ...Ht.position }
    }), [s, l] = he({
      width: -1,
      height: -1,
      imageAngle: 0,
      top: -1,
      left: -1
    });
    o.current.activeUuid = t, o.current.position = { ...s };
    const [f] = We(() => {
      const a = (m) => {
        const p = Math.round(m.clientX - o.current.moveStartX), E = Math.round(m.clientY - o.current.moveStartY);
        l((b) => {
          const N = o.current.resizePosition;
          let { top: v, left: C, width: M, height: R } = b;
          return [
            de.topRight,
            de.topLeft,
            de.top
          ].includes(N) ? (R -= E, v += E) : [
            de.bottomRight,
            de.bottom,
            de.bottomLeft
          ].includes(N) && (R += E), [
            de.topLeft,
            de.bottomLeft,
            de.left
          ].includes(N) ? (M -= p, C += p) : [
            de.topRight,
            de.bottomRight,
            de.right
          ].includes(N) && (M += p), {
            ...Qs(v, C, n),
            imageAngle: b.imageAngle,
            width: M,
            height: R
          };
        });
      }, u = (m) => {
        m.stopPropagation(), m.preventDefault();
        const p = o.current.activeUuid;
        if (!p || !o.current.resizePosition) {
          o.current = {
            ...Ht,
            position: { ...Ht.position }
          };
          return;
        }
        const E = o.current.position;
        if (o.current.resizePosition === de.rotate)
          n.updateDrawing(p, {
            imageAngle: E.imageAngle
          });
        else if (o.current.resizePosition === de.active) {
          const b = n.getCanvasSize(), { left: N, top: v } = E;
          if (N >= 0 && v >= 0 && N < b.width && v < b.height) {
            const C = Ta(n, N, v);
            C && n.updateDrawing(p, {
              fromCol: C.col,
              fromRow: C.row,
              marginX: C.marginX,
              marginY: C.marginY
            });
          }
        } else
          E.height > 0 && E.width > 0 && n.updateDrawing(p, {
            height: E.height,
            width: E.width
          });
        o.current = {
          ...Ht,
          position: { ...Ht.position }
        };
      }, h = (m) => {
        if (!(!o.current.resizePosition || m.buttons <= 0))
          if (m.stopPropagation(), m.preventDefault(), o.current.resizePosition === de.rotate) {
            const p = m.clientX - o.current.moveStartX, E = m.clientY - o.current.moveStartY, b = Math.atan2(E, p) * 180 / Math.PI;
            l((N) => ({
              ...N,
              imageAngle: b
            }));
            return;
          } else if (o.current.resizePosition === de.active) {
            const p = Math.round(m.clientX - o.current.moveStartX), E = Math.round(m.clientY - o.current.moveStartY);
            l((b) => {
              const N = b.top + E, v = b.left + p;
              return {
                ...Qs(N, v, n),
                imageAngle: b.imageAngle,
                width: b.width,
                height: b.height
              };
            }), o.current.moveStartX = m.clientX, o.current.moveStartY = m.clientY;
          } else {
            a(m), o.current.moveStartX = m.clientX, o.current.moveStartY = m.clientY;
            return;
          }
      };
      function g(m) {
        const p = m ? document.addEventListener : document.removeEventListener;
        p("pointerup", u), p("pointermove", h);
      }
      return [g];
    }, []);
    Re(() => f, [f]);
    const d = z(
      (a) => {
        if (a.stopPropagation(), a.preventDefault(), a.buttons <= 0)
          return;
        const u = a.currentTarget.dataset.position || "";
        u && (o.current.resizePosition = u, o.current.moveStartX = a.clientX, o.current.moveStartY = a.clientY);
      },
      []
    ), c = z((a) => {
      l((u) => ({ ...u, ...a }));
    }, []), i = z(() => {
      o.current = {
        ...Ht,
        position: { ...Ht.position }
      }, r(""), n.setFloatElementUuid(""), f(!1);
    }, [f]);
    return /* @__PURE__ */ I(rn, { children: [
      /* @__PURE__ */ I(
        "div",
        {
          className: ae(se["float-element-mask"], {
            [se.active]: !!t
          }),
          "data-testid": "float-element-mask",
          onPointerDown: i
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/containers/FloatElement/index.tsx",
          lineNumber: 233,
          columnNumber: 9
        },
        void 0
      ),
      e.map((a) => {
        const u = a.uuid === t, h = {
          ...a,
          ...u ? s : {}
        };
        return /* @__PURE__ */ I(
          Lv,
          {
            ...h,
            active: u,
            resetResize: c,
            pointerDown: (g) => {
              g.stopPropagation(), g.preventDefault(), !(g.buttons <= 0) && (o.current.resizePosition = de.active, o.current.moveStartX = g.clientX, o.current.moveStartY = g.clientY, o.current.activeUuid = a.uuid, n.setFloatElementUuid(a.uuid), r(a.uuid), l({
                top: a.top,
                left: a.left,
                width: a.width,
                height: a.height,
                imageAngle: a.imageAngle || 0
              }), f(!0));
            },
            resizeDown: d
          },
          a.uuid,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/FloatElement/index.tsx",
            lineNumber: 247,
            columnNumber: 13
          },
          void 0
        );
      })
    ] }, void 0, !0, {
      fileName: "/home/user/excel-collab/src/containers/FloatElement/index.tsx",
      lineNumber: 232,
      columnNumber: 7
    }, void 0);
  }
);
Fc.displayName = "FloatElementContainer";
const Fv = 300;
class Pv {
  lastTimeStamp = 0;
  pointerMove(t) {
    const { controller: n, x: r, y: o, position: s } = t, l = n.getHeaderSize();
    if (!s)
      return !1;
    const { range: f, isMerged: d } = n.getActiveRange({
      row: s.row,
      col: s.col,
      colCount: 1,
      rowCount: 1,
      sheetId: ""
    }), c = n.getActiveRange().range;
    if (c.row === f.row && c.col === f.col)
      return !1;
    let i = 0, a = 0;
    if (r > l.width && o > l.height) {
      if (d)
        return n.setActiveRange(f), !1;
      a = Math.abs(s.col - c.col) + 1, i = Math.abs(s.row - c.row) + 1;
    }
    return l.width > r && l.height <= o && (i = Math.abs(s.row - c.row) + 1), l.width <= r && l.height > o && (a = Math.abs(s.col - c.col) + 1), n.setActiveRange({
      row: Math.min(s.row, c.row),
      col: Math.min(s.col, c.col),
      rowCount: i,
      colCount: a,
      sheetId: ""
    }), !1;
  }
  pointerDown(t, n) {
    const { controller: r, x: o, y: s, position: l } = t, f = r.getHeaderSize(), { timeStamp: d } = n;
    if (!l)
      return !1;
    if (f.width > o && f.height > s)
      return r.setActiveRange({
        row: 0,
        col: 0,
        colCount: 0,
        rowCount: 0,
        sheetId: ""
      }), !1;
    if (f.width > o && f.height <= s)
      return r.setActiveRange({
        row: l.row,
        col: l.col,
        rowCount: 1,
        colCount: 0,
        sheetId: ""
      }), !1;
    if (f.width <= o && f.height > s)
      return r.setActiveRange({
        row: l.row,
        col: l.col,
        rowCount: 0,
        colCount: 1,
        sheetId: ""
      }), !1;
    const { range: c } = r.getActiveRange({
      row: l.row,
      col: l.col,
      colCount: 1,
      rowCount: 1,
      sheetId: r.getCurrentSheetId()
    }), i = r.getActiveRange().range;
    return Ia(i, c) ? d - this.lastTimeStamp < Fv && be.getState().setEditorStatus(dt.EDIT_CELL) : (Ke() && Bl(r), r.setActiveRange({
      row: l.row,
      col: l.col,
      rowCount: 1,
      colCount: 1,
      sheetId: ""
    })), this.lastTimeStamp = d, !1;
  }
}
class Uv {
  pointerDown(t, n) {
    if (!t.position)
      return !1;
    const { controller: r, position: o } = t, { row: s, col: l, marginX: f, marginY: d } = o;
    if (s !== 0)
      return !1;
    const c = r.getFilter();
    if (!c)
      return !1;
    const { range: i } = c;
    if (l >= i.col && l < i.col + i.colCount) {
      const a = r.getRowHeight(s), u = r.getColWidth(l);
      if (f > u - Yo && d > a - Yo)
        return n.stopPropagation(), {
          type: "filter",
          row: s,
          col: l,
          x: n.clientX,
          y: n.clientY
        };
    }
    return !1;
  }
  pointerMove() {
    return !1;
  }
}
const js = [new Uv(), new Pv()], _v = "_modalContainer_7sk22_1", Bv = "_listItem_7sk22_16", Qn = {
  modalContainer: _v,
  listItem: Bv,
  "dialog-title": "_dialog-title_7sk22_19",
  "dialog-button": "_dialog-button_7sk22_22"
};
function zv(e, t) {
  const n = e.getCurrentSheetId(), r = e.getSheetInfo(n);
  if (!r)
    return;
  const o = e.getFilter(n);
  if (!o)
    return;
  const { range: s } = o;
  let l = /* @__PURE__ */ new Set();
  o.value && o.value.type === "normal" && (l = new Set(o.value.value));
  const f = /* @__PURE__ */ new Map(), d = [], c = s.rowCount === 0 ? r.rowCount : s.row + s.rowCount;
  for (let a = s.row + 1; a < c; a++) {
    const u = e.getCell({
      row: a,
      col: t,
      rowCount: 1,
      colCount: 1,
      sheetId: n
    }), h = u ? u?.value : mt, g = u?.value === void 0;
    f.has(h) || d.push({
      checked: !0,
      label: g ? k.t("filter-empty") : String(h),
      value: h,
      count: 1
    }), f.set(h, (f.get(h) || 0) + 1);
  }
  const i = d.findIndex((a) => a.value === mt);
  return i >= 0 && d.unshift(...d.splice(i, 1)), {
    dataList: d.map((a) => {
      const u = l.size === 0 ? !0 : l.has(a.value);
      return { ...a, count: f.get(a.value) || 0, checked: u };
    })
  };
}
const Hv = ({ col: e, hide: t }) => {
  const { controller: n } = Se(), [r, o] = he([]);
  return Re(() => {
    const s = zv(n, e);
    s && o(s.dataList);
  }, [n, e]), /* @__PURE__ */ I("div", { "data-testid": "filter-modal", children: [
    /* @__PURE__ */ I("div", { className: Qn["dialog-title"], children: k.t("filter") }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
      lineNumber: 85,
      columnNumber: 7
    }, void 0),
    /* @__PURE__ */ I("div", { children: [
      /* @__PURE__ */ I("div", { children: [
        /* @__PURE__ */ I("label", { htmlFor: "modal_all", children: [
          /* @__PURE__ */ I(
            "input",
            {
              type: "checkbox",
              id: "modal_all",
              name: "all",
              value: "all",
              "data-testid": "filter-model-all",
              checked: r.every((s) => s.checked),
              onChange: (s) => {
                o(
                  (l) => l.map((f) => ({ ...f, checked: s.target.checked }))
                );
              }
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
              lineNumber: 89,
              columnNumber: 13
            },
            void 0
          ),
          k.t("filter-all")
        ] }, void 0, !0, {
          fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
          lineNumber: 88,
          columnNumber: 11
        }, void 0),
        /* @__PURE__ */ I("span", { children: [
          "(",
          r.length,
          ")"
        ] }, void 0, !0, {
          fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
          lineNumber: 104,
          columnNumber: 11
        }, void 0)
      ] }, void 0, !0, {
        fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
        lineNumber: 87,
        columnNumber: 9
      }, void 0),
      r.map((s, l) => /* @__PURE__ */ I("div", { className: Qn.listItem, children: [
        /* @__PURE__ */ I("label", { htmlFor: `modal_${s.label}`, children: [
          /* @__PURE__ */ I(
            "input",
            {
              type: "checkbox",
              id: `modal_${s.label}`,
              name: s.label,
              value: s.label,
              checked: s.checked,
              "data-testid": "filter-model-item",
              onChange: (f) => o((d) => (d[l].checked = f.target.checked, [...d]))
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
              lineNumber: 109,
              columnNumber: 15
            },
            void 0
          ),
          s.label
        ] }, void 0, !0, {
          fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
          lineNumber: 108,
          columnNumber: 13
        }, void 0),
        /* @__PURE__ */ I("span", { children: [
          "(",
          s.count,
          ")"
        ] }, void 0, !0, {
          fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
          lineNumber: 125,
          columnNumber: 13
        }, void 0)
      ] }, s.label, !0, {
        fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
        lineNumber: 107,
        columnNumber: 11
      }, void 0))
    ] }, void 0, !0, {
      fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
      lineNumber: 86,
      columnNumber: 7
    }, void 0),
    /* @__PURE__ */ I("div", { className: Qn["dialog-button"], children: [
      /* @__PURE__ */ I(V, { onClick: t, children: k.t("cancel") }, void 0, !1, {
        fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
        lineNumber: 131,
        columnNumber: 9
      }, void 0),
      /* @__PURE__ */ I(
        V,
        {
          className: Qn["dialog-cancel"],
          type: "primary",
          onClick: () => {
            const s = r.filter((l) => l.checked).map((l) => l.value);
            if (s.length === r.length) {
              const l = n.getFilter();
              typeof l?.col == "number" && l?.value && n.updateFilter("", {
                col: void 0,
                value: void 0
              }), t();
              return;
            }
            n.updateFilter("", {
              col: e,
              value: { type: "normal", value: s }
            }), t();
          },
          children: k.t("confirm")
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
          lineNumber: 132,
          columnNumber: 9
        },
        void 0
      )
    ] }, void 0, !0, {
      fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
      lineNumber: 130,
      columnNumber: 7
    }, void 0)
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/canvas/modal/Filter.tsx",
    lineNumber: 84,
    columnNumber: 5
  }, void 0);
}, Vv = (e) => {
  let t = null;
  if (e.type === "filter")
    t = /* @__PURE__ */ I(Hv, { ...e }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/canvas/modal/index.tsx",
      lineNumber: 8,
      columnNumber: 16
    }, void 0);
  else
    throw new Error(`can't handle the modal type: ${e.type}`);
  return /* @__PURE__ */ I(
    "div",
    {
      style: { top: e.y, left: e.x },
      className: Qn.modalContainer,
      children: t
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/containers/canvas/modal/index.tsx",
      lineNumber: 14,
      columnNumber: 5
    },
    void 0
  );
};
function Gv(e) {
  const t = [];
  for (let n = 0; n < 4; n++) {
    const r = e >> n * 8 & 255;
    t.push(r);
  }
  return t;
}
const $v = () => {
  const e = vn((s) => s.users), t = pn((s) => s.row), n = pn((s) => s.col), r = be((s) => s.currentSheetId), o = e.filter((s) => s?.range?.sheetId === r);
  return /* @__PURE__ */ I(je.Fragment, { children: o.map((s) => /* @__PURE__ */ I(qv, { row: t, col: n, ...s }, s.clientId, !1, {
    fileName: "/home/user/excel-collab/src/containers/canvas/Collaboration.tsx",
    lineNumber: 26,
    columnNumber: 9
  }, void 0)) }, void 0, !1, {
    fileName: "/home/user/excel-collab/src/containers/canvas/Collaboration.tsx",
    lineNumber: 24,
    columnNumber: 5
  }, void 0);
}, qv = ({ range: e, clientId: t, row: n, col: r }) => {
  const { controller: o } = Se(), s = We(() => {
    const d = Gv(t), c = Math.max(...d), [i, a, u] = d.filter((h) => h !== c);
    return `rgba(${i},${a},${u},${c / 255})`;
  }, []), [l, f] = he({
    top: -9999,
    left: -9999,
    width: 0,
    height: 0
  });
  return Re(() => {
    const { top: d, left: c } = o.computeCellPosition(e), i = o.getColWidth(e.col), a = o.getRowHeight(e.row);
    f({ left: c, top: d, width: i, height: a });
  }, [e, n, r, o]), /* @__PURE__ */ I(
    "div",
    {
      className: $e.user,
      style: { ...l, border: `2px solid ${s}` },
      children: /* @__PURE__ */ I("div", { className: $e.userContent, style: { color: s }, children: `${k.t("user-name")} ${t}` }, void 0, !1, {
        fileName: "/home/user/excel-collab/src/containers/canvas/Collaboration.tsx",
        lineNumber: 61,
        columnNumber: 7
      }, void 0)
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/containers/canvas/Collaboration.tsx",
      lineNumber: 57,
      columnNumber: 5
    },
    void 0
  );
};
function Zs(e, t) {
  const n = t.getCanvasSize(), { clientX: r = 0, clientY: o = 0 } = e, s = r - n.left, l = o - n.top;
  return {
    position: Ta(t, s, l),
    x: s,
    y: l,
    controller: t
  };
}
const Pc = ue(() => {
  const { controller: e } = Se(), [t, n] = he(null), r = be((u) => u.activeUuid), o = be((u) => u.drawings), [s, l] = he({
    top: ht,
    left: ht
  }), f = Xe(null);
  Re(() => {
    if (f.current)
      return Iv(e, f.current);
  }, []);
  const d = z(
    (u) => {
      u.preventDefault(), l({
        top: u.clientY,
        left: u.clientX
      });
    },
    []
  ), c = () => {
    l({
      top: ht,
      left: ht
    });
  }, i = z(
    (u) => {
      if (u.buttons <= 0)
        return;
      const h = Zs(u, e);
      for (const g of js) {
        const m = g.pointerMove(h, u);
        if (m) {
          typeof m != "boolean" && n(m);
          break;
        }
      }
    },
    []
  ), a = z(
    (u) => {
      if (u.buttons <= 0)
        return;
      n(null);
      const h = Zs(u, e);
      for (const g of js) {
        const m = g.pointerDown(h, u);
        if (m) {
          typeof m != "boolean" && n(m);
          break;
        }
      }
    },
    []
  );
  return /* @__PURE__ */ I(rn, { children: [
    /* @__PURE__ */ I(
      "div",
      {
        className: $e["canvas-container"],
        "data-testid": "canvas-container",
        children: [
          /* @__PURE__ */ I(
            "canvas",
            {
              className: $e["canvas-content"],
              onContextMenu: d,
              onPointerMove: i,
              onPointerDown: a,
              ref: f,
              "data-testid": "canvas-main"
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/index.tsx",
              lineNumber: 120,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(Yl, {}, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/canvas/index.tsx",
            lineNumber: 128,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(kc, {}, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/canvas/index.tsx",
            lineNumber: 129,
            columnNumber: 9
          }, void 0),
          o.length > 0 && /* @__PURE__ */ I(
            Fc,
            {
              floatElementList: o,
              activeUuid: r
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/canvas/index.tsx",
              lineNumber: 131,
              columnNumber: 11
            },
            void 0
          ),
          /* @__PURE__ */ I($v, {}, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/canvas/index.tsx",
            lineNumber: 136,
            columnNumber: 9
          }, void 0)
        ]
      },
      void 0,
      !0,
      {
        fileName: "/home/user/excel-collab/src/containers/canvas/index.tsx",
        lineNumber: 116,
        columnNumber: 7
      },
      void 0
    ),
    s.top >= 0 && s.left >= 0 && /* @__PURE__ */ I(Xl, { ...s, hideContextMenu: c }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/canvas/index.tsx",
      lineNumber: 139,
      columnNumber: 9
    }, void 0),
    t && /* @__PURE__ */ I(Vv, { ...t, hide: () => n(null) }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/canvas/index.tsx",
      lineNumber: 141,
      columnNumber: 22
    }, void 0)
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/canvas/index.tsx",
    lineNumber: 115,
    columnNumber: 5
  }, void 0);
});
Pc.displayName = "CanvasContainer";
const Wv = "_active_1d9ea_30", Je = {
  "sheet-bar-wrapper": "_sheet-bar-wrapper_1d9ea_1",
  "sheet-bar-list": "_sheet-bar-list_1d9ea_11",
  "sheet-bar-item": "_sheet-bar-item_1d9ea_22",
  active: Wv,
  "sheet-bar-item-text": "_sheet-bar-item-text_1d9ea_35",
  "sheet-bar-item-color": "_sheet-bar-item-color_1d9ea_39",
  "sheet-bar-context-menu": "_sheet-bar-context-menu_1d9ea_48",
  "sheet-bar-input": "_sheet-bar-input_1d9ea_63",
  "sheet-bar-unhide": "_sheet-bar-unhide_1d9ea_79",
  "add-button": "_add-button_1d9ea_83",
  "menu-button": "_menu-button_1d9ea_90",
  "select-popup": "_select-popup_1d9ea_98",
  "unhide-select": "_unhide-select_1d9ea_103"
}, Uc = ue(
  ({ position: e, sheetList: t, currentSheetId: n, hideMenu: r, editSheetName: o }) => {
    const { controller: s } = Se(), l = bn(!0, r), f = We(() => t.find((a) => a.sheetId === n)?.tabColor || "", [t, n]), d = We(() => t.filter((a) => a.isHide).map((a) => ({
      value: String(a.sheetId),
      label: a.name,
      disabled: !1
    })), [t]), c = () => {
      let a = String(d[0]?.value) || "";
      Zn({
        visible: !0,
        title: k.t("unhide-sheet"),
        testId: "sheet-bar-context-menu-unhide-dialog",
        children: /* @__PURE__ */ I(
          ln,
          {
            data: d,
            onChange: (u) => {
              a = String(u);
            },
            className: Je["unhide-select"],
            defaultValue: a,
            testId: "sheet-bar-context-menu-unhide-dialog-select"
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/SheetBar/SheetBarContextMenu.tsx",
            lineNumber: 42,
            columnNumber: 11
          },
          void 0
        ),
        onCancel: r,
        onOk() {
          if (!a)
            return Pe.error(k.t("sheet-id-can-not-be-empty"));
          s.unhideSheet(a), r();
        }
      });
    }, i = (a) => {
      s.updateSheetInfo({ tabColor: a }), r();
    };
    return /* @__PURE__ */ I(
      "div",
      {
        className: Je["sheet-bar-context-menu"],
        style: { left: e },
        ref: l,
        "data-testid": "sheet-bar-context-menu",
        children: [
          /* @__PURE__ */ I(
            V,
            {
              testId: "sheet-bar-context-menu-insert",
              onClick: () => {
                r(), s.addSheet();
              },
              children: k.t("insert")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/SheetBar/SheetBarContextMenu.tsx",
              lineNumber: 73,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "sheet-bar-context-menu-delete",
              onClick: () => {
                r(), s.deleteSheet();
              },
              children: k.t("delete")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/SheetBar/SheetBarContextMenu.tsx",
              lineNumber: 82,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "sheet-bar-context-menu-rename",
              onClick: () => {
                r(), o();
              },
              children: k.t("rename")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/SheetBar/SheetBarContextMenu.tsx",
              lineNumber: 91,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "sheet-bar-context-menu-hide",
              onClick: () => {
                r(), s.hideSheet();
              },
              children: k.t("hide")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/SheetBar/SheetBarContextMenu.tsx",
              lineNumber: 100,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            V,
            {
              testId: "sheet-bar-context-menu-unhide",
              className: Je["sheet-bar-unhide"],
              disabled: d.length === 0,
              onClick: c,
              children: k.t("unhide")
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/SheetBar/SheetBarContextMenu.tsx",
              lineNumber: 109,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(
            ur,
            {
              color: f,
              onChange: i,
              position: "top",
              testId: "sheet-bar-context-menu-tab-color",
              children: /* @__PURE__ */ I(
                V,
                {
                  className: Je["sheet-bar-unhide"],
                  testId: "sheet-bar-context-menu-tab-color",
                  children: k.t("tab-color")
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/SheetBar/SheetBarContextMenu.tsx",
                  lineNumber: 123,
                  columnNumber: 11
                },
                void 0
              )
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/SheetBar/SheetBarContextMenu.tsx",
              lineNumber: 117,
              columnNumber: 9
            },
            void 0
          )
        ]
      },
      void 0,
      !0,
      {
        fileName: "/home/user/excel-collab/src/containers/SheetBar/SheetBarContextMenu.tsx",
        lineNumber: 67,
        columnNumber: 7
      },
      void 0
    );
  }
);
Uc.displayName = "SheetBarContextMenu";
const _c = ue(({ children: e }) => {
  const { controller: t } = Se(), n = be((v) => v.sheetList), r = We(() => n.filter((v) => !v.isHide), [n]), o = We(() => n.filter((v) => !v.isHide).map((v) => ({ value: v.sheetId, label: v.name, disabled: !1 })), [n]), [s, l] = he(!1), f = be((v) => v.currentSheetId), [d, c] = he(ht), [i, a] = he(!1), u = z(
    (v) => {
      v.preventDefault();
      const C = v.clientX - 30;
      return c(C), !1;
    },
    []
  ), h = z(
    (v) => {
      if (v.stopPropagation(), v.key === "Enter") {
        const C = v.currentTarget.value;
        if (a(!1), !C)
          return;
        t.renameSheet(C);
      }
    },
    []
  ), g = z((v) => {
    l(!1), t.setCurrentSheetId(v);
  }, []), m = z(() => {
    t.addSheet();
  }, []), p = z(() => {
    c(ht);
  }, []), E = z(() => {
    a(!0);
  }, []), b = z(() => {
    l((v) => !v);
  }, []), N = bn(s, () => {
    l(!1);
  });
  return /* @__PURE__ */ I("div", { className: Je["sheet-bar-wrapper"], "data-testid": "sheet-bar", children: [
    /* @__PURE__ */ I("div", { ref: N, children: [
      /* @__PURE__ */ I(
        V,
        {
          onClick: b,
          className: Je["menu-button"],
          testId: "sheet-bar-select-sheet",
          children: /* @__PURE__ */ I(Ae, { name: "menu" }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
            lineNumber: 83,
            columnNumber: 13
          }, void 0)
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
          lineNumber: 78,
          columnNumber: 11
        },
        void 0
      ),
      s && /* @__PURE__ */ I(
        Ki,
        {
          data: o,
          onChange: g,
          active: !0,
          position: "top",
          value: f,
          testId: "sheet-bar-select-sheet-popup",
          className: Je["select-popup"]
        },
        void 0,
        !1,
        {
          fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
          lineNumber: 86,
          columnNumber: 13
        },
        void 0
      )
    ] }, void 0, !0, {
      fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
      lineNumber: 77,
      columnNumber: 9
    }, void 0),
    /* @__PURE__ */ I("div", { className: Je["sheet-bar-list"], "data-testid": "sheet-bar-list", children: r.map((v) => {
      const C = f === v.sheetId, M = C && i, R = v.tabColor || "", A = ae(Je["sheet-bar-item"], {
        [Je.active]: C
      });
      let x;
      !C && R && (x = { backgroundColor: R });
      const L = C ? "sheet-bar-active-item" : void 0;
      return /* @__PURE__ */ I(
        "div",
        {
          "data-testid": `${fa}${v.sheetId}`,
          className: A,
          style: x,
          onContextMenu: u,
          onClick: () => {
            f !== v.sheetId && (a(!1), t.setCurrentSheetId(v.sheetId));
          },
          children: M ? /* @__PURE__ */ I(
            "input",
            {
              className: Je["sheet-bar-input"],
              defaultValue: v.name,
              onKeyDown: h,
              type: "text",
              spellCheck: !0,
              maxLength: or,
              "data-testid": "sheet-bar-rename-input"
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
              lineNumber: 126,
              columnNumber: 19
            },
            void 0
          ) : /* @__PURE__ */ I(je.Fragment, { children: [
            C && R && /* @__PURE__ */ I(
              "span",
              {
                className: Je["sheet-bar-item-color"],
                style: { backgroundColor: R },
                "data-testid": "sheet-bar-tab-color-item"
              },
              void 0,
              !1,
              {
                fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
                lineNumber: 138,
                columnNumber: 23
              },
              void 0
            ),
            /* @__PURE__ */ I(
              "span",
              {
                className: Je["sheet-bar-item-text"],
                "data-testid": L,
                children: v.name
              },
              void 0,
              !1,
              {
                fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
                lineNumber: 144,
                columnNumber: 21
              },
              void 0
            )
          ] }, void 0, !0, {
            fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
            lineNumber: 136,
            columnNumber: 19
          }, void 0)
        },
        v.sheetId,
        !1,
        {
          fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
          lineNumber: 111,
          columnNumber: 15
        },
        void 0
      );
    }) }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
      lineNumber: 97,
      columnNumber: 9
    }, void 0),
    /* @__PURE__ */ I(
      V,
      {
        onClick: m,
        type: "circle",
        className: Je["add-button"],
        testId: "sheet-bar-add-sheet",
        children: /* @__PURE__ */ I(Ae, { name: "plus" }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
          lineNumber: 162,
          columnNumber: 11
        }, void 0)
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
        lineNumber: 156,
        columnNumber: 9
      },
      void 0
    ),
    d >= 0 && /* @__PURE__ */ I(
      Uc,
      {
        position: d,
        sheetList: n,
        currentSheetId: f,
        hideMenu: p,
        editSheetName: E
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
        lineNumber: 165,
        columnNumber: 11
      },
      void 0
    ),
    e
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/SheetBar/index.tsx",
    lineNumber: 76,
    columnNumber: 7
  }, void 0);
});
_c.displayName = "SheetBarContainer";
const _e = {
  "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>`,
  "docProps/app.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
      <Application>Microsoft Excel</Application>
      <DocSecurity>0</DocSecurity>
      <ScaleCrop>false</ScaleCrop>
      <HeadingPairs>
        <vt:vector size="2" baseType="variant">
          <vt:variant>
            <vt:lpstr>工作表</vt:lpstr>
          </vt:variant>
          <vt:variant>
            <vt:i4>{size}</vt:i4>
          </vt:variant>
        </vt:vector>
      </HeadingPairs>
      <TitlesOfParts>
        <vt:vector size="{size}" baseType="lpstr">
          {children}
        </vt:vector>
      </TitlesOfParts>
      <Company></Company>
      <LinksUpToDate>false</LinksUpToDate>
      <SharedDoc>false</SharedDoc>
      <HyperlinksChanged>false</HyperlinksChanged>
      <AppVersion>16.0300</AppVersion>
    </Properties>`,
  "docProps/core.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
      xmlns:dc="http://purl.org/dc/elements/1.1/"
      xmlns:dcterms="http://purl.org/dc/terms/"
      xmlns:dcmitype="http://purl.org/dc/dcmitype/"
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <dc:creator>Xu Steve</dc:creator>
      <cp:lastModifiedBy>Xu Steve</cp:lastModifiedBy>
      <dcterms:created xsi:type="dcterms:W3CDTF">{children}</dcterms:created>
      <dcterms:modified xsi:type="dcterms:W3CDTF">{children}</dcterms:modified>
    </cp:coreProperties>`,
  "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      {children}
      {size}
    </Relationships>`,
  "xl/charts/_rels/chart1.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors{children}.xml"/>
      <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style{children}.xml"/>
    </Relationships>`,
  "xl/charts/chart1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">
      <c:date1904 val="0"/>
      <c:lang val="zh-CN"/>
      <c:roundedCorners val="0"/>
      <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:Choice Requires="c14"
          xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">
          <c14:style val="102"/>
        </mc:Choice>
        <mc:Fallback>
          <c:style val="2"/>
        </mc:Fallback>
      </mc:AlternateContent>
      <c:chart>
        <c:title>
          <c:tx>
            <c:rich>
              <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
              <a:lstStyle/>
              <a:p>
                <a:pPr>
                  <a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:lumMod val="65000"/>
                        <a:lumOff val="35000"/>
                      </a:schemeClr>
                    </a:solidFill>
                    <a:latin typeface="+mn-lt"/>
                    <a:ea typeface="+mn-ea"/>
                    <a:cs typeface="+mn-cs"/>
                  </a:defRPr>
                </a:pPr>
                <a:r>
                  <a:rPr lang="en-US" altLang="zh-CN"/>
                  <a:t>{size}</a:t>
                </a:r>
              </a:p>
              <a:p>
                <a:pPr>
                  <a:defRPr/>
                </a:pPr>
                <a:endParaRPr lang="zh-CN" altLang="en-US"/>
              </a:p>
            </c:rich>
          </c:tx>
          <c:overlay val="0"/>
          <c:spPr>
            <a:noFill/>
            <a:ln>
              <a:noFill/>
            </a:ln>
            <a:effectLst/>
          </c:spPr>
          <c:txPr>
            <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr sz="1400" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">
                  <a:solidFill>
                    <a:schemeClr val="tx1">
                      <a:lumMod val="65000"/>
                      <a:lumOff val="35000"/>
                    </a:schemeClr>
                  </a:solidFill>
                  <a:latin typeface="+mn-lt"/>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </a:pPr>
              <a:endParaRPr lang="zh-CN"/>
            </a:p>
          </c:txPr>
        </c:title>
        <c:autoTitleDeleted val="0"/>
        <c:plotArea>
          <c:layout/>
          <c:{large}Chart>
            <c:barDir val="col"/>
            <c:grouping val="clustered"/>
            <c:varyColors val="0"/>
            {children}
            <c:dLbls>
              <c:showLegendKey val="0"/>
              <c:showVal val="0"/>
              <c:showCatName val="0"/>
              <c:showSerName val="0"/>
              <c:showPercent val="0"/>
              <c:showBubbleSize val="0"/>
            </c:dLbls>
            <c:gapWidth val="219"/>
            <c:overlap val="-27"/>
            <c:axId val="152326191"/>
            <c:axId val="152327631"/>
          </c:barChart>
          <c:catAx>
            <c:axId val="152326191"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="b"/>
            <c:majorTickMark val="none"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:spPr>
              <a:noFill/>
              <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                <a:solidFill>
                  <a:schemeClr val="tx1">
                    <a:lumMod val="15000"/>
                    <a:lumOff val="85000"/>
                  </a:schemeClr>
                </a:solidFill>
                <a:round/>
              </a:ln>
              <a:effectLst/>
            </c:spPr>
            <c:txPr>
              <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
              <a:lstStyle/>
              <a:p>
                <a:pPr>
                  <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:lumMod val="65000"/>
                        <a:lumOff val="35000"/>
                      </a:schemeClr>
                    </a:solidFill>
                    <a:latin typeface="+mn-lt"/>
                    <a:ea typeface="+mn-ea"/>
                    <a:cs typeface="+mn-cs"/>
                  </a:defRPr>
                </a:pPr>
                <a:endParaRPr lang="zh-CN"/>
              </a:p>
            </c:txPr>
            <c:crossAx val="152327631"/>
            <c:crosses val="autoZero"/>
            <c:auto val="1"/>
            <c:lblAlgn val="ctr"/>
            <c:lblOffset val="100"/>
            <c:noMultiLvlLbl val="0"/>
          </c:catAx>
          <c:valAx>
            <c:axId val="152327631"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="l"/>
            <c:majorGridlines>
              <c:spPr>
                <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                  <a:solidFill>
                    <a:schemeClr val="tx1">
                      <a:lumMod val="15000"/>
                      <a:lumOff val="85000"/>
                    </a:schemeClr>
                  </a:solidFill>
                  <a:round/>
                </a:ln>
                <a:effectLst/>
              </c:spPr>
            </c:majorGridlines>
            <c:numFmt formatCode="General" sourceLinked="1"/>
            <c:majorTickMark val="none"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:spPr>
              <a:noFill/>
              <a:ln>
                <a:noFill/>
              </a:ln>
              <a:effectLst/>
            </c:spPr>
            <c:txPr>
              <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
              <a:lstStyle/>
              <a:p>
                <a:pPr>
                  <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
                    <a:solidFill>
                      <a:schemeClr val="tx1">
                        <a:lumMod val="65000"/>
                        <a:lumOff val="35000"/>
                      </a:schemeClr>
                    </a:solidFill>
                    <a:latin typeface="+mn-lt"/>
                    <a:ea typeface="+mn-ea"/>
                    <a:cs typeface="+mn-cs"/>
                  </a:defRPr>
                </a:pPr>
                <a:endParaRPr lang="zh-CN"/>
              </a:p>
            </c:txPr>
            <c:crossAx val="152326191"/>
            <c:crosses val="autoZero"/>
            <c:crossBetween val="between"/>
          </c:valAx>
          <c:spPr>
            <a:noFill/>
            <a:ln>
              <a:noFill/>
            </a:ln>
            <a:effectLst/>
          </c:spPr>
        </c:plotArea>
        <c:legend>
          <c:legendPos val="b"/>
          <c:overlay val="0"/>
          <c:spPr>
            <a:noFill/>
            <a:ln>
              <a:noFill/>
            </a:ln>
            <a:effectLst/>
          </c:spPr>
          <c:txPr>
            <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">
                  <a:solidFill>
                    <a:schemeClr val="tx1">
                      <a:lumMod val="65000"/>
                      <a:lumOff val="35000"/>
                    </a:schemeClr>
                  </a:solidFill>
                  <a:latin typeface="+mn-lt"/>
                  <a:ea typeface="+mn-ea"/>
                  <a:cs typeface="+mn-cs"/>
                </a:defRPr>
              </a:pPr>
              <a:endParaRPr lang="zh-CN"/>
            </a:p>
          </c:txPr>
        </c:legend>
        <c:plotVisOnly val="1"/>
        <c:dispBlanksAs val="gap"/>
        <c:extLst>
          <c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}"
            xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart">
            <c16r3:dataDisplayOptions16>
              <c16r3:dispNaAsBlank val="1"/>
            </c16r3:dataDisplayOptions16>
          </c:ext>
        </c:extLst>
        <c:showDLblsOverMax val="0"/>
      </c:chart>
      <c:spPr>
        <a:solidFill>
          <a:schemeClr val="bg1"/>
        </a:solidFill>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
        <a:effectLst/>
      </c:spPr>
      <c:txPr>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
          <a:pPr>
            <a:defRPr/>
          </a:pPr>
          <a:endParaRPr lang="zh-CN"/>
        </a:p>
      </c:txPr>
      <c:printSettings>
        <c:headerFooter/>
        <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
        <c:pageSetup/>
      </c:printSettings>
    </c:chartSpace>`,
  "xl/charts/colors1.xml": `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">
<a:schemeClr val="accent1"/>
<a:schemeClr val="accent2"/>
<a:schemeClr val="accent3"/>
<a:schemeClr val="accent4"/>
<a:schemeClr val="accent5"/>
<a:schemeClr val="accent6"/>
<cs:variation/>
<cs:variation>
  <a:lumMod val="60000"/>
</cs:variation>
<cs:variation>
  <a:lumMod val="80000"/>
  <a:lumOff val="20000"/>
</cs:variation>
<cs:variation>
  <a:lumMod val="80000"/>
</cs:variation>
<cs:variation>
  <a:lumMod val="60000"/>
  <a:lumOff val="40000"/>
</cs:variation>
<cs:variation>
  <a:lumMod val="50000"/>
</cs:variation>
<cs:variation>
  <a:lumMod val="70000"/>
  <a:lumOff val="30000"/>
</cs:variation>
<cs:variation>
  <a:lumMod val="70000"/>
</cs:variation>
<cs:variation>
  <a:lumMod val="50000"/>
  <a:lumOff val="50000"/>
</cs:variation>
</cs:colorStyle>`,
  "xl/charts/style1.xml": `<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="201">
    <cs:axisTitle>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="1000" kern="1200"/>
    </cs:axisTitle>
    <cs:categoryAxis>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
      <cs:defRPr sz="900" kern="1200"/>
    </cs:categoryAxis>
    <cs:chartArea mods="allowNoFillOverride allowNoLineOverride">
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="bg1"/>
        </a:solidFill>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
      <cs:defRPr sz="1000" kern="1200"/>
    </cs:chartArea>
    <cs:dataLabel>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="75000"/>
          <a:lumOff val="25000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="900" kern="1200"/>
    </cs:dataLabel>
    <cs:dataLabelCallout>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="dk1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="lt1"/>
        </a:solidFill>
        <a:ln>
          <a:solidFill>
            <a:schemeClr val="dk1">
              <a:lumMod val="25000"/>
              <a:lumOff val="75000"/>
            </a:schemeClr>
          </a:solidFill>
        </a:ln>
      </cs:spPr>
      <cs:defRPr sz="900" kern="1200"/>
      <cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1">
        <a:spAutoFit/>
      </cs:bodyPr>
    </cs:dataLabelCallout>
    <cs:dataPoint>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="1">
        <cs:styleClr val="auto"/>
      </cs:fillRef>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
    </cs:dataPoint>
    <cs:dataPoint3D>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="1">
        <cs:styleClr val="auto"/>
      </cs:fillRef>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
    </cs:dataPoint3D>
    <cs:dataPointLine>
      <cs:lnRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:lnRef>
      <cs:fillRef idx="1"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="28575" cap="rnd">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:dataPointLine>
    <cs:dataPointMarker>
      <cs:lnRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:lnRef>
      <cs:fillRef idx="1">
        <cs:styleClr val="auto"/>
      </cs:fillRef>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
        </a:ln>
      </cs:spPr>
    </cs:dataPointMarker>
    <cs:dataPointMarkerLayout symbol="circle" size="5"/>
    <cs:dataPointWireframe>
      <cs:lnRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:lnRef>
      <cs:fillRef idx="1"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="rnd">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:dataPointWireframe>
    <cs:dataTable>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:spPr>
        <a:noFill/>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
      <cs:defRPr sz="900" kern="1200"/>
    </cs:dataTable>
    <cs:downBar>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="dk1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="dk1">
            <a:lumMod val="65000"/>
            <a:lumOff val="35000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:ln w="9525">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="65000"/>
              <a:lumOff val="35000"/>
            </a:schemeClr>
          </a:solidFill>
        </a:ln>
      </cs:spPr>
    </cs:downBar>
    <cs:dropLine>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="35000"/>
              <a:lumOff val="65000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:dropLine>
    <cs:errorBar>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="65000"/>
              <a:lumOff val="35000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:errorBar>
    <cs:floor>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:noFill/>
        <a:ln>
          <a:noFill/>
        </a:ln>
      </cs:spPr>
    </cs:floor>
    <cs:gridlineMajor>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:gridlineMajor>
    <cs:gridlineMinor>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="5000"/>
              <a:lumOff val="95000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:gridlineMinor>
    <cs:hiLoLine>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="75000"/>
              <a:lumOff val="25000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:hiLoLine>
    <cs:leaderLine>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="35000"/>
              <a:lumOff val="65000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:leaderLine>
    <cs:legend>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="900" kern="1200"/>
    </cs:legend>
    <cs:plotArea mods="allowNoFillOverride allowNoLineOverride">
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
    </cs:plotArea>
    <cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride">
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
    </cs:plotArea3D>
    <cs:seriesAxis>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="900" kern="1200"/>
    </cs:seriesAxis>
    <cs:seriesLine>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="35000"/>
              <a:lumOff val="65000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:round/>
        </a:ln>
      </cs:spPr>
    </cs:seriesLine>
    <cs:title>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="1400" b="0" kern="1200" spc="0" baseline="0"/>
    </cs:title>
    <cs:trendline>
      <cs:lnRef idx="0">
        <cs:styleClr val="auto"/>
      </cs:lnRef>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:ln w="19050" cap="rnd">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="sysDot"/>
        </a:ln>
      </cs:spPr>
    </cs:trendline>
    <cs:trendlineLabel>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="900" kern="1200"/>
    </cs:trendlineLabel>
    <cs:upBar>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="dk1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:solidFill>
          <a:schemeClr val="lt1"/>
        </a:solidFill>
        <a:ln w="9525">
          <a:solidFill>
            <a:schemeClr val="tx1">
              <a:lumMod val="15000"/>
              <a:lumOff val="85000"/>
            </a:schemeClr>
          </a:solidFill>
        </a:ln>
      </cs:spPr>
    </cs:upBar>
    <cs:valueAxis>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </cs:fontRef>
      <cs:defRPr sz="900" kern="1200"/>
    </cs:valueAxis>
    <cs:wall>
      <cs:lnRef idx="0"/>
      <cs:fillRef idx="0"/>
      <cs:effectRef idx="0"/>
      <cs:fontRef idx="minor">
        <a:schemeClr val="tx1"/>
      </cs:fontRef>
      <cs:spPr>
        <a:noFill/>
        <a:ln>
          <a:noFill/>
        </a:ln>
      </cs:spPr>
    </cs:wall>
  </cs:chartStyle>`,
  "xl/drawings/_rels/drawing1.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      {children}
    </Relationships>`,
  "xl/drawings/drawing1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      {children}
    </xdr:wsDr>`,
  "xl/theme/theme1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题​​">
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:dk1>
            <a:sysClr val="windowText" lastClr="000000"/>
          </a:dk1>
          <a:lt1>
            <a:sysClr val="window" lastClr="FFFFFF"/>
          </a:lt1>
          <a:dk2>
            <a:srgbClr val="44546A"/>
          </a:dk2>
          <a:lt2>
            <a:srgbClr val="E7E6E6"/>
          </a:lt2>
          <a:accent1>
            <a:srgbClr val="4472C4"/>
          </a:accent1>
          <a:accent2>
            <a:srgbClr val="ED7D31"/>
          </a:accent2>
          <a:accent3>
            <a:srgbClr val="A5A5A5"/>
          </a:accent3>
          <a:accent4>
            <a:srgbClr val="FFC000"/>
          </a:accent4>
          <a:accent5>
            <a:srgbClr val="5B9BD5"/>
          </a:accent5>
          <a:accent6>
            <a:srgbClr val="70AD47"/>
          </a:accent6>
          <a:hlink>
            <a:srgbClr val="0563C1"/>
          </a:hlink>
          <a:folHlink>
            <a:srgbClr val="954F72"/>
          </a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
          <a:majorFont>
            <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
            <a:font script="Jpan" typeface="游ゴシック Light"/>
            <a:font script="Hang" typeface="맑은 고딕"/>
            <a:font script="Hans" typeface="等线 Light"/>
            <a:font script="Hant" typeface="新細明體"/>
            <a:font script="Arab" typeface="Times New Roman"/>
            <a:font script="Hebr" typeface="Times New Roman"/>
            <a:font script="Thai" typeface="Tahoma"/>
            <a:font script="Ethi" typeface="Nyala"/>
            <a:font script="Beng" typeface="Vrinda"/>
            <a:font script="Gujr" typeface="Shruti"/>
            <a:font script="Khmr" typeface="MoolBoran"/>
            <a:font script="Knda" typeface="Tunga"/>
            <a:font script="Guru" typeface="Raavi"/>
            <a:font script="Cans" typeface="Euphemia"/>
            <a:font script="Cher" typeface="Plantagenet Cherokee"/>
            <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
            <a:font script="Tibt" typeface="Microsoft Himalaya"/>
            <a:font script="Thaa" typeface="MV Boli"/>
            <a:font script="Deva" typeface="Mangal"/>
            <a:font script="Telu" typeface="Gautami"/>
            <a:font script="Taml" typeface="Latha"/>
            <a:font script="Syrc" typeface="Estrangelo Edessa"/>
            <a:font script="Orya" typeface="Kalinga"/>
            <a:font script="Mlym" typeface="Kartika"/>
            <a:font script="Laoo" typeface="DokChampa"/>
            <a:font script="Sinh" typeface="Iskoola Pota"/>
            <a:font script="Mong" typeface="Mongolian Baiti"/>
            <a:font script="Viet" typeface="Times New Roman"/>
            <a:font script="Uigh" typeface="Microsoft Uighur"/>
            <a:font script="Geor" typeface="Sylfaen"/>
            <a:font script="Armn" typeface="Arial"/>
            <a:font script="Bugi" typeface="Leelawadee UI"/>
            <a:font script="Bopo" typeface="Microsoft JhengHei"/>
            <a:font script="Java" typeface="Javanese Text"/>
            <a:font script="Lisu" typeface="Segoe UI"/>
            <a:font script="Mymr" typeface="Myanmar Text"/>
            <a:font script="Nkoo" typeface="Ebrima"/>
            <a:font script="Olck" typeface="Nirmala UI"/>
            <a:font script="Osma" typeface="Ebrima"/>
            <a:font script="Phag" typeface="Phagspa"/>
            <a:font script="Syrn" typeface="Estrangelo Edessa"/>
            <a:font script="Syrj" typeface="Estrangelo Edessa"/>
            <a:font script="Syre" typeface="Estrangelo Edessa"/>
            <a:font script="Sora" typeface="Nirmala UI"/>
            <a:font script="Tale" typeface="Microsoft Tai Le"/>
            <a:font script="Talu" typeface="Microsoft New Tai Lue"/>
            <a:font script="Tfng" typeface="Ebrima"/>
          </a:majorFont>
          <a:minorFont>
            <a:latin typeface="Calibri" panose="020F0502020204030204"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
            <a:font script="Jpan" typeface="游ゴシック"/>
            <a:font script="Hang" typeface="맑은 고딕"/>
            <a:font script="Hans" typeface="等线"/>
            <a:font script="Hant" typeface="新細明體"/>
            <a:font script="Arab" typeface="Arial"/>
            <a:font script="Hebr" typeface="Arial"/>
            <a:font script="Thai" typeface="Tahoma"/>
            <a:font script="Ethi" typeface="Nyala"/>
            <a:font script="Beng" typeface="Vrinda"/>
            <a:font script="Gujr" typeface="Shruti"/>
            <a:font script="Khmr" typeface="DaunPenh"/>
            <a:font script="Knda" typeface="Tunga"/>
            <a:font script="Guru" typeface="Raavi"/>
            <a:font script="Cans" typeface="Euphemia"/>
            <a:font script="Cher" typeface="Plantagenet Cherokee"/>
            <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
            <a:font script="Tibt" typeface="Microsoft Himalaya"/>
            <a:font script="Thaa" typeface="MV Boli"/>
            <a:font script="Deva" typeface="Mangal"/>
            <a:font script="Telu" typeface="Gautami"/>
            <a:font script="Taml" typeface="Latha"/>
            <a:font script="Syrc" typeface="Estrangelo Edessa"/>
            <a:font script="Orya" typeface="Kalinga"/>
            <a:font script="Mlym" typeface="Kartika"/>
            <a:font script="Laoo" typeface="DokChampa"/>
            <a:font script="Sinh" typeface="Iskoola Pota"/>
            <a:font script="Mong" typeface="Mongolian Baiti"/>
            <a:font script="Viet" typeface="Arial"/>
            <a:font script="Uigh" typeface="Microsoft Uighur"/>
            <a:font script="Geor" typeface="Sylfaen"/>
            <a:font script="Armn" typeface="Arial"/>
            <a:font script="Bugi" typeface="Leelawadee UI"/>
            <a:font script="Bopo" typeface="Microsoft JhengHei"/>
            <a:font script="Java" typeface="Javanese Text"/>
            <a:font script="Lisu" typeface="Segoe UI"/>
            <a:font script="Mymr" typeface="Myanmar Text"/>
            <a:font script="Nkoo" typeface="Ebrima"/>
            <a:font script="Olck" typeface="Nirmala UI"/>
            <a:font script="Osma" typeface="Ebrima"/>
            <a:font script="Phag" typeface="Phagspa"/>
            <a:font script="Syrn" typeface="Estrangelo Edessa"/>
            <a:font script="Syrj" typeface="Estrangelo Edessa"/>
            <a:font script="Syre" typeface="Estrangelo Edessa"/>
            <a:font script="Sora" typeface="Nirmala UI"/>
            <a:font script="Tale" typeface="Microsoft Tai Le"/>
            <a:font script="Talu" typeface="Microsoft New Tai Lue"/>
            <a:font script="Tfng" typeface="Ebrima"/>
          </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
          <a:fillStyleLst>
            <a:solidFill>
              <a:schemeClr val="phClr"/>
            </a:solidFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0">
                  <a:schemeClr val="phClr">
                    <a:lumMod val="110000"/>
                    <a:satMod val="105000"/>
                    <a:tint val="67000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="50000">
                  <a:schemeClr val="phClr">
                    <a:lumMod val="105000"/>
                    <a:satMod val="103000"/>
                    <a:tint val="73000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="100000">
                  <a:schemeClr val="phClr">
                    <a:lumMod val="105000"/>
                    <a:satMod val="109000"/>
                    <a:tint val="81000"/>
                  </a:schemeClr>
                </a:gs>
              </a:gsLst>
              <a:lin ang="5400000" scaled="0"/>
            </a:gradFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0">
                  <a:schemeClr val="phClr">
                    <a:satMod val="103000"/>
                    <a:lumMod val="102000"/>
                    <a:tint val="94000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="50000">
                  <a:schemeClr val="phClr">
                    <a:satMod val="110000"/>
                    <a:lumMod val="100000"/>
                    <a:shade val="100000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="100000">
                  <a:schemeClr val="phClr">
                    <a:lumMod val="99000"/>
                    <a:satMod val="120000"/>
                    <a:shade val="78000"/>
                  </a:schemeClr>
                </a:gs>
              </a:gsLst>
              <a:lin ang="5400000" scaled="0"/>
            </a:gradFill>
          </a:fillStyleLst>
          <a:lnStyleLst>
            <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:schemeClr val="phClr"/>
              </a:solidFill>
              <a:prstDash val="solid"/>
              <a:miter lim="800000"/>
            </a:ln>
            <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:schemeClr val="phClr"/>
              </a:solidFill>
              <a:prstDash val="solid"/>
              <a:miter lim="800000"/>
            </a:ln>
            <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:schemeClr val="phClr"/>
              </a:solidFill>
              <a:prstDash val="solid"/>
              <a:miter lim="800000"/>
            </a:ln>
          </a:lnStyleLst>
          <a:effectStyleLst>
            <a:effectStyle>
              <a:effectLst/>
            </a:effectStyle>
            <a:effectStyle>
              <a:effectLst/>
            </a:effectStyle>
            <a:effectStyle>
              <a:effectLst>
                <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                  <a:srgbClr val="000000">
                    <a:alpha val="63000"/>
                  </a:srgbClr>
                </a:outerShdw>
              </a:effectLst>
            </a:effectStyle>
          </a:effectStyleLst>
          <a:bgFillStyleLst>
            <a:solidFill>
              <a:schemeClr val="phClr"/>
            </a:solidFill>
            <a:solidFill>
              <a:schemeClr val="phClr">
                <a:tint val="95000"/>
                <a:satMod val="170000"/>
              </a:schemeClr>
            </a:solidFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0">
                  <a:schemeClr val="phClr">
                    <a:tint val="93000"/>
                    <a:satMod val="150000"/>
                    <a:shade val="98000"/>
                    <a:lumMod val="102000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="50000">
                  <a:schemeClr val="phClr">
                    <a:tint val="98000"/>
                    <a:satMod val="130000"/>
                    <a:shade val="90000"/>
                    <a:lumMod val="103000"/>
                  </a:schemeClr>
                </a:gs>
                <a:gs pos="100000">
                  <a:schemeClr val="phClr">
                    <a:shade val="63000"/>
                    <a:satMod val="120000"/>
                  </a:schemeClr>
                </a:gs>
              </a:gsLst>
              <a:lin ang="5400000" scaled="0"/>
            </a:gradFill>
          </a:bgFillStyleLst>
        </a:fmtScheme>
      </a:themeElements>
      <a:objectDefaults/>
      <a:extraClrSchemeLst/>
      <a:extLst>
        <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
          <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>
        </a:ext>
      </a:extLst>
    </a:theme>`,
  "xl/worksheets/_rels/sheet1.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/{children}"/>
    </Relationships>`,
  "xl/worksheets/sheet1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3"
      xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
      xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
      xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
      xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{9B2CBB7B-9BFB-4EF9-A18C-0B0DED5460D3}">
      {larger}
      <dimension ref="A1:D5"/>
      <sheetViews>
        {large}
      </sheetViews>
      <sheetFormatPr defaultRowHeight="13.9" x14ac:dyDescent="0.4"/>
      {size}
      {children}
      <phoneticPr fontId="0" type="noConversion"/>
      <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
    </worksheet>`,
  "xl/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr"
      xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
      xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"
      xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
      {children}
      <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0">
          <alignment vertical="center"/>
        </xf>
      </cellStyleXfs>
      <cellStyles count="1">
        <cellStyle name="常规" xfId="0" builtinId="0"/>
      </cellStyles>
      <dxfs count="0"/>
      <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
      <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"
          xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
          <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
        </ext>
        <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}"
          xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
          <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/>
        </ext>
      </extLst>
    </styleSheet>`,
  "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2"
      xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
      xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
      xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"
      xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"
      xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
      <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27328"/>
      <workbookPr defaultThemeVersion="166925"/>
      <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:Choice Requires="x15">
          <x15ac:absPath url="C:UsersstevexuDesktop"
            xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/>
        </mc:Choice>
      </mc:AlternateContent>
      <xr:revisionPtr revIDLastSave="0" documentId="13_ncr:1_{39800FFC-1B5B-45A9-B956-6FB55C475C9F}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>
      <bookViews>
      <workbookView xWindow="-98" yWindow="-98" windowWidth="19396" windowHeight="11475" {size} xr2:uid="{11AE31A3-10C3-4738-A21D-56F9E9832A43}"/>
      </bookViews>
      <sheets>
        {children}
      </sheets>
      {large}
      <calcPr calcId="191029"/>
      <extLst>
        <ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}"
          xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
          <x15:workbookPr chartTrackingRefBase="1"/>
        </ext>
      </extLst>
    </workbook>`,
  "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      {children}
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
      <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
      <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    </Types>`
}, Ho = "xl", Yv = "xl/styles.xml", Xv = "xl/workbook.xml", Kv = "xl/_rels/workbook.xml.rels", Jv = "xl/theme/theme1.xml", Qv = "xl/sharedStrings.xml", Vt = "#text", ea = "xl/drawings/", ta = "../drawings/", Bc = 8;
function jv(e) {
  const t = parseFloat(e);
  return isNaN(t) ? 0 : Math.floor(t * Bc);
}
function Zv(e) {
  return Math.floor(e / Bc);
}
const zc = [
  // 'area',
  "bar",
  // 'bubble',
  // 'doughnut',
  "line",
  "pie"
  // 'ofPie',
  // 'radar',
  // 'scatter',
  // 'surface',
];
function ke(e, t, n) {
  const r = tt(e, t, n);
  return Array.isArray(r) ? r : [r];
}
function gi(e) {
  let t = {};
  if (e.nodeType === Node.ELEMENT_NODE) {
    if (e.attributes.length > 0)
      for (const n of e.attributes)
        t[n.nodeName] = n.nodeValue;
  } else e.nodeType === Node.TEXT_NODE && (t = e.nodeValue);
  if (e.childNodes.length > 0)
    for (const n of e.childNodes) {
      const r = n.nodeName;
      typeof n.nodeValue == "string" && !n.nodeValue.trim() || (typeof t[r] > "u" ? t[r] = gi(n) : (typeof t[r].push > "u" && (t[r] = [t[r]]), t[r].push(gi(n))));
    }
  return t;
}
function e6(e) {
  const n = new DOMParser().parseFromString(e, "text/xml");
  return gi(n);
}
function Vo(e) {
  if (!e)
    return "";
  const t = "#";
  return e.length === 6 ? t + e : e.length === 8 && e.startsWith("FF") ? t + e.slice(2) : "";
}
function tr(e, t) {
  if (!t)
    return "";
  if (t.theme) {
    const n = [
      "a:lt1",
      "a:dk1",
      "a:lt2",
      "a:dk2",
      "a:accent1",
      "a:accent2",
      "a:accent3",
      "a:accent4",
      "a:accent5",
      "a:accent6"
    ], r = parseInt(t.theme, 10);
    if (r >= 0 && r <= 1)
      return Vo(e[n[r]]?.["a:sysClr"]?.lastClr);
    if (r > 1 && r < n.length)
      return Vo(e[n[r]]?.["a:srgbClr"]?.val);
  }
  return Vo(t.rgb);
}
function Lr(e, t) {
  if (!e?.style)
    return;
  const n = e?.style;
  if (typeof oo[n] != "number")
    return;
  const r = tr(t, e?.color);
  return {
    type: n,
    color: r
  };
}
function t6(e, t, n) {
  const r = {}, o = ke(e, "styleSheet.cellXfs.xf", []), s = ke(e, "styleSheet.cellStyles.xf", []), l = o.length > 0 ? o : s;
  if (!t || l.length === 0 || !l[t])
    return r;
  const f = l[t];
  if (f.applyAlignment && f.alignment) {
    if (f.alignment.horizontal) {
      const d = {
        left: xe.LEFT,
        center: xe.CENTER,
        right: xe.RIGHT
      };
      r.horizontalAlign = d[f.alignment.horizontal];
    }
    if (f.alignment.vertical) {
      const d = {
        top: Te.TOP,
        center: Te.MIDDLE,
        bottom: Te.BOTTOM
      };
      r.verticalAlign = d[f.alignment.vertical];
    } else
      r.verticalAlign = Te.BOTTOM;
    r.isWrapText = !!f.alignment.wrapText;
  }
  if (f.applyFont && f.fontId) {
    const d = ke(e, "styleSheet.fonts.font", []), c = parseInt(f.fontId, 10);
    if (d[c]) {
      const i = d[c], a = i?.sz?.val ? parseInt(i?.sz?.val, 10) : void 0;
      r.fontSize = a || void 0, r.isBold = !!i?.b, r.isItalic = !!i?.i, r.isStrike = !!i?.strike, r.underline = i.u ? Ce.SINGLE : Ce.NONE, r.fontFamily = i?.name?.val;
      const u = tr(n, i.color);
      u && (r.fontColor = u);
    }
  }
  if (f.applyNumberFormat && f.numFmtId) {
    const c = ke(e, "styleSheet.numFmts.numFmt", []).find((i) => i.numFmtId === f.numFmtId);
    if (c)
      r.numberFormat = c.formatCode;
    else {
      const i = parseInt(f.numFmtId, 10), a = Ei.find((u) => u.id === i);
      a && (r.numberFormat = a.formatCode);
    }
  }
  if (f.applyFill && f.fillId) {
    const d = ke(e, "styleSheet.fills.fill", []), c = parseInt(f.fillId, 10);
    if (d[c]) {
      const i = d[c].gradientFill, a = d[c].patternFill;
      if (i && i.stop[0]) {
        const u = tr(n, i.stop[0].color);
        u && (r.fillColor = u);
      } else if (a) {
        const u = tr(n, a.fgColor);
        u && (r.fillColor = u);
      }
    }
  }
  if (f.applyBorder && f.borderId) {
    const d = ke(
      e,
      "styleSheet.borders.border",
      []
    ), c = parseInt(f.borderId, 10);
    if (d[c]) {
      const i = d[c];
      r.borderLeft = Lr(i.left, n), r.borderRight = Lr(i.right, n), r.borderTop = Lr(i.top, n), r.borderBottom = Lr(i.bottom, n);
    }
  }
  return r;
}
function n6(e, t) {
  const n = e[Xv], r = ke(
    e[Qv],
    "sst.si",
    []
  ), o = tt(
    e[Jv],
    "a:theme.a:themeElements.a:clrScheme",
    {}
  ), s = {
    workbook: {},
    mergeCells: {},
    customHeight: {},
    customWidth: {},
    definedNames: {},
    currentSheetId: "",
    drawings: {},
    rangeMap: {},
    worksheets: {},
    autoFilter: {},
    scroll: {}
  }, l = ke(
    e[Kv],
    "Relationships.Relationship",
    []
  ), f = {}, d = {};
  let c = 0;
  const i = ke(
    n,
    "workbook.sheets.sheet",
    []
  );
  let a = 0;
  for (const m of i) {
    if (!m)
      continue;
    const p = l.find((x) => x.Id === m["r:id"])?.Target || "", E = "worksheets/", b = p.slice(E.length), N = `${Ho}/${E}_rels/${b}.rels`;
    if (e[N]) {
      const x = ke(
        e[N],
        "Relationships.Relationship",
        []
      );
      for (const L of x)
        d[m.sheetId] || (d[m.sheetId] = []), d[m.sheetId].push(L.Target), L.Target.startsWith(ta) && c++;
    }
    const v = `${Ho}/${p}`;
    f[m.sheetId] = v;
    const C = Ot(
      tt(
        e[v],
        "worksheet.sheetViews.sheetView.selection.sqref",
        ""
      )
    ), M = tt(
      e[v],
      "worksheet.sheetViews.sheetView.tabSelected",
      ""
    ), R = tr(
      o,
      tt(e[v], "worksheet.sheetPr.tabColor", {
        rgb: ""
      })
    );
    M === "1" && (s.currentSheetId = m.sheetId);
    const A = {
      sheetId: m.sheetId,
      name: m.name,
      isHide: m.state === "hidden",
      rowCount: 200,
      colCount: 200,
      sort: a++,
      tabColor: R === "" ? void 0 : R
    };
    s.workbook[A.sheetId] = A, C && (C.sheetId = A.sheetId, s.rangeMap[m.sheetId] = C.toIRange());
  }
  const u = Object.values(s.workbook);
  u.sort((m, p) => m.sort - p.sort), s.currentSheetId = s.currentSheetId || u?.[0]?.sheetId;
  const h = (m) => Object.values(s.workbook).find((E) => E.name === m)?.sheetId || "";
  for (const m of u) {
    const p = f[m.sheetId], E = ke(
      e[p],
      "worksheet.sheetData.row",
      []
    ), b = ke(
      e[p],
      "worksheet.mergeCells.mergeCell",
      []
    );
    for (const R of b) {
      const A = Ot(R.ref, h);
      if (A) {
        A.sheetId = m.sheetId;
        const x = `${m.name}!${R.ref}`;
        s.mergeCells[x] = A.toIRange();
      }
    }
    const N = ke(
      e[p],
      "worksheet.cols.col",
      []
    ), v = tt(e[p], "worksheet.sheetFormatPr", {
      defaultColWidth: "",
      defaultRowHeight: "",
      outlineLevelRow: ""
    });
    if (N.length > 0) {
      for (const R of N)
        if (R && R.customWidth && R.width && R.min && R.max) {
          const x = v.defaultColWidth === R.width ? rr : jv(R.width), L = !!R.hidden;
          for (let D = parseInt(R.min, 10) - 1, W = parseInt(R.max, 10); D < W; D++)
            s.customWidth[bt(m.sheetId, D)] = {
              len: x,
              isHide: L
            };
        }
    }
    if (E.length === 0)
      continue;
    let { colCount: C } = m, { rowCount: M } = m;
    for (const R of E) {
      if (!R)
        continue;
      const A = parseInt(R.r, 10) - 1;
      if (M = Math.max(M, A + 1), M > mr)
        continue;
      if (R.customHeight && R.ht) {
        const L = v.defaultRowHeight === R.ht;
        s.customHeight[bt(m.sheetId, A)] = {
          len: L ? nr : Math.floor(parseInt(R.ht, 10)),
          isHide: !!R.hidden
        };
      }
      const x = Array.isArray(R.c) ? R.c : [R.c];
      for (const L of x) {
        if (!L)
          continue;
        const D = Ot(L.r);
        if (!D || (C = Math.max(C, D.col + 1), C > hr))
          continue;
        const W = L?.v?.[Vt] ?? "", J = parseInt(L.s, 10), K = {
          ...t6(e[Yv], J, o),
          value: W
        }, q = L?.f?.[Vt] || "";
        if (q && (q.startsWith(Fn) ? K.formula = q : K.formula = Fn + q), L.t === "s") {
          const H = parseInt(W, 10);
          if (!isNaN(H)) {
            const Y = r[H];
            K.value = Y?.t?.[Vt] ?? "";
          }
        }
        L.t === "b" && (K.value = W === "1"), s.worksheets[nt(m.sheetId, A, D.col)] = K, C = Math.max(C, D.col + 1);
      }
    }
    m.rowCount = Math.max(m.rowCount, M), m.colCount = Math.max(m.colCount, C);
  }
  const g = ke(
    n,
    "workbook.definedNames.definedName",
    []
  );
  for (const m of g) {
    const p = Ot(m[Vt], h);
    p && m?.name && (s.definedNames[m.name.toLowerCase()] = p.toIRange());
  }
  for (let m = 1; m <= c; m++) {
    const p = `drawing${m}.xml`, E = `${ea}${p}`, b = `${ea}_rels/${p}.rels`;
    if (!e[E] || !e[b])
      break;
    const N = ke(
      e[b],
      "Relationships.Relationship",
      []
    );
    let v = "";
    for (const [M, R] of Object.entries(d))
      if (R.some((A) => A === ta + p)) {
        v = M;
        break;
      }
    const C = ke(
      e[E],
      "xdr:wsDr.xdr:twoCellAnchor",
      []
    );
    for (const M of C) {
      const R = tt(
        M,
        "xdr:graphicFrame.a:graphic.a:graphicData.c:chart.r:id",
        ""
      ), A = tt(M, "xdr:pic.xdr:blipFill.a:blip.r:embed", ""), x = !!A, L = x ? A : R, D = N.find((H) => H.Id === L)?.Target || "";
      if (!D)
        continue;
      const W = Ho + D.slice(2), J = M["xdr:from"]["xdr:col"][Vt], ne = M["xdr:from"]["xdr:row"][Vt];
      if (!J || !ne)
        continue;
      const K = String(m), q = {
        title: "",
        type: x ? "floating-picture" : "chart",
        uuid: K,
        width: 300,
        height: 300,
        originHeight: 300,
        originWidth: 300,
        fromCol: parseInt(ne, 10),
        fromRow: parseInt(J, 10),
        sheetId: v,
        marginX: 0,
        marginY: 0
      };
      if (x) {
        if (!e[W])
          continue;
        const H = tt(M, "xdr:pic.xdr:nvPicPr.xdr:cNvPr.name", ""), Y = tt(M, "xdr:pic.xdr:nvPicPr.xdr:cNvPr.title", "");
        q.title = Y || H, q.imageAngle = 0, q.imageSrc = e[W] || "";
        const te = t[W];
        te && te.width > 0 && te.height > 0 && (q.width = te.width, q.height = te.height, q.originWidth = te.width, q.originHeight = te.height);
      } else {
        q.title = tt(
          e[W],
          "c:chartSpace.c:chart.c:title.c:tx.c:rich.a:p.0.a:r.a:t." + Vt,
          ""
        );
        let H = [];
        for (const X of zc) {
          const le = tt(
            e[W],
            `c:chartSpace.c:chart.c:plotArea.c:${X}Chart`,
            ""
          );
          if (le) {
            q.chartType = X, H = ke(le, "c:ser", []).map((j) => tt(j, `c:val.c:numRef.c:f.${Vt}`, "")).filter((j) => j);
            break;
          }
        }
        if (H.length === 0)
          continue;
        const Y = (X) => Object.values(s.workbook).find(
          ($) => $.name === X
        )?.sheetId || "";
        let te;
        for (const X of H) {
          const le = Ot(X, Y);
          le && (te ? te = Ii(te, le) : te = le);
        }
        if (!te)
          continue;
        te.sheetId = te.sheetId || v, q.chartRange = te.toIRange();
      }
      s.drawings[K] = q;
    }
  }
  return s;
}
async function r6(e) {
  const n = await (await import("./jszip.min-COPV1wbk.mjs").then((f) => f.j)).default.loadAsync(e), { files: r } = n, o = {}, s = {};
  for (const f of Object.keys(r))
    if (!r[f].dir)
      if (f.includes(".xml")) {
        const d = await r[f].async("string");
        d && (o[f] = e6(d));
      } else {
        let d = "";
        for (const [i, a] of Object.entries(La))
          if (a.some((u) => f.endsWith(u))) {
            d = i;
            break;
          }
        if (!d)
          continue;
        const c = await r[f].async("base64");
        c && (o[f] = `data:${d};base64,${c}`, s[f] = await yi(o[f]));
      }
  return n6(o, s);
}
const Fr = "/";
function ro(e) {
  return `FF${ti(e).slice(1, -2)}`;
}
function Pr(e, t = "left") {
  if (!e)
    return `<${t}/>`;
  const n = e.color || Ee("black");
  let r = "";
  return n === Ee("black") ? r = 'auto="1"' : r = `rgb="${ro(n)}"`, `<${t} style="${e.type}">
  <color ${r}/>
</${t}>`;
}
function o6(e, t) {
  if (!t || On(t))
    return;
  const n = [], r = {
    fontId: "0",
    fillId: "0",
    numFmtId: "0",
    borderId: "0"
  };
  t.fillColor && (r.fillId = String(e.fills.length), n.push('applyFill="1"'), e.fills.push(`<fill>
    <patternFill patternType="solid">
      <fgColor rgb="${ro(t.fillColor)}"/>
      <bgColor indexed="64"/>
    </patternFill>
  </fill>`));
  const o = [];
  if ((t.underline === Ce.SINGLE || t.underline === Ce.DOUBLE) && o.push("<u/>"), t.isBold && o.push("<b/>"), t.isItalic && o.push("<i/>"), t.isStrike && o.push("<strike/>"), typeof t.fontSize < "u" && o.push(`<sz val="${t.fontSize}"/>`), t.fontColor && o.push(`<color rgb="${ro(t.fontColor)}"/>`), t.fontFamily && o.push(`<name val="${t.fontFamily}"/>`), o.length > 0 && (r.fontId = String(e.fonts.length), n.push('applyFont="1"'), e.fonts.push(
    `<font>
${o.join(
      `
`
    )}
<charset val="0"/>
<scheme val="minor"/>
</font>`
  )), t.numberFormat) {
    const d = Ei.find(
      (c) => c.formatCode === t.numberFormat
    );
    d && (n.push('applyNumberFormat="1"'), r.numFmtId = String(d.id));
  }
  let s = '<alignment vertical="center"/>';
  const l = [];
  if (n.push('applyAlignment="1"'), t.isWrapText && l.push('wrapText="1"'), t.horizontalAlign !== void 0) {
    const d = {
      [xe.LEFT]: "left",
      [xe.CENTER]: "center",
      [xe.RIGHT]: "right"
    };
    l.push(`horizontal="${d[t.horizontalAlign]}"`);
  }
  if (t.verticalAlign !== void 0) {
    const d = {
      [Te.TOP]: "top",
      [Te.MIDDLE]: "center",
      [Te.BOTTOM]: "bottom"
    };
    t.verticalAlign !== Te.BOTTOM && l.push(`vertical="${d[t.verticalAlign]}"`);
  }
  l.length > 0 && (s = `<alignment ${l.join(" ")}/>`), (t.borderBottom || t.borderLeft || t.borderTop || t.borderRight) && (r.borderId = String(e.borders.length), n.push('applyBorder="1"'), e.borders.push(`<border>
    ${Pr(t.borderLeft, "left")}
    ${Pr(t.borderRight, "right")}
    ${Pr(t.borderTop, "top")}
    ${Pr(t.borderBottom, "bottom")}
    <diagonal/>
  </border>`));
  const f = `<xf numFmtId="${r.numFmtId}" fontId="${r.fontId}" fillId="${r.fillId}" borderId="${r.borderId}" xfId="0" ${n.join(" ")}>
${s}
</xf>`;
  e.cellXfs.push(f);
}
function lt(e, t = {}) {
  return e.replace(/{([a-z]+)}/gi, function(r, o) {
    if (o in t)
      return t[o];
    throw new Error(`compileTemplate not found key: "${o}"`);
  });
}
function i6(e) {
  return e.includes(" ") ? `'${e}'` : e;
}
function s6() {
  const e = /* @__PURE__ */ new Date(), t = e.getFullYear(), n = ("0" + (e.getMonth() + 1)).slice(-2), r = ("0" + e.getDate()).slice(-2), o = ("0" + e.getHours()).slice(-2), s = ("0" + e.getMinutes()).slice(-2), l = ("0" + e.getSeconds()).slice(-2);
  return `${t}-${n}-${r}T${o}:${s}:${l}Z`;
}
function a6(e, t) {
  const n = [];
  for (const [r, o] of Object.entries(e)) {
    const { sheetId: s, rowOrCol: l } = ir(r);
    if (t !== s || !o)
      continue;
    const f = l + 1, d = Zv(o.len);
    n.push(
      `<col min="${f}" max="${f}" width="${d}" customWidth="1" ${o.isHide ? 'hidden="1"' : ""}/>`
    );
  }
  return n.length === 0 ? "" : `<cols>${n.join("")}</cols>`;
}
function l6(e) {
  const t = [], n = Object.entries(e);
  n.sort((r, o) => r[0] > o[0] ? -1 : 1);
  for (const [r, o] of n)
    o.length > 0 && t.push(
      `<${r} count="${o.length}">
${o.join(`
`)}
</${r}>`
    );
  return t.join(`
`);
}
function c6(e, t) {
  const { col: n, row: r, sheetId: o, value: s, formula: l, ...f } = e, d = ut({
    col: n,
    row: r,
    rowCount: 1,
    colCount: 1,
    sheetId: ""
  }), c = l ? `<f>${l.slice(1)}</f>` : "";
  let i = `<v>${Nn(s, e?.numberFormat)}</v>`, a = "";
  return On(f) || (a = `s="${t.cellXfs.length}"`, o6(t, f)), typeof e.value == "boolean" && (i = `<v>${Number(e.value)}</v>`, a += ' t="b"'), `<c r="${d}" ${a}>${c}${i}</c>`;
}
function u6(e) {
  const t = e.toJSON(), n = Object.values(t.workbook);
  n.sort((a, u) => a.sort - u.sort);
  const r = [], o = {};
  for (let a = 0; a < n.length; a++) {
    const u = n[a], h = a + 1;
    o[u.sheetId] = {
      rid: `rId${h}`,
      target: `sheet${h}.xml`,
      name: u.name
    }, r.push(
      `<Override PartName="/xl/worksheets/sheet${h}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );
  }
  const s = (a) => {
    const u = a || t.currentSheetId, h = n.find((g) => g.sheetId === u)?.name || "";
    return i6(h);
  }, l = [];
  for (const [a, u] of Object.entries(t.definedNames)) {
    const h = ut(
      u,
      "absolute",
      s
    );
    l.push(`<definedName name="${a}">${h}</definedName>`);
  }
  const f = n.findIndex(
    (a) => a.sheetId === t.currentSheetId
  ), d = {};
  if (d["_rels/.rels"] = _e["_rels/.rels"], d["docProps/app.xml"] = lt(_e["docProps/app.xml"], {
    children: n.map((a) => `<vt:lpstr>${a.name}</vt:lpstr>`).join(""),
    size: String(n.length)
  }), d["docProps/core.xml"] = lt(_e["docProps/core.xml"], {
    children: s6()
  }), d["xl/_rels/workbook.xml.rels"] = lt(
    _e["xl/_rels/workbook.xml.rels"],
    {
      children: n.map((a) => {
        const u = o[a.sheetId];
        return `<Relationship Id="${u.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${u.target}"/>`;
      }).join(""),
      size: `<Relationship Id="rId${String(
        n.length + 1
      )}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
      <Relationship Id="rId${String(
        n.length + 2
      )}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      <Relationship Id="rId${String(
        n.length + 3
      )}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`
    }
  ), Object.keys(t.drawings).length > 0) {
    const a = Object.values(t.drawings);
    let u = 1;
    for (const h of n) {
      const g = a.filter((M) => M.sheetId === h.sheetId);
      if (g.length === 0)
        continue;
      const m = o[h.sheetId], p = `drawing${u}.xml`;
      d[`xl/worksheets/_rels/${m.target}.rels`] = lt(
        _e["xl/worksheets/_rels/sheet1.xml.rels"],
        { children: p }
      );
      let E = 1, b = 1, N = 1;
      const v = [], C = [];
      for (const M of g) {
        let R = M.fromCol, A = M.fromRow;
        for (let L = 0; L < M.width; L++)
          L += e.getColWidth(R++);
        for (let L = 0; L < M.height; L++)
          L += e.getRowHeight(A++);
        const x = `<xdr:from>
        <xdr:col>${M.fromCol}</xdr:col>
        <xdr:colOff>304800</xdr:colOff>
        <xdr:row>${M.fromRow}</xdr:row>
        <xdr:rowOff>165100</xdr:rowOff>
      </xdr:from>
      <xdr:to>
        <xdr:col>${R}</xdr:col>
        <xdr:colOff>457200</xdr:colOff>
        <xdr:row>${A}</xdr:row>
        <xdr:rowOff>63500</xdr:rowOff>
      </xdr:to>`;
        if (M.type === "chart") {
          v.push(
            `<Relationship Id="rId${E}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${b}.xml"/>`
          ), d[`xl/charts/_rels/chart${b}.xml.rels`] = lt(_e["xl/charts/_rels/chart1.xml.rels"], {
            children: String(b)
          }), d[`xl/charts/colors${b}.xml`] = _e["xl/charts/colors1.xml"], d[`xl/charts/style${b}.xml`] = _e["xl/charts/style1.xml"], r.push(
            `<Override PartName="/xl/charts/chart${b}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
          ), r.push(
            `<Override PartName="/xl/charts/colors${b}.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>`
          ), r.push(
            `<Override PartName="/xl/charts/style${b}.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>`
          );
          const L = [], D = M.chartRange;
          for (let W = 0; W < D.colCount; W++) {
            const J = ut(
              {
                row: D.row,
                col: D.col + W,
                rowCount: 1,
                colCount: 1,
                sheetId: ""
              },
              "absolute"
            ), ne = ut(
              {
                row: D.row + D.rowCount - 1,
                col: D.col + W,
                rowCount: 1,
                colCount: 1,
                sheetId: ""
              },
              "absolute"
            );
            L.push(` <c:ser>
            <c:idx val="${b}"/>
            <c:order val="${b}"/>
            <c:spPr>
              <a:solidFill>
                <a:schemeClr val="accent${b}"/>
              </a:solidFill>
              <a:ln>
                <a:noFill/>
              </a:ln>
              <a:effectLst/>
            </c:spPr>
            <c:invertIfNegative val="0"/>
            <c:dLbls>
              <c:delete val="1"/>
            </c:dLbls>
            <c:val>
              <c:numRef>
                <c:f>${s(
              D.sheetId
            )}!${J}:${ne}</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="${D.rowCount}"/>
                </c:numCache>
              </c:numRef>
            </c:val>
          </c:ser>`);
          }
          d[`xl/charts/chart${b}.xml`] = lt(
            _e["xl/charts/chart1.xml"],
            {
              size: M.title,
              children: L.join(`
`),
              large: zc.some((W) => W === M.chartType) ? M.chartType : "bar"
            }
          ), C.push(`<xdr:twoCellAnchor>
          ${x}
          <xdr:graphicFrame>
            <xdr:nvGraphicFramePr>
              <xdr:cNvPr id="${M.uuid}" name="Chart ${E}"/>
              <xdr:cNvGraphicFramePr/>
            </xdr:nvGraphicFramePr>
            <xdr:xfrm>
              <a:off x="8256270" y="1445260"/>
              <a:ext cx="4541520" cy="2885440"/>
            </xdr:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${E}"/>
              </a:graphicData>
            </a:graphic>
          </xdr:graphicFrame>
          <xdr:clientData/>
        </xdr:twoCellAnchor>`);
        } else if (M.type === "floating-picture") {
          const L = Fa(M.imageSrc);
          r.unshift(
            `<Default Extension="${L.ext.slice(1)}" ContentType="${L.type}"/>`
          );
          const D = `image${N}${L.ext}`;
          v.push(
            `<Relationship Id="rId${E}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${D}"/>`
          ), d[`xl/media/${D}`] = L.base64, C.push(`<xdr:twoCellAnchor editAs="oneCell">
          ${x}
          <xdr:pic>
            <xdr:nvPicPr>
              <xdr:cNvPr id="${M.uuid}" name="Picture ${E}" descr="face"/>
              <xdr:cNvPicPr>
                <a:picLocks noChangeAspect="1"/>
              </xdr:cNvPicPr>
            </xdr:nvPicPr>
            <xdr:blipFill>
              <a:blip r:embed="rId${E}"/>
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </xdr:blipFill>
            <xdr:spPr>
              <a:xfrm>
                <a:off x="744220" y="4693920"/>
                <a:ext cx="4770120" cy="5057140"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </xdr:spPr>
          </xdr:pic>
          <xdr:clientData/>
        </xdr:twoCellAnchor>`);
        }
        E++, N++, b++;
      }
      d[`xl/drawings/${p}`] = lt(
        _e["xl/drawings/drawing1.xml"],
        { children: C.join(`
`) }
      ), r.push(
        `<Override PartName="/xl/drawings/${p}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
      ), d[`xl/drawings/_rels/${p}.rels`] = lt(
        _e["xl/drawings/_rels/drawing1.xml.rels"],
        { children: v.join(`
`) }
      ), u++;
    }
  }
  d["xl/theme/theme1.xml"] = _e["xl/theme/theme1.xml"], d["xl/workbook.xml"] = lt(_e["xl/workbook.xml"], {
    size: f >= 0 ? ` activeTab="${f}" ` : "",
    large: l.length > 0 ? `<definedNames>
${l.join(`
`)}
</definedNames>` : "",
    children: n.map((a) => {
      const u = o[a.sheetId];
      return `<sheet name="${a.name}" sheetId="${a.sheetId}" r:id="${u.rid}"/>`;
    }).join(`
`)
  }), d["[Content_Types].xml"] = lt(
    _e["[Content_Types].xml"],
    {
      children: r.join(`
`)
    }
  );
  const c = {
    numFmts: [],
    fonts: [
      `<font>
    <sz val="11"/>
    <color theme="1"/>
    <name val="Calibri"/>
    <charset val="134"/>
    <scheme val="minor"/>
  </font>`
    ],
    fills: [
      `<fill>
    <patternFill patternType="none"/>
  </fill>`
    ],
    cellXfs: [
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0">
    <alignment vertical="center"/>
  </xf>`
    ],
    borders: [
      `<border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>`
    ]
  }, i = [];
  for (const [a, u] of Object.entries(t.worksheets)) {
    const h = jn(a);
    h && i.push({ ...u, ...h });
  }
  for (const a of n) {
    const u = Object.values(t.mergeCells).filter((N) => N.sheetId === a.sheetId).map((N) => `<mergeCell ref="${ut({
      ...N,
      sheetId: ""
    })}"/>`), h = o[a.sheetId], g = t.rangeMap[a.sheetId] || {
      row: 0,
      col: 0,
      rowCount: 1,
      colCount: 1,
      sheetId: ""
    };
    g.sheetId = "";
    const m = {
      children: "<sheetData/>",
      size: a6(t.customWidth, a.sheetId),
      large: `<sheetView ${a.sheetId === t.currentSheetId ? 'tabSelected="1"' : ""} workbookViewId="0">
    <selection activeCell="${ut({
        ...g,
        rowCount: 1,
        colCount: 1
      })}" sqref="${ut(g)}"/>
  </sheetView>`
    }, p = /* @__PURE__ */ new Map(), E = [];
    for (const N of i)
      if (N.sheetId === a.sheetId) {
        const v = p.get(N.row) || [];
        v.push(N), p.set(N.row, v);
      }
    const b = Array.from(p.keys());
    b.sort();
    for (const N of b) {
      const v = t.customHeight[bt(a.sheetId, N)];
      let C = "";
      v && (C = `ht="${v.len}" customHeight="1" ${v.isHide ? 'hidden="1"' : ""}`);
      const M = p.get(N).map((R) => c6(R, c));
      E.push(
        `<row r="${N + 1}" ${C} x14ac:dyDescent="0.4">
${M.join(
          `
`
        )}
</row>`
      );
    }
    E.length > 0 && (m.children = `<sheetData>
${E.join(`
`)}
</sheetData>`), u.length > 0 && (m.children += `<mergeCells count="${u.length}">
${u.join(`
`)}
</mergeCells>`), a.tabColor ? m.larger = `<sheetPr>
      <tabColor rgb="${ro(a.tabColor)}"/>
    </sheetPr>` : m.larger = "<sheetPr/>", d[`xl/worksheets/${h.target}`] = lt(
      _e["xl/worksheets/sheet1.xml"],
      m
    );
  }
  return d["xl/styles.xml"] = lt(_e["xl/styles.xml"], {
    children: l6(c)
  }), d;
}
async function f6(e, t) {
  const n = u6(t), r = (await import("./jszip.min-COPV1wbk.mjs").then((d) => d.j)).default, o = new r(), s = /* @__PURE__ */ new Map(), l = (d) => {
    if (d.length !== 0)
      for (let c = 0; c < d.length; c++)
        if (c === 0)
          s.has(d[c]) || s.set(d[c], o.folder(d[c]));
        else {
          const i = d.slice(0, c).join(Fr), a = d.slice(0, c + 1).join(Fr);
          s.has(a) || s.set(a, s.get(i).folder(d[c]));
        }
  };
  for (const [d, c] of Object.entries(n)) {
    const i = d.split(Fr).map((h) => h.trim()).filter((h) => h), a = i.pop();
    l(i);
    let u = o;
    if (i.length > 0 && (u = s.get(i.join(Fr))), d.startsWith("xl/media"))
      u.file(a, c, { base64: !0 });
    else {
      const h = c.trim().split("/n").map((g) => g.trim()).filter((g) => g).join("");
      u.file(a, h);
    }
  }
  const f = await o.generateAsync({ type: "blob" });
  ar(f, e);
}
const Hr = ua, Hc = dr;
function d6(e) {
  const r = new RegExp('"', "g"), o = new RegExp('"', "g");
  let s = "";
  for (let l = 0; l < e.length; l++) {
    let f = e[l];
    if (f === "")
      s += f;
    else if (f) {
      const d = Hr.length && f.indexOf(Hr) >= 0, c = f.indexOf('"') >= 0, i = f.indexOf(Hc) >= 0, a = c || d || i;
      a && (f = f.replace(r, '""')), c && (f = f.replace(o, '""')), a && (f = '"' + f + '"'), s += f;
    }
    l !== e.length - 1 && (s += Hr);
  }
  return s;
}
function h6(e) {
  const t = e.getCurrentSheetId(), n = [], r = e.getSheetInfo(t);
  if (!r)
    return "";
  for (let s = 0; s < r.rowCount; s++) {
    const l = [];
    for (let f = 0; f < r.colCount; f++) {
      const d = e.getCell({
        row: s,
        col: f,
        colCount: 1,
        rowCount: 1,
        sheetId: t
      });
      l.push(Nn(d?.value, d?.numberFormat));
    }
    n.push(d6(l));
  }
  const o = Array.from({ length: r.colCount }).fill("").join(Hr);
  for (; n.length > 0 && n[n.length - 1] === o; )
    n.pop();
  return n.join(Hc);
}
async function m6(e, t) {
  const n = await Mi(e, !1), r = Ea(n, ua), o = [], s = t.getCurrentSheetId();
  let l = 0;
  for (const f of r) {
    let d = 0;
    for (const c of f)
      o.push({ row: l, col: d, value: c, sheetId: s }), d++;
    l++;
  }
  t.transaction(() => {
    t.deleteAll(), t.setWorksheet(o);
  });
}
const g6 = "_i18n_ow4wb_21", p6 = "_theme_ow4wb_16", v6 = "_user_ow4wb_26", N6 = "_status_ow4wb_31", C6 = "_file_ow4wb_40", At = {
  "menubar-container": "_menubar-container_ow4wb_1",
  "menubar-menu": "_menubar-menu_ow4wb_10",
  "theme-button": "_theme-button_ow4wb_16",
  i18n: g6,
  theme: p6,
  user: v6,
  status: N6,
  file: C6
};
function pi(e) {
  const t = Object.keys(e);
  for (const n of t) {
    const r = `--${n}`, o = String(e[n] || "");
    document.documentElement.style.setProperty(r, o);
  }
}
function E6(e) {
  pi(e === "dark" ? Ml : ji);
}
const Vc = ue(() => {
  const { controller: e } = Se(), [t, n] = he("light");
  Re(() => {
    pi(Kr);
  }, []), Re(() => {
    n(Zi()), typeof window.matchMedia == "function" && window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", (o) => {
      n(o.matches ? "dark" : "light");
    });
  }, []), Re(() => {
    Wp(t), E6(t), e.emit("renderChange", {
      changeSet: /* @__PURE__ */ new Set(["cellStyle"])
    });
  }, [t, e]);
  const r = z(() => {
    n((o) => o === "dark" ? "light" : "dark");
  }, []);
  return /* @__PURE__ */ I("div", { "data-testid": "menubar-theme", className: At.theme, children: /* @__PURE__ */ I(
    V,
    {
      onClick: r,
      className: At["theme-button"],
      testId: "menubar-theme-toggle",
      children: /* @__PURE__ */ I(Ae, { name: t === "dark" ? "sun" : "moon" }, void 0, !1, {
        fileName: "/home/user/excel-collab/src/containers/MenuBar/Theme.tsx",
        lineNumber: 68,
        columnNumber: 9
      }, void 0)
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/containers/MenuBar/Theme.tsx",
      lineNumber: 63,
      columnNumber: 7
    },
    void 0
  ) }, void 0, !1, {
    fileName: "/home/user/excel-collab/src/containers/MenuBar/Theme.tsx",
    lineNumber: 62,
    columnNumber: 5
  }, void 0);
});
Vc.displayName = "Theme";
const b6 = Xo.map((e) => ({
  value: e,
  label: e,
  disabled: !1
})), Gc = ue(() => {
  const e = z((t) => {
    k.changeLanguage(String(t));
  }, []);
  return /* @__PURE__ */ I("div", { className: At.i18n, "data-testid": "menubar-i18n", children: /* @__PURE__ */ I(
    ln,
    {
      data: b6,
      defaultValue: k.current,
      onChange: e,
      testId: "menubar-i18n-select"
    },
    void 0,
    !1,
    {
      fileName: "/home/user/excel-collab/src/containers/MenuBar/I18N.tsx",
      lineNumber: 20,
      columnNumber: 7
    },
    void 0
  ) }, void 0, !1, {
    fileName: "/home/user/excel-collab/src/containers/MenuBar/I18N.tsx",
    lineNumber: 19,
    columnNumber: 5
  }, void 0);
});
Gc.displayName = "I18N";
const $c = () => {
  const e = vn((n) => n.clientId), t = navigator.onLine ? qo.ONLINE : qo.LOCAL;
  return /* @__PURE__ */ I("div", { className: At.user, children: [
    /* @__PURE__ */ I("div", { className: At.userName, children: `${k.t("user-name")} ${e}` }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/MenuBar/User.tsx",
      lineNumber: 13,
      columnNumber: 7
    }, void 0),
    /* @__PURE__ */ I(Ae, { name: t, className: At.status }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/MenuBar/User.tsx",
      lineNumber: 14,
      columnNumber: 7
    }, void 0)
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/MenuBar/User.tsx",
    lineNumber: 12,
    columnNumber: 5
  }, void 0);
};
$c.displayName = "User";
const w6 = ({ visible: e, setVisible: t }) => {
  const { provider: n, controller: r } = Se(), [o, s] = he(""), l = vn((a) => a.fileName), f = vn((a) => a.setFileName);
  Re(() => {
    s(l || k.t("default-name"));
  }, [l]);
  const d = z(() => {
    t(!0);
  }, []), c = z(
    (a) => {
      s(a.currentTarget.value.trim());
    },
    []
  ), i = z(() => {
    o && (n?.updateDocument?.(r.getHooks().doc.guid, { name: o }), f(o), t(!1));
  }, [o, n, r]);
  return /* @__PURE__ */ I(je.Fragment, { children: [
    /* @__PURE__ */ I("div", { className: At.file, onClick: d, children: l || k.t("default-name") }, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/MenuBar/File.tsx",
      lineNumber: 44,
      columnNumber: 7
    }, void 0),
    /* @__PURE__ */ I(
      Ji,
      {
        title: k.t("change-file-name"),
        visible: e,
        onCancel: () => t(!1),
        onOk: i,
        children: /* @__PURE__ */ I(
          "input",
          {
            type: "text",
            value: o,
            onChange: c,
            maxLength: 50,
            autoFocus: !0
          },
          void 0,
          !1,
          {
            fileName: "/home/user/excel-collab/src/containers/MenuBar/File.tsx",
            lineNumber: 53,
            columnNumber: 9
          },
          void 0
        )
      },
      void 0,
      !1,
      {
        fileName: "/home/user/excel-collab/src/containers/MenuBar/File.tsx",
        lineNumber: 47,
        columnNumber: 7
      },
      void 0
    )
  ] }, void 0, !0, {
    fileName: "/home/user/excel-collab/src/containers/MenuBar/File.tsx",
    lineNumber: 43,
    columnNumber: 5
  }, void 0);
}, qc = ue(
  ({ leftChildren: e, rightChildren: t, hideNewFile: n, hideRenameFile: r }) => {
    const { controller: o, provider: s } = Se(), [l, f] = he(!1), d = z(() => {
      f6(`excel_${Date.now()}.xlsx`, o);
    }, []), c = z(() => {
      const g = h6(o), m = new Blob([g], {
        type: "text/csv;charset=utf-8;"
      });
      ar(m, `excel_${Date.now()}.csv`);
    }, []), i = z(
      async (g) => {
        const m = g.target.files?.[0];
        if (!m)
          return;
        const p = await r6(m);
        await o.fromJSON(p), g.target.value = "", g.target.blur();
      },
      []
    ), a = z(
      async (g) => {
        const m = g.target.files?.[0];
        m && (await m6(m, o), g.target.value = "", g.target.blur());
      },
      []
    ), u = z(() => {
      const g = new Blob([JSON.stringify(o.toJSON())], {
        type: "application/json"
      });
      ar(g, `excel_${Date.now()}.json`);
    }, []), h = z(() => {
      const g = fr();
      s?.addDocument?.(g);
    }, []);
    return /* @__PURE__ */ I("div", { className: At["menubar-container"], "data-testid": "menubar", children: [
      /* @__PURE__ */ I("div", { className: At["menubar-menu"], children: [
        !r && /* @__PURE__ */ I(w6, { visible: l, setVisible: f }, void 0, !1, {
          fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
          lineNumber: 74,
          columnNumber: 13
        }, void 0),
        /* @__PURE__ */ I(
          Qi,
          {
            label: k.t("file"),
            className: At.menu,
            testId: "menubar-excel",
            children: [
              !n && /* @__PURE__ */ I(Le, { onClick: h, testId: "menubar-new-excel", children: k.t("new-file") }, void 0, !1, {
                fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                lineNumber: 82,
                columnNumber: 15
              }, void 0),
              !r && /* @__PURE__ */ I(
                Le,
                {
                  onClick: () => f(!0),
                  testId: "menubar-rename-excel",
                  children: k.t("rename-file")
                },
                void 0,
                !1,
                {
                  fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                  lineNumber: 87,
                  columnNumber: 15
                },
                void 0
              ),
              /* @__PURE__ */ I(Le, { testId: "menubar-import-xlsx", children: [
                /* @__PURE__ */ I(
                  "input",
                  {
                    type: "file",
                    hidden: !0,
                    onChange: i,
                    accept: ".xlsx",
                    id: "import_xlsx",
                    "data-testid": "menubar-import-xlsx-input"
                  },
                  void 0,
                  !1,
                  {
                    fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                    lineNumber: 95,
                    columnNumber: 15
                  },
                  void 0
                ),
                /* @__PURE__ */ I("label", { htmlFor: "import_xlsx", children: k.t("import", { format: "XLSX" }) }, void 0, !1, {
                  fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                  lineNumber: 103,
                  columnNumber: 15
                }, void 0)
              ] }, void 0, !0, {
                fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                lineNumber: 94,
                columnNumber: 13
              }, void 0),
              /* @__PURE__ */ I(Le, { testId: "menubar-import-csv", children: [
                /* @__PURE__ */ I(
                  "input",
                  {
                    type: "file",
                    hidden: !0,
                    onChange: a,
                    accept: ".csv",
                    id: "import_csv",
                    "data-testid": "menubar-import-csv-input"
                  },
                  void 0,
                  !1,
                  {
                    fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                    lineNumber: 108,
                    columnNumber: 15
                  },
                  void 0
                ),
                /* @__PURE__ */ I("label", { htmlFor: "import_csv", children: k.t("import", { format: "CSV" }) }, void 0, !1, {
                  fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                  lineNumber: 116,
                  columnNumber: 15
                }, void 0)
              ] }, void 0, !0, {
                fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                lineNumber: 107,
                columnNumber: 13
              }, void 0),
              /* @__PURE__ */ I(Le, { onClick: d, testId: "menubar-export-xlsx", children: k.t("export", { format: "XLSX" }) }, void 0, !1, {
                fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                lineNumber: 120,
                columnNumber: 13
              }, void 0),
              /* @__PURE__ */ I(Le, { testId: "menubar-export-csv", onClick: c, children: k.t("export", { format: "CSV" }) }, void 0, !1, {
                fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                lineNumber: 123,
                columnNumber: 13
              }, void 0),
              /* @__PURE__ */ I(Le, { testId: "menubar-export-json", onClick: u, children: k.t("export", { format: "JSON" }) }, void 0, !1, {
                fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
                lineNumber: 126,
                columnNumber: 13
              }, void 0)
            ]
          },
          void 0,
          !0,
          {
            fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
            lineNumber: 76,
            columnNumber: 11
          },
          void 0
        ),
        e
      ] }, void 0, !0, {
        fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
        lineNumber: 72,
        columnNumber: 9
      }, void 0),
      t,
      /* @__PURE__ */ I($c, {}, void 0, !1, {
        fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
        lineNumber: 133,
        columnNumber: 9
      }, void 0),
      /* @__PURE__ */ I(Gc, {}, void 0, !1, {
        fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
        lineNumber: 134,
        columnNumber: 9
      }, void 0),
      /* @__PURE__ */ I(Vc, {}, void 0, !1, {
        fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
        lineNumber: 135,
        columnNumber: 9
      }, void 0)
    ] }, void 0, !0, {
      fileName: "/home/user/excel-collab/src/containers/MenuBar/index.tsx",
      lineNumber: 71,
      columnNumber: 7
    }, void 0);
  }
);
qc.displayName = "MenuBarContainer";
function I6() {
  const [e, t] = he(!0), n = vn((d) => d.setFileInfo), { provider: r, controller: o, awareness: s } = Se(), l = Xe(!1), f = Xe(!1);
  return Re(() => {
    if (l.current)
      return;
    l.current = !0;
    async function d() {
      if (!r) {
        o.getSheetList().length === 0 && o.addFirstSheet(), t(!1);
        return;
      }
      t(!0);
      const c = o.getHooks().doc, i = await r?.getDocument?.(c.guid), a = i?.content ?? "";
      a && o.fromJSON(JSON.parse(a)), n(i?.id ?? c.guid, i?.name ?? ""), o.getSheetList().length === 0 && o.addFirstSheet(), t(!1);
    }
    d();
  }, []), Re(() => {
    if (!s || f.current)
      return;
    f.current = !0;
    const d = o.getHooks().doc;
    s.on("update", () => {
      const c = [];
      for (const i of s.getStates().entries()) {
        const [a, u] = i;
        !u.range || a === d.clientID || c.push({ clientId: a, range: u.range });
      }
      vn.getState().setUsers(c);
    }), d.on("update", (c, i, a, u) => {
      const h = vu(u);
      o.emit("renderChange", { changeSet: h });
    }), o.on("rangeChange", (c) => {
      s.setLocalStateField("range", c);
    });
  }, []), {
    isLoading: e
  };
}
const Wc = ue(
  ({
    style: e,
    menubarLeftChildren: t,
    menubarRightChildren: n,
    toolbarChildren: r,
    sheetBarChildren: o,
    hideNewFile: s,
    hideRenameFile: l
  }) => {
    const { isLoading: f } = I6();
    return f ? /* @__PURE__ */ I(Co, {}, void 0, !1, {
      fileName: "/home/user/excel-collab/src/containers/index.tsx",
      lineNumber: 108,
      columnNumber: 14
    }, void 0) : /* @__PURE__ */ I(
      "div",
      {
        className: Tp["app-container"],
        "data-testid": "app-container",
        style: e,
        children: [
          /* @__PURE__ */ I(
            qc,
            {
              leftChildren: t,
              rightChildren: n,
              hideNewFile: s,
              hideRenameFile: l
            },
            void 0,
            !1,
            {
              fileName: "/home/user/excel-collab/src/containers/index.tsx",
              lineNumber: 117,
              columnNumber: 9
            },
            void 0
          ),
          /* @__PURE__ */ I(Wl, { children: r }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/index.tsx",
            lineNumber: 123,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(Hl, {}, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/index.tsx",
            lineNumber: 124,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(Pc, {}, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/index.tsx",
            lineNumber: 125,
            columnNumber: 9
          }, void 0),
          /* @__PURE__ */ I(_c, { children: o }, void 0, !1, {
            fileName: "/home/user/excel-collab/src/containers/index.tsx",
            lineNumber: 126,
            columnNumber: 9
          }, void 0)
        ]
      },
      void 0,
      !0,
      {
        fileName: "/home/user/excel-collab/src/containers/index.tsx",
        lineNumber: 112,
        columnNumber: 7
      },
      void 0
    );
  }
);
Wc.displayName = "ExcelEditor";
const na = /* @__PURE__ */ new Map();
function M6(e, t) {
  const n = `${t}__${e.font}`, r = na.get(n);
  if (r)
    return r;
  const o = e.measureText(t), { actualBoundingBoxAscent: s, actualBoundingBoxDescent: l } = o, f = s + l, d = o.width, c = Math.ceil(d / sr()), i = Math.ceil(f / sr()), a = { width: c, height: i };
  return na.set(n, a), a;
}
function ct(e, t, n, r, o) {
  e.fillRect(ce(t), ce(n), ce(r), ce(o));
}
function vi(e, t, n, r, o) {
  e.strokeRect(ce(t), ce(n), ce(r), ce(o));
}
function ra(e, t, n, r, o) {
  e.clearRect(ce(t), ce(n), ce(r), ce(o));
}
function Ni(e, t, n, r) {
  e.fillText(t, ce(n), ce(r));
}
function qt(e, t) {
  if (t.length !== 0) {
    e.beginPath();
    for (let n = 0; n < t.length; n += 2) {
      const r = t[n], o = t[n + 1];
      e.moveTo(ce(r[0]), ce(r[1])), e.lineTo(ce(o[0]), ce(o[1]));
    }
    e.stroke();
  }
}
function oa(e, t, n, r) {
  e.beginPath(), e.moveTo(ce(t[0]), ce(t[1])), e.lineTo(ce(n[0]), ce(n[1])), e.lineTo(ce(r[0]), ce(r[1])), e.fill();
}
function Yc(e, t, n) {
  const r = e.slice(), [o, s] = e, l = Z * 2, f = n ? l : 0;
  return t === "bottom" ? r.push([o[0] + f, o[1] - l], [s[0] - f, s[1] - l]) : t === "top" ? r.push([o[0] + f, o[1] + l], [s[0] - f, s[1] + l]) : t === "left" ? r.push([o[0] + l, o[1] + f], [s[0] + l, s[1] - f]) : t === "right" && r.push([o[0] - l, o[1] + f], [s[0] - l, s[1] - f]), r;
}
function y6(e, t, n) {
  return n ? t ? e.split(mt) : Ms(e.replaceAll(mt, "")) : Ms(e);
}
function S6(e, t, n, r, o) {
  e.setLineDash([ce(8), ce(6)]);
  const s = Z;
  vi(
    e,
    t + s,
    n + s,
    r - s * 2,
    o - s * 2
  ), e.setLineDash([]);
}
function R6(e) {
  let t = [];
  return e === "hair" ? t = [Z, Z] : e === "dotted" || e === "mediumDashed" ? t = [Z * 2, Z * 2] : e === "dashed" ? t = [Z * 4, Z * 4] : e === "dashDot" || e === "mediumDashDot" ? t = [
    Z * 4,
    Z * 4,
    Z * 8,
    Z * 4
  ] : (e === "dashDotDot" || e === "mediumDashDotDot") && (t = [
    Z * 4,
    Z * 4,
    Z * 8,
    Z * 4,
    Z * 4,
    Z * 4
  ]), t;
}
function Ur(e, t, n, r, o) {
  if (!n)
    return;
  const { top: s, left: l, width: f, height: d } = t;
  let c = [];
  r === "top" ? c = [
    [l, s],
    [l + f, s]
  ] : r === "bottom" ? c = [
    [l, s + d],
    [l + f, s + d]
  ] : r === "left" ? c = [
    [l, s],
    [l, s + d]
  ] : r === "right" && (c = [
    [l + f, s],
    [l + f, s + d]
  ]);
  const { type: i, color: a } = n;
  e.lineWidth = oo[i], e.strokeStyle = a || Ee("black", o);
  const u = R6(i);
  i === "double" && (c = Yc(c, r, !0)), u.length > 0 && e.setLineDash(u.map((h) => ce(h))), qt(e, c), u.length > 0 && e.setLineDash([]);
}
function A6(e, t, n, r, o) {
  const { value: s, formula: l, ...f } = n, d = { height: 0, width: 0 };
  if (s === "" && On(f))
    return d;
  const c = n?.numberFormat || cn, i = c === cn && typeof s == "number", a = Nn(s, c), u = n?.fontSize ? n.fontSize : Cn, h = Ca(
    n?.isItalic ? "italic" : "normal",
    n?.isBold ? "bold" : "500",
    ce(u),
    n?.fontFamily
  );
  n?.fillColor && (e.fillStyle = n?.fillColor, ct(e, t.left, t.top, t.width, t.height));
  let g = n?.fontColor || Ee("contentColor", o);
  No(a) && (g = Ee("errorFormulaColor", o)), e.font = h, e.fillStyle = g;
  const m = { ...n };
  let p = m?.horizontalAlign;
  m?.horizontalAlign === void 0 && i && (p = xe.RIGHT), m.horizontalAlign = p;
  const N = (!m?.isWrapText && us(c) ? [a] : y6(a, n?.isWrapText, r)).map((R) => {
    const A = M6(e, R);
    return {
      str: R,
      width: A.width,
      height: A.height === 0 ? u : A.height
    };
  }), { width: v, height: C, resultList: M } = T6(
    N,
    t,
    m,
    wi(!!r, a)
  );
  if (v > 0 && C > 0) {
    const R = Math.ceil(u * (Kr.lineHeight - 1) / 2);
    let A = [];
    for (const x of M) {
      if (Ni(e, x.text, x.x, x.y), m?.underline) {
        e.strokeStyle = g;
        const L = x.y + x.height + R / 2, D = [
          [x.x, L],
          [x.x + x.width, L]
        ];
        m?.underline === Ce.DOUBLE ? A = A.concat(Yc(D, "bottom", !1)) : A = A.concat(D);
      }
      if (m?.isStrike) {
        e.strokeStyle = g;
        const L = x.y + x.height / 2 + R / 2;
        A = A.concat([
          [x.x, L],
          [x.x + x.width, L]
        ]);
      }
    }
    qt(e, A);
  }
  return d.height = Math.ceil(C), d.width = Math.ceil(v), d;
}
function T6(e, t, n, r) {
  const o = n?.fontSize ? n.fontSize : Cn, s = Math.ceil(o * (Kr.lineHeight - 1) / 2), l = n?.verticalAlign ?? Te.MIDDLE, { left: f, top: d, height: c } = t, i = Math.max(t.width, ...e.map((g) => g.width)), a = [];
  let u = 0, h = 0;
  if (n?.isWrapText && r) {
    let g = 0;
    for (let m = 0; m < e.length; m++) {
      const p = e[m];
      u = Math.max(u, p.width);
      const E = p.height + s * 2;
      h += E, a.push({
        text: p.str,
        x: 0,
        y: g,
        width: p.width,
        height: p.height
      }), g += E;
    }
  } else if (n?.isWrapText) {
    let g = 0;
    for (let m = 0; m < e.length; ) {
      let p = i, E = "", b = 0, N = 0;
      for (; m < e.length; ) {
        const v = e[m];
        if (p >= v.width)
          b += v.width, N = Math.max(N, v.height), p -= v.width, E += v.str, m++;
        else
          break;
      }
      if (E) {
        u = Math.max(u, b);
        const v = N + s * 2;
        h += v, a.push({
          text: E,
          x: 0,
          y: g,
          width: b,
          height: N
        }), g += v;
      }
    }
  } else {
    let g = "", m = i;
    for (let p = 0; p < e.length; p++) {
      const E = e[p];
      if (m >= E.width)
        u += E.width, h = Math.max(h, E.height), g += E.str, m -= E.width;
      else
        break;
    }
    a.push({
      text: g,
      x: 0,
      y: 0,
      width: u,
      height: h
    });
  }
  if (h = Math.max(h, o * Kr.lineHeight), u += s, h += s, u <= i && h <= c) {
    let g = f + s, m = d + (c - h) / 2 + s;
    l === Te.TOP ? m = d + s : l === Te.BOTTOM && (m = d + (c - h) + s), n?.horizontalAlign === xe.CENTER ? g = f + (i - u) / 2 : n?.horizontalAlign === xe.RIGHT && (g = f + (i - u) - s);
    for (const p of a)
      p.x = p.x + g, p.y = p.y + m;
  }
  return {
    width: u,
    height: h,
    resultList: a
  };
}
function ia(e) {
  return {
    textAlign: "center",
    textBaseline: "middle",
    font: Ca(void 0, "500", ce(Cn)),
    fillStyle: Ee("black", e),
    lineWidth: Z,
    strokeStyle: Ee("borderColor", e)
  };
}
const sa = Math.max(...Object.values(oo));
class x6 {
  canvas;
  ctx;
  width = 0;
  height = 0;
  isRendering = !1;
  rowMap = {};
  colMap = {};
  eventData = {
    theme: "light",
    sheetData: {},
    canvasSize: {
      top: 0,
      left: 0,
      width: 0,
      height: 0
    },
    headerSize: {
      width: 0,
      height: 0
    },
    currentSheetInfo: {
      isHide: !1,
      rowCount: 0,
      colCount: 0,
      name: "",
      sheetId: "",
      tabColor: "",
      sort: 1
    },
    scroll: {
      row: 0,
      col: 0,
      top: 0,
      left: 0,
      scrollLeft: 0,
      scrollTop: 0
    },
    range: {
      row: 0,
      col: 0,
      rowCount: 1,
      colCount: 1,
      sheetId: ""
    },
    copyRange: void 0,
    customHeight: {},
    customWidth: {},
    currentMergeCells: [],
    autoFilter: void 0
  };
  constructor(t) {
    this.canvas = t, this.ctx = t.getContext("2d");
    const n = sr();
    this.ctx.scale(n, n);
  }
  render(t) {
    if (t.changeSet.size === 0 || this.isRendering)
      return;
    this.isRendering = !0, this.eventData = t, this.clear();
    const { ctx: n } = this;
    n.strokeStyle = Ee("primaryColor", t.theme), n.fillStyle = Ee("white", t.theme), n.lineWidth = Z * 2;
    const { width: r, height: o } = this.eventData.canvasSize, s = this.eventData.headerSize, { endRow: l, contentHeight: f } = this.renderRowsHeader(o), { endCol: d, contentWidth: c } = this.renderColsHeader(r);
    this.renderGrid(r - s.width, o - s.height), this.renderTriangle(), this.renderMergeCell(), this.ctx.fillStyle = Ee("selectionColor", this.eventData.theme);
    const i = this.renderSelection({
      endRow: l,
      endCol: d,
      contentHeight: f,
      contentWidth: c
    });
    return this.renderAntLine(i), this.renderContent({ endRow: l, endCol: d, contentHeight: f, contentWidth: c }), this.ctx.lineWidth = sa, vi(this.ctx, i.left, i.top, i.width, i.height), this.isRendering = !1, {
      rowMap: { ...this.rowMap },
      colMap: { ...this.colMap }
    };
  }
  resize(t) {
    this.width = t.width, this.height = t.height, this.canvas.width = ce(t.width), this.canvas.height = ce(t.height);
  }
  clear() {
    ra(this.ctx, 0, 0, this.width, this.height);
  }
  renderRowsHeader(t) {
    const { row: n } = this.eventData.scroll, r = this.eventData.headerSize, { rowCount: o } = this.eventData.currentSheetInfo;
    this.ctx.save();
    const s = this.eventData.range;
    ct(this.ctx, 0, r.height, r.width, t), Object.assign(this.ctx, ia(this.eventData.theme));
    const l = [];
    let f = r.height, d = n;
    for (; d < o && f < t; d++) {
      const i = this.getRowHeight(d);
      let a = f;
      if (d === n && (a += Z / 2), l.push([0, a], [r.width, a]), i > 0) {
        const u = this.isHighlightRow(s, d);
        this.ctx.fillStyle = u ? Ee("primaryColor", this.eventData.theme) : Ee("black", this.eventData.theme), Ni(
          this.ctx,
          String(d + 1),
          r.width / 2,
          a + i / 2
        );
      }
      f += i;
    }
    l.push([0, f], [r.width, f]), l.push([0, 0], [0, f]), qt(this.ctx, l), this.ctx.restore();
    const c = d >= o ? f : t;
    return { endRow: d, contentHeight: Math.floor(c) };
  }
  renderColsHeader(t) {
    const { col: n, row: r } = this.eventData.scroll, o = this.eventData.headerSize, { colCount: s } = this.eventData.currentSheetInfo, l = this.eventData.range, f = [];
    this.ctx.save(), ct(this.ctx, o.width, 0, t, o.height), Object.assign(this.ctx, ia());
    let d = o.width, c = n;
    const i = this.eventData.autoFilter;
    let a = 0;
    for (let h = r; h < this.eventData.currentSheetInfo.rowCount && (a = this.getRowHeight(r), a === 0); h++)
      ;
    for (; c < s && d <= t; c++) {
      const h = this.getColWidth(c);
      let g = d;
      if (c === n && (g += Z / 2), f.push([g, 0], [g, o.height]), h > 0) {
        const m = this.isHighlightCol(l, c);
        this.ctx.fillStyle = m ? Ee("primaryColor", this.eventData.theme) : Ee("black", this.eventData.theme), Ni(
          this.ctx,
          wa(c),
          g + h / 2,
          o.height / 2
        );
      }
      i && c >= i.range.col && c < i.range.col + i.range.colCount && this.renderFilter(
        d + h,
        o.height + a,
        i.col === c
      ), d += h;
    }
    f.push([d, 0], [d, o.height]), f.push([0, 0], [d, 0]), qt(this.ctx, f), this.ctx.restore();
    const u = c >= s ? d : t;
    return {
      endCol: c,
      contentWidth: Math.floor(u)
    };
  }
  renderFilter(t, n, r) {
    const o = Yo, s = Z * 4, l = t - o - s, f = n - o - s;
    this.ctx.fillStyle = r ? Ee("primaryColor", this.eventData.theme) : Ee("borderColor", this.eventData.theme), this.ctx.strokeStyle = r ? Ee("primaryColor", this.eventData.theme) : Ee("borderColor", this.eventData.theme), vi(this.ctx, l, f, o, o), oa(
      this.ctx,
      [l + o / 4, f + o * 3 / 8],
      [l + o * 3 / 4, f + o * 3 / 8],
      [l + o / 2, f + o * 5 / 8]
    );
  }
  renderGrid(t, n) {
    const r = this.eventData.headerSize, { row: o, col: s } = this.eventData.scroll, { rowCount: l, colCount: f } = this.eventData.currentSheetInfo;
    this.ctx.save(), this.ctx.lineWidth = Z, this.ctx.strokeStyle = Ee("borderColor", this.eventData.theme), this.ctx.translate(ce(r.width), ce(r.height));
    const d = [];
    let c = 0, i = 0;
    const a = [], u = [];
    for (let h = o; h < l && c <= n; h++) {
      for (; h < l && this.getRowHeight(h) === 0; )
        h++;
      u.push([0, c]);
      const g = this.getRowHeight(h);
      c += g;
    }
    for (let h = s; h < f && i <= t; h++) {
      for (; h < f && this.getColWidth(h) === 0; )
        h++;
      a.push([i, 0]);
      const g = this.getColWidth(h);
      i += g;
    }
    for (const h of a)
      d.push(h, [h[0], c]);
    for (const h of u)
      d.push(h, [i, h[1]]);
    d.push([0, c], [i, c]), d.push([i, 0], [i, c]), qt(this.ctx, d), this.ctx.restore();
  }
  renderTriangle() {
    const t = this.eventData.headerSize;
    this.ctx.save(), ct(this.ctx, 0, 0, t.width, t.height), this.ctx.fillStyle = Ee(
      "triangleFillColor",
      this.eventData.theme
    );
    const n = 2, r = Math.floor(n), o = Math.floor(t.height - n), s = Math.floor(t.width * 0.4), l = Math.floor(t.width - n);
    oa(this.ctx, [l, r], [s, o], [l, o]), this.ctx.restore();
  }
  getRowHeight(t) {
    const n = bt(
      this.eventData.currentSheetInfo.sheetId,
      t
    ), r = this.eventData.customHeight[n];
    return r ? r.isHide ? Wo : r.len : nr;
  }
  getColWidth(t) {
    const n = bt(
      this.eventData.currentSheetInfo.sheetId,
      t
    ), r = this.eventData.customWidth[n];
    return r ? r.isHide ? Wo : r.len : rr;
  }
  isHighlightRow(t, n) {
    return !!(An(t) || Yt(t) || n >= t.row && n < t.row + t.rowCount);
  }
  isHighlightCol(t, n) {
    return !!(An(t) || Wt(t) || n >= t.col && n < t.col + t.colCount);
  }
  renderAntLine(t) {
    const n = this.eventData.copyRange;
    !n || n.sheetId !== this.eventData.currentSheetInfo.sheetId || (this.ctx.lineWidth = sa, this.ctx.strokeStyle = Ee("primaryColor", this.eventData.theme), S6(
      this.ctx,
      t.left,
      t.top,
      t.width,
      t.height
    ));
  }
  renderMergeCell() {
    const t = this.eventData.currentMergeCells;
    if (t.length === 0)
      return;
    const n = this.eventData.range;
    for (const r of t)
      n.row === r.row && n.col === r.col || this.clearRect(r);
  }
  /* jscpd:ignore-start */
  getCellSize(t) {
    let { row: n, col: r, colCount: o, rowCount: s } = t, l = n, f = r, d = n + s, c = r + o;
    const i = this.eventData.currentSheetInfo;
    An(t) ? (f = 0, c = i.colCount, l = 0, d = i.rowCount) : Yt(t) ? (l = 0, d = i.rowCount) : Wt(t) && (f = 0, c = i.colCount);
    let a = 0, u = 0;
    for (; l < d; l++)
      u += this.getRowHeight(l);
    for (; f < c; f++)
      a += this.getColWidth(f);
    return { width: a, height: u };
  }
  computeCellPosition(t) {
    const { row: n, col: r } = t, o = this.eventData.headerSize, s = this.eventData.scroll;
    let l = o.width, f = o.height, d = s.row, c = s.col;
    if (r >= s.col)
      for (; c < r; )
        l += this.getColWidth(c), c++;
    else
      for (l = -o.width; c > r; )
        l -= this.getColWidth(c), c--;
    if (n >= s.row)
      for (; d < n; )
        f += this.getRowHeight(d), d++;
    else
      for (f = -o.height; d > n; )
        f -= this.getRowHeight(d), d--;
    return {
      top: f,
      left: l
    };
  }
  getActiveRange(t) {
    const n = t || this.eventData.range, r = this.eventData.currentMergeCells;
    for (const o of r)
      if (Ma(n, o))
        return {
          range: {
            ...o,
            sheetId: o.sheetId
          },
          isMerged: !0
        };
    return {
      range: n,
      isMerged: !1
    };
  }
  /* jscpd:ignore-end */
  clearRect(t) {
    const n = this.getCellSize(t);
    if (n.width <= 0 || n.height <= 0)
      return;
    const r = this.computeCellPosition(t), o = Z;
    ra(
      this.ctx,
      r.left + o,
      r.top + o,
      n.width - o * 2,
      n.height - o * 2
    );
  }
  renderContent(t) {
    const { endCol: n, endRow: r, contentHeight: o, contentWidth: s } = t, { ctx: l } = this;
    l.textAlign = "left", l.textBaseline = "top", l.lineWidth = Z * 2;
    const f = this.eventData.headerSize, { row: d, col: c } = this.eventData.scroll, i = Math.floor(s - f.width), a = Math.floor(o - f.height);
    l.save(), this.rowMap = {}, this.colMap = {};
    const u = this.eventData.currentMergeCells;
    for (let h = d; h < r; h++)
      for (let g = c; g < n; g++) {
        const m = u.find(
          (p) => p.row === h && p.col === g
        );
        this.renderCell(h, g, m, i, a);
      }
    l.restore();
  }
  renderCell(t, n, r, o, s) {
    const { ctx: l } = this, f = {
      row: t,
      col: n,
      rowCount: 1,
      colCount: 1,
      sheetId: ""
    }, d = nt(
      this.eventData.currentSheetInfo.sheetId,
      t,
      n
    ), c = this.eventData.sheetData[d];
    if (!c)
      return;
    const i = this.getCellSize(r || f);
    if (i.width <= 0 || i.height <= 0)
      return;
    const a = this.computeCellPosition(f);
    l.lineWidth = Z * 2;
    const u = this.eventData.theme, h = A6(
      l,
      {
        top: a.top,
        left: a.left,
        width: Math.min(i.width, o),
        height: Math.min(i.height, s)
      },
      c,
      !!r,
      u
    ), g = Math.max(this.rowMap[t] ?? 0, h.height), m = Math.max(this.colMap[n] ?? 0, h.width);
    g >= nr && (this.rowMap[t] = g), m >= rr && (this.colMap[n] = m);
    const p = {
      top: a.top,
      left: a.left,
      height: Math.max(g, i.height),
      width: Math.max(m, i.width)
    };
    Ur(l, p, c.borderTop, "top", u), Ur(l, p, c.borderBottom, "bottom", u), Ur(l, p, c.borderLeft, "left", u), Ur(l, p, c.borderRight, "right", u);
  }
  renderSelection(t) {
    const n = this.eventData.range;
    return An(n) ? this.renderSelectAll(t) : Yt(n) ? this.renderSelectCol(t) : Wt(n) ? this.renderSelectRow(t) : this.renderSelectRange();
  }
  renderSelectRange() {
    const t = this.eventData.headerSize, n = this.eventData.range, r = this.computeCellPosition({
      row: n.row,
      col: n.col,
      rowCount: 1,
      colCount: 1,
      sheetId: ""
    }), o = n.row + n.rowCount - 1, s = n.col + n.colCount - 1, l = {
      row: o,
      col: s,
      rowCount: 1,
      colCount: 1,
      sheetId: ""
    }, f = this.computeCellPosition(l), d = this.getCellSize(l), c = f.left + d.width - r.left, i = f.top + d.height - r.top;
    ct(this.ctx, r.left, 0, c, t.height), ct(this.ctx, 0, r.top, t.width, i);
    const a = n.rowCount > 1 || n.colCount > 1;
    a && ct(this.ctx, r.left, r.top, c, i);
    const u = [
      [r.left, t.height],
      [r.left + c, t.height]
    ];
    return u.push(
      [t.width, r.top],
      [t.width, r.top + i]
    ), qt(this.ctx, u), a && this.renderActiveCell(), {
      left: r.left,
      top: r.top,
      width: c,
      height: i
    };
  }
  renderSelectAll(t) {
    const { contentHeight: n, contentWidth: r } = t, o = this.eventData.headerSize;
    ct(this.ctx, 0, 0, r, n), this.renderActiveCell();
    const s = r - o.width, l = n - o.height;
    return {
      left: o.width,
      top: o.height,
      width: s,
      height: l
    };
  }
  renderSelectCol({ contentHeight: t }) {
    const n = this.eventData.headerSize, r = this.eventData.range, o = this.computeCellPosition(r);
    let s = 0;
    for (let d = r.col, c = r.col + r.colCount; d < c; d++)
      s += this.getColWidth(d);
    const l = t - n.height;
    ct(this.ctx, o.left, 0, s, t), ct(this.ctx, 0, o.top, n.width, l);
    const f = [
      [n.width, n.height],
      [n.width, t]
    ];
    return qt(this.ctx, f), this.renderActiveCell(), {
      left: o.left,
      top: n.height,
      width: s,
      height: l
    };
  }
  renderSelectRow({ contentWidth: t }) {
    const n = this.eventData.headerSize, r = this.eventData.range, o = this.computeCellPosition(r);
    let s = 0;
    for (let d = r.row, c = r.row + r.rowCount; d < c; d++)
      s += this.getRowHeight(d);
    const l = t - n.width - Z;
    ct(this.ctx, o.left, 0, l, n.height), ct(this.ctx, 0, o.top, t, s);
    const f = [
      [o.left, n.height],
      [t, n.height]
    ];
    return qt(this.ctx, f), this.renderActiveCell(), {
      left: n.width,
      top: o.top,
      width: l,
      height: s
    };
  }
  renderActiveCell() {
    const t = this.eventData.range, n = this.getActiveRange({
      row: t.row,
      col: t.col,
      rowCount: 1,
      colCount: 1,
      sheetId: t.sheetId
    }).range;
    this.clearRect(n);
  }
}
let Go = null;
const D6 = {
  init(e) {
    Go = new x6(e.canvas), Mu(e.dpr);
  },
  resize(e) {
    Go?.resize(e);
  },
  async render(e, t) {
    const n = Go?.render(e);
    n && await t(n);
  },
  async computeFormulas(e, t) {
    return await B2(e, t);
  }
};
function oe(e = sn.MODEL) {
  return function(t, n, r) {
    const o = r.value;
    return r.value = function(...s) {
      const l = this;
      return l.getReadOnly() ? (jo("no auth method:", n), l.emit("toastMessage", {
        type: "error",
        message: k.t("no-login-editing")
      })) : l.getHooks().doc.transact(() => o.apply(l, s), e);
    }, r;
  };
}
var O6 = Object.defineProperty, k6 = Object.getOwnPropertyDescriptor, ie = (e, t, n, r) => {
  for (var o = k6(t, n), s = e.length - 1, l; s >= 0; s--)
    (l = e[s]) && (o = l(t, n, o) || o);
  return o && O6(t, n, o), o;
};
const L6 = {
  top: 0,
  left: 0,
  scrollLeft: 0,
  scrollTop: 0
};
class re extends ka {
  scrollValue = {};
  model;
  changeSet = /* @__PURE__ */ new Set();
  floatElementUuid = "";
  copyRange = void 0;
  isCut = !1;
  hooks;
  sheetViewSize = null;
  readonly = !1;
  canvasSize = {
    top: 0,
    left: 0,
    width: 0,
    height: 0
  };
  constructor(t, n) {
    super(), this.model = t, this.hooks = n;
  }
  setReadOnly(t) {
    this.readonly = t;
  }
  getReadOnly() {
    return this.readonly;
  }
  setCanvasSize(t) {
    this.canvasSize = { ...t };
  }
  getCanvasSize() {
    return { ...this.canvasSize };
  }
  setSheetViewSize() {
    const t = this.getHeaderSize(), n = this.getSheetInfo();
    if (!n)
      return;
    let { width: r, height: o } = t;
    for (let s = 0; s < n.colCount; s++) {
      const l = this.getCol(s);
      r += l.isHide ? 0 : l.len;
    }
    for (let s = 0; s < n.rowCount; s++) {
      const l = this.getRow(s);
      o += l.isHide ? 0 : l.len;
    }
    this.sheetViewSize = {
      width: r,
      height: o
    };
  }
  getSheetViewSize() {
    return this.sheetViewSize || this.setSheetViewSize(), this.sheetViewSize || { width: 0, height: 0 };
  }
  getHeaderSize() {
    const { row: t } = this.model.getScroll();
    return t > 9999 ? {
      width: Math.floor(ds * 2),
      height: fs
    } : {
      width: ds,
      height: fs
    };
  }
  clearHistory() {
    this.model.clearHistory();
  }
  getHooks() {
    return this.hooks;
  }
  validateDefinedName(t) {
    return this.model.validateDefinedName(t);
  }
  validateDrawing(t) {
    return this.model.validateDrawing(t);
  }
  validateRange(t) {
    return this.model.validateRange(t);
  }
  getCurrentSheetId() {
    return this.model.getCurrentSheetId();
  }
  getSheetList() {
    return this.model.getSheetList();
  }
  getSheetInfo(t) {
    return this.model.getSheetInfo(t);
  }
  transaction = (t, n) => this.model.transaction(t, n);
  emitChange() {
    const t = new Set(this.changeSet);
    this.changeSet = /* @__PURE__ */ new Set(), (t.has("customHeight") || t.has("customWidth")) && this.setSheetViewSize(), (t.has("worksheets") || t.has("definedNames") || t.has("currentSheetId")) && this.model.computeFormulas().then((n) => {
      n && (jo("computeFormulas"), this.emit("renderChange", { changeSet: /* @__PURE__ */ new Set(["worksheets"]) }));
    }), t.size > 0 && (jo("emitChange", t), this.emit("renderChange", { changeSet: t }));
  }
  getActiveRange(t) {
    return this.model.getActiveRange(t);
  }
  setNextActiveCell(t) {
    const { range: n, isMerged: r } = this.getActiveRange();
    let o = n.col, s = n.row;
    const l = this.getSheetInfo(n.sheetId);
    if (!l)
      return;
    const f = {
      ...n,
      rowCount: 1,
      colCount: 1
    };
    if (t === "left") {
      for (o--; o > 0 && this.getColWidth(o) <= 0; )
        o--;
      f.col = o;
    }
    if (t === "right") {
      for (r ? o = n.col + n.colCount : o++; o < l.colCount && this.getColWidth(o) <= 0; )
        o++;
      f.col = o;
    }
    if (t === "up") {
      for (s--; s > 0 && this.getRowHeight(s) <= 0; )
        s--;
      f.row = s;
    }
    if (t === "down") {
      for (r ? s = n.row + n.rowCount : s++; s < l.rowCount && this.getRowHeight(s) <= 0; )
        s++;
      f.row = s;
    }
    this.setActiveRange(f);
  }
  setActiveRange(t) {
    this.model.setActiveRange(t), this.changeSet.add("rangeMap");
    const n = {
      row: t.row,
      col: t.col,
      sheetId: t.sheetId || this.model.getCurrentSheetId(),
      rowCount: t.rowCount,
      colCount: t.colCount
    };
    this.emit("rangeChange", n), this.emitChange();
  }
  setCurrentSheetId(t) {
    this.model.setCurrentSheetId(t), this.changeSet.add("currentSheetId"), this.changeSet.add("customHeight"), this.changeSet.add("scroll"), this.emitChange();
  }
  getWorksheet(t) {
    return this.model.getWorksheet(t);
  }
  setWorksheet(t) {
    this.model.setWorksheet(t), this.changeSet.add("worksheets"), this.changeSet.add("cellStyle"), this.emitChange();
  }
  addSheet() {
    const t = this.model.addSheet();
    return this.changeSet.add("workbook"), this.changeSet.add("customHeight"), this.changeSet.add("currentSheetId"), this.emitChange(), t;
  }
  addFirstSheet() {
    return this.addSheet();
  }
  deleteSheet(t) {
    this.model.deleteSheet(t), this.changeSet.add("workbook"), this.changeSet.add("customHeight"), this.changeSet.add("currentSheetId"), this.emitChange();
  }
  updateSheetInfo(t, n) {
    this.model.updateSheetInfo(t, n), this.changeSet.add("workbook"), this.changeSet.add("customHeight"), this.emitChange();
  }
  hideSheet(t) {
    this.model.hideSheet(t), this.changeSet.add("workbook"), this.changeSet.add("customHeight"), this.changeSet.add("currentSheetId"), this.emitChange();
  }
  unhideSheet(t) {
    this.model.unhideSheet(t), this.changeSet.add("workbook"), this.changeSet.add("customHeight"), this.changeSet.add("currentSheetId"), this.emitChange();
  }
  renameSheet(t, n) {
    this.model.renameSheet(t, n), this.changeSet.add("workbook"), this.emitChange();
  }
  fromJSON(t) {
    this.changeSet.add("currentSheetId"), this.changeSet.add("cellStyle"), this.changeSet.add("definedNames"), this.changeSet.add("mergeCells"), this.changeSet.add("autoFilter"), this.changeSet.add("drawings"), this.changeSet.add("workbook"), this.changeSet.add("worksheets"), this.changeSet.add("rangeMap"), this.changeSet.add("scroll"), this.changeSet.add("customHeight"), this.changeSet.add("customWidth"), this.model.fromJSON(t), this.emitChange();
  }
  toJSON() {
    return this.model.toJSON();
  }
  setCell(t, n, r) {
    this.model.setCell(t, n, r), this.changeSet.add("worksheets"), this.changeSet.add("cellStyle"), this.emitChange();
  }
  deleteCell(t) {
    this.model.deleteCell(t), this.changeSet.add("worksheets"), this.changeSet.add("cellStyle"), this.emitChange();
  }
  setCellValue(t, n) {
    const r = this.getCell(n);
    if (us(r?.numberFormat) && typeof t == "string" && !isNaN(Date.parse(t))) {
      const o = Ev(new Date(t));
      if (o === null)
        return;
      this.model.setCellValue(o, n);
    } else
      this.model.setCellValue(t, n);
    this.changeSet.add("worksheets"), this.emitChange();
  }
  updateCellStyle(t, n) {
    this.model.updateCellStyle(t, n), this.changeSet.add("cellStyle"), this.emitChange();
  }
  getCell = (t) => this.model.getCell(t);
  canRedo() {
    return this.model.canRedo();
  }
  canUndo() {
    return this.model.canUndo();
  }
  undo() {
    this.model.undo(), this.changeSet.add("undo"), this.emitChange();
  }
  redo() {
    this.model.redo(), this.changeSet.add("redo"), this.emitChange();
  }
  getCol(t, n) {
    return this.model.getCol(t, n);
  }
  getColWidth(t, n) {
    const r = this.model.getCol(t, n);
    return r.isHide ? 0 : r.len;
  }
  setColWidth(t, n, r) {
    this.model.setColWidth(t, n, r), this.changeSet.add("customWidth"), this.emitChange();
  }
  getRow(t, n) {
    return this.model.getRow(t, n);
  }
  getRowHeight(t, n) {
    const r = this.model.getRow(t, n);
    return r.isHide ? 0 : r.len;
  }
  setRowHeight(t, n, r) {
    this.model.setRowHeight(t, n, r), this.changeSet.add("customHeight"), this.emitChange();
  }
  getCellSize(t) {
    return t.sheetId = t.sheetId || this.getCurrentSheetId(), this.getRangeSize(t);
  }
  getRangeSize(t) {
    let { row: n, col: r, colCount: o, rowCount: s } = t, l = n, f = r, d = n + s, c = r + o;
    const i = t.sheetId || this.getCurrentSheetId(), a = this.getSheetInfo(i);
    if (!a)
      return { width: 0, height: 0 };
    An(t) ? (f = 0, c = a.colCount, l = 0, d = a.rowCount) : Yt(t) ? (l = 0, d = a.rowCount) : Wt(t) && (f = 0, c = a.colCount);
    let u = 0, h = 0;
    for (; l < d; l++)
      h += this.getRowHeight(l, i);
    for (; f < c; f++)
      u += this.getColWidth(f, i);
    return { width: u, height: h };
  }
  computeCellPosition(t) {
    const { row: n, col: r } = t, o = t.sheetId || this.model.getCurrentSheetId(), s = this.getHeaderSize(), l = this.getScroll(o);
    let f = s.width, d = s.height, c = l.row, i = l.col;
    if (r >= l.col)
      for (; i < r; )
        f += this.getColWidth(i, o), i++;
    else
      for (f = -s.width; i > r; )
        f -= this.getColWidth(i, o), i--;
    if (n >= l.row)
      for (; c < n; )
        d += this.getRowHeight(c, o), c++;
    else
      for (d = -s.height; c > n; )
        d -= this.getRowHeight(c, o), c--;
    return {
      top: d,
      left: f
    };
  }
  addRow(t, n, r = !1) {
    this.model.addRow(t, n, r), this.changeSet.add("customHeight"), this.emitChange();
  }
  addCol(t, n, r = !1) {
    this.model.addCol(t, n, r), this.changeSet.add("customWidth"), this.emitChange();
  }
  deleteCol(t, n) {
    this.model.deleteCol(t, n), this.changeSet.add("customWidth"), this.emitChange();
  }
  deleteRow(t, n) {
    this.model.deleteRow(t, n), this.changeSet.add("customHeight"), this.emitChange();
  }
  hideCol(t, n) {
    this.model.hideCol(t, n), this.changeSet.add("customWidth"), this.emitChange();
  }
  hideRow(t, n) {
    this.model.hideRow(t, n), this.changeSet.add("customHeight"), this.emitChange();
  }
  unhideRow(t, n) {
    this.model.unhideRow(t, n), this.changeSet.add("customHeight"), this.emitChange();
  }
  unhideCol(t, n) {
    this.model.unhideCol(t, n), this.changeSet.add("customWidth"), this.emitChange();
  }
  getScroll(t) {
    const n = t || this.model.getCurrentSheetId(), r = this.model.getScroll(n);
    return { ...this.scrollValue[n] || { ...L6 }, ...r };
  }
  setScroll(t, n) {
    const r = n || this.model.getCurrentSheetId();
    let { row: o, col: s, ...l } = t;
    const f = this.getSheetInfo(r);
    f && (o < 0 ? o = 0 : o >= f.rowCount - 1 && (o = f.rowCount - 1), s < 0 ? s = 0 : s >= f.colCount - 1 && (s = f.colCount - 1)), this.scrollValue[r] = l, this.model.setScroll({
      row: o,
      col: s
    }) && this.changeSet.add("scroll"), this.emitChange();
  }
  getPasteRange(t) {
    const n = t.length, r = this.getActiveRange().range, o = Math.max(...t.map((s) => s.length));
    return {
      ...r,
      rowCount: n,
      colCount: o
    };
  }
  parseText(t) {
    const n = Ea(t), r = this.getPasteRange(n);
    if (n.length > 0)
      return this.model.setCell(n, [], r), r;
  }
  parseHTML(t) {
    const { textList: n, styleList: r, rowMap: o, colMap: s } = wu(t), l = this.getPasteRange(n);
    for (const [f, d] of o.entries())
      this.model.setRowHeight(f, d);
    for (const [f, d] of s.entries())
      this.model.setColWidth(f, d);
    if (n.length > 0 || r.length > 0)
      return this.model.setCell(n, r, l), l;
  }
  async getCopyData(t) {
    const { range: n, isMerged: r } = this.getActiveRange(), { row: o, col: s, rowCount: l, colCount: f } = n, d = [], c = [];
    let i = 1;
    const a = [], u = this.model.getCurrentSheetId(), h = /* @__PURE__ */ new Set();
    for (let v = o, C = o + l; v < C; v++) {
      const M = [], R = [], A = ps(this.getRow(v).len, "");
      for (let x = s, L = s + f; x < L; x++) {
        const D = this.model.getCell({
          row: v,
          col: x,
          rowCount: 1,
          colCount: 1,
          sheetId: u
        });
        if (!D)
          continue;
        const W = Nn(D.value, D?.numberFormat);
        M.push(W);
        const J = ps(this.getCol(x).len, ""), ne = `height=${A} width=${J} style='height:${A}pt;width:${J}pt;'`;
        if (h.add(`<col width=${J} style='width:${J}pt;'>`), D) {
          let K = Cu(D);
          W || (K += "mso-pattern:black none;");
          const q = `xl${i++}`;
          a.push(`.${q}{${K}}`), R.push(`<td ${ne} class="${q}"> ${W} </td>`);
        } else
          R.push(`<td ${ne}> ${W} </td>`);
        if (r)
          break;
      }
      if (d.push(M), c.push(
        `<tr height=${A} style='height:${A}pt;'>${R.join(`
`)}</tr>`
      ), r)
        break;
    }
    const g = {
      range: n,
      type: t,
      floatElementUuid: this.floatElementUuid
    }, m = JSON.stringify(g), p = Lu(
      a.join(`
`),
      `
      ${Array.from(h).join(`
`)}
      ${c.join(`
`)}`,
      Fu(m)
    ), E = `${d.map((v) => v.join("	")).join(`\r
`)}\r
`, b = {
      [Be]: E,
      [ze]: p,
      [Ge]: g,
      [Qe]: null
    };
    let N = "";
    if (this.floatElementUuid) {
      const C = this.getDrawingList(this.getCurrentSheetId()).find((M) => M.uuid === this.floatElementUuid);
      if (C)
        if (C.type === "floating-picture") {
          const M = document.querySelector(
            `canvas[data-uuid="${C.uuid}"]`
          );
          M && (M.toBlob((R) => {
            R && (b[Qe] = R);
          }), N = M.toDataURL());
        } else
          N = C.imageSrc || "", b[Qe] = qu(N);
    }
    return [b, N];
  }
  pasteFloatElement(t, n) {
    if (!t)
      return !1;
    if (n) {
      const c = this.getActiveRange().range;
      return this.model.updateDrawing(t, {
        fromCol: c.col,
        fromRow: c.row,
        marginX: 0,
        marginY: 0
      }), this.floatElementUuid = "", this.isCut = !1, this.emitChange(), !0;
    }
    const o = this.getDrawingList(this.getCurrentSheetId()).find((c) => c.uuid === t);
    if (!o)
      return !1;
    const s = this.getCellSize({
      row: o.fromRow,
      col: o.fromCol,
      rowCount: 1,
      colCount: 1,
      sheetId: o.sheetId
    });
    let { marginX: l, marginY: f } = o;
    const d = 14;
    return l + d < s.width ? l += d : l - d >= 0 && (l -= d), f + d < s.width ? f += d : f - d >= 0 && (f -= d), this.addDrawing({
      ...o,
      uuid: fr(),
      marginX: l,
      marginY: f
    }), !0;
  }
  async paste(t) {
    if (this.floatElementUuid && this.pasteFloatElement(this.floatElementUuid, this.isCut))
      return this.changeSet.add("drawings"), this.emitChange(), Promise.resolve();
    await this.basePaste(t), this.changeSet.add("worksheets"), this.changeSet.add("cellStyle"), this.emitChange();
  }
  async basePaste(t) {
    let n = "", r = "", o = null, s = null;
    if (t) {
      n = t.clipboardData?.getData(ze) || "", r = t.clipboardData?.getData(Be) || "";
      const l = t.clipboardData?.getData(Ge) || "";
      l && (o = JSON.parse(l)), o && o.type === "cut" && (t.clipboardData?.setData(ze, ""), t.clipboardData?.setData(Be, ""), t.clipboardData?.setData(Ge, ""), t.clipboardData?.setData(Qe, "")), t.clipboardData?.files && (s = t.clipboardData?.files);
    } else {
      const l = await Ra();
      n = l[ze], r = l[Be], l[Ge] && (o = l[Ge]), o && o.type === "cut" && Es(
        {
          [Be]: "",
          [ze]: "",
          [Ge]: null,
          [Qe]: null
        },
        "copy"
      );
    }
    if (s) {
      const l = [];
      for (const f of s) {
        if (f.size <= 0)
          continue;
        let d = f.name;
        const c = f.type.slice(6);
        d = d.slice(0, -(c.length + 1));
        const i = await Mi(f, !0);
        if (!i)
          continue;
        const a = await yi(i), u = this.getActiveRange().range;
        l.push({
          width: a.width,
          height: a.height,
          originHeight: a.height,
          originWidth: a.width,
          title: d,
          type: "floating-picture",
          uuid: fr(),
          imageSrc: i,
          sheetId: u.sheetId,
          fromRow: u.row,
          fromCol: u.col,
          marginX: 0,
          marginY: 0
        });
      }
      if (l.length > 0) {
        this.addDrawing(...l);
        return;
      }
    }
    if (!(o && o.floatElementUuid && this.pasteFloatElement(o.floatElementUuid, o.type === "cut"))) {
      if (o) {
        const l = this.model.pasteRange(
          o.range,
          o.type === "cut"
        );
        this.setActiveRange(l);
        return;
      }
      if (n = n.trim(), r = r.trim(), n) {
        const l = this.parseHTML(n);
        if (l) {
          this.setActiveRange(l);
          return;
        }
      }
      if (r) {
        const l = this.parseText(r);
        if (l) {
          this.setActiveRange(l);
          return;
        }
      }
      if (this.copyRange) {
        const l = this.model.pasteRange(this.copyRange, this.isCut);
        this.setActiveRange(l), this.isCut && (this.copyRange = void 0, this.isCut = !1);
        return;
      }
    }
  }
  copy(t) {
    return this.copyRange = this.getActiveRange().range, this.isCut = !1, this.baseCopy("copy", t);
  }
  cut(t) {
    return this.copyRange = this.getActiveRange().range, this.isCut = !0, this.baseCopy("cut", t);
  }
  async baseCopy(t, n) {
    const [r, o] = await this.getCopyData(t);
    n ? (n.clipboardData?.setData(ze, r[ze]), n.clipboardData?.setData(Be, r[Be]), n.clipboardData?.setData(
      Ge,
      r[Ge] ? JSON.stringify(r[Ge]) : ""
    ), n.clipboardData?.setData(Qe, o)) : Es(r, t), this.changeSet.add("antLine"), this.emitChange();
  }
  getCopyRange() {
    return this.copyRange;
  }
  setCopyRange(t) {
    this.copyRange = t;
  }
  deleteAll(t) {
    this.model.deleteAll(t), this.changeSet.add("definedNames"), this.changeSet.add("mergeCells"), this.changeSet.add("cellStyle"), this.changeSet.add("autoFilter"), this.changeSet.add("drawings"), this.changeSet.add("customHeight"), this.changeSet.add("customWidth"), this.changeSet.add("worksheets"), this.emitChange();
  }
  getDefineName(t) {
    return this.model.getDefineName(t);
  }
  getDefineNameList() {
    return this.model.getDefineNameList();
  }
  setDefineName(t, n) {
    const r = this.model.setDefineName(t, n);
    return this.changeSet.add("definedNames"), this.emitChange(), r;
  }
  checkDefineName(t) {
    return this.model.checkDefineName(t);
  }
  getDrawingList(t) {
    return this.model.getDrawingList(t);
  }
  addDrawing(...t) {
    this.model.addDrawing(...t), this.changeSet.add("drawings"), this.emitChange();
  }
  updateDrawing(t, n) {
    this.model.updateDrawing(t, n), this.changeSet.add("drawings"), this.emitChange();
  }
  deleteDrawing(t) {
    this.model.deleteDrawing(t), this.changeSet.add("drawings"), this.emitChange();
  }
  getMergeCellList(t) {
    return this.model.getMergeCellList(t);
  }
  addMergeCell(t, n) {
    this.model.addMergeCell(t, n), this.changeSet.add("mergeCells"), this.emitChange();
  }
  deleteMergeCell(t) {
    this.model.deleteMergeCell(t), this.changeSet.add("mergeCells"), this.emitChange();
  }
  setFloatElementUuid(t) {
    this.floatElementUuid = t;
  }
  getFilter(t) {
    return this.model.getFilter(t);
  }
  addFilter(t) {
    this.model.addFilter(t), this.changeSet.add("autoFilter"), this.emitChange();
  }
  deleteFilter(t) {
    this.model.deleteFilter(t), this.changeSet.add("autoFilter"), this.emitChange();
  }
  updateFilter(t, n) {
    this.model.updateFilter(t, n), this.changeSet.add("autoFilter"), this.emitChange();
  }
}
ie([
  oe()
], re.prototype, "setWorksheet");
ie([
  oe()
], re.prototype, "addSheet");
ie([
  oe(sn.SKIP_UNDO_REDO)
], re.prototype, "addFirstSheet");
ie([
  oe()
], re.prototype, "deleteSheet");
ie([
  oe()
], re.prototype, "updateSheetInfo");
ie([
  oe()
], re.prototype, "hideSheet");
ie([
  oe()
], re.prototype, "unhideSheet");
ie([
  oe()
], re.prototype, "renameSheet");
ie([
  oe(sn.SKIP_UNDO_REDO)
], re.prototype, "fromJSON");
ie([
  oe()
], re.prototype, "setCell");
ie([
  oe()
], re.prototype, "deleteCell");
ie([
  oe()
], re.prototype, "setCellValue");
ie([
  oe()
], re.prototype, "updateCellStyle");
ie([
  oe()
], re.prototype, "setColWidth");
ie([
  oe()
], re.prototype, "setRowHeight");
ie([
  oe()
], re.prototype, "addRow");
ie([
  oe()
], re.prototype, "addCol");
ie([
  oe()
], re.prototype, "deleteCol");
ie([
  oe()
], re.prototype, "deleteRow");
ie([
  oe()
], re.prototype, "hideCol");
ie([
  oe()
], re.prototype, "hideRow");
ie([
  oe()
], re.prototype, "unhideRow");
ie([
  oe()
], re.prototype, "unhideCol");
ie([
  oe()
], re.prototype, "paste");
ie([
  oe()
], re.prototype, "setCopyRange");
ie([
  oe()
], re.prototype, "deleteAll");
ie([
  oe()
], re.prototype, "setDefineName");
ie([
  oe()
], re.prototype, "addDrawing");
ie([
  oe()
], re.prototype, "updateDrawing");
ie([
  oe()
], re.prototype, "deleteDrawing");
ie([
  oe()
], re.prototype, "addMergeCell");
ie([
  oe()
], re.prototype, "deleteMergeCell");
ie([
  oe()
], re.prototype, "addFilter");
ie([
  oe()
], re.prototype, "deleteFilter");
ie([
  oe()
], re.prototype, "updateFilter");
class F6 {
  model;
  currentSheetId = "";
  constructor(t) {
    this.model = t;
  }
  get workbook() {
    return this.model.getRoot().get("workbook");
  }
  validateSheet(t) {
    return !(!t || !t.sheetId || !t.name || typeof t.isHide != "boolean" || typeof t.sort != "number" || typeof t.colCount != "number" || typeof t.rowCount != "number" || t.rowCount > mr || t.colCount > hr || t.rowCount <= 0 || t.colCount <= 0);
  }
  fromJSON(t) {
    const n = t.workbook || {}, r = t.currentSheetId || "", o = new we();
    for (const f of Object.values(n)) {
      if (!this.validateSheet(f))
        continue;
      const d = new we(Object.entries(f));
      o.set(f.sheetId, d);
    }
    this.model.getRoot().set("workbook", o);
    let s = this.getSheetId() || "";
    const l = o.get(r);
    l && !l?.get("isHide") && (s = r), this.setCurrentSheetId(s);
  }
  updateSheetInfo(t, n) {
    const r = n || this.getCurrentSheetId(), o = this.workbook?.get(r);
    if (!o)
      return;
    const s = Object.keys(t);
    for (const l of s)
      l === "sheetId" || o.get(l) === t[l] || o.set(l, t[l]);
  }
  getSheetList() {
    if (!this.workbook)
      return [];
    const t = Array.from(this.workbook.values()).map((n) => n.toJSON());
    return t.sort((n, r) => n.sort - r.sort), t.slice();
  }
  addSheet() {
    this.workbook || this.model.getRoot().set("workbook", new we());
    const t = this.getSheetList(), r = {
      ...mu(t),
      isHide: !1,
      colCount: ou,
      rowCount: ru
    };
    return this.workbook.set(
      r.sheetId,
      new we(Object.entries(r))
    ), r;
  }
  deleteSheet(t) {
    const n = t || this.getCurrentSheetId();
    this.checkSheetSize(t) && this.workbook?.delete(n);
  }
  hideSheet(t) {
    this.checkSheetSize(t) && this.updateSheetInfo({ isHide: !0 }, t);
  }
  unhideSheet(t) {
    this.updateSheetInfo({ isHide: !1 }, t);
  }
  renameSheet(t, n) {
    if (!t) {
      this.model.emit("toastMessage", {
        type: "error",
        message: k.t("the-value-cannot-be-empty")
      });
      return;
    }
    const r = n || this.getCurrentSheetId(), s = this.getSheetList().find((l) => l.name === t);
    if (s) {
      if (s.sheetId === r)
        return;
      this.model.emit("toastMessage", {
        type: "error",
        message: k.t("sheet-name-is-duplicate")
      });
      return;
    }
    this.updateSheetInfo({ name: t }, r);
  }
  getSheetInfo(t) {
    const n = t || this.getCurrentSheetId(), r = this.workbook?.get(n);
    if (r)
      return { ...r.toJSON() };
  }
  setCurrentSheetId(t) {
    this.workbook?.get(t) ? this.currentSheetId = t : this.currentSheetId = this.getSheetId() || "";
  }
  getCurrentSheetId() {
    return this.workbook?.get(this.currentSheetId) ? this.currentSheetId : (this.currentSheetId = this.getSheetId() || "", this.currentSheetId);
  }
  getSheetId() {
    return this.getSheetList().filter((r) => !r.isHide)[0]?.sheetId;
  }
  deleteAll(t) {
    this.deleteSheet(t);
  }
  checkSheetSize(t) {
    const n = t || this.getCurrentSheetId();
    return this.workbook?.get(n) ? this.getSheetList().filter((s) => !s.isHide).length < 2 ? (this.model.emit("toastMessage", {
      type: "error",
      message: k.t("a-workbook-must-contains-at-least-one-visible-worksheet")
    }), !1) : !0 : !1;
  }
}
class P6 {
  model;
  rangeMap = {};
  constructor(t) {
    this.model = t;
  }
  validateRange(t) {
    if (!t)
      return !1;
    t.sheetId = t.sheetId || this.model.getCurrentSheetId();
    const n = this.model.getSheetInfo(t.sheetId);
    return !(!n || t.row < 0 || t.col < 0 || t.colCount < 0 || t.rowCount < 0 || t.row >= n.rowCount || t.col >= n.colCount || t.colCount > n.colCount || t.rowCount > n.rowCount || t.row + t.rowCount > n.rowCount || t.col + t.colCount > n.colCount);
  }
  fromJSON(t) {
    const n = t.rangeMap || {}, r = {};
    for (const o of Object.values(n))
      o.sheetId = o.sheetId || this.model.getCurrentSheetId(), this.model.validateRange(o) && (r[o.sheetId] = { ...o });
    this.rangeMap = r;
  }
  toJSON() {
    return { ...this.rangeMap };
  }
  getActiveRange() {
    const t = this.model.getCurrentSheetId(), n = this.rangeMap[t];
    return n ? (n.sheetId = n.sheetId || t, {
      range: { ...n },
      isMerged: !1
    }) : {
      range: {
        row: 0,
        col: 0,
        rowCount: 1,
        colCount: 1,
        sheetId: t
      },
      isMerged: !1
    };
  }
  setActiveRange(t) {
    t.sheetId = t.sheetId || this.model.getCurrentSheetId(), this.validateRange(t) && (this.rangeMap[t.sheetId] = { ...t });
  }
  deleteAll(t) {
    delete this.rangeMap[t || this.model.getCurrentSheetId()];
  }
}
class U6 {
  model;
  constructor(t) {
    this.model = t;
  }
  get drawings() {
    return this.model.getRoot().get("drawings");
  }
  validateDrawing(t) {
    if (!t?.uuid || !this.model.getSheetInfo(t.sheetId))
      return !1;
    if (t.type === "floating-picture") {
      if (!t.imageSrc)
        return !1;
    } else if (t.chartType === void 0 || t.chartRange === void 0 || Ko.findIndex((r) => r.value === t.chartType) < 0 || !this.model.validateRange(t.chartRange))
      return !1;
    return !0;
  }
  fromJSON(t) {
    const n = t.drawings || {}, r = new we();
    for (const [o, s] of Object.entries(n)) {
      if (!this.validateDrawing(s) || !o)
        return;
      const l = new we(Object.entries(s));
      r.set(o, l);
    }
    this.model.getRoot().set("drawings", r);
  }
  getDrawingList(t) {
    if (!this.drawings)
      return [];
    const n = t || this.model.getCurrentSheetId();
    return Array.from(this.drawings.values()).filter((o) => o.get("sheetId") === n).map((o) => o.toJSON()).slice();
  }
  addDrawing(t) {
    if (this.drawings?.get(t.uuid))
      return this.model.emit("toastMessage", {
        type: "error",
        message: k.t("uuid-is-duplicate")
      });
    if (this.validateDrawing(t)) {
      if (t.type === "chart") {
        const r = t.chartRange;
        let o = !1;
        const s = this.model.getSheetInfo(r.sheetId);
        if (!s)
          return this.model.emit("toastMessage", {
            type: "error",
            message: k.t("sheet-is-not-exist")
          });
        if (Tn(
          r,
          s?.rowCount,
          s?.colCount,
          (l, f) => this.model.getCell({
            row: l,
            col: f,
            rowCount: 1,
            colCount: 1,
            sheetId: ""
          })?.value ? (o = !0, !0) : !1
        ), !o)
          return this.model.emit("toastMessage", {
            type: "error",
            message: k.t("cells-must-contain-data")
          });
        t.chartRange = un(t.chartRange);
      } else t.type === "floating-picture" && typeof t.imageAngle != "number" && (t.imageAngle = 0);
      this.drawings || this.model.getRoot().set("drawings", new we()), this.drawings.set(
        t.uuid,
        new we(Object.entries(t))
      );
    }
  }
  updateDrawing(t, n) {
    const r = Object.keys(n), o = this.drawings?.get(t);
    if (!(r.length === 0 || !o)) {
      for (const s of r)
        if (o.get(s) !== n[s] && s === "chartType" && Ko.findIndex(
          (f) => f.value === n.chartType
        ) < 0)
          return this.model.emit("toastMessage", {
            type: "error",
            message: k.t("unsupported-chart-types")
          });
      for (const s of r) {
        const l = n[s];
        o.get(s) !== l && o.set(
          s,
          s === "chartRange" ? un(l) : l
        );
      }
    }
  }
  deleteDrawing(t) {
    this.drawings?.delete(t);
  }
  addCol(t, n, r = !1) {
    if (!this.drawings)
      return;
    const o = this.model.getSheetInfo();
    if (!o)
      return;
    const s = o.colCount;
    for (const [l, f] of this.drawings.entries()) {
      const d = f.toJSON(), c = {}, i = r ? t + 1 : t;
      if (d.fromCol >= i && (c.fromCol = Math.min(s - 1, d.fromCol + n)), d.type === "chart" && d.chartRange.col >= i) {
        const a = {
          ...d.chartRange
        }, u = a.col;
        u + n <= s - 1 ? a.col += n : (a.col = s - 1, a.colCount = Math.max(
          s - 1 - n - u + a.colCount,
          1
        )), c.chartRange = a;
      }
      this.updateDrawing(l, c);
    }
  }
  addRow(t, n, r = !1) {
    if (n <= 0 || !this.drawings)
      return;
    const o = this.model.getSheetInfo();
    if (!o)
      return;
    const s = o.rowCount;
    for (const [l, f] of this.drawings.entries()) {
      const d = f.toJSON(), c = {}, i = r ? t : t + 1;
      if (d.fromRow >= i && (c.fromRow = Math.min(s - 1, d.fromRow + n)), d.type === "chart" && d.chartRange.row >= i) {
        const a = {
          ...d.chartRange
        }, u = a.row;
        u + n <= s - 1 ? a.row += n : (a.row = s - 1, a.rowCount = Math.max(
          s - 1 - n - u + a.rowCount,
          1
        )), c.chartRange = a;
      }
      this.updateDrawing(l, c);
    }
  }
  deleteCol(t, n) {
    if (this.drawings)
      for (const [r, o] of this.drawings.entries()) {
        const s = o.toJSON(), l = {};
        if (s.fromCol >= t && (l.fromCol = Math.max(s.fromCol - n, 0)), s.type === "chart" && s.chartRange.col >= t) {
          const f = {
            ...s.chartRange
          }, d = f.col;
          d >= n ? f.col -= n : (f.col = 0, f.colCount = Math.max(d - n + f.colCount, 1)), l.chartRange = f;
        }
        this.updateDrawing(r, l);
      }
  }
  /* jscpd:ignore-start */
  deleteRow(t, n) {
    if (this.drawings)
      for (const [r, o] of this.drawings.entries()) {
        const s = o.toJSON(), l = {};
        if (s.fromRow >= t && (l.fromRow = Math.max(s.fromRow - n, 0)), s.type === "chart" && s.chartRange.row >= t) {
          const f = {
            ...s.chartRange
          }, d = f.row;
          d >= n ? f.row = d - n : (f.row = 0, f.rowCount = Math.max(d - n + f.rowCount, 1)), l.chartRange = f;
        }
        this.updateDrawing(r, l);
      }
  }
  /* jscpd:ignore-end */
  deleteAll(t) {
    if (!this.drawings)
      return;
    const n = t || this.model.getCurrentSheetId();
    for (const [r, o] of this.drawings.entries())
      o.get("sheetId") === n && this.drawings.delete(r);
  }
}
class _6 {
  model;
  constructor(t) {
    this.model = t;
  }
  get definedNames() {
    return this.model.getRoot().get("definedNames");
  }
  validateDefinedName(t) {
    if (!t)
      return !1;
    const n = Ot(t, (r) => this.model.getSheetList().find((l) => l.name === r)?.sheetId || "");
    if (n) {
      const r = this.model.getSheetInfo(n.sheetId);
      if (r && n.col < r.colCount && n.row < r.rowCount)
        return !1;
    }
    return !!(Ci.test(t) && t.length <= da);
  }
  fromJSON(t) {
    const n = t.definedNames || {}, r = new we();
    for (const [o, s] of Object.entries(n))
      !this.model.validateRange(s) || !this.validateDefinedName(o) || r.set(o, un(s));
    this.model.getRoot().set("definedNames", r);
  }
  getDefineNameList() {
    if (!this.definedNames)
      return [];
    const t = Array.from(this.definedNames.entries());
    return t.length === 0 ? [] : t.map((n) => ({
      name: n[0],
      range: n[1]
    }));
  }
  getDefineName(t) {
    if (!this.definedNames)
      return "";
    const n = t.sheetId || this.model.getCurrentSheetId();
    for (const [r, o] of this.definedNames.entries())
      if (o.row === t.row && o.col === t.col && o.sheetId === n)
        return r;
    return "";
  }
  setDefineName(t, n) {
    if (!n)
      return !1;
    const r = this.getDefineName(t);
    if (r === n)
      return !1;
    this.definedNames || this.model.getRoot().set("definedNames", new we());
    const o = {
      row: t.row,
      col: t.col,
      sheetId: t.sheetId || this.model.getCurrentSheetId(),
      colCount: 1,
      rowCount: 1
    };
    return this.definedNames.set(n, o), r && this.definedNames?.delete(r), !0;
  }
  checkDefineName(t) {
    const n = this.definedNames?.get(t);
    if (n)
      return { ...n };
  }
  deleteAll(t) {
    if (!this.definedNames)
      return;
    const n = t || this.model.getCurrentSheetId();
    for (const [r, o] of this.definedNames.entries())
      o.sheetId === n && this.definedNames.delete(r);
  }
}
class B6 {
  model;
  worker;
  constructor(t, n) {
    this.model = t, this.worker = n;
  }
  get worksheets() {
    return this.model.getRoot().get("worksheets");
  }
  fromJSON(t) {
    const n = t.worksheets || {}, r = [];
    for (const [o, s] of Object.entries(n)) {
      const l = jn(o);
      if (!l)
        continue;
      const f = {
        ...l,
        ...s
      };
      r.push(f);
    }
    this.model.getRoot().set("worksheets", new we()), this.setWorksheet(r);
  }
  addMergeCell(t, n) {
    const r = t.sheetId || this.model.getCurrentSheetId(), o = n === Ct.MERGE_CONTENT, s = [], l = this.model.getSheetInfo(t.sheetId);
    if (!l || (Tn(t, l.rowCount, l.colCount, (a, u) => {
      const h = this.getCell({
        row: a,
        col: u,
        rowCount: 1,
        colCount: 1,
        sheetId: r
      });
      return h && s.push({ value: { ...h }, row: a, col: u }), !1;
    }), s.length === 0))
      return;
    let f = [
      Nn(s[0].value.value, s[0].value?.numberFormat)
    ];
    for (let a = 1; a < s.length; a++) {
      const u = s[a], h = nt(r, u.row, u.col), g = { ...u.value };
      o && f.push(Nn(g.value, g?.numberFormat)), this.worksheets?.delete(h);
    }
    s.some((a) => a.value.formula) && (f = [f[0]]);
    let d = $n(s[0], [
      "row",
      "col",
      "value"
    ]), c = s[0].value.value;
    n === Ct.MERGE_CELL || n === Ct.MERGE_CENTER ? d.horizontalAlign = xe.CENTER : n === Ct.MERGE_CONTENT && (d.isWrapText = !0);
    const i = {
      row: t.row,
      col: t.col,
      rowCount: 1,
      colCount: 1,
      sheetId: r
    };
    if (o ? c = f.join(mt) : c = f[0], t.row === s[0].row && t.col === s[0].col)
      o && this.setCellValue(c, i);
    else {
      const a = nt(r, s[0].row, s[0].col);
      this.worksheets?.delete(a), this.setCellValue(c, i);
    }
    this.updateCellStyle(d, i);
  }
  addRow(t, n, r = !1) {
    const o = this.getCoordinateList(), s = this.model.getCurrentSheetId(), l = r ? t : t + 1;
    for (let f = o.length - 1; f >= 0; f--) {
      const d = o[f];
      if (d.row < l)
        continue;
      const c = nt(s, d.row, d.col), i = $n(d, ["sheetId", "row", "col"]);
      this.worksheets?.delete(c), this.setCellModel(
        { sheetId: s, row: d.row + n, col: d.col },
        i
      );
    }
  }
  deleteRow(t, n) {
    const r = this.model.getCurrentSheetId(), o = this.getCoordinateList();
    for (const s of o) {
      if (s.row < t)
        continue;
      const l = nt(r, s.row, s.col), f = $n(s, ["sheetId", "row", "col"]);
      s.row >= t + n && this.setCellModel(
        { sheetId: r, row: s.row - n, col: s.col },
        f
      ), this.worksheets?.delete(l);
    }
  }
  addCol(t, n, r = !1) {
    const o = this.getCoordinateList(!0), s = this.model.getCurrentSheetId(), l = r ? t + 1 : t;
    for (let f = o.length - 1; f >= 0; f--) {
      const d = o[f];
      if (d.col < l)
        continue;
      const c = nt(s, d.row, d.col), i = $n(d, ["sheetId", "row", "col"]);
      this.worksheets?.delete(c), this.setCellModel(
        { sheetId: s, row: d.row, col: d.col + n },
        i
      );
    }
  }
  deleteCol(t, n) {
    const r = this.getCoordinateList(!0), o = this.model.getCurrentSheetId();
    for (const s of r) {
      if (s.col < t)
        continue;
      const l = nt(o, s.row, s.col), f = $n(s, ["sheetId", "row", "col"]);
      s.col >= t + n && this.setCellModel(
        { sheetId: o, row: s.row, col: s.col - n },
        f
      ), this.worksheets?.delete(l);
    }
  }
  getWorksheet(t) {
    const n = t || this.model.getCurrentSheetId();
    return this.getWorkData(n);
  }
  getWorkData(t = "") {
    const n = [];
    if (!this.worksheets)
      return n;
    for (const [r, o] of this.worksheets.entries()) {
      const s = jn(r);
      if (!s)
        continue;
      const l = o.toJSON();
      t && s.sheetId === t ? n.push({ ...s, ...l }) : n.push({ ...s, ...l });
    }
    return n;
  }
  setWorksheet(t) {
    const n = this.validateSheetData(t);
    for (const r of n) {
      const { sheetId: o, row: s, col: l, ...f } = r;
      this.setCellModel({ sheetId: o, row: s, col: l }, f);
    }
  }
  getCell(t) {
    const n = this.getCellModel(t);
    if (n)
      return {
        ...n.toJSON()
      };
  }
  deleteCell(t) {
    const n = t.sheetId || this.model.getCurrentSheetId(), r = this.model.getSheetInfo(n);
    if (!r || !this.worksheets)
      return;
    const o = [];
    Tn(t, r.rowCount, r.colCount, (s, l) => {
      const f = nt(n, s, l);
      return o.push(f), !1;
    });
    for (const s of o)
      this.worksheets.delete(s);
  }
  async setCell(t, n, r) {
    const { row: o, col: s } = r;
    for (let l = 0; l < t.length; l++)
      for (let f = 0; f < t[l].length; f++) {
        const d = {
          row: o + l,
          col: s + f
        };
        n[l] && n[l][f] && this.updateStyle(n[l][f], {
          ...d,
          sheetId: r.sheetId,
          rowCount: 1,
          colCount: 1
        }), this.setCellValue(t[l][f], { ...r, ...d });
      }
  }
  setCellValue(t, n) {
    const r = this.getCellModel(n, !0);
    typeof t == "string" && Na(t) && r?.get("numberFormat") !== ma ? this.setCellFormula(t, n) : (r?.get("formula") && this.setCellFormula("", n), this.setValue(t, n));
  }
  updateCellStyle(t, n) {
    if (On(t))
      return;
    const r = this.model.getSheetInfo(n.sheetId);
    r && Tn(n, r.rowCount, r.colCount, (o, s) => (this.updateStyle(t, {
      row: o,
      col: s,
      sheetId: n.sheetId,
      rowCount: 1,
      colCount: 1
    }), !1));
  }
  pasteRange(t, n) {
    const r = this.model.getCurrentSheetId(), o = this.model.getActiveRange().range, { row: s, col: l, rowCount: f, colCount: d, sheetId: c } = t, i = c || r, a = this.getWorksheet(i), u = {
      ...o,
      rowCount: f,
      colCount: d
    };
    if (a.length === 0)
      return u;
    const h = this.model.getSheetInfo(t.sheetId);
    return h && Tn(t, h.rowCount, h.colCount, (g, m) => {
      const p = nt(i, g, m), E = this.worksheets?.get(p), b = o.row + (g - s), N = o.col + (m - l), v = E ? { ...E.toJSON() } : { value: "" };
      return this.setCellModel(
        { sheetId: u.sheetId, row: b, col: N },
        v
      ), n && this.worksheets?.delete(p), !1;
    }), u;
  }
  deleteAll(t) {
    const n = t || this.model.getCurrentSheetId();
    if (this.worksheets)
      for (const r of this.worksheets.keys())
        jn(r)?.sheetId === n && this.worksheets.delete(r);
  }
  updateStyle(t, n) {
    const r = this.getCellModel(n, !0), o = Object.keys(t);
    for (const s of o) {
      const l = r.get(s), f = t[s];
      io(l, f) || r.set(s, f);
    }
  }
  setValue(t, n) {
    const r = this.getCellModel(n, !0);
    if (r?.get("value") === t)
      return;
    const s = ms(t ?? "");
    r?.set("value", s);
  }
  setCellFormula(t, n) {
    const r = this.getCellModel(n, !0);
    r?.get("formula") !== t && r?.set("formula", t);
  }
  validateSheetData(t) {
    const n = [], r = this.model.getCurrentSheetId();
    for (const o of t) {
      if (On(o) || o.row < 0 || o.col < 0)
        continue;
      const s = {
        ...o,
        sheetId: o.sheetId || r,
        value: ms(o.value ?? "")
      };
      o.formula && (o.formula.startsWith(Fn) || (s.formula = Fn + o.formula)), n.push(s);
    }
    return n;
  }
  getCoordinateList(t = !1) {
    const n = this.model.getCurrentSheetId(), r = this.getWorkData(n);
    return r.length === 0 ? [] : (r.sort((o, s) => t ? o.col - s.col : o.row - s.row), r);
  }
  getCellModel(t, n = !1) {
    const { row: r, col: o, sheetId: s } = t, l = s || this.model.getCurrentSheetId();
    if (!l)
      return;
    const f = nt(l, r, o), d = this.worksheets?.get(f);
    return n && !d ? this.setCellModel({ sheetId: l, row: r, col: o }, { value: "" }) : d;
  }
  setCellModel(t, n) {
    const { row: r, col: o, sheetId: s } = t, l = nt(
      s || this.model.getCurrentSheetId(),
      r,
      o
    );
    let f = this.worksheets;
    this.worksheets || (f = new we(), this.model.getRoot().set("worksheets", f));
    const d = new we(Object.entries(n));
    return f.set(l, d), d;
  }
  computeFormulasCallback = (t) => {
    const { list: n } = t;
    let r = !1;
    for (const o of n) {
      const s = this.worksheets?.get(o.key);
      if (!s)
        continue;
      const l = s.get("value"), f = s.get("formula");
      f && o.newValue === pa && f !== l ? (r = !0, s.set("value", f), s.set("formula", "")) : l !== o.newValue && (r = !0, s.set("value", o.newValue));
    }
    return r;
  };
  async computeFormulas() {
    const t = this.worksheets?.toJSON();
    if (!t || On(t))
      return Promise.resolve(!1);
    const n = {};
    for (const o of this.model.getDefineNameList())
      n[o.name] = o.range;
    const r = {
      worksheets: t,
      definedNames: n,
      currentSheetId: this.model.getCurrentSheetId(),
      workbook: this.model.getSheetList()
    };
    return this.worker.computeFormulas(
      r,
      ts(this.computeFormulasCallback)
    );
  }
}
class z6 {
  model;
  constructor(t) {
    this.model = t;
  }
  get mergeCells() {
    return this.model.getRoot().get("mergeCells");
  }
  fromJSON(t) {
    const n = t.mergeCells || {}, r = new we();
    for (const o of Object.values(n)) {
      if (o.sheetId = o.sheetId || this.model.getCurrentSheetId(), !this.model.validateRange(o))
        continue;
      const s = ut(
        o,
        "absolute",
        this.convertSheetIdToName
      );
      r.set(s, un(o));
    }
    this.model.getRoot().set("mergeCells", r);
  }
  getMergeCellList(t) {
    if (!this.mergeCells)
      return [];
    const n = t || this.model.getCurrentSheetId();
    return Array.from(this.mergeCells.values()).filter((r) => r.sheetId === n);
  }
  addMergeCell(t) {
    const n = ut(
      t,
      "absolute",
      this.convertSheetIdToName
    );
    if (this.mergeCells?.get(n)) {
      this.model.emit("toastMessage", {
        type: "error",
        message: k.t("merging-cell-is-duplicate")
      });
      return;
    }
    this.mergeCells || this.model.getRoot().set("mergeCells", new we()), this.mergeCells.set(n, un(t));
  }
  deleteMergeCell(t) {
    t.sheetId = t.sheetId || this.model.getCurrentSheetId();
    const n = ut(
      t,
      "absolute",
      this.convertSheetIdToName
    );
    this.mergeCells?.delete(n);
  }
  deleteAll(t) {
    if (!this.mergeCells)
      return;
    const n = t || this.model.getCurrentSheetId();
    for (const [r, o] of this.mergeCells.entries())
      o.sheetId === n && this.mergeCells.delete(r);
  }
  convertSheetIdToName = (t) => this.model.getSheetInfo(t)?.name || "";
}
class H6 {
  model;
  constructor(t) {
    this.model = t;
  }
  get customHeight() {
    return this.model.getRoot().get("customHeight");
  }
  getModel() {
    return this.customHeight || this.model.getRoot().set("customHeight", new we()), this.customHeight;
  }
  fromJSON(t) {
    const n = t.customHeight || {}, r = new we();
    for (const [o, s] of Object.entries(n)) {
      const { sheetId: l, rowOrCol: f } = ir(o);
      if (!l || f < 0)
        continue;
      const d = this.model.getSheetInfo(l);
      !d || f >= d.rowCount || typeof s.isHide == "boolean" && typeof s.len == "number" && r.set(o, s);
    }
    this.model.getRoot().set("customHeight", r);
  }
  hideRow(t, n) {
    this.toggleHideRow(t, n, !0);
  }
  toggleHideRow(t, n, r) {
    const o = this.model.getCurrentSheetId();
    for (let s = 0; s < n; s++) {
      const l = t + s, f = bt(o, l), d = this.getRow(l);
      if (d.isHide === r)
        continue;
      const c = { ...d, isHide: r };
      this.getModel().set(f, c);
    }
  }
  unhideRow(t, n) {
    this.toggleHideRow(t, n, !1);
  }
  getRow(t, n) {
    const r = n || this.model.getCurrentSheetId(), o = bt(r, t), s = this.customHeight?.get(o);
    return s ? { ...s } : {
      len: nr,
      isHide: !1
    };
  }
  /* jscpd:ignore-start */
  setRowHeight(t, n, r) {
    const o = r || this.model.getCurrentSheetId(), s = bt(o, t), l = this.getRow(t, r);
    if (l.len === n)
      return;
    const f = { ...l };
    f.len = n, this.getModel().set(s, f);
  }
  deleteAll(t) {
    if (!this.customHeight)
      return;
    const n = t || this.model.getCurrentSheetId();
    for (const r of this.customHeight.keys()) {
      const { sheetId: o } = ir(r);
      o === n && this.customHeight.delete(r);
    }
  }
  /* jscpd:ignore-end */
}
class V6 {
  model;
  constructor(t) {
    this.model = t;
  }
  get customWidth() {
    return this.model.getRoot().get("customWidth");
  }
  getModel() {
    return this.customWidth || this.model.getRoot().set("customWidth", new we()), this.customWidth;
  }
  fromJSON(t) {
    const n = t.customWidth || {}, r = new we();
    for (const [o, s] of Object.entries(n)) {
      const { sheetId: l, rowOrCol: f } = ir(o);
      if (!l || f < 0)
        continue;
      const d = this.model.getSheetInfo(l);
      !d || f >= d.colCount || typeof s.isHide == "boolean" && typeof s.len == "number" && r.set(o, s);
    }
    this.model.getRoot().set("customWidth", r);
  }
  hideCol(t, n) {
    this.toggleHideCol(t, n, !0);
  }
  toggleHideCol(t, n, r) {
    const o = this.model.getCurrentSheetId();
    for (let s = 0; s < n; s++) {
      const l = t + s, f = bt(o, l), d = this.getCol(l, o);
      if (d.isHide === r)
        continue;
      const c = { ...d, isHide: r };
      this.getModel().set(f, c);
    }
  }
  unhideCol(t, n) {
    this.toggleHideCol(t, n, !1);
  }
  getCol(t, n) {
    const r = n || this.model.getCurrentSheetId(), o = bt(r, t), s = this.customWidth?.get(o);
    return s ? { ...s } : {
      len: rr,
      isHide: !1
    };
  }
  setColWidth(t, n, r) {
    const o = r || this.model.getCurrentSheetId(), s = bt(o, t), l = this.getCol(t, r);
    if (l.len === n)
      return;
    const f = { ...l, len: n };
    this.getModel().set(s, f);
  }
  deleteAll(t) {
    if (!this.customWidth)
      return;
    const n = t || this.model.getCurrentSheetId();
    for (const r of this.customWidth.keys()) {
      const { sheetId: o } = ir(r);
      o === n && this.customWidth.delete(r);
    }
  }
}
class G6 {
  model;
  constructor(t) {
    this.model = t;
  }
  get autoFilter() {
    return this.model.getRoot().get("autoFilter");
  }
  fromJSON(t) {
    const n = t.autoFilter || {}, r = new we();
    for (const [o, s] of Object.entries(n)) {
      if (s.range.sheetId = o, !this.model.validateRange(s.range))
        continue;
      const l = new we(Object.entries(s));
      r.set(o, l);
    }
    this.model.getRoot().set("autoFilter", r);
  }
  deleteAll(t) {
    const n = t || this.model.getCurrentSheetId();
    this.autoFilter?.delete(n);
  }
  addFilter(t) {
    if (t.sheetId = t.sheetId || this.model.getCurrentSheetId(), !this.model.validateRange(t) || t.rowCount === 1 && t.colCount === 1)
      return;
    this.autoFilter || this.model.getRoot().set("autoFilter", new we());
    const n = {
      range: un(t)
    }, r = new we(Object.entries(n));
    this.autoFilter.set(t.sheetId, r);
  }
  deleteFilter(t) {
    this.deleteAll(t);
  }
  updateFilter(t, n) {
    const r = t || this.model.getCurrentSheetId(), o = Object.keys(n), s = this.autoFilter?.get(r);
    if (!s)
      return;
    for (const f of o)
      s.get(f) !== n[f] && s.set(f, n[f]);
    const l = s.get("range");
    if (l && typeof n.col == "number" && n.value) {
      this.applyFilter(l, n);
      return;
    }
    l && typeof n.col > "u" && !n.value && this.clearFilter(l);
  }
  clearFilter(t) {
    const n = this.model.getSheetInfo(t.sheetId);
    if (!n)
      return;
    let r = t.row + 1, o = t.rowCount === 0 ? n.rowCount : t.row + t.rowCount;
    for (; r < o; r++)
      this.model.unhideRow(r, 1);
  }
  applyFilter(t, n) {
    const r = this.model.getSheetInfo(t.sheetId);
    if (!r)
      return;
    const { col: o, value: s } = n;
    if (s?.type === "normal") {
      const l = new Set(s.value);
      let f = t.row + 1, d = t.rowCount === 0 ? r.rowCount : t.row + t.rowCount;
      for (; f < d; f++) {
        const c = this.model.getCell({
          row: f,
          col: o,
          rowCount: 1,
          colCount: 1,
          sheetId: t.sheetId
        }), i = c ? c.value : mt;
        l.has(i) ? this.model.unhideRow(f, 1) : this.model.hideRow(f, 1);
      }
    }
  }
  getFilter(t) {
    const n = t || this.model.getCurrentSheetId(), r = this.autoFilter?.get(n);
    if (r)
      return r.toJSON();
  }
}
class $6 {
  model;
  scroll = {};
  constructor(t) {
    this.model = t;
  }
  fromJSON(t) {
    const n = t.scroll || {}, r = {};
    for (const [o, s] of Object.entries(n))
      !o || s.row < 0 || s.col < 0 || (r[o] = s);
    this.scroll = r;
  }
  toJSON() {
    return { ...this.scroll };
  }
  deleteAll(t) {
    const n = t || this.model.getCurrentSheetId();
    delete this.scroll[n];
  }
  getScroll(t) {
    const n = this.scroll[t || this.model.getCurrentSheetId()];
    return n ? { ...n } : { row: 0, col: 0 };
  }
  setScroll(t, n) {
    const r = n || this.model.getCurrentSheetId(), o = this.model.getSheetInfo(r);
    if (!o)
      return !1;
    const s = o.rowCount - 1, l = o.colCount - 1;
    return t.row >= s ? t.row = s : t.row < 0 && (t.row = 0), t.col >= l ? t.col = l : t.col < 0 && (t.col = 0), this.scroll[r] = { ...t }, !0;
  }
}
class q6 extends ka {
  workbookManager;
  rangeMapManager;
  drawingsManager;
  definedNameManager;
  worksheetManager;
  mergeCellManager;
  rowManager;
  colManager;
  filterManager;
  scrollManager;
  doc;
  undoManager;
  constructor(t) {
    super();
    const { doc: n, worker: r } = t;
    this.doc = n;
    const o = this.getRoot();
    o.observeDeep((s) => {
      this.emit("modelChange", { event: s });
    }), this.undoManager = new eu(o, {
      trackedOrigins: /* @__PURE__ */ new Set([sn.MODEL, sn.SKIP_UPDATE]),
      captureTimeout: 100
    }), this.workbookManager = new F6(this), this.rangeMapManager = new P6(this), this.drawingsManager = new U6(this), this.definedNameManager = new _6(this), this.worksheetManager = new B6(this, r), this.mergeCellManager = new z6(this), this.rowManager = new H6(this), this.colManager = new V6(this), this.filterManager = new G6(this), this.scrollManager = new $6(this);
  }
  transaction = (t, n) => this.doc.transact(t, n || sn.MODEL);
  clearHistory() {
    this.undoManager.clear(!0, !0);
  }
  getRoot() {
    return this.doc.getMap("excel");
  }
  computeFormulas() {
    return this.worksheetManager.computeFormulas();
  }
  getSheetList() {
    return this.workbookManager.getSheetList();
  }
  getActiveRange(t) {
    const n = t || this.rangeMapManager.getActiveRange().range, r = this.getMergeCellList(n.sheetId);
    for (const o of r)
      if (Ma(n, o))
        return {
          range: {
            ...o,
            sheetId: o.sheetId || this.getCurrentSheetId()
          },
          isMerged: !0
        };
    return {
      range: n,
      isMerged: !1
    };
  }
  setActiveRange(t) {
    t.sheetId = t.sheetId || this.getCurrentSheetId();
    const { range: n } = this.getActiveRange(t);
    this.rangeMapManager.setActiveRange(n);
  }
  addSheet() {
    const t = this.workbookManager.addSheet();
    return this.worksheetManager.setWorksheet([]), this.workbookManager.setCurrentSheetId(t.sheetId), t;
  }
  deleteSheet(t) {
    const n = this.getNextSheetId(t);
    n && (this.deleteAll(t), this.workbookManager.deleteSheet(t), this.workbookManager.setCurrentSheetId(n));
  }
  updateSheetInfo(t, n) {
    this.workbookManager.updateSheetInfo(t, n);
  }
  hideSheet(t) {
    const n = this.getNextSheetId(t);
    n && (this.workbookManager.hideSheet(t), this.workbookManager.setCurrentSheetId(n));
  }
  unhideSheet(t) {
    t && (this.workbookManager.unhideSheet(t), this.workbookManager.setCurrentSheetId(t));
  }
  renameSheet(t, n) {
    this.workbookManager.renameSheet(t, n);
  }
  getSheetInfo(t) {
    return this.workbookManager.getSheetInfo(t);
  }
  setCurrentSheetId(t) {
    this.workbookManager.setCurrentSheetId(t);
  }
  getCurrentSheetId() {
    return this.workbookManager.getCurrentSheetId();
  }
  fromJSON = (t) => {
    this.workbookManager.fromJSON(t), this.rangeMapManager.fromJSON(t), this.drawingsManager.fromJSON(t), this.definedNameManager.fromJSON(t), this.worksheetManager.fromJSON(t), this.mergeCellManager.fromJSON(t), this.rowManager.fromJSON(t), this.colManager.fromJSON(t), this.filterManager.fromJSON(t), this.scrollManager.fromJSON(t), this.worksheetManager.computeFormulas();
  };
  toJSON = () => {
    const t = this.getRoot().toJSON(), n = {};
    for (const r of bi) {
      const o = t[r];
      if (r === "currentSheetId") {
        n.currentSheetId = o || "";
        continue;
      }
      n[r] = typeof o > "u" ? {} : o;
    }
    return {
      ...n,
      rangeMap: this.rangeMapManager.toJSON(),
      scroll: this.scrollManager.toJSON()
    };
  };
  deleteCell(t) {
    this.worksheetManager.deleteCell(t);
  }
  setCell(t, n, r) {
    return this.worksheetManager.setCell(t, n, r);
  }
  setCellValue(t, n) {
    return this.worksheetManager.setCellValue(t, n);
  }
  updateCellStyle(t, n) {
    return this.worksheetManager.updateCellStyle(t, n);
  }
  getCell = (t) => this.worksheetManager.getCell(t);
  getWorksheet(t) {
    return this.worksheetManager.getWorksheet(t);
  }
  setWorksheet(t) {
    this.worksheetManager.setWorksheet(t);
  }
  addCol(t, n, r = !1) {
    if (n <= 0)
      return;
    const o = this.getSheetInfo();
    if (!o || o.colCount >= hr)
      return;
    const s = this.getCurrentSheetId(), l = o.colCount + n;
    this.workbookManager.updateSheetInfo({ colCount: l }, s), this.drawingsManager.addCol(t, n, r), this.worksheetManager.addCol(t, n, r);
  }
  deleteCol(t, n) {
    if (n <= 0)
      return;
    const r = this.getSheetInfo();
    if (!r)
      return;
    const o = this.getCurrentSheetId(), s = r.colCount - n;
    this.workbookManager.updateSheetInfo({ colCount: s }, o), this.drawingsManager.deleteCol(t, n), this.worksheetManager.deleteCol(t, n);
  }
  hideCol(t, n) {
    this.colManager.hideCol(t, n);
  }
  getCol(t, n) {
    return this.colManager.getCol(t, n);
  }
  setColWidth(t, n, r) {
    this.colManager.setColWidth(t, n, r);
  }
  addRow(t, n, r = !1) {
    if (n <= 0)
      return;
    const o = this.getSheetInfo();
    if (!o || o.rowCount >= mr)
      return;
    const s = o.rowCount + n;
    this.workbookManager.updateSheetInfo({ rowCount: s }), this.drawingsManager.addRow(t, n, r), this.worksheetManager.addRow(t, n, r);
  }
  deleteRow(t, n) {
    if (n <= 0)
      return;
    const r = this.getSheetInfo();
    if (!r)
      return;
    const o = r.rowCount - n;
    this.workbookManager.updateSheetInfo({ rowCount: o }), this.drawingsManager.deleteRow(t, n), this.worksheetManager.deleteRow(t, n);
  }
  hideRow(t, n) {
    this.rowManager.hideRow(t, n);
  }
  unhideRow(t, n) {
    this.rowManager.unhideRow(t, n);
  }
  unhideCol(t, n) {
    this.colManager.unhideCol(t, n);
  }
  getRow(t, n) {
    return this.rowManager.getRow(t, n);
  }
  setRowHeight(t, n, r) {
    this.rowManager.setRowHeight(t, n, r);
  }
  canRedo() {
    return this.undoManager.canRedo();
  }
  canUndo() {
    return this.undoManager.canUndo();
  }
  undo() {
    this.undoManager.undo();
  }
  redo() {
    this.undoManager.redo();
  }
  pasteRange(t, n) {
    return this.worksheetManager.pasteRange(t, n);
  }
  deleteAll(t) {
    this.rowManager.deleteAll(t), this.colManager.deleteAll(t), this.worksheetManager.deleteAll(t), this.mergeCellManager.deleteAll(t), this.drawingsManager.deleteAll(t), this.definedNameManager.deleteAll(t), this.filterManager.deleteAll(t);
  }
  getDefineNameList() {
    return this.definedNameManager.getDefineNameList();
  }
  getDefineName(t) {
    return this.definedNameManager.getDefineName(t);
  }
  setDefineName(t, n) {
    return this.definedNameManager.setDefineName(t, n);
  }
  checkDefineName(t) {
    return this.definedNameManager.checkDefineName(t);
  }
  getDrawingList(t) {
    return this.drawingsManager.getDrawingList(t);
  }
  addDrawing(...t) {
    for (const n of t)
      this.drawingsManager.addDrawing(n);
  }
  updateDrawing(t, n) {
    this.drawingsManager.updateDrawing(t, n);
  }
  deleteDrawing(t) {
    this.drawingsManager.deleteDrawing(t);
  }
  getMergeCellList(t) {
    return this.mergeCellManager.getMergeCellList(t);
  }
  addMergeCell(t, n = Ct.MERGE_CENTER) {
    t.colCount === 1 && t.rowCount === 1 || (t.sheetId = t.sheetId || this.getCurrentSheetId(), this.mergeCellManager.addMergeCell(t), this.worksheetManager.addMergeCell(t, n));
  }
  deleteMergeCell(t) {
    this.mergeCellManager.deleteMergeCell(t);
  }
  validateRange(t) {
    return this.rangeMapManager.validateRange(t);
  }
  validateDrawing(t) {
    return this.drawingsManager.validateDrawing(t);
  }
  validateDefinedName(t) {
    return this.definedNameManager.validateDefinedName(t);
  }
  getFilter(t) {
    return this.filterManager.getFilter(t);
  }
  addFilter(t) {
    this.filterManager.addFilter(t);
  }
  deleteFilter(t) {
    this.filterManager.deleteFilter(t);
  }
  updateFilter(t, n) {
    this.filterManager.updateFilter(t, n);
  }
  getScroll(t) {
    return this.scrollManager.getScroll(t);
  }
  setScroll(t, n) {
    return this.scrollManager.setScroll(t, n);
  }
  getNextSheetId(t) {
    const n = t || this.getCurrentSheetId(), r = this.getSheetList(), o = r.findIndex((f) => f.sheetId === n);
    if (o < 0)
      return "";
    const s = o === r.length - 1;
    let l = s ? (o - 1 + r.length) % r.length : (o + 1) % r.length;
    for (; l !== o && r[l].isHide; )
      l = s ? (l - 1 + r.length) % r.length : (l + 1) % r.length;
    return r[l].sheetId;
  }
}
function W6(e) {
  return typeof DOMException < "u" && e instanceof DOMException && e.name === "DataCloneError";
}
function Y6(e) {
  if (e === null) return null;
  const t = typeof e;
  if (t === "function") {
    const n = e;
    return `function ${n.name || "(anonymous)"} ${typeof n.toString == "function" ? n.toString().slice(0, 80) : ""}`;
  }
  return t === "symbol" ? "symbol" : null;
}
function Rn(e, t, n, r) {
  if (r <= 0) return null;
  const o = Y6(e);
  if (o !== null) return { path: t, description: o };
  if (e === null || typeof e != "object") return null;
  const s = e;
  if (n.has(s) || (n.add(s), s instanceof ArrayBuffer || ArrayBuffer.isView(s) || s instanceof Date || s instanceof RegExp || s instanceof Blob || s instanceof File))
    return null;
  if (s instanceof Map) {
    let l = 0;
    for (const [f, d] of s) {
      const c = Rn(f, `${t}.key[${l}]`, n, r - 1) || Rn(d, `${t}.val[${l}]`, n, r - 1);
      if (c) return c;
      l++;
    }
    return null;
  }
  if (s instanceof Set) {
    let l = 0;
    for (const f of s) {
      const d = Rn(
        f,
        `${t}[${l}]`,
        n,
        r - 1
      );
      if (d) return d;
      l++;
    }
    return null;
  }
  if (Array.isArray(s)) {
    for (let l = 0; l < s.length; l++) {
      const f = Rn(
        s[l],
        `${t}[${l}]`,
        n,
        r - 1
      );
      if (f) return f;
    }
    return null;
  }
  for (const l of Object.keys(s)) {
    const f = Rn(
      s[l],
      `${t}.${l}`,
      n,
      r - 1
    );
    if (f) return f;
  }
  return null;
}
function X6(e) {
  const t = e.postMessage.bind(e);
  let n = !1;
  return e.postMessage = ((...r) => {
    try {
      t(...r);
    } catch (o) {
      if (!W6(o)) throw o;
      if (!n) {
        n = !0;
        const [s] = r, l = Rn(
          s,
          "payload",
          /* @__PURE__ */ new WeakSet(),
          32
        );
        console.error(
          "[excel-collab] worker.postMessage: DataCloneError. The render worker received a non-cloneable payload and will not render. First non-cloneable value found at:",
          l ? l.path : "(not found within traversal depth)",
          l ? `type: ${l.description}` : "",
          `
Raw payload:`,
          r[0]
        );
      }
      throw o;
    }
  }), e;
}
function K6() {
  return {
    worker: D6,
    doc: new $o()
  };
}
function J6(e = K6()) {
  const n = {
    worker: e.worker instanceof Worker ? Ll(X6(e.worker)) : e.worker,
    doc: e.doc
  }, r = new q6(n);
  return new re(r, n);
}
const Xc = '(function(){"use strict";class Te{type;value;constructor(t,n){this.type=t,this.value=n}toString(){return this.value}}var St=(e=>(e[e.TOP=0]="TOP",e[e.MIDDLE=1]="MIDDLE",e[e.BOTTOM=2]="BOTTOM",e))(St||{}),Pt=(e=>(e[e.LEFT=0]="LEFT",e[e.CENTER=1]="CENTER",e[e.RIGHT=2]="RIGHT",e))(Pt||{}),Rr=(e=>(e[e.NONE=0]="NONE",e[e.SINGLE=1]="SINGLE",e[e.DOUBLE=2]="DOUBLE",e))(Rr||{}),S=(e=>(e[e.EQUAL=0]="EQUAL",e[e.NOT_EQUAL=1]="NOT_EQUAL",e[e.PLUS=2]="PLUS",e[e.MINUS=3]="MINUS",e[e.STAR=4]="STAR",e[e.SLASH=5]="SLASH",e[e.EXPONENT=6]="EXPONENT",e[e.GREATER=7]="GREATER",e[e.GREATER_EQUAL=8]="GREATER_EQUAL",e[e.CONCATENATE=9]="CONCATENATE",e[e.COLON=10]="COLON",e[e.COMMA=11]="COMMA",e[e.EMPTY_CHAR=12]="EMPTY_CHAR",e[e.PERCENT=13]="PERCENT",e[e.LESS=14]="LESS",e[e.LESS_EQUAL=15]="LESS_EQUAL",e[e.STRING=16]="STRING",e[e.NUMBER=17]="NUMBER",e[e.BOOL=18]="BOOL",e[e.ERROR=19]="ERROR",e[e.ERROR_REF=20]="ERROR_REF",e[e.LEFT_BRACKET=21]="LEFT_BRACKET",e[e.RIGHT_BRACKET=22]="RIGHT_BRACKET",e[e.lEFT_BRACE=23]="lEFT_BRACE",e[e.RIGHT_BRACE=24]="RIGHT_BRACE",e[e.SEMICOLON=25]="SEMICOLON",e[e.EXCLAMATION=26]="EXCLAMATION",e[e.R1C1=27]="R1C1",e[e.SHEET_NAME=28]="SHEET_NAME",e[e.CELL=29]="CELL",e[e.COLUMN=30]="COLUMN",e[e.ROW=31]="ROW",e[e.DEFINED_NAME=32]="DEFINED_NAME",e[e.EXCEL_FUNCTION=33]="EXCEL_FUNCTION",e[e.REF_FUNCTION=34]="REF_FUNCTION",e[e.REF_FUNCTION_COND=35]="REF_FUNCTION_COND",e[e.EOF=36]="EOF",e))(S||{});const Ut=12,Ho="sans-serif",Mr=22,Or=76,Dr=0,Nn="&#10;",wr=16384,Lr=1048576,Vo=16,Sr="=",yt="_",An="General",ko="#GETTING_DATA",Pr=new Set(["#DIV/0!","#NULL!","#NUM!","#REF!","#VALUE!","#N/A","#NAME?",ko]),q=1,Ur={thin:q,hair:q,dotted:q,dashed:q,dashDot:q,dashDotDot:q,double:q,medium:q*2,mediumDashed:q*2,mediumDashDot:q*2,mediumDashDotDot:q*2,thick:q*4},Bo=/^[$]?[A-Za-z]{1,3}[$]?[1-9][0-9]*$/,Go=/^[$]?[A-Za-z]{1,3}$/,zo=/^\\$[1-9][0-9]*$/,Yo=/^[a-zA-Z_][a-zA-Z0-9_.?]*$/,Wo=new Set(["ABS","ACCRINT","ACCRINTM","ACOS","ACOSH","ACOT","ACOTH","ADDRESS","AGGREGATE","AMORDEGRC","AMORLINC","AND","ARABIC","AREAS","ARRAYTOTEXT","ASC","ASIN","ASINH","ATAN","ATAN2","ATANH","AVEDEV","AVERAGE","AVERAGEA","AVERAGEIF","AVERAGEIFS","BAHTTEXT","BASE","BESSELI","BESSELJ","BESSELK","BESSELY","BETA.DIST","BETA.INV","BETADIST","BETAINV","BIN2DEC","BIN2HEX","BIN2OCT","BINOM.DIST","BINOM.DIST.RANGE","BINOM.INV","BINOMDIST","BITAND","BITLSHIFT","BITOR","BITRSHIFT","BITXOR","CALL","CEILING","CEILING.MATH","CEILING.PRECISE","CELL","CHAR","CHIDIST","CHIINV","CHISQ.DIST","CHISQ.DIST.RT","CHISQ.INV","CHISQ.INV.RT","CHISQ.TEST","CHITEST","CLEAN","CODE","COLUMN","COLUMNS","COMBIN","COMBINA","COMPLEX","CONCAT","CONCATENATE","CONFIDENCE","CONFIDENCE.NORM","CONFIDENCE.T","CONVERT","CORREL","COS","COSH","COT","COTH","COUNT","COUNTA","COUNTBLANK","COUNTIF","COUNTIFS","COUPDAYBS","COUPDAYS","COUPDAYSNC","COUPNCD","COUPNUM","COUPPCD","COVAR","COVARIANCE.P","COVARIANCE.S","CRITBINOM","CSC","CSCH","CUBEKPIMEMBER","CUBEMEMBER","CUBEMEMBERPROPERTY","CUBERANKEDMEMBER","CUBESET","CUBESETCOUNT","CUBEVALUE","CUMIPMT","CUMPRINC","DATE","DATEDIF","DATEVALUE","DAVERAGE","DAY","DAYS","DAYS360","DB","DBCS","DCOUNT","DCOUNTA","DDB","DEC2BIN","DEC2HEX","DEC2OCT","DECIMAL","DEGREES","DELTA","DEVSQ","DGET","DISC","DMAX","DMIN","DOLLAR","DOLLARDE","DOLLARFR","DPRODUCT","DSTDEV","DSTDEVP","DSUM","DURATION","DVAR","DVARP","EDATE","EFFECT","ENCODEURL","EOMONTH","ERF","ERF.PRECISE","ERFC","ERFC.PRECISE","ERROR.TYPE","EUROCONVERT","EVEN","EXACT","EXP","EXPON.DIST","EXPONDIST","F.DIST","F.DIST.RT","F.INV","F.INV.RT","F.TEST","FACT","FACTDOUBLE","FALSE","FDIST","FILTER","FILTERXML","FIND","FINDB","FINV","FISHER","FISHERINV","FIXED","FLOOR","FLOOR.MATH","FLOOR.PRECISE","FORECAST","FORECAST.ETS","FORECAST.ETS.CONFINT","FORECAST.ETS.SEASONALITY","FORECAST.ETS.STAT","FORECAST.LINEAR","FORMULATEXT","FREQUENCY","FTEST","FV","FVSCHEDULE","GAMMA","GAMMA.DIST","GAMMA.INV","GAMMADIST","GAMMAINV","GAMMALN","GAMMALN.PRECISE","GAUSS","GCD","GEOMEAN","GESTEP","GETPIVOTDATA","GROWTH","HARMEAN","HEX2BIN","HEX2DEC","HEX2OCT","HLOOKUP","HOUR","HYPERLINK","HYPGEOM.DIST","HYPGEOMDIST","IFERROR","IFNA","IFS","IMABS","IMAGINARY","IMARGUMENT","IMCONJUGATE","IMCOS","IMCOSH","IMCOT","IMCSC","IMCSCH","IMDIV","IMEXP","IMLN","IMLOG10","IMLOG2","IMPOWER","IMPRODUCT","IMREAL","IMSEC","IMSECH","IMSIN","IMSINH","IMSQRT","IMSUB","IMSUM","IMTAN","INFO","INT","INTERCEPT","INTRATE","IPMT","IRR","ISBLANK","ISERR","ISERROR","ISEVEN","ISFORMULA","ISLOGICAL","ISNA","ISNONTEXT","ISNUMBER","ISO.CEILING","ISODD","ISOWEEKNUM","ISPMT","ISREF","ISTEXT","JIS","KURT","LARGE","LCM","LEFT","LEFTB","LEN","LENB","LET","LINEST","LN","LOG","LOG10","LOGEST","LOGINV","LOGNORM.DIST","LOGNORM.INV","LOGNORMDIST","LOOKUP","LOWER","MATCH","MAX","MAXA","MAXIFS","MDETERM","MDURATION","MEDIAN","MID","MIDB","MIN","MINA","MINIFS","MINUTE","MINVERSE","MIRR","MMULT","MOD","MODE","MODE.MULT","MODE.SNGL","MONTH","MROUND","MULTINOMIAL","MUNIT","N","NA","NEGBINOM.DIST","NEGBINOMDIST","NETWORKDAYS","NETWORKDAYS.INTL","NOMINAL","NORM.DIST","NORM.INV","NORM.S.DIST","NORM.S.INV","NORMDIST","NORMINV","NORMSDIST","NORMSINV","NOT","NOW","NPER","NPV","NUMBERVALUE","OCT2BIN","OCT2DEC","OCT2HEX","ODD","ODDFPRICE","ODDFYIELD","ODDLPRICE","ODDLYIELD","OR","PDURATION","PEARSON","PERCENTILE","PERCENTILE.EXC","PERCENTILE.INC","PERCENTRANK","PERCENTRANK.EXC","PERCENTRANK.INC","PERMUT","PERMUTATIONA","PHI","PHONETIC","PI","PMT","POISSON","POISSON.DIST","POWER","PPMT","PRICE","PRICEDISC","PRICEMAT","PROB","PRODUCT","PROPER","PV","QUARTILE","QUARTILE.EXC","QUARTILE.INC","QUOTIENT","RADIANS","RAND","RANDARRAY","RANDBETWEEN","RANK","RANK.AVG","RANK.EQ","RATE","RECEIVED","REGISTER.ID","REPLACE","REPLACEB","REPT","RIGHT","RIGHTB","ROMAN","ROUND","ROUNDDOWN","ROUNDUP","ROW","ROWS","RRI","RSQ","RTD","SEARCH","SEARCHB","SEC","SECH","SECOND","SEQUENCE","SERIESSUM","SHEET","SHEETS","SIGN","SIN","SINH","SKEW","SKEW.P","SLN","SLOPE","SMALL","SORT","SORTBY","SQRT","SQRTPI","STANDARDIZE","STDEV","STDEV.P","STDEV.S","STDEVA","STDEVP","STDEVPA","STEYX","SUBSTITUTE","SUBTOTAL","SUM","SUMIF","SUMIFS","SUMPRODUCT","SUMSQ","SUMX2MY2","SUMX2PY2","SUMXMY2","SWITCH","SYD","T","T.DIST","T.DIST.2T","T.DIST.RT","T.INV","T.INV.2T","T.TEST","TAN","TANH","TBILLEQ","TBILLPRICE","TBILLYIELD","TDIST","TEXT","TEXTJOIN","TIME","TIMEVALUE","TINV","TODAY","TRANSPOSE","TREND","TRIM","TRIMMEAN","TRUE","TRUNC","TTEST","TYPE","UNICHAR","UNICODE","UNIQUE","UPPER","VALUE","VALUETOTEXT","VAR","VAR.P","VAR.S","VARA","VARP","VARPA","VDB","VLOOKUP","WEBSERVICE","WEEKDAY","WEEKNUM","WEIBULL","WEIBULL.DIST","WORKDAY","WORKDAY.INTL","XIRR","XLOOKUP","XMATCH","XNPV","XOR","YEAR","YEARFRAC","YIELD","YIELDDISC","YIELDMAT","Z.TEST","ZTEST"]);function qo(e){const t=e||{};return[Object,Array].includes(t.constructor)&&!Object.entries(t).length}function yr(e){if(typeof e=="boolean")return[!0,Number(e)];if(typeof e=="number"&&!isNaN(e))return[!0,e];if(typeof e!="string")return[!1,0];if(e.length>12)return[!1,0];const t=Number(e);return isNaN(t)?[!1,0]:[!0,t]}function Fr(e,t,n){return`${e}${yt}${t}${yt}${n}`}function Ko(e){const[t,n,r]=e.split(yt),o=parseInt(n,10),l=parseInt(r,10),c={row:isNaN(o)?-1:o,col:isNaN(l)?-1:l,sheetId:t};return c.row<0||c.col<0||!t?null:c}function Hr(e,t){return`${e}${yt}${t}`}function Xo(e){return!!(typeof e=="string"&&e&&e.startsWith(Sr)&&e.length>1)}function $o(e,t){return e&&t.includes(Nn)}function xo(){const e=Math.floor(Math.random()*256).toString(16).padStart(2,"0"),t=Math.floor(Math.random()*256).toString(16).padStart(2,"0"),n=Math.floor(Math.random()*256).toString(16).padStart(2,"0");return`#${e}${t}${n}`}function Vr(e="normal",t="normal",n=Ut,r=""){const o=`${e} ${t} ${n}px `,l=["Source Code Pro",r,Ho].filter(c=>c.trim());return o+l.join(",")}let Cn;const Qo=2,kr=e=>Math.max(Math.floor(e||1),Qo);function Ft(e=typeof devicePixelRatio>"u"?void 0:devicePixelRatio){return typeof Cn=="number"?Cn:kr(e)}function x(e){return Math.floor(e*Ft())}function Jo(e){Cn=kr(e)}function Zo(e){const t=e.toUpperCase();let n=0;for(let r=0;r<t.length;r++)n=t.charCodeAt(r)-64+n*26;return n-1}function bo(e){const t=e+1;let n="",r=Math.floor(Math.abs(t)),o;for(;r>0;)o=(r-1)%26,n=String.fromCharCode(65+o)+n,r=Math.floor((r-o)/26);return n.toUpperCase()}function _o(e){let t=parseInt(e,10);return isNaN(t)?t=-1:t=Math.max(t-1,-1),t}function dt(e){return Et(e)&&mt(e)}function Et(e){return e.colCount===0}function mt(e){return e.rowCount===0}function jo(e,t){if(e.colCount===0&&t.colCount!==0||e.rowCount===0&&t.rowCount!==0)return!1;if(t.colCount===0&&t.row===e.row||t.rowCount===0&&t.col===e.col)return!0;const{row:n,col:r}=e;return n>=t.row&&n<t.row+t.rowCount&&r>=t.col&&r<t.col+t.colCount}class Le{row=0;col=0;colCount=0;rowCount=0;sheetId="";constructor(t,n,r,o,l){this.row=t,this.col=n,this.colCount=o,this.rowCount=r,this.sheetId=l}static makeRange(t){return new Le(t.row,t.col,t.rowCount,t.colCount,t.sheetId)}toIRange(){return{row:this.row,col:this.col,rowCount:this.rowCount,colCount:this.colCount,sheetId:this.sheetId}}}function es(e,t,n,r){const{row:o,col:l,rowCount:c,colCount:h}=e;if(dt(e)){for(let i=0;i<t;i++)for(let s=0;s<n;s++)if(r(i,s))return;return}if(Et(e)){for(let i=0;i<n;i++)if(r(o,i))return;return}if(mt(e)){for(let i=0;i<t;i++)if(r(i,l))return;return}const g=o+c,u=l+h;for(let i=o;i<g;i++)for(let s=l;s<u;s++)if(r(i,s))return}const vn=e=>e>="a"&&e<="z"||e>="A"&&e<="Z",Be=e=>e>="0"&&e<="9";function ts(e){return e}function Br(e,t){if(!e)return;const n=e.trim();let[r,o=""]=n.split("!");n.includes("!")||(r="",o=n),r&&r[0]==="\'"&&r[r.length-1]==="\'"&&(r=r.slice(1,-1));let l=0,c="",h="";for(o[l]==="$"&&l++;l<o.length&&vn(o[l]);)h+=o[l++];for(o[l]==="$"&&l++;l<o.length&&Be(o[l]);)c+=o[l++];if(l!==o.length||!c&&!h)return;let g=1,u=1,i=-1,s=-1;return c===""?(g=0,i=0):i=_o(c),h===""?(u=0,s=0):s=Zo(h),i<0||s<0||s>=wr||i>=Lr?void 0:new Le(i,s,g,u,t(r))}function ns(e,t=ts){const[n="",r=""]=e?.split(":")||[],o=Br(n,t);if(!o)return;const l=Br(r,t);return l?(l.sheetId=o.sheetId,Gr(o,l)):o}function Gr(e,t){if(e.sheetId!==t.sheetId)return;if(e.row===t.row&&e.col===t.col&&e.rowCount===t.rowCount&&e.colCount===t.colCount)return Le.makeRange(e);if(e.rowCount===0&&t.rowCount!==0||e.rowCount!==0&&t.rowCount===0||e.colCount===0&&t.colCount!==0||e.colCount!==0&&t.colCount===0)return;const n=Math.abs(e.row-t.row)+1,r=Math.abs(e.col-t.col)+1,o=e.row<t.row?e.row:t.row,l=e.col<t.col?e.col:t.col;return new Le(o,l,n,r,e.sheetId)}function zr(e,t){let n=-1;if(e.startsWith("[")){const r=parseInt(e.slice(1,-1),10);isNaN(r)||(n=t+r)}else{const r=parseInt(e,10);isNaN(r)||(n=r-1)}return isNaN(n)?-1:n}function rs(e,t={row:-1,col:-1,sheetId:""}){const n=e.toUpperCase();if(n[0]!=="R")return;const r=n.slice(1).split("C");if(r.length!==2)return;const[o,l]=r;let c=-1,h=-1;return o?c=zr(o,t.row):c=t.row,l?h=zr(l,t.col):h=t.col,h>=wr||c>=Lr||c<0||h<0?void 0:new Le(c,h,1,1,t.sheetId)}class Ge{namespace;static colorMap=new Map;constructor(t){this.namespace=t}init=()=>(this.setColor(),this.log);log=(...t)=>{if(!this.enable())return;const{namespace:n}=this,r=Ge.colorMap.get(n),o=[`%c ${n}:`,`color:${r};`,...t];console.log(...o)};enable(){return window.localStorage.getItem("debug")}setColor(){Ge.colorMap.has(this.namespace)||Ge.colorMap.set(this.namespace,xo())}}new Ge("model").init(),new Ge("controller").init(),new Ge("react").init(),new Ge("collaboration").init();function Yr(e){return e?typeof Intl>"u"||typeof Intl.Segmenter!="function"?[...e]:[...new Intl.Segmenter([],{granularity:"word"}).segment(e)].map(r=>r.segment):[]}function Wr(e){return e&&e.__esModule&&Object.prototype.hasOwnProperty.call(e,"default")?e.default:e}var Ht={exports:{}},is=Ht.exports,qr;function os(){return qr||(qr=1,(function(e,t){(function(n,r){e.exports=r()})(is,function(){var n=(function(r,o){var l=Array.prototype.concat,c=Array.prototype.slice,h=Object.prototype.toString;function g(v,p){var I=v>p?v:p;return r.pow(10,17-~~(r.log(I>0?I:-I)*r.LOG10E))}var u=Array.isArray||function(p){return h.call(p)==="[object Array]"};function i(v){return h.call(v)==="[object Function]"}function s(v){return typeof v=="number"?v-v===0:!1}function f(v){return l.apply([],v)}function a(){return new a._init(arguments)}a.fn=a.prototype,a._init=function(p){if(u(p[0]))if(u(p[0][0])){i(p[1])&&(p[0]=a.map(p[0],p[1]));for(var I=0;I<p[0].length;I++)this[I]=p[0][I];this.length=p[0].length}else this[0]=i(p[1])?a.map(p[0],p[1]):p[0],this.length=1;else if(s(p[0]))this[0]=a.seq.apply(null,p),this.length=1;else{if(p[0]instanceof a)return a(p[0].toArray());this[0]=[],this.length=1}return this},a._init.prototype=a.prototype,a._init.constructor=a,a.utils={calcRdx:g,isArray:u,isFunction:i,isNumber:s,toVector:f},a._random_fn=r.random,a.setRandom=function(p){if(typeof p!="function")throw new TypeError("fn is not a function");a._random_fn=p},a.extend=function(p){var I,A;if(arguments.length===1){for(A in p)a[A]=p[A];return this}for(I=1;I<arguments.length;I++)for(A in arguments[I])p[A]=arguments[I][A];return p},a.rows=function(p){return p.length||1},a.cols=function(p){return p[0].length||1},a.dimensions=function(p){return{rows:a.rows(p),cols:a.cols(p)}},a.row=function(p,I){return u(I)?I.map(function(A){return a.row(p,A)}):p[I]},a.rowa=function(p,I){return a.row(p,I)},a.col=function(p,I){if(u(I)){var A=a.arange(p.length).map(function(){return new Array(I.length)});return I.forEach(function(L,U){a.arange(p.length).forEach(function(H){A[H][U]=p[H][L]})}),A}for(var M=new Array(p.length),O=0;O<p.length;O++)M[O]=[p[O][I]];return M},a.cola=function(p,I){return a.col(p,I).map(function(A){return A[0]})},a.diag=function(p){for(var I=a.rows(p),A=new Array(I),M=0;M<I;M++)A[M]=[p[M][M]];return A},a.antidiag=function(p){for(var I=a.rows(p)-1,A=new Array(I),M=0;I>=0;I--,M++)A[M]=[p[M][I]];return A},a.transpose=function(p){var I=[],A,M,O,L,U;for(u(p[0])||(p=[p]),M=p.length,O=p[0].length,U=0;U<O;U++){for(A=new Array(M),L=0;L<M;L++)A[L]=p[L][U];I.push(A)}return I.length===1?I[0]:I},a.map=function(p,I,A){var M,O,L,U,H;for(u(p[0])||(p=[p]),O=p.length,L=p[0].length,U=A?p:new Array(O),M=0;M<O;M++)for(U[M]||(U[M]=new Array(L)),H=0;H<L;H++)U[M][H]=I(p[M][H],M,H);return U.length===1?U[0]:U},a.cumreduce=function(p,I,A){var M,O,L,U,H;for(u(p[0])||(p=[p]),O=p.length,L=p[0].length,U=A?p:new Array(O),M=0;M<O;M++)for(U[M]||(U[M]=new Array(L)),L>0&&(U[M][0]=p[M][0]),H=1;H<L;H++)U[M][H]=I(U[M][H-1],p[M][H]);return U.length===1?U[0]:U},a.alter=function(p,I){return a.map(p,I,!0)},a.create=function(p,I,A){var M=new Array(p),O,L;for(i(I)&&(A=I,I=p),O=0;O<p;O++)for(M[O]=new Array(I),L=0;L<I;L++)M[O][L]=A(O,L);return M};function d(){return 0}a.zeros=function(p,I){return s(I)||(I=p),a.create(p,I,d)};function E(){return 1}a.ones=function(p,I){return s(I)||(I=p),a.create(p,I,E)},a.rand=function(p,I){return s(I)||(I=p),a.create(p,I,a._random_fn)};function m(v,p){return v===p?1:0}a.identity=function(p,I){return s(I)||(I=p),a.create(p,I,m)},a.symmetric=function(p){var I=p.length,A,M;if(p.length!==p[0].length)return!1;for(A=0;A<I;A++)for(M=0;M<I;M++)if(p[M][A]!==p[A][M])return!1;return!0},a.clear=function(p){return a.alter(p,d)},a.seq=function(p,I,A,M){i(M)||(M=!1);var O=[],L=g(p,I),U=(I*L-p*L)/((A-1)*L),H=p,y;for(y=0;H<=I&&y<A;y++,H=(p*L+U*L*y)/L)O.push(M?M(H,y):H);return O},a.arange=function(p,I,A){var M=[],O;if(A=A||1,I===o&&(I=p,p=0),p===I||A===0)return[];if(p<I&&A<0)return[];if(p>I&&A>0)return[];if(A>0)for(O=p;O<I;O+=A)M.push(O);else for(O=p;O>I;O+=A)M.push(O);return M},a.slice=(function(){function v(I,A,M,O){var L,U=[],H=I.length;if(A===o&&M===o&&O===o)return a.copy(I);if(A=A||0,M=M||I.length,A=A>=0?A:H+A,M=M>=0?M:H+M,O=O||1,A===M||O===0)return[];if(A<M&&O<0)return[];if(A>M&&O>0)return[];if(O>0)for(L=A;L<M;L+=O)U.push(I[L]);else for(L=A;L>M;L+=O)U.push(I[L]);return U}function p(I,A){var M,O;if(A=A||{},s(A.row)){if(s(A.col))return I[A.row][A.col];var L=a.rowa(I,A.row);return M=A.col||{},v(L,M.start,M.end,M.step)}if(s(A.col)){var U=a.cola(I,A.col);return O=A.row||{},v(U,O.start,O.end,O.step)}O=A.row||{},M=A.col||{};var H=v(I,O.start,O.end,O.step);return H.map(function(y){return v(y,M.start,M.end,M.step)})}return p})(),a.sliceAssign=function(p,I,A){var M,O;if(s(I.row)){if(s(I.col))return p[I.row][I.col]=A;I.col=I.col||{},I.col.start=I.col.start||0,I.col.end=I.col.end||p[0].length,I.col.step=I.col.step||1,M=a.arange(I.col.start,r.min(p.length,I.col.end),I.col.step);var L=I.row;return M.forEach(function(H,y){p[L][H]=A[y]}),p}if(s(I.col)){I.row=I.row||{},I.row.start=I.row.start||0,I.row.end=I.row.end||p.length,I.row.step=I.row.step||1,O=a.arange(I.row.start,r.min(p[0].length,I.row.end),I.row.step);var U=I.col;return O.forEach(function(H,y){p[H][U]=A[y]}),p}return A[0].length===o&&(A=[A]),I.row.start=I.row.start||0,I.row.end=I.row.end||p.length,I.row.step=I.row.step||1,I.col.start=I.col.start||0,I.col.end=I.col.end||p[0].length,I.col.step=I.col.step||1,O=a.arange(I.row.start,r.min(p.length,I.row.end),I.row.step),M=a.arange(I.col.start,r.min(p[0].length,I.col.end),I.col.step),O.forEach(function(H,y){M.forEach(function(Z,Q){p[H][Z]=A[y][Q]})}),p},a.diagonal=function(p){var I=a.zeros(p.length,p.length);return p.forEach(function(A,M){I[M][M]=A}),I},a.copy=function(p){return p.map(function(I){return s(I)?I:I.map(function(A){return A})})};var C=a.prototype;return C.length=0,C.push=Array.prototype.push,C.sort=Array.prototype.sort,C.splice=Array.prototype.splice,C.slice=Array.prototype.slice,C.toArray=function(){return this.length>1?c.call(this):c.call(this)[0]},C.map=function(p,I){return a(a.map(this,p,I))},C.cumreduce=function(p,I){return a(a.cumreduce(this,p,I))},C.alter=function(p){return a.alter(this,p),this},(function(v){for(var p=0;p<v.length;p++)(function(I){C[I]=function(A){var M=this,O;return A?(setTimeout(function(){A.call(M,C[I].call(M))}),this):(O=a[I](this),u(O)?a(O):O)}})(v[p])})("transpose clear symmetric rows cols dimensions diag antidiag".split(" ")),(function(v){for(var p=0;p<v.length;p++)(function(I){C[I]=function(A,M){var O=this;return M?(setTimeout(function(){M.call(O,C[I].call(O,A))}),this):a(a[I](this,A))}})(v[p])})("row col".split(" ")),(function(v){for(var p=0;p<v.length;p++)(function(I){C[I]=function(){return a(a[I].apply(null,arguments))}})(v[p])})("create zeros ones rand identity".split(" ")),a})(Math);return(function(r,o){var l=r.utils.isFunction;function c(u,i){return u-i}function h(u,i,s){return o.max(i,o.min(u,s))}r.sum=function(i){for(var s=0,f=i.length;--f>=0;)s+=i[f];return s},r.sumsqrd=function(i){for(var s=0,f=i.length;--f>=0;)s+=i[f]*i[f];return s},r.sumsqerr=function(i){for(var s=r.mean(i),f=0,a=i.length,d;--a>=0;)d=i[a]-s,f+=d*d;return f},r.sumrow=function(i){for(var s=0,f=i.length;--f>=0;)s+=i[f];return s},r.product=function(i){for(var s=1,f=i.length;--f>=0;)s*=i[f];return s},r.min=function(i){for(var s=i[0],f=0;++f<i.length;)i[f]<s&&(s=i[f]);return s},r.max=function(i){for(var s=i[0],f=0;++f<i.length;)i[f]>s&&(s=i[f]);return s},r.unique=function(i){for(var s={},f=[],a=0;a<i.length;a++)s[i[a]]||(s[i[a]]=!0,f.push(i[a]));return f},r.mean=function(i){return r.sum(i)/i.length},r.meansqerr=function(i){return r.sumsqerr(i)/i.length},r.geomean=function(i){var s=i.map(o.log),f=r.mean(s);return o.exp(f)},r.median=function(i){var s=i.length,f=i.slice().sort(c);return s&1?f[s/2|0]:(f[s/2-1]+f[s/2])/2},r.cumsum=function(i){return r.cumreduce(i,function(s,f){return s+f})},r.cumprod=function(i){return r.cumreduce(i,function(s,f){return s*f})},r.diff=function(i){var s=[],f=i.length,a;for(a=1;a<f;a++)s.push(i[a]-i[a-1]);return s},r.rank=function(u){var i,s=[],f={};for(i=0;i<u.length;i++){var a=u[i];f[a]?f[a]++:(f[a]=1,s.push(a))}var d=s.sort(c),E={},m=1;for(i=0;i<d.length;i++){var a=d[i],C=f[a],v=m,p=m+C-1,I=(v+p)/2;E[a]=I,m+=C}return u.map(function(A){return E[A]})},r.mode=function(i){var s=i.length,f=i.slice().sort(c),a=1,d=0,E=0,m=[],C;for(C=0;C<s;C++)f[C]===f[C+1]?a++:(a>d?(m=[f[C]],d=a,E=0):a===d&&(m.push(f[C]),E++),a=1);return E===0?m[0]:m},r.range=function(i){return r.max(i)-r.min(i)},r.variance=function(i,s){return r.sumsqerr(i)/(i.length-(s?1:0))},r.pooledvariance=function(i){var s=i.reduce(function(a,d){return a+r.sumsqerr(d)},0),f=i.reduce(function(a,d){return a+d.length},0);return s/(f-i.length)},r.deviation=function(u){for(var i=r.mean(u),s=u.length,f=new Array(s),a=0;a<s;a++)f[a]=u[a]-i;return f},r.stdev=function(i,s){return o.sqrt(r.variance(i,s))},r.pooledstdev=function(i){return o.sqrt(r.pooledvariance(i))},r.meandev=function(i){for(var s=r.mean(i),f=[],a=i.length-1;a>=0;a--)f.push(o.abs(i[a]-s));return r.mean(f)},r.meddev=function(i){for(var s=r.median(i),f=[],a=i.length-1;a>=0;a--)f.push(o.abs(i[a]-s));return r.median(f)},r.coeffvar=function(i){return r.stdev(i)/r.mean(i)},r.quartiles=function(i){var s=i.length,f=i.slice().sort(c);return[f[o.round(s/4)-1],f[o.round(s/2)-1],f[o.round(s*3/4)-1]]},r.quantiles=function(i,s,f,a){var d=i.slice().sort(c),E=[s.length],m=i.length,C,v,p,I,A,M;for(typeof f>"u"&&(f=3/8),typeof a>"u"&&(a=3/8),C=0;C<s.length;C++)v=s[C],p=f+v*(1-f-a),I=m*v+p,A=o.floor(h(I,1,m-1)),M=h(I-A,0,1),E[C]=(1-M)*d[A-1]+M*d[A];return E},r.percentile=function(i,s,f){var a=i.slice().sort(c),d=s*(a.length+(f?1:-1))+(f?0:1),E=parseInt(d),m=d-E;return E+1<a.length?a[E-1]+m*(a[E]-a[E-1]):a[E-1]},r.percentileOfScore=function(i,s,f){var a=0,d=i.length,E=!1,m,C;for(f==="strict"&&(E=!0),C=0;C<d;C++)m=i[C],(E&&m<s||!E&&m<=s)&&a++;return a/d},r.histogram=function(i,s){s=s||4;var f=r.min(i),a=(r.max(i)-f)/s,d=i.length,E=[],m;for(m=0;m<s;m++)E[m]=0;for(m=0;m<d;m++)E[o.min(o.floor((i[m]-f)/a),s-1)]+=1;return E},r.covariance=function(i,s){var f=r.mean(i),a=r.mean(s),d=i.length,E=new Array(d),m;for(m=0;m<d;m++)E[m]=(i[m]-f)*(s[m]-a);return r.sum(E)/(d-1)},r.corrcoeff=function(i,s){return r.covariance(i,s)/r.stdev(i,1)/r.stdev(s,1)},r.spearmancoeff=function(u,i){return u=r.rank(u),i=r.rank(i),r.corrcoeff(u,i)},r.stanMoment=function(i,s){for(var f=r.mean(i),a=r.stdev(i),d=i.length,E=0,m=0;m<d;m++)E+=o.pow((i[m]-f)/a,s);return E/i.length},r.skewness=function(i){return r.stanMoment(i,3)},r.kurtosis=function(i){return r.stanMoment(i,4)-3};var g=r.prototype;(function(u){for(var i=0;i<u.length;i++)(function(s){g[s]=function(f,a){var d=[],E=0,m=this;if(l(f)&&(a=f,f=!1),a)return setTimeout(function(){a.call(m,g[s].call(m,f))}),this;if(this.length>1){for(m=f===!0?this:this.transpose();E<m.length;E++)d[E]=r[s](m[E]);return d}return r[s](this[0],f)}})(u[i])})("cumsum cumprod".split(" ")),(function(u){for(var i=0;i<u.length;i++)(function(s){g[s]=function(f,a){var d=[],E=0,m=this;if(l(f)&&(a=f,f=!1),a)return setTimeout(function(){a.call(m,g[s].call(m,f))}),this;if(this.length>1){for(s!=="sumrow"&&(m=f===!0?this:this.transpose());E<m.length;E++)d[E]=r[s](m[E]);return f===!0?r[s](r.utils.toVector(d)):d}return r[s](this[0],f)}})(u[i])})("sum sumsqrd sumsqerr sumrow product min max unique mean meansqerr geomean median diff rank mode range variance deviation stdev meandev meddev coeffvar quartiles histogram skewness kurtosis".split(" ")),(function(u){for(var i=0;i<u.length;i++)(function(s){g[s]=function(){var f=[],a=0,d=this,E=Array.prototype.slice.call(arguments),m;if(l(E[E.length-1])){m=E[E.length-1];var C=E.slice(0,E.length-1);return setTimeout(function(){m.call(d,g[s].apply(d,C))}),this}else{m=void 0;var v=function(I){return r[s].apply(d,[I].concat(E))}}if(this.length>1){for(d=d.transpose();a<d.length;a++)f[a]=v(d[a]);return f}return v(this[0])}})(u[i])})("quantiles percentileOfScore".split(" "))})(n,Math),(function(r,o){r.gammaln=function(c){var h=0,g=[76.18009172947146,-86.50532032941678,24.01409824083091,-1.231739572450155,.001208650973866179,-5395239384953e-18],u=1.000000000190015,i,s,f;for(f=(s=i=c)+5.5,f-=(i+.5)*o.log(f);h<6;h++)u+=g[h]/++s;return o.log(2.5066282746310007*u/i)-f},r.loggam=function(c){var h,g,u,i,s,f,a,d=[.08333333333333333,-.002777777777777778,.0007936507936507937,-.0005952380952380952,.0008417508417508418,-.001917526917526918,.00641025641025641,-.02955065359477124,.1796443723688307,-1.3924322169059];if(h=c,a=0,c==1||c==2)return 0;for(c<=7&&(a=o.floor(7-c),h=c+a),g=1/(h*h),u=2*o.PI,s=d[9],f=8;f>=0;f--)s*=g,s+=d[f];if(i=s/h+.5*o.log(u)+(h-.5)*o.log(h)-h,c<=7)for(f=1;f<=a;f++)i-=o.log(h-1),h-=1;return i},r.gammafn=function(c){var h=[-1.716185138865495,24.76565080557592,-379.80425647094563,629.3311553128184,866.9662027904133,-31451.272968848367,-36144.413418691176,66456.14382024054],g=[-30.8402300119739,315.35062697960416,-1015.1563674902192,-3107.771671572311,22538.11842098015,4755.846277527881,-134659.9598649693,-115132.2596755535],u=!1,i=0,s=0,f=0,a=c,d,E,m,C;if(c>171.6243769536076)return 1/0;if(a<=0)if(C=a%1+36e-17,C)u=(a&1?-1:1)*o.PI/o.sin(o.PI*C),a=1-a;else return 1/0;for(m=a,a<1?E=a++:E=(a-=i=(a|0)-1)-1,d=0;d<8;++d)f=(f+h[d])*E,s=s*E+g[d];if(C=f/s+1,m<a)C/=m;else if(m>a)for(d=0;d<i;++d)C*=a,a++;return u&&(C=u/C),C},r.gammap=function(c,h){return r.lowRegGamma(c,h)*r.gammafn(c)},r.lowRegGamma=function(c,h){var g=r.gammaln(c),u=c,i=1/c,s=i,f=h+1-c,a=1/1e-30,d=1/f,E=d,m=1,C=-~(o.log(c>=1?c:1/c)*8.5+c*.4+17),v;if(h<0||c<=0)return NaN;if(h<c+1){for(;m<=C;m++)i+=s*=h/++u;return i*o.exp(-h+c*o.log(h)-g)}for(;m<=C;m++)v=-m*(m-c),f+=2,d=v*d+f,a=f+v/a,d=1/d,E*=d*a;return 1-E*o.exp(-h+c*o.log(h)-g)},r.factorialln=function(c){return c<0?NaN:r.gammaln(c+1)},r.factorial=function(c){return c<0?NaN:r.gammafn(c+1)},r.combination=function(c,h){return c>170||h>170?o.exp(r.combinationln(c,h)):r.factorial(c)/r.factorial(h)/r.factorial(c-h)},r.combinationln=function(c,h){return r.factorialln(c)-r.factorialln(h)-r.factorialln(c-h)},r.permutation=function(c,h){return r.factorial(c)/r.factorial(c-h)},r.betafn=function(c,h){if(!(c<=0||h<=0))return c+h>170?o.exp(r.betaln(c,h)):r.gammafn(c)*r.gammafn(h)/r.gammafn(c+h)},r.betaln=function(c,h){return r.gammaln(c)+r.gammaln(h)-r.gammaln(c+h)},r.betacf=function(c,h,g){var u=1e-30,i=1,s=h+g,f=h+1,a=h-1,d=1,E=1-s*c/f,m,C,v,p;for(o.abs(E)<u&&(E=u),E=1/E,p=E;i<=100&&(m=2*i,C=i*(g-i)*c/((a+m)*(h+m)),E=1+C*E,o.abs(E)<u&&(E=u),d=1+C/d,o.abs(d)<u&&(d=u),E=1/E,p*=E*d,C=-(h+i)*(s+i)*c/((h+m)*(f+m)),E=1+C*E,o.abs(E)<u&&(E=u),d=1+C/d,o.abs(d)<u&&(d=u),E=1/E,v=E*d,p*=v,!(o.abs(v-1)<3e-7));i++);return p},r.gammapinv=function(c,h){var g=0,u=h-1,i=1e-8,s=r.gammaln(h),f,a,d,E,m,C,v;if(c>=1)return o.max(100,h+100*o.sqrt(h));if(c<=0)return 0;for(h>1?(C=o.log(u),v=o.exp(u*(C-1)-s),m=c<.5?c:1-c,d=o.sqrt(-2*o.log(m)),f=(2.30753+d*.27061)/(1+d*(.99229+d*.04481))-d,c<.5&&(f=-f),f=o.max(.001,h*o.pow(1-1/(9*h)-f/(3*o.sqrt(h)),3))):(d=1-h*(.253+h*.12),c<d?f=o.pow(c/d,1/h):f=1-o.log(1-(c-d)/(1-d)));g<12;g++){if(f<=0)return 0;if(a=r.lowRegGamma(h,f)-c,h>1?d=v*o.exp(-(f-u)+u*(o.log(f)-C)):d=o.exp(-f+u*o.log(f)-s),E=a/d,f-=d=E/(1-.5*o.min(1,E*((h-1)/f-1))),f<=0&&(f=.5*(f+d)),o.abs(d)<i*f)break}return f},r.erf=function(c){var h=[-1.3026537197817094,.6419697923564902,.019476473204185836,-.00956151478680863,-.000946595344482036,.000366839497852761,42523324806907e-18,-20278578112534e-18,-1624290004647e-18,130365583558e-17,15626441722e-18,-85238095915e-18,6529054439e-18,5059343495e-18,-991364156e-18,-227365122e-18,96467911e-18,2394038e-18,-6886027e-18,894487e-18,313092e-18,-112708e-18,381e-18,7106e-18,-1523e-18,-94e-18,121e-18,-28e-18],g=h.length-1,u=!1,i=0,s=0,f,a,d,E;for(c<0&&(c=-c,u=!0),f=2/(2+c),a=4*f-2;g>0;g--)d=i,i=a*i-s+h[g],s=d;return E=f*o.exp(-c*c+.5*(h[0]+a*i)-s),u?E-1:1-E},r.erfc=function(c){return 1-r.erf(c)},r.erfcinv=function(c){var h=0,g,u,i,s;if(c>=2)return-100;if(c<=0)return 100;for(s=c<1?c:2-c,i=o.sqrt(-2*o.log(s/2)),g=-.70711*((2.30753+i*.27061)/(1+i*(.99229+i*.04481))-i);h<2;h++)u=r.erfc(g)-s,g+=u/(1.1283791670955126*o.exp(-g*g)-g*u);return c<1?g:-g},r.ibetainv=function(c,h,g){var u=1e-8,i=h-1,s=g-1,f=0,a,d,E,m,C,v,p,I,A,M,O;if(c<=0)return 0;if(c>=1)return 1;for(h>=1&&g>=1?(E=c<.5?c:1-c,m=o.sqrt(-2*o.log(E)),p=(2.30753+m*.27061)/(1+m*(.99229+m*.04481))-m,c<.5&&(p=-p),I=(p*p-3)/6,A=2/(1/(2*h-1)+1/(2*g-1)),M=p*o.sqrt(I+A)/A-(1/(2*g-1)-1/(2*h-1))*(I+5/6-2/(3*A)),p=h/(h+g*o.exp(2*M))):(a=o.log(h/(h+g)),d=o.log(g/(h+g)),m=o.exp(h*a)/h,C=o.exp(g*d)/g,M=m+C,c<m/M?p=o.pow(h*M*c,1/h):p=1-o.pow(g*M*(1-c),1/g)),O=-r.gammaln(h)-r.gammaln(g)+r.gammaln(h+g);f<10;f++){if(p===0||p===1)return p;if(v=r.ibeta(p,h,g)-c,m=o.exp(i*o.log(p)+s*o.log(1-p)+O),C=v/m,p-=m=C/(1-.5*o.min(1,C*(i/p-s/(1-p)))),p<=0&&(p=.5*(p+m)),p>=1&&(p=.5*(p+m+1)),o.abs(m)<u*p&&f>0)break}return p},r.ibeta=function(c,h,g){var u=c===0||c===1?0:o.exp(r.gammaln(h+g)-r.gammaln(h)-r.gammaln(g)+h*o.log(c)+g*o.log(1-c));return c<0||c>1?!1:c<(h+1)/(h+g+2)?u*r.betacf(c,h,g)/h:1-u*r.betacf(1-c,g,h)/g},r.randn=function(c,h){var g,u,i,s,f;if(h||(h=c),c)return r.create(c,h,function(){return r.randn()});do g=r._random_fn(),u=1.7156*(r._random_fn()-.5),i=g-.449871,s=o.abs(u)+.386595,f=i*i+s*(.196*s-.25472*i);while(f>.27597&&(f>.27846||u*u>-4*o.log(g)*g*g));return u/g},r.randg=function(c,h,g){var u=c,i,s,f,a,d,E;if(g||(g=h),c||(c=1),h)return E=r.zeros(h,g),E.alter(function(){return r.randg(c)}),E;c<1&&(c+=1),i=c-1/3,s=1/o.sqrt(9*i);do{do d=r.randn(),a=1+s*d;while(a<=0);a=a*a*a,f=r._random_fn()}while(f>1-.331*o.pow(d,4)&&o.log(f)>.5*d*d+i*(1-a+o.log(a)));if(c==u)return i*a;do f=r._random_fn();while(f===0);return o.pow(f,1/u)*i*a},(function(l){for(var c=0;c<l.length;c++)(function(h){r.fn[h]=function(){return r(r.map(this,function(g){return r[h](g)}))}})(l[c])})("gammaln gammafn factorial factorialln".split(" ")),(function(l){for(var c=0;c<l.length;c++)(function(h){r.fn[h]=function(){return r(r[h].apply(null,arguments))}})(l[c])})("randn".split(" "))})(n,Math),(function(r,o){(function(u){for(var i=0;i<u.length;i++)(function(s){r[s]=function f(a,d,E){return this instanceof f?(this._a=a,this._b=d,this._c=E,this):new f(a,d,E)},r.fn[s]=function(f,a,d){var E=r[s](f,a,d);return E.data=this,E},r[s].prototype.sample=function(f){var a=this._a,d=this._b,E=this._c;return f?r.alter(f,function(){return r[s].sample(a,d,E)}):r[s].sample(a,d,E)},(function(f){for(var a=0;a<f.length;a++)(function(d){r[s].prototype[d]=function(E){var m=this._a,C=this._b,v=this._c;return!E&&E!==0&&(E=this.data),typeof E!="number"?r.fn.map.call(E,function(p){return r[s][d](p,m,C,v)}):r[s][d](E,m,C,v)}})(f[a])})("pdf cdf inv".split(" ")),(function(f){for(var a=0;a<f.length;a++)(function(d){r[s].prototype[d]=function(){return r[s][d](this._a,this._b,this._c)}})(f[a])})("mean median mode variance".split(" "))})(u[i])})("beta centralF cauchy chisquare exponential gamma invgamma kumaraswamy laplace lognormal noncentralt normal pareto studentt weibull uniform binomial negbin hypgeom poisson triangular tukey arcsine".split(" ")),r.extend(r.beta,{pdf:function(i,s,f){return i>1||i<0?0:s==1&&f==1?1:s<512&&f<512?o.pow(i,s-1)*o.pow(1-i,f-1)/r.betafn(s,f):o.exp((s-1)*o.log(i)+(f-1)*o.log(1-i)-r.betaln(s,f))},cdf:function(i,s,f){return i>1||i<0?(i>1)*1:r.ibeta(i,s,f)},inv:function(i,s,f){return r.ibetainv(i,s,f)},mean:function(i,s){return i/(i+s)},median:function(i,s){return r.ibetainv(.5,i,s)},mode:function(i,s){return(i-1)/(i+s-2)},sample:function(i,s){var f=r.randg(i);return f/(f+r.randg(s))},variance:function(i,s){return i*s/(o.pow(i+s,2)*(i+s+1))}}),r.extend(r.centralF,{pdf:function(i,s,f){var a,d,E;return i<0?0:s<=2?i===0&&s<2?1/0:i===0&&s===2?1:1/r.betafn(s/2,f/2)*o.pow(s/f,s/2)*o.pow(i,s/2-1)*o.pow(1+s/f*i,-(s+f)/2):(a=s*i/(f+i*s),d=f/(f+i*s),E=s*d/2,E*r.binomial.pdf((s-2)/2,(s+f-2)/2,a))},cdf:function(i,s,f){return i<0?0:r.ibeta(s*i/(s*i+f),s/2,f/2)},inv:function(i,s,f){return f/(s*(1/r.ibetainv(i,s/2,f/2)-1))},mean:function(i,s){return s>2?s/(s-2):void 0},mode:function(i,s){return i>2?s*(i-2)/(i*(s+2)):void 0},sample:function(i,s){var f=r.randg(i/2)*2,a=r.randg(s/2)*2;return f/i/(a/s)},variance:function(i,s){if(!(s<=4))return 2*s*s*(i+s-2)/(i*(s-2)*(s-2)*(s-4))}}),r.extend(r.cauchy,{pdf:function(i,s,f){return f<0?0:f/(o.pow(i-s,2)+o.pow(f,2))/o.PI},cdf:function(i,s,f){return o.atan((i-s)/f)/o.PI+.5},inv:function(u,i,s){return i+s*o.tan(o.PI*(u-.5))},median:function(i){return i},mode:function(i){return i},sample:function(i,s){return r.randn()*o.sqrt(1/(2*r.randg(.5)))*s+i}}),r.extend(r.chisquare,{pdf:function(i,s){return i<0?0:i===0&&s===2?.5:o.exp((s/2-1)*o.log(i)-i/2-s/2*o.log(2)-r.gammaln(s/2))},cdf:function(i,s){return i<0?0:r.lowRegGamma(s/2,i/2)},inv:function(u,i){return 2*r.gammapinv(u,.5*i)},mean:function(u){return u},median:function(i){return i*o.pow(1-2/(9*i),3)},mode:function(i){return i-2>0?i-2:0},sample:function(i){return r.randg(i/2)*2},variance:function(i){return 2*i}}),r.extend(r.exponential,{pdf:function(i,s){return i<0?0:s*o.exp(-s*i)},cdf:function(i,s){return i<0?0:1-o.exp(-s*i)},inv:function(u,i){return-o.log(1-u)/i},mean:function(u){return 1/u},median:function(u){return 1/u*o.log(2)},mode:function(){return 0},sample:function(i){return-1/i*o.log(r._random_fn())},variance:function(u){return o.pow(u,-2)}}),r.extend(r.gamma,{pdf:function(i,s,f){return i<0?0:i===0&&s===1?1/f:o.exp((s-1)*o.log(i)-i/f-r.gammaln(s)-s*o.log(f))},cdf:function(i,s,f){return i<0?0:r.lowRegGamma(s,i/f)},inv:function(u,i,s){return r.gammapinv(u,i)*s},mean:function(u,i){return u*i},mode:function(i,s){if(i>1)return(i-1)*s},sample:function(i,s){return r.randg(i)*s},variance:function(i,s){return i*s*s}}),r.extend(r.invgamma,{pdf:function(i,s,f){return i<=0?0:o.exp(-(s+1)*o.log(i)-f/i-r.gammaln(s)+s*o.log(f))},cdf:function(i,s,f){return i<=0?0:1-r.lowRegGamma(s,f/i)},inv:function(u,i,s){return s/r.gammapinv(1-u,i)},mean:function(u,i){return u>1?i/(u-1):void 0},mode:function(i,s){return s/(i+1)},sample:function(i,s){return s/r.randg(i)},variance:function(i,s){if(!(i<=2))return s*s/((i-1)*(i-1)*(i-2))}}),r.extend(r.kumaraswamy,{pdf:function(i,s,f){return i===0&&s===1?f:i===1&&f===1?s:o.exp(o.log(s)+o.log(f)+(s-1)*o.log(i)+(f-1)*o.log(1-o.pow(i,s)))},cdf:function(i,s,f){return i<0?0:i>1?1:1-o.pow(1-o.pow(i,s),f)},inv:function(i,s,f){return o.pow(1-o.pow(1-i,1/f),1/s)},mean:function(u,i){return i*r.gammafn(1+1/u)*r.gammafn(i)/r.gammafn(1+1/u+i)},median:function(i,s){return o.pow(1-o.pow(2,-1/s),1/i)},mode:function(i,s){if(i>=1&&s>=1&&i!==1&&s!==1)return o.pow((i-1)/(i*s-1),1/i)},variance:function(){throw new Error("variance not yet implemented")}}),r.extend(r.lognormal,{pdf:function(i,s,f){return i<=0?0:o.exp(-o.log(i)-.5*o.log(2*o.PI)-o.log(f)-o.pow(o.log(i)-s,2)/(2*f*f))},cdf:function(i,s,f){return i<0?0:.5+.5*r.erf((o.log(i)-s)/o.sqrt(2*f*f))},inv:function(u,i,s){return o.exp(-1.4142135623730951*s*r.erfcinv(2*u)+i)},mean:function(i,s){return o.exp(i+s*s/2)},median:function(i){return o.exp(i)},mode:function(i,s){return o.exp(i-s*s)},sample:function(i,s){return o.exp(r.randn()*s+i)},variance:function(i,s){return(o.exp(s*s)-1)*o.exp(2*i+s*s)}}),r.extend(r.noncentralt,{pdf:function(i,s,f){var a=1e-14;return o.abs(f)<a?r.studentt.pdf(i,s):o.abs(i)<a?o.exp(r.gammaln((s+1)/2)-f*f/2-.5*o.log(o.PI*s)-r.gammaln(s/2)):s/i*(r.noncentralt.cdf(i*o.sqrt(1+2/s),s+2,f)-r.noncentralt.cdf(i,s,f))},cdf:function(i,s,f){var a=1e-14,d=200;if(o.abs(f)<a)return r.studentt.cdf(i,s);var E=!1;i<0&&(E=!0,f=-f);for(var m=r.normal.cdf(-f,0,1),C=a+1,v=C,p=i*i/(i*i+s),I=0,A=o.exp(-f*f/2),M=o.exp(-f*f/2-.5*o.log(2)-r.gammaln(3/2))*f;I<d||v>a||C>a;)v=C,I>0&&(A*=f*f/(2*I),M*=f*f/(2*(I+1/2))),C=A*r.beta.cdf(p,I+.5,s/2)+M*r.beta.cdf(p,I+1,s/2),m+=.5*C,I++;return E?1-m:m}}),r.extend(r.normal,{pdf:function(i,s,f){return o.exp(-.5*o.log(2*o.PI)-o.log(f)-o.pow(i-s,2)/(2*f*f))},cdf:function(i,s,f){return .5*(1+r.erf((i-s)/o.sqrt(2*f*f)))},inv:function(u,i,s){return-1.4142135623730951*s*r.erfcinv(2*u)+i},mean:function(u){return u},median:function(i){return i},mode:function(u){return u},sample:function(i,s){return r.randn()*s+i},variance:function(u,i){return i*i}}),r.extend(r.pareto,{pdf:function(i,s,f){return i<s?0:f*o.pow(s,f)/o.pow(i,f+1)},cdf:function(i,s,f){return i<s?0:1-o.pow(s/i,f)},inv:function(i,s,f){return s/o.pow(1-i,1/f)},mean:function(i,s){if(!(s<=1))return s*o.pow(i,s)/(s-1)},median:function(i,s){return i*(s*o.SQRT2)},mode:function(i){return i},variance:function(u,i){if(!(i<=2))return u*u*i/(o.pow(i-1,2)*(i-2))}}),r.extend(r.studentt,{pdf:function(i,s){return s=s>1e100?1e100:s,1/(o.sqrt(s)*r.betafn(.5,s/2))*o.pow(1+i*i/s,-((s+1)/2))},cdf:function(i,s){var f=s/2;return r.ibeta((i+o.sqrt(i*i+s))/(2*o.sqrt(i*i+s)),f,f)},inv:function(u,i){var s=r.ibetainv(2*o.min(u,1-u),.5*i,.5);return s=o.sqrt(i*(1-s)/s),u>.5?s:-s},mean:function(i){return i>1?0:void 0},median:function(){return 0},mode:function(){return 0},sample:function(i){return r.randn()*o.sqrt(i/(2*r.randg(i/2)))},variance:function(i){return i>2?i/(i-2):i>1?1/0:void 0}}),r.extend(r.weibull,{pdf:function(i,s,f){return i<0||s<0||f<0?0:f/s*o.pow(i/s,f-1)*o.exp(-o.pow(i/s,f))},cdf:function(i,s,f){return i<0?0:1-o.exp(-o.pow(i/s,f))},inv:function(u,i,s){return i*o.pow(-o.log(1-u),1/s)},mean:function(u,i){return u*r.gammafn(1+1/i)},median:function(i,s){return i*o.pow(o.log(2),1/s)},mode:function(i,s){return s<=1?0:i*o.pow((s-1)/s,1/s)},sample:function(i,s){return i*o.pow(-o.log(r._random_fn()),1/s)},variance:function(i,s){return i*i*r.gammafn(1+2/s)-o.pow(r.weibull.mean(i,s),2)}}),r.extend(r.uniform,{pdf:function(i,s,f){return i<s||i>f?0:1/(f-s)},cdf:function(i,s,f){return i<s?0:i<f?(i-s)/(f-s):1},inv:function(u,i,s){return i+u*(s-i)},mean:function(i,s){return .5*(i+s)},median:function(i,s){return r.mean(i,s)},mode:function(){throw new Error("mode is not yet implemented")},sample:function(i,s){return i/2+s/2+(s/2-i/2)*(2*r._random_fn()-1)},variance:function(i,s){return o.pow(s-i,2)/12}});function l(u,i,s,f){for(var a=0,d=1,E=1,m=1,C=0,v=0,p;o.abs((E-v)/E)>f;)v=E,p=-(i+C)*(i+s+C)*u/(i+2*C)/(i+2*C+1),a=E+p*a,d=m+p*d,C=C+1,p=C*(s-C)*u/(i+2*C-1)/(i+2*C),E=a+p*E,m=d+p*m,a=a/m,d=d/m,E=E/m,m=1;return E/i}r.extend(r.binomial,{pdf:function(i,s,f){return f===0||f===1?s*f===i?1:0:r.combination(s,i)*o.pow(f,i)*o.pow(1-f,s-i)},cdf:function(i,s,f){var a,d=1e-10;if(i<0)return 0;if(i>=s)return 1;if(f<0||f>1||s<=0)return NaN;i=o.floor(i);var E=f,m=i+1,C=s-i,v=m+C,p=o.exp(r.gammaln(v)-r.gammaln(C)-r.gammaln(m)+m*o.log(E)+C*o.log(1-E));return E<(m+1)/(v+2)?a=p*l(E,m,C,d):a=1-p*l(1-E,C,m,d),o.round((1-a)*(1/d))/(1/d)}}),r.extend(r.negbin,{pdf:function(i,s,f){return i!==i>>>0?!1:i<0?0:r.combination(i+s-1,s-1)*o.pow(1-f,i)*o.pow(f,s)},cdf:function(i,s,f){var a=0,d=0;if(i<0)return 0;for(;d<=i;d++)a+=r.negbin.pdf(d,s,f);return a}}),r.extend(r.hypgeom,{pdf:function(i,s,f,a){if(i!==i|0)return!1;if(i<0||i<f-(s-a))return 0;if(i>a||i>f)return 0;if(f*2>s)return a*2>s?r.hypgeom.pdf(s-f-a+i,s,s-f,s-a):r.hypgeom.pdf(a-i,s,s-f,a);if(a*2>s)return r.hypgeom.pdf(f-i,s,f,s-a);if(f<a)return r.hypgeom.pdf(i,s,a,f);for(var d=1,E=0,m=0;m<i;m++){for(;d>1&&E<a;)d*=1-f/(s-E),E++;d*=(a-m)*(f-m)/((m+1)*(s-f-a+m+1))}for(;E<a;E++)d*=1-f/(s-E);return o.min(1,o.max(0,d))},cdf:function(i,s,f,a){if(i<0||i<f-(s-a))return 0;if(i>=a||i>=f)return 1;if(f*2>s)return a*2>s?r.hypgeom.cdf(s-f-a+i,s,s-f,s-a):1-r.hypgeom.cdf(a-i-1,s,s-f,a);if(a*2>s)return 1-r.hypgeom.cdf(f-i-1,s,f,s-a);if(f<a)return r.hypgeom.cdf(i,s,a,f);for(var d=1,E=1,m=0,C=0;C<i;C++){for(;d>1&&m<a;){var v=1-f/(s-m);E*=v,d*=v,m++}E*=(a-C)*(f-C)/((C+1)*(s-f-a+C+1)),d+=E}for(;m<a;m++)d*=1-f/(s-m);return o.min(1,o.max(0,d))}}),r.extend(r.poisson,{pdf:function(i,s){return s<0||i%1!==0||i<0?0:o.pow(s,i)*o.exp(-s)/r.factorial(i)},cdf:function(i,s){var f=[],a=0;if(i<0)return 0;for(;a<=i;a++)f.push(r.poisson.pdf(a,s));return r.sum(f)},mean:function(u){return u},variance:function(u){return u},sampleSmall:function(i){var s=1,f=0,a=o.exp(-i);do f++,s*=r._random_fn();while(s>a);return f-1},sampleLarge:function(i){var s=i,f,a,d,E,m,C,v,p,I,A;for(E=o.sqrt(s),m=o.log(s),v=.931+2.53*E,C=-.059+.02483*v,p=1.1239+1.1328/(v-3.4),I=.9277-3.6224/(v-2);;){if(a=o.random()-.5,d=o.random(),A=.5-o.abs(a),f=o.floor((2*C/A+v)*a+s+.43),A>=.07&&d<=I)return f;if(!(f<0||A<.013&&d>A)&&o.log(d)+o.log(p)-o.log(C/(A*A)+v)<=-s+f*m-r.loggam(f+1))return f}},sample:function(i){return i<10?this.sampleSmall(i):this.sampleLarge(i)}}),r.extend(r.triangular,{pdf:function(i,s,f,a){return f<=s||a<s||a>f?NaN:i<s||i>f?0:i<a?2*(i-s)/((f-s)*(a-s)):i===a?2/(f-s):2*(f-i)/((f-s)*(f-a))},cdf:function(i,s,f,a){return f<=s||a<s||a>f?NaN:i<=s?0:i>=f?1:i<=a?o.pow(i-s,2)/((f-s)*(a-s)):1-o.pow(f-i,2)/((f-s)*(f-a))},inv:function(i,s,f,a){return f<=s||a<s||a>f?NaN:i<=(a-s)/(f-s)?s+(f-s)*o.sqrt(i*((a-s)/(f-s))):s+(f-s)*(1-o.sqrt((1-i)*(1-(a-s)/(f-s))))},mean:function(i,s,f){return(i+s+f)/3},median:function(i,s,f){if(f<=(i+s)/2)return s-o.sqrt((s-i)*(s-f))/o.sqrt(2);if(f>(i+s)/2)return i+o.sqrt((s-i)*(f-i))/o.sqrt(2)},mode:function(i,s,f){return f},sample:function(i,s,f){var a=r._random_fn();return a<(f-i)/(s-i)?i+o.sqrt(a*(s-i)*(f-i)):s-o.sqrt((1-a)*(s-i)*(s-f))},variance:function(i,s,f){return(i*i+s*s+f*f-i*s-i*f-s*f)/18}}),r.extend(r.arcsine,{pdf:function(i,s,f){return f<=s?NaN:i<=s||i>=f?0:2/o.PI*o.pow(o.pow(f-s,2)-o.pow(2*i-s-f,2),-.5)},cdf:function(i,s,f){return i<s?0:i<f?2/o.PI*o.asin(o.sqrt((i-s)/(f-s))):1},inv:function(u,i,s){return i+(.5-.5*o.cos(o.PI*u))*(s-i)},mean:function(i,s){return s<=i?NaN:(i+s)/2},median:function(i,s){return s<=i?NaN:(i+s)/2},mode:function(){throw new Error("mode is not yet implemented")},sample:function(i,s){return(i+s)/2+(s-i)/2*o.sin(2*o.PI*r.uniform.sample(0,1))},variance:function(i,s){return s<=i?NaN:o.pow(s-i,2)/8}});function c(u){return u/o.abs(u)}r.extend(r.laplace,{pdf:function(i,s,f){return f<=0?0:o.exp(-o.abs(i-s)/f)/(2*f)},cdf:function(i,s,f){return f<=0?0:i<s?.5*o.exp((i-s)/f):1-.5*o.exp(-(i-s)/f)},mean:function(u){return u},median:function(u){return u},mode:function(u){return u},variance:function(u,i){return 2*i*i},sample:function(i,s){var f=r._random_fn()-.5;return i-s*c(f)*o.log(1-2*o.abs(f))}});function h(u,i,s){var f=12,a=6,d=-30,E=-50,m=60,C=8,v=3,p=2,I=3,A=[.9815606342467192,.9041172563704749,.7699026741943047,.5873179542866175,.3678314989981802,.1252334085114689],M=[.04717533638651183,.10693932599531843,.16007832854334622,.20316742672306592,.2334925365383548,.24914704581340277],O=u*.5;if(O>=C)return 1;var L=2*r.normal.cdf(O,0,1,1,0)-1;L>=o.exp(E/s)?L=o.pow(L,s):L=0;var U;u>v?U=p:U=I;for(var H=O,y=(C-O)/U,Z=H+y,Q=0,te=s-1,X=1;X<=U;X++){for(var $=0,G=.5*(Z+H),z=.5*(Z-H),ue=1;ue<=f;ue++){var W,re;a<ue?(W=f-ue+1,re=A[W-1]):(W=ue,re=-A[W-1]);var le=z*re,ne=G+le,we=ne*ne;if(we>m)break;var Xe=2*r.normal.cdf(ne,0,1,1,0),gt=2*r.normal.cdf(ne,u,1,1,0),ke=Xe*.5-gt*.5;ke>=o.exp(d/te)&&(ke=M[W-1]*o.exp(-(.5*we))*o.pow(ke,te),$+=ke)}$*=2*z*s/o.sqrt(2*o.PI),Q+=$,H=Z,Z+=y}return L+=Q,L<=o.exp(d/i)?0:(L=o.pow(L,i),L>=1?1:L)}function g(u,i,s){var f=.322232421088,a=.099348462606,d=-1,E=.588581570495,m=-.342242088547,C=.531103462366,v=-.204231210125,p=.10353775285,I=-453642210148e-16,A=.0038560700634,M=.8832,O=.2368,L=1.208,U=1.4142,H=120,y=.5-.5*u,Z=o.sqrt(o.log(1/(y*y))),Q=Z+((((Z*I+v)*Z+m)*Z+d)*Z+f)/((((Z*A+p)*Z+C)*Z+E)*Z+a);s<H&&(Q+=(Q*Q*Q+Q)/s/4);var te=M-O*Q;return s<H&&(te+=-1.214/s+L*Q/s),Q*(te*o.log(i-1)+U)}r.extend(r.tukey,{cdf:function(i,s,f){var a=1,d=s,E=16,m=8,C=-30,v=1e-14,p=100,I=800,A=5e3,M=25e3,O=1,L=.5,U=.25,H=.125,y=[.9894009349916499,.9445750230732326,.8656312023878318,.755404408355003,.6178762444026438,.45801677765722737,.2816035507792589,.09501250983763744],Z=[.027152459411754096,.062253523938647894,.09515851168249279,.12462897125553388,.14959598881657674,.16915651939500254,.18260341504492358,.1894506104550685];if(i<=0)return 0;if(f<2||a<1||d<2)return NaN;if(!Number.isFinite(i))return 1;if(f>M)return h(i,a,d);var Q=f*.5,te=Q*o.log(f)-f*o.log(2)-r.gammaln(Q),X=Q-1,$=f*.25,G;f<=p?G=O:f<=I?G=L:f<=A?G=U:G=H,te+=o.log(G);for(var z=0,ue=1;ue<=50;ue++){for(var W=0,re=(2*ue-1)*G,le=1;le<=E;le++){var ne,we;m<le?(ne=le-m-1,we=te+X*o.log(re+y[ne]*G)-(y[ne]*G+re)*$):(ne=le-1,we=te+X*o.log(re-y[ne]*G)+(y[ne]*G-re)*$);var Xe;if(we>=C){m<le?Xe=i*o.sqrt((y[ne]*G+re)*.5):Xe=i*o.sqrt((-(y[ne]*G)+re)*.5);var gt=h(Xe,a,d),ke=gt*Z[ne]*o.exp(we);W+=ke}}if(ue*G>=1&&W<=v)break;z+=W}if(W>v)throw new Error("tukey.cdf failed to converge");return z>1&&(z=1),z},inv:function(u,i,s){var f=1,a=i,d=1e-4,E=50;if(s<2||f<1||a<2)return NaN;if(u<0||u>1)return NaN;if(u===0)return 0;if(u===1)return 1/0;var m=g(u,a,s),C=r.tukey.cdf(m,i,s)-u,v;C>0?v=o.max(0,m-1):v=m+1;for(var p=r.tukey.cdf(v,i,s)-u,I,A=1;A<E;A++){I=v-p*(v-m)/(p-C),C=p,m=v,I<0&&(I=0,p=-u),p=r.tukey.cdf(I,i,s)-u,v=I;var M=o.abs(v-m);if(M<d)return I}throw new Error("tukey.inv failed to converge")}})})(n,Math),(function(r,o){var l=Array.prototype.push,c=r.utils.isArray;function h(g){return c(g)||g instanceof r}r.extend({add:function(u,i){return h(i)?(h(i[0])||(i=[i]),r.map(u,function(s,f,a){return s+i[f][a]})):r.map(u,function(s){return s+i})},subtract:function(u,i){return h(i)?(h(i[0])||(i=[i]),r.map(u,function(s,f,a){return s-i[f][a]||0})):r.map(u,function(s){return s-i})},divide:function(u,i){return h(i)?(h(i[0])||(i=[i]),r.multiply(u,r.inv(i))):r.map(u,function(s){return s/i})},multiply:function(u,i){var s,f,a,d,E,m,C,v;if(u.length===void 0&&i.length===void 0)return u*i;if(E=u.length,m=u[0].length,C=r.zeros(E,a=h(i)?i[0].length:m),v=0,h(i)){for(;v<a;v++)for(s=0;s<E;s++){for(d=0,f=0;f<m;f++)d+=u[s][f]*i[f][v];C[s][v]=d}return E===1&&v===1?C[0][0]:C}return r.map(u,function(p){return p*i})},outer:function(u,i){return r.multiply(u.map(function(s){return[s]}),[i])},dot:function(u,i){h(u[0])||(u=[u]),h(i[0])||(i=[i]);for(var s=u[0].length===1&&u.length!==1?r.transpose(u):u,f=i[0].length===1&&i.length!==1?r.transpose(i):i,a=[],d=0,E=s.length,m=s[0].length,C,v;d<E;d++){for(a[d]=[],C=0,v=0;v<m;v++)C+=s[d][v]*f[d][v];a[d]=C}return a.length===1?a[0]:a},pow:function(u,i){return r.map(u,function(s){return o.pow(s,i)})},exp:function(u){return r.map(u,function(i){return o.exp(i)})},log:function(u){return r.map(u,function(i){return o.log(i)})},abs:function(u){return r.map(u,function(i){return o.abs(i)})},norm:function(u,i){var s=0,f=0;for(isNaN(i)&&(i=2),h(u[0])&&(u=u[0]);f<u.length;f++)s+=o.pow(o.abs(u[f]),i);return o.pow(s,1/i)},angle:function(u,i){return o.acos(r.dot(u,i)/(r.norm(u)*r.norm(i)))},aug:function(u,i){var s=[],f;for(f=0;f<u.length;f++)s.push(u[f].slice());for(f=0;f<s.length;f++)l.apply(s[f],i[f]);return s},inv:function(u){for(var i=u.length,s=u[0].length,f=r.identity(i,s),a=r.gauss_jordan(u,f),d=[],E=0,m;E<i;E++)for(d[E]=[],m=s;m<a[0].length;m++)d[E][m-s]=a[E][m];return d},det:function g(u){if(u.length===2)return u[0][0]*u[1][1]-u[0][1]*u[1][0];for(var i=0,s=0;s<u.length;s++){for(var f=[],a=1;a<u.length;a++){f[a-1]=[];for(var d=0;d<u.length;d++)d<s?f[a-1][d]=u[a][d]:d>s&&(f[a-1][d-1]=u[a][d])}var E=s%2?-1:1;i+=g(f)*u[0][s]*E}return i},gauss_elimination:function(u,i){var s=0,f=0,a=u.length,d=u[0].length,E=1,m=0,C=[],v,p,I,A;for(u=r.aug(u,i),v=u[0].length,s=0;s<a;s++){for(p=u[s][s],f=s,A=s+1;A<d;A++)p<o.abs(u[A][s])&&(p=u[A][s],f=A);if(f!=s)for(A=0;A<v;A++)I=u[s][A],u[s][A]=u[f][A],u[f][A]=I;for(f=s+1;f<a;f++)for(E=u[f][s]/u[s][s],A=s;A<v;A++)u[f][A]=u[f][A]-E*u[s][A]}for(s=a-1;s>=0;s--){for(m=0,f=s+1;f<=a-1;f++)m=m+C[f]*u[s][f];C[s]=(u[s][v-1]-m)/u[s][s]}return C},gauss_jordan:function(u,i){var s=r.aug(u,i),f=s.length,a=s[0].length,d=0,E,m,C;for(m=0;m<f;m++){var v=m;for(C=m+1;C<f;C++)o.abs(s[C][m])>o.abs(s[v][m])&&(v=C);var p=s[m];for(s[m]=s[v],s[v]=p,C=m+1;C<f;C++)for(d=s[C][m]/s[m][m],E=m;E<a;E++)s[C][E]-=s[m][E]*d}for(m=f-1;m>=0;m--){for(d=s[m][m],C=0;C<m;C++)for(E=a-1;E>m-1;E--)s[C][E]-=s[m][E]*s[C][m]/d;for(s[m][m]/=d,E=f;E<a;E++)s[m][E]/=d}return s},triaUpSolve:function(u,i){var s=u[0].length,f=r.zeros(1,s)[0],a,d=!1;return i[0].length!=null&&(i=i.map(function(E){return E[0]}),d=!0),r.arange(s-1,-1,-1).forEach(function(E){a=r.arange(E+1,s).map(function(m){return f[m]*u[E][m]}),f[E]=(i[E]-r.sum(a))/u[E][E]}),d?f.map(function(E){return[E]}):f},triaLowSolve:function(u,i){var s=u[0].length,f=r.zeros(1,s)[0],a,d=!1;return i[0].length!=null&&(i=i.map(function(E){return E[0]}),d=!0),r.arange(s).forEach(function(E){a=r.arange(E).map(function(m){return u[E][m]*f[m]}),f[E]=(i[E]-r.sum(a))/u[E][E]}),d?f.map(function(E){return[E]}):f},lu:function(u){var i=u.length,s=r.identity(i),f=r.zeros(u.length,u[0].length),a;return r.arange(i).forEach(function(d){f[0][d]=u[0][d]}),r.arange(1,i).forEach(function(d){r.arange(d).forEach(function(E){a=r.arange(E).map(function(m){return s[d][m]*f[m][E]}),s[d][E]=(u[d][E]-r.sum(a))/f[E][E]}),r.arange(d,i).forEach(function(E){a=r.arange(d).map(function(m){return s[d][m]*f[m][E]}),f[d][E]=u[a.length][E]-r.sum(a)})}),[s,f]},cholesky:function(u){var i=u.length,s=r.zeros(u.length,u[0].length),f;return r.arange(i).forEach(function(a){f=r.arange(a).map(function(d){return o.pow(s[a][d],2)}),s[a][a]=o.sqrt(u[a][a]-r.sum(f)),r.arange(a+1,i).forEach(function(d){f=r.arange(a).map(function(E){return s[a][E]*s[d][E]}),s[d][a]=(u[a][d]-r.sum(f))/s[a][a]})}),s},gauss_jacobi:function(u,i,s,f){for(var a=0,d=0,E=u.length,m=[],C=[],v=[],p,I,A,M;a<E;a++)for(m[a]=[],C[a]=[],v[a]=[],d=0;d<E;d++)a>d?(m[a][d]=u[a][d],C[a][d]=v[a][d]=0):a<d?(C[a][d]=u[a][d],m[a][d]=v[a][d]=0):(v[a][d]=u[a][d],m[a][d]=C[a][d]=0);for(A=r.multiply(r.multiply(r.inv(v),r.add(m,C)),-1),I=r.multiply(r.inv(v),i),p=s,M=r.add(r.multiply(A,s),I),a=2;o.abs(r.norm(r.subtract(M,p)))>f;)p=M,M=r.add(r.multiply(A,p),I),a++;return M},gauss_seidel:function(u,i,s,f){for(var a=0,d=u.length,E=[],m=[],C=[],v,p,I,A,M;a<d;a++)for(E[a]=[],m[a]=[],C[a]=[],v=0;v<d;v++)a>v?(E[a][v]=u[a][v],m[a][v]=C[a][v]=0):a<v?(m[a][v]=u[a][v],E[a][v]=C[a][v]=0):(C[a][v]=u[a][v],E[a][v]=m[a][v]=0);for(A=r.multiply(r.multiply(r.inv(r.add(C,E)),m),-1),I=r.multiply(r.inv(r.add(C,E)),i),p=s,M=r.add(r.multiply(A,s),I),a=2;o.abs(r.norm(r.subtract(M,p)))>f;)p=M,M=r.add(r.multiply(A,p),I),a=a+1;return M},SOR:function(u,i,s,f,a){for(var d=0,E=u.length,m=[],C=[],v=[],p,I,A,M,O;d<E;d++)for(m[d]=[],C[d]=[],v[d]=[],p=0;p<E;p++)d>p?(m[d][p]=u[d][p],C[d][p]=v[d][p]=0):d<p?(C[d][p]=u[d][p],m[d][p]=v[d][p]=0):(v[d][p]=u[d][p],m[d][p]=C[d][p]=0);for(M=r.multiply(r.inv(r.add(v,r.multiply(m,a))),r.subtract(r.multiply(v,1-a),r.multiply(C,a))),A=r.multiply(r.multiply(r.inv(r.add(v,r.multiply(m,a))),i),a),I=s,O=r.add(r.multiply(M,s),A),d=2;o.abs(r.norm(r.subtract(O,I)))>f;)I=O,O=r.add(r.multiply(M,I),A),d++;return O},householder:function(u){for(var i=u.length,s=u[0].length,f=0,a=[],d=[],E,m,C,v,p;f<i-1;f++){for(E=0,v=f+1;v<s;v++)E+=u[v][f]*u[v][f];for(p=u[f+1][f]>0?-1:1,E=p*o.sqrt(E),m=o.sqrt((E*E-u[f+1][f]*E)/2),a=r.zeros(i,1),a[f+1][0]=(u[f+1][f]-E)/(2*m),C=f+2;C<i;C++)a[C][0]=u[C][f]/(2*m);d=r.subtract(r.identity(i,s),r.multiply(r.multiply(a,r.transpose(a)),2)),u=r.multiply(d,r.multiply(u,d))}return u},QR:(function(){var g=r.sum,u=r.arange;function i(s){var f=s.length,a=s[0].length,d=r.zeros(a,a);s=r.copy(s);var E,m,C;for(m=0;m<a;m++){for(d[m][m]=o.sqrt(g(u(f).map(function(v){return s[v][m]*s[v][m]}))),E=0;E<f;E++)s[E][m]=s[E][m]/d[m][m];for(C=m+1;C<a;C++)for(d[m][C]=g(u(f).map(function(v){return s[v][m]*s[v][C]})),E=0;E<f;E++)s[E][C]=s[E][C]-s[E][m]*d[m][C]}return[s,d]}return i})(),lstsq:(function(){function g(i){i=r.copy(i);var s=i.length,f=r.identity(s);return r.arange(s-1,-1,-1).forEach(function(a){r.sliceAssign(f,{row:a},r.divide(r.slice(f,{row:a}),i[a][a])),r.sliceAssign(i,{row:a},r.divide(r.slice(i,{row:a}),i[a][a])),r.arange(a).forEach(function(d){var E=r.multiply(i[d][a],-1),m=r.slice(i,{row:d}),C=r.multiply(r.slice(i,{row:a}),E);r.sliceAssign(i,{row:d},r.add(m,C));var v=r.slice(f,{row:d}),p=r.multiply(r.slice(f,{row:a}),E);r.sliceAssign(f,{row:d},r.add(v,p))})}),f}function u(i,s){var f=!1;s[0].length===void 0&&(s=s.map(function(M){return[M]}),f=!0);var a=r.QR(i),d=a[0],E=a[1],m=i[0].length,C=r.slice(d,{col:{end:m}}),v=r.slice(E,{row:{end:m}}),p=g(v),I=r.transpose(C);I[0].length===void 0&&(I=[I]);var A=r.multiply(r.multiply(p,I),s);return A.length===void 0&&(A=[[A]]),f?A.map(function(M){return M[0]}):A}return u})(),jacobi:function(u){for(var i=1,s=u.length,f=r.identity(s,s),a=[],d,E,m,C,v,p,I,A;i===1;){for(p=u[0][1],C=0,v=1,E=0;E<s;E++)for(m=0;m<s;m++)E!=m&&p<o.abs(u[E][m])&&(p=o.abs(u[E][m]),C=E,v=m);for(u[C][C]===u[v][v]?I=u[C][v]>0?o.PI/4:-o.PI/4:I=o.atan(2*u[C][v]/(u[C][C]-u[v][v]))/2,A=r.identity(s,s),A[C][C]=o.cos(I),A[C][v]=-o.sin(I),A[v][C]=o.sin(I),A[v][v]=o.cos(I),f=r.multiply(f,A),d=r.multiply(r.multiply(r.inv(A),u),A),u=d,i=0,E=1;E<s;E++)for(m=1;m<s;m++)E!=m&&o.abs(u[E][m])>.001&&(i=1)}for(E=0;E<s;E++)a.push(u[E][E]);return[f,a]},rungekutta:function(u,i,s,f,a,d){var E,m,C,v,p;if(d===2)for(;f<=s;)E=i*u(f,a),m=i*u(f+i,a+E),C=a+(E+m)/2,a=C,f=f+i;if(d===4)for(;f<=s;)E=i*u(f,a),m=i*u(f+i/2,a+E/2),v=i*u(f+i/2,a+m/2),p=i*u(f+i,a+v),C=a+(E+2*m+2*v+p)/6,a=C,f=f+i;return a},romberg:function(u,i,s,f){for(var a=0,d=(s-i)/2,E=[],m=[],C=[],v,p,I,A,M;a<f/2;){for(M=u(i),I=i,A=0;I<=s;I=I+d,A++)E[A]=I;for(v=E.length,I=1;I<v-1;I++)M+=(I%2!==0?4:2)*u(E[I]);M=d/3*(M+u(s)),C[a]=M,d/=2,a++}for(p=C.length,v=1;p!==1;){for(I=0;I<p-1;I++)m[I]=(o.pow(4,v)*C[I+1]-C[I])/(o.pow(4,v)-1);p=m.length,C=m,m=[],v++}return C},richardson:function(u,i,s,f){function a(O,L){for(var U=0,H=O.length,y;U<H;U++)O[U]===L&&(y=U);return y}for(var d=o.abs(s-u[a(u,s)+1]),E=0,m=[],C=[],v,p,I,A,M;f>=d;)v=a(u,s+f),p=a(u,s),m[E]=(i[v]-2*i[p]+i[2*p-v])/(f*f),f/=2,E++;for(A=m.length,I=1;A!=1;){for(M=0;M<A-1;M++)C[M]=(o.pow(4,I)*m[M+1]-m[M])/(o.pow(4,I)-1);A=C.length,m=C,C=[],I++}return m},simpson:function(u,i,s,f){for(var a=(s-i)/f,d=u(i),E=[],m=i,C=0,v=1,p;m<=s;m=m+a,C++)E[C]=m;for(p=E.length;v<p-1;v++)d+=(v%2!==0?4:2)*u(E[v]);return a/3*(d+u(s))},hermite:function(u,i,s,f){for(var a=u.length,d=0,E=0,m=[],C=[],v=[],p=[],I;E<a;E++){for(m[E]=1,I=0;I<a;I++)E!=I&&(m[E]*=(f-u[I])/(u[E]-u[I]));for(C[E]=0,I=0;I<a;I++)E!=I&&(C[E]+=1/(u[E]-u[I]));v[E]=(1-2*(f-u[E])*C[E])*(m[E]*m[E]),p[E]=(f-u[E])*(m[E]*m[E]),d+=v[E]*i[E]+p[E]*s[E]}return d},lagrange:function(u,i,s){for(var f=0,a=0,d,E,m=u.length;a<m;a++){for(E=i[a],d=0;d<m;d++)a!=d&&(E*=(s-u[d])/(u[a]-u[d]));f+=E}return f},cubic_spline:function(u,i,s){for(var f=u.length,a=0,d,E=[],m=[],C=[],v=[],p=[],I=[],A=[];a<f-1;a++)p[a]=u[a+1]-u[a];for(C[0]=0,a=1;a<f-1;a++)C[a]=3/p[a]*(i[a+1]-i[a])-3/p[a-1]*(i[a]-i[a-1]);for(a=1;a<f-1;a++)E[a]=[],m[a]=[],E[a][a-1]=p[a-1],E[a][a]=2*(p[a-1]+p[a]),E[a][a+1]=p[a],m[a][0]=C[a];for(v=r.multiply(r.inv(E),m),d=0;d<f-1;d++)I[d]=(i[d+1]-i[d])/p[d]-p[d]*(v[d+1][0]+2*v[d][0])/3,A[d]=(v[d+1][0]-v[d][0])/(3*p[d]);for(d=0;d<f&&!(u[d]>s);d++);return d-=1,i[d]+(s-u[d])*I[d]+r.sq(s-u[d])*v[d]+(s-u[d])*r.sq(s-u[d])*A[d]},gauss_quadrature:function(){throw new Error("gauss_quadrature not yet implemented")},PCA:function(u){var i=u.length,s=u[0].length,f=0,a,d,E=[],m=[],C=[],v=[],p=[],I=[],A=[],M=[],O=[],L=[];for(f=0;f<i;f++)E[f]=r.sum(u[f])/s;for(f=0;f<s;f++)for(A[f]=[],a=0;a<i;a++)A[f][a]=u[a][f]-E[a];for(A=r.transpose(A),f=0;f<i;f++)for(M[f]=[],a=0;a<i;a++)M[f][a]=r.dot([A[f]],[A[a]])/(s-1);for(C=r.jacobi(M),O=C[0],m=C[1],L=r.transpose(O),f=0;f<m.length;f++)for(a=f;a<m.length;a++)m[f]<m[a]&&(d=m[f],m[f]=m[a],m[a]=d,v=L[f],L[f]=L[a],L[a]=v);for(I=r.transpose(A),f=0;f<i;f++)for(p[f]=[],a=0;a<I.length;a++)p[f][a]=r.dot([L[f]],[I[a]]);return[u,m,L,p]}}),(function(g){for(var u=0;u<g.length;u++)(function(i){r.fn[i]=function(s,f){var a=this;return f?(setTimeout(function(){f.call(a,r.fn[i].call(a,s))},15),this):typeof r[i](this,s)=="number"?r[i](this,s):r(r[i](this,s))}})(g[u])})("add divide multiply subtract dot pow exp log abs norm angle".split(" "))})(n,Math),(function(r,o){var l=[].slice,c=r.utils.isNumber,h=r.utils.isArray;r.extend({zscore:function(){var i=l.call(arguments);return c(i[1])?(i[0]-i[1])/i[2]:(i[0]-r.mean(i[1]))/r.stdev(i[1],i[2])},ztest:function(){var i=l.call(arguments),s;return h(i[1])?(s=r.zscore(i[0],i[1],i[3]),i[2]===1?r.normal.cdf(-o.abs(s),0,1):r.normal.cdf(-o.abs(s),0,1)*2):i.length>2?(s=r.zscore(i[0],i[1],i[2]),i[3]===1?r.normal.cdf(-o.abs(s),0,1):r.normal.cdf(-o.abs(s),0,1)*2):(s=i[0],i[1]===1?r.normal.cdf(-o.abs(s),0,1):r.normal.cdf(-o.abs(s),0,1)*2)}}),r.extend(r.fn,{zscore:function(i,s){return(i-this.mean())/this.stdev(s)},ztest:function(i,s,f){var a=o.abs(this.zscore(i,f));return s===1?r.normal.cdf(-a,0,1):r.normal.cdf(-a,0,1)*2}}),r.extend({tscore:function(){var i=l.call(arguments);return i.length===4?(i[0]-i[1])/(i[2]/o.sqrt(i[3])):(i[0]-r.mean(i[1]))/(r.stdev(i[1],!0)/o.sqrt(i[1].length))},ttest:function(){var i=l.call(arguments),s;return i.length===5?(s=o.abs(r.tscore(i[0],i[1],i[2],i[3])),i[4]===1?r.studentt.cdf(-s,i[3]-1):r.studentt.cdf(-s,i[3]-1)*2):c(i[1])?(s=o.abs(i[0]),i[2]==1?r.studentt.cdf(-s,i[1]-1):r.studentt.cdf(-s,i[1]-1)*2):(s=o.abs(r.tscore(i[0],i[1])),i[2]==1?r.studentt.cdf(-s,i[1].length-1):r.studentt.cdf(-s,i[1].length-1)*2)}}),r.extend(r.fn,{tscore:function(i){return(i-this.mean())/(this.stdev(!0)/o.sqrt(this.cols()))},ttest:function(i,s){return s===1?1-r.studentt.cdf(o.abs(this.tscore(i)),this.cols()-1):r.studentt.cdf(-o.abs(this.tscore(i)),this.cols()-1)*2}}),r.extend({anovafscore:function(){var i=l.call(arguments),s,f,a,d,E,m,C,v;if(i.length===1){for(E=new Array(i[0].length),C=0;C<i[0].length;C++)E[C]=i[0][C];i=E}for(f=new Array,C=0;C<i.length;C++)f=f.concat(i[C]);for(a=r.mean(f),s=0,C=0;C<i.length;C++)s=s+i[C].length*o.pow(r.mean(i[C])-a,2);for(s/=i.length-1,m=0,C=0;C<i.length;C++)for(d=r.mean(i[C]),v=0;v<i[C].length;v++)m+=o.pow(i[C][v]-d,2);return m/=f.length-i.length,s/m},anovaftest:function(){var i=l.call(arguments),s,f,a,d;if(c(i[0]))return 1-r.centralF.cdf(i[0],i[1],i[2]);var E=r.anovafscore(i);for(s=i.length-1,a=0,d=0;d<i.length;d++)a=a+i[d].length;return f=a-s-1,1-r.centralF.cdf(E,s,f)},ftest:function(i,s,f){return 1-r.centralF.cdf(i,s,f)}}),r.extend(r.fn,{anovafscore:function(){return r.anovafscore(this.toArray())},anovaftes:function(){var i=0,s;for(s=0;s<this.length;s++)i=i+this[s].length;return r.ftest(this.anovafscore(),this.length-1,i-this.length)}}),r.extend({qscore:function(){var i=l.call(arguments),s,f,a,d,E;return c(i[0])?(s=i[0],f=i[1],a=i[2],d=i[3],E=i[4]):(s=r.mean(i[0]),f=r.mean(i[1]),a=i[0].length,d=i[1].length,E=i[2]),o.abs(s-f)/(E*o.sqrt((1/a+1/d)/2))},qtest:function(){var i=l.call(arguments),s;i.length===3?(s=i[0],i=i.slice(1)):i.length===7?(s=r.qscore(i[0],i[1],i[2],i[3],i[4]),i=i.slice(5)):(s=r.qscore(i[0],i[1],i[2]),i=i.slice(3));var f=i[0],a=i[1];return 1-r.tukey.cdf(s,a,f-a)},tukeyhsd:function(i){for(var s=r.pooledstdev(i),f=i.map(function(v){return r.mean(v)}),a=i.reduce(function(v,p){return v+p.length},0),d=[],E=0;E<i.length;++E)for(var m=E+1;m<i.length;++m){var C=r.qtest(f[E],f[m],i[E].length,i[m].length,s,a,i.length);d.push([[E,m],C])}return d}}),r.extend({normalci:function(){var i=l.call(arguments),s=new Array(2),f;return i.length===4?f=o.abs(r.normal.inv(i[1]/2,0,1)*i[2]/o.sqrt(i[3])):f=o.abs(r.normal.inv(i[1]/2,0,1)*r.stdev(i[2])/o.sqrt(i[2].length)),s[0]=i[0]-f,s[1]=i[0]+f,s},tci:function(){var i=l.call(arguments),s=new Array(2),f;return i.length===4?f=o.abs(r.studentt.inv(i[1]/2,i[3]-1)*i[2]/o.sqrt(i[3])):f=o.abs(r.studentt.inv(i[1]/2,i[2].length-1)*r.stdev(i[2],!0)/o.sqrt(i[2].length)),s[0]=i[0]-f,s[1]=i[0]+f,s},significant:function(i,s){return i<s}}),r.extend(r.fn,{normalci:function(i,s){return r.normalci(i,s,this.toArray())},tci:function(i,s){return r.tci(i,s,this.toArray())}});function g(u,i,s,f){if(u>1||s>1||u<=0||s<=0)throw new Error("Proportions should be greater than 0 and less than 1");var a=(u*i+s*f)/(i+f),d=o.sqrt(a*(1-a)*(1/i+1/f));return(u-s)/d}r.extend(r.fn,{oneSidedDifferenceOfProportions:function(i,s,f,a){var d=g(i,s,f,a);return r.ztest(d,1)},twoSidedDifferenceOfProportions:function(i,s,f,a){var d=g(i,s,f,a);return r.ztest(d,2)}})})(n,Math),n.models=(function(){function r(g){var u=g[0].length,i=n.arange(u).map(function(s){var f=n.arange(u).filter(function(a){return a!==s});return o(n.col(g,s).map(function(a){return a[0]}),n.col(g,f))});return i}function o(g,u){var i=g.length,s=u[0].length-1,f=i-s-1,a=n.lstsq(u,g),d=n.multiply(u,a.map(function(A){return[A]})).map(function(A){return A[0]}),E=n.subtract(g,d),m=n.mean(g),C=n.sum(d.map(function(A){return Math.pow(A-m,2)})),v=n.sum(g.map(function(A,M){return Math.pow(A-d[M],2)})),p=C+v,I=C/p;return{exog:u,endog:g,nobs:i,df_model:s,df_resid:f,coef:a,predict:d,resid:E,ybar:m,SST:p,SSE:C,SSR:v,R2:I}}function l(g){var u=r(g.exog),i=Math.sqrt(g.SSR/g.df_resid),s=u.map(function(m){var C=m.SST,v=m.R2;return i/Math.sqrt(C*(1-v))}),f=g.coef.map(function(m,C){return(m-0)/s[C]}),a=f.map(function(m){var C=n.studentt.cdf(m,g.df_resid);return(C>.5?1-C:C)*2}),d=n.studentt.inv(.975,g.df_resid),E=g.coef.map(function(m,C){var v=d*s[C];return[m-v,m+v]});return{se:s,t:f,p:a,sigmaHat:i,interval95:E}}function c(g){var u=g.R2/g.df_model/((1-g.R2)/g.df_resid),i=function(f,a,d){return n.beta.cdf(f/(d/a+f),a/2,d/2)},s=1-i(u,g.df_model,g.df_resid);return{F_statistic:u,pvalue:s}}function h(g,u){var i=o(g,u),s=l(i),f=c(i),a=1-(1-i.R2)*((i.nobs-1)/i.df_resid);return i.t=s,i.f=f,i.adjust_R2=a,i}return{ols:h}})(),n.extend({buildxmatrix:function(){for(var o=new Array(arguments.length),l=0;l<arguments.length;l++){var c=[1];o[l]=c.concat(arguments[l])}return n(o)},builddxmatrix:function(){for(var o=new Array(arguments[0].length),l=0;l<arguments[0].length;l++){var c=[1];o[l]=c.concat(arguments[0][l])}return n(o)},buildjxmatrix:function(o){for(var l=new Array(o.length),c=0;c<o.length;c++)l[c]=o[c];return n.builddxmatrix(l)},buildymatrix:function(o){return n(o).transpose()},buildjymatrix:function(o){return o.transpose()},matrixmult:function(o,l){var c,h,g,u,i;if(o.cols()==l.rows()){if(l.rows()>1){for(u=[],c=0;c<o.rows();c++)for(u[c]=[],h=0;h<l.cols();h++){for(i=0,g=0;g<o.cols();g++)i+=o.toArray()[c][g]*l.toArray()[g][h];u[c][h]=i}return n(u)}for(u=[],c=0;c<o.rows();c++)for(u[c]=[],h=0;h<l.cols();h++){for(i=0,g=0;g<o.cols();g++)i+=o.toArray()[c][g]*l.toArray()[h];u[c][h]=i}return n(u)}},regress:function(o,l){var c=n.xtranspxinv(o),h=o.transpose(),g=n.matrixmult(n(c),h);return n.matrixmult(g,l)},regresst:function(o,l,c){var h=n.regress(o,l),g={};g.anova={};var u=n.jMatYBar(o,h);g.yBar=u;var i=l.mean();g.anova.residuals=n.residuals(l,u),g.anova.ssr=n.ssr(u,i),g.anova.msr=g.anova.ssr/(o[0].length-1),g.anova.sse=n.sse(l,u),g.anova.mse=g.anova.sse/(l.length-(o[0].length-1)-1),g.anova.sst=n.sst(l,i),g.anova.mst=g.anova.sst/(l.length-1),g.anova.r2=1-g.anova.sse/g.anova.sst,g.anova.r2<0&&(g.anova.r2=0),g.anova.fratio=g.anova.msr/g.anova.mse,g.anova.pvalue=n.anovaftest(g.anova.fratio,o[0].length-1,l.length-(o[0].length-1)-1),g.anova.rmse=Math.sqrt(g.anova.mse),g.anova.r2adj=1-g.anova.mse/g.anova.mst,g.anova.r2adj<0&&(g.anova.r2adj=0),g.stats=new Array(o[0].length);for(var s=n.xtranspxinv(o),f,a,d,E=0;E<h.length;E++)f=Math.sqrt(g.anova.mse*Math.abs(s[E][E])),a=Math.abs(h[E]/f),d=n.ttest(a,l.length-o[0].length-1,c),g.stats[E]=[h[E],f,a,d];return g.regress=h,g},xtranspx:function(o){return n.matrixmult(o.transpose(),o)},xtranspxinv:function(o){var l=n.matrixmult(o.transpose(),o),c=n.inv(l);return c},jMatYBar:function(o,l){var c=n.matrixmult(o,l);return new n(c)},residuals:function(o,l){return n.matrixsubtract(o,l)},ssr:function(o,l){for(var c=0,h=0;h<o.length;h++)c+=Math.pow(o[h]-l,2);return c},sse:function(o,l){for(var c=0,h=0;h<o.length;h++)c+=Math.pow(o[h]-l[h],2);return c},sst:function(o,l){for(var c=0,h=0;h<o.length;h++)c+=Math.pow(o[h]-l,2);return c},matrixsubtract:function(o,l){for(var c=new Array(o.length),h=0;h<o.length;h++){c[h]=new Array(o[h].length);for(var g=0;g<o[h].length;g++)c[h][g]=o[h][g]-l[h][g]}return n(c)}}),n.jStat=n,n})})(Ht)),Ht.exports}var ss=os(),F=Wr(ss),Tn={},Kr;function us(){return Kr||(Kr=1,(function(e){(function(t){t(typeof DO_NOT_EXPORT_BESSEL>"u"?e:{})})(function(t){t.version="1.0.2";var n=Math;function r(i,s){for(var f=0,a=0;f<i.length;++f)a=s*a+i[f];return a}function o(i,s,f,a,d){if(s===0)return f;if(s===1)return a;for(var E=2/i,m=a,C=1;C<s;++C)m=a*C*E+d*f,f=a,a=m;return m}function l(i,s,f,a,d){return function(m,C){if(a){if(m===0)return a==1?-1/0:1/0;if(m<0)return NaN}if(C===0)return i(m);if(C===1)return s(m);if(C<0)return NaN;C|=0;var v=i(m),p=s(m);return o(m,C,v,p,d)}}var c=(function(){var i=.636619772,s=[57568490574,-13362590354,6516196407e-1,-1121442418e-2,77392.33017,-184.9052456].reverse(),f=[57568490411,1029532985,9494680718e-3,59272.64853,267.8532712,1].reverse(),a=[1,-.001098628627,2734510407e-14,-2073370639e-15,2093887211e-16].reverse(),d=[-.01562499995,.0001430488765,-6911147651e-15,7621095161e-16,-934935152e-16].reverse();function E(A){var M=0,O=0,L=0,U=A*A;if(A<8)O=r(s,U),L=r(f,U),M=O/L;else{var H=A-.785398164;U=64/U,O=r(a,U),L=r(d,U),M=n.sqrt(i/A)*(n.cos(H)*O-n.sin(H)*L*8/A)}return M}var m=[72362614232,-7895059235,2423968531e-1,-2972611439e-3,15704.4826,-30.16036606].reverse(),C=[144725228442,2300535178,1858330474e-2,99447.43394,376.9991397,1].reverse(),v=[1,.00183105,-3516396496e-14,2457520174e-15,-240337019e-15].reverse(),p=[.04687499995,-.0002002690873,8449199096e-15,-88228987e-14,105787412e-15].reverse();function I(A){var M=0,O=0,L=0,U=A*A,H=n.abs(A)-2.356194491;return Math.abs(A)<8?(O=A*r(m,U),L=r(C,U),M=O/L):(U=64/U,O=r(v,U),L=r(p,U),M=n.sqrt(i/n.abs(A))*(n.cos(H)*O-n.sin(H)*L*8/n.abs(A)),A<0&&(M=-M)),M}return function A(M,O){if(O=Math.round(O),!isFinite(M))return isNaN(M)?M:0;if(O<0)return(O%2?-1:1)*A(M,-O);if(M<0)return(O%2?-1:1)*A(-M,O);if(O===0)return E(M);if(O===1)return I(M);if(M===0)return 0;var L=0;if(M>O)L=o(M,O,E(M),I(M),-1);else{for(var U=2*n.floor((O+n.floor(n.sqrt(40*O)))/2),H=!1,y=0,Z=0,Q=1,te=0,X=2/M,$=U;$>0;$--)te=$*X*Q-y,y=Q,Q=te,n.abs(Q)>1e10&&(Q*=1e-10,y*=1e-10,L*=1e-10,Z*=1e-10),H&&(Z+=Q),H=!H,$==O&&(L=y);Z=2*Z-Q,L/=Z}return L}})(),h=(function(){var i=.636619772,s=[-2957821389,7062834065,-5123598036e-1,1087988129e-2,-86327.92757,228.4622733].reverse(),f=[40076544269,7452499648e-1,7189466438e-3,47447.2647,226.1030244,1].reverse(),a=[1,-.001098628627,2734510407e-14,-2073370639e-15,2093887211e-16].reverse(),d=[-.01562499995,.0001430488765,-6911147651e-15,7621095161e-16,-934945152e-16].reverse();function E(A){var M=0,O=0,L=0,U=A*A,H=A-.785398164;return A<8?(O=r(s,U),L=r(f,U),M=O/L+i*c(A,0)*n.log(A)):(U=64/U,O=r(a,U),L=r(d,U),M=n.sqrt(i/A)*(n.sin(H)*O+n.cos(H)*L*8/A)),M}var m=[-4900604943e3,127527439e4,-51534381390,7349264551e-1,-4237922726e-3,8511.937935].reverse(),C=[249958057e5,424441966400,3733650367,2245904002e-2,102042.605,354.9632885,1].reverse(),v=[1,.00183105,-3516396496e-14,2457520174e-15,-240337019e-15].reverse(),p=[.04687499995,-.0002002690873,8449199096e-15,-88228987e-14,105787412e-15].reverse();function I(A){var M=0,O=0,L=0,U=A*A,H=A-2.356194491;return A<8?(O=A*r(m,U),L=r(C,U),M=O/L+i*(c(A,1)*n.log(A)-1/A)):(U=64/U,O=r(v,U),L=r(p,U),M=n.sqrt(i/A)*(n.sin(H)*O+n.cos(H)*L*8/A)),M}return l(E,I,"BESSELY",1,-1)})(),g=(function(){var i=[1,3.5156229,3.0899424,1.2067492,.2659732,.0360768,.0045813].reverse(),s=[.39894228,.01328592,.00225319,-.00157565,.00916281,-.02057706,.02635537,-.01647633,.00392377].reverse();function f(m){return m<=3.75?r(i,m*m/(3.75*3.75)):n.exp(n.abs(m))/n.sqrt(n.abs(m))*r(s,3.75/n.abs(m))}var a=[.5,.87890594,.51498869,.15084934,.02658733,.00301532,32411e-8].reverse(),d=[.39894228,-.03988024,-.00362018,.00163801,-.01031555,.02282967,-.02895312,.01787654,-.00420059].reverse();function E(m){return m<3.75?m*r(a,m*m/(3.75*3.75)):(m<0?-1:1)*n.exp(n.abs(m))/n.sqrt(n.abs(m))*r(d,3.75/n.abs(m))}return function m(C,v){if(v=Math.round(v),v===0)return f(C);if(v===1)return E(C);if(v<0)return NaN;if(n.abs(C)===0)return 0;if(C==1/0)return 1/0;var p=0,I,A=2/n.abs(C),M=0,O=1,L=0,U=2*n.round((v+n.round(n.sqrt(40*v)))/2);for(I=U;I>0;I--)L=I*A*O+M,M=O,O=L,n.abs(O)>1e10&&(O*=1e-10,M*=1e-10,p*=1e-10),I==v&&(p=M);return p*=m(C,0)/O,C<0&&v%2?-p:p}})(),u=(function(){var i=[-.57721566,.4227842,.23069756,.0348859,.00262698,1075e-7,74e-7].reverse(),s=[1.25331414,-.07832358,.02189568,-.01062446,.00587872,-.0025154,53208e-8].reverse();function f(m){return m<=2?-n.log(m/2)*g(m,0)+r(i,m*m/4):n.exp(-m)/n.sqrt(m)*r(s,2/m)}var a=[1,.15443144,-.67278579,-.18156897,-.01919402,-.00110404,-4686e-8].reverse(),d=[1.25331414,.23498619,-.0365562,.01504268,-.00780353,.00325614,-68245e-8].reverse();function E(m){return m<=2?n.log(m/2)*g(m,1)+1/m*r(a,m*m/4):n.exp(-m)/n.sqrt(m)*r(d,2/m)}return l(f,E,"BESSELK",2,1)})();t.besselj=c,t.bessely=h,t.besseli=g,t.besselk=u})})(Tn)),Tn}var fs=us(),Vt=Wr(fs);const Rn=new Error("#NULL!"),fe=new Error("#DIV/0!"),R=new Error("#VALUE!"),nt=new Error("#REF!"),kt=new Error("#NAME?"),D=new Error("#NUM!"),B=new Error("#N/A"),Mn=new Error("#ERROR!"),Xr=new Error("#GETTING_DATA"),On=new Error("#CALC!");var ls=Object.freeze({__proto__:null,calc:On,data:Xr,div0:fe,error:Mn,na:B,name:kt,nil:Rn,num:D,ref:nt,value:R});let Pe=!1;function cs(){Pe=!0}function as(){Pe=!1}function $r(e){e<60&&(e+=1);const n=Math.floor(e-25569)*86400,r=new Date(n*1e3),o=e-Math.floor(e)+1e-7;let l=Math.floor(86400*o);const c=l%60;l-=c;const h=Math.floor(l/3600),g=Math.floor(l/60)%60;let u=r.getUTCDate(),i=r.getUTCMonth();return e>=60&&e<61&&(u=29,i=1),new Date(r.getUTCFullYear(),i,u,h,g,c)}function Ue(e){const t=new Date(1900,0,1),n=e>-22038912e5?2:1;return Math.ceil((e-t)/864e5)+n}var hs=Object.freeze({__proto__:null,dateToSerial:Ue,get returnSerial(){return Pe},serialToDate:$r,useDate:as,useSerial:cs});const gs="=",ds=[">",">=","<","<=","=","<>"],xr="operator",Qr="literal",Es=[xr,Qr],Bt=xr,ze=Qr;function ye(e,t){if(Es.indexOf(t)===-1)throw new Error("Unsupported token type: "+t);return{value:e,type:t}}function ms(e){return typeof e!="string"||/^\\d+(\\.\\d+)?$/.test(e)&&(e=e.indexOf(".")===-1?parseInt(e,10):parseFloat(e)),e}function ps(e){const t=e.length,n=[];let r=0,o="",l="";for(;r<t;){const c=e.charAt(r);switch(c){case">":case"<":case"=":l=l+c,o.length>0&&(n.push(o),o="");break;default:l.length>0&&(n.push(l),l=""),o=o+c;break}r++}return o.length>0&&n.push(o),l.length>0&&n.push(l),n}function Is(e){let t="";const n=[];for(let r=0;r<e.length;r++){const o=e[r];r===0&&ds.indexOf(o)>=0?n.push(ye(o,Bt)):t+=o}return t.length>0&&n.push(ye(ms(t),ze)),n.length>0&&n[0].type!==Bt&&n.unshift(ye(gs,Bt)),n}function Ns(e){const t=[];let n;for(let r=0;r<e.length;r++){const o=e[r];switch(o.type){case Bt:n=o.value;break;case ze:t.push(o.value);break}}return As(t,n)}function As(e,t){let n=!1;switch(t){case">":n=e[0]>e[1];break;case">=":n=e[0]>=e[1];break;case"<":n=e[0]<e[1];break;case"<=":n=e[0]<=e[1];break;case"=":n=e[0]==e[1];break;case"<>":n=e[0]!=e[1];break}return n}function rt(e){return Is(ps(e))}const it=Ns;function Gt(e){const t=[];return ce(e,n=>{t.push(n)}),t}function ce(e,t){let n=-1;const r=e.length;for(;++n<r&&t(e[n],n,e)!==!1;);return e}function Dn(e){let t=e.length,n;for(;t--;)if(n=e[t],typeof n!="number"){if(n===!0){e[t]=1;continue}if(n===!1){e[t]=0;continue}if(typeof n=="string"){const r=N(n);e[t]=r instanceof Error?0:r}}return e}function pt(e){const t=e.length,n=e.reduce((r,o)=>Math.max(r,o.length),0);return[t,n]}function Cs(e,t){const n=[e,t];return n.some(r=>!he(r))?B:n.some(r=>r<=0)?R:Array.from({length:e},()=>Array.from({length:t},()=>{}))}function Jr(e,t){if(!e)return R;(!e.every(o=>Array.isArray(o))||e.length===0)&&(e=[[...e]]),e.map((o,l)=>{o.map((c,h)=>{c||(e[l][h]=0)})});const n=e.reduce((o,l,c)=>l.length>e[o].length?c:o,0),r=e[n].length;return e.map(o=>[...o,...Array(r-o.length).fill(0)])}function P(){let e;if(arguments.length===1){const t=arguments[0];e=Ts(t)?Gt.apply(null,arguments):[t]}else e=Array.from(arguments);for(;!Rs(e);)e=wn(e);return e}function wn(e){return!e||!e.reduce?[e]:e.reduce((t,n)=>{const r=Array.isArray(t),o=Array.isArray(n);return r&&o?t.concat(n):r?(t.push(n),t):o?[t].concat(n):[t,n]})}function vs(e,t){return t=t||1,!e||typeof e.slice!="function"?e:e.slice(0,e.length-t)}function Ts(e){return e!=null&&typeof e.length=="number"&&typeof e!="string"}function Rs(e){if(!e)return!1;for(let t=0;t<e.length;++t)if(Array.isArray(e[t]))return!1;return!0}function ee(e,t){return t=t||1,!e||typeof e.slice!="function"?e:e.slice(t)}function zt(e){return e?e[0].map((t,n)=>e.map(r=>r[n])):R}function Ae(e,t){let n=null;return ce(e,(r,o)=>{if(r[0]===t)return n=o,!1}),n??R}function Y(){for(let e=0;e<arguments.length;e++)if(arguments[e]instanceof Error)return arguments[e]}function w(){let e=arguments.length;for(;e--;)if(arguments[e]instanceof Error)return!0;return!1}function Zr(e){return Math.round(e*1e14)/1e14}function Fe(){return P.apply(null,arguments).filter(t=>typeof t=="number")}function Ln(e){if(typeof e=="boolean"||e instanceof Error)return e;if(typeof e=="number")return e!==0;if(typeof e=="string"){const t=e.toUpperCase();if(t==="TRUE")return!0;if(t==="FALSE")return!1}return e instanceof Date&&!isNaN(e)?!0:R}function K(e){if(!isNaN(e)){if(e instanceof Date)return new Date(e);const t=parseFloat(e);return t<0||t>=2958466?D:$r(t)}return typeof e=="string"&&(e=/(\\d{4})-(\\d\\d?)-(\\d\\d?)$/.test(e)?new Date(e+"T00:00:00.000"):new Date(e),!isNaN(e))?e:R}function br(e){let t=e.length,n;for(;t--;){if(n=K(e[t]),n===R)return n;e[t]=n}return e}function N(e){return e instanceof Error?e:e==null?0:(typeof e=="boolean"&&(e=+e),!isNaN(e)&&e!==""?parseFloat(e):R)}function V(e){let t;if(!e||(t=e.length)===0)return R;let n;for(;t--;){if(e[t]instanceof Error)return e[t];if(n=N(e[t]),n instanceof Error)return n;e[t]=n}return e}function ie(e){return e instanceof Error?e:e==null?"":e.toString()}function Yt(){let e=arguments.length;for(;e--;)if(typeof arguments[e]=="string")return!0;return!1}function Wt(){const e=Gt(arguments),t=V(P(e.shift()));if(t instanceof Error)return t;const n=e,r=n.length/2;for(let l=0;l<r;l++)n[l*2]=P(n[l*2]);let o=[];for(let l=0;l<t.length;l++){let c=!1;for(let h=0;h<r;h++){const g=n[h*2][l],u=n[h*2+1],i=u===void 0||u==="*";let s=!1;if(i)s=!0;else{const f=rt(u+""),a=[ye(g,ze)].concat(f);s=it(a)}if(!s){c=!1;break}c=!0}c&&o.push(t[l])}return o}function he(e){return e!=null}const _r={};_r.TYPE=e=>{switch(e){case Rn:return 1;case fe:return 2;case R:return 3;case nt:return 4;case kt:return 5;case D:return 6;case B:return 7;case Xr:return 8}return B};function Ms(e){return e===null}function jr(e){return[R,nt,fe,D,kt,Rn].indexOf(e)>=0||typeof e=="number"&&(isNaN(e)||!isFinite(e))}function qt(e){return jr(e)||e===B}function Os(e){return!(Math.floor(Math.abs(e))&1)}function ei(e){return e===!0||e===!1}function Ds(e){return e===B}function ws(e){return typeof e!="string"}function Kt(e){return typeof e=="number"&&!isNaN(e)&&isFinite(e)}function Ls(e){return!!(Math.floor(Math.abs(e))&1)}function ti(e){return typeof e=="string"}function Ss(e){return Kt(e)?e:e instanceof Date?e.getTime():e===!0?1:e===!1?0:qt(e)?e:0}function Ps(){return B}function Us(e){if(Kt(e))return 1;if(ti(e))return 2;if(ei(e))return 4;if(qt(e))return 16;if(Array.isArray(e))return 64}function ys(){if(arguments.length<2)return B;const e=arguments[0];return e<1||e>254||arguments.length<e+1?R:arguments[e]}function Fs(e,t,...n){if(!Array.isArray(e))return R;const r=P([t,...n]),[o,l]=pt(e);if(r.some(u=>{if(!he(u))return!0;const i=Math.abs(u);return!(0<i&&i<=l)}))return R;const h=r.map(u=>u<0?u+l+1:u),g=new Array(o);for(let u=0;u<o;u++){g[u]=new Array(h.length);for(let i=0;i<h.length;i++)g[u][i]=e[u][h[i]-1]??0}return g}function Hs(e,t,...n){if(!Array.isArray(e))return R;const r=P([t,...n]);return r.some(c=>{if(!he(c))return!0;const h=Math.abs(c);return!(0<h&&h<=e.length)})?R:r.map(c=>c<0?c+e.length+1:c).reduce((c,h)=>{const g=[...e[--h]].map(u=>he(u)?u:0);return c.push(g),c},[])}function Vs(e,t){if(arguments.length!==2)return B;if(t<0)return D;if(!(e instanceof Array)||typeof t!="number")return R;if(e.length!==0)return F.col(e,t)}function ks(e){return arguments.length!==1?B:e instanceof Array?e.length===0?0:F.cols(e):R}function Bs(e,t,n){if(w(e,t,n))return Y(e,t,n);if(!(e instanceof Array))return R;const r=[],[o,l]=pt(e);let c=0,h=o;if(typeof t=="number"){if(Math.abs(t)>o)return On;t>0?c=t:t<0&&(h+=t)}let g=0,u=l;if(typeof n=="number"){if(Math.abs(n)>l)return On;n>0?g=n:n<0&&(u+=n)}for(let i=c;i<h;i++){const s=[...e[i].slice(g,u).map(f=>f??0)];r.push(s)}return r}function Gs(e,t,n,r){const[o,l]=pt(e);if(he(t)||(t=o),he(n)||(n=l),t<o||n<l)return R;const c=Cs(t,n);for(let h=0;h<t;h++)for(let g=0;g<n;g++){let u=e[h]?.[g];he(u)||(u=h<o&&g<l?0:r??B),c[h][g]=u}return c}function zs(e,t,n,r){return ni(e,zt(t),n,r)}function Ys(e,t,n){const r=Y(e,t,n);if(r)return r;if(!Array.isArray(e))return R;const o=e.length>0&&!Array.isArray(e[0]);return o&&!n?(n=t,t=1):(n=n||1,t=t||1),n<0||t<0?R:o&&t===1&&n<=e.length?e[n-1]:t<=e.length&&n<=e[t-1].length?e[t-1][n-1]:nt}function Ws(e,t,n){t=P(t),n=n?P(n):t;const r=typeof e=="number";let o=B;for(let l=0;l<t.length;l++){if(t[l]===e)return n[l];if(r&&t[l]<=e||typeof t[l]=="string"&&t[l].localeCompare(e)<0)o=n[l];else if(r&&t[l]>e)return o}return o}function qs(e,t,n){if(!e&&e!==0||!t||(arguments.length===2&&(n=1),t=P(t),!(t instanceof Array))||n!==-1&&n!==0&&n!==1)return B;let r,o;for(let l=0;l<t.length;l++)if(n===1){if(t[l]===e)return l+1;t[l]<e&&(o?t[l]>o&&(r=l+1,o=t[l]):(r=l+1,o=t[l]))}else if(n===0){if(typeof e=="string"&&typeof t[l]=="string"){const c=e.toLowerCase().replace(/\\?/g,".").replace(/\\*/g,".*").replace(/~/g,"\\\\").replace(/\\+/g,"\\\\+").replace(/\\(/g,"\\\\(").replace(/\\)/g,"\\\\)").replace(/\\[/g,"\\\\[").replace(/\\]/g,"\\\\]");if(new RegExp("^"+c+"$").test(t[l].toLowerCase()))return l+1}else if(t[l]===e)return l+1}else if(n===-1){if(t[l]===e)return l+1;t[l]>e&&(o?t[l]<o&&(r=l+1,o=t[l]):(r=l+1,o=t[l]))}return r||B}function Ks(e){return arguments.length!==1?B:e instanceof Array?e.length===0?0:F.rows(e):R}function Xs(e,t=1,n=1,r=!1){if(!e||!Array.isArray(e))return B;if(e.length===0)return 0;if(t=N(t),!t||t<1||(n=N(n),n!==1&&n!==-1))return R;if(r=Ln(r),typeof r!="boolean")return kt;const o=h=>h.sort((g,u)=>(g=ie(g[t-1]),u=ie(u[t-1]),n===1?g<u?n*-1:n:g>u?n:n*-1)),l=Jr(e),c=r?zt(l):l;return t>=1&&t<=c[0].length?r?zt(o(c)):o(c):R}function $s(e){if(!e)return B;const t=Jr(e);return zt(t)}function Sn(){const e=[];for(let t=0;t<arguments.length;++t){let n=!1;const r=arguments[t];for(let o=0;o<e.length&&(n=e[o]===r,!n);++o);n||e.push(r)}return e}function ni(e,t,n,r){if(!t||!n)return B;r=!(r===0||r===!1);let o=B,l=!1;const c=typeof e=="number",h=typeof e=="string"?e.toLowerCase():e;for(let g=0;g<t.length;g++){const u=t[g],i=typeof u[0]=="string"?u[0].toLowerCase():u[0];if(i===h){o=n<u.length+1?u[n-1]:nt;break}else!l&&(c&&r&&i<=e||r&&typeof i=="string"&&i.localeCompare(e)<0)&&(o=n<u.length+1?u[n-1]:nt);c&&i>e&&(l=!0)}return o}function xs(e,...t){const n=[e,...t];let r=0,o=0;const l=[];for(const g of n){const[u,i]=pt(g);r=Math.max(r,u),o+=i,l.push(i)}const c=Array.from({length:r},()=>new Array(o));let h=0;for(let g=0;g<n.length;g++){const u=n[g],i=l[g];for(let s=0;s<i;s++){for(let f=0;f<r;f++)c[f][h]=u[f]?.[s]??(f<u.length?0:B);h++}}return c}function Qs(e,...t){const n=[e,...t];let r=0,o=0;for(const g of n){const[u,i]=pt(g);r=Math.max(r,i),o+=u}const l=new Array(o);let c,h=-1;for(;(c=n.shift())!==void 0;)for(let g=0;g<c.length;g++){l[++h]=new Array(r);for(let u=0;u<r;u++)l[h][u]=c[g][u]??(u<c[g].length?0:B)}return l}function ri(e){return e=N(e),e===0?R:e instanceof Error?e:String.fromCharCode(e)}function Js(e){if(w(e))return e;e=e||"";const t=/[\\0-\\x1F]/g;return e.replace(t,"")}function ii(e){if(w(e))return e;e=e||"";let t=e.charCodeAt(0);return isNaN(t)&&(t=R),t}function oi(){const e=P(arguments),t=Y.apply(void 0,e);if(t)return t;let n=0;for(;(n=e.indexOf(!0))>-1;)e[n]="TRUE";let r=0;for(;(r=e.indexOf(!1))>-1;)e[r]="FALSE";return e.join("")}const Zs=oi;function bs(e,t=2){if(e=N(e),isNaN(e))return R;e=pi(e,t);const n={style:"currency",currency:"USD",minimumFractionDigits:t>=0?t:0,maximumFractionDigits:t>=0?t:0},r=e.toLocaleString("en-US",n);return e<0?"$("+r.slice(2)+")":r}function _s(e,t){if(arguments.length!==2)return B;const n=Y(e,t);return n||(e=ie(e),t=ie(t),e===t)}function js(e,t,n){if(arguments.length<2)return B;e=ie(e),t=ie(t),n=n===void 0?0:n;const r=t.indexOf(e,n-1);return r===-1?R:r+1}function si(e,t=2,n=!1){if(e=N(e),isNaN(e)||(t=N(t),isNaN(t)))return R;if(t<0){const r=Math.pow(10,-t);e=Math.round(e/r)*r}else e=e.toFixed(t);if(n)e=e.toString().replace(/,/g,"");else{const r=e.toString().split(".");r[0]=r[0].replace(/\\B(?=(\\d{3})+$)/g,","),e=r.join(".")}return e}function eu(e,t){const n=Y(e,t);return n||(e=ie(e),t=t===void 0?1:t,t=N(t),t instanceof Error||typeof e!="string"?R:e.substring(0,t))}function tu(e){return arguments.length===0?Mn:e instanceof Error?e:Array.isArray(e)?R:ie(e).length}function nu(e){return arguments.length!==1?R:(e=ie(e),w(e)?e:e.toLowerCase())}function ru(e,t,n){if(w(e,t,n))return Y(e,t,n);const r=[t,n].map(h=>N(h));if(!r.every(h=>!(h instanceof Error)&&h>0))return R;he(e)?typeof e!="string"&&(e=String(e)):e="",[t,n]=r;const l=t-1,c=l+n;return e.substring(l,c)}function iu(e,t,n){return e=he(e)?e:"",typeof e=="number"?e:typeof e!="string"?B:(t=typeof t>"u"?".":t,n=typeof n>"u"?",":n,Number(e.replace(t,".").replace(n,"")))}function ou(e){return w(e)?e:isNaN(e)&&typeof e=="number"?R:(e=ie(e),e.replace(/\\w\\S*/g,t=>t.charAt(0).toUpperCase()+t.substr(1).toLowerCase()))}function su(e,t,n,r){return t=N(t),n=N(n),w(t,n)||typeof e!="string"||typeof r!="string"?R:e.substr(0,t-1)+r+e.substr(t-1+n)}function me(e,t){const n=Y(e,t);return n||(e=ie(e),t=N(t),t instanceof Error?t:new Array(t+1).join(e))}function uu(e,t){const n=Y(e,t);return n||(e=ie(e),t=t===void 0?1:t,t=N(t),t instanceof Error?t:e.substring(e.length-t))}function fu(e,t,n){let r;return typeof e!="string"||typeof t!="string"?R:(n=n===void 0?0:n,r=t.toLowerCase().indexOf(e.toLowerCase(),n-1)+1,r===0?R:r)}function lu(e,t,n,r){if(arguments.length<3)return B;if(!e||!t)return e;if(r===void 0)return e.split(t).join(n);{if(r=Math.floor(Number(r)),Number.isNaN(r)||r<=0)return R;let o=0,l=0;for(;o>-1&&e.indexOf(t,o)>-1;)if(o=e.indexOf(t,o+1),l++,o>-1&&l===r)return e.substring(0,o)+n+e.substring(o+t.length);return e}}function de(e){return e instanceof Error||typeof e=="string"?e:""}function cu(e,t){if(e===void 0||e instanceof Error||t instanceof Error)return B;if(e instanceof Date)return e.toISOString().slice(0,10);if(t==null)return"";if(typeof t=="number")return String(t);if(typeof t!="string")return R;const n=t.startsWith("$")?"$":"",r=t.endsWith("%");t=t.replace(/%/g,"").replace(/\\$/g,"");const o=t.includes(".")?t.split(".")[1].match(/0/g).length:0,l=!t.includes(",");return r&&(e=e*100),e=si(e,o,l),e.startsWith("-")?(e=e.replace("-",""),e="-"+n+e):e=n+e,r&&(e=e+"%"),e}function au(e,t,...n){if(typeof t!="boolean"&&(t=Ln(t)),arguments.length<3)return B;e=e??"";let r=P(n),o=t?r.filter(l=>l):r;if(Array.isArray(e)){e=P(e);let l=o.map(h=>[h]),c=0;for(let h=0;h<l.length-1;h++)l[h].push(e[c]),c++,c===e.length&&(c=0);return o=P(l),o.join("")}return o.join(e)}function hu(e){return e=ie(e),e instanceof Error?e:e.replace(/\\s+/g," ").trim()}const gu=ri,du=ii;function Eu(e){return e=ie(e),e instanceof Error?e:e.toUpperCase()}function mu(e){const t=Y(e);if(t)return t;if(typeof e=="number")return e;if(he(e)||(e=""),typeof e!="string")return R;const n=/(%)$/.test(e)||/^(%)/.test(e);if(e=e.replace(/^[^0-9-]{0,3}/,""),e=e.replace(/[^0-9]{0,3}$/,""),e=e.replace(/[ ,]/g,""),e==="")return 0;let r=Number(e);return isNaN(r)?R:(r=r||0,n&&(r=r*.01),r)}const pu=2.5066282746310002;function Iu(){const t=P(arguments).filter(he);if(t.length===0)return D;const n=V(t);return n instanceof Error?n:F.sum(F(n).subtract(F.mean(n)).abs()[0])/n.length}function $e(){const t=P(arguments).filter(he);if(t.length===0)return fe;const n=Y.apply(void 0,t);if(n)return n;const r=Fe(t),o=r.length;let l=0,c=0,h;for(let g=0;g<o;g++)l+=r[g],c+=1;return h=l/c,isNaN(h)&&(h=D),h}function Pn(){const t=P(arguments).filter(he);if(t.length===0)return fe;const n=Y.apply(void 0,t);if(n)return n;const r=t,o=r.length;let l=0,c=0,h;for(let g=0;g<o;g++){const u=r[g];typeof u=="number"&&(l+=u),u===!0&&l++,u!==null&&c++}return h=l/c,isNaN(h)&&(h=D),h}function Nu(e,t,n){if(arguments.length<=1)return B;n=n||e;const o=P(n).filter(he);if(n=V(o),e=P(e),n instanceof Error)return n;let l=0,c=0;const h=t===void 0||t==="*",g=h?null:rt(t+"");for(let u=0;u<e.length;u++){const i=e[u];if(h)c+=n[u],l++;else{const s=[ye(i,ze)].concat(g);it(s)&&(c+=n[u],l++)}}return c/l}function Au(){const e=Wt(...arguments),n=e.reduce((r,o)=>r+o,0)/e.length;return isNaN(n)?0:n}const It={};It.DIST=function(e,t,n,r,o,l){return arguments.length<4||(o=o===void 0?0:o,l=l===void 0?1:l,e=N(e),t=N(t),n=N(n),o=N(o),l=N(l),w(e,t,n,o,l))?R:(e=(e-o)/(l-o),r?F.beta.cdf(e,t,n):F.beta.pdf(e,t,n))},It.INV=(e,t,n,r,o)=>(r=r===void 0?0:r,o=o===void 0?1:o,e=N(e),t=N(t),n=N(n),r=N(r),o=N(o),w(e,t,n,r,o)?R:F.beta.inv(e,t,n)*(o-r)+r);const ot={};ot.DIST=(e,t,n,r)=>(e=N(e),t=N(t),n=N(n),r=N(r),w(e,t,n,r)?R:r?F.binomial.cdf(e,t,n):F.binomial.pdf(e,t,n)),ot.DIST.RANGE=(e,t,n,r)=>{if(r=r===void 0?n:r,e=N(e),t=N(t),n=N(n),r=N(r),w(e,t,n,r))return R;let o=0;for(let l=n;l<=r;l++)o+=ft(e,l)*Math.pow(t,l)*Math.pow(1-t,e-l);return o},ot.INV=(e,t,n)=>{if(e=N(e),t=N(t),n=N(n),w(e,t,n))return R;let r=0;for(;r<=e;){if(F.binomial.cdf(r,e,t)>=n)return r;r++}};const Re={};Re.DIST=(e,t,n)=>(e=N(e),t=N(t),w(e,t)?R:n?F.chisquare.cdf(e,t):F.chisquare.pdf(e,t)),Re.DIST.RT=(e,t)=>!e|!t?B:e<1||t>Math.pow(10,10)?D:typeof e!="number"||typeof t!="number"?R:1-F.chisquare.cdf(e,t),Re.INV=(e,t)=>(e=N(e),t=N(t),w(e,t)?R:F.chisquare.inv(e,t)),Re.INV.RT=(e,t)=>!e|!t?B:e<0||e>1||t<1||t>Math.pow(10,10)?D:typeof e!="number"||typeof t!="number"?R:F.chisquare.inv(1-e,t),Re.TEST=function(e,t){if(arguments.length!==2)return B;if(!(e instanceof Array)||!(t instanceof Array)||e.length!==t.length||e[0]&&t[0]&&e[0].length!==t[0].length)return R;const n=e.length;let r,o,l;for(o=0;o<n;o++)e[o]instanceof Array||(r=e[o],e[o]=[],e[o].push(r)),t[o]instanceof Array||(r=t[o],t[o]=[],t[o].push(r));const c=e[0].length,h=c===1?n-1:(n-1)*(c-1);let g=0;const u=Math.PI;for(o=0;o<n;o++)for(l=0;l<c;l++)g+=Math.pow(e[o][l]-t[o][l],2)/t[o][l];function i(s,f){let a=Math.exp(-.5*s);f%2===1&&(a=a*Math.sqrt(2*s/u));let d=f;for(;d>=2;)a=a*s/d,d=d-2;let E=a,m=f;for(;E>1e-10*a;)m=m+2,E=E*s/m,a=a+E;return 1-a}return Math.round(i(g,h)*1e6)/1e6};const Un={};Un.NORM=(e,t,n)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:F.normalci(1,e,t,n)[1]-1),Un.T=(e,t,n)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:F.tci(1,e,t,n)[1]-1);function Cu(e,t){return e=V(P(e)),t=V(P(t)),w(e,t)?R:F.corrcoeff(e,t)}function Nt(){const e=P(arguments);return Fe(e).length}function At(){const e=P(arguments);return e.length-ui(e)}function ui(){const e=P(arguments);let t=0,n;for(let r=0;r<e.length;r++)n=e[r],(n==null||n==="")&&t++;return t}function vu(e,t){if(e=P(e),t===void 0||t==="*")return e.length;let r=0;const o=rt(t+"");for(let l=0;l<e.length;l++){const c=e[l],h=[ye(c,ze)].concat(o);it(h)&&r++}return r}function Tu(){const e=Gt(arguments),t=new Array(P(e[0]).length);for(let r=0;r<t.length;r++)t[r]=!0;for(let r=0;r<e.length;r+=2){const o=P(e[r]),l=e[r+1];if(!(l===void 0||l==="*")){const h=rt(l+"");for(let g=0;g<o.length;g++){const u=o[g],i=[ye(u,ze)].concat(h);t[g]=t[g]&&it(i)}}}let n=0;for(let r=0;r<t.length;r++)t[r]&&n++;return n}const st={};st.P=(e,t)=>{if(e=V(P(e)),t=V(P(t)),w(e,t))return R;const n=F.mean(e),r=F.mean(t);let o=0;const l=e.length;for(let c=0;c<l;c++)o+=(e[c]-n)*(t[c]-r);return o/l},st.S=(e,t)=>(e=V(P(e)),t=V(P(t)),w(e,t)?R:F.covariance(e,t));function Ru(){const e=V(P(arguments));if(e instanceof Error)return e;const t=F.mean(e);let n=0;for(let r=0;r<e.length;r++)n+=Math.pow(e[r]-t,2);return n}const yn={};yn.DIST=(e,t,n)=>(e=N(e),t=N(t),w(e,t)?R:n?F.exponential.cdf(e,t):F.exponential.pdf(e,t));const Me={};Me.DIST=(e,t,n,r)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:r?F.centralF.cdf(e,t,n):F.centralF.pdf(e,t,n)),Me.DIST.RT=function(e,t,n){return arguments.length!==3?B:e<0||t<1||n<1?D:typeof e!="number"||typeof t!="number"||typeof n!="number"?R:1-F.centralF.cdf(e,t,n)},Me.INV=(e,t,n)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:e<=0||e>1?D:F.centralF.inv(e,t,n)),Me.INV.RT=function(e,t,n){return arguments.length!==3?B:e<0||e>1||t<1||t>Math.pow(10,10)||n<1||n>Math.pow(10,10)?D:typeof e!="number"||typeof t!="number"||typeof n!="number"?R:F.centralF.inv(1-e,t,n)},Me.TEST=(e,t)=>{if(!e||!t||!(e instanceof Array)||!(t instanceof Array))return B;if(e.length<2||t.length<2)return fe;const n=(h,g)=>{let u=0;for(let i=0;i<h.length;i++)u+=Math.pow(h[i]-g,2);return u},r=Ve(e)/e.length,o=Ve(t)/t.length,l=n(e,r)/(e.length-1),c=n(t,o)/(t.length-1);return l/c};function Mu(e){return e=N(e),e instanceof Error?e:Math.log((1+e)/(1-e))/2}function Ou(e){if(e=N(e),e instanceof Error)return e;const t=Math.exp(2*e);return(t-1)/(t+1)}function fi(e,t,n){if(e=N(e),t=V(P(t)),n=V(P(n)),w(e,t,n))return R;const r=F.mean(n),o=F.mean(t),l=n.length;let c=0,h=0;for(let i=0;i<l;i++)c+=(n[i]-r)*(t[i]-o),h+=Math.pow(n[i]-r,2);const g=c/h;return o-g*r+g*e}function Du(e,t){if(e=V(P(e)),t=V(P(t)),w(e,t))return R;const n=e.length,r=t.length,o=[];for(let l=0;l<=r;l++){o[l]=0;for(let c=0;c<n;c++)l===0?e[c]<=t[0]&&(o[0]+=1):l<r?e[c]>t[l-1]&&e[c]<=t[l]&&(o[l]+=1):l===r&&e[c]>t[r-1]&&(o[r]+=1)}return o}function Ct(e){return e=N(e),e instanceof Error?e:e===0||parseInt(e,10)===e&&e<0?D:F.gammafn(e)}Ct.DIST=function(e,t,n,r){return arguments.length!==4?B:e<0||t<=0||n<=0||typeof e!="number"||typeof t!="number"||typeof n!="number"?R:r?F.gamma.cdf(e,t,n,!0):F.gamma.pdf(e,t,n,!1)},Ct.INV=function(e,t,n){return arguments.length!==3?B:e<0||e>1||t<=0||n<=0?D:typeof e!="number"||typeof t!="number"||typeof n!="number"?R:F.gamma.inv(e,t,n)};function Fn(e){return e=N(e),e instanceof Error?e:F.gammaln(e)}Fn.PRECISE=function(e){return arguments.length!==1?B:e<=0?D:typeof e!="number"?R:F.gammaln(e)};function wu(e){return e=N(e),e instanceof Error?e:F.normal.cdf(e,0,1)-.5}function Lu(){const e=V(P(arguments));return e instanceof Error?e:F.geomean(e)}function Su(e,t,n,r){if(e=V(P(e)),e instanceof Error)return e;let o;if(t===void 0)for(t=[],o=1;o<=e.length;o++)t.push(o);if(n===void 0&&(n=t),t=V(P(t)),n=V(P(n)),w(t,n))return R;r===void 0&&(r=!0);const l=e.length;let c=0,h=0,g=0,u=0;for(o=0;o<l;o++){const a=t[o],d=Math.log(e[o]);c+=a,h+=d,g+=a*d,u+=a*a}c/=l,h/=l,g/=l,u/=l;let i,s;r?(i=(g-c*h)/(u-c*c),s=h-i*c):(i=g/u,s=0);const f=[];for(o=0;o<n.length;o++)f.push(Math.exp(s+i*n[o]));return f}function Pu(){const e=V(P(arguments));if(e instanceof Error)return e;const t=e.length;let n=0;for(let r=0;r<t;r++)n+=1/e[r];return t/n}const Hn={};Hn.DIST=(e,t,n,r,o)=>{if(e=N(e),t=N(t),n=N(n),r=N(r),w(e,t,n,r))return R;function l(h,g,u,i){return ft(u,h)*ft(i-u,g-h)/ft(i,g)}function c(h,g,u,i){let s=0;for(let f=0;f<=h;f++)s+=l(f,g,u,i);return s}return o?c(e,t,n,r):l(e,t,n,r)};function Uu(e,t){return e=V(e),t=V(t),w(e,t)?R:e.length!==t.length?B:fi(0,e,t)}function yu(){const e=V(P(arguments));if(e instanceof Error)return e;const t=F.mean(e),n=e.length;let r=0;for(let o=0;o<n;o++)r+=Math.pow(e[o]-t,4);return r=r/Math.pow(F.stdev(e,!0),4),n*(n+1)/((n-1)*(n-2)*(n-3))*r-3*(n-1)*(n-1)/((n-2)*(n-3))}function li(e,t){const n=Y.apply(void 0,e);return n||(w(t)?t:(e=Fe(P(e)),t=N(t),t<0||e.length<t?R:e.sort((r,o)=>o-r)[t-1]))}function Vn(e,t){if(e=V(P(e)),t=V(P(t)),w(e,t))return R;const n=F.mean(e),r=F.mean(t),o=t.length;let l=0,c=0;for(let u=0;u<o;u++)l+=(t[u]-r)*(e[u]-n),c+=Math.pow(t[u]-r,2);const h=l/c,g=n-h*r;return[h,g]}function Fu(e,t){if(e=V(P(e)),t=V(P(t)),w(e,t)||e.length!==t.length)return R;for(let r=0;r<e.length;r++)e[r]=Math.log(e[r]);const n=Vn(e,t);return n[0]=Math.round(Math.exp(n[0])*1e6)/1e6,n[1]=Math.round(Math.exp(n[1])*1e6)/1e6,n}const ut={};ut.DIST=(e,t,n,r)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:r?F.lognormal.cdf(e,t,n):F.lognormal.pdf(e,t,n)),ut.INV=(e,t,n)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:F.lognormal.inv(e,t,n));function Xt(){const e=P(arguments),t=Y.apply(void 0,e);if(t)return t;const n=Fe(e);return n.length===0?0:Math.max.apply(Math,n)}function Hu(){const e=P(arguments),t=Y.apply(void 0,e);if(t)return t;let n=Dn(e);return n=n.map(r=>r??0),n.length===0?0:Math.max.apply(Math,n)}function Vu(){const e=Wt(...arguments);return e.length===0?0:Math.max.apply(Math,e)}function ci(){const e=P(arguments),t=Y.apply(void 0,e);if(t)return t;const n=Dn(e);let r=F.median(n);return isNaN(r)&&(r=D),r}function $t(){const e=P(arguments),t=Y.apply(void 0,e);if(t)return t;const n=Fe(e);return n.length===0?0:Math.min.apply(Math,n)}function ku(){const e=P(arguments),t=Y.apply(void 0,e);if(t)return t;let n=Dn(e);return n=n.map(r=>r??0),n.length===0?0:Math.min.apply(Math,n)}function Bu(){const e=Wt(...arguments);return e.length===0?0:Math.min.apply(Math,e)}const xe={};xe.MULT=function(){const e=V(P(arguments));if(e instanceof Error)return e;const t=e.length,n={};let r=[],o=0,l;for(let c=0;c<t;c++)l=e[c],n[l]=n[l]?n[l]+1:1,n[l]>o&&(o=n[l],r=[]),n[l]===o&&(r[r.length]=l);return r},xe.SNGL=function(){const e=V(P(arguments));return e instanceof Error?e:xe.MULT(e).sort((t,n)=>t-n)[0]};const kn={};kn.DIST=(e,t,n,r)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:r?F.negbin.cdf(e,t,n):F.negbin.pdf(e,t,n));const Oe={};Oe.DIST=(e,t,n,r)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:n<=0?D:r?F.normal.cdf(e,t,n):F.normal.pdf(e,t,n)),Oe.INV=(e,t,n)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:F.normal.inv(e,t,n)),Oe.S={},Oe.S.DIST=(e,t)=>(e=N(e),e instanceof Error?R:t?F.normal.cdf(e,0,1):F.normal.pdf(e,0,1)),Oe.S.INV=e=>(e=N(e),e instanceof Error?R:F.normal.inv(e,0,1));function ai(e,t){if(t=V(P(t)),e=V(P(e)),w(t,e))return R;const n=F.mean(e),r=F.mean(t),o=e.length;let l=0,c=0,h=0;for(let g=0;g<o;g++)l+=(e[g]-n)*(t[g]-r),c+=Math.pow(e[g]-n,2),h+=Math.pow(t[g]-r,2);return l/Math.sqrt(c*h)}const pe={};pe.EXC=(e,t)=>{if(e=V(P(e)),t=N(t),w(e,t))return R;e=e.sort((l,c)=>l-c);const n=e.length;if(t<1/(n+1)||t>1-1/(n+1))return D;const r=t*(n+1)-1,o=Math.floor(r);return Zr(r===o?e[r]:e[o]+(r-o)*(e[o+1]-e[o]))},pe.INC=(e,t)=>{if(e=V(P(e)),t=N(t),w(e,t))return R;e=e.sort((l,c)=>l-c);const n=e.length,r=t*(n-1),o=Math.floor(r);return Zr(r===o?e[r]:e[o]+(r-o)*(e[o+1]-e[o]))};const vt={};vt.EXC=(e,t,n)=>{if(n=n===void 0?3:n,e=V(P(e)),t=N(t),n=N(n),w(e,t,n))return R;e=e.sort((i,s)=>i-s);const r=Sn.apply(null,e),o=e.length,l=r.length,c=Math.pow(10,n);let h=0,g=!1,u=0;for(;!g&&u<l;)t===r[u]?(h=(e.indexOf(r[u])+1)/(o+1),g=!0):t>=r[u]&&(t<r[u+1]||u===l-1)&&(h=(e.indexOf(r[u])+1+(t-r[u])/(r[u+1]-r[u]))/(o+1),g=!0),u++;return Math.floor(h*c)/c},vt.INC=(e,t,n)=>{if(n=n===void 0?3:n,e=V(P(e)),t=N(t),n=N(n),w(e,t,n))return R;e=e.sort((i,s)=>i-s);const r=Sn.apply(null,e),o=e.length,l=r.length,c=Math.pow(10,n);let h=0,g=!1,u=0;for(;!g&&u<l;)t===r[u]?(h=e.indexOf(r[u])/(o-1),g=!0):t>=r[u]&&(t<r[u+1]||u===l-1)&&(h=(e.indexOf(r[u])+(t-r[u])/(r[u+1]-r[u]))/(o-1),g=!0),u++;return Math.floor(h*c)/c};function Gu(e,t){return e=N(e),t=N(t),w(e,t)?R:He(e)/He(e-t)}function zu(e,t){return e=N(e),t=N(t),w(e,t)?R:Math.pow(e,t)}function Yu(e){return e=N(e),e instanceof Error?R:Math.exp(-.5*e*e)/pu}const Bn={};Bn.DIST=(e,t,n)=>(e=N(e),t=N(t),w(e,t)?R:n?F.poisson.cdf(e,t):F.poisson.pdf(e,t));function Wu(e,t,n,r){if(n===void 0)return 0;if(r=r===void 0?n:r,e=V(P(e)),t=V(P(t)),n=N(n),r=N(r),w(e,t,n,r))return R;if(n===r)return e.indexOf(n)>=0?t[e.indexOf(n)]:0;const o=e.sort((h,g)=>h-g),l=o.length;let c=0;for(let h=0;h<l;h++)o[h]>=n&&o[h]<=r&&(c+=t[e.indexOf(o[h])]);return c}const Qe={};Qe.EXC=(e,t)=>{if(e=V(Fe(P(e))),t=N(t),w(e,t))return R;switch(t){case 1:return pe.EXC(e,.25);case 2:return pe.EXC(e,.5);case 3:return pe.EXC(e,.75);default:return D}},Qe.INC=(e,t)=>{if(e=V(Fe(P(e))),t=N(t),w(e,t))return R;switch(t){case 1:return pe.INC(e,.25);case 2:return pe.INC(e,.5);case 3:return pe.INC(e,.75);default:return D}};const Tt={};Tt.AVG=(e,t,n)=>{if(e=N(e),t=V(P(t)),w(e,t))return R;t=P(t),n=n||!1;const r=n?(c,h)=>c-h:(c,h)=>h-c;t=t.sort(r);const o=t.length;let l=0;for(let c=0;c<o;c++)t[c]===e&&l++;return l>1?(2*t.indexOf(e)+l+1)/2:t.indexOf(e)+1},Tt.EQ=(e,t,n)=>{if(e=N(e),t=V(P(t)),w(e,t))return R;n=n||!1;const r=n?(o,l)=>o-l:(o,l)=>l-o;return t=t.sort(r),t.indexOf(e)+1};function qu(e,t){if(arguments.length!==2)return B;if(t<0)return D;if(!(e instanceof Array)||typeof t!="number")return R;if(e.length!==0)return F.row(e,t)}function Ku(e,t){return e=V(P(e)),t=V(P(t)),w(e,t)?R:Math.pow(ai(e,t),2)}function Gn(){const e=V(P(arguments));if(e instanceof Error)return e;const t=F.mean(e),n=e.length;let r=0;for(let o=0;o<n;o++)r+=Math.pow(e[o]-t,3);return n*r/((n-1)*(n-2)*Math.pow(F.stdev(e,!0),3))}Gn.P=function(){const e=V(P(arguments));if(e instanceof Error)return e;const t=F.mean(e),n=e.length;let r=0,o=0;for(let l=0;l<n;l++)o+=Math.pow(e[l]-t,3),r+=Math.pow(e[l]-t,2);return o=o/n,r=r/n,o/Math.pow(r,3/2)};function Xu(e,t){if(e=V(P(e)),t=V(P(t)),w(e,t))return R;const n=F.mean(t),r=F.mean(e),o=t.length;let l=0,c=0;for(let h=0;h<o;h++)l+=(t[h]-n)*(e[h]-r),c+=Math.pow(t[h]-n,2);return l/c}function hi(e,t){return e=V(P(e)),t=N(t),w(e,t)?e:e.sort((n,r)=>n-r)[t-1]}function $u(e,t,n){return e=N(e),t=N(t),n=N(n),w(e,t,n)?R:(e-t)/n}const Ee={};Ee.P=function(){const e=ge.P.apply(this,arguments);let t=Math.sqrt(e);return isNaN(t)&&(t=D),t},Ee.S=function(){const e=ge.S.apply(this,arguments);return Math.sqrt(e)};function xu(){const e=gi.apply(this,arguments);return Math.sqrt(e)}function Qu(){const e=di.apply(this,arguments);let t=Math.sqrt(e);return isNaN(t)&&(t=D),t}function Ju(e,t){if(e=V(P(e)),t=V(P(t)),w(e,t))return R;const n=F.mean(t),r=F.mean(e),o=t.length;let l=0,c=0,h=0;for(let g=0;g<o;g++)l+=Math.pow(e[g]-r,2),c+=(t[g]-n)*(e[g]-r),h+=Math.pow(t[g]-n,2);return Math.sqrt((l-c*c/h)/(o-2))}de.DIST=(e,t,n)=>n!==1&&n!==2?D:n===1?de.DIST.RT(e,t):de.DIST["2T"](e,t),de.DIST["2T"]=function(e,t){return arguments.length!==2?B:e<0||t<1?D:typeof e!="number"||typeof t!="number"?R:(1-F.studentt.cdf(e,t))*2},de.DIST.RT=function(e,t){return arguments.length!==2?B:e<0||t<1?D:typeof e!="number"||typeof t!="number"?R:1-F.studentt.cdf(e,t)},de.INV=(e,t)=>(e=N(e),t=N(t),w(e,t)?R:F.studentt.inv(e,t)),de.INV["2T"]=(e,t)=>(e=N(e),t=N(t),e<=0||e>1||t<1?D:w(e,t)?R:Math.abs(F.studentt.inv(e/2,t))),de.TEST=(e,t)=>{if(e=V(P(e)),t=V(P(t)),w(e,t))return R;const n=F.mean(e),r=F.mean(t);let o=0,l=0,c;for(c=0;c<e.length;c++)o+=Math.pow(e[c]-n,2);for(c=0;c<t.length;c++)l+=Math.pow(t[c]-r,2);o=o/(e.length-1),l=l/(t.length-1);const h=Math.abs(n-r)/Math.sqrt(o/e.length+l/t.length);return de.DIST["2T"](h,e.length+t.length-2)};function Zu(e,t,n){if(e=V(P(e)),t=V(P(t)),n=V(P(n)),w(e,t,n))return R;const r=Vn(e,t),o=r[0],l=r[1],c=[];return n.forEach(h=>{c.push(o*h+l)}),c}function bu(e,t){if(e=V(P(e)),t=N(t),w(e,t))return R;const n=Je(e.length*t,2)/2;return F.mean(vs(ee(e.sort((r,o)=>r-o),n),n))}const ge={};ge.P=function(){const e=Fe(P(arguments)),t=e.length;let n=0;const r=$e(e);let o;for(let l=0;l<t;l++)n+=Math.pow(e[l]-r,2);return o=n/t,isNaN(o)&&(o=D),o},ge.S=function(){const e=Fe(P(arguments)),t=e.length;let n=0;const r=$e(e);for(let o=0;o<t;o++)n+=Math.pow(e[o]-r,2);return n/(t-1)};function gi(){const e=P(arguments),t=e.length;let n=0,r=0;const o=Pn(e);for(let l=0;l<t;l++){const c=e[l];typeof c=="number"?n+=Math.pow(c-o,2):c===!0?n+=Math.pow(1-o,2):n+=Math.pow(0-o,2),c!==null&&r++}return n/(r-1)}function di(){const e=P(arguments),t=e.length;let n=0,r=0;const o=Pn(e);let l;for(let c=0;c<t;c++){const h=e[c];typeof h=="number"?n+=Math.pow(h-o,2):h===!0?n+=Math.pow(1-o,2):n+=Math.pow(0-o,2),h!==null&&r++}return l=n/r,isNaN(l)&&(l=D),l}const zn={};zn.DIST=(e,t,n,r)=>(e=N(e),t=N(t),n=N(n),w(e,t,n)?R:r?1-Math.exp(-Math.pow(e/n,t)):Math.pow(e,t-1)*Math.exp(-Math.pow(e/n,t))*t/Math.pow(n,t));const Yn={};Yn.TEST=(e,t,n)=>{if(e=V(P(e)),t=N(t),w(e,t))return R;n=n||Ee.S(e);const r=e.length;return 1-Oe.S.DIST(($e(e)-t)/(n/Math.sqrt(r)),!0)};function _u(e){return e=N(e),e instanceof Error?e:Math.abs(e)}function ju(e){if(e=N(e),e instanceof Error)return e;let t=Math.acos(e);return isNaN(t)&&(t=D),t}function ef(e){if(e=N(e),e instanceof Error)return e;let t=Math.log(e+Math.sqrt(e*e-1));return isNaN(t)&&(t=D),t}function tf(e){return e=N(e),e instanceof Error?e:Math.atan(1/e)}function nf(e){if(e=N(e),e instanceof Error)return e;let t=.5*Math.log((e+1)/(e-1));return isNaN(t)&&(t=D),t}function rf(e,t,n,r){if(e=N(e),t=N(e),w(e,t))return R;switch(e){case 1:return $e(n);case 2:return Nt(n);case 3:return At(n);case 4:return Xt(n);case 5:return $t(n);case 6:return Qt(n);case 7:return Ee.S(n);case 8:return Ee.P(n);case 9:return Ve(n);case 10:return ge.S(n);case 11:return ge.P(n);case 12:return ci(n);case 13:return xe.SNGL(n);case 14:return li(n,r);case 15:return hi(n,r);case 16:return pe.INC(n,r);case 17:return Qe.INC(n,r);case 18:return pe.EXC(n,r);case 19:return Qe.EXC(n,r)}}function of(e){if(e==null)return 0;if(e instanceof Error)return e;if(!/^M*(?:D?C{0,3}|C[MD])(?:L?X{0,3}|X[CL])(?:V?I{0,3}|I[XV])$/.test(e))return R;let t=0;return e.replace(/[MDLV]|C[MD]?|X[CL]?|I[XV]?/g,n=>{t+={M:1e3,CM:900,D:500,CD:400,C:100,XC:90,L:50,XL:40,X:10,IX:9,V:5,IV:4,I:1}[n]}),t}function sf(e){if(e=N(e),e instanceof Error)return e;let t=Math.asin(e);return isNaN(t)&&(t=D),t}function uf(e){return e=N(e),e instanceof Error?e:Math.log(e+Math.sqrt(e*e+1))}function ff(e){return e=N(e),e instanceof Error?e:Math.atan(e)}function lf(e,t){e=N(e),t=N(t);const n=Y(e,t);return n||Math.atan2(e,t)}function cf(e){if(e=N(e),e instanceof Error)return e;let t=Math.log((1+e)/(1-e))/2;return isNaN(t)&&(t=D),t}function af(e,t,n){e=N(e),t=N(t),n=N(n);const r=Y(e,t,n);if(r)return r;if(t===0)return D;const o=e.toString(t);return new Array(Math.max(n+1-o.length,0)).join("0")+o}function Ye(e,t){e=N(e),t=N(t);const n=Y(e,t);return n||(t===0?0:e>0&&t<0?D:Math.ceil(e/t)*t)}Ye.MATH=(e,t,n=0)=>{t===void 0&&(t=e>0?1:-1),e=N(e),t=N(t),n=N(n);const r=Y(e,t,n);return r||(t===0?0:(t=Math.abs(t),n===0||e>0?Math.ceil(e/t)*t:Math.floor(e/t)*t))},Ye.PRECISE=(e,t)=>Ye.MATH(e,t);function ft(e,t){e=N(e),t=N(t);const n=Y(e,t);return n||(e<t?D:He(e)/(He(t)*He(e-t)))}function hf(e,t){e=N(e),t=N(t);const n=Y(e,t);return n||(e<t?D:e===0&&t===0?1:ft(e+t-1,e-1))}function gf(e){return e=N(e),e instanceof Error?e:Math.cos(e)}function df(e){return e=N(e),e instanceof Error?e:(Math.exp(e)+Math.exp(-e))/2}function Ef(e){return e=N(e),e instanceof Error?e:e===0?fe:1/Math.tan(e)}function mf(e){if(e=N(e),e instanceof Error)return e;if(e===0)return fe;const t=Math.exp(2*e);return(t+1)/(t-1)}function pf(e){return e=N(e),e instanceof Error?e:e===0?fe:1/Math.sin(e)}function If(e){return e=N(e),e instanceof Error?e:e===0?fe:2/(Math.exp(e)-Math.exp(-e))}function Nf(e,t){if(arguments.length<2)return B;e=e||"0",t=N(t);const n=Y(e,t);if(n)return n;if(t===0)return D;const r=parseInt(e,t);return isNaN(r)?D:r}function Af(e){return e=N(e),e instanceof Error?e:e*180/Math.PI}function Cf(e){return e=N(e),e instanceof Error?e:Ye.MATH(e,-2,-1)}function vf(e){return arguments.length<1?B:arguments.length>1?Mn:(e=N(e),e instanceof Error||(e=Math.exp(e)),e)}const xt=[];function He(e){if(e=N(e),e instanceof Error)return e;const t=Math.floor(e);return t===0||t===1?1:(xt[t]>0||(xt[t]=He(t-1)*t),xt[t])}function Ei(e){if(e=N(e),e instanceof Error)return e;const t=Math.floor(e);return t<=0?1:t*Ei(t-2)}function Je(e,t){e=N(e),t=N(t);const n=Y(e,t);return n||(t?e>0&&t<0?D:Math.floor(e/t)*t:fe)}Je.MATH=(e,t=1,n=0)=>{e=N(e),t=N(t),n=N(n);const r=Y(e,t,n);return r||(t===0?0:(t=Math.abs(t),n===0||e>0?Math.floor(e/t)*t:Math.ceil(e/t)*t))},Je.PRECISE=(e,t)=>Je.MATH(e,t);function Tf(){const e=V(P(arguments));if(e instanceof Error)return e;const t=e.length,n=e[0];let r=n<0?-n:n;for(let o=1;o<t;o++){const l=e[o];let c=l<0?-l:l;for(;r&&c;)r>c?r%=c:c%=r;r+=c}return r}function Rf(e){return e=N(e),e instanceof Error?e:Math.floor(e)}const Mf={CEILING:Ye};function Of(){const e=V(P(arguments));if(e instanceof Error)return e;for(var t,n,r,o,l=1;(r=e.pop())!==void 0;){if(r===0)return 0;for(;r>1;){if(r%2){for(t=3,n=Math.floor(Math.sqrt(r));t<=n&&r%t;t+=2);o=t<=n?t:r}else o=2;for(r/=o,l*=o,t=e.length;t;e[--t]%o===0&&(e[t]/=o)===1&&e.splice(t,1));}}return l}function Df(e){return e=N(e),e instanceof Error?e:e===0?D:Math.log(e)}function wf(e,t){e=N(e),t=t?N(t):10;const n=Y(e,t);return n||(e===0||t===0?D:Math.log(e)/Math.log(t))}function Lf(e){return e=N(e),e instanceof Error?e:e===0?D:Math.log(e)/Math.log(10)}function Sf(e,t){return!Array.isArray(e)||!Array.isArray(t)||e.some(r=>!r.length)||t.some(r=>!r.length)||wn(e).some(r=>typeof r!="number")||wn(t).some(r=>typeof r!="number")||e[0].length!==t.length?R:Array(e.length).fill(0).map(()=>Array(t[0].length).fill(0)).map((r,o)=>r.map((l,c)=>e[o].reduce((h,g,u)=>h+g*t[u][c],0)))}function Pf(e,t){e=N(e),t=N(t);const n=Y(e,t);if(n)return n;if(t===0)return fe;let r=Math.abs(e%t);return r=e<0?t-r:r,t>0?r:-r}function Uf(e,t){e=N(e),t=N(t);const n=Y(e,t);return n||(e*t===0?0:e*t<0?D:Math.round(e/t)*t)}function yf(){const e=V(P(arguments));if(e instanceof Error)return e;let t=0,n=1;for(let r=0;r<e.length;r++)t+=e[r],n*=He(e[r]);return He(t)/n}function Ff(e){return arguments.length>1?B:(e=parseInt(e),!e||e<=0?R:Array(e).fill(0).map(()=>Array(e).fill(0)).map((t,n)=>(t[n]=1,t)))}function Hf(e){if(e=N(e),e instanceof Error)return e;let t=Math.ceil(Math.abs(e));return t=t&1?t:t+1,e>=0?t:-t}function Vf(){return Math.PI}function mi(e,t){e=N(e),t=N(t);const n=Y(e,t);if(n)return n;if(e===0&&t===0)return D;const r=Math.pow(e,t);return isNaN(r)?D:r}function Qt(){const t=P(arguments).filter(o=>o!=null);if(t.length===0)return 0;const n=V(t);if(n instanceof Error)return n;let r=1;for(let o=0;o<n.length;o++)r*=n[o];return r}function kf(e,t){e=N(e),t=N(t);const n=Y(e,t);return n||parseInt(e/t,10)}function Bf(e){return e=N(e),e instanceof Error?e:e*Math.PI/180}function Gf(){return Math.random()}function zf(e,t){e=N(e),t=N(t);const n=Y(e,t);return n||e+Math.ceil((t-e+1)*Math.random())-1}function Yf(e){if(e=N(e),e instanceof Error)return e;const t=String(e).split(""),n=["","C","CC","CCC","CD","D","DC","DCC","DCCC","CM","","X","XX","XXX","XL","L","LX","LXX","LXXX","XC","","I","II","III","IV","V","VI","VII","VIII","IX"];let r="",o=3;for(;o--;)r=(n[+t.pop()+o*10]||"")+r;return new Array(+t.join("")+1).join("M")+r}function Wn(e,t,n){e=N(e),t=N(t);const r=Y(e,t);if(r)return r;const o=e>=0?1:-1;let c=(Math.abs(e)+"e"+t).split("e");return c=(n(c[0]+"e"+c[1])+"e"+-t).split("e"),+(c[0]+"e"+c[1])*o}function pi(e,t){return Wn(e,t,Math.round)}function Wf(e,t){return Wn(e,t,Math.floor)}function qf(e,t){return Wn(e,t,Math.ceil)}function Kf(e){return e=N(e),e instanceof Error?e:1/Math.cos(e)}function Xf(e){return e=N(e),e instanceof Error?e:2/(Math.exp(e)+Math.exp(-e))}function $f(e,t,n,r){if(e=N(e),t=N(t),n=N(n),r=V(r),w(e,t,n,r))return R;let o=r[0]*Math.pow(e,t);for(let l=1;l<r.length;l++)o+=r[l]*Math.pow(e,t+l*n);return o}function xf(e){return e=N(e),e instanceof Error?e:e<0?-1:e===0?0:1}function Qf(e){return e=N(e),e instanceof Error?e:Math.sin(e)}function Jf(e){return e=N(e),e instanceof Error?e:(Math.exp(e)-Math.exp(-e))/2}function Zf(e){return e=N(e),e instanceof Error?e:e<0?D:Math.sqrt(e)}function bf(e){return e=N(e),e instanceof Error?e:Math.sqrt(e*Math.PI)}function _f(e,t){if(e=N(e),e instanceof Error)return e;switch(e){case 1:return $e(t);case 2:return Nt(t);case 3:return At(t);case 4:return Xt(t);case 5:return $t(t);case 6:return Qt(t);case 7:return Ee.S(t);case 8:return Ee.P(t);case 9:return Ve(t);case 10:return ge.S(t);case 11:return ge.P(t);case 101:return $e(t);case 102:return Nt(t);case 103:return At(t);case 104:return Xt(t);case 105:return $t(t);case 106:return Qt(t);case 107:return Ee.S(t);case 108:return Ee.P(t);case 109:return Ve(t);case 110:return ge.S(t);case 111:return ge.P(t)}}function Ve(){let e=0;return ce(Gt(arguments),t=>{if(e instanceof Error)return!1;if(t instanceof Error)e=t;else if(typeof t=="number")e+=t;else if(typeof t=="string"){const n=parseFloat(t);!isNaN(n)&&(e+=n)}else if(Array.isArray(t)){const n=Ve.apply(null,t);n instanceof Error?e=n:e+=n}}),e}function jf(e,t,n){if(e=P(e),n=n?P(n):e,e instanceof Error)return e;if(t==null||t instanceof Error)return 0;let r=0;const o=t==="*",l=o?null:rt(t+"");for(let c=0;c<e.length;c++){const h=e[c],g=n[c];if(o)r+=h;else{const u=[ye(h,ze)].concat(l);r+=it(u)?g:0}}return r}function el(){const e=Wt(...arguments);return Ve(e)}function tl(){if(!arguments||arguments.length===0)return R;const e=arguments.length+1;let t=0,n,r,o,l;for(let c=0;c<arguments[0].length;c++)if(arguments[0][c]instanceof Array)for(let h=0;h<arguments[0][c].length;h++){for(n=1,r=1;r<e;r++){const g=arguments[r-1][c][h];if(g instanceof Error)return g;if(l=N(g),l instanceof Error)return l;n*=l}t+=n}else{for(n=1,r=1;r<e;r++){const h=arguments[r-1][c];if(h instanceof Error)return h;if(o=N(h),o instanceof Error)return o;n*=o}t+=n}return t}function nl(){const e=V(P(arguments));if(e instanceof Error)return e;let t=0;const n=e.length;for(let r=0;r<n;r++)t+=Kt(e[r])?e[r]*e[r]:0;return t}function rl(e,t){if(e=V(P(e)),t=V(P(t)),w(e,t))return R;let n=0;for(let r=0;r<e.length;r++)n+=e[r]*e[r]-t[r]*t[r];return n}function il(e,t){if(e=V(P(e)),t=V(P(t)),w(e,t))return R;let n=0;e=V(P(e)),t=V(P(t));for(let r=0;r<e.length;r++)n+=e[r]*e[r]+t[r]*t[r];return n}function ol(e,t){if(e=V(P(e)),t=V(P(t)),w(e,t))return R;let n=0;e=P(e),t=P(t);for(let r=0;r<e.length;r++)n+=Math.pow(e[r]-t[r],2);return n}function sl(e){return e=N(e),e instanceof Error?e:Math.tan(e)}function ul(e){if(e=N(e),e instanceof Error)return e;const t=Math.exp(2*e);return(t-1)/(t+1)}function fl(e,t){e=N(e),t=N(t);const n=Y(e,t);return n||(e>0?1:-1)*Math.floor(Math.abs(e)*Math.pow(10,t))/Math.pow(10,t)}function ll(e,t){if(arguments.length!==2)return B;e=N(e),t=N(t);const n=Y(e,t);return n||e+t}function cl(e,t){if(arguments.length!==2)return B;e=N(e),t=N(t);const n=Y(e,t);return n||(t===0?fe:e/t)}function al(e,t){return arguments.length!==2?B:e instanceof Error?e:t instanceof Error?t:(e===null&&(e=void 0),t===null&&(t=void 0),e===t)}function hl(e,t){if(arguments.length!==2)return B;if(e instanceof Error)return e;if(t instanceof Error)return t;Yt(e,t)?(e=ie(e),t=ie(t)):(e=N(e),t=N(t));const n=Y(e,t);return n||e>t}function gl(e,t){if(arguments.length!==2)return B;Yt(e,t)?(e=ie(e),t=ie(t)):(e=N(e),t=N(t));const n=Y(e,t);return n||e>=t}function dl(e,t){if(arguments.length!==2)return B;Yt(e,t)?(e=ie(e),t=ie(t)):(e=N(e),t=N(t));const n=Y(e,t);return n||e<t}function El(e,t){if(arguments.length!==2)return B;Yt(e,t)?(e=ie(e),t=ie(t)):(e=N(e),t=N(t));const n=Y(e,t);return n||e<=t}function ml(e,t){if(arguments.length!==2)return B;e=N(e),t=N(t);const n=Y(e,t);return n||e-t}function pl(e,t){if(arguments.length!==2)return B;e=N(e),t=N(t);const n=Y(e,t);return n||e*t}function Il(e,t){return arguments.length!==2?B:e instanceof Error?e:t instanceof Error?t:(e===null&&(e=void 0),t===null&&(t=void 0),e!==t)}function Nl(e,t){return arguments.length!==2?B:mi(e,t)}var Al=Object.freeze({__proto__:null,ADD:ll,DIVIDE:cl,EQ:al,GT:hl,GTE:gl,LT:dl,LTE:El,MINUS:ml,MULTIPLY:pl,NE:Il,POW:Nl});const Cl=[void 0,0,1,void 0,void 0,void 0,void 0,void 0,void 0,void 0,void 0,void 0,1,2,3,4,5,6,0],vl=[[],[1,2,3,4,5,6,7],[7,1,2,3,4,5,6],[6,0,1,2,3,4,5],[],[],[],[],[],[],[],[7,1,2,3,4,5,6],[6,7,1,2,3,4,5],[5,6,7,1,2,3,4],[4,5,6,7,1,2,3],[3,4,5,6,7,1,2],[2,3,4,5,6,7,1],[1,2,3,4,5,6,7]],Jt=[[],[6,0],[0,1],[1,2],[2,3],[3,4],[4,5],[5,6],void 0,void 0,void 0,[0,0],[1,1],[2,2],[3,3],[4,4],[5,5],[6,6]];function Tl(e,t,n){let r;return e=N(e),t=N(t),n=N(n),w(e,t,n)?r=R:(r=new Date(e,t-1,n),r.getFullYear()<0&&(r=D)),Pe?Ue(r):r}function We(e,t,n){n=n.toUpperCase(),e=K(e),t=K(t);const r=e.getFullYear(),o=e.getMonth(),l=e.getDate(),c=t.getFullYear(),h=t.getMonth(),g=t.getDate();let u;switch(n){case"Y":u=Math.floor(Kn(e,t));break;case"D":u=Ze(t,e);break;case"M":u=h-o+12*(c-r),g<l&&u--;break;case"MD":l<=g?u=g-l:(h===0?(e.setFullYear(c-1),e.setMonth(12)):(e.setFullYear(c),e.setMonth(h-1)),u=Ze(t,e));break;case"YM":u=h-o+12*(c-r),g<l&&u--,u=u%12;break;case"YD":h>o||h===o&&g<l?e.setFullYear(c):e.setFullYear(c-1),u=Ze(t,e);break}return u}function Rl(e){if(typeof e!="string")return R;const t=Date.parse(e);if(isNaN(t))return R;const n=new Date(e);return Pe?Ue(n):n}function Ml(e){const t=K(e);return t instanceof Error?t:t.getDate()}function Zt(e){const t=new Date(e);return t.setHours(0,0,0,0),t}function Ze(e,t){return e=K(e),t=K(t),e instanceof Error?e:t instanceof Error?t:Ue(Zt(e))-Ue(Zt(t))}function qe(e,t,n){if(n=Ln(n||"false"),e=K(e),t=K(t),e instanceof Error)return e;if(t instanceof Error)return t;if(n instanceof Error)return n;const r=e.getMonth();let o=t.getMonth(),l,c;if(n)l=e.getDate()===31?30:e.getDate(),c=t.getDate()===31?30:t.getDate();else{const h=new Date(e.getFullYear(),r+1,0).getDate(),g=new Date(t.getFullYear(),o+1,0).getDate();l=e.getDate()===h?30:e.getDate(),t.getDate()===g?l<30?(o++,c=1):c=30:c=t.getDate()}return 360*(t.getFullYear()-e.getFullYear())+30*(o-r)+(c-l)}function Ol(e,t){if(e=K(e),e instanceof Error)return e;if(isNaN(t))return R;let n=e.getDate();e.setDate(1),t=parseInt(t,10),e.setMonth(e.getMonth()+t);let r=e.getMonth();if(n>28){let o=[31,28,31,30,31,30,31,31,30,31,30,31][r],l=e.getFullYear();r===1&&(l%4===0&&l%100!==0||l%400===0)&&(o=29),n=Math.min(n,o)}return e.setDate(n),Pe?Ue(e):e}function Dl(e,t){if(e=K(e),e instanceof Error)return e;if(isNaN(t))return R;t=parseInt(t,10);const n=new Date(e.getFullYear(),e.getMonth()+t+1,0);return Pe?Ue(n):n}function wl(e){return e=K(e),e instanceof Error?e:e.getHours()}function Ii(e){if(e=K(e),e instanceof Error)return e;e=Zt(e),e.setDate(e.getDate()+4-(e.getDay()||7));const t=new Date(e.getFullYear(),0,1);return Math.ceil(((e-t)/864e5+1)/7)}function Ll(e){return e=K(e),e instanceof Error?e:e.getMinutes()}function Sl(e){return e=K(e),e instanceof Error?e:e.getMonth()+1}function bt(e,t,n){return bt.INTL(e,t,1,n)}bt.INTL=(e,t,n,r)=>{if(e=K(e),e instanceof Error)return e;if(t=K(t),t instanceof Error)return t;let o=!1;const l=[],c=[1,2,3,4,5,6,0],h=new RegExp("^[0|1]{7}$");if(n===void 0)n=Jt[1];else if(typeof n=="string"&&h.test(n)){o=!0,n=n.split("");for(let s=0;s<n.length;s++)n[s]==="1"&&l.push(c[s])}else n=Jt[n];if(!(n instanceof Array))return R;r===void 0?r=[]:r instanceof Array||(r=[r]);for(let s=0;s<r.length;s++){const f=K(r[s]);if(f instanceof Error)return f;r[s]=f}const g=Math.round((t-e)/(1e3*60*60*24))+1;let u=g;const i=e;for(let s=0;s<g;s++){const f=new Date().getTimezoneOffset()>0?i.getUTCDay():i.getDay();let a=o?l.includes(f):f===n[0]||f===n[1];for(let d=0;d<r.length;d++){const E=r[d];if(E.getDate()===i.getDate()&&E.getMonth()===i.getMonth()&&E.getFullYear()===i.getFullYear()){a=!0;break}}a&&u--,i.setDate(i.getDate()+1)}return u};function Pl(){return Pe?Ue(new Date):new Date}function Ul(e){return e=K(e),e instanceof Error?e:e.getSeconds()}function yl(e,t,n){return e=N(e),t=N(t),n=N(n),w(e,t,n)?R:e<0||t<0||n<0?D:(3600*e+60*t+n)/86400}function Fl(e){return e=K(e),e instanceof Error?e:(3600*e.getHours()+60*e.getMinutes()+e.getSeconds())/86400}function Hl(){const e=Zt(new Date);return Pe?Ue(e):e}function Vl(e,t){if(e=K(e),e instanceof Error)return e;t===void 0&&(t=1);const n=e.getDay();return vl[t][n]}function kl(e,t){if(e=K(e),e instanceof Error)return e;if(t===void 0&&(t=1),t===21)return Ii(e);const n=Cl[t];let r=new Date(e.getFullYear(),0,1);const o=r.getDay()<n?1:0;return r-=Math.abs(r.getDay()-n)*24*60*60*1e3,Math.floor((e-r)/(1e3*60*60*24)/7+1)+o}function _t(e,t,n){return _t.INTL(e,t,1,n)}_t.INTL=(e,t,n,r)=>{if(e=K(e),e instanceof Error)return e;if(t=N(t),t instanceof Error)return t;if(n===void 0?n=Jt[1]:n=Jt[n],!(n instanceof Array))return R;r===void 0?r=[]:r instanceof Array||(r=[r]);for(let c=0;c<r.length;c++){const h=K(r[c]);if(h instanceof Error)return h;r[c]=h}let o=0;const l=Math.sign(t);for(;o<t*l;){e.setDate(e.getDate()+l);const c=e.getDay();if(!(c===n[0]||c===n[1])){for(let h=0;h<r.length;h++){const g=r[h];if(g.getDate()===e.getDate()&&g.getMonth()===e.getMonth()&&g.getFullYear()===e.getFullYear()){o--;break}}o++}}return e.getFullYear()<1900?R:e};function Bl(e){return e=K(e),e instanceof Error?e:e.getFullYear()}function qn(e){return new Date(e,1,29).getMonth()===1}function jt(e,t){return Math.ceil((t-e)/1e3/60/60/24)}function Kn(e,t,n){if(e=K(e),e instanceof Error)return e;if(t=K(t),t instanceof Error)return t;n=n||0;let r=e.getDate();const o=e.getMonth()+1,l=e.getFullYear();let c=t.getDate();const h=t.getMonth()+1,g=t.getFullYear();switch(n){case 0:return r===31&&c===31?(r=30,c=30):r===31?r=30:r===30&&c===31&&(c=30),(c+h*30+g*360-(r+o*30+l*360))/360;case 1:{const u=(d,E)=>{const m=d.getFullYear(),C=new Date(m,2,1);if(qn(m)&&d<C&&E>=C)return!0;const v=E.getFullYear(),p=new Date(v,2,1);return qn(v)&&E>=p&&d<p};let i=365;if(l===g||l+1===g&&(o>h||o===h&&r>=c))return(l===g&&qn(l)||u(e,t)||h===1&&c===29)&&(i=366),jt(e,t)/i;const s=g-l+1,a=(new Date(g+1,0,1)-new Date(l,0,1))/1e3/60/60/24/s;return jt(e,t)/a}case 2:return jt(e,t)/360;case 3:return jt(e,t)/365;case 4:return(c+h*30+g*360-(r+o*30+l*360))/360}}function Xn(e){return/^[01]{1,10}$/.test(e)}function Gl(e,t){return e=N(e),t=N(t),w(e,t)?R:Vt.besseli(e,t)}function zl(e,t){return e=N(e),t=N(t),w(e,t)?R:Vt.besselj(e,t)}function Yl(e,t){return e=N(e),t=N(t),w(e,t)?R:Vt.besselk(e,t)}function Wl(e,t){return e=N(e),t=N(t),w(e,t)?R:Vt.bessely(e,t)}function ql(e){if(!Xn(e))return D;const t=parseInt(e,2),n=e.toString();return n.length===10&&n.substring(0,1)==="1"?parseInt(n.substring(1),2)-512:t}function Kl(e,t){if(!Xn(e))return D;const n=e.toString();if(n.length===10&&n.substring(0,1)==="1")return(0xfffffffe00+parseInt(n.substring(1),2)).toString(16);const r=parseInt(e,2).toString(16);return t===void 0?r:isNaN(t)?R:t<0?D:(t=Math.floor(t),t>=r.length?me("0",t-r.length)+r:D)}function Xl(e,t){if(!Xn(e))return D;const n=e.toString();if(n.length===10&&n.substring(0,1)==="1")return(1073741312+parseInt(n.substring(1),2)).toString(8);const r=parseInt(e,2).toString(8);return t===void 0?r:isNaN(t)?R:t<0?D:(t=Math.floor(t),t>=r.length?me("0",t-r.length)+r:D)}function $l(e,t){return e=N(e),t=N(t),w(e,t)?R:e<0||t<0||Math.floor(e)!==e||Math.floor(t)!==t||e>0xffffffffffff||t>0xffffffffffff?D:e&t}function xl(e,t){return e=N(e),t=N(t),w(e,t)?R:e<0||Math.floor(e)!==e||e>0xffffffffffff||Math.abs(t)>53?D:t>=0?e<<t:e>>-t}function Ql(e,t){return e=N(e),t=N(t),w(e,t)?R:e<0||t<0||Math.floor(e)!==e||Math.floor(t)!==t||e>0xffffffffffff||t>0xffffffffffff?D:e|t}function Jl(e,t){return e=N(e),t=N(t),w(e,t)?R:e<0||Math.floor(e)!==e||e>0xffffffffffff||Math.abs(t)>53?D:t>=0?e>>t:e<<-t}function Zl(e,t){return e=N(e),t=N(t),w(e,t)?R:e<0||t<0||Math.floor(e)!==e||Math.floor(t)!==t||e>0xffffffffffff||t>0xffffffffffff?D:e^t}function ae(e,t,n){if(e=N(e),t=N(t),w(e,t))return e;if(n=n===void 0?"i":n,n!=="i"&&n!=="j")return R;if(e===0&&t===0)return 0;if(e===0)return t===1?n:t.toString()+n;if(t===0)return e.toString();{const r=t>0?"+":"";return e.toString()+r+(t===1?n:t.toString()+n)}}function bl(e,t,n){if(e=N(e),e instanceof Error)return e;const r=[["a.u. of action","?",null,"action",!1,!1,105457168181818e-48],["a.u. of charge","e",null,"electric_charge",!1,!1,160217653141414e-33],["a.u. of energy","Eh",null,"energy",!1,!1,435974417757576e-32],["a.u. of length","a?",null,"length",!1,!1,529177210818182e-25],["a.u. of mass","m?",null,"mass",!1,!1,910938261616162e-45],["a.u. of time","?/Eh",null,"time",!1,!1,241888432650516e-31],["admiralty knot","admkn",null,"speed",!1,!0,.514773333],["ampere","A",null,"electric_current",!0,!1,1],["ampere per meter","A/m",null,"magnetic_field_intensity",!0,!1,1],["ångström","Å",["ang"],"length",!1,!0,1e-10],["are","ar",null,"area",!1,!0,100],["astronomical unit","ua",null,"length",!1,!1,149597870691667e-25],["bar","bar",null,"pressure",!1,!1,1e5],["barn","b",null,"area",!1,!1,1e-28],["becquerel","Bq",null,"radioactivity",!0,!1,1],["bit","bit",["b"],"information",!1,!0,1],["btu","BTU",["btu"],"energy",!1,!0,1055.05585262],["byte","byte",null,"information",!1,!0,8],["candela","cd",null,"luminous_intensity",!0,!1,1],["candela per square metre","cd/m?",null,"luminance",!0,!1,1],["coulomb","C",null,"electric_charge",!0,!1,1],["cubic ångström","ang3",["ang^3"],"volume",!1,!0,1e-30],["cubic foot","ft3",["ft^3"],"volume",!1,!0,.028316846592],["cubic inch","in3",["in^3"],"volume",!1,!0,16387064e-12],["cubic light-year","ly3",["ly^3"],"volume",!1,!0,846786664623715e-61],["cubic metre","m3",["m^3"],"volume",!0,!0,1],["cubic mile","mi3",["mi^3"],"volume",!1,!0,416818182544058e-5],["cubic nautical mile","Nmi3",["Nmi^3"],"volume",!1,!0,6352182208],["cubic Pica","Pica3",["Picapt3","Pica^3","Picapt^3"],"volume",!1,!0,758660370370369e-22],["cubic yard","yd3",["yd^3"],"volume",!1,!0,.764554857984],["cup","cup",null,"volume",!1,!0,.0002365882365],["dalton","Da",["u"],"mass",!1,!1,166053886282828e-41],["day","d",["day"],"time",!1,!0,86400],["degree","°",null,"angle",!1,!1,.0174532925199433],["degrees Rankine","Rank",null,"temperature",!1,!0,.555555555555556],["dyne","dyn",["dy"],"force",!1,!0,1e-5],["electronvolt","eV",["ev"],"energy",!1,!0,1.60217656514141],["ell","ell",null,"length",!1,!0,1.143],["erg","erg",["e"],"energy",!1,!0,1e-7],["farad","F",null,"electric_capacitance",!0,!1,1],["fluid ounce","oz",null,"volume",!1,!0,295735295625e-16],["foot","ft",null,"length",!1,!0,.3048],["foot-pound","flb",null,"energy",!1,!0,1.3558179483314],["gal","Gal",null,"acceleration",!1,!1,.01],["gallon","gal",null,"volume",!1,!0,.003785411784],["gauss","G",["ga"],"magnetic_flux_density",!1,!0,1],["grain","grain",null,"mass",!1,!0,647989e-10],["gram","g",null,"mass",!1,!0,.001],["gray","Gy",null,"absorbed_dose",!0,!1,1],["gross registered ton","GRT",["regton"],"volume",!1,!0,2.8316846592],["hectare","ha",null,"area",!1,!0,1e4],["henry","H",null,"inductance",!0,!1,1],["hertz","Hz",null,"frequency",!0,!1,1],["horsepower","HP",["h"],"power",!1,!0,745.69987158227],["horsepower-hour","HPh",["hh","hph"],"energy",!1,!0,2684519538e-3],["hour","h",["hr"],"time",!1,!0,3600],["imperial gallon (U.K.)","uk_gal",null,"volume",!1,!0,.00454609],["imperial hundredweight","lcwt",["uk_cwt","hweight"],"mass",!1,!0,50.802345],["imperial quart (U.K)","uk_qt",null,"volume",!1,!0,.0011365225],["imperial ton","brton",["uk_ton","LTON"],"mass",!1,!0,1016.046909],["inch","in",null,"length",!1,!0,.0254],["international acre","uk_acre",null,"area",!1,!0,4046.8564224],["IT calorie","cal",null,"energy",!1,!0,4.1868],["joule","J",null,"energy",!0,!0,1],["katal","kat",null,"catalytic_activity",!0,!1,1],["kelvin","K",["kel"],"temperature",!0,!0,1],["kilogram","kg",null,"mass",!0,!0,1],["knot","kn",null,"speed",!1,!0,.514444444444444],["light-year","ly",null,"length",!1,!0,9460730472580800],["litre","L",["l","lt"],"volume",!1,!0,.001],["lumen","lm",null,"luminous_flux",!0,!1,1],["lux","lx",null,"illuminance",!0,!1,1],["maxwell","Mx",null,"magnetic_flux",!1,!1,1e-18],["measurement ton","MTON",null,"volume",!1,!0,1.13267386368],["meter per hour","m/h",["m/hr"],"speed",!1,!0,.00027777777777778],["meter per second","m/s",["m/sec"],"speed",!0,!0,1],["meter per second squared","m?s??",null,"acceleration",!0,!1,1],["parsec","pc",["parsec"],"length",!1,!0,0x6da012f958ee1c],["meter squared per second","m?/s",null,"kinematic_viscosity",!0,!1,1],["metre","m",null,"length",!0,!0,1],["miles per hour","mph",null,"speed",!1,!0,.44704],["millimetre of mercury","mmHg",null,"pressure",!1,!1,133.322],["minute","?",null,"angle",!1,!1,.000290888208665722],["minute","min",["mn"],"time",!1,!0,60],["modern teaspoon","tspm",null,"volume",!1,!0,5e-6],["mole","mol",null,"amount_of_substance",!0,!1,1],["morgen","Morgen",null,"area",!1,!0,2500],["n.u. of action","?",null,"action",!1,!1,105457168181818e-48],["n.u. of mass","m?",null,"mass",!1,!1,910938261616162e-45],["n.u. of speed","c?",null,"speed",!1,!1,299792458],["n.u. of time","?/(me?c??)",null,"time",!1,!1,128808866778687e-35],["nautical mile","M",["Nmi"],"length",!1,!0,1852],["newton","N",null,"force",!0,!0,1],["œrsted","Oe ",null,"magnetic_field_intensity",!1,!1,79.5774715459477],["ohm","Ω",null,"electric_resistance",!0,!1,1],["ounce mass","ozm",null,"mass",!1,!0,.028349523125],["pascal","Pa",null,"pressure",!0,!1,1],["pascal second","Pa?s",null,"dynamic_viscosity",!0,!1,1],["pferdestärke","PS",null,"power",!1,!0,735.49875],["phot","ph",null,"illuminance",!1,!1,1e-4],["pica (1/6 inch)","pica",null,"length",!1,!0,.00035277777777778],["pica (1/72 inch)","Pica",["Picapt"],"length",!1,!0,.00423333333333333],["poise","P",null,"dynamic_viscosity",!1,!1,.1],["pond","pond",null,"force",!1,!0,.00980665],["pound force","lbf",null,"force",!1,!0,4.4482216152605],["pound mass","lbm",null,"mass",!1,!0,.45359237],["quart","qt",null,"volume",!1,!0,.000946352946],["radian","rad",null,"angle",!0,!1,1],["second","?",null,"angle",!1,!1,484813681109536e-20],["second","s",["sec"],"time",!0,!0,1],["short hundredweight","cwt",["shweight"],"mass",!1,!0,45.359237],["siemens","S",null,"electrical_conductance",!0,!1,1],["sievert","Sv",null,"equivalent_dose",!0,!1,1],["slug","sg",null,"mass",!1,!0,14.59390294],["square ångström","ang2",["ang^2"],"area",!1,!0,1e-20],["square foot","ft2",["ft^2"],"area",!1,!0,.09290304],["square inch","in2",["in^2"],"area",!1,!0,64516e-8],["square light-year","ly2",["ly^2"],"area",!1,!0,895054210748189e17],["square meter","m?",null,"area",!0,!0,1],["square mile","mi2",["mi^2"],"area",!1,!0,2589988110336e-6],["square nautical mile","Nmi2",["Nmi^2"],"area",!1,!0,3429904],["square Pica","Pica2",["Picapt2","Pica^2","Picapt^2"],"area",!1,!0,1792111111111e-17],["square yard","yd2",["yd^2"],"area",!1,!0,.83612736],["statute mile","mi",null,"length",!1,!0,1609.344],["steradian","sr",null,"solid_angle",!0,!1,1],["stilb","sb",null,"luminance",!1,!1,1e-4],["stokes","St",null,"kinematic_viscosity",!1,!1,1e-4],["stone","stone",null,"mass",!1,!0,6.35029318],["tablespoon","tbs",null,"volume",!1,!0,147868e-10],["teaspoon","tsp",null,"volume",!1,!0,492892e-11],["tesla","T",null,"magnetic_flux_density",!0,!0,1],["thermodynamic calorie","c",null,"energy",!1,!0,4.184],["ton","ton",null,"mass",!1,!0,907.18474],["tonne","t",null,"mass",!1,!1,1e3],["U.K. pint","uk_pt",null,"volume",!1,!0,.00056826125],["U.S. bushel","bushel",null,"volume",!1,!0,.03523907],["U.S. oil barrel","barrel",null,"volume",!1,!0,.158987295],["U.S. pint","pt",["us_pt"],"volume",!1,!0,.000473176473],["U.S. survey mile","survey_mi",null,"length",!1,!0,1609.347219],["U.S. survey/statute acre","us_acre",null,"area",!1,!0,4046.87261],["volt","V",null,"voltage",!0,!1,1],["watt","W",null,"power",!0,!0,1],["watt-hour","Wh",["wh"],"energy",!1,!0,3600],["weber","Wb",null,"magnetic_flux",!0,!1,1],["yard","yd",null,"length",!1,!0,.9144],["year","yr",null,"time",!1,!0,31557600]],o={Yi:["yobi",80,12089258196146292e8,"Yi","yotta"],Zi:["zebi",70,11805916207174113e5,"Zi","zetta"],Ei:["exbi",60,1152921504606847e3,"Ei","exa"],Pi:["pebi",50,0x4000000000000,"Pi","peta"],Ti:["tebi",40,1099511627776,"Ti","tera"],Gi:["gibi",30,1073741824,"Gi","giga"],Mi:["mebi",20,1048576,"Mi","mega"],ki:["kibi",10,1024,"ki","kilo"]},l={Y:["yotta",1e24,"Y"],Z:["zetta",1e21,"Z"],E:["exa",1e18,"E"],P:["peta",1e15,"P"],T:["tera",1e12,"T"],G:["giga",1e9,"G"],M:["mega",1e6,"M"],k:["kilo",1e3,"k"],h:["hecto",100,"h"],e:["dekao",10,"e"],d:["deci",.1,"d"],c:["centi",.01,"c"],m:["milli",.001,"m"],u:["micro",1e-6,"u"],n:["nano",1e-9,"n"],p:["pico",1e-12,"p"],f:["femto",1e-15,"f"],a:["atto",1e-18,"a"],z:["zepto",1e-21,"z"],y:["yocto",1e-24,"y"]};let c=null,h=null,g=t,u=n,i=1,s=1,f;for(let a=0;a<r.length;a++)f=r[a][2]===null?[]:r[a][2],(r[a][1]===g||f.indexOf(g)>=0)&&(c=r[a]),(r[a][1]===u||f.indexOf(u)>=0)&&(h=r[a]);if(c===null){const a=o[t.substring(0,2)];let d=l[t.substring(0,1)];t.substring(0,2)==="da"&&(d=["dekao",10,"da"]),a?(i=a[2],g=t.substring(2)):d&&(i=d[1],g=t.substring(d[2].length));for(let E=0;E<r.length;E++)f=r[E][2]===null?[]:r[E][2],(r[E][1]===g||f.indexOf(g)>=0)&&(c=r[E])}if(h===null){const a=o[n.substring(0,2)];let d=l[n.substring(0,1)];n.substring(0,2)==="da"&&(d=["dekao",10,"da"]),a?(s=a[2],u=n.substring(2)):d&&(s=d[1],u=n.substring(d[2].length));for(let E=0;E<r.length;E++)f=r[E][2]===null?[]:r[E][2],(r[E][1]===u||f.indexOf(u)>=0)&&(h=r[E])}return c===null||h===null||c[3]!==h[3]?B:e*c[6]*i/(h[6]*s)}function _l(e,t){if(e=N(e),e instanceof Error)return e;if(!/^-?[0-9]{1,3}$/.test(e)||e<-512||e>511)return D;if(e<0)return"1"+me("0",9-(512+e).toString(2).length)+(512+e).toString(2);const n=parseInt(e,10).toString(2);return typeof t>"u"?n:isNaN(t)?R:t<0?D:(t=Math.floor(t),t>=n.length?me("0",t-n.length)+n:D)}function jl(e,t){if(e=N(e),e instanceof Error)return e;if(!/^-?[0-9]{1,12}$/.test(e)||e<-549755813888||e>549755813887)return D;if(e<0)return(1099511627776+e).toString(16);const n=parseInt(e,10).toString(16);return typeof t>"u"?n:isNaN(t)?R:t<0?D:(t=Math.floor(t),t>=n.length?me("0",t-n.length)+n:D)}function ec(e,t){if(e=N(e),e instanceof Error)return e;if(!/^-?[0-9]{1,9}$/.test(e)||e<-536870912||e>536870911)return D;if(e<0)return(1073741824+e).toString(8);const n=parseInt(e,10).toString(8);return typeof t>"u"?n:isNaN(t)?R:t<0?D:(t=Math.floor(t),t>=n.length?me("0",t-n.length)+n:D)}function tc(e,t){return t=t===void 0?0:t,e=N(e),t=N(t),w(e,t)?R:e===t?1:0}function Ni(e,t){return t=t===void 0?0:t,e=N(e),t=N(t),w(e,t)?R:F.erf(e)}function Ai(e){return isNaN(e)?R:F.erfc(e)}function nc(e,t){return t=t||0,e=N(e),w(t,e)?e:e>=t?1:0}function rc(e,t){if(!/^[0-9A-Fa-f]{1,10}$/.test(e))return D;const n=e.length===10&&e.substring(0,1).toLowerCase()==="f",r=n?parseInt(e,16)-1099511627776:parseInt(e,16);if(r<-512||r>511)return D;if(n)return"1"+me("0",9-(512+r).toString(2).length)+(512+r).toString(2);const o=r.toString(2);return t===void 0?o:isNaN(t)?R:t<0?D:(t=Math.floor(t),t>=o.length?me("0",t-o.length)+o:D)}function ic(e){if(!/^[0-9A-Fa-f]{1,10}$/.test(e))return D;const t=parseInt(e,16);return t>=549755813888?t-1099511627776:t}function oc(e,t){if(!/^[0-9A-Fa-f]{1,10}$/.test(e))return D;const n=parseInt(e,16);if(n>536870911&&n<0xffe0000000)return D;if(n>=0xffe0000000)return(n-0xffc0000000).toString(8);const r=n.toString(8);return t===void 0?r:isNaN(t)?R:t<0?D:(t=Math.floor(t),t>=r.length?me("0",t-r.length)+r:D)}function $n(e){const t=_(e),n=b(e);return w(t,n)?R:Math.sqrt(Math.pow(t,2)+Math.pow(n,2))}function b(e){if(e===void 0||e===!0||e===!1)return R;if(e===0||e==="0")return 0;if(["i","j"].indexOf(e)>=0)return 1;e=e+"",e=e.replace("+i","+1i").replace("-i","-1i").replace("+j","+1j").replace("-j","-1j");let t=e.indexOf("+"),n=e.indexOf("-");t===0&&(t=e.indexOf("+",1)),n===0&&(n=e.indexOf("-",1));const r=e.substring(e.length-1,e.length),o=r==="i"||r==="j";return t>=0||n>=0?o?t>=0?isNaN(e.substring(0,t))||isNaN(e.substring(t+1,e.length-1))?D:Number(e.substring(t+1,e.length-1)):isNaN(e.substring(0,n))||isNaN(e.substring(n+1,e.length-1))?D:-Number(e.substring(n+1,e.length-1)):D:o?isNaN(e.substring(0,e.length-1))?D:e.substring(0,e.length-1):isNaN(e)?D:0}function xn(e){const t=_(e),n=b(e);return w(t,n)?R:t===0&&n===0?fe:t===0&&n>0?Math.PI/2:t===0&&n<0?-Math.PI/2:n===0&&t>0?0:n===0&&t<0?-Math.PI:t>0?Math.atan(n/t):t<0&&n>=0?Math.atan(n/t)+Math.PI:Math.atan(n/t)-Math.PI}function sc(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);return r=r==="i"||r==="j"?r:"i",n!==0?ae(t,-n,r):e}function en(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);return r=r==="i"||r==="j"?r:"i",ae(Math.cos(t)*(Math.exp(n)+Math.exp(-n))/2,-Math.sin(t)*(Math.exp(n)-Math.exp(-n))/2,r)}function Ci(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);return r=r==="i"||r==="j"?r:"i",ae(Math.cos(n)*(Math.exp(t)+Math.exp(-t))/2,Math.sin(n)*(Math.exp(t)-Math.exp(-t))/2,r)}function uc(e){const t=_(e),n=b(e);return w(t,n)?R:be(en(e),tn(e))}function be(e,t){const n=_(e),r=b(e),o=_(t),l=b(t);if(w(n,r,o,l))return R;const c=e.substring(e.length-1),h=t.substring(t.length-1);let g="i";if((c==="j"||h==="j")&&(g="j"),o===0&&l===0)return D;const u=o*o+l*l;return ae((n*o+r*l)/u,(r*o-n*l)/u,g)}function fc(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);r=r==="i"||r==="j"?r:"i";const o=Math.exp(t);return ae(o*Math.cos(n),o*Math.sin(n),r)}function lc(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);return r=r==="i"||r==="j"?r:"i",ae(Math.log(Math.sqrt(t*t+n*n)),Math.atan(n/t),r)}function cc(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);return r=r==="i"||r==="j"?r:"i",ae(Math.log(Math.sqrt(t*t+n*n))/Math.log(10),Math.atan(n/t)/Math.log(10),r)}function ac(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);return r=r==="i"||r==="j"?r:"i",ae(Math.log(Math.sqrt(t*t+n*n))/Math.log(2),Math.atan(n/t)/Math.log(2),r)}function hc(e,t){t=N(t);const n=_(e),r=b(e);if(w(t,n,r))return R;let o=e.substring(e.length-1);o=o==="i"||o==="j"?o:"i";const l=Math.pow($n(e),t),c=xn(e);return ae(l*Math.cos(t*c),l*Math.sin(t*c),o)}function gc(){let e=arguments[0];if(!arguments.length)return R;for(let t=1;t<arguments.length;t++){const n=_(e),r=b(e),o=_(arguments[t]),l=b(arguments[t]);if(w(n,r,o,l))return R;e=ae(n*o-r*l,n*l+r*o)}return e}function _(e){if(e===void 0||e===!0||e===!1)return R;if(e===0||e==="0"||["i","+i","1i","+1i","-i","-1i","j","+j","1j","+1j","-j","-1j"].indexOf(e)>=0)return 0;e=e+"";let t=e.indexOf("+"),n=e.indexOf("-");t===0&&(t=e.indexOf("+",1)),n===0&&(n=e.indexOf("-",1));const r=e.substring(e.length-1,e.length),o=r==="i"||r==="j";return t>=0||n>=0?o?t>=0?isNaN(e.substring(0,t))||isNaN(e.substring(t+1,e.length-1))?D:Number(e.substring(0,t)):isNaN(e.substring(0,n))||isNaN(e.substring(n+1,e.length-1))?D:Number(e.substring(0,n)):D:o?isNaN(e.substring(0,e.length-1))?D:0:isNaN(e)?D:e}function dc(e){if(e===!0||e===!1)return R;const t=_(e),n=b(e);return w(t,n)?R:be("1",en(e))}function Ec(e){const t=_(e),n=b(e);return w(t,n)?R:be("1",Ci(e))}function tn(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);return r=r==="i"||r==="j"?r:"i",ae(Math.sin(t)*(Math.exp(n)+Math.exp(-n))/2,Math.cos(t)*(Math.exp(n)-Math.exp(-n))/2,r)}function vi(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);return r=r==="i"||r==="j"?r:"i",ae(Math.cos(n)*(Math.exp(t)-Math.exp(-t))/2,Math.sin(n)*(Math.exp(t)+Math.exp(-t))/2,r)}function mc(e){const t=_(e),n=b(e);if(w(t,n))return R;let r=e.substring(e.length-1);r=r==="i"||r==="j"?r:"i";const o=Math.sqrt($n(e)),l=xn(e);return ae(o*Math.cos(l/2),o*Math.sin(l/2),r)}function pc(e){if(e===!0||e===!1)return R;const t=_(e),n=b(e);return w(t,n)?D:be("1",tn(e))}function Ic(e){if(e===!0||e===!1)return R;const t=_(e),n=b(e);return w(t,n)?D:be("1",vi(e))}function Nc(e,t){const n=_(e),r=b(e),o=_(t),l=b(t);if(w(n,r,o,l))return R;const c=e.substring(e.length-1),h=t.substring(t.length-1);let g="i";return(c==="j"||h==="j")&&(g="j"),ae(n-o,r-l,g)}function Ac(){if(!arguments.length)return R;const e=P(arguments);let t=0,n=0;for(const r of e){const o=+_(r),l=+b(r);if(w(o,l))return R;t+=o,n+=l}return ae(t,n,"i")}function Cc(e){if(e===!0||e===!1)return R;const t=_(e),n=b(e);return w(t,n)?R:be(tn(e),en(e))}function vc(e,t){if(!/^[0-7]{1,10}$/.test(e))return D;const n=e.length===10&&e.substring(0,1)==="7",r=n?parseInt(e,8)-1073741824:parseInt(e,8);if(r<-512||r>511)return D;if(n)return"1"+me("0",9-(512+r).toString(2).length)+(512+r).toString(2);const o=r.toString(2);return typeof t>"u"?o:isNaN(t)?R:t<0?D:(t=Math.floor(t),t>=o.length?me("0",t-o.length)+o:D)}function Tc(e){if(!/^[0-7]{1,10}$/.test(e))return D;const t=parseInt(e,8);return t>=536870912?t-1073741824:t}function Rc(e,t){if(!/^[0-7]{1,10}$/.test(e))return D;const n=parseInt(e,8);if(n>=536870912)return"ff"+(n+3221225472).toString(16);const r=n.toString(16);return t===void 0?r:isNaN(t)?R:t<0?D:(t=Math.floor(t),t>=r.length?me("0",t-r.length)+r:D)}const Mc=It.DIST,Oc=It.INV,Dc=ot.DIST,wc=Ye.MATH,Lc=Ye.PRECISE,Sc=Re.DIST,Pc=Re.DIST.RT,Uc=Re.INV,yc=Re.INV.RT,Fc=Re.TEST,Hc=st.P,Vc=st.P,kc=st.S,Bc=ot.INV,Gc=Ai.PRECISE,zc=Ni.PRECISE,Yc=yn.DIST,Wc=Me.DIST,qc=Me.DIST.RT,Kc=Me.INV,Xc=Me.INV.RT,$c=Je.MATH,xc=Je.PRECISE,Qc=Me.TEST,Jc=Ct.DIST,Zc=Ct.INV,bc=Fn.PRECISE,_c=Hn.DIST,jc=ut.INV,e0=ut.DIST,t0=ut.INV,n0=xe.MULT,r0=xe.SNGL,i0=kn.DIST,o0=bt.INTL,s0=Oe.DIST,u0=Oe.INV,f0=Oe.S.DIST,l0=Oe.S.INV,c0=pe.EXC,a0=pe.INC,h0=vt.EXC,g0=vt.INC,d0=Bn.DIST,E0=Qe.EXC,m0=Qe.INC,p0=Tt.AVG,I0=Tt.EQ,N0=Gn.P,A0=Ee.P,C0=Ee.S,v0=de.DIST,T0=de.DIST.RT,R0=de.INV,M0=de.TEST,O0=ge.P,D0=ge.S,w0=zn.DIST,L0=_t.INTL,S0=Yn.TEST;function Qn(e){const t=[];return ce(e,n=>{n&&t.push(n)}),t}function Ce(e,t){const n={};for(let l=1;l<e[0].length;++l)n[l]=!0;let r=t[0].length;for(let l=1;l<t.length;++l)t[l].length>r&&(r=t[l].length);for(let l=1;l<e.length;++l)for(let c=1;c<e[l].length;++c){let h=!1,g=!1;for(let u=0;u<t.length;++u){const i=t[u];if(i.length<r)continue;const s=i[0];if(e[l][0]===s){g=!0;for(let f=1;f<i.length;++f)if(!h)if(i[f]===void 0||i[f]==="*")h=!0;else{const d=rt(i[f]+""),E=[ye(e[l][c],ze)].concat(d);h=it(E)}}}g&&(n[c]=n[c]&&h)}const o=[];for(let l=0;l<e[0].length;++l)n[l]&&o.push(l-1);return o}function P0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);let l=0;return ce(r,c=>{l+=o[c]}),r.length===0?fe:l/r.length}function U0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);const l=[];return ce(r,c=>{l.push(o[c])}),Nt(l)}function y0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);const l=[];return ce(r,c=>{l.push(o[c])}),At(l)}function F0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const l=Ae(e,t);o=ee(e[l])}else o=ee(e[t]);return r.length===0?R:r.length>1?D:o[r[0]]}function H0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);let l=o[r[0]];return ce(r,c=>{l<o[c]&&(l=o[c])}),l}function V0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);let l=o[r[0]];return ce(r,c=>{l>o[c]&&(l=o[c])}),l}function k0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const h=Ae(e,t);o=ee(e[h])}else o=ee(e[t]);let l=[];ce(r,h=>{l.push(o[h])}),l=Qn(l);let c=1;return ce(l,h=>{c*=h}),c}function B0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);let l=[];return ce(r,c=>{l.push(o[c])}),l=Qn(l),Ee.S(l)}function G0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);let l=[];return ce(r,c=>{l.push(o[c])}),l=Qn(l),Ee.P(l)}function z0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);const l=[];return ce(r,c=>{l.push(o[c])}),Ve(l)}function Y0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);const l=[];return ce(r,c=>{l.push(o[c])}),ge.S(l)}function W0(e,t,n){if(isNaN(t)&&typeof t!="string")return R;const r=Ce(e,n);let o=[];if(typeof t=="string"){const c=Ae(e,t);o=ee(e[c])}else o=ee(e[t]);const l=[];return ce(r,c=>{l.push(o[c])}),ge.P(l)}function Jn(e){return e&&e.getTime&&!isNaN(e.getTime())}function Zn(e){return e instanceof Date?e:new Date(e)}function q0(e,t,n){let r=K(t);for(r.setFullYear(e.getFullYear()),r<e&&r.setFullYear(r.getFullYear()+1);r>e;)r.setMonth(r.getMonth()+-12/n);return r}function Ti(e){return e=N(e),[1,2,4].indexOf(e)===-1?D:e}function Ri(e){return e=N(e),[0,1,2,3,4].indexOf(e)===-1?D:e}function K0(e,t,n,r,o,l,c){return e=Zn(e),t=Zn(t),n=Zn(n),l=Ti(l),c=Ri(c),Y(l,c)?D:!Jn(e)||!Jn(t)||!Jn(n)?R:r<=0||o<=0||n<=e?D:(o=o||0,c=c||0,o*r*Kn(e,n,c))}function X0(e,t,n,r){if(r=Ri(r),n=Ti(n),e=K(e),t=K(t),Y(e,t))return R;if(Y(n,r)||e>=t)return D;if(r===1){let l=q0(e,t,n),c=K(l);return c.setMonth(c.getMonth()+12/n),We(l,c,"D")}let o;switch(r){case 0:case 2:case 4:o=360;break;case 3:o=365;break;default:return D}return o/n}function $0(e,t,n,r,o,l){if(e=N(e),t=N(t),n=N(n),w(e,t,n))return R;if(e<=0||t<=0||n<=0||r<1||o<1||r>o||l!==0&&l!==1)return D;const c=Rt(e,t,n,0,l);let h=0;r===1&&(l===0&&(h=-n),r++);for(let g=r;g<=o;g++)h+=l===1?_e(e,g-2,c,n,1)-c:_e(e,g-1,c,n,0);return h*=e,h}function x0(e,t,n,r,o,l){if(e=N(e),t=N(t),n=N(n),w(e,t,n))return R;if(e<=0||t<=0||n<=0||r<1||o<1||r>o||l!==0&&l!==1)return D;const c=Rt(e,t,n,0,l);let h=0;r===1&&(h=l===0?c+n*e:c,r++);for(let g=r;g<=o;g++)h+=l>0?c-(_e(e,g-2,c,n,1)-c)*e:c-_e(e,g-1,c,n,0)*e;return h}function Q0(e,t,n,r,o){if(o=o===void 0?12:o,e=N(e),t=N(t),n=N(n),r=N(r),o=N(o),w(e,t,n,r,o))return R;if(e<0||t<0||n<0||r<0||[1,2,3,4,5,6,7,8,9,10,11,12].indexOf(o)===-1||r>n)return D;if(t>=e)return 0;const l=(1-Math.pow(t/e,1/n)).toFixed(3),c=e*l*o/12;let h=c,g=0;const u=r===n?n-1:r;for(let i=2;i<=u;i++)g=(e-h)*l,h+=g;return r===1?c:r===n?(e-h)*l:g}function J0(e,t,n,r,o){if(o=o===void 0?2:o,e=N(e),t=N(t),n=N(n),r=N(r),o=N(o),w(e,t,n,r,o))return R;if(e<0||t<0||n<0||r<0||o<=0||r>n)return D;if(t>=e)return 0;let l=0,c=0;for(let h=1;h<=r;h++)c=Math.min((e-l)*(o/n),e-t-l),l+=c;return c}function Z0(e,t,n,r,o){if(e=K(e),t=K(t),n=N(n),r=N(r),o=N(o),o=o||0,w(e,t,n,r,o))return R;if(n<=0||r<=0)return D;if(e>=t)return R;let l,c;switch(o){case 0:l=360,c=qe(e,t,!1);break;case 1:l=365,c=We(e,t,"D");break;case 2:l=360,c=We(e,t,"D");break;case 3:l=365,c=We(e,t,"D");break;case 4:l=360,c=qe(e,t,!0);break;default:return D}return(r-n)/r*l/c}function b0(e,t){if(e=N(e),t=N(t),w(e,t))return R;if(t<0)return D;if(t>=0&&t<1)return fe;t=parseInt(t,10);let n=parseInt(e,10);n+=e%1*Math.pow(10,Math.ceil(Math.log(t)/Math.LN10))/t;const r=Math.pow(10,Math.ceil(Math.log(t)/Math.LN2)+1);return n=Math.round(n*r)/r,n}function _0(e,t){if(e=N(e),t=N(t),w(e,t))return R;if(t<0)return D;if(t>=0&&t<1)return fe;t=parseInt(t,10);let n=parseInt(e,10);return n+=e%1*Math.pow(10,-Math.ceil(Math.log(t)/Math.LN10))*t,n}function j0(e,t){return e=N(e),t=N(t),w(e,t)?R:e<=0||t<1?D:(t=parseInt(t,10),Math.pow(1+e/t,t)-1)}function _e(e,t,n,r,o){if(r=r||0,o=o||0,e=N(e),t=N(t),n=N(n),r=N(r),o=N(o),w(e,t,n,r,o))return R;let l;if(e===0)l=r+n*t;else{const c=Math.pow(1+e,t);l=o===1?r*c+n*(1+e)*(c-1)/e:r*c+n*(c-1)/e}return-l}function ea(e,t){if(e=N(e),t=V(P(t)),w(e,t))return R;const n=t.length;let r=e;for(let o=0;o<n;o++)r*=1+t[o];return r}function Mi(e,t,n,r,o,l){if(o=o||0,l=l||0,e=N(e),t=N(t),n=N(n),r=N(r),o=N(o),l=N(l),w(e,t,n,r,o,l))return R;const c=Rt(e,n,r,o,l);return(t===1?l===1?0:-r:l===1?_e(e,t-2,c,r,1)-c:_e(e,t-1,c,r,0))*e}function ta(e,t){if(t=typeof t=="number"?t:typeof t>"u"?.1:N(t),e=P(e).filter(he),e=V(e),w(e,t))return R;const n=new Float64Array(e.length);let r=!1,o=!1;for(let u=0;u<e.length;u++)n[u]=e[u],n[u]>0&&(r=!0),n[u]<0&&(o=!0);if(!r||!o)return D;const l=u=>{u<=-1&&(u=-.999999999);let i=n[0];const s=1+u;let f=1;for(let a=1;a<n.length;a++)f*=s,i+=n[a]/f;return i},c=new Map,h=function(u){const i=Math.round(u*1e10)/1e10;if(c.has(i))return c.get(i);const s=l(i);return c.set(i,s),s};return function(){let s=t,f=s,a=0;for(;a<1e3;){const v=h(s);if(Math.abs(v)<1e-10)return s;if(a>0&&Math.abs(s-f)<1e-10*10)break;const p=Math.max(1e-4,Math.abs(s*1e-4)),I=(h(s+p)-v)/p;if(Math.abs(I)<1e-10)break;f=s;const A=v/I,M=Math.max(.1,Math.abs(s)*.5);Math.abs(A)>M?s-=Math.sign(A)*M:s-=A,s<=-1&&(s=-.99999999),s>1e3&&(s=1e3),a++}let d=h(s);if(Math.abs(d)<1e-10)return s;let E,m;if(d>0){for(E=s,m=s+.1;h(m)>0&&m<1e3;)m=m*2+.1;if(m>=1e3)return s}else{for(m=s,E=Math.max(-.99999999,s-.1);h(E)<0&&E>-.99999999;)E=Math.max(-.99999999,E-.1);if(E<=-.99999999)return s}let C;for(let v=0;v<1e3;v++){C=(E+m)/2;const p=h(C);if(Math.abs(p)<1e-10||Math.abs(m-E)<1e-10)return C;p*h(E)<0?m=C:E=C}return C}()}function na(e,t,n,r){return e=N(e),t=N(t),n=N(n),r=N(r),w(e,t,n,r)?R:r*e*(t/n-1)}function ra(e,t,n){if(e=V(P(e)),t=N(t),n=N(n),w(e,t,n))return R;const r=e.length,o=[],l=[];for(let g=0;g<r;g++)e[g]<0?o.push(e[g]):l.push(e[g]);const c=-bn(n,l)*Math.pow(1+n,r-1),h=bn(t,o)*(1+t);return Math.pow(c/h,1/(r-1))-1}function ia(e,t){return e=N(e),t=N(t),w(e,t)?R:e<=0||t<1?D:(t=parseInt(t,10),(Math.pow(e+1,1/t)-1)*t)}function oa(e,t,n,r,o){if(o=o===void 0?0:o,r=r===void 0?0:r,e=N(e),t=N(t),n=N(n),r=N(r),o=N(o),w(e,t,n,r,o))return R;if(e===0)return-(n+r)/t;{const l=t*(1+e*o)-r*e,c=n*e+t*(1+e*o);return Math.log(l/c)/Math.log(1+e)}}function bn(){const e=V(P(arguments));if(e instanceof Error)return e;const t=e[0];let n=0;for(let r=1;r<e.length;r++)n+=e[r]/Math.pow(1+t,r);return n}function sa(e,t,n){return e=N(e),t=N(t),n=N(n),w(e,t,n)?R:e<=0?D:(Math.log(n)-Math.log(t))/Math.log(1+e)}function Rt(e,t,n,r,o){if(r=r||0,o=o||0,e=N(e),t=N(t),n=N(n),r=N(r),o=N(o),w(e,t,n,r,o))return R;let l;if(e===0)l=(n+r)/t;else{const c=Math.pow(1+e,t);l=o===1?(r*e/(c-1)+n*e/(1-1/c))/(1+e):r*e/(c-1)+n*e/(1-1/c)}return-l}function ua(e,t,n,r,o,l){return o=o||0,l=l||0,e=N(e),n=N(n),r=N(r),o=N(o),l=N(l),w(e,n,r,o,l)?R:Rt(e,n,r,o,l)-Mi(e,t,n,r,o,l)}function fa(e,t,n,r,o){if(e=K(e),t=K(t),n=N(n),r=N(r),o=N(o),o=o||0,w(e,t,n,r,o))return R;if(n<=0||r<=0)return D;if(e>=t)return R;let l,c;switch(o){case 0:l=360,c=qe(e,t,!1);break;case 1:l=365,c=We(e,t,"D");break;case 2:l=360,c=We(e,t,"D");break;case 3:l=365,c=We(e,t,"D");break;case 4:l=360,c=qe(e,t,!0);break;default:return D}return r-n*r*c/l}function la(e,t,n,r,o){return r=r||0,o=o||0,e=N(e),t=N(t),n=N(n),r=N(r),o=N(o),w(e,t,n,r,o)?R:e===0?-n*t-r:((1-Math.pow(1+e,t))/e*n*(1+e*o)-r)/Math.pow(1+e,t)}function ca(e,t,n,r,o,l){if(l=l===void 0?.1:l,r=r===void 0?0:r,o=o===void 0?0:o,e=N(e),t=N(t),n=N(n),r=N(r),o=N(o),l=N(l),w(e,t,n,r,o,l))return R;const c=1e-10,h=100;let g=l;o=o?1:0;for(let u=0;u<h;u++){if(g<=-1)return D;let i,s;if(Math.abs(g)<c?i=n*(1+e*g)+t*(1+g*o)*e+r:(s=Math.pow(1+g,e),i=n*s+t*(1/g+o)*(s-1)+r),Math.abs(i)<c)return g;let f;if(Math.abs(g)<c)f=n*e+t*o*e;else{s=Math.pow(1+g,e);const a=e*Math.pow(1+g,e-1);f=n*a+t*(1/g+o)*a+t*(-1/(g*g))*(s-1)}g-=i/f}return g}function aa(e,t,n){return e=N(e),t=N(t),n=N(n),w(e,t,n)?R:e===0||t===0?D:Math.pow(n/t,1/e)-1}function ha(e,t,n){return e=N(e),t=N(t),n=N(n),w(e,t,n)?R:n===0?D:(e-t)/n}function ga(e,t,n,r){return e=N(e),t=N(t),n=N(n),r=N(r),w(e,t,n,r)?R:n===0||r<1||r>n?D:(r=parseInt(r,10),(e-t)*(n-r+1)*2/(n*(n+1)))}function da(e,t,n){return e=K(e),t=K(t),n=N(n),w(e,t,n)?R:n<=0||e>t||t-e>365*24*60*60*1e3?D:365*n/(360-n*qe(e,t,!1))}function Ea(e,t,n){return e=K(e),t=K(t),n=N(n),w(e,t,n)?R:n<=0||e>t||t-e>365*24*60*60*1e3?D:100*(1-n*qe(e,t,!1)/360)}function ma(e,t,n){return e=K(e),t=K(t),n=N(n),w(e,t,n)?R:n<=0||e>t||t-e>365*24*60*60*1e3?D:(100-n)*360/(n*qe(e,t,!1))}function pa(e,t,n){if(e=V(P(e)),t=br(P(t)),n=N(n),w(e,t,n))return R;const r=(a,d,E)=>{const m=E+1;let C=a[0];for(let v=1;v<a.length;v++)C+=a[v]/Math.pow(m,Ze(d[v],d[0])/365);return C},o=(a,d,E)=>{const m=E+1;let C=0;for(let v=1;v<a.length;v++){const p=Ze(d[v],d[0])/365;C-=p*a[v]/Math.pow(m,p+1)}return C};let l=!1,c=!1;for(let a=0;a<e.length;a++)e[a]>0&&(l=!0),e[a]<0&&(c=!0);if(!l||!c)return D;n=n||.1;let h=n;const g=1e-10;let u,i,s,f=!0;do s=r(e,t,h),u=h-s/o(e,t,h),i=Math.abs(u-h),h=u,f=i>g&&Math.abs(s)>g;while(f);return h}function Ia(e,t,n){if(e=N(e),t=V(P(t)),n=br(P(n)),w(e,t,n))return R;let r=0;for(let o=0;o<t.length;o++)r+=t[o]/Math.pow(1+e,Ze(n[o],n[0])/365);return r}function Na(){const e=P(arguments);let t=R;for(let n=0;n<e.length;n++){if(e[n]instanceof Error)return e[n];e[n]===void 0||e[n]===null||typeof e[n]=="string"||(t===R&&(t=!0),e[n]||(t=!1))}return t}function Aa(){return!1}function Ca(e,t,n){return e instanceof Error?e:(t=arguments.length>=2?t:!0,t==null&&(t=0),n=arguments.length===3?n:!1,n==null&&(n=0),e?t:n)}function va(){for(let e=0;e<arguments.length/2;e++)if(arguments[e*2])return arguments[e*2+1];return B}function Ta(e,t){return qt(e)?t:e}function Ra(e,t){return e===B?t:e}function Ma(e){return typeof e=="string"?R:e instanceof Error?e:!e}function Oa(){const e=P(arguments);let t=R;for(let n=0;n<e.length;n++){if(e[n]instanceof Error)return e[n];e[n]===void 0||e[n]===null||typeof e[n]=="string"||(t===R&&(t=!1),e[n]&&(t=!0))}return t}function Da(){return!0}function wa(){const e=P(arguments);let t=R;for(let n=0;n<e.length;n++){if(e[n]instanceof Error)return e[n];e[n]===void 0||e[n]===null||typeof e[n]=="string"||(t===R&&(t=0),e[n]&&t++)}return t===R?t:!!(Math.floor(Math.abs(t))&1)}function La(){let e;if(arguments.length>0){const t=arguments[0],n=arguments.length-1,r=Math.floor(n/2);let o=!1;const l=n%2!==0,c=n%2===0?null:arguments[arguments.length-1];if(r){for(let h=0;h<r;h++)if(t===arguments[h*2+1]){e=arguments[h*2+2],o=!0;break}}o||(e=l?c:B)}else e=R;return e}const Oi={errors:ls,symbols:Al,date:hs};var Sa=Object.freeze({__proto__:null,ABS:_u,ACCRINT:K0,ACOS:ju,ACOSH:ef,ACOT:tf,ACOTH:nf,AGGREGATE:rf,AND:Na,ARABIC:of,ASIN:sf,ASINH:uf,ATAN:ff,ATAN2:lf,ATANH:cf,AVEDEV:Iu,AVERAGE:$e,AVERAGEA:Pn,AVERAGEIF:Nu,AVERAGEIFS:Au,BASE:af,BESSELI:Gl,BESSELJ:zl,BESSELK:Yl,BESSELY:Wl,BETA:It,BETADIST:Mc,BETAINV:Oc,BIN2DEC:ql,BIN2HEX:Kl,BIN2OCT:Xl,BINOM:ot,BINOMDIST:Dc,BITAND:$l,BITLSHIFT:xl,BITOR:Ql,BITRSHIFT:Jl,BITXOR:Zl,CEILING:Ye,CEILINGMATH:wc,CEILINGPRECISE:Lc,CHAR:ri,CHIDIST:Sc,CHIDISTRT:Pc,CHIINV:Uc,CHIINVRT:yc,CHISQ:Re,CHITEST:Fc,CHOOSE:ys,CHOOSECOLS:Fs,CHOOSEROWS:Hs,CLEAN:Js,CODE:ii,COLUMN:Vs,COLUMNS:ks,COMBIN:ft,COMBINA:hf,COMPLEX:ae,CONCAT:Zs,CONCATENATE:oi,CONFIDENCE:Un,CONVERT:bl,CORREL:Cu,COS:gf,COSH:df,COT:Ef,COTH:mf,COUNT:Nt,COUNTA:At,COUNTBLANK:ui,COUNTIF:vu,COUNTIFS:Tu,COUPDAYS:X0,COVAR:Hc,COVARIANCE:st,COVARIANCEP:Vc,COVARIANCES:kc,CRITBINOM:Bc,CSC:pf,CSCH:If,CUMIPMT:$0,CUMPRINC:x0,DATE:Tl,DATEDIF:We,DATEVALUE:Rl,DAVERAGE:P0,DAY:Ml,DAYS:Ze,DAYS360:qe,DB:Q0,DCOUNT:U0,DCOUNTA:y0,DDB:J0,DEC2BIN:_l,DEC2HEX:jl,DEC2OCT:ec,DECIMAL:Nf,DEGREES:Af,DELTA:tc,DEVSQ:Ru,DGET:F0,DISC:Z0,DMAX:H0,DMIN:V0,DOLLAR:bs,DOLLARDE:b0,DOLLARFR:_0,DPRODUCT:k0,DROP:Bs,DSTDEV:B0,DSTDEVP:G0,DSUM:z0,DVAR:Y0,DVARP:W0,EDATE:Ol,EFFECT:j0,EOMONTH:Dl,ERF:Ni,ERFC:Ai,ERFCPRECISE:Gc,ERFPRECISE:zc,ERROR:_r,EVEN:Cf,EXACT:_s,EXP:vf,EXPAND:Gs,EXPON:yn,EXPONDIST:Yc,F:Me,FACT:He,FACTDOUBLE:Ei,FALSE:Aa,FDIST:Wc,FDISTRT:qc,FIND:js,FINV:Kc,FINVRT:Xc,FISHER:Mu,FISHERINV:Ou,FIXED:si,FLOOR:Je,FLOORMATH:$c,FLOORPRECISE:xc,FORECAST:fi,FREQUENCY:Du,FTEST:Qc,FV:_e,FVSCHEDULE:ea,GAMMA:Ct,GAMMADIST:Jc,GAMMAINV:Zc,GAMMALN:Fn,GAMMALNPRECISE:bc,GAUSS:wu,GCD:Tf,GEOMEAN:Lu,GESTEP:nc,GROWTH:Su,HARMEAN:Pu,HEX2BIN:rc,HEX2DEC:ic,HEX2OCT:oc,HLOOKUP:zs,HOUR:wl,HSTACK:xs,HYPGEOM:Hn,HYPGEOMDIST:_c,IF:Ca,IFERROR:Ta,IFNA:Ra,IFS:va,IMABS:$n,IMAGINARY:b,IMARGUMENT:xn,IMCONJUGATE:sc,IMCOS:en,IMCOSH:Ci,IMCOT:uc,IMCSC:pc,IMCSCH:Ic,IMDIV:be,IMEXP:fc,IMLN:lc,IMLOG10:cc,IMLOG2:ac,IMPOWER:hc,IMPRODUCT:gc,IMREAL:_,IMSEC:dc,IMSECH:Ec,IMSIN:tn,IMSINH:vi,IMSQRT:mc,IMSUB:Nc,IMSUM:Ac,IMTAN:Cc,INDEX:Ys,INT:Rf,INTERCEPT:Uu,IPMT:Mi,IRR:ta,ISBLANK:Ms,ISERR:jr,ISERROR:qt,ISEVEN:Os,ISLOGICAL:ei,ISNA:Ds,ISNONTEXT:ws,ISNUMBER:Kt,ISO:Mf,ISODD:Ls,ISOWEEKNUM:Ii,ISPMT:na,ISTEXT:ti,KURT:yu,LARGE:li,LCM:Of,LEFT:eu,LEN:tu,LINEST:Vn,LN:Df,LOG:wf,LOG10:Lf,LOGEST:Fu,LOGINV:jc,LOGNORM:ut,LOGNORMDIST:e0,LOGNORMINV:t0,LOOKUP:Ws,LOWER:nu,MATCH:qs,MAX:Xt,MAXA:Hu,MAXIFS:Vu,MEDIAN:ci,MID:ru,MIN:$t,MINA:ku,MINIFS:Bu,MINUTE:Ll,MIRR:ra,MMULT:Sf,MOD:Pf,MODE:xe,MODEMULT:n0,MODESNGL:r0,MONTH:Sl,MROUND:Uf,MULTINOMIAL:yf,MUNIT:Ff,N:Ss,NA:Ps,NEGBINOM:kn,NEGBINOMDIST:i0,NETWORKDAYS:bt,NETWORKDAYSINTL:o0,NOMINAL:ia,NORM:Oe,NORMDIST:s0,NORMINV:u0,NORMSDIST:f0,NORMSINV:l0,NOT:Ma,NOW:Pl,NPER:oa,NPV:bn,NUMBERVALUE:iu,OCT2BIN:vc,OCT2DEC:Tc,OCT2HEX:Rc,ODD:Hf,OR:Oa,PDURATION:sa,PEARSON:ai,PERCENTILE:pe,PERCENTILEEXC:c0,PERCENTILEINC:a0,PERCENTRANK:vt,PERCENTRANKEXC:h0,PERCENTRANKINC:g0,PERMUT:Gu,PERMUTATIONA:zu,PHI:Yu,PI:Vf,PMT:Rt,POISSON:Bn,POISSONDIST:d0,POWER:mi,PPMT:ua,PRICEDISC:fa,PROB:Wu,PRODUCT:Qt,PROPER:ou,PV:la,QUARTILE:Qe,QUARTILEEXC:E0,QUARTILEINC:m0,QUOTIENT:kf,RADIANS:Bf,RAND:Gf,RANDBETWEEN:zf,RANK:Tt,RANKAVG:p0,RANKEQ:I0,RATE:ca,REPLACE:su,REPT:me,RIGHT:uu,ROMAN:Yf,ROUND:pi,ROUNDDOWN:Wf,ROUNDUP:qf,ROW:qu,ROWS:Ks,RRI:aa,RSQ:Ku,SEARCH:fu,SEC:Kf,SECH:Xf,SECOND:Ul,SERIESSUM:$f,SIGN:xf,SIN:Qf,SINH:Jf,SKEW:Gn,SKEWP:N0,SLN:ha,SLOPE:Xu,SMALL:hi,SORT:Xs,SQRT:Zf,SQRTPI:bf,STANDARDIZE:$u,STDEV:Ee,STDEVA:xu,STDEVP:A0,STDEVPA:Qu,STDEVS:C0,STEYX:Ju,SUBSTITUTE:lu,SUBTOTAL:_f,SUM:Ve,SUMIF:jf,SUMIFS:el,SUMPRODUCT:tl,SUMSQ:nl,SUMX2MY2:rl,SUMX2PY2:il,SUMXMY2:ol,SWITCH:La,SYD:ga,T:de,TAN:sl,TANH:ul,TBILLEQ:da,TBILLPRICE:Ea,TBILLYIELD:ma,TDIST:v0,TDISTRT:T0,TEXT:cu,TEXTJOIN:au,TIME:yl,TIMEVALUE:Fl,TINV:R0,TODAY:Hl,TRANSPOSE:$s,TREND:Zu,TRIM:hu,TRIMMEAN:bu,TRUE:Da,TRUNC:fl,TTEST:M0,TYPE:Us,UNICHAR:gu,UNICODE:du,UNIQUE:Sn,UPPER:Eu,VALUE:mu,VAR:ge,VARA:gi,VARP:O0,VARPA:di,VARS:D0,VLOOKUP:ni,VSTACK:Qs,WEEKDAY:Vl,WEEKNUM:kl,WEIBULL:zn,WEIBULLDIST:w0,WORKDAY:_t,WORKDAYINTL:L0,XIRR:pa,XNPV:Ia,XOR:wa,YEAR:Bl,YEARFRAC:Kn,Z:Yn,ZTEST:S0,utils:Oi});const Di=e=>{const t=Sa[e];if(!t)throw J("#NAME?");return t},J=e=>{const{errors:t}=Oi;return{"#GETTING_DATA":t.data,"#DIV/0!":t.div0,"#NULL!":t.nil,"#NUM!":t.num,"#REF!":t.ref,"#VALUE!":t.value,"#N/A":t.na,"#NAME?":t.name}[e]},nn=e=>{let t;return e instanceof Error&&Pr.has(e.message)&&(t=e.message),e instanceof String&&Pr.has(e)&&(t=e),t},wi="";class Pa{list;current=0;start=0;tokens=[];constructor(t){this.list=[...t]}scan(){for(;!this.isAtEnd();)this.start=this.current,this.scanToken();return this.tokens.push(new Te(S.EOF,"")),this.tokens.length>0&&this.tokens[0].type===S.EQUAL&&this.tokens.shift(),this.tokens}peek(){return this.isAtEnd()?wi:this.list[this.current]}match(t){return this.peek()!==t?!1:(this.next(),!0)}next(){return this.isAtEnd()?wi:this.list[this.current++]}isAtEnd(){return this.current>=this.list.length}addToken(t){const n=this.list.slice(this.start,this.current).join("");this.tokens.push(new Te(t,n))}string(t){for(;!this.isAtEnd()&&this.peek()!==t;)this.next();if(this.peek()!==t)throw J("#VALUE!");this.next();const n=this.list.slice(this.start+1,this.current-1).join("");if(t==="\'"&&this.match("!")){this.tokens.push(new Te(S.SHEET_NAME,n));return}this.tokens.push(new Te(S.STRING,n))}getDigits(){for(;!this.isAtEnd()&&Be(this.peek());)this.next()}getAlphas(){for(;!this.isAtEnd()&&vn(this.peek());)this.next()}matchR1C1(){if(this.match("[")){if(this.match("-"),Be(this.peek()))this.getDigits();else throw J("#VALUE!");if(this.peek()!=="]")throw J("#VALUE!");this.next()}else this.getDigits()}matchScientificCounting(){if(this.match("E")||this.match("e")){if(this.match("+")||this.match("-"))return this.getDigits(),this.addToken(S.NUMBER),!0;if(Be(this.peek()))return this.getDigits(),this.addToken(S.NUMBER),!0;throw J("#VALUE!")}return!1}number(){this.getDigits(),!(this.matchScientificCounting()||(this.match(".")&&this.getDigits(),this.matchScientificCounting()))&&this.addToken(S.NUMBER)}addIdentifier(){for(;!this.isAtEnd()&&this.anyChar(this.peek());)this.next();let t=this.list.slice(this.start,this.current).join("");const n=t.toUpperCase();if(Wo.has(n)&&this.match("(")){this.tokens.push(new Te(S.EXCEL_FUNCTION,n));return}if(["IF","CHOOSE"].includes(n)&&this.match("(")){this.tokens.push(new Te(S.REF_FUNCTION_COND,n));return}if(["INDEX","OFFSET","INDIRECT"].includes(n)&&this.match("(")){this.tokens.push(new Te(S.REF_FUNCTION,n));return}if(this.match("!")){this.tokens.push(new Te(S.SHEET_NAME,t));return}let r=S.BOOL;if(n==="TRUE"||n==="FALSE")t=n,r=S.BOOL;else if(nn(n))t=n,r=n==="#REF!"?S.ERROR_REF:S.ERROR;else if(Bo.test(n))r=S.CELL;else if(Go.test(n))r=S.COLUMN;else if(zo.test(n))r=S.ROW;else if(Yo.test(n))r=S.DEFINED_NAME;else throw J("#NAME?");this.tokens.push(new Te(r,t))}scanToken(){const t=this.next();switch(t){case"$":{vn(this.peek())?(this.getAlphas(),this.match("$")?Be(this.peek())?(this.getDigits(),this.addToken(S.CELL)):this.addIdentifier():Be(this.peek())?(this.getDigits(),this.addToken(S.CELL)):this.addToken(S.COLUMN)):Be(this.peek())?(this.getDigits(),this.addToken(S.ROW)):this.addIdentifier();break}case"r":case"R":{if(this.matchR1C1(),this.match("C")||this.match("c")){this.matchR1C1();const n=this.list.slice(this.start,this.current).join("").toUpperCase();this.tokens.push(new Te(S.R1C1,n))}else this.addIdentifier();break}case"(":this.addToken(S.LEFT_BRACKET);break;case")":this.addToken(S.RIGHT_BRACKET);break;case",":this.addToken(S.COMMA);break;case":":this.addToken(S.COLON);break;case Sr:this.addToken(S.EQUAL);break;case"<":this.match(">")?this.addToken(S.NOT_EQUAL):this.match("=")?this.addToken(S.LESS_EQUAL):this.addToken(S.LESS);break;case">":this.match("=")?this.addToken(S.GREATER_EQUAL):this.addToken(S.GREATER);break;case"+":this.addToken(S.PLUS);break;case"-":this.addToken(S.MINUS);break;case"*":this.addToken(S.STAR);break;case"/":this.addToken(S.SLASH);break;case"^":this.addToken(S.EXPONENT);break;case"&":this.addToken(S.CONCATENATE);break;case"%":this.addToken(S.PERCENT);break;case\'"\':this.string(t);break;case"\'":this.string(t);break;case"!":this.addToken(S.EXCLAMATION);break;case";":this.addToken(S.SEMICOLON);break;case"{":this.addToken(S.lEFT_BRACE);break;case"}":this.addToken(S.RIGHT_BRACE);break;case" ":break;case"\\r":case"	":case`\n`:break;default:Be(t)?this.number():this.addIdentifier();break}}anyChar(t){return!\'(),:=<>+-*/^&%"{}!\'.includes(t)&&!this.isWhiteSpace(t)}isWhiteSpace(t){return t===" "||t==="\\r"||t===`\n`||t==="	"}}class Mt{left;right;operator;constructor(t,n,r){this.left=t,this.operator=n,this.right=r}accept(t){return t.visitBinaryExpression(this)}handleConcatenate(t){const n=t.toString();return this.operator.type===S.CONCATENATE&&t instanceof _n&&t.value.type===S.STRING?JSON.stringify(n):n}toString(){const t=this.handleConcatenate(this.left),n=this.handleConcatenate(this.right);return`${t}${this.operator.toString()}${n}`}}class Ua{right;operator;constructor(t,n){this.operator=t,this.right=n}accept(t){return t.visitUnaryExpression(this)}toString(){return this.operator.toString()+this.right.toString()}}class ya{left;operator;constructor(t,n){this.operator=t,this.left=n}accept(t){return t.visitPostUnaryExpression(this)}toString(){return this.left.toString()+this.operator.toString()}}class _n{value;constructor(t){this.value=t}accept(t){return t.visitLiteralExpression(this)}toString(){return this.value.toString()}}class Ot{value;sheetName;constructor(t,n){this.value=t,this.sheetName=n}accept(t){return t.visitCellExpression(this)}toString(){const t=this.value.toString().toUpperCase();if(this.sheetName){let n=this.sheetName.toString();return n.includes(" ")&&(n=`\'${n}\'`),`${n}!${t}`}else return t}}class Fa{name;params;constructor(t,n){this.name=t,this.params=n}accept(t){return t.visitCallExpression(this)}toString(){return`${this.name.toString()}(${this.params.map(t=>t.toString()).join(",")})`}}class Ha{right;operator;left;constructor(t,n,r){this.left=t,this.operator=n,this.right=r}accept(t){return t.visitCellRangeExpression(this)}toString(){return this.left.toString()+this.operator.toString()+this.right.toString()}}class Va{value;constructor(t){this.value=t}accept(t){return t.visitGroupExpression(this)}toString(){return`(${this.value.toString()})`}}class ka{value;constructor(t){this.value=t}accept(t){return t.visitArrayExpression(this)}toString(){return`{${this.value.map(n=>n.toString()).join(",")}}`}}class Ba{tokens;current=0;constructor(t){this.tokens=t}parse(){const t=[];for(;!this.isAtEnd();)t.push(this.expression());return t}expression(){return this.comparison()}comparison(){let t=this.concatenate();for(;this.match(S.EQUAL,S.NOT_EQUAL,S.GREATER,S.GREATER_EQUAL,S.LESS,S.LESS_EQUAL);){const n=this.previous(),r=this.concatenate();t=new Mt(t,n,r)}return t}concatenate(){let t=this.term();for(;this.match(S.CONCATENATE);){const n=this.previous(),r=this.term();t=new Mt(t,n,r)}return t}term(){let t=this.factor();for(;this.match(S.PLUS,S.MINUS);){const n=this.previous(),r=this.factor();t=new Mt(t,n,r)}return t}factor(){let t=this.expo();for(;this.match(S.SLASH,S.STAR);){const n=this.previous(),r=this.expo();t=new Mt(t,n,r)}return t}expo(){let t=this.unary();for(;this.match(S.EXPONENT);){const n=this.previous(),r=this.unary();t=new Mt(t,n,r)}return t}unary(){if(this.match(S.PLUS,S.MINUS)){const t=this.previous(),n=this.unary();return new Ua(t,n)}return this.postUnary()}postUnary(){let t=this.cellRange();if(this.match(S.PERCENT)){const n=this.previous();t=new ya(n,t)}return t}cellRange(){let t=this.sheetRange();for(;this.match(S.COLON);){const n=this.previous(),r=this.sheetRange(),o=this.convertToCellExpression(r),l=this.convertToCellExpression(t);t=new Ha(l,n,o)}return t}convertToCellExpression(t){if(t instanceof Ot)return t;if(t instanceof _n&&t.value.type===S.NUMBER)return new Ot(t.value,void 0);throw J("#NAME?")}sheetRange(){if(this.match(S.SHEET_NAME)){const t=this.previous();let n=this.call();if(n instanceof Ot)return new Ot(n.value,t);throw J("#REF!")}return this.call()}call(){for(;this.match(S.EXCEL_FUNCTION,S.REF_FUNCTION,S.REF_FUNCTION_COND);){const t=this.previous(),n=[];if(!this.check(S.RIGHT_BRACKET))do{if(this.peek().type==S.RIGHT_BRACKET)break;n.push(this.expression())}while(this.match(S.COMMA));return this.expect(S.RIGHT_BRACKET),new Fa(t,n)}return this.primary()}primary(){if(this.match(S.lEFT_BRACE)){const t=[];do{if(this.peek().type==S.RIGHT_BRACE)break;t.push(this.expression())}while(this.match(S.COMMA));return this.expect(S.RIGHT_BRACE),new ka(t)}if(this.match(S.LEFT_BRACKET)){const t=this.expression();return this.expect(S.RIGHT_BRACKET),new Va(t)}if(this.match(S.NUMBER,S.STRING,S.BOOL))return new _n(this.previous());if(this.match(S.CELL,S.COLUMN,S.ROW,S.DEFINED_NAME,S.R1C1)){const t=this.previous();return new Ot(t,void 0)}throw this.match(S.ERROR,S.ERROR_REF)?J(this.previous().value):J("#VALUE!")}match(...t){const{type:n}=this.peek();return t.includes(n)?(this.next(),!0):!1}previous(){return this.tokens[this.current-1]}check(t){return this.peek().type===t}expect(t){if(this.check(t))return this.next(),this.previous();throw J("#VALUE!")}next(){this.current++}isAtEnd(){return this.peek().type===S.EOF}peek(){return this.current<this.tokens.length?this.tokens[this.current]:new Te(S.EOF,"")}}class Ga{expressions;cellDataMap;currentCoord;constructor(t,n,r){this.expressions=t,this.cellDataMap=r,this.currentCoord=n}interpret(){let t=[];for(const n of this.expressions){const r=this.evaluate(n),o=this.getRangeCellValue(r);Array.isArray(o)?t=t.concat(o):t.push(o)}return t}visitBinaryExpression(t){let n=this.evaluate(t.left),r=this.evaluate(t.right);switch(n=this.getOneValue(n),r=this.getOneValue(r),[S.MINUS,S.PLUS,S.SLASH,S.STAR,S.EXPONENT].includes(t.operator.type)&&(n=this.parseNumber(n),r=this.parseNumber(r)),t.operator.type){case S.MINUS:return n-r;case S.PLUS:return n+r;case S.SLASH:{if(r===0)throw J("#DIV/0!");return n/r}case S.STAR:return n*r;case S.EXPONENT:return Math.pow(n,r);case S.EQUAL:return n===r;case S.NOT_EQUAL:return n!==r;case S.GREATER:return n>r;case S.GREATER_EQUAL:return n>=r;case S.LESS:return n<r;case S.LESS_EQUAL:return n<=r;case S.CONCATENATE:return`${n}${r}`;default:throw J("#GETTING_DATA")}}visitCallExpression(t){const n=this.cellDataMap.getFunction(t.name.value);if(n&&typeof n=="function"){let r=[];for(const c of t.params){const h=this.evaluate(c);if(h instanceof Le){const g=this.getCellValue(h,!1);r=r.concat(g)}else r.push(h)}const o=n(...r),l=nn(o);if(l)throw J(l);return o}throw J("#NAME?")}visitCellExpression(t){let n="";if(t.sheetName){const o=this.cellDataMap.getSheetInfo("",t.sheetName.value);if(!o?.sheetId)throw J("#REF!");n=o.sheetId}let r;if([S.DEFINED_NAME,S.CELL,S.COLUMN,S.R1C1].includes(t.value.type)&&(r=this.cellDataMap.getDefinedName(t.value.value)),r||(t.value.type===S.R1C1?r=rs(t.value.value,this.currentCoord):r=ns(t.value.value)),!r)throw J("#REF!");return n&&(r.sheetId=n),r}visitLiteralExpression(t){const{type:n,value:r}=t.value;switch(n){case S.STRING:return r;case S.NUMBER:{const[o,l]=yr(r);if(o)return l;throw J("#VALUE!")}case S.BOOL:return r==="TRUE";default:throw J("#VALUE!")}}visitUnaryExpression(t){const n=this.evaluate(t.right),r=this.parseNumber(this.getOneValue(n));switch(t.operator.type){case S.MINUS:return-r;case S.PLUS:return r;default:throw J("#VALUE!")}}visitCellRangeExpression(t){switch(t.operator.type){case S.COLON:{const n=this.visitCellExpression(t.left),r=this.visitCellExpression(t.right),o=Gr(n,r);if(!o)throw J("#NAME?");return o}default:throw J("#NAME?")}}visitGroupExpression(t){return this.evaluate(t.value)}visitArrayExpression(t){const n=[];for(const r of t.value)n.push(this.evaluate(r));return n}visitPostUnaryExpression(t){const n=this.evaluate(t.left),r=this.parseNumber(this.getOneValue(n));if(t.operator.type===S.PERCENT)return r*.01;throw J("#VALUE!")}evaluate(t){return t.accept(this)}getRangeCellValue(t){if(t instanceof Le){if(t.colCount===t.rowCount&&t.colCount===1)return this.getCellValue(t,!0);throw J("#REF!")}return t}getOneValue(t){const n=this.getRangeCellValue(t);return Array.isArray(n)?n[0]:n}parseNumber(t){const[n,r]=yr(t);if(!n)throw J("#VALUE!");return r}getCellValue(t,n){let r=[];if(n){const g=this.cellDataMap.getCell(t),u=this.cellDataMap.handleCell(g,{row:t.row,col:t.col,sheetId:t.sheetId});return r=r.concat(u),r}const{row:o,col:l,sheetId:c}=t,h=this.cellDataMap.getSheetInfo(c);if(!h||o>=h.rowCount||l>=h.colCount)throw J("#REF!");return es(t,h.rowCount,h.colCount,(g,u)=>{const i=this.cellDataMap.getCell({row:g,col:u,rowCount:1,colCount:1,sheetId:c}),s=this.cellDataMap.handleCell(i,{row:g,col:u,sheetId:c});return r=r.concat(s),!1}),r}}function Li(e,t={row:0,col:0,sheetId:""},n=new za,r=new Map){if(r.has(e)){const o=r.get(e);return o.result.length===0?{result:[0]}:o}try{n.handleCell=(u,i)=>u?u.formula?Li(u.formula,i,n,r).result:[u.value]:[];const o={result:[],expressionStr:""};r.set(e,o);const l=new Pa(e).scan(),c=new Ba(l).parse(),h=new Ga(c,t,n).interpret(),g=[];for(const u of c)g.push(u.toString());return o.expressionStr=g.join(""),o.result=h,o}catch(o){const l=nn(o);if(l){const c={result:[l]};return r.set(e,c),c}throw o}}const jn="_test_";class za{map=new Map;definedNameMap=new Map;currentSheetId=jn;sheetList=[{sheetId:jn,name:"_test_",rowCount:200,colCount:30,isHide:!1,sort:1}];getKey(t,n,r){return`${t}_${n}_${r||jn}`}getFunction=Di;setSheetList(t){this.sheetList=t}set(t,n){const{row:r,col:o,sheetId:l}=t;for(let c=0;c<n.length;c++)for(let h=0;h<n[c].length;h++){const g=this.getKey(r+c,o+h,l),u=n[c][h];typeof u=="string"&&Xo(u)?this.map.set(g,{formula:u,value:""}):this.map.set(g,{value:u})}}getCell(t){const{row:n,col:r,sheetId:o}=t,l=this.getKey(n,r,o);return this.map.get(l)}getSheetInfo(t,n){if(n)return this.sheetList.find(o=>o.name===n);const r=t||this.currentSheetId;return this.sheetList.find(o=>o.sheetId===r)}handleCell=(t,n)=>t?[t.value]:[];setDefinedName(t,n){const r=Le.makeRange(n);this.definedNameMap.set(t,r)}getDefinedName(t){return this.definedNameMap.get(t)}}function Ya(e,t){const{currentSheetId:n,worksheets:r,workbook:o,definedNames:l}=e,c=new Map,h={list:[]},g={handleCell:()=>[],getFunction:Di,getCell:u=>{const{row:i,col:s,sheetId:f}=u,a=Fr(f||n,i,s);return r[a]},set:()=>{throw J("#REF!")},getSheetInfo:(u,i)=>{if(i)return o.find(f=>f.name===i);const s=u||n;return o.find(f=>f.sheetId===s)},setDefinedName:()=>{throw J("#REF!")},getDefinedName:u=>{const i=l[u];return i&&Le.makeRange(i)}};for(const[u,i]of Object.entries(r)){const s=Ko(u);if(!s||s.sheetId!==n||!i?.formula)continue;const f=Li(i?.formula,s,g,c);if(!f)continue;const a=f.result[0];a!==i.value&&h.list.push({key:u,newValue:a,sheetId:n})}return t(h)}const rn=2,on=2**2,er=2**3,sn=2**4,un=2**5,fn=2**6,tr=2**7,nr=2**8,rr=2**9,Wa=0,qa=2958466,Ka=-694324,Xa=35830291,$a=-1,Si=1,lt=6,Pi="general",Ui="hash",Dt="zero",yi="qmark",Fi="slash",ln="group",ir="scale",Hi="comma",Vi="break",or="text",ki="plus",Bi="minus",sr="point",Gi="space",zi="percent",ur="digit",Yi="calendar",cn="error",Wi="datetime",qi="duration",Ki="condition",Xi="dbnum",$i="natnum",xi="locale",Qi="color",xa="modifier",Ji="ampm",Zi="escaped",bi="string",_i="skip",ji="exp",eo="fill",to="paren",fr="char";var lr=Object.freeze({1078:"af",1052:"sq",1118:"am",5121:"ar_DZ",15361:"ar_BH",3073:"ar_EG",2049:"ar_IQ",11265:"ar_JO",13313:"ar_KW",12289:"ar_LB",4097:"ar_LY",6145:"ar_MA",8193:"ar_OM",16385:"ar_QA",1025:"ar_SA",10241:"ar_SY",7169:"ar_TN",14337:"ar_AE",9217:"ar_YE",1067:"hy",1101:"as",2092:"az_AZ",1068:"az_AZ",1069:"eu",1059:"be",2117:"bn",1093:"bn_IN",5146:"bs",1026:"bg",1109:"my",1027:"ca",2052:"zh_CN",3076:"zh_HK",5124:"zh_MO",4100:"zh_SG",1028:"zh_TW",1050:"hr",1029:"cs",1030:"da",1125:"dv",2067:"nl_BE",1043:"nl_NL",1126:"bin",3081:"en_AU",10249:"en_BZ",4105:"en_CA",9225:"en_CB",2057:"en_GB",16393:"en_IN",6153:"en_IE",8201:"en_JM",5129:"en_NZ",13321:"en_PH",7177:"en_ZA",11273:"en_TT",1033:"en_US",12297:"en_ZW",1061:"et",1071:"mk",1080:"fo",1065:"fa",1124:"fil",1035:"fi",2060:"fr_BE",11276:"fr_CM",3084:"fr_CA",9228:"fr_CG",12300:"fr_CI",1036:"fr_FR",5132:"fr_LU",13324:"fr_ML",6156:"fr_MC",14348:"fr_MA",10252:"fr_SN",4108:"fr_CH",7180:"fr",1122:"fy_NL",2108:"gd_IE",1084:"gd",1110:"gl",1079:"ka",3079:"de_AT",1031:"de_DE",5127:"de_LI",4103:"de_LU",2055:"de_CH",1032:"el",1140:"gn",1095:"gu",1279:"en",1037:"he",1081:"hi",1038:"hu",1039:"is",1136:"ig_NG",1057:"id",1040:"it_IT",2064:"it_CH",1041:"ja",1099:"kn",1120:"ks",1087:"kk",1107:"km",1111:"kok",1042:"ko",1088:"ky",1108:"lo",1142:"la",1062:"lv",1063:"lt",2110:"ms_BN",1086:"ms_MY",1100:"ml",1082:"mt",1112:"mni",1153:"mi",1102:"mr",1104:"mn",2128:"mn",1121:"ne",1044:"no_NO",2068:"no_NO",1096:"or",1045:"pl",1046:"pt_BR",2070:"pt_PT",1094:"pa",1047:"rm",2072:"ro_MO",1048:"ro_RO",1049:"ru",2073:"ru_MO",1083:"se",1103:"sa",3098:"sr_SP",2074:"sr_SP",1072:"st",1074:"tn",1113:"sd",1115:"si",1051:"sk",1060:"sl",1143:"so",1070:"sb",11274:"es_AR",16394:"es_BO",13322:"es_CL",9226:"es_CO",5130:"es_CR",7178:"es_DO",12298:"es_EC",17418:"es_SV",4106:"es_GT",18442:"es_HN",2058:"es_MX",19466:"es_NI",6154:"es_PA",15370:"es_PY",10250:"es_PE",20490:"es_PR",1034:"es_ES",14346:"es_UY",8202:"es_VE",1089:"sw",2077:"sv_FI",1053:"sv_SE",1114:"syc",1064:"tg",1097:"ta",1092:"tt",1098:"te",1054:"th",1105:"bo",1073:"ts",1055:"tr",1090:"tk",1058:"uk",1056:"ur",2115:"uz_UZ",1091:"uz_UZ",1075:"ve",1066:"vi",1106:"cy",1076:"xh",1085:"yi",1077:"zu"});const no=/^([a-z\\d]+)(?:[_-]([a-z\\d]+))?(?:\\.([a-z\\d]+))?(?:@([a-z\\d]+))?$/i,ct={},T=(e,t="")=>e.replace(/~/g,t).split(";"),j=(e,t=0,n=0)=>(e.mmm||(e.mmm=t<1?e.mmmm.concat():e.mmmm.map(r=>{const o=r.slice(0,t%10);return o+(t<10||r===o?"":".")})),e.ddd||(e.ddd=n<1?e.dddd.concat():e.dddd.map(r=>{const o=r.slice(0,n%10);return o+(n<10||r===o?"":".")})),!e.mmm6&&e.mmmm6&&(e.mmm6=e.mmmm6),e),Qa={group:" ",decimal:".",positive:"+",negative:"-",percent:"%",exponent:"E",nan:"NaN",infinity:"∞",ampm:T("AM;PM"),mmmm6:T("Muharram;Safar;Rabiʻ I;Rabiʻ II;Jumada I;Jumada II;Rajab;Shaʻban;Ramadan;Shawwal;Dhuʻl-Qiʻdah;Dhuʻl-Hijjah"),mmm6:T("Muh.;Saf.;Rab. I;Rab. II;Jum. I;Jum. II;Raj.;Sha.;Ram.;Shaw.;Dhuʻl-Q.;Dhuʻl-H."),mmmm:T("January;February;March;April;May;June;July;August;September;October;November;December"),mmm:T("Jan;Feb;Mar;Apr;May;Jun;Jul;Aug;Sep;Oct;Nov;Dec"),dddd:T("Sunday;Monday;Tuesday;Wednesday;Thursday;Friday;Saturday"),ddd:T("Sun;Mon;Tue;Wed;Thu;Fri;Sat"),bool:T("TRUE;FALSE"),preferMDY:!1};function ro(e){const t=no.exec(e);if(!t)throw new SyntaxError(`Malformed locale: ${e}`);return{lang:t[1]+(t[2]?"_"+t[2]:""),language:t[1],territory:t[2]||""}}function cr(e){if(typeof e=="number")return lr[e&65535]||null;const t=parseInt(e,16);return isFinite(t)&&lr[t&65535]?lr[t&65535]||null:no.test(e)?e:null}function Ja(e){const t=cr(e);let n=null;if(t){const r=ro(t);n=ct[r.lang]||ct[r.language]||null}return n}function ar(e){return Object.assign({},Qa,e)}function k(e,t){const n=typeof t=="object"?t:ro(t);return ct[n.lang]=ar(e),n.language!==n.lang&&!ct[n.language]&&(ct[n.language]=ar(e)),ct[n.lang]}const an=ar({group:",",preferMDY:!0});an.isDefault=!0,k({group:",",ampm:T("上午;下午"),mmmm:T("一月;二月;三月;四月;五月;六月;七月;八月;九月;十月;十一月;十二月"),mmm:T("1月;2月;3月;4月;5月;6月;7月;8月;9月;10月;11月;12月"),dddd:T("~日;~一;~二;~三;~四;~五;~六","星期"),ddd:T("周日;周一;周二;周三;周四;周五;周六")},"zh_CN");const hr={group:",",ampm:T("上午;下午"),mmmm:T("1月;2月;3月;4月;5月;6月;7月;8月;9月;10月;11月;12月"),mmm:T("1月;2月;3月;4月;5月;6月;7月;8月;9月;10月;11月;12月"),dddd:T("~日;~一;~二;~三;~四;~五;~六","星期"),ddd:T("周日;周一;周二;周三;周四;周五;周六")};k({...hr,nan:"非數值",dddd:T("~日;~一;~二;~三;~四;~五;~六","星期")},"zh_TW"),k({...hr,dddd:T("~日;~一;~二;~三;~四;~五;~六","星期")},"zh_HK"),k({...hr,ampm:T("午前;午後"),dddd:T("日~;月~;火~;水~;木~;金~;土~","曜日"),ddd:T("日;月;火;水;木;金;土")},"ja"),k({group:",",ampm:T("오전;오후"),mmmm:T("1월;2월;3월;4월;5월;6월;7월;8월;9월;10월;11월;12월"),mmm:T("1월;2월;3월;4월;5월;6월;7월;8월;9월;10월;11월;12월"),dddd:T("일요일;월요일;화요일;수요일;목요일;금요일;토요일"),ddd:T("일;월;화;수;목;금;토")},"ko"),k({group:",",ampm:T("ก่อนเที่ยง;หลังเที่ยง"),mmmm:T("มกร~;กุมภาพันธ์;มีน~;เมษายน;พฤษภ~;มิถุนายน;กรกฎ~;สิงห~;กันยายน;ตุล~;พฤศจิกายน;ธันว~","าคม"),mmm:T("ม.ค.;ก.พ.;มี.ค.;เม.ย.;พ.ค.;มิ.ย.;ก.ค.;ส.ค.;ก.ย.;ต.ค.;พ.ย.;ธ.ค."),dddd:T("วันอาทิตย์;วันจันทร์;วันอังคาร;วันพุธ;วันพฤหัสบดี;วันศุกร์;วันเสาร์"),ddd:T("อา.;จ.;อ.;พ.;พฤ.;ศ.;ส.")},"th"),k(j({decimal:",",ampm:T("dop.;odp."),mmmm:T("ledna;února;března;dubna;května;června;července;srpna;září;října;listopadu;prosince"),mmm:T("I;II;III;IV;V;VI;VII;VIII;IX;X;XI;XII"),dddd:T("neděle;pondělí;úterý;středa;čtvrtek;pátek;sobota"),bool:T("PRAVDA;NEPRAVDA")},-1,2),"cs"),k(j({group:".",decimal:",",mmmm:T("januar;februar;marts;april;maj;juni;juli;august;september;oktober;november;december"),dddd:T("søn~;man~;tirs~;ons~;tors~;fre~;lør~","dag"),bool:T("SAND;FALSK")},13,13),"da"),k(j({group:".",decimal:",",ampm:T("a.m.;p.m."),mmmm:T("januari;februari;maart;april;mei;juni;juli;augustus;september;oktober;november;december"),mmm:T("jan.;feb.;mrt.;apr.;mei;jun.;jul.;aug.;sep.;okt.;nov.;dec."),dddd:T("zondag;maandag;dinsdag;woensdag;donderdag;vrijdag;zaterdag"),bool:T("WAAR;ONWAAR")},-1,2),"nl"),k({group:",",preferMDY:!0},"en"),k({group:",",preferMDY:!0},"en_US"),k({group:","},"en_AU"),k({group:","},"en_CA"),k({group:","},"en_GB"),k({group:",",mmm:T("Jan;Feb;Mar;Apr;May;Jun;Jul;Aug;Sept;Oct;Nov;Dec")},"en_IE"),k(j({decimal:",",nan:"epäluku",ampm:T("ap.;ip."),mmmm:T("tammi~;helmi~;maalis~;huhti~;touko~;kesä~;heinä~;elo~;syys~;loka~;marras~;joulu~","kuuta"),mmm:T("tammik.;helmik.;maalisk.;huhtik.;toukok.;kesäk.;heinäk.;elok.;syysk.;lokak.;marrask.;jouluk."),dddd:T("sunnun~;maanan~;tiis~;keskiviikkona;tors~;perjan~;lauan~","taina"),bool:T("TOSI;EPÄTOSI")},-1,2),"fi");const gr=j({group:" ",decimal:",",mmmm:T("janvier;février;mars;avril;mai;juin;juillet;août;septembre;octobre;novembre;décembre"),mmm:T("janv.;févr.;mars;avr.;mai;juin;juil.;août;sept.;oct.;nov.;déc."),dddd:T("~manche;lun~;mar~;mercre~;jeu~;vendre~;same~","di"),bool:T("VRAI;FAUX")},-1,13);k({...gr},"fr"),k({...gr,mmm:T("janv.;févr.;mars;avr.;mai;juin;juill.;août;sept.;oct.;nov.;déc.")},"fr_CA"),k({group:"\'",decimal:".",...gr},"fr_CH");const io=j({mmmm:T("Januar;Februar;März;April;Mai;Juni;Juli;August;September;Oktober;November;Dezember"),mmm:T("Jan.;Feb.;März;Apr.;Mai;Juni;Juli;Aug.;Sept.;Okt.;Nov.;Dez."),dddd:T("Sonn~;Mon~;Diens~;Mittwoch;Donners~;Frei~;Sams~","tag"),bool:T("WAHR;FALSCH")},-1,12);k({group:".",decimal:",",...io},"de"),k({group:"\'",decimal:".",...io},"de_CH"),k(j({group:".",decimal:",",ampm:T("π.μ.;μ.μ."),mmmm:T("Ιανουαρ~;Φεβρουαρ~;Μαρτ~;Απριλ~;Μαΐου;Ιουν~;Ιουλ~;Αυγούστου;Σεπτεμβρ~;Οκτωβρ~;Νοεμβρ~;Δεκεμβρ~","ίου"),mmm:T("Ιαν;Φεβ;Μαρ;Απρ;Μαΐ;Ιουν;Ιουλ;Αυγ;Σεπ;Οκτ;Νοε;Δεκ"),dddd:T("Κυριακή;Δευτέρα;Τρίτη;Τετάρτη;Πέμπτη;Παρασκευή;Σάββατο")},-1,3),"el"),k({decimal:",",ampm:T("de.;du."),mmmm:T("január;február;március;április;május;június;július;augusztus;szeptember;október;november;december"),mmm:T("jan.;febr.;márc.;ápr.;máj.;jún.;júl.;aug.;szept.;okt.;nov.;dec."),dddd:T("vasárnap;hétfő;kedd;szerda;csütörtök;péntek;szombat"),ddd:T("V;H;K;Sze;Cs;P;Szo"),bool:T("IGAZ;HAMIS")},"hu"),k(j({group:".",decimal:",",ampm:T("f.h.;e.h."),mmmm:T("janúar;febrúar;mars;apríl;maí;júní;júlí;ágúst;september;október;nóvember;desember"),dddd:T("sunnu~;mánu~;þriðju~;miðviku~;fimmtu~;föstu~;laugar~","dagur")},13,13),"is"),k(j({group:".",decimal:",",mmmm:T("Januari;Februari;Maret;April;Mei;Juni;Juli;Agustus;September;Oktober;November;Desember"),dddd:T("Minggu;Senin;Selasa;Rabu;Kamis;Jumat;Sabtu")},3,3),"id");const oo=j({mmmm:T("gennaio;febbraio;marzo;aprile;maggio;giugno;luglio;agosto;settembre;ottobre;novembre;dicembre"),dddd:T("domenica;lunedì;martedì;mercoledì;giovedì;venerdì;sabato"),bool:T("VERO;FALSO")},3,3);k({group:".",decimal:",",...oo},"it"),k({group:"\'",decimal:".",...oo},"it_CH");const so={decimal:",",ampm:T("a.m.;p.m."),mmmm:T("januar;februar;mars;april;mai;juni;juli;august;september;oktober;november;desember"),mmm:T("jan.;feb.;mar.;apr.;mai;jun.;jul.;aug.;sep.;okt.;nov.;des."),dddd:T("søn~;man~;tirs~;ons~;tors~;fre~;lør~","dag"),bool:T("SANN;USANN")};k(j({...so},-1,13),"nb"),k(j({...so},-1,13),"no"),k(j({decimal:",",mmmm:T("stycznia;lutego;marca;kwietnia;maja;czerwca;lipca;sierpnia;września;października;listopada;grudnia"),dddd:T("niedziela;poniedziałek;wtorek;środa;czwartek;piątek;sobota"),ddd:T("niedz.;pon.;wt.;śr.;czw.;pt.;sob."),bool:T("PRAWDA;FAŁSZ")},3,-1),"pl");const uo={group:".",decimal:",",mmmm:T("janeiro;fevereiro;março;abril;maio;junho;julho;agosto;setembro;outubro;novembro;dezembro"),dddd:T("domingo;segunda-feira;terça-feira;quarta-feira;quinta-feira;sexta-feira;sábado"),bool:T("VERDADEIRO;FALSO")};k(j(uo,13,13),"pt"),k(j(uo,13,13),"pt_BR"),k({decimal:",",nan:"не число",mmmm:T("января;февраля;марта;апреля;мая;июня;июля;августа;сентября;октября;ноября;декабря"),mmm:T("янв.;февр.;мар.;апр.;мая;июн.;июл.;авг.;сент.;окт.;нояб.;дек."),dddd:T("воскресенье;понедельник;вторник;среда;четверг;пятница;суббота"),ddd:T("вс;пн;вт;ср;чт;пт;сб"),mmmm6:T("рамадан;шавваль;зуль-каада;зуль-хиджжа;мухаррам;раби-уль-авваль;раби-уль-ахир;джумад-уль-авваль;джумад-уль-ахир;раджаб;шаабан;рамадан"),mmm6:T("рам.;шав.;зуль-к.;зуль-х.;мух.;раб. I;раб. II;джум. I;джум. II;радж.;шааб.;рам."),bool:T("ИСТИНА;ЛОЖЬ")},"ru"),k(j({decimal:",",mmmm:T("januára;februára;marca;apríla;mája;júna;júla;augusta;septembra;októbra;novembra;decembra"),dddd:T("nedeľa;pondelok;utorok;streda;štvrtok;piatok;sobota")},3,2),"sk");const Se={group:".",decimal:",",ampm:T("a. m.;p. m."),mmmm:T("enero;febrero;marzo;abril;mayo;junio;julio;agosto;septiem~;octu~;noviem~;diciem~","bre"),mmm:T("ene;feb;mar;abr;may;jun;jul;ago;sept;oct;nov;dic"),dddd:T("domingo;lunes;martes;miércoles;jueves;viernes;sábado"),ddd:T("dom;lun;mar;mié;jue;vie;sáb"),bool:T("VERDADERO;FALSO")},Za=T("ene;feb;mar;abr;may;jun;jul;ago;sep;oct;nov;dic"),dr=T("ene.;feb.;mar.;abr.;may.;jun.;jul.;ago.;sept.;oct.;nov.;dic.");k({...Se},"es"),k({...Se},"es_AR"),k({...Se},"es_BO"),k({...Se},"es_CL"),k({...Se},"es_CO"),k({...Se},"es_EC"),k({...Se,mmm:Za,ampm:T("a.m.;p.m.")},"es_MX"),k({...Se,mmm:dr},"es_PY"),k({...Se,mmm:dr},"es_UY"),k({...Se,mmm:dr,mmmm:T("enero;febrero;marzo;abril;mayo;junio;julio;agosto;setiembre;octubre;noviembre;diciembre")},"es_VE"),k({decimal:",",ampm:T("fm;em"),mmmm:T("januari;februari;mars;april;maj;juni;juli;augusti;september;oktober;november;december"),mmm:T("jan.;feb.;mars;apr.;maj;juni;juli;aug.;sep.;okt.;nov.;dec."),dddd:T("sön~;mån~;tis~;ons~;tors~;fre~;lör~","dag"),ddd:T("sön;mån;tis;ons;tors;fre;lör")},"sv"),k(j({group:".",decimal:",",ampm:T("ÖÖ;ÖS"),mmmm:T("Ocak;Şubat;Mart;Nisan;Mayıs;Haziran;Temmuz;Ağustos;Eylül;Ekim;Kasım;Aralık"),mmm:T("Oca;Şub;Mar;Nis;May;Haz;Tem;Ağu;Eyl;Eki;Kas;Ara"),dddd:T("Pazar;Pazartesi;Salı;Çarşamba;Perşembe;Cuma;Cumartesi"),ddd:T("Paz;Pzt;Sal;Çar;Per;Cum;Cmt"),bool:T("DOĞRU;YANLIŞ")},3,-1),"tr"),k({group:",",ampm:T("yb;yh"),mmmm:T("Ionawr;Chwefror;Mawrth;Ebrill;Mai;Mehefin;Gorffennaf;Awst;Medi;Hydref;Tachwedd;Rhagfyr"),mmm:T("Ion;Chwef;Maw;Ebr;Mai;Meh;Gorff;Awst;Medi;Hyd;Tach;Rhag"),dddd:T("Dydd Sul;Dydd Llun;Dydd Mawrth;Dydd Mercher;Dydd Iau;Dydd Gwener;Dydd Sadwrn"),ddd:T("Sul;Llun;Maw;Mer;Iau;Gwen;Sad")},"cy"),k({group:".",decimal:",",mmmm:T("yanvar;fevral;mart;aprel;may;iyun;iyul;avqust;sentyabr;oktyabr;noyabr;dekabr"),mmm:T("yan;fev;mar;apr;may;iyn;iyl;avq;sen;okt;noy;dek"),dddd:T("bazar;bazar ertəsi;çərşənbə axşamı;çərşənbə;cümə axşamı;cümə;şənbə"),ddd:T("B.;B.e.;Ç.a.;Ç.;C.a.;C.;Ş.")},"az"),k(j({decimal:",",mmmm:T("студзеня;лютага;сакавіка;красавіка;мая;чэрвеня;ліпеня;жніўня;верасня;кастрычніка;лістапада;снежня"),dddd:T("нядзеля;панядзелак;аўторак;серада;чацвер;пятніца;субота"),ddd:T("нд;пн;аў;ср;чц;пт;сб")},3,-1),"be"),k({decimal:",",ampm:T("пр.об.;сл.об."),mmmm:T("януари;февруари;март;април;май;юни;юли;август;септември;октомври;ноември;декември"),mmm:T("яну;фев;март;апр;май;юни;юли;авг;сеп;окт;ное;дек"),dddd:T("неделя;понеделник;вторник;сряда;четвъртък;петък;събота"),ddd:T("нд;пн;вт;ср;чт;пт;сб"),bool:T("ИСТИНА;ЛОЖЬ")},"bg"),k({group:".",decimal:",",mmmm:T("de gener;de febrer;de març;d’abril;de maig;de juny;de juliol;d’agost;de setembre;d’octubre;de novembre;de desembre"),mmm:T("de gen.;de febr.;de març;d’abr.;de maig;de juny;de jul.;d’ag.;de set.;d’oct.;de nov.;de des."),dddd:T("diumenge;dilluns;dimarts;dimecres;dijous;divendres;dissabte"),ddd:T("dg.;dl.;dt.;dc.;dj.;dv.;ds."),ampm:T("a. m.;p. m.")},"ca"),k(j({group:",",decimal:".",mmmm:T("Enero;Pebrero;Marso;Abril;Mayo;Hunyo;Hulyo;Agosto;Setyembre;Oktubre;Nobyembre;Disyembre"),dddd:T("Linggo;Lunes;Martes;Miyerkules;Huwebes;Biyernes;Sabado")},3,3),"fil"),k({group:",",decimal:".",mmmm:T("જાન્યુઆરી;ફેબ્રુઆરી;માર્ચ;એપ્રિલ;મે;જૂન;જુલાઈ;ઑગસ્ટ;સપ્ટેમ્બર;ઑક્ટોબર;નવેમ્બર;ડિસેમ્બર"),mmm:T("જાન્યુ;ફેબ્રુ;માર્ચ;એપ્રિલ;મે;જૂન;જુલાઈ;ઑગસ્ટ;સપ્ટે;ઑક્ટો;નવે;ડિસે"),dddd:T("રવિ~;સોમ~;મંગળ~;બુધ~;ગુરુ~;શુક્ર~;શનિ~","વાર"),ddd:T("રવિ;સોમ;મંગળ;બુધ;ગુરુ;શુક્ર;શનિ")},"gu"),k({group:",",decimal:".",ampm:T("לפנה״צ;אחה״צ"),dddd:T("~ראשון;~שני;~שלישי;~רביעי;~חמישי;~שישי;~שבת","יום "),ddd:T("~א׳;~ב׳;~ג׳;~ד׳;~ה׳;~ו׳;שבת","יום "),mmmm:T("ינואר;פברואר;מרץ;אפריל;מאי;יוני;יולי;אוגוסט;ספטמבר;אוקטובר;נובמבר;דצמבר"),mmm:T("ינו׳;פבר׳;מרץ;אפר׳;מאי;יוני;יולי;אוג׳;ספט׳;אוק׳;נוב׳;דצמ׳"),mmmm6:T("רמדאן;שוואל;ד׳ו אל־קעדה;ד׳ו אל־חיג׳ה;מוחרם;רביע אל־אוול;רביע א־ת׳אני;ג׳ומאדא אל־אולא;ג׳ומאדא א־ת׳אניה;רג׳ב;שעבאן;רמדאן"),mmm6:T("רמדאן;שוואל;ד׳ו אל־קעדה;ד׳ו אל־חיג׳ה;מוחרם;רביע א׳;רביע ב׳;ג׳ומאדא א׳;ג׳ומאדא ב׳;רג׳ב;שעבאן;רמדאן")},"he"),k(j({group:".",decimal:",",mmmm:T("siječnja;veljače;ožujka;travnja;svibnja;lipnja;srpnja;kolovoza;rujna;listopada;studenoga;prosinca"),mmm:T("sij;velj;ožu;tra;svi;lip;srp;kol;ruj;lis;stu;pro"),dddd:T("nedjelja;ponedjeljak;utorak;srijeda;četvrtak;petak;subota")},-1,3),"hr"),k({decimal:",",mmmm:T("հունվարի;փետրվարի;մարտի;ապրիլի;մայիսի;հունիսի;հուլիսի;օգոստոսի;սեպտեմբերի;հոկտեմբերի;նոյեմբերի;դեկտեմբերի"),mmm:T("հնվ;փտվ;մրտ;ապր;մյս;հնս;հլս;օգս;սեպ;հոկ;նոյ;դեկ"),dddd:T("կիրակի;երկուշաբթի;երեքշաբթի;չորեքշաբթի;հինգշաբթի;ուրբաթ;շաբաթ"),ddd:T("կիր;երկ;երք;չրք;հնգ;ուր;շբթ")},"hy"),k(j({decimal:",",mmmm:T("იანვარი;თებერვალი;მარტი;აპრილი;მაისი;ივნისი;ივლისი;აგვისტო;სექტემბერი;ოქტომბერი;ნოემბერი;დეკემბერი"),dddd:T("კვირა;ორშაბათი;სამშაბათი;ოთხშაბათი;ხუთშაბათი;პარასკევი;შაბათი")},3,3),"ka"),k(j({decimal:",",mmmm:T("қаңтар;ақпан;наурыз;сәуір;мамыр;маусым;шілде;тамыз;қыркүйек;қазан;қараша;желтоқсан"),dddd:T("жексенбі;дүйсенбі;сейсенбі;сәрсенбі;бейсенбі;жұма;сенбі"),ddd:T("жс;дс;сс;ср;бс;жм;сб")},13,-1),"kk"),k({group:",",mmmm:T("ಜನವರಿ;ಫೆಬ್ರವರಿ;ಮಾರ್ಚ್;ಏಪ್ರಿಲ್;ಮೇ;ಜೂನ್;ಜುಲೈ;ಆಗಸ್ಟ್;ಸೆಪ್ಟೆಂಬರ್;ಅಕ್ಟೋಬರ್;ನವೆಂಬರ್;ಡಿಸೆಂಬರ್"),mmm:T("ಜನವರಿ;ಫೆಬ್ರವರಿ;ಮಾರ್ಚ್;ಏಪ್ರಿ;ಮೇ;ಜೂನ್;ಜುಲೈ;ಆಗಸ್ಟ್;ಸೆಪ್ಟೆಂ;ಅಕ್ಟೋ;ನವೆಂ;ಡಿಸೆಂ"),dddd:T("ಭಾನು~;ಸೋಮ~;ಮಂಗಳ~;ಬುಧ~;ಗುರು~;ಶುಕ್ರ~;ಶನಿ~","ವಾರ"),ddd:T("ಭಾನು;ಸೋಮ;ಮಂಗಳ;ಬುಧ;ಗುರು;ಶುಕ್ರ;ಶನಿ"),ampm:T("ಪೂರ್ವಾಹ್ನ;ಅಪರಾಹ್ನ")},"kn"),k({decimal:",",mmmm:T("sausio;vasario;kovo;balandžio;gegužės;birželio;liepos;rugpjūčio;rugsėjo;spalio;lapkričio;gruodžio"),mmm:T("saus.;vas.;kov.;bal.;geg.;birž.;liep.;rugp.;rugs.;spal.;lapkr.;gruod."),dddd:T("sekmadienis;pirmadienis;antradienis;trečiadienis;ketvirtadienis;penktadienis;šeštadienis"),ddd:T("sk;pr;an;tr;kt;pn;št"),ampm:T("priešpiet;popiet")},"lt"),k({decimal:",",mmmm:T("janvāris;februāris;marts;aprīlis;maijs;jūnijs;jūlijs;augusts;septembris;oktobris;novembris;decembris"),mmm:T("janv.;febr.;marts;apr.;maijs;jūn.;jūl.;aug.;sept.;okt.;nov.;dec."),dddd:T("svētdiena;pirmdiena;otrdiena;trešdiena;ceturtdiena;piektdiena;sestdiena"),ddd:T("svētd.;pirmd.;otrd.;trešd.;ceturtd.;piektd.;sestd."),ampm:T("priekšpusdienā;pēcpusdienā")},"lv"),k({group:",",decimal:".",mmmm:T("ജനുവരി;ഫെബ്രുവരി;മാർച്ച്;ഏപ്രിൽ;മേയ്;ജൂൺ;ജൂലൈ;ഓഗസ്റ്റ്;സെപ്റ്റംബർ;ഒക്‌ടോബർ;നവംബർ;ഡിസംബർ"),mmm:T("ജനു;ഫെബ്രു;മാർ;ഏപ്രി;മേയ്;ജൂൺ;ജൂലൈ;ഓഗ;സെപ്റ്റം;ഒക്ടോ;നവം;ഡിസം"),dddd:T("ഞായറാഴ്‌ച;തിങ്കളാഴ്‌ച;ചൊവ്വാഴ്ച;ബുധനാഴ്‌ച;വ്യാഴാഴ്‌ച;വെള്ളിയാഴ്‌ച;ശനിയാഴ്‌ച"),ddd:T("ഞായർ;തിങ്കൾ;ചൊവ്വ;ബുധൻ;വ്യാഴം;വെള്ളി;ശനി")},"ml"),k({group:",",decimal:".",mmmm:T("нэгдүгээ~;хоёрдугаа~;гуравдугаа~;дөрөвдүгээ~;тавдугаа~;зургаадугаа~;долоодугаа~;наймдугаа~;есдүгээ~;аравдугаа~;арван нэгдүгээ~;арван хоёрдугаа~","р сар"),mmm:T("1~;2~;3~;4~;5~;6~;7~;8~;9~;10~;11~;12~","-р сар"),dddd:T("ням;даваа;мягмар;лхагва;пүрэв;баасан;бямба"),ddd:T("Ня;Да;Мя;Лх;Пү;Ба;Бя"),ampm:T("ү.ө.;ү.х.")},"mn"),k({group:",",decimal:".",mmmm:T("जानेवारी;फेब्रुवारी;मार्च;एप्रिल;मे;जून;जुलै;ऑगस्ट;सप्टेंबर;ऑक्टोबर;नोव्हेंबर;डिसेंबर"),mmm:T("जाने;फेब्रु;मार्च;एप्रि;मे;जून;जुलै;ऑग;सप्टें;ऑक्टो;नोव्हें;डिसें"),dddd:T("रविवार;सोमवार;मंगळवार;बुधवार;गुरुवार;शुक्रवार;शनिवार"),ddd:T("रवि;सोम;मंगळ;बुध;गुरु;शुक्र;शनि")},"mr"),k(j({group:",",decimal:".",mmmm:T("ဇန်နဝါရီ;ဖေဖော်ဝါရီ;မတ်;ဧပြီ;မေ;ဇွန်;ဇူလိုင်;ဩဂုတ်;စက်တင်ဘာ;အောက်တိုဘာ;နိုဝင်ဘာ;ဒီဇင်ဘာ"),mmm:T("ဇန်;ဖေ;မတ်;ဧ;မေ;ဇွန်;ဇူ;ဩ;စက်;အောက်;နို;ဒီ"),dddd:T("တနင်္ဂနွေ;တနင်္လာ;အင်္ဂါ;ဗုဒ္ဓဟူး;ကြာသပတေး;သောကြာ;စနေ"),ampm:T("နံနက်;ညနေ")},-1,0),"my"),k({group:",",decimal:".",mmmm:T("ਜਨਵਰੀ;ਫ਼ਰਵਰੀ;ਮਾਰਚ;ਅਪ੍ਰੈਲ;ਮਈ;ਜੂਨ;ਜੁਲਾਈ;ਅਗਸਤ;ਸਤੰਬਰ;ਅਕਤੂਬਰ;ਨਵੰਬਰ;ਦਸੰਬਰ"),mmm:T("ਜਨ;ਫ਼ਰ;ਮਾਰਚ;ਅਪ੍ਰੈ;ਮਈ;ਜੂਨ;ਜੁਲਾ;ਅਗ;ਸਤੰ;ਅਕਤੂ;ਨਵੰ;ਦਸੰ"),dddd:T("ਐਤਵਾਰ;ਸੋਮਵਾਰ;ਮੰਗਲਵਾਰ;ਬੁੱਧਵਾਰ;ਵੀਰਵਾਰ;ਸ਼ੁੱਕਰਵਾਰ;ਸ਼ਨਿੱਚਰਵਾਰ"),ddd:T("ਐਤ;ਸੋਮ;ਮੰਗਲ;ਬੁੱਧ;ਵੀਰ;ਸ਼ੁੱਕਰ;ਸ਼ਨਿੱਚਰ"),ampm:T("ਪੂ.ਦੁ.;ਬਾ.ਦੁ.")},"pa"),k({group:".",decimal:",",mmmm:T("ianuarie;februarie;martie;aprilie;mai;iunie;iulie;august;septem~;octom~;noiem~;decem~","brie"),mmm:T("ian.;feb.;mar.;apr.;mai;iun.;iul.;aug.;sept.;oct.;nov.;dec."),dddd:T("duminică;luni;marți;miercuri;joi;vineri;sâmbătă"),ddd:T("dum.;lun.;mar.;mie.;joi;vin.;sâm."),ampm:T("a.m.;p.m.")},"ro"),k(j({group:".",decimal:",",mmmm:T("januar;februar;marec;april;maj;junij;julij;avgust;september;oktober;november;december"),mmm:T("jan.;feb.;mar.;apr.;maj;jun.;jul.;avg.;sep.;okt.;nov.;dec."),dddd:T("nedelja;ponedeljek;torek;sreda;četrtek;petek;sobota"),ampm:T("dop.;pop.")},-1,13),"sl"),k(j({group:".",decimal:",",mmmm:T("јануар;фебруар;март;април;мај;јун;јул;август;септембар;октобар;новембар;децембар"),dddd:T("недеља;понедељак;уторак;среда;четвртак;петак;субота")},3,3),"sr"),k({group:",",decimal:".",mmmm:T("ஜனவரி;பிப்ரவரி;மார்ச்;ஏப்ரல்;மே;ஜூன்;ஜூலை;ஆகஸ்ட்;செப்டம்பர்;அக்டோபர்;நவம்பர்;டிசம்பர்"),mmm:T("ஜன.;பிப்.;மார்.;ஏப்.;மே;ஜூன்;ஜூலை;ஆக.;செப்.;அக்.;நவ.;டிச."),dddd:T("ஞாயிறு;திங்கள்;செவ்வாய்;புதன்;வியாழன்;வெள்ளி;சனி"),ddd:T("ஞாயி.;திங்.;செவ்.;புத.;வியா.;வெள்.;சனி")},"ta"),k({group:",",decimal:".",mmmm:T("జనవరి;ఫిబ్రవరి;మార్చి;ఏప్రిల్;మే;జూన్;జులై;ఆగస్టు;సెప్టెంబర్;అక్టోబర్;నవంబర్;డిసెంబర్"),mmm:T("జన;ఫిబ్ర;మార్చి;ఏప్రి;మే;జూన్;జులై;ఆగ;సెప్టెం;అక్టో;నవం;డిసెం"),dddd:T("ఆదివారం;సోమవారం;మంగళవారం;బుధవారం;గురువారం;శుక్రవారం;శనివారం"),ddd:T("ఆది;సోమ;మంగళ;బుధ;గురు;శుక్ర;శని")},"te"),k({decimal:",",mmmm:T("січня;лютого;березня;квітня;травня;червня;липня;серпня;вересня;жовтня;листопада;грудня"),mmm:T("січ.;лют.;бер.;квіт.;трав.;черв.;лип.;серп.;вер.;жовт.;лист.;груд."),dddd:T("неділю;понеділок;вівторок;середу;четвер;пʼятницю;суботу"),ddd:T("нд;пн;вт;ср;чт;пт;сб"),ampm:T("дп;пп")},"uk"),k({group:".",decimal:",",mmmm:T("~1;~2;~3;~4;~5;~6;~7;~8;~9;~10;~11;~12","tháng "),mmm:T("~1;~2;~3;~4;~5;~6;~7;~8;~9;~10;~11;~12","thg "),dddd:T("Chủ Nhật;Thứ Hai;Thứ Ba;Thứ Tư;Thứ Năm;Thứ Sáu;Thứ Bảy"),ddd:T("CN;Th 2;Th 3;Th 4;Th 5;Th 6;Th 7"),ampm:T("SA;CH")},"vi"),k(j({group:"٬",decimal:"٫",ampm:T("ص;م"),mmmm:T("يناير;فبراير;مارس;أبريل;مايو;يونيو;يوليو;أغسطس;سبتمبر;أكتوبر;نوفمبر;ديسمبر"),dddd:T("الأحد;الاثنين;الثلاثاء;الأربعاء;الخميس;الجمعة;السبت"),mmmm6:T("رمضان;شوال;ذو القعدة;ذو الحجة;محرم;ربيع الأول;ربيع الآخرة;جمادى الأولى;جمادى الآخرة;رجب;شعبان;رمضان")},0,0),"ar"),k({group:",",decimal:".",mmmm:T("জানুয়ারী;ফেব্রুয়ারী;মার্চ;এপ্রিল;মে;জুন;জুলাই;আগস্ট;সেপ্টেম্বর;অক্টোবর;নভেম্বর;ডিসেম্বর"),mmm:T("জানু;ফেব;মার্চ;এপ্রি;মে;জুন;জুল;আগ;সেপ্টেঃ;অক্টোঃ;নভেঃ;ডিসেঃ"),dddd:T("রবিবার;সোমবার;মঙ্গলবার;বুধবার;বৃহস্পতিবার;শুক্রবার;শনিবার"),ddd:T("রবি;সোম;মঙ্গল;বুধ;বৃহস্পতি;শুক্র;শনি")},"bn"),k({group:",",decimal:".",mmmm:T("जनवरी;फ़रवरी;मार्च;अप्रैल;मई;जून;जुलाई;अगस्त;सितंबर;अक्तूबर;नवंबर;दिसंबर"),mmm:T("जन॰;फ़र॰;मार्च;अप्रैल;मई;जून;जुल॰;अग॰;सित॰;अक्तू॰;नव॰;दिस॰"),dddd:T("रविवार;सोमवार;मंगलवार;बुधवार;गुरुवार;शुक्रवार;शनिवार"),ddd:T("रवि;सोम;मंगल;बुध;गुरु;शुक्र;शनि"),ampm:T("am;pm")},"hi");const ba={overflow:"######",dateErrorThrows:!1,dateErrorNumber:!0,bigintErrorNumber:!1,dateSpanLarge:!0,leap1900:!0,nbsp:!1,throws:!0,invalid:"######",locale:"",ignoreTimezone:!1,grouping:[3,3],indexColors:!0,skipChar:"",repeatChar:""};function De(e,t=0){if(typeof e!="number")return e;if(e<0)return-De(-e,t);if(t){const n=10**t||1;return De(e*n,0)/n}return Math.round(e)}const _a=1e-13;function ja(e,t=2,n=2){const r=e<0?-1:1,o=10**(t||2),l=10**(n||2);let c=Math.abs(e),h=0,g=0,u=0,i=1,s,f;if(e=c,e%1===0)f=[e*r,1];else if(e<1e-19)f=[r,1e19];else if(e>1e19)f=[1e19*r,1];else{do if(c=1/(c-Math.floor(c)),s=i,i=i*Math.floor(c)+h,h=s,g=u,u=Math.floor(e*i+.5),u>=o||i>=l)return[r*g,h];while(Math.abs(e-u/i)>=_a&&c!==Math.floor(c));f=[r*u,i]}return f}const Ie=Math.floor;function fo(e,t=!0){if(t&&e>=0){if(e===0)return[1900,1,0];if(e===60)return[1900,2,29];if(e<60)return[1900,e<32?1:2,(e-1)%31+1]}let n=e+68569+2415019;const r=Ie(4*n/146097);n=n-Ie((146097*r+3)/4);const o=Ie(4e3*(n+1)/1461001);n=n-Ie(1461*o/4)+31;const l=Ie(80*n/2447),c=n-Ie(2447*l/80);n=Ie(l/11);const h=l+2-12*n;return[100*(r-49)+o+n|0,h|0,c|0]}function e1(e){return fo(e+1462)}function t1(e){if(e===60)throw new Error("#VALUE!");if(e<=1)return[1317,8,29];if(e<60)return[1317,e<32?9:10,1+(e-2)%30];const t=10631/30,n=8.01/60;let r=e+466935;const o=Ie(r/10631);r=r-10631*o;const l=Ie((r-n)/t);r=r-Ie(l*t+n);const c=Ie((r+28.5001)/29.5);return c===13?[30*o+l,12,30]:[30*o+l,c,r-Ie(29.5001*c-29)]}function n1(e,t=0,n=!0){const r=Ie(e);return t===lt?t1(r):t===$a?e1(r):fo(r,n)}function r1(e,t){let n=null;if(Array.isArray(e)){const[r,o,l,c,h,g]=e;n=Date.UTC(r,o==null?0:o-1,l??1,c||0,h||0,g||0)}else if(e instanceof Date&&(n=e*1,!t?.ignoreTimezone)){const r=new Date;r.setUTCFullYear(e.getFullYear(),e.getMonth(),e.getDate()),r.setUTCHours(e.getHours(),e.getMinutes(),e.getSeconds(),e.getMilliseconds()),n=r*1}if(n!=null&&isFinite(n)){const r=n/864e5;return r-(r<=-25509?-25568:-25569)}return null}const i1=["!d-m-y","!d-m-Y","!j-m-y","!j-m-Y","!d-n-y","!d-n-Y","!j-n-y","!j-n-Y","?m-d-y","?m-d-Y","?m-j-y","?m-j-Y","?n-d-y","?n-d-Y","?n-j-y","?n-j-Y","d-M-y","d-M-Y","j-M-y","j-M-Y","M-d-y","M-d-Y","M-j-y","M-j-Y","d-F-y","d-F-Y","F-d-y","F-d-Y","F-j-y","F-j-Y","j-F-y","j-F-Y","y-F-d","y-F-j","y-M-d","y-M-j","Y-F-d","Y-F-j","Y-M-d","Y-m-d","Y-M-j","Y-m-j","Y-n-d","Y-n-j","j-F","j-M","d-F","d-M","n-d","n-j","n-Y","m-d","m-j","m-Y","M-Y","M-y","F-y","F-Y","Y-M","Y-n","Y-m","Y-F","Y-M"],o1={},s1={};function Ne(e,t,n=1){if(e){const r=e[0],o=e.slice(1);r==="!"?Ne(o,t,4):r==="?"?Ne(o,t,2):(t[r]=t[r]||{},Ne(o,t[r],n))}else t.$=n}function lo(e,t){Ne(e,t),Ne(e+" x",t),Ne(e+" l",t),Ne(e+" l x",t),Ne("l "+e,t),Ne("l "+e+" x",t),Ne(e+" D",t),Ne(e+" D x",t),Ne("D "+e,t),Ne("D "+e+" x",t)}i1.forEach(e=>{e.startsWith("?")||lo(e,o1),e.startsWith("!")||lo(e,s1)}),new Date().getUTCFullYear();function je(e,t,n){return t[n+"_max"]=e.length,t[n+"_min"]=e.replace(/#/g,"").length,t}function oe(e,t){typeof e=="string"?t.push({type:"string",value:e}):t.push(e)}function co(e,t){const n=e?.type;return n===Ui||n===Dt||n===yi||n===ur&&t==="den"}function wt(e){const t=[],n={scale:1,percent:!1,text:!1,date:0,date_eval:!1,date_system:Si,sec_decimals:0,general:!1,clock:24,int_pattern:[],frac_pattern:[],man_pattern:[],den_pattern:[],num_pattern:[],tokens:t};let r="int",o=null;const l=[];let c,h=!1,g=-1,u=!1,i="",s=!1;for(;++g<e.length&&!u;){const p=e[g],I=p.type||cn;if(i+=p.raw,I===Pi)n.general=!0,oe(p,t);else if(co(p,r)){const A=n[r+"_pattern"];co(c,r)||c?.type===ln?(A.push((A.pop()||"")+p.value),o.num+=p.value):(A.push(p.value),o={type:r,num:p.value},oe(o,t))}else if(I===to)p.value==="("&&(n.parens=!0),oe(p.value,t);else if(I===ur)oe(p.value,t);else if(I===Fi)if(s=!0,n[r+"_pattern"].length){if(!o)throw new SyntaxError("Format pattern is missing a numerator");n.fractions=!0,n.num_pattern.push(n[r+"_pattern"].pop()),o.type="num",r="den",oe({type:"div"},t)}else oe(p.value,t);else if(I===Hi)oe(",",t);else if(I===ir)n.scale=.001**p.raw.length;else if(I===ln){if(r==="int"&&(n.grouping=!0),r==="den")throw new SyntaxError("Cannot group denominator digits")}else if(I===Gi)oe(p,t);else if(I===Vi){u=!0;break}else if(I===or)n.text=!0,oe(p,t);else if(I===ki||I===Bi)oe(p,t);else if(I===qi){const A=p.value.toLowerCase(),M=A[0],O={type:"",size:0,date:1,pad:A.length};M==="h"?(O.size=sn,O.type="hour-elap"):M==="m"?(O.size=un,O.type="min-elap"):(O.size=fn,O.type="sec-elap"),n.date=n.date|O.size,l.push(O),oe(O,t)}else if(n.date&&I===sr&&e[g+1]?.type===Dt){let A=1;g++;let M="0";e[g+1]?.type===Dt&&(M+="0",A=2,g++),e[g+1]?.type===Dt&&(M+="0",A=3,g++),i+=M;const O=[fn,tr,nr,rr][A];n.date=n.date|O,n.date_eval=!0,n.sec_decimals=Math.max(n.sec_decimals,A),oe({type:"subsec",size:O,decimals:A,date:1},t)}else if(I===Yi)h||(p.value==="B2"||p.value==="b2"?n.date_system=lt:n.date_system=Si);else if(I===Wi){const A={type:"",size:0,date:1},M=p.value.toLowerCase(),O=M[0];if(M==="y"||M==="yy")A.size=rn,A.type="year-short";else if(O==="y"||O==="e")A.size=rn,A.type="year";else if(M==="b"||M==="bb")A.size=rn,A.type="b-year-short";else if(O==="b")A.size=rn,A.type="b-year";else if(M==="d"||M==="dd")A.size=er,A.type="day",A.pad=/dd/.test(M);else if(M==="ddd"||M==="aaa")A.size=er,A.type="weekday-short";else if(O==="d"||O==="a")A.size=er,A.type="weekday";else if(O==="h")A.size=sn,A.type="hour",A.pad=/hh/i.test(M);else if(O==="m"){M.length===3?(A.size=on,A.type="monthname-short"):M.length===5?(A.size=on,A.type="monthname-single"):M.length>=4&&(A.size=on,A.type="monthname");const L=l[l.length-1];!A.type&&L&&!L.used&&L.size&(sn|fn)&&(L.used=!0,A.size=un,A.type="min",A.pad=/mm/.test(M)),A.type||(A.size=on,A.type="month",A.pad=/mm/.test(M),A.indeterminate=!0)}else if(O==="s"){A.size=fn,A.type="sec",A.pad=/ss/.test(M);const L=l[l.length-1];L&&L.size&un?A.used=!0:L?.indeterminate&&(delete L.indeterminate,L.size=un,L.type="min",A.used=!0)}n.date=n.date|A.size,n.date_eval=!0,l.push(A),oe(A,t)}else if(I===Ji)n.clock=12,n.date=n.date|sn,n.date_eval=!0,p.short=p.value==="A/P",oe(p,t);else if(I===bi||I===Zi||I===fr)oe(p.value,t);else if(I===Ki)n.condition=[p.value[0],parseFloat(p.value[1])];else if(I===xi){const A=p.value.split("-"),M=A.length<2?"":A.slice(1).join("-"),O=A[0];O&&oe(O,t);const L=cr(M);L&&(n.locale=L);const U=parseInt(M,16);isFinite(U)&&U&16711680&&(U>>16&255)===6&&(n.date_system=lt),h=!0}else if(I===Qi){let A,M=p.value.toLowerCase();(A=/^color\\s*(\\d+)$/i.exec(M))&&(M=parseInt(A[1],10)),n.color=M}else if(I===zi)n.scale=100,n.percent=!0,oe("%",t);else if(I===sr)oe(p,t),n.date||(n.dec_fractions=!0,r="frac");else if(I===ji)n.exponential=!0,n.exp_plus=p.value.includes("+"),r="man",oe({type:"exp",plus:n.exp_plus},t);else if(I===_i)oe(p,t);else if(I===eo)oe(p,t);else if(!(I===Xi||I===$i))throw I===cn?new SyntaxError(`Illegal character: ${i}`):new SyntaxError(`Unknown token ${I} in ${i}`);c=p}if(n.tokensUsed=g,n.pattern=i,/^((?:\\[[^\\]]+\\])+)(;|$)/.test(n.pattern)&&!/^\\[(?:h+|m+|s+)\\]/.test(n.pattern)&&oe({type:"text"},t),n.fractions&&n.dec_fractions||n.grouping&&!n.int_pattern.length||n.fractions&&n.exponential||n.fractions&&n.den_pattern.length*n.num_pattern.length===0||s&&!n.fractions&&!n.date||n.exponential&&(n.int_pattern.length||n.frac_pattern.length)*n.man_pattern.length===0)throw new SyntaxError(`Invalid pattern: ${i}`);const f=n.int_pattern.join(""),a=n.man_pattern.join(""),d=n.frac_pattern.join("");je(f,n,"int");let E=0;for(let p=0;p<f.length;p++){const I=f[f.length-1-p];/^[0-9?]/.test(I)&&(E=p+1)}n.int_min=E,je(d,n,"frac"),je(a,n,"man");let m=n.num_pattern.join(""),C=n.den_pattern[0]||"";if(C.includes("?")||m.includes("?")?(C=C.replace(/\\d/g,"?"),C=C.replace(/#$/g,"?"),je(m,n,"num"),je(C,n,"den"),m=m.replace(/#$/g,"?")):(je(m,n,"num"),je(C,n,"den")),n.int_p=f,n.man_p=a,n.num_p=m,n.den_p=C,n.den_pattern.length&&(n.denominator=parseInt(n.den_pattern.join("").replace(/\\D/g,""),10)),n.integer=!!f.length,!n.integer&&!n.exponential&&d.length){const p=n.tokens.findIndex(I=>I.type==="point");n.tokens.splice(p,0,{type:"int",value:"#"}),n.integer=!0,n.int_pattern=["#"],n.int_p="#"}if(n.fractions)for(let p=0;p<t.length-1;p++){const I=t[p];if(I.type!=="string"&&I.type!=="space")continue;const A=t[p+1].type;A==="num"?I.rule="num+int":A==="div"?I.rule="num":A==="den"&&(I.rule="den")}return n.grouping&&n.int_pattern.length>1&&(n.grouping=!1),n}function u1(e){if(e===0)return e;const t=Math.ceil(Math.log10(e<0?-e:e)),n=10**(16-Math.floor(t));return isFinite(n)?Math.round(e*n)/n:0}function ao(e,t=0){const n=Math.floor(Math.log10(e));return t>1?Math.floor(n/t)*t:n}function ho(e,t=1){return t<-300?parseFloat(e.toExponential().split("e")[0]):e*10**-t}const f1={total:1,sign:0,period:0,int:1,frac:0};function l1(e,t=!0){const n=Math.abs(e);if(!n)return f1;const r=t&&e<0?1:0,o=Math.floor(n),l=Math.floor(Math.log10(n)+1);let c=0,h=0;if(o!==n){c=1;const g=String(De(n*10**-l,15));let u=g.length,i=!0,s=0;for(;s<=g.length;){if(g[s]==="."){u--;break}else g[s]==="0"&&i?u--:i=!1;s++}h=u-l,h<0&&(h=0,c=0)}return{total:r+Math.max(l,1)+c+h,digits:Math.max(l,0)+h,sign:r,period:c,int:Math.max(l,1),frac:h}}const at=(e,t)=>e.replace(/\\./,t.decimal),go=(e,t,n)=>{const r=Math.abs(t);let o;return e===1?o=e:o=De(e,5),[at(o+"",n),n.exponent,t<0?n.negative:n.positive,r<10?"0":"",r]};function Eo(e,t,n,r){const o=n|0;if(typeof n=="string")e.push(n);else if(n===o)e.push(Math.abs(o));else{const l=Math.abs(n);let c=ao(l),h=ho(l,c);h===10&&(h=1,c++);const g=l1(l);if(c>=-4&&c<=-1){const u=l.toPrecision(10+c).replace(/\\.?0+$/,"");e.push(at(u,r))}else if(c===10){const u=l.toFixed(10).slice(0,12).replace(/\\.$/,"");e.push(at(u,r))}else if(Math.abs(c)<=9)if(g.total<=11){const u=De(l,9).toFixed(g.frac);e.push(at(u,r))}else c===9?e.push(Math.floor(l)):c>=0&&c<9?e.push(at(String(De(l,9-c)),r)):e.push(...go(h,c,r));else g.total>=12?e.push(...go(h,c,r)):e.push(at(De(l,9).toFixed(g.frac),r))}return e}function et(e,t=!1){return e==="0"?"0":e==="?"?t?" ":" ":""}const Lt=86400,c1=(e,t,n)=>n?e<Ka||t>=Xa:e<Wa||t>=qa;function mo(e,t,n,r){let o="",l="",c="",h="",g="",u="",i=0,s=0;if(typeof e=="bigint"){if(e<=Number.MAX_SAFE_INTEGER&&e>=Number.MIN_SAFE_INTEGER)e=Number(e);else return n.bigintErrorNumber?String(e):n.overflow;s=e}else s=Math.trunc(e);let f=0,a=0,d=1,E=0,m=0,C=0,v=0,p=0,I=0;const A=r||an;if(!t.text&&isFinite(t.scale)&&t.scale!==1&&(e=u1(e*t.scale)),t.exponential){let X=Math.abs(e);X&&(i=ao(X,t.int_max)),e&&!t.integer&&i++,X=ho(X,i),t.int_max===1&&De(X,t.frac_max)===10&&(X=1,i++),e=e<0?-X:X,o+=Math.abs(i)}if(t.integer){const X=Math.abs(De(e,t.fractions?1:t.frac_max));u+=X<1?"":Math.floor(X)}const M=n.grouping[0]??3,O=n.grouping[1]??M;t.dec_fractions&&(g=String(De(e,t.frac_max)).split(".")[1]||"");const L=!t.error&&(t.num_p.includes("0")||t.den_p.includes("0"));let U=L;if(t.fractions){U=L||!!(e%1);const X=Math.abs(t.integer?e%1:e);if(X)if(U=!0,t.denominator&&isFinite(t.denominator))h+=t.denominator,c+=De(X*t.denominator),c==="0"&&(c="",h="",U=L);else{const $=ja(X,1/0,t.den_max);c+=$[0],h+=$[1],t.integer&&c==="0"&&(c="",h="",U=L)}else!e&&!t.integer&&(U=!0,c="0",h="1");t.integer&&!U&&!Math.trunc(e)&&(u="0")}if(t.date){s=Math.trunc(e);const X=Lt*(e-s);if(f=Math.floor(X),I=X-f,Math.abs(I)<1e-6?I=0:I>.9999&&(I=0,f+=1,f===Lt&&(f=0,s+=1)),I){const $=t.date&rr||t.date&nr||t.date&tr;($===rr&&I>.9995||$===nr&&I>.995||$===tr&&I>.95||!$&&I>=.5)&&(f++,I=0)}if(s||t.date_system){const $=n1(e,t.date_system,n.leap1900);a=$[0],d=$[1],E=$[2]}if(f){const $=f<0?Lt+f:f;p=Math.floor($)%60,v=Math.floor($/60)%60,C=Math.floor($/60/60)%60}if(m=(6+s)%7,t.date_eval&&c1(e,s+f/Lt,n.dateSpanLarge)){if(n.dateErrorThrows)throw new Error("Date out of bounds");if(n.dateErrorNumber){const $=e<0?[A.negative]:[];return Eo($,{},e,A).join("")}return n.overflow}}const H=et("?",n.nbsp);i<0?l="-":t.exp_plus&&(l="+");const y=[],Z=(X,$,G,z)=>{const ue=!z&&X.length>$.length?G.length+X.length-$.length:G.length;X.length<$.length&&(z+=X.length-$.length);for(let W=0;W<ue;W++)y.push(X[W+z]||et(G[W],n.nbsp));return ue};let Q=!1;const te={int:0,frac:0,man:0,num:0,den:0};for(let X=0,$=t.tokens.length;X<$;X++){const G=t.tokens[X],z=G.type,ue=G.num?G.num.length:0;if(z==="string")G.rule?G.rule==="num"?U?y.push(G.value.replace(/ /g,H)):(t.num_min>0||t.den_min>0)&&y.push(G.value.replace(/./g,H)):G.rule==="num+int"?U&&u?y.push(G.value.replace(/ /g,H)):t.den_min>0&&(u||t.num_min)&&y.push(G.value.replace(/./g,H)):G.rule==="den"&&(U?y.push(G.value.replace(/ /g,H)):(t.den_min>0||t.den_min>0)&&y.push(G.value.replace(/./g,H))):y.push(G.value.replace(/ /g,H));else if(z==="space")G.rule==="num+int"?(U||t.num_min||t.den_min)&&(u||t.num_min)&&y.push(H):y.push(H);else if(z==="error")y.push(n.invalid);else if(z==="point")y.push(t.date?G.value:A.decimal);else if(z==="general")Eo(y,t,e,A);else if(z==="exp")y.push(A.exponent);else if(z==="minus")G.volatile&&t.date||G.volatile&&(e>=0||typeof e!="number")||(G.volatile&&!t.fractions&&(t.integer||t.dec_fractions)?(e<0&&u&&u!=="0"||g)&&y.push(A.negative):y.push(A.negative));else if(z==="plus")y.push(A.positive);else if(z==="text")y.push(e);else if(z==="fill")n.fillChar&&y.push(n.fillChar,G.value);else if(z==="skip")n.skipChar?y.push(n.skipChar,G.value):y.push(n.nbsp?" ":" ");else if(z==="div")U?y.push("/"):t.num_min>0||t.den_min>0?y.push(H):y.push(et("#",n.nbsp));else if(z==="int")if(t.int_pattern.length===1){const W=t.int_p,re=Math.max(t.int_min,u.length);let le="";for(let ne=re;ne>0;ne--){const we=u.charAt(u.length-ne),Xe=we?"":W.charAt(W.length-ne)||W[0];let gt="";if(t.grouping){const ke=ne-1-M;ke>=0&&!(ke%O)&&(gt=we||Xe==="0"?A.group:et("?",n.nbsp))}le+=(we||et(Xe,n.nbsp))+gt}y.push(le)}else te.int+=Z(u,t.int_p,G.num,te.int);else if(z==="frac"){const W=te.frac;for(let re=0;re<ue;re++)y.push(g[re+W]||et(G.num[re],n.nbsp));te.frac+=ue}else if(z==="man")!te[z]&&!te.man&&y.push(l),te.man+=Z(o,t.man_p,G.num,te.man);else if(z==="num")te.num+=Z(c,t.num_p,G.num,te.num);else if(z==="den"){const W=te.den;for(let re=0;re<ue;re++){let le=h[re+W];if(!le){const ne=G.num[re];"123456789".includes(ne)||Q&&ne==="0"?(Q=!0,le=n.nbsp?" ":" "):!Q&&re===ue-1&&ne==="0"&&!h?le="1":le=et(ne,n.nbsp)}y.push(le)}te.den+=ue}else if(z==="year")a<0&&y.push(A.negative),y.push(String(Math.abs(a)).padStart(4,"0"));else if(z==="year-short"){const W=a%100;y.push(W<10?"0":"",W)}else if(z==="month")y.push(G.pad&&d<10?"0":"",d);else if(z==="monthname-single")t.date_system===lt?y.push(A.mmmm6[d-1].charAt(0)):y.push(A.mmmm[d-1].charAt(0));else if(z==="monthname-short")t.date_system===lt?y.push(A.mmm6[d-1]):y.push(A.mmm[d-1]);else if(z==="monthname")t.date_system===lt?y.push(A.mmmm6[d-1]):y.push(A.mmmm[d-1]);else if(G.type==="weekday-short")y.push(A.ddd[m]);else if(z==="weekday")y.push(A.dddd[m]);else if(z==="day")y.push(G.pad&&E<10?"0":"",E);else if(z==="hour"){const W=C%t.clock||(t.clock<24?t.clock:0);y.push(G.pad&&W<10?"0":"",W)}else if(z==="min")y.push(G.pad&&v<10?"0":"",v);else if(z==="sec")y.push(G.pad&&p<10?"0":"",p);else if(z==="subsec"){y.push(A.decimal);const W=I.toFixed(t.sec_decimals);y.push(W.slice(2,2+G.decimals))}else if(z==="ampm"){const W=C<12?0:1;G.short&&!r?y.push("AP"[W]):y.push(A.ampm[W])}else if(z==="hour-elap"){e<0&&y.push(A.negative);const W=s*24+Math.floor(Math.abs(f)/3600);y.push(String(Math.abs(W)).padStart(G.pad,"0"))}else if(z==="min-elap"){e<0&&y.push(A.negative);const W=s*1440+Math.floor(Math.abs(f)/60);y.push(String(Math.abs(W)).padStart(G.pad,"0"))}else if(z==="sec-elap"){e<0&&y.push(A.negative);const W=s*Lt+Math.abs(f);y.push(String(Math.abs(W)).padStart(G.pad,"0"))}else if(z==="b-year")y.push(a+543);else if(z==="b-year-short"){const W=(a+543)%100;y.push(W<10?"0":"",W)}}return y.join("")}const a1=wt([{type:or,value:"@",raw:"@"}]);function h1(e,t){for(let n=0;n<3;n++){const r=t[n];if(r){let o;if(r.condition){const l=r.condition[0],c=r.condition[1];l==="="?o=e===c:l===">"?o=e>c:l==="<"?o=e<c:l===">="?o=e>=c:l==="<="?o=e<=c:l==="<>"&&(o=e!==c)}else o=!0;if(o)return r}}}function g1(e,t,n){const r=t.partitions,o=Ja(t.locale||n.locale),l=r[3]?r[3]:a1;if(typeof e=="boolean"&&(e=(o||an).bool[e?0:1]),e==null)return"";const c=typeof e=="bigint";if(typeof e!="number"&&!c)return mo(e,l,n,o);if(!c&&!isFinite(e)){const g=o||an;return isNaN(e)?g.nan:(e<0?g.negative:"")+g.infinity}const h=h1(e,r);return h?mo(e,h,n,o):n.overflow}function d1(e){return!!(e[0]?.date||e[1]?.date||e[2]?.date||e[3]?.date)}const E1=[[Pi,/^General/i,0],[Ui,/^#/,0],[Dt,/^0/,0],[yi,/^\\?/,0],[Fi,/^\\//,0],[Vi,/^;/,0],[or,/^@/,0],[ki,/^\\+/,0],[Bi,/^-/,0],[sr,/^\\./,0],[Gi,/^ /,0],[zi,/^%/,0],[ur,/^[1-9]/,0],[Yi,/^(?:B[12])/i,0],[cn,/^B$/,0],[Wi,/^(?:[hH]+|[mM]+|[sS]+|[yY]+|[bB]+|[dD]+|[gG]+|[aA]{3,}|e+)/,0],[qi,/^(?:\\[(h+|m+|s+)\\])/i,1],[Ki,/^\\[(<[=>]?|>=?|=)\\s*(-?[.\\d]+)\\]/,[1,2]],[Xi,/^\\[(DBNum[0-4]?\\d)\\]/i,1],[$i,/^\\[(NatNum[0-4]?\\d)\\]/i,1],[xi,/^\\[\\$([^\\]]+)\\]/,1],[Qi,/^\\[(black|blue|cyan|green|magenta|red|white|yellow|color\\s*\\d+)\\]/i,1],[xa,/^\\[([^\\]]+)\\]/,1],[Ji,/^(?:AM\\/PM|am\\/pm|A\\/P)/,0],[Zi,/^\\\\(.)/,1],[bi,/^"([^"]*?)"/,1],[_i,/^_(\\\\.|.)/,1],[ji,/^[Ee]([+-])/,1],[eo,/^\\*(\\\\.|.)/,1],[to,/^[()]/,0],[cn,/^[EÈÉÊËèéêëĒēĔĕĖėĘęĚěȄȅȆȇȨȩNnÑñŃńŅņŇňǸǹ["*/\\\\_]/,0],[fr,/^./,0]],m1=63,p1=35,I1=48,N1=57,Er=e=>{const t=(e||"\\0").charCodeAt(0);return t===m1||t===p1||t>=I1&&t<=N1};function hn(e){let t=0;const n=[],r=[];for(;t<e.length;){const o=e.slice(t);let l=0;const c=/^(,+)(.)?/.exec(o);if(c){const h=c[1];l=h.length;const g=e[t-1]||"";let u=!1,i=!1;Er(g)?(u=!0,i=!0):g==="."&&(i=!0);const s=c[2]||"";if(u&&(!s||s===";")&&(u=!1),i&&Er(s)&&(i=!1),u&&!i)n.push({type:ln,value:",",raw:h});else if(!u&&i)n.push({type:ir,value:",",raw:h});else if(u&&i){const f={type:ir,value:",",raw:h};n.push(f),r.push(f)}else n.push({type:Hi,value:",",raw:h})}else{let h;for(const[g,u,i]of E1){const s=u.exec(o);if(s){const f=Array.isArray(i)?i.map(a=>s[a]):s[i||0];h={type:g,value:f,raw:s[0]},n.push(h),l=s[0].length;break}}r.length&&h.raw===";"&&(r.length=0),r.length&&Er(h.raw)&&(r.forEach(g=>g.type=ln),r.length=0)}if(!l){const h=o[0];l=1,n.push({type:fr,value:h,raw:h})}t+=l}return n}const mr=e=>{const[t,n]=e.condition??[];n<0&&(t==="<"||t==="<="||t==="=")||n===0&&t==="<"||e.tokens.unshift({type:"minus",volatile:!0})},po=(e,t=null)=>{const n={};for(const r in e)Array.isArray(e[r])?n[r]=[...e[r]]:n[r]=e[r];return t&&n.tokens.unshift(t),n.generated=!0,n};function A1(e){const t=[];let n=!1,r,o=null,l=0,c=!1,h=0,g=0,u=hn(e);do{if(c=wt(u),(c.date||c.general)&&(c.int_pattern.length||c.frac_pattern.length||c.scale!==1||c.text))throw new Error("Illegal format");if(c.condition&&(g++,n=!0),c.text){if(o)throw new Error("Unexpected partition");o=c}c.locale&&(r=cr(c.locale)),t.push(c),l=u[c.tokensUsed]?.type==="break"?1:0,u=u.slice(c.tokensUsed+l),h++}while(l&&h<4&&g<3);if(l)throw new Error("Unexpected partition");if(g>2)throw new Error("Unexpected condition");const i=t[3];if(i&&(i.int_pattern.length||i.frac_pattern.length||i.date))throw new Error("Unexpected partition");if(n){t[0].condition||(t[0].condition=[">",0]);const s=t.length;if(s===1&&(t[1]=wt(hn("General")),t[1].generated=!0),s<3){const f=t[0],a=t[1];if(mr(f),a.condition)mr(a);else{const d=f.condition;(d[0]==="="||d[1]>=0&&(d[0]===">"||d[0]===">="))&&a.tokens.unshift({type:"minus",volatile:!0})}}else t.forEach(mr)}else{if(t.length<4&&o)for(let s=0,f=t.length;s<f;s++)t[s]===o&&t.splice(s,1);if(t.length<1&&o&&(t[0]=wt(hn("General")),t[0].generated=!0),t.length<2){const s={type:"minus",volatile:!0};t.push(po(t[0],s))}if(t.length<3&&t.push(po(t[0])),t.length<4)if(o)t.push(o);else{const s=wt(hn("@"));s.generated=!0,t.push(s)}t[0].condition=[">",0],t[1].condition=["<",0],t[2].condition=null}return{pattern:e,partitions:t,locale:r}}const Io=Object.create({});function No(e,t=!1){e||(e="General");let n=Io[e];if(!n)try{n=A1(e),Io[e]=n}catch(r){if(t)throw r;const o={tokens:[{type:"error"}],error:r.message};n={pattern:e,partitions:[o,o,o,o],error:r.message,locale:null}}return n}function C1(e,t,n={}){const r=Object.assign({},ba,n),o=No(e,r.throws),l=r1(t,r)??t;return g1(l,o,r)}function v1(e){const t=No(e,!1);return d1(t.partitions)}function T1(e,t,n={}){return C1(t??An,e,n)}function R1(e){return v1(e)}const M1={smallFont:"10px",font:"12px",largeFont:"14px",padding:"12px",lineHeight:1.5,mediumPadding:"8px",largePadding:"20px",borderRadius:"4px",borderLargeRadius:"8px",tinyPadding:"4px",scrollBarContent:"30px",formulaBarHeight:"25px"},O1={lowZIndex:2,middleZIndex:3,highZIndex:4,highestZIndex:9};function D1(e,t,n,r){t=t/100,n=n/100;const o=(1-Math.abs(2*n-1))*t,l=o*(1-Math.abs(e/60%2-1)),c=n-o/2;let h,g,u;if(e>=0&&e<60?(h=o,g=l,u=0):e>=60&&e<120?(h=l,g=o,u=0):e>=120&&e<180?(h=0,g=o,u=l):e>=180&&e<240?(h=0,g=l,u=o):e>=240&&e<300?(h=l,g=0,u=o):(h=o,g=0,u=l),r!==1)return`rgba(${h},${g},${u},${r})`;const i=d=>Math.round(d*255).toString(16).padStart(2,"0"),s=i(h+c),f=i(g+c),a=i(u+c);return`#${s}${f}${a}`}const w1=(e,t,n)=>{const r=e/255,o=t/255,l=n/255,c=Math.min(r,o,l),h=Math.max(r,o,l),g=h-c,u=L1(r,o,l),i=P1(h,c),s=S1(g,i);return[u,s*100,i*100]},L1=(e,t,n)=>{const r=Math.max(e,t,n),o=Math.min(e,t,n),l=r-o;let c=0;if(l==0)c=0;else switch(r){case e:{let u=(t-n)/l,i=0;u<0&&(i=360/60),c=u+i;break}case t:{let u=(n-e)/l,i=120/60;c=u+i;break}case n:let h=(e-t)/l,g=240/60;c=h+g;break}return c*60},S1=(e,t)=>e==0?0:e/(1-Math.abs(2*t-1)),P1=(e,t)=>(e+t)/2,Ao=([e,t,n,r])=>{const[o,l,c]=w1(e,t,n);return D1(o,l*.9,100-c,r)};function U1(e){e.startsWith("#")&&(e=e.slice(1));let t=parseInt(e.substring(0,2),16),n=parseInt(e.substring(2,4),16),r=parseInt(e.substring(4,6),16),o=parseInt(e.substring(6,8)||"ff",16);return[t,n,r,o/255]}const y1=e=>{if(e.startsWith("rgb")){const n=e.replace("rgb","").replace("a","").replace("(","").replace(")","").split(",").map(Number);return n.length<=3&&n.push(1),Ao(n)}else return Ao(U1(e))},Co=Object.freeze({primaryColor:"#217346",buttonActiveColor:"#c6c6c6",selectionColor:"rgba(198,198,198,0.3)",backgroundColor:"#e6e6e6",hoverColor:"rgba(0, 0, 0, 0.04)",scrollbarColor:"rgba(0, 0, 0, 0.1)",scrollbarHoveColor:"rgba(0, 0, 0, 0.2)",dialogBackground:"rgba(0, 0, 0, 0.6)",white:"#ffffff",black:"#000000",triangleFillColor:"#dddddd",contentColor:"#333333",borderColor:"#cccccc",activeBorderColor:"#808080",errorFormulaColor:"#ff0000"}),F1=Object.freeze(Object.fromEntries(Object.entries(Co).map(([e,t])=>[e,y1(t)]))),pr={...M1,...O1},H1="data-theme";function V1(){if(typeof sessionStorage<"u"){const e=sessionStorage.getItem(H1);if(e&&(e==="dark"||e==="light"))return e}return typeof matchMedia=="function"&&matchMedia("(prefers-color-scheme: dark)").matches?"dark":"light"}function se(e,t){return t==="dark"||V1()==="dark"?F1[e]:Co[e]}const vo=new Map;function k1(e,t){const n=`${t}__${e.font}`,r=vo.get(n);if(r)return r;const o=e.measureText(t),{actualBoundingBoxAscent:l,actualBoundingBoxDescent:c}=o,h=l+c,g=o.width,u=Math.ceil(g/Ft()),i=Math.ceil(h/Ft()),s={width:u,height:i};return vo.set(n,s),s}function ve(e,t,n,r,o){e.fillRect(x(t),x(n),x(r),x(o))}function Ir(e,t,n,r,o){e.strokeRect(x(t),x(n),x(r),x(o))}function To(e,t,n,r,o){e.clearRect(x(t),x(n),x(r),x(o))}function Nr(e,t,n,r){e.fillText(t,x(n),x(r))}function Ke(e,t){if(t.length!==0){e.beginPath();for(let n=0;n<t.length;n+=2){const r=t[n],o=t[n+1];e.moveTo(x(r[0]),x(r[1])),e.lineTo(x(o[0]),x(o[1]))}e.stroke()}}function Ro(e,t,n,r){e.beginPath(),e.moveTo(x(t[0]),x(t[1])),e.lineTo(x(n[0]),x(n[1])),e.lineTo(x(r[0]),x(r[1])),e.fill()}function Mo(e,t,n){const r=e.slice(),[o,l]=e,c=q*2,h=n?c:0;return t==="bottom"?r.push([o[0]+h,o[1]-c],[l[0]-h,l[1]-c]):t==="top"?r.push([o[0]+h,o[1]+c],[l[0]-h,l[1]+c]):t==="left"?r.push([o[0]+c,o[1]+h],[l[0]+c,l[1]-h]):t==="right"&&r.push([o[0]-c,o[1]+h],[l[0]-c,l[1]-h]),r}function B1(e,t,n){return n?t?e.split(Nn):Yr(e.replaceAll(Nn,"")):Yr(e)}function G1(e,t,n,r,o){e.setLineDash([x(8),x(6)]);const l=q;Ir(e,t+l,n+l,r-l*2,o-l*2),e.setLineDash([])}function z1(e){let t=[];return e==="hair"?t=[q,q]:e==="dotted"||e==="mediumDashed"?t=[q*2,q*2]:e==="dashed"?t=[q*4,q*4]:e==="dashDot"||e==="mediumDashDot"?t=[q*4,q*4,q*8,q*4]:(e==="dashDotDot"||e==="mediumDashDotDot")&&(t=[q*4,q*4,q*8,q*4,q*4,q*4]),t}function gn(e,t,n,r,o){if(!n)return;const{top:l,left:c,width:h,height:g}=t;let u=[];r==="top"?u=[[c,l],[c+h,l]]:r==="bottom"?u=[[c,l+g],[c+h,l+g]]:r==="left"?u=[[c,l],[c,l+g]]:r==="right"&&(u=[[c+h,l],[c+h,l+g]]);const{type:i,color:s}=n;e.lineWidth=Ur[i],e.strokeStyle=s||se("black",o);const f=z1(i);i==="double"&&(u=Mo(u,r,!0)),f.length>0&&e.setLineDash(f.map(a=>x(a))),Ke(e,u),f.length>0&&e.setLineDash([])}function Y1(e,t,n,r,o){const{value:l,formula:c,...h}=n,g={height:0,width:0};if(l===""&&qo(h))return g;const u=n?.numberFormat||An,i=u===An&&typeof l=="number",s=T1(l,u),f=n?.fontSize?n.fontSize:Ut,a=Vr(n?.isItalic?"italic":"normal",n?.isBold?"bold":"500",x(f),n?.fontFamily);n?.fillColor&&(e.fillStyle=n?.fillColor,ve(e,t.left,t.top,t.width,t.height));let d=n?.fontColor||se("contentColor",o);nn(s)&&(d=se("errorFormulaColor",o)),e.font=a,e.fillStyle=d;const E={...n};let m=E?.horizontalAlign;E?.horizontalAlign===void 0&&i&&(m=Pt.RIGHT),E.horizontalAlign=m;const p=(!E?.isWrapText&&R1(u)?[s]:B1(s,n?.isWrapText,r)).map(O=>{const L=k1(e,O);return{str:O,width:L.width,height:L.height===0?f:L.height}}),{width:I,height:A,resultList:M}=W1(p,t,E,$o(!!r,s));if(I>0&&A>0){const O=Math.ceil(f*(pr.lineHeight-1)/2);let L=[];for(const U of M){if(Nr(e,U.text,U.x,U.y),E?.underline){e.strokeStyle=d;const H=U.y+U.height+O/2,y=[[U.x,H],[U.x+U.width,H]];E?.underline===Rr.DOUBLE?L=L.concat(Mo(y,"bottom",!1)):L=L.concat(y)}if(E?.isStrike){e.strokeStyle=d;const H=U.y+U.height/2+O/2;L=L.concat([[U.x,H],[U.x+U.width,H]])}}Ke(e,L)}return g.height=Math.ceil(A),g.width=Math.ceil(I),g}function W1(e,t,n,r){const o=n?.fontSize?n.fontSize:Ut,l=Math.ceil(o*(pr.lineHeight-1)/2),c=n?.verticalAlign??St.MIDDLE,{left:h,top:g,height:u}=t,i=Math.max(t.width,...e.map(d=>d.width)),s=[];let f=0,a=0;if(n?.isWrapText&&r){let d=0;for(let E=0;E<e.length;E++){const m=e[E];f=Math.max(f,m.width);const C=m.height+l*2;a+=C,s.push({text:m.str,x:0,y:d,width:m.width,height:m.height}),d+=C}}else if(n?.isWrapText){let d=0;for(let E=0;E<e.length;){let m=i,C="",v=0,p=0;for(;E<e.length;){const I=e[E];if(m>=I.width)v+=I.width,p=Math.max(p,I.height),m-=I.width,C+=I.str,E++;else break}if(C){f=Math.max(f,v);const I=p+l*2;a+=I,s.push({text:C,x:0,y:d,width:v,height:p}),d+=I}}}else{let d="",E=i;for(let m=0;m<e.length;m++){const C=e[m];if(E>=C.width)f+=C.width,a=Math.max(a,C.height),d+=C.str,E-=C.width;else break}s.push({text:d,x:0,y:0,width:f,height:a})}if(a=Math.max(a,o*pr.lineHeight),f+=l,a+=l,f<=i&&a<=u){let d=h+l,E=g+(u-a)/2+l;c===St.TOP?E=g+l:c===St.BOTTOM&&(E=g+(u-a)+l),n?.horizontalAlign===Pt.CENTER?d=h+(i-f)/2:n?.horizontalAlign===Pt.RIGHT&&(d=h+(i-f)-l);for(const m of s)m.x=m.x+d,m.y=m.y+E}return{width:f,height:a,resultList:s}}function Oo(e){return{textAlign:"center",textBaseline:"middle",font:Vr(void 0,"500",x(Ut)),fillStyle:se("black",e),lineWidth:q,strokeStyle:se("borderColor",e)}}const Do=Math.max(...Object.values(Ur));class q1{canvas;ctx;width=0;height=0;isRendering=!1;rowMap={};colMap={};eventData={theme:"light",sheetData:{},canvasSize:{top:0,left:0,width:0,height:0},headerSize:{width:0,height:0},currentSheetInfo:{isHide:!1,rowCount:0,colCount:0,name:"",sheetId:"",tabColor:"",sort:1},scroll:{row:0,col:0,top:0,left:0,scrollLeft:0,scrollTop:0},range:{row:0,col:0,rowCount:1,colCount:1,sheetId:""},copyRange:void 0,customHeight:{},customWidth:{},currentMergeCells:[],autoFilter:void 0};constructor(t){this.canvas=t,this.ctx=t.getContext("2d");const n=Ft();this.ctx.scale(n,n)}render(t){if(t.changeSet.size===0||this.isRendering)return;this.isRendering=!0,this.eventData=t,this.clear();const{ctx:n}=this;n.strokeStyle=se("primaryColor",t.theme),n.fillStyle=se("white",t.theme),n.lineWidth=q*2;const{width:r,height:o}=this.eventData.canvasSize,l=this.eventData.headerSize,{endRow:c,contentHeight:h}=this.renderRowsHeader(o),{endCol:g,contentWidth:u}=this.renderColsHeader(r);this.renderGrid(r-l.width,o-l.height),this.renderTriangle(),this.renderMergeCell(),this.ctx.fillStyle=se("selectionColor",this.eventData.theme);const i=this.renderSelection({endRow:c,endCol:g,contentHeight:h,contentWidth:u});return this.renderAntLine(i),this.renderContent({endRow:c,endCol:g,contentHeight:h,contentWidth:u}),this.ctx.lineWidth=Do,Ir(this.ctx,i.left,i.top,i.width,i.height),this.isRendering=!1,{rowMap:{...this.rowMap},colMap:{...this.colMap}}}resize(t){this.width=t.width,this.height=t.height,this.canvas.width=x(t.width),this.canvas.height=x(t.height)}clear(){To(this.ctx,0,0,this.width,this.height)}renderRowsHeader(t){const{row:n}=this.eventData.scroll,r=this.eventData.headerSize,{rowCount:o}=this.eventData.currentSheetInfo;this.ctx.save();const l=this.eventData.range;ve(this.ctx,0,r.height,r.width,t),Object.assign(this.ctx,Oo(this.eventData.theme));const c=[];let h=r.height,g=n;for(;g<o&&h<t;g++){const i=this.getRowHeight(g);let s=h;if(g===n&&(s+=q/2),c.push([0,s],[r.width,s]),i>0){const f=this.isHighlightRow(l,g);this.ctx.fillStyle=f?se("primaryColor",this.eventData.theme):se("black",this.eventData.theme),Nr(this.ctx,String(g+1),r.width/2,s+i/2)}h+=i}c.push([0,h],[r.width,h]),c.push([0,0],[0,h]),Ke(this.ctx,c),this.ctx.restore();const u=g>=o?h:t;return{endRow:g,contentHeight:Math.floor(u)}}renderColsHeader(t){const{col:n,row:r}=this.eventData.scroll,o=this.eventData.headerSize,{colCount:l}=this.eventData.currentSheetInfo,c=this.eventData.range,h=[];this.ctx.save(),ve(this.ctx,o.width,0,t,o.height),Object.assign(this.ctx,Oo());let g=o.width,u=n;const i=this.eventData.autoFilter;let s=0;for(let a=r;a<this.eventData.currentSheetInfo.rowCount&&(s=this.getRowHeight(r),s===0);a++);for(;u<l&&g<=t;u++){const a=this.getColWidth(u);let d=g;if(u===n&&(d+=q/2),h.push([d,0],[d,o.height]),a>0){const E=this.isHighlightCol(c,u);this.ctx.fillStyle=E?se("primaryColor",this.eventData.theme):se("black",this.eventData.theme),Nr(this.ctx,bo(u),d+a/2,o.height/2)}i&&u>=i.range.col&&u<i.range.col+i.range.colCount&&this.renderFilter(g+a,o.height+s,i.col===u),g+=a}h.push([g,0],[g,o.height]),h.push([0,0],[g,0]),Ke(this.ctx,h),this.ctx.restore();const f=u>=l?g:t;return{endCol:u,contentWidth:Math.floor(f)}}renderFilter(t,n,r){const o=Vo,l=q*4,c=t-o-l,h=n-o-l;this.ctx.fillStyle=r?se("primaryColor",this.eventData.theme):se("borderColor",this.eventData.theme),this.ctx.strokeStyle=r?se("primaryColor",this.eventData.theme):se("borderColor",this.eventData.theme),Ir(this.ctx,c,h,o,o),Ro(this.ctx,[c+o/4,h+o*3/8],[c+o*3/4,h+o*3/8],[c+o/2,h+o*5/8])}renderGrid(t,n){const r=this.eventData.headerSize,{row:o,col:l}=this.eventData.scroll,{rowCount:c,colCount:h}=this.eventData.currentSheetInfo;this.ctx.save(),this.ctx.lineWidth=q,this.ctx.strokeStyle=se("borderColor",this.eventData.theme),this.ctx.translate(x(r.width),x(r.height));const g=[];let u=0,i=0;const s=[],f=[];for(let a=o;a<c&&u<=n;a++){for(;a<c&&this.getRowHeight(a)===0;)a++;f.push([0,u]);const d=this.getRowHeight(a);u+=d}for(let a=l;a<h&&i<=t;a++){for(;a<h&&this.getColWidth(a)===0;)a++;s.push([i,0]);const d=this.getColWidth(a);i+=d}for(const a of s)g.push(a,[a[0],u]);for(const a of f)g.push(a,[i,a[1]]);g.push([0,u],[i,u]),g.push([i,0],[i,u]),Ke(this.ctx,g),this.ctx.restore()}renderTriangle(){const t=this.eventData.headerSize;this.ctx.save(),ve(this.ctx,0,0,t.width,t.height),this.ctx.fillStyle=se("triangleFillColor",this.eventData.theme);const n=2,r=Math.floor(n),o=Math.floor(t.height-n),l=Math.floor(t.width*.4),c=Math.floor(t.width-n);Ro(this.ctx,[c,r],[l,o],[c,o]),this.ctx.restore()}getRowHeight(t){const n=Hr(this.eventData.currentSheetInfo.sheetId,t),r=this.eventData.customHeight[n];return r?r.isHide?Dr:r.len:Mr}getColWidth(t){const n=Hr(this.eventData.currentSheetInfo.sheetId,t),r=this.eventData.customWidth[n];return r?r.isHide?Dr:r.len:Or}isHighlightRow(t,n){return!!(dt(t)||mt(t)||n>=t.row&&n<t.row+t.rowCount)}isHighlightCol(t,n){return!!(dt(t)||Et(t)||n>=t.col&&n<t.col+t.colCount)}renderAntLine(t){const n=this.eventData.copyRange;!n||n.sheetId!==this.eventData.currentSheetInfo.sheetId||(this.ctx.lineWidth=Do,this.ctx.strokeStyle=se("primaryColor",this.eventData.theme),G1(this.ctx,t.left,t.top,t.width,t.height))}renderMergeCell(){const t=this.eventData.currentMergeCells;if(t.length===0)return;const n=this.eventData.range;for(const r of t)n.row===r.row&&n.col===r.col||this.clearRect(r)}getCellSize(t){let{row:n,col:r,colCount:o,rowCount:l}=t,c=n,h=r,g=n+l,u=r+o;const i=this.eventData.currentSheetInfo;dt(t)?(h=0,u=i.colCount,c=0,g=i.rowCount):mt(t)?(c=0,g=i.rowCount):Et(t)&&(h=0,u=i.colCount);let s=0,f=0;for(;c<g;c++)f+=this.getRowHeight(c);for(;h<u;h++)s+=this.getColWidth(h);return{width:s,height:f}}computeCellPosition(t){const{row:n,col:r}=t,o=this.eventData.headerSize,l=this.eventData.scroll;let c=o.width,h=o.height,g=l.row,u=l.col;if(r>=l.col)for(;u<r;)c+=this.getColWidth(u),u++;else for(c=-o.width;u>r;)c-=this.getColWidth(u),u--;if(n>=l.row)for(;g<n;)h+=this.getRowHeight(g),g++;else for(h=-o.height;g>n;)h-=this.getRowHeight(g),g--;return{top:h,left:c}}getActiveRange(t){const n=t||this.eventData.range,r=this.eventData.currentMergeCells;for(const o of r)if(jo(n,o))return{range:{...o,sheetId:o.sheetId},isMerged:!0};return{range:n,isMerged:!1}}clearRect(t){const n=this.getCellSize(t);if(n.width<=0||n.height<=0)return;const r=this.computeCellPosition(t),o=q;To(this.ctx,r.left+o,r.top+o,n.width-o*2,n.height-o*2)}renderContent(t){const{endCol:n,endRow:r,contentHeight:o,contentWidth:l}=t,{ctx:c}=this;c.textAlign="left",c.textBaseline="top",c.lineWidth=q*2;const h=this.eventData.headerSize,{row:g,col:u}=this.eventData.scroll,i=Math.floor(l-h.width),s=Math.floor(o-h.height);c.save(),this.rowMap={},this.colMap={};const f=this.eventData.currentMergeCells;for(let a=g;a<r;a++)for(let d=u;d<n;d++){const E=f.find(m=>m.row===a&&m.col===d);this.renderCell(a,d,E,i,s)}c.restore()}renderCell(t,n,r,o,l){const{ctx:c}=this,h={row:t,col:n,rowCount:1,colCount:1,sheetId:""},g=Fr(this.eventData.currentSheetInfo.sheetId,t,n),u=this.eventData.sheetData[g];if(!u)return;const i=this.getCellSize(r||h);if(i.width<=0||i.height<=0)return;const s=this.computeCellPosition(h);c.lineWidth=q*2;const f=this.eventData.theme,a=Y1(c,{top:s.top,left:s.left,width:Math.min(i.width,o),height:Math.min(i.height,l)},u,!!r,f),d=Math.max(this.rowMap[t]??0,a.height),E=Math.max(this.colMap[n]??0,a.width);d>=Mr&&(this.rowMap[t]=d),E>=Or&&(this.colMap[n]=E);const m={top:s.top,left:s.left,height:Math.max(d,i.height),width:Math.max(E,i.width)};gn(c,m,u.borderTop,"top",f),gn(c,m,u.borderBottom,"bottom",f),gn(c,m,u.borderLeft,"left",f),gn(c,m,u.borderRight,"right",f)}renderSelection(t){const n=this.eventData.range;return dt(n)?this.renderSelectAll(t):mt(n)?this.renderSelectCol(t):Et(n)?this.renderSelectRow(t):this.renderSelectRange()}renderSelectRange(){const t=this.eventData.headerSize,n=this.eventData.range,r=this.computeCellPosition({row:n.row,col:n.col,rowCount:1,colCount:1,sheetId:""}),o=n.row+n.rowCount-1,l=n.col+n.colCount-1,c={row:o,col:l,rowCount:1,colCount:1,sheetId:""},h=this.computeCellPosition(c),g=this.getCellSize(c),u=h.left+g.width-r.left,i=h.top+g.height-r.top;ve(this.ctx,r.left,0,u,t.height),ve(this.ctx,0,r.top,t.width,i);const s=n.rowCount>1||n.colCount>1;s&&ve(this.ctx,r.left,r.top,u,i);const f=[[r.left,t.height],[r.left+u,t.height]];return f.push([t.width,r.top],[t.width,r.top+i]),Ke(this.ctx,f),s&&this.renderActiveCell(),{left:r.left,top:r.top,width:u,height:i}}renderSelectAll(t){const{contentHeight:n,contentWidth:r}=t,o=this.eventData.headerSize;ve(this.ctx,0,0,r,n),this.renderActiveCell();const l=r-o.width,c=n-o.height;return{left:o.width,top:o.height,width:l,height:c}}renderSelectCol({contentHeight:t}){const n=this.eventData.headerSize,r=this.eventData.range,o=this.computeCellPosition(r);let l=0;for(let g=r.col,u=r.col+r.colCount;g<u;g++)l+=this.getColWidth(g);const c=t-n.height;ve(this.ctx,o.left,0,l,t),ve(this.ctx,0,o.top,n.width,c);const h=[[n.width,n.height],[n.width,t]];return Ke(this.ctx,h),this.renderActiveCell(),{left:o.left,top:n.height,width:l,height:c}}renderSelectRow({contentWidth:t}){const n=this.eventData.headerSize,r=this.eventData.range,o=this.computeCellPosition(r);let l=0;for(let g=r.row,u=r.row+r.rowCount;g<u;g++)l+=this.getRowHeight(g);const c=t-n.width-q;ve(this.ctx,o.left,0,c,n.height),ve(this.ctx,0,o.top,t,l);const h=[[o.left,n.height],[t,n.height]];return Ke(this.ctx,h),this.renderActiveCell(),{left:n.width,top:o.top,width:c,height:l}}renderActiveCell(){const t=this.eventData.range,n=this.getActiveRange({row:t.row,col:t.col,rowCount:1,colCount:1,sheetId:t.sheetId}).range;this.clearRect(n)}}let Ar=null;const K1={init(e){Ar=new q1(e.canvas),Jo(e.dpr)},resize(e){Ar?.resize(e)},async render(e,t){const n=Ar?.render(e);n&&await t(n)},async computeFormulas(e,t){return await Ya(e,t)}};const wo=Symbol("Comlink.proxy"),X1=Symbol("Comlink.endpoint"),$1=Symbol("Comlink.releaseProxy"),Cr=Symbol("Comlink.finalizer"),dn=Symbol("Comlink.thrown"),Lo=e=>typeof e=="object"&&e!==null||typeof e=="function",x1={canHandle:e=>Lo(e)&&e[wo],serialize(e){const{port1:t,port2:n}=new MessageChannel;return vr(e,t),[n,[n]]},deserialize(e){return e.start(),b1(e)}},Q1={canHandle:e=>Lo(e)&&dn in e,serialize({value:e}){let t;return e instanceof Error?t={isError:!0,value:{message:e.message,name:e.name,stack:e.stack}}:t={isError:!1,value:e},[t,[]]},deserialize(e){throw e.isError?Object.assign(new Error(e.value.message),e.value):e.value}},So=new Map([["proxy",x1],["throw",Q1]]);function J1(e,t){for(const n of e)if(t===n||n==="*"||n instanceof RegExp&&n.test(t))return!0;return!1}function vr(e,t=globalThis,n=["*"]){t.addEventListener("message",function r(o){if(!o||!o.data)return;if(!J1(n,o.origin)){console.warn(`Invalid origin \'${o.origin}\' for comlink proxy`);return}const{id:l,type:c,path:h}=Object.assign({path:[]},o.data),g=(o.data.argumentList||[]).map(tt);let u;try{const i=h.slice(0,-1).reduce((f,a)=>f[a],e),s=h.reduce((f,a)=>f[a],e);switch(c){case"GET":u=s;break;case"SET":i[h.slice(-1)[0]]=tt(o.data.value),u=!0;break;case"APPLY":u=s.apply(i,g);break;case"CONSTRUCT":{const f=new s(...g);u=nh(f)}break;case"ENDPOINT":{const{port1:f,port2:a}=new MessageChannel;vr(e,a),u=th(f,[f])}break;case"RELEASE":u=void 0;break;default:return}}catch(i){u={value:i,[dn]:0}}Promise.resolve(u).catch(i=>({value:i,[dn]:0})).then(i=>{const[s,f]=In(i);t.postMessage(Object.assign(Object.assign({},s),{id:l}),f),c==="RELEASE"&&(t.removeEventListener("message",r),Po(t),Cr in e&&typeof e[Cr]=="function"&&e[Cr]())}).catch(i=>{const[s,f]=In({value:new TypeError("Unserializable return value"),[dn]:0});t.postMessage(Object.assign(Object.assign({},s),{id:l}),f)})}),t.start&&t.start()}function Z1(e){return e.constructor.name==="MessagePort"}function Po(e){Z1(e)&&e.close()}function b1(e,t){const n=new Map;return e.addEventListener("message",function(o){const{data:l}=o;if(!l||!l.id)return;const c=n.get(l.id);if(c)try{c(l)}finally{n.delete(l.id)}}),Tr(e,n,[],t)}function En(e){if(e)throw new Error("Proxy has been released and is not useable")}function Uo(e){return ht(e,new Map,{type:"RELEASE"}).then(()=>{Po(e)})}const mn=new WeakMap,pn="FinalizationRegistry"in globalThis&&new FinalizationRegistry(e=>{const t=(mn.get(e)||0)-1;mn.set(e,t),t===0&&Uo(e)});function _1(e,t){const n=(mn.get(t)||0)+1;mn.set(t,n),pn&&pn.register(e,t,e)}function j1(e){pn&&pn.unregister(e)}function Tr(e,t,n=[],r=function(){}){let o=!1;const l=new Proxy(r,{get(c,h){if(En(o),h===$1)return()=>{j1(l),Uo(e),t.clear(),o=!0};if(h==="then"){if(n.length===0)return{then:()=>l};const g=ht(e,t,{type:"GET",path:n.map(u=>u.toString())}).then(tt);return g.then.bind(g)}return Tr(e,t,[...n,h])},set(c,h,g){En(o);const[u,i]=In(g);return ht(e,t,{type:"SET",path:[...n,h].map(s=>s.toString()),value:u},i).then(tt)},apply(c,h,g){En(o);const u=n[n.length-1];if(u===X1)return ht(e,t,{type:"ENDPOINT"}).then(tt);if(u==="bind")return Tr(e,t,n.slice(0,-1));const[i,s]=yo(g);return ht(e,t,{type:"APPLY",path:n.map(f=>f.toString()),argumentList:i},s).then(tt)},construct(c,h){En(o);const[g,u]=yo(h);return ht(e,t,{type:"CONSTRUCT",path:n.map(i=>i.toString()),argumentList:g},u).then(tt)}});return _1(l,e),l}function eh(e){return Array.prototype.concat.apply([],e)}function yo(e){const t=e.map(In);return[t.map(n=>n[0]),eh(t.map(n=>n[1]))]}const Fo=new WeakMap;function th(e,t){return Fo.set(e,t),e}function nh(e){return Object.assign(e,{[wo]:!0})}function In(e){for(const[t,n]of So)if(n.canHandle(e)){const[r,o]=n.serialize(e);return[{type:"HANDLER",name:t,value:r},o]}return[{type:"RAW",value:e},Fo.get(e)||[]]}function tt(e){switch(e.type){case"HANDLER":return So.get(e.name).deserialize(e.value);case"RAW":return e.value}}function ht(e,t,n,r){return new Promise(o=>{const l=rh();t.set(l,o),e.start&&e.start(),e.postMessage(Object.assign({id:l},n),r)})}function rh(){return new Array(4).fill(0).map(()=>Math.floor(Math.random()*Number.MAX_SAFE_INTEGER).toString(16)).join("-")}vr(K1)})();\n//# sourceMappingURL=worker-D-tO9fGm.js.map\n', aa = typeof self < "u" && self.Blob && new Blob(["(self.URL || self.webkitURL).revokeObjectURL(self.location.href);", Xc], { type: "text/javascript;charset=utf-8" });
function Q6(e) {
  let t;
  try {
    if (t = aa && (self.URL || self.webkitURL).createObjectURL(aa), !t) throw "";
    const n = new Worker(t, {
      name: e?.name
    });
    return n.addEventListener("error", () => {
      (self.URL || self.webkitURL).revokeObjectURL(t);
    }), n;
  } catch {
    return new Worker(
      "data:text/javascript;charset=utf-8," + encodeURIComponent(Xc),
      {
        name: e?.name
      }
    );
  }
}
const j6 = () => k.current, Z6 = () => "en-US", eN = ue((e) => {
  const { doc: t, provider: n, awareness: r, docConfig: o } = e, [s, l] = he(void 0), f = Xe(!1), d = jc(
    k.subscribe,
    j6,
    Z6
  );
  return Re(() => {
    if (f.current)
      return;
    f.current = !0, k.init();
    const c = J6({
      worker: new Q6(),
      doc: t || (o ? new $o(o) : new $o())
    });
    l({ controller: c, provider: n, awareness: r });
  }, []), Re(() => {
    const c = ga.includes(d);
    document.documentElement.setAttribute(
      "data-layout-direction",
      c ? "rtl" : "ltr"
    );
  }, [d]), /* @__PURE__ */ I(yl.Provider, { value: s, children: s && /* @__PURE__ */ I(Wc, { ...e }, d, !1, {
    fileName: "/home/user/excel-collab/src/editor.tsx",
    lineNumber: 66,
    columnNumber: 17
  }, void 0) }, void 0, !1, {
    fileName: "/home/user/excel-collab/src/editor.tsx",
    lineNumber: 65,
    columnNumber: 5
  }, void 0);
});
eN.displayName = "Excel";
export {
  dr as $,
  Ct as A,
  oo as B,
  iu as C,
  ou as D,
  xe as E,
  hs as F,
  Ce as G,
  Te as H,
  dt as I,
  ka as J,
  eN as K,
  Wc as L,
  Yo as M,
  Uu as N,
  Nu as O,
  ha as P,
  sN as Q,
  Fn as R,
  El as S,
  Hl as T,
  Wo as U,
  ze as V,
  Qe as W,
  La as X,
  Ae as Y,
  bi as Z,
  Xo as _,
  Ee as a,
  bt as a$,
  Aa as a0,
  Co as a1,
  Eo as a2,
  or as a3,
  da as a4,
  mt as a5,
  ca as a6,
  Qi as a7,
  qc as a8,
  Le as a9,
  mr as aA,
  No as aB,
  dN as aC,
  du as aD,
  ae as aE,
  gN as aF,
  bs as aG,
  yu as aH,
  Ma as aI,
  jo as aJ,
  qu as aK,
  ti as aL,
  Mi as aM,
  ps as aN,
  ms as aO,
  Cu as aP,
  gs as aQ,
  ut as aR,
  jn as aS,
  Es as aT,
  sr as aU,
  Pu as aV,
  Fa as aW,
  Fu as aX,
  Lu as aY,
  tt as aZ,
  fe as a_,
  Ei as aa,
  Be as ab,
  qo as ac,
  ei as ad,
  hN as ae,
  lu as af,
  fs as ag,
  ga as ah,
  fa as ai,
  tu as aj,
  Pn as ak,
  sn as al,
  tn as am,
  ln as an,
  Xr as ao,
  Ki as ap,
  _c as aq,
  wt as ar,
  yl as as,
  Wg as at,
  nu as au,
  pa as av,
  ip as aw,
  F as ax,
  Wl as ay,
  hr as az,
  cu as b,
  mu as b0,
  xt as b1,
  Cl as b2,
  Ta as b3,
  yi as b4,
  hu as b5,
  pu as b6,
  nt as b7,
  Zn as b8,
  J6 as b9,
  Au as bA,
  Su as bB,
  ar as bC,
  Mu as bD,
  cN as bE,
  Ms as bF,
  aN as bG,
  uN as bH,
  uu as bI,
  un as bJ,
  Pe as bK,
  fN as bL,
  xn as bM,
  I6 as bN,
  be as bO,
  Se as bP,
  pn as bQ,
  es as bR,
  vn as bS,
  ir as bT,
  D6 as bU,
  wa as ba,
  Qo as bb,
  Yt as bc,
  Gt as bd,
  On as be,
  Na as bf,
  gu as bg,
  wi as bh,
  lN as bi,
  Wt as bj,
  Ia as bk,
  An as bl,
  va as bm,
  Tn as bn,
  Ca as bo,
  Ii as bp,
  mN as bq,
  vu as br,
  fu as bs,
  ce as bt,
  $n as bu,
  wu as bv,
  Ru as bw,
  Ot as bx,
  Ea as by,
  Ra as bz,
  pN as c,
  io as d,
  V as e,
  nr as f,
  Pa as g,
  su as h,
  rr as i,
  Ko as j,
  au as k,
  ds as l,
  ua as m,
  Ge as n,
  Pc as o,
  Un as p,
  ur as q,
  Cn as r,
  cn as s,
  Z as t,
  ht as u,
  ru as v,
  ma as w,
  Ci as x,
  Xt as y,
  Ji as z
};
//# sourceMappingURL=index-WklHmZ9I.mjs.map
