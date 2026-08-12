import {
  SpreadsheetOperationError,
  cellAddressToLabel,
  cellLabelToAddress,
  type CellAddress,
  type CellModel,
  type CellPrimitive,
  type CellStyle,
  type ChartRangeBinding,
  type ColumnSchema,
  type RowSchema,
  type SheetMerge,
  type WorkbookModel,
  type WorksheetChartObject,
  type WorksheetChartType
} from "@excelsior/core";
import { createRequire } from "node:module";
import { basename, dirname, join, posix as posixPath } from "node:path";
import readExcelFile, {
  parseSheetData,
  type ParseSheetDataResult,
  type Schema,
  type Sheet as ReadSheet
} from "read-excel-file/node";
import writeExcelFile, {
  getSheetData,
  type Cell as WriterCell,
  type Column as WriterColumn,
  type Feature as WriterFeature,
  type Sheet as WriterSheet
} from "write-excel-file/node";

const createXlsxOperationError = (
  code: string,
  message: string,
  details?: Record<string, unknown>
): SpreadsheetOperationError =>
  new SpreadsheetOperationError({
    code,
    message,
    area: "wrapper",
    recoverable: true,
    details
  });

type BorderStyleName = NonNullable<NonNullable<CellStyle["border"]>["top"]>["style"];
type WriterCellObject = {
  align?: CellStyle["align"];
  alignVertical?: CellStyle["alignVertical"];
  backgroundColor?: string;
  textColor?: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: "bold";
  fontStyle?: "italic";
  textDecoration?: { underline: true };
  wrap?: boolean;
  format?: string;
  indent?: number;
  leftBorderColor?: string;
  leftBorderStyle?: BorderStyleName;
  rightBorderColor?: string;
  rightBorderStyle?: BorderStyleName;
  topBorderColor?: string;
  topBorderStyle?: BorderStyleName;
  bottomBorderColor?: string;
  bottomBorderStyle?: BorderStyleName;
  height?: number;
  columnSpan?: number;
  rowSpan?: number;
  type?: "Formula";
  value?: string | number | boolean;
};

const DEFAULT_SHEET_NAME = "Sheet1";
const EXCEL_COLUMN_WIDTH_UNIT = 8;

const workbookDefaults = {
  maxRows: 1000,
  maxColumns: 100,
  maxCellLength: 5000,
  maxFormulaLength: 2048,
  maxPasteCells: 10000,
  maxPivotSourceRows: 5000,
  rowHeight: 28,
  columnWidth: 120,
  viewportBuffer: 6,
  maxHistorySize: 100,
  enableFormulas: true,
  clipboardPolicy: "text-only" as const
};

const EMU_PER_PIXEL = 9525;
const MAX_EXPORT_CHARTS_PER_SHEET = 64;
const MAX_IMPORT_CHARTS_PER_SHEET = 64;
const MAX_CHART_RANGE_CELLS = 100000;
const MAX_CHART_SERIES = 64;
const MAX_CHART_FORMULA_REFERENCES = 512;
const CHART_MIN_SIZE = 40;
const DEFAULT_CHART_WIDTH = 420;
const DEFAULT_CHART_HEIGHT = 260;
const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DRAWING_MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

type WorkbookSheet = WorkbookModel["sheets"][number];

interface ParsedRelationship {
  id: string;
  type: string;
  target: string;
}

interface CellBounds {
  start: CellAddress;
  end: CellAddress;
}

interface ParsedFormulaRange {
  sheetName: string;
  start: CellAddress;
  end: CellAddress;
}

interface ParsedChartAnchor {
  from: {
    row: number;
    col: number;
    rowOffsetEmu: number;
    colOffsetEmu: number;
  };
  to?: {
    row: number;
    col: number;
    rowOffsetEmu: number;
    colOffsetEmu: number;
  };
  ext?: {
    widthEmu: number;
    heightEmu: number;
  };
  chartRelationId: string;
}

interface ParsedChartDefinition {
  chartTypeTag: string;
  mappedType: WorksheetChartType;
  title?: string;
  legendVisible: boolean;
  formulas: string[];
  xAxisTitle?: string;
  yAxisTitle?: string;
  xAxisVisible: boolean;
  yAxisVisible: boolean;
  xAxisType: "linear" | "log" | "date" | "category";
  yAxisType: "linear" | "log" | "date" | "category";
  seriesNames: string[];
}

interface ChartExportSeries {
  name?: string;
  nameRef?: string;
  categoryRef?: string;
  valueRef: string;
}

interface ChartExportAnchor {
  row: number;
  col: number;
  rowOffsetPx: number;
  colOffsetPx: number;
  widthPx: number;
  heightPx: number;
}

interface ChartExportEntryBase {
  chartId: string;
  drawingObjectId: number;
  anchor: ChartExportAnchor;
}

interface ChartExportChartEntry extends ChartExportEntryBase {
  kind: "chart";
  chartPath: string;
  relationId: string;
  chartXml: string;
}

interface ChartExportPlaceholderEntry extends ChartExportEntryBase {
  kind: "placeholder";
  placeholderText: string;
}

type ChartExportEntry = ChartExportChartEntry | ChartExportPlaceholderEntry;

type WriterSheetWithCharts = WriterSheet<any> & {
  chartsForExport?: ChartExportEntry[];
};

const SUPPORTED_IMPORT_TYPES = new Set<WorksheetChartType>(["column", "bar", "line", "area", "pie", "donut", "scatter"]);

const EXPORT_PROFILE_BY_TYPE: Record<WorksheetChartType, "line" | "area" | "bar-col" | "bar-row" | "pie" | "donut" | "scatter"> = {
  column: "bar-col",
  bar: "bar-row",
  line: "line",
  area: "area",
  pie: "pie",
  donut: "donut",
  scatter: "scatter",
  histogram: "bar-col",
  box: "line",
  heatmap: "bar-col",
  candlestick: "line",
  waterfall: "bar-col",
  funnel: "bar-col",
  polar: "line",
  treemap: "bar-col",
  sunburst: "pie",
  sankey: "bar-row",
  surface: "line",
  surface3d: "line",
  scatter3d: "line",
  unknown: "line"
};

const createSheetId = (index: number): string => `sheet-import-${index + 1}`;

const getCellKey = (row: number, col: number): string => `${row}:${col}`;

const normalizeImportedFormula = (formula: string): string =>
  formula.startsWith("=") ? formula : `=${formula}`;

const toBuffer = (input: Uint8Array | ArrayBuffer): Buffer =>
  Buffer.from(input instanceof Uint8Array ? input : new Uint8Array(input));

const getElementChildren = (element: Element, tagName?: string): Element[] => {
  const children: Element[] = [];
  for (let index = 0; index < element.childNodes.length; index += 1) {
    const child = element.childNodes[index];
    if (child.nodeType !== 1) {
      continue;
    }

    const childElement = child as Element;
    const normalizedTag = childElement.tagName.replace(/^.+:/, "");
    if (!tagName || normalizedTag === tagName) {
      children.push(childElement);
    }
  }
  return children;
};

const getFirstElementChild = (element: Element, tagName: string): Element | undefined =>
  getElementChildren(element, tagName)[0];

const getColorHex = (value: string | null | undefined): string | undefined => {
  if (!value) {
    return undefined;
  }

  const normalized = value.length === 8 ? value.slice(2) : value;
  return /^[0-9a-fA-F]{6}$/.test(normalized) ? `#${normalized.toUpperCase()}` : undefined;
};

const getColorFromElement = (element?: Element): string | undefined =>
  getColorHex(element?.getAttribute("rgb"));

const toWriterColor = (color?: string): string | undefined => {
  if (!color) {
    return undefined;
  }

  return color.startsWith("#") ? color.toUpperCase() : `#${color.toUpperCase()}`;
};

const parseBorderSide = (element?: Element) => {
  if (!element) {
    return undefined;
  }

  const style = element.getAttribute("style") as BorderStyleName | null;
  const color = getColorFromElement(getFirstElementChild(element, "color"));
  if (!style && !color) {
    return undefined;
  }

  return {
    style: style ?? undefined,
    color
  };
};

const stripUndefinedStyle = (style: CellStyle | undefined): CellStyle | undefined => {
  if (!style) {
    return undefined;
  }

  const cleaned = Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined)
  ) as CellStyle;

  if (
    cleaned.border &&
    !cleaned.border.left &&
    !cleaned.border.right &&
    !cleaned.border.top &&
    !cleaned.border.bottom
  ) {
    delete cleaned.border;
  }

  return Object.keys(cleaned).length ? cleaned : undefined;
};

const getCellStyle = (sheetStyle: CellStyle | undefined, rowStyle?: CellStyle, columnStyle?: CellStyle): CellStyle | undefined =>
  stripUndefinedStyle({
    ...columnStyle,
    ...rowStyle,
    ...sheetStyle,
    border: {
      ...columnStyle?.border,
      ...rowStyle?.border,
      ...sheetStyle?.border
    }
  });

const mapStyleToWriterCell = (style?: CellStyle): WriterCellObject => {
  if (!style) {
    return {};
  }

  return {
    align: style.align,
    alignVertical: style.alignVertical,
    backgroundColor: toWriterColor(style.backgroundColor),
    textColor: toWriterColor(style.textColor),
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight === "normal" ? undefined : style.fontWeight,
    fontStyle: style.fontStyle === "normal" ? undefined : style.fontStyle,
    textDecoration: style.underline ? { underline: true } : undefined,
    wrap: style.wrap,
    format: style.format,
    indent: style.indent,
    leftBorderColor: toWriterColor(style.border?.left?.color),
    leftBorderStyle: style.border?.left?.style,
    rightBorderColor: toWriterColor(style.border?.right?.color),
    rightBorderStyle: style.border?.right?.style,
    topBorderColor: toWriterColor(style.border?.top?.color),
    topBorderStyle: style.border?.top?.style,
    bottomBorderColor: toWriterColor(style.border?.bottom?.color),
    bottomBorderStyle: style.border?.bottom?.style
  };
};

const toWriterCell = (
  cell: CellModel | undefined,
  style: CellStyle | undefined,
  rowHeight?: number,
  merge?: SheetMerge
): WriterCell => {
  const writerCell: WriterCellObject = {
    ...mapStyleToWriterCell(style)
  };

  if (typeof rowHeight === "number") {
    writerCell.height = rowHeight;
  }

  if (merge) {
    writerCell.columnSpan = merge.end.col - merge.start.col + 1;
    writerCell.rowSpan = merge.end.row - merge.start.row + 1;
  }

  if (cell?.formula) {
    writerCell.type = "Formula";
    writerCell.value = cell.formula;
    return writerCell as WriterCell;
  }

  if (cell?.value !== undefined && cell.value !== null) {
    writerCell.value = cell.value as string | number | boolean;
    return writerCell as WriterCell;
  }

  if (Object.keys(writerCell).length) {
    return writerCell as WriterCell;
  }

  return null;
};

const findMergeAnchorMap = (merges: SheetMerge[]) => {
  const anchors = new Map<string, SheetMerge>();
  const covered = new Set<string>();
  for (const merge of merges) {
    anchors.set(getCellKey(merge.start.row, merge.start.col), merge);
    for (let row = merge.start.row; row <= merge.end.row; row += 1) {
      for (let col = merge.start.col; col <= merge.end.col; col += 1) {
        if (row === merge.start.row && col === merge.start.col) {
          continue;
        }
        covered.add(getCellKey(row, col));
      }
    }
  }
  return { anchors, covered };
};

const getBounds = (sheet: WorkbookModel["sheets"][number]) => {
  let rowCount = sheet.rowCount;
  let columnCount = sheet.columnCount;

  for (const key of Object.keys(sheet.cells)) {
    const [rowText, colText] = key.split(":");
    rowCount = Math.max(rowCount, Number(rowText) + 1);
    columnCount = Math.max(columnCount, Number(colText) + 1);
  }

  for (const merge of sheet.merges) {
    rowCount = Math.max(rowCount, merge.end.row + 1);
    columnCount = Math.max(columnCount, merge.end.col + 1);
  }

  for (const key of Object.keys(sheet.rows)) {
    rowCount = Math.max(rowCount, Number(key) + 1);
  }

  for (const key of Object.keys(sheet.columns)) {
    columnCount = Math.max(columnCount, Number(key) + 1);
  }

  return {
    rowCount: Math.max(1, rowCount),
    columnCount: Math.max(1, columnCount)
  };
};

const normalizeArchivePath = (path: string): string => path.replace(/\\/g, "/").replace(/^\/+/, "");

const resolveArchivePath = (basePath: string, targetPath: string): string => {
  if (!targetPath) {
    return "";
  }
  if (targetPath.startsWith("/")) {
    return normalizeArchivePath(targetPath.slice(1));
  }
  return normalizeArchivePath(posixPath.normalize(posixPath.join(posixPath.dirname(basePath), targetPath)));
};

const toEmu = (pixels: number): number => Math.max(0, Math.round((Number.isFinite(pixels) ? pixels : 0) * EMU_PER_PIXEL));

const fromEmu = (emus: number): number => Math.max(0, Math.round((Number.isFinite(emus) ? emus : 0) / EMU_PER_PIXEL));

const clampNumber = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

const sanitizeXmlText = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&apos;");

const sanitizeXmlAttribute = (value: string): string => sanitizeXmlText(value);

const sanitizeSpreadsheetText = (value: string, maxLength = 240): string => {
  const normalized = value
    .replace(/[\u0000-\u001f\u007f]/g, " ")
    .replace(/</g, "‹")
    .replace(/>/g, "›")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLength);
  if (!normalized) {
    return "";
  }
  return /^[=+\-@]/.test(normalized) ? `'${normalized}` : normalized;
};

const isSafeFormulaReference = (value: string): boolean =>
  /^(?:'((?:[^']|'')+)'|[^'!]+)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$/i.test(value.trim().replace(/^=/, ""));

const escapeSheetNameForFormula = (sheetName: string): string => `'${sheetName.replaceAll("'", "''")}'`;

const isFiniteInteger = (value: number): boolean => Number.isFinite(value) && Number.isInteger(value);

const normalizeCellLabel = (label: string): string => label.replaceAll("$", "").trim().toUpperCase();

const parseRangeAddress = (rangeAddress: string): CellBounds | undefined => {
  const parts = rangeAddress
    .split(":")
    .map((part) => normalizeCellLabel(part))
    .filter(Boolean);
  if (!parts.length || parts.length > 2) {
    return undefined;
  }

  try {
    const start = cellLabelToAddress(parts[0] as string);
    const end = cellLabelToAddress((parts[1] ?? parts[0]) as string);
    return {
      start: {
        row: Math.min(start.row, end.row),
        col: Math.min(start.col, end.col)
      },
      end: {
        row: Math.max(start.row, end.row),
        col: Math.max(start.col, end.col)
      }
    };
  } catch {
    return undefined;
  }
};

const getCellAddressRangeLabel = (start: CellAddress, end: CellAddress): string => {
  const normalizedStart = {
    row: Math.min(start.row, end.row),
    col: Math.min(start.col, end.col)
  };
  const normalizedEnd = {
    row: Math.max(start.row, end.row),
    col: Math.max(start.col, end.col)
  };
  return `${cellAddressToLabel(normalizedStart)}:${cellAddressToLabel(normalizedEnd)}`;
};

const toAbsoluteCellReference = (sheetName: string, address: CellAddress): string => {
  const label = cellAddressToLabel(address);
  const absoluteLabel = label.replace(/([A-Z]+)(\d+)/, (_match, colPart, rowPart) => `$${colPart}$${rowPart}`);
  return `${escapeSheetNameForFormula(sheetName)}!${absoluteLabel}`;
};

const toAbsoluteRangeReference = (sheetName: string, start: CellAddress, end: CellAddress): string => {
  const startLabel = toAbsoluteCellReference(sheetName, start);
  const endLabel = toAbsoluteCellReference(sheetName, end).replace(/^.+!/, "");
  return `${startLabel}:${endLabel}`;
};

const parseFormulaRangeReference = (formula: string): ParsedFormulaRange | undefined => {
  const normalizedFormula = formula.trim().replace(/^=/, "");
  const match = /^(?:'((?:[^']|'')+)'|([^'!]+))!(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?$/i.exec(normalizedFormula);
  if (!match) {
    return undefined;
  }

  const rawSheetName = (match[1] ?? match[2] ?? "").replace(/^(\[[^\]]+\])/, "");
  const sheetName = rawSheetName.replaceAll("''", "'").trim();
  if (!sheetName) {
    return undefined;
  }

  try {
    const start = cellLabelToAddress(normalizeCellLabel(match[3] as string));
    const end = cellLabelToAddress(normalizeCellLabel((match[4] ?? match[3]) as string));
    return {
      sheetName,
      start: {
        row: Math.min(start.row, end.row),
        col: Math.min(start.col, end.col)
      },
      end: {
        row: Math.max(start.row, end.row),
        col: Math.max(start.col, end.col)
      }
    };
  } catch {
    return undefined;
  }
};

const toCellPrimitive = (cell?: CellModel): CellPrimitive => {
  if (!cell) {
    return null;
  }
  return cell.formula ? cell.computedValue ?? null : cell.value;
};

const toNumericValue = (value: CellPrimitive): number | undefined => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }
  if (typeof value === "string") {
    const normalized = value.trim().replaceAll(",", ".");
    if (!normalized) {
      return undefined;
    }
    const numeric = Number(normalized);
    return Number.isFinite(numeric) ? numeric : undefined;
  }
  return undefined;
};

const toPointValue = (value: CellPrimitive): string | number => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }
  return sanitizeSpreadsheetText(String(value ?? ""), 120);
};

export interface XlsxTableExportOptions<Object extends object> {
  sheet?: string;
  columns: WriterColumn<Object>[];
}

export interface XlsxTableImportOptions<Object extends object, ColumnTitle extends string = string> {
  sheet?: number | string;
  schema: Schema<Object, ColumnTitle>;
}

export interface XlsxChartInteropOptions {
  onChartImported?: (payload: { sheetId: string; chart: WorksheetChartObject }) => void;
  onChartExported?: (payload: { sheetId: string; chartId: string }) => void;
  onChartUnsupportedFeature?: (payload: { sheetId: string; chartId: string; feature: string }) => void;
  onChartError?: (payload: {
    sheetId?: string;
    chartId?: string;
    errorCode: string;
    message: string;
  }) => void;
}

export interface XlsxAdapter {
  exportWorkbookToXlsx(workbook: WorkbookModel, options?: XlsxChartInteropOptions): Promise<Uint8Array>;
  importWorkbookFromXlsx(input: Uint8Array | ArrayBuffer, options?: XlsxChartInteropOptions): Promise<WorkbookModel>;
  exportTableToXlsx<Object extends object>(rows: Object[], options: XlsxTableExportOptions<Object>): Promise<Uint8Array>;
  importTableFromXlsx<Object extends object, ColumnTitle extends string = string>(
    input: Uint8Array | ArrayBuffer,
    options: XlsxTableImportOptions<Object, ColumnTitle>
  ): Promise<ParseSheetDataResult<Object, ColumnTitle>>;
}

export const createXlsxAdapter = (currentModulePath: string): XlsxAdapter => {
  const require = createRequire(currentModulePath);
  const readExcelFileRoot = dirname(require.resolve("read-excel-file/package.json"));

  const xmlModule = require(join(readExcelFileRoot, "commonjs/xml/xml.js")).default as {
    createDocument: (content: string) => Document;
  };
  const parseFilePaths = require(join(readExcelFileRoot, "commonjs/xlsx/parseFilePaths.js")).default as (
    content: string,
    xml: typeof xmlModule
  ) => {
    sheets: Record<string, string>;
    sharedStrings?: string;
    styles?: string;
  };
  const parseSpreadsheetInfo = require(join(readExcelFileRoot, "commonjs/xlsx/parseSpreadsheetInfo.js")).default as (
    content: string,
    xml: typeof xmlModule
  ) => {
    epoch1904: boolean;
    sheets: Array<{ name: string; relationId: string }>;
  };
  const parseSharedStrings = require(join(readExcelFileRoot, "commonjs/xlsx/parseSharedStrings.js")).default as (
    content: string | undefined,
    xml: typeof xmlModule
  ) => string[];
  const parseValueStyles = require(join(readExcelFileRoot, "commonjs/xlsx/parseStyles.js")).default as (
    content: string | undefined,
    xml: typeof xmlModule
  ) => unknown[];
  const parseCellValue = require(join(readExcelFileRoot, "commonjs/xlsx/parseCellValue.js")).default as (
    value: string | undefined,
    type: string | null,
    context: {
      getInlineStringValue: () => string | undefined;
      getInlineStringXml: () => string;
      getStyleId: () => string | null;
      styles: unknown[];
      sharedStrings: string[];
      epoch1904: boolean;
      options: Record<string, unknown>;
    }
  ) => CellModel["value"];
  const parseCellCoordinates = require(join(readExcelFileRoot, "commonjs/xlsx/parseCellCoordinates.js")).default as (
    coordinate: string
  ) => [number, number];
  const unpackXlsxFile = require(join(readExcelFileRoot, "commonjs/export/unpackXlsxFileNode.js")).default as (
    input: Buffer
  ) => Promise<Record<string, string>>;

  const parseStyleTable = (content?: string): CellStyle[] => {
    if (!content) {
      return [];
    }

    const document = xmlModule.createDocument(content);
    const styleSheet = document.documentElement;
    const fonts = getFirstElementChild(styleSheet, "fonts");
    const fills = getFirstElementChild(styleSheet, "fills");
    const borders = getFirstElementChild(styleSheet, "borders");
    const numFmts = getFirstElementChild(styleSheet, "numFmts");
    const cellXfs = getFirstElementChild(styleSheet, "cellXfs");

    const fontStyles = getElementChildren(fonts ?? styleSheet, "font").map((font): Partial<CellStyle> => ({
      fontFamily: getFirstElementChild(font, "name")?.getAttribute("val") ?? undefined,
      fontSize: Number(getFirstElementChild(font, "sz")?.getAttribute("val") ?? Number.NaN),
      fontWeight: getFirstElementChild(font, "b") ? "bold" : undefined,
      fontStyle: getFirstElementChild(font, "i") ? "italic" : undefined,
      underline: Boolean(getFirstElementChild(font, "u")),
      textColor: getColorFromElement(getFirstElementChild(font, "color"))
    }));

    const fillStyles = getElementChildren(fills ?? styleSheet, "fill").map((fill) => {
      const patternFill = getFirstElementChild(fill, "patternFill");
      return {
        backgroundColor:
          getColorFromElement(getFirstElementChild(patternFill ?? fill, "fgColor")) ??
          getColorFromElement(getFirstElementChild(patternFill ?? fill, "bgColor"))
      };
    });

    const borderStyles = getElementChildren(borders ?? styleSheet, "border").map((border) => ({
      border: {
        left: parseBorderSide(getFirstElementChild(border, "left")),
        right: parseBorderSide(getFirstElementChild(border, "right")),
        top: parseBorderSide(getFirstElementChild(border, "top")),
        bottom: parseBorderSide(getFirstElementChild(border, "bottom"))
      }
    }));

    const numberFormats = new Map<string, string>();
    for (const formatNode of getElementChildren(numFmts ?? styleSheet, "numFmt")) {
      const formatId = formatNode.getAttribute("numFmtId");
      const formatCode = formatNode.getAttribute("formatCode");
      if (formatId && formatCode) {
        numberFormats.set(formatId, formatCode);
      }
    }

    return getElementChildren(cellXfs ?? styleSheet, "xf").map((xf) => {
      const alignment = getFirstElementChild(xf, "alignment");
      const fontStyle = fontStyles[Number(xf.getAttribute("fontId") ?? -1)] ?? {};
      const fillStyle = fillStyles[Number(xf.getAttribute("fillId") ?? -1)] ?? {};
      const borderStyle = borderStyles[Number(xf.getAttribute("borderId") ?? -1)] ?? {};
      const format = numberFormats.get(xf.getAttribute("numFmtId") ?? "");
      const merged: CellStyle = {
        ...fontStyle,
        ...fillStyle,
        ...borderStyle,
        align: (alignment?.getAttribute("horizontal") as CellStyle["align"] | null) ?? undefined,
        alignVertical: (alignment?.getAttribute("vertical") as CellStyle["alignVertical"] | null) ?? undefined,
        wrap: alignment?.getAttribute("wrapText") === "1",
        indent: alignment?.hasAttribute("indent") ? Number(alignment.getAttribute("indent")) : undefined,
        format
      };

      if (Number.isNaN(merged.fontSize)) {
        delete merged.fontSize;
      }

      if (!merged.border?.left && !merged.border?.right && !merged.border?.top && !merged.border?.bottom) {
        delete merged.border;
      }

      return merged;
    });
  };

  const extractInlineString = (cell: Element): string | undefined => {
    const inlineString = getFirstElementChild(cell, "is");
    if (!inlineString) {
      return undefined;
    }

    const directText = getFirstElementChild(inlineString, "t")?.textContent ?? undefined;
    if (directText) {
      return directText;
    }

    const richRuns = getElementChildren(inlineString, "r")
      .map((run) => getFirstElementChild(run, "t")?.textContent ?? "")
      .join("");
    return richRuns || undefined;
  };

  const parseMergeReference = (reference: string): SheetMerge => {
    const [startRef, endRef = startRef] = reference.split(":");
    const [startRow, startCol] = parseCellCoordinates(startRef);
    const [endRow, endCol] = parseCellCoordinates(endRef);
    return {
      start: { row: startRow - 1, col: startCol - 1 },
      end: { row: endRow - 1, col: endCol - 1 }
    };
  };

  const parseSheetDimensionReference = (document: Document): CellBounds | undefined => {
    const reference = getFirstElementChild(document.documentElement, "dimension")?.getAttribute("ref")?.trim();
    if (!reference) {
      return undefined;
    }
    const [startRef, endRef = startRef] = reference.split(":");
    try {
      const [startRow, startCol] = parseCellCoordinates(startRef);
      const [endRow, endCol] = parseCellCoordinates(endRef);
      return {
        start: {
          row: Math.min(startRow, endRow) - 1,
          col: Math.min(startCol, endCol) - 1
        },
        end: {
          row: Math.max(startRow, endRow) - 1,
          col: Math.max(startCol, endCol) - 1
        }
      };
    } catch {
      return undefined;
    }
  };

  const parseColumnSchema = (document: Document, styleTable: CellStyle[]): Record<number, ColumnSchema> => {
    const schema: Record<number, ColumnSchema> = {};
    const cols = getFirstElementChild(document.documentElement, "cols");
    for (const col of getElementChildren(cols ?? document.documentElement, "col")) {
      const min = Number(col.getAttribute("min") ?? 1);
      const max = Number(col.getAttribute("max") ?? min);
      const width = Number(col.getAttribute("width") ?? Number.NaN);
      const hidden = col.getAttribute("hidden") === "1";
      const style = styleTable[Number(col.getAttribute("style") ?? -1)];
      for (let index = min - 1; index <= max - 1; index += 1) {
        schema[index] = {
          width: Number.isNaN(width) ? undefined : Math.round(width * EXCEL_COLUMN_WIDTH_UNIT),
          hidden: hidden || undefined,
          style: style && Object.keys(style).length ? { ...style } : undefined
        };
      }
    }
    return schema;
  };

  const parseRowSchema = (document: Document, styleTable: CellStyle[]): Record<number, RowSchema> => {
    const schema: Record<number, RowSchema> = {};
    const sheetData = getFirstElementChild(document.documentElement, "sheetData");
    for (const row of getElementChildren(sheetData ?? document.documentElement, "row")) {
      const rowIndex = Number(row.getAttribute("r") ?? 1) - 1;
      const height = Number(row.getAttribute("ht") ?? Number.NaN);
      const hidden = row.getAttribute("hidden") === "1";
      const style = styleTable[Number(row.getAttribute("s") ?? -1)];
      if (!Number.isNaN(height) || hidden || (style && Object.keys(style).length)) {
        schema[rowIndex] = {
          height: Number.isNaN(height) ? undefined : height,
          hidden: hidden || undefined,
          style: style && Object.keys(style).length ? { ...style } : undefined
        };
      }
    }
    return schema;
  };

  const parseSheetCells = (
    document: Document,
    sharedStrings: string[],
    valueStyles: unknown[],
    styleTable: CellStyle[],
    epoch1904: boolean
  ): Record<string, CellModel> => {
    const cells: Record<string, CellModel> = {};
    const sheetData = getFirstElementChild(document.documentElement, "sheetData");
    if (!sheetData) {
      return cells;
    }

    for (const row of getElementChildren(sheetData, "row")) {
      for (const cell of getElementChildren(row, "c")) {
        const coordinate = cell.getAttribute("r");
        if (!coordinate) {
          continue;
        }

        const [rowIndex, colIndex] = parseCellCoordinates(coordinate);
        const styleId = cell.getAttribute("s");
        const valueElement = getFirstElementChild(cell, "v");
        const formulaElement = getFirstElementChild(cell, "f");
        const formulaText = formulaElement?.textContent?.trim();

        let parsedValue: CellModel["value"] = null;
        try {
          parsedValue = parseCellValue(valueElement?.textContent ?? undefined, cell.getAttribute("t"), {
            getInlineStringValue: () => extractInlineString(cell),
            getInlineStringXml: () => cell.toString(),
            getStyleId: () => styleId,
            styles: valueStyles,
            sharedStrings,
            epoch1904,
            options: {}
          });
        } catch {
          parsedValue = (valueElement?.textContent ?? extractInlineString(cell) ?? null) as CellModel["value"];
        }

        const style = stripUndefinedStyle(styleTable[Number(styleId ?? -1)]);
        if (parsedValue == null && !formulaText && !style) {
          continue;
        }

        cells[getCellKey(rowIndex - 1, colIndex - 1)] = formulaText
          ? {
              value: normalizeImportedFormula(formulaText),
              formula: normalizeImportedFormula(formulaText),
              computedValue: parsedValue,
              style
            }
          : {
              value: parsedValue,
              computedValue: parsedValue,
              style
            };
      }
    }

    return cells;
  };

  const parseRelationships = (content?: string): ParsedRelationship[] => {
    if (!content) {
      return [];
    }
    const document = xmlModule.createDocument(content);
    const root = document.documentElement;
    return getElementChildren(root, "Relationship")
      .map((node) => ({
        id: node.getAttribute("Id") ?? "",
        type: node.getAttribute("Type") ?? "",
        target: node.getAttribute("Target") ?? ""
      }))
      .filter((relationship) => Boolean(relationship.id && relationship.target));
  };

  const findDescendantsByTagName = (element: Element, tagName: string): Element[] => {
    const queue: Element[] = [element];
    const result: Element[] = [];
    while (queue.length) {
      const current = queue.shift() as Element;
      const children = getElementChildren(current);
      for (const child of children) {
        if (child.tagName.replace(/^.+:/, "") === tagName) {
          result.push(child);
        }
        queue.push(child);
      }
    }
    return result;
  };

  const findFirstDescendantByTagName = (element: Element, tagName: string): Element | undefined =>
    findDescendantsByTagName(element, tagName)[0];

  const parseChartTypeTag = (plotArea: Element): string | undefined => {
    const children = getElementChildren(plotArea);
    for (const child of children) {
      const tag = child.tagName.replace(/^.+:/, "");
      if (tag.endsWith("Chart")) {
        return tag;
      }
    }
    return undefined;
  };

  const mapChartTypeTag = (chartTypeTag: string | undefined, chartNode?: Element): WorksheetChartType => {
    if (!chartTypeTag) {
      return "unknown";
    }
    switch (chartTypeTag) {
      case "barChart": {
        const barDirection = chartNode ? findFirstDescendantByTagName(chartNode, "barDir")?.getAttribute("val") : undefined;
        return barDirection === "bar" ? "bar" : "column";
      }
      case "lineChart":
        return "line";
      case "areaChart":
        return "area";
      case "pieChart":
        return "pie";
      case "doughnutChart":
        return "donut";
      case "scatterChart":
      case "bubbleChart":
        return "scatter";
      case "stockChart":
        return "candlestick";
      case "surfaceChart":
        return "surface";
      case "sunburstChart":
        return "sunburst";
      case "treemapChart":
        return "treemap";
      default:
        return "unknown";
    }
  };

  const parseTitleElement = (titleElement?: Element): string | undefined => {
    if (!titleElement) {
      return undefined;
    }
    const directValue = getFirstElementChild(getFirstElementChild(titleElement, "tx") ?? titleElement, "v")?.textContent?.trim();
    if (directValue) {
      return sanitizeSpreadsheetText(directValue, 180);
    }
    const textRuns = findDescendantsByTagName(titleElement, "t")
      .map((node) => node.textContent ?? "")
      .join("")
      .trim();
    return textRuns ? sanitizeSpreadsheetText(textRuns, 180) : undefined;
  };

  const parseChartTitle = (chartElement: Element): string | undefined => parseTitleElement(getFirstElementChild(chartElement, "title"));

  const parseChartLegendVisibility = (chartElement: Element): boolean => {
    const legend = getFirstElementChild(chartElement, "legend");
    if (!legend) {
      return false;
    }
    return getFirstElementChild(legend, "delete")?.getAttribute("val") !== "1";
  };

  const parseAxisType = (axisNode: Element | undefined): "linear" | "log" | "date" | "category" => {
    if (!axisNode) {
      return "linear";
    }
    const axisTag = axisNode.tagName.replace(/^.+:/, "");
    if (axisTag === "dateAx") {
      return "date";
    }
    if (axisTag === "catAx") {
      return "category";
    }
    if (axisTag === "valAx") {
      const scaling = getFirstElementChild(axisNode, "scaling");
      return getFirstElementChild(scaling ?? axisNode, "logBase") ? "log" : "linear";
    }
    return "linear";
  };

  const resolveAxisNodes = (
    plotArea: Element | undefined,
    chartNode: Element | undefined
  ): { xAxisNode?: Element; yAxisNode?: Element } => {
    if (!plotArea) {
      return {};
    }
    const axisById = new Map<string, Element>();
    for (const axisTag of ["catAx", "dateAx", "valAx"] as const) {
      for (const axisNode of getElementChildren(plotArea, axisTag)) {
        const axisId = getFirstElementChild(axisNode, "axId")?.getAttribute("val");
        if (axisId) {
          axisById.set(axisId, axisNode);
        }
      }
    }

    const axisIds = chartNode
      ? findDescendantsByTagName(chartNode, "axId")
          .map((node) => node.getAttribute("val") ?? "")
          .filter(Boolean)
      : [];
    let xAxisNode = axisIds[0] ? axisById.get(axisIds[0]) : undefined;
    let yAxisNode = axisIds[1] ? axisById.get(axisIds[1]) : undefined;

    if (!xAxisNode) {
      xAxisNode =
        getElementChildren(plotArea, "dateAx")[0] ??
        getElementChildren(plotArea, "catAx")[0] ??
        getElementChildren(plotArea, "valAx")[0];
    }
    if (!yAxisNode) {
      const valueAxes = getElementChildren(plotArea, "valAx");
      yAxisNode = valueAxes.find((axisNode) => axisNode !== xAxisNode) ?? valueAxes[0];
    }
    return {
      xAxisNode,
      yAxisNode
    };
  };

  const parseAxisMetadata = (axisNode: Element | undefined): {
    title?: string;
    visible: boolean;
    type: "linear" | "log" | "date" | "category";
  } => {
    if (!axisNode) {
      return {
        title: undefined,
        visible: true,
        type: "linear"
      };
    }
    const title = parseTitleElement(getFirstElementChild(axisNode, "title"));
    const visible = getFirstElementChild(axisNode, "delete")?.getAttribute("val") !== "1";
    return {
      title,
      visible,
      type: parseAxisType(axisNode)
    };
  };

  const parseChartDefinition = (chartXml: string): ParsedChartDefinition => {
    const document = xmlModule.createDocument(chartXml);
    const chartRoot = document.documentElement;
    const chart = getFirstElementChild(chartRoot, "chart");
    const plotArea = chart ? getFirstElementChild(chart, "plotArea") : undefined;
    const chartTypeTag = plotArea ? parseChartTypeTag(plotArea) : undefined;
    const chartNode = plotArea && chartTypeTag ? getFirstElementChild(plotArea, chartTypeTag) : undefined;
    const formulas = findDescendantsByTagName(chartRoot, "f")
      .map((formulaNode) => (formulaNode.textContent ?? "").trim())
      .filter(Boolean)
      .slice(0, MAX_CHART_FORMULA_REFERENCES);
    const seriesNames = chartNode
      ? getElementChildren(chartNode, "ser")
          .map((seriesNode) => {
            const txNode = getFirstElementChild(seriesNode, "tx");
            const literal = getFirstElementChild(txNode ?? seriesNode, "v")?.textContent?.trim();
            if (literal) {
              return sanitizeSpreadsheetText(literal, 80);
            }
            const textRuns = findDescendantsByTagName(txNode ?? seriesNode, "t")
              .map((node) => node.textContent ?? "")
              .join("")
              .trim();
            return textRuns ? sanitizeSpreadsheetText(textRuns, 80) : "";
          })
          .filter(Boolean)
      : [];
    const { xAxisNode, yAxisNode } = resolveAxisNodes(plotArea, chartNode);
    const xAxisMeta = parseAxisMetadata(xAxisNode);
    const yAxisMeta = parseAxisMetadata(yAxisNode);
    return {
      chartTypeTag: chartTypeTag ?? "unknownChart",
      mappedType: mapChartTypeTag(chartTypeTag, chartNode),
      title: chart ? parseChartTitle(chart) : undefined,
      legendVisible: chart ? parseChartLegendVisibility(chart) : false,
      formulas,
      xAxisTitle: xAxisMeta.title,
      yAxisTitle: yAxisMeta.title,
      xAxisVisible: xAxisMeta.visible,
      yAxisVisible: yAxisMeta.visible,
      xAxisType: xAxisMeta.type,
      yAxisType: yAxisMeta.type,
      seriesNames
    };
  };

  const parseDrawingAnchor = (anchorElement: Element): ParsedChartAnchor | undefined => {
    const fromNode = getFirstElementChild(anchorElement, "from");
    if (!fromNode) {
      return undefined;
    }
    const chartNode = findFirstDescendantByTagName(anchorElement, "chart");
    const chartRelationId = chartNode?.getAttribute("r:id") ?? chartNode?.getAttribute("id");
    if (!chartRelationId) {
      return undefined;
    }

    const parseMarker = (markerNode: Element | undefined) => {
      if (!markerNode) {
        return undefined;
      }
      const row = Number(getFirstElementChild(markerNode, "row")?.textContent ?? Number.NaN);
      const col = Number(getFirstElementChild(markerNode, "col")?.textContent ?? Number.NaN);
      const rowOff = Number(getFirstElementChild(markerNode, "rowOff")?.textContent ?? 0);
      const colOff = Number(getFirstElementChild(markerNode, "colOff")?.textContent ?? 0);
      if (!isFiniteInteger(row) || !isFiniteInteger(col)) {
        return undefined;
      }
      return {
        row,
        col,
        rowOffsetEmu: Number.isFinite(rowOff) ? rowOff : 0,
        colOffsetEmu: Number.isFinite(colOff) ? colOff : 0
      };
    };

    const from = parseMarker(fromNode);
    if (!from) {
      return undefined;
    }
    const to = parseMarker(getFirstElementChild(anchorElement, "to"));
    const extNode = getFirstElementChild(anchorElement, "ext");
    const widthEmu = Number(extNode?.getAttribute("cx") ?? Number.NaN);
    const heightEmu = Number(extNode?.getAttribute("cy") ?? Number.NaN);

    return {
      from,
      to,
      ext:
        Number.isFinite(widthEmu) && Number.isFinite(heightEmu)
          ? {
              widthEmu: widthEmu as number,
              heightEmu: heightEmu as number
            }
          : undefined,
      chartRelationId
    };
  };

  const getSheetColumnWidth = (sheet: WorkbookSheet, col: number): number =>
    sheet.columns[col]?.width ?? workbookDefaults.columnWidth;

  const getSheetRowHeight = (sheet: WorkbookSheet, row: number): number => sheet.rows[row]?.height ?? workbookDefaults.rowHeight;

  const getDistanceToColumn = (sheet: WorkbookSheet, col: number): number => {
    let offset = 0;
    for (let index = 0; index < col; index += 1) {
      offset += getSheetColumnWidth(sheet, index);
    }
    return offset;
  };

  const getDistanceToRow = (sheet: WorkbookSheet, row: number): number => {
    let offset = 0;
    for (let index = 0; index < row; index += 1) {
      offset += getSheetRowHeight(sheet, index);
    }
    return offset;
  };

  const estimateToAddressFromSize = (
    sheet: WorkbookSheet,
    from: CellAddress,
    widthPx: number,
    heightPx: number
  ): CellAddress => {
    let remainingWidth = Math.max(CHART_MIN_SIZE, widthPx);
    let remainingHeight = Math.max(CHART_MIN_SIZE, heightPx);
    let targetCol = from.col;
    let targetRow = from.row;
    const maxCol = Math.max(sheet.columnCount - 1, from.col + 512);
    const maxRow = Math.max(sheet.rowCount - 1, from.row + 1024);

    while (remainingWidth > getSheetColumnWidth(sheet, targetCol) && targetCol < maxCol) {
      remainingWidth -= getSheetColumnWidth(sheet, targetCol);
      targetCol += 1;
    }
    while (remainingHeight > getSheetRowHeight(sheet, targetRow) && targetRow < maxRow) {
      remainingHeight -= getSheetRowHeight(sheet, targetRow);
      targetRow += 1;
    }
    return {
      row: Math.max(0, targetRow),
      col: Math.max(0, targetCol)
    };
  };

  const resolvePositionFromAnchor = (sheet: WorkbookSheet, anchor: ParsedChartAnchor, zIndex: number) => {
    const fromAddress = {
      row: Math.max(0, anchor.from.row),
      col: Math.max(0, anchor.from.col)
    };
    const fromX = getDistanceToColumn(sheet, fromAddress.col) + fromEmu(anchor.from.colOffsetEmu);
    const fromY = getDistanceToRow(sheet, fromAddress.row) + fromEmu(anchor.from.rowOffsetEmu);
    const toAddress = anchor.to
      ? {
          row: Math.max(0, anchor.to.row),
          col: Math.max(0, anchor.to.col)
        }
      : estimateToAddressFromSize(sheet, fromAddress, fromEmu(anchor.ext?.widthEmu ?? DEFAULT_CHART_WIDTH), fromEmu(anchor.ext?.heightEmu ?? DEFAULT_CHART_HEIGHT));

    const toX = anchor.to
      ? getDistanceToColumn(sheet, toAddress.col) + fromEmu(anchor.to.colOffsetEmu)
      : fromX + fromEmu(anchor.ext?.widthEmu ?? DEFAULT_CHART_WIDTH);
    const toY = anchor.to
      ? getDistanceToRow(sheet, toAddress.row) + fromEmu(anchor.to.rowOffsetEmu)
      : fromY + fromEmu(anchor.ext?.heightEmu ?? DEFAULT_CHART_HEIGHT);

    return {
      fromCell: cellAddressToLabel(fromAddress),
      toCell: cellAddressToLabel(toAddress),
      offsetX: fromEmu(anchor.from.colOffsetEmu),
      offsetY: fromEmu(anchor.from.rowOffsetEmu),
      width: Math.max(CHART_MIN_SIZE, Math.round(Math.abs(toX - fromX))),
      height: Math.max(CHART_MIN_SIZE, Math.round(Math.abs(toY - fromY))),
      zIndex
    };
  };

  const createPlaceholderFigure = (chartType: WorksheetChartType, title: string | undefined, originalType: string) => ({
    data: [],
    layout: {
      title: sanitizeSpreadsheetText(title ?? `Chart (${chartType})`, 180),
      legend: {
        visible: false
      }
    },
    metadata: {
      placeholder: true,
      message: "Gráfico não suportado nesta versão",
      originalType
    }
  });

  const resolveSeriesFromRangeBinding = (
    sourceSheet: WorkbookSheet,
    sourceSheetName: string,
    binding: Pick<ChartRangeBinding, "rangeAddress" | "orientation" | "firstRowAsHeader" | "firstColumnAsLabel">,
    nameOverrides?: string[]
  ): ChartExportSeries[] => {
    const bounds = parseRangeAddress(binding.rangeAddress);
    if (!bounds) {
      return [];
    }

    const rows = bounds.end.row - bounds.start.row + 1;
    const cols = bounds.end.col - bounds.start.col + 1;
    if (rows <= 0 || cols <= 0) {
      return [];
    }
    if (rows * cols > MAX_CHART_RANGE_CELLS) {
      return [];
    }

    const orientation = binding.orientation === "columns" ? "columns" : "rows";
    const firstRowAsHeader = binding.firstRowAsHeader === true;
    const firstColumnAsLabel = binding.firstColumnAsLabel === true;
    const series: ChartExportSeries[] = [];

    if (orientation === "rows") {
      const dataStartRow = bounds.start.row + (firstRowAsHeader ? 1 : 0);
      const dataEndRow = bounds.end.row;
      const labelColumn = firstColumnAsLabel ? bounds.start.col : undefined;
      const seriesStartCol = bounds.start.col + (firstColumnAsLabel ? 1 : 0);
      for (let col = seriesStartCol; col <= bounds.end.col; col += 1) {
        if (series.length >= MAX_CHART_SERIES) {
          break;
        }
        if (dataStartRow > dataEndRow) {
          continue;
        }
        const headerCell = toCellPrimitive(sourceSheet.cells[getCellKey(bounds.start.row, col)]);
        const fallbackName = firstRowAsHeader
          ? sanitizeSpreadsheetText(String(headerCell ?? `Series ${col - seriesStartCol + 1}`), 80)
          : `Series ${col - seriesStartCol + 1}`;
        const customName = sanitizeSpreadsheetText(nameOverrides?.[series.length] ?? "", 80);
        const name = customName || fallbackName;
        const nameRef = customName
          ? undefined
          : firstRowAsHeader
          ? toAbsoluteCellReference(sourceSheetName, {
              row: bounds.start.row,
              col
            })
          : undefined;
        const categoryRef =
          labelColumn !== undefined
            ? toAbsoluteRangeReference(
                sourceSheetName,
                { row: dataStartRow, col: labelColumn },
                { row: dataEndRow, col: labelColumn }
              )
            : undefined;
        const valueRef = toAbsoluteRangeReference(
          sourceSheetName,
          { row: dataStartRow, col },
          { row: dataEndRow, col }
        );
        if (!isSafeFormulaReference(valueRef)) {
          continue;
        }
        series.push({
          name,
          nameRef,
          categoryRef,
          valueRef
        });
      }
      return series;
    }

    const dataStartCol = bounds.start.col + (firstColumnAsLabel ? 1 : 0);
    const dataEndCol = bounds.end.col;
    const labelRow = firstRowAsHeader ? bounds.start.row : undefined;
    const seriesStartRow = bounds.start.row + (firstRowAsHeader ? 1 : 0);
    for (let row = seriesStartRow; row <= bounds.end.row; row += 1) {
      if (series.length >= MAX_CHART_SERIES) {
        break;
      }
      if (dataStartCol > dataEndCol) {
        continue;
      }
      const headerCell = toCellPrimitive(sourceSheet.cells[getCellKey(row, bounds.start.col)]);
      const fallbackName = firstColumnAsLabel
        ? sanitizeSpreadsheetText(String(headerCell ?? `Series ${row - seriesStartRow + 1}`), 80)
        : `Series ${row - seriesStartRow + 1}`;
      const customName = sanitizeSpreadsheetText(nameOverrides?.[series.length] ?? "", 80);
      const name = customName || fallbackName;
      const nameRef = customName
        ? undefined
        : firstColumnAsLabel
        ? toAbsoluteCellReference(sourceSheetName, {
            row,
            col: bounds.start.col
          })
        : undefined;
      const categoryRef =
        labelRow !== undefined
          ? toAbsoluteRangeReference(
              sourceSheetName,
              { row: labelRow, col: dataStartCol },
              { row: labelRow, col: dataEndCol }
            )
          : undefined;
      const valueRef = toAbsoluteRangeReference(
        sourceSheetName,
        { row, col: dataStartCol },
        { row, col: dataEndCol }
      );
      if (!isSafeFormulaReference(valueRef)) {
        continue;
      }
      series.push({
        name,
        nameRef,
        categoryRef,
        valueRef
      });
    }
    return series;
  };

  const buildChartFigureFromSourceRange = (
    sheetById: Map<string, WorkbookSheet>,
    chartType: WorksheetChartType,
    sourceRange: ChartRangeBinding | undefined,
    title: string | undefined,
    originalType: string
  ): WorksheetChartObject["figure"] => {
    if (!sourceRange) {
      return createPlaceholderFigure(chartType, title, originalType);
    }
    const sourceSheet = sheetById.get(sourceRange.sheetId);
    const bounds = parseRangeAddress(sourceRange.rangeAddress);
    if (!sourceSheet || !bounds) {
      return createPlaceholderFigure(chartType, title, originalType);
    }

    const matrix: CellPrimitive[][] = [];
    for (let row = bounds.start.row; row <= bounds.end.row; row += 1) {
      const nextRow: CellPrimitive[] = [];
      for (let col = bounds.start.col; col <= bounds.end.col; col += 1) {
        nextRow.push(toCellPrimitive(sourceSheet.cells[getCellKey(row, col)]));
      }
      matrix.push(nextRow);
    }
    if (!matrix.length || !matrix[0]?.length) {
      return createPlaceholderFigure(chartType, title, originalType);
    }

    const orientationRows = sourceRange.orientation !== "columns";
    const oriented = orientationRows
      ? matrix
      : matrix[0].map((_value, columnIndex) => matrix.map((row) => row[columnIndex] ?? null));
    const firstRowAsHeader = sourceRange.firstRowAsHeader === true;
    const firstColumnAsLabel = sourceRange.firstColumnAsLabel === true;
    const headers = firstRowAsHeader
      ? oriented[0]!.map((value, index) => String(value ?? `Series ${index + 1}`))
      : oriented[0]!.map((_value, index) => `Series ${index + 1}`);
    const dataRows = firstRowAsHeader ? oriented.slice(1) : oriented.slice();
    if (!dataRows.length) {
      return createPlaceholderFigure(chartType, title, originalType);
    }

    if (chartType === "pie" || chartType === "donut") {
      const labelIndex = firstColumnAsLabel ? 0 : -1;
      const valueIndex = firstColumnAsLabel ? 1 : 0;
      const labels: string[] = [];
      const values: number[] = [];
      for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex += 1) {
        const row = dataRows[rowIndex] ?? [];
        const numeric = toNumericValue((row[valueIndex] ?? null) as CellPrimitive);
        if (numeric === undefined) {
          continue;
        }
        labels.push(
          sanitizeSpreadsheetText(
            labelIndex >= 0 ? String(row[labelIndex] ?? `Item ${rowIndex + 1}`) : `Item ${rowIndex + 1}`,
            100
          )
        );
        values.push(numeric);
      }
      if (!values.length) {
        return createPlaceholderFigure(chartType, title, originalType);
      }
      return {
        data: [
          {
            type: chartType,
            labels,
            values,
            name: sanitizeSpreadsheetText(headers[valueIndex] ?? "Value", 80)
          }
        ],
        layout: {
          title: sanitizeSpreadsheetText(title ?? headers[valueIndex] ?? "Chart", 180),
          legend: {
            visible: true
          }
        },
        metadata: {
          source: "xlsx-import"
        }
      };
    }

    if (!SUPPORTED_IMPORT_TYPES.has(chartType)) {
      return createPlaceholderFigure(chartType, title, originalType);
    }

    const labelIndex = firstColumnAsLabel ? 0 : -1;
    const xValues = dataRows.map((row, rowIndex) =>
      labelIndex >= 0 ? toPointValue((row[labelIndex] ?? null) as CellPrimitive) : rowIndex + 1
    );
    const startSeriesColumn = firstColumnAsLabel ? 1 : 0;
    const traces: Array<Record<string, unknown>> = [];
    for (let col = startSeriesColumn; col < headers.length; col += 1) {
      const yValues = dataRows.map((row) => {
        const numeric = toNumericValue((row[col] ?? null) as CellPrimitive);
        return numeric ?? 0;
      });
      if (!yValues.some((value) => Number.isFinite(value))) {
        continue;
      }
      traces.push({
        type: chartType === "column" ? "bar" : chartType,
        name: sanitizeSpreadsheetText(headers[col] ?? `Series ${col + 1}`, 80),
        x: xValues,
        y: yValues,
        mode: chartType === "scatter" ? "markers" : chartType === "line" ? "lines" : "lines+markers"
      });
    }
    if (!traces.length) {
      return createPlaceholderFigure(chartType, title, originalType);
    }

    return {
      data: traces,
      layout: {
        title: sanitizeSpreadsheetText(title ?? "Chart", 180),
        legend: {
          visible: true
        }
      },
      metadata: {
        source: "xlsx-import"
      }
    };
  };

  const resolveChartRangeBinding = (
    formulas: string[],
    fallbackSheetName: string,
    fallbackSheetId: string,
    sheetNameToId: Map<string, string>
  ): ChartRangeBinding | undefined => {
    const parsedRanges = formulas
      .slice(0, MAX_CHART_FORMULA_REFERENCES)
      .map(parseFormulaRangeReference)
      .filter((value): value is ParsedFormulaRange => Boolean(value));
    if (!parsedRanges.length) {
      return undefined;
    }

    const scoreBySheet = new Map<string, number>();
    for (const parsedRange of parsedRanges) {
      const key = parsedRange.sheetName.toLowerCase();
      scoreBySheet.set(key, (scoreBySheet.get(key) ?? 0) + 1);
    }
    const fallbackKey = fallbackSheetName.toLowerCase();
    let selectedSheetKey = fallbackKey;
    let selectedScore = -1;
    for (const [sheetKey, score] of scoreBySheet.entries()) {
      if (sheetKey === fallbackKey) {
        selectedSheetKey = sheetKey;
        selectedScore = score;
        break;
      }
      if (score > selectedScore) {
        selectedSheetKey = sheetKey;
        selectedScore = score;
      }
    }

    const selectedRanges = parsedRanges.filter((range) => range.sheetName.toLowerCase() === selectedSheetKey);
    if (!selectedRanges.length) {
      return undefined;
    }

    const bounds = selectedRanges.reduce(
      (accumulator, range) => ({
        start: {
          row: Math.min(accumulator.start.row, range.start.row),
          col: Math.min(accumulator.start.col, range.start.col)
        },
        end: {
          row: Math.max(accumulator.end.row, range.end.row),
          col: Math.max(accumulator.end.col, range.end.col)
        }
      }),
      {
        start: { ...selectedRanges[0].start },
        end: { ...selectedRanges[0].end }
      }
    );

    const selectedSheetId = sheetNameToId.get(selectedSheetKey) ?? fallbackSheetId;
    return {
      chartId: "",
      sheetId: selectedSheetId,
      rangeAddress: getCellAddressRangeLabel(bounds.start, bounds.end),
      orientation: "rows",
      firstRowAsHeader: bounds.end.row > bounds.start.row,
      firstColumnAsLabel: bounds.end.col > bounds.start.col,
      autoRefresh: true
    };
  };

  const collectWorksheetChartsFromDrawing = (input: {
    sheet: WorkbookSheet;
    sheetPath: string;
    contents: Record<string, string>;
    sheetNameToId: Map<string, string>;
    sheetById: Map<string, WorkbookSheet>;
    startZIndex: number;
  }): WorksheetChartObject[] => {
    const { sheet, sheetPath, contents, sheetNameToId, sheetById, startZIndex } = input;
    const drawingResults: WorksheetChartObject[] = [];
    const sheetDocument = xmlModule.createDocument(contents[sheetPath] ?? "<worksheet/>");
    const drawingReferences = getElementChildren(sheetDocument.documentElement, "drawing").map((node) => node.getAttribute("r:id"));
    if (!drawingReferences.length) {
      return drawingResults;
    }

    const sheetRelsPath = normalizeArchivePath(
      `${posixPath.dirname(sheetPath)}/_rels/${basename(sheetPath)}.rels`
    );
    const sheetRelationships = parseRelationships(contents[sheetRelsPath]);
    const relationshipById = new Map(sheetRelationships.map((relationship) => [relationship.id, relationship]));
    const drawingPaths = drawingReferences
      .filter((relationshipId): relationshipId is string => Boolean(relationshipId))
      .map((relationshipId) => relationshipById.get(relationshipId))
      .filter((relationship): relationship is ParsedRelationship => Boolean(relationship && relationship.type === DRAWING_REL_TYPE))
      .map((relationship) => resolveArchivePath(sheetPath, relationship.target))
      .filter(Boolean);

    let zIndex = Math.max(1, startZIndex);
    for (const drawingPath of drawingPaths) {
      const drawingXml = contents[drawingPath];
      if (!drawingXml) {
        continue;
      }

      const drawingDocument = xmlModule.createDocument(drawingXml);
      const drawingRoot = drawingDocument.documentElement;
      const drawingRelsPath = normalizeArchivePath(
        `${posixPath.dirname(drawingPath)}/_rels/${basename(drawingPath)}.rels`
      );
      const drawingRelationships = parseRelationships(contents[drawingRelsPath]);
      const drawingRelationshipById = new Map(drawingRelationships.map((relationship) => [relationship.id, relationship]));
      const anchorElements = getElementChildren(drawingRoot).filter((element) => {
        const tag = element.tagName.replace(/^.+:/, "");
        return tag === "twoCellAnchor" || tag === "oneCellAnchor" || tag === "absoluteAnchor";
      });

      for (let anchorIndex = 0; anchorIndex < anchorElements.length; anchorIndex += 1) {
        if (drawingResults.length >= MAX_IMPORT_CHARTS_PER_SHEET) {
          break;
        }
        const anchorElement = anchorElements[anchorIndex] as Element;
        const anchor = parseDrawingAnchor(anchorElement);
        if (!anchor) {
          continue;
        }

        const chartRelationship = drawingRelationshipById.get(anchor.chartRelationId);
        if (!chartRelationship || chartRelationship.type !== CHART_REL_TYPE) {
          continue;
        }
        const chartPath = resolveArchivePath(drawingPath, chartRelationship.target);
        const chartXml = contents[chartPath];
        if (!chartXml) {
          continue;
        }

        const definition = parseChartDefinition(chartXml);
        const sourceRange = resolveChartRangeBinding(definition.formulas, sheet.name, sheet.id, sheetNameToId);
        const chartId = `chart-import-${sheet.id}-${anchorIndex + 1}`;
        const unsupportedFeatures: string[] = [];
        let normalizedSourceRange = sourceRange
          ? {
              ...sourceRange,
              chartId
            }
          : undefined;
        if (normalizedSourceRange) {
          const rangeBounds = parseRangeAddress(normalizedSourceRange.rangeAddress);
          const rangeCellCount =
            rangeBounds != null
              ? (rangeBounds.end.row - rangeBounds.start.row + 1) * (rangeBounds.end.col - rangeBounds.start.col + 1)
              : 0;
          if (!rangeBounds || rangeCellCount > MAX_CHART_RANGE_CELLS) {
            unsupportedFeatures.push("range:too-large");
            normalizedSourceRange = undefined;
          }
        }
        const figure = buildChartFigureFromSourceRange(
          sheetById,
          definition.mappedType,
          normalizedSourceRange,
          definition.title,
          definition.chartTypeTag
        );
        const figureLayout =
          figure.layout && typeof figure.layout === "object"
            ? ({ ...(figure.layout as Record<string, unknown>) } as Record<string, unknown>)
            : {};
        const xAxis = (figureLayout.xAxis && typeof figureLayout.xAxis === "object"
          ? { ...(figureLayout.xAxis as Record<string, unknown>) }
          : {}) as Record<string, unknown>;
        const yAxis = (figureLayout.yAxis && typeof figureLayout.yAxis === "object"
          ? { ...(figureLayout.yAxis as Record<string, unknown>) }
          : {}) as Record<string, unknown>;
        if (definition.xAxisTitle) {
          xAxis.title = sanitizeSpreadsheetText(definition.xAxisTitle, 120);
        }
        if (definition.yAxisTitle) {
          yAxis.title = sanitizeSpreadsheetText(definition.yAxisTitle, 120);
        }
        xAxis.type = definition.xAxisType;
        yAxis.type = definition.yAxisType;
        xAxis.visible = definition.xAxisVisible;
        yAxis.visible = definition.yAxisVisible;
        figureLayout.xAxis = xAxis;
        figureLayout.yAxis = yAxis;
        figure.layout = figureLayout;
        if (definition.seriesNames.length && Array.isArray(figure.data)) {
          figure.data = figure.data.map((trace, index) => {
            if (!trace || typeof trace !== "object") {
              return trace;
            }
            const overrideName = definition.seriesNames[index];
            if (!overrideName) {
              return trace;
            }
            return {
              ...(trace as Record<string, unknown>),
              name: sanitizeSpreadsheetText(overrideName, 80)
            };
          });
        }
        const position = resolvePositionFromAnchor(sheet, anchor, zIndex++);
        if (!SUPPORTED_IMPORT_TYPES.has(definition.mappedType)) {
          unsupportedFeatures.push(`type:${definition.chartTypeTag}`);
        }
        if (!normalizedSourceRange) {
          unsupportedFeatures.push("range:missing");
        }

        drawingResults.push({
          id: chartId,
          sheetId: sheet.id,
          type: definition.mappedType,
          title: definition.title,
          sourceRange: normalizedSourceRange,
          figure,
          position,
          state: {
            selected: false,
            visible: true,
            locked: false,
            lastRenderedAt: Date.now()
          },
          style: {
            professionalPreset: "spreadsheet"
          },
          excelInterop: {
            originalChartType: definition.chartTypeTag,
            originalChartId: chartPath,
            originalAnchor: anchor,
            unsupportedFeatures: unsupportedFeatures.length ? unsupportedFeatures : undefined,
            fallbackImage: unsupportedFeatures.length > 0,
            preservedRawMetadata: {
              formulas: definition.formulas,
              seriesNames: definition.seriesNames,
              xAxisType: definition.xAxisType,
              yAxisType: definition.yAxisType,
              xAxisVisible: definition.xAxisVisible,
              yAxisVisible: definition.yAxisVisible
            }
          }
        });
      }
    }

    return drawingResults;
  };

  const buildChartTitleXml = (title: string): string => {
    const safeTitle = sanitizeXmlText(sanitizeSpreadsheetText(title, 180));
    if (!safeTitle) {
      return "";
    }
    return (
      "<c:title>" +
      "<c:tx>" +
      "<c:rich>" +
      "<a:bodyPr/>" +
      "<a:lstStyle/>" +
      "<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>" +
      safeTitle +
      "</a:t></a:r></a:p>" +
      "</c:rich>" +
      "</c:tx>" +
      "</c:title>"
    );
  };

  const buildSeriesXml = (series: ChartExportSeries, index: number): string => {
    const safeNameRef = series.nameRef && isSafeFormulaReference(series.nameRef) ? sanitizeXmlText(series.nameRef) : undefined;
    const safeCategoryRef =
      series.categoryRef && isSafeFormulaReference(series.categoryRef) ? sanitizeXmlText(series.categoryRef) : undefined;
    const safeValueRef = isSafeFormulaReference(series.valueRef) ? sanitizeXmlText(series.valueRef) : undefined;
    if (!safeValueRef) {
      return "";
    }
    const tx = safeNameRef
      ? `<c:tx><c:strRef><c:f>${safeNameRef}</c:f></c:strRef></c:tx>`
      : series.name
        ? `<c:tx><c:v>${sanitizeXmlText(sanitizeSpreadsheetText(series.name, 80))}</c:v></c:tx>`
        : "";
    const cat = safeCategoryRef ? `<c:cat><c:strRef><c:f>${safeCategoryRef}</c:f></c:strRef></c:cat>` : "";
    return (
      "<c:ser>" +
      `<c:idx val="${index}"/>` +
      `<c:order val="${index}"/>` +
      tx +
      cat +
      `<c:val><c:numRef><c:f>${safeValueRef}</c:f></c:numRef></c:val>` +
      "</c:ser>"
    );
  };

  const buildPieSeriesXml = (series: ChartExportSeries, index: number): string => {
    const safeNameRef = series.nameRef && isSafeFormulaReference(series.nameRef) ? sanitizeXmlText(series.nameRef) : undefined;
    const safeCategoryRef =
      series.categoryRef && isSafeFormulaReference(series.categoryRef) ? sanitizeXmlText(series.categoryRef) : undefined;
    const safeValueRef = isSafeFormulaReference(series.valueRef) ? sanitizeXmlText(series.valueRef) : undefined;
    if (!safeValueRef) {
      return "";
    }
    const tx = safeNameRef
      ? `<c:tx><c:strRef><c:f>${safeNameRef}</c:f></c:strRef></c:tx>`
      : series.name
        ? `<c:tx><c:v>${sanitizeXmlText(sanitizeSpreadsheetText(series.name, 80))}</c:v></c:tx>`
        : "";
    const cat = safeCategoryRef ? `<c:cat><c:strRef><c:f>${safeCategoryRef}</c:f></c:strRef></c:cat>` : "";
    return (
      "<c:ser>" +
      `<c:idx val="${index}"/>` +
      `<c:order val="${index}"/>` +
      tx +
      cat +
      `<c:val><c:numRef><c:f>${safeValueRef}</c:f></c:numRef></c:val>` +
      "</c:ser>"
    );
  };

  const buildScatterSeriesXml = (series: ChartExportSeries, index: number): string => {
    const safeNameRef = series.nameRef && isSafeFormulaReference(series.nameRef) ? sanitizeXmlText(series.nameRef) : undefined;
    const safeCategoryRef =
      series.categoryRef && isSafeFormulaReference(series.categoryRef) ? sanitizeXmlText(series.categoryRef) : undefined;
    const safeValueRef = isSafeFormulaReference(series.valueRef) ? sanitizeXmlText(series.valueRef) : undefined;
    if (!safeCategoryRef || !safeValueRef) {
      return "";
    }
    const tx = safeNameRef
      ? `<c:tx><c:strRef><c:f>${safeNameRef}</c:f></c:strRef></c:tx>`
      : series.name
        ? `<c:tx><c:v>${sanitizeXmlText(sanitizeSpreadsheetText(series.name, 80))}</c:v></c:tx>`
        : "";
    return (
      "<c:ser>" +
      `<c:idx val="${index}"/>` +
      `<c:order val="${index}"/>` +
      tx +
      `<c:xVal><c:numRef><c:f>${safeCategoryRef}</c:f></c:numRef></c:xVal>` +
      `<c:yVal><c:numRef><c:f>${safeValueRef}</c:f></c:numRef></c:yVal>` +
      "</c:ser>"
    );
  };

  const buildCartesianChartXml = (input: {
    profile: "line" | "area" | "bar-col" | "bar-row" | "scatter";
    title?: string;
    legendVisible: boolean;
    series: ChartExportSeries[];
    chartSequence: number;
    xAxisTitle?: string;
    yAxisTitle?: string;
    xAxisVisible: boolean;
    yAxisVisible: boolean;
    xAxisType: "linear" | "log" | "date" | "category";
    yAxisType: "linear" | "log" | "date" | "category";
  }): string => {
    const {
      profile,
      title,
      legendVisible,
      series,
      chartSequence,
      xAxisTitle,
      yAxisTitle,
      xAxisVisible,
      yAxisVisible,
      xAxisType,
      yAxisType
    } = input;
    const axisId1 = 120000 + chartSequence * 10;
    const axisId2 = axisId1 + 1;
    const seriesXml = series
      .map((entry, index) => buildSeriesXml(entry, index))
      .filter(Boolean)
      .join("");
    const legendXml = legendVisible
      ? "<c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend>"
      : "<c:legend><c:legendPos val=\"r\"/><c:layout/><c:delete val=\"1\"/></c:legend>";

    let chartTypeXml = "";
    if (profile === "line") {
      chartTypeXml =
        "<c:lineChart><c:grouping val=\"standard\"/>" +
        seriesXml +
        `<c:axId val="${axisId1}"/><c:axId val="${axisId2}"/>` +
        "</c:lineChart>";
    } else if (profile === "scatter") {
      const scatterSeriesXml = series
        .map((entry, index) => buildScatterSeriesXml(entry, index))
        .filter(Boolean)
        .join("");
      chartTypeXml =
        "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/>" +
        scatterSeriesXml +
        `<c:axId val="${axisId1}"/><c:axId val="${axisId2}"/>` +
        "</c:scatterChart>";
    } else if (profile === "area") {
      chartTypeXml =
        "<c:areaChart><c:grouping val=\"standard\"/>" +
        seriesXml +
        `<c:axId val="${axisId1}"/><c:axId val="${axisId2}"/>` +
        "</c:areaChart>";
    } else {
      chartTypeXml =
        `<c:barChart><c:barDir val="${profile === "bar-row" ? "bar" : "col"}"/><c:grouping val="clustered"/>` +
        seriesXml +
        `<c:axId val="${axisId1}"/><c:axId val="${axisId2}"/>` +
        "</c:barChart>";
    }

    const xAxisTag = profile === "scatter" ? "valAx" : xAxisType === "date" ? "dateAx" : "catAx";
    const xAxisScalingXml =
      xAxisTag === "valAx"
        ? `<c:scaling><c:orientation val="minMax"/>${xAxisType === "log" ? "<c:logBase val=\"10\"/>" : ""}</c:scaling>`
        : "<c:scaling><c:orientation val=\"minMax\"/></c:scaling>";
    const xAxisXml =
      `<c:${xAxisTag}><c:axId val="${axisId1}"/>` +
      xAxisScalingXml +
      `<c:delete val="${xAxisVisible ? 0 : 1}"/>` +
      (xAxisTitle ? buildChartTitleXml(xAxisTitle) : "") +
      `<c:axPos val="${profile === "bar-row" ? "l" : "b"}"/>` +
      `<c:crossAx val="${axisId2}"/><c:tickLblPos val="nextTo"/><c:crosses val="autoZero"/></c:${xAxisTag}>`;
    const yAxisXml =
      `<c:valAx><c:axId val="${axisId2}"/><c:scaling><c:orientation val="minMax"/>${
        yAxisType === "log" ? "<c:logBase val=\"10\"/>" : ""
      }</c:scaling><c:delete val="${yAxisVisible ? 0 : 1}"/>` +
      (yAxisTitle ? buildChartTitleXml(yAxisTitle) : "") +
      `<c:axPos val="${profile === "bar-row" ? "b" : "l"}"/>` +
      `<c:crossAx val="${axisId1}"/><c:tickLblPos val="nextTo"/><c:crosses val="autoZero"/></c:valAx>`;

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${DOC_REL_NS}">` +
      "<c:lang val=\"en-US\"/>" +
      "<c:chart>" +
      (title ? buildChartTitleXml(title) : "") +
      "<c:plotArea><c:layout/>" +
      chartTypeXml +
      xAxisXml +
      yAxisXml +
      "</c:plotArea>" +
      legendXml +
      "<c:plotVisOnly val=\"1\"/>" +
      "</c:chart>" +
      "</c:chartSpace>"
    );
  };

  const buildPieChartXml = (input: {
    profile: "pie" | "donut";
    title?: string;
    legendVisible: boolean;
    series: ChartExportSeries[];
  }): string => {
    const { profile, title, legendVisible, series } = input;
    const seriesXml = series
      .map((entry, index) => buildPieSeriesXml(entry, index))
      .filter(Boolean)
      .join("");
    const legendXml = legendVisible
      ? "<c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend>"
      : "<c:legend><c:legendPos val=\"r\"/><c:layout/><c:delete val=\"1\"/></c:legend>";
    const chartNode =
      profile === "donut"
        ? `<c:doughnutChart><c:varyColors val="1"/>${seriesXml}<c:holeSize val="60"/></c:doughnutChart>`
        : `<c:pieChart><c:varyColors val="1"/>${seriesXml}<c:firstSliceAng val="0"/></c:pieChart>`;

    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${DOC_REL_NS}">` +
      "<c:lang val=\"en-US\"/>" +
      "<c:chart>" +
      (title ? buildChartTitleXml(title) : "") +
      "<c:plotArea><c:layout/>" +
      chartNode +
      "</c:plotArea>" +
      legendXml +
      "<c:plotVisOnly val=\"1\"/>" +
      "</c:chart>" +
      "</c:chartSpace>"
    );
  };

  const resolveChartLegendVisible = (chart: WorksheetChartObject): boolean => {
    const layout = chart.figure.layout;
    if (!layout || typeof layout !== "object") {
      return true;
    }
    const legend = (layout as Record<string, unknown>).legend;
    if (!legend || typeof legend !== "object") {
      return true;
    }
    const visible = (legend as Record<string, unknown>).visible;
    return visible !== false;
  };

  const resolveChartAxisConfig = (
    chart: WorksheetChartObject,
    axis: "x" | "y"
  ): {
    title?: string;
    visible: boolean;
    type: "linear" | "log" | "date" | "category";
  } => {
    const layout = chart.figure.layout;
    if (!layout || typeof layout !== "object") {
      return {
        title: undefined,
        visible: true,
        type: axis === "x" ? "category" : "linear"
      };
    }
    const key = axis === "x" ? "xAxis" : "yAxis";
    const axisRecord =
      key in layout && typeof (layout as Record<string, unknown>)[key] === "object"
        ? ((layout as Record<string, unknown>)[key] as Record<string, unknown>)
        : undefined;
    const titleValue = axisRecord?.title;
    let title: string | undefined;
    if (typeof titleValue === "string") {
      title = sanitizeSpreadsheetText(titleValue, 120);
    } else if (titleValue && typeof titleValue === "object" && typeof (titleValue as Record<string, unknown>).text === "string") {
      title = sanitizeSpreadsheetText((titleValue as Record<string, unknown>).text as string, 120);
    }
    const typeValue = typeof axisRecord?.type === "string" ? axisRecord.type.toLowerCase() : "";
    const type = (["linear", "log", "date", "category"] as const).includes(typeValue as "linear" | "log" | "date" | "category")
      ? (typeValue as "linear" | "log" | "date" | "category")
      : axis === "x"
        ? "category"
        : "linear";
    const visible = axisRecord?.visible !== false;
    return {
      title,
      visible,
      type
    };
  };

  const buildChartExportEntries = (input: {
    workbook: WorkbookModel;
    sheet: WorkbookSheet;
    sheetNameById: Map<string, string>;
    globalStartIndex: number;
  }): { entries: ChartExportEntry[]; nextGlobalIndex: number } => {
    const { workbook, sheet, sheetNameById, globalStartIndex } = input;
    const entries: ChartExportEntry[] = [];
    let nextGlobalIndex = globalStartIndex;
    const charts = (sheet.charts ?? []).filter((chart) => chart.state.visible !== false).slice(0, MAX_EXPORT_CHARTS_PER_SHEET);
    const buildAnchor = (chart: WorksheetChartObject): ChartExportAnchor => {
      const fromAddress = (() => {
        try {
          return cellLabelToAddress(normalizeCellLabel(chart.position.fromCell));
        } catch {
          return { row: 0, col: 0 };
        }
      })();
      const safeFromAddress = {
        row: clampNumber(fromAddress.row, 0, Math.max(0, sheet.rowCount - 1)),
        col: clampNumber(fromAddress.col, 0, Math.max(0, sheet.columnCount - 1))
      };
      return {
        row: safeFromAddress.row,
        col: safeFromAddress.col,
        rowOffsetPx: Math.max(0, chart.position.offsetY),
        colOffsetPx: Math.max(0, chart.position.offsetX),
        widthPx: Math.max(CHART_MIN_SIZE, chart.position.width || DEFAULT_CHART_WIDTH),
        heightPx: Math.max(CHART_MIN_SIZE, chart.position.height || DEFAULT_CHART_HEIGHT)
      };
    };

    for (let chartIndex = 0; chartIndex < charts.length; chartIndex += 1) {
      const chart = charts[chartIndex] as WorksheetChartObject;
      const sourceRange = chart.sourceRange;
      const anchor = buildAnchor(chart);
      const pushPlaceholder = (reason: string): void => {
        entries.push({
          kind: "placeholder",
          chartId: chart.id,
          drawingObjectId: chartIndex + 1,
          anchor,
          placeholderText: `Gráfico não suportado para exportação (${reason}).`
        });
      };
      if (!sourceRange) {
        pushPlaceholder("range ausente");
        continue;
      }
      const sourceSheet = workbook.sheets.find((item) => item.id === sourceRange.sheetId);
      const sourceSheetName = sourceSheet ? sourceSheet.name : sheetNameById.get(sourceRange.sheetId);
      if (!sourceSheet || !sourceSheetName) {
        pushPlaceholder("sheet de origem inválida");
        continue;
      }
      const traceNameOverrides = Array.isArray(chart.figure.data)
        ? chart.figure.data.map((trace) => {
            if (!trace || typeof trace !== "object") {
              return "";
            }
            const nameValue = (trace as Record<string, unknown>).name;
            return typeof nameValue === "string" ? sanitizeSpreadsheetText(nameValue, 80) : "";
          })
        : undefined;
      const series = resolveSeriesFromRangeBinding(sourceSheet, sourceSheetName, sourceRange, traceNameOverrides);
      if (!series.length) {
        pushPlaceholder("séries não encontradas");
        continue;
      }
      const exportProfile = EXPORT_PROFILE_BY_TYPE[chart.type] ?? "line";
      const exportSeries = exportProfile === "scatter" ? series.filter((entry) => Boolean(entry.categoryRef)) : series;
      if (!exportSeries.length) {
        pushPlaceholder("séries incompatíveis com perfil");
        continue;
      }
      const legendVisible = resolveChartLegendVisible(chart);
      const xAxisConfig = resolveChartAxisConfig(chart, "x");
      const yAxisConfig = resolveChartAxisConfig(chart, "y");
      const title = chart.title?.trim() || undefined;
      const chartSequence = nextGlobalIndex;
      const chartPath = `xl/charts/chart${chartSequence}.xml`;
      const relationId = `rId-chart-${chartIndex + 1}`;
      const drawingObjectId = chartIndex + 1;
      const chartXml =
        exportProfile === "pie" || exportProfile === "donut"
          ? buildPieChartXml({
              profile: exportProfile,
              title,
              legendVisible,
              series: [exportSeries[0] as ChartExportSeries]
            })
          : buildCartesianChartXml({
              profile: exportProfile,
              title,
              legendVisible,
              series: exportSeries,
              chartSequence,
              xAxisTitle: xAxisConfig.title,
              yAxisTitle: yAxisConfig.title,
              xAxisVisible: xAxisConfig.visible,
              yAxisVisible: yAxisConfig.visible,
              xAxisType: xAxisConfig.type,
              yAxisType: yAxisConfig.type
            });

      entries.push({
        kind: "chart",
        chartId: chart.id,
        chartPath,
        relationId,
        drawingObjectId,
        anchor,
        chartXml
      });
      nextGlobalIndex += 1;
    }

    return { entries, nextGlobalIndex };
  };

  const buildDrawingChartAnchorXml = (entry: ChartExportChartEntry): string =>
    "<xdr:oneCellAnchor>" +
    "<xdr:from>" +
    `<xdr:col>${entry.anchor.col}</xdr:col>` +
    `<xdr:colOff>${toEmu(entry.anchor.colOffsetPx)}</xdr:colOff>` +
    `<xdr:row>${entry.anchor.row}</xdr:row>` +
    `<xdr:rowOff>${toEmu(entry.anchor.rowOffsetPx)}</xdr:rowOff>` +
    "</xdr:from>" +
    `<xdr:ext cx="${toEmu(entry.anchor.widthPx)}" cy="${toEmu(entry.anchor.heightPx)}"/>` +
    "<xdr:graphicFrame macro=\"\">" +
    "<xdr:nvGraphicFramePr>" +
    `<xdr:cNvPr id="${entry.drawingObjectId}" name="${sanitizeXmlAttribute(`Chart ${entry.drawingObjectId}`)}"/>` +
    "<xdr:cNvGraphicFramePr/>" +
    "</xdr:nvGraphicFramePr>" +
    "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>" +
    "<a:graphic>" +
    "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" +
    `<c:chart xmlns:c="${CHART_NS}" xmlns:r="${DOC_REL_NS}" r:id="${sanitizeXmlAttribute(entry.relationId)}"/>` +
    "</a:graphicData>" +
    "</a:graphic>" +
    "</xdr:graphicFrame>" +
    "<xdr:clientData/>" +
    "</xdr:oneCellAnchor>";

  const buildDrawingPlaceholderAnchorXml = (entry: ChartExportPlaceholderEntry): string =>
    "<xdr:oneCellAnchor>" +
    "<xdr:from>" +
    `<xdr:col>${entry.anchor.col}</xdr:col>` +
    `<xdr:colOff>${toEmu(entry.anchor.colOffsetPx)}</xdr:colOff>` +
    `<xdr:row>${entry.anchor.row}</xdr:row>` +
    `<xdr:rowOff>${toEmu(entry.anchor.rowOffsetPx)}</xdr:rowOff>` +
    "</xdr:from>" +
    `<xdr:ext cx="${toEmu(entry.anchor.widthPx)}" cy="${toEmu(entry.anchor.heightPx)}"/>` +
    "<xdr:sp macro=\"\">" +
    "<xdr:nvSpPr>" +
    `<xdr:cNvPr id="${entry.drawingObjectId}" name="${sanitizeXmlAttribute(`Chart Placeholder ${entry.drawingObjectId}`)}"/>` +
    "<xdr:cNvSpPr txBox=\"1\"/>" +
    "<xdr:nvPr/>" +
    "</xdr:nvSpPr>" +
    "<xdr:spPr>" +
    "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>" +
    "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>" +
    "<a:solidFill><a:srgbClr val=\"F8FAFC\"/></a:solidFill>" +
    "<a:ln><a:solidFill><a:srgbClr val=\"94A3B8\"/></a:solidFill></a:ln>" +
    "</xdr:spPr>" +
    "<xdr:txBody>" +
    "<a:bodyPr/><a:lstStyle/>" +
    `<a:p><a:r><a:rPr lang="en-US"/><a:t>${sanitizeXmlText(sanitizeSpreadsheetText(entry.placeholderText, 180))}</a:t></a:r></a:p>` +
    "</xdr:txBody>" +
    "</xdr:sp>" +
    "<xdr:clientData/>" +
    "</xdr:oneCellAnchor>";

  const createChartExportFeature = (): WriterFeature<any> => ({
    files: {
      transform: {
        "[Content_Types].xml": {
          insert: (sheetOptions) => {
            const typedSheetOptions = sheetOptions as Array<{ chartsForExport?: ChartExportEntry[] }>;
            const chartPaths = Array.from(
              new Set(
                typedSheetOptions.flatMap((item) =>
                  (item.chartsForExport ?? [])
                    .filter((chart): chart is ChartExportChartEntry => chart.kind === "chart")
                    .map((chart) => chart.chartPath)
                )
              )
            );
            return chartPaths
              .map(
                (chartPath) =>
                  `<Override ContentType="${CHART_CONTENT_TYPE}" PartName="/${sanitizeXmlAttribute(chartPath)}"/>`
              )
              .join("");
          }
        },
        "xl/drawings/drawing{id}.xml": {
          insert: (sheetOptions) =>
            ((sheetOptions as { chartsForExport?: ChartExportEntry[] }).chartsForExport ?? [])
              .map((entry) => (entry.kind === "chart" ? buildDrawingChartAnchorXml(entry) : buildDrawingPlaceholderAnchorXml(entry)))
              .join("")
        },
        "xl/drawings/_rels/drawing{id}.xml.rels": {
          insert: (sheetOptions) =>
            ((sheetOptions as { chartsForExport?: ChartExportEntry[] }).chartsForExport ?? [])
              .filter((entry): entry is ChartExportChartEntry => entry.kind === "chart")
              .map(
                (entry) =>
                  `<Relationship Id="${sanitizeXmlAttribute(entry.relationId)}" Type="${CHART_REL_TYPE}" Target="../charts/${sanitizeXmlAttribute(
                    basename(entry.chartPath)
                  )}"/>`
              )
              .join("")
        }
      },
      write: {
        files: (sheetOptions) => {
          const typedSheetOptions = sheetOptions as Array<{ chartsForExport?: ChartExportEntry[] }>;
          const files: Record<string, string> = {};
          for (const option of typedSheetOptions) {
            for (const chart of (option.chartsForExport ?? []).filter(
              (entry): entry is ChartExportChartEntry => entry.kind === "chart"
            )) {
              files[chart.chartPath] = chart.chartXml;
            }
          }
          return Object.keys(files).length ? files : undefined;
        }
      }
    }
  });

  const exportWorkbookToXlsx = async (
    workbook: WorkbookModel,
    options?: XlsxChartInteropOptions
  ): Promise<Uint8Array> => {
    const sheetNameById = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet.name]));
    let chartSequence = 1;
    let hasCharts = false;
    const sheets: WriterSheetWithCharts[] = workbook.sheets.map((sheet) => {
      const sourceCharts = (sheet.charts ?? []).filter((chart) => chart.state.visible !== false);
      if (sourceCharts.length > MAX_EXPORT_CHARTS_PER_SHEET) {
        for (const hiddenByLimit of sourceCharts.slice(MAX_EXPORT_CHARTS_PER_SHEET)) {
          options?.onChartUnsupportedFeature?.({
            sheetId: sheet.id,
            chartId: hiddenByLimit.id,
            feature: "export-limit:sheet-chart-cap"
          });
        }
      }
      const bounds = getBounds(sheet);
      const { anchors, covered } = findMergeAnchorMap(sheet.merges);
      const data = Array.from({ length: bounds.rowCount }, (_, row) =>
        Array.from({ length: bounds.columnCount }, (_, col) => {
          const key = getCellKey(row, col);
          if (covered.has(key)) {
            return null;
          }

          const merge = anchors.get(key);
          const cell = sheet.cells[key];
          const style = getCellStyle(cell?.style, sheet.rows[row]?.style, sheet.columns[col]?.style);
          const rowHeight = sheet.rows[row]?.height;
          return toWriterCell(cell, style, rowHeight, merge);
        })
      );

      const chartExport = buildChartExportEntries({
        workbook,
        sheet,
        sheetNameById,
        globalStartIndex: chartSequence
      });
      chartSequence = chartExport.nextGlobalIndex;
      hasCharts = hasCharts || chartExport.entries.length > 0;
      for (const entry of chartExport.entries) {
        options?.onChartExported?.({
          sheetId: sheet.id,
          chartId: entry.chartId
        });
        if (entry.kind === "placeholder") {
          options?.onChartUnsupportedFeature?.({
            sheetId: sheet.id,
            chartId: entry.chartId,
            feature: "export-fallback:placeholder"
          });
        }
        const sourceChart = sourceCharts.find((chart) => chart.id === entry.chartId);
        for (const unsupportedFeature of sourceChart?.excelInterop?.unsupportedFeatures ?? []) {
          options?.onChartUnsupportedFeature?.({
            sheetId: sheet.id,
            chartId: entry.chartId,
            feature: unsupportedFeature
          });
        }
      }

      return {
        sheet: sheet.name,
        columns: Array.from({ length: bounds.columnCount }, (_, col) => ({
          width: ((sheet.columns[col]?.width ?? workbook.settings.columnWidth) / EXCEL_COLUMN_WIDTH_UNIT) || undefined
        })),
        data,
        chartsForExport: chartExport.entries
      };
    });

    try {
      const buffer = await writeExcelFile(
        sheets,
        hasCharts
          ? {
              features: [createChartExportFeature()]
            }
          : undefined
      ).toBuffer();
      return new Uint8Array(buffer);
    } catch (error) {
      options?.onChartError?.({
        errorCode: "XLSX_CHART_EXPORT_FAILED",
        message: error instanceof Error ? error.message : "Failed to export workbook charts."
      });
      throw error;
    }
  };

  const importWorkbookFromXlsx = async (
    input: Uint8Array | ArrayBuffer,
    options?: XlsxChartInteropOptions
  ): Promise<WorkbookModel> => {
    try {
      const contents = await unpackXlsxFile(toBuffer(input));
      const filePaths = parseFilePaths(contents["xl/_rels/workbook.xml.rels"], xmlModule);
      const workbookInfo = parseSpreadsheetInfo(contents["xl/workbook.xml"], xmlModule);
      const sharedStrings = parseSharedStrings(
        filePaths.sharedStrings ? contents[filePaths.sharedStrings] : undefined,
        xmlModule
      );
      const valueStyles = parseValueStyles(filePaths.styles ? contents[filePaths.styles] : undefined, xmlModule);
      const styleTable = parseStyleTable(filePaths.styles ? contents[filePaths.styles] : undefined);

    const parsedSheets = workbookInfo.sheets.map((sheetInfo, index) => {
      const sheetPath = filePaths.sheets[sheetInfo.relationId];
      const normalizedSheetPath = sheetPath ? normalizeArchivePath(sheetPath) : "";
      const document =
        normalizedSheetPath && contents[normalizedSheetPath]
          ? xmlModule.createDocument(contents[normalizedSheetPath])
          : xmlModule.createDocument("<worksheet/>");
      const cells = parseSheetCells(document, sharedStrings, valueStyles, styleTable, workbookInfo.epoch1904);
      const merges = getElementChildren(getFirstElementChild(document.documentElement, "mergeCells") ?? document.documentElement, "mergeCell")
        .map((merge) => merge.getAttribute("ref"))
        .filter((reference): reference is string => Boolean(reference))
        .map(parseMergeReference);
      const columns = parseColumnSchema(document, styleTable);
      const rows = parseRowSchema(document, styleTable);
      const dimension = parseSheetDimensionReference(document);

      let rowCount = (dimension?.end.row ?? 0) + 1;
      let columnCount = (dimension?.end.col ?? 0) + 1;
      for (const key of Object.keys(cells)) {
        const [rowText, colText] = key.split(":");
        rowCount = Math.max(rowCount, Number(rowText) + 1);
        columnCount = Math.max(columnCount, Number(colText) + 1);
      }
      for (const merge of merges) {
        rowCount = Math.max(rowCount, merge.end.row + 1);
        columnCount = Math.max(columnCount, merge.end.col + 1);
      }
      for (const key of Object.keys(rows)) {
        rowCount = Math.max(rowCount, Number(key) + 1);
      }
      for (const key of Object.keys(columns)) {
        columnCount = Math.max(columnCount, Number(key) + 1);
      }

      const sheet: WorkbookSheet = {
        id: createSheetId(index),
        name: sheetInfo.name || DEFAULT_SHEET_NAME,
        rowCount,
        columnCount,
        cells,
        merges,
        columns,
        rows,
        selection: {
          start: { row: 0, col: 0 },
          end: { row: 0, col: 0 }
        }
      };

      return {
        sheet,
        sheetPath: normalizedSheetPath
      };
    });

    const sheetNameToId = new Map(parsedSheets.map(({ sheet }) => [sheet.name.toLowerCase(), sheet.id]));
    const sheetById = new Map(parsedSheets.map(({ sheet }) => [sheet.id, sheet]));
    for (const parsedSheet of parsedSheets) {
      if (!parsedSheet.sheetPath || !contents[parsedSheet.sheetPath]) {
        continue;
      }
      const importedCharts = collectWorksheetChartsFromDrawing({
        sheet: parsedSheet.sheet,
        sheetPath: parsedSheet.sheetPath,
        contents,
        sheetNameToId,
        sheetById,
        startZIndex: (parsedSheet.sheet.charts ?? []).length + 1
      });
      if (importedCharts.length) {
        parsedSheet.sheet.charts = importedCharts;
        for (const chart of importedCharts) {
          options?.onChartImported?.({
            sheetId: parsedSheet.sheet.id,
            chart
          });
          for (const unsupportedFeature of chart.excelInterop?.unsupportedFeatures ?? []) {
            options?.onChartUnsupportedFeature?.({
              sheetId: parsedSheet.sheet.id,
              chartId: chart.id,
              feature: unsupportedFeature
            });
          }
        }
        for (const chart of importedCharts) {
          try {
            const fromAddress = cellLabelToAddress(normalizeCellLabel(chart.position.fromCell));
            parsedSheet.sheet.rowCount = Math.max(parsedSheet.sheet.rowCount, fromAddress.row + 1);
            parsedSheet.sheet.columnCount = Math.max(parsedSheet.sheet.columnCount, fromAddress.col + 1);
            if (chart.position.toCell) {
              const toAddress = cellLabelToAddress(normalizeCellLabel(chart.position.toCell));
              parsedSheet.sheet.rowCount = Math.max(parsedSheet.sheet.rowCount, toAddress.row + 1);
              parsedSheet.sheet.columnCount = Math.max(parsedSheet.sheet.columnCount, toAddress.col + 1);
            }
          } catch {
            options?.onChartError?.({
              sheetId: parsedSheet.sheet.id,
              chartId: chart.id,
              errorCode: "XLSX_CHART_IMPORT_MALFORMED_POSITION",
              message: "Ignored malformed chart anchor metadata."
            });
            // Ignore malformed labels from imported metadata.
          }
        }
      }
    }
    const sheets = parsedSheets.map(({ sheet }) => sheet);

      return {
        id: "workbook-imported",
        sheets: sheets.length
          ? sheets
          : [
              {
                id: createSheetId(0),
                name: DEFAULT_SHEET_NAME,
                rowCount: 1,
                columnCount: 1,
                cells: {},
                merges: [],
                columns: {},
                rows: {},
                selection: {
                  start: { row: 0, col: 0 },
                  end: { row: 0, col: 0 }
                }
              }
            ],
        activeSheetId: sheets[0]?.id ?? createSheetId(0),
        metadata: {},
        settings: workbookDefaults
      };
    } catch (error) {
      options?.onChartError?.({
        errorCode: "XLSX_CHART_IMPORT_FAILED",
        message: error instanceof Error ? error.message : "Failed to import workbook charts."
      });
      throw error;
    }
  };

  const exportTableToXlsx = async <Object extends object>(
    rows: Object[],
    options: XlsxTableExportOptions<Object>
  ): Promise<Uint8Array> => {
    const data = getSheetData(rows, options.columns);
    const sheet: WriterSheet<Buffer> = {
      sheet: options.sheet ?? DEFAULT_SHEET_NAME,
      columns: options.columns.map((column) => ({ width: column.width })),
      data
    };
    const buffer = await writeExcelFile([sheet]).toBuffer();
    return new Uint8Array(buffer);
  };

  const importTableFromXlsx = async <Object extends object, ColumnTitle extends string = string>(
    input: Uint8Array | ArrayBuffer,
    options: XlsxTableImportOptions<Object, ColumnTitle>
  ): Promise<ParseSheetDataResult<Object, ColumnTitle>> => {
    const workbook = await readExcelFile(toBuffer(input));
    const resolvedSheet =
      typeof options.sheet === "string"
        ? workbook.find((sheet) => sheet.sheet === options.sheet)
        : workbook[(typeof options.sheet === "number" ? options.sheet : 1) - 1];

    if (!resolvedSheet) {
      throw createXlsxOperationError("XLSX_SHEET_NOT_FOUND", `Sheet not found: ${String(options.sheet ?? 1)}`, {
        sheet: options.sheet ?? 1
      });
    }

    return parseSheetData(resolvedSheet.data, options.schema);
  };

  return {
    exportWorkbookToXlsx,
    importWorkbookFromXlsx,
    exportTableToXlsx,
    importTableFromXlsx
  };
};

export type { ParseSheetDataResult, ReadSheet, Schema, WriterColumn };