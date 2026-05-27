import { SpreadsheetOperationError, type CellModel, type CellStyle, type ColumnSchema, type RowSchema, type SheetMerge, type WorkbookModel } from "@excelsior/core";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
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

export interface XlsxTableExportOptions<Object extends object> {
  sheet?: string;
  columns: WriterColumn<Object>[];
}

export interface XlsxTableImportOptions<Object extends object, ColumnTitle extends string = string> {
  sheet?: number | string;
  schema: Schema<Object, ColumnTitle>;
}

export interface XlsxAdapter {
  exportWorkbookToXlsx(workbook: WorkbookModel): Promise<Uint8Array>;
  importWorkbookFromXlsx(input: Uint8Array | ArrayBuffer): Promise<WorkbookModel>;
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

  const exportWorkbookToXlsx = async (workbook: WorkbookModel): Promise<Uint8Array> => {
    const sheets: WriterSheet<Buffer>[] = workbook.sheets.map((sheet) => {
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

      return {
        sheet: sheet.name,
        columns: Array.from({ length: bounds.columnCount }, (_, col) => ({
          width: ((sheet.columns[col]?.width ?? workbook.settings.columnWidth) / EXCEL_COLUMN_WIDTH_UNIT) || undefined
        })),
        data
      };
    });

    const buffer = await writeExcelFile(sheets).toBuffer();
    return new Uint8Array(buffer);
  };

  const importWorkbookFromXlsx = async (input: Uint8Array | ArrayBuffer): Promise<WorkbookModel> => {
    const contents = await unpackXlsxFile(toBuffer(input));
    const filePaths = parseFilePaths(contents["xl/_rels/workbook.xml.rels"], xmlModule);
    const workbookInfo = parseSpreadsheetInfo(contents["xl/workbook.xml"], xmlModule);
    const sharedStrings = parseSharedStrings(
      filePaths.sharedStrings ? contents[filePaths.sharedStrings] : undefined,
      xmlModule
    );
    const valueStyles = parseValueStyles(filePaths.styles ? contents[filePaths.styles] : undefined, xmlModule);
    const styleTable = parseStyleTable(filePaths.styles ? contents[filePaths.styles] : undefined);

    const sheets = workbookInfo.sheets.map((sheetInfo, index) => {
      const sheetPath = filePaths.sheets[sheetInfo.relationId];
      const document = sheetPath ? xmlModule.createDocument(contents[sheetPath]) : xmlModule.createDocument("<worksheet/>");
      const cells = parseSheetCells(document, sharedStrings, valueStyles, styleTable, workbookInfo.epoch1904);
      const merges = getElementChildren(getFirstElementChild(document.documentElement, "mergeCells") ?? document.documentElement, "mergeCell")
        .map((merge) => merge.getAttribute("ref"))
        .filter((reference): reference is string => Boolean(reference))
        .map(parseMergeReference);
      const columns = parseColumnSchema(document, styleTable);
      const rows = parseRowSchema(document, styleTable);

      let rowCount = 1;
      let columnCount = 1;
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

      return {
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
    });

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