import {
  SpreadsheetOperationError,
  cellAddressToLabel,
  cellLabelToAddress,
  type CellAddress,
  type CellComment,
  type CellModel,
  type CellPrimitive,
  type CellStyle,
  type CellValidationConfig,
  type CellValidationRule,
  type ColumnSchema,
  type RowSchema,
  type SheetMerge,
  type WorkbookModel
} from "@excelsior/core";
import { buildWorkbookFiles, readXlsxArchive, writeXlsxArchive, type NativeSheet } from "./native-ooxml";
import { createXmlDocument } from "./xml-tree";
import type {
  ParseSheetDataError,
  ParseSheetDataResult,
  ReadSheet,
  Schema,
  SheetCellValue,
  WriterCell,
  WriterColumn
} from "./xlsx-types";

const basename = (path: string): string => path.slice(path.lastIndexOf("/") + 1);
const dirname = (path: string): string => {
  const index = path.lastIndexOf("/");
  return index < 0 ? "." : path.slice(0, index);
};
const normalizePosixPath = (path: string): string => {
  const segments: string[] = [];
  for (const segment of path.replaceAll("\\", "/").split("/")) {
    if (!segment || segment === ".") continue;
    if (segment === "..") segments.pop();
    else segments.push(segment);
  }
  return segments.join("/");
};
const posixPath = {
  dirname,
  join: (...parts: string[]) => parts.join("/"),
  normalize: normalizePosixPath
};
const decodeBase64Utf8 = (payload: string): string => {
  const binary = atob(payload);
  return new TextDecoder().decode(Uint8Array.from(binary, (character) => character.charCodeAt(0)));
};
const encodeBase64Utf8 = (value: string): string => {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

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

const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const VML_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const EXCELSIOR_METADATA_PATH = "customXml/excelsior.xml";
const EXCELSIOR_METADATA_CONTENT_TYPE = "application/vnd.excelsior.metadata+xml";
const MAX_EXCELSIOR_METADATA_LENGTH = 2_000_000;
const MAX_EXCELSIOR_METADATA_CELLS = 10_000;

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

type WriterSheetWithFeatures = NativeSheet & {
  excelsiorMetadata?: string;
  nativeDataValidations?: string;
  nativeComments?: { commentsXml: string; vmlXml: string };
};

interface ExcelsiorWorkbookMetadata {
  version: 1;
  sheets: Array<{
    name: string;
    cells: Record<string, { note?: string; comments?: CellComment[]; validation?: CellValidationConfig }>;
  }>;
}

const buildExcelsiorMetadata = (workbook: WorkbookModel): string | undefined => {
  const sheets = workbook.sheets.map((sheet) => ({
    name: sheet.name,
    cells: Object.fromEntries(
      Object.entries(sheet.cells)
        .filter(([, cell]) => cell.note !== undefined || cell.comments?.length || cell.validation)
        .map(([address, cell]) => [address, {
          ...(cell.note !== undefined ? { note: cell.note } : {}),
          ...(cell.comments?.length ? { comments: cell.comments } : {}),
          ...(cell.validation ? { validation: cell.validation } : {})
        }])
    )
  })).filter((sheet) => Object.keys(sheet.cells).length > 0);
  if (!sheets.length) {
    return undefined;
  }
  const serialized = JSON.stringify({ version: 1, sheets } satisfies ExcelsiorWorkbookMetadata);
  return serialized.length <= MAX_EXCELSIOR_METADATA_LENGTH ? serialized : undefined;
};

const parseExcelsiorMetadata = (content: string | undefined, maxCellLength: number): ExcelsiorWorkbookMetadata | undefined => {
  if (!content || content.length > MAX_EXCELSIOR_METADATA_LENGTH) {
    return undefined;
  }
  try {
    const payload = content.match(/<excelsiorMetadata[^>]*>([A-Za-z0-9+/=]+)<\/excelsiorMetadata>/)?.[1];
    if (!payload) {
      return undefined;
    }
    const decoded = decodeBase64Utf8(payload);
    if (decoded.length > MAX_EXCELSIOR_METADATA_LENGTH) {
      return undefined;
    }
    const parsed = JSON.parse(decoded) as Partial<ExcelsiorWorkbookMetadata>;
    if (parsed.version !== 1 || !Array.isArray(parsed.sheets)) {
      return undefined;
    }
    let cellCount = 0;
    const sheets = parsed.sheets.flatMap((sheet) => {
      if (!sheet || typeof sheet.name !== "string" || !sheet.cells || typeof sheet.cells !== "object") {
        return [];
      }
      const cells: ExcelsiorWorkbookMetadata["sheets"][number]["cells"] = {};
      for (const [address, rawCell] of Object.entries(sheet.cells)) {
        if (++cellCount > MAX_EXCELSIOR_METADATA_CELLS || !/^\d+:\d+$/.test(address) || !rawCell || typeof rawCell !== "object") {
          continue;
        }
        const candidate = rawCell as { note?: unknown; comments?: unknown; validation?: unknown };
        const note = typeof candidate.note === "string" ? candidate.note.slice(0, maxCellLength) : undefined;
        const comments = Array.isArray(candidate.comments)
          ? candidate.comments.slice(0, 100).flatMap((rawComment): CellComment[] => {
              const comment = rawComment as Partial<CellComment>;
              if (typeof comment.id !== "string" || typeof comment.content !== "string" || typeof comment.author?.id !== "string") {
                return [];
              }
              return [{
                id: comment.id.slice(0, 120),
                author: { id: comment.author.id.slice(0, 120), name: comment.author.name?.slice(0, 120) },
                content: comment.content.slice(0, maxCellLength),
                createdAt: Number.isFinite(comment.createdAt) ? Number(comment.createdAt) : 0,
                updatedAt: Number.isFinite(comment.updatedAt) ? Number(comment.updatedAt) : undefined,
                resolved: comment.resolved === true,
                replies: Array.isArray(comment.replies) ? comment.replies.slice(0, 100).flatMap((rawReply) => {
                  const reply = rawReply as CellComment["replies"][number];
                  return typeof reply?.id === "string" && typeof reply.content === "string" && typeof reply.author?.id === "string"
                    ? [{ ...reply, id: reply.id.slice(0, 120), content: reply.content.slice(0, maxCellLength), author: { id: reply.author.id.slice(0, 120), name: reply.author.name?.slice(0, 120) } }]
                    : [];
                }) : []
              }];
            })
          : undefined;
        const validation = candidate.validation && typeof candidate.validation === "object" && Array.isArray((candidate.validation as CellValidationConfig).rules)
          ? { rules: (candidate.validation as CellValidationConfig).rules.slice(0, 32) }
          : undefined;
        cells[address] = { ...(note !== undefined ? { note } : {}), ...(comments?.length ? { comments } : {}), ...(validation ? { validation } : {}) };
      }
      return [{ name: sheet.name.slice(0, 120), cells }];
    });
    return { version: 1, sheets };
  } catch {
    return undefined;
  }
};

const createSheetId = (index: number): string => `sheet-import-${index + 1}`;

const getCellKey = (row: number, col: number): string => `${row}:${col}`;

const normalizeImportedFormula = (formula: string): string =>
  formula.startsWith("=") ? formula : `=${formula}`;

const toBuffer = (input: Uint8Array | ArrayBuffer): Uint8Array =>
  input instanceof Uint8Array ? new Uint8Array(input) : new Uint8Array(input);

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
): WriterCell | null => {
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
    return writerCell;
  }

  if (cell?.value !== undefined && cell.value !== null) {
    writerCell.value = cell.value as string | number | boolean;
    return writerCell;
  }

  if (Object.keys(writerCell).length) {
    return writerCell;
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

const buildNativeValidationRuleXml = (rule: CellValidationRule, cellLabel: string): string | undefined => {
  let type: "decimal" | "whole" | "textLength" | "list";
  let operator: string | undefined;
  let formula1: string | undefined;
  let formula2: string | undefined;
  if (rule.type === "number" || rule.type === "range") {
    type = rule.type === "range" || (Number.isInteger(rule.min) && Number.isInteger(rule.max)) ? "whole" : "decimal";
    if (rule.min !== undefined && rule.max !== undefined) {
      operator = "between";
      formula1 = String(rule.min);
      formula2 = String(rule.max);
    } else if (rule.min !== undefined) {
      operator = "greaterThanOrEqual";
      formula1 = String(rule.min);
    } else if (rule.max !== undefined) {
      operator = "lessThanOrEqual";
      formula1 = String(rule.max);
    } else {
      return undefined;
    }
  } else if (rule.type === "length") {
    type = "textLength";
    if (rule.min !== undefined && rule.max !== undefined) {
      operator = "between";
      formula1 = String(rule.min);
      formula2 = String(rule.max);
    } else if (rule.min !== undefined) {
      operator = "greaterThanOrEqual";
      formula1 = String(rule.min);
    } else if (rule.max !== undefined) {
      operator = "lessThanOrEqual";
      formula1 = String(rule.max);
    } else {
      return undefined;
    }
  } else if (rule.type === "list" || rule.type === "dropdown") {
    const values = rule.values.slice(0, 128).map((value) => String(value ?? "").replaceAll('"', '""'));
    const serialized = `"${values.join(",")}"`;
    if (!values.length || serialized.length > 255) {
      return undefined;
    }
    type = "list";
    formula1 = serialized;
  } else {
    return undefined;
  }
  const attributes = [
    `type="${type}"`,
    ...(operator ? [`operator="${operator}"`] : []),
    `allowBlank="1"`,
    `showErrorMessage="1"`,
    `errorStyle="${rule.severity === "warning" ? "warning" : "stop"}"`,
    ...(rule.message ? [`error="${sanitizeXmlAttribute(sanitizeSpreadsheetText(rule.message, 225))}"`] : []),
    `sqref="${sanitizeXmlAttribute(cellLabel)}"`
  ];
  return `<dataValidation ${attributes.join(" ")}><formula1>${sanitizeXmlText(formula1)}</formula1>${
    formula2 === undefined ? "" : `<formula2>${sanitizeXmlText(formula2)}</formula2>`
  }</dataValidation>`;
};

const buildNativeDataValidationsXml = (sheet: WorkbookSheet): string | undefined => {
  const entries = Object.entries(sheet.cells).flatMap(([key, cell]) => {
    const [row, col] = key.split(":").map(Number);
    if (!Number.isInteger(row) || !Number.isInteger(col) || !cell.validation?.rules.length) {
      return [];
    }
    const ruleXml = cell.validation.rules.map((rule) => buildNativeValidationRuleXml(rule, cellAddressToLabel({ row: row!, col: col! }))).find(Boolean);
    return ruleXml ? [ruleXml] : [];
  }).slice(0, MAX_EXCELSIOR_METADATA_CELLS);
  return entries.length ? `<dataValidations count="${entries.length}">${entries.join("")}</dataValidations>` : undefined;
};

const buildNativeComments = (sheet: WorkbookSheet): { commentsXml: string; vmlXml: string } | undefined => {
  const comments = Object.entries(sheet.cells).flatMap(([key, cell]) => {
    const [row, col] = key.split(":").map(Number);
    const sourceComment = cell.comments?.[0];
    const content = sourceComment?.content ?? cell.note;
    if (!Number.isInteger(row) || !Number.isInteger(col) || !content) {
      return [];
    }
    return [{
      row: row!,
      col: col!,
      ref: cellAddressToLabel({ row: row!, col: col! }),
      author: sanitizeSpreadsheetText(sourceComment?.author.name ?? sourceComment?.author.id ?? "Excelsior", 120),
      content: content.slice(0, workbookDefaults.maxCellLength)
    }];
  }).slice(0, MAX_EXCELSIOR_METADATA_CELLS);
  if (!comments.length) {
    return undefined;
  }
  const authors = Array.from(new Set(comments.map((comment) => comment.author)));
  const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>${authors.map((author) => `<author>${sanitizeXmlText(author)}</author>`).join("")}</authors><commentList>${comments.map((comment) => `<comment ref="${comment.ref}" authorId="${authors.indexOf(comment.author)}"><text><t xml:space="preserve">${sanitizeXmlText(comment.content)}</t></text></comment>`).join("")}</commentList></comments>`;
  const vmlXml = `<?xml version="1.0" encoding="UTF-8"?><xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${comments.map((comment, index) => `<v:shape id="_x0000_s${1025 + index}" type="#_x0000_t202" style="position:absolute;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${comment.col}, 15, ${comment.row}, 2, ${comment.col + 2}, 15, ${comment.row + 4}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${comment.row}</x:Row><x:Column>${comment.col}</x:Column></x:ClientData></v:shape>`).join("")}</xml>`;
  return { commentsXml, vmlXml };
};

const parseNativeDataValidations = (document: Document): Record<string, CellValidationConfig> => {
  const result: Record<string, CellValidationConfig> = {};
  const container = getFirstElementChild(document.documentElement, "dataValidations");
  if (!container) {
    return result;
  }
  let cellCount = 0;
  for (const element of getElementChildren(container, "dataValidation")) {
    const type = element.getAttribute("type");
    const operator = element.getAttribute("operator") ?? "between";
    const formula1 = getFirstElementChild(element, "formula1")?.textContent?.trim();
    const formula2 = getFirstElementChild(element, "formula2")?.textContent?.trim();
    const message = element.getAttribute("error")?.slice(0, 225) || undefined;
    const severity = element.getAttribute("errorStyle") === "warning" ? "warning" as const : "error" as const;
    let rule: CellValidationRule | undefined;
    if (type === "list" && formula1?.startsWith('"') && formula1.endsWith('"')) {
      rule = { type: "dropdown", values: formula1.slice(1, -1).split(",").map((value) => value.replaceAll('""', '"')).slice(0, 128), message, severity };
    } else if ((type === "whole" || type === "decimal" || type === "textLength") && formula1 !== undefined) {
      const first = Number(formula1);
      const second = formula2 === undefined ? undefined : Number(formula2);
      if (Number.isFinite(first) && (second === undefined || Number.isFinite(second))) {
        const bounds = operator === "greaterThanOrEqual"
          ? { min: first }
          : operator === "lessThanOrEqual"
            ? { max: first }
            : { min: first, ...(second === undefined ? {} : { max: second }) };
        rule = type === "textLength" ? { type: "length", ...bounds, message, severity } : { type: "number", ...bounds, message, severity };
      }
    }
    if (!rule) {
      continue;
    }
    for (const reference of (element.getAttribute("sqref") ?? "").split(/\s+/).filter(Boolean)) {
      const bounds = parseRangeAddress(reference);
      if (!bounds) {
        continue;
      }
      for (let row = bounds.start.row; row <= bounds.end.row && cellCount < MAX_EXCELSIOR_METADATA_CELLS; row += 1) {
        for (let col = bounds.start.col; col <= bounds.end.col && cellCount < MAX_EXCELSIOR_METADATA_CELLS; col += 1) {
          if (row >= workbookDefaults.maxRows || col >= workbookDefaults.maxColumns) {
            continue;
          }
          const key = getCellKey(row, col);
          result[key] = { rules: [...(result[key]?.rules ?? []), rule].slice(0, 32) };
          cellCount += 1;
        }
      }
    }
  }
  return result;
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

export interface XlsxAdapter {
  exportWorkbookToXlsx(workbook: WorkbookModel): Promise<Uint8Array>;
  importWorkbookFromXlsx(input: Uint8Array | ArrayBuffer): Promise<WorkbookModel>;
  exportTableToXlsx<Object extends object>(rows: Object[], options: XlsxTableExportOptions<Object>): Promise<Uint8Array>;
  importTableFromXlsx<Object extends object, ColumnTitle extends string = string>(
    input: Uint8Array | ArrayBuffer,
    options: XlsxTableImportOptions<Object, ColumnTitle>
  ): Promise<ParseSheetDataResult<Object, ColumnTitle>>;
}

export const createXlsxAdapter = (): XlsxAdapter => {
  const xmlModule = { createDocument: createXmlDocument };

  const parseCellCoordinates = (coordinate: string): [number, number] => {
    const address = cellLabelToAddress(normalizeCellLabel(coordinate));
    return [address.row + 1, address.col + 1];
  };

  const parseWorkbookPaths = (relationshipsContent?: string) => {
    const relationships = parseRelationships(relationshipsContent);
    const sheets: Record<string, string> = {};
    let sharedStrings: string | undefined;
    let styles: string | undefined;
    for (const relationship of relationships) {
      const path = resolveArchivePath("xl/workbook.xml", relationship.target);
      if (relationship.type.endsWith("/worksheet")) sheets[relationship.id] = path;
      else if (relationship.type.endsWith("/sharedStrings")) sharedStrings = path;
      else if (relationship.type.endsWith("/styles")) styles = path;
    }
    return { sheets, sharedStrings, styles };
  };

  const parseWorkbookInfo = (content?: string) => {
    const document = xmlModule.createDocument(content ?? "<workbook/>");
    const workbookProperties = getFirstElementChild(document.documentElement, "workbookPr");
    const sheetsElement = getFirstElementChild(document.documentElement, "sheets");
    return {
      epoch1904: workbookProperties?.getAttribute("date1904") === "1",
      sheets: getElementChildren(sheetsElement ?? document.documentElement, "sheet").map((sheet) => ({
        name: sheet.getAttribute("name") ?? DEFAULT_SHEET_NAME,
        relationId: sheet.getAttribute("r:id") ?? ""
      }))
    };
  };

  const parseSharedStrings = (content?: string): string[] => {
    if (!content) return [];
    const document = xmlModule.createDocument(content);
    return getElementChildren(document.documentElement, "si").map((item) =>
      findDescendantsByTagName(item, "t").map((text) => text.textContent ?? "").join("")
    );
  };

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

        const rawValue = valueElement?.textContent;
        const valueType = cell.getAttribute("t");
        let parsedValue: CellModel["value"] = null;
        if (valueType === "inlineStr") {
          parsedValue = extractInlineString(cell) ?? "";
        } else if (valueType === "s") {
          parsedValue = sharedStrings[Number(rawValue)] ?? "";
        } else if (valueType === "b") {
          parsedValue = rawValue === "1";
        } else if (valueType === "str") {
          parsedValue = rawValue ?? "";
        } else if (valueType !== "e" && rawValue !== undefined && rawValue !== "") {
          const numeric = Number(rawValue);
          const format = styleTable[Number(styleId ?? -1)]?.format ?? "";
          const isDateFormat = /(?:^|[^\\])[dmyhs]/i.test(format.replace(/\[[^\]]+]/g, ""));
          parsedValue = Number.isFinite(numeric)
            ? isDateFormat
              ? new Date((numeric - (epoch1904 ? 24_107 : 25_569)) * 86_400_000).toISOString()
              : numeric
            : rawValue;
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

  const parseNativeComments = (sheetPath: string, contents: Record<string, string>): Record<string, CellComment[]> => {
    const relationshipsPath = normalizeArchivePath(`${dirname(sheetPath)}/_rels/${basename(sheetPath)}.rels`);
    const commentsRelationship = parseRelationships(contents[relationshipsPath]).find((relationship) => relationship.type === COMMENTS_REL_TYPE);
    if (!commentsRelationship) {
      return {};
    }
    const commentsPath = resolveArchivePath(sheetPath, commentsRelationship.target);
    const content = contents[commentsPath];
    if (!content) {
      return {};
    }
    const document = xmlModule.createDocument(content);
    const authorsElement = getFirstElementChild(document.documentElement, "authors");
    const authors = getElementChildren(authorsElement ?? document.documentElement, "author").map((author) => (author.textContent ?? "").slice(0, 120));
    const commentList = getFirstElementChild(document.documentElement, "commentList");
    const result: Record<string, CellComment[]> = {};
    for (const comment of getElementChildren(commentList ?? document.documentElement, "comment").slice(0, MAX_EXCELSIOR_METADATA_CELLS)) {
      const reference = comment.getAttribute("ref");
      if (!reference) {
        continue;
      }
      const bounds = parseRangeAddress(reference);
      if (!bounds || bounds.start.row !== bounds.end.row || bounds.start.col !== bounds.end.col) {
        continue;
      }
      const contentText = findDescendantsByTagName(comment, "t").map((text) => text.textContent ?? "").join("").slice(0, workbookDefaults.maxCellLength);
      if (!contentText) {
        continue;
      }
      const authorId = Number(comment.getAttribute("authorId") ?? 0);
      const authorName = authors[authorId] || "Excel";
      const key = getCellKey(bounds.start.row, bounds.start.col);
      result[key] = [{
        id: `xlsx-comment-${bounds.start.row}-${bounds.start.col}`,
        author: { id: `xlsx-author-${Number.isInteger(authorId) ? authorId : 0}`, name: authorName },
        content: contentText,
        createdAt: 0,
        resolved: false,
        replies: []
      }];
    }
    return result;
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

  const addWorkbookFeatures = (
    files: Record<string, string | Uint8Array>,
    sheetOptions: WriterSheetWithFeatures[]
  ): void => {
    const appendBefore = (path: string, closingTag: string, markup: string): void => {
      const content = files[path];
      if (typeof content !== "string" || !content.includes(closingTag)) throw new Error(`Missing OOXML insertion point: ${path}`);
      files[path] = content.replace(closingTag, `${markup}${closingTag}`);
    };
    const contentTypeEntries: string[] = [];
    for (const [index, option] of sheetOptions.entries()) {
      const sheetId = index + 1;
      const sheetRelations: string[] = [];
      const sheetMarkup: string[] = [];
      if (option.nativeDataValidations) sheetMarkup.unshift(option.nativeDataValidations);
      if (option.nativeComments) {
        files[`xl/comments${sheetId}.xml`] = option.nativeComments.commentsXml;
        files[`xl/drawings/vmlDrawing${sheetId}.vml`] = option.nativeComments.vmlXml;
        contentTypeEntries.push(`<Override ContentType="${COMMENTS_CONTENT_TYPE}" PartName="/xl/comments${sheetId}.xml"/>`);
        sheetRelations.push(`<Relationship Id="rId-excelsior-comments" Type="${COMMENTS_REL_TYPE}" Target="../comments${sheetId}.xml"/><Relationship Id="rId-excelsior-vml" Type="${VML_REL_TYPE}" Target="../drawings/vmlDrawing${sheetId}.vml"/>`);
        sheetMarkup.push('<legacyDrawing r:id="rId-excelsior-vml"/>');
      }
      if (sheetMarkup.length) appendBefore(`xl/worksheets/sheet${sheetId}.xml`, "</worksheet>", sheetMarkup.join(""));
      if (sheetRelations.length) {
        files[`xl/worksheets/_rels/sheet${sheetId}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheetRelations.join("")}</Relationships>`;
      }
    }
    if (sheetOptions.some((item) => item.nativeComments)) contentTypeEntries.push('<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>');
    const metadata = sheetOptions.find((option) => option.excelsiorMetadata)?.excelsiorMetadata;
    if (metadata) {
      files[EXCELSIOR_METADATA_PATH] = `<excelsiorMetadata version="1">${encodeBase64Utf8(metadata)}</excelsiorMetadata>`;
      contentTypeEntries.push(`<Override ContentType="${EXCELSIOR_METADATA_CONTENT_TYPE}" PartName="/${EXCELSIOR_METADATA_PATH}"/>`);
    }
    if (contentTypeEntries.length) appendBefore("[Content_Types].xml", "</Types>", contentTypeEntries.join(""));
  };

  const exportWorkbookToXlsx = async (workbook: WorkbookModel): Promise<Uint8Array> => {
    let hasNativeDataValidations = false;
    let hasNativeComments = false;
    const excelsiorMetadata = buildExcelsiorMetadata(workbook);
    const sheets: WriterSheetWithFeatures[] = workbook.sheets.map((sheet) => {
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

      const nativeDataValidations = buildNativeDataValidationsXml(sheet);
      hasNativeDataValidations = hasNativeDataValidations || Boolean(nativeDataValidations);
      const nativeComments = buildNativeComments(sheet);
      hasNativeComments = hasNativeComments || Boolean(nativeComments);
      return {
        sheet: sheet.name,
        columns: Array.from({ length: bounds.columnCount }, (_, col) => ({
          width: ((sheet.columns[col]?.width ?? workbook.settings.columnWidth) / EXCEL_COLUMN_WIDTH_UNIT) || undefined
        })),
        data,
        ...(nativeDataValidations ? { nativeDataValidations } : {}),
        ...(nativeComments ? { nativeComments } : {}),
        ...(excelsiorMetadata ? { excelsiorMetadata } : {})
      };
    });

    const files = buildWorkbookFiles(sheets);
    if (hasNativeDataValidations || hasNativeComments || excelsiorMetadata) {
      addWorkbookFeatures(files, sheets);
    }
    return writeXlsxArchive(files);
  };

  const importWorkbookFromXlsx = async (input: Uint8Array | ArrayBuffer): Promise<WorkbookModel> => {
    const contents = await readXlsxArchive(toBuffer(input));
    const filePaths = parseWorkbookPaths(contents["xl/_rels/workbook.xml.rels"]);
    const workbookInfo = parseWorkbookInfo(contents["xl/workbook.xml"]);
    const sharedStrings = parseSharedStrings(filePaths.sharedStrings ? contents[filePaths.sharedStrings] : undefined);
    const styleTable = parseStyleTable(filePaths.styles ? contents[filePaths.styles] : undefined);

    const parsedSheets = workbookInfo.sheets.map((sheetInfo, index) => {
      const sheetPath = filePaths.sheets[sheetInfo.relationId];
      const normalizedSheetPath = sheetPath ? normalizeArchivePath(sheetPath) : "";
      const document =
        normalizedSheetPath && contents[normalizedSheetPath]
          ? xmlModule.createDocument(contents[normalizedSheetPath])
          : xmlModule.createDocument("<worksheet/>");
      const cells = parseSheetCells(document, sharedStrings, styleTable, workbookInfo.epoch1904);
      const nativeDataValidations = parseNativeDataValidations(document);
      for (const [address, validation] of Object.entries(nativeDataValidations)) {
        const existingCell = cells[address];
        cells[address] = existingCell ? { ...existingCell, validation } : { value: null, validation };
      }
      const nativeComments = normalizedSheetPath ? parseNativeComments(normalizedSheetPath, contents) : {};
      for (const [address, comments] of Object.entries(nativeComments)) {
        const existingCell = cells[address];
        cells[address] = existingCell ? { ...existingCell, comments } : { value: null, comments };
      }
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

    const excelsiorMetadata = parseExcelsiorMetadata(contents[EXCELSIOR_METADATA_PATH], workbookDefaults.maxCellLength);
    for (const metadataSheet of excelsiorMetadata?.sheets ?? []) {
      const targetSheet = parsedSheets.find(({ sheet }) => sheet.name === metadataSheet.name)?.sheet;
      if (!targetSheet) {
        continue;
      }
      for (const [address, metadataCell] of Object.entries(metadataSheet.cells)) {
        const [row, col] = address.split(":").map(Number);
        if (!Number.isInteger(row) || !Number.isInteger(col) || row! < 0 || col! < 0 || row! >= workbookDefaults.maxRows || col! >= workbookDefaults.maxColumns) {
          continue;
        }
        const existingCell = targetSheet.cells[address];
        targetSheet.cells[address] = existingCell ? { ...existingCell, ...metadataCell } : { value: null, ...metadataCell };
        targetSheet.rowCount = Math.max(targetSheet.rowCount, row! + 1);
        targetSheet.columnCount = Math.max(targetSheet.columnCount, col! + 1);
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
  };

  const exportTableToXlsx = async <Object extends object>(
    rows: Object[],
    options: XlsxTableExportOptions<Object>
  ): Promise<Uint8Array> => {
    const hasHeaders = options.columns.some((column) => column.header !== undefined);
    const data = [
      ...(hasHeaders ? [options.columns.map((column) => column.header ?? null)] : []),
      ...rows.map((row, rowIndex) => options.columns.map((column) => column.cell(row, rowIndex)))
    ];
    const sheet: NativeSheet = {
      sheet: options.sheet ?? DEFAULT_SHEET_NAME,
      columns: options.columns.map((column) => ({ width: column.width })),
      data
    };
    return writeXlsxArchive(buildWorkbookFiles([sheet]));
  };

  const parseTabularValue = (value: SheetCellValue, type: unknown): unknown => {
    if (value == null) return null;
    if (!type || type === String) return String(value);
    if (type === Number) {
      const parsed = typeof value === "number" ? value : Number(value);
      if (!Number.isFinite(parsed)) throw new Error("not_a_number");
      return parsed;
    }
    if (type === Boolean) {
      if (typeof value === "boolean") return value;
      if (value === 1 || String(value).toLowerCase() === "true") return true;
      if (value === 0 || String(value).toLowerCase() === "false") return false;
      throw new Error("not_a_boolean");
    }
    if (type === Date) {
      const parsed = value instanceof Date ? value : new Date(value as string | number);
      if (Number.isNaN(parsed.getTime())) throw new Error("not_a_date");
      return parsed;
    }
    if (typeof type === "function") return type(value);
    return value;
  };

  const parseTableData = <Object extends object, ColumnTitle extends string>(
    data: ReadSheet["data"],
    schema: Schema<Object, ColumnTitle>
  ): ParseSheetDataResult<Object, ColumnTitle> => {
    const headers = (data[0] ?? []).map((value) => String(value ?? ""));
    const errors: ParseSheetDataError<ColumnTitle>[] = [];
    const parseObject = (targetSchema: Record<string, unknown>, row: ReadSheet["data"][number], rowNumber: number, topLevel: Record<string, unknown>): Record<string, unknown> => {
      const object: Record<string, unknown> = {};
      for (const [property, rawEntry] of Object.entries(targetSchema)) {
        if (!rawEntry || typeof rawEntry !== "object") continue;
        const entry = rawEntry as { column?: ColumnTitle; type?: unknown; oneOf?: unknown[]; required?: boolean | ((value: Record<string, unknown>) => boolean); validate?: (value: unknown) => void; schema?: Record<string, unknown> };
        if (entry.schema) {
          object[property] = parseObject(entry.schema, row, rowNumber, topLevel);
          continue;
        }
        const column = entry.column as ColumnTitle;
        const columnIndex = headers.indexOf(String(column));
        const rawValue = columnIndex < 0 ? undefined : row[columnIndex];
        const required = typeof entry.required === "function" ? entry.required(topLevel) : entry.required === true;
        if (rawValue == null || rawValue === "") {
          if (required) errors.push({ row: rowNumber, column, columnIndex, error: "required", reason: undefined, value: rawValue, type: entry.type as never });
          object[property] = null;
          continue;
        }
        try {
          const parsed = parseTabularValue(rawValue, entry.type);
          if (entry.oneOf && !entry.oneOf.includes(parsed)) throw new Error("invalid");
          entry.validate?.(parsed);
          object[property] = parsed ?? null;
        } catch (error) {
          errors.push({ row: rowNumber, column, columnIndex, error: error instanceof Error ? error.message : "invalid", reason: undefined, value: rawValue, type: entry.type as never });
        }
      }
      return object;
    };
    const objects = data.slice(1).map((row, index) => {
      const topLevel: Record<string, unknown> = {};
      Object.assign(topLevel, parseObject(schema as Record<string, unknown>, row, index + 2, topLevel));
      return topLevel as Object;
    });
    return errors.length ? { objects: undefined, errors } : { objects, errors: undefined };
  };

  const importTableFromXlsx = async <Object extends object, ColumnTitle extends string = string>(
    input: Uint8Array | ArrayBuffer,
    options: XlsxTableImportOptions<Object, ColumnTitle>
  ): Promise<ParseSheetDataResult<Object, ColumnTitle>> => {
    const workbook = await importWorkbookFromXlsx(input);
    const sourceSheet = typeof options.sheet === "string"
      ? workbook.sheets.find((sheet) => sheet.name === options.sheet)
      : workbook.sheets[(typeof options.sheet === "number" ? options.sheet : 1) - 1];

    if (!sourceSheet) {
      throw createXlsxOperationError("XLSX_SHEET_NOT_FOUND", `Sheet not found: ${String(options.sheet ?? 1)}`, {
        sheet: options.sheet ?? 1
      });
    }
    const data: ReadSheet["data"] = Array.from({ length: sourceSheet.rowCount }, (_, row) =>
      Array.from({ length: sourceSheet.columnCount }, (_, col) => {
        const cell = sourceSheet.cells[getCellKey(row, col)];
        return (cell?.formula ? cell.computedValue : cell?.value) as ReadSheet["data"][number][number] ?? null;
      })
    );
    return parseTableData(data, options.schema);
  };

  return {
    exportWorkbookToXlsx,
    importWorkbookFromXlsx,
    exportTableToXlsx,
    importTableFromXlsx
  };
};

export type { ParseSheetDataResult, ReadSheet, Schema, WriterColumn };