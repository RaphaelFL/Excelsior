import { packZip, unpackZip } from "./native-zip";
import type { WriterCell, WriterCellValue } from "./xlsx-types";

export interface NativeSheet {
  sheet: string;
  columns?: Array<{ width?: number }>;
  data: WriterCellValue[][];
}

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const escapeXml = (value: string): string => value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&apos;");

const columnLabel = (index: number): string => {
  let label = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    label = String.fromCodePoint(65 + ((value - 1) % 26)) + label;
  }
  return label;
};

const normalizeCell = (cell: WriterCellValue): WriterCell | undefined => {
  if (cell == null) return undefined;
  if (cell instanceof Date) return { value: cell, type: Date };
  if (typeof cell === "object") return cell;
  return { value: cell };
};

const colorValue = (value?: string): string | undefined => {
  const normalized = value?.replace(/^#/, "").toUpperCase();
  return normalized && /^[0-9A-F]{6}$/.exec(normalized) ? normalized : undefined;
};

const styleKey = (cell: WriterCell): string => JSON.stringify({
  align: cell.align,
  alignVertical: cell.alignVertical,
  backgroundColor: colorValue(cell.backgroundColor),
  textColor: colorValue(cell.textColor),
  fontFamily: cell.fontFamily,
  fontSize: cell.fontSize,
  fontWeight: cell.fontWeight,
  fontStyle: cell.fontStyle,
  textDecoration: cell.textDecoration,
  wrap: cell.wrap,
  format: cell.format,
  indent: cell.indent,
  leftBorderColor: colorValue(cell.leftBorderColor),
  leftBorderStyle: cell.leftBorderStyle,
  rightBorderColor: colorValue(cell.rightBorderColor),
  rightBorderStyle: cell.rightBorderStyle,
  topBorderColor: colorValue(cell.topBorderColor),
  topBorderStyle: cell.topBorderStyle,
  bottomBorderColor: colorValue(cell.bottomBorderColor),
  bottomBorderStyle: cell.bottomBorderStyle
});

const hasStyle = (cell: WriterCell): boolean => styleKey(cell) !== "{}";

const buildStyles = (cells: WriterCell[]): { xml: string; ids: Map<string, number> } => {
  const unique = Array.from(new Map(cells.filter(hasStyle).map((cell) => [styleKey(cell), cell])).values());
  const ids = new Map(unique.map((cell, index) => [styleKey(cell), index + 1]));
  const fonts = [{}, ...unique.map((cell) => cell)];
  const fills = [undefined, undefined, ...unique.map((cell) => colorValue(cell.backgroundColor))];
  const borders = [{}, ...unique.map((cell) => cell)];
  const customFormats = Array.from(new Set(unique.map((cell) => cell.format).filter((value): value is string => Boolean(value))));
  const numFmtId = (format?: string): number => format ? 164 + customFormats.indexOf(format) : 0;
  const fontXml = fonts.map((cell) => {
    const typed = cell as WriterCell;
    return `<font>${typed.fontWeight ? "<b/>" : ""}${typed.fontStyle ? "<i/>" : ""}${typed.textDecoration?.underline ? "<u/>" : ""}${typed.textDecoration?.strikethrough ? "<strike/>" : ""}${typed.fontSize ? `<sz val="${typed.fontSize}"/>` : ""}${colorValue(typed.textColor) ? `<color rgb="FF${colorValue(typed.textColor)}"/>` : ""}${typed.fontFamily ? `<name val="${escapeXml(typed.fontFamily)}"/>` : ""}</font>`;
  }).join("");
  const fillXml = fills.map((color, index) => index < 2
    ? `<fill><patternFill patternType="${index === 0 ? "none" : "gray125"}"/></fill>`
    : `<fill><patternFill patternType="solid"><fgColor rgb="FF${color ?? "FFFFFF"}"/><bgColor indexed="64"/></patternFill></fill>`).join("");
  const side = (name: string, style?: string, color?: string): string => `<${name}${style ? ` style="${escapeXml(style)}"` : ""}>${color ? `<color rgb="FF${color}"/>` : ""}</${name}>`;
  const borderXml = borders.map((cell) => {
    const typed = cell as WriterCell;
    return `<border>${side("left", typed.leftBorderStyle, colorValue(typed.leftBorderColor))}${side("right", typed.rightBorderStyle, colorValue(typed.rightBorderColor))}${side("top", typed.topBorderStyle, colorValue(typed.topBorderColor))}${side("bottom", typed.bottomBorderStyle, colorValue(typed.bottomBorderColor))}<diagonal/></border>`;
  }).join("");
  const xfXml = ["<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>", ...unique.map((cell, index) => {
    const alignment = cell.align || cell.alignVertical || cell.wrap || cell.indent !== undefined
      ? `<alignment${cell.align ? ` horizontal="${cell.align}"` : ""}${cell.alignVertical ? ` vertical="${cell.alignVertical}"` : ""}${cell.wrap ? " wrapText=\"1\"" : ""}${cell.indent !== undefined ? ` indent="${cell.indent}"` : ""}/>`
      : "";
    return `<xf numFmtId="${numFmtId(cell.format)}" fontId="${index + 1}" fillId="${index + 2}" borderId="${index + 1}" xfId="0" applyFont="1" applyFill="1" applyBorder="1"${cell.format ? " applyNumberFormat=\"1\"" : ""}${alignment ? " applyAlignment=\"1\"" : ""}>${alignment}</xf>`;
  })].join("");
  return {
    ids,
    xml: `${XML_HEADER}<styleSheet xmlns="${MAIN_NS}">${customFormats.length ? `<numFmts count="${customFormats.length}">${customFormats.map((format, index) => `<numFmt numFmtId="${164 + index}" formatCode="${escapeXml(format)}"/>`).join("")}</numFmts>` : ""}<fonts count="${fonts.length}">${fontXml}</fonts><fills count="${fills.length}">${fillXml}</fills><borders count="${borders.length}">${borderXml}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${unique.length + 1}">${xfXml}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`
  };
};

const dateSerial = (date: Date): number => date.getTime() / 86_400_000 + 25_569;

const buildCellXml = (cell: WriterCell, reference: string, styleId?: number): string => {
  const style = styleId ? ` s="${styleId}"` : "";
  if (cell.type === "Formula") {
    const formula = String(cell.value ?? "").replace(/^=/, "");
    return `<c r="${reference}"${style}><f>${escapeXml(formula)}</f><v></v></c>`;
  }
  const value = cell.value;
  if (value instanceof Date || cell.type === Date) return `<c r="${reference}"${style}><v>${dateSerial(value instanceof Date ? value : new Date(String(value)))}</v></c>`;
  if (typeof value === "boolean") return `<c r="${reference}" t="b"${style}><v>${value ? 1 : 0}</v></c>`;
  if (typeof value === "number") return `<c r="${reference}"${style}><v>${Number.isFinite(value) ? value : 0}</v></c>`;
  if (value === undefined) return `<c r="${reference}"${style}/>`;
  return `<c r="${reference}" t="inlineStr"${style}><is><t xml:space="preserve">${escapeXml(String(value))}</t></is></c>`;
};

const buildSheetXml = (sheet: NativeSheet, styleIds: Map<string, number>): string => {
  const merges: string[] = [];
  const rows = sheet.data.map((row, rowIndex) => {
    let height: number | undefined;
    const cells = row.map((rawCell, colIndex) => {
      const cell = normalizeCell(rawCell);
      if (!cell) return "";
      if (cell.height !== undefined) height = Math.max(height ?? 0, cell.height);
      if ((cell.columnSpan ?? 1) > 1 || (cell.rowSpan ?? 1) > 1) {
        merges.push(`${columnLabel(colIndex)}${rowIndex + 1}:${columnLabel(colIndex + (cell.columnSpan ?? 1) - 1)}${rowIndex + (cell.rowSpan ?? 1)}`);
      }
      return buildCellXml(cell, `${columnLabel(colIndex)}${rowIndex + 1}`, styleIds.get(styleKey(cell)));
    }).join("");
    return cells || height !== undefined ? `<row r="${rowIndex + 1}"${height !== undefined ? ` ht="${height}" customHeight="1"` : ""}>${cells}</row>` : "";
  }).join("");
  const maxColumns = Math.max(1, sheet.data.reduce((max, row) => Math.max(max, row.length), sheet.columns?.length ?? 0));
  const dimension = `A1:${columnLabel(maxColumns - 1)}${Math.max(1, sheet.data.length)}`;
  const columns = sheet.columns?.map((column, index) => column.width === undefined ? "" : `<col min="${index + 1}" max="${index + 1}" width="${column.width}" customWidth="1"/>`).join("") ?? "";
  return `${XML_HEADER}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><dimension ref="${dimension}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>${columns ? `<cols>${columns}</cols>` : ""}<sheetData>${rows}</sheetData>${merges.length ? `<mergeCells count="${merges.length}">${merges.map((merge) => `<mergeCell ref="${merge}"/>`).join("")}</mergeCells>` : ""}</worksheet>`;
};

export const buildWorkbookFiles = (sheets: NativeSheet[]): Record<string, string | Uint8Array> => {
  const safeSheets = sheets.length ? sheets : [{ sheet: "Sheet1", data: [] }];
  const allCells = safeSheets.flatMap((sheet) => sheet.data.flatMap((row) => row.map(normalizeCell).filter((cell): cell is WriterCell => Boolean(cell))));
  const styles = buildStyles(allCells);
  const files: Record<string, string | Uint8Array> = {
    "[Content_Types].xml": `${XML_HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${safeSheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`,
    "_rels/.rels": `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `${XML_HEADER}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><bookViews><workbookView/></bookViews><sheets>${safeSheets.map((sheet, index) => `<sheet name="${escapeXml(sheet.sheet || `Sheet${index + 1}`)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}</sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${safeSheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("")}<Relationship Id="rId${safeSheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": styles.xml
  };
  safeSheets.forEach((sheet, index) => {
    files[`xl/worksheets/sheet${index + 1}.xml`] = buildSheetXml(sheet, styles.ids);
  });
  return files;
};

export const writeXlsxArchive = (files: Record<string, string | Uint8Array>): Uint8Array => packZip(files);

export const readXlsxArchive = async (input: Uint8Array): Promise<Record<string, string>> => Object.fromEntries(
  Object.entries(await unpackZip(input)).map(([path, data]) => [path, new TextDecoder().decode(data)])
);
