import type { ChartFigureInput } from "../model/Figure";
import type { CartesianTrace, TracePoint } from "../model/Trace";

export type SpreadsheetCellValue = string | number | boolean | Date | null | undefined;

export interface SpreadsheetRangeInput {
  sheetName?: string;
  headers: string[];
  rows: SpreadsheetCellValue[][];
}

export interface SpreadsheetAdapterOptions {
  title?: string;
  traceType?: "line" | "scatter" | "bar";
  xColumn?: string | number;
  seriesColumns?: Array<string | number>;
}

export const createFigureFromSpreadsheetRange = (
  range: SpreadsheetRangeInput,
  options: SpreadsheetAdapterOptions = {}
): ChartFigureInput => {
  const headers = range.headers.map((header) => String(header ?? "").trim());
  if (headers.length === 0) {
    throw new Error("Spreadsheet range requires at least one header.");
  }
  if (range.rows.length === 0) {
    throw new Error("Spreadsheet range requires at least one data row.");
  }

  const xColumnIndex = resolveColumnIndex(options.xColumn ?? 0, headers);
  const seriesColumnIndexes = (options.seriesColumns ?? inferSeriesColumns(headers.length, xColumnIndex)).map((column) =>
    resolveColumnIndex(column, headers)
  );

  const xValues = range.rows.map((row) => normalizePointValue(row[xColumnIndex], row.length > xColumnIndex ? row[xColumnIndex] : ""));
  const traces: CartesianTrace[] = [];

  for (const seriesIndex of seriesColumnIndexes) {
    const yValues = range.rows.map((row) => toNumber(row[seriesIndex]));
    if (yValues.every((value) => !Number.isFinite(value))) {
      continue;
    }

    traces.push({
      type: options.traceType ?? "line",
      name: headers[seriesIndex] || `Serie ${seriesIndex + 1}`,
      x: xValues,
      y: yValues.map((value) => (Number.isFinite(value) ? value : 0)),
      mode: options.traceType === "scatter" ? "markers" : options.traceType === "line" ? "lines" : "lines+markers"
    });
  }

  if (traces.length === 0) {
    throw new Error("Spreadsheet range did not produce any numeric series for chart traces.");
  }

  return {
    data: traces,
    layout: {
      title: options.title ?? range.sheetName ?? "Chart from spreadsheet range",
      xAxis: {
        title: headers[xColumnIndex] || "X"
      },
      yAxis: {
        title: "Value"
      }
    },
    metadata: {
      source: "spreadsheet-range",
      sheetName: range.sheetName ?? "",
      columns: headers.length,
      rows: range.rows.length
    }
  };
};

const resolveColumnIndex = (value: string | number, headers: string[]): number => {
  if (typeof value === "number") {
    if (!Number.isInteger(value) || value < 0 || value >= headers.length) {
      throw new Error(`Invalid column index '${value}'.`);
    }
    return value;
  }

  const byHeader = headers.findIndex((header) => header.toLowerCase() === value.trim().toLowerCase());
  if (byHeader < 0) {
    throw new Error(`Column '${value}' was not found in headers.`);
  }
  return byHeader;
};

const inferSeriesColumns = (headersLength: number, xColumnIndex: number): number[] =>
  Array.from({ length: headersLength }, (_, index) => index).filter((index) => index !== xColumnIndex);

const normalizePointValue = (value: SpreadsheetCellValue, fallback: SpreadsheetCellValue): TracePoint => {
  if (value instanceof Date) {
    return value;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  const source = value ?? fallback;
  return String(source ?? "");
};

const toNumber = (value: SpreadsheetCellValue): number => {
  if (typeof value === "number") {
    return value;
  }
  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }
  if (value instanceof Date) {
    return value.getTime();
  }
  if (typeof value === "string") {
    const normalized = value.replace(",", ".").trim();
    const numeric = Number(normalized);
    return Number.isFinite(numeric) ? numeric : Number.NaN;
  }
  return Number.NaN;
};
