import type { SheetModel, WorkbookModel, WorksheetChartObject } from "../domain/types";
import { createCoreOperationError, createSheetNotFoundError } from "../errors/spreadsheet-operation-error";

export const getSheetByIdOrThrow = (workbook: WorkbookModel, sheetId: string): SheetModel => {
  const sheet = workbook.sheets.find((item) => item.id === sheetId);
  if (!sheet) {
    throw createSheetNotFoundError(sheetId);
  }
  return sheet;
};

export const ensureSheetCharts = (sheet: SheetModel): WorksheetChartObject[] => {
  if (!sheet.charts) {
    sheet.charts = [];
  }
  return sheet.charts;
};

export const findChartIndexOrThrow = (sheet: SheetModel, chartId: string): number => {
  const charts = ensureSheetCharts(sheet);
  const chartIndex = charts.findIndex((item) => item.id === chartId);
  if (chartIndex < 0) {
    throw createCoreOperationError("CORE_CHART_NOT_FOUND", `Chart not found: ${chartId}`, {
      sheetId: sheet.id,
      chartId
    });
  }
  return chartIndex;
};

export const assertChartBelongsToSheet = (chart: WorksheetChartObject, sheetId: string): void => {
  if (chart.sheetId !== sheetId) {
    throw createCoreOperationError(
      "CORE_CHART_INVALID_SHEET",
      `Chart '${chart.id}' belongs to '${chart.sheetId}', but command targeted '${sheetId}'.`,
      {
        chartId: chart.id,
        chartSheetId: chart.sheetId,
        sheetId
      }
    );
  }
};
