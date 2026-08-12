import type { CommandResult, WorksheetChartObject } from "../domain/types";
import { cloneValue } from "../utils/clone";
import type { CommandContext, SpreadsheetCommand } from "./command-types";
import { assertChartBelongsToSheet, ensureSheetCharts, getSheetByIdOrThrow } from "./chart-command-utils";

export class CreateChartCommand implements SpreadsheetCommand {
  readonly type = "CreateChartCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      chart: WorksheetChartObject;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext): CommandResult {
    const workbook = cloneValue(context.workbook);
    const sheet = getSheetByIdOrThrow(workbook, this.input.sheetId);
    const charts = ensureSheetCharts(sheet);
    const chart = cloneValue(this.input.chart);
    assertChartBelongsToSheet(chart, sheet.id);

    const existingIndex = charts.findIndex((item) => item.id === chart.id);
    if (existingIndex >= 0) {
      charts[existingIndex] = chart;
      return {
        workbook,
        operations: [
          {
            op: "replace",
            id: sheet.id,
            path: ["charts", existingIndex],
            value: chart
          }
        ],
        affectedRanges: [sheet.selection],
        recordHistory: true
      };
    }

    charts.push(chart);
    return {
      workbook,
      operations: [
        {
          op: "add",
          id: sheet.id,
          path: ["charts", charts.length - 1],
          value: chart
        }
      ],
      affectedRanges: [sheet.selection],
      recordHistory: true
    };
  }
}
