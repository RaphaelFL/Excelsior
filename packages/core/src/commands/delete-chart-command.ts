import type { CommandResult } from "../domain/types";
import { cloneValue } from "../utils/clone";
import type { CommandContext, SpreadsheetCommand } from "./command-types";
import { ensureSheetCharts, findChartIndexOrThrow, getSheetByIdOrThrow } from "./chart-command-utils";

export class DeleteChartCommand implements SpreadsheetCommand {
  readonly type = "DeleteChartCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      chartId: string;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext): CommandResult {
    const workbook = cloneValue(context.workbook);
    const sheet = getSheetByIdOrThrow(workbook, this.input.sheetId);
    const charts = ensureSheetCharts(sheet);
    const chartIndex = findChartIndexOrThrow(sheet, this.input.chartId);

    charts.splice(chartIndex, 1);
    return {
      workbook,
      operations: [
        {
          op: "remove",
          id: sheet.id,
          path: ["charts", chartIndex],
          value: null
        }
      ],
      affectedRanges: [sheet.selection],
      recordHistory: true
    };
  }
}
