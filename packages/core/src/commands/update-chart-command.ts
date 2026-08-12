import type { CommandResult, WorksheetChartObject } from "../domain/types";
import { cloneValue } from "../utils/clone";
import type { CommandContext, SpreadsheetCommand } from "./command-types";
import { ensureSheetCharts, findChartIndexOrThrow, getSheetByIdOrThrow } from "./chart-command-utils";

export class UpdateChartCommand implements SpreadsheetCommand {
  readonly type = "UpdateChartCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      chartId: string;
      patch: Partial<Omit<WorksheetChartObject, "id" | "sheetId">>;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext): CommandResult {
    const workbook = cloneValue(context.workbook);
    const sheet = getSheetByIdOrThrow(workbook, this.input.sheetId);
    const charts = ensureSheetCharts(sheet);
    const chartIndex = findChartIndexOrThrow(sheet, this.input.chartId);
    const currentChart = charts[chartIndex];

    const nextChart: WorksheetChartObject = {
      ...currentChart,
      ...cloneValue(this.input.patch),
      id: currentChart.id,
      sheetId: currentChart.sheetId,
      sourceRange: this.input.patch.sourceRange
        ? {
            ...currentChart.sourceRange,
            ...cloneValue(this.input.patch.sourceRange)
          }
        : currentChart.sourceRange,
      position: this.input.patch.position
        ? {
            ...currentChart.position,
            ...cloneValue(this.input.patch.position)
          }
        : currentChart.position,
      style: this.input.patch.style
        ? {
            ...currentChart.style,
            ...cloneValue(this.input.patch.style)
          }
        : currentChart.style,
      state: this.input.patch.state
        ? {
            ...currentChart.state,
            ...cloneValue(this.input.patch.state)
          }
        : currentChart.state,
      excelInterop: this.input.patch.excelInterop
        ? {
            ...currentChart.excelInterop,
            ...cloneValue(this.input.patch.excelInterop)
          }
        : currentChart.excelInterop
    };

    charts[chartIndex] = nextChart;
    return {
      workbook,
      operations: [
        {
          op: "replace",
          id: sheet.id,
          path: ["charts", chartIndex],
          value: nextChart
        }
      ],
      affectedRanges: [sheet.selection],
      recordHistory: true
    };
  }
}
