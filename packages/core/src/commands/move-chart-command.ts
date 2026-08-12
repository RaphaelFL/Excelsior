import type { ChartPosition, CommandResult } from "../domain/types";
import { cloneValue } from "../utils/clone";
import type { CommandContext, SpreadsheetCommand } from "./command-types";
import { ensureSheetCharts, findChartIndexOrThrow, getSheetByIdOrThrow } from "./chart-command-utils";

export class MoveChartCommand implements SpreadsheetCommand {
  readonly type = "MoveChartCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      chartId: string;
      position: Pick<ChartPosition, "fromCell" | "toCell" | "offsetX" | "offsetY" | "zIndex">;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext): CommandResult {
    const workbook = cloneValue(context.workbook);
    const sheet = getSheetByIdOrThrow(workbook, this.input.sheetId);
    const charts = ensureSheetCharts(sheet);
    const chartIndex = findChartIndexOrThrow(sheet, this.input.chartId);
    const chart = charts[chartIndex];
    const nextChart = {
      ...chart,
      position: {
        ...chart.position,
        ...cloneValue(this.input.position)
      }
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
