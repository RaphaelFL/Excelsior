import type { CommandResult, WorksheetChartObject } from "../domain/types";
import { cloneValue } from "../utils/clone";
import type { CommandContext, SpreadsheetCommand } from "./command-types";
import { ensureSheetCharts, findChartIndexOrThrow, getSheetByIdOrThrow } from "./chart-command-utils";

type MutableRecord = Record<string, unknown>;

export class ChangeChartLegendCommand implements SpreadsheetCommand {
  readonly type = "ChangeChartLegendCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      chartId: string;
      visible: boolean;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext): CommandResult {
    const workbook = cloneValue(context.workbook);
    const sheet = getSheetByIdOrThrow(workbook, this.input.sheetId);
    const charts = ensureSheetCharts(sheet);
    const chartIndex = findChartIndexOrThrow(sheet, this.input.chartId);
    const current = charts[chartIndex];
    const figure = cloneValue(current.figure) as unknown as MutableRecord;
    const layout = toMutableRecord(figure.layout);
    const legend = toMutableRecord(layout.legend);
    const nextChart: WorksheetChartObject = {
      ...current,
      figure: {
        ...current.figure,
        ...figure,
        layout: {
          ...layout,
          legend: {
            ...legend,
            visible: this.input.visible
          }
        }
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

const toMutableRecord = (value: unknown): MutableRecord => {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return value as MutableRecord;
  }
  return {};
};
