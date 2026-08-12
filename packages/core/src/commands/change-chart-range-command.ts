import type { ChartRangeBinding, CommandResult } from "../domain/types";
import type { CommandContext, SpreadsheetCommand } from "./command-types";
import { UpdateChartCommand } from "./update-chart-command";

export class ChangeChartRangeCommand implements SpreadsheetCommand {
  readonly type = "ChangeChartRangeCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      chartId: string;
      sourceRange: ChartRangeBinding;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext): CommandResult {
    return new UpdateChartCommand({
      sheetId: this.input.sheetId,
      chartId: this.input.chartId,
      patch: {
        sourceRange: this.input.sourceRange
      }
    }).execute(context);
  }
}
