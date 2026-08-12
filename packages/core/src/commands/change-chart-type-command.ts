import type { CommandResult, WorksheetChartType } from "../domain/types";
import type { CommandContext, SpreadsheetCommand } from "./command-types";
import { UpdateChartCommand } from "./update-chart-command";

export class ChangeChartTypeCommand implements SpreadsheetCommand {
  readonly type = "ChangeChartTypeCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      chartId: string;
      chartType: WorksheetChartType;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext): CommandResult {
    return new UpdateChartCommand({
      sheetId: this.input.sheetId,
      chartId: this.input.chartId,
      patch: {
        type: this.input.chartType
      }
    }).execute(context);
  }
}
