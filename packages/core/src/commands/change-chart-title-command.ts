import type { CommandResult } from "../domain/types";
import type { CommandContext, SpreadsheetCommand } from "./command-types";
import { UpdateChartCommand } from "./update-chart-command";

export class ChangeChartTitleCommand implements SpreadsheetCommand {
  readonly type = "ChangeChartTitleCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      chartId: string;
      title?: string;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext): CommandResult {
    return new UpdateChartCommand({
      sheetId: this.input.sheetId,
      chartId: this.input.chartId,
      patch: {
        title: this.input.title
      }
    }).execute(context);
  }
}
