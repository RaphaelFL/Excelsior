import type { CommandResult } from "../domain/types";
import { createSheetNotFoundError } from "../errors/spreadsheet-operation-error";
import { cloneValue } from "../utils/clone";
import { deleteAxis, type Axis } from "../utils/sheet-structure";
import type { CommandContext, SpreadsheetCommand } from "./command-types";

export class DeleteAxisCommand implements SpreadsheetCommand {
  readonly type = "DeleteAxisCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      axis: Axis;
      start: number;
      end: number;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext): CommandResult {
    const workbook = cloneValue(context.workbook);
    const sheet = workbook.sheets.find((item) => item.id === this.input.sheetId);

    if (!sheet) {
      throw createSheetNotFoundError(this.input.sheetId);
    }

    const count = this.input.end - this.input.start + 1;
    deleteAxis(sheet, this.input.axis, this.input.start, count);

    return {
      workbook,
      operations: [
        {
          op: "deleteRowCol",
          id: sheet.id,
          path: [],
          value: {
            type: this.input.axis,
            start: this.input.start,
            end: this.input.end,
            id: sheet.id
          }
        }
      ],
      affectedRanges: [sheet.selection],
      recordHistory: true
    };
  }
}