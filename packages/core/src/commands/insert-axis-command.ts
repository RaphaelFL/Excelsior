import type { CommandResult } from "../domain/types";
import { createSheetNotFoundError } from "../errors/spreadsheet-operation-error";
import { cloneValue } from "../utils/clone";
import { insertAxis, type Axis } from "../utils/sheet-structure";
import type { CommandContext, SpreadsheetCommand } from "./command-types";

export class InsertAxisCommand implements SpreadsheetCommand {
  readonly type = "InsertAxisCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      axis: Axis;
      index: number;
      count: number;
      direction?: "lefttop" | "rightbottom";
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

    insertAxis(sheet, this.input.axis, this.input.index, this.input.count);

    return {
      workbook,
      operations: [
        {
          op: "insertRowCol",
          id: sheet.id,
          path: [],
          value: {
            type: this.input.axis,
            index: this.input.index,
            count: this.input.count,
            direction: this.input.direction ?? "rightbottom",
            id: sheet.id
          }
        }
      ],
      affectedRanges: [sheet.selection],
      recordHistory: true
    };
  }
}