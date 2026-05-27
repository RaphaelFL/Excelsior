import { createSheetNotFoundError } from "../errors/spreadsheet-operation-error";
import { cloneValue } from "../utils/clone";
import type { CommandContext, SpreadsheetCommand } from "./command-types";

export class SelectRangeCommand implements SpreadsheetCommand {
  readonly type = "SelectRangeCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      rowStart: number;
      rowEnd: number;
      colStart: number;
      colEnd: number;
    }
  ) {
    this.sheetId = input.sheetId;
  }

  execute(context: CommandContext) {
    const workbook = cloneValue(context.workbook);
    const sheet = workbook.sheets.find((item) => item.id === this.input.sheetId);

    if (!sheet) {
      throw createSheetNotFoundError(this.input.sheetId);
    }

    sheet.selection = {
      start: {
        row: Math.max(0, Math.min(this.input.rowStart, this.input.rowEnd)),
        col: Math.max(0, Math.min(this.input.colStart, this.input.colEnd))
      },
      end: {
        row: Math.min(sheet.rowCount - 1, Math.max(this.input.rowStart, this.input.rowEnd)),
        col: Math.min(sheet.columnCount - 1, Math.max(this.input.colStart, this.input.colEnd))
      }
    };

    return {
      workbook,
      operations: [],
      affectedRanges: [sheet.selection],
      recordHistory: false
    };
  }
}