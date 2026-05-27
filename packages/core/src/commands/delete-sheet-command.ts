import type { CommandResult } from "../domain/types";
import { createCoreOperationError, createSheetNotFoundError } from "../errors/spreadsheet-operation-error";
import { cloneValue } from "../utils/clone";
import type { CommandContext, SpreadsheetCommand } from "./command-types";

export class DeleteSheetCommand implements SpreadsheetCommand {
  readonly type = "DeleteSheetCommand";

  readonly sheetId: string;

  constructor(sheetId: string) {
    this.sheetId = sheetId;
  }

  execute(context: CommandContext): CommandResult {
    const workbook = cloneValue(context.workbook);

    if (workbook.sheets.length === 1) {
      throw createCoreOperationError("CORE_WORKBOOK_MIN_SHEETS", "Workbook must keep at least one sheet.");
    }

    const currentIndex = workbook.sheets.findIndex((sheet) => sheet.id === this.sheetId);
    if (currentIndex < 0) {
      throw createSheetNotFoundError(this.sheetId);
    }

    const [deletedSheet] = workbook.sheets.splice(currentIndex, 1);
    const nextActiveSheet = workbook.sheets[currentIndex] ?? workbook.sheets[currentIndex - 1];
    workbook.activeSheetId = nextActiveSheet.id;

    return {
      workbook,
      operations: [
        {
          op: "deleteSheet",
          id: deletedSheet.id,
          path: ["sheets", currentIndex],
          value: { id: deletedSheet.id }
        }
      ],
      affectedRanges: [deletedSheet.selection],
      recordHistory: true
    };
  }
}