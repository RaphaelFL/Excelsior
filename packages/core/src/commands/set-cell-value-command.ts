import type { CellModel, CommandResult, SpreadsheetError, SpreadsheetOperation } from "../domain/types";
import type { CommandContext, SpreadsheetCommand } from "./command-types";
import { createSheetNotFoundError } from "../errors/spreadsheet-operation-error";
import { getCellKey } from "../utils/cell-key";
import { cloneValue } from "../utils/clone";
import { recalculateWorkbookFormulas } from "../utils/recalculate-formulas";
import { CellValidationError } from "../validation/cell-validation-error";

const createRangeError = (message: string): SpreadsheetError => ({
  code: "CORE_RANGE_INVALID",
  message,
  area: "core",
  recoverable: true
});

export class SetCellValueCommand implements SpreadsheetCommand {
  readonly type = "SetCellValueCommand";

  readonly sheetId: string;

  constructor(
    private readonly input: {
      sheetId: string;
      row: number;
      col: number;
      value: string | number | boolean | null;
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

    if (
      this.input.row < 0 ||
      this.input.col < 0 ||
      this.input.row >= sheet.rowCount ||
      this.input.col >= sheet.columnCount
    ) {
      throw createRangeError("Cell address is outside the current sheet bounds.");
    }

    const key = getCellKey(this.input.row, this.input.col);
    const previousCell = sheet.cells[key];
    const validationResult = context.validationRegistry.validateCellValue({
      workbook,
      sheet,
      cell: previousCell,
      row: this.input.row,
      col: this.input.col,
      value: this.input.value,
      validation: previousCell?.validation
    });

    if (!validationResult.valid && validationResult.error) {
      throw new CellValidationError(this.input.sheetId, { row: this.input.row, col: this.input.col }, validationResult.error);
    }

    const nextCell: CellModel = {
      ...previousCell,
      value: this.input.value,
      computedValue: this.input.value,
      error: undefined,
      formula: undefined
    };

    if (typeof this.input.value === "string" && this.input.value.startsWith("=")) {
      nextCell.formula = this.input.value;
      nextCell.computedValue = null;
    }

    sheet.cells[key] = nextCell;

    const { formulaEngine } = context;

    recalculateWorkbookFormulas(workbook, formulaEngine, [
      {
        sheetId: sheet.id,
        row: this.input.row,
        col: this.input.col
      }
    ]);

    const operations: SpreadsheetOperation[] = [
      {
        op: previousCell ? "replace" : "add",
        id: sheet.id,
        path: ["cells", key],
        value: sheet.cells[key]
      }
    ];

    return {
      workbook,
      operations,
      affectedRanges: [
        {
          start: { row: this.input.row, col: this.input.col },
          end: { row: this.input.row, col: this.input.col }
        }
      ],
      recordHistory: true
    };
  }
}