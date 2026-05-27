import type { CellRange, CommandResult, SpreadsheetOperation } from "../domain/types";
import { cloneValue } from "../utils/clone";
import { applyOperationsToWorkbook } from "../utils/apply-operations";
import { deriveFormulaRecalculationTargets, recalculateWorkbookFormulas } from "../utils/recalculate-formulas";
import type { CommandContext, SpreadsheetCommand } from "./command-types";

export class UpdateSheetOperationsCommand implements SpreadsheetCommand {
  readonly type = "UpdateSheetOperationsCommand";

  constructor(
    readonly sheetId: string,
    private readonly operations: SpreadsheetOperation[],
    private readonly affectedRanges: CellRange[]
  ) {}

  execute(context: CommandContext): CommandResult {
    const workbook = cloneValue(context.workbook);
    const dirtyCells = deriveFormulaRecalculationTargets(this.operations);
    const nextWorkbook = recalculateWorkbookFormulas(
      applyOperationsToWorkbook(workbook, this.operations),
      context.formulaEngine,
      dirtyCells
    );

    return {
      workbook: nextWorkbook,
      operations: this.operations,
      affectedRanges: this.affectedRanges,
      recordHistory: true
    };
  }
}