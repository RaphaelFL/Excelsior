import type { CommandResult, FormulaEngine, WorkbookModel } from "../domain/types";
import type { ValidationRegistry } from "../validation/validation-registry";

export interface CommandContext {
  workbook: WorkbookModel;
  formulaEngine?: FormulaEngine;
  validationRegistry: ValidationRegistry;
  now: () => number;
}

export interface SpreadsheetCommand {
  readonly type: string;
  readonly sheetId?: string;
  execute: (context: CommandContext) => CommandResult;
}