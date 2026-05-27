import type { SpreadsheetError } from "../domain/types";

export class SpreadsheetOperationError extends Error {
  readonly details: SpreadsheetError;

  constructor(details: SpreadsheetError) {
    super(details.message);
    this.name = details.code;
    this.details = details;
  }
}

export const createCoreOperationError = (
  code: string,
  message: string,
  details?: Record<string, unknown>
): SpreadsheetOperationError =>
  new SpreadsheetOperationError({
    code,
    message,
    area: "core",
    recoverable: true,
    details
  });

export const createSheetNotFoundError = (
  sheetId: string,
  details?: Record<string, unknown>
): SpreadsheetOperationError =>
  createCoreOperationError("CORE_SHEET_NOT_FOUND", `Sheet not found: ${sheetId}`, {
    ...details,
    sheetId
  });

export const createWorkbookSheetRequiredError = (
  message = "Workbook requires at least one sheet."
): SpreadsheetOperationError => createCoreOperationError("CORE_WORKBOOK_REQUIRES_SHEET", message);