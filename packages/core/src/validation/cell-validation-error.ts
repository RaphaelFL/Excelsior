import type { CellAddress, SpreadsheetError } from "../domain/types";

export class CellValidationError extends Error {
  readonly details: SpreadsheetError;

  constructor(
    readonly sheetId: string,
    readonly address: CellAddress,
    details: SpreadsheetError
  ) {
    super(details.message);
    this.name = details.code;
    this.details = details;
  }
}