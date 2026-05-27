import type { SpreadsheetOperation, WorkbookModel } from "../domain/types";
import { createSheetNotFoundError } from "../errors/spreadsheet-operation-error";
import { deleteAxis, insertAxis } from "./sheet-structure";

const getSheet = (workbook: WorkbookModel, sheetRef: string) => {
  const sheet = workbook.sheets.find((item) => item.id === sheetRef);
  if (!sheet) {
    throw createSheetNotFoundError(sheetRef, { operation: "applyOperations" });
  }
  return sheet;
};

const setAtPath = (target: Record<string, unknown>, path: Array<string | number>, value: unknown): void => {
  if (!path.length) {
    return;
  }

  let cursor: Record<string | number, unknown> = target;
  for (let index = 0; index < path.length - 1; index += 1) {
    const segment = path[index];
    const nextSegment = path[index + 1];
    const existing = cursor[segment];

    if (existing == null || typeof existing !== "object") {
      cursor[segment] = typeof nextSegment === "number" ? [] : {};
    }

    cursor = cursor[segment] as Record<string | number, unknown>;
  }

  const lastSegment = path.at(-1);
  if (lastSegment == null) {
    return;
  }

  cursor[lastSegment] = value;
};

const removeAtPath = (target: Record<string, unknown>, path: Array<string | number>): void => {
  if (!path.length) {
    return;
  }

  let cursor: Record<string | number, unknown> | undefined = target;
  for (let index = 0; index < path.length - 1; index += 1) {
    const segment = path[index];
    const next = cursor?.[segment];
    if (next == null || typeof next !== "object") {
      return;
    }
    cursor = next as Record<string | number, unknown>;
  }

  if (!cursor) {
    return;
  }

  const lastSegment = path.at(-1);
  if (lastSegment == null) {
    return;
  }

  if (Array.isArray(cursor) && typeof lastSegment === "number") {
    cursor.splice(lastSegment, 1);
    return;
  }

  delete cursor[lastSegment];
};

export const applyOperationsToWorkbook = (
  workbook: WorkbookModel,
  operations: SpreadsheetOperation[]
): WorkbookModel => {
  for (const operation of operations) {
    switch (operation.op) {
      case "add":
      case "replace": {
        const sheet = getSheet(workbook, operation.id);
        setAtPath(sheet as unknown as Record<string, unknown>, operation.path, operation.value);
        break;
      }
      case "remove": {
        const sheet = getSheet(workbook, operation.id);
        removeAtPath(sheet as unknown as Record<string, unknown>, operation.path);
        break;
      }
      case "insertRowCol": {
        const sheet = getSheet(workbook, operation.id);
        const payload = operation.value as {
          type: "row" | "column";
          index: number;
          count: number;
        };
        insertAxis(sheet, payload.type, payload.index, payload.count);
        break;
      }
      case "deleteRowCol": {
        const sheet = getSheet(workbook, operation.id);
        const payload = operation.value as {
          type: "row" | "column";
          start: number;
          end: number;
        };
        deleteAxis(sheet, payload.type, payload.start, payload.end - payload.start + 1);
        break;
      }
      case "addSheet": {
        workbook.sheets.push(operation.value as WorkbookModel["sheets"][number]);
        workbook.activeSheetId = operation.id;
        break;
      }
      case "deleteSheet": {
        const index = workbook.sheets.findIndex((sheet) => sheet.id === operation.id);
        if (index >= 0) {
          workbook.sheets.splice(index, 1);
          workbook.activeSheetId = workbook.sheets[index]?.id ?? workbook.sheets[index - 1]?.id ?? workbook.activeSheetId;
        }
        break;
      }
    }
  }

  return workbook;
};