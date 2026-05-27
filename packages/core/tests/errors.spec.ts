import { describe, expect, it } from "vitest";
import { SpreadsheetOperationError, WorkbookEngine, cellLabelToAddress, type GridPlugin } from "../src/index";

const captureError = (callback: () => void): unknown => {
  try {
    callback();
  } catch (error) {
    return error;
  }

  throw new Error("Expected callback to throw.");
};

const expectTypedError = (error: unknown, code: string): SpreadsheetOperationError => {
  expect(error).toBeInstanceOf(SpreadsheetOperationError);
  const typed = error as SpreadsheetOperationError;
  expect(typed.name).toBe(code);
  expect(typed.details.code).toBe(code);
  expect(typed.details.area).toBe("core");
  return typed;
};

describe("typed core errors", () => {
  it("returns typed sheet-not-found errors from the workbook api", () => {
    const engine = new WorkbookEngine();
    const error = captureError(() => {
      engine.getSelection("missing-sheet");
    });

    const typed = expectTypedError(error, "CORE_SHEET_NOT_FOUND");
    expect(typed.details.details).toMatchObject({
      sheetId: "missing-sheet"
    });
  });

  it("returns typed duplicate-plugin errors", () => {
    const engine = new WorkbookEngine();
    const plugin: GridPlugin = {
      id: "audit-plugin",
      setup() {}
    };

    engine.registerPlugin(plugin);
    const error = captureError(() => {
      engine.registerPlugin(plugin);
    });

    const typed = expectTypedError(error, "CORE_PLUGIN_ALREADY_REGISTERED");
    expect(typed.details.details).toMatchObject({
      pluginId: "audit-plugin"
    });
  });

  it("returns typed duplicate-validator errors", () => {
    const engine = new WorkbookEngine();
    const validator = () => undefined;

    engine.registerValidator("money", validator);
    const error = captureError(() => {
      engine.registerValidator("money", validator);
    });

    const typed = expectTypedError(error, "CORE_VALIDATOR_ALREADY_REGISTERED");
    expect(typed.details.details).toMatchObject({
      validatorId: "money"
    });
  });

  it("returns typed invalid-address errors", () => {
    const error = captureError(() => {
      cellLabelToAddress("not-a-cell");
    });

    const typed = expectTypedError(error, "CORE_ADDRESS_INVALID_LABEL");
    expect(typed.details.details).toMatchObject({
      label: "not-a-cell"
    });
  });
});