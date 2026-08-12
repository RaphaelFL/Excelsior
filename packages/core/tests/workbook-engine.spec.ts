import { describe, expect, it, vi } from "vitest";
import type {
  FormulaEngine,
  FormulaEvaluationContext,
  FormulaEvaluationResult,
  FormulaReference,
  PivotBuildProgress,
  PivotModule
} from "../src/index";
import { CellValidationError, ServerSideRowModel, WorkbookEngine } from "../src/index";
import { BasicFormulaEngine } from "@excelsior/formulas";

class TrackingFormulaEngine implements FormulaEngine {
  private readonly base = new BasicFormulaEngine();

  readonly evaluations: string[] = [];

  evaluate(expression: string, context: FormulaEvaluationContext): FormulaEvaluationResult {
    this.evaluations.push(`${context.currentSheetName}!${context.currentCell.row}:${context.currentCell.col}`);
    return this.base.evaluate(expression, context);
  }

  collectReferences(expression: string): FormulaReference[] {
    return this.base.collectReferences?.(expression) ?? [];
  }
}

describe("WorkbookEngine", () => {
  const waitFor = async (predicate: () => boolean, message: string) => {
    for (let attempt = 0; attempt < 20; attempt += 1) {
      if (predicate()) {
        return;
      }

      await Promise.resolve();
    }

    throw new Error(message);
  };

  it("creates a workbook with one default sheet", () => {
    const engine = new WorkbookEngine();
    const snapshot = engine.getSnapshot();

    expect(snapshot.sheets).toHaveLength(1);
    expect(snapshot.sheets[0]?.name).toBe("Sheet1");
  });

  it("sets cell values through typed commands and emits serializable ops", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const operations = engine.setCellValue({
      sheetId: sheet.id,
      row: 1,
      col: 1,
      value: "42"
    });

    expect(engine.getDisplayValue(sheet.id, 1, 1)).toBe("42");
    expect(operations[0]).toMatchObject({
      op: "add",
      id: sheet.id,
      path: ["cells", "1:1"]
    });
  });

  it("supports undo and redo for value changes", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "A" });
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("A");

    expect(engine.undo()).toBe(true);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");

    expect(engine.redo()).toBe(true);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("A");
  });

  it("rejects cell values and formulas that exceed workbook input limits", () => {
    const engine = new WorkbookEngine({
      settings: {
        maxCellLength: 3,
        maxFormulaLength: 5
      }
    });
    const sheet = engine.getActiveSheet();

    expect(engine.validateCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "ABCD" })).toMatchObject({
      valid: false,
      error: {
        code: "CORE_VALIDATION_CELL_MAX_LENGTH"
      }
    });

    expect(engine.validateCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "=12345" })).toMatchObject({
      valid: false,
      error: {
        code: "CORE_VALIDATION_FORMULA_MAX_LENGTH"
      }
    });

    expect(() => engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "=12345" })).toThrow(CellValidationError);
  });

  it("treats redundant single-cell writes as no-op", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "A" });
    const operations = engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "A" });

    expect(operations).toEqual([]);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("A");
    expect(engine.undo()).toBe(true);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(engine.undo()).toBe(false);
  });

  it("preserves cell style metadata when value changes", () => {
    const engine = WorkbookEngine.fromJSON(
      {
        id: "workbook-1",
        activeSheetId: "sheet-1",
        metadata: {},
        settings: {
          maxRows: 100,
          maxColumns: 50,
          maxCellLength: 5000,
          maxFormulaLength: 2048,
          maxPasteCells: 10000,
          rowHeight: 28,
          columnWidth: 120,
          viewportBuffer: 6,
          maxHistorySize: 100,
          enableFormulas: true,
          clipboardPolicy: "text-only"
        },
        sheets: [
          {
            id: "sheet-1",
            name: "Sheet1",
            cells: {
              "0:0": {
                value: "draft",
                computedValue: "draft",
                style: {
                  backgroundColor: "#ffeecc",
                  fontWeight: "bold",
                  align: "center"
                },
                metadata: {
                  source: "import"
                }
              }
            },
            merges: [],
            columns: {},
            rows: {},
            rowCount: 10,
            columnCount: 10,
            selection: {
              start: { row: 0, col: 0 },
              end: { row: 0, col: 0 }
            }
          }
        ]
      },
      new BasicFormulaEngine()
    );

    engine.setCellValue({ sheetId: "sheet-1", row: 0, col: 0, value: "published" });

    expect(engine.getCell("sheet-1", 0, 0)).toMatchObject({
      value: "published",
      style: {
        backgroundColor: "#ffeecc",
        fontWeight: "bold",
        align: "center"
      },
      metadata: {
        source: "import"
      }
    });
  });

  it("keeps read models isolated while updating cached display values across revisions", () => {
    const engine = WorkbookEngine.fromJSON(
      {
        id: "workbook-1",
        activeSheetId: "sheet-1",
        metadata: {},
        settings: {
          maxRows: 100,
          maxColumns: 50,
          maxCellLength: 5000,
          maxFormulaLength: 2048,
          maxPasteCells: 10000,
          rowHeight: 28,
          columnWidth: 120,
          viewportBuffer: 6,
          maxHistorySize: 100,
          enableFormulas: true,
          clipboardPolicy: "text-only"
        },
        sheets: [
          {
            id: "sheet-1",
            name: "Sheet1",
            cells: {
              "0:0": {
                value: "draft",
                computedValue: "draft"
              }
            },
            merges: [
              {
                start: { row: 0, col: 0 },
                end: { row: 0, col: 1 }
              }
            ],
            columns: {
              0: { width: 180 }
            },
            rows: {
              0: { height: 40 }
            },
            rowCount: 10,
            columnCount: 10,
            selection: {
              start: { row: 0, col: 0 },
              end: { row: 0, col: 0 }
            }
          }
        ]
      },
      new BasicFormulaEngine()
    );

    const cell = engine.getCell("sheet-1", 0, 0);
    const merge = engine.getMerge("sheet-1", 0, 0);
    const column = engine.getColumnSchema("sheet-1", 0);
    const row = engine.getRowSchema("sheet-1", 0);

    expect(engine.getDisplayValue("sheet-1", 0, 0)).toBe("draft");
    expect(engine.getDisplayValue("sheet-1", 0, 0)).toBe("draft");

    if (cell) {
      cell.value = "mutated";
    }
    if (merge) {
      merge.end.col = 9;
    }
    if (column) {
      column.width = 999;
    }
    if (row) {
      row.height = 999;
    }

    expect(engine.getCell("sheet-1", 0, 0)?.value).toBe("draft");
    expect(engine.getMerge("sheet-1", 0, 0)).toEqual({
      start: { row: 0, col: 0 },
      end: { row: 0, col: 1 }
    });
    expect(engine.getColumnSchema("sheet-1", 0)?.width).toBe(180);
    expect(engine.getRowSchema("sheet-1", 0)?.height).toBe(40);

    engine.setCellValue({ sheetId: "sheet-1", row: 0, col: 0, value: "published" });

    expect(engine.getDisplayValue("sheet-1", 0, 0)).toBe("published");
  });

  it("applies cell styles, merges and column or row sizing through serializable ops", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    const styleOperations = engine.setCellStyle({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      style: {
        fontWeight: "bold",
        backgroundColor: "#d1fae5",
        wrap: true
      }
    });
    engine.resizeColumn(sheet.id, 0, 180);
    engine.resizeRow(sheet.id, 0, 40);
    const mergeOperations = engine.mergeCells({
      sheetId: sheet.id,
      start: { row: 0, col: 0 },
      end: { row: 1, col: 1 }
    });

    expect(styleOperations[0]).toMatchObject({
      op: "add",
      id: sheet.id,
      path: ["cells", "0:0"]
    });
    expect(mergeOperations[0]).toMatchObject({
      op: "replace",
      id: sheet.id,
      path: ["merges"]
    });
    expect(engine.getCell(sheet.id, 0, 0)?.style).toMatchObject({
      fontWeight: "bold",
      backgroundColor: "#d1fae5",
      wrap: true
    });
    expect(engine.getMerge(sheet.id, 1, 1)).toEqual({
      start: { row: 0, col: 0 },
      end: { row: 1, col: 1 }
    });
    expect(engine.getColumnSchema(sheet.id, 0)?.width).toBe(180);
    expect(engine.getRowSchema(sheet.id, 0)?.height).toBe(40);

    engine.unmergeCells({ sheetId: sheet.id, row: 0, col: 0 });
    expect(engine.getMerge(sheet.id, 0, 0)).toBeUndefined();
  });

  it("evaluates formulas without dynamic execution", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 5 });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 1, value: 7 });
    engine.setCellValue({ sheetId: sheet.id, row: 1, col: 0, value: "=A1+B1" });

    expect(engine.getDisplayValue(sheet.id, 1, 0)).toBe("12");
  });

  it("adds and removes sheets while keeping one active sheet", () => {
    const engine = new WorkbookEngine();
    const firstSheet = engine.getActiveSheet();

    const secondSheetId = engine.addSheet({ name: "Sheet2" });

    expect(engine.getSnapshot().sheets).toHaveLength(2);
    expect(engine.getActiveSheet().id).toBe(secondSheetId);

    engine.deleteSheet(secondSheetId);

    expect(engine.getSnapshot().sheets).toHaveLength(1);
    expect(engine.getActiveSheet().id).toBe(firstSheet.id);
  });

  it("shifts cell contents when rows are inserted and deleted", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 2, col: 0, value: "kept" });
    engine.insertRows(sheet.id, 1, 2);

    expect(engine.getDisplayValue(sheet.id, 4, 0)).toBe("kept");

    engine.deleteRows(sheet.id, 1, 2);
    expect(engine.getDisplayValue(sheet.id, 2, 0)).toBe("kept");
  });

  it("shifts cell contents when columns are inserted and deleted", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 2, value: "kept" });
    engine.insertColumns(sheet.id, 1, 2);

    expect(engine.getDisplayValue(sheet.id, 0, 4)).toBe("kept");

    engine.deleteColumns(sheet.id, 1, 2);
    expect(engine.getDisplayValue(sheet.id, 0, 2)).toBe("kept");
  });

  it("serializes and reloads workbook state as json", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "persisted" });
    const snapshot = engine.toJSON();
    const rehydrated = WorkbookEngine.fromJSON(snapshot, new BasicFormulaEngine());

    expect(rehydrated.getDisplayValue(sheet.id, 0, 0)).toBe("persisted");
    expect(rehydrated.getSnapshot().activeSheetId).toBe(snapshot.activeSheetId);
  });

  it("evaluates formulas with references between sheets", () => {
    const engine = new WorkbookEngine(
      {
        data: [{ name: "Revenue" }, { name: "Summary" }]
      },
      new BasicFormulaEngine()
    );
    const [revenueSheet, summarySheet] = engine.getSnapshot().sheets;

    engine.setCellValue({ sheetId: revenueSheet.id, row: 0, col: 0, value: 10 });
    engine.setCellValue({ sheetId: revenueSheet.id, row: 0, col: 1, value: 5 });
    engine.setCellValue({ sheetId: summarySheet.id, row: 0, col: 0, value: "='Revenue'!A1+'Revenue'!B1" });

    expect(engine.getDisplayValue(summarySheet.id, 0, 0)).toBe("15");
  });

  it("recalculates only formulas impacted by a value edit when the engine exposes references", () => {
    const formulaEngine = new TrackingFormulaEngine();
    const engine = new WorkbookEngine({}, formulaEngine);
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 5 });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 1, value: "=A1*2" });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 2, value: "=99" });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 3, value: "=B1+1" });

    formulaEngine.evaluations.splice(0);
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 7 });

    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("14");
    expect(engine.getDisplayValue(sheet.id, 0, 2)).toBe("99");
    expect(engine.getDisplayValue(sheet.id, 0, 3)).toBe("15");
    expect(formulaEngine.evaluations).toContain("Sheet1!0:1");
    expect(formulaEngine.evaluations).toContain("Sheet1!0:3");
    expect(formulaEngine.evaluations).not.toContain("Sheet1!0:2");
  });

  it("marks affected formulas with an error when recalc workload exceeds the configured limit", () => {
    const engine = new WorkbookEngine(
      {
        settings: {
          maxRecalcCells: 1
        }
      },
      new BasicFormulaEngine()
    );
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 2 });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 1, value: "=A1*2" });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 2, value: "=B1+1" });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 3 });

    expect(engine.getCell(sheet.id, 0, 1)?.error).toMatchObject({
      code: "FORMULA_RECALC_LIMIT_EXCEEDED"
    });
    expect(engine.getCell(sheet.id, 0, 2)?.error).toMatchObject({
      code: "FORMULA_RECALC_LIMIT_EXCEEDED"
    });
  });

  it("recalculates only formulas impacted by batch updates when the engine exposes references", () => {
    const formulaEngine = new TrackingFormulaEngine();
    const engine = new WorkbookEngine({}, formulaEngine);
    const sheet = engine.getActiveSheet();

    engine.updateCells({
      sheetId: sheet.id,
      updates: [
        { row: 0, col: 0, value: 3 },
        { row: 0, col: 1, value: "=A1*2" },
        { row: 0, col: 2, value: "=42" },
        { row: 0, col: 3, value: "=B1+1" }
      ]
    });

    formulaEngine.evaluations.splice(0);
    engine.updateCells({
      sheetId: sheet.id,
      updates: [{ row: 0, col: 0, value: 6 }]
    });

    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("12");
    expect(engine.getDisplayValue(sheet.id, 0, 2)).toBe("42");
    expect(engine.getDisplayValue(sheet.id, 0, 3)).toBe("13");
    expect(formulaEngine.evaluations).toContain("Sheet1!0:1");
    expect(formulaEngine.evaluations).toContain("Sheet1!0:3");
    expect(formulaEngine.evaluations).not.toContain("Sheet1!0:2");
  });

  it("skips history and operations for no-op batched cell updates", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.updateCells({
      sheetId: sheet.id,
      updates: [{ row: 0, col: 0, value: 10 }]
    });

    const operations = engine.updateCells({
      sheetId: sheet.id,
      updates: [
        { row: 0, col: 0, value: 10 },
        { row: 0, col: 0, value: 10 }
      ]
    });

    expect(operations).toEqual([]);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("10");
    expect(engine.undo()).toBe(true);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(engine.undo()).toBe(false);
  });

  it("applies keyed cell transactions as a single undoable batch", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 1, col: 1, value: "legacy" });

    const operations = engine.applyCellTransaction({
      sheetId: sheet.id,
      changes: [
        { type: "upsert", key: "0:0", value: 10 },
        { type: "upsert", key: "0:1", value: "=A1*2" },
        { type: "remove", key: "1:1" }
      ]
    });

    expect(operations).toHaveLength(3);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("10");
    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("20");
    expect(engine.getDisplayValue(sheet.id, 1, 1)).toBe("");
    expect(engine.undo()).toBe(true);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("");
    expect(engine.getDisplayValue(sheet.id, 1, 1)).toBe("legacy");
  });

  it("treats redundant keyed cell transactions as no-op", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 10 });

    const operations = engine.applyCellTransaction({
      sheetId: sheet.id,
      changes: [
        { type: "upsert", key: "0:0", value: 10 },
        { type: "remove", key: "9:9" }
      ]
    });

    expect(operations).toEqual([]);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("10");
    expect(engine.undo()).toBe(true);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(engine.undo()).toBe(false);
  });

  it("preserves the active selection during high-frequency updates", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.selectRange({
      sheetId: sheet.id,
      rowStart: 4,
      rowEnd: 5,
      colStart: 2,
      colEnd: 3
    });

    engine.updateCells({
      sheetId: sheet.id,
      updates: [
        { row: 0, col: 0, value: 10 },
        { row: 0, col: 1, value: "=A1*2" }
      ]
    });
    engine.applyCellTransaction({
      sheetId: sheet.id,
      changes: [
        { type: "upsert", key: "1:0", value: 20 },
        { type: "remove", key: "9:9" }
      ]
    });

    expect(engine.getSelection(sheet.id)).toEqual({
      start: { row: 4, col: 2 },
      end: { row: 5, col: 3 }
    });
  });

  it("returns controlled error for circular references across sheets", () => {
    const engine = new WorkbookEngine(
      {
        data: [{ name: "First" }, { name: "Second" }]
      },
      new BasicFormulaEngine()
    );
    const [firstSheet, secondSheet] = engine.getSnapshot().sheets;

    engine.setCellValue({ sheetId: firstSheet.id, row: 0, col: 0, value: "='Second'!A1" });
    engine.setCellValue({ sheetId: secondSheet.id, row: 0, col: 0, value: "='First'!A1" });

    expect(engine.getDisplayValue(firstSheet.id, 0, 0)).toBe("#FORMULA_INVALID");
    expect(engine.getDisplayValue(secondSheet.id, 0, 0)).toBe("#FORMULA_INVALID");
  });

  it("applies serializable operations from another source", () => {
    const source = new WorkbookEngine();
    const target = new WorkbookEngine();
    const sourceSheet = source.getActiveSheet();
    const targetSheet = target.getActiveSheet();
    target.setActiveSheet(targetSheet.id);

    const operations = source.setCellValue({ sheetId: sourceSheet.id, row: 2, col: 2, value: "sync" });
    target.applyOperations(
      operations.map((operation) => ({
        ...operation,
        id: targetSheet.id
      }))
    );

    expect(target.getDisplayValue(targetSheet.id, 2, 2)).toBe("sync");
  });

  it("recalculates dependent formulas when remote operations are applied", () => {
    const source = new WorkbookEngine(
      {
        data: [{ name: "Revenue" }, { name: "Summary" }]
      },
      new BasicFormulaEngine()
    );
    const [sourceRevenue, sourceSummary] = source.getSnapshot().sheets;
    source.setCellValue({ sheetId: sourceSummary.id, row: 0, col: 0, value: "='Revenue'!A1*2" });
    const target = WorkbookEngine.fromJSON(source.toJSON(), new BasicFormulaEngine());
    const [targetRevenue, targetSummary] = target.getSnapshot().sheets;
    const operations = source.setCellValue({ sheetId: sourceRevenue.id, row: 0, col: 0, value: 12 });

    target.applyOperations(
      operations.map((operation) => ({
        ...operation,
        id: targetRevenue.id
      }))
    );

    target.setActiveSheet(targetSummary.id);
    expect(target.getDisplayValue(targetSummary.id, 0, 0)).toBe("24");
  });

  it("rejects invalid edits before confirming the value and preserves the previous content", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.setCellValidation({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      validation: {
        rules: [{ type: "number", min: 10, max: 20 }]
      }
    });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "12" });

    expect(() => {
      engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "abc" });
    }).toThrow(CellValidationError);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("12");
  });

  it("supports custom validators registered through the public API", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.registerValidator("starts-with", ({ value, params }) => {
      const expectedPrefix = String(params?.prefix ?? "");
      return typeof value === "string" && value.startsWith(expectedPrefix)
        ? undefined
        : {
            code: "CORE_VALIDATION_CUSTOM_PREFIX",
            message: `O valor deve começar com ${expectedPrefix}.`,
            ruleType: "custom",
            validator: "starts-with"
          };
    });

    engine.setCellValidation({
      sheetId: sheet.id,
      row: 1,
      col: 0,
      validation: {
        rules: [{ type: "custom", validator: "starts-with", params: { prefix: "INV-" } }]
      }
    });

    expect(engine.getRegisteredValidators()).toContainEqual({ id: "starts-with" });
    expect(engine.validateCellValue({ sheetId: sheet.id, row: 1, col: 0, value: "INV-001" }).valid).toBe(true);
    expect(engine.validateCellValue({ sheetId: sheet.id, row: 1, col: 0, value: "001" })).toMatchObject({
      valid: false,
      issue: {
        validator: "starts-with"
      }
    });
  });

  it("evaluates conditional formatting with deterministic priority and safe formulas", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 12 });
    engine.setConditionalFormattingRules(sheet.id, [
      {
        id: "gt-5",
        type: "greaterThan",
        range: {
          start: { row: 0, col: 0 },
          end: { row: 0, col: 0 }
        },
        value: 5,
        priority: 80,
        style: {
          backgroundColor: "#bfdbfe",
          fontWeight: "bold"
        }
      },
      {
        id: "formula-hot",
        type: "formula",
        range: {
          start: { row: 0, col: 0 },
          end: { row: 0, col: 0 }
        },
        formula: "=A1-10",
        priority: 10,
        style: {
          backgroundColor: "#fecaca"
        }
      }
    ]);

    expect(engine.getConditionalFormattingRules(sheet.id)).toHaveLength(2);
    expect(engine.getConditionalStyle(sheet.id, 0, 0)).toMatchObject({
      backgroundColor: "#fecaca",
      fontWeight: "bold"
    });
  });

  it("supports duplicate highlights and color scales in conditional formatting", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 10 });
    engine.setCellValue({ sheetId: sheet.id, row: 1, col: 0, value: 10 });
    engine.setCellValue({ sheetId: sheet.id, row: 2, col: 0, value: 30 });
    engine.setConditionalFormattingRules(sheet.id, [
      {
        id: "scale-a",
        type: "colorScale",
        range: {
          start: { row: 0, col: 0 },
          end: { row: 2, col: 0 }
        },
        priority: 100,
        minColor: "#000000",
        maxColor: "#ffffff"
      },
      {
        id: "dupe-a",
        type: "duplicates",
        range: {
          start: { row: 0, col: 0 },
          end: { row: 2, col: 0 }
        },
        priority: 20,
        style: {
          textColor: "#166534"
        }
      }
    ]);

    expect(engine.getConditionalStyle(sheet.id, 0, 0)).toMatchObject({
      backgroundColor: "#000000",
      textColor: "#166534"
    });
    expect(engine.getConditionalStyle(sheet.id, 2, 0)).toMatchObject({
      backgroundColor: "#ffffff"
    });
  });

  it("persists frozen panes and hidden rows or columns through the public core API", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.freezeRows(sheet.id, 2);
    engine.freezeColumns(sheet.id, 1);
    engine.setRowsHidden(sheet.id, 3, 4, true);
    engine.setColumnsHidden(sheet.id, 2, 2, true);

    expect(engine.getFrozenPane(sheet.id)).toEqual({ rows: 2, columns: 1 });
    expect(engine.getRowSchema(sheet.id, 3)?.hidden).toBe(true);
    expect(engine.getRowSchema(sheet.id, 4)?.hidden).toBe(true);
    expect(engine.getColumnSchema(sheet.id, 2)?.hidden).toBe(true);

    engine.setRowsHidden(sheet.id, 3, 4, false);
    engine.setColumnsHidden(sheet.id, 2, 2, false);

    expect(engine.getRowSchema(sheet.id, 3)?.hidden).toBeUndefined();
    expect(engine.getColumnSchema(sheet.id, 2)?.hidden).toBeUndefined();
  });

  it("creates and updates embedded chart objects with undoable commands", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const createdEvents: Array<{ chartId: string; sheetId: string }> = [];
    const movedEvents: Array<{ chartId: string; x: number; y: number }> = [];
    const resizedEvents: Array<{ chartId: string; width: number; height: number }> = [];
    const deletedEvents: string[] = [];
    const selectedEvents: string[] = [];
    const unselectedEvents: string[] = [];

    engine.on("chart:created", ({ chartId, sheetId }) => createdEvents.push({ chartId, sheetId }));
    engine.on("chart:moved", ({ chartId, position }) => movedEvents.push({ chartId, x: position.offsetX, y: position.offsetY }));
    engine.on("chart:resized", ({ chartId, position }) =>
      resizedEvents.push({ chartId, width: position.width, height: position.height })
    );
    engine.on("chart:deleted", ({ chartId }) => deletedEvents.push(chartId));
    engine.on("chart:selected", ({ chartId }) => selectedEvents.push(chartId));
    engine.on("chart:unselected", ({ chartId }) => unselectedEvents.push(chartId));

    const createOperations = engine.createChart({
      sheetId: sheet.id,
      chart: {
        type: "line",
        title: "Revenue",
        figure: {
          data: [
            {
              type: "line",
              x: ["Jan", "Fev", "Mar"],
              y: [10, 20, 15]
            }
          ]
        },
        sourceRange: {
          rangeAddress: "A1:B4",
          orientation: "columns",
          firstRowAsHeader: true,
          firstColumnAsLabel: true,
          autoRefresh: true
        },
        position: {
          fromCell: "D2",
          offsetX: 12,
          offsetY: 18,
          width: 360,
          height: 220
        }
      }
    });

    expect(createOperations[0]).toMatchObject({
      op: "add",
      id: sheet.id,
      path: ["charts", 0]
    });
    expect(createdEvents).toHaveLength(1);

    const createdChart = engine.getCharts(sheet.id)[0];
    expect(createdChart?.sourceRange?.chartId).toBe(createdChart?.id);
    expect(createdChart?.sourceRange?.sheetId).toBe(sheet.id);

    if (!createdChart) {
      throw new Error("Expected created chart.");
    }

    engine.moveChart({
      sheetId: sheet.id,
      chartId: createdChart.id,
      position: {
        fromCell: "E3",
        toCell: "K16",
        offsetX: 20,
        offsetY: 32,
        zIndex: 5
      }
    });
    engine.resizeChart({
      sheetId: sheet.id,
      chartId: createdChart.id,
      position: {
        width: 420,
        height: 260,
        toCell: "L18"
      }
    });
    engine.changeChartType({
      sheetId: sheet.id,
      chartId: createdChart.id,
      chartType: "bar"
    });
    engine.changeChartTitle({
      sheetId: sheet.id,
      chartId: createdChart.id,
      title: "Revenue by Month"
    });
    engine.changeChartLegend({
      sheetId: sheet.id,
      chartId: createdChart.id,
      visible: false
    });
    engine.changeChartRange({
      sheetId: sheet.id,
      chartId: createdChart.id,
      sourceRange: {
        rangeAddress: "A1:C8",
        orientation: "columns",
        firstRowAsHeader: true,
        firstColumnAsLabel: true,
        autoRefresh: true
      }
    });

    const updated = engine.getChart(sheet.id, createdChart.id);
    expect(updated).toMatchObject({
      type: "bar",
      title: "Revenue by Month",
      position: {
        fromCell: "E3",
        toCell: "L18",
        offsetX: 20,
        offsetY: 32,
        width: 420,
        height: 260,
        zIndex: 5
      },
      sourceRange: {
        rangeAddress: "A1:C8"
      }
    });
    expect(movedEvents.at(-1)).toMatchObject({
      chartId: createdChart.id,
      x: 20,
      y: 32
    });
    expect(resizedEvents.at(-1)).toMatchObject({
      chartId: createdChart.id,
      width: 420,
      height: 260
    });

    engine.updateChart({
      sheetId: sheet.id,
      chartId: createdChart.id,
      patch: {
        state: {
          ...createdChart.state,
          selected: true
        }
      }
    });
    expect(selectedEvents.at(-1)).toBe(createdChart.id);

    engine.deleteChart({ sheetId: sheet.id, chartId: createdChart.id });
    expect(engine.getCharts(sheet.id)).toHaveLength(0);
    expect(deletedEvents).toContain(createdChart.id);
    expect(unselectedEvents.at(-1)).toBe(createdChart.id);

    expect(engine.undo()).toBe(true);
    expect(engine.getCharts(sheet.id)).toHaveLength(1);
  });

  it("emits range binding updates only for charts affected by changed cells", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const rangeEvents: Array<{ chartId: string; reason: string; range: string }> = [];

    engine.on("chart:rangeChanged", ({ chartId, reason, range }) => {
      rangeEvents.push({ chartId, reason, range: range.rangeAddress });
    });

    engine.createChart({
      sheetId: sheet.id,
      chart: {
        type: "line",
        figure: {
          data: [{ type: "line", x: ["A", "B"], y: [1, 2] }]
        },
        sourceRange: {
          rangeAddress: "A1:B3",
          orientation: "columns",
          firstRowAsHeader: true,
          firstColumnAsLabel: true,
          autoRefresh: true
        },
        position: {
          fromCell: "E2",
          offsetX: 0,
          offsetY: 0,
          width: 300,
          height: 180
        }
      }
    });

    const chart = engine.getCharts(sheet.id)[0];
    expect(chart).toBeDefined();
    expect(rangeEvents.at(-1)).toMatchObject({
      chartId: chart?.id,
      reason: "binding-updated",
      range: "A1:B3"
    });

    rangeEvents.splice(0);
    engine.setCellValue({ sheetId: sheet.id, row: 1, col: 1, value: 99 });
    expect(rangeEvents).toEqual([
      {
        chartId: chart!.id,
        reason: "source-cells-updated",
        range: "A1:B3"
      }
    ]);

    engine.setCellValue({ sheetId: sheet.id, row: 8, col: 8, value: 77 });
    expect(rangeEvents).toHaveLength(1);
  });

  it("emits chart import/export and runtime diagnostic events", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const imported: string[] = [];
    const exported: string[] = [];
    const unsupported: string[] = [];
    const renderStarted: string[] = [];
    const renderFinished: string[] = [];
    const renderSkipped: string[] = [];
    const chartErrors: string[] = [];
    engine.on("chart:imported", ({ chartId }) => imported.push(chartId));
    engine.on("chart:exported", ({ chartId }) => exported.push(chartId));
    engine.on("chart:unsupportedFeature", ({ feature }) => unsupported.push(feature));
    engine.on("chart:renderStarted", ({ chartId }) => renderStarted.push(chartId));
    engine.on("chart:renderFinished", ({ chartId }) => renderFinished.push(chartId));
    engine.on("chart:renderSkipped", ({ reason }) => renderSkipped.push(reason));
    engine.on("chart:error", ({ errorCode }) => chartErrors.push(errorCode));

    const snapshot = engine.toJSON();
    snapshot.sheets[0]!.charts = [
      {
        id: "chart-imported-1",
        sheetId: sheet.id,
        type: "surface3d",
        title: "3D chart",
        sourceRange: {
          chartId: "chart-imported-1",
          sheetId: sheet.id,
          rangeAddress: "A1:B3",
          orientation: "rows",
          firstRowAsHeader: true,
          firstColumnAsLabel: true,
          autoRefresh: true
        },
        figure: {
          data: [],
          layout: {}
        },
        position: {
          fromCell: "C2",
          toCell: "H14",
          offsetX: 0,
          offsetY: 0,
          width: 320,
          height: 220,
          zIndex: 1
        },
        state: {
          selected: false,
          visible: true,
          locked: false
        },
        excelInterop: {
          unsupportedFeatures: ["type:surface3d"]
        }
      }
    ];
    engine.loadFromJSON(snapshot);
    expect(imported).toContain("chart-imported-1");
    expect(unsupported).toContain("type:surface3d");

    const exportedSnapshot = engine.toJSON({ emitChartExportEvents: true });
    expect(exportedSnapshot.sheets[0]?.charts).toHaveLength(1);
    expect(exported).toContain("chart-imported-1");

    engine.reportChartRenderStarted(sheet.id, "chart-imported-1");
    engine.reportChartRenderFinished(sheet.id, "chart-imported-1", 12);
    engine.reportChartRenderSkipped(sheet.id, "chart-imported-1", "outside-viewport");
    engine.reportChartError({
      sheetId: sheet.id,
      chartId: "chart-imported-1",
      errorCode: "TEST_CHART_ERROR",
      message: "Synthetic error"
    });

    expect(renderStarted).toContain("chart-imported-1");
    expect(renderFinished).toContain("chart-imported-1");
    expect(renderSkipped).toContain("outside-viewport");
    expect(chartErrors).toContain("TEST_CHART_ERROR");
  });

  it("emits controlled chart data warnings for invalid range bindings", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const dataInvalidEvents: Array<{ chartId: string; reason: string }> = [];

    engine.on("chart:dataInvalid", ({ chartId, reason }) => {
      dataInvalidEvents.push({ chartId, reason });
    });

    engine.createChart({
      sheetId: sheet.id,
      chart: {
        type: "bar",
        figure: {
          data: [{ type: "bar", x: ["A", "B"], y: [3, 4] }]
        },
        sourceRange: {
          rangeAddress: "A1::B3",
          orientation: "columns",
          firstRowAsHeader: true,
          firstColumnAsLabel: true,
          autoRefresh: true
        },
        position: {
          fromCell: "C2",
          offsetX: 0,
          offsetY: 0,
          width: 280,
          height: 180
        }
      }
    });

    const chart = engine.getCharts(sheet.id)[0];
    if (!chart) {
      throw new Error("Expected chart for invalid range test.");
    }

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 1 });
    expect(dataInvalidEvents.at(-1)).toMatchObject({
      chartId: chart.id
    });
    expect(dataInvalidEvents.at(-1)?.reason).toContain("Invalid chart range");
  });

  it("creates a client-side pivot sheet as a derived view", () => {
    const engine = new WorkbookEngine({
      data: [
        {
          name: "Sales",
          rowCount: 8,
          columnCount: 3,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 10, computedValue: 10 },
            "2:0": { value: "North", computedValue: "North" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 20, computedValue: 20 },
            "3:0": { value: "South", computedValue: "South" },
            "3:1": { value: "Q1", computedValue: "Q1" },
            "3:2": { value: 5, computedValue: 5 },
            "4:0": { value: "South", computedValue: "South" },
            "4:1": { value: "Q2", computedValue: "Q2" },
            "4:2": { value: 15, computedValue: 15 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();

    const pivot = engine.createPivotSheet({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 4, col: 2 }
      },
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum" }]
    });

    expect(pivot.name).toBe("Sales Pivot");
    expect(pivot.cells?.["0:0"]?.value).toBe("Region");
    expect(pivot.cells?.["0:1"]?.value).toBe("Q1");
    expect(pivot.cells?.["0:2"]?.value).toBe("Q2");
    expect(pivot.cells?.["0:3"]?.value).toBe("Total");
    expect(pivot.cells?.["1:0"]?.value).toBe("North");
    expect(pivot.cells?.["1:1"]?.value).toBe(10);
    expect(pivot.cells?.["1:2"]?.value).toBe(20);
    expect(pivot.cells?.["1:3"]?.value).toBe(30);
    expect(pivot.cells?.["2:0"]?.value).toBe("South");
    expect(pivot.cells?.["2:1"]?.value).toBe(5);
    expect(pivot.cells?.["2:2"]?.value).toBe(15);
    expect(pivot.cells?.["2:3"]?.value).toBe(20);
    expect(pivot.cells?.["3:0"]?.value).toBe("Grand Total");
    expect(pivot.cells?.["3:1"]?.value).toBe(15);
    expect(pivot.cells?.["3:2"]?.value).toBe(35);
    expect(pivot.cells?.["3:3"]?.value).toBe(50);
  });

  it("materializes a pivot sheet with subtotals through the workbook engine", () => {
    const engine = new WorkbookEngine({
      data: [
        {
          name: "Tickets",
          rowCount: 8,
          columnCount: 3,
          cells: {
            "0:0": { value: "Team", computedValue: "Team" },
            "0:1": { value: "Status", computedValue: "Status" },
            "0:2": { value: "Count", computedValue: "Count" },
            "1:0": { value: "Eng", computedValue: "Eng" },
            "1:1": { value: "Open", computedValue: "Open" },
            "1:2": { value: 1, computedValue: 1 },
            "2:0": { value: "Eng", computedValue: "Eng" },
            "2:1": { value: "Closed", computedValue: "Closed" },
            "2:2": { value: 1, computedValue: 1 },
            "3:0": { value: "Support", computedValue: "Support" },
            "3:1": { value: "Open", computedValue: "Open" },
            "3:2": { value: 1, computedValue: 1 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();

    const pivotSheetId = engine.addPivotSheet({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      },
      rows: ["Team", "Status"],
      values: [{ field: "Count", aggregate: "sum", as: "Tickets" }],
      includeSubtotals: true
    });

    expect(engine.getActiveSheet().id).toBe(pivotSheetId);
    expect(engine.getDisplayValue(pivotSheetId, 0, 0)).toBe("Team");
    expect(engine.getDisplayValue(pivotSheetId, 0, 1)).toBe("Status");
    expect(engine.getDisplayValue(pivotSheetId, 0, 2)).toBe("Tickets");
    expect(engine.getDisplayValue(pivotSheetId, 1, 0)).toBe("Eng");
    expect(engine.getDisplayValue(pivotSheetId, 1, 1)).toBe("Open");
    expect(engine.getDisplayValue(pivotSheetId, 1, 2)).toBe("1");
    expect(engine.getDisplayValue(pivotSheetId, 2, 0)).toBe("Eng");
    expect(engine.getDisplayValue(pivotSheetId, 2, 1)).toBe("Closed");
    expect(engine.getDisplayValue(pivotSheetId, 2, 2)).toBe("1");
    expect(engine.getDisplayValue(pivotSheetId, 3, 0)).toBe("Eng Total");
    expect(engine.getDisplayValue(pivotSheetId, 3, 2)).toBe("2");
    expect(engine.getDisplayValue(pivotSheetId, 5, 0)).toBe("Support Total");
    expect(engine.getDisplayValue(pivotSheetId, 5, 2)).toBe("1");
    expect(engine.getDisplayValue(pivotSheetId, 6, 0)).toBe("Grand Total");
    expect(engine.getDisplayValue(pivotSheetId, 6, 2)).toBe("3");
  });

  it("infers a pivot config that uses column dimensions when the range has multiple dimensions", () => {
    const engine = new WorkbookEngine({
      data: [
        {
          name: "Sales",
          rowCount: 8,
          columnCount: 3,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 10, computedValue: 10 },
            "2:0": { value: "North", computedValue: "North" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 20, computedValue: 20 },
            "3:0": { value: "South", computedValue: "South" },
            "3:1": { value: "Q1", computedValue: "Q1" },
            "3:2": { value: 5, computedValue: 5 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();

    const inferred = engine.inferPivotSheet({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      }
    });

    expect(inferred).toMatchObject({
      sourceSheetId: sourceSheet.id,
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum" }],
      includeSubtotals: true
    });
  });

  it("allows disabling the pivot module per workbook", () => {
    const engine = new WorkbookEngine({
      pivotModule: false,
      data: [{ name: "Sales" }]
    });
    const sourceSheet = engine.getActiveSheet();
    const input = {
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 1, col: 1 }
      },
      rows: ["Region"],
      values: [{ field: "Sales", aggregate: "sum" as const }]
    };

    expect(() => engine.inferPivotSheet({ sourceSheetId: sourceSheet.id, sourceRange: input.sourceRange })).toThrow(
      "Pivot module is not enabled for this workbook."
    );
    expect(() => engine.createPivotSheet(input)).toThrow("Pivot module is not enabled for this workbook.");
    expect(() => engine.addPivotSheet(input)).toThrow("Pivot module is not enabled for this workbook.");
  });

  it("delegates pivot operations to a configured pivot module", () => {
    const calls: string[] = [];
    const pivotModule: PivotModule = {
      inferPivotSheet(workbook, input) {
        calls.push(`infer:${workbook.id}:${input.sourceSheetId}`);
        return {
          sourceSheetId: input.sourceSheetId,
          sourceRange: input.sourceRange,
          rows: ["Region"],
          values: [{ field: "Sales", aggregate: "count" }],
          sheetName: "Custom Pivot"
        };
      },
      createPivotSheet(workbook, input) {
        calls.push(`create:${workbook.id}:${input.sourceSheetId}`);
        return {
          name: input.sheetName ?? "Custom Pivot",
          rowCount: 1,
          columnCount: 1,
          cells: {
            "0:0": { value: "stub", computedValue: "stub" }
          },
          merges: [],
          columns: {},
          rows: {}
        };
      }
    };
    const engine = new WorkbookEngine({
      pivotModule,
      data: [
        {
          name: "Sales",
          rowCount: 2,
          columnCount: 2,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: 10, computedValue: 10 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();
    const inferred = engine.inferPivotSheet({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 1, col: 1 }
      }
    });
    const pivot = engine.createPivotSheet(inferred);
    const pivotSheetId = engine.addPivotSheet(inferred);

    expect(inferred.sheetName).toBe("Custom Pivot");
    expect(pivot.name).toBe("Custom Pivot");
    expect(engine.getActiveSheet().id).toBe(pivotSheetId);
    expect(calls).toEqual([
      `infer:${engine.getSnapshot().id}:${sourceSheet.id}`,
      `create:${engine.getSnapshot().id}:${sourceSheet.id}`,
      `create:${engine.getSnapshot().id}:${sourceSheet.id}`
    ]);
  });

  it("builds pivot asynchronously in chunks through the workbook engine", async () => {
    const engine = new WorkbookEngine({
      data: [
        {
          name: "Sales",
          rowCount: 5,
          columnCount: 3,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 10, computedValue: 10 },
            "2:0": { value: "North", computedValue: "North" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 20, computedValue: 20 },
            "3:0": { value: "South", computedValue: "South" },
            "3:1": { value: "Q1", computedValue: "Q1" },
            "3:2": { value: 5, computedValue: 5 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();
    let yields = 0;

    const pivot = await engine.createPivotSheetAsync(
      {
        sourceSheetId: sourceSheet.id,
        sourceRange: {
          start: { row: 0, col: 0 },
          end: { row: 3, col: 2 }
        },
        rows: ["Region"],
        columns: ["Quarter"],
        values: [{ field: "Sales", aggregate: "sum" }]
      },
      {
        chunkSize: 1,
        yieldControl: async () => {
          yields += 1;
          await Promise.resolve();
        }
      }
    );

    expect(yields).toBeGreaterThan(0);
    expect(pivot.cells?.["0:3"]?.value).toBe("Total");
    const pivotSheetId = await engine.addPivotSheetAsync({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      },
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum" }]
    });
    expect(engine.getActiveSheet().id).toBe(pivotSheetId);
  });

  it("supports progress and cancellation during async pivot builds", async () => {
    const engine = new WorkbookEngine({
      data: [
        {
          name: "Sales",
          rowCount: 5,
          columnCount: 3,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 10, computedValue: 10 },
            "2:0": { value: "North", computedValue: "North" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 20, computedValue: 20 },
            "3:0": { value: "South", computedValue: "South" },
            "3:1": { value: "Q1", computedValue: "Q1" },
            "3:2": { value: 5, computedValue: 5 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();
    const controller = new AbortController();
    const progress: PivotBuildProgress[] = [];

    await expect(
      engine.createPivotSheetAsync(
        {
          sourceSheetId: sourceSheet.id,
          sourceRange: {
            start: { row: 0, col: 0 },
            end: { row: 3, col: 2 }
          },
          rows: ["Region"],
          columns: ["Quarter"],
          values: [{ field: "Sales", aggregate: "sum" }]
        },
        {
          chunkSize: 1,
          signal: controller.signal,
          yieldControl: async () => {},
          onProgress: (step) => {
            progress.push(step);
            if (step.phase === "aggregate" && step.completed === 1) {
              controller.abort();
            }
          }
        }
      )
    ).rejects.toMatchObject({ name: "AbortError" });

    expect(progress.some((step) => step.phase === "aggregate" && step.completed === 1)).toBe(true);
    expect(progress.some((step) => step.phase === "materialize")).toBe(false);
  });

  it("rejects oversized client-side pivot sources with a controlled error", async () => {
    const engine = new WorkbookEngine({
      settings: {
        maxPivotSourceRows: 2
      },
      data: [
        {
          name: "Sales",
          rowCount: 6,
          columnCount: 3,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 10, computedValue: 10 },
            "2:0": { value: "North", computedValue: "North" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 20, computedValue: 20 },
            "3:0": { value: "South", computedValue: "South" },
            "3:1": { value: "Q1", computedValue: "Q1" },
            "3:2": { value: 5, computedValue: 5 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();
    const input = {
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      },
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum" as const }]
    };

    expect(() => engine.createPivotSheet(input)).toThrow(/configured limit of 2 rows/i);
    await expect(engine.createPivotSheetAsync(input)).rejects.toMatchObject({
      name: "CORE_PIVOT_CLIENT_ROW_LIMIT_EXCEEDED"
    });
  });

  it("persists pivot definitions on the sheet and refreshes the derived view in place", async () => {
    const engine = new WorkbookEngine({
      data: [
        {
          name: "Sales",
          rowCount: 5,
          columnCount: 3,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 10, computedValue: 10 },
            "2:0": { value: "North", computedValue: "North" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 20, computedValue: 20 },
            "3:0": { value: "South", computedValue: "South" },
            "3:1": { value: "Q1", computedValue: "Q1" },
            "3:2": { value: 5, computedValue: 5 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();
    const pivotSheetId = engine.addPivotSheet({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      },
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum" }],
      sheetName: "Sales Pivot"
    });

    expect(engine.getPivotSheetDefinition(pivotSheetId)).toMatchObject({
      sourceSheetId: sourceSheet.id,
      rows: ["Region"],
      columns: ["Quarter"]
    });
    expect(engine.getDisplayValue(pivotSheetId, 1, 1)).toBe("10");

    engine.setCellValue({ sheetId: sourceSheet.id, row: 1, col: 2, value: 15 });
    await engine.refreshPivotSheet(pivotSheetId);

    expect(engine.getActiveSheet().id).toBe(pivotSheetId);
    expect(engine.getDisplayValue(pivotSheetId, 1, 1)).toBe("15");
    expect(engine.getPivotSheetDefinition(pivotSheetId)?.sheetName).toBe("Sales Pivot");
  });

  it("auto-refreshes persisted pivot views and marks manual pivots as stale", async () => {
    const engine = new WorkbookEngine({
      data: [
        {
          name: "Sales",
          rowCount: 5,
          columnCount: 3,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 10, computedValue: 10 },
            "2:0": { value: "North", computedValue: "North" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 20, computedValue: 20 },
            "3:0": { value: "South", computedValue: "South" },
            "3:1": { value: "Q1", computedValue: "Q1" },
            "3:2": { value: 5, computedValue: 5 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();
    const pivotSheetId = engine.addPivotSheet({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      },
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum" }],
      sheetName: "Sales Pivot"
    });

    engine.setCellValue({ sheetId: sourceSheet.id, row: 1, col: 2, value: 15 });

    await waitFor(
      () => engine.getDisplayValue(pivotSheetId, 1, 1) === "15",
      "Timed out waiting for derived pivot auto-refresh."
    );

    expect(engine.getPivotSheetViewDefinition(pivotSheetId)).toMatchObject({
      autoRefresh: true,
      stale: false,
      refreshStatus: "idle",
      result: {
        executionMode: "client",
        remote: false
      }
    });

    engine.setPivotSheetAutoRefresh(pivotSheetId, false);
    engine.setCellValue({ sheetId: sourceSheet.id, row: 1, col: 2, value: 18 });
    await Promise.resolve();

    expect(engine.getDisplayValue(pivotSheetId, 1, 1)).toBe("15");
    expect(engine.getPivotSheetViewDefinition(pivotSheetId)).toMatchObject({
      autoRefresh: false,
      stale: true,
      refreshStatus: "idle"
    });
  });

  it("updates a persisted pivot view in place while preserving its refresh policy", async () => {
    const engine = new WorkbookEngine({
      data: [
        {
          name: "Sales",
          rowCount: 5,
          columnCount: 3,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 10, computedValue: 10 },
            "2:0": { value: "North", computedValue: "North" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 20, computedValue: 20 },
            "3:0": { value: "South", computedValue: "South" },
            "3:1": { value: "Q1", computedValue: "Q1" },
            "3:2": { value: 5, computedValue: 5 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();
    const pivotSheetId = engine.addPivotSheet({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      },
      rows: ["Region"],
      values: [{ field: "Sales", aggregate: "sum" }],
      sheetName: "Sales Pivot"
    });

    engine.setPivotSheetAutoRefresh(pivotSheetId, false);
    await engine.updatePivotSheet(pivotSheetId, {
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      },
      rows: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum", as: "Revenue" }],
      sheetName: "Sales by Quarter"
    });

    expect(engine.getDisplayValue(pivotSheetId, 0, 0)).toBe("Quarter");
    expect(engine.getDisplayValue(pivotSheetId, 0, 1)).toBe("Revenue");
    expect(engine.getPivotSheetDefinition(pivotSheetId)).toMatchObject({
      rows: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum", as: "Revenue" }],
      sheetName: "Sales by Quarter"
    });
    expect(engine.getPivotSheetViewDefinition(pivotSheetId)).toMatchObject({
      autoRefresh: false,
      stale: false,
      refreshStatus: "idle"
    });
  });

  it("materializes pivots through a server-side row model when requested", async () => {
    const engine = new WorkbookEngine({
      pivotModule: false,
      data: [{ name: "Remote Sales", rowCount: 1000, columnCount: 5, cells: {}, merges: [], columns: {}, rows: {} }]
    });
    const sourceSheet = engine.getActiveSheet();
    const getRows = vi.fn(async (request: {
      requestKind?: string;
      groupKeys?: string[];
      pivotModel?: Array<{ field: string }>;
      aggregateModel?: Array<{ field: string; function: string; as?: string }>;
    }) => {
      if (request.requestKind === "pivotSheet") {
        return {
          rows: [],
          pivotSheet: {
            name: "Remote Sales Pivot",
            rowCount: 2,
            columnCount: 2,
            cells: {
              "0:0": { value: "Region", computedValue: "Region" },
              "0:1": { value: "Revenue", computedValue: "Revenue" },
              "1:0": { value: "North", computedValue: "North" },
              "1:1": { value: 42, computedValue: 42 }
            },
            merges: [],
            columns: {},
            rows: {}
          }
        };
      }

      return {
        totalRows: 1000,
        rows: [{ index: 0 }]
      };
    });

    engine.setRowModel(
      sourceSheet.id,
      new ServerSideRowModel({
        rowCount: "unknown",
        dataSource: {
          getRows
        }
      })
    );

    const pivot = await engine.createPivotSheetAsync({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 999, col: 4 }
      },
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Revenue", aggregate: "sum", as: "Revenue" }],
      executionMode: "server"
    });

    expect(pivot.name).toBe("Remote Sales Pivot");
    expect(pivot.cells?.["1:1"]?.value).toBe(42);
    expect(getRows).toHaveBeenCalledWith(
      expect.objectContaining({
        requestKind: "pivotSheet",
        groupKeys: ["Region"],
        pivotModel: [{ field: "Quarter" }],
        aggregateModel: [{ field: "Revenue", function: "sum", as: "Revenue" }]
      }),
      expect.any(Object)
    );
  });
});