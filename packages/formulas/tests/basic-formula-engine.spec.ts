import { describe, expect, it, vi } from "vitest";
import { BasicFormulaEngine } from "../src/index";

describe("BasicFormulaEngine", () => {
  const engine = new BasicFormulaEngine();

  it("evaluates arithmetic expressions", () => {
    const result = engine.evaluate("=2+3*4", {
      currentCell: { row: 0, col: 0 },
      currentSheetId: "sheet-1",
      currentSheetName: "Sheet1",
      getCell: () => undefined,
      evaluateCell: () => ({ value: null })
    });

    expect(result.value).toBe(14);
  });

  it("evaluates range functions", () => {
    const values = new Map([
      ["0:0", 1],
      ["1:0", 2],
      ["2:0", 3]
    ]);

    const result = engine.evaluate("=SUM(A1:A3)", {
      currentCell: { row: 3, col: 0 },
      currentSheetId: "sheet-1",
      currentSheetName: "Sheet1",
      getCell: () => undefined,
      evaluateCell: (row, col) => ({ value: values.get(`${row}:${col}`) ?? null })
    });

    expect(result.value).toBe(6);
  });

  it("evaluates cross-sheet references and additional functions", () => {
    const values = new Map([
      ["Revenue:0:0", 12.345],
      ["Revenue:0:1", 20],
      ["Summary:0:0", 1],
      ["Summary:0:1", 0]
    ]);

    const result = engine.evaluate("=ROUND('Revenue'!A1,2)+IF(Summary!A1,ABS(-2),0)+COUNT('Revenue'!A1:B1)", {
      currentCell: { row: 0, col: 0 },
      currentSheetId: "summary-id",
      currentSheetName: "Summary",
      getCell: () => undefined,
      evaluateCell: (row, col, _trail, sheetRef) => ({
        value: values.get(`${sheetRef ?? "Summary"}:${row}:${col}`) ?? null
      })
    });

    expect(result.value).toBe(16.35);
  });

  it("accepts absolute references without breaking evaluation", () => {
    const values = new Map([
      ["0:0", 2],
      ["1:0", 3]
    ]);

    const result = engine.evaluate("=$A$1+$A2", {
      currentCell: { row: 1, col: 1 },
      currentSheetId: "sheet-1",
      currentSheetName: "Sheet1",
      getCell: () => undefined,
      evaluateCell: (row, col) => ({ value: values.get(`${row}:${col}`) ?? null })
    });

    expect(result.value).toBe(5);
  });

  it("collects direct, range and cross-sheet references without duplication", () => {
    expect(engine.collectReferences?.("=SUM(A1:A2)+'Revenue'!B3+$A$1")).toEqual([
      { row: 0, col: 0 },
      { row: 1, col: 0 },
      { row: 2, col: 1, sheetRef: "Revenue" }
    ]);
  });

  it("rejects unsupported tokens instead of evaluating dynamic expressions", () => {
    const result = engine.evaluate("=SUM(A1)+window.alert(1)", {
      currentCell: { row: 0, col: 0 },
      currentSheetId: "sheet-1",
      currentSheetName: "Sheet1",
      getCell: () => undefined,
      evaluateCell: () => ({ value: null })
    });

    expect(result).toMatchObject({
      value: null,
      error: {
        code: "FORMULA_PARSE_INVALID"
      }
    });
  });

  it("rejects unsupported functions explicitly", () => {
    const result = engine.evaluate("=HYPERLINK(1,2)", {
      currentCell: { row: 0, col: 0 },
      currentSheetId: "sheet-1",
      currentSheetName: "Sheet1",
      getCell: () => undefined,
      evaluateCell: () => ({ value: null })
    });

    expect(result).toMatchObject({
      value: null,
      error: {
        code: "FORMULA_FUNCTION_UNSUPPORTED"
      }
    });
  });

  it("continues evaluating formulas when eval and Function are blocked by CSP-like restrictions", () => {
    const originalEval = globalThis.eval;
    const originalFunction = globalThis.Function;
    const evalSpy = vi.fn(() => {
      throw new Error("eval should not run");
    }) as unknown as typeof globalThis.eval;
    const blockedFunction = function Function(): never {
      throw new Error("Function constructor should not run");
    } as unknown as FunctionConstructor;

    globalThis.eval = evalSpy;
    globalThis.Function = blockedFunction;

    try {
      const result = engine.evaluate("=SUM(1,2,3)", {
        currentCell: { row: 0, col: 0 },
        currentSheetId: "sheet-1",
        currentSheetName: "Sheet1",
        getCell: () => undefined,
        evaluateCell: () => ({ value: null })
      });

      expect(result.value).toBe(6);
      expect(evalSpy).not.toHaveBeenCalled();
    } finally {
      globalThis.eval = originalEval;
      globalThis.Function = originalFunction;
    }
  });

  it("rejects hostile payloads cleanly even when global execution primitives are unavailable", () => {
    const originalEval = globalThis.eval;
    const originalFunction = globalThis.Function;
    globalThis.eval = (() => {
      throw new Error("eval should not run");
    }) as typeof globalThis.eval;
    globalThis.Function = function Function(): never {
      throw new Error("Function constructor should not run");
    } as unknown as FunctionConstructor;

    try {
      const result = engine.evaluate("=SUM(A1)+globalThis.process.exit(1)", {
        currentCell: { row: 0, col: 0 },
        currentSheetId: "sheet-1",
        currentSheetName: "Sheet1",
        getCell: () => undefined,
        evaluateCell: () => ({ value: 1 })
      });

      expect(result).toMatchObject({
        value: null,
        error: {
          code: "FORMULA_PARSE_INVALID"
        }
      });
    } finally {
      globalThis.eval = originalEval;
      globalThis.Function = originalFunction;
    }
  });
});