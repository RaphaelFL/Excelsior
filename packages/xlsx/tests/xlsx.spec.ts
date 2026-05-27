import { describe, expect, it } from "vitest";
import { SpreadsheetOperationError } from "@excelsior/core";
import { exportTableToXlsx, exportWorkbookToXlsx, importTableFromXlsx, importWorkbookFromXlsx } from "../src/index";

describe("xlsx adapter", () => {
  it("roundtrips workbook data including formulas, styles, merges and column sizing", async () => {
    const bytes = await exportWorkbookToXlsx({
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
          name: "Revenue",
          rowCount: 3,
          columnCount: 3,
          selection: {
            start: { row: 0, col: 0 },
            end: { row: 0, col: 0 }
          },
          merges: [
            {
              start: { row: 0, col: 0 },
              end: { row: 1, col: 1 }
            }
          ],
          columns: {
            0: { width: 180 }
          },
          rows: {
            0: { height: 40 }
          },
          cells: {
            "0:0": {
              value: 10,
              computedValue: 10,
              style: {
                backgroundColor: "#ffeecc",
                fontWeight: "bold",
                align: "center"
              }
            },
            "0:1": { value: 20, computedValue: 20 },
            "2:0": { value: "=A1+B1", formula: "=A1+B1", computedValue: 30 }
          }
        },
        {
          id: "sheet-2",
          name: "Summary",
          rowCount: 1,
          columnCount: 1,
          selection: {
            start: { row: 0, col: 0 },
            end: { row: 0, col: 0 }
          },
          merges: [],
          columns: {},
          rows: {},
          cells: {
            "0:0": { value: "ready", computedValue: "ready" }
          }
        }
      ]
    });

    const imported = await importWorkbookFromXlsx(bytes);

    expect(imported.sheets[0]?.name).toBe("Revenue");
    expect(imported.sheets[1]?.name).toBe("Summary");
    expect(imported.sheets[0]?.cells["0:0"]?.value).toBe(10);
    expect(imported.sheets[0]?.cells["0:0"]?.style).toMatchObject({
      backgroundColor: "#FFEECC",
      fontWeight: "bold",
      align: "center"
    });
    expect(imported.sheets[0]?.cells["2:0"]?.formula).toBe("=A1+B1");
    expect(imported.sheets[0]?.merges).toEqual([
      {
        start: { row: 0, col: 0 },
        end: { row: 1, col: 1 }
      }
    ]);
    expect(imported.sheets[0]?.columns[0]?.width).toBe(180);
    expect(imported.sheets[0]?.rows[0]?.height).toBe(40);
    expect(imported.sheets[1]?.cells["0:0"]?.value).toBe("ready");
  });

  it("imports tabular rows through a column schema", async () => {
    const bytes = await exportTableToXlsx(
      [
        { name: "Ada", score: 99 },
        { name: "Linus", score: 87 }
      ],
      {
        sheet: "People",
        columns: [
          {
            width: 24,
            header: { value: "Name", fontWeight: "bold" },
            cell: (row) => ({ value: row.name })
          },
          {
            width: 12,
            header: { value: "Score", fontWeight: "bold" },
            cell: (row) => ({ value: row.score, type: Number })
          }
        ]
      }
    );

    const result = await importTableFromXlsx<{ name: string; score: number }, "Name" | "Score">(bytes, {
      sheet: "People",
      schema: {
        name: {
          column: "Name",
          type: String,
          required: true
        },
        score: {
          column: "Score",
          type: Number,
          required: true
        }
      }
    });

    expect(result.errors).toBeUndefined();
    expect(result.objects).toEqual([
      { name: "Ada", score: 99 },
      { name: "Linus", score: 87 }
    ]);
  });

  it("returns a typed wrapper error when the requested sheet is missing", async () => {
    const bytes = await exportTableToXlsx(
      [{ name: "Ada" }],
      {
        sheet: "People",
        columns: [
          {
            width: 24,
            header: { value: "Name", fontWeight: "bold" },
            cell: (row) => ({ value: row.name })
          }
        ]
      }
    );

    await expect(
      importTableFromXlsx<{ name: string }, "Name">(bytes, {
        sheet: "Missing",
        schema: {
          name: {
            column: "Name",
            type: String,
            required: true
          }
        }
      })
    ).rejects.toMatchObject({
      name: "XLSX_SHEET_NOT_FOUND",
      details: {
        code: "XLSX_SHEET_NOT_FOUND",
        area: "wrapper"
      }
    });
  });
});