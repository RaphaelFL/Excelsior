import { describe, expect, it } from "vitest";
import { WorkbookEngine } from "@excelsior/core";

describe("pivot build performance regression", () => {
  it("builds a large pivot asynchronously with observable chunk progress", async () => {
    const cells: Record<string, { value: string | number; computedValue: string | number }> = {
      "0:0": { value: "Region", computedValue: "Region" },
      "0:1": { value: "Quarter", computedValue: "Quarter" },
      "0:2": { value: "Revenue", computedValue: "Revenue" }
    };
    const regions = ["North", "South", "East", "West"];
    const quarters = ["Q1", "Q2", "Q3", "Q4"];
    const rowCount = 4001;

    for (let row = 1; row < rowCount; row += 1) {
      cells[`${row}:0`] = { value: regions[row % regions.length]!, computedValue: regions[row % regions.length]! };
      cells[`${row}:1`] = { value: quarters[row % quarters.length]!, computedValue: quarters[row % quarters.length]! };
      cells[`${row}:2`] = { value: row, computedValue: row };
    }

    const engine = new WorkbookEngine({
      data: [
        {
          name: "Sales",
          rowCount,
          columnCount: 3,
          cells,
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();
    let aggregateSteps = 0;
    let yieldCount = 0;

    const pivot = await engine.createPivotSheetAsync(
      {
        sourceSheetId: sourceSheet.id,
        sourceRange: {
          start: { row: 0, col: 0 },
          end: { row: rowCount - 1, col: 2 }
        },
        rows: ["Region"],
        columns: ["Quarter"],
        values: [{ field: "Revenue", aggregate: "sum", as: "Revenue" }]
      },
      {
        chunkSize: 250,
        yieldControl: async () => {
          yieldCount += 1;
          await Promise.resolve();
        },
        onProgress: (progress) => {
          if (progress.phase === "aggregate") {
            aggregateSteps += 1;
          }
        }
      }
    );

    expect(aggregateSteps).toBeGreaterThan(5);
    expect(yieldCount).toBeGreaterThan(5);
    expect(pivot.cells?.["0:0"]?.value).toBe("Region");
    expect(pivot.cells?.["0:5"]?.value).toBe("Total");
  });
});
