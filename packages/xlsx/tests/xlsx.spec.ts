import { describe, expect, it } from "vitest";
import { BasicFormulaEngine } from "@excelsior/formulas";
import { SpreadsheetOperationError, WorkbookEngine } from "@excelsior/core";
import { packZip, unpackZip } from "../src/native-zip";
import { createXmlDocument } from "../src/xml-tree";
import {
  exportTableToXlsx,
  exportWorkbookEngineToXlsx,
  exportWorkbookToXlsx,
  importTableFromXlsx,
  importWorkbookFromXlsx,
  importWorkbookIntoEngineFromXlsx
} from "../src/index";

describe("xlsx adapter", () => {
  it("rejects unsafe ZIP paths on write and read", async () => {
    expect(() => packZip({ "../escape.xml": "unsafe" })).toThrow("Unsafe ZIP entry path");

    const archive = packZip({ "safe": "content" });
    const mutated = new Uint8Array(archive);
    for (let index = 0; index <= mutated.length - 4; index += 1) {
      if (String.fromCharCode(...mutated.subarray(index, index + 4)) === "safe") {
        mutated.set(Buffer.from("../x"), index);
      }
    }
    await expect(unpackZip(mutated)).rejects.toThrow("Unsafe ZIP entry path");
  });

  it("rejects XML documents with entity declarations", () => {
    expect(() => createXmlDocument('<!DOCTYPE x [<!ENTITY payload "unsafe">]><x>&payload;</x>')).toThrow(
      "Unsafe or oversized XML document"
    );
  });

  it("rejects ZIP entries that announce zip-bomb-sized output", async () => {
    const archive = packZip({ "safe.xml": "content" });
    const mutated = new Uint8Array(archive);
    const centralOffset = mutated.findIndex((byte, index) =>
      byte === 0x50 && mutated[index + 1] === 0x4b && mutated[index + 2] === 0x01 && mutated[index + 3] === 0x02
    );
    expect(centralOffset).toBeGreaterThanOrEqual(0);
    new DataView(mutated.buffer, mutated.byteOffset).setUint32(centralOffset + 24, 32 * 1024 * 1024 + 1, true);
    await expect(unpackZip(mutated)).rejects.toThrow("Unsupported or oversized ZIP entry");
  });

  it("exports and imports native OOXML data validation without private metadata", async () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 5 });
    engine.setCellValidation({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      validation: { rules: [{ type: "number", min: 1, max: 10, message: "Use 1 a 10", severity: "warning" }] }
    });

    const bytes = await exportWorkbookToXlsx(engine.getSnapshot());
    const files = await unpackZip(bytes);
    const sheetXml = new TextDecoder().decode(files["xl/worksheets/sheet1.xml"]!);
    expect(sheetXml).toContain("<dataValidations count=\"1\">");
    expect(sheetXml).toContain("errorStyle=\"warning\"");
    delete files["customXml/excelsior.xml"];

    const imported = await importWorkbookFromXlsx(packZip(files));
    expect(imported.sheets[0]?.cells["0:0"]?.validation).toEqual({
      rules: [{ type: "number", min: 1, max: 10, message: "Use 1 a 10", severity: "warning" }]
    });
  });

  it("exports and imports native OOXML comments without private metadata", async () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "Revisar" });
    engine.createCellComment({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      comment: { content: "Comentário compatível", author: { id: "ana", name: "Ana" } }
    });

    const bytes = await exportWorkbookToXlsx(engine.getSnapshot());
    const files = await unpackZip(bytes);
    expect(new TextDecoder().decode(files["xl/comments1.xml"]!)).toContain("Comentário compatível");
    expect(files["xl/drawings/vmlDrawing1.vml"]).toBeDefined();
    delete files["customXml/excelsior.xml"];

    const imported = await importWorkbookFromXlsx(packZip(files));
    expect(imported.sheets[0]?.cells["0:0"]?.comments?.[0]).toMatchObject({
      content: "Comentário compatível",
      author: { name: "Ana" },
      resolved: false,
      replies: []
    });
  });

  it("roundtrips notes, comment threads and validation metadata", async () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "5" });
    engine.setCellNote({ sheetId: sheet.id, row: 0, col: 0, note: "Confira este valor" });
    engine.setCellValidation({ sheetId: sheet.id, row: 0, col: 0, validation: { rules: [{ type: "number", min: 10, severity: "warning" }] } });
    engine.createCellComment({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      comment: { content: "Revisar", author: { id: "user-1", name: "Ana" } }
    });

    const bytes = await exportWorkbookToXlsx(engine.getSnapshot());
    const imported = await importWorkbookFromXlsx(bytes);
    expect(imported.sheets[0]?.cells["0:0"]).toMatchObject({
      note: "Confira este valor",
      validation: { rules: [{ type: "number", min: 10, severity: "warning" }] },
      comments: [{ content: "Revisar", author: { id: "user-1", name: "Ana" }, resolved: false }]
    });
  });

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

  it("roundtrips worksheet charts with source range and anchor metadata", async () => {
    const workbook = {
      id: "workbook-charts",
      activeSheetId: "sheet-1",
      metadata: {},
      settings: {
        maxRows: 1000,
        maxColumns: 100,
        maxCellLength: 5000,
        maxFormulaLength: 2048,
        maxPasteCells: 10000,
        rowHeight: 28,
        columnWidth: 120,
        viewportBuffer: 6,
        maxHistorySize: 100,
        enableFormulas: true,
        clipboardPolicy: "text-only" as const
      },
      sheets: [
        {
          id: "sheet-1",
          name: "Dashboard",
          rowCount: 12,
          columnCount: 8,
          selection: {
            start: { row: 0, col: 0 },
            end: { row: 0, col: 0 }
          },
          merges: [],
          columns: {},
          rows: {},
          cells: {
            "0:0": { value: "Month", computedValue: "Month" },
            "0:1": { value: "North", computedValue: "North" },
            "0:2": { value: "South", computedValue: "South" },
            "1:0": { value: "Jan", computedValue: "Jan" },
            "1:1": { value: 10, computedValue: 10 },
            "1:2": { value: 8, computedValue: 8 },
            "2:0": { value: "Feb", computedValue: "Feb" },
            "2:1": { value: 14, computedValue: 14 },
            "2:2": { value: 11, computedValue: 11 },
            "3:0": { value: "Mar", computedValue: "Mar" },
            "3:1": { value: 20, computedValue: 20 },
            "3:2": { value: 13, computedValue: 13 },
            "0:3": { value: "X", computedValue: "X" },
            "0:4": { value: "Y", computedValue: "Y" },
            "1:3": { value: 1, computedValue: 1 },
            "1:4": { value: 4, computedValue: 4 },
            "2:3": { value: 2, computedValue: 2 },
            "2:4": { value: 6, computedValue: 6 },
            "3:3": { value: 3, computedValue: 3 },
            "3:4": { value: 9, computedValue: 9 }
          },
          charts: [
            {
              id: "chart-line-1",
              sheetId: "sheet-1",
              type: "line" as const,
              title: "Revenue trend",
              sourceRange: {
                chartId: "chart-line-1",
                sheetId: "sheet-1",
                rangeAddress: "A1:C4",
                orientation: "rows" as const,
                firstRowAsHeader: true,
                firstColumnAsLabel: true,
                autoRefresh: true
              },
              figure: {
                data: [],
                layout: {
                  title: "Revenue trend",
                  xAxis: {
                    title: "Month",
                    visible: true,
                    type: "date"
                  },
                  yAxis: {
                    title: "Revenue",
                    visible: false,
                    type: "log"
                  }
                }
              },
              position: {
                fromCell: "E2",
                toCell: "J14",
                offsetX: 12,
                offsetY: 10,
                width: 500,
                height: 280,
                zIndex: 1
              },
              state: {
                selected: false,
                visible: true,
                locked: false
              },
              excelInterop: {}
            },
            {
              id: "chart-pie-1",
              sheetId: "sheet-1",
              type: "pie" as const,
              title: "Regional split",
              sourceRange: {
                chartId: "chart-pie-1",
                sheetId: "sheet-1",
                rangeAddress: "A1:B4",
                orientation: "rows" as const,
                firstRowAsHeader: true,
                firstColumnAsLabel: true,
                autoRefresh: true
              },
              figure: {
                data: [],
                layout: {
                  title: "Regional split"
                }
              },
              position: {
                fromCell: "A7",
                toCell: "E16",
                offsetX: 6,
                offsetY: 4,
                width: 360,
                height: 240,
                zIndex: 2
              },
              state: {
                selected: false,
                visible: true,
                locked: false
              },
              excelInterop: {}
            },
            {
              id: "chart-scatter-1",
              sheetId: "sheet-1",
              type: "scatter" as const,
              title: "Correlation",
              sourceRange: {
                chartId: "chart-scatter-1",
                sheetId: "sheet-1",
                rangeAddress: "D1:E4",
                orientation: "rows" as const,
                firstRowAsHeader: true,
                firstColumnAsLabel: true,
                autoRefresh: true
              },
              figure: {
                data: [{ type: "scatter", name: "Corr XY", x: [1, 2, 3], y: [4, 6, 9] }],
                layout: {
                  title: "Correlation",
                  xAxis: { title: "X", type: "linear" },
                  yAxis: { title: "Y", type: "linear" }
                }
              },
              position: {
                fromCell: "F7",
                toCell: "L16",
                offsetX: 4,
                offsetY: 2,
                width: 380,
                height: 250,
                zIndex: 3
              },
              state: {
                selected: false,
                visible: true,
                locked: false
              },
              excelInterop: {}
            }
          ]
        }
      ]
    };

    const bytes = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(bytes);
    const importedSheet = imported.sheets[0];

    expect(importedSheet?.charts).toHaveLength(3);
    expect(importedSheet?.charts?.map((chart) => chart.type)).toEqual(["line", "pie", "scatter"]);
    expect(importedSheet?.charts?.map((chart) => chart.title)).toEqual(["Revenue trend", "Regional split", "Correlation"]);
    expect(importedSheet?.charts?.[0]?.sourceRange?.rangeAddress).toBe("A1:C4");
    expect(importedSheet?.charts?.[1]?.sourceRange?.rangeAddress).toBe("A1:B4");
    expect(importedSheet?.charts?.[0]?.position.fromCell).toBe("E2");
    expect(importedSheet?.charts?.[1]?.position.fromCell).toBe("A7");
    expect(importedSheet?.charts?.[0]?.position.width).toBeGreaterThan(0);
    expect(importedSheet?.charts?.[0]?.position.height).toBeGreaterThan(0);
    expect(Array.isArray(importedSheet?.charts?.[0]?.figure.data)).toBe(true);
    expect(Array.isArray(importedSheet?.charts?.[1]?.figure.data)).toBe(true);
    const firstLayout = importedSheet?.charts?.[0]?.figure.layout as
      | { xAxis?: { title?: string; visible?: boolean }; yAxis?: { title?: string; visible?: boolean } }
      | undefined;
    expect(firstLayout?.xAxis?.title).toBe("Month");
    expect(firstLayout?.yAxis?.title).toBe("Revenue");
    expect(firstLayout?.yAxis?.visible).toBe(false);
    expect((firstLayout?.xAxis as { type?: string } | undefined)?.type).toBe("date");
    expect((firstLayout?.yAxis as { type?: string } | undefined)?.type).toBe("log");
    const scatterTraceName = (importedSheet?.charts?.[2]?.figure.data?.[0] as { name?: string } | undefined)?.name;
    expect(scatterTraceName).toBe("Corr XY");
  });

  it("emits chart interop callbacks during workbook export/import", async () => {
    const workbook = {
      id: "workbook-callbacks",
      activeSheetId: "sheet-1",
      metadata: {},
      settings: {
        maxRows: 1000,
        maxColumns: 100,
        maxCellLength: 5000,
        maxFormulaLength: 2048,
        maxPasteCells: 10000,
        rowHeight: 28,
        columnWidth: 120,
        viewportBuffer: 6,
        maxHistorySize: 100,
        enableFormulas: true,
        clipboardPolicy: "text-only" as const
      },
      sheets: [
        {
          id: "sheet-1",
          name: "Dash",
          rowCount: 6,
          columnCount: 4,
          selection: {
            start: { row: 0, col: 0 },
            end: { row: 0, col: 0 }
          },
          merges: [],
          columns: {},
          rows: {},
          cells: {
            "0:0": { value: "M", computedValue: "M" },
            "0:1": { value: "V", computedValue: "V" },
            "1:0": { value: "Jan", computedValue: "Jan" },
            "1:1": { value: 10, computedValue: 10 }
          },
          charts: [
            {
              id: "chart-callback",
              sheetId: "sheet-1",
              type: "line" as const,
              title: "Callback chart",
              sourceRange: {
                chartId: "chart-callback",
                sheetId: "sheet-1",
                rangeAddress: "A1:B2",
                orientation: "rows" as const,
                firstRowAsHeader: true,
                firstColumnAsLabel: true,
                autoRefresh: true
              },
              figure: {
                data: [],
                layout: {
                  title: "Callback chart"
                }
              },
              position: {
                fromCell: "C2",
                toCell: "H12",
                offsetX: 0,
                offsetY: 0,
                width: 360,
                height: 220,
                zIndex: 1
              },
              state: {
                selected: false,
                visible: true,
                locked: false
              },
              excelInterop: {
                unsupportedFeatures: ["custom:legacy-axis"]
              }
            }
          ]
        }
      ]
    };

    const exportedChartIds: string[] = [];
    const importedChartIds: string[] = [];
    const unsupportedFeatures: string[] = [];
    const bytes = await exportWorkbookToXlsx(workbook, {
      onChartExported: ({ chartId }) => exportedChartIds.push(chartId),
      onChartUnsupportedFeature: ({ feature }) => unsupportedFeatures.push(feature)
    });

    expect(exportedChartIds).toContain("chart-callback");
    expect(unsupportedFeatures).toContain("custom:legacy-axis");

    await importWorkbookFromXlsx(bytes, {
      onChartImported: ({ chart }) => importedChartIds.push(chart.id)
    });

    expect(importedChartIds.length).toBeGreaterThan(0);
  });

  it("bridges workbook engine chart events during xlsx export and import", async () => {
    const sourceEngine = new WorkbookEngine(
      {
        data: [
          {
            name: "Source",
            rowCount: 6,
            columnCount: 4,
            cells: {
              "0:0": { value: "Month", computedValue: "Month" },
              "0:1": { value: "Revenue", computedValue: "Revenue" },
              "1:0": { value: 1, computedValue: 1 },
              "1:1": { value: 10, computedValue: 10 },
              "2:0": { value: 2, computedValue: 2 },
              "2:1": { value: 18, computedValue: 18 }
            }
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const sourceSheet = sourceEngine.getActiveSheet();
    sourceEngine.createChart({
      sheetId: sourceSheet.id,
      chart: {
        id: "chart-engine-1",
        type: "line",
        title: "Revenue",
        sourceRange: {
          rangeAddress: "A1:B3",
          orientation: "rows",
          firstRowAsHeader: true,
          firstColumnAsLabel: true,
          autoRefresh: true
        },
        figure: {
          data: [{ type: "line", name: "Revenue", x: [1, 2], y: [10, 18] }],
          layout: { title: "Revenue" }
        },
        position: {
          fromCell: "C2",
          toCell: "H12",
          offsetX: 0,
          offsetY: 0,
          width: 360,
          height: 220,
          zIndex: 1
        },
        state: {
          selected: false,
          visible: true,
          locked: false
        }
      }
    });

    const exportedChartIds: string[] = [];
    sourceEngine.on("chart:exported", ({ chartId }) => exportedChartIds.push(chartId));
    const bytes = await exportWorkbookEngineToXlsx(sourceEngine);
    expect(bytes.length).toBeGreaterThan(0);
    expect(exportedChartIds).toContain("chart-engine-1");

    const targetEngine = new WorkbookEngine({}, new BasicFormulaEngine());
    const importedChartIds: string[] = [];
    targetEngine.on("chart:imported", ({ chartId }) => importedChartIds.push(chartId));
    await importWorkbookIntoEngineFromXlsx(targetEngine, bytes);
    expect(importedChartIds.length).toBeGreaterThan(0);
    const importedSheet = targetEngine.getActiveSheet();
    expect(targetEngine.getCharts(importedSheet.id).length).toBeGreaterThan(0);
  });

  it("sanitizes chart text payload to reduce formula injection risk", async () => {
    const workbook = {
      id: "workbook-sanitize",
      activeSheetId: "sheet-1",
      metadata: {},
      settings: {
        maxRows: 1000,
        maxColumns: 100,
        maxCellLength: 5000,
        maxFormulaLength: 2048,
        maxPasteCells: 10000,
        rowHeight: 28,
        columnWidth: 120,
        viewportBuffer: 6,
        maxHistorySize: 100,
        enableFormulas: true,
        clipboardPolicy: "text-only" as const
      },
      sheets: [
        {
          id: "sheet-1",
          name: "Safe",
          rowCount: 6,
          columnCount: 4,
          selection: {
            start: { row: 0, col: 0 },
            end: { row: 0, col: 0 }
          },
          merges: [],
          columns: {},
          rows: {},
          cells: {
            "0:0": { value: "Month", computedValue: "Month" },
            "0:1": { value: "+Revenue", computedValue: "+Revenue" },
            "1:0": { value: "Jan", computedValue: "Jan" },
            "1:1": { value: 10, computedValue: 10 }
          },
          charts: [
            {
              id: "chart-sanitize",
              sheetId: "sheet-1",
              type: "line" as const,
              title: "=cmd|' /C calc'!A0",
              sourceRange: {
                chartId: "chart-sanitize",
                sheetId: "sheet-1",
                rangeAddress: "A1:B2",
                orientation: "rows" as const,
                firstRowAsHeader: true,
                firstColumnAsLabel: true,
                autoRefresh: true
              },
              figure: {
                data: [],
                layout: {
                  title: "=cmd|' /C calc'!A0",
                  xAxis: {
                    title: "<script>alert(1)</script>"
                  },
                  yAxis: {
                    title: "@SUM(A1:A2)"
                  }
                }
              },
              position: {
                fromCell: "C2",
                toCell: "H12",
                offsetX: 0,
                offsetY: 0,
                width: 360,
                height: 220,
                zIndex: 1
              },
              state: {
                selected: false,
                visible: true,
                locked: false
              },
              excelInterop: {}
            }
          ]
        }
      ]
    };

    const bytes = await exportWorkbookToXlsx(workbook);
    const imported = await importWorkbookFromXlsx(bytes);
    const chart = imported.sheets[0]?.charts?.[0];

    expect(chart?.title?.startsWith("'")).toBe(true);
    const traceName = (chart?.figure.data?.[0] as { name?: string } | undefined)?.name;
    expect(typeof traceName).toBe("string");
    expect((traceName as string).startsWith("'")).toBe(true);
    const importedLayout = chart?.figure.layout as
      | { xAxis?: { title?: string }; yAxis?: { title?: string } }
      | undefined;
    expect(importedLayout?.xAxis?.title).not.toContain("<");
    expect(importedLayout?.xAxis?.title).not.toContain(">");
    expect(importedLayout?.yAxis?.title?.startsWith("'")).toBe(true);
  });

  it("exports unsupported chart payload as safe placeholder drawing with warning callback", async () => {
    const workbook = {
      id: "workbook-placeholder-export",
      activeSheetId: "sheet-1",
      metadata: {},
      settings: {
        maxRows: 1000,
        maxColumns: 100,
        maxCellLength: 5000,
        maxFormulaLength: 2048,
        maxPasteCells: 10000,
        rowHeight: 28,
        columnWidth: 120,
        viewportBuffer: 6,
        maxHistorySize: 100,
        enableFormulas: true,
        clipboardPolicy: "text-only" as const
      },
      sheets: [
        {
          id: "sheet-1",
          name: "Sheet",
          rowCount: 8,
          columnCount: 6,
          selection: {
            start: { row: 0, col: 0 },
            end: { row: 0, col: 0 }
          },
          merges: [],
          columns: {},
          rows: {},
          cells: {
            "0:0": { value: "A", computedValue: "A" },
            "0:1": { value: "B", computedValue: "B" },
            "1:0": { value: 1, computedValue: 1 },
            "1:1": { value: 2, computedValue: 2 }
          },
          charts: [
            {
              id: "chart-no-range",
              sheetId: "sheet-1",
              type: "surface3d" as const,
              title: "Unsupported",
              figure: {
                data: [],
                layout: { title: "Unsupported" }
              },
              position: {
                fromCell: "C2",
                toCell: "H14",
                offsetX: 0,
                offsetY: 0,
                width: 360,
                height: 220,
                zIndex: 1
              },
              state: {
                selected: false,
                visible: true,
                locked: false
              },
              excelInterop: {}
            }
          ]
        }
      ]
    };
    const unsupported: string[] = [];
    const bytes = await exportWorkbookToXlsx(workbook, {
      onChartUnsupportedFeature: ({ feature }) => unsupported.push(feature)
    });
    expect(bytes.length).toBeGreaterThan(0);
    expect(unsupported).toContain("export-fallback:placeholder");
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