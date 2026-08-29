import { describe, expect, it } from "vitest";
import { SpreadsheetOperationError, WorkbookEngine } from "@excelsior/core";
import { packZip, unpackZip } from "../src/native-zip";
import { createXmlDocument } from "../src/xml-tree";
import {
  exportTableToXlsx,
  exportWorkbookToXlsx,
  importTableFromXlsx,
  importWorkbookFromXlsx
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