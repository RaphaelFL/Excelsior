import { describe, expect, it } from "vitest";
import { WorkbookEngine } from "@excelsior/core";
import { BasicFormulaEngine } from "@excelsior/formulas";
import { attachWorkbookDevtools } from "../src/index";

describe("devtools", () => {
  it("captures public engine events and exposes snapshot access", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const session = attachWorkbookDevtools(engine);
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "42" });
    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });

    expect(session.events.some((event) => event.name === "command:completed")).toBe(true);
    expect(session.events.some((event) => event.name === "cell:updated")).toBe(true);
    expect(session.events.some((event) => event.name === "selection:changed")).toBe(true);
    expect(session.snapshot().sheets[0]?.cells["0:0"]?.value).toBe("42");

    session.stop();
    const eventCount = session.events.length;
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 1, value: "43" });
    expect(session.events).toHaveLength(eventCount);
  });
});