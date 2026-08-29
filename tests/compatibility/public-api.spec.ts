import { describe, expect, it, vi } from "vitest";
import { ServerSideRowModel, ViewportRowModel, WorkbookEngine } from "@excelsior/core";
import { BasicFormulaEngine } from "@excelsior/formulas";
import type { CellRichTextSegment, CellStyle, ClientSideQueryState, GridPlugin, SheetSplitPane } from "@excelsior/core";
import { createSpreadsheet } from "@excelsior/vanilla";

describe("public API compatibility", () => {
  const selectPivotOptions = (select: HTMLSelectElement | null, values: string[]) => {
    if (!select) {
      return;
    }

    for (const option of Array.from(select.options)) {
      option.selected = values.includes(option.value);
    }

    select.dispatchEvent(new Event("input", { bubbles: true }));
  };

  const setPivotSelectValue = (select: HTMLSelectElement | null, value: string) => {
    if (!select) {
      return;
    }

    select.value = value;
    select.dispatchEvent(new Event("input", { bubbles: true }));
  };

  const setPivotInputValue = (input: HTMLInputElement | null, value: string) => {
    if (!input) {
      return;
    }

    input.value = value;
    input.dispatchEvent(new Event("input", { bubbles: true }));
  };

  const setPivotToggle = (input: HTMLInputElement | null, checked: boolean) => {
    if (!input) {
      return;
    }

    input.checked = checked;
    input.dispatchEvent(new Event("input", { bubbles: true }));
  };

  const waitForActiveSheetName = async (engine: WorkbookEngine, expectedName: string) => {
    for (let attempt = 0; attempt < 10; attempt += 1) {
      if (engine.getActiveSheet().name === expectedName) {
        return;
      }

      await new Promise((resolve) => globalThis.setTimeout(resolve, 0));
    }

    throw new Error(`Timed out waiting for active sheet ${expectedName}`);
  };

  it("accepts workbook-style initialization data through the core and vanilla wrapper", () => {
    const engine = new WorkbookEngine({
      data: [
        {
          id: "sheet-1",
          name: "Sheet1",
          rowCount: 4,
          columnCount: 4,
          cells: {
            "0:0": { value: 10, computedValue: 10 },
            "1:0": { value: "=A1*2", formula: "=A1*2", computedValue: 20 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });

    expect(engine.getDisplayValue(engine.getActiveSheet().id, 0, 0)).toBe("10");
    expect(engine.getSnapshot().sheets[0]?.name).toBe("Sheet1");

    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const instance = createSpreadsheet(container, {
      data: [{ name: "Sheet1" }]
    });

    expect(instance.engine.getSnapshot().sheets[0]?.name).toBe("Sheet1");
    expect(container.querySelector(".excelsior-shell")).not.toBeNull();

    instance.destroy();
    container.remove();
  });

  it("exposes rich text and overflow through the public core API", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const richText: CellRichTextSegment[] = [{ text: "Docs", hyperlink: "https://example.com" }];
    const overflow: CellStyle["overflow"] = "ellipsis";

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "Docs" });
    engine.setCellRichText({ sheetId: sheet.id, row: 0, col: 0, richText });
    engine.setCellStyle({ sheetId: sheet.id, row: 0, col: 0, style: { overflow } });

    expect(engine.getCellRichText(sheet.id, 0, 0)?.[0]).toEqual({ text: "Docs", hyperlink: "https://example.com/" });
    expect(engine.getCell(sheet.id, 0, 0)?.style?.overflow).toBe("ellipsis");
  });

  it("exposes serializable client-side multi-column query descriptors", () => {
    const engine = new WorkbookEngine({ data: [{ rowCount: 2, columnCount: 2 }] });
    const sheet = engine.getActiveSheet();
    engine.applyClientSideSortFilter({
      sheetId: sheet.id,
      sort: [{ column: 0, direction: "asc" }, { column: 1, direction: "desc" }],
      filters: [{ column: 1, type: "number", operator: "gte", value: 10 }]
    });
    const state: ClientSideQueryState | undefined = engine.getClientSideQuery(sheet.id);
    expect(JSON.parse(JSON.stringify(state))).toEqual(state);
  });

  it("exposes serializable split pane state through the public core API", () => {
    const engine = new WorkbookEngine({ data: [{ rowCount: 20, columnCount: 10 }] });
    const sheet = engine.getActiveSheet();
    const splitPane: SheetSplitPane = { horizontalRow: 5, verticalColumn: 3 };

    engine.setSplitPane(sheet.id, splitPane);

    expect(engine.getSplitPane(sheet.id)).toEqual(splitPane);
    expect(JSON.parse(JSON.stringify(engine.toJSON().sheets[0]?.splitPane))).toEqual(splitPane);
  });

  it("accepts localization options through the public vanilla wrapper", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const instance = createSpreadsheet(container, {
      data: [{ name: "Sheet1" }],
      localization: {
        direction: "rtl",
        messages: {
          findReplace: "Localizar"
        },
        shortcuts: {
          openFindReplace: ["Alt+F"]
        }
      }
    });

    expect(container.querySelector(".excelsior-shell")?.getAttribute("dir")).toBe("rtl");
    expect(container.querySelector<HTMLButtonElement>("[data-action='find-replace']")?.textContent).toBe("Localizar");

    instance.destroy();
    container.remove();
  });

  it("accepts renderDebounceMs through the public vanilla wrapper", () => {
    vi.useFakeTimers();

    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const instance = createSpreadsheet(container, {
      data: [{ name: "Sheet1" }],
      renderDebounceMs: 5
    });
    const sheet = instance.engine.getActiveSheet();
    const renderSpy = vi.spyOn(instance.renderer, "render");
    renderSpy.mockClear();

    instance.engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });
    instance.engine.selectRange({ sheetId: sheet.id, rowStart: 1, rowEnd: 1, colStart: 1, colEnd: 1 });

    expect(renderSpy).not.toHaveBeenCalled();

    vi.advanceTimersByTime(5);

    expect(renderSpy).toHaveBeenCalledTimes(1);

    renderSpy.mockRestore();
    instance.destroy();
    container.remove();
    vi.useRealTimers();
  });

  it("materializes a configured pivot sheet through the public vanilla toolbar", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const instance = createSpreadsheet(container, {
      data: [
        {
          name: "Sales",
          rowCount: 6,
          columnCount: 2,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Sales", computedValue: "Sales" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: 10, computedValue: 10 },
            "2:0": { value: "South", computedValue: "South" },
            "2:1": { value: 5, computedValue: 5 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });

    container.querySelector<HTMLButtonElement>("[data-action='create-pivot']")?.click();
    expect(container.querySelector<HTMLElement>(".excelsior-pivot-panel")?.hidden).toBe(false);
    container.querySelector<HTMLButtonElement>("[data-pivot-action='apply']")?.click();
    await waitForActiveSheetName(instance.engine, "Sales Pivot");

    expect(instance.engine.getActiveSheet().name).toBe("Sales Pivot");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 0, 1)).toBe("SUM Sales");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 1, 0)).toBe("North");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 1, 1)).toBe("10");

    instance.destroy();
    container.remove();
  });

  it("supports multiple pivot row dimensions through the public vanilla toolbar", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const instance = createSpreadsheet(container, {
      data: [
        {
          name: "Tickets",
          rowCount: 5,
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

    container.querySelector<HTMLButtonElement>("[data-action='create-pivot']")?.click();
    selectPivotOptions(container.querySelector<HTMLSelectElement>("[data-pivot-role='row']"), ["Team", "Status"]);
    selectPivotOptions(container.querySelector<HTMLSelectElement>("[data-pivot-role='column']"), []);
    container.querySelector<HTMLButtonElement>("[data-pivot-action='apply']")?.click();
    await waitForActiveSheetName(instance.engine, "Tickets Pivot");

    expect(instance.engine.getActiveSheet().name).toBe("Tickets Pivot");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 0, 0)).toBe("Team");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 0, 1)).toBe("Status");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 0, 2)).toBe("SUM Count");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 3, 0)).toBe("Eng Total");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 6, 0)).toBe("Grand Total");

    instance.destroy();
    container.remove();
  });

  it("supports multiple pivot values and total toggles through the public vanilla toolbar", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const instance = createSpreadsheet(container, {
      data: [
        {
          name: "Sales",
          rowCount: 6,
          columnCount: 4,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Sales", computedValue: "Sales" },
            "0:3": { value: "Orders", computedValue: "Orders" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 10, computedValue: 10 },
            "1:3": { value: 1, computedValue: 1 },
            "2:0": { value: "North", computedValue: "North" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 20, computedValue: 20 },
            "2:3": { value: 2, computedValue: 2 },
            "3:0": { value: "South", computedValue: "South" },
            "3:1": { value: "Q1", computedValue: "Q1" },
            "3:2": { value: 5, computedValue: 5 },
            "3:3": { value: 1, computedValue: 1 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });

    container.querySelector<HTMLButtonElement>("[data-action='create-pivot']")?.click();
    setPivotSelectValue(
      container.querySelector<HTMLSelectElement>("[data-pivot-role='value-field'][data-pivot-value-index='0']"),
      "Sales"
    );
    setPivotInputValue(
      container.querySelector<HTMLInputElement>("[data-pivot-role='value-alias'][data-pivot-value-index='0']"),
      "Revenue"
    );
    container.querySelector<HTMLButtonElement>("[data-pivot-action='add-value']")?.click();
    setPivotSelectValue(
      container.querySelector<HTMLSelectElement>("[data-pivot-role='value-field'][data-pivot-value-index='1']"),
      "Orders"
    );
    setPivotInputValue(
      container.querySelector<HTMLInputElement>("[data-pivot-role='value-alias'][data-pivot-value-index='1']"),
      "Deals"
    );
    setPivotToggle(container.querySelector<HTMLInputElement>("[data-pivot-role='include-row-totals']"), false);
    setPivotToggle(container.querySelector<HTMLInputElement>("[data-pivot-role='include-column-totals']"), false);
    setPivotToggle(container.querySelector<HTMLInputElement>("[data-pivot-role='include-subtotals']"), false);
    container.querySelector<HTMLButtonElement>("[data-pivot-action='apply']")?.click();
    await waitForActiveSheetName(instance.engine, "Sales Pivot");

    expect(instance.engine.getActiveSheet().name).toBe("Sales Pivot");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 0, 1)).toBe("Q1 • Revenue");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 0, 2)).toBe("Q1 • Deals");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 0, 3)).toBe("Q2 • Revenue");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 0, 4)).toBe("Q2 • Deals");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 1, 0)).toBe("North");
    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 1, 4)).toBe("2");

    instance.destroy();
    container.remove();
  });

  it("accepts an explicit row model through the public vanilla wrapper", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "160px";
    document.body.append(container);

    const rowModel = new ViewportRowModel({
      rowCount: 200,
      getRows: ({ startRow }) => ({
        rowCount: 200,
        rows: [{ index: startRow }]
      })
    });

    const instance = createSpreadsheet(container, {
      data: [{ name: "Sheet1" }],
      rowModel
    });

    expect(instance.engine.getRowModel(instance.engine.getActiveSheet().id)).toBe(rowModel);

    instance.destroy();
    container.remove();
  });

  it("exposes remote pivot toggling through the public vanilla toolbar", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "160px";
    document.body.append(container);

    const dataSource = {
      getRows: vi.fn(async ({ startRow, endRow, pivotModel }) => ({
        totalRows: 200,
        rows: Array.from({ length: endRow - startRow + 1 }, (_value, index) => ({ index: startRow + index })),
        pivotModel
      }))
    };

    const rowModel = new ServerSideRowModel({
      rowCount: 200,
      dataSource
    });

    const instance = createSpreadsheet(container, {
      data: [{ name: "Sheet1" }],
      rowModel
    });

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(1);
    });

    container.querySelector<HTMLButtonElement>("[data-action='pivot-column']")?.click();

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(2);
    });
    expect(dataSource.getRows.mock.calls[1][0]).toMatchObject({
      pivotModel: [{ field: "A" }]
    });
    expect(container.querySelector("[data-action='pivot-column']")?.getAttribute("aria-pressed")).toBe("true");

    instance.destroy();
    container.remove();
  });

  it("updates sort, filter, grouping, expansion, pivot and aggregation through the public core API without replacing the remote row model", async () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const rowModel = new ServerSideRowModel({
      rowCount: 100,
      dataSource: {
        getRows: async ({ sortModel, filterModel, groupKeys, expandedGroupPaths, pivotModel, aggregateModel }) => ({
          totalRows: 100,
          rows: [{ index: 0 }],
          groupInfo: [{ key: "engineering", path: ["engineering"], level: 0, childCount: 12, expanded: true }],
          sortModel,
          filterModel,
          groupKeys,
          expandedGroupPaths,
          pivotModel,
          aggregateModel
        })
      }
    });

    engine.setRowModel(sheet.id, rowModel);
    engine.updateRemoteRowModel(sheet.id, {
      sortModel: [{ field: "price", direction: "desc" }],
      filterModel: {
        status: { operator: "equals", value: "active" }
      },
      groupKeys: ["team", "status"],
      expandedGroupPaths: [["engineering"], ["engineering", "active"]],
      pivotModel: [{ field: "quarter" }],
      aggregateModel: [
        { field: "revenue", function: "sum", as: "revenueSum" },
        { field: "id", function: "count", as: "rowCount" }
      ]
    });

    const result = await rowModel.getRows({ sheetId: sheet.id, startRow: 0, endRow: 0 });

    expect(engine.getRowModel(sheet.id)).toBe(rowModel);
    expect(engine.getRemoteRowModelRequest(sheet.id)).toEqual({
      sortModel: [{ field: "price", direction: "desc" }],
      filterModel: {
        status: { operator: "equals", value: "active" }
      },
      groupKeys: ["team", "status"],
      expandedGroupPaths: [["engineering"], ["engineering", "active"]],
      pivotModel: [{ field: "quarter" }],
      aggregateModel: [
        { field: "revenue", function: "sum", as: "revenueSum" },
        { field: "id", function: "count", as: "rowCount" }
      ]
    });
    expect(result).toEqual({
      rowCount: 100,
      rows: [{ index: 0, group: { key: "engineering", path: ["engineering"], level: 0, childCount: 12, expanded: true } }]
    });
  });

  it("exposes remote group expansion state through the public core API", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const rowModel = new ServerSideRowModel({
      rowCount: 100,
      dataSource: {
        getRows: async () => ({
          totalRows: 100,
          rows: [{ index: 0 }]
        })
      }
    });

    engine.setRowModel(sheet.id, rowModel);
    engine.setRemoteGroupExpanded(sheet.id, ["engineering"], true);
    engine.setRemoteGroupExpanded(sheet.id, ["engineering", "active"], true);
    engine.setRemoteGroupExpanded(sheet.id, ["engineering"], false);

    expect(engine.getRemoteRowModelRequest(sheet.id)).toEqual({
      expandedGroupPaths: [["engineering", "active"]]
    });
  });

  it("exposes the plugin engine through the public core API", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const plugin: GridPlugin = {
      id: "compat-plugin",
      setup(context) {
        context.setState({ updates: 0 });
        return context.on("cell:updated", () => {
          context.setState<{ updates: number }>((previous) => ({
            updates: (previous?.updates ?? 0) + 1
          }));
        });
      }
    };

    engine.registerPlugin(plugin);
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "ok" });

    expect(engine.getRegisteredPlugins()).toContainEqual({ id: "compat-plugin", enabled: true });
    expect(engine.getPluginState<{ updates: number }>("compat-plugin")).toEqual({ updates: 1 });

    engine.disablePlugin("compat-plugin");
    expect(engine.isPluginEnabled("compat-plugin")).toBe(false);
  });

  it("exposes validation rules and safe custom validators through the public core API", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.registerValidator("prefix", ({ value, params }) => {
      const prefix = String(params?.prefix ?? "");
      return typeof value === "string" && value.startsWith(prefix)
        ? undefined
        : {
            code: "CORE_VALIDATION_CUSTOM_PREFIX",
            message: `O valor deve começar com ${prefix}.`,
            ruleType: "custom",
            validator: "prefix"
          };
    });

    engine.setCellValidation({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      validation: {
        rules: [{ type: "custom", validator: "prefix", params: { prefix: "SKU-" } }]
      }
    });

    expect(engine.getCellValidation(sheet.id, 0, 0)?.rules).toHaveLength(1);
    expect(engine.validateCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "SKU-001" }).valid).toBe(true);
    expect(engine.validateCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "001" }).valid).toBe(false);
  });

  it("exposes conditional formatting through the public core API", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 42 });
    engine.setConditionalFormattingRules(sheet.id, [
      {
        id: "compat-conditional",
        type: "greaterThan",
        range: {
          start: { row: 0, col: 0 },
          end: { row: 0, col: 0 }
        },
        value: 10,
        priority: 10,
        style: {
          backgroundColor: "#fde68a"
        }
      }
    ]);

    expect(engine.getConditionalFormattingRules(sheet.id)).toHaveLength(1);
    expect(engine.getConditionalStyle(sheet.id, 0, 0)).toMatchObject({
      backgroundColor: "#fde68a"
    });
  });

  it("exposes undoable batch operations through the public core API", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.applyBatchOperations({
      anchorSheetId: sheet.id,
      operations: [
        {
          op: "add",
          id: sheet.id,
          path: ["cells", "0:0"],
          value: { value: "A", computedValue: "A" }
        },
        {
          op: "add",
          id: sheet.id,
          path: ["cells", "1:0"],
          value: { value: "B", computedValue: "B" }
        }
      ],
      affectedRanges: [
        {
          start: { row: 0, col: 0 },
          end: { row: 1, col: 0 }
        }
      ]
    });

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("A");
    expect(engine.getDisplayValue(sheet.id, 1, 0)).toBe("B");
    expect(engine.undo()).toBe(true);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(engine.getDisplayValue(sheet.id, 1, 0)).toBe("");
  });

  it("exposes undoable batched cell updates through the public core API", () => {
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();

    engine.updateCells({
      sheetId: sheet.id,
      updates: [
        { row: 0, col: 0, value: 10 },
        { row: 0, col: 1, value: "=A1*2" }
      ]
    });

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("10");
    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("20");
    expect(engine.undo()).toBe(true);
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("");
  });

  it("exposes keyed cell transactions through the public core API", () => {
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
    expect(engine.getDisplayValue(sheet.id, 1, 1)).toBe("legacy");
  });

  it("preserves selection during public high-frequency update APIs", () => {
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
      updates: [{ row: 0, col: 0, value: 10 }]
    });
    engine.applyCellTransaction({
      sheetId: sheet.id,
      changes: [{ type: "upsert", key: "1:0", value: 20 }]
    });

    expect(engine.getSelection(sheet.id)).toEqual({
      start: { row: 4, col: 2 },
      end: { row: 5, col: 3 }
    });
  });

  it("exposes client-side pivot sheets through the public core API", () => {
    const engine = new WorkbookEngine({
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

    const pivot = engine.createPivotSheet({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      },
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum" }]
    });
    const pivotSheetId = engine.addPivotSheet({
      sourceSheetId: sourceSheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 3, col: 2 }
      },
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum" }],
      sheetName: "Sales Pivot View"
    });

    expect(pivot.name).toBe("Sales Pivot");
    expect(pivot.cells?.["1:1"]?.value).toBe(10);
    expect(engine.getActiveSheet().id).toBe(pivotSheetId);
    expect(engine.getActiveSheet().name).toBe("Sales Pivot View");
    expect(engine.getDisplayValue(pivotSheetId, 1, 1)).toBe("10");
  });

  it("exposes pivot inference through the public core API", () => {
    const engine = new WorkbookEngine({
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
            "2:0": { value: "South", computedValue: "South" },
            "2:1": { value: "Q2", computedValue: "Q2" },
            "2:2": { value: 5, computedValue: 5 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sheet = engine.getActiveSheet();

    const inferred = engine.inferPivotSheet({
      sourceSheetId: sheet.id,
      sourceRange: {
        start: { row: 0, col: 0 },
        end: { row: 2, col: 2 }
      }
    });

    expect(inferred).toMatchObject({
      rows: ["Region"],
      columns: ["Quarter"],
      values: [{ field: "Sales", aggregate: "sum" }],
      includeSubtotals: true
    });
  });

  it("treats redundant setCellValue writes as a public no-op", () => {
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

  it("treats no-op batched cell updates as a public no-op", () => {
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

  it("accepts custom renderers and editors through the public vanilla wrapper", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const instance = createSpreadsheet(container, {
      data: [{ name: "Sheet1" }],
      cellRenderers: [
        {
          id: "compat-renderer",
          matches: ({ row, col }) => row === 0 && col === 0,
          render: () => ({ text: "Compat", classNames: ["compat-renderer"] })
        }
      ],
      cellEditors: [
        {
          id: "compat-editor",
          matches: ({ row, col }) => row === 0 && col === 1,
          create: () => {
            let input: HTMLInputElement | undefined;
            return {
              mount(host) {
                input = document.createElement("input");
                input.value = "ok";
                host.replaceChildren(input);
              },
              getValue() {
                return input?.value ?? "";
              }
            };
          }
        }
      ]
    });

    expect(container.querySelector(".compat-renderer")?.textContent).toBe("Compat");

    container.querySelector<HTMLElement>("[data-row='0'][data-col='1']")?.dispatchEvent(
      new MouseEvent("dblclick", { bubbles: true })
    );
    container.querySelector<HTMLInputElement>(".excelsior-custom-editor-host input")?.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true })
    );

    expect(instance.engine.getDisplayValue(instance.engine.getActiveSheet().id, 0, 1)).toBe("ok");
    expect(container.querySelector(".excelsior-fill-handle")).not.toBeNull();

    instance.destroy();
    container.remove();
  });

  it("accepts autofill configuration through the public vanilla wrapper", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const instance = createSpreadsheet(container, {
      data: [{ name: "Sheet1" }],
      autofill: {
        maxCells: 128,
        copyStyle: true
      }
    });

    expect(container.querySelector(".excelsior-fill-handle")).not.toBeNull();

    instance.destroy();
    container.remove();
  });
});