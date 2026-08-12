import { describe, expect, it, vi } from "vitest";
import {
  cellAddressToLabel,
  defaultPivotModule,
  ServerSideRowModel,
  ViewportRowModel,
  WorkbookEngine,
  type PivotBuildAsyncOptions,
  type PivotModule,
  type PivotSheetInput,
  type SpreadsheetOperation,
  type WorkbookDataInput
} from "@excelsior/core";
import { BasicFormulaEngine } from "@excelsior/formulas";
import { DomSpreadsheetRenderer } from "../src/index";

const createDeferred = <T>() => {
  let resolve: (value: T) => void;
  let reject: (reason?: unknown) => void;
  const promise = new Promise<T>((innerResolve, innerReject) => {
    resolve = innerResolve;
    reject = innerReject;
  });

  return {
    promise,
    resolve: resolve!,
    reject: reject!
  };
};

describe("DomSpreadsheetRenderer", () => {
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

  const dispatchMouse = (target: EventTarget | null | undefined, type: string, init: MouseEventInit) => {
    if (!target) {
      return;
    }
    target.dispatchEvent(
      new MouseEvent(type, {
        bubbles: true,
        cancelable: true,
        ...init
      })
    );
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

  it("switches tabs and adds sheets from the UI", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [{ name: "One" }, { name: "Two" }]
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const tabs = Array.from(container.querySelectorAll<HTMLButtonElement>("[data-sheet-id]"));
    tabs.find((tab) => tab.textContent?.includes("Two"))?.click();

    expect(engine.getActiveSheet().name).toBe("Two");

    container.querySelector<HTMLButtonElement>("[data-action='add-sheet']")?.click();
    expect(engine.getSnapshot().sheets).toHaveLength(3);

    renderer.dispose();
    container.remove();
  });

  it("renders row headers and keeps the grid offset from the gutter", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Rows",
            rowCount: 4,
            columnCount: 2,
            cells: {
              "0:0": { value: "Cliente", computedValue: "Cliente" },
              "1:0": { value: "Delta", computedValue: "Delta" },
              "2:0": { value: "Omega", computedValue: "Omega" }
            }
          }
        ]
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const headers = Array.from(container.querySelectorAll<HTMLElement>(".excelsior-row-header"));
    const firstCell = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");

    expect(headers.slice(0, 3).map((header) => header.textContent)).toEqual(["1", "2", "3"]);
    expect(Number.parseFloat(firstCell?.style.left ?? "0")).toBeGreaterThanOrEqual(56);

    headers[2]?.click();
    expect(engine.getActiveSheet().selection.end.row).toBe(2);

    renderer.dispose();
    container.remove();
  });

  it("selects the full sheet when the top-left corner is clicked", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Select All",
            rowCount: 6,
            columnCount: 4
          }
        ]
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    container.querySelector<HTMLButtonElement>("[data-corner-action='select-all']")?.click();

    expect(engine.getActiveSheet().selection.start).toEqual({ row: 0, col: 0 });
    expect(engine.getActiveSheet().selection.end).toEqual({ row: 5, col: 3 });

    renderer.dispose();
    container.remove();
  });

  it("renders column headers for columns beyond L", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Wide",
            rowCount: 4,
            columnCount: 16
          }
        ]
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    const headerM = container.querySelector<HTMLElement>("[data-column-header-col='12']");
    const headerP = container.querySelector<HTMLElement>("[data-column-header-col='15']");

    expect(headerM?.textContent).toBe("M");
    expect(headerP?.textContent).toBe("P");

    renderer.dispose();
    container.remove();
  });

  it("opens a pivot panel and materializes a configured pivot from the toolbar", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
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
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    container.querySelector<HTMLButtonElement>("[data-action='create-pivot']")?.click();
    const panel = container.querySelector<HTMLElement>(".excelsior-pivot-panel");
    const rowSelect = container.querySelector<HTMLSelectElement>("[data-pivot-role='row']");
    const columnSelect = container.querySelector<HTMLSelectElement>("[data-pivot-role='column']");
    const valueSelect = container.querySelector<HTMLSelectElement>("[data-pivot-role='value-field'][data-pivot-value-index='0']");

    expect(panel?.hidden).toBe(false);
    expect(Array.from(rowSelect?.selectedOptions ?? []).map((option) => option.value)).toEqual(["Region"]);
    expect(Array.from(columnSelect?.selectedOptions ?? []).map((option) => option.value)).toEqual(["Quarter"]);
    expect(valueSelect?.value).toBe("Sales");

    container.querySelector<HTMLButtonElement>("[data-pivot-action='apply']")?.click();
    await waitForActiveSheetName(engine, "Sales Pivot");

    const activeSheet = engine.getActiveSheet();
    expect(activeSheet.name).toBe("Sales Pivot");
    expect(engine.getDisplayValue(activeSheet.id, 0, 0)).toBe("Region");
    expect(engine.getDisplayValue(activeSheet.id, 0, 1)).toBe("Q1");
    expect(engine.getDisplayValue(activeSheet.id, 0, 2)).toBe("Q2");
    expect(engine.getDisplayValue(activeSheet.id, 0, 3)).toBe("Total");
    expect(engine.getDisplayValue(activeSheet.id, 1, 0)).toBe("North");
    expect(engine.getDisplayValue(activeSheet.id, 1, 1)).toBe("10");
    expect(engine.getDisplayValue(activeSheet.id, 1, 2)).toBe("20");
    expect(engine.getDisplayValue(activeSheet.id, 1, 3)).toBe("30");
    expect(engine.getDisplayValue(activeSheet.id, 2, 0)).toBe("South");
    expect(engine.getDisplayValue(activeSheet.id, 2, 1)).toBe("5");
    expect(engine.getDisplayValue(activeSheet.id, 2, 2)).toBe("");
    expect(engine.getDisplayValue(activeSheet.id, 2, 3)).toBe("5");

    renderer.dispose();
    container.remove();
  });

  it("supports multiple row dimensions from the pivot panel", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
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
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    container.querySelector<HTMLButtonElement>("[data-action='create-pivot']")?.click();
    selectPivotOptions(container.querySelector<HTMLSelectElement>("[data-pivot-role='row']"), ["Team", "Status"]);
    selectPivotOptions(container.querySelector<HTMLSelectElement>("[data-pivot-role='column']"), []);
    container.querySelector<HTMLButtonElement>("[data-pivot-action='apply']")?.click();
    await waitForActiveSheetName(engine, "Tickets Pivot");

    const activeSheet = engine.getActiveSheet();
    expect(activeSheet.name).toBe("Tickets Pivot");
    expect(engine.getDisplayValue(activeSheet.id, 0, 0)).toBe("Team");
    expect(engine.getDisplayValue(activeSheet.id, 0, 1)).toBe("Status");
    expect(engine.getDisplayValue(activeSheet.id, 0, 2)).toBe("SUM Count");
    expect(engine.getDisplayValue(activeSheet.id, 1, 0)).toBe("Eng");
    expect(engine.getDisplayValue(activeSheet.id, 1, 1)).toBe("Open");
    expect(engine.getDisplayValue(activeSheet.id, 1, 2)).toBe("1");
    expect(engine.getDisplayValue(activeSheet.id, 3, 0)).toBe("Eng Total");
    expect(engine.getDisplayValue(activeSheet.id, 3, 2)).toBe("2");
    expect(engine.getDisplayValue(activeSheet.id, 6, 0)).toBe("Grand Total");
    expect(engine.getDisplayValue(activeSheet.id, 6, 2)).toBe("3");

    renderer.dispose();
    container.remove();
  });

  it("supports multiple pivot values, aliases and total toggles from the panel", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
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
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

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
    await waitForActiveSheetName(engine, "Sales Pivot");

    const activeSheet = engine.getActiveSheet();
    expect(activeSheet.name).toBe("Sales Pivot");
    expect(engine.getDisplayValue(activeSheet.id, 0, 0)).toBe("Region");
    expect(engine.getDisplayValue(activeSheet.id, 0, 1)).toBe("Q1 • Revenue");
    expect(engine.getDisplayValue(activeSheet.id, 0, 2)).toBe("Q1 • Deals");
    expect(engine.getDisplayValue(activeSheet.id, 0, 3)).toBe("Q2 • Revenue");
    expect(engine.getDisplayValue(activeSheet.id, 0, 4)).toBe("Q2 • Deals");
    expect(engine.getDisplayValue(activeSheet.id, 0, 5)).toBe("");
    expect(engine.getDisplayValue(activeSheet.id, 1, 0)).toBe("North");
    expect(engine.getDisplayValue(activeSheet.id, 1, 1)).toBe("10");
    expect(engine.getDisplayValue(activeSheet.id, 1, 2)).toBe("1");
    expect(engine.getDisplayValue(activeSheet.id, 1, 3)).toBe("20");
    expect(engine.getDisplayValue(activeSheet.id, 1, 4)).toBe("2");

    renderer.dispose();
    container.remove();
  });

  it("shows pivot progress and allows canceling a pending pivot build", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const deferred = createDeferred<WorkbookDataInput>();
    const pivotModule: PivotModule = {
      createPivotSheet: defaultPivotModule.createPivotSheet,
      inferPivotSheet: defaultPivotModule.inferPivotSheet,
      createPivotSheetAsync: async (
        workbook,
        _input: PivotSheetInput,
        options?: PivotBuildAsyncOptions
      ): Promise<WorkbookDataInput> => {
        options?.onProgress?.({
          phase: "aggregate",
          completed: 1,
          total: 4
        });

        return new Promise<WorkbookDataInput>((resolve, reject) => {
          const abortHandler = () => reject(new DOMException("Pivot build was aborted.", "AbortError"));
          if (options?.signal?.aborted) {
            abortHandler();
            return;
          }

          options?.signal?.addEventListener("abort", abortHandler, { once: true });
          deferred.promise.then(resolve, reject).finally(() => {
            options?.signal?.removeEventListener("abort", abortHandler);
          });
        });
      }
    };
    const engine = new WorkbookEngine(
      {
        pivotModule,
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
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const status = container.querySelector<HTMLElement>(".excelsior-status-message");

    container.querySelector<HTMLButtonElement>("[data-action='create-pivot']")?.click();
    container.querySelector<HTMLButtonElement>("[data-pivot-action='apply']")?.click();

    await vi.waitFor(() => {
      expect(status?.textContent).toBe("Criando pivot... 25%");
      expect(container.querySelector<HTMLButtonElement>("[data-pivot-action='close']")?.textContent).toBe("Cancelar");
    });

    container.querySelector<HTMLButtonElement>("[data-pivot-action='close']")?.click();

    await vi.waitFor(() => {
      expect(status?.textContent).toBe("Criação da pivot cancelada.");
      expect(container.querySelector<HTMLElement>(".excelsior-pivot-panel")?.hidden).toBe(true);
    });

    deferred.resolve(
      defaultPivotModule.createPivotSheet(engine.getSnapshot(), {
        sourceSheetId: engine.getSnapshot().sheets[0]!.id,
        sourceRange: {
          start: { row: 0, col: 0 },
          end: { row: 3, col: 2 }
        },
        rows: ["Region"],
        columns: ["Quarter"],
        values: [{ field: "Sales", aggregate: "sum" }]
      })
    );

    renderer.dispose();
    container.remove();
  });

  it("shows a friendly message when a pivot source exceeds the client-side limit", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
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
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const status = container.querySelector<HTMLElement>(".excelsior-status-message");

    container.querySelector<HTMLButtonElement>("[data-action='create-pivot']")?.click();
    container.querySelector<HTMLButtonElement>("[data-pivot-action='apply']")?.click();

    await vi.waitFor(() => {
      expect(status?.textContent).toContain("Fonte grande demais para pivot local");
      expect(status?.textContent).toContain("Limite atual: 2 linhas");
      expect(container.querySelector<HTMLElement>(".excelsior-pivot-panel")?.hidden).toBe(false);
    });

    renderer.dispose();
    container.remove();
  });

  it("supports keyboard navigation on sheet tabs without nested interactive controls", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [{ name: "One" }, { name: "Two" }, { name: "Three" }]
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const tabs = Array.from(container.querySelectorAll<HTMLButtonElement>("[data-sheet-id]"));
    const firstTab = tabs[0];

    expect(container.querySelector(".excelsior-sheet-tab [data-close-sheet-id]")).toBeNull();

    firstTab?.focus();
    firstTab?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));

    const refreshedTabs = Array.from(container.querySelectorAll<HTMLButtonElement>("[data-sheet-id]"));
    expect(engine.getActiveSheet().name).toBe("Two");
    expect(document.activeElement).toBe(refreshedTabs[1]);
    expect(refreshedTabs[1]?.getAttribute("aria-selected")).toBe("true");

    renderer.dispose();
    container.remove();
  });

  it("commits edits from the formula bar to the selected cell", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });
    const input = container.querySelector<HTMLInputElement>(".excelsior-formula-input");
    input!.value = "123";
    input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("123");

    renderer.dispose();
    container.remove();
  });

  it("renders merged cells with persisted style and sizing", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

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
          viewportBuffer: 2,
          maxHistorySize: 100,
          enableFormulas: true,
          clipboardPolicy: "text-only"
        },
        sheets: [
          {
            id: "sheet-1",
            name: "Styled",
            rowCount: 3,
            columnCount: 3,
            cells: {
              "0:0": {
                value: "Merged",
                computedValue: "Merged",
                style: {
                  fontWeight: "bold",
                  align: "center",
                  backgroundColor: "#ffeecc"
                }
              }
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
            selection: {
              start: { row: 0, col: 0 },
              end: { row: 0, col: 0 }
            }
          }
        ]
      },
      new BasicFormulaEngine()
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const mergedCell = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");

    expect(mergedCell).not.toBeNull();
    expect(mergedCell?.classList.contains("is-merged")).toBe(true);
    expect(mergedCell?.style.width).toBe("300px");
    expect(mergedCell?.style.height).toBe("68px");
    expect(mergedCell?.style.fontWeight).toBe("bold");
    expect(mergedCell?.style.textAlign).toBe("center");
    expect(container.querySelector("[data-row='0'][data-col='1']")).toBeNull();

    renderer.dispose();
    container.remove();
  });

  it("shows validation feedback and blocks invalid formula bar commits", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    engine.setCellValidation({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      validation: {
        rules: [{ type: "number", min: 10, max: 20 }]
      }
    });

    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });
    const input = container.querySelector<HTMLInputElement>(".excelsior-formula-input");
    input!.value = "abc";
    input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(container.querySelector<HTMLElement>(".excelsior-status-message")?.textContent).toContain("numérico");
    expect(container.querySelector("[data-row='0'][data-col='0']")?.classList.contains("is-error")).toBe(true);

    renderer.dispose();
    container.remove();
  });

  it("coalesces burst renders when renderDebounceMs is configured", () => {
    vi.useFakeTimers();

    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine, { renderDebounceMs: 5 });
    const renderSpy = vi.spyOn(renderer, "render");
    renderSpy.mockClear();
    const sheet = engine.getActiveSheet();

    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });
    engine.selectRange({ sheetId: sheet.id, rowStart: 1, rowEnd: 1, colStart: 1, colEnd: 1 });

    expect(renderSpy).not.toHaveBeenCalled();

    vi.advanceTimersByTime(5);

    expect(renderSpy).toHaveBeenCalledTimes(1);

    renderSpy.mockRestore();
    renderer.dispose();
    container.remove();
    vi.useRealTimers();
  });

  it("announces status updates and exposes pressed state on toolbar toggles", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    engine.setCellValidation({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      validation: {
        rules: [{ type: "number", min: 10, max: 20 }]
      }
    });

    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });

    const boldButton = container.querySelector<HTMLButtonElement>("[data-action='bold']");
    const formulaInput = container.querySelector<HTMLInputElement>(".excelsior-formula-input");
    const statusMessage = container.querySelector<HTMLElement>(".excelsior-status-message");
    const findButton = container.querySelector<HTMLButtonElement>("[data-action='find-replace']");

    expect(boldButton?.getAttribute("aria-pressed")).toBe("false");

    boldButton?.click();
    const pressedBoldButton = container.querySelector<HTMLButtonElement>("[data-action='bold']");
    expect(pressedBoldButton?.getAttribute("aria-pressed")).toBe("true");

    formulaInput!.value = "abc";
    formulaInput!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));

    expect(statusMessage?.getAttribute("role")).toBe("status");
    expect(statusMessage?.getAttribute("aria-live")).toBe("polite");
    expect(statusMessage?.textContent).toContain("numérico");

    findButton?.click();
    const findResults = container.querySelector<HTMLElement>(".excelsior-find-replace-results");
    expect(findResults?.getAttribute("role")).toBe("status");
    expect(findResults?.getAttribute("aria-live")).toBe("polite");

    renderer.dispose();
    container.remove();
  });

  it("exposes grid semantics and announces the active cell for assistive tech", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "Alpha" });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 1, value: "Beta" });
    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 1 });

    const viewport = container.querySelector<HTMLElement>(".excelsior-viewport");
    const activeCell = container.querySelector<HTMLElement>("[data-row='0'][data-col='1']");
    const selectedCell = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");
    const announcer = container.querySelector<HTMLElement>(".excelsior-visually-hidden");
    const columnHeader = container.querySelector<HTMLElement>(".excelsior-column-header");
    const tablist = container.querySelector<HTMLElement>(".excelsior-sheet-tabs");
    const activeTab = container.querySelector<HTMLElement>(".excelsior-sheet-tab.is-active");

    expect(viewport?.getAttribute("role")).toBe("grid");
    expect(viewport?.getAttribute("aria-activedescendant")).toBe(activeCell?.id ?? null);
    expect(viewport?.getAttribute("aria-rowcount")).toBe(String(sheet.rowCount));
    expect(viewport?.getAttribute("aria-colcount")).toBe(String(sheet.columnCount));
    expect(activeCell?.getAttribute("role")).toBe("gridcell");
    expect(activeCell?.getAttribute("aria-selected")).toBe("true");
    expect(activeCell?.getAttribute("aria-colindex")).toBe("2");
    expect(activeCell?.getAttribute("aria-label")).toContain("Linha 1");
    expect(activeCell?.getAttribute("aria-label")).toContain("Coluna B");
    expect(selectedCell?.getAttribute("aria-selected")).toBe("true");
    expect(announcer?.textContent).toContain("Célula ativa");
    expect(announcer?.textContent).toContain("Linha 1");
    expect(announcer?.textContent).toContain("Coluna B");
    expect(columnHeader?.getAttribute("role")).toBe("columnheader");
    expect(columnHeader?.getAttribute("aria-label")).toBe("Coluna A");
    expect(tablist?.getAttribute("role")).toBe("tablist");
    expect(activeTab?.getAttribute("role")).toBe("tab");
    expect(activeTab?.getAttribute("aria-selected")).toBe("true");

    renderer.dispose();
    container.remove();
  });

  it("validates paste in batch before applying any cell update", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();
    const viewport = container.querySelector<HTMLElement>(".excelsior-viewport");

    engine.setCellValidation({
      sheetId: sheet.id,
      row: 0,
      col: 1,
      validation: {
        rules: [{ type: "number", min: 1, max: 5 }]
      }
    });
    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });

    const pasteEvent = new Event("paste", { bubbles: true, cancelable: true });
    Object.defineProperty(pasteEvent, "clipboardData", {
      value: {
        getData: (type: string) => (type === "text/plain" ? "A\tabc" : "")
      }
    });

    viewport!.dispatchEvent(pasteEvent);

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("");
    expect(container.querySelector<HTMLElement>(".excelsior-status-message")?.textContent).toContain("numérico");

    renderer.dispose();
    container.remove();
  });

  it("blocks paste payloads that exceed the configured maxPasteCells", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        settings: {
          maxPasteCells: 2
        }
      },
      new BasicFormulaEngine()
    );
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();
    const viewport = container.querySelector<HTMLElement>(".excelsior-viewport");

    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });

    const pasteEvent = new Event("paste", { bubbles: true, cancelable: true });
    Object.defineProperty(pasteEvent, "clipboardData", {
      value: {
        getData: (type: string) => (type === "text/plain" ? "A\tB\n1\t2" : "")
      }
    });

    viewport!.dispatchEvent(pasteEvent);

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("");

    renderer.dispose();
    container.remove();
  });

  it("applies conditional formatting styles to rendered cells", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 15 });
    engine.setConditionalFormattingRules(sheet.id, [
      {
        id: "positive",
        type: "greaterThan",
        range: {
          start: { row: 0, col: 0 },
          end: { row: 0, col: 0 }
        },
        value: 10,
        priority: 10,
        style: {
          backgroundColor: "#dcfce7",
          fontWeight: "bold"
        }
      }
    ]);

    const cell = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");

    expect(cell?.style.backgroundColor).toBe("rgb(220, 252, 231)");
    expect(cell?.style.fontWeight).toBe("bold");

    renderer.dispose();
    container.remove();
  });

  it("renders dropdown affordance and commits selection from list validation UI", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    engine.setCellValidation({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      validation: {
        rules: [{ type: "dropdown", values: ["Backlog", "Doing", "Done"] }]
      }
    });

    const cell = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");
    expect(cell?.querySelector(".excelsior-cell-affordance")?.textContent).toBe("▾");

    cell?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));

    const select = container.querySelector<HTMLSelectElement>(".excelsior-select-editor");
    expect(select?.hidden).toBe(false);
    expect(select?.options).toHaveLength(3);

    select!.selectedIndex = 1;
    select!.dispatchEvent(new Event("change", { bubbles: true }));

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("Doing");

    renderer.dispose();
    container.remove();
  });

  it("renders checkbox affordance and toggles value on edit", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    engine.setCellValidation({
      sheetId: sheet.id,
      row: 0,
      col: 1,
      validation: {
        rules: [{ type: "checkbox" }]
      }
    });

    const cell = container.querySelector<HTMLElement>("[data-row='0'][data-col='1']");
    expect(cell?.querySelector(".excelsior-cell-affordance")?.textContent).toBe("☐");

    cell?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));

    expect(engine.getCell(sheet.id, 0, 1)?.value).toBe(true);
    expect(container.querySelector<HTMLElement>("[data-row='0'][data-col='1'] .excelsior-cell-affordance")?.textContent).toBe("☑");

    renderer.dispose();
    container.remove();
  });

  it("renders custom cell output with safe fallback when the renderer matches", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 88 });

    const renderer = new DomSpreadsheetRenderer(container, engine, {
      cellRenderers: [
        {
          id: "score-pill",
          matches: ({ row, col }) => row === 0 && col === 0,
          render: ({ displayValue }) => ({
            parts: [
              { text: "Score", tone: "muted" },
              { text: displayValue, tone: "accent" }
            ],
            accessoryText: "pts",
            classNames: ["custom-score-cell"]
          })
        }
      ]
    });

    const cell = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");
    expect(cell?.classList.contains("custom-score-cell")).toBe(true);
    expect(cell?.textContent).toContain("Score");
    expect(cell?.textContent).toContain("88");
    expect(cell?.textContent).toContain("pts");

    renderer.dispose();
    container.remove();
  });

  it("renders custom renderer text literally and drops unsafe class names", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "seed" });

    const renderer = new DomSpreadsheetRenderer(container, engine, {
      cellRenderers: [
        {
          id: "safe-html-proof",
          matches: ({ row, col }) => row === 0 && col === 0,
          render: () => ({
            text: "<img src=x onerror=alert(1)>",
            accessoryText: "<svg/onload=alert(1)>",
            classNames: ["safe-pill", "bad class", 'x\"onclick=1'],
            ariaLabel: "<b>danger</b>",
            title: "<script>alert(1)</script>"
          })
        }
      ]
    });

    const cell = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");
    expect(cell?.querySelector("img")).toBeNull();
    expect(cell?.querySelector("svg")).toBeNull();
    expect(cell?.textContent).toContain("<img src=x onerror=alert(1)>");
    expect(cell?.textContent).toContain("<svg/onload=alert(1)>");
    expect(cell?.classList.contains("safe-pill")).toBe(true);
    expect(cell?.className).not.toContain("bad class");
    expect(cell?.className).not.toContain("onclick");
    expect(cell?.getAttribute("aria-label")).toBe("<b>danger</b>");
    expect(cell?.title).toBe("<script>alert(1)</script>");

    renderer.dispose();
    container.remove();
  });

  it("falls back safely when a custom renderer throws", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "plain" });

    const renderer = new DomSpreadsheetRenderer(container, engine, {
      cellRenderers: [
        {
          id: "broken",
          matches: () => true,
          render: () => {
            throw new Error("boom");
          }
        }
      ]
    });

    expect(container.querySelector<HTMLElement>("[data-row='0'][data-col='0']")?.textContent).toBe("plain");

    renderer.dispose();
    container.remove();
  });

  it("commits value through a custom editor and preserves validation on failure", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    engine.setCellValidation({
      sheetId: sheet.id,
      row: 0,
      col: 0,
      validation: {
        rules: [{ type: "number", min: 10, max: 20 }]
      }
    });

    const renderer = new DomSpreadsheetRenderer(container, engine, {
      cellEditors: [
        {
          id: "range-editor",
          matches: ({ row, col }) => row === 0 && col === 0,
          create: () => {
            let input: HTMLInputElement | undefined;
            return {
              mount(host) {
                input = document.createElement("input");
                input.value = "abc";
                host.replaceChildren(input);
              },
              getValue() {
                return input?.value ?? "";
              },
              focus() {
                input?.focus();
              }
            };
          }
        }
      ]
    });

    const cell = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");
    cell?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));

    const customInput = container.querySelector<HTMLInputElement>(".excelsior-custom-editor-host input");
    expect(customInput).not.toBeNull();
    customInput?.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("");
    expect(container.querySelector<HTMLElement>(".excelsior-status-message")?.textContent).toContain("numérico");

    customInput!.value = "15";
    customInput?.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("15");

    renderer.dispose();
    container.remove();
  });

  it("keeps frozen rows and columns visible while scrolling", () => {
    const container = document.createElement("div");
    container.style.width = "320px";
    container.style.height = "240px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "TL" });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 3, value: "Top" });
    engine.setCellValue({ sheetId: sheet.id, row: 4, col: 0, value: "Left" });
    engine.freezeRows(sheet.id, 1);
    engine.freezeColumns(sheet.id, 1);

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const viewport = container.querySelector<HTMLElement>(".excelsior-viewport");
    viewport!.scrollTop = 120;
    viewport!.scrollLeft = 180;
    viewport!.dispatchEvent(new Event("scroll"));

    const topLeft = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");
    const topFrozen = container.querySelector<HTMLElement>("[data-row='0'][data-col='3']");
    const leftFrozen = container.querySelector<HTMLElement>("[data-row='4'][data-col='0']");

    expect(topLeft?.classList.contains("is-frozen")).toBe(true);
    expect(topFrozen?.classList.contains("is-frozen-row")).toBe(true);
    expect(leftFrozen?.classList.contains("is-frozen-column")).toBe(true);
    expect(topFrozen?.style.top).toBe("120px");
    expect(leftFrozen?.style.left).toBe("236px");

    renderer.dispose();
    container.remove();
  });

  it("skips hidden rows and columns during keyboard navigation and clipboard copy by default", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const viewport = container.querySelector<HTMLElement>(".excelsior-viewport");

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "A" });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 1, value: "B" });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 2, value: "C" });
    engine.setColumnsHidden(sheet.id, 1, 1, true);
    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });

    viewport!.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));

    expect(engine.getSelection(sheet.id)).toEqual({
      start: { row: 0, col: 2 },
      end: { row: 0, col: 2 }
    });

    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 2 });
    const clipboard = new Map<string, string>();
    const copyEvent = new Event("copy", { bubbles: true, cancelable: true });
    Object.defineProperty(copyEvent, "clipboardData", {
      value: {
        setData: (type: string, value: string) => {
          clipboard.set(type, value);
        }
      }
    });

    viewport!.dispatchEvent(copyEvent);

    expect(clipboard.get("text/plain")).toBe("A\tC");

    renderer.dispose();
    container.remove();
  });

  it("fills numeric sequences vertically with batch operations from the fill handle", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const onChange = vi.fn<(operations: SpreadsheetOperation[]) => void>();
    const renderer = new DomSpreadsheetRenderer(container, engine, { onChange });

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 1 });
    engine.setCellValue({ sheetId: sheet.id, row: 1, col: 0, value: 2 });
    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 0 });
    onChange.mockClear();

    const handle = container.querySelector<HTMLElement>(".excelsior-fill-handle");
    const target = container.querySelector<HTMLElement>("[data-row='3'][data-col='0']");

    handle?.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    target?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
    globalThis.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));

    expect(engine.getDisplayValue(sheet.id, 2, 0)).toBe("3");
    expect(engine.getDisplayValue(sheet.id, 3, 0)).toBe("4");
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0]?.[0]).toHaveLength(3);
    expect(onChange.mock.calls[0]?.[0].at(-1)).toMatchObject({
      op: "replace",
      path: ["selection"]
    });

    renderer.dispose();
    container.remove();
  });

  it("copies formulas with relative and absolute reference adjustments from the fill handle", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const renderer = new DomSpreadsheetRenderer(container, engine);

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 1 });
    engine.setCellValue({ sheetId: sheet.id, row: 1, col: 0, value: 2 });
    engine.setCellValue({ sheetId: sheet.id, row: 2, col: 0, value: 3 });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 1, value: "=A1+$A$1+A$1+$A1" });
    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 1, colEnd: 1 });

    const handle = container.querySelector<HTMLElement>(".excelsior-fill-handle");
    const target = container.querySelector<HTMLElement>("[data-row='2'][data-col='1']");

    handle?.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    target?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
    globalThis.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));

    expect(engine.getCell(sheet.id, 1, 1)?.formula).toBe("=A2+$A$1+A$1+$A2");
    expect(engine.getCell(sheet.id, 2, 1)?.formula).toBe("=A3+$A$1+A$1+$A3");
    expect(engine.getDisplayValue(sheet.id, 2, 1)).toBe("8");

    renderer.dispose();
    container.remove();
  });

  it("supports localization formatters, custom shortcuts and RTL horizontal navigation", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const renderer = new DomSpreadsheetRenderer(container, engine, {
      localization: {
        direction: "rtl",
        messages: {
          findReplace: "Localizar"
        },
        shortcuts: {
          openFindReplace: ["Alt+F"]
        },
        formatters: {
          number: (value) => `num:${value.toFixed(1)}`,
          date: (value) => `date:${value}`
        }
      }
    });

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 12.5 });
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 1, value: "2026-05-26" });
    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 1, colEnd: 1 });

    const numberCell = container.querySelector<HTMLElement>("[data-row='0'][data-col='0']");
    const dateCell = container.querySelector<HTMLElement>("[data-row='0'][data-col='1']");
    const viewport = container.querySelector<HTMLElement>(".excelsior-viewport");

    expect(container.querySelector<HTMLElement>(".excelsior-shell")?.getAttribute("dir")).toBe("rtl");
    expect(container.querySelector<HTMLButtonElement>("[data-action='find-replace']")?.textContent).toBe("Localizar");
    expect(numberCell?.textContent).toBe("num:12.5");
    expect(dateCell?.textContent).toBe("date:2026-05-26");

    viewport?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", altKey: false, bubbles: true }));
    expect(engine.getSelection(sheet.id)).toEqual({
      start: { row: 0, col: 0 },
      end: { row: 0, col: 0 }
    });

    viewport?.dispatchEvent(new KeyboardEvent("keydown", { key: "f", altKey: true, bubbles: true }));
    expect(container.querySelector<HTMLElement>(".excelsior-find-replace")?.hidden).toBe(false);

    renderer.dispose();
    container.remove();
  });

  it("renders visible rows through the configured row model strategy", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "120px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const requests: Array<{ startRow: number; endRow: number }> = [];
    engine.setRowModel(
      sheet.id,
      new ViewportRowModel({
        rowCount: sheet.rowCount,
        getRows: ({ startRow, endRow, sheetId }) => {
          requests.push({ startRow, endRow });
          return {
            rowCount: sheet.rowCount,
            rows: [{ index: 0 }, { index: Math.min(2, endRow) }]
          };
        }
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    expect(requests.length).toBeGreaterThan(0);
    expect(container.querySelector("[data-row='0'][data-col='0']")).not.toBeNull();
    expect(container.querySelector("[data-row='1'][data-col='0']")).toBeNull();
    expect(container.querySelector("[data-row='2'][data-col='0']")).not.toBeNull();

    renderer.dispose();
    container.remove();
  });

  it("reuses the cached row model window while the viewport stays unchanged", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "120px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const getRows = vi.fn(({ startRow, endRow }: { startRow: number; endRow: number }) => ({
      rowCount: sheet.rowCount,
      rows: Array.from({ length: endRow - startRow + 1 }, (_value, index) => ({
        index: startRow + index
      }))
    }));

    engine.setRowModel(
      sheet.id,
      new ViewportRowModel({
        rowCount: sheet.rowCount,
        getRows
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const callCountAfterMount = getRows.mock.calls.length;

    renderer.render();

    expect(getRows.mock.calls.length).toBe(callCountAfterMount);

    renderer.dispose();
    container.remove();
  });

  it("shows loading feedback while an async row model request is pending", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "120px";
    document.body.append(container);

    const deferred = createDeferred<{ rowCount: number; rows: Array<{ index: number }> }>();
    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    engine.setRowModel(
      sheet.id,
      new ViewportRowModel({
        rowCount: sheet.rowCount,
        getRows: () => deferred.promise
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const status = container.querySelector<HTMLElement>(".excelsior-status-message");

    expect(status?.textContent).toBe("Carregando linhas...");
    expect(status?.classList.contains("is-error")).toBe(false);

    deferred.resolve({
      rowCount: sheet.rowCount,
      rows: [{ index: 0 }, { index: 1 }]
    });

    await vi.waitFor(() => {
      expect(status?.textContent).toBe("");
    });

    renderer.dispose();
    container.remove();
  });

  it("shows row model load failures in the status region", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "120px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    engine.setRowModel(
      sheet.id,
      new ViewportRowModel({
        rowCount: sheet.rowCount,
        getRows: () => Promise.reject(new Error("backend offline"))
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);
    const status = container.querySelector<HTMLElement>(".excelsior-status-message");

    await vi.waitFor(() => {
      expect(status?.textContent).toBe("backend offline");
    });
    expect(status?.classList.contains("is-error")).toBe(true);

    renderer.dispose();
    container.remove();
  });

  it("refetches the visible window when the remote row model request model changes", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "120px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const dataSource = {
      getRows: vi.fn(async ({ startRow, endRow, sortModel, filterModel }) => ({
        totalRows: sheet.rowCount,
        rows: Array.from({ length: endRow - startRow + 1 }, (_value, index) => ({
          index: startRow + index
        })),
        sortModel,
        filterModel
      }))
    };

    engine.setRowModel(
      sheet.id,
      new ServerSideRowModel({
        rowCount: sheet.rowCount,
        dataSource
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(1);
    });

    engine.updateRemoteRowModel(sheet.id, {
      sortModel: [{ field: "price", direction: "asc" }],
      filterModel: {
        status: { operator: "equals", value: "draft" }
      }
    });

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(2);
    });
    expect(dataSource.getRows.mock.calls[1][0]).toMatchObject({
      sortModel: [{ field: "price", direction: "asc" }],
      filterModel: {
        status: { operator: "equals", value: "draft" }
      }
    });

    renderer.dispose();
    container.remove();
  });

  it("updates sort and filter for the active column through the toolbar on remote row models", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "120px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const dataSource = {
      getRows: vi.fn(async ({ startRow, endRow, sortModel, filterModel }) => ({
        totalRows: sheet.rowCount,
        rows: Array.from({ length: endRow - startRow + 1 }, (_value, index) => ({
          index: startRow + index
        })),
        sortModel,
        filterModel
      }))
    };

    engine.setRowModel(
      sheet.id,
      new ServerSideRowModel({
        rowCount: sheet.rowCount,
        dataSource
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(1);
    });

    container.querySelector<HTMLButtonElement>("[data-action='sort-desc']")?.click();

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(2);
    });
    expect(dataSource.getRows.mock.calls[1][0]).toMatchObject({
      sortModel: [{ field: "A", direction: "desc" }]
    });

    const filterInput = container.querySelector<HTMLInputElement>("[data-remote-filter-input]");
    filterInput!.value = "draft";
    filterInput?.dispatchEvent(new Event("input", { bubbles: true }));
    filterInput?.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(3);
    });
    expect(dataSource.getRows.mock.calls[2][0]).toMatchObject({
      sortModel: [{ field: "A", direction: "desc" }],
      filterModel: {
        A: { operator: "equals", value: "draft" }
      }
    });

    renderer.dispose();
    container.remove();
  });

  it("supports explicit server-side pivot materialization from the panel", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({
      pivotModule: false,
      data: [
        {
          name: "Remote Sales",
          rowCount: 1000,
          columnCount: 5,
          cells: {
            "0:0": { value: "Region", computedValue: "Region" },
            "0:1": { value: "Quarter", computedValue: "Quarter" },
            "0:2": { value: "Revenue", computedValue: "Revenue" },
            "1:0": { value: "North", computedValue: "North" },
            "1:1": { value: "Q1", computedValue: "Q1" },
            "1:2": { value: 42, computedValue: 42 }
          },
          merges: [],
          columns: {},
          rows: {}
        }
      ]
    });
    const sourceSheet = engine.getActiveSheet();
    const getRows = vi.fn(async (request: { requestKind?: string }) => {
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

      return { totalRows: 1000, rows: [{ index: 0 }] };
    });

    engine.setRowModel(
      sourceSheet.id,
      new ServerSideRowModel({
        rowCount: "unknown",
        dataSource: { getRows }
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    container.querySelector<HTMLButtonElement>("[data-action='create-pivot']")?.click();
    setPivotSelectValue(container.querySelector<HTMLSelectElement>("[data-pivot-role='execution-mode']"), "server");
    setPivotToggle(container.querySelector<HTMLInputElement>("[data-pivot-role='auto-refresh']"), false);
    container.querySelector<HTMLButtonElement>("[data-pivot-action='apply']")?.click();
    await waitForActiveSheetName(engine, "Remote Sales Pivot");

    const activeSheet = engine.getActiveSheet();
    expect(activeSheet.name).toBe("Remote Sales Pivot");
    expect(engine.getPivotSheetViewDefinition(activeSheet.id)).toMatchObject({
      autoRefresh: false,
      result: {
        executionMode: "server",
        remote: true
      }
    });
    expect(getRows).toHaveBeenCalledWith(expect.objectContaining({ requestKind: "pivotSheet" }), expect.any(Object));

    renderer.dispose();
    container.remove();
  });

  it("reopens the pivot panel for a persisted pivot and updates it in place", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

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

    const renderer = new DomSpreadsheetRenderer(container, engine);

    container.querySelector<HTMLButtonElement>("[data-action='create-pivot']")?.click();

    expect(container.querySelector<HTMLSelectElement>("[data-pivot-role='execution-mode']")?.value).toBe("auto");
    expect(container.querySelector<HTMLInputElement>("[data-pivot-role='auto-refresh']")?.checked).toBe(false);

    selectPivotOptions(container.querySelector<HTMLSelectElement>("[data-pivot-role='row']"), ["Quarter"]);
    selectPivotOptions(container.querySelector<HTMLSelectElement>("[data-pivot-role='column']"), []);
    setPivotInputValue(
      container.querySelector<HTMLInputElement>("[data-pivot-role='value-alias'][data-pivot-value-index='0']"),
      "Revenue"
    );
    setPivotToggle(container.querySelector<HTMLInputElement>("[data-pivot-role='auto-refresh']"), true);
    container.querySelector<HTMLButtonElement>("[data-pivot-action='apply']")?.click();

    for (let attempt = 0; attempt < 10; attempt += 1) {
      if (engine.getDisplayValue(pivotSheetId, 0, 0) === "Quarter") {
        break;
      }

      await new Promise((resolve) => globalThis.setTimeout(resolve, 0));
    }

    expect(engine.getActiveSheet().id).toBe(pivotSheetId);
    expect(engine.getDisplayValue(pivotSheetId, 0, 0)).toBe("Quarter");
    expect(engine.getDisplayValue(pivotSheetId, 0, 1)).toBe("Revenue");
    expect(engine.getPivotSheetViewDefinition(pivotSheetId)).toMatchObject({
      autoRefresh: true,
      input: {
        rows: ["Quarter"],
        values: [{ field: "Sales", aggregate: "sum", as: "Revenue" }]
      }
    });

    renderer.dispose();
    container.remove();
  });

  it("shows stale status for persisted pivots with auto-refresh disabled", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

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

    const renderer = new DomSpreadsheetRenderer(container, engine);

    engine.setCellValue({ sheetId: sourceSheet.id, row: 1, col: 2, value: 15 });
    await new Promise((resolve) => globalThis.setTimeout(resolve, 0));

    expect(container.querySelector<HTMLElement>(".excelsior-status-message")?.textContent).toContain(
      "Pivot derivada desatualizada"
    );

    renderer.dispose();
    container.remove();
  });

  it("updates grouping, pivoting and aggregations for the active column through the toolbar on remote row models", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "120px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const dataSource = {
      getRows: vi.fn(async ({ startRow, endRow, groupKeys, pivotModel, aggregateModel }) => ({
        totalRows: sheet.rowCount,
        rows: Array.from({ length: endRow - startRow + 1 }, (_value, index) => ({
          index: startRow + index
        })),
        groupKeys,
        pivotModel,
        aggregateModel
      }))
    };

    engine.setRowModel(
      sheet.id,
      new ServerSideRowModel({
        rowCount: sheet.rowCount,
        dataSource
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(1);
    });

    container.querySelector<HTMLButtonElement>("[data-action='group-column']")?.click();

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(2);
    });
    expect(dataSource.getRows.mock.calls[1][0]).toMatchObject({
      groupKeys: ["A"]
    });
    expect(container.querySelector("[data-action='group-column']")?.getAttribute("aria-pressed")).toBe("true");

    container.querySelector<HTMLButtonElement>("[data-action='pivot-column']")?.click();

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(3);
    });
    expect(dataSource.getRows.mock.calls[2][0]).toMatchObject({
      groupKeys: ["A"],
      pivotModel: [{ field: "A" }]
    });
    expect(container.querySelector("[data-action='pivot-column']")?.getAttribute("aria-pressed")).toBe("true");

    container.querySelector<HTMLButtonElement>("[data-action='aggregate-sum']")?.click();

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(4);
    });
    expect(dataSource.getRows.mock.calls[3][0]).toMatchObject({
      groupKeys: ["A"],
      pivotModel: [{ field: "A" }],
      aggregateModel: [{ field: "A", function: "sum", as: "A_sum" }]
    });
    expect(container.querySelector("[data-action='aggregate-sum']")?.getAttribute("aria-pressed")).toBe("true");

    container.querySelector<HTMLButtonElement>("[data-action='aggregate-avg']")?.click();
    container.querySelector<HTMLButtonElement>("[data-action='aggregate-min']")?.click();
    container.querySelector<HTMLButtonElement>("[data-action='aggregate-max']")?.click();
    container.querySelector<HTMLButtonElement>("[data-action='aggregate-count']")?.click();

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(8);
    });
    expect(dataSource.getRows.mock.calls[7][0]).toMatchObject({
      groupKeys: ["A"],
      pivotModel: [{ field: "A" }],
      aggregateModel: [
        { field: "A", function: "sum", as: "A_sum" },
        { field: "A", function: "avg", as: "A_avg" },
        { field: "A", function: "min", as: "A_min" },
        { field: "A", function: "max", as: "A_max" },
        { field: "A", function: "count", as: "A_count" }
      ]
    });
    expect(container.querySelector("[data-action='aggregate-avg']")?.getAttribute("aria-pressed")).toBe("true");
    expect(container.querySelector("[data-action='aggregate-min']")?.getAttribute("aria-pressed")).toBe("true");
    expect(container.querySelector("[data-action='aggregate-max']")?.getAttribute("aria-pressed")).toBe("true");
    expect(container.querySelector("[data-action='aggregate-count']")?.getAttribute("aria-pressed")).toBe("true");

    renderer.dispose();
    container.remove();
  });

  it("cycles remote sort from the column header and updates the active column", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "120px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const dataSource = {
      getRows: vi.fn(async ({ startRow, endRow, sortModel }) => ({
        totalRows: sheet.rowCount,
        rows: Array.from({ length: endRow - startRow + 1 }, (_value, index) => ({
          index: startRow + index
        })),
        sortModel
      }))
    };

    engine.setRowModel(
      sheet.id,
      new ServerSideRowModel({
        rowCount: sheet.rowCount,
        dataSource
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(1);
    });

    const headerB = container.querySelector<HTMLElement>("[data-column-header-col='1']");
    const initialCallCount = dataSource.getRows.mock.calls.length;
    headerB?.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    await vi.waitFor(() => {
      expect(dataSource.getRows.mock.calls.length).toBeGreaterThan(initialCallCount);
      expect(
        dataSource.getRows.mock.calls.some(([request]) =>
          expect.objectContaining({
            sortModel: [{ field: "B", direction: "asc" }]
          }).asymmetricMatch(request)
        )
      ).toBe(true);
    });
    expect(engine.getSelection(sheet.id)).toEqual({
      start: { row: 0, col: 1 },
      end: { row: 0, col: 1 }
    });
    expect(container.querySelector("[data-column-header-col='1']")?.getAttribute("aria-sort")).toBe("ascending");

    const headerBAfterClick = container.querySelector<HTMLElement>("[data-column-header-col='1']");
    const callCountAfterClick = dataSource.getRows.mock.calls.length;
    headerBAfterClick?.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));

    await vi.waitFor(() => {
      expect(dataSource.getRows.mock.calls.length).toBeGreaterThan(callCountAfterClick);
      expect(
        dataSource.getRows.mock.calls.some(([request]) =>
          expect.objectContaining({
            sortModel: [{ field: "B", direction: "desc" }]
          }).asymmetricMatch(request)
        )
      ).toBe(true);
    });
    expect(container.querySelector("[data-column-header-col='1']")?.getAttribute("aria-sort")).toBe("descending");

    renderer.dispose();
    container.remove();
  });

  it("toggles remote group expansion from grouped remote rows", async () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "120px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "Engineering" });

    const dataSource = {
      getRows: vi.fn(async ({ startRow, endRow, expandedGroupPaths }) => ({
        totalRows: sheet.rowCount,
        rows: Array.from({ length: endRow - startRow + 1 }, (_value, index) => ({
          index: startRow + index
        })),
        groupInfo:
          startRow === 0
            ? [
                {
                  key: "engineering",
                  path: ["engineering"],
                  level: 0,
                  childCount: 12,
                  expanded:
                    expandedGroupPaths?.some((path: string[]) => path.length === 1 && path[0] === "engineering") ?? false
                }
              ]
            : undefined,
        expandedGroupPaths
      }))
    };

    engine.setRowModel(
      sheet.id,
      new ServerSideRowModel({
        rowCount: sheet.rowCount,
        dataSource
      })
    );

    const renderer = new DomSpreadsheetRenderer(container, engine);

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(1);
    });

    await vi.waitFor(() => {
      expect(container.querySelector("[data-remote-group-toggle='true']")).not.toBeNull();
    });

    container.querySelector<HTMLButtonElement>("[data-remote-group-toggle='true']")?.click();

    await vi.waitFor(() => {
      expect(dataSource.getRows).toHaveBeenCalledTimes(2);
    });
    expect(dataSource.getRows.mock.calls[1][0]).toMatchObject({
      expandedGroupPaths: [["engineering"]]
    });

    await vi.waitFor(() => {
      expect(container.querySelector("[data-remote-group-toggle='true']")?.getAttribute("aria-expanded")).toBe("true");
    });

    renderer.dispose();
    container.remove();
  });

  it("does not commit text editing while IME composition is active", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const viewport = container.querySelector<HTMLElement>(".excelsior-viewport");
    const editor = container.querySelector<HTMLInputElement>(".excelsior-editor");

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "antes" });
    engine.selectRange({ sheetId: sheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });

    viewport?.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    editor!.value = "durante-ime";
    editor?.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true, data: "du" }));
    editor?.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("antes");
    expect(editor?.hidden).toBe(false);

    editor?.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true, data: "durante-ime" }));
    editor?.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));

    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("durante-ime");

    renderer.dispose();
    container.remove();
  });

  it("finds matches in the current sheet with navigation and search options", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const sheet = engine.getActiveSheet();
    const renderer = new DomSpreadsheetRenderer(container, engine);

    try {
      engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "Task" });
      engine.setCellValue({ sheetId: sheet.id, row: 1, col: 0, value: "task" });
      engine.setCellValue({ sheetId: sheet.id, row: 2, col: 0, value: "Task - next" });

      container.querySelector<HTMLButtonElement>("[data-action='find-replace']")?.click();
      const [findInput] = Array.from(container.querySelectorAll<HTMLInputElement>(".excelsior-find-replace-input"));
      findInput!.value = "task";
      findInput!.dispatchEvent(new Event("input", { bubbles: true }));
      expect(container.querySelector<HTMLElement>(".excelsior-find-replace-results")?.textContent).toBe("1/3");

      expect(engine.getSelection(sheet.id)).toEqual({
        start: { row: 0, col: 0 },
        end: { row: 0, col: 0 }
      });

      renderer.focus();
      container.querySelector<HTMLElement>(".excelsior-viewport")?.dispatchEvent(
        new KeyboardEvent("keydown", { key: "F3", bubbles: true })
      );
      expect(engine.getSelection(sheet.id)).toEqual({
        start: { row: 1, col: 0 },
        end: { row: 1, col: 0 }
      });

      const caseToggle = container.querySelector<HTMLInputElement>("[data-find-role='case-sensitive']");
      const wholeToggle = container.querySelector<HTMLInputElement>("[data-find-role='whole-cell']");
      findInput!.value = "Task";
      findInput!.dispatchEvent(new Event("input", { bubbles: true }));
      caseToggle!.checked = true;
      caseToggle!.dispatchEvent(new Event("input", { bubbles: true }));
      wholeToggle!.checked = true;
      wholeToggle!.dispatchEvent(new Event("input", { bubbles: true }));
      expect(container.querySelector<HTMLElement>(".excelsior-find-replace-results")?.textContent).toBe("1/1");

      expect(engine.getSelection(sheet.id)).toEqual({
        start: { row: 0, col: 0 },
        end: { row: 0, col: 0 }
      });
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("replaces the current match and replaces all workbook matches with undoable batch history", () => {
    const container = document.createElement("div");
    container.style.width = "800px";
    container.style.height = "480px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [{ name: "One" }, { name: "Two" }]
      },
      new BasicFormulaEngine()
    );
    const firstSheet = engine.getSnapshot().sheets[0]!;
    const secondSheet = engine.getSnapshot().sheets[1]!;
    const renderer = new DomSpreadsheetRenderer(container, engine);

    try {
      engine.setCellValue({ sheetId: firstSheet.id, row: 0, col: 0, value: "Todo" });
      engine.setCellValue({ sheetId: firstSheet.id, row: 1, col: 0, value: "Todo" });
      engine.setCellValue({ sheetId: secondSheet.id, row: 0, col: 0, value: "Todo" });

      container.querySelector<HTMLButtonElement>("[data-action='find-replace']")?.click();
      const [findInput, replaceInput] = Array.from(container.querySelectorAll<HTMLInputElement>(".excelsior-find-replace-input"));
      const workbookToggle = container.querySelector<HTMLInputElement>("[data-find-role='scope-workbook']");

      workbookToggle!.checked = true;
      workbookToggle!.dispatchEvent(new Event("input", { bubbles: true }));
      findInput!.value = "Todo";
      findInput!.dispatchEvent(new Event("input", { bubbles: true }));
      replaceInput!.value = "Done";
      replaceInput!.dispatchEvent(new Event("input", { bubbles: true }));
      expect(container.querySelector<HTMLElement>(".excelsior-find-replace-results")?.textContent).toBe("1/3");

      container.querySelector<HTMLButtonElement>("[data-find-action='replace']")?.click();
      expect(engine.getDisplayValue(firstSheet.id, 0, 0)).toBe("Done");

      expect(engine.getDisplayValue(firstSheet.id, 0, 0)).toBe("Done");
      expect(engine.getDisplayValue(firstSheet.id, 1, 0)).toBe("Todo");
      expect(engine.getDisplayValue(secondSheet.id, 0, 0)).toBe("Todo");

      container.querySelector<HTMLButtonElement>("[data-find-action='replace-all']")?.click();
      expect(engine.getDisplayValue(firstSheet.id, 1, 0)).toBe("Done");
      expect(engine.getDisplayValue(secondSheet.id, 0, 0)).toBe("Done");

      expect(engine.getDisplayValue(firstSheet.id, 1, 0)).toBe("Done");
      expect(engine.getDisplayValue(secondSheet.id, 0, 0)).toBe("Done");

      expect(engine.undo()).toBe(true);
      expect(engine.getDisplayValue(firstSheet.id, 0, 0)).toBe("Done");
      expect(engine.getDisplayValue(firstSheet.id, 1, 0)).toBe("Todo");
      expect(engine.getDisplayValue(secondSheet.id, 0, 0)).toBe("Todo");
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("creates a chart object from the selected range using the toolbar", () => {
    const container = document.createElement("div");
    container.style.width = "1100px";
    container.style.height = "620px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Sales",
            rowCount: 8,
            columnCount: 4,
            cells: {
              "0:0": { value: "Month", computedValue: "Month" },
              "0:1": { value: "North", computedValue: "North" },
              "0:2": { value: "South", computedValue: "South" },
              "1:0": { value: "Jan", computedValue: "Jan" },
              "1:1": { value: 10, computedValue: 10 },
              "1:2": { value: 8, computedValue: 8 },
              "2:0": { value: "Feb", computedValue: "Feb" },
              "2:1": { value: 14, computedValue: 14 },
              "2:2": { value: 9, computedValue: 9 },
              "3:0": { value: "Mar", computedValue: "Mar" },
              "3:1": { value: 18, computedValue: 18 },
              "3:2": { value: 12, computedValue: 12 }
            },
            merges: [],
            columns: {},
            rows: {}
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    try {
      engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        colStart: 0,
        rowEnd: 3,
        colEnd: 2
      });
      container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.click();

      const charts = engine.getCharts(sheet.id);
      expect(charts).toHaveLength(1);
      expect(charts[0]?.type).toBe("line");
      expect(charts[0]?.sourceRange?.rangeAddress).toBe("A1:C4");
      expect(container.querySelector<HTMLElement>(`[data-chart-id='${charts[0]?.id ?? ""}']`)).not.toBeNull();
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("recomputes the chart figure when source cells are updated", () => {
    const container = document.createElement("div");
    container.style.width = "1100px";
    container.style.height = "620px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Forecast",
            rowCount: 8,
            columnCount: 4,
            cells: {
              "0:0": { value: "Month", computedValue: "Month" },
              "0:1": { value: "Value", computedValue: "Value" },
              "1:0": { value: "Jan", computedValue: "Jan" },
              "1:1": { value: 100, computedValue: 100 },
              "2:0": { value: "Feb", computedValue: "Feb" },
              "2:1": { value: 120, computedValue: 120 },
              "3:0": { value: "Mar", computedValue: "Mar" },
              "3:1": { value: 130, computedValue: 130 }
            },
            merges: [],
            columns: {},
            rows: {}
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    try {
      engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        colStart: 0,
        rowEnd: 3,
        colEnd: 1
      });
      container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.click();
      const chart = engine.getCharts(sheet.id)[0]!;
      const originalTrace = chart.figure.data[0] as { y?: unknown };
      expect(Array.isArray(originalTrace.y) ? originalTrace.y[0] : undefined).toBe(100);

      engine.setCellValue({
        sheetId: sheet.id,
        row: 1,
        col: 1,
        value: 250
      });

      const refreshed = engine.getChart(sheet.id, chart.id)!;
      const refreshedTrace = refreshed.figure.data[0] as { y?: unknown };
      expect(Array.isArray(refreshedTrace.y) ? refreshedTrace.y[0] : undefined).toBe(250);
      expect(refreshed.state.lastRenderedAt).toBeTypeOf("number");
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("supports moving, resizing and deleting chart objects from the sheet", () => {
    const container = document.createElement("div");
    container.style.width = "1100px";
    container.style.height = "620px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Ops",
            rowCount: 10,
            columnCount: 4,
            cells: {
              "0:0": { value: "Day", computedValue: "Day" },
              "0:1": { value: "Tickets", computedValue: "Tickets" },
              "1:0": { value: "Mon", computedValue: "Mon" },
              "1:1": { value: 12, computedValue: 12 },
              "2:0": { value: "Tue", computedValue: "Tue" },
              "2:1": { value: 8, computedValue: 8 },
              "3:0": { value: "Wed", computedValue: "Wed" },
              "3:1": { value: 16, computedValue: 16 }
            },
            merges: [],
            columns: {},
            rows: {}
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    try {
      engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        colStart: 0,
        rowEnd: 3,
        colEnd: 1
      });
      container.querySelector<HTMLButtonElement>("[data-action='chart-column']")?.click();
      const chartId = engine.getCharts(sheet.id)[0]?.id as string;
      const chartBefore = engine.getChart(sheet.id, chartId)!;
      const sourceRangeBefore = chartBefore.sourceRange ? JSON.parse(JSON.stringify(chartBefore.sourceRange)) : undefined;

      const chartElement = container.querySelector<HTMLElement>(`[data-chart-id='${chartId}']`);
      dispatchMouse(chartElement?.querySelector("[data-chart-body='true']"), "mousedown", {
        button: 0,
        clientX: 300,
        clientY: 220
      });
      dispatchMouse(globalThis, "mousemove", {
        button: 0,
        clientX: 380,
        clientY: 300
      });
      dispatchMouse(globalThis, "mouseup", {
        button: 0,
        clientX: 380,
        clientY: 300
      });
      const afterBodyDrag = engine.getChart(sheet.id, chartId)!;
      expect(afterBodyDrag.position).toEqual(chartBefore.position);

      dispatchMouse(chartElement?.querySelector(".excelsior-chart-object-header"), "mousedown", {
        button: 0,
        clientX: 280,
        clientY: 180
      });
      dispatchMouse(globalThis, "mousemove", {
        button: 0,
        clientX: 340,
        clientY: 240
      });
      dispatchMouse(globalThis, "mouseup", {
        button: 0,
        clientX: 340,
        clientY: 240
      });

      const moved = engine.getChart(sheet.id, chartId)!;
      expect(
        moved.position.fromCell !== chartBefore.position.fromCell ||
          moved.position.offsetX !== chartBefore.position.offsetX ||
          moved.position.offsetY !== chartBefore.position.offsetY
      ).toBe(true);
      expect(moved.sourceRange).toEqual(sourceRangeBefore);

      const resizeHandle = container.querySelector<HTMLElement>(`[data-chart-id='${chartId}'] [data-chart-resize='true']`);
      const widthBefore = moved.position.width;
      const heightBefore = moved.position.height;
      dispatchMouse(resizeHandle, "mousedown", {
        button: 0,
        clientX: 520,
        clientY: 360
      });
      dispatchMouse(globalThis, "mousemove", {
        button: 0,
        clientX: 600,
        clientY: 420
      });
      dispatchMouse(globalThis, "mouseup", {
        button: 0,
        clientX: 600,
        clientY: 420
      });

      const resized = engine.getChart(sheet.id, chartId)!;
      expect(resized.position.width).toBeGreaterThanOrEqual(widthBefore);
      expect(resized.position.height).toBeGreaterThanOrEqual(heightBefore);
      expect(resized.sourceRange).toEqual(sourceRangeBefore);

      container.querySelector<HTMLButtonElement>(`[data-chart-id='${chartId}'] [data-chart-action='delete']`)?.click();
      expect(engine.getCharts(sheet.id)).toHaveLength(0);
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("normalizes runtime chart dimensions before rendering", () => {
    const container = document.createElement("div");
    container.style.width = "1100px";
    container.style.height = "620px";
    document.body.append(container);

    const engine = new WorkbookEngine({}, new BasicFormulaEngine());
    const renderer = new DomSpreadsheetRenderer(container, engine);

    try {
      const rendererState = renderer as unknown as {
        withRuntimeFigureSize: (figure: { data: unknown[]; layout?: Record<string, unknown> }, width: number, height: number) => {
          data: unknown[];
          layout?: Record<string, unknown>;
        };
      };
      const sized = rendererState.withRuntimeFigureSize(
        {
          data: [],
          layout: {
            title: "Sizing Check"
          }
        },
        459.9,
        279.2
      );
      expect(sized.layout?.title).toBe("Sizing Check");
      expect(sized.layout?.width).toBe(460);
      expect(sized.layout?.height).toBe(279);
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("edits selected charts from the panel and emits selection events", () => {
    const container = document.createElement("div");
    container.style.width = "1100px";
    container.style.height = "620px";
    document.body.append(container);

    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Panel",
            rowCount: 12,
            columnCount: 4,
            cells: {
              "0:0": { value: "Month", computedValue: "Month" },
              "0:1": { value: "Revenue", computedValue: "Revenue" },
              "0:2": { value: "Cost", computedValue: "Cost" },
              "1:0": { value: "Jan", computedValue: "Jan" },
              "1:1": { value: 120, computedValue: 120 },
              "1:2": { value: 95, computedValue: 95 },
              "2:0": { value: "Feb", computedValue: "Feb" },
              "2:1": { value: 180, computedValue: 180 },
              "2:2": { value: 130, computedValue: 130 },
              "3:0": { value: "Mar", computedValue: "Mar" },
              "3:1": { value: 90, computedValue: 90 },
              "3:2": { value: 70, computedValue: 70 }
            },
            merges: [],
            columns: {},
            rows: {}
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const selectedEvents: string[] = [];
    const unselectedEvents: string[] = [];
    engine.on("chart:selected", ({ chartId }) => selectedEvents.push(chartId));
    engine.on("chart:unselected", ({ chartId }) => unselectedEvents.push(chartId));
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();

    try {
      engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        colStart: 0,
        rowEnd: 3,
        colEnd: 2
      });
      container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.click();

      const chartId = engine.getCharts(sheet.id)[0]?.id as string;
      expect(selectedEvents).toContain(chartId);

      container.querySelector<HTMLElement>("[data-row='1'][data-col='1']")?.click();
      expect(unselectedEvents).toContain(chartId);

      container.querySelector<HTMLElement>(`[data-chart-id='${chartId}'] .excelsior-chart-object-header`)?.click();
      const panel = container.querySelector<HTMLElement>(".excelsior-chart-edit-panel");
      expect(panel?.hidden).toBe(false);

      const titleInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='title']");
      const typeSelect = panel?.querySelector<HTMLSelectElement>("[data-chart-role='type']");
      const rangeInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='range']");
      const legendToggle = panel?.querySelector<HTMLInputElement>("[data-chart-role='legend']");
      const categoryColumnInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='category-column']");
      const seriesColumnsInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='series-columns']");
      const valueColumnInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='value-column']");
      titleInput!.value = "Receita Trimestral";
      titleInput?.dispatchEvent(new Event("input", { bubbles: true }));
      typeSelect!.value = "bar";
      typeSelect?.dispatchEvent(new Event("input", { bubbles: true }));
      rangeInput!.value = "A1:C3";
      rangeInput?.dispatchEvent(new Event("input", { bubbles: true }));
      legendToggle!.checked = false;
      legendToggle?.dispatchEvent(new Event("input", { bubbles: true }));
      categoryColumnInput!.value = "1";
      categoryColumnInput?.dispatchEvent(new Event("input", { bubbles: true }));
      seriesColumnsInput!.value = "3";
      seriesColumnsInput?.dispatchEvent(new Event("input", { bubbles: true }));
      valueColumnInput!.value = "2";
      valueColumnInput?.dispatchEvent(new Event("input", { bubbles: true }));
      panel?.querySelector<HTMLButtonElement>("[data-chart-action='apply']")?.click();

      const edited = engine.getChart(sheet.id, chartId);
      expect(edited?.type).toBe("bar");
      expect(edited?.title).toBe("Receita Trimestral");
      expect(edited?.sourceRange?.rangeAddress).toBe("A1:C3");
      expect(edited?.sourceRange?.categoryColumnIndex).toBe(0);
      expect(edited?.sourceRange?.seriesColumnIndexes).toEqual([2]);
      expect(edited?.sourceRange?.valueColumnIndex).toBe(1);
      expect(Array.isArray(edited?.figure.data) ? edited?.figure.data.length : 0).toBe(1);
      expect((edited?.figure.layout as { legend?: { visible?: boolean } } | undefined)?.legend?.visible).toBe(false);
      expect(selectedEvents.filter((value) => value === chartId).length).toBeGreaterThan(1);
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("renders chart categories and disables in-development chart actions", () => {
    const container = document.createElement("div");
    container.style.width = "1000px";
    container.style.height = "520px";
    document.body.append(container);
    const engine = new WorkbookEngine();
    const renderer = new DomSpreadsheetRenderer(container, engine);

    try {
      expect(container.querySelector("[data-chart-category='common']")).not.toBeNull();
      expect(container.querySelector("[data-chart-category='statistical']")).not.toBeNull();
      expect(container.querySelector("[data-chart-category='financial']")).not.toBeNull();
      expect(container.querySelector("[data-chart-category='advanced']")).not.toBeNull();
      expect(container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.disabled).toBe(false);
      expect(container.querySelector<HTMLButtonElement>("[data-action='chart-candlestick']")?.disabled).toBe(true);
      expect(container.querySelector<HTMLButtonElement>("[data-action='chart-surface3d']")?.disabled).toBe(true);
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("updates axis titles from chart edit panel", () => {
    const container = document.createElement("div");
    container.style.width = "1100px";
    container.style.height = "620px";
    document.body.append(container);
    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Axis",
            rowCount: 12,
            columnCount: 3,
            cells: {
              "0:0": { value: "M", computedValue: "M" },
              "0:1": { value: "V", computedValue: "V" },
              "1:0": { value: "Jan", computedValue: "Jan" },
              "1:1": { value: 10, computedValue: 10 },
              "2:0": { value: "Feb", computedValue: "Feb" },
              "2:1": { value: 12, computedValue: 12 }
            }
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();
    try {
      engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        colStart: 0,
        rowEnd: 2,
        colEnd: 1
      });
      container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.click();
      const chartId = engine.getCharts(sheet.id)[0]?.id as string;
      container.querySelector<HTMLElement>(`[data-chart-id='${chartId}'] .excelsior-chart-object-header`)?.click();
      const panel = container.querySelector<HTMLElement>(".excelsior-chart-edit-panel");
      const xAxisInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='x-axis-title']");
      const yAxisInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='y-axis-title']");
      const xAxisTypeSelect = panel?.querySelector<HTMLSelectElement>("[data-chart-role='x-axis-type']");
      const yAxisTypeSelect = panel?.querySelector<HTMLSelectElement>("[data-chart-role='y-axis-type']");
      const xAxisVisibleToggle = panel?.querySelector<HTMLInputElement>("[data-chart-role='x-axis-visible']");
      const yAxisVisibleToggle = panel?.querySelector<HTMLInputElement>("[data-chart-role='y-axis-visible']");
      xAxisInput!.value = "Período";
      xAxisInput?.dispatchEvent(new Event("input", { bubbles: true }));
      yAxisInput!.value = "Receita";
      yAxisInput?.dispatchEvent(new Event("input", { bubbles: true }));
      xAxisTypeSelect!.value = "category";
      xAxisTypeSelect?.dispatchEvent(new Event("input", { bubbles: true }));
      yAxisTypeSelect!.value = "log";
      yAxisTypeSelect?.dispatchEvent(new Event("input", { bubbles: true }));
      xAxisVisibleToggle!.checked = false;
      xAxisVisibleToggle?.dispatchEvent(new Event("input", { bubbles: true }));
      yAxisVisibleToggle!.checked = true;
      yAxisVisibleToggle?.dispatchEvent(new Event("input", { bubbles: true }));
      panel?.querySelector<HTMLButtonElement>("[data-chart-action='apply']")?.click();

      const chart = engine.getChart(sheet.id, chartId);
      const layout = chart?.figure.layout as
        | {
            xAxis?: { title?: string; type?: string; visible?: boolean };
            yAxis?: { title?: string; type?: string; visible?: boolean };
          }
        | undefined;
      expect(layout?.xAxis?.title).toBe("Período");
      expect(layout?.yAxis?.title).toBe("Receita");
      expect(layout?.xAxis?.type).toBe("category");
      expect(layout?.yAxis?.type).toBe("log");
      expect(layout?.xAxis?.visible).toBe(false);
      expect(layout?.yAxis?.visible).toBe(true);
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("supports optional chart preview before insertion", () => {
    const container = document.createElement("div");
    container.style.width = "1100px";
    container.style.height = "620px";
    document.body.append(container);
    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Preview",
            rowCount: 10,
            columnCount: 4,
            cells: {
              "0:0": { value: "Month", computedValue: "Month" },
              "0:1": { value: "Revenue", computedValue: "Revenue" },
              "1:0": { value: "Jan", computedValue: "Jan" },
              "1:1": { value: 120, computedValue: 120 },
              "2:0": { value: "Feb", computedValue: "Feb" },
              "2:1": { value: 180, computedValue: 180 }
            },
            merges: [],
            columns: {},
            rows: {}
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const renderer = new DomSpreadsheetRenderer(container, engine, {
      chartInsertPreview: true
    });
    const sheet = engine.getActiveSheet();
    try {
      engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        colStart: 0,
        rowEnd: 2,
        colEnd: 1
      });
      container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.click();
      expect(engine.getCharts(sheet.id)).toHaveLength(0);
      const previewPanel = container.querySelector<HTMLElement>(".excelsior-chart-preview-panel");
      expect(previewPanel?.hidden).toBe(false);
      previewPanel?.querySelector<HTMLButtonElement>("[data-chart-action='preview-insert']")?.click();
      expect(engine.getCharts(sheet.id)).toHaveLength(1);
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("sanitizes malicious chart edit payloads from the panel", () => {
    const container = document.createElement("div");
    container.style.width = "1100px";
    container.style.height = "620px";
    document.body.append(container);
    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Security",
            rowCount: 12,
            columnCount: 3,
            cells: {
              "0:0": { value: "M", computedValue: "M" },
              "0:1": { value: "V", computedValue: "V" },
              "1:0": { value: "Jan", computedValue: "Jan" },
              "1:1": { value: 10, computedValue: 10 },
              "2:0": { value: "Feb", computedValue: "Feb" },
              "2:1": { value: 12, computedValue: 12 }
            }
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();
    try {
      engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        colStart: 0,
        rowEnd: 2,
        colEnd: 1
      });
      container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.click();
      const chartId = engine.getCharts(sheet.id)[0]?.id as string;
      container.querySelector<HTMLElement>(`[data-chart-id='${chartId}'] .excelsior-chart-object-header`)?.click();

      const panel = container.querySelector<HTMLElement>(".excelsior-chart-edit-panel");
      const titleInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='title']");
      const xAxisInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='x-axis-title']");
      const yAxisInput = panel?.querySelector<HTMLInputElement>("[data-chart-role='y-axis-title']");
      titleInput!.value = "<svg onload=alert(1)>";
      titleInput?.dispatchEvent(new Event("input", { bubbles: true }));
      xAxisInput!.value = "=cmd|' /C calc'!A0";
      xAxisInput?.dispatchEvent(new Event("input", { bubbles: true }));
      yAxisInput!.value = "<script>alert('x')</script>";
      yAxisInput?.dispatchEvent(new Event("input", { bubbles: true }));
      panel?.querySelector<HTMLButtonElement>("[data-chart-action='apply']")?.click();

      const chart = engine.getChart(sheet.id, chartId);
      expect(chart?.title).not.toContain("<");
      expect(chart?.title).not.toContain(">");
      const layout = chart?.figure.layout as
        | {
            xAxis?: { title?: string };
            yAxis?: { title?: string };
          }
        | undefined;
      expect(layout?.xAxis?.title?.startsWith("'")).toBe(true);
      expect(layout?.yAxis?.title).not.toContain("<");
      expect(layout?.yAxis?.title).not.toContain(">");
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it.each([
    { cells: 1000, maxRangeCells: 100000, shouldCreate: true, maxMs: 8000 },
    { cells: 10000, maxRangeCells: 100000, shouldCreate: true, maxMs: 26000 },
    { cells: 100000, maxRangeCells: 50000, shouldCreate: false, maxMs: 8000 }
  ])(
    "handles chart creation for $cells selected cells",
    ({ cells, maxRangeCells, shouldCreate, maxMs }) => {
    const container = document.createElement("div");
    container.style.width = "1200px";
    container.style.height = "700px";
    document.body.append(container);
    const rowCount = Math.floor(cells / 2);
    const blockedReasons: string[] = [];
    const engine = new WorkbookEngine(
      {
        settings: {
          maxRows: Math.max(2000, rowCount + 10),
          maxColumns: 20
        },
        data: [
          {
            name: "RangePerf",
            rowCount: Math.max(10, rowCount + 1),
            columnCount: 3,
            cells: {
              "0:0": { value: "X", computedValue: "X" },
              "0:1": { value: "Y", computedValue: "Y" },
              "1:0": { value: 1, computedValue: 1 },
              "1:1": { value: 2, computedValue: 2 }
            },
            merges: [],
            columns: {},
            rows: {}
          }
        ]
      },
      new BasicFormulaEngine()
    );
    engine.on("security:blocked-input", ({ reason }) => blockedReasons.push(reason));
    const renderer = new DomSpreadsheetRenderer(container, engine, {
      chartLimits: {
        maxRangeCells,
        maxSeriesPerChart: 128,
        maxPointsPerChart: 200000
      }
    });
    const sheet = engine.getActiveSheet();
    try {
      engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        colStart: 0,
        rowEnd: rowCount - 1,
        colEnd: 1
      });
      const startedAt = Date.now();
      container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.click();
      const elapsedMs = Date.now() - startedAt;
      if (shouldCreate) {
        expect(engine.getCharts(sheet.id)).toHaveLength(1);
      } else {
        expect(engine.getCharts(sheet.id)).toHaveLength(0);
        expect(blockedReasons).toContain("chart-range-too-large");
      }
      expect(elapsedMs).toBeLessThan(maxMs);
    } finally {
      renderer.dispose();
      container.remove();
    }
    },
    30000
  );

  it("blocks oversized ranges for chart creation and reports security", () => {
    const container = document.createElement("div");
    container.style.width = "1100px";
    container.style.height = "620px";
    document.body.append(container);
    const engine = new WorkbookEngine(
      {
        settings: {
          maxRows: 1000,
          maxColumns: 1000
        },
        data: [
          {
            name: "Big",
            rowCount: 400,
            columnCount: 400,
            cells: {}
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const securityReasons: string[] = [];
    engine.on("security:blocked-input", ({ reason }) => securityReasons.push(reason));
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();
    try {
      engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        colStart: 0,
        rowEnd: 399,
        colEnd: 399
      });
      container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.click();
      expect(engine.getCharts(sheet.id)).toHaveLength(0);
      expect(securityReasons).toContain("chart-range-too-large");
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("emits render lifecycle events and skips offscreen chart previews", () => {
    const container = document.createElement("div");
    container.style.width = "700px";
    container.style.height = "420px";
    document.body.append(container);
    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Render",
            rowCount: 200,
            columnCount: 40,
            cells: {
              "0:0": { value: "X", computedValue: "X" },
              "0:1": { value: "Y", computedValue: "Y" },
              "1:0": { value: 1, computedValue: 1 },
              "1:1": { value: 2, computedValue: 2 }
            }
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const sheet = engine.getActiveSheet();
    const renderStarted: string[] = [];
    const renderFinished: string[] = [];
    const renderSkipped: string[] = [];
    engine.on("chart:renderStarted", ({ chartId }) => renderStarted.push(chartId));
    engine.on("chart:renderFinished", ({ chartId }) => renderFinished.push(chartId));
    engine.on("chart:renderSkipped", ({ reason }) => renderSkipped.push(reason));

    engine.createChart({
      sheetId: sheet.id,
      chart: {
        id: "chart-visible",
        type: "line",
        title: "Visible",
        sourceRange: {
          rangeAddress: "A1:B2",
          orientation: "rows",
          firstRowAsHeader: true,
          firstColumnAsLabel: true,
          autoRefresh: true
        },
        figure: {
          data: [{ type: "line", x: [1, 2], y: [2, 3] }],
          layout: { title: "Visible" }
        },
        position: {
          fromCell: "C2",
          toCell: "I12",
          offsetX: 0,
          offsetY: 0,
          width: 300,
          height: 180
        },
        state: {
          selected: false,
          visible: true,
          locked: false
        }
      }
    });
    engine.createChart({
      sheetId: sheet.id,
      chart: {
        id: "chart-offscreen",
        type: "line",
        title: "Offscreen",
        sourceRange: {
          rangeAddress: "A1:B2",
          orientation: "rows",
          firstRowAsHeader: true,
          firstColumnAsLabel: true,
          autoRefresh: true
        },
        figure: {
          data: [{ type: "line", x: [1, 2], y: [2, 4] }],
          layout: { title: "Offscreen" }
        },
        position: {
          fromCell: "Z80",
          toCell: "AF96",
          offsetX: 0,
          offsetY: 0,
          width: 320,
          height: 200
        },
        state: {
          selected: false,
          visible: true,
          locked: false
        }
      }
    });

    const renderer = new DomSpreadsheetRenderer(container, engine);
    try {
      expect(renderStarted).toContain("chart-visible");
      expect(renderFinished).toContain("chart-visible");
      expect(renderSkipped).toContain("outside-viewport");
      expect(container.querySelector("[data-chart-id='chart-offscreen'].is-offscreen")).not.toBeNull();
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it.each([1, 5, 10, 25, 50])("keeps viewport stable with %i chart objects", (chartCount) => {
    const container = document.createElement("div");
    container.style.width = "1200px";
    container.style.height = "700px";
    document.body.append(container);
    const engine = new WorkbookEngine(
      {
        settings: {
          maxRows: 2000,
          maxColumns: 200
        },
        data: [
          {
            id: "sheet-perf",
            name: "Perf",
            rowCount: 300,
            columnCount: 40,
            cells: {
              "0:0": { value: "X", computedValue: "X" },
              "0:1": { value: "Y", computedValue: "Y" },
              "1:0": { value: 1, computedValue: 1 },
              "1:1": { value: 2, computedValue: 2 },
              "2:0": { value: 2, computedValue: 2 },
              "2:1": { value: 4, computedValue: 4 }
            },
            merges: [],
            columns: {},
            rows: {}
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const sheet = engine.getActiveSheet();
    for (let index = 0; index < chartCount; index += 1) {
      const rowBase = 2 + index * 2;
      const colBase = (index % 8) * 4;
      engine.createChart({
        sheetId: sheet.id,
        chart: {
          id: `chart-perf-${index}`,
          type: "line",
          title: `Perf ${index}`,
          sourceRange: {
            rangeAddress: "A1:B3",
            orientation: "rows",
            firstRowAsHeader: true,
            firstColumnAsLabel: true,
            autoRefresh: true
          },
          figure: {
            data: [
              {
                type: "line",
                x: [1, 2],
                y: [index + 1, index + 2]
              }
            ],
            layout: {
              title: `Perf ${index}`
            }
          },
          position: {
            fromCell: cellAddressToLabel({ row: rowBase, col: colBase }),
            toCell: cellAddressToLabel({ row: rowBase + 6, col: colBase + 3 }),
            offsetX: 0,
            offsetY: 0,
            width: 220,
            height: 150,
            zIndex: index + 1
          },
          state: {
            selected: false,
            visible: true,
            locked: false
          }
        }
      });
    }

    const renderer = new DomSpreadsheetRenderer(container, engine, {
      chartPerformance: {
        skipOffscreenPreview: true
      }
    });
    try {
      const objects = container.querySelectorAll(".excelsior-chart-object");
      expect(objects.length).toBe(chartCount);
      const viewport = container.querySelector<HTMLElement>(".excelsior-viewport");
      viewport!.scrollTop = 1200;
      viewport?.dispatchEvent(new Event("scroll"));
      expect(container.querySelectorAll(".excelsior-chart-object").length).toBe(chartCount);
    } finally {
      renderer.dispose();
      container.remove();
    }
  });

  it("releases chart runtimes and caches on dispose", () => {
    const container = document.createElement("div");
    container.style.width = "900px";
    container.style.height = "560px";
    document.body.append(container);
    const engine = new WorkbookEngine(
      {
        data: [
          {
            name: "Cleanup",
            rowCount: 8,
            columnCount: 4,
            cells: {
              "0:0": { value: "X", computedValue: "X" },
              "0:1": { value: "Y", computedValue: "Y" },
              "1:0": { value: 1, computedValue: 1 },
              "1:1": { value: 2, computedValue: 2 }
            }
          }
        ]
      },
      new BasicFormulaEngine()
    );
    const renderer = new DomSpreadsheetRenderer(container, engine);
    const sheet = engine.getActiveSheet();
    engine.selectRange({
      sheetId: sheet.id,
      rowStart: 0,
      colStart: 0,
      rowEnd: 2,
      colEnd: 1
    });
    container.querySelector<HTMLButtonElement>("[data-action='chart-line']")?.click();
    const rendererState = renderer as unknown as {
      chartRuntimeById: Map<string, unknown>;
      chartObjectElementById: Map<string, unknown>;
      chartBodyElementById: Map<string, unknown>;
    };
    expect(rendererState.chartObjectElementById.size).toBeGreaterThan(0);
    renderer.dispose();
    expect(rendererState.chartRuntimeById.size).toBe(0);
    expect(rendererState.chartObjectElementById.size).toBe(0);
    expect(rendererState.chartBodyElementById.size).toBe(0);
    expect(container.childElementCount).toBe(0);
    container.remove();
  });
});