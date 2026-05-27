import { describe, expect, it, vi } from "vitest";
import { ClientSideRowModel, InfiniteRowModel, ServerSideRowModel, ViewportRowModel, WorkbookEngine } from "../src/index";

const createDeferred = <T>() => {
  let resolve: (value: T) => void;
  const promise = new Promise<T>((innerResolve) => {
    resolve = innerResolve;
  });

  return {
    promise,
    resolve: resolve!
  };
};

describe("row models", () => {
  it("exposes a default client-side row model per sheet", () => {
    const engine = WorkbookEngine.fromJSON({
      id: "workbook-1",
      activeSheetId: "sheet-1",
      metadata: {},
      settings: {
        maxRows: 100,
        maxColumns: 26,
        maxCellLength: 5000,
        maxFormulaLength: 2048,
        maxPasteCells: 10000,
        rowHeight: 28,
        columnWidth: 120,
        viewportBuffer: 4,
        maxHistorySize: 100,
        enableFormulas: true,
        clipboardPolicy: "text-only"
      },
      sheets: [
        {
          id: "sheet-1",
          name: "Sheet1",
          cells: {},
          merges: [],
          columns: {},
          rows: {
            1: { hidden: true },
            2: { height: 44 }
          },
          rowCount: 5,
          columnCount: 5,
          selection: {
            start: { row: 0, col: 0 },
            end: { row: 0, col: 0 }
          }
        }
      ]
    });

    const rowModel = engine.getRowModel("sheet-1");
    const result = rowModel.getRows({ sheetId: "sheet-1", startRow: 0, endRow: 2 });

    expect(rowModel).toBeInstanceOf(ClientSideRowModel);
    expect(rowModel.getRowCount()).toBe(5);
    expect(result).toEqual({
      rowCount: 5,
      rows: [
        { index: 0, hidden: false, height: 28 },
        { index: 1, hidden: true, height: 28 },
        { index: 2, hidden: false, height: 44 }
      ]
    });
  });

  it("allows explicit row model replacement per sheet", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const refreshSpy = vi.fn();
    const disposeSpy = vi.fn();
    const rowModel = new ViewportRowModel({
      rowCount: 10,
      getRows: ({ startRow, endRow }) => ({
        rowCount: 10,
        rows: [{ index: startRow }, { index: endRow }]
      }),
      onRefresh: refreshSpy,
      onDispose: disposeSpy
    });

    engine.setRowModel(sheet.id, rowModel);

    expect(engine.getRowModel(sheet.id)).toBe(rowModel);
    expect(refreshSpy).toHaveBeenCalledWith("initial");
    expect(rowModel.getRows({ sheetId: sheet.id, startRow: 2, endRow: 4 })).toEqual({
      rowCount: 10,
      rows: [{ index: 2 }, { index: 4 }]
    });

    engine.clearRowModel(sheet.id);

    expect(disposeSpy).toHaveBeenCalledTimes(1);
    expect(engine.getRowModel(sheet.id)).toBeInstanceOf(ClientSideRowModel);
  });

  it("loads infinite row model data in blocks and reuses cached blocks", async () => {
    const getRows = vi.fn(async ({ startRow, endRow }: { startRow: number; endRow: number }) => ({
      totalRows: 24,
      rows: Array.from({ length: endRow - startRow + 1 }, (_value, offset) => ({
        index: startRow + offset
      }))
    }));

    const rowModel = new InfiniteRowModel({
      rowCount: "unknown",
      blockSize: 5,
      dataSource: {
        getRows
      }
    });

    await expect(rowModel.getRows({ sheetId: "sheet-1", startRow: 2, endRow: 7 })).resolves.toEqual({
      rowCount: 24,
      rows: [{ index: 2 }, { index: 3 }, { index: 4 }, { index: 5 }, { index: 6 }, { index: 7 }]
    });

    await expect(rowModel.getRows({ sheetId: "sheet-1", startRow: 3, endRow: 4 })).resolves.toEqual({
      rowCount: 24,
      rows: [{ index: 3 }, { index: 4 }]
    });

    expect(getRows.mock.calls).toHaveLength(2);
    expect(getRows.mock.calls[0][0]).toMatchObject({ startRow: 0, endRow: 4 });
    expect(getRows.mock.calls[1][0]).toMatchObject({ startRow: 5, endRow: 9 });
    expect(rowModel.getRowCount()).toBe(24);
  });

  it("ignores stale server-side responses that resolve after a newer request", async () => {
    const first = createDeferred<{ rows: Array<{ index: number }>; totalRows: number }>();
    const second = createDeferred<{ rows: Array<{ index: number }>; totalRows: number }>();
    const dataSource = {
      getRows: vi
        .fn()
        .mockImplementationOnce(() => first.promise)
        .mockImplementationOnce(() => second.promise)
    };
    const rowModel = new ServerSideRowModel({
      rowCount: "unknown",
      dataSource
    });

    const olderRequest = rowModel.getRows({ sheetId: "sheet-1", startRow: 0, endRow: 1 });
    const newerRequest = rowModel.getRows({ sheetId: "sheet-1", startRow: 10, endRow: 11 });

    second.resolve({
      totalRows: 200,
      rows: [{ index: 10 }, { index: 11 }]
    });

    await expect(newerRequest).resolves.toEqual({
      rowCount: 200,
      rows: [{ index: 10 }, { index: 11 }]
    });

    first.resolve({
      totalRows: 50,
      rows: [{ index: 0 }, { index: 1 }]
    });

    await expect(olderRequest).resolves.toEqual({
      rowCount: 200,
      rows: [{ index: 10 }, { index: 11 }]
    });
    expect(rowModel.getRowCount()).toBe(200);
  });

  it("aborts an older server-side request when a newer one starts", async () => {
    let firstSignal: AbortSignal | undefined;
    const second = createDeferred<{ rows: Array<{ index: number }>; totalRows: number }>();
    const dataSource = {
      getRows: vi
        .fn()
        .mockImplementationOnce((_request, context?: { signal?: AbortSignal }) => {
          firstSignal = context?.signal;
          return new Promise<{ rows: Array<{ index: number }>; totalRows: number }>((_resolve, reject) => {
            context?.signal?.addEventListener("abort", () => {
              reject(new DOMException("Request aborted", "AbortError"));
            });
          });
        })
        .mockImplementationOnce(() => second.promise)
    };
    const rowModel = new ServerSideRowModel({
      rowCount: "unknown",
      dataSource
    });

    const olderRequest = rowModel.getRows({ sheetId: "sheet-1", startRow: 0, endRow: 1 });
    const newerRequest = rowModel.getRows({ sheetId: "sheet-1", startRow: 10, endRow: 11 });

    expect(firstSignal?.aborted).toBe(true);

    second.resolve({
      totalRows: 200,
      rows: [{ index: 10 }, { index: 11 }]
    });

    await expect(newerRequest).resolves.toEqual({
      rowCount: 200,
      rows: [{ index: 10 }, { index: 11 }]
    });
    await expect(olderRequest).resolves.toEqual({
      rowCount: "unknown",
      rows: []
    });
  });

  it("updates sort, filter, grouping, expansion, pivot and aggregation through the engine without replacing the remote row model", async () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const dataSource = {
      getRows: vi.fn(async ({ sortModel, filterModel, groupKeys, expandedGroupPaths, pivotModel, aggregateModel }: {
        sortModel?: unknown;
        filterModel?: unknown;
        groupKeys?: unknown;
        expandedGroupPaths?: unknown;
        pivotModel?: unknown;
        aggregateModel?: unknown;
      }) => ({
        totalRows: 20,
        rows: [{ index: 0 }],
        groupInfo: [{ key: "engineering", path: ["engineering"], level: 0, childCount: 12, expanded: true }],
        sortModel,
        filterModel,
        groupKeys,
        expandedGroupPaths,
        pivotModel,
        aggregateModel
      }))
    };
    const rowModel = new ServerSideRowModel({
      rowCount: "unknown",
      dataSource
    });
    const changedSpy = vi.fn();

    engine.on("row-model:changed", changedSpy);
    engine.setRowModel(sheet.id, rowModel);

    await rowModel.getRows({ sheetId: sheet.id, startRow: 0, endRow: 0 });
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
    await rowModel.getRows({ sheetId: sheet.id, startRow: 0, endRow: 0 });

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
    expect(dataSource.getRows.mock.calls[1][0]).toMatchObject({
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
    expect(changedSpy).toHaveBeenCalledTimes(2);
  });

  it("preserves expanded remote group paths through the helper API", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    const rowModel = new ServerSideRowModel({
      rowCount: "unknown",
      dataSource: {
        getRows: vi.fn(async () => ({ totalRows: 10, rows: [{ index: 0 }] }))
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

  it("accepts initial expanded group paths in remote row model options", () => {
    const rowModel = new ServerSideRowModel({
      rowCount: "unknown",
      expandedGroupPaths: [["engineering"], ["engineering", "active"]],
      dataSource: {
        getRows: vi.fn(async () => ({ totalRows: 10, rows: [{ index: 0 }] }))
      }
    });

    expect(rowModel.getRequestModel()).toEqual({
      expandedGroupPaths: [["engineering"], ["engineering", "active"]]
    });
  });

  it("maps remote group metadata onto row model rows", async () => {
    const rowModel = new ServerSideRowModel({
      rowCount: "unknown",
      dataSource: {
        getRows: vi.fn(async () => ({
          totalRows: 10,
          rows: [{ index: 3 }],
          groupInfo: [{ key: "engineering", path: ["engineering", "active"], level: 1, childCount: 4, expanded: false }]
        }))
      }
    });

    await expect(rowModel.getRows({ sheetId: "sheet-1", startRow: 3, endRow: 3 })).resolves.toEqual({
      rowCount: 10,
      rows: [
        {
          index: 3,
          group: { key: "engineering", path: ["engineering", "active"], level: 1, childCount: 4, expanded: false }
        }
      ]
    });
  });

  it("materializes a pivot sheet through the remote data source contract", async () => {
    const getRows = vi.fn(async (request: {
      requestKind?: string;
      groupKeys?: string[];
      pivotModel?: Array<{ field: string }>;
      aggregateModel?: Array<{ field: string; function: string; as?: string }>;
      pivotInput?: { sheetName?: string };
    }) => ({
      rows: [],
      pivotSheet: {
        name: request.pivotInput?.sheetName ?? "Remote Pivot",
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
    }));
    const rowModel = new ServerSideRowModel({
      rowCount: "unknown",
      sortModel: [{ field: "revenue", direction: "desc" }],
      filterModel: {
        status: { operator: "equals", value: "active" }
      },
      dataSource: {
        getRows
      }
    });

    await expect(
      rowModel.buildPivotSheet({
        sourceSheetId: "sheet-1",
        sourceRange: {
          start: { row: 0, col: 0 },
          end: { row: 100, col: 4 }
        },
        rows: ["Region"],
        columns: ["Quarter"],
        values: [{ field: "Revenue", aggregate: "sum", as: "Revenue" }],
        sheetName: "Remote Pivot",
        executionMode: "server"
      })
    ).resolves.toMatchObject({
      name: "Remote Pivot",
      cells: {
        "1:1": { value: 42, computedValue: 42 }
      }
    });

    expect(getRows).toHaveBeenCalledWith(
      expect.objectContaining({
        requestKind: "pivotSheet",
        sortModel: [{ field: "revenue", direction: "desc" }],
        filterModel: {
          status: { operator: "equals", value: "active" }
        },
        groupKeys: ["Region"],
        pivotModel: [{ field: "Quarter" }],
        aggregateModel: [{ field: "Revenue", function: "sum", as: "Revenue" }]
      }),
      expect.any(Object)
    );
  });
});