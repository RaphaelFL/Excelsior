import type { PivotSheetInput, RowSchema, WorkbookDataInput, WorkbookModel } from "./domain/types";
import type {
  DataAggregateModel,
  DataFilterModel,
  DataGroupInfo,
  DataPivotModel,
  DataSourceRequestContext,
  DataRequest,
  DataResponse,
  DataSortModel,
  DataSource
} from "./data-source";
import { SpreadsheetOperationError } from "./errors/spreadsheet-operation-error";

export type RowModelKind = "clientSide" | "viewport" | "infinite" | "serverSide";

export type RowModelRefreshReason = "initial" | "manual" | "sheetChanged" | "viewportChanged";

export interface RowRequest {
  sheetId: string;
  startRow: number;
  endRow: number;
}

export interface RowModelRow {
  index: number;
  group?: DataGroupInfo;
  hidden?: boolean;
  height?: number;
}

export interface RowResult {
  rows: RowModelRow[];
  rowCount?: number | "unknown";
}

export interface RowModel {
  readonly kind: RowModelKind;
  getRowCount(): number | "unknown";
  getRows(request: RowRequest): RowResult | Promise<RowResult>;
  refresh(reason: RowModelRefreshReason): void;
  dispose(): void;
}

export interface RemotePivotCapableRowModel extends RowModel {
  buildPivotSheet(input: PivotSheetInput): Promise<WorkbookDataInput>;
  getRequestModel(): RemoteRowModelUpdate;
}

export interface RowModelWorkbookSource {
  getSnapshot(): WorkbookModel;
  getRowSchema(sheetId: string, row: number): RowSchema | undefined;
}

export interface RemoteRowModelOptions {
  dataSource: DataSource;
  rowCount?: number | "unknown";
  sortModel?: DataSortModel;
  filterModel?: DataFilterModel;
  groupKeys?: string[];
  expandedGroupPaths?: string[][];
  pivotModel?: DataPivotModel;
  aggregateModel?: DataAggregateModel;
  visibleColumns?: string[];
  onRefresh?: (reason: RowModelRefreshReason) => void;
  onDispose?: () => void;
}

export interface RemoteRowModelUpdate {
  sortModel?: DataSortModel;
  filterModel?: DataFilterModel;
  groupKeys?: string[];
  expandedGroupPaths?: string[][];
  pivotModel?: DataPivotModel;
  aggregateModel?: DataAggregateModel;
}

export interface InfiniteRowModelOptions extends RemoteRowModelOptions {
  blockSize?: number;
}

export interface ServerSideRowModelOptions extends RemoteRowModelOptions {}

const normalizeRequestRange = (
  request: RowRequest,
  rowCount: number | "unknown"
): { startRow: number; endRow: number } => {
  const startRow = Math.max(0, Math.min(request.startRow, request.endRow));
  const requestedEnd = Math.max(request.startRow, request.endRow);
  if (rowCount === "unknown") {
    return {
      startRow,
      endRow: Math.max(startRow, requestedEnd)
    };
  }

  if (rowCount <= 0) {
    return {
      startRow: 0,
      endRow: -1
    };
  }

  return {
    startRow: Math.min(startRow, rowCount - 1),
    endRow: Math.max(0, Math.min(rowCount - 1, requestedEnd))
  };
};

const toRowModelRows = (response: DataResponse): RowModelRow[] =>
  response.rows.map((row, index) => ({
    index: row.index,
    group: response.groupInfo?.[index] ? { ...response.groupInfo[index] } : undefined,
    hidden: row.hidden,
    height: row.height
  }));

const cloneRowResult = (result: RowResult): RowResult => ({
  rowCount: result.rowCount,
  rows: result.rows.map((row) => ({
    ...row,
    group: row.group ? { ...row.group } : undefined
  }))
});

const createEmptyRowResult = (rowCount: number | "unknown"): RowResult => ({
  rows: [],
  rowCount
});

const createRemoteRequest = (
  kind: Extract<RowModelKind, "infinite" | "serverSide">,
  requestState: RemoteRowModelUpdate,
  options: RemoteRowModelOptions,
  request: RowRequest,
  startRow: number,
  endRow: number,
  requestVersion: number
): DataRequest => ({
  sheetId: request.sheetId,
  startRow,
  endRow,
  sortModel: requestState.sortModel,
  filterModel: requestState.filterModel,
  groupKeys: requestState.groupKeys,
  expandedGroupPaths: requestState.expandedGroupPaths?.map((path) => [...path]),
  pivotModel: requestState.pivotModel,
  aggregateModel: requestState.aggregateModel,
  visibleColumns: options.visibleColumns,
  requestId: `${kind}:${request.sheetId}:${requestVersion}:${startRow}:${endRow}`
});

const createPivotSheetRequest = (
  kind: Extract<RowModelKind, "infinite" | "serverSide">,
  requestState: RemoteRowModelUpdate,
  options: RemoteRowModelOptions,
  input: PivotSheetInput,
  requestVersion: number
): DataRequest => ({
  sheetId: input.sourceSheetId,
  startRow: input.sourceRange.start.row,
  endRow: input.sourceRange.end.row,
  requestKind: "pivotSheet",
  sortModel: requestState.sortModel,
  filterModel: requestState.filterModel,
  groupKeys: input.rows ?? requestState.groupKeys,
  expandedGroupPaths: requestState.expandedGroupPaths?.map((path) => [...path]),
  pivotModel: input.columns?.map((field) => ({ field })) ?? requestState.pivotModel,
  aggregateModel: input.values.map((value) => ({
    field: value.field,
    function: value.aggregate,
    as: value.as
  })),
  pivotInput: {
    ...input,
    sourceRange: {
      start: { ...input.sourceRange.start },
      end: { ...input.sourceRange.end }
    },
    rows: input.rows ? [...input.rows] : undefined,
    columns: input.columns ? [...input.columns] : undefined,
    values: input.values.map((value) => ({ ...value }))
  },
  visibleColumns: options.visibleColumns,
  requestId: `${kind}:pivot:${input.sourceSheetId}:${requestVersion}`
});

const createRemotePivotError = (code: string, message: string, details?: Record<string, unknown>): SpreadsheetOperationError =>
  new SpreadsheetOperationError({
    code,
    message,
    area: "pivot",
    recoverable: true,
    details
  });

const requestRemotePivotSheet = async (
  kind: Extract<RowModelKind, "infinite" | "serverSide">,
  requestState: RemoteRowModelUpdate,
  options: RemoteRowModelOptions,
  input: PivotSheetInput,
  requestVersion: number,
  getCurrentRequestVersion: () => number,
  contextFactory: () => { controller: AbortController; context: DataSourceRequestContext },
  releaseController: (controller: AbortController) => void
): Promise<WorkbookDataInput> => {
  const { controller, context } = contextFactory();

  try {
    const response = await options.dataSource.getRows(
      createPivotSheetRequest(kind, requestState, options, input, requestVersion),
      context
    );

    if (requestVersion !== getCurrentRequestVersion()) {
      throw createRemotePivotError("PIVOT_REMOTE_STALE", "Remote pivot response was superseded by a newer request.");
    }

    if (!response.pivotSheet) {
      throw createRemotePivotError("PIVOT_REMOTE_UNSUPPORTED", "Remote data source did not return a pivot sheet.", {
        sheetId: input.sourceSheetId
      });
    }

    return response.pivotSheet;
  } catch (error) {
    if (controller.signal.aborted && requestVersion !== getCurrentRequestVersion() && isAbortError(error)) {
      throw createRemotePivotError("PIVOT_REMOTE_ABORTED", "Remote pivot request was aborted.", {
        sheetId: input.sourceSheetId
      });
    }

    throw error;
  } finally {
    releaseController(controller);
  }
};

const applyRemoteRowModelUpdate = (
  current: RemoteRowModelUpdate,
  update: RemoteRowModelUpdate
): RemoteRowModelUpdate => ({
  sortModel: Object.hasOwn(update, "sortModel") ? update.sortModel : current.sortModel,
  filterModel: Object.hasOwn(update, "filterModel") ? update.filterModel : current.filterModel,
  groupKeys: Object.hasOwn(update, "groupKeys") ? update.groupKeys : current.groupKeys,
  expandedGroupPaths: Object.hasOwn(update, "expandedGroupPaths")
    ? update.expandedGroupPaths
    : current.expandedGroupPaths,
  pivotModel: Object.hasOwn(update, "pivotModel") ? update.pivotModel : current.pivotModel,
  aggregateModel: Object.hasOwn(update, "aggregateModel")
    ? update.aggregateModel
    : current.aggregateModel
});

const cloneRemoteRowModelUpdate = (update: RemoteRowModelUpdate): RemoteRowModelUpdate => ({
  sortModel: update.sortModel?.map((item) => ({ ...item })),
  filterModel: update.filterModel
    ? Object.fromEntries(Object.entries(update.filterModel).map(([field, descriptor]) => [field, { ...descriptor }]))
    : undefined,
  groupKeys: update.groupKeys ? [...update.groupKeys] : undefined,
  expandedGroupPaths: update.expandedGroupPaths?.map((path) => [...path]),
  pivotModel: update.pivotModel?.map((item) => ({ ...item })),
  aggregateModel: update.aggregateModel?.map((item) => ({ ...item }))
});

const applyRemoteResponse = (
  currentRowCount: number | "unknown",
  response: DataResponse
): { rowCount: number | "unknown"; result: RowResult } => {
  const rowCount = typeof response.totalRows === "number" ? response.totalRows : currentRowCount;
  return {
    rowCount,
    result: {
      rows: toRowModelRows(response),
      rowCount
    }
  };
};

const getBlockRanges = (startRow: number, endRow: number, blockSize: number): Array<{ startRow: number; endRow: number }> => {
  if (endRow < startRow) {
    return [];
  }

  const ranges: Array<{ startRow: number; endRow: number }> = [];
  let blockStart = Math.floor(startRow / blockSize) * blockSize;
  while (blockStart <= endRow) {
    ranges.push({
      startRow: blockStart,
      endRow: blockStart + blockSize - 1
    });
    blockStart += blockSize;
  }
  return ranges;
};

const isAbortError = (error: unknown): boolean =>
  error instanceof DOMException
    ? error.name === "AbortError"
    : typeof error === "object" && error !== null && "name" in error && (error as { name?: unknown }).name === "AbortError";

export class ClientSideRowModel implements RowModel {
  readonly kind = "clientSide" as const;

  constructor(
    private readonly source: RowModelWorkbookSource,
    private readonly sheetId: string
  ) {}

  getRowCount(): number {
    const sheet = this.source.getSnapshot().sheets.find((item) => item.id === this.sheetId);
    return sheet?.rowCount ?? 0;
  }

  getRows(request: RowRequest): RowResult {
    const snapshot = this.source.getSnapshot();
    const sheet = snapshot.sheets.find((item) => item.id === this.sheetId);
    if (!sheet) {
      return { rows: [], rowCount: 0 };
    }

    const { startRow, endRow } = normalizeRequestRange(request, sheet.rowCount);
    const rows: RowModelRow[] = [];
    for (let index = startRow; index <= endRow; index += 1) {
      const schema = this.source.getRowSchema(this.sheetId, index);
      rows.push({
        index,
        hidden: Boolean(schema?.hidden),
        height: schema?.height ?? snapshot.settings.rowHeight
      });
    }

    return {
      rows,
      rowCount: sheet.rowCount
    };
  }

  refresh(_reason: RowModelRefreshReason): void {}

  dispose(): void {}
}

export interface ViewportRowModelOptions {
  rowCount: number | "unknown";
  getRows: (request: RowRequest) => RowResult | Promise<RowResult>;
  onRefresh?: (reason: RowModelRefreshReason) => void;
  onDispose?: () => void;
}

export class ViewportRowModel implements RowModel {
  readonly kind = "viewport" as const;

  constructor(private readonly options: ViewportRowModelOptions) {}

  getRowCount(): number | "unknown" {
    return this.options.rowCount;
  }

  getRows(request: RowRequest): RowResult | Promise<RowResult> {
    return this.options.getRows(request);
  }

  refresh(reason: RowModelRefreshReason): void {
    this.options.onRefresh?.(reason);
  }

  dispose(): void {
    this.options.onDispose?.();
  }
}

export class InfiniteRowModel implements RowModel {
  readonly kind = "infinite" as const;

  private knownRowCount: number | "unknown";

  private requestVersion = 0;

  private lastResolvedResult: RowResult;

  private readonly blockCache = new Map<number, RowModelRow[]>();

  private readonly pendingRequestControllers = new Set<AbortController>();

  private readonly blockSize: number;

  private requestState: RemoteRowModelUpdate;

  constructor(private readonly options: InfiniteRowModelOptions) {
    this.knownRowCount = options.rowCount ?? "unknown";
    this.lastResolvedResult = createEmptyRowResult(this.knownRowCount);
    this.blockSize = Math.max(1, options.blockSize ?? 100);
    this.requestState = {
      sortModel: options.sortModel,
      filterModel: options.filterModel,
      groupKeys: options.groupKeys,
      expandedGroupPaths: options.expandedGroupPaths,
      pivotModel: options.pivotModel,
      aggregateModel: options.aggregateModel
    };
  }

  getRowCount(): number | "unknown" {
    return this.knownRowCount;
  }

  private abortPendingRequests(): void {
    for (const controller of this.pendingRequestControllers) {
      controller.abort();
    }
    this.pendingRequestControllers.clear();
  }

  private createRequestContext(): { controller: AbortController; context: DataSourceRequestContext } {
    const controller = new AbortController();
    this.pendingRequestControllers.add(controller);
    return {
      controller,
      context: {
        signal: controller.signal
      }
    };
  }

  async getRows(request: RowRequest): Promise<RowResult> {
    const requestVersion = ++this.requestVersion;
    this.abortPendingRequests();
    const { startRow, endRow } = normalizeRequestRange(request, this.knownRowCount);
    if (endRow < startRow) {
      return createEmptyRowResult(this.knownRowCount);
    }

    const blockRanges = getBlockRanges(startRow, endRow, this.blockSize);
    const missingBlockRanges = blockRanges.filter((range) => !this.blockCache.has(range.startRow));

    await Promise.all(
      missingBlockRanges.map(async (range) => {
        const { controller, context } = this.createRequestContext();

        try {
          const response = await this.options.dataSource.getRows(
            createRemoteRequest(this.kind, this.requestState, this.options, request, range.startRow, range.endRow, requestVersion),
            context
          );
          if (requestVersion !== this.requestVersion) {
            return;
          }

          const { rowCount, result } = applyRemoteResponse(this.knownRowCount, response);
          this.knownRowCount = rowCount;
          this.blockCache.set(range.startRow, result.rows);
        } catch (error) {
          if (controller.signal.aborted && requestVersion !== this.requestVersion && isAbortError(error)) {
            return;
          }
          throw error;
        } finally {
          this.pendingRequestControllers.delete(controller);
        }
      })
    );

    if (requestVersion !== this.requestVersion) {
      return cloneRowResult(this.lastResolvedResult);
    }

    const rows = blockRanges.flatMap((range) =>
      (this.blockCache.get(range.startRow) ?? []).filter((row) => row.index >= startRow && row.index <= endRow)
    );
    const result: RowResult = {
      rows,
      rowCount: this.knownRowCount
    };
    this.lastResolvedResult = cloneRowResult(result);
    return result;
  }

  refresh(reason: RowModelRefreshReason): void {
    this.requestVersion += 1;
    this.abortPendingRequests();
    this.blockCache.clear();
    this.options.onRefresh?.(reason);
  }

  updateRequest(update: RemoteRowModelUpdate): void {
    this.requestState = applyRemoteRowModelUpdate(this.requestState, update);
    this.refresh("manual");
  }

  getRequestModel(): RemoteRowModelUpdate {
    return cloneRemoteRowModelUpdate(this.requestState);
  }

  async buildPivotSheet(input: PivotSheetInput): Promise<WorkbookDataInput> {
    const requestVersion = ++this.requestVersion;
    this.abortPendingRequests();
    return requestRemotePivotSheet(
      this.kind,
      this.requestState,
      this.options,
      input,
      requestVersion,
      () => this.requestVersion,
      () => this.createRequestContext(),
      (controller) => {
        this.pendingRequestControllers.delete(controller);
      }
    );
  }

  dispose(): void {
    this.requestVersion += 1;
    this.abortPendingRequests();
    this.blockCache.clear();
    this.options.onDispose?.();
  }
}

export class ServerSideRowModel implements RowModel {
  readonly kind = "serverSide" as const;

  private knownRowCount: number | "unknown";

  private requestVersion = 0;

  private lastResolvedResult: RowResult;

  private readonly pendingRequestControllers = new Set<AbortController>();

  private requestState: RemoteRowModelUpdate;

  constructor(private readonly options: ServerSideRowModelOptions) {
    this.knownRowCount = options.rowCount ?? "unknown";
    this.lastResolvedResult = createEmptyRowResult(this.knownRowCount);
    this.requestState = {
      sortModel: options.sortModel,
      filterModel: options.filterModel,
      groupKeys: options.groupKeys,
      expandedGroupPaths: options.expandedGroupPaths,
      pivotModel: options.pivotModel,
      aggregateModel: options.aggregateModel
    };
  }

  getRowCount(): number | "unknown" {
    return this.knownRowCount;
  }

  private abortPendingRequests(): void {
    for (const controller of this.pendingRequestControllers) {
      controller.abort();
    }
    this.pendingRequestControllers.clear();
  }

  private createRequestContext(): { controller: AbortController; context: DataSourceRequestContext } {
    const controller = new AbortController();
    this.pendingRequestControllers.add(controller);
    return {
      controller,
      context: {
        signal: controller.signal
      }
    };
  }

  async getRows(request: RowRequest): Promise<RowResult> {
    const requestVersion = ++this.requestVersion;
    this.abortPendingRequests();
    const { startRow, endRow } = normalizeRequestRange(request, this.knownRowCount);
    if (endRow < startRow) {
      return createEmptyRowResult(this.knownRowCount);
    }

    const { controller, context } = this.createRequestContext();

    try {
      const response = await this.options.dataSource.getRows(
        createRemoteRequest(this.kind, this.requestState, this.options, request, startRow, endRow, requestVersion),
        context
      );
      if (requestVersion !== this.requestVersion) {
        return cloneRowResult(this.lastResolvedResult);
      }

      const { rowCount, result } = applyRemoteResponse(this.knownRowCount, response);
      this.knownRowCount = rowCount;
      this.lastResolvedResult = cloneRowResult(result);
      return result;
    } catch (error) {
      if (controller.signal.aborted && requestVersion !== this.requestVersion && isAbortError(error)) {
        return cloneRowResult(this.lastResolvedResult);
      }
      throw error;
    } finally {
      this.pendingRequestControllers.delete(controller);
    }
  }

  refresh(reason: RowModelRefreshReason): void {
    this.requestVersion += 1;
    this.abortPendingRequests();
    this.options.onRefresh?.(reason);
  }

  updateRequest(update: RemoteRowModelUpdate): void {
    this.requestState = applyRemoteRowModelUpdate(this.requestState, update);
    this.refresh("manual");
  }

  getRequestModel(): RemoteRowModelUpdate {
    return cloneRemoteRowModelUpdate(this.requestState);
  }

  async buildPivotSheet(input: PivotSheetInput): Promise<WorkbookDataInput> {
    const requestVersion = ++this.requestVersion;
    this.abortPendingRequests();
    return requestRemotePivotSheet(
      this.kind,
      this.requestState,
      this.options,
      input,
      requestVersion,
      () => this.requestVersion,
      () => this.createRequestContext(),
      (controller) => {
        this.pendingRequestControllers.delete(controller);
      }
    );
  }

  dispose(): void {
    this.requestVersion += 1;
    this.abortPendingRequests();
    this.options.onDispose?.();
  }
}