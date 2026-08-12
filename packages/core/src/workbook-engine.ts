import { CommandBus } from "./command-bus";
import { AddSheetCommand } from "./commands/add-sheet-command";
import { ChangeChartLegendCommand } from "./commands/change-chart-legend-command";
import { ChangeChartRangeCommand } from "./commands/change-chart-range-command";
import { ChangeChartTitleCommand } from "./commands/change-chart-title-command";
import { ChangeChartTypeCommand } from "./commands/change-chart-type-command";
import { CreateChartCommand } from "./commands/create-chart-command";
import { DeleteAxisCommand } from "./commands/delete-axis-command";
import { DeleteChartCommand } from "./commands/delete-chart-command";
import { DeleteSheetCommand } from "./commands/delete-sheet-command";
import { InsertAxisCommand } from "./commands/insert-axis-command";
import { MoveChartCommand } from "./commands/move-chart-command";
import { ResizeChartCommand } from "./commands/resize-chart-command";
import { SetCellValueCommand } from "./commands/set-cell-value-command";
import { SelectRangeCommand } from "./commands/select-range-command";
import { UpdateChartCommand } from "./commands/update-chart-command";
import { UpdateSheetOperationsCommand } from "./commands/update-sheet-operations-command";
import { CellValidationError } from "./validation/cell-validation-error";
import type {
  CellAddress,
  CellBatchUpdate,
  CellPrimitive,
  CellTransactionChange,
  CellModel,
  CellRange,
  CellStyle,
  CellValidationConfig,
  CellValidationResult,
  ChartPosition,
  ChartRangeBinding,
  ColumnSchema,
  ConditionalFormattingRule,
  FormulaEngine,
  PivotDerivedViewDefinition,
  PivotDerivedViewResult,
  PivotBuildAsyncOptions,
  PivotInferenceInput,
  PivotModule,
  PivotSheetInput,
  RegisteredCellValidator,
  RowSchema,
  SafeCellValidator,
  SheetMerge,
  WorksheetChartObject,
  WorksheetChartObjectInput,
  WorksheetChartType,
  SpreadsheetEventMap,
  SpreadsheetOperation,
  WorkbookConfig,
  WorkbookDataInput,
  WorkbookModel,
  WorkbookSettings
} from "./domain/types";
import { TypedEventEmitter } from "./events/typed-event-emitter";
import { HistoryManager } from "./history/history-manager";
import { PluginManager } from "./plugins/plugin-manager";
import type { GridPlugin, PluginState, RegisteredGridPlugin } from "./plugins/types";
import { cellLabelToAddress } from "./utils/address";
import { createId } from "./utils/id";
import { cloneValue } from "./utils/clone";
import { getCellKey } from "./utils/cell-key";
import { applyOperationsToWorkbook } from "./utils/apply-operations";
import { getConditionalFormattingStyle } from "./conditional-formatting/evaluator";
import {
  createCoreOperationError,
  createSheetNotFoundError,
  createWorkbookSheetRequiredError,
  SpreadsheetOperationError
} from "./errors/spreadsheet-operation-error";
import { defaultPivotModule } from "./pivot/module";
import { deriveFormulaRecalculationTargets, recalculateWorkbookFormulas } from "./utils/recalculate-formulas";
import { ValidationRegistry } from "./validation/validation-registry";
import { ClientSideRowModel, InfiniteRowModel, ServerSideRowModel } from "./row-model";
import type { RemotePivotCapableRowModel, RemoteRowModelUpdate, RowModel } from "./row-model";

const normalizeRange = (start: CellAddress, end: CellAddress): CellRange => ({
  start: {
    row: Math.min(start.row, end.row),
    col: Math.min(start.col, end.col)
  },
  end: {
    row: Math.max(start.row, end.row),
    col: Math.max(start.col, end.col)
  }
});

const isAddressWithinRange = (address: CellAddress, range: CellRange): boolean =>
  address.row >= range.start.row &&
  address.row <= range.end.row &&
  address.col >= range.start.col &&
  address.col <= range.end.col;

const rangesOverlap = (left: CellRange, right: CellRange): boolean =>
  left.start.row <= right.end.row &&
  left.end.row >= right.start.row &&
  left.start.col <= right.end.col &&
  left.end.col >= right.start.col;

const areGroupPathsEqual = (left: string[], right: string[]): boolean =>
  left.length === right.length && left.every((segment, index) => segment === right[index]);

const DEFAULT_SETTINGS: WorkbookSettings = {
  maxRows: 1000,
  maxColumns: 100,
  maxCellLength: 5000,
  maxFormulaLength: 2048,
  maxPasteCells: 10000,
  maxRecalcCells: 10000,
  maxPivotSourceRows: 5000,
  rowHeight: 28,
  columnWidth: 120,
  viewportBuffer: 6,
  maxHistorySize: 100,
  enableFormulas: true,
  clipboardPolicy: "text-only"
};

const createDefaultSelection = (): CellRange => ({
  start: { row: 0, col: 0 },
  end: { row: 0, col: 0 }
});

const createRangeError = (message: string): Error => {
  const error = new Error(message);
  error.name = "RangeError";
  return error;
};

const clonePivotInput = (input: PivotSheetInput): PivotSheetInput => ({
  ...input,
  sourceRange: {
    start: { ...input.sourceRange.start },
    end: { ...input.sourceRange.end }
  },
  rows: input.rows ? [...input.rows] : undefined,
  columns: input.columns ? [...input.columns] : undefined,
  values: input.values.map((value) => ({ ...value }))
});

const clonePivotDerivedViewDefinition = (definition: PivotDerivedViewDefinition): PivotDerivedViewDefinition => ({
  ...definition,
  input: clonePivotInput(definition.input),
  result: definition.result ? { ...definition.result } : undefined
});

const createPivotDerivedViewResult = (
  sheet: WorkbookDataInput,
  executionMode: PivotDerivedViewResult["executionMode"],
  remote: boolean
): PivotDerivedViewResult => ({
  executionMode,
  remote,
  rowCount: sheet.rowCount,
  columnCount: sheet.columnCount
});

const createPivotDerivedViewDefinition = (
  input: PivotSheetInput,
  options?: Partial<Omit<PivotDerivedViewDefinition, "kind" | "input" | "refreshedAt">>
): PivotDerivedViewDefinition => ({
  kind: "pivot",
  input: clonePivotInput(input),
  refreshedAt: Date.now(),
  autoRefresh: options?.autoRefresh ?? true,
  stale: options?.stale ?? false,
  refreshStatus: options?.refreshStatus ?? "idle",
  ...(options?.lastError ? { lastError: options.lastError } : {}),
  ...(options?.result ? { result: { ...options.result } } : {})
});

const readPivotDerivedViewDefinition = (metadata: Record<string, unknown> | undefined): PivotDerivedViewDefinition | undefined => {
  const derivedView = metadata?.derivedView;
  if (!derivedView || typeof derivedView !== "object") {
    return undefined;
  }

  const candidate = derivedView as Partial<PivotDerivedViewDefinition>;
  return candidate.kind === "pivot" && candidate.input ? (candidate as PivotDerivedViewDefinition) : undefined;
};

const isNoOpCellWrite = (previousCell: CellModel | undefined, value: CellPrimitive): boolean => {
  const nextIsFormula = typeof value === "string" && value.startsWith("=");
  if (!previousCell) {
    return value === null;
  }

  return nextIsFormula
    ? previousCell.formula === value
    : previousCell.formula === undefined && previousCell.value === value;
};

const parseCellKey = (key: string): CellAddress | undefined => {
  const match = /^(\d+):(\d+)$/.exec(key);
  if (!match) {
    return undefined;
  }

  return {
    row: Number(match[1]),
    col: Number(match[2])
  };
};

const parseChartRangeAddress = (address: string): CellRange | undefined => {
  const normalizedAddress = address.trim();
  if (!normalizedAddress) {
    return undefined;
  }
  const withoutSheetRef = normalizedAddress.includes("!") ? normalizedAddress.slice(normalizedAddress.lastIndexOf("!") + 1) : normalizedAddress;
  const parts = withoutSheetRef.split(":").map((part) => part.trim());
  if (parts.length === 0 || parts.length > 2) {
    return undefined;
  }
  const startLabel = parts[0];
  const endLabel = parts[1];
  if (!startLabel || (parts.length === 2 && !endLabel)) {
    return undefined;
  }
  try {
    const start = cellLabelToAddress(startLabel);
    const end = endLabel ? cellLabelToAddress(endLabel) : start;
    return normalizeRange(start, end);
  } catch {
    return undefined;
  }
};

const createWorkbook = (config: WorkbookConfig): WorkbookModel => {
  const settings: WorkbookSettings = {
    ...DEFAULT_SETTINGS,
    ...config.settings
  };

  const inputSheets = config.data?.length
    ? config.data
    : [
        {
          name: "Sheet1"
        }
      ];

  const sheets = inputSheets.map((sheetInput, index) => ({
    id: sheetInput.id ?? createId(`sheet-${index + 1}`),
    name: sheetInput.name ?? `Sheet${index + 1}`,
    cells: cloneValue(sheetInput.cells ?? {}),
    merges: cloneValue(sheetInput.merges ?? []),
    conditionalFormats: cloneValue(sheetInput.conditionalFormats ?? []),
    frozenRows: Math.max(0, sheetInput.frozenRows ?? 0),
    frozenColumns: Math.max(0, sheetInput.frozenColumns ?? 0),
    columns: cloneValue(sheetInput.columns ?? {}),
    rows: cloneValue(sheetInput.rows ?? {}),
    rowCount: Math.min(sheetInput.rowCount ?? 200, settings.maxRows),
    columnCount: Math.min(sheetInput.columnCount ?? 26, settings.maxColumns),
    selection: createDefaultSelection(),
    charts: cloneValue(sheetInput.charts ?? []),
    metadata: cloneValue(sheetInput.metadata ?? {})
  }));

  if (sheets.length === 0) {
    throw createWorkbookSheetRequiredError();
  }

  return {
    id: createId("workbook"),
    sheets,
    activeSheetId: sheets[0].id,
    metadata: config.metadata ?? {},
    settings
  };
};

const normalizeWorkbookSnapshot = (snapshot: WorkbookModel): WorkbookModel => {
  if (!snapshot.sheets.length) {
    throw createWorkbookSheetRequiredError();
  }

  const cloned = cloneValue(snapshot);
  for (const sheet of cloned.sheets) {
    sheet.merges ??= [];
    sheet.conditionalFormats ??= [];
    sheet.frozenRows ??= 0;
    sheet.frozenColumns ??= 0;
    sheet.columns ??= {};
    sheet.rows ??= {};
    sheet.charts ??= [];
    sheet.metadata ??= {};
  }

  return {
    ...cloned,
    activeSheetId: snapshot.sheets.some((sheet) => sheet.id === snapshot.activeSheetId)
      ? snapshot.activeSheetId
      : snapshot.sheets[0].id,
    settings: {
      ...DEFAULT_SETTINGS,
      ...snapshot.settings
    }
  };
};

export class WorkbookEngine {
  private readonly events = new TypedEventEmitter<SpreadsheetEventMap>();

  private readonly historyManager: HistoryManager;

  private readonly commandBus: CommandBus;

  private readonly pluginManager: PluginManager;

  private readonly validationRegistry: ValidationRegistry;

  private readonly explicitRowModels = new Map<string, RowModel>();

  private readonly clientSideRowModels = new Map<string, ClientSideRowModel>();

  private readonly pivotModule?: PivotModule;

  private readonly pendingDerivedPivotRefreshes = new Map<string, Promise<void>>();

  private readonly invalidatedDerivedPivotSheets = new Set<string>();

  private displayValueCacheRevision = -1;

  private readonly displayValueCache = new Map<string, string>();

  private disposed = false;

  constructor(
    config: WorkbookConfig = {},
    private readonly formulaEngine?: FormulaEngine
  ) {
    const workbook = createWorkbook(config);
    this.historyManager = new HistoryManager(workbook.settings.maxHistorySize);
    this.validationRegistry = new ValidationRegistry();
    this.commandBus = new CommandBus(
      workbook,
      this.historyManager,
      this.events,
      workbook.settings.enableFormulas ? this.formulaEngine : undefined,
      this.validationRegistry
    );
    this.pluginManager = new PluginManager(this);
    this.pivotModule = config.pivotModule === false ? undefined : (config.pivotModule ?? defaultPivotModule);
    this.events.on("command:completed", ({ sheetId, commandType }) => {
      if (!sheetId || commandType === "SelectRangeCommand" || this.disposed) {
        return;
      }

      this.handleDerivedPivotSourceMutation(sheetId);
    });
    this.events.emit("engine:created", {
      timestamp: Date.now(),
      workbookId: workbook.id
    });
  }

  static fromJSON(snapshot: WorkbookModel, formulaEngine?: FormulaEngine): WorkbookEngine {
    const engine = new WorkbookEngine({}, formulaEngine);
    engine.loadFromJSON(snapshot);
    return engine;
  }

  on<TKey extends keyof SpreadsheetEventMap>(
    eventName: TKey,
    listener: (payload: SpreadsheetEventMap[TKey]) => void
  ): () => void {
    return this.events.on(eventName, listener);
  }

  registerPlugin(plugin: GridPlugin, enabled = true): void {
    this.pluginManager.register(plugin, enabled);
  }

  unregisterPlugin(pluginId: string): void {
    this.pluginManager.unregister(pluginId);
  }

  enablePlugin(pluginId: string): void {
    this.pluginManager.enable(pluginId);
  }

  disablePlugin(pluginId: string): void {
    this.pluginManager.disable(pluginId);
  }

  isPluginEnabled(pluginId: string): boolean {
    return this.pluginManager.isEnabled(pluginId);
  }

  getRegisteredPlugins(): RegisteredGridPlugin[] {
    return this.pluginManager.list();
  }

  getPluginState<TState extends PluginState = PluginState>(pluginId: string): TState | undefined {
    return this.pluginManager.getState<TState>(pluginId);
  }

  getSnapshot(): WorkbookModel {
    return this.commandBus.getSnapshot();
  }

  private getWorkbookState(): Readonly<WorkbookModel> {
    return this.commandBus.peekWorkbook();
  }

  private getWorkbookRevision(): number {
    return this.commandBus.getRevision();
  }

  private getSheetState(sheetId: string) {
    return this.getWorkbookState().sheets.find((item) => item.id === sheetId);
  }

  private getCellState(sheetId: string, row: number, col: number): CellModel | undefined {
    return this.getSheetState(sheetId)?.cells[getCellKey(row, col)];
  }

  private syncDisplayValueCache(): void {
    const revision = this.getWorkbookRevision();
    if (revision === this.displayValueCacheRevision) {
      return;
    }

    this.displayValueCacheRevision = revision;
    this.displayValueCache.clear();
  }

  getActiveSheet() {
    const workbook = this.getWorkbookState();
    return cloneValue(workbook.sheets.find((sheet) => sheet.id === workbook.activeSheetId) ?? workbook.sheets[0]);
  }

  getCell(sheetId: string, row: number, col: number): CellModel | undefined {
    const cell = this.getCellState(sheetId, row, col);
    return cell ? cloneValue(cell) : undefined;
  }

  getCellValidation(sheetId: string, row: number, col: number): CellValidationConfig | undefined {
    return this.getCell(sheetId, row, col)?.validation;
  }

  getConditionalFormattingRules(sheetId: string): ConditionalFormattingRule[] {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    return cloneValue(sheet.conditionalFormats ?? []);
  }

  getConditionalStyle(sheetId: string, row: number, col: number): CellStyle | undefined {
    return getConditionalFormattingStyle(this.getSnapshot(), sheetId, row, col, this.formulaEngine);
  }

  validateCellValue(input: {
    sheetId: string;
    row: number;
    col: number;
    value: string | number | boolean | null;
  }): CellValidationResult {
    const workbook = this.getSnapshot();
    const sheet = workbook.sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    const cell = sheet.cells[getCellKey(input.row, input.col)];
    return this.validationRegistry.validateCellValue({
      workbook,
      sheet,
      cell,
      row: input.row,
      col: input.col,
      value: input.value,
      validation: cell?.validation
    });
  }

  registerValidator(id: string, validator: SafeCellValidator): void {
    this.validationRegistry.registerValidator(id, validator);
  }

  unregisterValidator(id: string): void {
    this.validationRegistry.unregisterValidator(id);
  }

  getRegisteredValidators(): RegisteredCellValidator[] {
    return this.validationRegistry.listValidators();
  }

  getDisplayValue(sheetId: string, row: number, col: number): string {
    this.syncDisplayValueCache();
    const cacheKey = `${sheetId}:${row}:${col}`;
    const cached = this.displayValueCache.get(cacheKey);
    if (cached !== undefined) {
      return cached;
    }

    const cell = this.getCellState(sheetId, row, col);

    if (!cell) {
      return "";
    }

    if (cell.error) {
      const errorValue = `#${cell.error.code}`;
      this.displayValueCache.set(cacheKey, errorValue);
      return errorValue;
    }

    const value = cell.formula ? cell.computedValue : cell.value;
    const displayValue = value == null ? "" : String(value);
    this.displayValueCache.set(cacheKey, displayValue);
    return displayValue;
  }

  getMerge(sheetId: string, row: number, col: number): SheetMerge | undefined {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      return undefined;
    }

    const merge = sheet.merges.find((item) => isAddressWithinRange({ row, col }, item));
    return merge ? cloneValue(merge) : undefined;
  }

  getColumnSchema(sheetId: string, col: number): ColumnSchema | undefined {
    const schema = this.getSheetState(sheetId)?.columns[col];
    return schema ? cloneValue(schema) : undefined;
  }

  getRowSchema(sheetId: string, row: number): RowSchema | undefined {
    const schema = this.getSheetState(sheetId)?.rows[row];
    return schema ? cloneValue(schema) : undefined;
  }

  getFrozenPane(sheetId: string): { rows: number; columns: number } {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    return {
      rows: sheet.frozenRows ?? 0,
      columns: sheet.frozenColumns ?? 0
    };
  }

  private emitRowModelChanged(sheetId: string): void {
    this.events.emit("row-model:changed", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId
    });
  }

  setRowModel(sheetId: string, rowModel: RowModel): void {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    const previous = this.explicitRowModels.get(sheetId);
    if (previous && previous !== rowModel) {
      previous.dispose();
    }

    this.explicitRowModels.set(sheetId, rowModel);
    rowModel.refresh("initial");
    this.emitRowModelChanged(sheetId);
  }

  clearRowModel(sheetId: string): void {
    const previous = this.explicitRowModels.get(sheetId);
    previous?.dispose();
    this.explicitRowModels.delete(sheetId);
    this.emitRowModelChanged(sheetId);
  }

  getRowModel(sheetId: string): RowModel {
    const explicit = this.explicitRowModels.get(sheetId);
    if (explicit) {
      return explicit;
    }

    let clientSide = this.clientSideRowModels.get(sheetId);
    if (!clientSide) {
      clientSide = new ClientSideRowModel(this, sheetId);
      this.clientSideRowModels.set(sheetId, clientSide);
    }

    return clientSide;
  }

  updateRemoteRowModel(sheetId: string, update: RemoteRowModelUpdate): void {
    const rowModel = this.explicitRowModels.get(sheetId);
    if (!rowModel) {
      throw createCoreOperationError(
        "CORE_REMOTE_ROW_MODEL_REQUIRED",
        `Sheet does not have an explicit remote row model: ${sheetId}`,
        { sheetId }
      );
    }

    if (rowModel instanceof InfiniteRowModel || rowModel instanceof ServerSideRowModel) {
      rowModel.updateRequest(update);
      this.emitRowModelChanged(sheetId);
      return;
    }

    throw createCoreOperationError(
      "CORE_REMOTE_ROW_MODEL_NOT_UPDATEABLE",
      `Sheet does not use an updateable remote row model: ${sheetId}`,
      { sheetId }
    );
  }

  setRemoteGroupExpanded(sheetId: string, path: string[], expanded: boolean): void {
    if (!path.length) {
      throw createCoreOperationError("CORE_REMOTE_GROUP_PATH_EMPTY", "Remote group path cannot be empty.");
    }

    const current = this.getRemoteRowModelRequest(sheetId);
    if (!current) {
      throw createCoreOperationError(
        "CORE_REMOTE_ROW_MODEL_REQUIRED",
        `Sheet does not have an explicit remote row model: ${sheetId}`,
        { sheetId }
      );
    }

    const currentPaths = current.expandedGroupPaths ?? [];
    const hasCurrentPath = currentPaths.some((item) => areGroupPathsEqual(item, path));
    let nextPaths = currentPaths;

    if (expanded) {
      if (!hasCurrentPath) {
        nextPaths = [...currentPaths, [...path]];
      }
    } else {
      nextPaths = currentPaths.filter((item) => !areGroupPathsEqual(item, path));
    }

    this.updateRemoteRowModel(sheetId, {
      expandedGroupPaths: nextPaths.length ? nextPaths : undefined
    });
  }

  getRemoteRowModelRequest(sheetId: string): RemoteRowModelUpdate | undefined {
    const rowModel = this.explicitRowModels.get(sheetId);
    if (rowModel instanceof InfiniteRowModel || rowModel instanceof ServerSideRowModel) {
      return rowModel.getRequestModel();
    }

    return undefined;
  }

  private getRemotePivotRowModel(sheetId: string): RemotePivotCapableRowModel | undefined {
    const rowModel = this.explicitRowModels.get(sheetId);
    if (rowModel instanceof InfiniteRowModel || rowModel instanceof ServerSideRowModel) {
      return rowModel;
    }

    return undefined;
  }

  private attachPivotDefinition(
    sheet: WorkbookDataInput,
    input: PivotSheetInput,
    options?: Partial<Omit<PivotDerivedViewDefinition, "kind" | "input" | "refreshedAt">>
  ): WorkbookDataInput {
    const executionMode = options?.result?.executionMode ?? input.executionMode ?? "client";
    const derivedView = createPivotDerivedViewDefinition(input, {
      ...options,
      result:
        options?.result ??
        createPivotDerivedViewResult(sheet, executionMode, executionMode === "server")
    });

    return {
      ...sheet,
      metadata: sheet.metadata
        ? {
            ...sheet.metadata,
            derivedView
          }
        : {
            derivedView
          }
    };
  }

  private updateSheetMetadata(sheetId: string, metadata: Record<string, unknown>): void {
    const workbook = this.getSnapshot();
    const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
    if (sheetIndex < 0) {
      throw createCoreOperationError("CORE_SHEET_NOT_FOUND", `Sheet not found: ${sheetId}`, { sheetId });
    }

    workbook.sheets[sheetIndex] = {
      ...workbook.sheets[sheetIndex],
      metadata: cloneValue(metadata)
    };

    this.commandBus.replaceWorkbook(workbook);
    this.emitRowModelChanged(sheetId);
  }

  private getPivotSheetViewDefinitionInternal(sheetId: string): PivotDerivedViewDefinition | undefined {
    const sheet = this.getSheetState(sheetId);
    return readPivotDerivedViewDefinition(sheet?.metadata);
  }

  private updatePivotSheetViewDefinition(
    sheetId: string,
    patch: Partial<Omit<PivotDerivedViewDefinition, "kind" | "input" | "refreshedAt">>
  ): void {
    const sheet = this.getSheetState(sheetId);
    const definition = readPivotDerivedViewDefinition(sheet?.metadata);
    if (!sheet || !definition) {
      return;
    }

    const nextDefinition = createPivotDerivedViewDefinition(definition.input, {
      autoRefresh: patch.autoRefresh ?? definition.autoRefresh ?? true,
      stale: patch.stale ?? definition.stale ?? false,
      refreshStatus: patch.refreshStatus ?? definition.refreshStatus ?? "idle",
      lastError: patch.lastError ?? definition.lastError,
      result: patch.result ?? definition.result
    });
    nextDefinition.refreshedAt = patch.refreshStatus === "refreshing" ? definition.refreshedAt : Date.now();

    this.updateSheetMetadata(sheetId, {
      ...(sheet.metadata ?? {}),
      derivedView: nextDefinition
    });
  }

  private getDerivedPivotSheetsForSource(sourceSheetId: string): Array<{ sheetId: string; definition: PivotDerivedViewDefinition }> {
    return this.getSnapshot().sheets.flatMap((sheet) => {
      const definition = readPivotDerivedViewDefinition(sheet.metadata);
      if (!definition || definition.input.sourceSheetId !== sourceSheetId || sheet.id === sourceSheetId) {
        return [];
      }

      return [{ sheetId: sheet.id, definition }];
    });
  }

  private handleDerivedPivotSourceMutation(sourceSheetId: string): void {
    for (const { sheetId, definition } of this.getDerivedPivotSheetsForSource(sourceSheetId)) {
      if (definition.autoRefresh === false) {
        this.updatePivotSheetViewDefinition(sheetId, {
          stale: true,
          refreshStatus: "idle",
          lastError: undefined
        });
        continue;
      }

      void this.queueDerivedPivotRefresh(sheetId);
    }
  }

  private async queueDerivedPivotRefresh(sheetId: string): Promise<void> {
    const pendingRefresh = this.pendingDerivedPivotRefreshes.get(sheetId);
    if (pendingRefresh) {
      this.invalidatedDerivedPivotSheets.add(sheetId);
      this.updatePivotSheetViewDefinition(sheetId, {
        stale: true,
        refreshStatus: "refreshing",
        lastError: undefined
      });
      return pendingRefresh;
    }

    const refreshPromise = (async () => {
      do {
        this.invalidatedDerivedPivotSheets.delete(sheetId);
        this.updatePivotSheetViewDefinition(sheetId, {
          stale: true,
          refreshStatus: "refreshing",
          lastError: undefined
        });

        try {
          await this.refreshPivotSheet(sheetId);
        } catch (error) {
          this.updatePivotSheetViewDefinition(sheetId, {
            stale: true,
            refreshStatus: "error",
            lastError: error instanceof Error ? error.message : "Falha ao atualizar a pivot derivada."
          });
          return;
        }
      } while (!this.disposed && this.invalidatedDerivedPivotSheets.has(sheetId));
    })().finally(() => {
      this.pendingDerivedPivotRefreshes.delete(sheetId);
      this.invalidatedDerivedPivotSheets.delete(sheetId);
    });

    this.pendingDerivedPivotRefreshes.set(sheetId, refreshPromise);
    return refreshPromise;
  }

  private replaceSheetData(sheetId: string, nextSheetInput: WorkbookDataInput): void {
    const workbook = this.getSnapshot();
    const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.id === sheetId);
    if (sheetIndex < 0) {
      throw createCoreOperationError("CORE_SHEET_NOT_FOUND", `Sheet not found: ${sheetId}`, { sheetId });
    }

    const currentSheet = workbook.sheets[sheetIndex];
    workbook.sheets[sheetIndex] = {
      ...currentSheet,
      name: nextSheetInput.name ?? currentSheet.name,
      cells: cloneValue(nextSheetInput.cells ?? {}),
      merges: cloneValue(nextSheetInput.merges ?? []),
      conditionalFormats: cloneValue(nextSheetInput.conditionalFormats ?? []),
      frozenRows: Math.max(0, nextSheetInput.frozenRows ?? 0),
      frozenColumns: Math.max(0, nextSheetInput.frozenColumns ?? 0),
      columns: cloneValue(nextSheetInput.columns ?? {}),
      rows: cloneValue(nextSheetInput.rows ?? {}),
      rowCount: Math.min(nextSheetInput.rowCount ?? currentSheet.rowCount, workbook.settings.maxRows),
      columnCount: Math.min(nextSheetInput.columnCount ?? currentSheet.columnCount, workbook.settings.maxColumns),
      charts: cloneValue(nextSheetInput.charts ?? currentSheet.charts ?? []),
      metadata: cloneValue(nextSheetInput.metadata ?? currentSheet.metadata ?? {})
    };

    this.commandBus.replaceWorkbook(workbook);
    this.emitRowModelChanged(sheetId);
  }

  private executeSheetOperations(
    sheetId: string,
    operations: SpreadsheetOperation[],
    affectedRanges: CellRange[]
  ): SpreadsheetOperation[] {
    const result = this.commandBus.execute(new UpdateSheetOperationsCommand(sheetId, operations, affectedRanges));
    return result.operations;
  }

  applyBatchOperations(input: {
    anchorSheetId: string;
    operations: SpreadsheetOperation[];
    affectedRanges?: CellRange[];
  }): SpreadsheetOperation[] {
    if (!input.operations.length) {
      return [];
    }

    const sheet = this.getSnapshot().sheets.find((item) => item.id === input.anchorSheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.anchorSheetId);
    }

    const operations = this.executeSheetOperations(input.anchorSheetId, input.operations, input.affectedRanges ?? [sheet.selection]);
    this.emitChartRangeChangesFromOperations(operations);
    return operations;
  }

  updateCells(input: {
    sheetId: string;
    updates: CellBatchUpdate[];
    affectedRanges?: CellRange[];
  }): SpreadsheetOperation[] {
    if (!input.updates.length) {
      return [];
    }

    const workbook = this.getSnapshot();
    const sheet = workbook.sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    let minRow = Number.POSITIVE_INFINITY;
    let maxRow = Number.NEGATIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxCol = Number.NEGATIVE_INFINITY;

    const normalizedUpdates = new Map<string, CellBatchUpdate>();
    for (const update of input.updates) {
      normalizedUpdates.set(getCellKey(update.row, update.col), update);
    }

    const operations = Array.from(normalizedUpdates.values()).flatMap((update) => {
      if (update.row < 0 || update.col < 0 || update.row >= sheet.rowCount || update.col >= sheet.columnCount) {
        throw createRangeError("Cell address is outside the current sheet bounds.");
      }

      const key = getCellKey(update.row, update.col);
      const previousCell = sheet.cells[key];
      const validationResult = this.validationRegistry.validateCellValue({
        workbook,
        sheet,
        cell: previousCell,
        row: update.row,
        col: update.col,
        value: update.value,
        validation: previousCell?.validation
      });

      if (!validationResult.valid && validationResult.error) {
        throw new CellValidationError(input.sheetId, { row: update.row, col: update.col }, validationResult.error);
      }

      const isNoOp = isNoOpCellWrite(previousCell, update.value);

      if (isNoOp) {
        return [];
      }

      minRow = Math.min(minRow, update.row);
      maxRow = Math.max(maxRow, update.row);
      minCol = Math.min(minCol, update.col);
      maxCol = Math.max(maxCol, update.col);

      const nextCell: CellModel = {
        ...previousCell,
        value: update.value,
        computedValue: update.value,
        error: undefined,
        formula: undefined
      };

      if (typeof update.value === "string" && update.value.startsWith("=")) {
        nextCell.formula = update.value;
        nextCell.computedValue = null;
      }

      return [
        {
          op: previousCell ? "replace" : "add",
          id: sheet.id,
          path: ["cells", key],
          value: nextCell
        } satisfies SpreadsheetOperation
      ];
    });

    if (!operations.length) {
      return [];
    }

    return this.applyBatchOperations({
      anchorSheetId: input.sheetId,
      operations,
      affectedRanges: input.affectedRanges ?? [
        {
          start: { row: minRow, col: minCol },
          end: { row: maxRow, col: maxCol }
        }
      ]
    });
  }

  applyCellTransaction(input: {
    sheetId: string;
    changes: CellTransactionChange[];
    affectedRanges?: CellRange[];
  }): SpreadsheetOperation[] {
    if (!input.changes.length) {
      return [];
    }

    const workbook = this.getWorkbookState();
    const sheet = workbook.sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    let minRow = Number.POSITIVE_INFINITY;
    let maxRow = Number.NEGATIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxCol = Number.NEGATIVE_INFINITY;

    const normalizedChanges = new Map<string, CellTransactionChange>();
    for (const change of input.changes) {
      normalizedChanges.set(change.key, change);
    }

    const operations: SpreadsheetOperation[] = [];

    for (const change of Array.from(normalizedChanges.values())) {
      const address = parseCellKey(change.key);
      if (!address) {
        throw createCoreOperationError("CORE_CELL_KEY_INVALID", `Invalid cell key: ${change.key}`, { key: change.key });
      }

      if (address.row < 0 || address.col < 0 || address.row >= sheet.rowCount || address.col >= sheet.columnCount) {
        throw createRangeError("Cell address is outside the current sheet bounds.");
      }

      const previousCell = sheet.cells[change.key];
      if (change.type === "remove") {
        if (!previousCell) {
          continue;
        }

        minRow = Math.min(minRow, address.row);
        maxRow = Math.max(maxRow, address.row);
        minCol = Math.min(minCol, address.col);
        maxCol = Math.max(maxCol, address.col);

        operations.push({
          op: "remove",
          id: sheet.id,
          path: ["cells", change.key],
          value: null
        });
        continue;
      }

      const validationResult = this.validationRegistry.validateCellValue({
        workbook,
        sheet,
        cell: previousCell,
        row: address.row,
        col: address.col,
        value: change.value,
        validation: previousCell?.validation
      });

      if (!validationResult.valid && validationResult.error) {
        throw new CellValidationError(input.sheetId, { row: address.row, col: address.col }, validationResult.error);
      }

      if (isNoOpCellWrite(previousCell, change.value)) {
        continue;
      }

      minRow = Math.min(minRow, address.row);
      maxRow = Math.max(maxRow, address.row);
      minCol = Math.min(minCol, address.col);
      maxCol = Math.max(maxCol, address.col);

      const nextCell: CellModel = {
        ...previousCell,
        value: change.value,
        computedValue: change.value,
        error: undefined,
        formula: undefined
      };

      if (typeof change.value === "string" && change.value.startsWith("=")) {
        nextCell.formula = change.value;
        nextCell.computedValue = null;
      }

      operations.push({
        op: previousCell ? "replace" : "add",
        id: sheet.id,
        path: ["cells", change.key],
        value: nextCell
      });
    }

    if (!operations.length) {
      return [];
    }

    return this.applyBatchOperations({
      anchorSheetId: input.sheetId,
      operations,
      affectedRanges: input.affectedRanges ?? [
        {
          start: { row: minRow, col: minCol },
          end: { row: maxRow, col: maxCol }
        }
      ]
    });
  }

  setCellValue(input: {
    sheetId: string;
    row: number;
    col: number;
    value: string | number | boolean | null;
  }): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    if (input.row < 0 || input.col < 0 || input.row >= sheet.rowCount || input.col >= sheet.columnCount) {
      throw createRangeError("Cell address is outside the current sheet bounds.");
    }

    if (isNoOpCellWrite(sheet.cells[getCellKey(input.row, input.col)], input.value)) {
      return [];
    }

    const result = this.commandBus.execute(new SetCellValueCommand(input));
    this.events.emit("cell:updated", {
      timestamp: Date.now(),
      workbookId: result.workbook.id,
      sheetId: input.sheetId,
      address: { row: input.row, col: input.col }
    });

    const cell = this.getCell(input.sheetId, input.row, input.col);
    if (cell?.error) {
      this.events.emit("formula:failed", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        address: { row: input.row, col: input.col },
        errorCode: cell.error.code
      });
    }

    this.emitChartRangeChangesFromOperations(result.operations);

    return result.operations;
  }

  setCellStyle(input: {
    sheetId: string;
    row: number;
    col: number;
    style: Partial<CellStyle>;
    mode?: "merge" | "replace";
  }): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    const merge = this.getMerge(input.sheetId, input.row, input.col);
    const target = merge?.start ?? { row: input.row, col: input.col };
    const key = getCellKey(target.row, target.col);
    const previousCell = sheet.cells[key];
    const nextStyle = input.mode === "replace" ? { ...input.style } : { ...previousCell?.style, ...input.style };
    const nextCell: CellModel = {
      ...previousCell,
      value: previousCell?.value ?? null,
      computedValue: previousCell?.computedValue ?? previousCell?.value ?? null,
      style: nextStyle
    };

    const operations = this.executeSheetOperations(
      input.sheetId,
      [
        {
          op: previousCell ? "replace" : "add",
          id: input.sheetId,
          path: ["cells", key],
          value: nextCell
        }
      ],
      [{ start: target, end: target }]
    );

    this.events.emit("cell:updated", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId: input.sheetId,
      address: target
    });

    return operations;
  }

  setCellValidation(input: {
    sheetId: string;
    row: number;
    col: number;
    validation?: CellValidationConfig;
  }): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    const key = getCellKey(input.row, input.col);
    const previousCell = sheet.cells[key];
    if (!previousCell && !input.validation) {
      return [];
    }

    const nextCell: CellModel = {
      ...previousCell,
      value: previousCell?.value ?? null,
      computedValue: previousCell?.computedValue ?? previousCell?.value ?? null,
      validation: input.validation ? cloneValue(input.validation) : undefined
    };

    if (!input.validation) {
      delete nextCell.validation;
    }

    return this.executeSheetOperations(
      input.sheetId,
      [
        {
          op: previousCell ? "replace" : "add",
          id: input.sheetId,
          path: ["cells", key],
          value: nextCell
        }
      ],
      [
        {
          start: { row: input.row, col: input.col },
          end: { row: input.row, col: input.col }
        }
      ]
    );
  }

  setConditionalFormattingRules(sheetId: string, rules: ConditionalFormattingRule[]): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    return this.executeSheetOperations(
      sheetId,
      [
        {
          op: "replace",
          id: sheetId,
          path: ["conditionalFormats"],
          value: cloneValue(rules)
        }
      ],
      [sheet.selection]
    );
  }

  resizeColumn(sheetId: string, col: number, width: number): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    const nextSchema: ColumnSchema = {
      ...sheet.columns[col],
      width
    };

    return this.executeSheetOperations(
      sheetId,
      [
        {
          op: sheet.columns[col] ? "replace" : "add",
          id: sheetId,
          path: ["columns", String(col)],
          value: nextSchema
        }
      ],
      [sheet.selection]
    );
  }

  resizeRow(sheetId: string, row: number, height: number): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    const nextSchema: RowSchema = {
      ...sheet.rows[row],
      height
    };

    return this.executeSheetOperations(
      sheetId,
      [
        {
          op: sheet.rows[row] ? "replace" : "add",
          id: sheetId,
          path: ["rows", String(row)],
          value: nextSchema
        }
      ],
      [sheet.selection]
    );
  }

  freezeRows(sheetId: string, count: number): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    return this.executeSheetOperations(
      sheetId,
      [
        {
          op: "replace",
          id: sheetId,
          path: ["frozenRows"],
          value: Math.min(Math.max(count, 0), sheet.rowCount)
        }
      ],
      [sheet.selection]
    );
  }

  freezeColumns(sheetId: string, count: number): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    return this.executeSheetOperations(
      sheetId,
      [
        {
          op: "replace",
          id: sheetId,
          path: ["frozenColumns"],
          value: Math.min(Math.max(count, 0), sheet.columnCount)
        }
      ],
      [sheet.selection]
    );
  }

  setRowsHidden(sheetId: string, start: number, end = start, hidden = true): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    const rowStart = Math.max(0, Math.min(start, end));
    const rowEnd = Math.min(sheet.rowCount - 1, Math.max(start, end));
    const operations: SpreadsheetOperation[] = [];

    for (let row = rowStart; row <= rowEnd; row += 1) {
      const nextSchema: RowSchema = {
        ...sheet.rows[row],
        hidden
      };

      if (!hidden) {
        delete nextSchema.hidden;
      }

      operations.push({
        op: sheet.rows[row] ? "replace" : "add",
        id: sheetId,
        path: ["rows", String(row)],
        value: nextSchema
      });
    }

    return this.executeSheetOperations(sheetId, operations, [sheet.selection]);
  }

  setColumnsHidden(sheetId: string, start: number, end = start, hidden = true): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    const colStart = Math.max(0, Math.min(start, end));
    const colEnd = Math.min(sheet.columnCount - 1, Math.max(start, end));
    const operations: SpreadsheetOperation[] = [];

    for (let col = colStart; col <= colEnd; col += 1) {
      const nextSchema: ColumnSchema = {
        ...sheet.columns[col],
        hidden
      };

      if (!hidden) {
        delete nextSchema.hidden;
      }

      operations.push({
        op: sheet.columns[col] ? "replace" : "add",
        id: sheetId,
        path: ["columns", String(col)],
        value: nextSchema
      });
    }

    return this.executeSheetOperations(sheetId, operations, [sheet.selection]);
  }

  mergeCells(input: { sheetId: string; start: CellAddress; end: CellAddress }): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    const range = normalizeRange(input.start, input.end);
    const nextMerges = sheet.merges.filter((merge) => !rangesOverlap(merge, range));
    nextMerges.push(range);

    const operations: SpreadsheetOperation[] = [
      {
        op: "replace",
        id: input.sheetId,
        path: ["merges"],
        value: nextMerges
      }
    ];

    for (let row = range.start.row; row <= range.end.row; row += 1) {
      for (let col = range.start.col; col <= range.end.col; col += 1) {
        if (row === range.start.row && col === range.start.col) {
          continue;
        }

        const key = getCellKey(row, col);
        if (sheet.cells[key]) {
          operations.push({
            op: "remove",
            id: input.sheetId,
            path: ["cells", key],
            value: null
          });
        }
      }
    }

    return this.executeSheetOperations(input.sheetId, operations, [range]);
  }

  unmergeCells(input: { sheetId: string; row: number; col: number }): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    const nextMerges = sheet.merges.filter((merge) => !isAddressWithinRange({ row: input.row, col: input.col }, merge));
    return this.executeSheetOperations(
      input.sheetId,
      [
        {
          op: "replace",
          id: input.sheetId,
          path: ["merges"],
          value: nextMerges
        }
      ],
      [sheet.selection]
    );
  }

  selectRange(input: {
    sheetId: string;
    rowStart: number;
    rowEnd: number;
    colStart: number;
    colEnd: number;
  }): void {
    this.commandBus.execute(new SelectRangeCommand(input));
    this.events.emit("selection:changed", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId: input.sheetId,
      range: {
        start: {
          row: Math.min(input.rowStart, input.rowEnd),
          col: Math.min(input.colStart, input.colEnd)
        },
        end: {
          row: Math.max(input.rowStart, input.rowEnd),
          col: Math.max(input.colStart, input.colEnd)
        }
      }
    });
  }

  undo(): boolean {
    const workbook = this.commandBus.undo();
    return Boolean(workbook);
  }

  redo(): boolean {
    const workbook = this.commandBus.redo();
    return Boolean(workbook);
  }

  setActiveSheet(sheetId: string): void {
    const workbook = this.getSnapshot();
    if (!workbook.sheets.some((sheet) => sheet.id === sheetId)) {
      throw createSheetNotFoundError(sheetId);
    }
    workbook.activeSheetId = sheetId;
    this.commandBus.replaceWorkbook(workbook);
  }

  reportSecurityEvent(reason: string, details?: Record<string, unknown>): void {
    this.events.emit("security:blocked-input", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      reason,
      details
    });
  }

  reportChartImported(sheetId: string, chart: WorksheetChartObject): void {
    this.events.emit("chart:imported", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId,
      chartId: chart.id,
      chart: cloneValue(chart)
    });
  }

  reportChartExported(sheetId: string, chartId: string): void {
    this.events.emit("chart:exported", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId,
      chartId
    });
  }

  reportChartUnsupportedFeature(sheetId: string, chartId: string, feature: string): void {
    this.events.emit("chart:unsupportedFeature", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId,
      chartId,
      feature
    });
  }

  reportChartRenderStarted(sheetId: string, chartId: string): void {
    this.events.emit("chart:renderStarted", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId,
      chartId
    });
  }

  reportChartRenderFinished(sheetId: string, chartId: string, durationMs?: number): void {
    this.events.emit("chart:renderFinished", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId,
      chartId,
      durationMs
    });
  }

  reportChartRenderSkipped(sheetId: string, chartId: string, reason: string): void {
    this.events.emit("chart:renderSkipped", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId,
      chartId,
      reason
    });
  }

  reportChartError(input: {
    errorCode: string;
    message: string;
    sheetId?: string;
    chartId?: string;
  }): void {
    this.events.emit("chart:error", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId: input.sheetId,
      chartId: input.chartId,
      errorCode: input.errorCode,
      message: input.message
    });
  }

  private emitChartExportedEvents(workbook: WorkbookModel): void {
    for (const sheet of workbook.sheets) {
      for (const chart of sheet.charts ?? []) {
        this.reportChartExported(sheet.id, chart.id);
      }
    }
  }

  private emitChartImportedEvents(workbook: WorkbookModel): void {
    for (const sheet of workbook.sheets) {
      for (const chart of sheet.charts ?? []) {
        this.reportChartImported(sheet.id, chart);
        for (const unsupportedFeature of chart.excelInterop?.unsupportedFeatures ?? []) {
          this.reportChartUnsupportedFeature(sheet.id, chart.id, unsupportedFeature);
        }
      }
    }
  }

  toJSON(options?: { emitChartExportEvents?: boolean }): WorkbookModel {
    const snapshot = this.getSnapshot();
    if (options?.emitChartExportEvents) {
      this.emitChartExportedEvents(snapshot);
    }
    return snapshot;
  }

  loadFromJSON(snapshot: WorkbookModel): void {
    const normalized = normalizeWorkbookSnapshot(snapshot);
    this.commandBus.replaceWorkbook(recalculateWorkbookFormulas(normalized, this.formulaEngine));
    this.emitChartImportedEvents(normalized);
  }

  applyOperations(operations: SpreadsheetOperation[]): void {
    const snapshot = this.getSnapshot();
    const nextWorkbook = recalculateWorkbookFormulas(
      applyOperationsToWorkbook(snapshot, operations),
      this.formulaEngine,
      deriveFormulaRecalculationTargets(operations)
    );
    this.commandBus.replaceWorkbook(nextWorkbook);
    this.events.emit("command:completed", {
      timestamp: Date.now(),
      workbookId: nextWorkbook.id,
      durationMs: 0,
      commandType: "ApplyOperations",
      operations
    });
    this.emitChartRangeChangesFromOperations(operations);
  }

  getSelection(sheetId: string): CellRange {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }
    return sheet.selection;
  }

  getCharts(sheetId: string): WorksheetChartObject[] {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }
    return cloneValue(sheet.charts ?? []);
  }

  getChart(sheetId: string, chartId: string): WorksheetChartObject | undefined {
    return this.getCharts(sheetId).find((chart) => chart.id === chartId);
  }

  createChart(input: { sheetId: string; chart: WorksheetChartObjectInput }): SpreadsheetOperation[] {
    const sheet = this.getSheetState(input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    const chartId = input.chart.id?.trim() || createId("chart");
    const sourceRange = input.chart.sourceRange
      ? {
          ...cloneValue(input.chart.sourceRange),
          chartId,
          sheetId: input.sheetId
        }
      : undefined;
    const chart: WorksheetChartObject = {
      id: chartId,
      sheetId: input.sheetId,
      type: input.chart.type,
      title: input.chart.title,
      sourceRange,
      figure: cloneValue(input.chart.figure),
      position: {
        fromCell: input.chart.position.fromCell,
        toCell: input.chart.position.toCell,
        offsetX: Number(input.chart.position.offsetX) || 0,
        offsetY: Number(input.chart.position.offsetY) || 0,
        width: Math.max(40, Number(input.chart.position.width) || 320),
        height: Math.max(40, Number(input.chart.position.height) || 220),
        zIndex: Math.max(0, Math.round(Number(input.chart.position.zIndex ?? 1) || 1))
      },
      style: input.chart.style ? cloneValue(input.chart.style) : undefined,
      state: {
        selected: input.chart.state?.selected ?? false,
        visible: input.chart.state?.visible ?? true,
        locked: input.chart.state?.locked ?? false,
        ...(typeof input.chart.state?.lastRenderedAt === "number"
          ? {
              lastRenderedAt: input.chart.state.lastRenderedAt
            }
          : {})
      },
      excelInterop: {
        ...(input.chart.excelInterop ? cloneValue(input.chart.excelInterop) : {})
      }
    };

    const result = this.commandBus.execute(
      new CreateChartCommand({
        sheetId: input.sheetId,
        chart
      })
    );
    this.events.emit("chart:created", {
      timestamp: Date.now(),
      workbookId: result.workbook.id,
      sheetId: input.sheetId,
      chartId,
      chart: cloneValue(chart)
    });
    if (chart.state.selected) {
      this.events.emit("chart:selected", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId
      });
    }
    if (sourceRange) {
      this.events.emit("chart:rangeChanged", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId,
        range: cloneValue(sourceRange),
        reason: "binding-updated"
      });
    }
    return result.operations;
  }

  updateChart(input: {
    sheetId: string;
    chartId: string;
    patch: Partial<Omit<WorksheetChartObject, "id" | "sheetId">>;
  }): SpreadsheetOperation[] {
    const previousChart = this.getChart(input.sheetId, input.chartId);
    const result = this.commandBus.execute(new UpdateChartCommand(input));
    const nextChart = this.getChart(input.sheetId, input.chartId);
    if (!nextChart) {
      throw createCoreOperationError("CORE_CHART_NOT_FOUND", `Chart not found: ${input.chartId}`, {
        sheetId: input.sheetId,
        chartId: input.chartId
      });
    }

    this.events.emit("chart:updated", {
      timestamp: Date.now(),
      workbookId: result.workbook.id,
      sheetId: input.sheetId,
      chartId: input.chartId,
      chart: cloneValue(nextChart)
    });

    if (previousChart && previousChart.state.selected !== nextChart.state.selected) {
      this.events.emit(nextChart.state.selected ? "chart:selected" : "chart:unselected", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId: input.chartId
      });
    }

    if (input.patch.sourceRange) {
      this.events.emit("chart:rangeChanged", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId: input.chartId,
        range: cloneValue(input.patch.sourceRange),
        reason: "binding-updated"
      });
    }

    return result.operations;
  }

  moveChart(input: {
    sheetId: string;
    chartId: string;
    position: Pick<ChartPosition, "fromCell" | "toCell" | "offsetX" | "offsetY" | "zIndex">;
  }): SpreadsheetOperation[] {
    const result = this.commandBus.execute(new MoveChartCommand(input));
    const nextChart = this.getChart(input.sheetId, input.chartId);
    if (!nextChart) {
      throw createCoreOperationError("CORE_CHART_NOT_FOUND", `Chart not found: ${input.chartId}`, {
        sheetId: input.sheetId,
        chartId: input.chartId
      });
    }
    this.events.emit("chart:moved", {
      timestamp: Date.now(),
      workbookId: result.workbook.id,
      sheetId: input.sheetId,
      chartId: input.chartId,
      position: cloneValue(nextChart.position)
    });
    return result.operations;
  }

  resizeChart(input: {
    sheetId: string;
    chartId: string;
    position: Pick<ChartPosition, "width" | "height" | "toCell">;
  }): SpreadsheetOperation[] {
    const result = this.commandBus.execute(new ResizeChartCommand(input));
    const nextChart = this.getChart(input.sheetId, input.chartId);
    if (!nextChart) {
      throw createCoreOperationError("CORE_CHART_NOT_FOUND", `Chart not found: ${input.chartId}`, {
        sheetId: input.sheetId,
        chartId: input.chartId
      });
    }
    this.events.emit("chart:resized", {
      timestamp: Date.now(),
      workbookId: result.workbook.id,
      sheetId: input.sheetId,
      chartId: input.chartId,
      position: cloneValue(nextChart.position)
    });
    return result.operations;
  }

  deleteChart(input: { sheetId: string; chartId: string }): SpreadsheetOperation[] {
    const previousChart = this.getChart(input.sheetId, input.chartId);
    const result = this.commandBus.execute(new DeleteChartCommand(input));
    if (previousChart?.state.selected) {
      this.events.emit("chart:unselected", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId: input.chartId
      });
    }
    this.events.emit("chart:deleted", {
      timestamp: Date.now(),
      workbookId: result.workbook.id,
      sheetId: input.sheetId,
      chartId: input.chartId
    });
    return result.operations;
  }

  changeChartType(input: { sheetId: string; chartId: string; chartType: WorksheetChartType }): SpreadsheetOperation[] {
    const result = this.commandBus.execute(new ChangeChartTypeCommand(input));
    const nextChart = this.getChart(input.sheetId, input.chartId);
    if (nextChart) {
      this.events.emit("chart:updated", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId: input.chartId,
        chart: cloneValue(nextChart)
      });
    }
    return result.operations;
  }

  changeChartRange(input: {
    sheetId: string;
    chartId: string;
    sourceRange: Omit<ChartRangeBinding, "chartId" | "sheetId">;
  }): SpreadsheetOperation[] {
    const nextRange: ChartRangeBinding = {
      ...cloneValue(input.sourceRange),
      chartId: input.chartId,
      sheetId: input.sheetId
    };
    const result = this.commandBus.execute(
      new ChangeChartRangeCommand({
        sheetId: input.sheetId,
        chartId: input.chartId,
        sourceRange: nextRange
      })
    );
    const nextChart = this.getChart(input.sheetId, input.chartId);
    if (nextChart) {
      this.events.emit("chart:updated", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId: input.chartId,
        chart: cloneValue(nextChart)
      });
      this.events.emit("chart:rangeChanged", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId: input.chartId,
        range: cloneValue(nextRange),
        reason: "binding-updated"
      });
    }
    return result.operations;
  }

  changeChartTitle(input: { sheetId: string; chartId: string; title?: string }): SpreadsheetOperation[] {
    const result = this.commandBus.execute(new ChangeChartTitleCommand(input));
    const nextChart = this.getChart(input.sheetId, input.chartId);
    if (nextChart) {
      this.events.emit("chart:updated", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId: input.chartId,
        chart: cloneValue(nextChart)
      });
    }
    return result.operations;
  }

  changeChartLegend(input: { sheetId: string; chartId: string; visible: boolean }): SpreadsheetOperation[] {
    const result = this.commandBus.execute(new ChangeChartLegendCommand(input));
    const nextChart = this.getChart(input.sheetId, input.chartId);
    if (nextChart) {
      this.events.emit("chart:updated", {
        timestamp: Date.now(),
        workbookId: result.workbook.id,
        sheetId: input.sheetId,
        chartId: input.chartId,
        chart: cloneValue(nextChart)
      });
    }
    return result.operations;
  }

  private emitChartRangeChangesFromOperations(operations: SpreadsheetOperation[]): void {
    if (!operations.length) {
      return;
    }
    const rangesBySheet = new Map<string, CellRange[]>();
    for (const operation of operations) {
      if (operation.path[0] !== "cells" || typeof operation.path[1] !== "string") {
        continue;
      }
      const address = parseCellKey(operation.path[1]);
      if (!address) {
        continue;
      }
      const list = rangesBySheet.get(operation.id) ?? [];
      list.push({
        start: { ...address },
        end: { ...address }
      });
      rangesBySheet.set(operation.id, list);
    }

    for (const [sheetId, ranges] of rangesBySheet.entries()) {
      this.emitLinkedChartRangeChanges(sheetId, ranges);
    }
  }

  private emitLinkedChartRangeChanges(sheetId: string, changedRanges: CellRange[]): void {
    if (!changedRanges.length) {
      return;
    }
    const sheet = this.getSheetState(sheetId);
    const charts = sheet?.charts;
    if (!sheet || !charts?.length) {
      return;
    }

    const workbookId = this.getSnapshot().id;
    const emitted = new Set<string>();
    for (const chart of charts) {
      if (!chart.sourceRange?.autoRefresh || chart.sourceRange.sheetId !== sheetId) {
        continue;
      }
      const sourceRange = parseChartRangeAddress(chart.sourceRange.rangeAddress);
      if (!sourceRange) {
        this.events.emit("chart:dataInvalid", {
          timestamp: Date.now(),
          workbookId,
          sheetId,
          chartId: chart.id,
          reason: `Invalid chart range '${chart.sourceRange.rangeAddress}'.`
        });
        continue;
      }
      if (emitted.has(chart.id)) {
        continue;
      }
      if (!changedRanges.some((changedRange) => rangesOverlap(changedRange, sourceRange))) {
        continue;
      }
      emitted.add(chart.id);
      this.events.emit("chart:rangeChanged", {
        timestamp: Date.now(),
        workbookId,
        sheetId,
        chartId: chart.id,
        range: cloneValue(chart.sourceRange),
        reason: "source-cells-updated"
      });
    }
  }

  private getPivotModule(): PivotModule {
    if (!this.pivotModule) {
      throw createCoreOperationError("CORE_PIVOT_MODULE_DISABLED", "Pivot module is not enabled for this workbook.");
    }

    return this.pivotModule;
  }

  createPivotSheet(input: PivotSheetInput): WorkbookDataInput {
    if (input.executionMode === "server") {
      throw createCoreOperationError(
        "CORE_PIVOT_SERVER_REQUIRES_ASYNC",
        "Server-side pivot materialization requires createPivotSheetAsync.",
        { sourceSheetId: input.sourceSheetId }
      );
    }

    return this.attachPivotDefinition(this.getPivotModule().createPivotSheet(this.getWorkbookState(), input), input);
  }

  async createPivotSheetAsync(input: PivotSheetInput, options?: PivotBuildAsyncOptions): Promise<WorkbookDataInput> {
    const remoteRowModel = this.getRemotePivotRowModel(input.sourceSheetId);
    if (input.executionMode === "server") {
      if (!remoteRowModel) {
        throw createCoreOperationError(
          "CORE_PIVOT_REMOTE_ROW_MODEL_REQUIRED",
          "Server-side pivot materialization requires an explicit remote row model.",
          { sourceSheetId: input.sourceSheetId }
        );
      }

      const remotePivotSheet = await remoteRowModel.buildPivotSheet(input);
      return this.attachPivotDefinition(remotePivotSheet, input, {
        result: createPivotDerivedViewResult(remotePivotSheet, "server", true)
      });
    }

    if (input.executionMode !== "client" && remoteRowModel) {
      const remotePivotSheet = await remoteRowModel.buildPivotSheet(input);
      return this.attachPivotDefinition(remotePivotSheet, input, {
        result: createPivotDerivedViewResult(remotePivotSheet, "server", true)
      });
    }

    const pivotModule = this.getPivotModule();
    if (pivotModule.createPivotSheetAsync) {
      const nextSheet = await pivotModule.createPivotSheetAsync(this.getWorkbookState(), input, options);
      return this.attachPivotDefinition(nextSheet, input, {
        result: createPivotDerivedViewResult(nextSheet, input.executionMode ?? "client", false)
      });
    }

    const nextSheet = pivotModule.createPivotSheet(this.getWorkbookState(), input);
    return this.attachPivotDefinition(nextSheet, input, {
      result: createPivotDerivedViewResult(nextSheet, input.executionMode ?? "client", false)
    });
  }

  inferPivotSheet(input: PivotInferenceInput): PivotSheetInput {
    return this.getPivotModule().inferPivotSheet(this.getWorkbookState(), input);
  }

  addPivotSheet(input: PivotSheetInput): string {
    return this.addSheet(this.createPivotSheet(input));
  }

  async addPivotSheetAsync(input: PivotSheetInput, options?: PivotBuildAsyncOptions): Promise<string> {
    return this.addSheet(await this.createPivotSheetAsync(input, options));
  }

  getPivotSheetViewDefinition(sheetId: string): PivotDerivedViewDefinition | undefined {
    const definition = this.getPivotSheetViewDefinitionInternal(sheetId);
    return definition ? clonePivotDerivedViewDefinition(definition) : undefined;
  }

  getPivotSheetDefinition(sheetId: string): PivotSheetInput | undefined {
    const definition = this.getPivotSheetViewDefinitionInternal(sheetId);
    return definition ? clonePivotInput(definition.input) : undefined;
  }

  setPivotSheetAutoRefresh(sheetId: string, autoRefresh: boolean): void {
    const definition = this.getPivotSheetViewDefinitionInternal(sheetId);
    if (!definition) {
      throw createCoreOperationError("CORE_PIVOT_SHEET_REQUIRED", `Sheet is not a persisted pivot view: ${sheetId}`, { sheetId });
    }

    this.updatePivotSheetViewDefinition(sheetId, {
      autoRefresh,
      stale: autoRefresh ? definition.stale ?? false : true,
      refreshStatus: autoRefresh ? definition.refreshStatus ?? "idle" : "idle",
      lastError: autoRefresh ? undefined : definition.lastError
    });
  }

  replaceSheet(sheetId: string, input: WorkbookDataInput): void {
    this.replaceSheetData(sheetId, input);
  }

  async updatePivotSheet(sheetId: string, input: PivotSheetInput, options?: PivotBuildAsyncOptions): Promise<void> {
    const currentDefinition = this.getPivotSheetViewDefinitionInternal(sheetId);
    if (!currentDefinition) {
      throw createCoreOperationError("CORE_PIVOT_SHEET_REQUIRED", `Sheet is not a persisted pivot view: ${sheetId}`, { sheetId });
    }

    const nextSheet = await this.createPivotSheetAsync(input, options);
    const nextDefinition = readPivotDerivedViewDefinition(nextSheet.metadata) ?? createPivotDerivedViewDefinition(input);
    this.replaceSheetData(sheetId, {
      ...nextSheet,
      id: sheetId,
      name: nextSheet.name ?? this.getSheetState(sheetId)?.name,
      metadata: {
        ...(nextSheet.metadata ?? {}),
        derivedView: createPivotDerivedViewDefinition(input, {
          autoRefresh: currentDefinition.autoRefresh ?? true,
          stale: false,
          refreshStatus: "idle",
          lastError: undefined,
          result: nextDefinition.result
        })
      }
    });
  }

  async refreshPivotSheet(sheetId: string, options?: PivotBuildAsyncOptions): Promise<void> {
    const definition = this.getPivotSheetViewDefinitionInternal(sheetId);
    if (!definition) {
      throw createCoreOperationError("CORE_PIVOT_SHEET_REQUIRED", `Sheet is not a persisted pivot view: ${sheetId}`, { sheetId });
    }

    const nextSheet = await this.createPivotSheetAsync(definition.input, options);
    const nextDefinition = readPivotDerivedViewDefinition(nextSheet.metadata) ?? createPivotDerivedViewDefinition(definition.input);
    this.replaceSheetData(sheetId, {
      ...nextSheet,
      id: sheetId,
      name: nextSheet.name ?? this.getSheetState(sheetId)?.name,
      metadata: {
        ...(nextSheet.metadata ?? {}),
        derivedView: createPivotDerivedViewDefinition(definition.input, {
          autoRefresh: definition.autoRefresh ?? true,
          stale: false,
          refreshStatus: "idle",
          lastError: undefined,
          result: nextDefinition.result
        })
      }
    });
  }

  addSheet(input?: WorkbookDataInput): string {
    const result = this.commandBus.execute(new AddSheetCommand(input));
    return result.workbook.activeSheetId;
  }

  deleteSheet(sheetId: string): void {
    this.clearRowModel(sheetId);
    this.clientSideRowModels.delete(sheetId);
    this.commandBus.execute(new DeleteSheetCommand(sheetId));
  }

  insertRows(sheetId: string, index: number, count = 1): void {
    this.commandBus.execute(
      new InsertAxisCommand({
        sheetId,
        axis: "row",
        index,
        count
      })
    );
  }

  deleteRows(sheetId: string, start: number, end = start): void {
    this.commandBus.execute(
      new DeleteAxisCommand({
        sheetId,
        axis: "row",
        start,
        end
      })
    );
  }

  insertColumns(sheetId: string, index: number, count = 1): void {
    this.commandBus.execute(
      new InsertAxisCommand({
        sheetId,
        axis: "column",
        index,
        count
      })
    );
  }

  deleteColumns(sheetId: string, start: number, end = start): void {
    this.commandBus.execute(
      new DeleteAxisCommand({
        sheetId,
        axis: "column",
        start,
        end
      })
    );
  }

  dispose(): void {
    if (this.disposed) {
      return;
    }
    this.disposed = true;
    this.pendingDerivedPivotRefreshes.clear();
    this.invalidatedDerivedPivotSheets.clear();
    for (const rowModel of this.explicitRowModels.values()) {
      rowModel.dispose();
    }
    this.explicitRowModels.clear();
    this.clientSideRowModels.clear();
    this.pluginManager.clear();
    this.events.emit("engine:disposed", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id
    });
    this.events.clear();
  }
}

export type { CellAddress, CellModel };