import { CommandBus } from "./command-bus";
import { AddSheetCommand } from "./commands/add-sheet-command";
import { DeleteAxisCommand } from "./commands/delete-axis-command";
import { DeleteSheetCommand } from "./commands/delete-sheet-command";
import { InsertAxisCommand } from "./commands/insert-axis-command";
import { SetCellValueCommand } from "./commands/set-cell-value-command";
import { SelectRangeCommand } from "./commands/select-range-command";
import { UpdateSheetOperationsCommand } from "./commands/update-sheet-operations-command";
import { CellValidationError } from "./validation/cell-validation-error";
import type {
  CellAddress,
  CellBatchUpdate,
  CellComment,
  CellCommentReply,
  ClientSideFilterDescriptor,
  ClientSideQueryState,
  ClientSideSortDescriptor,
  CellPrimitive,
  CellRichTextSegment,
  CellTransactionChange,
  CellModel,
  CellRange,
  CellStyle,
  CellValidationConfig,
  CellValidationResult,
  CollaborationAdapter,
  CollaborationConflictPolicy,
  CollaborationEnvelope,
  CollaborationPresence,
  CollaborationPresenceMessage,
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
  SheetSplitPane,
  SheetModel,
  SheetMerge,
  WorksheetObjectPosition,
  WorksheetImageObject,
  WorksheetImageObjectInput,
  WorksheetWidgetObject,
  WorksheetWidgetObjectInput,
  JsonValue,
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

const CLIENT_SIDE_QUERY_METADATA_KEY = "clientSideQuery";

const toComparableNumber = (value: CellPrimitive, type: "number" | "date"): number | undefined => {
  if (type === "date" && typeof value === "string") {
    const timestamp = Date.parse(value);
    return Number.isNaN(timestamp) ? undefined : timestamp;
  }

  const numeric = typeof value === "number" ? value : Number(value);
  return Number.isFinite(numeric) ? numeric : undefined;
};

const matchesClientSideFilter = (value: CellPrimitive, descriptor: ClientSideFilterDescriptor): boolean => {
  if (descriptor.type === "text") {
    const candidate = String(value ?? "");
    const expected = String(descriptor.value);
    const left = descriptor.caseSensitive ? candidate : candidate.toLocaleLowerCase();
    const right = descriptor.caseSensitive ? expected : expected.toLocaleLowerCase();
    if (descriptor.operator === "contains") return left.includes(right);
    if (descriptor.operator === "startsWith") return left.startsWith(right);
    if (descriptor.operator === "equals") return left === right;
    return false;
  }

  const candidate = toComparableNumber(value, descriptor.type);
  const expected = toComparableNumber(descriptor.value, descriptor.type);
  if (candidate === undefined || expected === undefined) return false;
  if (descriptor.operator === "equals") return candidate === expected;
  if (descriptor.operator === "gt") return candidate > expected;
  if (descriptor.operator === "gte") return candidate >= expected;
  if (descriptor.operator === "lt") return candidate < expected;
  if (descriptor.operator === "lte") return candidate <= expected;
  if (descriptor.operator === "between") {
    const upper = descriptor.valueTo === undefined ? undefined : toComparableNumber(descriptor.valueTo, descriptor.type);
    return upper !== undefined && candidate >= Math.min(expected, upper) && candidate <= Math.max(expected, upper);
  }
  return false;
};

const compareClientSideValues = (left: CellPrimitive, right: CellPrimitive): number => {
  if (left == null && right == null) return 0;
  if (left == null) return 1;
  if (right == null) return -1;
  if (typeof left === "number" && typeof right === "number") return left - right;
  return String(left).localeCompare(String(right), undefined, { numeric: true, sensitivity: "base" });
};

const DEFAULT_SETTINGS: WorkbookSettings = {
  maxRows: 1000,
  maxColumns: 100,
  maxCellLength: 5000,
  maxFormulaLength: 2048,
  maxPasteCells: 10000,
  maxImageSourceLength: 2_000_000,
  maxWidgetDataLength: 100_000,
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

const normalizeRichTextColor = (color: unknown): string | undefined => {
  if (typeof color !== "string" || color.length > 128 || /[\u0000-\u001f]/.test(color)) {
    return undefined;
  }

  const trimmed = color.trim();
  return /^(?:#[\da-f]{3,8}|[a-z]+|(?:rgb|hsl)a?\([\d.% ,/+\-]+\))$/i.test(trimmed) ? trimmed : undefined;
};

const normalizeRichTextHyperlink = (hyperlink: unknown): string | undefined => {
  if (typeof hyperlink !== "string" || !hyperlink || /[\s\u007f]/.test(hyperlink)) {
    return undefined;
  }

  try {
    const url = new URL(hyperlink);
    if (url.protocol === "https:") {
      return url.href;
    }
    if (url.protocol === "mailto:" && url.pathname.includes("@")) {
      return url.href;
    }
  } catch {
    return undefined;
  }

  return undefined;
};

const normalizeRichTextSegments = (segments: CellRichTextSegment[], maxLength: number): CellRichTextSegment[] => {
  if (!Array.isArray(segments)) {
    throw createCoreOperationError("CORE_CELL_RICH_TEXT_INVALID", "Cell rich text must be an array of segments.");
  }

  let totalLength = 0;
  return segments.map((segment) => {
    if (!segment || typeof segment.text !== "string") {
      throw createCoreOperationError("CORE_CELL_RICH_TEXT_INVALID", "Each rich text segment must contain text.");
    }
    totalLength += segment.text.length;
    if (totalLength > maxLength) {
      throw createCoreOperationError("CORE_CELL_RICH_TEXT_MAX_LENGTH", "Cell rich text exceeds the configured maximum length.", {
        maxLength
      });
    }

    const hyperlink = segment.hyperlink === undefined ? undefined : normalizeRichTextHyperlink(segment.hyperlink);
    if (segment.hyperlink !== undefined && hyperlink === undefined) {
      throw createCoreOperationError("CORE_CELL_RICH_TEXT_LINK_INVALID", "Rich text hyperlinks must use HTTPS or mailto.");
    }
    const color = segment.style?.color === undefined ? undefined : normalizeRichTextColor(segment.style.color);
    if (segment.style?.color !== undefined && color === undefined) {
      throw createCoreOperationError("CORE_CELL_RICH_TEXT_COLOR_INVALID", "Rich text color is not supported.");
    }

    const style = segment.style
      ? {
          ...(segment.style.bold === true ? { bold: true } : {}),
          ...(segment.style.italic === true ? { italic: true } : {}),
          ...(segment.style.underline === true ? { underline: true } : {}),
          ...(segment.style.strike === true ? { strike: true } : {}),
          ...(color ? { color } : {})
        }
      : undefined;
    return {
      text: segment.text,
      ...(style && Object.keys(style).length ? { style } : {}),
      ...(hyperlink ? { hyperlink } : {})
    };
  });
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

interface CollaborationClock {
  sequence: number;
  clientId: string;
}

const isNewerCollaborationClock = (next: CollaborationClock, current?: CollaborationClock): boolean =>
  !current || next.sequence > current.sequence || (next.sequence === current.sequence && next.clientId > current.clientId);

const getCollaborationOperationKey = (operation: SpreadsheetOperation): string =>
  `${operation.id}:${JSON.stringify(operation.path)}`;

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
    splitPane: sheetInput.splitPane ? cloneValue(sheetInput.splitPane) : undefined,
    columns: cloneValue(sheetInput.columns ?? {}),
    rows: cloneValue(sheetInput.rows ?? {}),
    rowCount: Math.min(sheetInput.rowCount ?? 200, settings.maxRows),
    columnCount: Math.min(sheetInput.columnCount ?? 26, settings.maxColumns),
    selection: createDefaultSelection(),
    images: cloneValue(sheetInput.images ?? []),
    widgets: cloneValue(sheetInput.widgets ?? []),
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
    if (sheet.splitPane && sheet.splitPane.horizontalRow === undefined && sheet.splitPane.verticalColumn === undefined) {
      delete sheet.splitPane;
    }
    sheet.columns ??= {};
    sheet.rows ??= {};
    sheet.images ??= [];
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

  private readonly collaborationAdapter?: CollaborationAdapter;

  private readonly collaborationClientId?: string;

  private readonly collaborationEnvelopeIds = new Set<string>();

  private readonly collaborationOperationClocks = new Map<string, CollaborationClock>();

  private readonly collaborationPresences = new Map<string, CollaborationPresence>();

  private readonly collaborationPresenceClocks = new Map<string, CollaborationClock>();

  private readonly collaborationPresenceTimers = new Map<string, ReturnType<typeof setTimeout>>();

  private readonly collaborationConflictPolicy: CollaborationConflictPolicy;

  private readonly collaborationPresenceTtlMs: number;

  private collaborationSequence = 0;

  private collaborationPresenceSequence = 0;

  private applyingRemoteOperations = false;

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
    this.collaborationAdapter = config.collaboration?.adapter;
    this.collaborationClientId = config.collaboration ? (config.collaboration.clientId ?? createId("client")) : undefined;
    this.collaborationConflictPolicy = config.collaboration?.conflictPolicy ?? "last-write-wins";
    this.collaborationPresenceTtlMs = Math.max(1, config.collaboration?.presenceTtlMs ?? 30_000);
    this.events.on("command:completed", ({ sheetId, commandType }) => {
      if (!sheetId || commandType === "SelectRangeCommand" || this.disposed) {
        return;
      }

      this.handleDerivedPivotSourceMutation(sheetId);
    });
    this.events.on("command:completed", ({ sheetId, operations }) => {
      if (!sheetId || !operations.some((operation) => operation.path[0] === "splitPane")) {
        return;
      }

      this.events.emit("split-pane:changed", {
        timestamp: Date.now(),
        workbookId: this.getSnapshot().id,
        sheetId,
        splitPane: this.getSplitPane(sheetId)
      });
    });
    this.events.on("command:completed", ({ sheetId, operations }) => {
      if (!this.collaborationAdapter || !this.collaborationClientId || this.applyingRemoteOperations || !operations.length) {
        return;
      }
      const envelope: CollaborationEnvelope = {
        id: `${this.collaborationClientId}:${++this.collaborationSequence}`,
        workbookId: workbook.id,
        clientId: this.collaborationClientId,
        sequence: this.collaborationSequence,
        timestamp: Date.now(),
        sheetId,
        operations: cloneValue(operations)
      };
      for (const operation of operations) {
        this.collaborationOperationClocks.set(getCollaborationOperationKey(operation), {
          sequence: envelope.sequence,
          clientId: envelope.clientId
        });
      }
      this.collaborationEnvelopeIds.add(envelope.id);
      void Promise.resolve(this.collaborationAdapter.send(envelope)).catch((error: unknown) => {
        this.emitCollaborationError("send", error);
      });
    });
    if (this.collaborationAdapter && this.collaborationClientId) {
      this.events.emit("collaboration:status", {
        timestamp: Date.now(),
        workbookId: workbook.id,
        clientId: this.collaborationClientId,
        status: "connecting"
      });
      void Promise.resolve(
        this.collaborationAdapter.connect({
          workbookId: workbook.id,
          clientId: this.collaborationClientId,
          receive: (envelope) => {
            this.applyCollaborationEnvelope(envelope);
          },
          receivePresence: (message) => {
            this.applyCollaborationPresenceMessage(message);
          },
          receiveError: (error) => {
            this.emitCollaborationError("receive", error);
          }
        })
      ).then(() => {
        if (!this.disposed) {
          this.events.emit("collaboration:status", {
            timestamp: Date.now(),
            workbookId: workbook.id,
            clientId: this.collaborationClientId!,
            status: "connected"
          });
          void Promise.resolve(this.collaborationAdapter?.getPresence?.(workbook.id) ?? []).then((presences) => {
            for (const presence of presences) {
              this.applyCollaborationPresenceMessage({
                type: "presence:update",
                workbookId: workbook.id,
                clientId: presence.clientId,
                sequence: presence.sequence,
                timestamp: presence.updatedAt,
                presence
              });
            }
          }).catch((error: unknown) => {
            this.emitCollaborationError("presence", error);
          });
        }
      }).catch((error: unknown) => {
        this.emitCollaborationError("connect", error);
      });
    }
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

  private emitCollaborationError(
    phase: "connect" | "send" | "receive" | "presence" | "disconnect",
    error: unknown
  ): void {
    if (!this.collaborationClientId) {
      return;
    }
    this.events.emit("collaboration:error", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      clientId: this.collaborationClientId,
      phase,
      message: error instanceof Error ? error.message : "Unknown collaboration error."
    });
    this.events.emit("collaboration:status", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      clientId: this.collaborationClientId,
      status: "error"
    });
  }

  applyCollaborationEnvelope(envelope: CollaborationEnvelope): boolean {
    if (this.disposed || this.collaborationEnvelopeIds.has(envelope.id)) {
      return false;
    }
    const workbook = this.getSnapshot();
    if (envelope.workbookId !== workbook.id || !envelope.operations.length) {
      this.emitCollaborationError("receive", new Error("Collaboration envelope does not match this workbook."));
      return false;
    }
    const anchorSheetId = envelope.sheetId && workbook.sheets.some((sheet) => sheet.id === envelope.sheetId)
      ? envelope.sheetId
      : workbook.activeSheetId;
    this.collaborationEnvelopeIds.add(envelope.id);
    const clock = { sequence: envelope.sequence, clientId: envelope.clientId };
    const operations = envelope.operations.filter((operation) =>
      isNewerCollaborationClock(clock, this.collaborationOperationClocks.get(getCollaborationOperationKey(operation)))
    );
    if (!operations.length) {
      return false;
    }
    this.applyingRemoteOperations = true;
    try {
      this.applyBatchOperations({ anchorSheetId, operations: cloneValue(operations) });
    } catch (error) {
      this.collaborationEnvelopeIds.delete(envelope.id);
      this.emitCollaborationError("receive", error);
      return false;
    } finally {
      this.applyingRemoteOperations = false;
    }
    for (const operation of operations) {
      this.collaborationOperationClocks.set(getCollaborationOperationKey(operation), clock);
    }
    this.events.emit("collaboration:operationsApplied", {
      timestamp: Date.now(),
      workbookId: workbook.id,
      clientId: envelope.clientId,
      envelopeId: envelope.id,
      operationCount: operations.length
    });
    return true;
  }

  getCollaborationConflictPolicy(): CollaborationConflictPolicy {
    return this.collaborationConflictPolicy;
  }

  private expireCollaborationPresences(now = Date.now()): void {
    for (const [clientId, presence] of this.collaborationPresences) {
      if (presence.expiresAt > now) {
        continue;
      }
      this.collaborationPresences.delete(clientId);
      const timer = this.collaborationPresenceTimers.get(clientId);
      if (timer) {
        clearTimeout(timer);
        this.collaborationPresenceTimers.delete(clientId);
      }
      this.events.emit("collaboration:presenceRemoved", {
        timestamp: now,
        workbookId: this.getWorkbookState().id,
        clientId,
        reason: "expired"
      });
    }
  }

  private applyCollaborationPresenceMessage(message: CollaborationPresenceMessage): boolean {
    if (message.workbookId !== this.getWorkbookState().id) {
      this.emitCollaborationError("presence", new Error("Collaboration presence does not match this workbook."));
      return false;
    }
    const clock = { sequence: message.sequence, clientId: message.clientId };
    if (!isNewerCollaborationClock(clock, this.collaborationPresenceClocks.get(message.clientId))) {
      return false;
    }
    this.collaborationPresenceClocks.set(message.clientId, clock);
    if (message.type === "presence:remove") {
      const removed = this.collaborationPresences.delete(message.clientId);
      const timer = this.collaborationPresenceTimers.get(message.clientId);
      if (timer) {
        clearTimeout(timer);
        this.collaborationPresenceTimers.delete(message.clientId);
      }
      if (removed) {
        this.events.emit("collaboration:presenceRemoved", {
          timestamp: Date.now(),
          workbookId: message.workbookId,
          clientId: message.clientId,
          reason: "removed"
        });
      }
      return removed;
    }
    const presence: CollaborationPresence = {
      ...cloneValue(message.presence),
      clientId: message.clientId,
      sequence: message.sequence
    };
    this.collaborationPresences.set(message.clientId, presence);
    const previousTimer = this.collaborationPresenceTimers.get(message.clientId);
    if (previousTimer) {
      clearTimeout(previousTimer);
    }
    const timer = setTimeout(() => {
      this.expireCollaborationPresences();
    }, Math.max(0, presence.expiresAt - Date.now()));
    this.collaborationPresenceTimers.set(message.clientId, timer);
    this.events.emit("collaboration:presenceChanged", {
      timestamp: Date.now(),
      workbookId: message.workbookId,
      presence: cloneValue(presence)
    });
    return true;
  }

  getPresence(clientId: string): CollaborationPresence | undefined {
    this.expireCollaborationPresences();
    const presence = this.collaborationPresences.get(clientId);
    return presence ? cloneValue(presence) : undefined;
  }

  getPresences(): CollaborationPresence[] {
    this.expireCollaborationPresences();
    return [...this.collaborationPresences.values()].map((presence) => cloneValue(presence));
  }

  updatePresence(input: Omit<CollaborationPresence, "clientId" | "sequence" | "updatedAt" | "expiresAt">): CollaborationPresence {
    if (!this.collaborationAdapter || !this.collaborationClientId) {
      throw createCoreOperationError("CORE_COLLABORATION_REQUIRED", "Collaboration must be configured to update presence.");
    }
    const now = Date.now();
    const presence: CollaborationPresence = {
      ...cloneValue(input),
      clientId: this.collaborationClientId,
      sequence: ++this.collaborationPresenceSequence,
      updatedAt: now,
      expiresAt: now + this.collaborationPresenceTtlMs
    };
    const message: Extract<CollaborationPresenceMessage, { type: "presence:update" }> = {
      type: "presence:update",
      workbookId: this.getWorkbookState().id,
      clientId: presence.clientId,
      sequence: presence.sequence,
      timestamp: now,
      presence
    };
    this.applyCollaborationPresenceMessage(message);
    void Promise.resolve(this.collaborationAdapter.updatePresence?.(message)).catch((error: unknown) => {
      this.emitCollaborationError("presence", error);
    });
    return cloneValue(presence);
  }

  removePresence(clientId = this.collaborationClientId): boolean {
    if (!this.collaborationAdapter || !this.collaborationClientId || !clientId) {
      return false;
    }
    const message: Extract<CollaborationPresenceMessage, { type: "presence:remove" }> = {
      type: "presence:remove",
      workbookId: this.getWorkbookState().id,
      clientId,
      sequence: ++this.collaborationPresenceSequence,
      timestamp: Date.now()
    };
    const removed = this.applyCollaborationPresenceMessage(message);
    void Promise.resolve(this.collaborationAdapter.removePresence?.(message)).catch((error: unknown) => {
      this.emitCollaborationError("presence", error);
    });
    return removed;
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

  getCellNote(sheetId: string, row: number, col: number): string | undefined {
    return this.getCell(sheetId, row, col)?.note;
  }

  getCellRichText(sheetId: string, row: number, col: number): CellRichTextSegment[] | undefined {
    return this.getCell(sheetId, row, col)?.richText;
  }

  getCellComments(sheetId: string, row: number, col: number): CellComment[] {
    return this.getCell(sheetId, row, col)?.comments ?? [];
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

  getSplitPane(sheetId: string): SheetSplitPane | undefined {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    return sheet.splitPane ? cloneValue(sheet.splitPane) : undefined;
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

  getClientSideQuery(sheetId: string): ClientSideQueryState | undefined {
    const state = this.getSheetState(sheetId)?.metadata?.[CLIENT_SIDE_QUERY_METADATA_KEY];
    if (!state || typeof state !== "object") {
      return undefined;
    }
    return cloneValue(state as ClientSideQueryState);
  }

  applyClientSideSortFilter(input: {
    sheetId: string;
    sort?: ClientSideSortDescriptor[];
    filters?: ClientSideFilterDescriptor[];
    hasHeader?: boolean;
  }): SpreadsheetOperation[] {
    const explicitRowModel = this.explicitRowModels.get(input.sheetId);
    if (explicitRowModel && !(explicitRowModel instanceof ClientSideRowModel)) {
      throw createCoreOperationError(
        "CORE_CLIENT_SIDE_ROW_MODEL_REQUIRED",
        `Sheet does not use local or client-side rows: ${input.sheetId}`,
        { sheetId: input.sheetId, rowModelKind: explicitRowModel.kind }
      );
    }

    const sheet = this.getSheetState(input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }

    const state: ClientSideQueryState = {
      sort: cloneValue(input.sort ?? []),
      filters: cloneValue(input.filters ?? []),
      hasHeader: input.hasHeader !== false
    };
    for (const descriptor of [...state.sort, ...state.filters]) {
      if (!Number.isInteger(descriptor.column) || descriptor.column < 0 || descriptor.column >= sheet.columnCount) {
        throw createRangeError("Client-side query column is outside the current sheet bounds.");
      }
    }

    const firstDataRow = state.hasHeader ? 1 : 0;
    const rowIndexes = Array.from({ length: Math.max(0, sheet.rowCount - firstDataRow) }, (_, index) => firstDataRow + index);
    const getValue = (row: number, column: number): CellPrimitive => {
      const cell = sheet.cells[getCellKey(row, column)];
      return cell?.computedValue ?? cell?.value ?? null;
    };
    rowIndexes.sort((leftRow, rightRow) => {
      for (const descriptor of state.sort) {
        const comparison = compareClientSideValues(getValue(leftRow, descriptor.column), getValue(rightRow, descriptor.column));
        if (comparison !== 0) return descriptor.direction === "asc" ? comparison : -comparison;
      }
      return leftRow - rightRow;
    });

    const nextCells: SheetModel["cells"] = {};
    const nextRows: SheetModel["rows"] = {};
    if (state.hasHeader) {
      for (let col = 0; col < sheet.columnCount; col += 1) {
        const header = sheet.cells[getCellKey(0, col)];
        if (header) nextCells[getCellKey(0, col)] = cloneValue(header);
      }
      if (sheet.rows[0]) nextRows[0] = cloneValue(sheet.rows[0]);
    }

    rowIndexes.forEach((sourceRow, offset) => {
      const targetRow = firstDataRow + offset;
      for (let col = 0; col < sheet.columnCount; col += 1) {
        const cell = sheet.cells[getCellKey(sourceRow, col)];
        if (cell) nextCells[getCellKey(targetRow, col)] = cloneValue(cell);
      }
      const matches = state.filters.every((filter) => matchesClientSideFilter(getValue(sourceRow, filter.column), filter));
      const schema = cloneValue(sheet.rows[sourceRow] ?? {});
      if (state.filters.length) schema.hidden = !matches;
      else delete schema.hidden;
      if (Object.keys(schema).length) nextRows[targetRow] = schema;
    });

    const metadata = cloneValue(sheet.metadata ?? {});
    metadata[CLIENT_SIDE_QUERY_METADATA_KEY] = cloneValue(state);
    const operations = this.applyBatchOperations({
      anchorSheetId: sheet.id,
      operations: [
        { op: "replace", id: sheet.id, path: ["cells"], value: nextCells },
        { op: "replace", id: sheet.id, path: ["rows"], value: nextRows },
        { op: "replace", id: sheet.id, path: ["metadata"], value: metadata }
      ],
      affectedRanges: [{ start: { row: firstDataRow, col: 0 }, end: { row: sheet.rowCount - 1, col: sheet.columnCount - 1 } }]
    });
    this.events.emit("client-side-query:applied", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId: sheet.id,
      state: cloneValue(state)
    });
    this.emitRowModelChanged(sheet.id);
    return operations;
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
      splitPane: nextSheetInput.splitPane ? cloneValue(nextSheetInput.splitPane) : undefined,
      columns: cloneValue(nextSheetInput.columns ?? {}),
      rows: cloneValue(nextSheetInput.rows ?? {}),
      rowCount: Math.min(nextSheetInput.rowCount ?? currentSheet.rowCount, workbook.settings.maxRows),
      columnCount: Math.min(nextSheetInput.columnCount ?? currentSheet.columnCount, workbook.settings.maxColumns),
      images: cloneValue(nextSheetInput.images ?? currentSheet.images ?? []),
      widgets: cloneValue(nextSheetInput.widgets ?? currentSheet.widgets ?? []),
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

    return this.executeSheetOperations(input.anchorSheetId, input.operations, input.affectedRanges ?? [sheet.selection]);
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

  setCellNote(input: { sheetId: string; row: number; col: number; note?: string }): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }
    if (input.row < 0 || input.col < 0 || input.row >= sheet.rowCount || input.col >= sheet.columnCount) {
      throw createRangeError("Cell address is outside the current sheet bounds.");
    }

    const note = input.note === undefined || input.note === "" ? undefined : input.note;
    if (note !== undefined && note.length > this.getWorkbookState().settings.maxCellLength) {
      throw createCoreOperationError("CORE_CELL_NOTE_MAX_LENGTH", "Cell note exceeds the configured maximum length.", {
        maxLength: this.getWorkbookState().settings.maxCellLength
      });
    }

    const key = getCellKey(input.row, input.col);
    const previousCell = sheet.cells[key];
    if (previousCell?.note === note || (!previousCell && note === undefined)) {
      return [];
    }

    const nextCell: CellModel = {
      ...previousCell,
      value: previousCell?.value ?? null,
      computedValue: previousCell?.computedValue ?? previousCell?.value ?? null,
      note
    };
    if (note === undefined) {
      delete nextCell.note;
    }

    const operations = this.executeSheetOperations(
      input.sheetId,
      [{ op: previousCell ? "replace" : "add", id: input.sheetId, path: ["cells", key], value: nextCell }],
      [{ start: { row: input.row, col: input.col }, end: { row: input.row, col: input.col } }]
    );
    this.events.emit("cell:noteChanged", {
      timestamp: Date.now(),
      workbookId: this.getWorkbookState().id,
      sheetId: input.sheetId,
      address: { row: input.row, col: input.col },
      note
    });
    return operations;
  }

  setCellRichText(input: {
    sheetId: string;
    row: number;
    col: number;
    richText?: CellRichTextSegment[];
  }): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }
    if (input.row < 0 || input.col < 0 || input.row >= sheet.rowCount || input.col >= sheet.columnCount) {
      throw createRangeError("Cell address is outside the current sheet bounds.");
    }

    const richText = input.richText?.length
      ? normalizeRichTextSegments(input.richText, this.getWorkbookState().settings.maxCellLength)
      : undefined;
    const key = getCellKey(input.row, input.col);
    const previousCell = sheet.cells[key];
    if (JSON.stringify(previousCell?.richText) === JSON.stringify(richText) || (!previousCell && richText === undefined)) {
      return [];
    }

    const nextCell: CellModel = {
      ...previousCell,
      value: previousCell?.value ?? null,
      computedValue: previousCell?.computedValue ?? previousCell?.value ?? null,
      richText
    };
    if (richText === undefined) {
      delete nextCell.richText;
    }

    const operations = this.executeSheetOperations(
      input.sheetId,
      [{ op: previousCell ? "replace" : "add", id: input.sheetId, path: ["cells", key], value: nextCell }],
      [{ start: { row: input.row, col: input.col }, end: { row: input.row, col: input.col } }]
    );
    this.events.emit("cell:richTextChanged", {
      timestamp: Date.now(),
      workbookId: this.getWorkbookState().id,
      sheetId: input.sheetId,
      address: { row: input.row, col: input.col },
      richText: richText ? cloneValue(richText) : undefined
    });
    return operations;
  }

  createCellComment(input: {
    sheetId: string;
    row: number;
    col: number;
    comment: { id?: string; author: CellComment["author"]; content: string };
  }): SpreadsheetOperation[] {
    const sheet = this.getSheetState(input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }
    const previousCell = this.getCellState(input.sheetId, input.row, input.col);
    const content = input.comment.content.trim();
    if (!content || content.length > this.getWorkbookState().settings.maxCellLength) {
      throw createCoreOperationError("CORE_CELL_COMMENT_INVALID", "Cell comment content is empty or exceeds the configured maximum length.");
    }
    const comment: CellComment = {
      id: input.comment.id?.trim() || createId("comment"),
      author: cloneValue(input.comment.author),
      content,
      createdAt: Date.now(),
      resolved: false,
      replies: []
    };
    const comments = [...(previousCell?.comments ?? []), comment];
    const operations = this.replaceCellComments(input.sheetId, input.row, input.col, previousCell, comments);
    this.events.emit("cell:commentCreated", {
      timestamp: Date.now(),
      workbookId: this.getWorkbookState().id,
      sheetId: input.sheetId,
      address: { row: input.row, col: input.col },
      comment: cloneValue(comment)
    });
    return operations;
  }

  replyToCellComment(input: {
    sheetId: string;
    row: number;
    col: number;
    commentId: string;
    reply: { id?: string; author: CellCommentReply["author"]; content: string };
  }): SpreadsheetOperation[] {
    const previousCell = this.getCellState(input.sheetId, input.row, input.col);
    const comments = cloneValue(previousCell?.comments ?? []);
    const comment = comments.find((item) => item.id === input.commentId);
    if (!comment) {
      throw createCoreOperationError("CORE_CELL_COMMENT_NOT_FOUND", `Cell comment not found: ${input.commentId}`);
    }
    const content = input.reply.content.trim();
    if (!content || content.length > this.getWorkbookState().settings.maxCellLength) {
      throw createCoreOperationError("CORE_CELL_COMMENT_REPLY_INVALID", "Cell comment reply is empty or exceeds the configured maximum length.");
    }
    const reply: CellCommentReply = {
      id: input.reply.id?.trim() || createId("reply"),
      author: cloneValue(input.reply.author),
      content,
      createdAt: Date.now()
    };
    comment.replies.push(reply);
    comment.updatedAt = reply.createdAt;
    const operations = this.replaceCellComments(input.sheetId, input.row, input.col, previousCell, comments);
    this.events.emit("cell:commentReplied", {
      timestamp: Date.now(),
      workbookId: this.getWorkbookState().id,
      sheetId: input.sheetId,
      address: { row: input.row, col: input.col },
      commentId: input.commentId,
      reply: cloneValue(reply)
    });
    return operations;
  }

  resolveCellComment(input: {
    sheetId: string;
    row: number;
    col: number;
    commentId: string;
    resolved?: boolean;
  }): SpreadsheetOperation[] {
    const previousCell = this.getCellState(input.sheetId, input.row, input.col);
    const comments = cloneValue(previousCell?.comments ?? []);
    const comment = comments.find((item) => item.id === input.commentId);
    if (!comment) {
      throw createCoreOperationError("CORE_CELL_COMMENT_NOT_FOUND", `Cell comment not found: ${input.commentId}`);
    }
    const resolved = input.resolved ?? true;
    if (comment.resolved === resolved) {
      return [];
    }
    comment.resolved = resolved;
    comment.updatedAt = Date.now();
    const operations = this.replaceCellComments(input.sheetId, input.row, input.col, previousCell, comments);
    this.events.emit("cell:commentResolved", {
      timestamp: Date.now(),
      workbookId: this.getWorkbookState().id,
      sheetId: input.sheetId,
      address: { row: input.row, col: input.col },
      commentId: input.commentId,
      resolved
    });
    return operations;
  }

  deleteCellComment(input: { sheetId: string; row: number; col: number; commentId: string }): SpreadsheetOperation[] {
    const previousCell = this.getCellState(input.sheetId, input.row, input.col);
    const comments = (previousCell?.comments ?? []).filter((comment) => comment.id !== input.commentId);
    if (comments.length === (previousCell?.comments ?? []).length) {
      throw createCoreOperationError("CORE_CELL_COMMENT_NOT_FOUND", `Cell comment not found: ${input.commentId}`);
    }
    const operations = this.replaceCellComments(input.sheetId, input.row, input.col, previousCell, comments);
    this.events.emit("cell:commentDeleted", {
      timestamp: Date.now(),
      workbookId: this.getWorkbookState().id,
      sheetId: input.sheetId,
      address: { row: input.row, col: input.col },
      commentId: input.commentId
    });
    return operations;
  }

  private replaceCellComments(
    sheetId: string,
    row: number,
    col: number,
    previousCell: CellModel | undefined,
    comments: CellComment[]
  ): SpreadsheetOperation[] {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }
    if (row < 0 || col < 0 || row >= sheet.rowCount || col >= sheet.columnCount) {
      throw createRangeError("Cell address is outside the current sheet bounds.");
    }
    const nextCell: CellModel = {
      ...previousCell,
      value: previousCell?.value ?? null,
      computedValue: previousCell?.computedValue ?? previousCell?.value ?? null,
      comments: cloneValue(comments)
    };
    if (!comments.length) {
      delete nextCell.comments;
    }
    return this.executeSheetOperations(
      sheetId,
      [{ op: previousCell ? "replace" : "add", id: sheetId, path: ["cells", getCellKey(row, col)], value: nextCell }],
      [{ start: { row, col }, end: { row, col } }]
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

  setSplitPane(sheetId: string, splitPane: SheetSplitPane): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }

    const horizontalRow = splitPane.horizontalRow;
    const verticalColumn = splitPane.verticalColumn;
    if (horizontalRow !== undefined && (!Number.isInteger(horizontalRow) || horizontalRow <= 0 || horizontalRow >= sheet.rowCount)) {
      throw new RangeError(`Horizontal split row must be between 1 and ${sheet.rowCount - 1}.`);
    }
    if (verticalColumn !== undefined && (!Number.isInteger(verticalColumn) || verticalColumn <= 0 || verticalColumn >= sheet.columnCount)) {
      throw new RangeError(`Vertical split column must be between 1 and ${sheet.columnCount - 1}.`);
    }
    if (horizontalRow === undefined && verticalColumn === undefined) {
      return this.clearSplitPane(sheetId);
    }

    const nextSplitPane: SheetSplitPane = { horizontalRow, verticalColumn };
    return this.executeSheetOperations(
      sheetId,
      [{
        op: sheet.splitPane ? "replace" : "add",
        id: sheetId,
        path: ["splitPane"],
        value: nextSplitPane
      }],
      [sheet.selection]
    );
  }

  clearSplitPane(sheetId: string): SpreadsheetOperation[] {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }
    if (!sheet.splitPane) {
      return [];
    }

    return this.executeSheetOperations(
      sheetId,
      [{ op: "remove", id: sheetId, path: ["splitPane"], value: undefined }],
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

  toJSON(): WorkbookModel {
    return this.getSnapshot();
  }

  loadFromJSON(snapshot: WorkbookModel): void {
    const normalized = normalizeWorkbookSnapshot(snapshot);
    this.commandBus.replaceWorkbook(recalculateWorkbookFormulas(normalized, this.formulaEngine));
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
  }

  getSelection(sheetId: string): CellRange {
    const sheet = this.getSnapshot().sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }
    return sheet.selection;
  }

  getImages(sheetId: string): WorksheetImageObject[] {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }
    return cloneValue(sheet.images ?? []);
  }

  getImage(sheetId: string, imageId: string): WorksheetImageObject | undefined {
    return this.getImages(sheetId).find((image) => image.id === imageId);
  }

  private assertImageSource(src: string): void {
    const settings = this.getSnapshot().settings;
    if (!src || src.length > (settings.maxImageSourceLength ?? DEFAULT_SETTINGS.maxImageSourceLength!)) {
      throw createCoreOperationError("CORE_IMAGE_SOURCE_INVALID", "Image source is empty or exceeds the configured limit.");
    }
    const isHttps = /^https:\/\/[^\s]+$/i.test(src);
    const isSafeData = /^data:image\/(?:png|jpeg|gif|webp);base64,[a-z0-9+/=]+$/i.test(src);
    if (!isHttps && !isSafeData) {
      throw createCoreOperationError("CORE_IMAGE_SOURCE_UNSAFE", "Image source must use HTTPS or a safe raster data URL.");
    }
  }

  createImage(input: { sheetId: string; image: WorksheetImageObjectInput }): SpreadsheetOperation[] {
    const sheet = this.getSheetState(input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }
    this.assertImageSource(input.image.src);
    const image: WorksheetImageObject = {
      id: input.image.id?.trim() || createId("image"),
      sheetId: input.sheetId,
      src: input.image.src,
      alt: String(input.image.alt ?? "").slice(0, this.getSnapshot().settings.maxCellLength),
      position: {
        fromCell: input.image.position.fromCell,
        toCell: input.image.position.toCell,
        offsetX: Number(input.image.position.offsetX) || 0,
        offsetY: Number(input.image.position.offsetY) || 0,
        width: Math.max(40, Number(input.image.position.width) || 320),
        height: Math.max(40, Number(input.image.position.height) || 220),
        zIndex: Math.max(0, Math.round(Number(input.image.position.zIndex ?? 1) || 1))
      },
      style: input.image.style ? cloneValue(input.image.style) : undefined,
      state: {
        selected: input.image.state?.selected ?? false,
        visible: input.image.state?.visible ?? true,
        locked: input.image.state?.locked ?? false
      }
    };
    const operations = this.executeSheetOperations(input.sheetId, [{
      op: "add",
      id: input.sheetId,
      path: ["images", (sheet.images ?? []).length],
      value: image
    }], [sheet.selection]);
    this.events.emit("image:created", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId: input.sheetId,
      imageId: image.id,
      image: cloneValue(image)
    });
    return operations;
  }

  updateImage(input: {
    sheetId: string;
    imageId: string;
    src?: string;
    alt?: string;
    position?: Partial<WorksheetObjectPosition>;
    style?: WorksheetImageObject["style"];
    state?: Partial<WorksheetImageObject["state"]>;
  }): SpreadsheetOperation[] {
    const sheet = this.getSheetState(input.sheetId);
    const imageIndex = sheet?.images?.findIndex((image) => image.id === input.imageId) ?? -1;
    if (!sheet || imageIndex < 0) {
      throw createCoreOperationError("CORE_IMAGE_NOT_FOUND", `Image not found: ${input.imageId}`, { sheetId: input.sheetId });
    }
    const current = sheet.images![imageIndex]!;
    if (input.src !== undefined) {
      this.assertImageSource(input.src);
    }
    const next: WorksheetImageObject = {
      ...cloneValue(current),
      src: input.src ?? current.src,
      alt: input.alt === undefined ? current.alt : String(input.alt).slice(0, this.getSnapshot().settings.maxCellLength),
      position: { ...current.position, ...(input.position ?? {}) },
      style: input.style === undefined ? current.style : cloneValue(input.style),
      state: { ...current.state, ...(input.state ?? {}) }
    };
    next.position.width = Math.max(40, Number(next.position.width) || 40);
    next.position.height = Math.max(40, Number(next.position.height) || 40);
    const operations = this.executeSheetOperations(input.sheetId, [{
      op: "replace",
      id: input.sheetId,
      path: ["images", imageIndex],
      value: next
    }], [sheet.selection]);
    this.events.emit("image:updated", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId: input.sheetId,
      imageId: input.imageId,
      image: cloneValue(next)
    });
    return operations;
  }

  deleteImage(sheetId: string, imageId: string): SpreadsheetOperation[] {
    const sheet = this.getSheetState(sheetId);
    const imageIndex = sheet?.images?.findIndex((image) => image.id === imageId) ?? -1;
    if (!sheet || imageIndex < 0) {
      throw createCoreOperationError("CORE_IMAGE_NOT_FOUND", `Image not found: ${imageId}`, { sheetId });
    }
    const operations = this.executeSheetOperations(sheetId, [{
      op: "remove",
      id: sheetId,
      path: ["images", imageIndex],
      value: cloneValue(sheet.images![imageIndex])
    }], [sheet.selection]);
    this.events.emit("image:deleted", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId,
      imageId
    });
    return operations;
  }

  moveImage(input: {
    sheetId: string;
    imageId: string;
    position: Pick<WorksheetObjectPosition, "fromCell" | "toCell" | "offsetX" | "offsetY" | "zIndex">;
  }): SpreadsheetOperation[] {
    const operations = this.updateImage(input);
    const image = this.getImage(input.sheetId, input.imageId)!;
    this.events.emit("image:moved", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId: input.sheetId,
      imageId: input.imageId,
      position: cloneValue(image.position)
    });
    return operations;
  }

  resizeImage(input: {
    sheetId: string;
    imageId: string;
    position: Pick<WorksheetObjectPosition, "width" | "height" | "toCell">;
  }): SpreadsheetOperation[] {
    const operations = this.updateImage(input);
    const image = this.getImage(input.sheetId, input.imageId)!;
    this.events.emit("image:resized", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id,
      sheetId: input.sheetId,
      imageId: input.imageId,
      position: cloneValue(image.position)
    });
    return operations;
  }

  selectImage(sheetId: string, imageId?: string): SpreadsheetOperation[] {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }
    if (imageId && !sheet.images?.some((image) => image.id === imageId)) {
      throw createCoreOperationError("CORE_IMAGE_NOT_FOUND", `Image not found: ${imageId}`, { sheetId });
    }
    const changed = (sheet.images ?? [])
      .map((image, index) => ({ image, index, selected: image.id === imageId }))
      .filter(({ image, selected }) => image.state.selected !== selected);
    if (changed.length === 0) {
      return [];
    }
    const operations = this.executeSheetOperations(sheetId, changed.map(({ image, index, selected }) => ({
      op: "replace" as const,
      id: sheetId,
      path: ["images", index],
      value: { ...cloneValue(image), state: { ...image.state, selected } }
    })), [sheet.selection]);
    for (const { image, selected } of changed) {
      this.events.emit(selected ? "image:selected" : "image:unselected", {
        timestamp: Date.now(),
        workbookId: this.getSnapshot().id,
        sheetId,
        imageId: image.id
      });
    }
    return operations;
  }

  getWidgets(sheetId: string): WorksheetWidgetObject[] {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }
    return cloneValue(sheet.widgets ?? []);
  }

  getWidget(sheetId: string, widgetId: string): WorksheetWidgetObject | undefined {
    return this.getWidgets(sheetId).find((widget) => widget.id === widgetId);
  }

  private assertWidgetJson(value: JsonValue | Record<string, JsonValue>, field: string): void {
    const seen = new Set<object>();
    const visit = (candidate: unknown): void => {
      if (candidate === null || typeof candidate === "string" || typeof candidate === "boolean") {
        return;
      }
      if (typeof candidate === "number") {
        if (Number.isFinite(candidate)) return;
        throw createCoreOperationError("CORE_WIDGET_DATA_INVALID", `${field} must contain finite JSON values.`);
      }
      if (!candidate || typeof candidate !== "object" || seen.has(candidate)) {
        throw createCoreOperationError("CORE_WIDGET_DATA_INVALID", `${field} must contain acyclic JSON values.`);
      }
      seen.add(candidate);
      const prototype = Object.getPrototypeOf(candidate);
      if (!Array.isArray(candidate) && prototype !== Object.prototype && prototype !== null) {
        throw createCoreOperationError("CORE_WIDGET_DATA_INVALID", `${field} must contain plain JSON objects.`);
      }
      for (const [key, child] of Object.entries(candidate)) {
        if (key === "__proto__" || key === "constructor" || key === "prototype") {
          throw createCoreOperationError("CORE_WIDGET_DATA_INVALID", `${field} contains a forbidden key.`);
        }
        visit(child);
      }
      seen.delete(candidate);
    };
    visit(value);
    const serialized = JSON.stringify(value);
    const limit = this.getSnapshot().settings.maxWidgetDataLength ?? DEFAULT_SETTINGS.maxWidgetDataLength!;
    if (serialized.length > limit) {
      throw createCoreOperationError("CORE_WIDGET_DATA_TOO_LARGE", `${field} exceeds the configured limit.`);
    }
  }

  createWidget(input: { sheetId: string; widget: WorksheetWidgetObjectInput }): SpreadsheetOperation[] {
    const sheet = this.getSheetState(input.sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(input.sheetId);
    }
    const type = input.widget.type.trim();
    if (!/^[a-z][a-z0-9._-]{0,127}$/i.test(type)) {
      throw createCoreOperationError("CORE_WIDGET_TYPE_INVALID", "Widget type must be a stable alphanumeric identifier.");
    }
    const config = input.widget.config ?? {};
    this.assertWidgetJson(config, "Widget config");
    if (input.widget.data !== undefined) this.assertWidgetJson(input.widget.data, "Widget data");
    const widget: WorksheetWidgetObject = {
      id: input.widget.id?.trim() || createId("widget"),
      sheetId: input.sheetId,
      type,
      label: String(input.widget.label ?? type).slice(0, this.getSnapshot().settings.maxCellLength),
      config: cloneValue(config),
      ...(input.widget.data === undefined ? {} : { data: cloneValue(input.widget.data) }),
      position: {
        fromCell: input.widget.position.fromCell,
        toCell: input.widget.position.toCell,
        offsetX: Number(input.widget.position.offsetX) || 0,
        offsetY: Number(input.widget.position.offsetY) || 0,
        width: Math.max(40, Number(input.widget.position.width) || 320),
        height: Math.max(40, Number(input.widget.position.height) || 220),
        zIndex: Math.max(0, Math.round(Number(input.widget.position.zIndex ?? 1) || 1))
      },
      state: {
        selected: input.widget.state?.selected ?? false,
        visible: input.widget.state?.visible ?? true,
        locked: input.widget.state?.locked ?? false
      }
    };
    const operations = this.executeSheetOperations(input.sheetId, [{
      op: "add",
      id: input.sheetId,
      path: ["widgets", (sheet.widgets ?? []).length],
      value: widget
    }], [sheet.selection]);
    this.events.emit("widget:created", {
      timestamp: Date.now(), workbookId: this.getSnapshot().id, sheetId: input.sheetId, widgetId: widget.id, widget: cloneValue(widget)
    });
    return operations;
  }

  updateWidget(input: {
    sheetId: string;
    widgetId: string;
    config?: Record<string, JsonValue>;
    data?: JsonValue;
    label?: string;
    position?: Partial<WorksheetObjectPosition>;
    state?: Partial<WorksheetWidgetObject["state"]>;
  }): SpreadsheetOperation[] {
    const sheet = this.getSheetState(input.sheetId);
    const widgetIndex = sheet?.widgets?.findIndex((widget) => widget.id === input.widgetId) ?? -1;
    if (!sheet || widgetIndex < 0) {
      throw createCoreOperationError("CORE_WIDGET_NOT_FOUND", `Widget not found: ${input.widgetId}`, { sheetId: input.sheetId });
    }
    if (input.config !== undefined) this.assertWidgetJson(input.config, "Widget config");
    if (input.data !== undefined) this.assertWidgetJson(input.data, "Widget data");
    const current = sheet.widgets![widgetIndex]!;
    const next: WorksheetWidgetObject = {
      ...cloneValue(current),
      label: input.label === undefined ? current.label : String(input.label).slice(0, this.getSnapshot().settings.maxCellLength),
      config: input.config === undefined ? current.config : cloneValue(input.config),
      ...(input.data === undefined ? {} : { data: cloneValue(input.data) }),
      position: { ...current.position, ...(input.position ?? {}) },
      state: { ...current.state, ...(input.state ?? {}) }
    };
    next.position.width = Math.max(40, Number(next.position.width) || 40);
    next.position.height = Math.max(40, Number(next.position.height) || 40);
    next.position.zIndex = Math.max(0, Math.round(Number(next.position.zIndex) || 0));
    const operations = this.executeSheetOperations(input.sheetId, [{
      op: "replace", id: input.sheetId, path: ["widgets", widgetIndex], value: next
    }], [sheet.selection]);
    this.events.emit("widget:updated", {
      timestamp: Date.now(), workbookId: this.getSnapshot().id, sheetId: input.sheetId, widgetId: input.widgetId, widget: cloneValue(next)
    });
    return operations;
  }

  moveWidget(input: {
    sheetId: string;
    widgetId: string;
    position: Pick<WorksheetObjectPosition, "fromCell" | "toCell" | "offsetX" | "offsetY" | "zIndex">;
  }): SpreadsheetOperation[] {
    const operations = this.updateWidget(input);
    const widget = this.getWidget(input.sheetId, input.widgetId)!;
    this.events.emit("widget:moved", {
      timestamp: Date.now(), workbookId: this.getSnapshot().id, sheetId: input.sheetId, widgetId: input.widgetId, position: cloneValue(widget.position)
    });
    return operations;
  }

  resizeWidget(input: {
    sheetId: string;
    widgetId: string;
    position: Pick<WorksheetObjectPosition, "width" | "height" | "toCell">;
  }): SpreadsheetOperation[] {
    const operations = this.updateWidget(input);
    const widget = this.getWidget(input.sheetId, input.widgetId)!;
    this.events.emit("widget:resized", {
      timestamp: Date.now(), workbookId: this.getSnapshot().id, sheetId: input.sheetId, widgetId: input.widgetId, position: cloneValue(widget.position)
    });
    return operations;
  }

  selectWidget(sheetId: string, widgetId?: string): SpreadsheetOperation[] {
    const sheet = this.getSheetState(sheetId);
    if (!sheet) {
      throw createSheetNotFoundError(sheetId);
    }
    if (widgetId && !sheet.widgets?.some((widget) => widget.id === widgetId)) {
      throw createCoreOperationError("CORE_WIDGET_NOT_FOUND", `Widget not found: ${widgetId}`, { sheetId });
    }
    const changed = (sheet.widgets ?? [])
      .map((widget, index) => ({ widget, index, selected: widget.id === widgetId }))
      .filter(({ widget, selected }) => widget.state.selected !== selected);
    if (changed.length === 0) {
      return [];
    }
    return this.executeSheetOperations(sheetId, changed.map(({ widget, index, selected }) => ({
      op: "replace" as const,
      id: sheetId,
      path: ["widgets", index],
      value: { ...cloneValue(widget), state: { ...widget.state, selected } }
    })), [sheet.selection]);
  }

  deleteWidget(sheetId: string, widgetId: string): SpreadsheetOperation[] {
    const sheet = this.getSheetState(sheetId);
    const widgetIndex = sheet?.widgets?.findIndex((widget) => widget.id === widgetId) ?? -1;
    if (!sheet || widgetIndex < 0) {
      throw createCoreOperationError("CORE_WIDGET_NOT_FOUND", `Widget not found: ${widgetId}`, { sheetId });
    }
    const operations = this.executeSheetOperations(sheetId, [{
      op: "remove", id: sheetId, path: ["widgets", widgetIndex], value: cloneValue(sheet.widgets![widgetIndex])
    }], [sheet.selection]);
    this.events.emit("widget:deleted", { timestamp: Date.now(), workbookId: this.getSnapshot().id, sheetId, widgetId });
    return operations;
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
    this.collaborationPresences.clear();
    this.collaborationPresenceClocks.clear();
    for (const timer of this.collaborationPresenceTimers.values()) {
      clearTimeout(timer);
    }
    this.collaborationPresenceTimers.clear();
    for (const rowModel of this.explicitRowModels.values()) {
      rowModel.dispose();
    }
    this.explicitRowModels.clear();
    this.clientSideRowModels.clear();
    this.pluginManager.clear();
    if (this.collaborationAdapter && this.collaborationClientId) {
      void Promise.resolve(this.collaborationAdapter.disconnect?.()).catch((error: unknown) => {
        this.emitCollaborationError("disconnect", error);
      });
      this.events.emit("collaboration:status", {
        timestamp: Date.now(),
        workbookId: this.getSnapshot().id,
        clientId: this.collaborationClientId,
        status: "disconnected"
      });
    }
    this.events.emit("engine:disposed", {
      timestamp: Date.now(),
      workbookId: this.getSnapshot().id
    });
    this.events.clear();
  }
}

export type { CellAddress, CellModel };