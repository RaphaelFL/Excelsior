import {
  CellValidationError,
  SpreadsheetOperationError,
  cellAddressToLabel,
  columnIndexToLabel,
  type CellAddress,
  type CellModel,
  type PivotAggregateFunction,
  type PivotBuildProgress,
  type CellPrimitive,
  type CellRange,
  type CellValidationConfig,
  type CellValidationRule,
  type CellStyle,
  type PivotExecutionMode,
  type RowModelRow,
  type SheetMerge,
  type RowResult,
  type SpreadsheetOperation,
  type WorkbookEngine
} from "@excelsior/core";
import { parseTabularText, resolveClipboardText } from "./clipboard";
import {
  collectFindReplaceEntries,
  collectFindReplaceMatches,
  createFindReplacePrepared,
  getSearchableCellText,
  type FindReplaceMatch,
  type FindReplaceScope
} from "./find-replace";

export interface CustomCellRenderContext {
  workbookId: string;
  sheetId: string;
  row: number;
  col: number;
  address: CellAddress;
  cell?: Readonly<CellModel>;
  value: CellPrimitive;
  displayValue: string;
  validation?: Readonly<CellValidationConfig>;
  selected: boolean;
  active: boolean;
}

export interface SafeCellRenderPart {
  text: string;
  tone?: "default" | "muted" | "accent";
}

export interface SafeCellRenderOutput {
  text?: string;
  parts?: SafeCellRenderPart[];
  accessoryText?: string;
  title?: string;
  ariaLabel?: string;
  classNames?: string[];
}

export interface CustomCellRenderer {
  id: string;
  matches(context: Readonly<CustomCellRenderContext>): boolean;
  render(context: Readonly<CustomCellRenderContext>): SafeCellRenderOutput;
}

export interface CustomCellEditorContext {
  workbookId: string;
  sheetId: string;
  row: number;
  col: number;
  address: CellAddress;
  cell?: Readonly<CellModel>;
  value: CellPrimitive;
  displayValue: string;
  validation?: Readonly<CellValidationConfig>;
}

export interface CustomCellEditorInstance {
  mount(container: HTMLElement): void;
  getValue(): string | number | boolean | null;
  focus?(): void;
  destroy?(): void;
}

export interface CustomCellEditor {
  id: string;
  matches(context: Readonly<CustomCellEditorContext>): boolean;
  create(context: Readonly<CustomCellEditorContext>): CustomCellEditorInstance;
}

export interface AutofillOptions {
  enabled?: boolean;
  maxCells?: number;
  copyStyle?: boolean;
}

export interface RendererMessages {
  gridLabel: string;
  sheetTabsLabel: string;
  formulaInputLabel: string;
  activeCellPrefix: string;
  rowPrefix: string;
  columnPrefix: string;
  sheetTabClose: string;
  blankCell: string;
  selectedCell: string;
  findPlaceholder: string;
  replacePlaceholder: string;
  findLabel: string;
  replaceLabel: string;
  scopeWorkbook: string;
  caseSensitive: string;
  wholeCell: string;
  regex: string;
  previousMatch: string;
  nextMatch: string;
  replaceOne: string;
  replaceAll: string;
  closePanel: string;
  searching: string;
  noResults: string;
  loadingRows: string;
  loadRowsFailed: string;
  sortAscending: string;
  sortDescending: string;
  groupColumn: string;
  pivotColumn: string;
  aggregateSum: string;
  aggregateAverage: string;
  aggregateMin: string;
  aggregateMax: string;
  aggregateCount: string;
  createPivot: string;
  pivotRow: string;
  pivotColumnField: string;
  pivotValue: string;
  pivotAggregate: string;
  pivotAlias: string;
  pivotAddValue: string;
  pivotRemoveValue: string;
  pivotRowTotals: string;
  pivotColumnTotals: string;
  pivotSubtotals: string;
  pivotCreating: string;
  pivotCancel: string;
  pivotCancelled: string;
  pivotApply: string;
  pivotUpdate: string;
  pivotExecutionMode: string;
  pivotExecutionAuto: string;
  pivotExecutionClient: string;
  pivotExecutionServer: string;
  pivotAutoRefresh: string;
  pivotStale: string;
  pivotRefreshing: string;
  pivotTooLarge: string;
  expandGroup: string;
  collapseGroup: string;
  clearColumnQuery: string;
  filterColumn: string;
  filterPlaceholder: string;
  undo: string;
  redo: string;
  bold: string;
  italic: string;
  wrap: string;
  alignLeft: string;
  alignCenter: string;
  alignRight: string;
  merge: string;
  unmerge: string;
  insertRow: string;
  deleteRow: string;
  insertColumn: string;
  deleteColumn: string;
  findReplace: string;
  addSheet: string;
  checkboxHint: string;
  dropdownHint: string;
  invalidCellValue: string;
  autofillHandle: string;
}

export interface RendererShortcutMap {
  openFindReplace: string[];
  findNext: string[];
  findPrevious: string[];
}

export interface RendererFormatters {
  number?: (value: number) => string;
  date?: (value: string) => string;
}

export interface RendererLocalizationOptions {
  locale?: string;
  direction?: "ltr" | "rtl";
  messages?: Partial<RendererMessages>;
  shortcuts?: Partial<RendererShortcutMap>;
  formatters?: RendererFormatters;
}

export interface DomSpreadsheetRendererOptions {
  onChange?: (operations: SpreadsheetOperation[]) => void;
  cellRenderers?: CustomCellRenderer[];
  cellEditors?: CustomCellEditor[];
  includeHiddenCellsInClipboard?: boolean;
  autofill?: AutofillOptions;
  localization?: RendererLocalizationOptions;
  renderDebounceMs?: number;
}

const DEFAULT_RENDERER_MESSAGES: RendererMessages = {
  gridLabel: "Planilha",
  sheetTabsLabel: "Abas da planilha",
  formulaInputLabel: "Barra de fórmulas",
  activeCellPrefix: "Célula ativa",
  rowPrefix: "Linha",
  columnPrefix: "Coluna",
  sheetTabClose: "Fechar aba",
  blankCell: "vazia",
  selectedCell: "selecionada",
  findPlaceholder: "Buscar",
  replacePlaceholder: "Substituir",
  findLabel: "Find",
  replaceLabel: "Replace",
  scopeWorkbook: "Workbook",
  caseSensitive: "Case",
  wholeCell: "Whole",
  regex: "Regex",
  previousMatch: "Prev",
  nextMatch: "Next",
  replaceOne: "Replace",
  replaceAll: "Replace All",
  closePanel: "Close",
  searching: "Buscando...",
  noResults: "0 resultados",
  loadingRows: "Carregando linhas...",
  loadRowsFailed: "Falha ao carregar linhas.",
  sortAscending: "Sort A-Z",
  sortDescending: "Sort Z-A",
  groupColumn: "Agrupar coluna",
  pivotColumn: "Pivot coluna",
  aggregateSum: "Somar coluna",
  aggregateAverage: "Media da coluna",
  aggregateMin: "Min da coluna",
  aggregateMax: "Max da coluna",
  aggregateCount: "Contar coluna",
  createPivot: "Pivot",
  pivotRow: "Linhas",
  pivotColumnField: "Colunas",
  pivotValue: "Valor",
  pivotAggregate: "Agregacao",
  pivotAlias: "Alias",
  pivotAddValue: "Adicionar medida",
  pivotRemoveValue: "Remover medida",
  pivotRowTotals: "Totais por linha",
  pivotColumnTotals: "Total geral",
  pivotSubtotals: "Subtotais",
  pivotCreating: "Criando pivot...",
  pivotCancel: "Cancelar",
  pivotCancelled: "Criação da pivot cancelada.",
  pivotApply: "Criar pivot",
  pivotUpdate: "Atualizar pivot",
  pivotExecutionMode: "Execução",
  pivotExecutionAuto: "Auto",
  pivotExecutionClient: "Client-side",
  pivotExecutionServer: "Server-side",
  pivotAutoRefresh: "Auto-refresh",
  pivotStale: "Pivot derivada desatualizada. Reabra o painel para atualizar ou reative o auto-refresh.",
  pivotRefreshing: "Atualizando pivot derivada...",
  pivotTooLarge: "Fonte grande demais para pivot local. Reduza o intervalo ou use pivot server-side.",
  expandGroup: "Expandir grupo",
  collapseGroup: "Recolher grupo",
  clearColumnQuery: "Limpar coluna",
  filterColumn: "Filtro",
  filterPlaceholder: "Filtrar coluna ativa",
  undo: "Undo",
  redo: "Redo",
  bold: "Bold",
  italic: "Italic",
  wrap: "Wrap",
  alignLeft: "Align Left",
  alignCenter: "Align Center",
  alignRight: "Align Right",
  merge: "Merge",
  unmerge: "Unmerge",
  insertRow: "+ Row",
  deleteRow: "- Row",
  insertColumn: "+ Col",
  deleteColumn: "- Col",
  findReplace: "Find",
  addSheet: "+ Sheet",
  checkboxHint: "Clique duplo para alternar o checkbox.",
  dropdownHint: "Clique duplo para escolher um valor da lista.",
  invalidCellValue: "Valor inválido para a célula selecionada.",
  autofillHandle: "Arrastar para preencher"
};

const DEFAULT_RENDERER_SHORTCUTS: RendererShortcutMap = {
  openFindReplace: ["Ctrl+F", "Meta+F"],
  findNext: ["F3"],
  findPrevious: ["Shift+F3"]
};

const isWithinRange = (row: number, col: number, range: CellRange): boolean =>
  row >= range.start.row &&
  row <= range.end.row &&
  col >= range.start.col &&
  col <= range.end.col;

const rangesOverlap = (left: CellRange, right: CellRange): boolean =>
  left.start.row <= right.end.row &&
  left.end.row >= right.start.row &&
  left.start.col <= right.end.col &&
  left.end.col >= right.start.col;

const getCellKey = (row: number, col: number): string => `${row}:${col}`;

const buildOffsets = (count: number, getSize: (index: number) => number): number[] => {
  const offsets = [0];
  for (let index = 0; index < count; index += 1) {
    offsets.push(offsets[index] + getSize(index));
  }
  return offsets;
};

const findVisibleBounds = (
  offsets: number[],
  viewportStart: number,
  viewportEnd: number,
  buffer: number
): { start: number; end: number } => {
  const maxIndex = Math.max(0, offsets.length - 2);
  let start = 0;
  while (start < maxIndex && offsets[start + 1] < viewportStart) {
    start += 1;
  }

  let end = start;
  while (end < maxIndex && offsets[end] <= viewportEnd) {
    end += 1;
  }

  return {
    start: Math.max(0, start - buffer),
    end: Math.min(maxIndex, end + buffer)
  };
};

const getSpanSize = (offsets: number[], start: number, end: number): number => offsets[end + 1] - offsets[start];

const normalizeStyle = (style?: CellStyle): CellStyle | undefined => {
  if (!style) {
    return undefined;
  }

  const nextStyle: CellStyle = {
    ...style,
    border: {
      ...style.border
    }
  };

  if (!nextStyle.border?.top && !nextStyle.border?.right && !nextStyle.border?.bottom && !nextStyle.border?.left) {
    delete nextStyle.border;
  }

  return Object.values(nextStyle).some((value) => value !== undefined) ? nextStyle : undefined;
};

type CellValidationListRule = Extract<CellValidationRule, { type: "list" | "dropdown" }>;
type CellValidationCheckboxRule = Extract<CellValidationRule, { type: "checkbox" }>;

type InteractiveValidationRule = CellValidationListRule | CellValidationCheckboxRule;

const isListValidationRule = (rule: CellValidationRule): rule is CellValidationListRule =>
  rule.type === "list" || rule.type === "dropdown";

const isCheckboxValidationRule = (rule: CellValidationRule): rule is CellValidationCheckboxRule => rule.type === "checkbox";

const toBoolean = (value: CellPrimitive): boolean => {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    return ["true", "1", "yes", "y"].includes(value.trim().toLowerCase());
  }

  if (typeof value === "number") {
    return value !== 0;
  }

  return false;
};

const sanitizeClassNames = (classNames: string[] | undefined): string[] =>
  (classNames ?? []).filter((className) => /^[a-zA-Z0-9_-]+$/.test(className));

const freezeRenderContext = (context: CustomCellRenderContext): Readonly<CustomCellRenderContext> =>
  Object.freeze({
    ...context,
    address: Object.freeze({ ...context.address }),
    cell: context.cell ? Object.freeze({ ...context.cell }) : undefined,
    validation: context.validation ? Object.freeze({ ...context.validation, rules: [...context.validation.rules] }) : undefined
  });

const freezeEditorContext = (context: CustomCellEditorContext): Readonly<CustomCellEditorContext> =>
  Object.freeze({
    ...context,
    address: Object.freeze({ ...context.address }),
    cell: context.cell ? Object.freeze({ ...context.cell }) : undefined,
    validation: context.validation ? Object.freeze({ ...context.validation, rules: [...context.validation.rules] }) : undefined
  });

const appendRangeIndices = (target: number[], start: number, end: number): void => {
  for (let index = start; index <= end; index += 1) {
    if (!target.includes(index)) {
      target.push(index);
    }
  }
};

const getElementFromEventTarget = (target: EventTarget | null): HTMLElement | undefined =>
  target instanceof HTMLElement ? target : undefined;

const cloneCellStyle = (style?: CellStyle): CellStyle | undefined => {
  if (!style) {
    return undefined;
  }

  return {
    ...style,
    border: style.border
      ? {
          ...style.border,
          top: style.border.top ? { ...style.border.top } : undefined,
          right: style.border.right ? { ...style.border.right } : undefined,
          bottom: style.border.bottom ? { ...style.border.bottom } : undefined,
          left: style.border.left ? { ...style.border.left } : undefined
        }
      : undefined
  };
};

const getRangeCellCount = (range: CellRange): number =>
  (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);

const columnLabelToIndex = (label: string): number => {
  let value = 0;
  for (const character of label) {
    value = value * 26 + ((character.codePointAt(0) ?? 64) - 64);
  }
  return value - 1;
};

const adjustFormulaReferences = (formula: string, rowOffset: number, colOffset: number): string =>
  formula.replace(/((?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!)?(\$?)([A-Z]+)(\$?)(\d+)/g, (_match, sheetPrefix, colLock, colLabel, rowLock, rowLabel) => {
    const resolvedSheetPrefix = sheetPrefix ?? "";
    const nextCol = colLock ? columnLabelToIndex(colLabel) : Math.max(0, columnLabelToIndex(colLabel) + colOffset);
    const nextRow = rowLock ? Number(rowLabel) - 1 : Math.max(0, Number(rowLabel) - 1 + rowOffset);
    return `${resolvedSheetPrefix}${colLock}${columnIndexToLabel(nextCol)}${rowLock}${nextRow + 1}`;
  });

const toNumericValue = (value: CellPrimitive): number | undefined => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return undefined;
    }
    const numeric = Number(trimmed);
    return Number.isFinite(numeric) ? numeric : undefined;
  }

  return undefined;
};

const toDateValue = (value: CellPrimitive): number | undefined => {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value.trim())) {
    return undefined;
  }

  const timestamp = Date.parse(`${value.trim()}T00:00:00.000Z`);
  return Number.isFinite(timestamp) ? timestamp : undefined;
};

const formatDateValue = (timestamp: number): string => new Date(timestamp).toISOString().slice(0, 10);

const matchesShortcut = (event: KeyboardEvent, shortcut: string): boolean => {
  const parts = shortcut
    .split("+")
    .map((part) => part.trim().toLowerCase())
    .filter(Boolean);
  if (!parts.length) {
    return false;
  }

  const key = parts.at(-1);
  if (!key) {
    return false;
  }

  return (
    event.ctrlKey === parts.includes("ctrl") &&
    event.metaKey === parts.includes("meta") &&
    event.altKey === parts.includes("alt") &&
    event.shiftKey === parts.includes("shift") &&
    event.key.toLowerCase() === key
  );
};

const toDomIdSegment = (value: string): string => value.replace(/[^a-zA-Z0-9_-]+/g, "-");

type AutofillAxis = "row" | "column";

interface AutofillPreview {
  axis: AutofillAxis;
  fillRange: CellRange;
}

interface FindReplaceState {
  open: boolean;
  query: string;
  replaceText: string;
  scope: FindReplaceScope;
  caseSensitive: boolean;
  wholeCell: boolean;
  regex: boolean;
  matches: FindReplaceMatch[];
  activeIndex: number;
  pending: boolean;
  error?: string;
}

interface PivotPanelState {
  open: boolean;
  pending: boolean;
  progress: number;
  sourceSheetId?: string;
  targetSheetId?: string;
  sourceRange?: CellRange;
  fields: string[];
  rowFields: string[];
  columnFields: string[];
  values: PivotPanelValueState[];
  includeRowTotals: boolean;
  includeColumnTotals: boolean;
  includeSubtotals: boolean;
  executionMode: PivotExecutionMode;
  autoRefresh: boolean;
}

interface PivotPanelValueState {
  field: string;
  aggregate: PivotAggregateFunction;
  as: string;
}

interface RemoteFilterDraft {
  sheetId: string;
  field: string;
  value: string;
}

type ScheduledFrameHandle = ReturnType<typeof globalThis.setTimeout>;

const scheduleFrame = (callback: () => void): ScheduledFrameHandle => {
  return globalThis.setTimeout(callback, 16);
};

const cancelScheduledFrame = (handle: ScheduledFrameHandle | undefined): void => {
  if (handle == null) {
    return;
  }

  globalThis.clearTimeout(handle);
};

const createFindReplaceActionButton = (action: string, label: string): HTMLButtonElement => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "excelsior-find-replace-button";
  button.dataset.findAction = action;
  button.textContent = label;
  return button;
};

const createFindReplaceToggle = (role: string, label: string): HTMLLabelElement => {
  const wrapper = document.createElement("label");
  wrapper.className = "excelsior-find-replace-toggle";
  const input = document.createElement("input");
  input.type = "checkbox";
  input.dataset.findRole = role;
  const text = document.createElement("span");
  text.textContent = label;
  wrapper.append(input, text);
  return wrapper;
};

const createPivotAggregateLabel = (messages: RendererMessages, aggregate: PivotAggregateFunction): string => {
  switch (aggregate) {
    case "sum":
      return messages.aggregateSum;
    case "avg":
      return messages.aggregateAverage;
    case "min":
      return messages.aggregateMin;
    case "max":
      return messages.aggregateMax;
    case "count":
      return messages.aggregateCount;
    default:
      return aggregate;
  }
};

export class DomSpreadsheetRenderer {
  private readonly root = document.createElement("div");

  private readonly chrome = document.createElement("div");

  private readonly toolbar = document.createElement("div");

  private readonly formulaBar = document.createElement("div");

  private readonly formulaAddress = document.createElement("span");

  private readonly formulaInput = document.createElement("input");

  private readonly statusMessage = document.createElement("span");

  private readonly findReplacePanel = document.createElement("div");

  private readonly findReplaceQueryInput = document.createElement("input");

  private readonly findReplaceValueInput = document.createElement("input");

  private readonly findReplaceResults = document.createElement("span");

  private readonly pivotPanel = document.createElement("div");

  private readonly pivotApplyButton = createFindReplaceActionButton("apply", "");

  private readonly pivotCloseButton = createFindReplaceActionButton("close", "");

  private readonly pivotRowSelect = document.createElement("select");

  private readonly pivotColumnSelect = document.createElement("select");

  private readonly pivotExecutionModeSelect = document.createElement("select");

  private readonly pivotValuesField = document.createElement("div");

  private readonly pivotValuesContainer = document.createElement("div");

  private readonly pivotIncludeRowTotals = document.createElement("input");

  private readonly pivotIncludeColumnTotals = document.createElement("input");

  private readonly pivotIncludeSubtotals = document.createElement("input");

  private readonly pivotAutoRefresh = document.createElement("input");

  private readonly activeCellAnnouncement = document.createElement("div");

  private readonly viewport = document.createElement("div");

  private readonly surface = document.createElement("div");

  private readonly cellsLayer = document.createElement("div");

  private readonly editor = document.createElement("input");

  private readonly selectEditor = document.createElement("select");

  private readonly customEditorHost = document.createElement("div");

  private readonly sheetTabs = document.createElement("div");

  private readonly messages: RendererMessages;

  private readonly shortcuts: RendererShortcutMap;

  private readonly direction: "ltr" | "rtl";

  private readonly unsubscribeCallbacks: Array<() => void> = [];

  private readonly composingTargets = new WeakSet<EventTarget>();

  private resizeObserver?: ResizeObserver;

  private editingCell?: { row: number; col: number; mode: "text" | "select" | "custom" };

  private autofillDrag?: {
    sourceRange: CellRange;
    preview?: AutofillPreview;
  };

  private activeCustomEditor?: CustomCellEditorInstance;

  private validationFeedback?: { sheetId: string; row: number; col: number; message: string };

  private rowModelFeedback?: { sheetId: string; error?: string };

  private pivotFeedback?: { sheetId: string; message: string; isError: boolean };

  private pivotBuildController?: AbortController;

  private remoteFilterDraft?: RemoteFilterDraft;

  private readonly findReplaceState: FindReplaceState = {
    open: false,
    query: "",
    replaceText: "",
    scope: "sheet",
    caseSensitive: false,
    wholeCell: false,
    regex: false,
    matches: [],
    activeIndex: -1,
    pending: false
  };

  private readonly pivotPanelState: PivotPanelState = {
    open: false,
    pending: false,
    progress: 0,
    fields: [],
    rowFields: [],
    columnFields: [],
    values: [],
    includeRowTotals: true,
    includeColumnTotals: true,
    includeSubtotals: true,
    executionMode: "auto",
    autoRefresh: true
  };

  private findReplaceSearchVersion = 0;

  private findReplaceSearchHandle?: ScheduledFrameHandle;

  private renderHandle?: ScheduledFrameHandle;

  private readonly rowModelWindowCache = new Map<string, RowModelRow[]>();

  private readonly pendingRowModelRequests = new Set<string>();

  private readonly failedRowModelRequests = new Set<string>();

  private getLocale(): string | undefined {
    return this.options.localization?.locale;
  }

  private getNumberFormatter(): ((value: number) => string) | undefined {
    return this.options.localization?.formatters?.number;
  }

  private getDateFormatter(): ((value: string) => string) | undefined {
    return this.options.localization?.formatters?.date;
  }

  private getRenderDebounceMs(): number {
    return Math.max(0, this.options.renderDebounceMs ?? 0);
  }

  private isRtl(): boolean {
    return this.direction === "rtl";
  }

  private getCellElementId(sheetId: string, row: number, col: number): string {
    return `excelsior-cell-${toDomIdSegment(sheetId)}-${row}-${col}`;
  }

  private getSheetTabElementId(sheetId: string): string {
    return `excelsior-sheet-tab-${toDomIdSegment(sheetId)}`;
  }

  private getColumnHeaderElementId(sheetId: string, col: number): string {
    return `excelsior-column-header-${toDomIdSegment(sheetId)}-${col}`;
  }

  private getColumnHeaderAccessibilityLabel(col: number): string {
    return `${this.messages.columnPrefix} ${columnIndexToLabel(col)}`;
  }

  private getCellAccessibilityLabel(sheetId: string, row: number, col: number, selected: boolean): string {
    const sheet = this.engine.getSnapshot().sheets.find((item) => item.id === sheetId);
    const displayValue = this.getRenderedCellDisplayValue(sheetId, row, col);
    const parts = [
      sheet?.name,
      `${this.messages.rowPrefix} ${row + 1}`,
      `${this.messages.columnPrefix} ${columnIndexToLabel(col)}`,
      displayValue || this.messages.blankCell
    ].filter(Boolean);

    if (selected) {
      parts.push(this.messages.selectedCell);
    }

    return parts.join(", ");
  }

  private updateAccessibilityState(sheetId: string, row: number, col: number, selected: boolean): void {
    this.viewport.setAttribute("aria-activedescendant", this.getCellElementId(sheetId, row, col));
    this.activeCellAnnouncement.textContent = `${this.messages.activeCellPrefix}: ${this.getCellAccessibilityLabel(
      sheetId,
      row,
      col,
      selected
    )}`;
  }

  private matchesShortcutList(event: KeyboardEvent, shortcuts: string[]): boolean {
    return shortcuts.some((shortcut) => matchesShortcut(event, shortcut));
  }

  private isCompositionInProgress(target: EventTarget | null, event?: KeyboardEvent): boolean {
    return event?.isComposing === true || (target != null && this.composingTargets.has(target));
  }

  private formatDisplayValue(value: CellPrimitive, rawDisplayValue: string): string {
    const dateFormatter = this.getDateFormatter();
    if (dateFormatter && typeof value === "string") {
      const timestamp = toDateValue(value);
      if (timestamp !== undefined) {
        return dateFormatter(value);
      }
    }

    const locale = this.getLocale();
    if (typeof value === "string") {
      const timestamp = toDateValue(value);
      if (timestamp !== undefined && locale) {
        return new Intl.DateTimeFormat(locale, {
          year: "numeric",
          month: "2-digit",
          day: "2-digit",
          timeZone: "UTC"
        }).format(new Date(timestamp));
      }
    }

    const numeric = toNumericValue(value);
    if (numeric !== undefined) {
      const numberFormatter = this.getNumberFormatter();
      if (numberFormatter) {
        return numberFormatter(numeric);
      }
      if (locale) {
        return new Intl.NumberFormat(locale).format(numeric);
      }
    }

    return rawDisplayValue;
  }

  private getRenderedCellDisplayValue(sheetId: string, row: number, col: number): string {
    return this.formatDisplayValue(
      this.getCellPrimitiveValue(sheetId, row, col),
      this.engine.getDisplayValue(sheetId, row, col)
    );
  }

  private getHorizontalNavigationDirection(key: "ArrowLeft" | "ArrowRight"): -1 | 1 {
    if (!this.isRtl()) {
      return key === "ArrowLeft" ? -1 : 1;
    }

    return key === "ArrowLeft" ? 1 : -1;
  }

  private isAutofillEnabled(): boolean {
    return this.options.autofill?.enabled !== false;
  }

  private getAutofillMaxCells(): number {
    return Math.max(1, this.options.autofill?.maxCells ?? 4000);
  }

  private shouldCopyStyleOnAutofill(): boolean {
    return this.options.autofill?.copyStyle === true;
  }

  private getColumnWidth(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>, col: number): number {
    return sheet.columns[col]?.hidden ? 0 : sheet.columns[col]?.width ?? this.engine.getSnapshot().settings.columnWidth;
  }

  private getResolvedRowCount(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>): number {
    const rowCount = this.engine.getRowModel(sheet.id).getRowCount();
    return rowCount === "unknown" ? sheet.rowCount : rowCount;
  }

  private getRemoteRequestField(col: number): string {
    return columnIndexToLabel(col);
  }

  private getRemoteRequestModel(sheetId: string): ReturnType<WorkbookEngine["getRemoteRowModelRequest"]> {
    return this.engine.getRemoteRowModelRequest(sheetId);
  }

  private getActiveRemoteSortDirection(sheetId: string, col: number): "asc" | "desc" | undefined {
    const field = this.getRemoteRequestField(col);
    return this.getRemoteRequestModel(sheetId)?.sortModel?.find((item) => item.field === field)?.direction;
  }

  private isActiveRemoteGrouped(sheetId: string, col: number): boolean {
    const field = this.getRemoteRequestField(col);
    return this.getRemoteRequestModel(sheetId)?.groupKeys?.includes(field) === true;
  }

  private isActiveRemotePivoted(sheetId: string, col: number): boolean {
    const field = this.getRemoteRequestField(col);
    return this.getRemoteRequestModel(sheetId)?.pivotModel?.some((item) => item.field === field) === true;
  }

  private hasActiveRemoteAggregate(sheetId: string, col: number, aggregateFunction: string): boolean {
    const field = this.getRemoteRequestField(col);
    return (
      this.getRemoteRequestModel(sheetId)?.aggregateModel?.some(
        (item) => item.field === field && item.function === aggregateFunction
      ) === true
    );
  }

  private getActiveRemoteFilterValue(sheetId: string, col: number): string {
    const field = this.getRemoteRequestField(col);
    if (this.remoteFilterDraft?.sheetId === sheetId && this.remoteFilterDraft.field === field) {
      return this.remoteFilterDraft.value;
    }

    const value = this.getRemoteRequestModel(sheetId)?.filterModel?.[field]?.value;
    return value == null ? "" : String(value);
  }

  private setRemoteFilterDraft(sheetId: string, col: number, value: string): void {
    this.remoteFilterDraft = {
      sheetId,
      field: this.getRemoteRequestField(col),
      value
    };
  }

  private applyRemoteSortForActiveColumn(direction: "asc" | "desc"): void {
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    if (this.getRemoteRequestModel(sheet.id) === undefined) {
      return;
    }

    this.engine.updateRemoteRowModel(sheet.id, {
      sortModel: [{ field: this.getRemoteRequestField(activeAddress.col), direction }]
    });
  }

  private cycleRemoteSortForColumn(sheetId: string, col: number): void {
    const current = this.getActiveRemoteSortDirection(sheetId, col);
    const field = this.getRemoteRequestField(col);
    if (current === undefined) {
      this.engine.updateRemoteRowModel(sheetId, {
        sortModel: [{ field, direction: "asc" }]
      });
      return;
    }

    if (current === "asc") {
      this.engine.updateRemoteRowModel(sheetId, {
        sortModel: [{ field, direction: "desc" }]
      });
      return;
    }

    this.engine.updateRemoteRowModel(sheetId, {
      sortModel: undefined
    });
  }

  private applyRemoteFilterForActiveColumn(rawValue: string): void {
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    const current = this.getRemoteRequestModel(sheet.id);
    if (current === undefined) {
      return;
    }

    const field = this.getRemoteRequestField(activeAddress.col);
    const nextFilterModel = {
      ...(current.filterModel ?? {})
    };
    const value = rawValue.trim();
    if (value) {
      nextFilterModel[field] = {
        operator: "equals",
        value
      };
    } else {
      delete nextFilterModel[field];
    }

    this.remoteFilterDraft = undefined;
    this.engine.updateRemoteRowModel(sheet.id, {
      filterModel: Object.keys(nextFilterModel).length ? nextFilterModel : undefined
    });
  }

  private clearRemoteQueryForActiveColumn(): void {
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    const current = this.getRemoteRequestModel(sheet.id);
    if (current === undefined) {
      return;
    }

    const field = this.getRemoteRequestField(activeAddress.col);
    const nextFilterModel = {
      ...(current.filterModel ?? {})
    };
    delete nextFilterModel[field];
    const nextSortModel = current.sortModel?.filter((item) => item.field !== field);

    this.remoteFilterDraft = undefined;
    this.engine.updateRemoteRowModel(sheet.id, {
      sortModel: nextSortModel?.length ? nextSortModel : undefined,
      filterModel: Object.keys(nextFilterModel).length ? nextFilterModel : undefined
    });
  }

  private toggleRemoteGroupingForActiveColumn(): void {
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    const current = this.getRemoteRequestModel(sheet.id);
    if (current === undefined) {
      return;
    }

    const field = this.getRemoteRequestField(activeAddress.col);
    const nextGroupKeys = current.groupKeys?.includes(field)
      ? current.groupKeys.filter((item) => item !== field)
      : [...(current.groupKeys ?? []), field];

    this.engine.updateRemoteRowModel(sheet.id, {
      groupKeys: nextGroupKeys.length ? nextGroupKeys : undefined
    });
  }

  private toggleRemotePivotForActiveColumn(): void {
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    const current = this.getRemoteRequestModel(sheet.id);
    if (current === undefined) {
      return;
    }

    const field = this.getRemoteRequestField(activeAddress.col);
    const currentPivotModel = current.pivotModel ?? [];
    const nextPivotModel = currentPivotModel.some((item) => item.field === field)
      ? currentPivotModel.filter((item) => item.field !== field)
      : [...currentPivotModel, { field }];

    this.engine.updateRemoteRowModel(sheet.id, {
      pivotModel: nextPivotModel.length ? nextPivotModel : undefined
    });
  }

  private toggleRemoteAggregateForActiveColumn(aggregateFunction: string): void {
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    const current = this.getRemoteRequestModel(sheet.id);
    if (current === undefined) {
      return;
    }

    const field = this.getRemoteRequestField(activeAddress.col);
    const currentAggregateModel = current.aggregateModel ?? [];
    const hasAggregate = currentAggregateModel.some((item) => item.field === field && item.function === aggregateFunction);
    const nextAggregateModel = hasAggregate
      ? currentAggregateModel.filter((item) => !(item.field === field && item.function === aggregateFunction))
      : [...currentAggregateModel, { field, function: aggregateFunction, as: `${field}_${aggregateFunction}` }];

    this.engine.updateRemoteRowModel(sheet.id, {
      aggregateModel: nextAggregateModel.length ? nextAggregateModel : undefined
    });
  }

  private clearRowModelWindowCache(): void {
    this.rowModelWindowCache.clear();
    this.pendingRowModelRequests.clear();
    this.failedRowModelRequests.clear();
    this.rowModelFeedback = undefined;
  }

  private hasPendingRowModelRequests(sheetId: string): boolean {
    const requestPrefix = `${sheetId}:`;
    for (const requestKey of this.pendingRowModelRequests) {
      if (requestKey.startsWith(requestPrefix)) {
        return true;
      }
    }

    return false;
  }

  private clearRowModelFeedback(sheetId: string): void {
    if (this.rowModelFeedback?.sheetId === sheetId) {
      this.rowModelFeedback = undefined;
    }
  }

  private setRowModelError(sheetId: string, error: unknown): void {
    this.rowModelFeedback = {
      sheetId,
      error: error instanceof Error && error.message ? error.message : this.messages.loadRowsFailed
    };
  }

  private resolveRowModelRows(sheetId: string, startRow: number, endRow: number): RowModelRow[] | undefined {
    const requestKey = `${sheetId}:${startRow}:${endRow}`;
    const cached = this.rowModelWindowCache.get(requestKey);
    if (cached && !this.pendingRowModelRequests.has(requestKey)) {
      return cached;
    }
    if (this.failedRowModelRequests.has(requestKey) && !this.pendingRowModelRequests.has(requestKey)) {
      return cached;
    }
    const rowModel = this.engine.getRowModel(sheetId);
    const applyResult = (result: RowResult): RowModelRow[] => {
      const rows = result.rows.map((row) => ({
        ...row,
        group: row.group
          ? {
              ...row.group,
              path: row.group.path ? [...row.group.path] : undefined
            }
          : undefined
      }));
      this.rowModelWindowCache.set(requestKey, rows);
      this.failedRowModelRequests.delete(requestKey);
      return rows;
    };
    const result = rowModel.getRows({ sheetId, startRow, endRow });

    if (typeof (result as Promise<RowResult>)?.then === "function") {
      if (!this.pendingRowModelRequests.has(requestKey)) {
        this.pendingRowModelRequests.add(requestKey);
        this.clearRowModelFeedback(sheetId);
        void (result as Promise<RowResult>)
          .then((resolved) => {
            const requestWasPending = this.pendingRowModelRequests.delete(requestKey);
            if (!requestWasPending) {
              return;
            }

            applyResult(resolved);
            if (!this.hasPendingRowModelRequests(sheetId)) {
              this.clearRowModelFeedback(sheetId);
            }
            this.render();
          })
          .catch((error) => {
            const requestWasPending = this.pendingRowModelRequests.delete(requestKey);
            if (!requestWasPending) {
              return;
            }

            this.failedRowModelRequests.add(requestKey);
            this.setRowModelError(sheetId, error);
            this.render();
          });
      }

      return cached;
    }

    return applyResult(result as RowResult);
  }

  private getRowHeight(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>, row: number): number {
    return sheet.rows[row]?.hidden ? 0 : sheet.rows[row]?.height ?? this.engine.getSnapshot().settings.rowHeight;
  }

  private isRowHidden(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>, row: number): boolean {
    return Boolean(sheet.rows[row]?.hidden);
  }

  private isColumnHidden(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>, col: number): boolean {
    return Boolean(sheet.columns[col]?.hidden);
  }

  private findNextVisibleIndex(
    axis: "row" | "column",
    sheet: ReturnType<WorkbookEngine["getActiveSheet"]>,
    current: number,
    direction: -1 | 1
  ): number {
    const max = axis === "row" ? sheet.rowCount - 1 : sheet.columnCount - 1;
    const isHidden = axis === "row" ? this.isRowHidden.bind(this) : this.isColumnHidden.bind(this);
    let candidate = Math.min(Math.max(current, 0), max);

    while (candidate >= 0 && candidate <= max && isHidden(sheet, candidate)) {
      candidate += direction;
    }

    return Math.min(Math.max(candidate, 0), max);
  }

  private shouldIncludeHiddenInClipboard(): boolean {
    return this.options.includeHiddenCellsInClipboard === true;
  }

  private getRenderedRowIndices(
    sheet: ReturnType<WorkbookEngine["getActiveSheet"]>,
    visibleRows: { start: number; end: number },
    resolvedRowCount: number,
    rowModelRows?: RowModelRow[]
  ): number[] {
    const rows: number[] = [];
    const frozen = this.engine.getFrozenPane(sheet.id).rows;
    appendRangeIndices(rows, 0, Math.max(-1, frozen - 1));
    if (rowModelRows?.length) {
      for (const row of rowModelRows) {
        if (!rows.includes(row.index)) {
          rows.push(row.index);
        }
      }
    } else {
      appendRangeIndices(rows, visibleRows.start, visibleRows.end);
    }
    return rows.filter((row) => row >= 0 && row < resolvedRowCount);
  }

  private getRenderedColumnIndices(
    sheet: ReturnType<WorkbookEngine["getActiveSheet"]>,
    visibleColumns: { start: number; end: number }
  ): number[] {
    const columns: number[] = [];
    const frozen = this.engine.getFrozenPane(sheet.id).columns;
    appendRangeIndices(columns, 0, Math.max(-1, frozen - 1));
    appendRangeIndices(columns, visibleColumns.start, visibleColumns.end);
    return columns.filter((col) => col >= 0 && col < sheet.columnCount);
  }

  private getFrozenAdjustedTop(sheetId: string, rowOffsets: number[], row: number): number {
    const frozenRows = this.engine.getFrozenPane(sheetId).rows;
    return row < frozenRows ? rowOffsets[row] + this.viewport.scrollTop : rowOffsets[row];
  }

  private getFrozenAdjustedLeft(sheetId: string, colOffsets: number[], col: number): number {
    const frozenColumns = this.engine.getFrozenPane(sheetId).columns;
    return col < frozenColumns ? colOffsets[col] + this.viewport.scrollLeft : colOffsets[col];
  }

  private getCellPrimitiveValue(sheetId: string, row: number, col: number): CellPrimitive {
    const cell = this.engine.getCell(sheetId, row, col);
    if (!cell) {
      return null;
    }

    return cell.formula ? cell.computedValue ?? null : cell.value;
  }

  private resolveAutofillPreview(sourceRange: CellRange, row: number, col: number): AutofillPreview | undefined {
    const rowDistance = row < sourceRange.start.row ? sourceRange.start.row - row : row > sourceRange.end.row ? row - sourceRange.end.row : 0;
    const colDistance = col < sourceRange.start.col ? sourceRange.start.col - col : col > sourceRange.end.col ? col - sourceRange.end.col : 0;

    if (rowDistance === 0 && colDistance === 0) {
      return undefined;
    }

    const axis: AutofillAxis =
      rowDistance > 0 && rowDistance >= colDistance ? "row" : rowDistance === 0 ? "column" : colDistance === 0 ? "row" : "column";

    if (axis === "row") {
      const startRow = row < sourceRange.start.row ? row : sourceRange.end.row + 1;
      const endRow = row < sourceRange.start.row ? sourceRange.start.row - 1 : row;
      if (startRow > endRow) {
        return undefined;
      }

      return {
        axis,
        fillRange: {
          start: { row: startRow, col: sourceRange.start.col },
          end: { row: endRow, col: sourceRange.end.col }
        }
      };
    }

    const startCol = col < sourceRange.start.col ? col : sourceRange.end.col + 1;
    const endCol = col < sourceRange.start.col ? sourceRange.start.col - 1 : col;
    if (startCol > endCol) {
      return undefined;
    }

    return {
      axis,
      fillRange: {
        start: { row: sourceRange.start.row, col: startCol },
        end: { row: sourceRange.end.row, col: endCol }
      }
    };
  }

  private getMappedAutofillSourceAddress(sourceRange: CellRange, preview: AutofillPreview, row: number, col: number): CellAddress {
    const sourceHeight = sourceRange.end.row - sourceRange.start.row + 1;
    const sourceWidth = sourceRange.end.col - sourceRange.start.col + 1;

    if (preview.axis === "row") {
      const rowOffset = row < sourceRange.start.row ? sourceRange.start.row - 1 - row : row - (sourceRange.end.row + 1);
      return {
        row: row < sourceRange.start.row ? sourceRange.end.row - (rowOffset % sourceHeight) : sourceRange.start.row + (rowOffset % sourceHeight),
        col: sourceRange.start.col + ((col - sourceRange.start.col) % sourceWidth)
      };
    }

    const colOffset = col < sourceRange.start.col ? sourceRange.start.col - 1 - col : col - (sourceRange.end.col + 1);
    return {
      row: sourceRange.start.row + ((row - sourceRange.start.row) % sourceHeight),
      col: col < sourceRange.start.col ? sourceRange.end.col - (colOffset % sourceWidth) : sourceRange.start.col + (colOffset % sourceWidth)
    };
  }

  private getAutofillSeriesValue(sheetId: string, sourceRange: CellRange, preview: AutofillPreview, row: number, col: number): CellPrimitive | undefined {
    const sourceHeight = sourceRange.end.row - sourceRange.start.row + 1;
    const sourceWidth = sourceRange.end.col - sourceRange.start.col + 1;

    if (preview.axis === "row" && sourceHeight > 1) {
      const seriesCol = sourceRange.start.col + ((col - sourceRange.start.col) % sourceWidth);
      const series = Array.from({ length: sourceHeight }, (_value, index) =>
        this.getCellPrimitiveValue(sheetId, sourceRange.start.row + index, seriesCol)
      );
      const relativeIndex = row - sourceRange.start.row;
      const numericSeries = series.map(toNumericValue);
      if (numericSeries.every((value) => value !== undefined)) {
        const step = (numericSeries[1] as number) - (numericSeries[0] as number);
        if (numericSeries.every((value, index) => index === 0 || value === (numericSeries[0] as number) + step * index)) {
          return (numericSeries[0] as number) + step * relativeIndex;
        }
      }

      const dateSeries = series.map(toDateValue);
      if (dateSeries.every((value) => value !== undefined)) {
        const step = (dateSeries[1] as number) - (dateSeries[0] as number);
        if (dateSeries.every((value, index) => index === 0 || value === (dateSeries[0] as number) + step * index)) {
          return formatDateValue((dateSeries[0] as number) + step * relativeIndex);
        }
      }
    }

    if (preview.axis === "column" && sourceWidth > 1) {
      const seriesRow = sourceRange.start.row + ((row - sourceRange.start.row) % sourceHeight);
      const series = Array.from({ length: sourceWidth }, (_value, index) =>
        this.getCellPrimitiveValue(sheetId, seriesRow, sourceRange.start.col + index)
      );
      const relativeIndex = col - sourceRange.start.col;
      const numericSeries = series.map(toNumericValue);
      if (numericSeries.every((value) => value !== undefined)) {
        const step = (numericSeries[1] as number) - (numericSeries[0] as number);
        if (numericSeries.every((value, index) => index === 0 || value === (numericSeries[0] as number) + step * index)) {
          return (numericSeries[0] as number) + step * relativeIndex;
        }
      }

      const dateSeries = series.map(toDateValue);
      if (dateSeries.every((value) => value !== undefined)) {
        const step = (dateSeries[1] as number) - (dateSeries[0] as number);
        if (dateSeries.every((value, index) => index === 0 || value === (dateSeries[0] as number) + step * index)) {
          return formatDateValue((dateSeries[0] as number) + step * relativeIndex);
        }
      }
    }

    return undefined;
  }

  private createRenderContext(
    sheetId: string,
    row: number,
    col: number,
    selected: boolean,
    active: boolean
  ): Readonly<CustomCellRenderContext> {
    const workbook = this.engine.getSnapshot();
    const cell = this.engine.getCell(sheetId, row, col);
    return freezeRenderContext({
      workbookId: workbook.id,
      sheetId,
      row,
      col,
      address: { row, col },
      cell,
      value: this.getCellPrimitiveValue(sheetId, row, col),
      displayValue: this.getRenderedCellDisplayValue(sheetId, row, col),
      validation: cell?.validation,
      selected,
      active
    });
  }

  private createEditorContext(sheetId: string, row: number, col: number): Readonly<CustomCellEditorContext> {
    const workbook = this.engine.getSnapshot();
    const cell = this.engine.getCell(sheetId, row, col);
    return freezeEditorContext({
      workbookId: workbook.id,
      sheetId,
      row,
      col,
      address: { row, col },
      cell,
      value: this.getCellPrimitiveValue(sheetId, row, col),
      displayValue: this.engine.getDisplayValue(sheetId, row, col),
      validation: cell?.validation
    });
  }

  private getCustomRenderer(context: Readonly<CustomCellRenderContext>): CustomCellRenderer | undefined {
    return this.options.cellRenderers?.find((renderer) => {
      try {
        return renderer.matches(context);
      } catch {
        return false;
      }
    });
  }

  private getCustomEditor(context: Readonly<CustomCellEditorContext>): CustomCellEditor | undefined {
    return this.options.cellEditors?.find((editor) => {
      try {
        return editor.matches(context);
      } catch {
        return false;
      }
    });
  }

  private getInteractiveValidationRule(sheetId: string, row: number, col: number): InteractiveValidationRule | undefined {
    return this.engine
      .getCellValidation(sheetId, row, col)
      ?.rules.find((rule): rule is InteractiveValidationRule => isListValidationRule(rule) || isCheckboxValidationRule(rule));
  }

  private getMergeAt(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>, row: number, col: number): SheetMerge | undefined {
    return sheet.merges.find((merge) => isWithinRange(row, col, merge));
  }

  private resolveCellAddress(
    sheet: ReturnType<WorkbookEngine["getActiveSheet"]>,
    row: number,
    col: number
  ): { address: CellAddress; merge?: SheetMerge } {
    const merge = this.getMergeAt(sheet, row, col);
    return {
      address: merge?.start ?? { row, col },
      merge
    };
  }

  private getResolvedSelectionRange(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>): CellRange {
    const merge = this.getMergeAt(sheet, sheet.selection.end.row, sheet.selection.end.col);
    return merge ?? sheet.selection;
  }

  private normalizeRange(range: CellRange): CellRange {
    return {
      start: {
        row: Math.min(range.start.row, range.end.row),
        col: Math.min(range.start.col, range.end.col)
      },
      end: {
        row: Math.max(range.start.row, range.end.row),
        col: Math.max(range.start.col, range.end.col)
      }
    };
  }

  private getUsedRange(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>): CellRange | undefined {
    const entries = Object.keys(sheet.cells);
    if (!entries.length) {
      return undefined;
    }

    let minRow = Number.POSITIVE_INFINITY;
    let maxRow = Number.NEGATIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxCol = Number.NEGATIVE_INFINITY;

    for (const entry of entries) {
      const match = /^(\d+):(\d+)$/.exec(entry);
      if (!match) {
        continue;
      }

      const row = Number(match[1]);
      const col = Number(match[2]);
      minRow = Math.min(minRow, row);
      maxRow = Math.max(maxRow, row);
      minCol = Math.min(minCol, col);
      maxCol = Math.max(maxCol, col);
    }

    if (!Number.isFinite(minRow) || !Number.isFinite(minCol)) {
      return undefined;
    }

    return {
      start: { row: minRow, col: minCol },
      end: { row: maxRow, col: maxCol }
    };
  }

  private isPivotCompatibleRange(range: CellRange | undefined): range is CellRange {
    return !!range && range.end.row > range.start.row && range.end.col >= range.start.col;
  }

  private getPivotSourceRange(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>): CellRange | undefined {
    const selectionRange = this.normalizeRange(this.getResolvedSelectionRange(sheet));
    if (this.isPivotCompatibleRange(selectionRange)) {
      return selectionRange;
    }

    const usedRange = this.getUsedRange(sheet);
    return this.isPivotCompatibleRange(usedRange) ? usedRange : undefined;
  }

  private createPivotFromActiveSheet(): void {
    this.openPivotPanel();
  }

  private getPivotPanelFields(sheetId: string, sourceRange: CellRange): string[] {
    return Array.from({ length: sourceRange.end.col - sourceRange.start.col + 1 }, (_value, index) =>
      this.engine.getDisplayValue(sheetId, sourceRange.start.row, sourceRange.start.col + index).trim()
    ).filter(Boolean);
  }

  private createPivotPanelValueState(
    fields: string[],
    value?: Partial<PivotPanelValueState>
  ): PivotPanelValueState {
    return {
      field: value?.field && fields.includes(value.field) ? value.field : (fields[0] ?? ""),
      aggregate: value?.aggregate ?? "sum",
      as: value?.as ?? ""
    };
  }

  private getPivotStatus(sheetId: string): { message?: string; isError: boolean } {
    const definition = this.engine.getPivotSheetViewDefinition(sheetId);
    if (!definition) {
      return { message: undefined, isError: false };
    }

    if (definition.refreshStatus === "refreshing") {
      return { message: this.messages.pivotRefreshing, isError: false };
    }

    if (definition.refreshStatus === "error") {
      return {
        message: definition.lastError ?? this.messages.pivotStale,
        isError: true
      };
    }

    if (definition.stale) {
      return { message: this.messages.pivotStale, isError: false };
    }

    return { message: undefined, isError: false };
  }

  private inferPivotPanelInput(sourceSheetId: string, sourceRange: CellRange) {
    try {
      return this.engine.inferPivotSheet({
        sourceSheetId,
        sourceRange
      });
    } catch {
      const fields = this.getPivotPanelFields(sourceSheetId, sourceRange);
      const valueField = fields.at(-1);
      if (!valueField) {
        throw new SpreadsheetOperationError({
          code: "RENDERER_PIVOT_SOURCE_INVALID",
          message: "Selecione um intervalo tabular com cabeçalho para criar a pivot.",
          area: "renderer",
          recoverable: true,
          details: {
            sourceSheetId,
            sourceRange
          }
        });
      }

      return {
        sourceSheetId,
        sourceRange,
        rows: fields[0] ? [fields[0]] : [],
        columns: fields.length > 2 && fields[1] ? [fields[1]] : [],
        values: [{ field: valueField, aggregate: "sum" as const }],
        includeRowTotals: true,
        includeColumnTotals: true,
        includeSubtotals: fields.length > 2
      };
    }
  }

  private openPivotPanel(): void {
    const activeSheet = this.engine.getActiveSheet();
    const persistedPivot = this.engine.getPivotSheetViewDefinition(activeSheet.id);
    const sourceSheetId = persistedPivot?.input.sourceSheetId ?? activeSheet.id;
    const sourceSheet = this.engine.getSnapshot().sheets.find((sheet) => sheet.id === sourceSheetId) ?? activeSheet;
    const sourceRange = persistedPivot?.input.sourceRange ?? this.getPivotSourceRange(sourceSheet);
    if (!sourceRange) {
      this.setPivotFeedback(activeSheet.id, "Selecione um intervalo tabular com cabeçalho para criar a pivot.", true);
      this.render();
      return;
    }

    try {
      const pivotInput = persistedPivot?.input ?? this.inferPivotPanelInput(sourceSheetId, sourceRange);
      this.findReplaceState.open = false;
      this.pivotPanelState.open = true;
      this.pivotPanelState.pending = false;
      this.pivotPanelState.progress = 0;
      this.pivotPanelState.sourceSheetId = sourceSheetId;
      this.pivotPanelState.targetSheetId = persistedPivot ? activeSheet.id : undefined;
      this.pivotPanelState.sourceRange = sourceRange;
      this.pivotPanelState.fields = this.getPivotPanelFields(sourceSheetId, sourceRange);
      this.pivotPanelState.rowFields = [...(pivotInput.rows ?? [])];
      this.pivotPanelState.columnFields = [...(pivotInput.columns ?? [])];
      this.pivotPanelState.values = (pivotInput.values.length ? pivotInput.values : [{}]).map((value) =>
        this.createPivotPanelValueState(this.pivotPanelState.fields, value)
      );
      this.pivotPanelState.includeRowTotals = pivotInput.includeRowTotals !== false;
      this.pivotPanelState.includeColumnTotals = pivotInput.includeColumnTotals !== false;
      this.pivotPanelState.includeSubtotals = pivotInput.includeSubtotals !== false;
      this.pivotPanelState.executionMode = pivotInput.executionMode ?? "auto";
      this.pivotPanelState.autoRefresh = persistedPivot?.autoRefresh !== false;
      this.pivotBuildController = undefined;
      this.clearPivotFeedback(activeSheet.id);
      this.render();
      this.pivotPanel.querySelector<HTMLElement>("[data-pivot-role='value-field']")?.focus();
    } catch (error) {
      this.setPivotFeedback(activeSheet.id, error instanceof Error ? error.message : "Falha ao criar a pivot.", true);
      this.render();
    }
  }

  private closePivotPanel(): void {
    if (this.pivotPanelState.pending) {
      this.pivotBuildController?.abort();
      return;
    }

    this.pivotPanelState.open = false;
    this.render();
    this.focus();
  }

  private setPivotFeedback(sheetId: string, message: string, isError = false): void {
    this.pivotFeedback = {
      sheetId,
      message,
      isError
    };
  }

  private clearPivotFeedback(sheetId?: string): void {
    if (!sheetId || this.pivotFeedback?.sheetId === sheetId) {
      this.pivotFeedback = undefined;
    }
  }

  private isAbortError(error: unknown): boolean {
    return error instanceof DOMException
      ? error.name === "AbortError"
      : typeof error === "object" && error !== null && "name" in error && (error as { name?: unknown }).name === "AbortError";
  }

  private formatPivotProgress(progress: PivotBuildProgress): string {
    if (progress.total <= 0) {
      return this.messages.pivotCreating;
    }

    const percentage = Math.max(0, Math.min(100, Math.round((progress.completed / progress.total) * 100)));
    return `${this.messages.pivotCreating} ${percentage}%`;
  }

  private resolvePivotErrorMessage(error: unknown): string {
    if (error instanceof Error && error.name === "CORE_PIVOT_CLIENT_ROW_LIMIT_EXCEEDED") {
      const details = "details" in error && typeof error.details === "object" && error.details !== null
        ? (error.details as { details?: { maxPivotSourceRows?: unknown } })
        : undefined;
      const limit = typeof details?.details?.maxPivotSourceRows === "number" ? details.details.maxPivotSourceRows : undefined;
      return limit ? `${this.messages.pivotTooLarge} Limite atual: ${limit} linhas.` : this.messages.pivotTooLarge;
    }

    return error instanceof Error ? error.message : "Falha ao criar a pivot.";
  }

  private renderPivotPanelSelect(
    select: HTMLSelectElement,
    value: string,
    options: string[],
    allowEmpty = true
  ): void {
    const nextOptions = allowEmpty ? ["", ...options] : options;
    select.replaceChildren(
      ...nextOptions.map((optionValue) => {
        const option = document.createElement("option");
        option.value = optionValue;
        option.textContent = optionValue || "Nenhum";
        return option;
      })
    );
    select.value = nextOptions.includes(value) ? value : nextOptions[0] ?? "";
  }

  private renderPivotPanelMultiSelect(select: HTMLSelectElement, values: string[], options: string[]): void {
    const selected = new Set(values);
    select.replaceChildren(
      ...options.map((optionValue) => {
        const option = document.createElement("option");
        option.value = optionValue;
        option.textContent = optionValue;
        option.selected = selected.has(optionValue);
        return option;
      })
    );
  }

  private renderPivotPanel(): void {
    this.pivotPanel.hidden = !this.pivotPanelState.open;
    if (!this.pivotPanelState.open) {
      return;
    }

    const fields = this.pivotPanelState.fields;
    this.renderPivotPanelMultiSelect(this.pivotRowSelect, this.pivotPanelState.rowFields, fields);
    this.renderPivotPanelMultiSelect(this.pivotColumnSelect, this.pivotPanelState.columnFields, fields);
    this.pivotValuesContainer.replaceChildren(
      ...this.pivotPanelState.values.map((valueState, index) => {
        const row = document.createElement("div");
        row.className = "excelsior-pivot-value-row";

        const fieldSelect = document.createElement("select");
        fieldSelect.className = "excelsior-find-replace-input";
        fieldSelect.dataset.pivotRole = "value-field";
        fieldSelect.dataset.pivotValueIndex = String(index);
        this.renderPivotPanelSelect(fieldSelect, valueState.field, fields, false);

        const aggregateSelect = document.createElement("select");
        aggregateSelect.className = "excelsior-find-replace-input";
        aggregateSelect.dataset.pivotRole = "value-aggregate";
        aggregateSelect.dataset.pivotValueIndex = String(index);
        aggregateSelect.replaceChildren(
          ...(["sum", "avg", "min", "max", "count"] as PivotAggregateFunction[]).map((aggregate) => {
            const option = document.createElement("option");
            option.value = aggregate;
            option.textContent = createPivotAggregateLabel(this.messages, aggregate);
            return option;
          })
        );
        aggregateSelect.value = valueState.aggregate;

        const aliasInput = document.createElement("input");
        aliasInput.type = "text";
        aliasInput.className = "excelsior-find-replace-input";
        aliasInput.dataset.pivotRole = "value-alias";
        aliasInput.dataset.pivotValueIndex = String(index);
        aliasInput.value = valueState.as;
        aliasInput.placeholder = this.messages.pivotAlias;

        const removeButton = createFindReplaceActionButton("remove-value", this.messages.pivotRemoveValue);
        removeButton.dataset.pivotAction = "remove-value";
        removeButton.dataset.pivotValueIndex = String(index);
        delete removeButton.dataset.findAction;
        removeButton.disabled = this.pivotPanelState.values.length <= 1;

        row.append(fieldSelect, aggregateSelect, aliasInput, removeButton);
        return row;
      })
    );
    this.pivotIncludeRowTotals.checked = this.pivotPanelState.includeRowTotals;
    this.pivotIncludeColumnTotals.checked = this.pivotPanelState.includeColumnTotals;
    this.pivotIncludeSubtotals.checked = this.pivotPanelState.includeSubtotals;
    this.pivotExecutionModeSelect.value = this.pivotPanelState.executionMode;
    this.pivotAutoRefresh.checked = this.pivotPanelState.autoRefresh;
    let pivotApplyLabel = this.pivotPanelState.targetSheetId ? this.messages.pivotUpdate : this.messages.pivotApply;
    if (this.pivotPanelState.pending) {
      pivotApplyLabel =
        this.pivotPanelState.progress > 0
          ? `${this.messages.pivotCreating} ${this.pivotPanelState.progress}%`
          : this.messages.pivotCreating;
    }
    this.pivotApplyButton.textContent = pivotApplyLabel;
    this.pivotCloseButton.textContent = this.pivotPanelState.pending ? this.messages.pivotCancel : this.messages.closePanel;

    for (const control of this.pivotPanel.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLButtonElement>(
      "input, select, button"
    )) {
      if (control.dataset.pivotAction === "close") {
        control.disabled = false;
        continue;
      }
      control.disabled = this.pivotPanelState.pending;
    }
  }

  private getPivotPanelSelection(select: HTMLSelectElement): string[] {
    return Array.from(select.selectedOptions, (option) => option.value).filter(Boolean);
  }

  private async applyPivotPanel(): Promise<void> {
    const { sourceSheetId, sourceRange, targetSheetId } = this.pivotPanelState;
    if (!sourceSheetId || !sourceRange || this.pivotPanelState.pending) {
      return;
    }

    const valueFields = new Set(this.pivotPanelState.values.map((value) => value.field).filter(Boolean));
    if (!valueFields.size) {
      return;
    }

    const rows = Array.from(new Set(this.pivotPanelState.rowFields)).filter((field) => field && !valueFields.has(field));
    const rowSet = new Set(rows);
    const columns = Array.from(new Set(this.pivotPanelState.columnFields)).filter(
      (field) => field && !valueFields.has(field) && !rowSet.has(field)
    );
    const values = this.pivotPanelState.values
      .filter((value) => value.field)
      .map((value) => ({
        field: value.field,
        aggregate: value.aggregate,
        ...(value.as.trim() ? { as: value.as.trim() } : {})
      }));

    if (!values.length) {
      return;
    }

    const feedbackSheetId = targetSheetId ?? sourceSheetId;
    const pivotInput = {
      sourceSheetId,
      sourceRange,
      rows,
      columns,
      values,
      includeRowTotals: this.pivotPanelState.includeRowTotals,
      includeColumnTotals: this.pivotPanelState.includeColumnTotals,
      includeSubtotals: this.pivotPanelState.includeSubtotals && rows.length + columns.length > 1,
      executionMode: this.pivotPanelState.executionMode
    } as const;

    try {
      const controller = new AbortController();
      const progressHandler = (progress: PivotBuildProgress) => {
        this.pivotPanelState.progress = progress.total > 0 ? Math.round((progress.completed / progress.total) * 100) : 0;
        this.setPivotFeedback(feedbackSheetId, this.formatPivotProgress(progress));
        this.render();
      };

      this.pivotBuildController = controller;
      this.pivotPanelState.pending = true;
      this.pivotPanelState.progress = 0;
      this.setPivotFeedback(feedbackSheetId, targetSheetId ? this.messages.pivotRefreshing : this.messages.pivotCreating);
      this.render();

      const pivotSheetId = targetSheetId
        ? (await this.engine.updatePivotSheet(targetSheetId, pivotInput, {
            signal: controller.signal,
            onProgress: progressHandler
          }), targetSheetId)
        : await this.engine.addPivotSheetAsync(pivotInput, {
            signal: controller.signal,
            onProgress: progressHandler
          });

      this.engine.setPivotSheetAutoRefresh(pivotSheetId, this.pivotPanelState.autoRefresh);
      this.pivotPanelState.pending = false;
      this.pivotPanelState.progress = 0;
      this.pivotPanelState.open = false;
      this.pivotBuildController = undefined;
      this.clearPivotFeedback(feedbackSheetId);
      this.render();
      this.focus();
    } catch (error) {
      this.pivotBuildController = undefined;
      this.pivotPanelState.pending = false;
      this.pivotPanelState.progress = 0;
      if (this.isAbortError(error)) {
        this.pivotPanelState.open = false;
        this.setPivotFeedback(feedbackSheetId, this.messages.pivotCancelled);
        this.render();
        this.focus();
        return;
      }

      this.setPivotFeedback(feedbackSheetId, this.resolvePivotErrorMessage(error), true);
      this.render();
    }
  }

  private getActiveAddress(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>): CellAddress {
    return this.resolveCellAddress(sheet, sheet.selection.end.row, sheet.selection.end.col).address;
  }

  private getCellStyle(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>, row: number, col: number): CellStyle | undefined {
    const resolved = this.resolveCellAddress(sheet, row, col).address;
    const cell = this.engine.getCell(sheet.id, resolved.row, resolved.col);
    const conditionalStyle = this.engine.getConditionalStyle(sheet.id, resolved.row, resolved.col);
    return normalizeStyle({
      ...sheet.columns[col]?.style,
      ...sheet.rows[row]?.style,
      ...cell?.style,
      ...conditionalStyle,
      border: {
        ...sheet.columns[col]?.style?.border,
        ...sheet.rows[row]?.style?.border,
        ...cell?.style?.border,
        ...conditionalStyle?.border
      }
    });
  }

  private setValidationError(error: CellValidationError): void {
    this.validationFeedback = {
      sheetId: error.sheetId,
      row: error.address.row,
      col: error.address.col,
      message: error.message
    };
  }

  private commitCellValue(
    sheetId: string,
    row: number,
    col: number,
    value: string | number | boolean | null,
    onValidationError: () => void
  ): boolean {
    try {
      this.engine.setCellValue({
        sheetId,
        row,
        col,
        value
      });
      this.clearValidationFeedback();
      return true;
    } catch (error) {
      if (error instanceof CellValidationError) {
        this.setValidationError(error);
        this.render();
        onValidationError();
        return false;
      }

      throw error;
    }
  }

  private renderCellContent(
    cell: HTMLElement,
    sheetId: string,
    row: number,
    col: number,
    rowModelRow?: RowModelRow
  ): void {
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    const context = this.createRenderContext(
      sheetId,
      row,
      col,
      rangesOverlap({ start: { row, col }, end: { row, col } }, sheet.selection),
      activeAddress.row === row && activeAddress.col === col
    );
    const renderer = this.getCustomRenderer(context);

    if (renderer) {
      try {
        const output = renderer.render(context);
        cell.replaceChildren();
        for (const className of sanitizeClassNames(output.classNames)) {
          cell.classList.add(className);
        }
        if (output.parts?.length) {
          for (const part of output.parts) {
            const node = document.createElement("span");
            node.className = "excelsior-cell-part";
            if (part.tone && part.tone !== "default") {
              node.classList.add(`is-${part.tone}`);
            }
            node.textContent = part.text;
            cell.append(node);
          }
        } else {
          const content = document.createElement("span");
          content.className = "excelsior-cell-content";
          content.textContent = output.text ?? context.displayValue;
          cell.append(content);
        }

        if (output.accessoryText) {
          const accessory = document.createElement("span");
          accessory.className = "excelsior-cell-affordance";
          accessory.setAttribute("aria-hidden", "true");
          accessory.textContent = output.accessoryText;
          cell.append(accessory);
        }

        if (output.title) {
          cell.title = output.title;
        }
        if (output.ariaLabel) {
          cell.setAttribute("aria-label", output.ariaLabel);
        }
        return;
      } catch {
        cell.replaceChildren();
      }
    }

    const rowModel = this.engine.getRowModel(sheetId);
    const visibleGroup = rowModel.kind !== "clientSide" ? this.getVisibleRemoteGroup(rowModelRow, col) : undefined;

    if (visibleGroup) {
      cell.classList.add("has-remote-group");
      const toggle = document.createElement("button");
      const path = visibleGroup.path?.length ? visibleGroup.path : [visibleGroup.key];
      const isExpanded = visibleGroup.expanded !== false;
      toggle.type = "button";
      toggle.className = "excelsior-group-toggle";
      toggle.dataset.remoteGroupToggle = "true";
      toggle.dataset.remoteGroupPath = JSON.stringify(path);
      toggle.dataset.remoteGroupExpanded = String(isExpanded);
      toggle.setAttribute("aria-expanded", String(isExpanded));
      toggle.setAttribute("aria-label", `${isExpanded ? this.messages.collapseGroup : this.messages.expandGroup}: ${visibleGroup.key}`);
      toggle.textContent = isExpanded ? "▾" : "▸";

      const content = document.createElement("span");
      content.className = "excelsior-cell-content excelsior-group-label";
      content.textContent = context.displayValue || visibleGroup.key;

      cell.replaceChildren(toggle, content);

      if (typeof visibleGroup.childCount === "number") {
        const meta = document.createElement("span");
        meta.className = "excelsior-cell-affordance excelsior-group-meta";
        meta.setAttribute("aria-hidden", "true");
        meta.textContent = String(visibleGroup.childCount);
        cell.append(meta);
      }
      return;
    }

    const value = context.displayValue;
    const validationRule = this.getInteractiveValidationRule(sheetId, row, col);

    if (!validationRule) {
      cell.textContent = value;
      return;
    }

    cell.classList.add("has-validation-ui");

    if (isCheckboxValidationRule(validationRule)) {
      cell.classList.add("is-checkbox-validation");
      const affordance = document.createElement("span");
      affordance.className = "excelsior-cell-affordance is-checkbox";
      affordance.setAttribute("aria-hidden", "true");
      affordance.textContent = toBoolean(this.getCellPrimitiveValue(sheetId, row, col)) ? "☑" : "☐";
      if (!cell.title) {
        cell.title = this.messages.checkboxHint;
      }
      cell.replaceChildren(affordance);
      return;
    }

    cell.classList.add("is-dropdown-validation");

    const content = document.createElement("span");
    content.className = "excelsior-cell-content";
    content.textContent = value;

    const affordance = document.createElement("span");
    affordance.className = "excelsior-cell-affordance";
    affordance.setAttribute("aria-hidden", "true");
    affordance.textContent = "▾";

    if (!cell.title) {
      cell.title = this.messages.dropdownHint;
    }

    cell.replaceChildren(content, affordance);
  }

  private hideInlineEditors(): void {
    this.editor.hidden = true;
    this.selectEditor.hidden = true;
    this.customEditorHost.hidden = true;
  }

  private destroyCustomEditor(): void {
    this.activeCustomEditor?.destroy?.();
    this.activeCustomEditor = undefined;
    this.customEditorHost.replaceChildren();
  }

  private positionEditorElement(element: HTMLElement, row: number, col: number): void {
    const sheet = this.engine.getActiveSheet();
    const rowOffsets = buildOffsets(sheet.rowCount, (index) => this.getRowHeight(sheet, index));
    const colOffsets = buildOffsets(sheet.columnCount, (index) => this.getColumnWidth(sheet, index));
    const merge = this.getMergeAt(sheet, row, col);
    const targetRange = merge ?? {
      start: { row, col },
      end: { row, col }
    };
    element.style.top = `${rowOffsets[targetRange.start.row]}px`;
    element.style.left = `${colOffsets[targetRange.start.col]}px`;
    element.style.width = `${getSpanSize(colOffsets, targetRange.start.col, targetRange.end.col) - 2}px`;
    element.style.height = `${getSpanSize(rowOffsets, targetRange.start.row, targetRange.end.row) - 2}px`;
  }

  private applyCellPresentation(cell: HTMLElement, style: CellStyle | undefined): void {
    if (!style) {
      return;
    }

    cell.style.textAlign = style.align ?? "left";
    cell.style.justifyContent =
      style.align === "center" ? "center" : style.align === "right" ? "flex-end" : "flex-start";
    cell.style.alignItems =
      style.alignVertical === "top"
        ? "flex-start"
        : style.alignVertical === "bottom"
          ? "flex-end"
          : "center";
    cell.style.backgroundColor = style.backgroundColor ?? "";
    cell.style.color = style.textColor ?? "";
    cell.style.fontFamily = style.fontFamily ?? "";
    cell.style.fontSize = style.fontSize ? `${style.fontSize}px` : "";
    cell.style.fontWeight = style.fontWeight ?? "";
    cell.style.fontStyle = style.fontStyle ?? "";
    cell.style.textDecoration = style.underline ? "underline" : "";
    cell.style.whiteSpace = style.wrap ? "normal" : "nowrap";
    cell.style.lineHeight = style.wrap ? "1.35" : "1";
    cell.style.paddingLeft = `${8 + (style.indent ?? 0) * 8}px`;
    cell.style.borderTopStyle = style.border?.top?.style ?? "solid";
    cell.style.borderTopColor = style.border?.top?.color ?? "";
    cell.style.borderRightStyle = style.border?.right?.style ?? "solid";
    cell.style.borderRightColor = style.border?.right?.color ?? "";
    cell.style.borderBottomStyle = style.border?.bottom?.style ?? "solid";
    cell.style.borderBottomColor = style.border?.bottom?.color ?? "";
    cell.style.borderLeftStyle = style.border?.left?.style ?? "solid";
    cell.style.borderLeftColor = style.border?.left?.color ?? "";

    if (style.wrap) {
      cell.classList.add("is-wrapped");
    }
  }

  private selectResolvedCell(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>, row: number, col: number): void {
    const { address, merge } = this.resolveCellAddress(sheet, row, col);
    this.engine.selectRange({
      sheetId: sheet.id,
      rowStart: merge?.start.row ?? address.row,
      rowEnd: merge?.end.row ?? address.row,
      colStart: merge?.start.col ?? address.col,
      colEnd: merge?.end.col ?? address.col
    });
  }

  private applyStyleToSelection(style: Partial<CellStyle>): void {
    const sheet = this.engine.getActiveSheet();
    const processed = new Set<string>();
    for (let row = sheet.selection.start.row; row <= sheet.selection.end.row; row += 1) {
      for (let col = sheet.selection.start.col; col <= sheet.selection.end.col; col += 1) {
        const { address } = this.resolveCellAddress(sheet, row, col);
        const key = getCellKey(address.row, address.col);
        if (processed.has(key)) {
          continue;
        }
        processed.add(key);
        this.engine.setCellStyle({
          sheetId: sheet.id,
          row: address.row,
          col: address.col,
          style
        });
      }
    }
  }

  private getActiveFindReplaceMatch(): FindReplaceMatch | undefined {
    if (this.findReplaceState.activeIndex < 0) {
      return undefined;
    }

    return this.findReplaceState.matches[this.findReplaceState.activeIndex];
  }

  private syncFindReplaceStateFromControls(): void {
    this.findReplaceState.query = this.findReplaceQueryInput.value;
    this.findReplaceState.replaceText = this.findReplaceValueInput.value;
    this.findReplaceState.scope = (this.findReplacePanel.querySelector<HTMLInputElement>("[data-find-role='scope-workbook']")?.checked ?? false)
      ? "workbook"
      : "sheet";
    this.findReplaceState.caseSensitive =
      this.findReplacePanel.querySelector<HTMLInputElement>("[data-find-role='case-sensitive']")?.checked ?? false;
    this.findReplaceState.wholeCell =
      this.findReplacePanel.querySelector<HTMLInputElement>("[data-find-role='whole-cell']")?.checked ?? false;
    this.findReplaceState.regex = this.findReplacePanel.querySelector<HTMLInputElement>("[data-find-role='regex']")?.checked ?? false;
  }

  private cancelFindReplaceSearch(): void {
    this.findReplaceSearchVersion += 1;
    cancelScheduledFrame(this.findReplaceSearchHandle);
    this.findReplaceSearchHandle = undefined;
    this.findReplaceState.pending = false;
  }

  private focusFindReplaceMatch(match: FindReplaceMatch): void {
    if (this.engine.getSnapshot().activeSheetId !== match.sheetId) {
      this.engine.setActiveSheet(match.sheetId);
    }

    this.engine.selectRange({
      sheetId: match.sheetId,
      rowStart: match.row,
      rowEnd: match.row,
      colStart: match.col,
      colEnd: match.col
    });
    this.focus();
  }

  private navigateFindReplaceResults(step: number): void {
    if (!this.findReplaceState.matches.length) {
      return;
    }

    const nextIndex =
      (this.findReplaceState.activeIndex + step + this.findReplaceState.matches.length) % this.findReplaceState.matches.length;
    this.findReplaceState.activeIndex = nextIndex;
    const match = this.getActiveFindReplaceMatch();
    if (match) {
      this.focusFindReplaceMatch(match);
    }
    this.render();
  }

  private scheduleFindReplaceSearch(): void {
    this.syncFindReplaceStateFromControls();
    this.cancelFindReplaceSearch();

    const query = this.findReplaceState.query.trim();
    if (!query) {
      this.findReplaceState.matches = [];
      this.findReplaceState.activeIndex = -1;
      this.findReplaceState.error = undefined;
      this.render();
      return;
    }

    const preparedResult = createFindReplacePrepared({
      query: this.findReplaceState.query,
      replaceText: this.findReplaceState.replaceText,
      scope: this.findReplaceState.scope,
      caseSensitive: this.findReplaceState.caseSensitive,
      wholeCell: this.findReplaceState.wholeCell,
      regex: this.findReplaceState.regex
    });

    if (!preparedResult.prepared) {
      this.findReplaceState.matches = [];
      this.findReplaceState.activeIndex = -1;
      this.findReplaceState.error = preparedResult.error;
      this.render();
      return;
    }

    const snapshot = this.engine.getSnapshot();
    const entries = collectFindReplaceEntries(snapshot, snapshot.activeSheetId, this.findReplaceState.scope);
    const prepared = preparedResult.prepared;
    const version = this.findReplaceSearchVersion + 1;
    this.findReplaceSearchVersion = version;
    this.findReplaceState.pending = true;
    this.findReplaceState.error = undefined;
    this.findReplaceState.matches = [];
    this.findReplaceState.activeIndex = -1;
    this.render();

    const matches: FindReplaceMatch[] = [];
    let index = 0;
    const chunkSize = 250;
    const finalizeSearch = (): void => {
      this.findReplaceSearchHandle = undefined;
      this.findReplaceState.pending = false;
      this.findReplaceState.matches = matches;
      this.findReplaceState.activeIndex = matches.length ? 0 : -1;
      if (matches[0]) {
        this.focusFindReplaceMatch(matches[0]);
      }
      this.render();
    };

    if (entries.length <= chunkSize) {
      matches.push(...collectFindReplaceMatches(entries, prepared, 0, entries.length));
      finalizeSearch();
      return;
    }

    const runChunk = (): void => {
      if (version !== this.findReplaceSearchVersion) {
        return;
      }

      matches.push(...collectFindReplaceMatches(entries, prepared, index, index + chunkSize));
      index += chunkSize;

      if (index < entries.length) {
        this.findReplaceSearchHandle = scheduleFrame(runChunk);
        return;
      }

      finalizeSearch();
    };

    this.findReplaceSearchHandle = scheduleFrame(runChunk);
  }

  private replaceActiveFindReplaceMatch(): void {
    const match = this.getActiveFindReplaceMatch();
    if (!match) {
      return;
    }

    const preparedResult = createFindReplacePrepared({
      query: this.findReplaceState.query,
      replaceText: this.findReplaceState.replaceText,
      scope: this.findReplaceState.scope,
      caseSensitive: this.findReplaceState.caseSensitive,
      wholeCell: this.findReplaceState.wholeCell,
      regex: this.findReplaceState.regex
    });

    if (!preparedResult.prepared) {
      this.findReplaceState.error = preparedResult.error;
      this.render();
      return;
    }

    const cell = this.engine.getCell(match.sheetId, match.row, match.col);
    const text = getSearchableCellText(cell);
    const nextValue = preparedResult.prepared.replace(text);
    this.engine.setCellValue({
      sheetId: match.sheetId,
      row: match.row,
      col: match.col,
      value: nextValue
    });
  }

  private replaceAllFindReplaceMatches(): void {
    if (!this.findReplaceState.matches.length) {
      return;
    }

    const preparedResult = createFindReplacePrepared({
      query: this.findReplaceState.query,
      replaceText: this.findReplaceState.replaceText,
      scope: this.findReplaceState.scope,
      caseSensitive: this.findReplaceState.caseSensitive,
      wholeCell: this.findReplaceState.wholeCell,
      regex: this.findReplaceState.regex
    });

    if (!preparedResult.prepared) {
      this.findReplaceState.error = preparedResult.error;
      this.render();
      return;
    }

    const operations: SpreadsheetOperation[] = [];
    const affectedRanges: CellRange[] = [];
    for (const match of this.findReplaceState.matches) {
      const cell = this.engine.getCell(match.sheetId, match.row, match.col);
      const nextValue = preparedResult.prepared.replace(getSearchableCellText(cell));
      const nextCell: CellModel = {
        ...cell,
        value: nextValue,
        computedValue: nextValue,
        formula: undefined,
        error: undefined
      };

      if (typeof nextValue === "string" && nextValue.startsWith("=")) {
        nextCell.formula = nextValue;
        nextCell.computedValue = null;
      }

      operations.push({
        op: cell ? "replace" : "add",
        id: match.sheetId,
        path: ["cells", getCellKey(match.row, match.col)],
        value: nextCell
      });
      affectedRanges.push({
        start: { row: match.row, col: match.col },
        end: { row: match.row, col: match.col }
      });
    }

    if (!operations.length) {
      return;
    }

    this.engine.applyBatchOperations({
      anchorSheetId: this.engine.getSnapshot().activeSheetId,
      operations,
      affectedRanges
    });
  }

  private openFindReplacePanel(): void {
    this.pivotPanelState.open = false;
    if (!this.findReplaceState.open) {
      this.findReplaceState.open = true;
      this.findReplaceQueryInput.value = this.findReplaceState.query;
      this.findReplaceValueInput.value = this.findReplaceState.replaceText;
    }

    this.render();
    this.findReplaceQueryInput.focus();
    this.findReplaceQueryInput.select();
    this.scheduleFindReplaceSearch();
  }

  private closeFindReplacePanel(): void {
    this.cancelFindReplaceSearch();
    this.findReplaceState.open = false;
    this.findReplaceState.error = undefined;
    this.findReplaceState.pending = false;
    this.findReplaceState.matches = [];
    this.findReplaceState.activeIndex = -1;
    this.render();
    this.focus();
  }

  private renderFindReplacePanel(): void {
    this.findReplacePanel.hidden = !this.findReplaceState.open;
    this.findReplaceQueryInput.value = this.findReplaceState.query;
    this.findReplaceValueInput.value = this.findReplaceState.replaceText;

    const workbookToggle = this.findReplacePanel.querySelector<HTMLInputElement>("[data-find-role='scope-workbook']");
    const caseToggle = this.findReplacePanel.querySelector<HTMLInputElement>("[data-find-role='case-sensitive']");
    const wholeCellToggle = this.findReplacePanel.querySelector<HTMLInputElement>("[data-find-role='whole-cell']");
    const regexToggle = this.findReplacePanel.querySelector<HTMLInputElement>("[data-find-role='regex']");

    if (workbookToggle) {
      workbookToggle.checked = this.findReplaceState.scope === "workbook";
    }
    if (caseToggle) {
      caseToggle.checked = this.findReplaceState.caseSensitive;
    }
    if (wholeCellToggle) {
      wholeCellToggle.checked = this.findReplaceState.wholeCell;
    }
    if (regexToggle) {
      regexToggle.checked = this.findReplaceState.regex;
    }

    const activeIndex = this.findReplaceState.activeIndex + 1;
    const count = this.findReplaceState.matches.length;
    this.findReplaceResults.textContent = this.findReplaceState.error
      ? this.findReplaceState.error
      : this.findReplaceState.pending
        ? this.messages.searching
        : count
          ? `${activeIndex}/${count}`
          : this.findReplaceState.query
            ? this.messages.noResults
            : "";
    this.findReplaceResults.classList.toggle("is-error", Boolean(this.findReplaceState.error));

    const disableNavigation = !count || this.findReplaceState.pending;
    for (const role of ["previous", "next", "replace", "replace-all"]) {
      const button = this.findReplacePanel.querySelector<HTMLButtonElement>(`[data-find-action='${role}']`);
      if (!button) {
        continue;
      }
      button.disabled = role === "replace-all" ? disableNavigation : disableNavigation;
    }
  }

  private applyAutofill(preview: AutofillPreview): void {
    const sheet = this.engine.getActiveSheet();
    const sourceRange = this.autofillDrag?.sourceRange;
    if (!sourceRange) {
      return;
    }

    if (getRangeCellCount(preview.fillRange) > this.getAutofillMaxCells()) {
      return;
    }

    const operations: SpreadsheetOperation[] = [];
    const copyStyle = this.shouldCopyStyleOnAutofill();

    for (let row = preview.fillRange.start.row; row <= preview.fillRange.end.row; row += 1) {
      for (let col = preview.fillRange.start.col; col <= preview.fillRange.end.col; col += 1) {
        const sourceAddress = this.getMappedAutofillSourceAddress(sourceRange, preview, row, col);
        const sourceCell = this.engine.getCell(sheet.id, sourceAddress.row, sourceAddress.col);
        const targetCell = this.engine.getCell(sheet.id, row, col);
        const seriesValue = this.getAutofillSeriesValue(sheet.id, sourceRange, preview, row, col);
        const nextValue = sourceCell?.formula
          ? adjustFormulaReferences(sourceCell.formula, row - sourceAddress.row, col - sourceAddress.col)
          : seriesValue ?? (sourceCell ? this.getCellPrimitiveValue(sheet.id, sourceAddress.row, sourceAddress.col) : null);
        const validation = this.engine.validateCellValue({
          sheetId: sheet.id,
          row,
          col,
          value: nextValue
        });

        if (!validation.valid) {
          this.validationFeedback = {
            sheetId: sheet.id,
            row,
            col,
            message: validation.error?.message ?? "Valor inválido para a célula selecionada."
          };
          this.render();
          return;
        }

        const nextCell: CellModel = {
          ...targetCell,
          value: nextValue,
          computedValue: nextValue,
          formula: undefined,
          error: undefined
        };

        if (typeof nextValue === "string" && nextValue.startsWith("=")) {
          nextCell.formula = nextValue;
          nextCell.computedValue = null;
        }

        if (copyStyle) {
          nextCell.style = cloneCellStyle(sourceCell?.style);
        }

        operations.push({
          op: targetCell ? "replace" : "add",
          id: sheet.id,
          path: ["cells", getCellKey(row, col)],
          value: nextCell
        });
      }
    }

    if (!operations.length) {
      return;
    }

    operations.push({
      op: "replace",
      id: sheet.id,
      path: ["selection"],
      value: {
        start: {
          row: Math.min(sourceRange.start.row, preview.fillRange.start.row),
          col: Math.min(sourceRange.start.col, preview.fillRange.start.col)
        },
        end: {
          row: Math.max(sourceRange.end.row, preview.fillRange.end.row),
          col: Math.max(sourceRange.end.col, preview.fillRange.end.col)
        }
      }
    });

    this.engine.applyBatchOperations({
      anchorSheetId: sheet.id,
      operations,
      affectedRanges: [sourceRange, preview.fillRange]
    });
    this.clearValidationFeedback();
  }

  private readonly handleAutofillMouseDown = (event: MouseEvent): void => {
    if (!this.isAutofillEnabled()) {
      return;
    }

    const handle = getElementFromEventTarget(event.target)?.closest<HTMLElement>(".excelsior-fill-handle");
    if (!handle) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    const sheet = this.engine.getActiveSheet();
    this.autofillDrag = {
      sourceRange: this.getResolvedSelectionRange(sheet)
    };
  };

  private readonly handleAutofillMouseMove = (event: MouseEvent): void => {
    if (!this.autofillDrag) {
      return;
    }

    const target = getElementFromEventTarget(event.target)?.closest<HTMLElement>("[data-row][data-col]");
    if (!target) {
      return;
    }

    const preview = this.resolveAutofillPreview(
      this.autofillDrag.sourceRange,
      Number(target.dataset.row),
      Number(target.dataset.col)
    );

    this.autofillDrag.preview =
      preview && getRangeCellCount(preview.fillRange) <= this.getAutofillMaxCells() ? preview : undefined;
    this.render();
  };

  private readonly handleAutofillMouseUp = (event: MouseEvent): void => {
    if (!this.autofillDrag) {
      return;
    }

    const target = getElementFromEventTarget(event.target)?.closest<HTMLElement>("[data-row][data-col]");
    if (target) {
      const preview = this.resolveAutofillPreview(
        this.autofillDrag.sourceRange,
        Number(target.dataset.row),
        Number(target.dataset.col)
      );
      if (preview && getRangeCellCount(preview.fillRange) <= this.getAutofillMaxCells()) {
        this.autofillDrag.preview = preview;
      }
    }

    const preview = this.autofillDrag.preview;
    if (preview) {
      this.applyAutofill(preview);
    }

    this.autofillDrag = undefined;
    this.render();
  };

  constructor(
    private readonly container: HTMLElement,
    private readonly engine: WorkbookEngine,
    private readonly options: DomSpreadsheetRendererOptions = {}
  ) {
    this.messages = {
      ...DEFAULT_RENDERER_MESSAGES,
      ...options.localization?.messages
    };
    this.shortcuts = {
      ...DEFAULT_RENDERER_SHORTCUTS,
      ...options.localization?.shortcuts
    };
    this.direction = options.localization?.direction ?? "ltr";
    this.root.className = "excelsior-shell";
    this.chrome.className = "excelsior-chrome";
    this.toolbar.className = "excelsior-toolbar";
    this.formulaBar.className = "excelsior-formula-bar";
    this.formulaAddress.className = "excelsior-formula-address";
    this.formulaInput.className = "excelsior-formula-input";
    this.statusMessage.className = "excelsior-status-message";
    this.statusMessage.setAttribute("role", "status");
    this.statusMessage.setAttribute("aria-live", "polite");
    this.statusMessage.setAttribute("aria-atomic", "true");
    this.findReplacePanel.className = "excelsior-find-replace";
    this.pivotPanel.className = "excelsior-find-replace excelsior-pivot-panel";
    this.viewport.className = "excelsior-viewport";
    this.viewport.tabIndex = 0;
    this.viewport.setAttribute("role", "grid");
    this.viewport.setAttribute("aria-label", this.messages.gridLabel);
    this.viewport.setAttribute("aria-multiselectable", "true");
    this.root.dir = this.direction;
    this.viewport.dir = this.direction;
    this.root.classList.toggle("is-rtl", this.isRtl());
    this.surface.className = "excelsior-surface";
    this.cellsLayer.className = "excelsior-cells";
    this.editor.className = "excelsior-editor";
    this.selectEditor.className = "excelsior-select-editor";
    this.customEditorHost.className = "excelsior-custom-editor-host";
    this.sheetTabs.className = "excelsior-sheet-tabs";
    this.editor.type = "text";
    this.editor.hidden = true;
    this.selectEditor.hidden = true;
    this.customEditorHost.hidden = true;
    this.formulaInput.type = "text";
    this.formulaInput.setAttribute("aria-label", this.messages.formulaInputLabel);
    this.findReplacePanel.hidden = true;
    this.findReplaceQueryInput.type = "text";
    this.findReplaceQueryInput.placeholder = this.messages.findPlaceholder;
    this.findReplaceQueryInput.className = "excelsior-find-replace-input";
    this.findReplaceValueInput.type = "text";
    this.findReplaceValueInput.placeholder = this.messages.replacePlaceholder;
    this.findReplaceValueInput.className = "excelsior-find-replace-input";
    this.findReplaceResults.className = "excelsior-find-replace-results";
    this.findReplaceResults.setAttribute("role", "status");
    this.findReplaceResults.setAttribute("aria-live", "polite");
    this.findReplaceResults.setAttribute("aria-atomic", "true");
    this.activeCellAnnouncement.className = "excelsior-visually-hidden";
    this.activeCellAnnouncement.setAttribute("aria-live", "polite");
    this.activeCellAnnouncement.setAttribute("aria-atomic", "true");
    this.sheetTabs.setAttribute("role", "tablist");
    this.sheetTabs.setAttribute("aria-label", this.messages.sheetTabsLabel);

    const queryField = document.createElement("label");
    queryField.className = "excelsior-find-replace-field";
    const queryLabel = document.createElement("span");
    queryLabel.textContent = this.messages.findLabel;
    queryField.append(queryLabel, this.findReplaceQueryInput);

    const replaceField = document.createElement("label");
    replaceField.className = "excelsior-find-replace-field";
    const replaceLabel = document.createElement("span");
    replaceLabel.textContent = this.messages.replaceLabel;
    replaceField.append(replaceLabel, this.findReplaceValueInput);

    const optionsRow = document.createElement("div");
    optionsRow.className = "excelsior-find-replace-options";
    optionsRow.append(
      createFindReplaceToggle("scope-workbook", this.messages.scopeWorkbook),
      createFindReplaceToggle("case-sensitive", this.messages.caseSensitive),
      createFindReplaceToggle("whole-cell", this.messages.wholeCell),
      createFindReplaceToggle("regex", this.messages.regex)
    );

    const actionsRow = document.createElement("div");
    actionsRow.className = "excelsior-find-replace-actions";
    actionsRow.append(
      createFindReplaceActionButton("previous", this.messages.previousMatch),
      createFindReplaceActionButton("next", this.messages.nextMatch),
      createFindReplaceActionButton("replace", this.messages.replaceOne),
      createFindReplaceActionButton("replace-all", this.messages.replaceAll),
      createFindReplaceActionButton("close", this.messages.closePanel)
    );

    this.findReplacePanel.append(queryField, replaceField, optionsRow, actionsRow, this.findReplaceResults);

    this.pivotPanel.hidden = true;
    this.pivotRowSelect.className = "excelsior-find-replace-input";
    this.pivotRowSelect.dataset.pivotRole = "row";
    this.pivotRowSelect.multiple = true;
    this.pivotRowSelect.size = 4;
    this.pivotColumnSelect.className = "excelsior-find-replace-input";
    this.pivotColumnSelect.dataset.pivotRole = "column";
    this.pivotColumnSelect.multiple = true;
    this.pivotColumnSelect.size = 4;
    this.pivotExecutionModeSelect.className = "excelsior-find-replace-input";
    this.pivotExecutionModeSelect.dataset.pivotRole = "execution-mode";
    this.pivotExecutionModeSelect.replaceChildren(
      ...([
        ["auto", this.messages.pivotExecutionAuto],
        ["client", this.messages.pivotExecutionClient],
        ["server", this.messages.pivotExecutionServer]
      ] as const).map(([value, label]) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = label;
        return option;
      })
    );
    this.pivotValuesField.className = "excelsior-find-replace-field excelsior-pivot-values-field";
    this.pivotValuesContainer.className = "excelsior-pivot-values";
    this.pivotIncludeRowTotals.type = "checkbox";
    this.pivotIncludeRowTotals.dataset.pivotRole = "include-row-totals";
    this.pivotIncludeColumnTotals.type = "checkbox";
    this.pivotIncludeColumnTotals.dataset.pivotRole = "include-column-totals";
    this.pivotIncludeSubtotals.type = "checkbox";
    this.pivotIncludeSubtotals.dataset.pivotRole = "include-subtotals";
    this.pivotAutoRefresh.type = "checkbox";
    this.pivotAutoRefresh.dataset.pivotRole = "auto-refresh";

    const pivotRowField = document.createElement("label");
    pivotRowField.className = "excelsior-find-replace-field";
    const pivotRowLabel = document.createElement("span");
    pivotRowLabel.textContent = this.messages.pivotRow;
    pivotRowField.append(pivotRowLabel, this.pivotRowSelect);

    const pivotColumnField = document.createElement("label");
    pivotColumnField.className = "excelsior-find-replace-field";
    const pivotColumnLabel = document.createElement("span");
    pivotColumnLabel.textContent = this.messages.pivotColumnField;
    pivotColumnField.append(pivotColumnLabel, this.pivotColumnSelect);

    const pivotExecutionField = document.createElement("label");
    pivotExecutionField.className = "excelsior-find-replace-field";
    const pivotExecutionLabel = document.createElement("span");
    pivotExecutionLabel.textContent = this.messages.pivotExecutionMode;
    pivotExecutionField.append(pivotExecutionLabel, this.pivotExecutionModeSelect);

    const pivotValuesLabel = document.createElement("span");
    pivotValuesLabel.textContent = this.messages.pivotValue;
    const pivotAddValue = createFindReplaceActionButton("add-value", this.messages.pivotAddValue);
    pivotAddValue.dataset.pivotAction = "add-value";
    delete pivotAddValue.dataset.findAction;
    const pivotValuesToolbar = document.createElement("div");
    pivotValuesToolbar.className = "excelsior-find-replace-actions";
    pivotValuesToolbar.append(pivotAddValue);
    this.pivotValuesField.append(pivotValuesLabel, this.pivotValuesContainer, pivotValuesToolbar);

    const pivotToggles = document.createElement("div");
    pivotToggles.className = "excelsior-find-replace-options excelsior-pivot-options";
    const pivotRowTotals = document.createElement("label");
    pivotRowTotals.className = "excelsior-find-replace-toggle";
    const pivotRowTotalsText = document.createElement("span");
    pivotRowTotalsText.textContent = this.messages.pivotRowTotals;
    pivotRowTotals.append(this.pivotIncludeRowTotals, pivotRowTotalsText);
    const pivotColumnTotals = document.createElement("label");
    pivotColumnTotals.className = "excelsior-find-replace-toggle";
    const pivotColumnTotalsText = document.createElement("span");
    pivotColumnTotalsText.textContent = this.messages.pivotColumnTotals;
    pivotColumnTotals.append(this.pivotIncludeColumnTotals, pivotColumnTotalsText);
    const pivotSubtotals = document.createElement("label");
    pivotSubtotals.className = "excelsior-find-replace-toggle";
    const pivotSubtotalsText = document.createElement("span");
    pivotSubtotalsText.textContent = this.messages.pivotSubtotals;
    pivotSubtotals.append(this.pivotIncludeSubtotals, pivotSubtotalsText);
    const pivotAutoRefresh = document.createElement("label");
    pivotAutoRefresh.className = "excelsior-find-replace-toggle";
    const pivotAutoRefreshText = document.createElement("span");
    pivotAutoRefreshText.textContent = this.messages.pivotAutoRefresh;
    pivotAutoRefresh.append(this.pivotAutoRefresh, pivotAutoRefreshText);
    pivotToggles.append(pivotRowTotals, pivotColumnTotals, pivotSubtotals, pivotAutoRefresh);

    const pivotActions = document.createElement("div");
    pivotActions.className = "excelsior-find-replace-actions";
    this.pivotApplyButton.textContent = this.messages.pivotApply;
    this.pivotApplyButton.dataset.pivotAction = "apply";
    delete this.pivotApplyButton.dataset.findAction;
    this.pivotCloseButton.textContent = this.messages.closePanel;
    this.pivotCloseButton.dataset.pivotAction = "close";
    delete this.pivotCloseButton.dataset.findAction;
    pivotActions.append(this.pivotApplyButton, this.pivotCloseButton);
    this.pivotPanel.append(pivotRowField, pivotColumnField, pivotExecutionField, this.pivotValuesField, pivotToggles, pivotActions);

    this.surface.append(this.cellsLayer, this.editor, this.selectEditor, this.customEditorHost);
    this.formulaBar.append(this.formulaAddress, this.formulaInput, this.statusMessage, this.findReplacePanel, this.pivotPanel);
    this.root.append(this.chrome, this.formulaBar, this.activeCellAnnouncement, this.viewport, this.sheetTabs);
    this.viewport.append(this.surface);
    this.container.replaceChildren(this.root);

    this.bindEvents();
    this.render();
  }

  focus(): void {
    this.viewport.focus();
  }

  dispose(): void {
    this.cancelScheduledRender();
    this.cancelFindReplaceSearch();
    this.resizeObserver?.disconnect();
    for (const unsubscribe of this.unsubscribeCallbacks) {
      unsubscribe();
    }
    this.viewport.removeEventListener("scroll", this.handleScroll);
    this.viewport.removeEventListener("keydown", this.handleKeyDown);
    this.viewport.removeEventListener("copy", this.handleCopy);
    this.viewport.removeEventListener("paste", this.handlePaste);
    this.toolbar.removeEventListener("click", this.handleToolbarClick);
    this.toolbar.removeEventListener("input", this.handleToolbarInput);
    this.toolbar.removeEventListener("keydown", this.handleToolbarKeyDown);
    this.chrome.removeEventListener("click", this.handleColumnHeaderClick);
    this.chrome.removeEventListener("keydown", this.handleColumnHeaderKeyDown);
    this.sheetTabs.removeEventListener("click", this.handleSheetTabClick);
    this.sheetTabs.removeEventListener("keydown", this.handleSheetTabsKeyDown);
    this.formulaInput.removeEventListener("keydown", this.handleFormulaInputKeyDown);
    this.formulaInput.removeEventListener("blur", this.handleFormulaInputBlur);
    this.formulaInput.removeEventListener("compositionstart", this.handleCompositionStart);
    this.formulaInput.removeEventListener("compositionend", this.handleCompositionEnd);
    this.findReplacePanel.removeEventListener("input", this.handleFindReplacePanelInput);
    this.findReplacePanel.removeEventListener("click", this.handleFindReplacePanelClick);
    this.findReplacePanel.removeEventListener("keydown", this.handleFindReplacePanelKeyDown);
    this.findReplacePanel.removeEventListener("compositionstart", this.handleCompositionStart);
    this.findReplacePanel.removeEventListener("compositionend", this.handleCompositionEnd);
    this.pivotPanel.removeEventListener("input", this.handlePivotPanelInput);
    this.pivotPanel.removeEventListener("click", this.handlePivotPanelClick);
    this.pivotPanel.removeEventListener("keydown", this.handlePivotPanelKeyDown);
    this.pivotPanel.removeEventListener("compositionstart", this.handleCompositionStart);
    this.pivotPanel.removeEventListener("compositionend", this.handleCompositionEnd);
    this.cellsLayer.removeEventListener("click", this.handleCellClick);
    this.cellsLayer.removeEventListener("mousedown", this.handleAutofillMouseDown);
    this.cellsLayer.removeEventListener("mousemove", this.handleAutofillMouseMove);
    this.cellsLayer.removeEventListener("mouseup", this.handleAutofillMouseUp);
    this.cellsLayer.removeEventListener("dblclick", this.handleCellDoubleClick);
    this.editor.removeEventListener("blur", this.handleEditorBlur);
    this.editor.removeEventListener("keydown", this.handleEditorKeyDown);
    this.editor.removeEventListener("compositionstart", this.handleCompositionStart);
    this.editor.removeEventListener("compositionend", this.handleCompositionEnd);
    this.selectEditor.removeEventListener("blur", this.commitSelectEditor);
    this.selectEditor.removeEventListener("change", this.commitSelectEditor);
    this.selectEditor.removeEventListener("keydown", this.handleSelectEditorKeyDown);
    this.customEditorHost.removeEventListener("keydown", this.handleCustomEditorKeyDown);
    this.customEditorHost.removeEventListener("compositionstart", this.handleCompositionStart);
    this.customEditorHost.removeEventListener("compositionend", this.handleCompositionEnd);
    globalThis.removeEventListener("mousemove", this.handleAutofillMouseMove);
    globalThis.removeEventListener("mouseup", this.handleAutofillMouseUp);
    this.destroyCustomEditor();
    this.container.replaceChildren();
  }

  private bindEvents(): void {
    this.unsubscribeCallbacks.push(
      this.engine.on("command:completed", ({ operations, commandType, sheetId }) => {
        if (commandType !== "ApplyOperations") {
          this.options.onChange?.(operations);
        }
        if (commandType !== "SelectRangeCommand") {
          const targetSheetId = sheetId ?? this.engine.getSnapshot().activeSheetId;
          this.engine.getRowModel(targetSheetId).refresh("sheetChanged");
          this.clearRowModelWindowCache();
        }
        if (this.findReplaceState.open && this.findReplaceState.query && commandType !== "SelectRangeCommand") {
          this.scheduleFindReplaceSearch();
        }
        this.requestRender();
      }),
      this.engine.on("selection:changed", () => {
        this.requestRender();
      }),
      this.engine.on("row-model:changed", () => {
        this.clearRowModelWindowCache();
        this.requestRender();
      })
    );

    this.viewport.addEventListener("scroll", this.handleScroll);
    this.viewport.addEventListener("keydown", this.handleKeyDown);
    this.viewport.addEventListener("copy", this.handleCopy);
    this.viewport.addEventListener("paste", this.handlePaste);
    this.toolbar.addEventListener("click", this.handleToolbarClick);
    this.toolbar.addEventListener("input", this.handleToolbarInput);
    this.toolbar.addEventListener("keydown", this.handleToolbarKeyDown);
    this.chrome.addEventListener("click", this.handleColumnHeaderClick);
    this.chrome.addEventListener("keydown", this.handleColumnHeaderKeyDown);
    this.sheetTabs.addEventListener("click", this.handleSheetTabClick);
    this.sheetTabs.addEventListener("keydown", this.handleSheetTabsKeyDown);
    this.formulaInput.addEventListener("keydown", this.handleFormulaInputKeyDown);
    this.formulaInput.addEventListener("blur", this.handleFormulaInputBlur);
    this.formulaInput.addEventListener("compositionstart", this.handleCompositionStart);
    this.formulaInput.addEventListener("compositionend", this.handleCompositionEnd);
    this.findReplacePanel.addEventListener("input", this.handleFindReplacePanelInput);
    this.findReplacePanel.addEventListener("click", this.handleFindReplacePanelClick);
    this.findReplacePanel.addEventListener("keydown", this.handleFindReplacePanelKeyDown);
    this.findReplacePanel.addEventListener("compositionstart", this.handleCompositionStart);
    this.findReplacePanel.addEventListener("compositionend", this.handleCompositionEnd);
    this.pivotPanel.addEventListener("input", this.handlePivotPanelInput);
    this.pivotPanel.addEventListener("click", this.handlePivotPanelClick);
    this.pivotPanel.addEventListener("keydown", this.handlePivotPanelKeyDown);
    this.pivotPanel.addEventListener("compositionstart", this.handleCompositionStart);
    this.pivotPanel.addEventListener("compositionend", this.handleCompositionEnd);
    this.cellsLayer.addEventListener("click", this.handleCellClick);
    this.cellsLayer.addEventListener("mousedown", this.handleAutofillMouseDown);
    this.cellsLayer.addEventListener("mousemove", this.handleAutofillMouseMove);
    this.cellsLayer.addEventListener("mouseup", this.handleAutofillMouseUp);
    this.cellsLayer.addEventListener("dblclick", this.handleCellDoubleClick);
    this.editor.addEventListener("blur", this.handleEditorBlur);
    this.editor.addEventListener("keydown", this.handleEditorKeyDown);
    this.editor.addEventListener("compositionstart", this.handleCompositionStart);
    this.editor.addEventListener("compositionend", this.handleCompositionEnd);
    this.selectEditor.addEventListener("blur", this.commitSelectEditor);
    this.selectEditor.addEventListener("change", this.commitSelectEditor);
    this.selectEditor.addEventListener("keydown", this.handleSelectEditorKeyDown);
    this.customEditorHost.addEventListener("keydown", this.handleCustomEditorKeyDown);
    this.customEditorHost.addEventListener("compositionstart", this.handleCompositionStart);
    this.customEditorHost.addEventListener("compositionend", this.handleCompositionEnd);
    globalThis.addEventListener("mousemove", this.handleAutofillMouseMove);
    globalThis.addEventListener("mouseup", this.handleAutofillMouseUp);

    if (typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => this.requestRender());
      this.resizeObserver.observe(this.viewport);
    } else {
      window.addEventListener("resize", this.handleWindowResize);
      this.unsubscribeCallbacks.push(() => window.removeEventListener("resize", this.handleWindowResize));
    }
  }

  private readonly handleScroll = (): void => {
    this.requestRender();
  };

  private readonly handleWindowResize = (): void => {
    this.requestRender();
  };

  private cancelScheduledRender(): void {
    cancelScheduledFrame(this.renderHandle);
    this.renderHandle = undefined;
  }

  private requestRender(): void {
    const debounceMs = this.getRenderDebounceMs();
    if (debounceMs <= 0) {
      this.cancelScheduledRender();
      this.render();
      return;
    }

    if (this.renderHandle != null) {
      return;
    }

    this.renderHandle = globalThis.setTimeout(() => {
      this.renderHandle = undefined;
      this.render();
    }, debounceMs);
  }

  private readonly handleCellClick = (event: Event): void => {
    const groupToggle = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-remote-group-toggle]");
    if (groupToggle) {
      const pathText = groupToggle.dataset.remoteGroupPath;
      const expanded = groupToggle.dataset.remoteGroupExpanded === "true";
      const path = pathText ? JSON.parse(pathText) : [];
      if (Array.isArray(path) && path.every((segment) => typeof segment === "string") && path.length > 0) {
        const sheet = this.engine.getActiveSheet();
        this.engine.setRemoteGroupExpanded(sheet.id, path, !expanded);
      }
      this.focus();
      return;
    }

    const target = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-row][data-col]");
    if (!target) {
      return;
    }

    const row = Number(target.dataset.row);
    const col = Number(target.dataset.col);
    const sheet = this.engine.getActiveSheet();
    this.selectResolvedCell(sheet, row, col);
    this.focus();
  };

  private readonly handleCellDoubleClick = (event: Event): void => {
    if ((event.target as HTMLElement | null)?.closest("[data-remote-group-toggle]")) {
      return;
    }

    const target = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-row][data-col]");
    if (!target) {
      return;
    }
    this.startEditing(Number(target.dataset.row), Number(target.dataset.col));
  };

  private readonly handleColumnHeaderClick = (event: Event): void => {
    const target = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-column-header-col]");
    if (!target) {
      return;
    }

    const col = Number(target.dataset.columnHeaderCol);
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    this.selectResolvedCell(sheet, activeAddress.row, col);
    if (this.getRemoteRequestModel(sheet.id) !== undefined) {
      this.cycleRemoteSortForColumn(sheet.id, col);
    }
    this.focus();
  };

  private readonly handleColumnHeaderKeyDown = (event: KeyboardEvent): void => {
    const target = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-column-header-col]");
    if (!target) {
      return;
    }

    if (event.key !== "Enter" && event.key !== " ") {
      return;
    }

    event.preventDefault();
    this.handleColumnHeaderClick(event);
  };

  private readonly handleCompositionStart = (event: CompositionEvent): void => {
    if (event.target) {
      this.composingTargets.add(event.target);
    }
  };

  private readonly handleCompositionEnd = (event: CompositionEvent): void => {
    if (event.target) {
      this.composingTargets.delete(event.target);
    }
  };

  private readonly handleKeyDown = (event: KeyboardEvent): void => {
    if (this.editingCell) {
      return;
    }

    if (this.isCompositionInProgress(event.target, event)) {
      return;
    }

    if (this.matchesShortcutList(event, this.shortcuts.openFindReplace)) {
      event.preventDefault();
      this.openFindReplacePanel();
      return;
    }

    if (this.matchesShortcutList(event, this.shortcuts.findNext)) {
      event.preventDefault();
      this.navigateFindReplaceResults(1);
      return;
    }

    if (this.matchesShortcutList(event, this.shortcuts.findPrevious)) {
      event.preventDefault();
      this.navigateFindReplaceResults(-1);
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const selection = this.getResolvedSelectionRange(sheet);
    let nextRow = selection.end.row;
    let nextCol = selection.end.col;

    switch (event.key) {
      case "ArrowUp":
        nextRow = this.findNextVisibleIndex("row", sheet, nextRow - 1, -1);
        break;
      case "ArrowDown":
        nextRow = this.findNextVisibleIndex("row", sheet, nextRow + 1, 1);
        break;
      case "ArrowLeft":
        nextCol = this.findNextVisibleIndex("column", sheet, nextCol + this.getHorizontalNavigationDirection("ArrowLeft"), this.getHorizontalNavigationDirection("ArrowLeft"));
        break;
      case "ArrowRight":
        nextCol = this.findNextVisibleIndex("column", sheet, nextCol + this.getHorizontalNavigationDirection("ArrowRight"), this.getHorizontalNavigationDirection("ArrowRight"));
        break;
      case "Enter":
      case "F2":
        event.preventDefault();
        this.startEditing(selection.end.row, selection.end.col);
        return;
      default:
        return;
    }

    event.preventDefault();
    this.selectResolvedCell(
      sheet,
      Math.min(Math.max(nextRow, 0), sheet.rowCount - 1),
      Math.min(Math.max(nextCol, 0), sheet.columnCount - 1)
    );
  };

  private readonly handleEditorBlur = (): void => {
    if (this.isCompositionInProgress(this.editor)) {
      return;
    }

    this.commitEditor();
  };

  private readonly handleCopy = (event: ClipboardEvent): void => {
    const clipboard = event.clipboardData;
    if (!clipboard) {
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const selection = sheet.selection;
    const rows: string[] = [];
    const includeHidden = this.shouldIncludeHiddenInClipboard();

    for (let row = selection.start.row; row <= selection.end.row; row += 1) {
      if (!includeHidden && this.isRowHidden(sheet, row)) {
        continue;
      }

      const cells: string[] = [];
      for (let col = selection.start.col; col <= selection.end.col; col += 1) {
        if (!includeHidden && this.isColumnHidden(sheet, col)) {
          continue;
        }
        cells.push(this.engine.getDisplayValue(sheet.id, row, col));
      }
      rows.push(cells.join("\t"));
    }

    clipboard.setData("text/plain", rows.join("\n"));
    event.preventDefault();
  };

  private readonly handlePaste = (event: ClipboardEvent): void => {
    const clipboard = event.clipboardData;
    if (!clipboard) {
      return;
    }

    const workbook = this.engine.getSnapshot();
    const sheet = this.engine.getActiveSheet();
    const analysis = resolveClipboardText(
      clipboard.getData("text/plain"),
      clipboard.getData("text/html"),
      workbook.settings.clipboardPolicy,
      workbook.settings.maxPasteCells
    );

    if (analysis.hardBlocked) {
      const reason = analysis.reasons.some((item) => item.startsWith("max-paste-cells-exceeded:"))
        ? "clipboard-paste-blocked"
        : "clipboard-html-blocked";
      this.engine.reportSecurityEvent(reason, { reasons: analysis.reasons });
      event.preventDefault();
      return;
    }

    if (analysis.blocked) {
      this.engine.reportSecurityEvent("clipboard-html-sanitized", { reasons: analysis.reasons });
    }

    const rows = parseTabularText(analysis.text);
    const selection = sheet.selection;
    const includeHidden = this.shouldIncludeHiddenInClipboard();

    for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
      const rowValues = rows[rowOffset] ?? [];
      for (let colOffset = 0; colOffset < rowValues.length; colOffset += 1) {
        const value = rowValues[colOffset];
        const row = selection.start.row + rowOffset;
        const col = selection.start.col + colOffset;

        if (row < sheet.rowCount && col < sheet.columnCount) {
          if (!includeHidden && (this.isRowHidden(sheet, row) || this.isColumnHidden(sheet, col))) {
            continue;
          }
          const validation = this.engine.validateCellValue({ sheetId: sheet.id, row, col, value });
          if (!validation.valid) {
            this.validationFeedback = {
              sheetId: sheet.id,
              row,
              col,
              message: validation.error?.message ?? this.messages.invalidCellValue
            };
            this.render();
            event.preventDefault();
            return;
          }
        }
      }
    }

    for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
      const rowValues = rows[rowOffset] ?? [];
      for (let colOffset = 0; colOffset < rowValues.length; colOffset += 1) {
        const value = rowValues[colOffset];
        const row = selection.start.row + rowOffset;
        const col = selection.start.col + colOffset;

        if (row < sheet.rowCount && col < sheet.columnCount) {
          if (!includeHidden && (this.isRowHidden(sheet, row) || this.isColumnHidden(sheet, col))) {
            continue;
          }
          this.engine.setCellValue({ sheetId: sheet.id, row, col, value });
        }
      }
    }

    this.clearValidationFeedback();

    event.preventDefault();
  };

  private readonly handleEditorKeyDown = (event: KeyboardEvent): void => {
    if (this.isCompositionInProgress(event.target, event)) {
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      this.commitEditor();
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      this.cancelEditing();
    }
  };

  private readonly handleSelectEditorKeyDown = (event: KeyboardEvent): void => {
    if (event.key === "Enter") {
      event.preventDefault();
      this.commitSelectEditor();
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      this.cancelEditing();
    }
  };

  private readonly handleCustomEditorKeyDown = (event: KeyboardEvent): void => {
    if (!this.editingCell || this.editingCell.mode !== "custom") {
      return;
    }

    if (this.isCompositionInProgress(event.target, event)) {
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      this.commitCustomEditor();
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      this.cancelEditing();
    }
  };

  private readonly handleFindReplacePanelInput = (): void => {
    this.scheduleFindReplaceSearch();
  };

  private readonly handleFindReplacePanelClick = (event: Event): void => {
    const button = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-find-action]");
    if (!button) {
      return;
    }

    switch (button.dataset.findAction) {
      case "previous":
        this.navigateFindReplaceResults(-1);
        break;
      case "next":
        this.navigateFindReplaceResults(1);
        break;
      case "replace":
        this.replaceActiveFindReplaceMatch();
        break;
      case "replace-all":
        this.replaceAllFindReplaceMatches();
        break;
      case "close":
        this.closeFindReplacePanel();
        break;
      default:
        break;
    }
  };

  private readonly handleFindReplacePanelKeyDown = (event: KeyboardEvent): void => {
    if (this.isCompositionInProgress(event.target, event)) {
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      this.closeFindReplacePanel();
      return;
    }

    if (event.key === "Enter" && event.target === this.findReplaceQueryInput) {
      event.preventDefault();
      this.navigateFindReplaceResults(event.shiftKey ? -1 : 1);
      return;
    }

    if (event.key === "Enter" && event.target === this.findReplaceValueInput) {
      event.preventDefault();
      this.replaceActiveFindReplaceMatch();
    }
  };

  private readonly handlePivotPanelInput = (): void => {
    this.pivotPanelState.rowFields = this.getPivotPanelSelection(this.pivotRowSelect);
    this.pivotPanelState.columnFields = this.getPivotPanelSelection(this.pivotColumnSelect);

    const valueFieldElements = Array.from(
      this.pivotPanel.querySelectorAll<HTMLSelectElement>("[data-pivot-role='value-field']")
    );
    const valueAggregateElements = Array.from(
      this.pivotPanel.querySelectorAll<HTMLSelectElement>("[data-pivot-role='value-aggregate']")
    );
    const valueAliasElements = Array.from(
      this.pivotPanel.querySelectorAll<HTMLInputElement>("[data-pivot-role='value-alias']")
    );

    this.pivotPanelState.values = valueFieldElements.map((fieldElement, index) => ({
      field: fieldElement.value,
      aggregate: (valueAggregateElements[index]?.value as PivotAggregateFunction | undefined) ?? "sum",
      as: valueAliasElements[index]?.value ?? ""
    }));
    this.pivotPanelState.includeRowTotals = this.pivotIncludeRowTotals.checked;
    this.pivotPanelState.includeColumnTotals = this.pivotIncludeColumnTotals.checked;
    this.pivotPanelState.includeSubtotals = this.pivotIncludeSubtotals.checked;
    this.pivotPanelState.executionMode = (this.pivotExecutionModeSelect.value as PivotExecutionMode | undefined) ?? "auto";
    this.pivotPanelState.autoRefresh = this.pivotAutoRefresh.checked;
  };

  private readonly handlePivotPanelClick = (event: Event): void => {
    const button = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-pivot-action]");
    if (!button) {
      return;
    }

    switch (button.dataset.pivotAction) {
      case "add-value":
        this.pivotPanelState.values.push(this.createPivotPanelValueState(this.pivotPanelState.fields));
        this.render();
        this.pivotPanel.querySelector<HTMLElement>(`[data-pivot-role='value-field'][data-pivot-value-index='${this.pivotPanelState.values.length - 1}']`)?.focus();
        break;
      case "remove-value": {
        const index = Number(button.dataset.pivotValueIndex);
        if (Number.isFinite(index) && this.pivotPanelState.values.length > 1) {
          this.pivotPanelState.values.splice(index, 1);
          this.render();
        }
        break;
      }
      case "apply":
        void this.applyPivotPanel();
        break;
      case "close":
        this.closePivotPanel();
        break;
      default:
        break;
    }
  };

  private readonly handlePivotPanelKeyDown = (event: KeyboardEvent): void => {
    if (this.isCompositionInProgress(event.target, event)) {
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      this.closePivotPanel();
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      void this.applyPivotPanel();
    }
  };

  private readonly handleToolbarClick = (event: Event): void => {
    const button = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-action]");
    if (!button) {
      return;
    }

    const action = button.dataset.action;
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    const activeCell = this.engine.getCell(sheet.id, activeAddress.row, activeAddress.col);
    const selection = sheet.selection.end;

    switch (action) {
      case "undo":
        this.engine.undo();
        break;
      case "redo":
        this.engine.redo();
        break;
      case "insert-row":
        this.engine.insertRows(sheet.id, selection.row, 1);
        break;
      case "delete-row":
        this.engine.deleteRows(sheet.id, selection.row, selection.row);
        break;
      case "insert-column":
        this.engine.insertColumns(sheet.id, selection.col, 1);
        break;
      case "delete-column":
        this.engine.deleteColumns(sheet.id, selection.col, selection.col);
        break;
      case "add-sheet":
        this.engine.addSheet();
        break;
      case "create-pivot":
        this.createPivotFromActiveSheet();
        return;
      case "find-replace":
        this.openFindReplacePanel();
        return;
      case "sort-asc":
        this.applyRemoteSortForActiveColumn("asc");
        this.focus();
        return;
      case "sort-desc":
        this.applyRemoteSortForActiveColumn("desc");
        this.focus();
        return;
      case "group-column":
        this.toggleRemoteGroupingForActiveColumn();
        this.focus();
        return;
      case "pivot-column":
        this.toggleRemotePivotForActiveColumn();
        this.focus();
        return;
      case "aggregate-sum":
        this.toggleRemoteAggregateForActiveColumn("sum");
        this.focus();
        return;
      case "aggregate-avg":
        this.toggleRemoteAggregateForActiveColumn("avg");
        this.focus();
        return;
      case "aggregate-min":
        this.toggleRemoteAggregateForActiveColumn("min");
        this.focus();
        return;
      case "aggregate-max":
        this.toggleRemoteAggregateForActiveColumn("max");
        this.focus();
        return;
      case "aggregate-count":
        this.toggleRemoteAggregateForActiveColumn("count");
        this.focus();
        return;
      case "clear-column-query":
        this.clearRemoteQueryForActiveColumn();
        this.focus();
        return;
      case "bold":
        this.applyStyleToSelection({
          fontWeight: activeCell?.style?.fontWeight === "bold" ? "normal" : "bold"
        });
        break;
      case "italic":
        this.applyStyleToSelection({
          fontStyle: activeCell?.style?.fontStyle === "italic" ? "normal" : "italic"
        });
        break;
      case "wrap":
        this.applyStyleToSelection({
          wrap: !activeCell?.style?.wrap
        });
        break;
      case "align-left":
        this.applyStyleToSelection({ align: "left" });
        break;
      case "align-center":
        this.applyStyleToSelection({ align: "center" });
        break;
      case "align-right":
        this.applyStyleToSelection({ align: "right" });
        break;
      case "merge":
        this.engine.mergeCells({
          sheetId: sheet.id,
          start: sheet.selection.start,
          end: sheet.selection.end
        });
        break;
      case "unmerge":
        this.engine.unmergeCells({
          sheetId: sheet.id,
          row: activeAddress.row,
          col: activeAddress.col
        });
        break;
      default:
        return;
    }

    this.render();
    this.focus();
  };

  private readonly handleToolbarInput = (event: Event): void => {
    const input = (event.target as HTMLElement | null)?.closest<HTMLInputElement>("[data-remote-filter-input]");
    if (!input) {
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    this.setRemoteFilterDraft(sheet.id, activeAddress.col, input.value);
  };

  private readonly handleToolbarKeyDown = (event: KeyboardEvent): void => {
    const input = (event.target as HTMLElement | null)?.closest<HTMLInputElement>("[data-remote-filter-input]");
    if (!input) {
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      this.applyRemoteFilterForActiveColumn(input.value);
      this.focus();
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      this.remoteFilterDraft = undefined;
      this.render();
      this.focus();
    }
  };

  private readonly handleFormulaInputKeyDown = (event: KeyboardEvent): void => {
    if (this.isCompositionInProgress(event.target, event)) {
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      this.commitFormulaInput();
    }

    if (event.key === "Escape") {
      event.preventDefault();
      this.renderFormulaBar();
      this.focus();
    }
  };

  private readonly handleFormulaInputBlur = (): void => {
    if (this.isCompositionInProgress(this.formulaInput)) {
      return;
    }

    this.commitFormulaInput();
  };

  private readonly handleSheetTabClick = (event: Event): void => {
    const closeButton = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-close-sheet-id]");
    if (closeButton) {
      const sheetId = closeButton.dataset.closeSheetId;
      if (sheetId && this.engine.getSnapshot().sheets.length > 1) {
        this.engine.deleteSheet(sheetId);
        this.render();
      }
      return;
    }

    const tabButton = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-sheet-id]");
    if (!tabButton) {
      return;
    }

    const sheetId = tabButton.dataset.sheetId;
    if (!sheetId) {
      return;
    }

    this.engine.setActiveSheet(sheetId);
    this.render();
    this.focus();
  };

  private readonly handleSheetTabsKeyDown = (event: KeyboardEvent): void => {
    const tabButtons = Array.from(this.sheetTabs.querySelectorAll<HTMLButtonElement>("[data-sheet-id]"));
    if (!tabButtons.length) {
      return;
    }

    const target = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-sheet-id]");
    if (!target) {
      return;
    }

    const currentIndex = tabButtons.indexOf(target);
    if (currentIndex < 0) {
      return;
    }

    const focusAndActivate = (nextIndex: number): void => {
      const nextTab = tabButtons[(nextIndex + tabButtons.length) % tabButtons.length];
      const sheetId = nextTab?.dataset.sheetId;
      if (!nextTab || !sheetId) {
        return;
      }

      this.engine.setActiveSheet(sheetId);
      this.render();
      this.sheetTabs.querySelector<HTMLButtonElement>(`[data-sheet-id='${sheetId}']`)?.focus();
    };

    switch (event.key) {
      case "ArrowLeft":
      case "ArrowUp":
        event.preventDefault();
        focusAndActivate(currentIndex - 1);
        return;
      case "ArrowRight":
      case "ArrowDown":
        event.preventDefault();
        focusAndActivate(currentIndex + 1);
        return;
      case "Home":
        event.preventDefault();
        focusAndActivate(0);
        return;
      case "End":
        event.preventDefault();
        focusAndActivate(tabButtons.length - 1);
        return;
      case "Enter":
      case " ": {
        event.preventDefault();
        const sheetId = target.dataset.sheetId;
        if (sheetId) {
          this.engine.setActiveSheet(sheetId);
          this.render();
          this.sheetTabs.querySelector<HTMLButtonElement>(`[data-sheet-id='${sheetId}']`)?.focus();
        }
        return;
      }
      default:
        return;
    }
  };

  private readonly commitEditor = (): void => {
    if (!this.editingCell || this.editingCell.mode !== "text") {
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.resolveCellAddress(sheet, this.editingCell.row, this.editingCell.col).address;
    if (
      !this.commitCellValue(sheet.id, activeAddress.row, activeAddress.col, this.editor.value, () => {
        this.editor.focus();
        this.editor.select();
      })
    ) {
      return;
    }

    this.editingCell = undefined;
    this.destroyCustomEditor();
    this.hideInlineEditors();
    this.focus();
  };

  private readonly commitSelectEditor = (): void => {
    if (!this.editingCell || this.editingCell.mode !== "select") {
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.resolveCellAddress(sheet, this.editingCell.row, this.editingCell.col).address;
    const validationRule = this.getInteractiveValidationRule(sheet.id, activeAddress.row, activeAddress.col);
    if (!validationRule || !isListValidationRule(validationRule)) {
      this.cancelEditing();
      return;
    }

    const nextValue = validationRule.values[this.selectEditor.selectedIndex] ?? this.selectEditor.value;
    if (!this.commitCellValue(sheet.id, activeAddress.row, activeAddress.col, nextValue, () => this.selectEditor.focus())) {
      return;
    }

    this.editingCell = undefined;
    this.destroyCustomEditor();
    this.hideInlineEditors();
    this.focus();
  };

  private readonly commitCustomEditor = (): void => {
    if (!this.editingCell || this.editingCell.mode !== "custom" || !this.activeCustomEditor) {
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.resolveCellAddress(sheet, this.editingCell.row, this.editingCell.col).address;
    if (
      !this.commitCellValue(
        sheet.id,
        activeAddress.row,
        activeAddress.col,
        this.activeCustomEditor.getValue(),
        () => this.activeCustomEditor?.focus?.()
      )
    ) {
      return;
    }

    this.editingCell = undefined;
    this.destroyCustomEditor();
    this.hideInlineEditors();
    this.focus();
  };

  private readonly commitFormulaInput = (): void => {
    const sheet = this.engine.getActiveSheet();
    const selection = this.getActiveAddress(sheet);
    try {
      this.engine.setCellValue({
        sheetId: sheet.id,
        row: selection.row,
        col: selection.col,
        value: this.formulaInput.value
      });
      this.clearValidationFeedback();
    } catch (error) {
      if (error instanceof CellValidationError) {
        this.validationFeedback = {
          sheetId: error.sheetId,
          row: error.address.row,
          col: error.address.col,
          message: error.details.message
        };
        this.render();
        this.formulaInput.focus();
        this.formulaInput.select();
        return;
      }

      throw error;
    }

    this.renderFormulaBar();
  };

  private cancelEditing(): void {
    this.editingCell = undefined;
    this.destroyCustomEditor();
    this.hideInlineEditors();
    this.focus();
  }

  private startEditing(row: number, col: number): void {
    const sheet = this.engine.getActiveSheet();
    const { address } = this.resolveCellAddress(sheet, row, col);
    const cell = this.engine.getCell(sheet.id, address.row, address.col);
    this.selectResolvedCell(sheet, row, col);
    const customEditorContext = this.createEditorContext(sheet.id, address.row, address.col);
    const customEditor = this.getCustomEditor(customEditorContext);

    if (customEditor) {
      try {
        this.hideInlineEditors();
        this.destroyCustomEditor();
        this.editingCell = { ...address, mode: "custom" };
        this.activeCustomEditor = customEditor.create(customEditorContext);
        this.customEditorHost.hidden = false;
        this.positionEditor(address.row, address.col);
        this.activeCustomEditor.mount(this.customEditorHost);
        this.activeCustomEditor.focus?.();
        return;
      } catch {
        this.destroyCustomEditor();
        this.hideInlineEditors();
      }
    }

    const validationRule = this.getInteractiveValidationRule(sheet.id, address.row, address.col);

    if (validationRule && isCheckboxValidationRule(validationRule)) {
      const currentValue = this.getCellPrimitiveValue(sheet.id, address.row, address.col);
      this.commitCellValue(sheet.id, address.row, address.col, !toBoolean(currentValue), () => this.focus());
      this.focus();
      return;
    }

    this.hideInlineEditors();

    if (validationRule && isListValidationRule(validationRule)) {
      const currentValue = this.getCellPrimitiveValue(sheet.id, address.row, address.col);
      this.editingCell = { ...address, mode: "select" };
      this.selectEditor.replaceChildren();

      for (const candidate of validationRule.values) {
        const option = document.createElement("option");
        option.textContent = String(candidate);
        this.selectEditor.append(option);
      }

      const selectedIndex = validationRule.values.findIndex(
        (candidate: CellPrimitive) => candidate === currentValue || String(candidate) === String(currentValue ?? "")
      );
      this.selectEditor.selectedIndex = selectedIndex >= 0 ? selectedIndex : 0;
      this.selectEditor.hidden = false;
      this.positionEditor(address.row, address.col);
      this.selectEditor.focus();
      return;
    }

    this.editingCell = { ...address, mode: "text" };

    this.editor.hidden = false;
    this.editor.value = cell?.formula ?? (cell?.value == null ? "" : String(cell.value));
    this.positionEditor(address.row, address.col);
    this.editor.focus();
    this.editor.select();
  }

  private positionEditor(row: number, col: number): void {
    if (this.editingCell?.mode === "custom") {
      this.positionEditorElement(this.customEditorHost, row, col);
      return;
    }

    if (this.editingCell?.mode === "select") {
      this.positionEditorElement(this.selectEditor, row, col);
      return;
    }

    this.positionEditorElement(this.editor, row, col);
  }

  private renderFormulaBar(): void {
    const sheet = this.engine.getActiveSheet();
    const selection = this.getActiveAddress(sheet);
    const label = cellAddressToLabel(selection);
    const cell = this.engine.getCell(sheet.id, selection.row, selection.col);
    const value = cell?.formula ?? this.engine.getDisplayValue(sheet.id, selection.row, selection.col);
    const validationMessage =
      this.validationFeedback && this.validationFeedback.sheetId === sheet.id ? this.validationFeedback.message : undefined;
    const rowModelError = this.rowModelFeedback?.sheetId === sheet.id ? this.rowModelFeedback.error : undefined;
    const derivedPivotStatus = this.getPivotStatus(sheet.id);
    const pivotStatus = this.pivotFeedback?.sheetId === sheet.id ? this.pivotFeedback.message : derivedPivotStatus.message;
    const pivotStatusIsError = this.pivotFeedback?.sheetId === sheet.id ? this.pivotFeedback.isError : derivedPivotStatus.isError;
    this.formulaAddress.textContent = label;
    this.statusMessage.textContent =
      validationMessage ?? rowModelError ?? pivotStatus ?? (this.hasPendingRowModelRequests(sheet.id) ? this.messages.loadingRows : "");
    this.statusMessage.classList.toggle("is-error", Boolean(validationMessage ?? rowModelError) || pivotStatusIsError);
    if (document.activeElement !== this.formulaInput) {
      this.formulaInput.value = value;
    }
    this.renderFindReplacePanel();
    this.renderPivotPanel();
  }

  private clearValidationFeedback(): void {
    this.validationFeedback = undefined;
  }

  private renderChrome(): void {
    const sheet = this.engine.getActiveSheet();
    this.chrome.replaceChildren();
    const fragment = document.createDocumentFragment();
    const corner = document.createElement("div");
    corner.className = "excelsior-corner";
    corner.textContent = sheet.name;
    corner.setAttribute("aria-hidden", "true");
    fragment.append(corner);

    const columnStrip = document.createElement("div");
    columnStrip.className = "excelsior-column-strip";
    const remoteRequestModel = this.getRemoteRequestModel(sheet.id);
    for (let col = 0; col < Math.min(sheet.columnCount, 12); col += 1) {
      const header = document.createElement("div");
      header.className = "excelsior-column-header";
      header.id = this.getColumnHeaderElementId(sheet.id, col);
      header.setAttribute("role", "columnheader");
      header.setAttribute("aria-colindex", String(col + 1));
      header.setAttribute("aria-label", this.getColumnHeaderAccessibilityLabel(col));
      header.dataset.columnHeaderCol = String(col);
      header.tabIndex = 0;
      if (remoteRequestModel !== undefined) {
        const direction = this.getActiveRemoteSortDirection(sheet.id, col);
        header.setAttribute(
          "aria-sort",
          direction === "asc" ? "ascending" : direction === "desc" ? "descending" : "none"
        );
      }
      header.style.width = `${this.getColumnWidth(sheet, col)}px`;
      header.textContent = columnIndexToLabel(col);
      columnStrip.append(header);
    }
    fragment.append(this.renderToolbar(), columnStrip);
    this.chrome.append(fragment);
  }

  private renderToolbar(): HTMLElement {
    this.toolbar.replaceChildren();
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    const activeCell = this.engine.getCell(sheet.id, activeAddress.row, activeAddress.col);
    const remoteRequestModel = this.getRemoteRequestModel(sheet.id);
    const pivotSourceRange = this.getPivotSourceRange(sheet);
    const activeRemoteSort = this.getActiveRemoteSortDirection(sheet.id, activeAddress.col);
    const toggleStates: Partial<Record<string, boolean>> = {
      bold: activeCell?.style?.fontWeight === "bold",
      italic: activeCell?.style?.fontStyle === "italic",
      wrap: activeCell?.style?.wrap === true,
      "group-column": this.isActiveRemoteGrouped(sheet.id, activeAddress.col),
      "pivot-column": this.isActiveRemotePivoted(sheet.id, activeAddress.col),
      "aggregate-sum": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "sum"),
      "aggregate-avg": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "avg"),
      "aggregate-min": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "min"),
      "aggregate-max": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "max"),
      "aggregate-count": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "count"),
      "sort-asc": activeRemoteSort === "asc",
      "sort-desc": activeRemoteSort === "desc"
    };
    const actions: Array<{ action: string; label: string }> = [
      { action: "undo", label: this.messages.undo },
      { action: "redo", label: this.messages.redo },
      { action: "sort-asc", label: this.messages.sortAscending },
      { action: "sort-desc", label: this.messages.sortDescending },
      { action: "group-column", label: this.messages.groupColumn },
      { action: "pivot-column", label: this.messages.pivotColumn },
      { action: "aggregate-sum", label: this.messages.aggregateSum },
      { action: "aggregate-avg", label: this.messages.aggregateAverage },
      { action: "aggregate-min", label: this.messages.aggregateMin },
      { action: "aggregate-max", label: this.messages.aggregateMax },
      { action: "aggregate-count", label: this.messages.aggregateCount },
      { action: "clear-column-query", label: this.messages.clearColumnQuery },
      { action: "bold", label: this.messages.bold },
      { action: "italic", label: this.messages.italic },
      { action: "wrap", label: this.messages.wrap },
      { action: "align-left", label: this.messages.alignLeft },
      { action: "align-center", label: this.messages.alignCenter },
      { action: "align-right", label: this.messages.alignRight },
      { action: "merge", label: this.messages.merge },
      { action: "unmerge", label: this.messages.unmerge },
      { action: "insert-row", label: this.messages.insertRow },
      { action: "delete-row", label: this.messages.deleteRow },
      { action: "insert-column", label: this.messages.insertColumn },
      { action: "delete-column", label: this.messages.deleteColumn },
      { action: "create-pivot", label: this.messages.createPivot },
      { action: "find-replace", label: this.messages.findReplace },
      { action: "add-sheet", label: this.messages.addSheet }
    ];

    if (remoteRequestModel !== undefined) {
      const filterField = document.createElement("label");
      filterField.className = "excelsior-toolbar-filter";
      const filterLabel = document.createElement("span");
      filterLabel.textContent = `${this.messages.filterColumn} ${this.getRemoteRequestField(activeAddress.col)}`;
      const filterInput = document.createElement("input");
      filterInput.type = "text";
      filterInput.className = "excelsior-toolbar-input";
      filterInput.dataset.remoteFilterInput = "true";
      filterInput.placeholder = this.messages.filterPlaceholder;
      filterInput.value = this.getActiveRemoteFilterValue(sheet.id, activeAddress.col);
      filterInput.setAttribute("aria-label", `${this.messages.filterColumn} ${this.getRemoteRequestField(activeAddress.col)}`);
      filterField.append(filterLabel, filterInput);
      this.toolbar.append(filterField);
    }

    for (const item of actions) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "excelsior-toolbar-button";
      button.dataset.action = item.action;
      button.textContent = item.label;
      if (
        [
          "sort-asc",
          "sort-desc",
          "group-column",
          "pivot-column",
          "aggregate-sum",
          "aggregate-avg",
          "aggregate-min",
          "aggregate-max",
          "aggregate-count",
          "clear-column-query"
        ].includes(item.action) && remoteRequestModel === undefined
      ) {
        button.disabled = true;
      }
      if (item.action === "create-pivot" && pivotSourceRange === undefined) {
        button.disabled = true;
      }
      if (item.action in toggleStates) {
        button.setAttribute("aria-pressed", String(Boolean(toggleStates[item.action])));
      }
      this.toolbar.append(button);
    }

    return this.toolbar;
  }

  private renderSheetTabs(): void {
    const workbook = this.engine.getSnapshot();
    this.sheetTabs.replaceChildren();

    for (const sheet of workbook.sheets) {
      const tabItem = document.createElement("div");
      tabItem.className = "excelsior-sheet-tab-item";
      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "excelsior-sheet-tab";
      tab.id = this.getSheetTabElementId(sheet.id);
      tab.setAttribute("role", "tab");
      if (sheet.id === workbook.activeSheetId) {
        tab.classList.add("is-active");
        tab.setAttribute("aria-selected", "true");
        tab.tabIndex = 0;
      } else {
        tab.setAttribute("aria-selected", "false");
        tab.tabIndex = -1;
      }
      tab.dataset.sheetId = sheet.id;

      const label = document.createElement("span");
      label.textContent = sheet.name;
      tab.append(label);
      tabItem.append(tab);

      if (workbook.sheets.length > 1) {
        const close = document.createElement("button");
        close.type = "button";
        close.className = "excelsior-sheet-tab-close";
        close.dataset.closeSheetId = sheet.id;
        close.setAttribute("aria-label", `${this.messages.sheetTabClose}: ${sheet.name}`);
        close.textContent = "×";
        tabItem.append(close);
      }

      this.sheetTabs.append(tabItem);
    }
  }

  readonly render = (): void => {
    this.cancelScheduledRender();
    const workbook = this.engine.getSnapshot();
    const sheet = this.engine.getActiveSheet();
    const resolvedRowCount = this.getResolvedRowCount(sheet);
    const { viewportBuffer } = workbook.settings;
    const viewportWidth = this.viewport.clientWidth || this.container.clientWidth || 800;
    const viewportHeight = this.viewport.clientHeight || this.container.clientHeight || 480;
    const rowOffsets = buildOffsets(resolvedRowCount, (index) => this.getRowHeight(sheet, index));
    const colOffsets = buildOffsets(sheet.columnCount, (index) => this.getColumnWidth(sheet, index));
    const visibleRows = findVisibleBounds(
      rowOffsets,
      this.viewport.scrollTop,
      this.viewport.scrollTop + viewportHeight,
      viewportBuffer
    );
    const visibleColumns = findVisibleBounds(
      colOffsets,
      this.viewport.scrollLeft,
      this.viewport.scrollLeft + viewportWidth,
      viewportBuffer
    );
    const rowModelRows = this.resolveRowModelRows(sheet.id, visibleRows.start, visibleRows.end);
    const rowModelRowsByIndex = new Map((rowModelRows ?? []).map((rowModelRow) => [rowModelRow.index, rowModelRow]));
    const rowsToRender = this.getRenderedRowIndices(sheet, visibleRows, resolvedRowCount, rowModelRows);
    const columnsToRender = this.getRenderedColumnIndices(sheet, visibleColumns);
    const { anchors, covered } = (() => {
      const anchorMap = new Map<string, SheetMerge>();
      const coveredCells = new Set<string>();
      for (const merge of sheet.merges) {
        anchorMap.set(getCellKey(merge.start.row, merge.start.col), merge);
        for (let row = merge.start.row; row <= merge.end.row; row += 1) {
          for (let col = merge.start.col; col <= merge.end.col; col += 1) {
            if (row === merge.start.row && col === merge.start.col) {
              continue;
            }
            coveredCells.add(getCellKey(row, col));
          }
        }
      }
      return { anchors: anchorMap, covered: coveredCells };
    })();
    const activeAddress = this.getActiveAddress(sheet);
    const selectionRange = this.getResolvedSelectionRange(sheet);
    const autofillPreview = this.autofillDrag?.preview?.fillRange;
    const activeFindMatch = this.getActiveFindReplaceMatch();
    const findMatchKeys = new Set(
      this.findReplaceState.matches
        .filter((match) => match.sheetId === sheet.id)
        .map((match) => getCellKey(match.row, match.col))
    );
    this.viewport.setAttribute("aria-rowcount", String(resolvedRowCount));
    this.viewport.setAttribute("aria-colcount", String(sheet.columnCount));
    this.updateAccessibilityState(
      sheet.id,
      activeAddress.row,
      activeAddress.col,
      isWithinRange(activeAddress.row, activeAddress.col, selectionRange)
    );
    this.renderChrome();
    this.renderFormulaBar();
    this.renderSheetTabs();

    const fragment = document.createDocumentFragment();

    for (const row of rowsToRender) {
      if (this.isRowHidden(sheet, row)) {
        continue;
      }

      for (const col of columnsToRender) {
        if (this.isColumnHidden(sheet, col)) {
          continue;
        }

        const key = getCellKey(row, col);
        if (covered.has(key)) {
          continue;
        }

        const merge = anchors.get(key);
        const cellRange = merge ?? {
          start: { row, col },
          end: { row, col }
        };
        const cell = document.createElement("div");
        cell.className = "excelsior-cell";
        if (merge) {
          cell.classList.add("is-merged");
        }
        if (rangesOverlap(cellRange, sheet.selection)) {
          cell.classList.add("is-selected");
        }
        if (autofillPreview && rangesOverlap(cellRange, autofillPreview)) {
          cell.classList.add("is-autofill-preview");
        }
        if (findMatchKeys.has(getCellKey(cellRange.start.row, cellRange.start.col))) {
          cell.classList.add("is-find-match");
        }
        if (
          activeFindMatch?.sheetId === sheet.id &&
          activeFindMatch.row === cellRange.start.row &&
          activeFindMatch.col === cellRange.start.col
        ) {
          cell.classList.add("is-find-match-active");
        }

        const frozenPane = this.engine.getFrozenPane(sheet.id);
        if (cellRange.start.row < frozenPane.rows || cellRange.start.col < frozenPane.columns) {
          cell.classList.add("is-frozen");
        }
        if (cellRange.start.row < frozenPane.rows) {
          cell.classList.add("is-frozen-row");
        }
        if (cellRange.start.col < frozenPane.columns) {
          cell.classList.add("is-frozen-column");
        }

        const model = this.engine.getCell(sheet.id, cellRange.start.row, cellRange.start.col);
        if (model?.error) {
          cell.classList.add("is-error");
          cell.title = model.error.message;
        }

        if (
          this.validationFeedback &&
          this.validationFeedback.sheetId === sheet.id &&
          this.validationFeedback.row === cellRange.start.row &&
          this.validationFeedback.col === cellRange.start.col
        ) {
          cell.classList.add("is-error");
          cell.title = this.validationFeedback.message;
        }

        cell.dataset.row = String(cellRange.start.row);
        cell.dataset.col = String(cellRange.start.col);
        cell.id = this.getCellElementId(sheet.id, cellRange.start.row, cellRange.start.col);
        cell.setAttribute("role", "gridcell");
        cell.setAttribute("aria-rowindex", String(cellRange.start.row + 1));
        cell.setAttribute("aria-colindex", String(cellRange.start.col + 1));
        cell.setAttribute("aria-selected", String(rangesOverlap(cellRange, sheet.selection)));
        if (cellRange.start.col < Math.min(sheet.columnCount, 12)) {
          cell.setAttribute("aria-describedby", this.getColumnHeaderElementId(sheet.id, cellRange.start.col));
        }
        cell.setAttribute(
          "aria-label",
          this.getCellAccessibilityLabel(
            sheet.id,
            cellRange.start.row,
            cellRange.start.col,
            rangesOverlap(cellRange, sheet.selection)
          )
        );
        if (merge) {
          cell.setAttribute("aria-rowspan", String(cellRange.end.row - cellRange.start.row + 1));
          cell.setAttribute("aria-colspan", String(cellRange.end.col - cellRange.start.col + 1));
        }
        cell.style.top = `${this.getFrozenAdjustedTop(sheet.id, rowOffsets, cellRange.start.row)}px`;
        cell.style.left = `${this.getFrozenAdjustedLeft(sheet.id, colOffsets, cellRange.start.col)}px`;
        cell.style.width = `${getSpanSize(colOffsets, cellRange.start.col, cellRange.end.col)}px`;
        cell.style.height = `${getSpanSize(rowOffsets, cellRange.start.row, cellRange.end.row)}px`;
        this.applyCellPresentation(cell, this.getCellStyle(sheet, cellRange.start.row, cellRange.start.col));
        this.renderCellContent(cell, sheet.id, cellRange.start.row, cellRange.start.col, rowModelRowsByIndex.get(cellRange.start.row));

        if (
          this.isAutofillEnabled() &&
          !this.editingCell &&
          cellRange.end.row === selectionRange.end.row &&
          cellRange.end.col === selectionRange.end.col
        ) {
          const handle = document.createElement("button");
          handle.type = "button";
          handle.className = "excelsior-fill-handle";
          handle.setAttribute("aria-label", this.messages.autofillHandle);
          cell.append(handle);
        }

        fragment.append(cell);
      }
    }

    this.cellsLayer.replaceChildren(fragment);

    if (this.editingCell) {
      this.positionEditor(this.editingCell.row, this.editingCell.col);
    }
  };

  private getVisibleRemoteGroup(rowModelRow: RowModelRow | undefined, col: number): RowModelRow["group"] | undefined {
    if (!rowModelRow?.group) {
      return undefined;
    }

    const targetCol = Math.max(0, rowModelRow.group.level ?? 0);
    return col === targetCol ? rowModelRow.group : undefined;
  }
}