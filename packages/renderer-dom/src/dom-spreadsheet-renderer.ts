import {
  CellValidationError,
  SpreadsheetOperationError,
  WorkbookEngine,
  cellAddressToLabel,
  cellLabelToAddress,
  columnIndexToLabel,
  type CellAddress,
  type CellModel,
  type CellPrimitive,
  type CellRange,
  type CellStyle,
  type CellValidationConfig,
  type CellValidationRule,
  type ClientSideFilterDescriptor,
  type CollaborationPresence,
  type CommentAuthor,
  type PivotAggregateFunction,
  type PivotBuildProgress,
  type PivotExecutionMode,
  type RowModelRow,
  type RowResult,
  type SheetMerge,
  type SheetModel,
  type SpreadsheetOperation,
  type WorksheetImageObject,
  type WorksheetObjectPosition,
  type WorksheetWidgetObject
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
  toolbarDataGroup: string;
  toolbarFontGroup: string;
  toolbarAlignmentGroup: string;
  toolbarStructureGroup: string;
  undo: string;
  redo: string;
  bold: string;
  italic: string;
  underline: string;
  fontFamily: string;
  fontSize: string;
  wrap: string;
  textColor: string;
  borderColor: string;
  fillColor: string;
  formatPainter: string;
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
  cellNote: string;
  cellNoteLabel: string;
  cellNoteSave: string;
  cellNoteRemove: string;
  cellNoteClose: string;
  cellNoteIndicator: string;
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
  widgetRenderers?: Readonly<Record<string, CustomWidgetRenderer>>;
  includeHiddenCellsInClipboard?: boolean;
  autofill?: AutofillOptions;
  localization?: RendererLocalizationOptions;
  renderDebounceMs?: number;
  collaboration?: {
    user: CommentAuthor;
    color?: string;
  };
  comments?: {
    author: CommentAuthor;
  };
}

export interface CustomWidgetRenderContext {
  host: HTMLElement;
  widget: Readonly<WorksheetWidgetObject>;
}

export type CustomWidgetRenderer = (context: CustomWidgetRenderContext) => void | (() => void);

const DEFAULT_PRESENCE_COLOR = "#2563eb";
const SAFE_PRESENCE_COLOR = /^#[0-9a-f]{6}$/i;

const getPresenceColor = (presence: CollaborationPresence): string => {
  const color = presence.metadata?.color;
  return typeof color === "string" && SAFE_PRESENCE_COLOR.test(color) ? color : DEFAULT_PRESENCE_COLOR;
};

const getPresenceName = (presence: CollaborationPresence): string =>
  presence.user?.name?.trim().slice(0, 80) || "Participante remoto";

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
  toolbarDataGroup: "Dados",
  toolbarFontGroup: "Fonte",
  toolbarAlignmentGroup: "Alinhamento",
  toolbarStructureGroup: "Estrutura",
  undo: "Undo",
  redo: "Redo",
  bold: "Bold",
  italic: "Italic",
  underline: "Underline",
  fontFamily: "Fonte",
  fontSize: "Tamanho da fonte",
  wrap: "Wrap",
  textColor: "Text Color",
  borderColor: "Border Color",
  fillColor: "Fill Color",
  formatPainter: "Copiar formatação",
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
  cellNote: "Nota",
  cellNoteLabel: "Nota da célula",
  cellNoteSave: "Salvar",
  cellNoteRemove: "Remover",
  cellNoteClose: "Fechar",
  cellNoteIndicator: "Abrir nota da célula",
  checkboxHint: "Clique duplo para alternar o checkbox.",
  dropdownHint: "Clique duplo para escolher um valor da lista.",
  invalidCellValue: "Valor inválido para a célula selecionada.",
  autofillHandle: "Arrastar para preencher",
};

const DEFAULT_RENDERER_SHORTCUTS: RendererShortcutMap = {
  openFindReplace: ["Ctrl+F", "Meta+F"],
  findNext: ["F3"],
  findPrevious: ["Shift+F3"]
};

const ROW_HEADER_WIDTH = 56;

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

const getSafeRichTextHref = (hyperlink: string | undefined): string | undefined => {
  if (!hyperlink || /[\s\u007f]/.test(hyperlink)) {
    return undefined;
  }

  try {
    const url = new URL(hyperlink);
    return url.protocol === "https:" || (url.protocol === "mailto:" && url.pathname.includes("@")) ? url.href : undefined;
  } catch {
    return undefined;
  }
};

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
type CellValidationDateRule = Extract<CellValidationRule, { type: "date" }>;

type InteractiveValidationRule = CellValidationListRule | CellValidationCheckboxRule | CellValidationDateRule;

const isListValidationRule = (rule: CellValidationRule): rule is CellValidationListRule =>
  rule.type === "list" || rule.type === "dropdown";

const isCheckboxValidationRule = (rule: CellValidationRule): rule is CellValidationCheckboxRule => rule.type === "checkbox";

const isDateValidationRule = (rule: CellValidationRule): rule is CellValidationDateRule => rule.type === "date";

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

interface VisualObjectSurfaceMetrics {
  sheetId: string;
  rowOffsets: number[];
  colOffsets: number[];
  rowCount: number;
  colCount: number;
}

interface VisualObjectRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

interface VisualObjectInteractionState {
  mode: "move" | "resize";
  kind: "image" | "widget";
  sheetId: string;
  objectId: string;
  pointerStartX: number;
  pointerStartY: number;
  originRect: VisualObjectRect;
  liveRect: VisualObjectRect;
}

const VISUAL_OBJECT_MIN_WIDTH = 220;
const VISUAL_OBJECT_MIN_HEIGHT = 150;

const clampNumeric = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));
const toFiniteNumber = (value: unknown, fallback = 0): number =>
  typeof value === "number" && Number.isFinite(value) ? value : fallback;
const cloneSerializable = <T>(value: T): T => {
  if (typeof structuredClone === "function") {
    try {
      return structuredClone(value);
    } catch {
      // Fallback to JSON clone for plain serializable payloads.
    }
  }
  return JSON.parse(JSON.stringify(value)) as T;
};

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

  private readonly gridPanel = document.createElement("div");

  private readonly chrome = document.createElement("div");

  private readonly toolbar = document.createElement("div");

  private readonly imageFileInput = document.createElement("input");

  private readonly formulaBar = document.createElement("div");

  private readonly formulaAddress = document.createElement("span");

  private readonly formulaInput = document.createElement("input");

  private readonly statusMessage = document.createElement("span");

  private readonly findReplacePanel = document.createElement("div");

  private readonly findReplaceQueryInput = document.createElement("input");

  private readonly findReplaceValueInput = document.createElement("input");

  private readonly findReplaceResults = document.createElement("span");

  private readonly notePanel = document.createElement("section");

  private readonly noteInput = document.createElement("textarea");

  private readonly commentList = document.createElement("div");

  private readonly commentInput = document.createElement("textarea");

  private readonly noteSaveButton = createFindReplaceActionButton("save", "");

  private readonly noteRemoveButton = createFindReplaceActionButton("remove", "");

  private readonly noteCloseButton = createFindReplaceActionButton("close", "");

  private readonly pivotPanel = document.createElement("div");

  private readonly pivotApplyButton = createFindReplaceActionButton("apply", "");

  private readonly pivotCloseButton = createFindReplaceActionButton("close", "");

  private readonly toolbarToolPanel = document.createElement("section");

  private readonly toolbarToolTitle = document.createElement("strong");

  private readonly toolbarToolModeField = document.createElement("label");

  private readonly toolbarToolModeLabel = document.createElement("span");

  private readonly toolbarToolModeSelect = document.createElement("select");

  private readonly toolbarToolValueField = document.createElement("label");

  private readonly toolbarToolValueLabel = document.createElement("span");

  private readonly toolbarToolValueInput = document.createElement("input");

  private activeToolbarTool?: "link" | "split" | "validation" | "conditional" | "find-special";

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

  private readonly rowHeaders = document.createElement("div");

  private readonly surface = document.createElement("div");

  private readonly cellsLayer = document.createElement("div");

  private readonly visualObjectsLayer = document.createElement("div");

  private sheetZoom = 1;

  private readonly splitPaneLayer = document.createElement("div");

  private readonly editor = document.createElement("input");

  private readonly selectEditor = document.createElement("select");

  private readonly customEditorHost = document.createElement("div");

  private readonly sheetTabs = document.createElement("div");

  private readonly textColorInput = document.createElement("input");

  private readonly borderColorInput = document.createElement("input");

  private readonly fillColorInput = document.createElement("input");

  private readonly colorPickerCard = document.createElement("section");

  private readonly colorPickerHandle = document.createElement("div");

  private colorPickerDragging?: { offsetX: number; offsetY: number };

  private colorPickerSelectionDragging = false;

  private activeColorPicker?: {
    input: HTMLInputElement;
    label: string;
    hue: number;
    saturation: number;
    value: number;
    left: string;
    top: string;
    transform: string;
  };

  private pendingColorStyle?: Partial<CellStyle>;

  private readonly messages: RendererMessages;

  private readonly shortcuts: RendererShortcutMap;

  private readonly direction: "ltr" | "rtl";

  private readonly unsubscribeCallbacks: Array<() => void> = [];

  private readonly composingTargets = new WeakSet<EventTarget>();

  private resizeObserver?: ResizeObserver;

  private editingCell?: { row: number; col: number; mode: "text" | "select" | "custom" };

  private renderedHeaderColumns = new Set<number>();

  private autofillDrag?: {
    sourceRange: CellRange;
    preview?: AutofillPreview;
  };

  private activeCustomEditor?: CustomCellEditorInstance;

  private validationFeedback?: { sheetId: string; row: number; col: number; message: string; isWarning?: boolean };

  private rowModelFeedback?: { sheetId: string; error?: string };

  private pivotFeedback?: { sheetId: string; message: string; isError: boolean };

  private visualObjectInteraction?: VisualObjectInteractionState;

  private visualObjectSurfaceMetrics?: VisualObjectSurfaceMetrics;

  private visualObjectLastInteractionMoveTs = 0;

  private lastViewportScrollRenderTs = 0;

  private splitPaneDrag?: { axis: "horizontal" | "vertical"; index: number };

  private readonly imageObjectElementById = new Map<string, HTMLElement>();

  private readonly widgetObjectElementById = new Map<string, HTMLElement>();

  private readonly widgetBodyElementById = new Map<string, HTMLElement>();

  private readonly widgetRenderSignatureById = new Map<string, string>();

  private readonly widgetCleanupById = new Map<string, () => void>();

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

  private noteEditorCell?: { sheetId: string; row: number; col: number };

  private localPresenceClientId?: string;

  private formatPainterStyle?: CellStyle;

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

  private formatDisplayValue(value: CellPrimitive, rawDisplayValue: string, style?: CellStyle): string {
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
      const format = style?.format;
      if (format && format !== "General") {
        const decimals = format.match(/\.(0+)/)?.[1].length ?? 0;
        if (format.includes("%")) {
          return new Intl.NumberFormat(locale, {
            style: "percent",
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
          }).format(numeric);
        }
        if (format.includes("R$")) {
          return new Intl.NumberFormat(locale ?? "pt-BR", {
            style: "currency",
            currency: "BRL",
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
          }).format(numeric);
        }
        if (/^[#,0]+(?:\.0+)?$/.test(format)) {
          return new Intl.NumberFormat(locale, {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals,
            useGrouping: format.includes(",")
          }).format(numeric);
        }
      }
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
    const cell = this.engine.getCell(sheetId, row, col);
    return this.formatDisplayValue(
      this.getCellPrimitiveValue(sheetId, row, col),
      this.engine.getDisplayValue(sheetId, row, col),
      cell?.style
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
    return sheet.columns[col]?.hidden
      ? 0
      : (sheet.columns[col]?.width ?? this.engine.getSnapshot().settings.columnWidth) * this.sheetZoom;
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
      if (this.engine.getRowModel(sheet.id).kind !== "clientSide") {
        return;
      }
      const current = this.engine.getClientSideQuery(sheet.id);
      const sort = (current?.sort ?? []).filter((item) => item.column !== activeAddress.col);
      sort.push({ column: activeAddress.col, direction });
      this.engine.applyClientSideSortFilter({
        sheetId: sheet.id,
        sort,
        filters: current?.filters,
        hasHeader: current?.hasHeader
      });
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
      if (this.engine.getRowModel(sheet.id).kind !== "clientSide") {
        return;
      }
      const local = this.engine.getClientSideQuery(sheet.id);
      this.engine.applyClientSideSortFilter({
        sheetId: sheet.id,
        sort: local?.sort.filter((item) => item.column !== activeAddress.col),
        filters: local?.filters.filter((item) => item.column !== activeAddress.col),
        hasHeader: local?.hasHeader
      });
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

  private applyLocalFilterForActiveColumn(): void {
    const sheet = this.engine.getActiveSheet();
    if (this.engine.getRowModel(sheet.id).kind !== "clientSide") {
      return;
    }
    const activeAddress = this.getActiveAddress(sheet);
    const type = this.toolbar.querySelector<HTMLSelectElement>("[data-local-filter-type]")?.value as ClientSideFilterDescriptor["type"];
    const operator = this.toolbar.querySelector<HTMLSelectElement>("[data-local-filter-operator]")?.value as ClientSideFilterDescriptor["operator"];
    const rawValue = this.toolbar.querySelector<HTMLInputElement>("[data-local-filter-value]")?.value.trim() ?? "";
    const rawValueTo = this.toolbar.querySelector<HTMLInputElement>("[data-local-filter-value-to]")?.value.trim() ?? "";
    if (!rawValue) {
      return;
    }
    const toValue = (value: string): string | number => {
      if (type !== "number") {
        return value;
      }
      const normalized = value.replaceAll(" ", "").replace(",", ".");
      return Number(normalized);
    };
    const value = toValue(rawValue);
    const valueTo = rawValueTo ? toValue(rawValueTo) : undefined;
    const valueIsInvalid = type === "number" && (typeof value !== "number" || !Number.isFinite(value));
    const valueToIsInvalid = type === "number" && (typeof valueTo !== "number" || !Number.isFinite(valueTo));
    if (valueIsInvalid || (operator === "between" && (valueTo === undefined || valueToIsInvalid))) {
      return;
    }
    const descriptor: ClientSideFilterDescriptor = {
      column: activeAddress.col,
      type,
      operator,
      value
    };
    if (operator === "between" && valueTo !== undefined) descriptor.valueTo = valueTo;
    const current = this.engine.getClientSideQuery(sheet.id);
    const filters = (current?.filters ?? []).filter((item) => item.column !== activeAddress.col);
    filters.push(descriptor);
    this.viewport.scrollTop = 0;
    this.engine.applyClientSideSortFilter({
      sheetId: sheet.id,
      sort: current?.sort,
      filters,
      hasHeader: current?.hasHeader
    });
  }

  private clearLocalFilters(): void {
    const sheet = this.engine.getActiveSheet();
    if (this.engine.getRowModel(sheet.id).kind !== "clientSide") {
      return;
    }
    const current = this.engine.getClientSideQuery(sheet.id);
    this.viewport.scrollTop = 0;
    this.engine.applyClientSideSortFilter({
      sheetId: sheet.id,
      sort: current?.sort,
      filters: [],
      hasHeader: current?.hasHeader
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
    return sheet.rows[row]?.hidden
      ? 0
      : (sheet.rows[row]?.height ?? this.engine.getSnapshot().settings.rowHeight) * this.sheetZoom;
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
    const split = this.engine.getSplitPane(sheet.id)?.horizontalRow ?? 0;
    appendRangeIndices(rows, 0, Math.max(-1, Math.max(frozen, split) - 1));
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
    const split = this.engine.getSplitPane(sheet.id)?.verticalColumn ?? 0;
    appendRangeIndices(columns, 0, Math.max(-1, Math.max(frozen, split) - 1));
    appendRangeIndices(columns, visibleColumns.start, visibleColumns.end);
    return columns.filter((col) => col >= 0 && col < sheet.columnCount);
  }

  private getFrozenAdjustedTop(sheetId: string, rowOffsets: number[], row: number): number {
    const frozenRows = this.engine.getFrozenPane(sheetId).rows;
    const splitRows = this.engine.getSplitPane(sheetId)?.horizontalRow ?? 0;
    return row < Math.max(frozenRows, splitRows) ? rowOffsets[row] + this.viewport.scrollTop : rowOffsets[row];
  }

  private getFrozenAdjustedLeft(sheetId: string, colOffsets: number[], col: number): number {
    const frozenColumns = this.engine.getFrozenPane(sheetId).columns;
    const splitColumns = this.engine.getSplitPane(sheetId)?.verticalColumn ?? 0;
    return col < Math.max(frozenColumns, splitColumns) ? colOffsets[col] + this.viewport.scrollLeft : colOffsets[col];
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
      ?.rules.find(
        (rule): rule is InteractiveValidationRule =>
          isListValidationRule(rule) || isCheckboxValidationRule(rule) || isDateValidationRule(rule)
      );
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
      const validation = this.engine.validateCellValue({ sheetId, row, col, value });
      this.engine.setCellValue({
        sheetId,
        row,
        col,
        value
      });
      if (validation.issue?.severity === "warning") {
        this.validationFeedback = { sheetId, row, col, message: validation.issue.message, isWarning: true };
        this.render();
      } else {
        this.clearValidationFeedback();
      }
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
      const model = this.engine.getCell(sheetId, row, col);
      if (!model?.formula && model?.richText?.length) {
        const style = this.getCellStyle(sheet, row, col);
        const content = document.createElement("span");
        content.className = "excelsior-cell-content excelsior-cell-rich-text";
        content.style.overflow = style?.overflow === "visible" ? "visible" : "hidden";
        content.style.textOverflow = style?.overflow === "ellipsis" ? "ellipsis" : "clip";
        for (const segment of model.richText) {
          const href = getSafeRichTextHref(segment.hyperlink);
          if (!href && !segment.style) {
            content.append(document.createTextNode(segment.text));
            continue;
          }

          const node = document.createElement(href ? "a" : "span");
          node.textContent = segment.text;
          node.style.fontWeight = segment.style?.bold ? "bold" : "";
          node.style.fontStyle = segment.style?.italic ? "italic" : "";
          node.style.textDecorationLine = [
            segment.style?.underline ? "underline" : "",
            segment.style?.strike ? "line-through" : ""
          ].filter(Boolean).join(" ");
          node.style.color = segment.style?.color ?? "";
          if (node instanceof HTMLAnchorElement && href) {
            node.href = href;
            node.rel = "noopener noreferrer";
            if (href.startsWith("https:")) {
              node.target = "_blank";
            }
          }
          content.append(node);
        }
        cell.replaceChildren(content);
        if (!cell.title && !style?.wrap && ["clip", "ellipsis"].includes(style?.overflow ?? "")) {
          cell.title = model.richText.map((segment) => segment.text).join("");
        }
        return;
      }
      const content = document.createElement("span");
      content.className = "excelsior-cell-content";
      content.textContent = value;
      cell.replaceChildren(content);
      if (!cell.title && !model?.style?.wrap && ["clip", "ellipsis"].includes(model?.style?.overflow ?? "")) {
        cell.title = value;
      }
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
    element.style.left = `${ROW_HEADER_WIDTH + colOffsets[targetRange.start.col]}px`;
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
    cell.style.textDecorationLine = [style.underline ? "underline" : "", style.strike ? "line-through" : ""]
      .filter(Boolean)
      .join(" ");
    cell.style.whiteSpace = style.wrap ? "normal" : "nowrap";
    cell.style.lineHeight = style.wrap ? "1.35" : "1";
    cell.style.overflow = style.overflow === "visible" ? "visible" : "hidden";
    cell.style.textOverflow = style.overflow === "ellipsis" ? "ellipsis" : "clip";
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

  private applyStyleToSelection(style: Partial<CellStyle>, mode: "merge" | "replace" = "merge"): void {
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
          style,
          mode
        });
      }
    }
  }

  private applyQuickSum(): void {
    const sheet = this.engine.getActiveSheet();
    const selection = sheet.selection;
    let source = selection;
    let target = { row: selection.end.row + 1, col: selection.end.col };
    if (selection.start.row === selection.end.row && selection.start.col === selection.end.col) {
      target = { ...selection.start };
      let startRow = target.row - 1;
      while (startRow >= 0 && toNumericValue(this.getCellPrimitiveValue(sheet.id, startRow, target.col)) !== undefined) {
        startRow -= 1;
      }
      source = {
        start: { row: startRow + 1, col: target.col },
        end: { row: target.row - 1, col: target.col }
      };
    }
    if (target.row < 0 || target.row >= sheet.rowCount || source.end.row < source.start.row) {
      return;
    }
    const from = `${columnIndexToLabel(source.start.col)}${source.start.row + 1}`;
    const to = `${columnIndexToLabel(source.end.col)}${source.end.row + 1}`;
    this.engine.setCellValue({ sheetId: sheet.id, row: target.row, col: target.col, value: `=SUM(${from}:${to})` });
    this.engine.selectRange({ sheetId: sheet.id, rowStart: target.row, rowEnd: target.row, colStart: target.col, colEnd: target.col });
  }

  private mergeSelectionByAxis(axis: "horizontal" | "vertical"): void {
    const sheet = this.engine.getActiveSheet();
    if (axis === "horizontal") {
      for (let row = sheet.selection.start.row; row <= sheet.selection.end.row; row += 1) {
        this.engine.mergeCells({
          sheetId: sheet.id,
          start: { row, col: sheet.selection.start.col },
          end: { row, col: sheet.selection.end.col }
        });
      }
      return;
    }
    for (let col = sheet.selection.start.col; col <= sheet.selection.end.col; col += 1) {
      this.engine.mergeCells({
        sheetId: sheet.id,
        start: { row: sheet.selection.start.row, col },
        end: { row: sheet.selection.end.row, col }
      });
    }
  }

  private applyBorder(position: "all" | "top" | "right" | "bottom" | "left" | "none"): void {
    if (position === "none") {
      this.applyStyleToSelection({ border: {} });
      return;
    }
    const edge = { color: "#334155", style: "thin" as const };
    this.applyStyleToSelection({
      border: position === "all"
        ? { top: edge, right: edge, bottom: edge, left: edge }
        : { [position]: edge }
    });
  }

  private insertLink(hyperlink: string, targetColumn: number): void {
    const sheet = this.engine.getActiveSheet();
    const address = this.getActiveAddress(sheet);
    const column = Number.isInteger(targetColumn) && targetColumn >= 0 && targetColumn < sheet.columnCount
      ? targetColumn
      : address.col;
    const current = this.engine.getDisplayValue(sheet.id, address.row, column);
    if (!hyperlink.trim()) return;
    this.engine.setCellRichText({
      sheetId: sheet.id,
      row: address.row,
      col: column,
      richText: [{ text: current || hyperlink, hyperlink: hyperlink.trim() }]
    });
  }

  private splitSelectedColumn(separatorInput: string, selectedColumn: number): void {
    if (!separatorInput) return;
    const sheet = this.engine.getActiveSheet();
    const sourceColumn = Number.isInteger(selectedColumn) && selectedColumn >= 0 && selectedColumn < sheet.columnCount
      ? selectedColumn
      : sheet.selection.start.col;
    const sampleSeparator = separatorInput.length > 1
      ? [",", ";", "|", "\t"].find((candidate) => separatorInput.includes(candidate))
      : undefined;
    const usesSample = sampleSeparator !== undefined && sheet.selection.start.row === sheet.selection.end.row;
    const separator = usesSample ? sampleSeparator : separatorInput;
    for (let row = sheet.selection.start.row; row <= sheet.selection.end.row; row += 1) {
      const sourceValue = usesSample ? separatorInput : String(this.getCellPrimitiveValue(sheet.id, row, sourceColumn) ?? "");
      const parts = sourceValue.split(separator);
      parts.forEach((part, index) => {
        if (sourceColumn + index < sheet.columnCount) {
          this.engine.setCellValue({ sheetId: sheet.id, row, col: sourceColumn + index, value: part.trim() });
        }
      });
    }
  }

  private exportVisibleGridSvg(): void {
    const bounds = this.viewport.getBoundingClientRect();
    const clone = this.viewport.cloneNode(true) as HTMLElement;
    clone.setAttribute("xmlns", "http://www.w3.org/1999/xhtml");
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${Math.max(1, bounds.width)}" height="${Math.max(1, bounds.height)}"><foreignObject width="100%" height="100%">${clone.outerHTML}</foreignObject></svg>`;
    const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${this.engine.getActiveSheet().name}.svg`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  private updateSheetZoom(delta?: number): void {
    this.sheetZoom = delta === undefined ? 1 : Math.max(0.5, Math.min(2, Math.round((this.sheetZoom + delta) * 10) / 10));
    this.render();
  }

  private configureSelectionValidation(modeInput: string, rawValue: string): void {
    const mode = modeInput.trim().toLocaleLowerCase();
    if (!mode) return;
    const sheet = this.engine.getActiveSheet();
    let validation: CellValidationConfig | undefined;
    if (mode === "lista") {
      const values = rawValue
        .split(",")
        .map((value) => value.trim())
        .filter(Boolean);
      if (!values.length) return;
      validation = { rules: [{ type: "dropdown", values }] };
    } else if (mode === "número" || mode === "numero") {
      validation = { rules: [{ type: "number" }] };
    } else if (mode === "data") {
      validation = { rules: [{ type: "date" }] };
    } else if (mode === "checkbox") {
      validation = { rules: [{ type: "checkbox" }] };
    } else if (mode !== "remover") {
      return;
    }
    for (let row = sheet.selection.start.row; row <= sheet.selection.end.row; row += 1) {
      for (let col = sheet.selection.start.col; col <= sheet.selection.end.col; col += 1) {
        this.engine.setCellValidation({ sheetId: sheet.id, row, col, validation });
      }
    }
  }

  private configureConditionalFormatting(modeInput: string, rawValue: string): void {
    const mode = modeInput.trim().toLocaleLowerCase();
    if (!mode) return;
    const sheet = this.engine.getActiveSheet();
    const range = this.normalizeRange(sheet.selection);
    const existing = this.engine.getConditionalFormattingRules(sheet.id);
    if (mode === "remover") {
      this.engine.setConditionalFormattingRules(sheet.id, existing.filter((rule) => !rangesOverlap(rule.range, range)));
      return;
    }
    const base = { id: `toolbar-${Date.now()}-${existing.length}`, range, priority: 10 };
    if (mode === "maior") {
      const value = Number(rawValue.replace(",", "."));
      if (!Number.isFinite(value)) return;
      this.engine.setConditionalFormattingRules(sheet.id, [...existing, {
        ...base,
        type: "greaterThan",
        value,
        style: { backgroundColor: "#fef08a", textColor: "#713f12", fontWeight: "bold" }
      }]);
    } else if (mode === "texto") {
      const text = rawValue.trim();
      if (!text) return;
      this.engine.setConditionalFormattingRules(sheet.id, [...existing, {
        ...base,
        type: "containsText",
        text,
        style: { backgroundColor: "#bfdbfe", textColor: "#1e3a8a" }
      }]);
    } else if (mode === "duplicados") {
      this.engine.setConditionalFormattingRules(sheet.id, [...existing, {
        ...base,
        type: "duplicates",
        style: { backgroundColor: "#fecaca", textColor: "#7f1d1d" }
      }]);
    } else if (mode === "escala") {
      this.engine.setConditionalFormattingRules(sheet.id, [...existing, {
        ...base,
        type: "colorScale",
        minColor: "#dcfce7",
        maxColor: "#ef4444"
      }]);
    }
  }

  private findSpecialCell(modeInput: string, queryInput = ""): void {
    const mode = modeInput.trim().toLocaleLowerCase();
    const query = queryInput.trim().toLocaleLowerCase();
    if (!mode) return;
    const sheet = this.engine.getActiveSheet();
    for (let row = 0; row < sheet.rowCount; row += 1) {
      for (let col = 0; col < sheet.columnCount; col += 1) {
        const cell = this.engine.getCell(sheet.id, row, col);
        const value = cell?.value;
        const matches = mode === "fórmulas" || mode === "formulas"
          ? typeof value === "string" && value.startsWith("=") && (!query || value.toLocaleLowerCase().includes(query))
          : mode === "vazias"
            ? value === undefined || value === null || value === ""
            : mode === "erros"
              ? Boolean(cell?.error)
              : mode === "constantes"
                ? value !== undefined && value !== null && !(typeof value === "string" && value.startsWith("="))
                : false;
        if (matches) {
          this.engine.selectRange({ sheetId: sheet.id, rowStart: row, rowEnd: row, colStart: col, colEnd: col });
          const rowOffsets = buildOffsets(sheet.rowCount, (index) => this.getRowHeight(sheet, index));
          const colOffsets = buildOffsets(sheet.columnCount, (index) => this.getColumnWidth(sheet, index));
          this.viewport.scrollTop = rowOffsets[row] ?? 0;
          this.viewport.scrollLeft = colOffsets[col] ?? 0;
          return;
        }
      }
    }
  }

  private openToolbarTool(tool: NonNullable<DomSpreadsheetRenderer["activeToolbarTool"]>): void {
    const sheet = this.engine.getActiveSheet();
    const activeColumn = this.getActiveAddress(sheet).col;
    const columnModes = Array.from({ length: sheet.columnCount }, (_, col) => {
      const label = columnIndexToLabel(col);
      const header = this.engine.getDisplayValue(sheet.id, 0, col).trim();
      return [String(col), header ? `${label} — ${header}` : label] as [string, string];
    });
    const definitions = {
      link: { title: "Inserir link", modeLabel: "Destino", modes: columnModes, valueLabel: "Endereço HTTPS ou mailto", value: "https://" },
      split: { title: "Dividir coluna", modeLabel: "Coluna", modes: columnModes, valueLabel: "Separador ou texto", value: "," },
      validation: {
        title: "Validação de dados",
        modeLabel: "Opção",
        modes: [["lista", "Lista"], ["numero", "Número"], ["data", "Data"], ["checkbox", "Checkbox"], ["remover", "Remover"]] as Array<[string, string]>,
        valueLabel: "Valores separados por vírgula",
        value: "Sim,Não"
      },
      conditional: {
        title: "Formatação condicional",
        modeLabel: "Opção",
        modes: [["maior", "Maior que"], ["texto", "Contém texto"], ["duplicados", "Duplicados"], ["escala", "Escala de cores"], ["remover", "Remover"]] as Array<[string, string]>,
        valueLabel: "Valor",
        value: "0"
      },
      "find-special": {
        title: "Localizar células especiais",
        modeLabel: "Opção",
        modes: [["formulas", "Fórmulas"], ["vazias", "Vazias"], ["erros", "Erros"], ["constantes", "Constantes"]] as Array<[string, string]>,
        valueLabel: "Buscar fórmula (opcional)",
        value: ""
      }
    } as const;
    const definition = definitions[tool];
    this.cancelFindReplaceSearch();
    this.findReplaceState.open = false;
    this.activeToolbarTool = tool;
    this.toolbarToolTitle.textContent = definition.title;
    this.toolbarToolModeLabel.textContent = definition.modeLabel;
    this.toolbarToolModeSelect.replaceChildren(...definition.modes.map(([value, label]) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      return option;
    }));
    if (tool === "link" || tool === "split") {
      this.toolbarToolModeSelect.value = String(activeColumn);
    }
    this.toolbarToolModeField.hidden = definition.modes.length === 0;
    this.toolbarToolValueLabel.textContent = definition.valueLabel;
    this.toolbarToolValueInput.value = definition.value;
    this.updateToolbarToolFields();
    this.toolbarToolPanel.hidden = false;
    (this.toolbarToolModeField.hidden ? this.toolbarToolValueInput : this.toolbarToolModeSelect).focus();
  }

  private readonly updateToolbarToolFields = (): void => {
    const mode = this.toolbarToolModeSelect.value;
    const needsValue = this.activeToolbarTool === "link" || this.activeToolbarTool === "split"
      || (this.activeToolbarTool === "validation" && mode === "lista")
      || (this.activeToolbarTool === "conditional" && (mode === "maior" || mode === "texto"))
      || (this.activeToolbarTool === "find-special" && mode === "formulas");
    this.toolbarToolValueField.hidden = !needsValue;
    if (this.activeToolbarTool === "conditional") {
      this.toolbarToolValueLabel.textContent = mode === "texto" ? "Texto" : "Valor";
    }
  };

  private closeToolbarTool(): void {
    this.toolbarToolPanel.hidden = true;
    this.activeToolbarTool = undefined;
    this.focus();
  }

  private applyToolbarTool(): void {
    const tool = this.activeToolbarTool;
    const mode = this.toolbarToolModeSelect.value;
    const value = this.toolbarToolValueInput.value;
    this.toolbarToolPanel.hidden = true;
    if (tool === "link") this.insertLink(value, Number(mode));
    else if (tool === "split") this.splitSelectedColumn(value, Number(mode));
    else if (tool === "validation") this.configureSelectionValidation(mode, value);
    else if (tool === "conditional") this.configureConditionalFormatting(mode, value);
    else if (tool === "find-special") this.findSpecialCell(mode, value);
    this.activeToolbarTool = undefined;
    this.render();
    this.focus();
  }

  private readonly handleToolbarToolClick = (event: Event): void => {
    const action = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-tool-action]")?.dataset.toolAction;
    if (action === "apply") this.applyToolbarTool();
    else if (action === "cancel") this.closeToolbarTool();
  };

  private readonly handleToolbarToolKeyDown = (event: KeyboardEvent): void => {
    if (event.key === "Escape") {
      event.preventDefault();
      this.closeToolbarTool();
    } else if (event.key === "Enter") {
      event.preventDefault();
      this.applyToolbarTool();
    }
  };

  private applyFormatPainterToSelection(): void {
    const sourceStyle = this.formatPainterStyle;
    if (!sourceStyle) {
      return;
    }
    const sheet = this.engine.getActiveSheet();
    const processed = new Set<string>();
    const operations: SpreadsheetOperation[] = [];
    for (let row = sheet.selection.start.row; row <= sheet.selection.end.row; row += 1) {
      for (let col = sheet.selection.start.col; col <= sheet.selection.end.col; col += 1) {
        const { address } = this.resolveCellAddress(sheet, row, col);
        const key = getCellKey(address.row, address.col);
        if (processed.has(key)) {
          continue;
        }
        processed.add(key);
        const previousCell = this.engine.getCell(sheet.id, address.row, address.col);
        operations.push({
          op: previousCell ? "replace" : "add",
          id: sheet.id,
          path: ["cells", key],
          value: {
            ...previousCell,
            value: previousCell?.value ?? null,
            computedValue: previousCell?.computedValue ?? previousCell?.value ?? null,
            style: cloneSerializable(sourceStyle)
          }
        });
      }
    }
    if (operations.length) {
      this.engine.applyBatchOperations({
        anchorSheetId: sheet.id,
        operations,
        affectedRanges: [sheet.selection]
      });
    }
    this.formatPainterStyle = undefined;
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
    this.toolbarToolPanel.hidden = true;
    this.activeToolbarTool = undefined;
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
    this.gridPanel.className = "excelsior-grid-panel";
    this.chrome.className = "excelsior-chrome";
    this.toolbar.className = "excelsior-toolbar";
    this.imageFileInput.type = "file";
    this.imageFileInput.accept = "image/png,image/jpeg,image/gif,image/webp";
    this.imageFileInput.hidden = true;
    this.imageFileInput.addEventListener("change", this.handleImageFileChange);
    this.formulaBar.className = "excelsior-formula-bar";
    this.formulaAddress.className = "excelsior-formula-address";
    this.formulaInput.className = "excelsior-formula-input";
    this.statusMessage.className = "excelsior-status-message";
    this.statusMessage.setAttribute("role", "status");
    this.statusMessage.setAttribute("aria-live", "polite");
    this.statusMessage.setAttribute("aria-atomic", "true");
    this.findReplacePanel.className = "excelsior-find-replace";
    this.notePanel.className = "excelsior-note-panel";
    this.pivotPanel.className = "excelsior-find-replace excelsior-pivot-panel";
    this.toolbarToolPanel.className = "excelsior-toolbar-tool-panel";
    this.viewport.className = "excelsior-viewport";
    this.rowHeaders.className = "excelsior-row-headers";
    this.viewport.tabIndex = 0;
    this.viewport.setAttribute("role", "grid");
    this.viewport.setAttribute("aria-label", this.messages.gridLabel);
    this.viewport.setAttribute("aria-multiselectable", "true");
    this.root.dir = this.direction;
    this.viewport.dir = this.direction;
    this.root.classList.toggle("is-rtl", this.isRtl());
    this.rowHeaders.setAttribute("aria-hidden", "true");
    this.surface.className = "excelsior-surface";
    this.cellsLayer.className = "excelsior-cells";
    this.visualObjectsLayer.className = "excelsior-visual-objects-layer";
    this.splitPaneLayer.className = "excelsior-split-pane-layer";
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
    this.notePanel.hidden = true;
    this.toolbarToolPanel.hidden = true;
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

    this.toolbarToolPanel.setAttribute("role", "dialog");
    this.toolbarToolPanel.setAttribute("aria-modal", "false");
    this.toolbarToolPanel.dataset.toolbarToolPanel = "true";
    this.toolbarToolTitle.className = "excelsior-toolbar-tool-title";
    this.toolbarToolModeField.className = "excelsior-find-replace-field";
    this.toolbarToolModeLabel.textContent = "Opção";
    this.toolbarToolModeSelect.className = "excelsior-find-replace-input";
    this.toolbarToolModeSelect.dataset.toolbarToolMode = "true";
    this.toolbarToolModeField.append(this.toolbarToolModeLabel, this.toolbarToolModeSelect);
    this.toolbarToolValueField.className = "excelsior-find-replace-field";
    this.toolbarToolValueInput.className = "excelsior-find-replace-input";
    this.toolbarToolValueInput.dataset.toolbarToolValue = "true";
    this.toolbarToolValueField.append(this.toolbarToolValueLabel, this.toolbarToolValueInput);
    const toolbarToolActions = document.createElement("div");
    toolbarToolActions.className = "excelsior-find-replace-actions";
    const toolbarToolApply = createFindReplaceActionButton("apply", "Aplicar");
    toolbarToolApply.dataset.toolAction = "apply";
    delete toolbarToolApply.dataset.findAction;
    const toolbarToolCancel = createFindReplaceActionButton("cancel", "Cancelar");
    toolbarToolCancel.dataset.toolAction = "cancel";
    delete toolbarToolCancel.dataset.findAction;
    toolbarToolActions.append(toolbarToolApply, toolbarToolCancel);
    this.toolbarToolPanel.append(this.toolbarToolTitle, this.toolbarToolModeField, this.toolbarToolValueField, toolbarToolActions);

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

    const noteTitle = document.createElement("strong");
    noteTitle.className = "excelsior-note-panel-title";
    noteTitle.textContent = this.messages.cellNoteLabel;
    this.noteInput.className = "excelsior-note-input";
    this.noteInput.setAttribute("aria-label", this.messages.cellNoteLabel);
    this.noteInput.maxLength = this.engine.getSnapshot().settings.maxCellLength;
    this.noteSaveButton.textContent = this.messages.cellNoteSave;
    this.noteSaveButton.dataset.noteAction = "save";
    delete this.noteSaveButton.dataset.findAction;
    this.noteRemoveButton.textContent = this.messages.cellNoteRemove;
    this.noteRemoveButton.dataset.noteAction = "remove";
    delete this.noteRemoveButton.dataset.findAction;
    this.noteCloseButton.textContent = this.messages.cellNoteClose;
    this.noteCloseButton.dataset.noteAction = "close";
    delete this.noteCloseButton.dataset.findAction;
    const noteActions = document.createElement("div");
    noteActions.className = "excelsior-find-replace-actions";
    noteActions.append(this.noteSaveButton, this.noteRemoveButton, this.noteCloseButton);
    const commentTitle = document.createElement("strong");
    commentTitle.className = "excelsior-note-panel-title";
    commentTitle.textContent = "Comentários";
    this.commentList.className = "excelsior-comment-list";
    this.commentList.setAttribute("aria-live", "polite");
    this.commentInput.className = "excelsior-note-input excelsior-comment-input";
    this.commentInput.dataset.commentInput = "new";
    this.commentInput.maxLength = this.engine.getSnapshot().settings.maxCellLength;
    this.commentInput.setAttribute("aria-label", "Adicionar comentário");
    const commentCreateButton = createFindReplaceActionButton("create", "Adicionar comentário");
    commentCreateButton.dataset.commentAction = "create";
    delete commentCreateButton.dataset.findAction;
    this.notePanel.append(noteTitle, this.noteInput, noteActions, commentTitle, this.commentList, this.commentInput, commentCreateButton);

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

    this.surface.append(
      this.cellsLayer,
      this.visualObjectsLayer,
      this.splitPaneLayer,
      this.editor,
      this.selectEditor,
      this.customEditorHost
    );
    this.formulaBar.append(
      this.formulaAddress,
      this.formulaInput,
      this.statusMessage,
      this.findReplacePanel,
      this.pivotPanel,
      this.toolbarToolPanel
    );
    this.root.append(this.chrome, this.formulaBar, this.activeCellAnnouncement, this.gridPanel, this.sheetTabs, this.imageFileInput);
    this.gridPanel.append(this.viewport, this.rowHeaders, this.notePanel);
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
    if (this.localPresenceClientId) {
      this.engine.removePresence(this.localPresenceClientId);
      this.localPresenceClientId = undefined;
    }
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
    this.toolbarToolPanel.removeEventListener("click", this.handleToolbarToolClick);
    this.toolbarToolPanel.removeEventListener("keydown", this.handleToolbarToolKeyDown);
    this.toolbarToolModeSelect.removeEventListener("input", this.updateToolbarToolFields);
    this.imageFileInput.removeEventListener("change", this.handleImageFileChange);
    this.chrome.removeEventListener("click", this.handleColumnHeaderClick);
    this.chrome.removeEventListener("keydown", this.handleColumnHeaderKeyDown);
    this.rowHeaders.removeEventListener("click", this.handleRowHeaderClick);
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
    this.notePanel.removeEventListener("click", this.handleNotePanelClick);
    this.notePanel.removeEventListener("keydown", this.handleNotePanelKeyDown);
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
    this.visualObjectsLayer.removeEventListener("mousedown", this.handleVisualObjectLayerMouseDown);
    this.visualObjectsLayer.removeEventListener("click", this.handleVisualObjectLayerClick);
    this.visualObjectsLayer.removeEventListener("keydown", this.handleVisualObjectKeyDown);
    this.splitPaneLayer.removeEventListener("mousedown", this.handleSplitPaneMouseDown);
    this.splitPaneLayer.removeEventListener("keydown", this.handleSplitPaneKeyDown);
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
    this.colorPickerHandle.removeEventListener("mousedown", this.handleColorPickerMouseDown);
    globalThis.removeEventListener("mousemove", this.handleColorPickerMouseMove);
    globalThis.removeEventListener("mouseup", this.handleColorPickerMouseUp);
    globalThis.removeEventListener("mousemove", this.handleAutofillMouseMove);
    globalThis.removeEventListener("mouseup", this.handleAutofillMouseUp);
    globalThis.removeEventListener("mousemove", this.handleVisualObjectInteractionMouseMove);
    globalThis.removeEventListener("mouseup", this.handleVisualObjectInteractionMouseUp);
    globalThis.removeEventListener("mousemove", this.handleSplitPaneMouseMove);
    globalThis.removeEventListener("mouseup", this.handleSplitPaneMouseUp);
    this.destroyAllWidgetRuntimes();
    this.destroyCustomEditor();
    this.colorPickerCard.remove();
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
        this.updateLocalPresence();
        this.requestRender();
      }),
      this.engine.on("collaboration:presenceChanged", () => {
        this.requestRender();
      }),
      this.engine.on("collaboration:presenceRemoved", () => {
        this.requestRender();
      }),
      this.engine.on("cell:commentCreated", () => {
        this.renderCommentThreads();
      }),
      this.engine.on("cell:commentReplied", () => {
        this.renderCommentThreads();
      }),
      this.engine.on("cell:commentResolved", () => {
        this.renderCommentThreads();
      }),
      this.engine.on("cell:commentDeleted", () => {
        this.renderCommentThreads();
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
    this.toolbarToolPanel.addEventListener("click", this.handleToolbarToolClick);
    this.toolbarToolPanel.addEventListener("keydown", this.handleToolbarToolKeyDown);
    this.toolbarToolModeSelect.addEventListener("input", this.updateToolbarToolFields);
    this.textColorInput.type = "color";
    this.textColorInput.value = "#000000";
    this.textColorInput.setAttribute("aria-hidden", "true");
    this.textColorInput.addEventListener("change", this.handleTextColorChange);
    this.borderColorInput.type = "color";
    this.borderColorInput.value = "#000000";
    this.borderColorInput.setAttribute("aria-hidden", "true");
    this.borderColorInput.addEventListener("change", this.handleBorderColorChange);
    this.fillColorInput.type = "color";
    this.fillColorInput.value = "#ffffff";
    this.fillColorInput.setAttribute("aria-hidden", "true");
    this.fillColorInput.addEventListener("change", this.handleFillColorChange);
    this.colorPickerHandle.className = "excelsior-color-picker-handle";
    this.colorPickerHandle.addEventListener("mousedown", this.handleColorPickerMouseDown);
    this.chrome.addEventListener("click", this.handleColumnHeaderClick);
    this.chrome.addEventListener("keydown", this.handleColumnHeaderKeyDown);
    this.rowHeaders.addEventListener("click", this.handleRowHeaderClick);
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
    this.notePanel.addEventListener("click", this.handleNotePanelClick);
    this.notePanel.addEventListener("keydown", this.handleNotePanelKeyDown);
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
    this.visualObjectsLayer.addEventListener("mousedown", this.handleVisualObjectLayerMouseDown);
    this.visualObjectsLayer.addEventListener("click", this.handleVisualObjectLayerClick);
    this.visualObjectsLayer.addEventListener("keydown", this.handleVisualObjectKeyDown);
    this.splitPaneLayer.addEventListener("mousedown", this.handleSplitPaneMouseDown);
    this.splitPaneLayer.addEventListener("keydown", this.handleSplitPaneKeyDown);
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
    globalThis.addEventListener("mousemove", this.handleVisualObjectInteractionMouseMove);
    globalThis.addEventListener("mouseup", this.handleVisualObjectInteractionMouseUp);
    globalThis.addEventListener("mousemove", this.handleSplitPaneMouseMove);
    globalThis.addEventListener("mouseup", this.handleSplitPaneMouseUp);

    if (typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => this.requestRender());
      this.resizeObserver.observe(this.viewport);
    } else {
      window.addEventListener("resize", this.handleWindowResize);
      this.unsubscribeCallbacks.push(() => window.removeEventListener("resize", this.handleWindowResize));
    }
  }

  private readonly handleScroll = (): void => {
    const throttleMs = 16;
    const now = Date.now();
    if (throttleMs > 0 && now - this.lastViewportScrollRenderTs < throttleMs) {
      return;
    }
    this.lastViewportScrollRenderTs = now;
    this.requestRender();
  };

  private readonly handleWindowResize = (): void => {
    this.requestRender();
  };

  private findSplitIndex(offsets: number[], pointerOffset: number): number {
    let bestIndex = 1;
    let bestDistance = Number.POSITIVE_INFINITY;
    for (let index = 1; index < offsets.length - 1; index += 1) {
      const distance = Math.abs(offsets[index] - pointerOffset);
      if (distance < bestDistance) {
        bestIndex = index;
        bestDistance = distance;
      }
    }
    return bestIndex;
  }

  private readonly handleSplitPaneMouseDown = (event: MouseEvent): void => {
    const divider = (event.target as HTMLElement).closest<HTMLElement>("[data-split-axis]");
    const axis = divider?.dataset.splitAxis;
    if (!divider || (axis !== "horizontal" && axis !== "vertical")) {
      return;
    }
    event.preventDefault();
    this.splitPaneDrag = { axis, index: Number(divider.dataset.splitIndex) };
  };

  private readonly handleSplitPaneMouseMove = (event: MouseEvent): void => {
    if (!this.splitPaneDrag || !this.visualObjectSurfaceMetrics) {
      return;
    }
    const rect = this.viewport.getBoundingClientRect();
    const offsets = this.splitPaneDrag.axis === "horizontal"
      ? this.visualObjectSurfaceMetrics.rowOffsets
      : this.visualObjectSurfaceMetrics.colOffsets;
    const pointerOffset = this.splitPaneDrag.axis === "horizontal"
      ? event.clientY - rect.top
      : event.clientX - rect.left - ROW_HEADER_WIDTH;
    this.splitPaneDrag.index = this.findSplitIndex(offsets, pointerOffset);
    const divider = this.splitPaneLayer.querySelector<HTMLElement>(`[data-split-axis='${this.splitPaneDrag.axis}']`);
    if (divider) {
      const position = offsets[this.splitPaneDrag.index];
      divider.style[this.splitPaneDrag.axis === "horizontal" ? "top" : "left"] = `${position + (
        this.splitPaneDrag.axis === "horizontal" ? this.viewport.scrollTop : this.viewport.scrollLeft + ROW_HEADER_WIDTH
      )}px`;
    }
  };

  private readonly handleSplitPaneMouseUp = (): void => {
    if (!this.splitPaneDrag) {
      return;
    }
    const sheet = this.engine.getActiveSheet();
    const current = this.engine.getSplitPane(sheet.id) ?? {};
    const next = this.splitPaneDrag.axis === "horizontal"
      ? { ...current, horizontalRow: this.splitPaneDrag.index }
      : { ...current, verticalColumn: this.splitPaneDrag.index };
    this.splitPaneDrag = undefined;
    this.engine.setSplitPane(sheet.id, next);
  };

  private readonly handleSplitPaneKeyDown = (event: KeyboardEvent): void => {
    const divider = (event.target as HTMLElement).closest<HTMLElement>("[data-split-axis]");
    const axis = divider?.dataset.splitAxis;
    if (!divider || (axis !== "horizontal" && axis !== "vertical")) {
      return;
    }
    const sheet = this.engine.getActiveSheet();
    const current = this.engine.getSplitPane(sheet.id) ?? {};
    if (event.key === "Delete") {
      event.preventDefault();
      const next = axis === "horizontal"
        ? { ...current, horizontalRow: undefined }
        : { ...current, verticalColumn: undefined };
      this.engine.setSplitPane(sheet.id, next);
      return;
    }
    let delta = 0;
    if (event.key === "ArrowUp" || event.key === "ArrowLeft") {
      delta = -1;
    } else if (event.key === "ArrowDown" || event.key === "ArrowRight") {
      delta = 1;
    }
    if (!delta) {
      return;
    }
    event.preventDefault();
    const currentIndex = Number(divider.dataset.splitIndex);
    const maxIndex = axis === "horizontal" ? sheet.rowCount - 1 : sheet.columnCount - 1;
    const nextIndex = Math.min(Math.max(currentIndex + delta, 1), maxIndex);
    this.engine.setSplitPane(sheet.id, axis === "horizontal"
      ? { ...current, horizontalRow: nextIndex }
      : { ...current, verticalColumn: nextIndex });
  };

  private renderSplitPanes(
    sheet: ReturnType<WorkbookEngine["getActiveSheet"]>,
    rowOffsets: number[],
    colOffsets: number[]
  ): void {
    const splitPane = this.engine.getSplitPane(sheet.id);
    const fragment = document.createDocumentFragment();
    const appendDivider = (axis: "horizontal" | "vertical", index: number, position: number): void => {
      const divider = document.createElement("div");
      divider.className = `excelsior-split-divider is-${axis}`;
      divider.dataset.splitAxis = axis;
      divider.dataset.splitIndex = String(index);
      divider.tabIndex = 0;
      divider.setAttribute("role", "separator");
      divider.setAttribute("aria-orientation", axis);
      divider.setAttribute("aria-valuemin", "1");
      divider.setAttribute("aria-valuemax", String(axis === "horizontal" ? sheet.rowCount - 1 : sheet.columnCount - 1));
      divider.setAttribute("aria-valuenow", String(index));
      divider.setAttribute("aria-label", axis === "horizontal" ? "Divisor horizontal" : "Divisor vertical");
      divider.style[axis === "horizontal" ? "top" : "left"] = `${position}px`;
      fragment.append(divider);
    };
    if (splitPane?.horizontalRow !== undefined) {
      appendDivider("horizontal", splitPane.horizontalRow, rowOffsets[splitPane.horizontalRow] + this.viewport.scrollTop);
    }
    if (splitPane?.verticalColumn !== undefined) {
      appendDivider(
        "vertical",
        splitPane.verticalColumn,
        ROW_HEADER_WIDTH + colOffsets[splitPane.verticalColumn] + this.viewport.scrollLeft
      );
    }
    this.splitPaneLayer.replaceChildren(fragment);
  }

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
    const noteIndicator = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-cell-note]");
    if (noteIndicator) {
      const cell = noteIndicator.closest<HTMLElement>("[data-row][data-col]");
      if (cell) {
        const sheet = this.engine.getActiveSheet();
        this.selectResolvedCell(sheet, Number(cell.dataset.row), Number(cell.dataset.col));
        this.openNotePanel(sheet.id, Number(cell.dataset.row), Number(cell.dataset.col));
      }
      event.preventDefault();
      return;
    }

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
    if (this.formatPainterStyle) {
      this.applyFormatPainterToSelection();
      this.render();
    }
    this.focus();
  };

  private readonly handleCellDoubleClick = (event: Event): void => {
    if ((event.target as HTMLElement | null)?.closest("[data-remote-group-toggle], [data-cell-note]")) {
      return;
    }

    const target = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-row][data-col]");
    if (!target) {
      return;
    }
    this.startEditing(Number(target.dataset.row), Number(target.dataset.col));
  };

  private readonly handleColumnHeaderClick = (event: Event): void => {
    const corner = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-corner-action='select-all']");
    if (corner) {
      const sheet = this.engine.getActiveSheet();
      this.engine.selectRange({
        sheetId: sheet.id,
        rowStart: 0,
        rowEnd: Math.max(0, sheet.rowCount - 1),
        colStart: 0,
        colEnd: Math.max(0, sheet.columnCount - 1)
      });
      this.focus();
      return;
    }

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

  private readonly handleRowHeaderClick = (event: Event): void => {
    const target = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-row-header-row]");
    if (!target) {
      return;
    }

    const row = Number(target.dataset.rowHeaderRow);
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    this.selectResolvedCell(sheet, row, activeAddress.col);
    this.focus();
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

  private readonly handleTextColorChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const color = input.value;
    if (color) {
      this.queueColorChange({ textColor: color });
    }
  };

  private readonly handleBorderColorChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const color = input.value;
    if (color) {
      this.queueColorChange({
        border: {
          top: { color, style: "thin" },
          right: { color, style: "thin" },
          bottom: { color, style: "thin" },
          left: { color, style: "thin" }
        }
      });
    }
  };

  private readonly handleFillColorChange = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    const color = input.value;
    if (color) {
      this.queueColorChange({ backgroundColor: color });
    }
  };

  private readonly queueColorChange = (style: Partial<CellStyle>): void => {
    this.pendingColorStyle = style;
  };

  private readonly buildPendingColorStyle = (input: HTMLInputElement, color: string): Partial<CellStyle> => {
    if (input === this.textColorInput) {
      return { textColor: color };
    }

    if (input === this.fillColorInput) {
      return { backgroundColor: color };
    }

    return {
      border: {
        top: { color, style: "thin" },
        right: { color, style: "thin" },
        bottom: { color, style: "thin" },
        left: { color, style: "thin" }
      }
    };
  };

  private readonly clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

  private readonly hexToRgb = (hex: string): { r: number; g: number; b: number } => {
    const normalized = hex.replace("#", "").trim();
    const value = normalized.length === 3
      ? normalized
          .split("")
          .map((part) => `${part}${part}`)
          .join("")
      : normalized.padEnd(6, "0").slice(0, 6);

    return {
      r: Number.parseInt(value.slice(0, 2), 16),
      g: Number.parseInt(value.slice(2, 4), 16),
      b: Number.parseInt(value.slice(4, 6), 16)
    };
  };

  private readonly rgbToHex = (r: number, g: number, b: number): string => {
    const toHex = (value: number) => this.clamp(Math.round(value), 0, 255).toString(16).padStart(2, "0");
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`.toUpperCase();
  };

  private readonly rgbToHsv = (
    r: number,
    g: number,
    b: number
  ): { hue: number; saturation: number; value: number } => {
    const red = r / 255;
    const green = g / 255;
    const blue = b / 255;
    const max = Math.max(red, green, blue);
    const min = Math.min(red, green, blue);
    const delta = max - min;
    let hue = 0;

    if (delta !== 0) {
      if (max === red) {
        hue = 60 * (((green - blue) / delta) % 6);
      } else if (max === green) {
        hue = 60 * (((blue - red) / delta) + 2);
      } else {
        hue = 60 * (((red - green) / delta) + 4);
      }
    }

    if (hue < 0) {
      hue += 360;
    }

    return {
      hue,
      saturation: max === 0 ? 0 : delta / max,
      value: max
    };
  };

  private readonly hsvToRgb = (
    hue: number,
    saturation: number,
    value: number
  ): { r: number; g: number; b: number } => {
    const chroma = value * saturation;
    const segment = hue / 60;
    const intermediate = chroma * (1 - Math.abs((segment % 2) - 1));
    const match = value - chroma;
    let red = 0;
    let green = 0;
    let blue = 0;

    if (segment >= 0 && segment < 1) {
      red = chroma;
      green = intermediate;
    } else if (segment < 2) {
      red = intermediate;
      green = chroma;
    } else if (segment < 3) {
      green = chroma;
      blue = intermediate;
    } else if (segment < 4) {
      green = intermediate;
      blue = chroma;
    } else if (segment < 5) {
      red = intermediate;
      blue = chroma;
    } else {
      red = chroma;
      blue = intermediate;
    }

    return {
      r: Math.round((red + match) * 255),
      g: Math.round((green + match) * 255),
      b: Math.round((blue + match) * 255)
    };
  };

  private readonly getActiveColorHex = (): string => {
    const state = this.activeColorPicker;
    if (!state) {
      return "#000000";
    }

    const { r, g, b } = this.hsvToRgb(state.hue, state.saturation, state.value);
    return this.rgbToHex(r, g, b);
  };

  private readonly syncActiveColorPicker = (): void => {
    const state = this.activeColorPicker;
    if (!state) {
      return;
    }

    const hex = this.getActiveColorHex();
    state.input.value = hex;
    this.queueColorChange(this.buildPendingColorStyle(state.input, hex));
  };

  private readonly closeColorPicker = (): void => {
    this.activeColorPicker = undefined;
    this.pendingColorStyle = undefined;
    this.colorPickerCard.remove();
    this.handleColorPickerMouseUp();
    this.handleColorSurfaceMouseUp();
  };

  private readonly renderColorPickerCard = (): void => {
    const state = this.activeColorPicker;
    if (!state) {
      this.colorPickerCard.remove();
      return;
    }

    this.syncActiveColorPicker();
    const { r, g, b } = this.hsvToRgb(state.hue, state.saturation, state.value);
    const currentHex = this.rgbToHex(r, g, b);

    this.colorPickerCard.replaceChildren();
    this.colorPickerCard.className = "excelsior-color-picker-card";
    this.colorPickerCard.style.left = state.left;
    this.colorPickerCard.style.top = state.top;
    this.colorPickerCard.style.transform = state.transform;

    const header = document.createElement("div");
    header.className = "excelsior-color-picker-header";
    this.colorPickerHandle.className = "excelsior-color-picker-handle";
    this.colorPickerHandle.textContent = state.label;
    const closeButton = document.createElement("button");
    closeButton.type = "button";
    closeButton.className = "excelsior-color-picker-close";
    closeButton.textContent = "×";
    closeButton.setAttribute("aria-label", "Fechar seletor de cor");
    closeButton.title = "Fechar";
    closeButton.addEventListener("click", () => {
      this.closeColorPicker();
      this.render();
      this.focus();
    });
    header.append(this.colorPickerHandle, closeButton);

    const surface = document.createElement("div");
    surface.className = "excelsior-color-picker-surface";
    surface.style.background = `linear-gradient(to top, #000000, transparent), linear-gradient(to right, #ffffff, hsl(${state.hue} 100% 50%))`;
    surface.addEventListener("mousedown", this.handleColorSurfaceMouseDown);

    const cursor = document.createElement("div");
    cursor.className = "excelsior-color-picker-cursor";
    cursor.style.left = `${state.saturation * 100}%`;
    cursor.style.top = `${(1 - state.value) * 100}%`;
    surface.append(cursor);

    const hueInput = document.createElement("input");
    hueInput.type = "range";
    hueInput.className = "excelsior-color-picker-hue";
    hueInput.min = "0";
    hueInput.max = "360";
    hueInput.value = String(Math.round(state.hue));
    hueInput.addEventListener("input", () => {
      if (!this.activeColorPicker) {
        return;
      }
      this.activeColorPicker.hue = Number(hueInput.value);
      this.renderColorPickerCard();
    });

    const details = document.createElement("div");
    details.className = "excelsior-color-picker-details";

    const preview = document.createElement("div");
    preview.className = "excelsior-color-picker-preview";
    preview.style.background = currentHex;
    details.append(preview);

    const channels = document.createElement("div");
    channels.className = "excelsior-color-picker-channels";
    const createChannel = (label: "R" | "G" | "B", value: number) => {
      const field = document.createElement("label");
      field.className = "excelsior-color-picker-channel";
      const input = document.createElement("input");
      input.type = "number";
      input.min = "0";
      input.max = "255";
      input.value = String(value);
      input.addEventListener("change", () => {
        const wrapper = input.closest(".excelsior-color-picker-channels");
        if (!wrapper || !this.activeColorPicker) {
          return;
        }
        const values = Array.from(wrapper.querySelectorAll<HTMLInputElement>("input")).map((element) =>
          this.clamp(Number(element.value || "0"), 0, 255)
        );
        const [nextR, nextG, nextB] = values;
        const hsv = this.rgbToHsv(nextR, nextG, nextB);
        this.activeColorPicker.hue = hsv.hue;
        this.activeColorPicker.saturation = hsv.saturation;
        this.activeColorPicker.value = hsv.value;
        this.renderColorPickerCard();
      });
      const caption = document.createElement("span");
      caption.textContent = label;
      field.append(input, caption);
      return field;
    };

    channels.append(createChannel("R", r), createChannel("G", g), createChannel("B", b));
    details.append(channels);

    const confirmButton = document.createElement("button");
    confirmButton.type = "button";
    confirmButton.className = "excelsior-color-confirmation-button";
    confirmButton.textContent = "Confirmar cor";
    confirmButton.setAttribute("aria-label", "Confirmar cor");
    confirmButton.addEventListener("click", () => {
      if (!this.pendingColorStyle) {
        return;
      }
      this.applyStyleToSelection(this.pendingColorStyle);
      this.closeColorPicker();
      this.render();
      this.focus();
    });

    this.colorPickerCard.append(header, surface, hueInput, details, confirmButton);
    if (!this.colorPickerCard.isConnected) {
      document.body.append(this.colorPickerCard);
    }
  };

  private readonly updateColorPickerFromSurface = (event: MouseEvent): void => {
    const state = this.activeColorPicker;
    const surface = this.colorPickerCard.querySelector<HTMLElement>(".excelsior-color-picker-surface");
    if (!state || !surface) {
      return;
    }

    const bounds = surface.getBoundingClientRect();
    state.saturation = this.clamp((event.clientX - bounds.left) / bounds.width, 0, 1);
    state.value = 1 - this.clamp((event.clientY - bounds.top) / bounds.height, 0, 1);
    this.renderColorPickerCard();
  };

  private readonly handleColorSurfaceMouseDown = (event: MouseEvent): void => {
    this.colorPickerSelectionDragging = true;
    this.updateColorPickerFromSurface(event);
    event.preventDefault();
    globalThis.addEventListener("mousemove", this.handleColorSurfaceMouseMove);
    globalThis.addEventListener("mouseup", this.handleColorSurfaceMouseUp);
  };

  private readonly handleColorSurfaceMouseMove = (event: MouseEvent): void => {
    if (!this.colorPickerSelectionDragging) {
      return;
    }

    this.updateColorPickerFromSurface(event);
  };

  private readonly handleColorSurfaceMouseUp = (): void => {
    this.colorPickerSelectionDragging = false;
    globalThis.removeEventListener("mousemove", this.handleColorSurfaceMouseMove);
    globalThis.removeEventListener("mouseup", this.handleColorSurfaceMouseUp);
  };

  private readonly handleColorPickerMouseDown = (event: MouseEvent): void => {
    const bounds = this.colorPickerCard.getBoundingClientRect();
    this.colorPickerDragging = {
      offsetX: event.clientX - bounds.left,
      offsetY: event.clientY - bounds.top
    };
    event.preventDefault();
    globalThis.addEventListener("mousemove", this.handleColorPickerMouseMove);
    globalThis.addEventListener("mouseup", this.handleColorPickerMouseUp);
  };

  private readonly handleColorPickerMouseMove = (event: MouseEvent): void => {
    if (!this.colorPickerDragging || !this.activeColorPicker) {
      return;
    }

    this.activeColorPicker.left = `${Math.max(0, event.clientX - this.colorPickerDragging.offsetX)}px`;
    this.activeColorPicker.top = `${Math.max(0, event.clientY - this.colorPickerDragging.offsetY)}px`;
    this.activeColorPicker.transform = "none";
    this.colorPickerCard.style.left = this.activeColorPicker.left;
    this.colorPickerCard.style.top = this.activeColorPicker.top;
    this.colorPickerCard.style.transform = this.activeColorPicker.transform;
  };

  private readonly handleColorPickerMouseUp = (): void => {
    this.colorPickerDragging = undefined;
    globalThis.removeEventListener("mousemove", this.handleColorPickerMouseMove);
    globalThis.removeEventListener("mouseup", this.handleColorPickerMouseUp);
  };

  private readonly openColorPicker = (input: HTMLInputElement, label: string): void => {
    const currentColor = input.value || (input === this.fillColorInput ? "#FFFFFF" : "#000000");
    const { r, g, b } = this.hexToRgb(currentColor);
    const hsv = this.rgbToHsv(r, g, b);
    const previousState = this.activeColorPicker;
    this.activeColorPicker = {
      input,
      label,
      hue: hsv.hue,
      saturation: hsv.saturation,
      value: hsv.value,
      left: previousState?.left ?? "50%",
      top: previousState?.top ?? "50%",
      transform: previousState?.transform ?? "translate(-50%, -50%)"
    };
    this.renderColorPickerCard();
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

  private openNotePanel(sheetId: string, row: number, col: number): void {
    this.noteEditorCell = { sheetId, row, col };
    this.noteInput.value = this.engine.getCellNote(sheetId, row, col) ?? "";
    this.noteRemoveButton.disabled = this.noteInput.value.length === 0;
    this.commentInput.value = "";
    this.renderCommentThreads();
    this.notePanel.hidden = false;
    this.noteInput.focus();
  }

  private closeNotePanel(): void {
    this.noteEditorCell = undefined;
    this.notePanel.hidden = true;
    this.noteInput.value = "";
    this.commentInput.value = "";
    this.commentList.replaceChildren();
  }

  private getCommentAuthor(): CommentAuthor {
    return this.options.comments?.author ?? this.options.collaboration?.user ?? { id: "local", name: "Você" };
  }

  private createCommentAction(label: string, action: string, commentId: string): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = label;
    button.dataset.commentAction = action;
    button.dataset.commentId = commentId;
    return button;
  }

  private renderCommentThreads(): void {
    const target = this.noteEditorCell;
    const fragment = document.createDocumentFragment();
    if (!target) {
      this.commentList.replaceChildren();
      return;
    }
    for (const comment of this.engine.getCellComments(target.sheetId, target.row, target.col)) {
      const thread = document.createElement("article");
      thread.className = "excelsior-comment-thread";
      thread.dataset.commentThread = comment.id;
      const header = document.createElement("div");
      header.className = "excelsior-comment-header";
      const author = document.createElement("strong");
      author.textContent = comment.author.name?.trim() || "Participante";
      const status = document.createElement("span");
      status.textContent = comment.resolved ? "Resolvido" : "Aberto";
      header.append(author, status);
      const content = document.createElement("p");
      content.textContent = comment.content;
      thread.append(header, content);
      for (const reply of comment.replies) {
        const replyElement = document.createElement("div");
        replyElement.className = "excelsior-comment-reply";
        const replyAuthor = document.createElement("strong");
        replyAuthor.textContent = reply.author.name?.trim() || "Participante";
        const replyContent = document.createElement("span");
        replyContent.textContent = reply.content;
        replyElement.append(replyAuthor, replyContent);
        thread.append(replyElement);
      }
      const replyInput = document.createElement("textarea");
      replyInput.className = "excelsior-note-input excelsior-comment-reply-input";
      replyInput.dataset.commentReply = comment.id;
      replyInput.maxLength = this.engine.getSnapshot().settings.maxCellLength;
      replyInput.setAttribute("aria-label", `Responder a ${comment.author.name?.trim() || "comentário"}`);
      const actions = document.createElement("div");
      actions.className = "excelsior-comment-actions";
      actions.append(
        this.createCommentAction("Responder", "reply", comment.id),
        this.createCommentAction(comment.resolved ? "Reabrir" : "Resolver", comment.resolved ? "reopen" : "resolve", comment.id),
        this.createCommentAction("Excluir", "delete", comment.id)
      );
      thread.append(replyInput, actions);
      fragment.append(thread);
    }
    this.commentList.replaceChildren(fragment);
  }

  private updateLocalPresence(): void {
    const collaboration = this.options.collaboration;
    if (!collaboration) {
      return;
    }
    const sheet = this.engine.getActiveSheet();
    try {
      const presence = this.engine.updatePresence({
        user: collaboration.user,
        cursor: { sheetId: sheet.id, ...sheet.selection.end },
        selection: { sheetId: sheet.id, range: sheet.selection },
        metadata: { color: SAFE_PRESENCE_COLOR.test(collaboration.color ?? "") ? collaboration.color : DEFAULT_PRESENCE_COLOR }
      });
      this.localPresenceClientId = presence.clientId;
    } catch (error) {
      if (!(error instanceof SpreadsheetOperationError) || error.details.code !== "CORE_COLLABORATION_REQUIRED") {
        throw error;
      }
    }
  }

  private saveNoteFromPanel(note?: string): void {
    const target = this.noteEditorCell;
    if (!target) {
      return;
    }
    this.engine.setCellNote({ ...target, note });
    this.closeNotePanel();
    this.render();
    this.focus();
  }

  private readonly handleNotePanelClick = (event: Event): void => {
    const commentButton = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-comment-action]");
    const target = this.noteEditorCell;
    if (commentButton && target) {
      const action = commentButton.dataset.commentAction;
      const commentId = commentButton.dataset.commentId;
      if (action === "create") {
        this.engine.createCellComment({ ...target, comment: { author: this.getCommentAuthor(), content: this.commentInput.value } });
        this.commentInput.value = "";
      } else if (commentId && action === "reply") {
        const input = this.commentList.querySelector<HTMLTextAreaElement>(`[data-comment-reply='${commentId}']`);
        this.engine.replyToCellComment({ ...target, commentId, reply: { author: this.getCommentAuthor(), content: input?.value ?? "" } });
      } else if (commentId && (action === "resolve" || action === "reopen")) {
        this.engine.resolveCellComment({ ...target, commentId, resolved: action === "resolve" });
      } else if (commentId && action === "delete") {
        this.engine.deleteCellComment({ ...target, commentId });
      }
      this.renderCommentThreads();
      this.requestRender();
      return;
    }
    const button = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-note-action]");
    if (!button) {
      return;
    }
    if (button.dataset.noteAction === "save") {
      this.saveNoteFromPanel(this.noteInput.value);
    } else if (button.dataset.noteAction === "remove") {
      this.saveNoteFromPanel();
    } else if (button.dataset.noteAction === "close") {
      this.closeNotePanel();
      this.focus();
    }
  };

  private readonly handleNotePanelKeyDown = (event: KeyboardEvent): void => {
    if (event.key === "Escape") {
      event.preventDefault();
      this.closeNotePanel();
      this.focus();
    } else if (event.target === this.noteInput && event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      this.saveNoteFromPanel(this.noteInput.value);
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
      case "cell-note":
        this.openNotePanel(sheet.id, activeAddress.row, activeAddress.col);
        return;
      case "insert-image":
        this.imageFileInput.click();
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
      case "apply-local-filter":
        this.applyLocalFilterForActiveColumn();
        this.focus();
        return;
      case "clear-local-filters":
        this.clearLocalFilters();
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
      case "underline":
        this.applyStyleToSelection({ underline: !activeCell?.style?.underline });
        break;
      case "strike":
        this.applyStyleToSelection({ strike: !activeCell?.style?.strike });
        break;
      case "clear-format":
        this.applyStyleToSelection({}, "replace");
        break;
      case "currency-format":
        this.applyStyleToSelection({ format: "R$ #,##0.00" });
        break;
      case "percentage-format":
        this.applyStyleToSelection({ format: "0.00%" });
        break;
      case "number-decrease":
      case "number-increase": {
        const currentDecimals = activeCell?.style?.format?.match(/\.(0+)/)?.[1].length ?? 0;
        const decimals = Math.max(0, Math.min(9, currentDecimals + (action === "number-increase" ? 1 : -1)));
        this.applyStyleToSelection({ format: decimals ? `#,##0.${"0".repeat(decimals)}` : "#,##0" });
        break;
      }
      case "text-color":
        this.openColorPicker(this.textColorInput, this.messages.textColor);
        return;
      case "border-color":
        this.openColorPicker(this.borderColorInput, this.messages.borderColor);
        return;
      case "fill-color":
        this.openColorPicker(this.fillColorInput, this.messages.fillColor);
        return;
      case "format-painter":
        this.formatPainterStyle = activeCell?.style ? cloneSerializable(activeCell.style) : {};
        this.render();
        this.focus();
        return;
      case "confirm-color":
        if (this.pendingColorStyle) {
          this.applyStyleToSelection(this.pendingColorStyle);
          this.pendingColorStyle = undefined;
          this.render();
          this.focus();
        }
        return;
      case "wrap":
        this.applyStyleToSelection({
          wrap: !activeCell?.style?.wrap
        });
        break;
      case "overflow": {
        const modes: Array<NonNullable<CellStyle["overflow"]>> = ["clip", "ellipsis", "visible"];
        const currentIndex = modes.indexOf(activeCell?.style?.overflow ?? "clip");
        this.applyStyleToSelection({ overflow: modes[(currentIndex + 1) % modes.length] });
        break;
      }
      case "align-left":
        this.applyStyleToSelection({ align: "left" });
        break;
      case "align-center":
        this.applyStyleToSelection({ align: "center" });
        break;
      case "align-right":
        this.applyStyleToSelection({ align: "right" });
        break;
      case "align-top":
        this.applyStyleToSelection({ alignVertical: "top" });
        break;
      case "align-middle":
        this.applyStyleToSelection({ alignVertical: "center" });
        break;
      case "align-bottom":
        this.applyStyleToSelection({ alignVertical: "bottom" });
        break;
      case "rotate-clockwise":
        this.applyStyleToSelection({ rotation: 45 });
        break;
      case "rotate-counterclockwise":
        this.applyStyleToSelection({ rotation: -45 });
        break;
      case "rotate-none":
        this.applyStyleToSelection({ rotation: 0 });
        break;
      case "border-all":
        this.applyBorder("all");
        break;
      case "border-top":
      case "border-right":
      case "border-bottom":
      case "border-left":
        this.applyBorder(action.slice("border-".length) as "top" | "right" | "bottom" | "left");
        break;
      case "border-none":
        this.applyBorder("none");
        break;
      case "insert-link":
        this.openToolbarTool("link");
        return;
      case "split-column":
        this.openToolbarTool("split");
        return;
      case "export-svg":
        this.exportVisibleGridSvg();
        break;
      case "zoom-in":
        this.updateSheetZoom(0.1);
        return;
      case "zoom-out":
        this.updateSheetZoom(-0.1);
        return;
      case "zoom-reset":
        this.updateSheetZoom();
        return;
      case "data-validation":
        this.openToolbarTool("validation");
        return;
      case "conditional-formatting":
        this.openToolbarTool("conditional");
        return;
      case "find-special":
        this.openToolbarTool("find-special");
        return;
      case "quick-sum":
        this.applyQuickSum();
        break;
      case "freeze-rows":
        this.engine.freezeRows(sheet.id, activeAddress.row + 1);
        break;
      case "freeze-columns":
        this.engine.freezeColumns(sheet.id, activeAddress.col + 1);
        break;
      case "unfreeze":
        this.engine.freezeRows(sheet.id, 0);
        this.engine.freezeColumns(sheet.id, 0);
        break;
      case "merge":
        this.engine.mergeCells({
          sheetId: sheet.id,
          start: sheet.selection.start,
          end: sheet.selection.end
        });
        break;
      case "merge-horizontal":
        this.mergeSelectionByAxis("horizontal");
        break;
      case "merge-vertical":
        this.mergeSelectionByAxis("vertical");
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

  private readonly handleImageFileChange = (): void => {
    const file = this.imageFileInput.files?.[0];
    this.imageFileInput.value = "";
    if (!file || !["image/png", "image/jpeg", "image/gif", "image/webp"].includes(file.type)) {
      return;
    }
    const reader = new FileReader();
    reader.addEventListener("load", () => {
      if (typeof reader.result !== "string") {
        return;
      }
      const sheet = this.engine.getActiveSheet();
      const address = this.getActiveAddress(sheet);
      this.engine.createImage({
        sheetId: sheet.id,
        image: {
          src: reader.result,
          alt: file.name,
          position: {
            fromCell: cellAddressToLabel(address),
            offsetX: 8,
            offsetY: 8,
            width: 320,
            height: 220
          }
        }
      });
      this.requestRender();
    }, { once: true });
    reader.readAsDataURL(file);
  };

  private readonly handleToolbarInput = (event: Event): void => {
    const fontFamilySelect = (event.target as HTMLElement | null)?.closest<HTMLSelectElement>("[data-font-family]");
    if (fontFamilySelect) {
      this.applyStyleToSelection({ fontFamily: fontFamilySelect.value });
      return;
    }
    const fontSizeSelect = (event.target as HTMLElement | null)?.closest<HTMLSelectElement>("[data-font-size]");
    if (fontSizeSelect) {
      this.applyStyleToSelection({ fontSize: Number(fontSizeSelect.value) });
      return;
    }
    const formatSelect = (event.target as HTMLElement | null)?.closest<HTMLSelectElement>("[data-number-format]");
    if (formatSelect) {
      this.applyStyleToSelection({ format: formatSelect.value });
      return;
    }
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

    this.editor.type = validationRule && isDateValidationRule(validationRule) ? "date" : "text";
    this.editor.min = validationRule && isDateValidationRule(validationRule) ? validationRule.min ?? "" : "";
    this.editor.max = validationRule && isDateValidationRule(validationRule) ? validationRule.max ?? "" : "";
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
      validationMessage ??
      rowModelError ??
      pivotStatus ??
      (this.hasPendingRowModelRequests(sheet.id) ? this.messages.loadingRows : "");
    this.statusMessage.classList.toggle(
      "is-error",
      Boolean((validationMessage && !this.validationFeedback?.isWarning) || rowModelError) || pivotStatusIsError
    );
    if (document.activeElement !== this.formulaInput) {
      this.formulaInput.value = value;
    }
    this.renderFindReplacePanel();
    this.renderPivotPanel();
  }

  private clearValidationFeedback(): void {
    this.validationFeedback = undefined;
  }

  private renderChrome(columnsToRender: number[]): void {
    const sheet = this.engine.getActiveSheet();
    const selectionRange = this.getResolvedSelectionRange(sheet);
    const activeAddress = this.getActiveAddress(sheet);
    const fullSheetSelected =
      selectionRange.start.row === 0 &&
      selectionRange.start.col === 0 &&
      selectionRange.end.row === this.getResolvedRowCount(sheet) - 1 &&
      selectionRange.end.col === sheet.columnCount - 1;
    this.chrome.replaceChildren();
    const fragment = document.createDocumentFragment();
    const corner = document.createElement("button");
    corner.type = "button";
    corner.className = "excelsior-corner";
    if (fullSheetSelected) {
      corner.classList.add("is-active");
    }
    corner.dataset.cornerAction = "select-all";
    corner.setAttribute("aria-label", `Selecionar toda a ${this.messages.gridLabel.toLowerCase()}`);
    corner.title = "Selecionar toda a planilha";

    const columnStrip = document.createElement("div");
    columnStrip.className = "excelsior-column-strip";
    const remoteRequestModel = this.getRemoteRequestModel(sheet.id);
    const fixedColumnCount = Math.max(
      this.engine.getFrozenPane(sheet.id).columns,
      this.engine.getSplitPane(sheet.id)?.verticalColumn ?? 0
    );
    const renderedColumns: number[] = [];
    for (let col = 0; col < sheet.columnCount; col += 1) {
      if (!this.isColumnHidden(sheet, col)) {
        renderedColumns.push(col);
      }
    }
    this.renderedHeaderColumns = new Set(renderedColumns);

    for (const col of renderedColumns) {
      const header = document.createElement("div");
      header.className = "excelsior-column-header";
      header.id = this.getColumnHeaderElementId(sheet.id, col);
      header.setAttribute("role", "columnheader");
      header.setAttribute("aria-colindex", String(col + 1));
      header.setAttribute("aria-label", this.getColumnHeaderAccessibilityLabel(col));
      header.dataset.columnHeaderCol = String(col);
      header.tabIndex = 0;
      if (col >= selectionRange.start.col && col <= selectionRange.end.col) {
        header.classList.add("is-selected");
      }
      if (col === activeAddress.col) {
        header.classList.add("is-active");
      }
      if (remoteRequestModel !== undefined) {
        const direction = this.getActiveRemoteSortDirection(sheet.id, col);
        header.setAttribute(
          "aria-sort",
          direction === "asc" ? "ascending" : direction === "desc" ? "descending" : "none"
        );
      }
      header.style.width = `${this.getColumnWidth(sheet, col)}px`;
      header.style.flex = "0 0 auto";
      if (col < fixedColumnCount) {
        header.style.zIndex = "1";
      } else if (this.viewport.scrollLeft > 0) {
        header.style.transform = `translateX(-${this.viewport.scrollLeft}px)`;
      }
      header.textContent = columnIndexToLabel(col);
      columnStrip.append(header);
    }
    const headerRow = document.createElement("div");
    headerRow.className = "excelsior-column-header-row";
    headerRow.append(corner, columnStrip);

    fragment.append(this.renderToolbar(), headerRow);
    this.chrome.append(fragment);
  }

  private renderRowHeaders(
    sheet: ReturnType<WorkbookEngine["getActiveSheet"]>,
    rowOffsets: number[],
    rowsToRender: number[],
    selectionRange: CellRange,
    activeAddress: CellAddress
  ): void {
    const fragment = document.createDocumentFragment();
    const frozenRows = this.engine.getFrozenPane(sheet.id).rows;

    for (const row of rowsToRender) {
      if (this.isRowHidden(sheet, row)) {
        continue;
      }

      const header = document.createElement("div");
      header.className = "excelsior-row-header";
      if (row >= selectionRange.start.row && row <= selectionRange.end.row) {
        header.classList.add("is-selected");
      }
      if (row === activeAddress.row) {
        header.classList.add("is-active");
      }
      if (row < frozenRows) {
        header.classList.add("is-frozen");
      }
      header.dataset.rowHeaderRow = String(row);
      header.style.top = `${this.getFrozenAdjustedTop(sheet.id, rowOffsets, row) - this.viewport.scrollTop}px`;
      header.style.height = `${this.getRowHeight(sheet, row)}px`;
      header.textContent = String(row + 1);
      fragment.append(header);
    }

    this.rowHeaders.replaceChildren(fragment);
  }

  private getToolbarActionIcon(action: string): string {
    const iconByAction: Record<string, string> = {
      undo: "↶",
      redo: "↷",
      "sort-asc": "A↑",
      "sort-desc": "Z↓",
      "group-column": "⊞",
      "pivot-column": "⥁",
      "aggregate-sum": "Σ",
      "aggregate-avg": "AVG",
      "aggregate-min": "MIN",
      "aggregate-max": "MAX",
      "aggregate-count": "#",
      "clear-column-query": "⌫",
      "clear-format": "◇",
      "currency-format": "R$",
      "percentage-format": "%",
      "number-decrease": ".0←",
      "number-increase": ".00→",
      bold: "B",
      italic: "I",
      strike: "S",
      "text-color": "A",
      "border-color": "⊞",
      "fill-color": "▣",
      "format-painter": "F",
      wrap: "↵",
      "align-left": "≡←",
      "align-center": "≡",
      "align-right": "→≡",
      "align-top": "↥",
      "align-middle": "↕",
      "align-bottom": "↧",
      "rotate-clockwise": "45°",
      "rotate-counterclockwise": "-45°",
      "rotate-none": "0°",
      "border-all": "▦",
      "border-top": "▔",
      "border-right": "▏",
      "border-bottom": "▁",
      "border-left": "▕",
      "border-none": "□",
      "insert-link": "L",
      "split-column": "C|C",
      "export-svg": "SVG",
      "zoom-in": "+",
      "zoom-out": "−",
      "zoom-reset": "100%",
      "data-validation": "✓",
      "conditional-formatting": "CF",
      "find-special": "⌕!",
      "quick-sum": "Σ",
      "freeze-rows": "▤",
      "freeze-columns": "▥",
      unfreeze: "□",
      merge: "⇆",
      "merge-horizontal": "⇔",
      "merge-vertical": "⇕",
      unmerge: "⇅",
      "insert-row": "+R",
      "delete-row": "-R",
      "insert-column": "+C",
      "delete-column": "-C",
      "create-pivot": "◫",
      "find-replace": "⌕",
      "cell-note": "N",
      "add-sheet": "+"
    };

    return iconByAction[action] ?? "•";
  }

  private renderToolbar(): HTMLElement {
    this.toolbar.replaceChildren();
    const sheet = this.engine.getActiveSheet();
    const activeAddress = this.getActiveAddress(sheet);
    const activeCell = this.engine.getCell(sheet.id, activeAddress.row, activeAddress.col);
    const remoteRequestModel = this.getRemoteRequestModel(sheet.id);
    const localQuery = this.engine.getRowModel(sheet.id).kind === "clientSide"
      ? this.engine.getClientSideQuery(sheet.id)
      : undefined;
    const pivotSourceRange = this.getPivotSourceRange(sheet);
    const activeRemoteSort = this.getActiveRemoteSortDirection(sheet.id, activeAddress.col);
    const activeLocalSort = localQuery?.sort.find((item) => item.column === activeAddress.col)?.direction;
    const toggleStates: Partial<Record<string, boolean>> = {
      bold: activeCell?.style?.fontWeight === "bold",
      italic: activeCell?.style?.fontStyle === "italic",
      underline: activeCell?.style?.underline === true,
      wrap: activeCell?.style?.wrap === true,
      "format-painter": this.formatPainterStyle !== undefined,
      "group-column": this.isActiveRemoteGrouped(sheet.id, activeAddress.col),
      "pivot-column": this.isActiveRemotePivoted(sheet.id, activeAddress.col),
      "aggregate-sum": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "sum"),
      "aggregate-avg": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "avg"),
      "aggregate-min": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "min"),
      "aggregate-max": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "max"),
      "aggregate-count": this.hasActiveRemoteAggregate(sheet.id, activeAddress.col, "count"),
      "sort-asc": (activeRemoteSort ?? activeLocalSort) === "asc",
      "sort-desc": (activeRemoteSort ?? activeLocalSort) === "desc"
    };

    type ToolbarGroupKey = "data" | "font" | "alignment" | "structure";
    const actions: Array<{ action: string; label: string; group: ToolbarGroupKey }> = [
      { action: "undo", label: this.messages.undo, group: "data" },
      { action: "redo", label: this.messages.redo, group: "data" },
      { action: "sort-asc", label: this.messages.sortAscending, group: "data" },
      { action: "sort-desc", label: this.messages.sortDescending, group: "data" },
      { action: "group-column", label: this.messages.groupColumn, group: "data" },
      { action: "pivot-column", label: this.messages.pivotColumn, group: "data" },
      { action: "aggregate-sum", label: this.messages.aggregateSum, group: "data" },
      { action: "aggregate-avg", label: this.messages.aggregateAverage, group: "data" },
      { action: "aggregate-min", label: this.messages.aggregateMin, group: "data" },
      { action: "aggregate-max", label: this.messages.aggregateMax, group: "data" },
      { action: "aggregate-count", label: this.messages.aggregateCount, group: "data" },
      { action: "clear-column-query", label: this.messages.clearColumnQuery, group: "data" },
      { action: "create-pivot", label: this.messages.createPivot, group: "data" },
      { action: "find-replace", label: this.messages.findReplace, group: "data" },
      { action: "cell-note", label: this.messages.cellNote, group: "data" },
      { action: "insert-image", label: "Inserir imagem", group: "structure" },
      { action: "add-sheet", label: this.messages.addSheet, group: "data" },
      { action: "bold", label: this.messages.bold, group: "font" },
      { action: "italic", label: this.messages.italic, group: "font" },
      { action: "underline", label: this.messages.underline, group: "font" },
      { action: "strike", label: "Tachado", group: "font" },
      { action: "clear-format", label: "Limpar formatação", group: "font" },
      { action: "currency-format", label: "Moeda", group: "font" },
      { action: "percentage-format", label: "Percentual", group: "font" },
      { action: "number-decrease", label: "Diminuir casas decimais", group: "font" },
      { action: "number-increase", label: "Aumentar casas decimais", group: "font" },
      { action: "text-color", label: this.messages.textColor, group: "font" },
      { action: "border-color", label: this.messages.borderColor, group: "font" },
      { action: "fill-color", label: this.messages.fillColor, group: "font" },
      { action: "format-painter", label: this.messages.formatPainter, group: "font" },
      { action: "wrap", label: this.messages.wrap, group: "alignment" },
      { action: "overflow", label: `Overflow: ${activeCell?.style?.overflow ?? "clip"}`, group: "alignment" },
      { action: "align-left", label: this.messages.alignLeft, group: "alignment" },
      { action: "align-center", label: this.messages.alignCenter, group: "alignment" },
      { action: "align-right", label: this.messages.alignRight, group: "alignment" },
      { action: "align-top", label: "Alinhar acima", group: "alignment" },
      { action: "align-middle", label: "Alinhar ao meio", group: "alignment" },
      { action: "align-bottom", label: "Alinhar abaixo", group: "alignment" },
      { action: "rotate-clockwise", label: "Girar texto 45 graus", group: "alignment" },
      { action: "rotate-counterclockwise", label: "Girar texto -45 graus", group: "alignment" },
      { action: "rotate-none", label: "Remover rotação", group: "alignment" },
      { action: "border-all", label: "Todas as bordas", group: "font" },
      { action: "border-top", label: "Borda superior", group: "font" },
      { action: "border-right", label: "Borda direita", group: "font" },
      { action: "border-bottom", label: "Borda inferior", group: "font" },
      { action: "border-left", label: "Borda esquerda", group: "font" },
      { action: "border-none", label: "Remover bordas", group: "font" },
      { action: "insert-link", label: "Inserir link", group: "data" },
      { action: "split-column", label: "Dividir coluna", group: "data" },
      { action: "export-svg", label: "Capturar planilha em SVG", group: "data" },
      { action: "zoom-out", label: "Diminuir zoom da planilha", group: "data" },
      { action: "zoom-reset", label: `Zoom ${Math.round(this.sheetZoom * 100)}%`, group: "data" },
      { action: "zoom-in", label: "Aumentar zoom da planilha", group: "data" },
      { action: "data-validation", label: "Validação de dados", group: "data" },
      { action: "conditional-formatting", label: "Formatação condicional", group: "font" },
      { action: "find-special", label: "Localizar células especiais", group: "data" },
      { action: "quick-sum", label: "AutoSoma", group: "data" },
      { action: "freeze-rows", label: "Congelar linhas até a célula", group: "structure" },
      { action: "freeze-columns", label: "Congelar colunas até a célula", group: "structure" },
      { action: "unfreeze", label: "Descongelar painéis", group: "structure" },
      { action: "merge", label: this.messages.merge, group: "alignment" },
      { action: "merge-horizontal", label: "Mesclar horizontalmente", group: "alignment" },
      { action: "merge-vertical", label: "Mesclar verticalmente", group: "alignment" },
      { action: "unmerge", label: this.messages.unmerge, group: "alignment" },
      { action: "insert-row", label: this.messages.insertRow, group: "structure" },
      { action: "delete-row", label: this.messages.deleteRow, group: "structure" },
      { action: "insert-column", label: this.messages.insertColumn, group: "structure" },
      { action: "delete-column", label: this.messages.deleteColumn, group: "structure" }
    ];
    const remoteOnlyActions = new Set([
      "group-column",
      "pivot-column",
      "aggregate-sum",
      "aggregate-avg",
      "aggregate-min",
      "aggregate-max",
      "aggregate-count",
    ]);

    const groupOrder: Array<{ key: ToolbarGroupKey; label: string }> = [
      { key: "data", label: this.messages.toolbarDataGroup },
      { key: "font", label: this.messages.toolbarFontGroup },
      { key: "alignment", label: this.messages.toolbarAlignmentGroup },
      { key: "structure", label: this.messages.toolbarStructureGroup }
    ];
    const ribbon = document.createElement("div");
    ribbon.className = "excelsior-toolbar-ribbon";
    const groupControls = new Map<ToolbarGroupKey, HTMLElement>();

    for (const groupItem of groupOrder) {
      const group = document.createElement("section");
      group.className = "excelsior-toolbar-group";
      group.dataset.toolbarGroup = groupItem.key;

      const controls = document.createElement("div");
      controls.className = "excelsior-toolbar-group-controls";

      const label = document.createElement("span");
      label.className = "excelsior-toolbar-group-label";
      label.textContent = groupItem.label;

      group.append(controls, label);
      groupControls.set(groupItem.key, controls);
      ribbon.append(group);
    }

    if (remoteRequestModel !== undefined) {
      const filterField = document.createElement("label");
      filterField.className = "excelsior-toolbar-filter is-inline";
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
      groupControls.get("data")?.append(filterField);
    } else if (this.engine.getRowModel(sheet.id).kind === "clientSide") {
      const activeFilter = localQuery?.filters.find((item) => item.column === activeAddress.col);
      const filterField = document.createElement("div");
      filterField.className = "excelsior-toolbar-filter is-inline";
      filterField.dataset.localFilterPanel = "true";
      const typeSelect = document.createElement("select");
      typeSelect.dataset.localFilterType = "true";
      typeSelect.setAttribute("aria-label", "Tipo do filtro local");
      for (const item of [
        { value: "text", label: "Texto" },
        { value: "number", label: "Número" },
        { value: "date", label: "Data" }
      ] as const) {
        typeSelect.add(new Option(item.label, item.value, false, activeFilter?.type === item.value));
      }
      const operatorSelect = document.createElement("select");
      operatorSelect.dataset.localFilterOperator = "true";
      operatorSelect.setAttribute("aria-label", "Operador do filtro local");
      const operatorItems = [
        { value: "equals", label: "Igual a" },
        { value: "contains", label: "Contém" },
        { value: "startsWith", label: "Começa com" },
        { value: "gt", label: "Maior que" },
        { value: "gte", label: "Maior ou igual" },
        { value: "lt", label: "Menor que" },
        { value: "lte", label: "Menor ou igual" },
        { value: "between", label: "Entre" }
      ] as const;
      const syncOperatorOptions = () => {
        const previousOperator = operatorSelect.value || activeFilter?.operator;
        const allowedOperators = typeSelect.value === "text"
          ? operatorItems.filter((item) => ["equals", "contains", "startsWith"].includes(item.value))
          : operatorItems.filter((item) => !["contains", "startsWith"].includes(item.value));
        operatorSelect.replaceChildren(...allowedOperators.map((item) => new Option(item.label, item.value)));
        operatorSelect.value = allowedOperators.some((item) => item.value === previousOperator)
          ? previousOperator ?? "equals"
          : "equals";
      };
      syncOperatorOptions();
      const valueInput = document.createElement("input");
      valueInput.dataset.localFilterValue = "true";
      valueInput.className = "excelsior-toolbar-input";
      valueInput.inputMode = typeSelect.value === "number" ? "decimal" : "text";
      valueInput.placeholder = `${this.messages.filterColumn} ${this.getRemoteRequestField(activeAddress.col)}`;
      valueInput.value = activeFilter === undefined ? "" : String(activeFilter.value);
      const valueToInput = document.createElement("input");
      valueToInput.dataset.localFilterValueTo = "true";
      valueToInput.className = "excelsior-toolbar-input";
      valueToInput.placeholder = "Até";
      valueToInput.value = activeFilter?.valueTo === undefined ? "" : String(activeFilter.valueTo);
      const syncValueToState = () => {
        valueToInput.disabled = operatorSelect.value !== "between";
        valueToInput.inputMode = typeSelect.value === "number" ? "decimal" : "text";
        valueToInput.setAttribute("aria-label", valueToInput.disabled ? "Valor final disponível para o operador Entre" : "Valor final do intervalo");
      };
      operatorSelect.addEventListener("change", syncValueToState);
      typeSelect.addEventListener("change", () => {
        syncOperatorOptions();
        valueInput.inputMode = typeSelect.value === "number" ? "decimal" : "text";
        syncValueToState();
      });
      syncValueToState();
      const applyButton = document.createElement("button");
      applyButton.type = "button";
      applyButton.className = "excelsior-toolbar-button";
      applyButton.dataset.action = "apply-local-filter";
      applyButton.textContent = "Aplicar filtro";
      const clearButton = document.createElement("button");
      clearButton.type = "button";
      clearButton.className = "excelsior-toolbar-button";
      clearButton.dataset.action = "clear-local-filters";
      clearButton.textContent = "Zerar filtros";
      filterField.append(typeSelect, operatorSelect, valueInput, valueToInput, applyButton, clearButton);
      groupControls.get("data")?.append(filterField);
    }

    for (const item of actions) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "excelsior-toolbar-button";
      button.dataset.action = item.action;
      button.textContent = item.label;
      button.setAttribute("aria-label", item.label);
      button.dataset.icon = this.getToolbarActionIcon(item.action);
      button.title = item.label;
      if (remoteOnlyActions.has(item.action) && remoteRequestModel === undefined) {
        button.disabled = true;
      }
      if (item.action === "create-pivot" && pivotSourceRange === undefined) {
        button.disabled = true;
      }
      if (item.action in toggleStates) {
        button.setAttribute("aria-pressed", String(Boolean(toggleStates[item.action])));
      }
      groupControls.get(item.group)?.append(button);
    }

    const fontFamilySelect = document.createElement("select");
    fontFamilySelect.dataset.fontFamily = "true";
    fontFamilySelect.setAttribute("aria-label", this.messages.fontFamily);
    for (const family of ["Arial", "Calibri", "Georgia", "Tahoma", "Verdana", "Courier New"]) {
      fontFamilySelect.add(new Option(family, family, false, (activeCell?.style?.fontFamily ?? "Arial") === family));
    }
    const fontSizeSelect = document.createElement("select");
    fontSizeSelect.dataset.fontSize = "true";
    fontSizeSelect.setAttribute("aria-label", this.messages.fontSize);
    for (const size of [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32]) {
      fontSizeSelect.add(new Option(String(size), String(size), false, (activeCell?.style?.fontSize ?? 12) === size));
    }
    const formatSelect = document.createElement("select");
    formatSelect.dataset.numberFormat = "true";
    formatSelect.setAttribute("aria-label", "Formato da célula");
    for (const item of [
      { label: "Geral", value: "General" },
      { label: "Número", value: "#,##0.00" },
      { label: "Moeda", value: "R$ #,##0.00" },
      { label: "Percentual", value: "0.00%" },
      { label: "Data", value: "dd/mm/yyyy" }
    ]) {
      formatSelect.add(new Option(item.label, item.value, false, (activeCell?.style?.format ?? "General") === item.value));
    }
    groupControls.get("font")?.prepend(formatSelect, fontFamilySelect, fontSizeSelect);

    this.toolbar.append(ribbon);
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
    this.visualObjectSurfaceMetrics = {
      sheetId: sheet.id,
      rowOffsets,
      colOffsets,
      rowCount: resolvedRowCount,
      colCount: sheet.columnCount
    };
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
    const remotePresences = this.engine
      .getPresences()
      .filter((presence) => presence.clientId !== this.localPresenceClientId);
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
    this.renderChrome(columnsToRender);
    this.renderFormulaBar();
    this.renderSheetTabs();
    this.renderRowHeaders(sheet, rowOffsets, rowsToRender, selectionRange, activeAddress);
    this.renderSplitPanes(sheet, rowOffsets, colOffsets);

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
        if (this.renderedHeaderColumns.has(cellRange.start.col)) {
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
        cell.style.left = `${ROW_HEADER_WIDTH + this.getFrozenAdjustedLeft(sheet.id, colOffsets, cellRange.start.col)}px`;
        cell.style.width = `${getSpanSize(colOffsets, cellRange.start.col, cellRange.end.col)}px`;
        cell.style.height = `${getSpanSize(rowOffsets, cellRange.start.row, cellRange.end.row)}px`;
        const cellStyle = this.getCellStyle(sheet, cellRange.start.row, cellRange.start.col);
        this.applyCellPresentation(cell, cellStyle);
        this.renderCellContent(cell, sheet.id, cellRange.start.row, cellRange.start.col, rowModelRowsByIndex.get(cellRange.start.row));
        const content = cell.querySelector<HTMLElement>(".excelsior-cell-content");
        if (content && cellStyle?.rotation) {
          content.style.transform = `rotate(${cellStyle.rotation}deg)`;
        }

        for (const presence of remotePresences) {
          const color = getPresenceColor(presence);
          if (presence.selection?.sheetId === sheet.id && rangesOverlap(cellRange, presence.selection.range)) {
            const remoteSelection = document.createElement("span");
            remoteSelection.className = "excelsior-remote-selection";
            remoteSelection.dataset.remoteSelection = presence.clientId;
            remoteSelection.setAttribute("aria-hidden", "true");
            remoteSelection.style.setProperty("--excelsior-presence-color", color);
            cell.append(remoteSelection);
          }
          if (
            presence.cursor?.sheetId === sheet.id &&
            isWithinRange(presence.cursor.row, presence.cursor.col, cellRange)
          ) {
            const remoteCursor = document.createElement("span");
            const name = getPresenceName(presence);
            remoteCursor.className = "excelsior-remote-cursor";
            remoteCursor.dataset.remoteCursor = presence.clientId;
            remoteCursor.textContent = name;
            remoteCursor.setAttribute("role", "note");
            remoteCursor.setAttribute("aria-label", `Cursor remoto de ${name}`);
            remoteCursor.style.setProperty("--excelsior-presence-color", color);
            cell.append(remoteCursor);
          }
        }

        if (model?.note || model?.comments?.length) {
          cell.classList.add("has-note");
          const noteIndicator = document.createElement("button");
          noteIndicator.type = "button";
          noteIndicator.className = "excelsior-cell-note-indicator";
          noteIndicator.dataset.cellNote = "true";
          noteIndicator.setAttribute("aria-label", this.messages.cellNoteIndicator);
          noteIndicator.title = model.note ?? `${model?.comments?.length ?? 0} comentário(s)`;
          cell.append(noteIndicator);
        }

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
    this.renderVisualObjects(sheet, rowOffsets, colOffsets);

    if (this.editingCell) {
      this.positionEditor(this.editingCell.row, this.editingCell.col);
    }
  };

  private resolveVisualObjectRect(
    position: WorksheetObjectPosition,
    rowOffsets: number[],
    colOffsets: number[],
    rowCount: number,
    colCount: number
  ): VisualObjectRect {
    const anchor = cellLabelToAddress(position.fromCell);
    const row = clampNumeric(anchor.row, 0, Math.max(0, rowCount - 1));
    const col = clampNumeric(anchor.col, 0, Math.max(0, colCount - 1));
    return {
      left: ROW_HEADER_WIDTH + (colOffsets[col] ?? 0) + position.offsetX,
      top: (rowOffsets[row] ?? 0) + position.offsetY,
      width: position.width,
      height: position.height
    };
  }

  private resolveWorksheetObjectPositionFromRect(
    rect: VisualObjectRect,
    metrics: VisualObjectSurfaceMetrics,
    previous: WorksheetObjectPosition
  ): WorksheetObjectPosition {
    const findAnchorIndex = (offsets: number[], coordinate: number, count: number): number => {
      let index = 0;
      while (index + 1 < count && (offsets[index + 1] ?? Number.POSITIVE_INFINITY) <= coordinate) {
        index += 1;
      }
      return clampNumeric(index, 0, Math.max(0, count - 1));
    };
    const relativeLeft = Math.max(0, rect.left - ROW_HEADER_WIDTH);
    const relativeTop = Math.max(0, rect.top);
    const fromCol = findAnchorIndex(metrics.colOffsets, relativeLeft, metrics.colCount);
    const fromRow = findAnchorIndex(metrics.rowOffsets, relativeTop, metrics.rowCount);
    const toCol = findAnchorIndex(metrics.colOffsets, relativeLeft + rect.width, metrics.colCount);
    const toRow = findAnchorIndex(metrics.rowOffsets, relativeTop + rect.height, metrics.rowCount);
    return {
      fromCell: cellAddressToLabel({ row: fromRow, col: fromCol }),
      toCell: cellAddressToLabel({ row: toRow, col: toCol }),
      offsetX: relativeLeft - (metrics.colOffsets[fromCol] ?? 0),
      offsetY: relativeTop - (metrics.rowOffsets[fromRow] ?? 0),
      width: rect.width,
      height: rect.height,
      zIndex: previous.zIndex
    };
  }

  private renderVisualObjects(
    sheet: SheetModel,
    rowOffsets: number[],
    colOffsets: number[]
  ): void {
    const fragment = document.createDocumentFragment();
    const imageIds = new Set<string>();
    for (const image of this.engine.getImages(sheet.id).filter((item) => item.state.visible !== false)) {
      imageIds.add(image.id);
      const rect = this.resolveVisualObjectRect(image.position, rowOffsets, colOffsets, sheet.rowCount, sheet.columnCount);
      const object = this.imageObjectElementById.get(image.id) ?? this.createImageObjectElement(image);
      object.style.left = `${rect.left}px`;
      object.style.top = `${rect.top}px`;
      object.style.width = `${rect.width}px`;
      object.style.height = `${rect.height}px`;
      object.style.zIndex = String(image.position.zIndex);
      object.classList.toggle("is-selected", image.state.selected);
      object.classList.toggle("is-locked", image.state.locked);
      object.setAttribute("aria-selected", String(image.state.selected));
      const picture = object.querySelector<HTMLImageElement>("[data-image-content='true']");
      if (picture) {
        picture.src = image.src;
        picture.alt = image.alt;
        picture.style.objectFit = image.style?.objectFit ?? "contain";
        picture.style.opacity = String(image.style?.opacity ?? 1);
      }
      fragment.append(object);
    }
    for (const [imageId, object] of this.imageObjectElementById) {
      if (!imageIds.has(imageId)) {
        object.remove();
        this.imageObjectElementById.delete(imageId);
      }
    }

    const widgetIds = new Set<string>();
    for (const widget of this.engine.getWidgets(sheet.id).filter((item) => item.state.visible !== false)) {
      widgetIds.add(widget.id);
      const rect = this.resolveVisualObjectRect(widget.position, rowOffsets, colOffsets, sheet.rowCount, sheet.columnCount);
      const object = this.widgetObjectElementById.get(widget.id) ?? this.createWidgetObjectElement(widget);
      object.style.left = `${rect.left}px`;
      object.style.top = `${rect.top}px`;
      object.style.width = `${rect.width}px`;
      object.style.height = `${rect.height}px`;
      object.style.zIndex = String(widget.position.zIndex);
      object.classList.toggle("is-selected", widget.state.selected);
      object.classList.toggle("is-locked", widget.state.locked);
      object.setAttribute("aria-selected", String(widget.state.selected));
      this.renderWidgetObject(widget);
      fragment.append(object);
    }
    for (const [widgetId, object] of this.widgetObjectElementById) {
      if (!widgetIds.has(widgetId)) {
        this.destroyWidgetRuntime(widgetId);
        object.remove();
        this.widgetObjectElementById.delete(widgetId);
        this.widgetBodyElementById.delete(widgetId);
      }
    }
    this.visualObjectsLayer.replaceChildren(fragment);
  }

  private createImageObjectElement(image: WorksheetImageObject): HTMLElement {
    const object = document.createElement("section");
    object.className = "excelsior-visual-object excelsior-image-object";
    object.dataset.imageId = image.id;
    object.tabIndex = 0;
    object.setAttribute("role", "group");
    object.setAttribute("aria-label", image.alt || "Imagem");
    const header = document.createElement("header");
    header.className = "excelsior-visual-object-header";
    header.dataset.imageMove = "true";
    const title = document.createElement("div");
    title.className = "excelsior-visual-object-title";
    title.textContent = image.alt || "Imagem";
    const actions = document.createElement("div");
    actions.className = "excelsior-visual-object-actions";
    actions.append(
      this.createVisualObjectAction("image", "back", "Enviar imagem para trás", "↓"),
      this.createVisualObjectAction("image", "front", "Trazer imagem para frente", "↑"),
      this.createVisualObjectAction("image", "lock", "Bloquear ou desbloquear imagem", "L"),
      this.createVisualObjectAction("image", "delete", "Excluir imagem", "×", true)
    );
    header.append(title, actions);
    const body = document.createElement("div");
    body.className = "excelsior-visual-object-body";
    const picture = document.createElement("img");
    picture.dataset.imageContent = "true";
    picture.draggable = false;
    body.append(picture);
    const resize = document.createElement("button");
    resize.type = "button";
    resize.className = "excelsior-visual-object-resize";
    resize.dataset.imageResize = "true";
    resize.setAttribute("aria-label", "Redimensionar imagem");
    object.append(header, body, resize);
    this.imageObjectElementById.set(image.id, object);
    return object;
  }

  private createVisualObjectAction(
    kind: "image" | "widget",
    action: string,
    label: string,
    text: string,
    destructive = false
  ): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = destructive ? "excelsior-visual-object-delete" : "excelsior-visual-object-action";
    button.dataset[`${kind}Action`] = action;
    button.setAttribute("aria-label", label);
    button.title = label;
    button.textContent = text;
    return button;
  }

  private createWidgetObjectElement(widget: WorksheetWidgetObject): HTMLElement {
    const object = document.createElement("section");
    object.className = "excelsior-visual-object excelsior-widget-object";
    object.dataset.widgetId = widget.id;
    object.tabIndex = 0;
    object.setAttribute("role", "group");
    object.setAttribute("aria-label", widget.label);
    const header = document.createElement("header");
    header.className = "excelsior-visual-object-header";
    header.dataset.widgetMove = "true";
    const title = document.createElement("div");
    title.className = "excelsior-visual-object-title";
    title.textContent = widget.label;
    const actions = document.createElement("div");
    actions.className = "excelsior-visual-object-actions";
    actions.append(
      this.createVisualObjectAction("widget", "back", `Enviar ${widget.label} para trás`, "↓"),
      this.createVisualObjectAction("widget", "front", `Trazer ${widget.label} para frente`, "↑"),
      this.createVisualObjectAction("widget", "lock", `Bloquear ou desbloquear ${widget.label}`, "L"),
      this.createVisualObjectAction("widget", "delete", `Excluir ${widget.label}`, "×", true)
    );
    header.append(title, actions);
    const body = document.createElement("div");
    body.className = "excelsior-visual-object-body excelsior-widget-object-body";
    body.dataset.widgetBody = "true";
    const resize = document.createElement("button");
    resize.type = "button";
    resize.className = "excelsior-visual-object-resize";
    resize.dataset.widgetResize = "true";
    resize.setAttribute("aria-label", `Redimensionar ${widget.label}`);
    object.append(header, body, resize);
    this.widgetObjectElementById.set(widget.id, object);
    this.widgetBodyElementById.set(widget.id, body);
    return object;
  }

  private renderWidgetObject(widget: WorksheetWidgetObject): void {
    const body = this.widgetBodyElementById.get(widget.id);
    if (!body) return;
    const signature = JSON.stringify({ type: widget.type, label: widget.label, config: widget.config, data: widget.data });
    if (this.widgetRenderSignatureById.get(widget.id) === signature) return;
    this.destroyWidgetRuntime(widget.id);
    body.replaceChildren();
    const renderer = this.options.widgetRenderers?.[widget.type];
    if (!renderer) {
      const placeholder = document.createElement("div");
      placeholder.className = "excelsior-widget-placeholder";
      placeholder.textContent = `Widget não registrado: ${widget.type}`;
      body.append(placeholder);
      this.widgetRenderSignatureById.set(widget.id, signature);
      return;
    }
    try {
      const cleanup = renderer({ host: body, widget });
      if (typeof cleanup === "function") this.widgetCleanupById.set(widget.id, cleanup);
      this.widgetRenderSignatureById.set(widget.id, signature);
    } catch {
      body.replaceChildren();
      const placeholder = document.createElement("div");
      placeholder.className = "excelsior-widget-placeholder";
      placeholder.textContent = `Falha ao renderizar widget: ${widget.type}`;
      body.append(placeholder);
    }
  }

  private destroyWidgetRuntime(widgetId: string): void {
    try {
      this.widgetCleanupById.get(widgetId)?.();
    } catch {
      // Best-effort cleanup for opt-in renderer code.
    }
    this.widgetCleanupById.delete(widgetId);
    this.widgetRenderSignatureById.delete(widgetId);
  }

  private destroyAllWidgetRuntimes(): void {
    for (const widgetId of this.widgetCleanupById.keys()) {
      this.destroyWidgetRuntime(widgetId);
    }
    this.widgetObjectElementById.clear();
    this.widgetBodyElementById.clear();
  }

  private readonly handleVisualObjectLayerMouseDown = (event: MouseEvent): void => {
    if (event.button !== 0) {
      return;
    }
    const target = event.target as HTMLElement | null;
    const objectElement = target?.closest<HTMLElement>("[data-image-id], [data-widget-id]");
    if (!objectElement) {
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const kind: VisualObjectInteractionState["kind"] = objectElement.dataset.imageId ? "image" : "widget";
    const objectId = objectElement.dataset.imageId ?? objectElement.dataset.widgetId;
    if (!objectId) {
      return;
    }

    if (target?.closest("[data-image-action], [data-widget-action]")) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }

    const visualObject = this.getVisualObject(kind, sheet.id, objectId);
    if (!visualObject || visualObject.state.locked) {
      return;
    }
    const metrics = this.visualObjectSurfaceMetrics;
    if (!metrics || metrics.sheetId !== sheet.id) {
      return;
    }

    const resizeHandle = target?.closest("[data-image-resize='true'], [data-widget-resize='true']");
    const moveHandle = target?.closest("[data-image-move='true'], [data-widget-move='true']");
    if (!resizeHandle && !moveHandle) {
      return;
    }

    const mode: VisualObjectInteractionState["mode"] = resizeHandle ? "resize" : "move";
    const originRect = this.resolveVisualObjectRect(visualObject.position, metrics.rowOffsets, metrics.colOffsets, metrics.rowCount, metrics.colCount);
    this.visualObjectInteraction = {
      mode,
      kind,
      sheetId: sheet.id,
      objectId,
      pointerStartX: event.clientX,
      pointerStartY: event.clientY,
      originRect,
      liveRect: { ...originRect }
    };
    if (kind === "image") {
      this.engine.selectImage(sheet.id, objectId);
    } else {
      this.engine.selectWidget(sheet.id, objectId);
    }
    event.preventDefault();
    event.stopPropagation();
    this.render();
    this.focus();
  };

  private readonly handleVisualObjectLayerClick = (event: Event): void => {
    const target = event.target as HTMLElement | null;
    const imageElement = target?.closest<HTMLElement>("[data-image-id]");
    if (imageElement) {
      const imageId = imageElement.dataset.imageId;
      const sheetId = this.engine.getActiveSheet().id;
      const action = target?.closest<HTMLElement>("[data-image-action]")?.dataset.imageAction;
      const image = imageId ? this.engine.getImage(sheetId, imageId) : undefined;
      if (imageId && image) {
        if (action === "delete") this.engine.deleteImage(sheetId, imageId);
        else if (action === "lock") this.engine.updateImage({ sheetId, imageId, state: { locked: !image.state.locked } });
        else if (action === "back") this.engine.updateImage({ sheetId, imageId, position: { zIndex: Math.max(0, image.position.zIndex - 1) } });
        else if (action === "front") this.engine.updateImage({ sheetId, imageId, position: { zIndex: image.position.zIndex + 1 } });
        else this.engine.selectImage(sheetId, imageId);
        this.requestRender();
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    }
    const widgetElement = target?.closest<HTMLElement>("[data-widget-id]");
    if (widgetElement) {
      const widgetId = widgetElement.dataset.widgetId;
      const sheetId = this.engine.getActiveSheet().id;
      const action = target?.closest<HTMLElement>("[data-widget-action]")?.dataset.widgetAction;
      const widget = widgetId ? this.engine.getWidget(sheetId, widgetId) : undefined;
      if (widgetId && widget) {
        if (action === "delete") this.engine.deleteWidget(sheetId, widgetId);
        else if (action === "lock") this.engine.updateWidget({ sheetId, widgetId, state: { locked: !widget.state.locked } });
        else if (action === "back") this.engine.updateWidget({ sheetId, widgetId, position: { zIndex: Math.max(0, widget.position.zIndex - 1) } });
        else if (action === "front") this.engine.updateWidget({ sheetId, widgetId, position: { zIndex: widget.position.zIndex + 1 } });
        else this.engine.selectWidget(sheetId, widgetId);
        this.requestRender();
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    }
  };

  private getVisualObject(kind: VisualObjectInteractionState["kind"], sheetId: string, objectId: string): WorksheetImageObject | WorksheetWidgetObject | undefined {
    if (kind === "image") return this.engine.getImage(sheetId, objectId);
    return this.engine.getWidget(sheetId, objectId);
  }

  private commitVisualObjectGeometry(
    kind: VisualObjectInteractionState["kind"],
    sheetId: string,
    objectId: string,
    mode: VisualObjectInteractionState["mode"],
    rect: VisualObjectRect
  ): void {
    const metrics = this.visualObjectSurfaceMetrics;
    const visualObject = this.getVisualObject(kind, sheetId, objectId);
    if (!metrics || metrics.sheetId !== sheetId || !visualObject) return;
    const position = this.resolveWorksheetObjectPositionFromRect(rect, metrics, visualObject.position);
    if (mode === "move") {
      const next = {
        fromCell: position.fromCell,
        toCell: position.toCell,
        offsetX: position.offsetX,
        offsetY: position.offsetY,
        zIndex: position.zIndex
      };
      if (kind === "image") this.engine.moveImage({ sheetId, imageId: objectId, position: next });
      else this.engine.moveWidget({ sheetId, widgetId: objectId, position: next });
      return;
    }
    const next = { width: position.width, height: position.height, toCell: position.toCell };
    if (kind === "image") this.engine.resizeImage({ sheetId, imageId: objectId, position: next });
    else this.engine.resizeWidget({ sheetId, widgetId: objectId, position: next });
  }

  private readonly handleVisualObjectKeyDown = (event: KeyboardEvent): void => {
    const target = event.target as HTMLElement | null;
    if (target?.closest("[data-image-action], [data-widget-action]")) return;
    const element = target?.closest<HTMLElement>("[data-image-id], [data-widget-id]");
    if (!element) return;
    const kind: VisualObjectInteractionState["kind"] = element.dataset.imageId ? "image" : "widget";
    const objectId = element.dataset.imageId ?? element.dataset.widgetId;
    const sheetId = this.engine.getActiveSheet().id;
    const visualObject = objectId ? this.getVisualObject(kind, sheetId, objectId) : undefined;
    if (!objectId || !visualObject) return;

    if (event.key.toLowerCase() === "l") {
      if (kind === "image") this.engine.updateImage({ sheetId, imageId: objectId, state: { locked: !visualObject.state.locked } });
      else this.engine.updateWidget({ sheetId, widgetId: objectId, state: { locked: !visualObject.state.locked } });
    } else if (event.key === "Delete" || event.key === "Backspace") {
      if (kind === "image") this.engine.deleteImage(sheetId, objectId);
      else this.engine.deleteWidget(sheetId, objectId);
    } else if (event.key === "PageUp" || event.key === "PageDown") {
      const zIndex = Math.max(0, visualObject.position.zIndex + (event.key === "PageUp" ? 1 : -1));
      if (kind === "image") this.engine.updateImage({ sheetId, imageId: objectId, position: { zIndex } });
      else this.engine.updateWidget({ sheetId, widgetId: objectId, position: { zIndex } });
    } else if (["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(event.key)) {
      if (visualObject.state.locked || !this.visualObjectSurfaceMetrics) return;
      const delta = event.ctrlKey ? 1 : 10;
      const rect = this.resolveVisualObjectRect(
        visualObject.position,
        this.visualObjectSurfaceMetrics.rowOffsets,
        this.visualObjectSurfaceMetrics.colOffsets,
        this.visualObjectSurfaceMetrics.rowCount,
        this.visualObjectSurfaceMetrics.colCount
      );
      const horizontal = event.key === "ArrowLeft" ? -delta : event.key === "ArrowRight" ? delta : 0;
      const vertical = event.key === "ArrowUp" ? -delta : event.key === "ArrowDown" ? delta : 0;
      if (event.shiftKey) {
        rect.width = Math.max(VISUAL_OBJECT_MIN_WIDTH, rect.width + horizontal);
        rect.height = Math.max(VISUAL_OBJECT_MIN_HEIGHT, rect.height + vertical);
      } else {
        rect.left += horizontal;
        rect.top += vertical;
      }
      this.commitVisualObjectGeometry(kind, sheetId, objectId, event.shiftKey ? "resize" : "move", rect);
    } else {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    this.requestRender();
  };

  private readonly handleVisualObjectInteractionMouseMove = (event: MouseEvent): void => {
    const interaction = this.visualObjectInteraction;
    if (!interaction) {
      return;
    }
    const throttleMs = 16;
    const now = Date.now();
    if (throttleMs > 0 && now - this.visualObjectLastInteractionMoveTs < throttleMs) {
      event.preventDefault();
      return;
    }
    this.visualObjectLastInteractionMoveTs = now;
    const metrics = this.visualObjectSurfaceMetrics;
    if (!metrics || metrics.sheetId !== interaction.sheetId) {
      return;
    }

    const deltaX = event.clientX - interaction.pointerStartX;
    const deltaY = event.clientY - interaction.pointerStartY;
    const maxSurfaceWidth = metrics.colOffsets[metrics.colOffsets.length - 1] ?? 0;
    const maxSurfaceHeight = metrics.rowOffsets[metrics.rowOffsets.length - 1] ?? 0;
    const nextRect: VisualObjectRect = {
      ...interaction.originRect
    };

    if (interaction.mode === "move") {
      nextRect.left = clampNumeric(
        interaction.originRect.left + deltaX,
        ROW_HEADER_WIDTH,
        Math.max(ROW_HEADER_WIDTH, ROW_HEADER_WIDTH + maxSurfaceWidth - interaction.originRect.width)
      );
      nextRect.top = clampNumeric(
        interaction.originRect.top + deltaY,
        0,
        Math.max(0, maxSurfaceHeight - interaction.originRect.height)
      );
    } else {
      nextRect.width = clampNumeric(
        interaction.originRect.width + deltaX,
        VISUAL_OBJECT_MIN_WIDTH,
        Math.max(VISUAL_OBJECT_MIN_WIDTH, ROW_HEADER_WIDTH + maxSurfaceWidth - interaction.originRect.left)
      );
      nextRect.height = clampNumeric(
        interaction.originRect.height + deltaY,
        VISUAL_OBJECT_MIN_HEIGHT,
        Math.max(VISUAL_OBJECT_MIN_HEIGHT, maxSurfaceHeight - interaction.originRect.top)
      );
    }

    interaction.liveRect = nextRect;
    this.visualObjectInteraction = interaction;
    const selector = interaction.kind === "image"
      ? `[data-image-id='${interaction.objectId}']`
      : `[data-widget-id='${interaction.objectId}']`;
    const objectElement = this.visualObjectsLayer.querySelector<HTMLElement>(selector);
    if (objectElement) {
      objectElement.style.left = `${nextRect.left}px`;
      objectElement.style.top = `${nextRect.top}px`;
      objectElement.style.width = `${nextRect.width}px`;
      objectElement.style.height = `${nextRect.height}px`;
    }
    event.preventDefault();
  };

  private readonly handleVisualObjectInteractionMouseUp = (): void => {
    const interaction = this.visualObjectInteraction;
    if (!interaction) {
      return;
    }
    this.visualObjectLastInteractionMoveTs = 0;
    this.visualObjectInteraction = undefined;
    const metrics = this.visualObjectSurfaceMetrics;
    if (!this.getVisualObject(interaction.kind, interaction.sheetId, interaction.objectId) || !metrics || metrics.sheetId !== interaction.sheetId) {
      this.requestRender();
      return;
    }

    this.commitVisualObjectGeometry(interaction.kind, interaction.sheetId, interaction.objectId, interaction.mode, interaction.liveRect);
    this.requestRender();
  };

  private getVisibleRemoteGroup(rowModelRow: RowModelRow | undefined, col: number): RowModelRow["group"] | undefined {
    if (!rowModelRow?.group) {
      return undefined;
    }

    const targetCol = Math.max(0, rowModelRow.group.level ?? 0);
    return col === targetCol ? rowModelRow.group : undefined;
  }
}