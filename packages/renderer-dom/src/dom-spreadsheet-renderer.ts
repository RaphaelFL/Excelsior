import {
  CellValidationError,
  SpreadsheetOperationError,
  cellAddressToLabel,
  cellLabelToAddress,
  columnIndexToLabel,
  type CellAddress,
  type ClientSideFilterDescriptor,
  type CellModel,
  type CollaborationPresence,
  type CommentAuthor,
  type PivotAggregateFunction,
  type PivotBuildProgress,
  type CellPrimitive,
  type CellRange,
  type CellValidationConfig,
  type CellValidationRule,
  type CellStyle,
  type ChartPosition,
  type PivotExecutionMode,
  type RowModelRow,
  type SheetMerge,
  type RowResult,
  type SpreadsheetOperation,
  type WorksheetChartObject,
  type WorksheetImageObject,
  type WorksheetWidgetObject,
  type WorksheetChartType,
  type WorkbookEngine
} from "@excelsior/core";
import {
  createFigure,
  createFigureFromSpreadsheetRange,
  type ChartFigureInput,
  type ChartHandle,
  type SpreadsheetRangeInput
} from "@excelsior/charts";
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
  toolbarChartsGroup: string;
  toolbarChartsCommonGroup: string;
  toolbarChartsStatisticalGroup: string;
  toolbarChartsFinancialGroup: string;
  toolbarChartsAdvancedGroup: string;
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
  chartColumn: string;
  chartBar: string;
  chartLine: string;
  chartArea: string;
  chartPie: string;
  chartDonut: string;
  chartScatter: string;
  chartHistogram: string;
  chartBox: string;
  chartViolin: string;
  chartHeatmap: string;
  chartContour: string;
  chartCandlestick: string;
  chartWaterfall: string;
  chartFunnel: string;
  chartPolar: string;
  chartTernary: string;
  chartGeo: string;
  chartTreemap: string;
  chartSunburst: string;
  chartSankey: string;
  chartSurface: string;
  chartSurface3d: string;
  chartScatter3d: string;
  chartUnsupportedType: string;
  chartTooManyObjects: string;
  chartRangeTooLarge: string;
  chartTooManySeries: string;
  chartTooManyPoints: string;
  chartInvalidRange: string;
  chartInserted: string;
  chartAutoUpdated: string;
  chartDelete: string;
  chartDeleted: string;
  chartMoveHandle: string;
  chartResizeHandle: string;
  chartInsertError: string;
  chartEditPanelTitle: string;
  chartEditTypeLabel: string;
  chartEditRangeLabel: string;
  chartEditTitleLabel: string;
  chartEditLegendLabel: string;
  chartEditXAxisTitleLabel: string;
  chartEditYAxisTitleLabel: string;
  chartEditXAxisTypeLabel: string;
  chartEditYAxisTypeLabel: string;
  chartEditXAxisVisibleLabel: string;
  chartEditYAxisVisibleLabel: string;
  chartEditOrientationLabel: string;
  chartEditFirstRowHeaderLabel: string;
  chartEditFirstColumnLabelLabel: string;
  chartEditAutoRefreshLabel: string;
  chartEditCategoryColumnLabel: string;
  chartEditSeriesColumnsLabel: string;
  chartEditValueColumnLabel: string;
  chartAxisTypeLinear: string;
  chartAxisTypeCategory: string;
  chartAxisTypeDate: string;
  chartAxisTypeLog: string;
  chartEditApply: string;
  chartEditClose: string;
  chartEditSaved: string;
  chartEditInvalidRange: string;
  chartEditInvalidBinding: string;
  chartPreviewTitle: string;
  chartPreviewInsert: string;
  chartPreviewCancel: string;
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
  chartLimits?: {
    maxChartsPerSheet?: number;
    maxRangeCells?: number;
    maxSeriesPerChart?: number;
    maxPointsPerChart?: number;
  };
  chartPerformance?: {
    interactionThrottleMs?: number;
    offscreenMarginPx?: number;
    skipOffscreenPreview?: boolean;
  };
  chartInsertPreview?: boolean;
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
  toolbarChartsGroup: "Gráficos",
  toolbarChartsCommonGroup: "Comuns",
  toolbarChartsStatisticalGroup: "Estatísticos",
  toolbarChartsFinancialGroup: "Financeiros",
  toolbarChartsAdvancedGroup: "Avançados",
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
  chartColumn: "Coluna",
  chartBar: "Barra",
  chartLine: "Linha",
  chartArea: "Área",
  chartPie: "Pizza",
  chartDonut: "Rosca",
  chartScatter: "Dispersão",
  chartHistogram: "Histograma",
  chartBox: "Box plot",
  chartViolin: "Violin",
  chartHeatmap: "Heatmap",
  chartContour: "Contorno",
  chartCandlestick: "Candlestick",
  chartWaterfall: "Waterfall",
  chartFunnel: "Funil",
  chartPolar: "Polar",
  chartTernary: "Ternário",
  chartGeo: "Mapa GeoJSON",
  chartTreemap: "Treemap",
  chartSunburst: "Sunburst",
  chartSankey: "Sankey",
  chartSurface: "Superfície",
  chartSurface3d: "Superfície 3D",
  chartScatter3d: "Dispersão 3D",
  chartUnsupportedType: "Em desenvolvimento",
  chartTooManyObjects: "Limite de gráficos por planilha atingido.",
  chartRangeTooLarge: "Intervalo muito grande para gráfico. Reduza a seleção.",
  chartTooManySeries: "Séries demais para renderizar com segurança.",
  chartTooManyPoints: "Pontos demais para renderizar com segurança.",
  chartInvalidRange: "Selecione um intervalo tabular com pelo menos 2 linhas e 2 colunas para criar o gráfico.",
  chartInserted: "Gráfico inserido na planilha.",
  chartAutoUpdated: "Gráfico atualizado com base nos dados da seleção.",
  chartDelete: "Excluir gráfico",
  chartDeleted: "Gráfico removido da planilha.",
  chartMoveHandle: "Mover gráfico",
  chartResizeHandle: "Redimensionar gráfico",
  chartInsertError: "Não foi possível criar o gráfico para o intervalo selecionado.",
  chartEditPanelTitle: "Editar gráfico selecionado",
  chartEditTypeLabel: "Tipo",
  chartEditRangeLabel: "Intervalo",
  chartEditTitleLabel: "Título",
  chartEditLegendLabel: "Exibir legenda",
  chartEditXAxisTitleLabel: "Título eixo X",
  chartEditYAxisTitleLabel: "Título eixo Y",
  chartEditXAxisTypeLabel: "Tipo eixo X",
  chartEditYAxisTypeLabel: "Tipo eixo Y",
  chartEditXAxisVisibleLabel: "Exibir eixo X",
  chartEditYAxisVisibleLabel: "Exibir eixo Y",
  chartEditOrientationLabel: "Orientação",
  chartEditFirstRowHeaderLabel: "1ª linha como cabeçalho",
  chartEditFirstColumnLabelLabel: "1ª coluna como categoria",
  chartEditAutoRefreshLabel: "Atualização automática",
  chartEditCategoryColumnLabel: "Coluna de categorias (X)",
  chartEditSeriesColumnsLabel: "Colunas de séries (Y)",
  chartEditValueColumnLabel: "Coluna de valores (pie/donut)",
  chartAxisTypeLinear: "Linear",
  chartAxisTypeCategory: "Categoria",
  chartAxisTypeDate: "Data",
  chartAxisTypeLog: "Log",
  chartEditApply: "Aplicar",
  chartEditClose: "Fechar",
  chartEditSaved: "Configurações do gráfico atualizadas.",
  chartEditInvalidRange: "Intervalo inválido. Use o formato A1:C10.",
  chartEditInvalidBinding: "Configuração de colunas inválida para o intervalo informado.",
  chartPreviewTitle: "Pré-visualização do gráfico",
  chartPreviewInsert: "Inserir gráfico",
  chartPreviewCancel: "Cancelar"
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

type ChartToolbarAction =
  | "chart-column"
  | "chart-bar"
  | "chart-line"
  | "chart-area"
  | "chart-pie"
  | "chart-donut"
  | "chart-scatter"
  | "chart-histogram"
  | "chart-box"
  | "chart-violin"
  | "chart-heatmap"
  | "chart-contour"
  | "chart-candlestick"
  | "chart-waterfall"
  | "chart-funnel"
  | "chart-polar"
  | "chart-ternary"
  | "chart-geo"
  | "chart-treemap"
  | "chart-sunburst"
  | "chart-sankey"
  | "chart-surface"
  | "chart-surface3d"
  | "chart-scatter3d";

type ChartToolbarCategory = "common" | "statistical" | "financial" | "advanced";

interface ChartToolbarDefinition {
  action: ChartToolbarAction;
  type: WorksheetChartType;
  category: ChartToolbarCategory;
  messageKey: keyof RendererMessages;
  enabled: boolean;
}

interface ChartSurfaceMetrics {
  sheetId: string;
  rowOffsets: number[];
  colOffsets: number[];
  rowCount: number;
  colCount: number;
}

interface ChartRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

interface ChartInteractionState {
  mode: "move" | "resize";
  kind: "chart" | "image" | "widget";
  sheetId: string;
  chartId: string;
  pointerStartX: number;
  pointerStartY: number;
  originRect: ChartRect;
  liveRect: ChartRect;
}

type ChartBindingOptions = Pick<
  NonNullable<WorksheetChartObject["sourceRange"]>,
  | "rangeAddress"
  | "orientation"
  | "firstRowAsHeader"
  | "firstColumnAsLabel"
  | "autoRefresh"
  | "categoryColumnIndex"
  | "seriesColumnIndexes"
  | "valueColumnIndex"
>;

const CHART_ACTION_TO_TYPE: Record<ChartToolbarAction, WorksheetChartType> = {
  "chart-column": "column",
  "chart-bar": "bar",
  "chart-line": "line",
  "chart-area": "area",
  "chart-pie": "pie",
  "chart-donut": "donut",
  "chart-scatter": "scatter",
  "chart-histogram": "histogram",
  "chart-box": "box",
  "chart-violin": "violin",
  "chart-heatmap": "heatmap",
  "chart-contour": "contour",
  "chart-candlestick": "candlestick",
  "chart-waterfall": "waterfall",
  "chart-funnel": "funnel",
  "chart-polar": "polar",
  "chart-ternary": "ternary",
  "chart-geo": "geo",
  "chart-treemap": "treemap",
  "chart-sunburst": "sunburst",
  "chart-sankey": "sankey",
  "chart-surface": "surface",
  "chart-surface3d": "surface3d",
  "chart-scatter3d": "scatter3d"
};

const CHART_TOOLBAR_DEFINITIONS: ChartToolbarDefinition[] = [
  { action: "chart-column", type: "column", category: "common", messageKey: "chartColumn", enabled: true },
  { action: "chart-bar", type: "bar", category: "common", messageKey: "chartBar", enabled: true },
  { action: "chart-line", type: "line", category: "common", messageKey: "chartLine", enabled: true },
  { action: "chart-area", type: "area", category: "common", messageKey: "chartArea", enabled: true },
  { action: "chart-pie", type: "pie", category: "common", messageKey: "chartPie", enabled: true },
  { action: "chart-donut", type: "donut", category: "common", messageKey: "chartDonut", enabled: true },
  { action: "chart-scatter", type: "scatter", category: "common", messageKey: "chartScatter", enabled: true },
  { action: "chart-histogram", type: "histogram", category: "statistical", messageKey: "chartHistogram", enabled: true },
  { action: "chart-box", type: "box", category: "statistical", messageKey: "chartBox", enabled: true },
  { action: "chart-violin", type: "violin", category: "statistical", messageKey: "chartViolin", enabled: true },
  { action: "chart-heatmap", type: "heatmap", category: "statistical", messageKey: "chartHeatmap", enabled: true },
  { action: "chart-contour", type: "contour", category: "statistical", messageKey: "chartContour", enabled: true },
  { action: "chart-candlestick", type: "candlestick", category: "financial", messageKey: "chartCandlestick", enabled: true },
  { action: "chart-waterfall", type: "waterfall", category: "financial", messageKey: "chartWaterfall", enabled: true },
  { action: "chart-funnel", type: "funnel", category: "financial", messageKey: "chartFunnel", enabled: true },
  { action: "chart-polar", type: "polar", category: "advanced", messageKey: "chartPolar", enabled: true },
  { action: "chart-ternary", type: "ternary", category: "advanced", messageKey: "chartTernary", enabled: true },
  { action: "chart-geo", type: "geo", category: "advanced", messageKey: "chartGeo", enabled: true },
  { action: "chart-treemap", type: "treemap", category: "advanced", messageKey: "chartTreemap", enabled: true },
  { action: "chart-sunburst", type: "sunburst", category: "advanced", messageKey: "chartSunburst", enabled: true },
  { action: "chart-sankey", type: "sankey", category: "advanced", messageKey: "chartSankey", enabled: true },
  { action: "chart-surface", type: "surface", category: "advanced", messageKey: "chartSurface", enabled: true },
  { action: "chart-surface3d", type: "surface3d", category: "advanced", messageKey: "chartSurface3d", enabled: true },
  { action: "chart-scatter3d", type: "scatter3d", category: "advanced", messageKey: "chartScatter3d", enabled: true }
];

const CHART_PLACEHOLDER_ACTIONS = new Set<ChartToolbarAction>(
  CHART_TOOLBAR_DEFINITIONS.filter((item) => !item.enabled).map((item) => item.action)
);
const CHART_DISABLED_ACTIONS = new Set<ChartToolbarAction>(
  CHART_TOOLBAR_DEFINITIONS.filter((item) => !item.enabled).map((item) => item.action)
);
const CHART_ACTIONS = new Set<ChartToolbarAction>(Object.keys(CHART_ACTION_TO_TYPE) as ChartToolbarAction[]);
const CHART_EDIT_TYPE_OPTIONS: Array<{ type: WorksheetChartType; messageKey: keyof RendererMessages }> = [
  { type: "column", messageKey: "chartColumn" },
  { type: "bar", messageKey: "chartBar" },
  { type: "line", messageKey: "chartLine" },
  { type: "area", messageKey: "chartArea" },
  { type: "pie", messageKey: "chartPie" },
  { type: "donut", messageKey: "chartDonut" },
  { type: "scatter", messageKey: "chartScatter" },
  { type: "histogram", messageKey: "chartHistogram" },
  { type: "box", messageKey: "chartBox" },
  { type: "violin", messageKey: "chartViolin" },
  { type: "heatmap", messageKey: "chartHeatmap" },
  { type: "contour", messageKey: "chartContour" },
  { type: "candlestick", messageKey: "chartCandlestick" },
  { type: "waterfall", messageKey: "chartWaterfall" },
  { type: "funnel", messageKey: "chartFunnel" },
  { type: "polar", messageKey: "chartPolar" },
  { type: "ternary", messageKey: "chartTernary" },
  { type: "geo", messageKey: "chartGeo" },
  { type: "treemap", messageKey: "chartTreemap" },
  { type: "sunburst", messageKey: "chartSunburst" },
  { type: "sankey", messageKey: "chartSankey" },
  { type: "surface", messageKey: "chartSurface" },
  { type: "surface3d", messageKey: "chartSurface3d" },
  { type: "scatter3d", messageKey: "chartScatter3d" }
];
const CHART_AXIS_TYPE_OPTIONS = ["linear", "category", "date", "log"] as const;
type ChartAxisTypeOption = (typeof CHART_AXIS_TYPE_OPTIONS)[number];
type ChartAxisLayoutType = ChartAxisTypeOption | "multicategory";

const CHART_MIN_WIDTH = 220;
const CHART_MIN_HEIGHT = 150;
const CHART_SVG_NS = "http://www.w3.org/2000/svg";
const MAX_CHARTS_PER_SHEET_DEFAULT = 50;
const MAX_CHART_RANGE_CELLS_DEFAULT = 100000;
const MAX_CHART_SERIES_DEFAULT = 64;
const MAX_CHART_POINTS_DEFAULT = 100000;
const CHART_INTERACTION_THROTTLE_MS_DEFAULT = 24;
const CHART_OFFSCREEN_MARGIN_PX_DEFAULT = 180;

const clampNumeric = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));
const toFiniteNumber = (value: unknown, fallback = 0): number =>
  typeof value === "number" && Number.isFinite(value) ? value : fallback;
const cloneSerializable = <T>(value: T): T => {
  if (typeof structuredClone === "function") {
    try {
      return structuredClone(value);
    } catch {
      // Fallback to JSON clone for plain chart payloads.
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

  private readonly geoJsonFileInput = document.createElement("input");

  private readonly chartLayoutImageFileInput = document.createElement("input");

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

  private readonly chartEditPanel = document.createElement("div");

  private readonly chartEditPanelHeader = document.createElement("div");

  private readonly pivotApplyButton = createFindReplaceActionButton("apply", "");

  private readonly pivotCloseButton = createFindReplaceActionButton("close", "");

  private readonly chartEditTypeSelect = document.createElement("select");

  private readonly chartEditRangeInput = document.createElement("input");

  private readonly chartEditTitleInput = document.createElement("input");

  private readonly chartEditLegendToggle = document.createElement("input");

  private readonly chartEditXAxisTitleInput = document.createElement("input");

  private readonly chartEditYAxisTitleInput = document.createElement("input");

  private readonly chartEditAnnotationTextInput = document.createElement("input");

  private readonly chartEditAnnotationXInput = document.createElement("input");

  private readonly chartEditAnnotationYInput = document.createElement("input");

  private readonly chartEditAnnotationArrowToggle = document.createElement("input");

  private readonly chartEditShapeTypeSelect = document.createElement("select");

  private readonly chartEditShapeX0Input = document.createElement("input");

  private readonly chartEditShapeY0Input = document.createElement("input");

  private readonly chartEditShapeX1Input = document.createElement("input");

  private readonly chartEditShapeY1Input = document.createElement("input");

  private readonly chartEditXAxisTypeSelect = document.createElement("select");

  private readonly chartEditYAxisTypeSelect = document.createElement("select");

  private readonly chartEditXAxisVisibleToggle = document.createElement("input");

  private readonly chartEditYAxisVisibleToggle = document.createElement("input");

  private readonly chartEditRangeSelectorToggle = document.createElement("input");

  private readonly chartEditRangeSliderToggle = document.createElement("input");

  private readonly chartEditOrientationSelect = document.createElement("select");

  private readonly chartEditFirstRowHeaderToggle = document.createElement("input");

  private readonly chartEditFirstColumnLabelToggle = document.createElement("input");

  private readonly chartEditAutoRefreshToggle = document.createElement("input");

  private readonly chartEditCategoryColumnInput = document.createElement("input");

  private readonly chartEditSeriesColumnsInput = document.createElement("input");

  private readonly chartEditValueColumnInput = document.createElement("input");

  private readonly chartEditApplyButton = createFindReplaceActionButton("apply", "");

  private readonly chartEditCloseButton = createFindReplaceActionButton("close", "");

  private readonly chartInsertPreviewPanel = document.createElement("div");

  private readonly chartInsertPreviewHeader = document.createElement("div");

  private readonly chartInsertPreviewHost = document.createElement("div");

  private readonly chartInsertPreviewInsertButton = createFindReplaceActionButton("insert", "");

  private readonly chartInsertPreviewCancelButton = createFindReplaceActionButton("cancel", "");

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

  private readonly chartsLayer = document.createElement("div");

  private sheetZoom = 1;

  private headerResize?: {
    axis: "column" | "row";
    index: number;
    sheetId: string;
    startPosition: number;
    startSize: number;
    currentSize: number;
  };

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

  private chartPanelDragging?: { panel: HTMLElement; offsetX: number; offsetY: number };

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

  private chartFeedback?: { sheetId: string; message: string; isError: boolean };

  private selectedChartId?: string;

  private chartInteraction?: ChartInteractionState;

  private chartSurfaceMetrics?: ChartSurfaceMetrics;

  private chartLastInteractionMoveTs = 0;

  private lastViewportScrollRenderTs = 0;

  private splitPaneDrag?: { axis: "horizontal" | "vertical"; index: number };

  private readonly chartRenderStatusById = new Map<string, string>();

  private readonly chartObjectElementById = new Map<string, HTMLElement>();

  private readonly chartBodyElementById = new Map<string, HTMLElement>();

  private readonly chartRuntimeById = new Map<string, ChartHandle>();

  private readonly chartRuntimeFigureById = new Map<string, string>();

  private readonly imageObjectElementById = new Map<string, HTMLElement>();

  private readonly widgetObjectElementById = new Map<string, HTMLElement>();

  private readonly widgetBodyElementById = new Map<string, HTMLElement>();

  private readonly widgetRenderSignatureById = new Map<string, string>();

  private readonly widgetCleanupById = new Map<string, () => void>();

  private pendingChartInsertion?: {
    sheetId: string;
    chartType: WorksheetChartType;
    sourceRange: CellRange;
    binding: ChartBindingOptions;
    figure: WorksheetChartObject["figure"];
    title: string;
    placeholderMode: boolean;
  };

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

  private getChartLimits(): {
    maxChartsPerSheet: number;
    maxRangeCells: number;
    maxSeriesPerChart: number;
    maxPointsPerChart: number;
  } {
    return {
      maxChartsPerSheet: Math.max(1, this.options.chartLimits?.maxChartsPerSheet ?? MAX_CHARTS_PER_SHEET_DEFAULT),
      maxRangeCells: Math.max(4, this.options.chartLimits?.maxRangeCells ?? MAX_CHART_RANGE_CELLS_DEFAULT),
      maxSeriesPerChart: Math.max(1, this.options.chartLimits?.maxSeriesPerChart ?? MAX_CHART_SERIES_DEFAULT),
      maxPointsPerChart: Math.max(10, this.options.chartLimits?.maxPointsPerChart ?? MAX_CHART_POINTS_DEFAULT)
    };
  }

  private getChartPerformanceOptions(): {
    interactionThrottleMs: number;
    offscreenMarginPx: number;
    skipOffscreenPreview: boolean;
  } {
    return {
      interactionThrottleMs: Math.max(0, this.options.chartPerformance?.interactionThrottleMs ?? CHART_INTERACTION_THROTTLE_MS_DEFAULT),
      offscreenMarginPx: Math.max(0, this.options.chartPerformance?.offscreenMarginPx ?? CHART_OFFSCREEN_MARGIN_PX_DEFAULT),
      skipOffscreenPreview: this.options.chartPerformance?.skipOffscreenPreview !== false
    };
  }

  private sanitizeChartText(rawValue: string, maxLength = 240): string {
    const normalized = rawValue
      .replace(/[\u0000-\u001f\u007f]/g, " ")
      .replace(/</g, "‹")
      .replace(/>/g, "›")
      .replace(/\s+/g, " ")
      .trim();
    if (!normalized) {
      return "";
    }
    const bounded = normalized.slice(0, maxLength);
    return /^[=+\-@]/.test(bounded) ? `'${bounded}` : bounded;
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

  private autoFitColumn(col: number): void {
    const sheet = this.engine.getActiveSheet();
    let width = 40;
    for (let row = 0; row < sheet.rowCount; row += 1) {
      const value = this.getRenderedCellDisplayValue(sheet.id, row, col);
      const fontSize = this.getCellStyle(sheet, row, col)?.fontSize ?? 12;
      width = Math.max(width, value.length * fontSize * 0.62 + 20);
    }
    this.engine.resizeColumn(sheet.id, col, Math.min(600, Math.ceil(width)));
  }

  private autoFitRow(row: number): void {
    const sheet = this.engine.getActiveSheet();
    let height = 20;
    for (let col = 0; col < sheet.columnCount; col += 1) {
      const value = this.getRenderedCellDisplayValue(sheet.id, row, col);
      const style = this.getCellStyle(sheet, row, col);
      const fontSize = style?.fontSize ?? 12;
      const lineHeight = fontSize * 1.35;
      const availableWidth = Math.max(1, this.getColumnWidth(sheet, col) / this.sheetZoom - 16);
      const textWidth = value.length * fontSize * 0.62;
      const lineCount = style?.wrap ? Math.max(1, Math.ceil(textWidth / availableWidth)) : 1;
      const rotation = Math.abs(style?.rotation ?? 0) * Math.PI / 180;
      const rotatedHeight = Math.abs(Math.sin(rotation)) * textWidth + Math.abs(Math.cos(rotation)) * lineHeight;
      height = Math.max(height, style?.rotation ? rotatedHeight + 12 : lineCount * lineHeight + 12);
    }
    this.engine.resizeRow(sheet.id, row, Math.min(400, Math.ceil(height)));
  }

  private readonly handleHeaderResizeMouseDown = (event: MouseEvent): void => {
    const element = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-column-resize], [data-row-resize]");
    if (!element) return;
    event.preventDefault();
    event.stopPropagation();
    const column = element.dataset.columnResize;
    const row = element.dataset.rowResize;
    const sheet = this.engine.getActiveSheet();
    const axis = column !== undefined ? "column" : "row";
    const index = Number(column ?? row);
    this.headerResize = {
      axis,
      index,
      sheetId: sheet.id,
      startPosition: axis === "column" ? event.clientX : event.clientY,
      startSize: axis === "column"
        ? this.getColumnWidth(sheet, index) / this.sheetZoom
        : this.getRowHeight(sheet, index) / this.sheetZoom,
      currentSize: axis === "column"
        ? this.getColumnWidth(sheet, index) / this.sheetZoom
        : this.getRowHeight(sheet, index) / this.sheetZoom
    };
    this.root.classList.add(axis === "column" ? "is-resizing-column" : "is-resizing-row");
    window.addEventListener("mousemove", this.handleHeaderResizeMouseMove);
    window.addEventListener("mouseup", this.handleHeaderResizeMouseUp, { once: true });
  };

  private previewHeaderResize(state: NonNullable<DomSpreadsheetRenderer["headerResize"]>, size: number): void {
    const delta = (size - state.currentSize) * this.sheetZoom;
    if (!delta) return;
    state.currentSize = size;

    if (state.axis === "column") {
      const header = this.chrome.querySelector<HTMLElement>(`[data-column-header-col='${state.index}']`);
      if (header) header.style.width = `${Math.max(0, Number.parseFloat(header.style.width) + delta)}px`;
      for (const cell of this.cellsLayer.querySelectorAll<HTMLElement>("[data-row][data-col]")) {
        const start = Number(cell.dataset.col);
        const span = Number(cell.getAttribute("aria-colspan") ?? 1);
        if (start <= state.index && state.index < start + span) {
          cell.style.width = `${Math.max(0, Number.parseFloat(cell.style.width) + delta)}px`;
        } else if (start > state.index) {
          cell.style.left = `${Number.parseFloat(cell.style.left) + delta}px`;
        }
      }
      this.surface.style.width = `${Math.max(0, Number.parseFloat(this.surface.style.width) + delta)}px`;
      return;
    }

    for (const header of this.rowHeaders.querySelectorAll<HTMLElement>("[data-row-header-row]")) {
      const row = Number(header.dataset.rowHeaderRow);
      if (row === state.index) header.style.height = `${Math.max(0, Number.parseFloat(header.style.height) + delta)}px`;
      else if (row > state.index) header.style.top = `${Number.parseFloat(header.style.top) + delta}px`;
    }
    for (const cell of this.cellsLayer.querySelectorAll<HTMLElement>("[data-row][data-col]")) {
      const start = Number(cell.dataset.row);
      const span = Number(cell.getAttribute("aria-rowspan") ?? 1);
      if (start <= state.index && state.index < start + span) {
        cell.style.height = `${Math.max(0, Number.parseFloat(cell.style.height) + delta)}px`;
      } else if (start > state.index) {
        cell.style.top = `${Number.parseFloat(cell.style.top) + delta}px`;
      }
    }
    this.surface.style.height = `${Math.max(0, Number.parseFloat(this.surface.style.height) + delta)}px`;
  }

  private readonly handleHeaderResizeMouseMove = (event: MouseEvent): void => {
    const state = this.headerResize;
    if (!state) return;
    const position = state.axis === "column" ? event.clientX : event.clientY;
    const minimum = state.axis === "column" ? 40 : 20;
    const maximum = state.axis === "column" ? 600 : 400;
    const size = Math.max(minimum, Math.min(maximum, state.startSize + (position - state.startPosition) / this.sheetZoom));
    this.previewHeaderResize(state, Math.round(size));
  };

  private readonly handleHeaderResizeMouseUp = (): void => {
    const state = this.headerResize;
    this.headerResize = undefined;
    window.removeEventListener("mousemove", this.handleHeaderResizeMouseMove);
    this.root.classList.remove("is-resizing-column", "is-resizing-row");
    if (!state || state.currentSize === state.startSize) return;
    if (state.axis === "column") this.engine.resizeColumn(state.sheetId, state.index, state.currentSize);
    else this.engine.resizeRow(state.sheetId, state.index, state.currentSize);
  };

  private readonly handleHeaderResizeDoubleClick = (event: MouseEvent): void => {
    const element = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-column-resize], [data-row-resize]");
    if (!element) return;
    event.preventDefault();
    event.stopPropagation();
    if (element.dataset.columnResize !== undefined) this.autoFitColumn(Number(element.dataset.columnResize));
    else this.autoFitRow(Number(element.dataset.rowResize));
  };

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
    this.findReplaceState.error = undefined;
    this.findReplaceState.matches = [];
    this.findReplaceState.activeIndex = -1;
    this.findReplacePanel.hidden = true;
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
    this.geoJsonFileInput.type = "file";
    this.geoJsonFileInput.accept = ".geojson,.json,application/geo+json,application/json";
    this.geoJsonFileInput.hidden = true;
    this.geoJsonFileInput.addEventListener("change", this.handleGeoJsonFileChange);
    this.chartLayoutImageFileInput.type = "file";
    this.chartLayoutImageFileInput.accept = "image/png,image/jpeg,image/gif,image/webp";
    this.chartLayoutImageFileInput.dataset.chartLayoutImage = "file";
    this.chartLayoutImageFileInput.setAttribute("aria-label", "Imagem do gráfico");
    this.chartLayoutImageFileInput.hidden = true;
    this.chartLayoutImageFileInput.addEventListener("change", this.handleChartLayoutImageFileChange);
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
    this.chartEditPanel.className = "excelsior-find-replace excelsior-chart-edit-panel";
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
    this.chartsLayer.className = "excelsior-charts-layer";
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
    this.chartEditPanel.hidden = true;
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
    this.chartEditTypeSelect.className = "excelsior-find-replace-input";
    this.chartEditTypeSelect.dataset.chartRole = "type";
    this.chartEditTypeSelect.replaceChildren(
      ...CHART_EDIT_TYPE_OPTIONS.map((optionDef) => {
        const option = document.createElement("option");
        option.value = optionDef.type;
        option.textContent = this.messages[optionDef.messageKey] ?? optionDef.type;
        option.disabled = !this.isChartTypeEnabled(optionDef.type);
        return option;
      })
    );
    this.chartEditRangeInput.type = "text";
    this.chartEditRangeInput.className = "excelsior-find-replace-input";
    this.chartEditRangeInput.dataset.chartRole = "range";
    this.chartEditTitleInput.type = "text";
    this.chartEditTitleInput.className = "excelsior-find-replace-input";
    this.chartEditTitleInput.dataset.chartRole = "title";
    this.chartEditLegendToggle.type = "checkbox";
    this.chartEditLegendToggle.dataset.chartRole = "legend";
    this.chartEditXAxisTitleInput.type = "text";
    this.chartEditXAxisTitleInput.className = "excelsior-find-replace-input";
    this.chartEditXAxisTitleInput.dataset.chartRole = "x-axis-title";
    this.chartEditYAxisTitleInput.type = "text";
    this.chartEditYAxisTitleInput.className = "excelsior-find-replace-input";
    this.chartEditYAxisTitleInput.dataset.chartRole = "y-axis-title";
    this.chartEditAnnotationTextInput.type = "text";
    this.chartEditAnnotationTextInput.className = "excelsior-find-replace-input";
    this.chartEditAnnotationTextInput.dataset.chartRole = "annotation-text";
    for (const [input, role] of [
      [this.chartEditAnnotationXInput, "annotation-x"],
      [this.chartEditAnnotationYInput, "annotation-y"],
      [this.chartEditShapeX0Input, "shape-x0"],
      [this.chartEditShapeY0Input, "shape-y0"],
      [this.chartEditShapeX1Input, "shape-x1"],
      [this.chartEditShapeY1Input, "shape-y1"]
    ] as const) {
      input.type = "number";
      input.min = "0";
      input.max = "1";
      input.step = "0.05";
      input.className = "excelsior-find-replace-input";
      input.dataset.chartRole = role;
    }
    this.chartEditAnnotationArrowToggle.type = "checkbox";
    this.chartEditAnnotationArrowToggle.dataset.chartRole = "annotation-arrow";
    this.chartEditShapeTypeSelect.className = "excelsior-find-replace-input";
    this.chartEditShapeTypeSelect.dataset.chartRole = "shape-type";
    this.chartEditShapeTypeSelect.replaceChildren(
      ...[["", "Nenhuma"], ["line", "Linha"], ["rect", "Retângulo"], ["circle", "Círculo"]].map(([value, label]) => {
        const option = document.createElement("option");
        option.value = value ?? "";
        option.textContent = label ?? "";
        return option;
      })
    );
    this.chartEditXAxisTypeSelect.className = "excelsior-find-replace-input";
    this.chartEditXAxisTypeSelect.dataset.chartRole = "x-axis-type";
    this.chartEditYAxisTypeSelect.className = "excelsior-find-replace-input";
    this.chartEditYAxisTypeSelect.dataset.chartRole = "y-axis-type";
    const axisTypeOptions = CHART_AXIS_TYPE_OPTIONS.map((axisType) => {
      const option = document.createElement("option");
      option.value = axisType;
      option.textContent =
        axisType === "linear"
          ? this.messages.chartAxisTypeLinear
          : axisType === "category"
            ? this.messages.chartAxisTypeCategory
            : axisType === "date"
              ? this.messages.chartAxisTypeDate
              : this.messages.chartAxisTypeLog;
      return option;
    });
    this.chartEditXAxisTypeSelect.replaceChildren(...axisTypeOptions.map((option) => option.cloneNode(true) as HTMLOptionElement));
    this.chartEditYAxisTypeSelect.replaceChildren(...axisTypeOptions.map((option) => option.cloneNode(true) as HTMLOptionElement));
    this.chartEditXAxisVisibleToggle.type = "checkbox";
    this.chartEditXAxisVisibleToggle.dataset.chartRole = "x-axis-visible";
    this.chartEditYAxisVisibleToggle.type = "checkbox";
    this.chartEditYAxisVisibleToggle.dataset.chartRole = "y-axis-visible";
    this.chartEditRangeSelectorToggle.type = "checkbox";
    this.chartEditRangeSelectorToggle.dataset.chartRole = "range-selector";
    this.chartEditRangeSliderToggle.type = "checkbox";
    this.chartEditRangeSliderToggle.dataset.chartRole = "range-slider";
    this.chartEditOrientationSelect.className = "excelsior-find-replace-input";
    this.chartEditOrientationSelect.dataset.chartRole = "orientation";
    this.chartEditOrientationSelect.replaceChildren(
      ...([
        ["rows", "Linhas"],
        ["columns", "Colunas"]
      ] as const).map(([value, label]) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = label;
        return option;
      })
    );
    this.chartEditFirstRowHeaderToggle.type = "checkbox";
    this.chartEditFirstRowHeaderToggle.dataset.chartRole = "first-row-header";
    this.chartEditFirstColumnLabelToggle.type = "checkbox";
    this.chartEditFirstColumnLabelToggle.dataset.chartRole = "first-column-label";
    this.chartEditAutoRefreshToggle.type = "checkbox";
    this.chartEditAutoRefreshToggle.dataset.chartRole = "auto-refresh";
    this.chartEditCategoryColumnInput.type = "text";
    this.chartEditCategoryColumnInput.className = "excelsior-find-replace-input";
    this.chartEditCategoryColumnInput.dataset.chartRole = "category-column";
    this.chartEditCategoryColumnInput.placeholder = "1";
    this.chartEditSeriesColumnsInput.type = "text";
    this.chartEditSeriesColumnsInput.className = "excelsior-find-replace-input";
    this.chartEditSeriesColumnsInput.dataset.chartRole = "series-columns";
    this.chartEditSeriesColumnsInput.placeholder = "2,3";
    this.chartEditValueColumnInput.type = "text";
    this.chartEditValueColumnInput.className = "excelsior-find-replace-input";
    this.chartEditValueColumnInput.dataset.chartRole = "value-column";
    this.chartEditValueColumnInput.placeholder = "2";

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

    this.chartEditApplyButton.textContent = this.messages.chartEditApply;
    this.chartEditApplyButton.dataset.chartAction = "apply";
    delete this.chartEditApplyButton.dataset.findAction;
    this.chartEditCloseButton.textContent = this.messages.chartEditClose;
    this.chartEditCloseButton.dataset.chartAction = "close";
    delete this.chartEditCloseButton.dataset.findAction;
    const chartTitleField = document.createElement("label");
    chartTitleField.className = "excelsior-find-replace-field";
    const chartTitleLabel = document.createElement("span");
    chartTitleLabel.textContent = this.messages.chartEditTitleLabel;
    chartTitleField.append(chartTitleLabel, this.chartEditTitleInput);
    const chartTypeField = document.createElement("label");
    chartTypeField.className = "excelsior-find-replace-field";
    const chartTypeLabel = document.createElement("span");
    chartTypeLabel.textContent = this.messages.chartEditTypeLabel;
    chartTypeField.append(chartTypeLabel, this.chartEditTypeSelect);
    const chartRangeField = document.createElement("label");
    chartRangeField.className = "excelsior-find-replace-field";
    const chartRangeLabel = document.createElement("span");
    chartRangeLabel.textContent = this.messages.chartEditRangeLabel;
    chartRangeField.append(chartRangeLabel, this.chartEditRangeInput);
    const chartXAxisTitleField = document.createElement("label");
    chartXAxisTitleField.className = "excelsior-find-replace-field";
    const chartXAxisTitleLabel = document.createElement("span");
    chartXAxisTitleLabel.textContent = this.messages.chartEditXAxisTitleLabel;
    chartXAxisTitleField.append(chartXAxisTitleLabel, this.chartEditXAxisTitleInput);
    const chartYAxisTitleField = document.createElement("label");
    chartYAxisTitleField.className = "excelsior-find-replace-field";
    const chartYAxisTitleLabel = document.createElement("span");
    chartYAxisTitleLabel.textContent = this.messages.chartEditYAxisTitleLabel;
    chartYAxisTitleField.append(chartYAxisTitleLabel, this.chartEditYAxisTitleInput);
    const createChartEditField = (label: string, input: HTMLElement): HTMLLabelElement => {
      const field = document.createElement("label");
      field.className = "excelsior-find-replace-field";
      const text = document.createElement("span");
      text.textContent = label;
      field.append(text, input);
      return field;
    };
    const chartAnnotationTextField = createChartEditField("Anotação", this.chartEditAnnotationTextInput);
    const chartAnnotationXField = createChartEditField("Anotação X", this.chartEditAnnotationXInput);
    const chartAnnotationYField = createChartEditField("Anotação Y", this.chartEditAnnotationYInput);
    const chartAnnotationArrowToggle = document.createElement("label");
    chartAnnotationArrowToggle.className = "excelsior-find-replace-toggle";
    const chartAnnotationArrowText = document.createElement("span");
    chartAnnotationArrowText.textContent = "Seta na anotação";
    chartAnnotationArrowToggle.append(this.chartEditAnnotationArrowToggle, chartAnnotationArrowText);
    const chartShapeTypeField = createChartEditField("Shape", this.chartEditShapeTypeSelect);
    const chartShapeX0Field = createChartEditField("Shape X inicial", this.chartEditShapeX0Input);
    const chartShapeY0Field = createChartEditField("Shape Y inicial", this.chartEditShapeY0Input);
    const chartShapeX1Field = createChartEditField("Shape X final", this.chartEditShapeX1Input);
    const chartShapeY1Field = createChartEditField("Shape Y final", this.chartEditShapeY1Input);
    const chartXAxisTypeField = document.createElement("label");
    chartXAxisTypeField.className = "excelsior-find-replace-field";
    const chartXAxisTypeLabel = document.createElement("span");
    chartXAxisTypeLabel.textContent = this.messages.chartEditXAxisTypeLabel;
    chartXAxisTypeField.append(chartXAxisTypeLabel, this.chartEditXAxisTypeSelect);
    const chartYAxisTypeField = document.createElement("label");
    chartYAxisTypeField.className = "excelsior-find-replace-field";
    const chartYAxisTypeLabel = document.createElement("span");
    chartYAxisTypeLabel.textContent = this.messages.chartEditYAxisTypeLabel;
    chartYAxisTypeField.append(chartYAxisTypeLabel, this.chartEditYAxisTypeSelect);
    const chartLegendToggle = document.createElement("label");
    chartLegendToggle.className = "excelsior-find-replace-toggle";
    const chartLegendText = document.createElement("span");
    chartLegendText.textContent = this.messages.chartEditLegendLabel;
    chartLegendToggle.append(this.chartEditLegendToggle, chartLegendText);
    const chartXAxisVisibleToggle = document.createElement("label");
    chartXAxisVisibleToggle.className = "excelsior-find-replace-toggle";
    const chartXAxisVisibleText = document.createElement("span");
    chartXAxisVisibleText.textContent = this.messages.chartEditXAxisVisibleLabel;
    chartXAxisVisibleToggle.append(this.chartEditXAxisVisibleToggle, chartXAxisVisibleText);
    const chartYAxisVisibleToggle = document.createElement("label");
    chartYAxisVisibleToggle.className = "excelsior-find-replace-toggle";
    const chartYAxisVisibleText = document.createElement("span");
    chartYAxisVisibleText.textContent = this.messages.chartEditYAxisVisibleLabel;
    chartYAxisVisibleToggle.append(this.chartEditYAxisVisibleToggle, chartYAxisVisibleText);
    const chartRangeSelectorToggle = document.createElement("label");
    chartRangeSelectorToggle.className = "excelsior-find-replace-toggle";
    const chartRangeSelectorText = document.createElement("span");
    chartRangeSelectorText.textContent = "Seletor de faixa";
    chartRangeSelectorToggle.append(this.chartEditRangeSelectorToggle, chartRangeSelectorText);
    const chartRangeSliderToggle = document.createElement("label");
    chartRangeSliderToggle.className = "excelsior-find-replace-toggle";
    const chartRangeSliderText = document.createElement("span");
    chartRangeSliderText.textContent = "Slider de faixa";
    chartRangeSliderToggle.append(this.chartEditRangeSliderToggle, chartRangeSliderText);
    const chartOrientationField = document.createElement("label");
    chartOrientationField.className = "excelsior-find-replace-field";
    const chartOrientationLabel = document.createElement("span");
    chartOrientationLabel.textContent = this.messages.chartEditOrientationLabel;
    chartOrientationField.append(chartOrientationLabel, this.chartEditOrientationSelect);
    const chartCategoryColumnField = document.createElement("label");
    chartCategoryColumnField.className = "excelsior-find-replace-field";
    const chartCategoryColumnLabel = document.createElement("span");
    chartCategoryColumnLabel.textContent = this.messages.chartEditCategoryColumnLabel;
    chartCategoryColumnField.append(chartCategoryColumnLabel, this.chartEditCategoryColumnInput);
    const chartSeriesColumnsField = document.createElement("label");
    chartSeriesColumnsField.className = "excelsior-find-replace-field";
    const chartSeriesColumnsLabel = document.createElement("span");
    chartSeriesColumnsLabel.textContent = this.messages.chartEditSeriesColumnsLabel;
    chartSeriesColumnsField.append(chartSeriesColumnsLabel, this.chartEditSeriesColumnsInput);
    const chartValueColumnField = document.createElement("label");
    chartValueColumnField.className = "excelsior-find-replace-field";
    const chartValueColumnLabel = document.createElement("span");
    chartValueColumnLabel.textContent = this.messages.chartEditValueColumnLabel;
    chartValueColumnField.append(chartValueColumnLabel, this.chartEditValueColumnInput);
    const chartFirstRowHeaderToggle = document.createElement("label");
    chartFirstRowHeaderToggle.className = "excelsior-find-replace-toggle";
    const chartFirstRowHeaderText = document.createElement("span");
    chartFirstRowHeaderText.textContent = this.messages.chartEditFirstRowHeaderLabel;
    chartFirstRowHeaderToggle.append(this.chartEditFirstRowHeaderToggle, chartFirstRowHeaderText);
    const chartFirstColumnLabelToggle = document.createElement("label");
    chartFirstColumnLabelToggle.className = "excelsior-find-replace-toggle";
    const chartFirstColumnLabelText = document.createElement("span");
    chartFirstColumnLabelText.textContent = this.messages.chartEditFirstColumnLabelLabel;
    chartFirstColumnLabelToggle.append(this.chartEditFirstColumnLabelToggle, chartFirstColumnLabelText);
    const chartAutoRefreshToggle = document.createElement("label");
    chartAutoRefreshToggle.className = "excelsior-find-replace-toggle";
    const chartAutoRefreshText = document.createElement("span");
    chartAutoRefreshText.textContent = this.messages.chartEditAutoRefreshLabel;
    chartAutoRefreshToggle.append(this.chartEditAutoRefreshToggle, chartAutoRefreshText);
    this.chartEditPanelHeader.className = "excelsior-chart-panel-header";
    const chartHeader = document.createElement("span");
    chartHeader.className = "excelsior-chart-edit-panel-title";
    chartHeader.textContent = this.messages.chartEditPanelTitle;
    this.chartEditPanelHeader.replaceChildren(chartHeader);
    const chartActions = document.createElement("div");
    chartActions.className = "excelsior-find-replace-actions";
    const chartAddImageButton = createFindReplaceActionButton("add-image", "Adicionar imagem");
    chartAddImageButton.dataset.chartAction = "add-image";
    delete chartAddImageButton.dataset.findAction;
    const chartRemoveImageButton = createFindReplaceActionButton("remove-image", "Remover imagem");
    chartRemoveImageButton.dataset.chartAction = "remove-image";
    delete chartRemoveImageButton.dataset.findAction;
    chartActions.append(chartAddImageButton, chartRemoveImageButton, this.chartEditApplyButton, this.chartEditCloseButton);
    this.chartEditPanel.append(
      this.chartEditPanelHeader,
      chartTitleField,
      chartTypeField,
      chartRangeField,
      chartXAxisTitleField,
      chartYAxisTitleField,
      chartAnnotationTextField,
      chartAnnotationXField,
      chartAnnotationYField,
      chartAnnotationArrowToggle,
      chartShapeTypeField,
      chartShapeX0Field,
      chartShapeY0Field,
      chartShapeX1Field,
      chartShapeY1Field,
      chartXAxisTypeField,
      chartYAxisTypeField,
      chartOrientationField,
      chartCategoryColumnField,
      chartSeriesColumnsField,
      chartValueColumnField,
      chartLegendToggle,
      chartFirstRowHeaderToggle,
      chartFirstColumnLabelToggle,
      chartAutoRefreshToggle,
      chartXAxisVisibleToggle,
      chartYAxisVisibleToggle,
      chartRangeSelectorToggle,
      chartRangeSliderToggle,
      chartActions
    );

    this.chartInsertPreviewPanel.className = "excelsior-find-replace excelsior-chart-preview-panel";
    this.chartInsertPreviewPanel.hidden = true;
    this.chartInsertPreviewHost.className = "excelsior-chart-preview-panel-host";
    this.chartInsertPreviewInsertButton.textContent = this.messages.chartPreviewInsert;
    this.chartInsertPreviewInsertButton.dataset.chartAction = "preview-insert";
    delete this.chartInsertPreviewInsertButton.dataset.findAction;
    this.chartInsertPreviewCancelButton.textContent = this.messages.chartPreviewCancel;
    this.chartInsertPreviewCancelButton.dataset.chartAction = "preview-cancel";
    delete this.chartInsertPreviewCancelButton.dataset.findAction;
    this.chartInsertPreviewHeader.className = "excelsior-chart-panel-header";
    const chartPreviewHeader = document.createElement("span");
    chartPreviewHeader.className = "excelsior-chart-edit-panel-title";
    chartPreviewHeader.textContent = this.messages.chartPreviewTitle;
    this.chartInsertPreviewHeader.replaceChildren(chartPreviewHeader);
    const chartPreviewActions = document.createElement("div");
    chartPreviewActions.className = "excelsior-find-replace-actions";
    chartPreviewActions.append(this.chartInsertPreviewInsertButton, this.chartInsertPreviewCancelButton);
    this.chartInsertPreviewPanel.append(this.chartInsertPreviewHeader, this.chartInsertPreviewHost, chartPreviewActions);

    this.surface.append(
      this.cellsLayer,
      this.chartsLayer,
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
    this.root.append(this.chrome, this.formulaBar, this.activeCellAnnouncement, this.gridPanel, this.sheetTabs, this.imageFileInput, this.geoJsonFileInput, this.chartLayoutImageFileInput);
    this.gridPanel.append(this.viewport, this.rowHeaders, this.notePanel, this.chartEditPanel, this.chartInsertPreviewPanel);
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
    this.viewport.removeEventListener("mousedown", this.handleViewportMouseDown);
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
    this.geoJsonFileInput.removeEventListener("change", this.handleGeoJsonFileChange);
    this.chartLayoutImageFileInput.removeEventListener("change", this.handleChartLayoutImageFileChange);
    this.chrome.removeEventListener("click", this.handleColumnHeaderClick);
    this.chrome.removeEventListener("keydown", this.handleColumnHeaderKeyDown);
    this.chrome.removeEventListener("mousedown", this.handleHeaderResizeMouseDown);
    this.chrome.removeEventListener("dblclick", this.handleHeaderResizeDoubleClick);
    this.rowHeaders.removeEventListener("click", this.handleRowHeaderClick);
    this.rowHeaders.removeEventListener("mousedown", this.handleHeaderResizeMouseDown);
    this.rowHeaders.removeEventListener("dblclick", this.handleHeaderResizeDoubleClick);
    window.removeEventListener("mousemove", this.handleHeaderResizeMouseMove);
    window.removeEventListener("mouseup", this.handleHeaderResizeMouseUp);
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
    this.chartEditPanel.removeEventListener("input", this.handleChartEditPanelInput);
    this.chartEditPanel.removeEventListener("click", this.handleChartEditPanelClick);
    this.chartEditPanel.removeEventListener("keydown", this.handleChartEditPanelKeyDown);
    this.chartEditPanel.removeEventListener("compositionstart", this.handleCompositionStart);
    this.chartEditPanel.removeEventListener("compositionend", this.handleCompositionEnd);
    this.chartInsertPreviewPanel.removeEventListener("click", this.handleChartInsertPreviewClick);
    this.chartInsertPreviewPanel.removeEventListener("keydown", this.handleChartInsertPreviewKeyDown);
    this.chartInsertPreviewPanel.removeEventListener("compositionstart", this.handleCompositionStart);
    this.chartInsertPreviewPanel.removeEventListener("compositionend", this.handleCompositionEnd);
    this.chartEditPanelHeader.removeEventListener("mousedown", this.handleChartEditPanelMouseDown);
    this.chartInsertPreviewHeader.removeEventListener("mousedown", this.handleChartPreviewPanelMouseDown);
    globalThis.removeEventListener("mousemove", this.handleChartPanelMouseMove);
    globalThis.removeEventListener("mouseup", this.handleChartPanelMouseUp);
    this.cellsLayer.removeEventListener("click", this.handleCellClick);
    this.cellsLayer.removeEventListener("mousedown", this.handleAutofillMouseDown);
    this.cellsLayer.removeEventListener("mousemove", this.handleAutofillMouseMove);
    this.cellsLayer.removeEventListener("mouseup", this.handleAutofillMouseUp);
    this.cellsLayer.removeEventListener("dblclick", this.handleCellDoubleClick);
    this.chartsLayer.removeEventListener("mousedown", this.handleChartLayerMouseDown);
    this.chartsLayer.removeEventListener("click", this.handleChartLayerClick);
    this.chartsLayer.removeEventListener("keydown", this.handleVisualObjectKeyDown);
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
    globalThis.removeEventListener("mousemove", this.handleChartInteractionMouseMove);
    globalThis.removeEventListener("mouseup", this.handleChartInteractionMouseUp);
    globalThis.removeEventListener("mousemove", this.handleSplitPaneMouseMove);
    globalThis.removeEventListener("mouseup", this.handleSplitPaneMouseUp);
    this.destroyAllChartRuntimes();
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
      }),
      this.engine.on("chart:rangeChanged", ({ sheetId, chartId, reason, range }) => {
        if (reason === "source-cells-updated") {
          this.refreshBoundChartFigure(sheetId, chartId, range);
        }
      }),
      this.engine.on("chart:dataInvalid", ({ sheetId }) => {
        this.setChartFeedback(sheetId, "Dados inválidos no intervalo de origem de um gráfico.", true);
        this.requestRender();
      }),
      this.engine.on("chart:selected", ({ sheetId, chartId }) => {
        if (this.engine.getSnapshot().activeSheetId === sheetId) {
          this.selectedChartId = chartId;
          this.requestRender();
        }
      }),
      this.engine.on("chart:unselected", ({ sheetId, chartId }) => {
        if (this.engine.getSnapshot().activeSheetId === sheetId && this.selectedChartId === chartId) {
          this.selectedChartId = undefined;
          this.requestRender();
        }
      }),
      this.engine.on("chart:deleted", ({ sheetId, chartId }) => {
        if (this.selectedChartId === chartId && this.engine.getSnapshot().activeSheetId === sheetId) {
          this.selectedChartId = undefined;
        }
        if (this.chartInteraction?.chartId === chartId) {
          this.chartInteraction = undefined;
        }
        this.destroyChartRuntime(chartId);
        this.chartObjectElementById.get(chartId)?.remove();
        this.chartObjectElementById.delete(chartId);
        this.chartBodyElementById.delete(chartId);
      })
    );

    this.viewport.addEventListener("scroll", this.handleScroll);
    this.viewport.addEventListener("mousedown", this.handleViewportMouseDown);
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
    this.chrome.addEventListener("mousedown", this.handleHeaderResizeMouseDown);
    this.chrome.addEventListener("dblclick", this.handleHeaderResizeDoubleClick);
    this.rowHeaders.addEventListener("click", this.handleRowHeaderClick);
    this.rowHeaders.addEventListener("mousedown", this.handleHeaderResizeMouseDown);
    this.rowHeaders.addEventListener("dblclick", this.handleHeaderResizeDoubleClick);
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
    this.chartEditPanel.addEventListener("input", this.handleChartEditPanelInput);
    this.chartEditPanel.addEventListener("click", this.handleChartEditPanelClick);
    this.chartEditPanel.addEventListener("keydown", this.handleChartEditPanelKeyDown);
    this.chartEditPanel.addEventListener("compositionstart", this.handleCompositionStart);
    this.chartEditPanel.addEventListener("compositionend", this.handleCompositionEnd);
    this.chartInsertPreviewPanel.addEventListener("click", this.handleChartInsertPreviewClick);
    this.chartInsertPreviewPanel.addEventListener("keydown", this.handleChartInsertPreviewKeyDown);
    this.chartInsertPreviewPanel.addEventListener("compositionstart", this.handleCompositionStart);
    this.chartInsertPreviewPanel.addEventListener("compositionend", this.handleCompositionEnd);
    this.chartEditPanelHeader.addEventListener("mousedown", this.handleChartEditPanelMouseDown);
    this.chartInsertPreviewHeader.addEventListener("mousedown", this.handleChartPreviewPanelMouseDown);
    this.cellsLayer.addEventListener("click", this.handleCellClick);
    this.cellsLayer.addEventListener("mousedown", this.handleAutofillMouseDown);
    this.cellsLayer.addEventListener("mousemove", this.handleAutofillMouseMove);
    this.cellsLayer.addEventListener("mouseup", this.handleAutofillMouseUp);
    this.cellsLayer.addEventListener("dblclick", this.handleCellDoubleClick);
    this.chartsLayer.addEventListener("mousedown", this.handleChartLayerMouseDown);
    this.chartsLayer.addEventListener("click", this.handleChartLayerClick);
    this.chartsLayer.addEventListener("keydown", this.handleVisualObjectKeyDown);
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
    globalThis.addEventListener("mousemove", this.handleChartInteractionMouseMove);
    globalThis.addEventListener("mouseup", this.handleChartInteractionMouseUp);
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
    const throttleMs = this.getChartPerformanceOptions().interactionThrottleMs;
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
    if (!this.splitPaneDrag || !this.chartSurfaceMetrics) {
      return;
    }
    const rect = this.viewport.getBoundingClientRect();
    const offsets = this.splitPaneDrag.axis === "horizontal"
      ? this.chartSurfaceMetrics.rowOffsets
      : this.chartSurfaceMetrics.colOffsets;
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
    this.setChartSelection(sheet.id, undefined);
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
      this.setChartSelection(sheet.id, undefined);
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
    this.setChartSelection(sheet.id, undefined);
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
    this.setChartSelection(sheet.id, undefined);
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

  private readonly startChartPanelDrag = (panel: HTMLElement, event: MouseEvent): void => {
    const panelBounds = panel.getBoundingClientRect();
    this.chartPanelDragging = {
      panel,
      offsetX: event.clientX - panelBounds.left,
      offsetY: event.clientY - panelBounds.top
    };
    event.preventDefault();
    globalThis.addEventListener("mousemove", this.handleChartPanelMouseMove);
    globalThis.addEventListener("mouseup", this.handleChartPanelMouseUp);
  };

  private readonly handleChartEditPanelMouseDown = (event: MouseEvent): void => {
    this.startChartPanelDrag(this.chartEditPanel, event);
  };

  private readonly handleChartPreviewPanelMouseDown = (event: MouseEvent): void => {
    this.startChartPanelDrag(this.chartInsertPreviewPanel, event);
  };

  private readonly handleChartPanelMouseMove = (event: MouseEvent): void => {
    if (!this.chartPanelDragging) {
      return;
    }

    const { panel, offsetX, offsetY } = this.chartPanelDragging;
    const gridBounds = this.gridPanel.getBoundingClientRect();
    const panelWidth = panel.offsetWidth;
    const panelHeight = panel.offsetHeight;
    const nextLeft = this.clamp(event.clientX - gridBounds.left - offsetX, 0, Math.max(0, gridBounds.width - panelWidth));
    const nextTop = this.clamp(event.clientY - gridBounds.top - offsetY, 0, Math.max(0, gridBounds.height - panelHeight));

    panel.style.left = `${nextLeft}px`;
    panel.style.top = `${nextTop}px`;
    panel.style.right = "auto";
  };

  private readonly handleChartPanelMouseUp = (): void => {
    this.chartPanelDragging = undefined;
    globalThis.removeEventListener("mousemove", this.handleChartPanelMouseMove);
    globalThis.removeEventListener("mouseup", this.handleChartPanelMouseUp);
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

  private readonly handleChartEditPanelInput = (): void => {
    const sheet = this.engine.getActiveSheet();
    this.clearChartFeedback(sheet.id);
  };

  private readonly handleChartEditPanelClick = (event: Event): void => {
    const button = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-chart-action]");
    if (!button) {
      return;
    }
    switch (button.dataset.chartAction) {
      case "apply":
        this.applyChartEditPanelChanges();
        break;
      case "add-image":
        this.chartLayoutImageFileInput.click();
        break;
      case "remove-image": {
        const sheet = this.engine.getActiveSheet();
        const chart = this.selectedChartId ? this.engine.getChart(sheet.id, this.selectedChartId) : undefined;
        if (chart) {
          this.engine.updateChart({
            sheetId: sheet.id,
            chartId: chart.id,
            patch: { figure: { ...chart.figure, layout: { ...(chart.figure.layout as Record<string, unknown>), images: [] } } }
          });
          this.requestRender();
        }
        break;
      }
      case "close": {
        const sheet = this.engine.getActiveSheet();
        this.setChartSelection(sheet.id, undefined);
        this.requestRender();
        this.focus();
        break;
      }
      default:
        break;
    }
  };

  private readonly handleChartEditPanelKeyDown = (event: KeyboardEvent): void => {
    if (this.isCompositionInProgress(event.target, event)) {
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      const sheet = this.engine.getActiveSheet();
      this.setChartSelection(sheet.id, undefined);
      this.requestRender();
      this.focus();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      this.applyChartEditPanelChanges();
    }
  };

  private readonly handleChartInsertPreviewClick = (event: Event): void => {
    const button = (event.target as HTMLElement | null)?.closest<HTMLButtonElement>("[data-chart-action]");
    if (!button) {
      return;
    }
    if (button.dataset.chartAction === "preview-cancel") {
      this.closeChartInsertPreview();
      this.focus();
      return;
    }
    if (button.dataset.chartAction === "preview-insert") {
      this.commitPendingChartInsertion();
      this.focus();
    }
  };

  private readonly handleChartInsertPreviewKeyDown = (event: KeyboardEvent): void => {
    if (this.isCompositionInProgress(event.target, event)) {
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      this.closeChartInsertPreview();
      this.focus();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      this.commitPendingChartInsertion();
      this.focus();
    }
  };

  private readonly handleViewportMouseDown = (event: MouseEvent): void => {
    if (!this.chartInsertPreviewPanel.hidden) {
      this.closeChartInsertPreview();
    }
    if (!this.selectedChartId || event.button !== 0) {
      return;
    }
    const target = event.target as HTMLElement | null;
    if (target?.closest("[data-chart-id]")) {
      return;
    }
    const sheet = this.engine.getActiveSheet();
    this.setChartSelection(sheet.id, undefined);
    this.requestRender();
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
        if (action && this.isChartToolbarAction(action)) {
          this.createChartFromSelection(action);
          this.focus();
        }
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

  private readonly handleGeoJsonFileChange = (): void => {
    const file = this.geoJsonFileInput.files?.[0];
    this.geoJsonFileInput.value = "";
    if (!file || file.size > 5_000_000) {
      return;
    }
    const reader = new FileReader();
    reader.addEventListener("load", () => {
      try {
        if (typeof reader.result !== "string") {
          throw new Error("Não foi possível ler o arquivo GeoJSON.");
        }
        const parsed = JSON.parse(reader.result) as {
          type?: unknown;
          features?: unknown;
        };
        if (parsed.type !== "FeatureCollection" || !Array.isArray(parsed.features) || !parsed.features.length) {
          throw new Error("O arquivo deve conter uma FeatureCollection GeoJSON.");
        }
        const features = parsed.features.map((raw, index) => {
          const feature = raw as { type?: unknown; id?: unknown; geometry?: { type?: unknown; coordinates?: unknown }; properties?: unknown };
          if (
            feature.type !== "Feature" ||
            !feature.geometry ||
            (feature.geometry.type !== "Polygon" && feature.geometry.type !== "MultiPolygon") ||
            !Array.isArray(feature.geometry.coordinates)
          ) {
            throw new Error(`Feature GeoJSON inválida na posição ${index + 1}.`);
          }
          const rawProperties = feature.properties && typeof feature.properties === "object" && !Array.isArray(feature.properties)
            ? feature.properties as Record<string, unknown>
            : {};
          const properties = Object.fromEntries(
            Object.entries(rawProperties).filter((entry): entry is [string, string | number | boolean | null] => {
              const value = entry[1];
              return value === null || typeof value === "string" || typeof value === "number" || typeof value === "boolean";
            })
          );
          const sourceId = feature.id ?? properties.id ?? index;
          const safeId = typeof sourceId === "string" || typeof sourceId === "number" ? String(sourceId) : String(index);
          return {
            type: "Feature" as const,
            geometry: {
              type: feature.geometry.type,
              coordinates: cloneSerializable(feature.geometry.coordinates)
            },
            properties: {
              ...properties,
              __excelsiorId: safeId
            }
          };
        });
        const sheet = this.engine.getActiveSheet();
        const sourceRange = this.getChartSourceRange(sheet);
        if (!sourceRange) {
          throw new Error(this.messages.chartInvalidRange);
        }
        const binding = this.createChartBindingOptions(sheet.id, sourceRange);
        const rangeInput = this.createSpreadsheetRangeInput(sheet.id, sourceRange, binding);
        const categoryIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
        const valueIndex = this.getNumericChartSeriesIndexes(rangeInput, binding, categoryIndex)[0];
        const pairs = rangeInput.rows
          .map((row) => ({ location: String(row[categoryIndex] ?? ""), value: toNumericValue((row[valueIndex] ?? null) as CellPrimitive) }))
          .filter((item): item is { location: string; value: number } => Boolean(item.location) && item.value !== undefined);
        const propertyKeys = Object.keys(features[0]?.properties ?? {});
        const locationSet = new Set(pairs.map((item) => item.location));
        const featureIdField = propertyKeys.reduce(
          (best, key) => {
            const score = features.filter((feature) => locationSet.has(String((feature.properties as Record<string, unknown>)[key] ?? ""))).length;
            return score > best.score ? { key, score } : best;
          },
          { key: "__excelsiorId", score: 0 }
        ).key;
        const title = this.sanitizeChartText(`${this.messages.chartGeo} (${binding.rangeAddress})`, 180);
        this.engine.createChart({
          sheetId: sheet.id,
          chart: {
            type: "geo",
            title,
            sourceRange: binding,
            figure: {
              data: [{
                type: "geo",
                geojson: { type: "FeatureCollection", features },
                locations: pairs.map((item) => item.location),
                values: pairs.map((item) => item.value),
                featureIdField,
                showColorLegend: true
              }],
              layout: { title },
              metadata: { source: "spreadsheet-range", chartType: "geo" }
            },
            position: this.createDefaultChartPosition(sheet, sourceRange),
            state: { selected: true, visible: true, locked: false }
          }
        });
        this.requestRender();
      } catch (error) {
        const sheet = this.engine.getActiveSheet();
        this.setChartFeedback(sheet.id, error instanceof Error ? error.message : this.messages.chartInsertError, true);
        this.requestRender();
      }
    }, { once: true });
    reader.readAsText(file);
  };

  private readonly handleChartLayoutImageFileChange = (): void => {
    const file = this.chartLayoutImageFileInput.files?.[0];
    this.chartLayoutImageFileInput.value = "";
    if (!file || file.size > 2_000_000 || !["image/png", "image/jpeg", "image/gif", "image/webp"].includes(file.type)) {
      return;
    }
    const sheet = this.engine.getActiveSheet();
    const chartId = this.selectedChartId;
    if (!chartId) {
      return;
    }
    const reader = new FileReader();
    reader.addEventListener("load", () => {
      if (typeof reader.result !== "string" || !reader.result.startsWith("data:image/")) {
        return;
      }
      const chart = this.engine.getChart(sheet.id, chartId);
      if (!chart) {
        return;
      }
      const layout = chart.figure.layout && typeof chart.figure.layout === "object"
        ? chart.figure.layout as Record<string, unknown>
        : {};
      this.engine.updateChart({
        sheetId: sheet.id,
        chartId,
        patch: {
          figure: {
            ...chart.figure,
            layout: {
              ...layout,
              images: [{ source: reader.result, x: 0.08, y: 0.08, width: 120, height: 80, xRef: "paper", yRef: "paper", opacity: 0.9 }]
            }
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

    this.closeChartInsertPreview();
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
    this.setChartSelection(sheet.id, undefined);
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
    const chartStatus = this.chartFeedback?.sheetId === sheet.id ? this.chartFeedback.message : undefined;
    const chartStatusIsError = this.chartFeedback?.sheetId === sheet.id ? this.chartFeedback.isError : false;
    this.formulaAddress.textContent = label;
    this.statusMessage.textContent =
      validationMessage ??
      rowModelError ??
      pivotStatus ??
      chartStatus ??
      (this.hasPendingRowModelRequests(sheet.id) ? this.messages.loadingRows : "");
    this.statusMessage.classList.toggle(
      "is-error",
      Boolean((validationMessage && !this.validationFeedback?.isWarning) || rowModelError) || pivotStatusIsError || chartStatusIsError
    );
    if (document.activeElement !== this.formulaInput) {
      this.formulaInput.value = value;
    }
    this.renderFindReplacePanel();
    this.renderPivotPanel();
    this.renderChartEditPanel();
  }

  private clearValidationFeedback(): void {
    this.validationFeedback = undefined;
  }

  private renderChrome(columnsToRender: number[]): void {
    const sheet = this.engine.getActiveSheet();
    const toolbarScrollLeft = this.toolbar.scrollLeft;
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
      const resizeHandle = document.createElement("span");
      resizeHandle.className = "excelsior-column-resize-handle";
      resizeHandle.dataset.columnResize = String(col);
      resizeHandle.setAttribute("aria-label", `Redimensionar coluna ${columnIndexToLabel(col)}; clique duplo para ajustar ao conteúdo`);
      resizeHandle.title = "Arraste para redimensionar; clique duplo para ajustar ao conteúdo";
      header.append(resizeHandle);
      columnStrip.append(header);
    }
    const headerRow = document.createElement("div");
    headerRow.className = "excelsior-column-header-row";
    headerRow.append(corner, columnStrip);

    fragment.append(this.renderToolbar(), headerRow);
    this.chrome.append(fragment);
    this.toolbar.scrollLeft = toolbarScrollLeft;
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
      const resizeHandle = document.createElement("span");
      resizeHandle.className = "excelsior-row-resize-handle";
      resizeHandle.dataset.rowResize = String(row);
      resizeHandle.setAttribute("aria-label", `Redimensionar linha ${row + 1}; clique duplo para ajustar ao conteúdo`);
      resizeHandle.title = "Arraste para redimensionar; clique duplo para ajustar ao conteúdo";
      header.append(resizeHandle);
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
      "zoom-reset": `${Math.round(this.sheetZoom * 100)}%`,
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
      "add-sheet": "+",
      "chart-column": "▥",
      "chart-bar": "☰",
      "chart-line": "╱",
      "chart-area": "◭",
      "chart-pie": "◔",
      "chart-donut": "◍",
      "chart-scatter": "·",
      "chart-histogram": "▤",
      "chart-heatmap": "▦"
    };

    return iconByAction[action] ?? "•";
  }

  private isChartToolbarAction(action: string): action is ChartToolbarAction {
    return CHART_ACTIONS.has(action as ChartToolbarAction);
  }

  private createChartToolbarIcon(action: ChartToolbarAction): SVGSVGElement {
    const svg = document.createElementNS(CHART_SVG_NS, "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("aria-hidden", "true");
    svg.classList.add("excelsior-toolbar-icon-svg");

    const stroke = document.createElementNS(CHART_SVG_NS, "path");
    stroke.setAttribute("fill", "none");
    stroke.setAttribute("stroke", "currentColor");
    stroke.setAttribute("stroke-width", "1.8");
    stroke.setAttribute("stroke-linecap", "round");
    stroke.setAttribute("stroke-linejoin", "round");

    const fill = document.createElementNS(CHART_SVG_NS, "path");
    fill.setAttribute("fill", "currentColor");

    switch (action) {
      case "chart-column":
        fill.setAttribute("d", "M4 19h3v-7H4zm6 0h3V6h-3zm6 0h3v-10h-3z");
        break;
      case "chart-bar":
        fill.setAttribute("d", "M5 7h11v3H5zm0 5h15v3H5zm0 5h8v3H5z");
        break;
      case "chart-line":
        stroke.setAttribute("d", "M4 18L9 12l4 3 7-9");
        svg.append(stroke);
        return svg;
      case "chart-area":
        fill.setAttribute("d", "M4 18V8l5 4 4-3 7 9H4z");
        break;
      case "chart-pie":
        stroke.setAttribute("d", "M12 3a9 9 0 1 0 9 9h-9z");
        fill.setAttribute("d", "M12 3v9h9A9 9 0 0 0 12 3z");
        svg.append(stroke, fill);
        return svg;
      case "chart-donut":
        stroke.setAttribute("d", "M12 4a8 8 0 1 0 8 8h-8z");
        fill.setAttribute("d", "M12 4v8h8A8 8 0 0 0 12 4z");
        const hole = document.createElementNS(CHART_SVG_NS, "circle");
        hole.setAttribute("cx", "12");
        hole.setAttribute("cy", "12");
        hole.setAttribute("r", "4");
        hole.setAttribute("fill", "white");
        svg.append(stroke, fill, hole);
        return svg;
      case "chart-scatter":
        fill.setAttribute("d", "M6 16a2 2 0 1 0 0.01 0zm6-6a2 2 0 1 0 0.01 0zm6 4a2 2 0 1 0 0.01 0z");
        break;
      case "chart-histogram":
        fill.setAttribute("d", "M4 19h3v-4H4zm4 0h3V9H8zm4 0h3V6h-3zm4 0h3v-7h-3z");
        break;
      case "chart-box":
        stroke.setAttribute("d", "M6 8h12v8H6zM12 4v4m0 8v4M8 12h8");
        svg.append(stroke);
        return svg;
      case "chart-heatmap":
        fill.setAttribute("d", "M4 4h5v5H4zm6 0h5v5h-5zm6 0h4v5h-4zM4 10h5v5H4zm6 0h5v5h-5zm6 0h4v5h-4zM4 16h5v4H4zm6 0h5v4h-5zm6 0h4v4h-4z");
        break;
      case "chart-candlestick":
        stroke.setAttribute("d", "M7 5v14m10-12v12M5 9h4v6H5zm10-4h4v8h-4z");
        svg.append(stroke);
        return svg;
      case "chart-waterfall":
        fill.setAttribute("d", "M4 17h4V9H4zm6 0h4V12h-4zm6 0h4V6h-4");
        stroke.setAttribute("d", "M6 9h6m0 3h6");
        svg.append(fill, stroke);
        return svg;
      case "chart-funnel":
        fill.setAttribute("d", "M4 5h16l-5 6v6l-6 2v-8z");
        break;
      case "chart-polar":
        stroke.setAttribute("d", "M12 4v16M4 12h16M6 6l12 12M18 6L6 18");
        const polarCircle = document.createElementNS(CHART_SVG_NS, "circle");
        polarCircle.setAttribute("cx", "12");
        polarCircle.setAttribute("cy", "12");
        polarCircle.setAttribute("r", "7");
        polarCircle.setAttribute("fill", "none");
        polarCircle.setAttribute("stroke", "currentColor");
        polarCircle.setAttribute("stroke-width", "1.2");
        svg.append(stroke, polarCircle);
        return svg;
      case "chart-treemap":
        fill.setAttribute("d", "M4 4h8v8H4zm9 0h7v5h-7zM13 10h7v10h-7zM4 13h8v7H4z");
        break;
      case "chart-sunburst":
        stroke.setAttribute("d", "M12 3a9 9 0 1 0 9 9");
        const ring = document.createElementNS(CHART_SVG_NS, "circle");
        ring.setAttribute("cx", "12");
        ring.setAttribute("cy", "12");
        ring.setAttribute("r", "5");
        ring.setAttribute("fill", "none");
        ring.setAttribute("stroke", "currentColor");
        ring.setAttribute("stroke-width", "1.2");
        svg.append(stroke, ring);
        return svg;
      case "chart-sankey":
        stroke.setAttribute("d", "M4 7h6c4 0 4 10 8 10h2M4 17h4c4 0 4-10 8-10h4");
        svg.append(stroke);
        return svg;
      case "chart-surface":
      case "chart-surface3d":
        stroke.setAttribute("d", "M4 16l5-8 5 4 6-6M4 18h16M6 18V8m6 10V10m6 8V6");
        svg.append(stroke);
        return svg;
      case "chart-scatter3d":
        fill.setAttribute("d", "M6 15a2 2 0 1 0 .01 0zm6-5a2 2 0 1 0 .01 0zm6 3a2 2 0 1 0 .01 0z");
        stroke.setAttribute("d", "M6 15l6-5 6 3");
        svg.append(stroke, fill);
        return svg;
      default:
        fill.setAttribute("d", "M4 4h16v16H4z");
    }

    svg.append(fill);
    return svg;
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

    type ToolbarGroupKey = "data" | "font" | "alignment" | "structure" | "charts";
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
      { key: "structure", label: this.messages.toolbarStructureGroup },
      { key: "charts", label: this.messages.toolbarChartsGroup }
    ];
    const ribbon = document.createElement("div");
    ribbon.className = "excelsior-toolbar-ribbon";
    const groupControls = new Map<ToolbarGroupKey, HTMLElement>();
    const chartCategoryControls = new Map<ChartToolbarCategory, HTMLElement>();

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

    const chartsContainer = groupControls.get("charts");
    if (chartsContainer) {
      const chartCategories: Array<{ key: ChartToolbarCategory; label: string }> = [
        { key: "common", label: this.messages.toolbarChartsCommonGroup },
        { key: "statistical", label: this.messages.toolbarChartsStatisticalGroup },
        { key: "financial", label: this.messages.toolbarChartsFinancialGroup },
        { key: "advanced", label: this.messages.toolbarChartsAdvancedGroup }
      ];
      for (const category of chartCategories) {
        const wrapper = document.createElement("div");
        wrapper.className = "excelsior-toolbar-chart-category";
        wrapper.dataset.chartCategory = category.key;
        const categoryLabel = document.createElement("span");
        categoryLabel.className = "excelsior-toolbar-chart-category-label";
        categoryLabel.textContent = category.label;
        const row = document.createElement("div");
        row.className = "excelsior-toolbar-chart-category-controls";
        wrapper.append(categoryLabel, row);
        chartsContainer.append(wrapper);
        chartCategoryControls.set(category.key, row);
      }
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
      const chartAction = this.isChartToolbarAction(item.action) ? item.action : undefined;
      const chartIsPlaceholder = chartAction ? CHART_PLACEHOLDER_ACTIONS.has(chartAction) : false;
      if (chartAction) {
        button.classList.add("has-svg-icon");
        button.append(this.createChartToolbarIcon(chartAction));
      } else {
        button.dataset.icon = this.getToolbarActionIcon(item.action);
      }
      button.title = chartIsPlaceholder ? `${item.label} (${this.messages.chartUnsupportedType})` : item.label;
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

    for (const chartDefinition of CHART_TOOLBAR_DEFINITIONS) {
      const button = document.createElement("button");
      const label = this.messages[chartDefinition.messageKey] ?? chartDefinition.type;
      button.type = "button";
      button.className = "excelsior-toolbar-button has-svg-icon";
      button.dataset.action = chartDefinition.action;
      button.setAttribute("aria-label", label);
      button.title = chartDefinition.enabled ? label : `${label} (${this.messages.chartUnsupportedType})`;
      button.append(this.createChartToolbarIcon(chartDefinition.action));
      if (!chartDefinition.enabled) {
        button.disabled = true;
      }
      chartCategoryControls.get(chartDefinition.category)?.append(button);
    }

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
    this.surface.style.width = `${ROW_HEADER_WIDTH + (colOffsets[colOffsets.length - 1] ?? 0)}px`;
    this.surface.style.height = `${rowOffsets[rowOffsets.length - 1] ?? 0}px`;
    this.chartSurfaceMetrics = {
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
          content.classList.add("is-rotated");
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
    this.renderChartObjects(sheet, rowOffsets, colOffsets);

    if (this.editingCell) {
      this.positionEditor(this.editingCell.row, this.editingCell.col);
    }
  };

  private setChartFeedback(sheetId: string, message: string, isError = false): void {
    this.chartFeedback = {
      sheetId,
      message,
      isError
    };
  }

  private clearChartFeedback(sheetId?: string): void {
    if (!sheetId || this.chartFeedback?.sheetId === sheetId) {
      this.chartFeedback = undefined;
    }
  }

  private isChartCompatibleRange(range: CellRange | undefined): range is CellRange {
    return !!range && range.end.row > range.start.row && range.end.col > range.start.col;
  }

  private getChartSourceRange(sheet: ReturnType<WorkbookEngine["getActiveSheet"]>): CellRange | undefined {
    const selectionRange = this.normalizeRange(this.getResolvedSelectionRange(sheet));
    if (this.isChartCompatibleRange(selectionRange)) {
      return selectionRange;
    }

    const usedRange = this.getUsedRange(sheet);
    return this.isChartCompatibleRange(usedRange) ? usedRange : undefined;
  }

  private parseChartRangeAddress(rangeAddress: string): CellRange | undefined {
    const parts = rangeAddress
      .split(":")
      .map((part) => part.trim())
      .filter(Boolean);
    if (parts.length === 0 || parts.length > 2) {
      return undefined;
    }

    try {
      const start = cellLabelToAddress(parts[0] as string);
      const end = cellLabelToAddress((parts[1] ?? parts[0]) as string);
      return this.normalizeRange({ start, end });
    } catch {
      return undefined;
    }
  }

  private buildChartRangeAddress(range: CellRange): string {
    return `${cellAddressToLabel(range.start)}:${cellAddressToLabel(range.end)}`;
  }

  private createChartBindingOptions(_sheetId: string, sourceRange: CellRange): ChartBindingOptions {
    return {
      rangeAddress: this.buildChartRangeAddress(sourceRange),
      orientation: "rows",
      firstRowAsHeader: true,
      firstColumnAsLabel: true,
      autoRefresh: true,
      categoryColumnIndex: 0
    };
  }

  private getChartBindingOptions(chart: WorksheetChartObject): ChartBindingOptions | undefined {
    const sourceRange = chart.sourceRange;
    if (!sourceRange) {
      return undefined;
    }
    return {
      rangeAddress: sourceRange.rangeAddress,
      orientation: sourceRange.orientation,
      firstRowAsHeader: sourceRange.firstRowAsHeader,
      firstColumnAsLabel: sourceRange.firstColumnAsLabel,
      autoRefresh: sourceRange.autoRefresh,
      categoryColumnIndex: sourceRange.categoryColumnIndex,
      seriesColumnIndexes: sourceRange.seriesColumnIndexes ? [...sourceRange.seriesColumnIndexes] : undefined,
      valueColumnIndex: sourceRange.valueColumnIndex
    };
  }

  private formatChartColumnIndexInput(value: number | undefined): string {
    return Number.isInteger(value) && (value as number) >= 0 ? String((value as number) + 1) : "";
  }

  private formatChartColumnIndexListInput(values: number[] | undefined): string {
    if (!Array.isArray(values) || values.length === 0) {
      return "";
    }
    return values.map((value) => String(value + 1)).join(",");
  }

  private parseChartColumnIndexInput(raw: string, fieldLabel: string): number | undefined {
    const normalized = raw.trim();
    if (!normalized) {
      return undefined;
    }
    if (!/^\d+$/.test(normalized)) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_BINDING_INVALID",
        message: `${this.messages.chartEditInvalidBinding} (${fieldLabel})`,
        area: "renderer",
        recoverable: true
      });
    }
    const parsed = Number(normalized);
    if (!Number.isInteger(parsed) || parsed < 1) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_BINDING_INVALID",
        message: `${this.messages.chartEditInvalidBinding} (${fieldLabel})`,
        area: "renderer",
        recoverable: true
      });
    }
    return parsed - 1;
  }

  private parseChartColumnIndexListInput(raw: string, fieldLabel: string): number[] | undefined {
    const normalized = raw.trim();
    if (!normalized) {
      return undefined;
    }
    const parts = normalized
      .split(/[,\s;]+/)
      .map((part) => part.trim())
      .filter(Boolean);
    if (!parts.length) {
      return undefined;
    }
    const unique: number[] = [];
    for (const part of parts) {
      if (!/^\d+$/.test(part)) {
        throw new SpreadsheetOperationError({
          code: "RENDERER_CHART_BINDING_INVALID",
          message: `${this.messages.chartEditInvalidBinding} (${fieldLabel})`,
          area: "renderer",
          recoverable: true
        });
      }
      const parsed = Number(part);
      if (!Number.isInteger(parsed) || parsed < 1) {
        throw new SpreadsheetOperationError({
          code: "RENDERER_CHART_BINDING_INVALID",
          message: `${this.messages.chartEditInvalidBinding} (${fieldLabel})`,
          area: "renderer",
          recoverable: true
        });
      }
      const normalizedIndex = parsed - 1;
      if (!unique.includes(normalizedIndex)) {
        unique.push(normalizedIndex);
      }
    }
    return unique.length ? unique : undefined;
  }

  private areChartSeriesColumnsEqual(left: number[] | undefined, right: number[] | undefined): boolean {
    if (!left?.length && !right?.length) {
      return true;
    }
    if (!left || !right || left.length !== right.length) {
      return false;
    }
    return left.every((value, index) => value === right[index]);
  }

  private areChartBindingOptionsEqual(left: ChartBindingOptions | undefined, right: ChartBindingOptions): boolean {
    if (!left) {
      return false;
    }
    return (
      left.rangeAddress === right.rangeAddress &&
      left.orientation === right.orientation &&
      left.firstRowAsHeader === right.firstRowAsHeader &&
      left.firstColumnAsLabel === right.firstColumnAsLabel &&
      left.autoRefresh === right.autoRefresh &&
      left.categoryColumnIndex === right.categoryColumnIndex &&
      left.valueColumnIndex === right.valueColumnIndex &&
      this.areChartSeriesColumnsEqual(left.seriesColumnIndexes, right.seriesColumnIndexes)
    );
  }

  private isChartActionEnabled(action: ChartToolbarAction): boolean {
    return !CHART_DISABLED_ACTIONS.has(action);
  }

  private isChartTypeEnabled(type: WorksheetChartType): boolean {
    const match = CHART_TOOLBAR_DEFINITIONS.find((definition) => definition.type === type);
    return match?.enabled ?? false;
  }

  private getChartLegendVisible(chart: WorksheetChartObject): boolean {
    const layout =
      typeof chart.figure.layout === "object" && chart.figure.layout !== null
        ? (chart.figure.layout as Record<string, unknown>)
        : undefined;
    const legend = layout?.legend;
    if (typeof legend === "object" && legend !== null && "visible" in legend) {
      return Boolean((legend as { visible?: unknown }).visible ?? true);
    }
    return true;
  }

  private getChartAxisRecord(chart: WorksheetChartObject, axis: "x" | "y"): Record<string, unknown> | undefined {
    const layout =
      typeof chart.figure.layout === "object" && chart.figure.layout !== null
        ? (chart.figure.layout as Record<string, unknown>)
        : undefined;
    if (!layout) {
      return undefined;
    }
    const keys = axis === "x" ? ["xAxis", "xaxis"] : ["yAxis", "yaxis"];
    for (const key of keys) {
      const axisConfig = layout[key];
      if (axisConfig && typeof axisConfig === "object") {
        return axisConfig as Record<string, unknown>;
      }
    }
    return undefined;
  }

  private getChartAxisTitle(chart: WorksheetChartObject, axis: "x" | "y"): string {
    const axisConfig = this.getChartAxisRecord(chart, axis);
    if (!axisConfig) {
      return "";
    }
    const title = axisConfig.title;
    if (typeof title === "string") {
      return title;
    }
    if (title && typeof title === "object") {
      const text = (title as Record<string, unknown>).text;
      if (typeof text === "string") {
        return text;
      }
    }
    return "";
  }

  private getChartAxisType(chart: WorksheetChartObject, axis: "x" | "y"): ChartAxisTypeOption {
    const axisConfig = this.getChartAxisRecord(chart, axis);
    const type = typeof axisConfig?.type === "string" ? axisConfig.type.toLowerCase() : "";
    if (CHART_AXIS_TYPE_OPTIONS.includes(type as ChartAxisTypeOption)) {
      return type as ChartAxisTypeOption;
    }
    return "linear";
  }

  private getChartAxisVisible(chart: WorksheetChartObject, axis: "x" | "y"): boolean {
    const axisConfig = this.getChartAxisRecord(chart, axis);
    const visible = axisConfig?.visible;
    return visible !== false;
  }

  private withChartAxisOptions(
    figure: WorksheetChartObject["figure"],
    options: {
      xAxisTitle?: string;
      yAxisTitle?: string;
      xAxisType?: ChartAxisTypeOption;
      yAxisType?: ChartAxisTypeOption;
      xAxisVisible?: boolean;
      yAxisVisible?: boolean;
    }
  ): WorksheetChartObject["figure"] {
    const nextFigure: WorksheetChartObject["figure"] = {
      ...figure,
      data: Array.isArray(figure.data) ? figure.data.map((trace) => ({ ...(trace as Record<string, unknown>) })) : []
    };
    const layout =
      typeof figure.layout === "object" && figure.layout !== null
        ? ({ ...(figure.layout as Record<string, unknown>) } as Record<string, unknown>)
        : {};
    const xAxis = (layout.xAxis && typeof layout.xAxis === "object" ? { ...(layout.xAxis as Record<string, unknown>) } : {}) as Record<
      string,
      unknown
    >;
    const yAxis = (layout.yAxis && typeof layout.yAxis === "object" ? { ...(layout.yAxis as Record<string, unknown>) } : {}) as Record<
      string,
      unknown
    >;
    if (options.xAxisTitle) {
      xAxis.title = options.xAxisTitle;
    } else {
      delete xAxis.title;
    }
    if (options.yAxisTitle) {
      yAxis.title = options.yAxisTitle;
    } else {
      delete yAxis.title;
    }
    if (options.xAxisType) {
      xAxis.type = options.xAxisType;
    }
    if (options.yAxisType) {
      yAxis.type = options.yAxisType;
    }
    if (options.xAxisVisible !== undefined) {
      xAxis.visible = options.xAxisVisible;
    }
    if (options.yAxisVisible !== undefined) {
      yAxis.visible = options.yAxisVisible;
    }
    layout.xAxis = xAxis;
    layout.yAxis = yAxis;
    nextFigure.layout = layout;
    return nextFigure;
  }

  private setChartSelection(sheetId: string, chartId?: string): void {
    const charts = this.engine.getCharts(sheetId);
    for (const chart of charts) {
      const shouldSelect = chart.id === chartId;
      if (chart.state.selected === shouldSelect) {
        continue;
      }
      try {
        this.engine.updateChart({
          sheetId,
          chartId: chart.id,
          patch: {
            state: {
              ...chart.state,
              selected: shouldSelect
            }
          }
        });
      } catch {
        // Selection sync should not crash renderer interactions.
      }
    }
    this.selectedChartId = chartId;
  }

  private renderChartEditPanel(): void {
    const sheet = this.engine.getActiveSheet();
    const selectedChart = this.selectedChartId ? this.engine.getChart(sheet.id, this.selectedChartId) : undefined;
    this.chartEditPanel.hidden = !selectedChart;
    if (!selectedChart) {
      return;
    }

    if (!Array.from(this.chartEditTypeSelect.options).some((option) => option.value === selectedChart.type)) {
      const option = document.createElement("option");
      option.value = selectedChart.type;
      option.textContent = `${selectedChart.type} (${this.messages.chartUnsupportedType})`;
      option.disabled = false;
      this.chartEditTypeSelect.append(option);
    }

    const binding = this.getChartBindingOptions(selectedChart);
    this.chartEditTypeSelect.value = selectedChart.type;
    if (document.activeElement !== this.chartEditRangeInput) {
      this.chartEditRangeInput.value = binding?.rangeAddress ?? "";
    }
    if (document.activeElement !== this.chartEditTitleInput) {
      this.chartEditTitleInput.value = selectedChart.title ?? "";
    }
    this.chartEditOrientationSelect.value = binding?.orientation ?? "rows";
    this.chartEditFirstRowHeaderToggle.checked = binding?.firstRowAsHeader ?? true;
    this.chartEditFirstColumnLabelToggle.checked = binding?.firstColumnAsLabel ?? true;
    this.chartEditAutoRefreshToggle.checked = binding?.autoRefresh ?? true;
    if (document.activeElement !== this.chartEditCategoryColumnInput) {
      this.chartEditCategoryColumnInput.value = this.formatChartColumnIndexInput(binding?.categoryColumnIndex);
    }
    if (document.activeElement !== this.chartEditSeriesColumnsInput) {
      this.chartEditSeriesColumnsInput.value = this.formatChartColumnIndexListInput(binding?.seriesColumnIndexes);
    }
    if (document.activeElement !== this.chartEditValueColumnInput) {
      this.chartEditValueColumnInput.value = this.formatChartColumnIndexInput(binding?.valueColumnIndex);
    }
    if (document.activeElement !== this.chartEditXAxisTitleInput) {
      this.chartEditXAxisTitleInput.value = this.getChartAxisTitle(selectedChart, "x");
    }
    if (document.activeElement !== this.chartEditYAxisTitleInput) {
      this.chartEditYAxisTitleInput.value = this.getChartAxisTitle(selectedChart, "y");
    }
    this.chartEditXAxisTypeSelect.value = this.getChartAxisType(selectedChart, "x");
    this.chartEditYAxisTypeSelect.value = this.getChartAxisType(selectedChart, "y");
    this.chartEditXAxisVisibleToggle.checked = this.getChartAxisVisible(selectedChart, "x");
    this.chartEditYAxisVisibleToggle.checked = this.getChartAxisVisible(selectedChart, "y");
    const selectedLayout = selectedChart.figure.layout as { xAxis?: { rangeSelector?: { visible?: boolean }; rangeSlider?: { visible?: boolean } } } | undefined;
    this.chartEditRangeSelectorToggle.checked = selectedLayout?.xAxis?.rangeSelector?.visible === true;
    this.chartEditRangeSliderToggle.checked = selectedLayout?.xAxis?.rangeSlider?.visible === true;
    this.chartEditLegendToggle.checked = this.getChartLegendVisible(selectedChart);
    const layout = selectedChart.figure.layout as { annotations?: Array<Record<string, unknown>>; shapes?: Array<Record<string, unknown>> } | undefined;
    const annotation = layout?.annotations?.[0];
    const shape = layout?.shapes?.[0];
    this.chartEditAnnotationTextInput.value = typeof annotation?.text === "string" ? annotation.text : "";
    this.chartEditAnnotationXInput.value = String(typeof annotation?.x === "number" ? annotation.x : 0.5);
    this.chartEditAnnotationYInput.value = String(typeof annotation?.y === "number" ? annotation.y : 0.5);
    this.chartEditAnnotationArrowToggle.checked = annotation?.showArrow === true;
    this.chartEditShapeTypeSelect.value = typeof shape?.type === "string" ? shape.type : "";
    this.chartEditShapeX0Input.value = String(typeof shape?.x0 === "number" ? shape.x0 : 0.2);
    this.chartEditShapeY0Input.value = String(typeof shape?.y0 === "number" ? shape.y0 : 0.2);
    this.chartEditShapeX1Input.value = String(typeof shape?.x1 === "number" ? shape.x1 : 0.8);
    this.chartEditShapeY1Input.value = String(typeof shape?.y1 === "number" ? shape.y1 : 0.8);
    this.chartEditValueColumnInput.disabled = !(selectedChart.type === "pie" || selectedChart.type === "donut");
  }

  private applyChartEditPanelChanges(): void {
    const sheet = this.engine.getActiveSheet();
    const selectedChartId = this.selectedChartId;
    if (!selectedChartId) {
      return;
    }
    let chart = this.engine.getChart(sheet.id, selectedChartId);
    if (!chart) {
      return;
    }

    const nextType = this.chartEditTypeSelect.value.trim() as WorksheetChartType;
    const nextTitle = this.sanitizeChartText(this.chartEditTitleInput.value, 180);
    const nextLegendVisible = this.chartEditLegendToggle.checked;
    const nextXAxisTitle = this.sanitizeChartText(this.chartEditXAxisTitleInput.value, 120);
    const nextYAxisTitle = this.sanitizeChartText(this.chartEditYAxisTitleInput.value, 120);
    const nextXAxisType = (this.chartEditXAxisTypeSelect.value as ChartAxisTypeOption | undefined) ?? "linear";
    const nextYAxisType = (this.chartEditYAxisTypeSelect.value as ChartAxisTypeOption | undefined) ?? "linear";
    const nextXAxisVisible = this.chartEditXAxisVisibleToggle.checked;
    const nextYAxisVisible = this.chartEditYAxisVisibleToggle.checked;
    const nextRangeSelectorVisible = this.chartEditRangeSelectorToggle.checked;
    const nextRangeSliderVisible = this.chartEditRangeSliderToggle.checked;
    const nextAnnotationText = this.sanitizeChartText(this.chartEditAnnotationTextInput.value, 240);
    const readPaperCoordinate = (input: HTMLInputElement, fallback: number): number =>
      Math.min(1, Math.max(0, toFiniteNumber(input.valueAsNumber, fallback)));
    const nextAnnotationX = readPaperCoordinate(this.chartEditAnnotationXInput, 0.5);
    const nextAnnotationY = readPaperCoordinate(this.chartEditAnnotationYInput, 0.5);
    const nextShapeType = this.chartEditShapeTypeSelect.value;
    const nextShapeX0 = readPaperCoordinate(this.chartEditShapeX0Input, 0.2);
    const nextShapeY0 = readPaperCoordinate(this.chartEditShapeY0Input, 0.2);
    const nextShapeX1 = readPaperCoordinate(this.chartEditShapeX1Input, 0.8);
    const nextShapeY1 = readPaperCoordinate(this.chartEditShapeY1Input, 0.8);
    const nextRangeInput = this.chartEditRangeInput.value.trim().toUpperCase().replace(/\s+/g, "");
    const nextOrientation = this.chartEditOrientationSelect.value === "columns" ? "columns" : "rows";
    const nextFirstRowAsHeader = this.chartEditFirstRowHeaderToggle.checked;
    const nextFirstColumnAsLabel = this.chartEditFirstColumnLabelToggle.checked;
    const nextAutoRefresh = this.chartEditAutoRefreshToggle.checked;
    const nextCategoryColumnInput = this.chartEditCategoryColumnInput.value;
    const nextSeriesColumnsInput = this.chartEditSeriesColumnsInput.value;
    const nextValueColumnInput = this.chartEditValueColumnInput.value;
    let refreshFromSourceRange = false;

    try {
      const nextCategoryColumnIndex = this.parseChartColumnIndexInput(
        nextCategoryColumnInput,
        this.messages.chartEditCategoryColumnLabel
      );
      const nextSeriesColumnIndexes = this.parseChartColumnIndexListInput(
        nextSeriesColumnsInput,
        this.messages.chartEditSeriesColumnsLabel
      );
      const nextValueColumnIndex = this.parseChartColumnIndexInput(
        nextValueColumnInput,
        this.messages.chartEditValueColumnLabel
      );
      if (nextType && nextType !== chart.type) {
        if (!this.isChartTypeEnabled(nextType)) {
          this.engine.reportChartUnsupportedFeature(sheet.id, selectedChartId, `chart-type:${nextType}`);
          this.setChartFeedback(sheet.id, `${this.messages.chartUnsupportedType}: ${nextType}`, true);
          this.render();
          return;
        }
        this.engine.changeChartType({
          sheetId: sheet.id,
          chartId: selectedChartId,
          chartType: nextType
        });
        chart = this.engine.getChart(sheet.id, selectedChartId) ?? chart;
        refreshFromSourceRange = true;
      }

      const normalizedTitle = nextTitle ? nextTitle : undefined;
      const currentTitle = chart.title?.trim() ? chart.title.trim() : undefined;
      if (normalizedTitle !== currentTitle) {
        this.engine.changeChartTitle({
          sheetId: sheet.id,
          chartId: selectedChartId,
          title: normalizedTitle
        });
        chart = this.engine.getChart(sheet.id, selectedChartId) ?? chart;
      }

      const currentLegendVisible = this.getChartLegendVisible(chart);
      if (nextLegendVisible !== currentLegendVisible) {
        this.engine.changeChartLegend({
          sheetId: sheet.id,
          chartId: selectedChartId,
          visible: nextLegendVisible
        });
        chart = this.engine.getChart(sheet.id, selectedChartId) ?? chart;
      }

      const currentXAxisTitle = this.getChartAxisTitle(chart, "x");
      const currentYAxisTitle = this.getChartAxisTitle(chart, "y");
      const currentXAxisType = this.getChartAxisType(chart, "x");
      const currentYAxisType = this.getChartAxisType(chart, "y");
      const currentXAxisVisible = this.getChartAxisVisible(chart, "x");
      const currentYAxisVisible = this.getChartAxisVisible(chart, "y");
      if (
        nextXAxisTitle !== currentXAxisTitle ||
        nextYAxisTitle !== currentYAxisTitle ||
        nextXAxisType !== currentXAxisType ||
        nextYAxisType !== currentYAxisType ||
        nextXAxisVisible !== currentXAxisVisible ||
        nextYAxisVisible !== currentYAxisVisible
      ) {
        this.engine.updateChart({
          sheetId: sheet.id,
          chartId: selectedChartId,
          patch: {
            figure: this.withChartAxisOptions(chart.figure, {
              xAxisTitle: nextXAxisTitle || undefined,
              yAxisTitle: nextYAxisTitle || undefined,
              xAxisType: nextXAxisType,
              yAxisType: nextYAxisType,
              xAxisVisible: nextXAxisVisible,
              yAxisVisible: nextYAxisVisible
            })
          }
        });
        chart = this.engine.getChart(sheet.id, selectedChartId) ?? chart;
      }

      const currentBinding = this.getChartBindingOptions(chart);
      const currentLayout = chart.figure.layout && typeof chart.figure.layout === "object"
        ? chart.figure.layout as Record<string, unknown>
        : {};
      const currentXAxis = currentLayout.xAxis && typeof currentLayout.xAxis === "object"
        ? currentLayout.xAxis as Record<string, unknown>
        : {};
      this.engine.updateChart({
        sheetId: sheet.id,
        chartId: selectedChartId,
        patch: {
          figure: {
            ...chart.figure,
            layout: {
              ...currentLayout,
              xAxis: {
                ...currentXAxis,
                rangeSelector: { visible: nextRangeSelectorVisible },
                rangeSlider: { visible: nextRangeSliderVisible, start: 0, end: 1 }
              },
              annotations: nextAnnotationText ? [{
                text: nextAnnotationText,
                x: nextAnnotationX,
                y: nextAnnotationY,
                xRef: "paper",
                yRef: "paper",
                showArrow: this.chartEditAnnotationArrowToggle.checked,
                arrowToX: nextAnnotationX,
                arrowToY: Math.max(0, nextAnnotationY - 0.12)
              }] : [],
              shapes: nextShapeType ? [{
                type: nextShapeType,
                x0: nextShapeX0,
                y0: nextShapeY0,
                x1: nextShapeX1,
                y1: nextShapeY1,
                xRef: "paper",
                yRef: "paper",
                stroke: "#2563eb",
                strokeWidth: 2,
                fill: nextShapeType === "line" ? "transparent" : "#dbeafe",
                opacity: 0.65
              }] : []
            }
          }
        }
      });
      chart = this.engine.getChart(sheet.id, selectedChartId) ?? chart;
      const rangeInputForBinding = nextRangeInput || currentBinding?.rangeAddress;
      if (rangeInputForBinding) {
        const parsedRange = this.parseChartRangeAddress(rangeInputForBinding);
        if (!parsedRange) {
          this.setChartFeedback(sheet.id, this.messages.chartEditInvalidRange, true);
          this.render();
          return;
        }
        const rangeCellCount = getRangeCellCount(parsedRange);
        if (rangeCellCount > this.getChartLimits().maxRangeCells) {
          this.engine.reportSecurityEvent("chart-range-too-large", {
            sheetId: sheet.id,
            chartId: selectedChartId,
            rangeCellCount
          });
          this.setChartFeedback(sheet.id, this.messages.chartRangeTooLarge, true);
          this.render();
          return;
        }
        const normalizedRangeAddress = this.buildChartRangeAddress(parsedRange);
        const nextBinding: ChartBindingOptions = {
          rangeAddress: normalizedRangeAddress,
          orientation: nextOrientation,
          firstRowAsHeader: nextFirstRowAsHeader,
          firstColumnAsLabel: nextFirstColumnAsLabel,
          autoRefresh: nextAutoRefresh,
          categoryColumnIndex: nextCategoryColumnIndex,
          seriesColumnIndexes: nextSeriesColumnIndexes,
          valueColumnIndex: nextValueColumnIndex
        };
        if (!this.areChartBindingOptionsEqual(currentBinding, nextBinding)) {
          this.engine.changeChartRange({
            sheetId: sheet.id,
            chartId: selectedChartId,
            sourceRange: {
              ...nextBinding
            }
          });
          chart = this.engine.getChart(sheet.id, selectedChartId) ?? chart;
          refreshFromSourceRange = true;
        }
      }

      if (refreshFromSourceRange && chart.sourceRange) {
        this.refreshBoundChartFigure(sheet.id, chart.id, chart.sourceRange);
      }

      this.setChartFeedback(sheet.id, this.messages.chartEditSaved, false);
      this.render();
      this.focus();
    } catch (error) {
      this.engine.reportChartError({
        sheetId: sheet.id,
        chartId: selectedChartId,
        errorCode: "RENDERER_CHART_EDIT_FAILED",
        message: error instanceof Error ? error.message : "Failed to edit chart."
      });
      this.setChartFeedback(sheet.id, error instanceof Error ? error.message : this.messages.chartInsertError, true);
      this.render();
    }
  }

  private transposeSpreadsheetRows(rows: CellPrimitive[][]): CellPrimitive[][] {
    if (!rows.length) {
      return [];
    }
    const width = Math.max(...rows.map((row) => row.length));
    const transposed: CellPrimitive[][] = [];
    for (let col = 0; col < width; col += 1) {
      const nextRow: CellPrimitive[] = [];
      for (let row = 0; row < rows.length; row += 1) {
        nextRow.push((rows[row] ?? [])[col] ?? null);
      }
      transposed.push(nextRow);
    }
    return transposed;
  }

  private createSpreadsheetRangeInput(
    sheetId: string,
    sourceRange: CellRange,
    binding: ChartBindingOptions
  ): SpreadsheetRangeInput {
    const chartLimits = this.getChartLimits();
    const rangeCellCount = getRangeCellCount(sourceRange);
    if (rangeCellCount > chartLimits.maxRangeCells) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_TOO_LARGE",
        message: this.messages.chartRangeTooLarge,
        area: "renderer",
        recoverable: true,
        details: {
          rangeCellCount,
          maxRangeCells: chartLimits.maxRangeCells
        }
      });
    }
    const rawRows: CellPrimitive[][] = [];
    for (let row = sourceRange.start.row; row <= sourceRange.end.row; row += 1) {
      const nextRow: CellPrimitive[] = [];
      for (let col = sourceRange.start.col; col <= sourceRange.end.col; col += 1) {
        nextRow.push(this.getCellPrimitiveValue(sheetId, row, col));
      }
      rawRows.push(nextRow);
    }

    const orientedRows = binding.orientation === "columns" ? this.transposeSpreadsheetRows(rawRows) : rawRows;
    if (!orientedRows.length || !orientedRows[0]?.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_EMPTY",
        message: this.messages.chartInvalidRange,
        area: "renderer",
        recoverable: true
      });
    }

    const width = orientedRows[0].length;
    const defaultHeaders = Array.from({ length: width }, (_value, index) => `Col ${index + 1}`);
    const headers = binding.firstRowAsHeader
      ? defaultHeaders.map((fallback, index) => {
          const value = orientedRows[0]?.[index];
          const label = value == null ? "" : String(value).trim();
          const safeLabel = this.sanitizeChartText(label, 80);
          return safeLabel || fallback;
        })
      : defaultHeaders;
    const dataRows = (binding.firstRowAsHeader ? orientedRows.slice(1) : orientedRows).map((row) =>
      headers.map((_header, index) => {
        const value = (row[index] ?? null) as SpreadsheetRangeInput["rows"][number][number];
        if (typeof value === "string") {
          return this.sanitizeChartText(value, 160) as SpreadsheetRangeInput["rows"][number][number];
        }
        return value;
      })
    );
    if (!dataRows.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_WITHOUT_DATA",
        message: this.messages.chartInvalidRange,
        area: "renderer",
        recoverable: true
      });
    }
    const configuredSeriesCount = Array.isArray(binding.seriesColumnIndexes)
      ? binding.seriesColumnIndexes.filter(
          (columnIndex) => Number.isInteger(columnIndex) && columnIndex >= 0 && columnIndex < headers.length
        ).length
      : undefined;
    const estimatedSeriesCount = Math.max(
      1,
      configuredSeriesCount ??
        (binding.orientation === "columns" ? dataRows.length : headers.length - (binding.firstColumnAsLabel ? 1 : 0))
    );
    const estimatedPointCount = dataRows.length * headers.length;
    if (estimatedSeriesCount > chartLimits.maxSeriesPerChart) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_SERIES_LIMIT_EXCEEDED",
        message: this.messages.chartTooManySeries,
        area: "renderer",
        recoverable: true
      });
    }
    if (estimatedPointCount > chartLimits.maxPointsPerChart) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_POINTS_LIMIT_EXCEEDED",
        message: this.messages.chartTooManyPoints,
        area: "renderer",
        recoverable: true
      });
    }

    const workbookSheet = this.engine.getSnapshot().sheets.find((sheet) => sheet.id === sheetId);
    return {
      sheetName: workbookSheet?.name,
      headers,
      rows: dataRows
    };
  }

  private toWorksheetChartFigure(input: ReturnType<typeof createFigureFromSpreadsheetRange>): WorksheetChartObject["figure"] {
    const metadata = input.metadata;
    return {
      data: (Array.isArray(input.data) ? input.data : []).map((trace) => ({ ...(trace as Record<string, unknown>) })),
      layout: input.layout ? ({ ...(input.layout as Record<string, unknown>) } as Record<string, unknown>) : undefined,
      config: input.config ? ({ ...(input.config as Record<string, unknown>) } as Record<string, unknown>) : undefined,
      frames: input.frames ? [...input.frames] : undefined,
      selection: input.selection,
      metadata:
        metadata && typeof metadata === "object" && !Array.isArray(metadata)
          ? ({ ...(metadata as Record<string, unknown>) } as Record<string, unknown>)
          : undefined,
      schemaVersion: input.schemaVersion
    };
  }

  private resolveChartCategoryColumnIndex(rangeInput: SpreadsheetRangeInput, binding: ChartBindingOptions): number {
    const resolved =
      Number.isInteger(binding.categoryColumnIndex) && (binding.categoryColumnIndex as number) >= 0
        ? (binding.categoryColumnIndex as number)
        : 0;
    if (resolved < 0 || resolved >= rangeInput.headers.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_BINDING_INVALID",
        message: this.messages.chartEditInvalidBinding,
        area: "renderer",
        recoverable: true
      });
    }
    return resolved;
  }

  private resolveChartSeriesColumnIndexes(
    rangeInput: SpreadsheetRangeInput,
    binding: ChartBindingOptions,
    categoryColumnIndex: number
  ): number[] | undefined {
    const configured = binding.seriesColumnIndexes;
    if (!configured?.length) {
      return undefined;
    }
    const normalized: number[] = [];
    for (const index of configured) {
      if (!Number.isInteger(index) || index < 0 || index >= rangeInput.headers.length) {
        throw new SpreadsheetOperationError({
          code: "RENDERER_CHART_BINDING_INVALID",
          message: this.messages.chartEditInvalidBinding,
          area: "renderer",
          recoverable: true
        });
      }
      if (index === categoryColumnIndex) {
        continue;
      }
      if (!normalized.includes(index)) {
        normalized.push(index);
      }
    }
    if (!normalized.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_BINDING_INVALID",
        message: this.messages.chartEditInvalidBinding,
        area: "renderer",
        recoverable: true
      });
    }
    return normalized;
  }

  private buildPieLikeFigure(
    rangeInput: SpreadsheetRangeInput,
    chartType: "pie" | "donut",
    title: string,
    binding: ChartBindingOptions
  ): WorksheetChartObject["figure"] {
    const width = rangeInput.headers.length;
    const labelColumnIndex =
      Number.isInteger(binding.categoryColumnIndex) && (binding.categoryColumnIndex as number) >= 0
        ? (binding.categoryColumnIndex as number)
        : binding.firstColumnAsLabel
          ? 0
          : -1;
    if (labelColumnIndex >= width) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_BINDING_INVALID",
        message: this.messages.chartEditInvalidBinding,
        area: "renderer",
        recoverable: true
      });
    }
    let valueColumnIndex =
      Number.isInteger(binding.valueColumnIndex) && (binding.valueColumnIndex as number) >= 0
        ? (binding.valueColumnIndex as number)
        : -1;
    if (valueColumnIndex >= width) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_BINDING_INVALID",
        message: this.messages.chartEditInvalidBinding,
        area: "renderer",
        recoverable: true
      });
    }
    if (valueColumnIndex >= 0 && valueColumnIndex === labelColumnIndex) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_BINDING_INVALID",
        message: this.messages.chartEditInvalidBinding,
        area: "renderer",
        recoverable: true
      });
    }
    if (valueColumnIndex < 0) {
      for (let index = 0; index < width; index += 1) {
        if (index === labelColumnIndex) {
          continue;
        }
        const hasNumeric = rangeInput.rows.some((row) => toNumericValue((row[index] ?? null) as CellPrimitive) !== undefined);
        if (hasNumeric) {
          valueColumnIndex = index;
          break;
        }
      }
    }

    if (valueColumnIndex < 0) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_WITHOUT_NUMERIC_SERIES",
        message: this.messages.chartInsertError,
        area: "renderer",
        recoverable: true
      });
    }

    const labels: string[] = [];
    const values: number[] = [];
    for (let index = 0; index < rangeInput.rows.length; index += 1) {
      const row = rangeInput.rows[index] ?? [];
      const numeric = toNumericValue((row[valueColumnIndex] ?? null) as CellPrimitive);
      if (numeric === undefined) {
        continue;
      }
      const labelSource = labelColumnIndex >= 0 ? row[labelColumnIndex] : `Item ${index + 1}`;
      const label = this.sanitizeChartText(String(labelSource ?? `Item ${index + 1}`).trim(), 100) || `Item ${index + 1}`;
      labels.push(label);
      values.push(numeric);
    }

    if (!values.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_WITHOUT_NUMERIC_VALUES",
        message: this.messages.chartInsertError,
        area: "renderer",
        recoverable: true
      });
    }

    return {
      data: [
        {
          type: chartType,
          labels,
          values,
          name: rangeInput.headers[valueColumnIndex] ?? "Valor"
        }
      ],
      layout: {
        title,
        legend: {
          visible: true
        }
      },
      metadata: {
        source: "spreadsheet-range",
        chartType,
        rows: rangeInput.rows.length,
        columns: rangeInput.headers.length
      }
    };
  }

  private getNumericChartSeriesIndexes(
    rangeInput: SpreadsheetRangeInput,
    binding: ChartBindingOptions,
    categoryColumnIndex: number
  ): number[] {
    const configured = this.resolveChartSeriesColumnIndexes(rangeInput, binding, categoryColumnIndex);
    const candidates = configured ?? rangeInput.headers.map((_header, index) => index).filter((index) => index !== categoryColumnIndex);
    const numeric = candidates.filter((index) =>
      rangeInput.rows.some((row) => toNumericValue((row[index] ?? null) as CellPrimitive) !== undefined)
    );
    if (!numeric.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_WITHOUT_NUMERIC_SERIES",
        message: this.messages.chartInsertError,
        area: "renderer",
        recoverable: true
      });
    }
    return numeric;
  }

  private buildStatisticalFigure(
    rangeInput: SpreadsheetRangeInput,
    chartType: "histogram" | "box" | "violin",
    title: string,
    binding: ChartBindingOptions
  ): WorksheetChartObject["figure"] {
    const categoryColumnIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
    const seriesIndexes = this.getNumericChartSeriesIndexes(rangeInput, binding, categoryColumnIndex);
    const data = seriesIndexes.map((seriesIndex) => ({
      type: chartType,
      name: rangeInput.headers[seriesIndex] ?? `Série ${seriesIndex + 1}`,
      values: rangeInput.rows
        .map((row) => toNumericValue((row[seriesIndex] ?? null) as CellPrimitive))
        .filter((value): value is number => value !== undefined)
    }));
    return {
      data,
      layout: { title, legend: { visible: data.length > 1 } },
      metadata: { source: "spreadsheet-range", chartType, rows: rangeInput.rows.length, columns: rangeInput.headers.length }
    };
  }

  private buildHeatmapFigure(
    rangeInput: SpreadsheetRangeInput,
    title: string,
    binding: ChartBindingOptions,
    chartType: "heatmap" | "contour" = "heatmap"
  ): WorksheetChartObject["figure"] {
    const categoryColumnIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
    const seriesIndexes = this.getNumericChartSeriesIndexes(rangeInput, binding, categoryColumnIndex);
    const z: number[][] = [];
    const y: string[] = [];
    for (let rowIndex = 0; rowIndex < rangeInput.rows.length; rowIndex += 1) {
      const row = rangeInput.rows[rowIndex] ?? [];
      const values = seriesIndexes.map((seriesIndex) => toNumericValue((row[seriesIndex] ?? null) as CellPrimitive));
      if (values.some((value) => value === undefined)) {
        continue;
      }
      z.push(values as number[]);
      y.push(this.sanitizeChartText(String(row[categoryColumnIndex] ?? `Item ${rowIndex + 1}`), 100));
    }
    if (!z.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_WITHOUT_NUMERIC_MATRIX",
        message: this.messages.chartInsertError,
        area: "renderer",
        recoverable: true
      });
    }
    return {
      data: [{ type: chartType, z, x: seriesIndexes.map((index) => rangeInput.headers[index] ?? `Série ${index + 1}`), y }],
      layout: { title, legend: { visible: false } },
      metadata: { source: "spreadsheet-range", chartType, rows: z.length, columns: seriesIndexes.length }
    };
  }

  private buildTernaryFigure(
    rangeInput: SpreadsheetRangeInput,
    title: string,
    binding: ChartBindingOptions
  ): WorksheetChartObject["figure"] {
    const categoryColumnIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
    const coordinateIndexes = [categoryColumnIndex, ...this.getNumericChartSeriesIndexes(rangeInput, binding, categoryColumnIndex)]
      .filter((index, position, indexes) => indexes.indexOf(index) === position)
      .slice(0, 3);
    const numericIndexes = coordinateIndexes.filter((index) =>
      rangeInput.rows.some((row) => toNumericValue((row[index] ?? null) as CellPrimitive) !== undefined)
    );
    if (numericIndexes.length < 3) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_TERNARY_REQUIRES_ABC",
        message: "Gráfico ternário requer três colunas numéricas: A, B e C.",
        area: "renderer",
        recoverable: true
      });
    }
    const rows = rangeInput.rows
      .map((row) => numericIndexes.map((index) => toNumericValue((row[index] ?? null) as CellPrimitive)))
      .filter((row) => row.every((value) => value !== undefined)) as number[][];
    return {
      data: [{ type: "ternary", a: rows.map((row) => row[0]), b: rows.map((row) => row[1]), c: rows.map((row) => row[2]) }],
      layout: { title, legend: { visible: false } },
      metadata: { source: "spreadsheet-range", chartType: "ternary", rows: rows.length, columns: 3 }
    };
  }

  private buildCategoryValueFigure(
    rangeInput: SpreadsheetRangeInput,
    chartType: "waterfall" | "funnel" | "polar",
    title: string,
    binding: ChartBindingOptions
  ): WorksheetChartObject["figure"] {
    const categoryColumnIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
    const valueColumnIndex = this.getNumericChartSeriesIndexes(rangeInput, binding, categoryColumnIndex)[0] as number;
    const labels: string[] = [];
    const values: number[] = [];
    for (let rowIndex = 0; rowIndex < rangeInput.rows.length; rowIndex += 1) {
      const row = rangeInput.rows[rowIndex] ?? [];
      const value = toNumericValue((row[valueColumnIndex] ?? null) as CellPrimitive);
      if (value === undefined) {
        continue;
      }
      labels.push(this.sanitizeChartText(String(row[categoryColumnIndex] ?? `Item ${rowIndex + 1}`), 100));
      values.push(value);
    }
    const trace =
      chartType === "waterfall"
        ? { type: chartType, x: labels, y: values, name: rangeInput.headers[valueColumnIndex] }
        : chartType === "funnel"
          ? { type: chartType, labels, values, name: rangeInput.headers[valueColumnIndex] }
          : { type: chartType, theta: labels, r: values, variant: "bar", name: rangeInput.headers[valueColumnIndex] };
    return {
      data: [trace],
      layout: { title, legend: { visible: false } },
      metadata: { source: "spreadsheet-range", chartType, rows: values.length, columns: 2 }
    };
  }

  private buildCandlestickFigure(
    rangeInput: SpreadsheetRangeInput,
    title: string,
    binding: ChartBindingOptions
  ): WorksheetChartObject["figure"] {
    const categoryColumnIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
    const seriesIndexes = this.getNumericChartSeriesIndexes(rangeInput, binding, categoryColumnIndex);
    if (seriesIndexes.length < 4) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_CANDLESTICK_REQUIRES_OHLC",
        message: "Candlestick requer uma coluna de categoria e quatro colunas numéricas: abertura, máxima, mínima e fechamento.",
        area: "renderer",
        recoverable: true
      });
    }
    const [openIndex, highIndex, lowIndex, closeIndex] = seriesIndexes;
    const x: string[] = [];
    const open: number[] = [];
    const high: number[] = [];
    const low: number[] = [];
    const close: number[] = [];
    for (let rowIndex = 0; rowIndex < rangeInput.rows.length; rowIndex += 1) {
      const row = rangeInput.rows[rowIndex] ?? [];
      const next = [openIndex, highIndex, lowIndex, closeIndex].map((index) =>
        toNumericValue((row[index as number] ?? null) as CellPrimitive)
      );
      if (next.some((value) => value === undefined)) {
        continue;
      }
      x.push(this.sanitizeChartText(String(row[categoryColumnIndex] ?? `Item ${rowIndex + 1}`), 100));
      open.push(next[0] as number);
      high.push(next[1] as number);
      low.push(next[2] as number);
      close.push(next[3] as number);
    }
    if (!x.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_WITHOUT_OHLC_VALUES",
        message: this.messages.chartInsertError,
        area: "renderer",
        recoverable: true
      });
    }
    return {
      data: [{ type: "candlestick", x, open, high, low, close }],
      layout: { title, legend: { visible: false } },
      metadata: { source: "spreadsheet-range", chartType: "candlestick", rows: x.length, columns: 5 }
    };
  }

  private buildHierarchyFigure(
    rangeInput: SpreadsheetRangeInput,
    chartType: "treemap" | "sunburst",
    title: string,
    binding: ChartBindingOptions
  ): WorksheetChartObject["figure"] {
    const labelColumnIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
    const valueColumnIndex = this.getNumericChartSeriesIndexes(rangeInput, binding, labelColumnIndex)[0] as number;
    const parentColumnIndex = rangeInput.headers.findIndex(
      (_header, index) => index !== labelColumnIndex && index !== valueColumnIndex
    );
    if (parentColumnIndex < 0) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_HIERARCHY_REQUIRES_PARENT",
        message: "Treemap e sunburst requerem colunas de item, pai e valor.",
        area: "renderer",
        recoverable: true
      });
    }
    const ids: string[] = [];
    const labels: string[] = [];
    const parents: string[] = [];
    const values: number[] = [];
    rangeInput.rows.forEach((row, rowIndex) => {
      const value = toNumericValue((row[valueColumnIndex] ?? null) as CellPrimitive);
      const label = this.sanitizeChartText(String(row[labelColumnIndex] ?? ""), 100);
      if (value === undefined || !label) {
        return;
      }
      ids.push(label);
      labels.push(label);
      parents.push(this.sanitizeChartText(String(row[parentColumnIndex] ?? ""), 100));
      values.push(value);
    });
    if (!ids.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_WITHOUT_HIERARCHY_VALUES",
        message: this.messages.chartInsertError,
        area: "renderer",
        recoverable: true
      });
    }
    return {
      data: [{ type: chartType, ids, labels, parents, values }],
      layout: { title, legend: { visible: false } },
      metadata: { source: "spreadsheet-range", chartType, rows: ids.length, columns: 3 }
    };
  }

  private buildSankeyFigure(
    rangeInput: SpreadsheetRangeInput,
    title: string,
    binding: ChartBindingOptions
  ): WorksheetChartObject["figure"] {
    const sourceColumnIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
    const valueColumnIndex = this.getNumericChartSeriesIndexes(rangeInput, binding, sourceColumnIndex)[0] as number;
    const targetColumnIndex = rangeInput.headers.findIndex(
      (_header, index) => index !== sourceColumnIndex && index !== valueColumnIndex
    );
    if (targetColumnIndex < 0) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_SANKEY_REQUIRES_TARGET",
        message: "Sankey requer colunas de origem, destino e valor.",
        area: "renderer",
        recoverable: true
      });
    }
    const nodeIds: string[] = [];
    const nodeIndexes = new Map<string, number>();
    const source: number[] = [];
    const target: number[] = [];
    const value: number[] = [];
    const resolveNode = (raw: CellPrimitive): number => {
      const id = this.sanitizeChartText(String(raw ?? ""), 100);
      const existing = nodeIndexes.get(id);
      if (existing !== undefined) {
        return existing;
      }
      const index = nodeIds.length;
      nodeIds.push(id);
      nodeIndexes.set(id, index);
      return index;
    };
    rangeInput.rows.forEach((row) => {
      const amount = toNumericValue((row[valueColumnIndex] ?? null) as CellPrimitive);
      const sourceValue = (row[sourceColumnIndex] ?? null) as CellPrimitive;
      const targetValue = (row[targetColumnIndex] ?? null) as CellPrimitive;
      if (amount === undefined || sourceValue === null || targetValue === null) {
        return;
      }
      source.push(resolveNode(sourceValue));
      target.push(resolveNode(targetValue));
      value.push(amount);
    });
    if (!value.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_WITHOUT_SANKEY_LINKS",
        message: this.messages.chartInsertError,
        area: "renderer",
        recoverable: true
      });
    }
    return {
      data: [{ type: "sankey", nodes: { ids: nodeIds, labels: nodeIds }, links: { source, target, value } }],
      layout: { title, legend: { visible: false } },
      metadata: { source: "spreadsheet-range", chartType: "sankey", rows: value.length, columns: 3 }
    };
  }

  private buildThreeDimensionalFigure(
    rangeInput: SpreadsheetRangeInput,
    chartType: "surface" | "surface3d" | "scatter3d",
    title: string,
    binding: ChartBindingOptions
  ): WorksheetChartObject["figure"] {
    const categoryColumnIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
    const seriesIndexes = this.getNumericChartSeriesIndexes(rangeInput, binding, categoryColumnIndex);
    if (chartType === "scatter3d") {
      const coordinateIndexes = [categoryColumnIndex, ...seriesIndexes.filter((index) => index !== categoryColumnIndex)].slice(0, 3);
      if (coordinateIndexes.length < 3) {
        throw new SpreadsheetOperationError({
          code: "RENDERER_CHART_SCATTER3D_REQUIRES_XYZ",
          message: "Scatter 3D requer três colunas numéricas: X, Y e Z.",
          area: "renderer",
          recoverable: true
        });
      }
      const coordinates = rangeInput.rows
        .map((row) => coordinateIndexes.map((index) => toNumericValue((row[index] ?? null) as CellPrimitive)))
        .filter((row) => row.every((value) => value !== undefined)) as number[][];
      return {
        data: [{ type: "scatter3d", x: coordinates.map((row) => row[0]), y: coordinates.map((row) => row[1]), z: coordinates.map((row) => row[2]), mode: "markers" }],
        layout: { title, legend: { visible: false } },
        metadata: { source: "spreadsheet-range", chartType, rows: coordinates.length, columns: 3 }
      };
    }
    const matrixIndexes = [categoryColumnIndex, ...seriesIndexes.filter((index) => index !== categoryColumnIndex)];
    const z = rangeInput.rows
      .map((row) => matrixIndexes.map((index) => toNumericValue((row[index] ?? null) as CellPrimitive)))
      .filter((row) => row.every((value) => value !== undefined)) as number[][];
    if (!z.length) {
      throw new SpreadsheetOperationError({
        code: "RENDERER_CHART_RANGE_WITHOUT_SURFACE_MATRIX",
        message: this.messages.chartInsertError,
        area: "renderer",
        recoverable: true
      });
    }
    return {
      data: [{ type: "surface", z }],
      layout: { title, legend: { visible: false } },
      metadata: { source: "spreadsheet-range", chartType, rows: z.length, columns: matrixIndexes.length }
    };
  }

  private buildChartFigureFromRange(input: {
    sheetId: string;
    chartType: WorksheetChartType;
    sourceRange: CellRange;
    binding: ChartBindingOptions;
    title: string;
    placeholderMode: boolean;
  }): WorksheetChartObject["figure"] {
    const rangeInput = this.createSpreadsheetRangeInput(input.sheetId, input.sourceRange, input.binding);

    if (input.chartType === "pie" || input.chartType === "donut") {
      return this.buildPieLikeFigure(rangeInput, input.chartType, input.title, input.binding);
    }
    if (input.chartType === "histogram" || input.chartType === "box" || input.chartType === "violin") {
      return this.buildStatisticalFigure(rangeInput, input.chartType, input.title, input.binding);
    }
    if (input.chartType === "heatmap" || input.chartType === "contour") {
      return this.buildHeatmapFigure(rangeInput, input.title, input.binding, input.chartType);
    }
    if (input.chartType === "waterfall" || input.chartType === "funnel" || input.chartType === "polar") {
      return this.buildCategoryValueFigure(rangeInput, input.chartType, input.title, input.binding);
    }
    if (input.chartType === "candlestick") {
      return this.buildCandlestickFigure(rangeInput, input.title, input.binding);
    }
    if (input.chartType === "ternary") {
      return this.buildTernaryFigure(rangeInput, input.title, input.binding);
    }
    if (input.chartType === "treemap" || input.chartType === "sunburst") {
      return this.buildHierarchyFigure(rangeInput, input.chartType, input.title, input.binding);
    }
    if (input.chartType === "sankey") {
      return this.buildSankeyFigure(rangeInput, input.title, input.binding);
    }
    if (input.chartType === "surface" || input.chartType === "surface3d" || input.chartType === "scatter3d") {
      return this.buildThreeDimensionalFigure(rangeInput, input.chartType, input.title, input.binding);
    }

    const categoryColumnIndex = this.resolveChartCategoryColumnIndex(rangeInput, input.binding);
    const seriesColumnIndexes = this.resolveChartSeriesColumnIndexes(rangeInput, input.binding, categoryColumnIndex);

    if (input.chartType === "area") {
      const base = createFigureFromSpreadsheetRange(rangeInput, {
        title: input.title,
        traceType: "line",
        xColumn: categoryColumnIndex,
        seriesColumns: seriesColumnIndexes
      });
      const figure = this.toWorksheetChartFigure(base);
      figure.data = figure.data.map((trace, index) => {
        const nextTrace = typeof trace === "object" && trace !== null ? { ...(trace as Record<string, unknown>) } : {};
        nextTrace.type = "area";
        nextTrace.mode = "lines";
        nextTrace.fill = index === 0 ? "tozeroy" : "tonexty";
        return nextTrace;
      });
      return figure;
    }

    const traceType =
      input.chartType === "column" ||
      input.chartType === "bar" ||
      input.placeholderMode
        ? "bar"
        : input.chartType === "scatter"
          ? "scatter"
          : "line";
    const base = createFigureFromSpreadsheetRange(rangeInput, {
      title: input.title,
      traceType,
      xColumn: categoryColumnIndex,
      seriesColumns: seriesColumnIndexes
    });
    const figure = this.toWorksheetChartFigure(base);
    if (input.placeholderMode) {
      const metadata = (figure.metadata ?? {}) as Record<string, unknown>;
      metadata.fallbackChartType = input.chartType;
      metadata.fallbackTraceType = traceType;
      figure.metadata = metadata;
    }
    return figure;
  }

  private estimateFigureSeriesAndPoints(figure: WorksheetChartObject["figure"]): {
    seriesCount: number;
    pointsCount: number;
  } {
    const traces = Array.isArray(figure.data) ? figure.data : [];
    let pointsCount = 0;
    for (const trace of traces) {
      if (!trace || typeof trace !== "object") {
        continue;
      }
      const record = trace as Record<string, unknown>;
      const candidates = [record.y, record.x, record.values, record.z, record.open, record.close, record.high, record.low];
      for (const candidate of candidates) {
        if (Array.isArray(candidate)) {
          pointsCount += candidate.length;
        }
      }
    }
    return {
      seriesCount: traces.length,
      pointsCount
    };
  }

  private isChartInsertPreviewEnabled(): boolean {
    return this.options.chartInsertPreview === true;
  }

  private closeChartInsertPreview(): void {
    this.pendingChartInsertion = undefined;
    this.chartInsertPreviewHost.replaceChildren();
    this.chartInsertPreviewPanel.hidden = true;
  }

  private openChartInsertPreview(input: {
    sheetId: string;
    chartType: WorksheetChartType;
    sourceRange: CellRange;
    binding: ChartBindingOptions;
    figure: WorksheetChartObject["figure"];
    title: string;
    placeholderMode: boolean;
  }): void {
    this.pendingChartInsertion = input;
    const previewChart: WorksheetChartObject = {
      id: "chart-preview",
      sheetId: input.sheetId,
      type: input.chartType,
      title: input.title,
      sourceRange: {
        chartId: "chart-preview",
        sheetId: input.sheetId,
        ...input.binding
      },
      figure: input.figure,
      position: {
        fromCell: "A1",
        toCell: "B2",
        offsetX: 0,
        offsetY: 0,
        width: 420,
        height: 240,
        zIndex: 1
      },
      state: {
        selected: false,
        visible: true,
        locked: false
      },
      excelInterop: {}
    };
    this.renderChartPreview(this.chartInsertPreviewHost, previewChart);
    this.chartInsertPreviewPanel.hidden = false;
  }

  private commitPendingChartInsertion(): void {
    const pending = this.pendingChartInsertion;
    if (!pending) {
      return;
    }
    const sheet = this.engine.getSnapshot().sheets.find((sheetItem) => sheetItem.id === pending.sheetId);
    if (!sheet) {
      this.closeChartInsertPreview();
      return;
    }
    this.engine.createChart({
      sheetId: pending.sheetId,
      chart: {
        type: pending.chartType,
        title: pending.title,
        sourceRange: pending.binding,
        figure: pending.figure,
        position: this.createDefaultChartPosition(sheet, pending.sourceRange),
        state: {
          selected: true,
          visible: true,
          locked: false,
          lastRenderedAt: Date.now()
        },
        excelInterop: pending.placeholderMode
          ? {
              unsupportedFeatures: [`toolbar:${pending.chartType}`],
              fallbackImage: true
            }
          : undefined
      }
    });
    const inserted = this.engine.getCharts(pending.sheetId).at(-1);
    this.setChartSelection(pending.sheetId, inserted?.id);
    if (pending.placeholderMode && inserted) {
      this.engine.reportChartUnsupportedFeature(pending.sheetId, inserted.id, `toolbar:${pending.chartType}`);
    }
    this.setChartFeedback(
      pending.sheetId,
      pending.placeholderMode ? `${this.messages.chartInserted} ${this.messages.chartUnsupportedType}.` : this.messages.chartInserted,
      false
    );
    this.closeChartInsertPreview();
    this.render();
  }

  private createChartFromSelection(action: ChartToolbarAction): void {
    const sheet = this.engine.getActiveSheet();
    if (action === "chart-geo") {
      this.geoJsonFileInput.click();
      return;
    }
    if (!this.isChartActionEnabled(action)) {
      const disabledType = CHART_ACTION_TO_TYPE[action];
      this.engine.reportChartUnsupportedFeature(sheet.id, "toolbar", `toolbar:${disabledType}`);
      this.setChartFeedback(sheet.id, `${this.messages.chartUnsupportedType}: ${disabledType}`, true);
      this.render();
      return;
    }
    const chartLimits = this.getChartLimits();
    const visibleChartsCount = this.engine.getCharts(sheet.id).filter((chart) => chart.state.visible !== false).length;
    if (visibleChartsCount >= chartLimits.maxChartsPerSheet) {
      this.engine.reportSecurityEvent("chart-limit-reached", {
        sheetId: sheet.id,
        limit: chartLimits.maxChartsPerSheet
      });
      this.setChartFeedback(sheet.id, this.messages.chartTooManyObjects, true);
      this.render();
      return;
    }
    const sourceRange = this.getChartSourceRange(sheet);
    if (!sourceRange) {
      this.setChartFeedback(sheet.id, this.messages.chartInvalidRange, true);
      this.render();
      return;
    }
    const rangeCellCount = getRangeCellCount(sourceRange);
    if (rangeCellCount > chartLimits.maxRangeCells) {
      this.engine.reportSecurityEvent("chart-range-too-large", {
        sheetId: sheet.id,
        rangeCellCount,
        maxRangeCells: chartLimits.maxRangeCells
      });
      this.setChartFeedback(sheet.id, this.messages.chartRangeTooLarge, true);
      this.render();
      return;
    }

    const chartType = CHART_ACTION_TO_TYPE[action];
    const binding = this.createChartBindingOptions(sheet.id, sourceRange);
    const placeholderMode = CHART_PLACEHOLDER_ACTIONS.has(action);
    const titleLabel =
      this.messages[(`chart${chartType[0]?.toUpperCase() ?? ""}${chartType.slice(1)}` as keyof RendererMessages)] ?? chartType;
    const title = this.sanitizeChartText(`${titleLabel} (${binding.rangeAddress})`, 180);
    try {
      const figure = this.buildChartFigureFromRange({
        sheetId: sheet.id,
        chartType,
        sourceRange,
        binding,
        title,
        placeholderMode
      });
      const figureStats = this.estimateFigureSeriesAndPoints(figure);
      if (figureStats.seriesCount > chartLimits.maxSeriesPerChart) {
        this.engine.reportSecurityEvent("chart-series-limit", {
          sheetId: sheet.id,
          seriesCount: figureStats.seriesCount,
          maxSeriesPerChart: chartLimits.maxSeriesPerChart
        });
        this.setChartFeedback(sheet.id, this.messages.chartTooManySeries, true);
        this.render();
        return;
      }
      if (figureStats.pointsCount > chartLimits.maxPointsPerChart) {
        this.engine.reportSecurityEvent("chart-points-limit", {
          sheetId: sheet.id,
          pointsCount: figureStats.pointsCount,
          maxPointsPerChart: chartLimits.maxPointsPerChart
        });
        this.setChartFeedback(sheet.id, this.messages.chartTooManyPoints, true);
        this.render();
        return;
      }
      if (this.isChartInsertPreviewEnabled()) {
        this.openChartInsertPreview({
          sheetId: sheet.id,
          chartType,
          sourceRange,
          binding,
          figure,
          title,
          placeholderMode
        });
        this.render();
      } else {
        this.pendingChartInsertion = {
          sheetId: sheet.id,
          chartType,
          sourceRange,
          binding,
          figure,
          title,
          placeholderMode
        };
        this.commitPendingChartInsertion();
      }
      this.focus();
    } catch (error) {
      this.engine.reportChartError({
        sheetId: sheet.id,
        errorCode: "RENDERER_CHART_CREATE_FAILED",
        message: error instanceof Error ? error.message : "Failed to create chart."
      });
      this.setChartFeedback(sheet.id, error instanceof Error ? error.message : this.messages.chartInsertError, true);
      this.render();
    }
  }

  private createDefaultChartPosition(
    sheet: ReturnType<WorkbookEngine["getActiveSheet"]>,
    sourceRange: CellRange
  ): Omit<ChartPosition, "zIndex"> & { zIndex?: number } {
    const targetRow = Math.min(Math.max(0, sheet.rowCount - 1), sourceRange.end.row + 1);
    const targetCol = sourceRange.start.col;
    const toRow = Math.min(Math.max(0, sheet.rowCount - 1), targetRow + 12);
    const toCol = Math.min(Math.max(0, sheet.columnCount - 1), targetCol + 7);
    const currentCharts = this.engine.getCharts(sheet.id);
    const maxZIndex = currentCharts.reduce((max, chart) => Math.max(max, chart.position.zIndex), 1);

    return {
      fromCell: cellAddressToLabel({ row: targetRow, col: targetCol }),
      toCell: cellAddressToLabel({ row: toRow, col: toCol }),
      offsetX: 12,
      offsetY: 8,
      width: 460,
      height: 280,
      zIndex: maxZIndex + 1
    };
  }

  private refreshBoundChartFigure(
    sheetId: string,
    chartId: string,
    range: NonNullable<WorksheetChartObject["sourceRange"]>
  ): void {
    const chart = this.engine.getChart(sheetId, chartId);
    if (!chart?.sourceRange || chart.sourceRange.autoRefresh === false || chart.state.locked) {
      return;
    }

    const sourceRange = this.parseChartRangeAddress(range.rangeAddress);
    if (!sourceRange) {
      this.setChartFeedback(sheetId, this.messages.chartInsertError, true);
      this.engine.reportChartError({
        sheetId,
        chartId,
        errorCode: "RENDERER_CHART_RANGE_INVALID",
        message: `Invalid chart range '${range.rangeAddress}'.`
      });
      return;
    }

    try {
      const placeholderMode = CHART_PLACEHOLDER_ACTIONS.has(`chart-${chart.type}` as ChartToolbarAction);
      const binding = {
          rangeAddress: chart.sourceRange.rangeAddress,
          orientation: chart.sourceRange.orientation,
          firstRowAsHeader: chart.sourceRange.firstRowAsHeader,
          firstColumnAsLabel: chart.sourceRange.firstColumnAsLabel,
          autoRefresh: chart.sourceRange.autoRefresh,
          categoryColumnIndex: chart.sourceRange.categoryColumnIndex,
          seriesColumnIndexes: chart.sourceRange.seriesColumnIndexes ? [...chart.sourceRange.seriesColumnIndexes] : undefined,
          valueColumnIndex: chart.sourceRange.valueColumnIndex
        };
      let nextFigure: WorksheetChartObject["figure"];
      if (chart.type === "geo") {
        const rangeInput = this.createSpreadsheetRangeInput(sheetId, sourceRange, binding);
        const categoryIndex = this.resolveChartCategoryColumnIndex(rangeInput, binding);
        const valueIndex = this.getNumericChartSeriesIndexes(rangeInput, binding, categoryIndex)[0];
        const pairs = rangeInput.rows
          .map((row) => ({ location: String(row[categoryIndex] ?? ""), value: toNumericValue((row[valueIndex] ?? null) as CellPrimitive) }))
          .filter((item): item is { location: string; value: number } => Boolean(item.location) && item.value !== undefined);
        nextFigure = cloneSerializable(chart.figure);
        const trace = nextFigure.data[0];
        if (!trace || typeof trace !== "object" || (trace as { type?: unknown }).type !== "geo") {
          throw new Error("O gráfico GeoJSON não possui um trace geo válido.");
        }
        nextFigure.data[0] = {
          ...(trace as Record<string, unknown>),
          locations: pairs.map((item) => item.location),
          values: pairs.map((item) => item.value)
        };
      } else {
        nextFigure = this.buildChartFigureFromRange({
          sheetId,
          chartType: chart.type,
          sourceRange,
          binding,
          title: chart.title ?? "Chart",
          placeholderMode
        });
      }
      if (placeholderMode) {
        this.engine.reportChartUnsupportedFeature(sheetId, chartId, `chart-type:${chart.type}`);
      }
      const nextLayout =
        typeof nextFigure.layout === "object" && nextFigure.layout !== null
          ? ({ ...(nextFigure.layout as Record<string, unknown>) } as Record<string, unknown>)
          : {};
      const nextLegend =
        typeof nextLayout.legend === "object" && nextLayout.legend !== null
          ? ({ ...(nextLayout.legend as Record<string, unknown>) } as Record<string, unknown>)
          : {};
      nextLegend.visible = this.getChartLegendVisible(chart);
      nextLayout.legend = nextLegend;
      nextFigure.layout = nextLayout;
      nextFigure = this.withChartAxisOptions(nextFigure, {
        xAxisTitle: this.getChartAxisTitle(chart, "x") || undefined,
        yAxisTitle: this.getChartAxisTitle(chart, "y") || undefined,
        xAxisType: this.getChartAxisType(chart, "x"),
        yAxisType: this.getChartAxisType(chart, "y"),
        xAxisVisible: this.getChartAxisVisible(chart, "x"),
        yAxisVisible: this.getChartAxisVisible(chart, "y")
      });
      this.engine.updateChart({
        sheetId,
        chartId,
        patch: {
          figure: nextFigure,
          state: {
            ...chart.state,
            lastRenderedAt: Date.now()
          }
        }
      });
      if (this.engine.getSnapshot().activeSheetId === sheetId) {
        this.setChartFeedback(sheetId, this.messages.chartAutoUpdated, false);
      }
    } catch (error) {
      this.engine.reportChartError({
        sheetId,
        chartId,
        errorCode: "RENDERER_CHART_REFRESH_FAILED",
        message: error instanceof Error ? error.message : "Failed to refresh chart."
      });
      this.setChartFeedback(sheetId, error instanceof Error ? error.message : this.messages.chartInsertError, true);
    }
  }

  private resolveSurfaceIndex(offsets: number[], value: number, maxIndex: number): number {
    const clampedValue = Math.max(0, value);
    let index = 0;
    while (index < maxIndex && offsets[index + 1] <= clampedValue) {
      index += 1;
    }
    return clampNumeric(index, 0, maxIndex);
  }

  private resolveChartRect(
    position: ChartPosition,
    rowOffsets: number[],
    colOffsets: number[],
    rowCount: number,
    colCount: number
  ): ChartRect {
    let fromAddress: CellAddress;
    try {
      fromAddress = cellLabelToAddress(position.fromCell);
    } catch {
      fromAddress = { row: 0, col: 0 };
    }
    const row = clampNumeric(fromAddress.row, 0, Math.max(0, rowCount - 1));
    const col = clampNumeric(fromAddress.col, 0, Math.max(0, colCount - 1));
    return {
      left: ROW_HEADER_WIDTH + (colOffsets[col] ?? 0) + toFiniteNumber(position.offsetX, 0),
      top: (rowOffsets[row] ?? 0) + toFiniteNumber(position.offsetY, 0),
      width: Math.max(CHART_MIN_WIDTH, toFiniteNumber(position.width, CHART_MIN_WIDTH)),
      height: Math.max(CHART_MIN_HEIGHT, toFiniteNumber(position.height, CHART_MIN_HEIGHT))
    };
  }

  private resolveChartPositionFromRect(rect: ChartRect, metrics: ChartSurfaceMetrics, previous: ChartPosition): ChartPosition {
    const maxRowIndex = Math.max(0, metrics.rowCount - 1);
    const maxColIndex = Math.max(0, metrics.colCount - 1);
    const row = this.resolveSurfaceIndex(metrics.rowOffsets, rect.top, maxRowIndex);
    const col = this.resolveSurfaceIndex(metrics.colOffsets, rect.left - ROW_HEADER_WIDTH, maxColIndex);
    const endRow = this.resolveSurfaceIndex(metrics.rowOffsets, rect.top + rect.height, maxRowIndex);
    const endCol = this.resolveSurfaceIndex(metrics.colOffsets, rect.left - ROW_HEADER_WIDTH + rect.width, maxColIndex);
    const fromCellBaseTop = metrics.rowOffsets[row] ?? 0;
    const fromCellBaseLeft = ROW_HEADER_WIDTH + (metrics.colOffsets[col] ?? 0);

    return {
      fromCell: cellAddressToLabel({ row, col }),
      toCell: cellAddressToLabel({ row: endRow, col: endCol }),
      offsetX: Math.round(rect.left - fromCellBaseLeft),
      offsetY: Math.round(rect.top - fromCellBaseTop),
      width: Math.round(rect.width),
      height: Math.round(rect.height),
      zIndex: previous.zIndex
    };
  }

  private getChartViewportBounds(): { left: number; top: number; right: number; bottom: number } {
    const width = this.viewport.clientWidth || this.container.clientWidth || 800;
    const height = this.viewport.clientHeight || this.container.clientHeight || 480;
    return {
      left: ROW_HEADER_WIDTH + this.viewport.scrollLeft,
      top: this.viewport.scrollTop,
      right: ROW_HEADER_WIDTH + this.viewport.scrollLeft + width,
      bottom: this.viewport.scrollTop + height
    };
  }

  private isChartRectWithinViewport(rect: ChartRect, marginPx: number): boolean {
    const viewport = this.getChartViewportBounds();
    return (
      rect.left + rect.width >= viewport.left - marginPx &&
      rect.left <= viewport.right + marginPx &&
      rect.top + rect.height >= viewport.top - marginPx &&
      rect.top <= viewport.bottom + marginPx
    );
  }

  private updateChartRenderStatus(
    sheetId: string,
    chartId: string,
    status: "rendered" | "skipped",
    reason?: string,
    durationMs?: number
  ): void {
    const key = `${sheetId}:${chartId}`;
    const marker = status === "skipped" ? `skipped:${reason ?? "unknown"}` : "rendered";
    if (status === "skipped" && this.chartRenderStatusById.get(key) === marker) {
      return;
    }
    this.chartRenderStatusById.set(key, marker);
    if (status === "skipped") {
      this.engine.reportChartRenderSkipped(sheetId, chartId, reason ?? "skipped");
      return;
    }
    this.engine.reportChartRenderFinished(sheetId, chartId, durationMs);
  }

  private createSvgElement<T extends keyof SVGElementTagNameMap>(tag: T): SVGElementTagNameMap[T] {
    return document.createElementNS(CHART_SVG_NS, tag);
  }

  private coerceNumericSeries(values: unknown): number[] {
    if (!Array.isArray(values)) {
      return [];
    }
    return values
      .map((value) => toNumericValue((value ?? null) as CellPrimitive))
      .filter((value): value is number => typeof value === "number" && Number.isFinite(value));
  }

  private renderLineLikeChartPreview(
    svg: SVGSVGElement,
    values: number[],
    mode: "line" | "area" | "scatter"
  ): void {
    if (!values.length) {
      return;
    }

    const width = 160;
    const height = 96;
    const innerX = 12;
    const innerY = 10;
    const innerW = width - 24;
    const innerH = height - 22;
    const minValue = Math.min(...values);
    const maxValue = Math.max(...values);
    const range = maxValue - minValue || 1;
    const points = values.slice(0, 24).map((value, index, array) => {
      const x = innerX + (array.length <= 1 ? innerW / 2 : (innerW * index) / (array.length - 1));
      const y = innerY + innerH - ((value - minValue) / range) * innerH;
      return { x, y };
    });

    if (mode === "area" && points.length) {
      const areaPath = this.createSvgElement("path");
      const lineSegments = points.map((point, index) => `${index === 0 ? "M" : "L"}${point.x.toFixed(2)} ${point.y.toFixed(2)}`).join(" ");
      areaPath.setAttribute("d", `${lineSegments} L ${points.at(-1)?.x ?? innerX} ${innerY + innerH} L ${points[0]?.x ?? innerX} ${innerY + innerH} Z`);
      areaPath.setAttribute("fill", "rgba(15,118,110,0.24)");
      areaPath.setAttribute("stroke", "none");
      svg.append(areaPath);
    }

    if (mode === "line" || mode === "area") {
      const linePath = this.createSvgElement("path");
      linePath.setAttribute(
        "d",
        points.map((point, index) => `${index === 0 ? "M" : "L"}${point.x.toFixed(2)} ${point.y.toFixed(2)}`).join(" ")
      );
      linePath.setAttribute("fill", "none");
      linePath.setAttribute("stroke", "currentColor");
      linePath.setAttribute("stroke-width", "2");
      linePath.setAttribute("stroke-linecap", "round");
      linePath.setAttribute("stroke-linejoin", "round");
      svg.append(linePath);
    }

    if (mode === "scatter") {
      for (const point of points) {
        const marker = this.createSvgElement("circle");
        marker.setAttribute("cx", point.x.toFixed(2));
        marker.setAttribute("cy", point.y.toFixed(2));
        marker.setAttribute("r", "2.5");
        marker.setAttribute("fill", "currentColor");
        svg.append(marker);
      }
    }
  }

  private renderBarChartPreview(svg: SVGSVGElement, values: number[]): void {
    if (!values.length) {
      return;
    }
    const bars = values.slice(0, 14);
    const width = 160;
    const height = 96;
    const baseLine = 86;
    const minValue = Math.min(...bars, 0);
    const maxValue = Math.max(...bars, 1);
    const range = maxValue - minValue || 1;
    const slot = 134 / bars.length;
    for (let index = 0; index < bars.length; index += 1) {
      const value = bars[index] as number;
      const normalized = (value - minValue) / range;
      const barHeight = Math.max(4, normalized * 72);
      const rect = this.createSvgElement("rect");
      rect.setAttribute("x", String(14 + index * slot));
      rect.setAttribute("y", String(baseLine - barHeight));
      rect.setAttribute("width", String(Math.max(3, slot - 2)));
      rect.setAttribute("height", String(barHeight));
      rect.setAttribute("rx", "1");
      rect.setAttribute("fill", "currentColor");
      rect.setAttribute("opacity", "0.75");
      svg.append(rect);
    }
  }

  private renderPieChartPreview(svg: SVGSVGElement, values: number[], donut: boolean): void {
    if (!values.length) {
      return;
    }
    const total = values.reduce((sum, value) => sum + Math.max(0, value), 0);
    if (total <= 0) {
      return;
    }
    const colors = ["#0f766e", "#2563eb", "#9333ea", "#d97706", "#dc2626", "#16a34a", "#475569"];
    let currentAngle = -Math.PI / 2;
    for (let index = 0; index < values.length; index += 1) {
      const value = Math.max(0, values[index] as number);
      if (value <= 0) {
        continue;
      }
      const angle = (value / total) * Math.PI * 2;
      const nextAngle = currentAngle + angle;
      const x1 = 80 + Math.cos(currentAngle) * 34;
      const y1 = 48 + Math.sin(currentAngle) * 34;
      const x2 = 80 + Math.cos(nextAngle) * 34;
      const y2 = 48 + Math.sin(nextAngle) * 34;
      const arc = this.createSvgElement("path");
      arc.setAttribute(
        "d",
        `M 80 48 L ${x1.toFixed(2)} ${y1.toFixed(2)} A 34 34 0 ${angle > Math.PI ? 1 : 0} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z`
      );
      arc.setAttribute("fill", colors[index % colors.length] as string);
      svg.append(arc);
      currentAngle = nextAngle;
    }
    if (donut) {
      const hole = this.createSvgElement("circle");
      hole.setAttribute("cx", "80");
      hole.setAttribute("cy", "48");
      hole.setAttribute("r", "16");
      hole.setAttribute("fill", "white");
      svg.append(hole);
    }
  }

  private renderChartPreview(host: HTMLElement, chart: WorksheetChartObject): void {
    host.replaceChildren();
    const data = Array.isArray(chart.figure.data) ? chart.figure.data : [];
    const firstTrace = data.find((trace) => typeof trace === "object" && trace !== null) as Record<string, unknown> | undefined;
    const primaryValues = this.coerceNumericSeries(firstTrace?.y);
    const pieValues = this.coerceNumericSeries(firstTrace?.values);
    const svg = this.createSvgElement("svg");
    svg.setAttribute("viewBox", "0 0 160 96");
    svg.classList.add("excelsior-chart-preview-svg");

    const bg = this.createSvgElement("rect");
    bg.setAttribute("x", "0");
    bg.setAttribute("y", "0");
    bg.setAttribute("width", "160");
    bg.setAttribute("height", "96");
    bg.setAttribute("fill", "rgba(248,250,252,0.86)");
    bg.setAttribute("stroke", "rgba(148,163,184,0.25)");
    svg.append(bg);

    switch (chart.type) {
      case "column":
      case "bar":
        this.renderBarChartPreview(svg, primaryValues.length ? primaryValues : pieValues);
        break;
      case "line":
        this.renderLineLikeChartPreview(svg, primaryValues, "line");
        break;
      case "area":
        this.renderLineLikeChartPreview(svg, primaryValues, "area");
        break;
      case "scatter":
        this.renderLineLikeChartPreview(svg, primaryValues, "scatter");
        break;
      case "pie":
        this.renderPieChartPreview(svg, pieValues, false);
        break;
      case "donut":
        this.renderPieChartPreview(svg, pieValues, true);
        break;
      default: {
        const placeholder = document.createElement("div");
        placeholder.className = "excelsior-chart-preview-placeholder";
        placeholder.textContent = `${this.messages.chartUnsupportedType}: ${chart.type}`;
        host.append(placeholder);
        return;
      }
    }

    host.append(svg);
  }

  private toChartFigureInput(figure: WorksheetChartObject["figure"]): ChartFigureInput {
    const normalizedData = Array.isArray(figure.data) ? cloneSerializable(figure.data as unknown[]) : [];
    const normalizedLayout =
      figure.layout && typeof figure.layout === "object" ? cloneSerializable(figure.layout as Record<string, unknown>) : undefined;
    const normalizedConfig =
      figure.config && typeof figure.config === "object" ? cloneSerializable(figure.config as Record<string, unknown>) : undefined;
    const normalizedFrames = Array.isArray(figure.frames) ? cloneSerializable(figure.frames as unknown[]) : undefined;
    const normalizedMetadata =
      figure.metadata && typeof figure.metadata === "object" ? cloneSerializable(figure.metadata as Record<string, unknown>) : undefined;
    const normalizedSelection =
      figure.selection && typeof figure.selection === "object" ? cloneSerializable(figure.selection as Record<string, unknown>) : undefined;
    return {
      data: normalizedData as ChartFigureInput["data"],
      layout: normalizedLayout as ChartFigureInput["layout"],
      config: normalizedConfig as ChartFigureInput["config"],
      frames: normalizedFrames as ChartFigureInput["frames"],
      metadata: normalizedMetadata as ChartFigureInput["metadata"],
      selection: normalizedSelection as ChartFigureInput["selection"],
      schemaVersion: typeof figure.schemaVersion === "string" ? figure.schemaVersion : undefined
    };
  }

  private getFigureLayoutRecord(figure: ChartFigureInput): Record<string, unknown> {
    return typeof figure.layout === "object" && figure.layout !== null
      ? ({ ...(figure.layout as Record<string, unknown>) } as Record<string, unknown>)
      : {};
  }

  private getFigureAxisConfig(
    layout: Record<string, unknown>,
    axis: "x" | "y",
    secondary = false
  ): Record<string, unknown> | undefined {
    const keys =
      axis === "x"
        ? secondary
          ? ["xAxis2", "xaxis2"]
          : ["xAxis", "xaxis"]
        : secondary
          ? ["yAxis2", "yaxis2"]
          : ["yAxis", "yaxis"];
    for (const key of keys) {
      const candidate = layout[key];
      if (candidate && typeof candidate === "object" && !Array.isArray(candidate)) {
        return candidate as Record<string, unknown>;
      }
    }
    return undefined;
  }

  private getFigureAxisType(layout: Record<string, unknown>, axis: "x" | "y", secondary = false): ChartAxisLayoutType {
    const axisConfig = this.getFigureAxisConfig(layout, axis, secondary);
    const rawType = typeof axisConfig?.type === "string" ? axisConfig.type.trim().toLowerCase() : "";
    if (
      rawType === "linear" ||
      rawType === "category" ||
      rawType === "date" ||
      rawType === "log" ||
      rawType === "multicategory"
    ) {
      return rawType;
    }
    return axis === "x" ? "category" : "linear";
  }

  private hasFigureAxisTitle(layout: Record<string, unknown>, axis: "x" | "y", secondary = false): boolean {
    const axisConfig = this.getFigureAxisConfig(layout, axis, secondary);
    if (!axisConfig) {
      return false;
    }
    const title = axisConfig.title;
    if (typeof title === "string") {
      return title.trim().length > 0;
    }
    if (title && typeof title === "object") {
      const text = (title as Record<string, unknown>).text;
      return typeof text === "string" && text.trim().length > 0;
    }
    return false;
  }

  private appendRuntimeSampleValues(target: unknown[], source: unknown, maxSamples = 72): void {
    if (!Array.isArray(source) || source.length === 0 || target.length >= maxSamples) {
      return;
    }
    const remaining = maxSamples - target.length;
    if (source.length <= remaining) {
      for (const value of source) {
        target.push(value);
      }
      return;
    }
    const step = Math.max(1, Math.floor(source.length / remaining));
    for (let index = 0; index < source.length && target.length < maxSamples; index += step) {
      target.push(source[index]);
    }
    if (target.length < maxSamples) {
      target.push(source[source.length - 1]);
    }
  }

  private collectRuntimeAxisSampleValues(figure: ChartFigureInput, axis: "x" | "y", secondary = false): unknown[] {
    const traces = Array.isArray(figure.data) ? figure.data : [];
    const values: unknown[] = [];
    for (const trace of traces) {
      if (!trace || typeof trace !== "object" || Array.isArray(trace)) {
        continue;
      }
      const record = trace as Record<string, unknown>;
      const axisRefKey = axis === "x" ? "xAxisRef" : "yAxisRef";
      const axisRefValue = typeof record[axisRefKey] === "string" ? (record[axisRefKey] as string).toLowerCase() : "";
      const belongsToSecondary = axis === "x" ? axisRefValue === "x2" : axisRefValue === "y2";
      if (secondary ? !belongsToSecondary : belongsToSecondary) {
        continue;
      }
      const candidates =
        axis === "x"
          ? [record.x, record.labels]
          : [record.y, record.values, record.open, record.close, record.high, record.low, record.z];
      for (const candidate of candidates) {
        this.appendRuntimeSampleValues(values, candidate, 72);
      }
      if (values.length >= 72) {
        break;
      }
    }
    return values;
  }

  private formatRuntimeAxisSampleLabel(value: unknown, axisType: ChartAxisLayoutType): string {
    if (value == null) {
      return "";
    }
    if (axisType === "date") {
      const timestamp =
        value instanceof Date
          ? value.getTime()
          : typeof value === "number"
            ? value
            : Date.parse(String(value).trim());
      if (Number.isFinite(timestamp)) {
        const date = new Date(timestamp);
        if (!Number.isNaN(date.getTime())) {
          const year = date.getUTCFullYear();
          const month = String(date.getUTCMonth() + 1).padStart(2, "0");
          const day = String(date.getUTCDate()).padStart(2, "0");
          return `${year}-${month}-${day}`;
        }
      }
    }
    if (typeof value === "number") {
      if (!Number.isFinite(value)) {
        return "";
      }
      const absolute = Math.abs(value);
      if (absolute >= 1000) {
        return value.toFixed(0);
      }
      if (absolute >= 100) {
        return value.toFixed(1);
      }
      return value.toFixed(2);
    }
    if (value instanceof Date) {
      const year = value.getUTCFullYear();
      const month = String(value.getUTCMonth() + 1).padStart(2, "0");
      const day = String(value.getUTCDate()).padStart(2, "0");
      return `${year}-${month}-${day}`;
    }
    if (typeof value === "boolean") {
      return value ? "TRUE" : "FALSE";
    }
    return this.sanitizeChartText(String(value).trim(), 48);
  }

  private estimateRuntimeAxisLabelChars(
    figure: ChartFigureInput,
    axisType: ChartAxisLayoutType,
    axis: "x" | "y",
    secondary = false
  ): number {
    const values = this.collectRuntimeAxisSampleValues(figure, axis, secondary);
    let maxChars = axisType === "date" ? 10 : 4;
    for (const value of values) {
      const label = this.formatRuntimeAxisSampleLabel(value, axisType);
      if (label) {
        maxChars = Math.max(maxChars, label.length);
      }
    }
    return clampNumeric(maxChars, 4, 24);
  }

  private estimateRuntimePointCount(figure: ChartFigureInput): number {
    const traces = Array.isArray(figure.data) ? figure.data : [];
    let maxPoints = 0;
    for (const trace of traces) {
      if (!trace || typeof trace !== "object" || Array.isArray(trace)) {
        continue;
      }
      const record = trace as Record<string, unknown>;
      const candidateLengths = [
        record.x,
        record.y,
        record.values,
        record.labels,
        record.z,
        record.open,
        record.close,
        record.high,
        record.low
      ]
        .filter(Array.isArray)
        .map((candidate) => (candidate as unknown[]).length);
      for (const length of candidateLengths) {
        maxPoints = Math.max(maxPoints, length);
      }
    }
    return maxPoints;
  }

  private hasBarLikeRuntimeTraces(figure: ChartFigureInput): boolean {
    const traces = Array.isArray(figure.data) ? figure.data : [];
    for (const trace of traces) {
      if (!trace || typeof trace !== "object" || Array.isArray(trace)) {
        continue;
      }
      const type = typeof (trace as Record<string, unknown>).type === "string" ? ((trace as Record<string, unknown>).type as string) : "";
      const normalized = type.trim().toLowerCase();
      if (
        normalized === "bar" ||
        normalized === "histogram" ||
        normalized === "waterfall" ||
        normalized === "candlestick" ||
        normalized === "ohlc" ||
        normalized === "funnel" ||
        normalized === "box" ||
        normalized === "violin"
      ) {
        return true;
      }
    }
    return false;
  }

  private estimateRuntimeVirtualSize(figure: ChartFigureInput, width: number, height: number): { width: number; height: number } {
    const safeWidth = Math.max(1, Math.round(toFiniteNumber(width, CHART_MIN_WIDTH)));
    const safeHeight = Math.max(1, Math.round(toFiniteNumber(height, CHART_MIN_HEIGHT)));
    const pointCount = this.estimateRuntimePointCount(figure);
    const axisLayout = this.getFigureLayoutRecord(figure);
    const xAxisType = this.getFigureAxisType(axisLayout, "x", false);
    const hasBarLike = this.hasBarLikeRuntimeTraces(figure);
    const pixelsPerPoint =
      hasBarLike || xAxisType === "category" || xAxisType === "multicategory" || xAxisType === "date" ? 36 : 20;

    let virtualWidth = safeWidth;
    if (pointCount > 12) {
      const estimatedWidth = Math.round(pointCount * pixelsPerPoint + 180);
      const maxVirtualWidth = Math.max(safeWidth, Math.min(9600, Math.max(safeWidth * 6, 3200)));
      virtualWidth = Math.min(maxVirtualWidth, Math.max(safeWidth, estimatedWidth));
    }

    const seriesCount = Array.isArray(figure.data) ? figure.data.length : 0;
    let virtualHeight = safeHeight;
    if (seriesCount > 8) {
      const extraHeight = Math.min(480, (seriesCount - 8) * 14);
      virtualHeight = Math.max(safeHeight, safeHeight + extraHeight);
    }
    if (xAxisType === "multicategory") {
      virtualHeight = Math.max(virtualHeight, safeHeight + 42);
    }

    return {
      width: Math.max(1, Math.round(virtualWidth)),
      height: Math.max(1, Math.round(virtualHeight))
    };
  }

  private resolveRuntimeFigureMargins(
    figure: ChartFigureInput,
    width: number,
    height: number
  ): { top: number; right: number; bottom: number; left: number } {
    const layout = this.getFigureLayoutRecord(figure);
    const rawMargin = layout.margin;
    const margin = rawMargin && typeof rawMargin === "object" ? (rawMargin as Record<string, unknown>) : {};
    const currentTop = Math.max(0, Math.round(toFiniteNumber(margin.top, 56)));
    const currentRight = Math.max(0, Math.round(toFiniteNumber(margin.right, 24)));
    const currentBottom = Math.max(0, Math.round(toFiniteNumber(margin.bottom, 48)));
    const currentLeft = Math.max(0, Math.round(toFiniteNumber(margin.left, 56)));

    const xAxisType = this.getFigureAxisType(layout, "x", false);
    const yAxisType = this.getFigureAxisType(layout, "y", false);
    const yAxis2Type = this.getFigureAxisType(layout, "y", true);
    const yChars = this.estimateRuntimeAxisLabelChars(figure, yAxisType, "y", false);
    const y2Chars = this.estimateRuntimeAxisLabelChars(figure, yAxis2Type, "y", true);
    const xChars = this.estimateRuntimeAxisLabelChars(figure, xAxisType, "x", false);
    const hasYTitle = this.hasFigureAxisTitle(layout, "y", false);
    const hasY2Title = this.hasFigureAxisTitle(layout, "y", true);
    const hasXTitle = this.hasFigureAxisTitle(layout, "x", false);
    const hasY2Values = this.collectRuntimeAxisSampleValues(figure, "y", true).length > 0;

    let left = Math.max(
      currentLeft,
      yAxisType === "date" ? 92 : 68,
      Math.min(148, 18 + yChars * 7 + (hasYTitle ? 18 : 0))
    );
    let right = Math.max(currentRight, hasY2Values ? Math.min(132, 14 + y2Chars * 7 + (hasY2Title ? 16 : 0)) : 24);
    let bottom = Math.max(
      currentBottom,
      xAxisType === "multicategory" ? 90 : xAxisType === "date" ? 70 : 50,
      Math.min(126, 18 + xChars * 6 + (hasXTitle ? 18 : 0))
    );
    let top = Math.max(currentTop, typeof layout.title === "string" && layout.title.trim() ? 64 : 40);

    const minPlotWidth = 120;
    if (width - left - right < minPlotWidth) {
      const overflow = minPlotWidth - (width - left - right);
      const reducibleLeft = Math.max(0, left - 56);
      const reduceLeft = Math.min(reducibleLeft, Math.ceil(overflow * 0.65));
      left -= reduceLeft;
      const remainingOverflow = overflow - reduceLeft;
      right -= Math.min(Math.max(0, right - 20), remainingOverflow);
    }
    const minPlotHeight = 96;
    if (height - top - bottom < minPlotHeight) {
      const overflow = minPlotHeight - (height - top - bottom);
      const reducibleBottom = Math.max(0, bottom - 40);
      const reduceBottom = Math.min(reducibleBottom, Math.ceil(overflow * 0.7));
      bottom -= reduceBottom;
      const remainingOverflow = overflow - reduceBottom;
      top -= Math.min(Math.max(0, top - 24), remainingOverflow);
    }

    return {
      top: Math.max(24, Math.round(top)),
      right: Math.max(18, Math.round(right)),
      bottom: Math.max(40, Math.round(bottom)),
      left: Math.max(56, Math.round(left))
    };
  }

  private withRuntimeFigureSize(figure: ChartFigureInput, width: number, height: number): ChartFigureInput {
    const safeWidth = Math.max(1, Math.round(toFiniteNumber(width, CHART_MIN_WIDTH)));
    const safeHeight = Math.max(1, Math.round(toFiniteNumber(height, CHART_MIN_HEIGHT)));
    const layout = this.getFigureLayoutRecord(figure);
    layout.width = safeWidth;
    layout.height = safeHeight;
    layout.margin = this.resolveRuntimeFigureMargins(figure, safeWidth, safeHeight);
    return {
      ...figure,
      layout: layout as ChartFigureInput["layout"]
    };
  }

  private destroyChartRuntime(chartId: string): void {
    const runtime = this.chartRuntimeById.get(chartId);
    if (!runtime) {
      this.chartRuntimeFigureById.delete(chartId);
      return;
    }
    try {
      runtime.destroy();
    } catch {
      // Best-effort cleanup.
    }
    this.chartRuntimeById.delete(chartId);
    this.chartRuntimeFigureById.delete(chartId);
  }

  private destroyAllChartRuntimes(): void {
    for (const chartId of this.chartRuntimeById.keys()) {
      this.destroyChartRuntime(chartId);
    }
    this.chartObjectElementById.clear();
    this.chartBodyElementById.clear();
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

  private renderChartEngineFigure(body: HTMLElement, chart: WorksheetChartObject): boolean {
    const width = body.clientWidth || chart.position.width || 0;
    const height = body.clientHeight || chart.position.height || 0;
    if (width < 32 || height < 32) {
      return false;
    }
    const baseFigure = this.toChartFigureInput(chart.figure);
    const runtimeSize = this.estimateRuntimeVirtualSize(baseFigure, width, height);
    const figureInput = this.withRuntimeFigureSize(baseFigure, runtimeSize.width, runtimeSize.height);
    if (!Array.isArray(figureInput.data) || figureInput.data.length === 0) {
      return false;
    }
    const runtimeHost = body.querySelector<HTMLElement>("[data-chart-runtime-host='true']") ?? document.createElement("div");
    runtimeHost.dataset.chartRuntimeHost = "true";
    runtimeHost.className = "excelsior-chart-runtime-host";
    runtimeHost.style.width = `${runtimeSize.width}px`;
    runtimeHost.style.height = `${runtimeSize.height}px`;
    runtimeHost.style.minWidth = `${Math.max(1, Math.round(width))}px`;
    runtimeHost.style.minHeight = `${Math.max(1, Math.round(height))}px`;
    body.classList.toggle("is-scrollable", runtimeSize.width > width || runtimeSize.height > height);
    if (!runtimeHost.isConnected) {
      body.replaceChildren(runtimeHost);
    }
    const runtimeSignature = JSON.stringify({
      figure: figureInput,
      width: runtimeSize.width,
      height: runtimeSize.height
    });
    const existingRuntime = this.chartRuntimeById.get(chart.id);
    if (!existingRuntime) {
      const runtime = createFigure(runtimeHost, figureInput, {
        containerClassName: "excelsior-chart-runtime"
      });
      this.chartRuntimeById.set(chart.id, runtime);
      this.chartRuntimeFigureById.set(chart.id, runtimeSignature);
      return true;
    }
    if (this.chartRuntimeFigureById.get(chart.id) !== runtimeSignature) {
      existingRuntime.update(figureInput);
      this.chartRuntimeFigureById.set(chart.id, runtimeSignature);
    } else {
      existingRuntime.resize();
    }
    return true;
  }

  private createChartObjectElement(chart: WorksheetChartObject): HTMLElement {
    const object = document.createElement("section");
    object.className = "excelsior-chart-object";
    object.dataset.chartId = chart.id;

    const header = document.createElement("header");
    header.className = "excelsior-chart-object-header";
    header.setAttribute("aria-label", this.messages.chartMoveHandle);
    header.dataset.chartMove = "true";

    const title = document.createElement("div");
    title.className = "excelsior-chart-object-title";
    title.dataset.chartTitle = "true";
    const subtitle = document.createElement("div");
    subtitle.className = "excelsior-chart-object-subtitle";
    subtitle.dataset.chartSubtitle = "true";
    title.append(subtitle);

    const actions = document.createElement("div");
    actions.className = "excelsior-chart-object-actions";
    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "excelsior-chart-object-delete";
    deleteButton.dataset.chartAction = "delete";
    deleteButton.setAttribute("aria-label", this.messages.chartDelete);
    deleteButton.title = this.messages.chartDelete;
    deleteButton.textContent = "×";
    actions.append(deleteButton);
    header.append(title, actions);

    const body = document.createElement("div");
    body.className = "excelsior-chart-object-body";
    body.dataset.chartBody = "true";

    const resize = document.createElement("button");
    resize.type = "button";
    resize.className = "excelsior-chart-object-resize";
    resize.dataset.chartResize = "true";
    resize.setAttribute("aria-label", this.messages.chartResizeHandle);
    resize.title = this.messages.chartResizeHandle;

    object.append(header, body, resize);
    this.chartObjectElementById.set(chart.id, object);
    this.chartBodyElementById.set(chart.id, body);
    return object;
  }

  private renderChartObjects(
    sheet: ReturnType<WorkbookEngine["getActiveSheet"]>,
    rowOffsets: number[],
    colOffsets: number[]
  ): void {
    const chartLimits = this.getChartLimits();
    const performanceOptions = this.getChartPerformanceOptions();
    const charts = this.engine
      .getCharts(sheet.id)
      .filter((chart) => chart.state.visible !== false)
      .sort((left, right) => left.position.zIndex - right.position.zIndex)
      .slice(0, chartLimits.maxChartsPerSheet);
    if (this.selectedChartId && !charts.some((chart) => chart.id === this.selectedChartId)) {
      this.selectedChartId = undefined;
    }
    if (!this.selectedChartId) {
      this.selectedChartId = charts.find((chart) => chart.state.selected)?.id;
    }
    const visibleChartIds = new Set(charts.map((chart) => chart.id));
    for (const [chartId, element] of this.chartObjectElementById.entries()) {
      if (visibleChartIds.has(chartId)) {
        continue;
      }
      this.destroyChartRuntime(chartId);
      this.chartRenderStatusById.delete(`${sheet.id}:${chartId}`);
      element.remove();
      this.chartObjectElementById.delete(chartId);
      this.chartBodyElementById.delete(chartId);
    }

    const fragment = document.createDocumentFragment();
    for (const chart of charts) {
      const rect = this.resolveChartRect(chart.position, rowOffsets, colOffsets, sheet.rowCount, sheet.columnCount);
      const object = this.chartObjectElementById.get(chart.id) ?? this.createChartObjectElement(chart);
      const body = this.chartBodyElementById.get(chart.id);
      if (!body) {
        continue;
      }

      object.style.left = `${rect.left}px`;
      object.style.top = `${rect.top}px`;
      object.style.width = `${rect.width}px`;
      object.style.height = `${rect.height}px`;
      object.style.zIndex = String(chart.position.zIndex);
      object.classList.toggle("is-selected", chart.id === this.selectedChartId || chart.state.selected);
      object.classList.toggle("is-locked", chart.state.locked);

      const titleElement = object.querySelector<HTMLElement>("[data-chart-title='true']");
      const subtitleElement = object.querySelector<HTMLElement>("[data-chart-subtitle='true']");
      if (titleElement) {
        const titleText = chart.title ?? `Chart (${chart.type})`;
        titleElement.childNodes[0]?.nodeType === Node.TEXT_NODE
          ? (titleElement.childNodes[0].textContent = titleText)
          : titleElement.prepend(document.createTextNode(titleText));
      }
      if (subtitleElement) {
        subtitleElement.textContent = chart.sourceRange?.rangeAddress ?? chart.type;
      }

      const shouldSkipPreview =
        performanceOptions.skipOffscreenPreview && !this.isChartRectWithinViewport(rect, performanceOptions.offscreenMarginPx);
      if (shouldSkipPreview) {
        object.classList.add("is-offscreen");
        this.destroyChartRuntime(chart.id);
        body.classList.remove("is-scrollable");
        body.replaceChildren();
        const placeholder = document.createElement("div");
        placeholder.className = "excelsior-chart-preview-placeholder";
        placeholder.textContent = "Prévia pausada fora da viewport.";
        body.append(placeholder);
        this.updateChartRenderStatus(sheet.id, chart.id, "skipped", "outside-viewport");
      } else {
        object.classList.remove("is-offscreen");
        this.engine.reportChartRenderStarted(sheet.id, chart.id);
        const renderStartedAt = Date.now();
        try {
          const rendered = this.renderChartEngineFigure(body, chart);
          if (!rendered) {
            this.destroyChartRuntime(chart.id);
            body.classList.remove("is-scrollable");
            this.renderChartPreview(body, chart);
          }
          this.updateChartRenderStatus(sheet.id, chart.id, "rendered", undefined, Date.now() - renderStartedAt);
        } catch (error) {
          this.destroyChartRuntime(chart.id);
          body.classList.remove("is-scrollable");
          this.engine.reportChartError({
            sheetId: sheet.id,
            chartId: chart.id,
            errorCode: "RENDERER_CHART_PREVIEW_FAILED",
            message: error instanceof Error ? error.message : "Chart preview failed."
          });
          try {
            this.renderChartPreview(body, chart);
          } catch {
            body.replaceChildren();
            const placeholder = document.createElement("div");
            placeholder.className = "excelsior-chart-preview-placeholder";
            placeholder.textContent = this.messages.chartInsertError;
            body.append(placeholder);
          }
          this.updateChartRenderStatus(sheet.id, chart.id, "skipped", "render-error");
        }
      }

      fragment.append(object);
    }

    const activeRenderStateKeys = new Set(charts.map((chart) => `${sheet.id}:${chart.id}`));
    for (const key of this.chartRenderStatusById.keys()) {
      if (key.startsWith(`${sheet.id}:`) && !activeRenderStateKeys.has(key)) {
        this.chartRenderStatusById.delete(key);
      }
    }

    const imageIds = new Set<string>();
    for (const image of this.engine.getImages(sheet.id).filter((item) => item.state.visible !== false)) {
      imageIds.add(image.id);
      const rect = this.resolveChartRect(image.position, rowOffsets, colOffsets, sheet.rowCount, sheet.columnCount);
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
      const rect = this.resolveChartRect(widget.position, rowOffsets, colOffsets, sheet.rowCount, sheet.columnCount);
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
    this.chartsLayer.replaceChildren(fragment);
  }

  private createImageObjectElement(image: WorksheetImageObject): HTMLElement {
    const object = document.createElement("section");
    object.className = "excelsior-chart-object excelsior-image-object";
    object.dataset.imageId = image.id;
    object.tabIndex = 0;
    object.setAttribute("role", "group");
    object.setAttribute("aria-label", image.alt || "Imagem");
    const header = document.createElement("header");
    header.className = "excelsior-chart-object-header";
    header.dataset.imageMove = "true";
    const title = document.createElement("div");
    title.className = "excelsior-chart-object-title";
    title.textContent = image.alt || "Imagem";
    const actions = document.createElement("div");
    actions.className = "excelsior-chart-object-actions";
    actions.append(
      this.createVisualObjectAction("image", "back", "Enviar imagem para trás", "↓"),
      this.createVisualObjectAction("image", "front", "Trazer imagem para frente", "↑"),
      this.createVisualObjectAction("image", "lock", "Bloquear ou desbloquear imagem", "L"),
      this.createVisualObjectAction("image", "delete", "Excluir imagem", "×", true)
    );
    header.append(title, actions);
    const body = document.createElement("div");
    body.className = "excelsior-chart-object-body";
    const picture = document.createElement("img");
    picture.dataset.imageContent = "true";
    picture.draggable = false;
    body.append(picture);
    const resize = document.createElement("button");
    resize.type = "button";
    resize.className = "excelsior-chart-object-resize";
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
    button.className = destructive ? "excelsior-chart-object-delete" : "excelsior-visual-object-action";
    button.dataset[`${kind}Action`] = action;
    button.setAttribute("aria-label", label);
    button.title = label;
    button.textContent = text;
    return button;
  }

  private createWidgetObjectElement(widget: WorksheetWidgetObject): HTMLElement {
    const object = document.createElement("section");
    object.className = "excelsior-chart-object excelsior-widget-object";
    object.dataset.widgetId = widget.id;
    object.tabIndex = 0;
    object.setAttribute("role", "group");
    object.setAttribute("aria-label", widget.label);
    const header = document.createElement("header");
    header.className = "excelsior-chart-object-header";
    header.dataset.widgetMove = "true";
    const title = document.createElement("div");
    title.className = "excelsior-chart-object-title";
    title.textContent = widget.label;
    const actions = document.createElement("div");
    actions.className = "excelsior-chart-object-actions";
    actions.append(
      this.createVisualObjectAction("widget", "back", `Enviar ${widget.label} para trás`, "↓"),
      this.createVisualObjectAction("widget", "front", `Trazer ${widget.label} para frente`, "↑"),
      this.createVisualObjectAction("widget", "lock", `Bloquear ou desbloquear ${widget.label}`, "L"),
      this.createVisualObjectAction("widget", "delete", `Excluir ${widget.label}`, "×", true)
    );
    header.append(title, actions);
    const body = document.createElement("div");
    body.className = "excelsior-chart-object-body excelsior-widget-object-body";
    body.dataset.widgetBody = "true";
    const resize = document.createElement("button");
    resize.type = "button";
    resize.className = "excelsior-chart-object-resize";
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
      placeholder.className = "excelsior-chart-preview-placeholder";
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
      placeholder.className = "excelsior-chart-preview-placeholder";
      placeholder.textContent = `Falha ao renderizar widget: ${widget.type}`;
      body.append(placeholder);
    }
  }

  private readonly handleChartLayerMouseDown = (event: MouseEvent): void => {
    if (event.button !== 0) {
      return;
    }
    const target = event.target as HTMLElement | null;
    const objectElement = target?.closest<HTMLElement>("[data-chart-id], [data-image-id], [data-widget-id]");
    if (!objectElement) {
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const kind: ChartInteractionState["kind"] = objectElement.dataset.chartId
      ? "chart"
      : objectElement.dataset.imageId
        ? "image"
        : "widget";
    const objectId = objectElement.dataset.chartId ?? objectElement.dataset.imageId ?? objectElement.dataset.widgetId;
    if (!objectId) {
      return;
    }

    if (target?.closest("[data-chart-action], [data-image-action], [data-widget-action]")) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }

    const visualObject = this.getVisualObject(kind, sheet.id, objectId);
    if (!visualObject || visualObject.state.locked) {
      return;
    }
    const metrics = this.chartSurfaceMetrics;
    if (!metrics || metrics.sheetId !== sheet.id) {
      return;
    }

    const resizeHandle = target?.closest("[data-chart-resize='true'], [data-image-resize='true'], [data-widget-resize='true']");
    const moveHandle = target?.closest("[data-chart-move='true'], [data-image-move='true'], [data-widget-move='true']");
    if (!resizeHandle && !moveHandle) {
      return;
    }

    const mode: ChartInteractionState["mode"] = resizeHandle ? "resize" : "move";
    const originRect = this.resolveChartRect(visualObject.position, metrics.rowOffsets, metrics.colOffsets, metrics.rowCount, metrics.colCount);
    this.chartInteraction = {
      mode,
      kind,
      sheetId: sheet.id,
      chartId: objectId,
      pointerStartX: event.clientX,
      pointerStartY: event.clientY,
      originRect,
      liveRect: { ...originRect }
    };
    if (kind === "chart") {
      this.setChartSelection(sheet.id, objectId);
      this.clearChartFeedback(sheet.id);
    } else if (kind === "image") {
      this.engine.selectImage(sheet.id, objectId);
    } else {
      this.engine.selectWidget(sheet.id, objectId);
    }
    event.preventDefault();
    event.stopPropagation();
    this.render();
    this.focus();
  };

  private readonly handleChartLayerClick = (event: Event): void => {
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
    const chartElement = target?.closest<HTMLElement>("[data-chart-id]");
    if (!chartElement) {
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const chartId = chartElement.dataset.chartId;
    if (!chartId) {
      return;
    }
    this.setChartSelection(sheet.id, chartId);

    if (target?.closest("[data-chart-action='delete']")) {
      if (this.engine.getChart(sheet.id, chartId)) {
        this.engine.deleteChart({ sheetId: sheet.id, chartId });
      }
      this.selectedChartId = undefined;
      this.setChartFeedback(sheet.id, this.messages.chartDeleted, false);
      this.render();
      this.focus();
      event.preventDefault();
      event.stopPropagation();
      return;
    }

    this.requestRender();
  };

  private getVisualObject(kind: ChartInteractionState["kind"], sheetId: string, objectId: string): WorksheetChartObject | WorksheetImageObject | WorksheetWidgetObject | undefined {
    if (kind === "chart") return this.engine.getChart(sheetId, objectId);
    if (kind === "image") return this.engine.getImage(sheetId, objectId);
    return this.engine.getWidget(sheetId, objectId);
  }

  private commitVisualObjectGeometry(
    kind: ChartInteractionState["kind"],
    sheetId: string,
    objectId: string,
    mode: ChartInteractionState["mode"],
    rect: ChartRect
  ): void {
    const metrics = this.chartSurfaceMetrics;
    const visualObject = this.getVisualObject(kind, sheetId, objectId);
    if (!metrics || metrics.sheetId !== sheetId || !visualObject) return;
    const position = this.resolveChartPositionFromRect(rect, metrics, visualObject.position);
    if (mode === "move") {
      const next = {
        fromCell: position.fromCell,
        toCell: position.toCell,
        offsetX: position.offsetX,
        offsetY: position.offsetY,
        zIndex: position.zIndex
      };
      if (kind === "chart") this.engine.moveChart({ sheetId, chartId: objectId, position: next });
      else if (kind === "image") this.engine.moveImage({ sheetId, imageId: objectId, position: next });
      else this.engine.moveWidget({ sheetId, widgetId: objectId, position: next });
      return;
    }
    const next = { width: position.width, height: position.height, toCell: position.toCell };
    if (kind === "chart") this.engine.resizeChart({ sheetId, chartId: objectId, position: next });
    else if (kind === "image") this.engine.resizeImage({ sheetId, imageId: objectId, position: next });
    else this.engine.resizeWidget({ sheetId, widgetId: objectId, position: next });
  }

  private readonly handleVisualObjectKeyDown = (event: KeyboardEvent): void => {
    const target = event.target as HTMLElement | null;
    if (target?.closest("[data-image-action], [data-widget-action]")) return;
    const element = target?.closest<HTMLElement>("[data-image-id], [data-widget-id]");
    if (!element) return;
    const kind: ChartInteractionState["kind"] = element.dataset.imageId ? "image" : "widget";
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
      if (visualObject.state.locked || !this.chartSurfaceMetrics) return;
      const delta = event.ctrlKey ? 1 : 10;
      const rect = this.resolveChartRect(
        visualObject.position,
        this.chartSurfaceMetrics.rowOffsets,
        this.chartSurfaceMetrics.colOffsets,
        this.chartSurfaceMetrics.rowCount,
        this.chartSurfaceMetrics.colCount
      );
      const horizontal = event.key === "ArrowLeft" ? -delta : event.key === "ArrowRight" ? delta : 0;
      const vertical = event.key === "ArrowUp" ? -delta : event.key === "ArrowDown" ? delta : 0;
      if (event.shiftKey) {
        rect.width = Math.max(CHART_MIN_WIDTH, rect.width + horizontal);
        rect.height = Math.max(CHART_MIN_HEIGHT, rect.height + vertical);
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

  private readonly handleChartInteractionMouseMove = (event: MouseEvent): void => {
    const interaction = this.chartInteraction;
    if (!interaction) {
      return;
    }
    const throttleMs = this.getChartPerformanceOptions().interactionThrottleMs;
    const now = Date.now();
    if (throttleMs > 0 && now - this.chartLastInteractionMoveTs < throttleMs) {
      event.preventDefault();
      return;
    }
    this.chartLastInteractionMoveTs = now;
    const metrics = this.chartSurfaceMetrics;
    if (!metrics || metrics.sheetId !== interaction.sheetId) {
      return;
    }

    const deltaX = event.clientX - interaction.pointerStartX;
    const deltaY = event.clientY - interaction.pointerStartY;
    const maxSurfaceWidth = metrics.colOffsets[metrics.colOffsets.length - 1] ?? 0;
    const maxSurfaceHeight = metrics.rowOffsets[metrics.rowOffsets.length - 1] ?? 0;
    const nextRect: ChartRect = {
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
        CHART_MIN_WIDTH,
        Math.max(CHART_MIN_WIDTH, ROW_HEADER_WIDTH + maxSurfaceWidth - interaction.originRect.left)
      );
      nextRect.height = clampNumeric(
        interaction.originRect.height + deltaY,
        CHART_MIN_HEIGHT,
        Math.max(CHART_MIN_HEIGHT, maxSurfaceHeight - interaction.originRect.top)
      );
    }

    interaction.liveRect = nextRect;
    this.chartInteraction = interaction;
    const selector = interaction.kind === "chart"
      ? `[data-chart-id='${interaction.chartId}']`
      : interaction.kind === "image"
        ? `[data-image-id='${interaction.chartId}']`
        : `[data-widget-id='${interaction.chartId}']`;
    const objectElement = this.chartsLayer.querySelector<HTMLElement>(selector);
    if (objectElement) {
      objectElement.style.left = `${nextRect.left}px`;
      objectElement.style.top = `${nextRect.top}px`;
      objectElement.style.width = `${nextRect.width}px`;
      objectElement.style.height = `${nextRect.height}px`;
    }
    event.preventDefault();
  };

  private readonly handleChartInteractionMouseUp = (): void => {
    const interaction = this.chartInteraction;
    if (!interaction) {
      return;
    }
    this.chartLastInteractionMoveTs = 0;
    this.chartInteraction = undefined;
    const metrics = this.chartSurfaceMetrics;
    if (!this.getVisualObject(interaction.kind, interaction.sheetId, interaction.chartId) || !metrics || metrics.sheetId !== interaction.sheetId) {
      this.requestRender();
      return;
    }

    try {
      this.commitVisualObjectGeometry(interaction.kind, interaction.sheetId, interaction.chartId, interaction.mode, interaction.liveRect);
    } catch (error) {
      if (interaction.kind === "chart") {
        this.engine.reportChartError({
          sheetId: interaction.sheetId,
          chartId: interaction.chartId,
          errorCode: "RENDERER_CHART_INTERACTION_FAILED",
          message: error instanceof Error ? error.message : "Chart interaction failed."
        });
      }
      // Renderer interaction should not crash runtime if chart update fails.
    }
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