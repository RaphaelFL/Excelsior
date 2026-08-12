import {
  CellValidationError,
  SpreadsheetOperationError,
  cellAddressToLabel,
  cellLabelToAddress,
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
  type ChartPosition,
  type PivotExecutionMode,
  type RowModelRow,
  type SheetMerge,
  type RowResult,
  type SpreadsheetOperation,
  type WorksheetChartObject,
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
  chartColumn: string;
  chartBar: string;
  chartLine: string;
  chartArea: string;
  chartPie: string;
  chartDonut: string;
  chartScatter: string;
  chartHistogram: string;
  chartBox: string;
  chartHeatmap: string;
  chartCandlestick: string;
  chartWaterfall: string;
  chartFunnel: string;
  chartPolar: string;
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
  chartHeatmap: "Heatmap",
  chartCandlestick: "Candlestick",
  chartWaterfall: "Waterfall",
  chartFunnel: "Funil",
  chartPolar: "Polar",
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
  | "chart-heatmap"
  | "chart-candlestick"
  | "chart-waterfall"
  | "chart-funnel"
  | "chart-polar"
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
  "chart-heatmap": "heatmap",
  "chart-candlestick": "candlestick",
  "chart-waterfall": "waterfall",
  "chart-funnel": "funnel",
  "chart-polar": "polar",
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
  { action: "chart-histogram", type: "histogram", category: "statistical", messageKey: "chartHistogram", enabled: false },
  { action: "chart-box", type: "box", category: "statistical", messageKey: "chartBox", enabled: false },
  { action: "chart-heatmap", type: "heatmap", category: "statistical", messageKey: "chartHeatmap", enabled: false },
  { action: "chart-candlestick", type: "candlestick", category: "financial", messageKey: "chartCandlestick", enabled: false },
  { action: "chart-waterfall", type: "waterfall", category: "financial", messageKey: "chartWaterfall", enabled: false },
  { action: "chart-funnel", type: "funnel", category: "financial", messageKey: "chartFunnel", enabled: false },
  { action: "chart-polar", type: "polar", category: "advanced", messageKey: "chartPolar", enabled: false },
  { action: "chart-treemap", type: "treemap", category: "advanced", messageKey: "chartTreemap", enabled: false },
  { action: "chart-sunburst", type: "sunburst", category: "advanced", messageKey: "chartSunburst", enabled: false },
  { action: "chart-sankey", type: "sankey", category: "advanced", messageKey: "chartSankey", enabled: false },
  { action: "chart-surface", type: "surface", category: "advanced", messageKey: "chartSurface", enabled: false },
  { action: "chart-surface3d", type: "surface3d", category: "advanced", messageKey: "chartSurface3d", enabled: false },
  { action: "chart-scatter3d", type: "scatter3d", category: "advanced", messageKey: "chartScatter3d", enabled: false }
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
  { type: "heatmap", messageKey: "chartHeatmap" },
  { type: "candlestick", messageKey: "chartCandlestick" },
  { type: "waterfall", messageKey: "chartWaterfall" },
  { type: "funnel", messageKey: "chartFunnel" },
  { type: "polar", messageKey: "chartPolar" },
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

  private readonly formulaBar = document.createElement("div");

  private readonly formulaAddress = document.createElement("span");

  private readonly formulaInput = document.createElement("input");

  private readonly statusMessage = document.createElement("span");

  private readonly findReplacePanel = document.createElement("div");

  private readonly findReplaceQueryInput = document.createElement("input");

  private readonly findReplaceValueInput = document.createElement("input");

  private readonly findReplaceResults = document.createElement("span");

  private readonly pivotPanel = document.createElement("div");

  private readonly chartEditPanel = document.createElement("div");

  private readonly pivotApplyButton = createFindReplaceActionButton("apply", "");

  private readonly pivotCloseButton = createFindReplaceActionButton("close", "");

  private readonly chartEditTypeSelect = document.createElement("select");

  private readonly chartEditRangeInput = document.createElement("input");

  private readonly chartEditTitleInput = document.createElement("input");

  private readonly chartEditLegendToggle = document.createElement("input");

  private readonly chartEditXAxisTitleInput = document.createElement("input");

  private readonly chartEditYAxisTitleInput = document.createElement("input");

  private readonly chartEditXAxisTypeSelect = document.createElement("select");

  private readonly chartEditYAxisTypeSelect = document.createElement("select");

  private readonly chartEditXAxisVisibleToggle = document.createElement("input");

  private readonly chartEditYAxisVisibleToggle = document.createElement("input");

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

  private readonly chartInsertPreviewHost = document.createElement("div");

  private readonly chartInsertPreviewInsertButton = createFindReplaceActionButton("insert", "");

  private readonly chartInsertPreviewCancelButton = createFindReplaceActionButton("cancel", "");

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

  private renderedHeaderColumns = new Set<number>();

  private autofillDrag?: {
    sourceRange: CellRange;
    preview?: AutofillPreview;
  };

  private activeCustomEditor?: CustomCellEditorInstance;

  private validationFeedback?: { sheetId: string; row: number; col: number; message: string };

  private rowModelFeedback?: { sheetId: string; error?: string };

  private pivotFeedback?: { sheetId: string; message: string; isError: boolean };

  private chartFeedback?: { sheetId: string; message: string; isError: boolean };

  private selectedChartId?: string;

  private chartInteraction?: ChartInteractionState;

  private chartSurfaceMetrics?: ChartSurfaceMetrics;

  private chartLastInteractionMoveTs = 0;

  private lastViewportScrollRenderTs = 0;

  private readonly chartRenderStatusById = new Map<string, string>();

  private readonly chartObjectElementById = new Map<string, HTMLElement>();

  private readonly chartBodyElementById = new Map<string, HTMLElement>();

  private readonly chartRuntimeById = new Map<string, ChartHandle>();

  private readonly chartRuntimeFigureById = new Map<string, string>();

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
    this.gridPanel.className = "excelsior-grid-panel";
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
    this.chartEditPanel.className = "excelsior-find-replace excelsior-chart-edit-panel";
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
    this.chartEditPanel.hidden = true;
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
    const chartHeader = document.createElement("span");
    chartHeader.className = "excelsior-chart-edit-panel-title";
    chartHeader.textContent = this.messages.chartEditPanelTitle;
    const chartActions = document.createElement("div");
    chartActions.className = "excelsior-find-replace-actions";
    chartActions.append(this.chartEditApplyButton, this.chartEditCloseButton);
    this.chartEditPanel.append(
      chartHeader,
      chartTitleField,
      chartTypeField,
      chartRangeField,
      chartXAxisTitleField,
      chartYAxisTitleField,
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
    const chartPreviewHeader = document.createElement("span");
    chartPreviewHeader.className = "excelsior-chart-edit-panel-title";
    chartPreviewHeader.textContent = this.messages.chartPreviewTitle;
    const chartPreviewActions = document.createElement("div");
    chartPreviewActions.className = "excelsior-find-replace-actions";
    chartPreviewActions.append(this.chartInsertPreviewInsertButton, this.chartInsertPreviewCancelButton);
    this.chartInsertPreviewPanel.append(chartPreviewHeader, this.chartInsertPreviewHost, chartPreviewActions);

    this.surface.append(this.cellsLayer, this.chartsLayer, this.editor, this.selectEditor, this.customEditorHost);
    this.formulaBar.append(
      this.formulaAddress,
      this.formulaInput,
      this.statusMessage,
      this.findReplacePanel,
      this.pivotPanel,
      this.chartEditPanel,
      this.chartInsertPreviewPanel
    );
    this.root.append(this.chrome, this.formulaBar, this.activeCellAnnouncement, this.gridPanel, this.sheetTabs);
    this.gridPanel.append(this.viewport, this.rowHeaders);
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
    this.viewport.removeEventListener("mousedown", this.handleViewportMouseDown);
    this.viewport.removeEventListener("keydown", this.handleKeyDown);
    this.viewport.removeEventListener("copy", this.handleCopy);
    this.viewport.removeEventListener("paste", this.handlePaste);
    this.toolbar.removeEventListener("click", this.handleToolbarClick);
    this.toolbar.removeEventListener("input", this.handleToolbarInput);
    this.toolbar.removeEventListener("keydown", this.handleToolbarKeyDown);
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
    this.cellsLayer.removeEventListener("click", this.handleCellClick);
    this.cellsLayer.removeEventListener("mousedown", this.handleAutofillMouseDown);
    this.cellsLayer.removeEventListener("mousemove", this.handleAutofillMouseMove);
    this.cellsLayer.removeEventListener("mouseup", this.handleAutofillMouseUp);
    this.cellsLayer.removeEventListener("dblclick", this.handleCellDoubleClick);
    this.chartsLayer.removeEventListener("mousedown", this.handleChartLayerMouseDown);
    this.chartsLayer.removeEventListener("click", this.handleChartLayerClick);
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
    globalThis.removeEventListener("mousemove", this.handleChartInteractionMouseMove);
    globalThis.removeEventListener("mouseup", this.handleChartInteractionMouseUp);
    this.destroyAllChartRuntimes();
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
    this.cellsLayer.addEventListener("click", this.handleCellClick);
    this.cellsLayer.addEventListener("mousedown", this.handleAutofillMouseDown);
    this.cellsLayer.addEventListener("mousemove", this.handleAutofillMouseMove);
    this.cellsLayer.addEventListener("mouseup", this.handleAutofillMouseUp);
    this.cellsLayer.addEventListener("dblclick", this.handleCellDoubleClick);
    this.chartsLayer.addEventListener("mousedown", this.handleChartLayerMouseDown);
    this.chartsLayer.addEventListener("click", this.handleChartLayerClick);
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
    this.setChartSelection(sheet.id, undefined);
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
        if (action && this.isChartToolbarAction(action)) {
          this.createChartFromSelection(action);
          this.focus();
        }
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
      Boolean(validationMessage ?? rowModelError) || pivotStatusIsError || chartStatusIsError
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
      bold: "B",
      italic: "I",
      wrap: "↵",
      "align-left": "≡←",
      "align-center": "≡",
      "align-right": "→≡",
      merge: "⇆",
      unmerge: "⇅",
      "insert-row": "+R",
      "delete-row": "-R",
      "insert-column": "+C",
      "delete-column": "-C",
      "create-pivot": "◫",
      "find-replace": "⌕",
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
      { action: "add-sheet", label: this.messages.addSheet, group: "data" },
      { action: "bold", label: this.messages.bold, group: "font" },
      { action: "italic", label: this.messages.italic, group: "font" },
      { action: "wrap", label: this.messages.wrap, group: "alignment" },
      { action: "align-left", label: this.messages.alignLeft, group: "alignment" },
      { action: "align-center", label: this.messages.alignCenter, group: "alignment" },
      { action: "align-right", label: this.messages.alignRight, group: "alignment" },
      { action: "merge", label: this.messages.merge, group: "alignment" },
      { action: "unmerge", label: this.messages.unmerge, group: "alignment" },
      { action: "insert-row", label: this.messages.insertRow, group: "structure" },
      { action: "delete-row", label: this.messages.deleteRow, group: "structure" },
      { action: "insert-column", label: this.messages.insertColumn, group: "structure" },
      { action: "delete-column", label: this.messages.deleteColumn, group: "structure" }
    ];
    const remoteOnlyActions = new Set([
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
    this.chartEditLegendToggle.checked = this.getChartLegendVisible(selectedChart);
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
      input.placeholderMode ||
      input.chartType === "heatmap" ||
      input.chartType === "histogram"
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
      let nextFigure = this.buildChartFigureFromRange({
        sheetId,
        chartType: chart.type,
        sourceRange,
        binding: {
          rangeAddress: chart.sourceRange.rangeAddress,
          orientation: chart.sourceRange.orientation,
          firstRowAsHeader: chart.sourceRange.firstRowAsHeader,
          firstColumnAsLabel: chart.sourceRange.firstColumnAsLabel,
          autoRefresh: chart.sourceRange.autoRefresh,
          categoryColumnIndex: chart.sourceRange.categoryColumnIndex,
          seriesColumnIndexes: chart.sourceRange.seriesColumnIndexes ? [...chart.sourceRange.seriesColumnIndexes] : undefined,
          valueColumnIndex: chart.sourceRange.valueColumnIndex
        },
        title: chart.title ?? "Chart",
        placeholderMode
      });
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

    this.chartsLayer.replaceChildren(fragment);
  }

  private readonly handleChartLayerMouseDown = (event: MouseEvent): void => {
    if (event.button !== 0) {
      return;
    }
    const target = event.target as HTMLElement | null;
    const chartElement = target?.closest<HTMLElement>("[data-chart-id]");
    if (!chartElement) {
      return;
    }

    const sheet = this.engine.getActiveSheet();
    const chartId = chartElement.dataset.chartId;
    if (!chartId) {
      return;
    }

    if (target?.closest("[data-chart-action='delete']")) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }

    const chart = this.engine.getChart(sheet.id, chartId);
    if (!chart || chart.state.locked) {
      return;
    }
    const metrics = this.chartSurfaceMetrics;
    if (!metrics || metrics.sheetId !== sheet.id) {
      return;
    }

    const resizeHandle = target?.closest("[data-chart-resize='true']");
    const moveHandle = target?.closest("[data-chart-move='true']");
    if (!resizeHandle && !moveHandle) {
      return;
    }

    const mode: ChartInteractionState["mode"] = resizeHandle ? "resize" : "move";
    const originRect = this.resolveChartRect(chart.position, metrics.rowOffsets, metrics.colOffsets, metrics.rowCount, metrics.colCount);
    this.chartInteraction = {
      mode,
      sheetId: sheet.id,
      chartId,
      pointerStartX: event.clientX,
      pointerStartY: event.clientY,
      originRect,
      liveRect: { ...originRect }
    };
    this.setChartSelection(sheet.id, chartId);
    this.clearChartFeedback(sheet.id);
    event.preventDefault();
    event.stopPropagation();
    this.render();
    this.focus();
  };

  private readonly handleChartLayerClick = (event: Event): void => {
    const target = event.target as HTMLElement | null;
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
    const chartElement = this.chartsLayer.querySelector<HTMLElement>(`[data-chart-id='${interaction.chartId}']`);
    if (chartElement) {
      chartElement.style.left = `${nextRect.left}px`;
      chartElement.style.top = `${nextRect.top}px`;
      chartElement.style.width = `${nextRect.width}px`;
      chartElement.style.height = `${nextRect.height}px`;
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
    const chart = this.engine.getChart(interaction.sheetId, interaction.chartId);
    const metrics = this.chartSurfaceMetrics;
    if (!chart || !metrics || metrics.sheetId !== interaction.sheetId) {
      this.requestRender();
      return;
    }

    const nextPosition = this.resolveChartPositionFromRect(interaction.liveRect, metrics, chart.position);
    try {
      if (interaction.mode === "move") {
        this.engine.moveChart({
          sheetId: interaction.sheetId,
          chartId: interaction.chartId,
          position: {
            fromCell: nextPosition.fromCell,
            toCell: nextPosition.toCell,
            offsetX: nextPosition.offsetX,
            offsetY: nextPosition.offsetY,
            zIndex: nextPosition.zIndex
          }
        });
      } else {
        this.engine.resizeChart({
          sheetId: interaction.sheetId,
          chartId: interaction.chartId,
          position: {
            width: nextPosition.width,
            height: nextPosition.height,
            toCell: nextPosition.toCell
          }
        });
      }
    } catch (error) {
      this.engine.reportChartError({
        sheetId: interaction.sheetId,
        chartId: interaction.chartId,
        errorCode: "RENDERER_CHART_INTERACTION_FAILED",
        message: error instanceof Error ? error.message : "Chart interaction failed."
      });
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