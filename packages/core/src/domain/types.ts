export type CellPrimitive = string | number | boolean | null;

export type ClipboardPolicy = "text-only" | "safe-html" | "blocked-html";

export type PivotExecutionMode = "auto" | "client" | "server";

export interface CellAddress {
  row: number;
  col: number;
}

export interface CellRange {
  start: CellAddress;
  end: CellAddress;
}

export interface CellBatchUpdate {
  row: number;
  col: number;
  value: CellPrimitive;
}

export interface CellKeyedUpsert {
  key: string;
  value: CellPrimitive;
}

export interface CellTransactionUpsertChange {
  type: "upsert";
  key: string;
  value: CellPrimitive;
}

export interface CellTransactionRemoveChange {
  type: "remove";
  key: string;
}

export type CellTransactionChange = CellTransactionUpsertChange | CellTransactionRemoveChange;

export interface SpreadsheetError {
  code: string;
  message: string;
  area: "core" | "renderer" | "formula" | "wrapper" | "security" | "pivot";
  recoverable: boolean;
  details?: Record<string, unknown>;
}

export interface CellValidationBaseRule {
  message?: string;
  allowFormula?: boolean;
}

export interface CellValidationTextRule extends CellValidationBaseRule {
  type: "text";
  minLength?: number;
  maxLength?: number;
}

export interface CellValidationNumberRule extends CellValidationBaseRule {
  type: "number";
  min?: number;
  max?: number;
}

export interface CellValidationDateRule extends CellValidationBaseRule {
  type: "date";
  min?: string;
  max?: string;
}

export interface CellValidationBooleanRule extends CellValidationBaseRule {
  type: "boolean";
}

export interface CellValidationListRule extends CellValidationBaseRule {
  type: "list" | "dropdown";
  values: CellPrimitive[];
}

export interface CellValidationCheckboxRule extends CellValidationBaseRule {
  type: "checkbox";
}

export interface CellValidationRequiredRule extends CellValidationBaseRule {
  type: "required";
}

export interface CellValidationRangeRule extends CellValidationBaseRule {
  type: "range";
  min: number;
  max: number;
}

export interface CellValidationLengthRule extends CellValidationBaseRule {
  type: "length";
  min?: number;
  max?: number;
}

export interface CellValidationRegexRule extends CellValidationBaseRule {
  type: "regex";
  pattern: string;
  flags?: string;
}

export interface CellValidationCustomRule extends CellValidationBaseRule {
  type: "custom";
  validator: string;
  params?: Record<string, unknown>;
}

export type CellValidationRule =
  | CellValidationTextRule
  | CellValidationNumberRule
  | CellValidationDateRule
  | CellValidationBooleanRule
  | CellValidationListRule
  | CellValidationCheckboxRule
  | CellValidationRequiredRule
  | CellValidationRangeRule
  | CellValidationLengthRule
  | CellValidationRegexRule
  | CellValidationCustomRule;

export interface CellValidationConfig {
  rules: CellValidationRule[];
}

export interface CellValidationIssue {
  code: string;
  message: string;
  ruleType: CellValidationRule["type"];
  validator?: string;
}

export interface CellValidationResult {
  valid: boolean;
  issue?: CellValidationIssue;
  error?: SpreadsheetError;
}

export interface CellValidationContext {
  workbook: WorkbookModel;
  sheet: SheetModel;
  cell?: CellModel;
  address: CellAddress;
  value: CellPrimitive;
  params?: Record<string, unknown>;
}

export type SafeCellValidator = (context: CellValidationContext) => CellValidationIssue | undefined;

export interface RegisteredCellValidator {
  id: string;
}

export interface CellBorderStyle {
  color?: string;
  style?: "thin" | "medium" | "thick" | "dashed" | "dotted" | "double";
}

export interface ConditionalFormattingBaseRule {
  id: string;
  range: CellRange;
  priority?: number;
}

export interface ConditionalFormattingComparisonRule extends ConditionalFormattingBaseRule {
  type: "greaterThan" | "lessThan";
  value: number;
  style: Partial<CellStyle>;
}

export interface ConditionalFormattingEqualityRule extends ConditionalFormattingBaseRule {
  type: "equal" | "notEqual";
  value: CellPrimitive;
  style: Partial<CellStyle>;
}

export interface ConditionalFormattingBetweenRule extends ConditionalFormattingBaseRule {
  type: "between";
  min: number;
  max: number;
  style: Partial<CellStyle>;
}

export interface ConditionalFormattingContainsTextRule extends ConditionalFormattingBaseRule {
  type: "containsText";
  text: string;
  style: Partial<CellStyle>;
}

export interface ConditionalFormattingDateRule extends ConditionalFormattingBaseRule {
  type: "dateBefore" | "dateAfter";
  date: string;
  style: Partial<CellStyle>;
}

export interface ConditionalFormattingDuplicatesRule extends ConditionalFormattingBaseRule {
  type: "duplicates";
  style: Partial<CellStyle>;
}

export interface ConditionalFormattingColorScaleRule extends ConditionalFormattingBaseRule {
  type: "colorScale";
  minColor: string;
  maxColor: string;
  textColor?: string;
}

export interface ConditionalFormattingFormulaRule extends ConditionalFormattingBaseRule {
  type: "formula";
  formula: string;
  style: Partial<CellStyle>;
}

export type ConditionalFormattingRule =
  | ConditionalFormattingComparisonRule
  | ConditionalFormattingEqualityRule
  | ConditionalFormattingBetweenRule
  | ConditionalFormattingContainsTextRule
  | ConditionalFormattingDateRule
  | ConditionalFormattingDuplicatesRule
  | ConditionalFormattingColorScaleRule
  | ConditionalFormattingFormulaRule;

export interface CellStyle {
  align?: "left" | "center" | "right";
  alignVertical?: "top" | "center" | "bottom";
  backgroundColor?: string;
  textColor?: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: "normal" | "bold";
  fontStyle?: "normal" | "italic";
  underline?: boolean;
  wrap?: boolean;
  format?: string;
  indent?: number;
  border?: {
    top?: CellBorderStyle;
    right?: CellBorderStyle;
    bottom?: CellBorderStyle;
    left?: CellBorderStyle;
  };
}

export interface SheetMerge {
  start: CellAddress;
  end: CellAddress;
}

export interface ColumnSchema {
  width?: number;
  hidden?: boolean;
  style?: CellStyle;
}

export interface RowSchema {
  height?: number;
  hidden?: boolean;
  style?: CellStyle;
}

export interface CellModel {
  value: CellPrimitive;
  formula?: string;
  computedValue?: CellPrimitive;
  error?: SpreadsheetError;
  validation?: CellValidationConfig;
  style?: CellStyle;
  note?: string;
  metadata?: Record<string, unknown>;
}

export type WorksheetChartType =
  | "column"
  | "bar"
  | "line"
  | "area"
  | "pie"
  | "donut"
  | "scatter"
  | "histogram"
  | "box"
  | "heatmap"
  | "candlestick"
  | "waterfall"
  | "funnel"
  | "polar"
  | "treemap"
  | "sunburst"
  | "sankey"
  | "surface"
  | "surface3d"
  | "scatter3d"
  | "unknown";

export interface ChartFigureSnapshot {
  data: unknown[];
  layout?: Record<string, unknown>;
  config?: Record<string, unknown>;
  frames?: unknown[];
  selection?: unknown;
  metadata?: Record<string, unknown>;
  schemaVersion?: string;
}

export interface ChartRangeBinding {
  chartId: string;
  sheetId: string;
  rangeAddress: string;
  orientation: "rows" | "columns";
  firstRowAsHeader: boolean;
  firstColumnAsLabel: boolean;
  autoRefresh: boolean;
  categoryColumnIndex?: number;
  seriesColumnIndexes?: number[];
  valueColumnIndex?: number;
}

export interface ChartPosition {
  fromCell: string;
  toCell?: string;
  offsetX: number;
  offsetY: number;
  width: number;
  height: number;
  zIndex: number;
}

export interface WorksheetChartObject {
  id: string;
  sheetId: string;
  type: WorksheetChartType;
  title?: string;
  sourceRange?: ChartRangeBinding;
  figure: ChartFigureSnapshot;
  position: ChartPosition;
  style?: {
    theme?: string;
    background?: string;
    borderColor?: string;
    borderWidth?: number;
    fontFamily?: string;
    professionalPreset?: "spreadsheet" | "report" | "dashboard";
  };
  state: {
    selected: boolean;
    visible: boolean;
    locked: boolean;
    lastRenderedAt?: number;
  };
  excelInterop: {
    originalChartType?: string;
    originalChartId?: string;
    originalAnchor?: unknown;
    unsupportedFeatures?: string[];
    fallbackImage?: boolean;
    preservedRawMetadata?: unknown;
  };
}

export interface WorksheetChartObjectInput {
  id?: string;
  type: WorksheetChartType;
  title?: string;
  sourceRange?: Omit<ChartRangeBinding, "chartId" | "sheetId">;
  figure: ChartFigureSnapshot;
  position: Omit<ChartPosition, "zIndex"> & { zIndex?: number };
  style?: WorksheetChartObject["style"];
  state?: Partial<WorksheetChartObject["state"]>;
  excelInterop?: Partial<WorksheetChartObject["excelInterop"]>;
}

export interface SheetModel {
  id: string;
  name: string;
  cells: Record<string, CellModel>;
  merges: SheetMerge[];
  conditionalFormats?: ConditionalFormattingRule[];
  frozenRows?: number;
  frozenColumns?: number;
  columns: Record<number, ColumnSchema>;
  rows: Record<number, RowSchema>;
  rowCount: number;
  columnCount: number;
  selection: CellRange;
  charts?: WorksheetChartObject[];
  metadata?: Record<string, unknown>;
}

export interface WorkbookSettings {
  maxRows: number;
  maxColumns: number;
  maxCellLength: number;
  maxFormulaLength: number;
  maxPasteCells: number;
  maxRecalcCells?: number;
  maxPivotSourceRows?: number;
  rowHeight: number;
  columnWidth: number;
  viewportBuffer: number;
  maxHistorySize: number;
  enableFormulas: boolean;
  clipboardPolicy: ClipboardPolicy;
}

export interface WorkbookModel {
  id: string;
  sheets: SheetModel[];
  activeSheetId: string;
  metadata: Record<string, unknown>;
  settings: WorkbookSettings;
}

export interface WorkbookDataInput {
  id?: string;
  name?: string;
  rowCount?: number;
  columnCount?: number;
  cells?: Record<string, CellModel>;
  merges?: SheetMerge[];
  conditionalFormats?: ConditionalFormattingRule[];
  frozenRows?: number;
  frozenColumns?: number;
  columns?: Record<number, ColumnSchema>;
  rows?: Record<number, RowSchema>;
  charts?: WorksheetChartObject[];
  metadata?: Record<string, unknown>;
}

export interface WorkbookConfig {
  data?: WorkbookDataInput[];
  settings?: Partial<WorkbookSettings>;
  metadata?: Record<string, unknown>;
  pivotModule?: PivotModule | false;
}

export type PivotAggregateFunction = "sum" | "avg" | "min" | "max" | "count";

export interface PivotValueDefinition {
  field: string;
  aggregate: PivotAggregateFunction;
  as?: string;
}

export interface PivotSheetInput {
  sourceSheetId: string;
  sourceRange: CellRange;
  rows?: string[];
  columns?: string[];
  values: PivotValueDefinition[];
  includeRowTotals?: boolean;
  includeColumnTotals?: boolean;
  includeSubtotals?: boolean;
  sheetName?: string;
  executionMode?: PivotExecutionMode;
}

export interface PivotInferenceInput {
  sourceSheetId: string;
  sourceRange: CellRange;
  sheetName?: string;
  executionMode?: PivotExecutionMode;
}

export interface PivotDerivedViewResult {
  executionMode: PivotExecutionMode;
  remote: boolean;
  rowCount?: number;
  columnCount?: number;
}

export interface PivotDerivedViewDefinition {
  kind: "pivot";
  input: PivotSheetInput;
  refreshedAt: number;
  autoRefresh?: boolean;
  stale?: boolean;
  refreshStatus?: "idle" | "refreshing" | "error";
  lastError?: string;
  result?: PivotDerivedViewResult;
}

export interface PivotBuildProgress {
  phase: "aggregate" | "materialize";
  completed: number;
  total: number;
}

export interface PivotBuildAsyncOptions {
  chunkSize?: number;
  yieldControl?: () => Promise<void>;
  signal?: AbortSignal;
  onProgress?: (progress: PivotBuildProgress) => void;
}

export interface PivotModule {
  createPivotSheet(workbook: Readonly<WorkbookModel>, input: PivotSheetInput): WorkbookDataInput;
  createPivotSheetAsync?(
    workbook: Readonly<WorkbookModel>,
    input: PivotSheetInput,
    options?: PivotBuildAsyncOptions
  ): Promise<WorkbookDataInput>;
  inferPivotSheet(workbook: Readonly<WorkbookModel>, input: PivotInferenceInput): PivotSheetInput;
}

export type SpreadsheetOperationName =
  | "add"
  | "replace"
  | "remove"
  | "insertRowCol"
  | "deleteRowCol"
  | "addSheet"
  | "deleteSheet";

export interface SpreadsheetOperation {
  op: SpreadsheetOperationName;
  id: string;
  path: Array<string | number>;
  value: unknown;
}

export interface SpreadsheetEventMap {
  "engine:created": {
    timestamp: number;
    workbookId: string;
  };
  "engine:disposed": {
    timestamp: number;
    workbookId: string;
  };
  "command:completed": {
    timestamp: number;
    workbookId: string;
    sheetId?: string;
    durationMs: number;
    commandType: string;
    operations: SpreadsheetOperation[];
  };
  "command:failed": {
    timestamp: number;
    workbookId: string;
    sheetId?: string;
    durationMs: number;
    commandType: string;
    errorCode: string;
  };
  "cell:updated": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    address: CellAddress;
  };
  "selection:changed": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    range: CellRange;
  };
  "chart:created": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    chart: WorksheetChartObject;
  };
  "chart:updated": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    chart: WorksheetChartObject;
  };
  "chart:deleted": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
  };
  "chart:moved": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    position: ChartPosition;
  };
  "chart:resized": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    position: ChartPosition;
  };
  "chart:selected": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
  };
  "chart:unselected": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
  };
  "chart:rangeChanged": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    range: ChartRangeBinding;
    reason: "binding-updated" | "source-cells-updated" | "manual-refresh";
  };
  "chart:dataInvalid": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    reason: string;
  };
  "chart:imported": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    chart: WorksheetChartObject;
  };
  "chart:exported": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
  };
  "chart:unsupportedFeature": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    feature: string;
  };
  "chart:error": {
    timestamp: number;
    workbookId: string;
    sheetId?: string;
    chartId?: string;
    errorCode: string;
    message: string;
  };
  "chart:renderStarted": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
  };
  "chart:renderFinished": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    durationMs?: number;
  };
  "chart:renderSkipped": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    chartId: string;
    reason: string;
  };
  "row-model:changed": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
  };
  "security:blocked-input": {
    timestamp: number;
    workbookId: string;
    reason: string;
    details?: Record<string, unknown>;
  };
  "formula:failed": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    address: CellAddress;
    errorCode: string;
  };
}

export interface FormulaEvaluationResult {
  value: CellPrimitive;
  error?: SpreadsheetError;
}

export interface FormulaEvaluationContext {
  currentCell: CellAddress;
  currentSheetId: string;
  currentSheetName: string;
  getCell: (row: number, col: number, sheetRef?: string) => CellModel | undefined;
  evaluateCell: (
    row: number,
    col: number,
    trail?: string[],
    sheetRef?: string
  ) => FormulaEvaluationResult;
}

export interface FormulaReference {
  row: number;
  col: number;
  sheetRef?: string;
}

export interface FormulaEngine {
  evaluate: (
    expression: string,
    context: FormulaEvaluationContext
  ) => FormulaEvaluationResult;
  collectReferences?: (expression: string) => FormulaReference[];
}

export interface CommandResult {
  workbook: WorkbookModel;
  operations: SpreadsheetOperation[];
  affectedRanges: CellRange[];
  recordHistory?: boolean;
}