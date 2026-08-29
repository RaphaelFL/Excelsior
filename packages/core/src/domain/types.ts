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
  severity?: "error" | "warning";
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

export interface CellValidationCellComparisonRule extends CellValidationBaseRule {
  type: "cellComparison";
  reference: CellAddress;
  operator: "equals" | "notEquals" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual";
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
  | CellValidationCustomRule
  | CellValidationCellComparisonRule;

export interface CellValidationConfig {
  rules: CellValidationRule[];
}

export interface CellValidationIssue {
  code: string;
  message: string;
  ruleType: CellValidationRule["type"];
  validator?: string;
  severity?: "error" | "warning";
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
  strike?: boolean;
  rotation?: 0 | 45 | -45 | 90 | -90;
  wrap?: boolean;
  overflow?: "clip" | "ellipsis" | "visible";
  format?: string;
  indent?: number;
  border?: {
    top?: CellBorderStyle;
    right?: CellBorderStyle;
    bottom?: CellBorderStyle;
    left?: CellBorderStyle;
  };
}

export interface CellRichTextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
}

export interface CellRichTextSegment {
  text: string;
  style?: CellRichTextStyle;
  hyperlink?: string;
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

export type ClientSideSortDirection = "asc" | "desc";

export interface ClientSideSortDescriptor {
  column: number;
  direction: ClientSideSortDirection;
}

export type ClientSideFilterType = "text" | "number" | "date";

export type ClientSideFilterOperator =
  | "equals"
  | "contains"
  | "startsWith"
  | "gt"
  | "gte"
  | "lt"
  | "lte"
  | "between";

export interface ClientSideFilterDescriptor {
  column: number;
  type: ClientSideFilterType;
  operator: ClientSideFilterOperator;
  value: string | number;
  valueTo?: string | number;
  caseSensitive?: boolean;
}

export interface ClientSideQueryState {
  sort: ClientSideSortDescriptor[];
  filters: ClientSideFilterDescriptor[];
  hasHeader: boolean;
}

export interface CommentAuthor {
  id: string;
  name?: string;
}

export interface CellCommentReply {
  id: string;
  author: CommentAuthor;
  content: string;
  createdAt: number;
  updatedAt?: number;
}

export interface CellComment {
  id: string;
  author: CommentAuthor;
  content: string;
  createdAt: number;
  updatedAt?: number;
  resolved: boolean;
  replies: CellCommentReply[];
}

export interface CellModel {
  value: CellPrimitive;
  formula?: string;
  computedValue?: CellPrimitive;
  error?: SpreadsheetError;
  validation?: CellValidationConfig;
  style?: CellStyle;
  note?: string;
  comments?: CellComment[];
  richText?: CellRichTextSegment[];
  metadata?: Record<string, unknown>;
}

export interface WorksheetObjectPosition {
  fromCell: string;
  toCell?: string;
  offsetX: number;
  offsetY: number;
  width: number;
  height: number;
  zIndex: number;
}

export interface WorksheetImageObject {
  id: string;
  sheetId: string;
  src: string;
  alt: string;
  position: WorksheetObjectPosition;
  style?: {
    objectFit?: "contain" | "cover" | "fill";
    borderColor?: string;
    borderWidth?: number;
    opacity?: number;
  };
  state: {
    selected: boolean;
    visible: boolean;
    locked: boolean;
  };
}

export interface WorksheetImageObjectInput {
  id?: string;
  src: string;
  alt?: string;
  position: Omit<WorksheetObjectPosition, "zIndex"> & { zIndex?: number };
  style?: WorksheetImageObject["style"];
  state?: Partial<WorksheetImageObject["state"]>;
}

export type JsonPrimitive = string | number | boolean | null;

export type JsonValue = JsonPrimitive | JsonValue[] | { [key: string]: JsonValue };

export interface WorksheetWidgetObject {
  id: string;
  sheetId: string;
  type: string;
  label: string;
  config: { [key: string]: JsonValue };
  data?: JsonValue;
  position: WorksheetObjectPosition;
  state: {
    selected: boolean;
    visible: boolean;
    locked: boolean;
  };
}

export interface WorksheetWidgetObjectInput {
  id?: string;
  type: string;
  label?: string;
  config?: { [key: string]: JsonValue };
  data?: JsonValue;
  position: Omit<WorksheetObjectPosition, "zIndex"> & { zIndex?: number };
  state?: Partial<WorksheetWidgetObject["state"]>;
}

export interface SheetSplitPane {
  horizontalRow?: number;
  verticalColumn?: number;
}

export interface SheetModel {
  id: string;
  name: string;
  cells: Record<string, CellModel>;
  merges: SheetMerge[];
  conditionalFormats?: ConditionalFormattingRule[];
  frozenRows?: number;
  frozenColumns?: number;
  splitPane?: SheetSplitPane;
  columns: Record<number, ColumnSchema>;
  rows: Record<number, RowSchema>;
  rowCount: number;
  columnCount: number;
  selection: CellRange;
  images?: WorksheetImageObject[];
  widgets?: WorksheetWidgetObject[];
  metadata?: Record<string, unknown>;
}

export interface WorkbookSettings {
  maxRows: number;
  maxColumns: number;
  maxCellLength: number;
  maxFormulaLength: number;
  maxPasteCells: number;
  maxImageSourceLength?: number;
  maxWidgetDataLength?: number;
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
  splitPane?: SheetSplitPane;
  columns?: Record<number, ColumnSchema>;
  rows?: Record<number, RowSchema>;
  images?: WorksheetImageObject[];
  widgets?: WorksheetWidgetObject[];
  metadata?: Record<string, unknown>;
}

export interface WorkbookConfig {
  data?: WorkbookDataInput[];
  settings?: Partial<WorkbookSettings>;
  metadata?: Record<string, unknown>;
  pivotModule?: PivotModule | false;
  collaboration?: CollaborationConfig;
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

export interface CollaborationEnvelope {
  id: string;
  workbookId: string;
  clientId: string;
  sequence: number;
  timestamp: number;
  sheetId?: string;
  operations: SpreadsheetOperation[];
}

export type CollaborationConflictPolicy = "last-write-wins";

export interface CollaborationPresence {
  clientId: string;
  sequence: number;
  updatedAt: number;
  expiresAt: number;
  user?: CommentAuthor;
  cursor?: CellAddress & { sheetId: string };
  selection?: { sheetId: string; range: CellRange };
  metadata?: Record<string, unknown>;
}

export type CollaborationPresenceMessage =
  | {
      type: "presence:update";
      workbookId: string;
      clientId: string;
      sequence: number;
      timestamp: number;
      presence: CollaborationPresence;
    }
  | {
      type: "presence:remove";
      workbookId: string;
      clientId: string;
      sequence: number;
      timestamp: number;
    };

export interface CollaborationConnection {
  workbookId: string;
  clientId: string;
  receive: (envelope: CollaborationEnvelope) => void;
  receivePresence?: (message: CollaborationPresenceMessage) => void;
  receiveError?: (error: unknown) => void;
}

export interface CollaborationAdapter {
  connect(connection: CollaborationConnection): void | Promise<void>;
  send(envelope: CollaborationEnvelope): void | Promise<void>;
  getPresence?(workbookId: string): CollaborationPresence[] | Promise<CollaborationPresence[]>;
  updatePresence?(message: Extract<CollaborationPresenceMessage, { type: "presence:update" }>): void | Promise<void>;
  removePresence?(message: Extract<CollaborationPresenceMessage, { type: "presence:remove" }>): void | Promise<void>;
  disconnect?(): void | Promise<void>;
}

export interface CollaborationConfig {
  adapter: CollaborationAdapter;
  clientId?: string;
  conflictPolicy?: CollaborationConflictPolicy;
  presenceTtlMs?: number;
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
  "collaboration:status": {
    timestamp: number;
    workbookId: string;
    clientId: string;
    status: "connecting" | "connected" | "disconnected" | "error";
  };
  "collaboration:operationsApplied": {
    timestamp: number;
    workbookId: string;
    clientId: string;
    envelopeId: string;
    operationCount: number;
  };
  "collaboration:error": {
    timestamp: number;
    workbookId: string;
    clientId: string;
    phase: "connect" | "send" | "receive" | "presence" | "disconnect";
    message: string;
  };
  "collaboration:presenceChanged": {
    timestamp: number;
    workbookId: string;
    presence: CollaborationPresence;
  };
  "collaboration:presenceRemoved": {
    timestamp: number;
    workbookId: string;
    clientId: string;
    reason: "removed" | "expired";
  };
  "cell:updated": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    address: CellAddress;
  };
  "cell:richTextChanged": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    address: CellAddress;
    richText?: CellRichTextSegment[];
  };
  "cell:noteChanged": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    address: CellAddress;
    note?: string;
  };
  "cell:commentCreated": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    address: CellAddress;
    comment: CellComment;
  };
  "cell:commentReplied": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    address: CellAddress;
    commentId: string;
    reply: CellCommentReply;
  };
  "cell:commentResolved": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    address: CellAddress;
    commentId: string;
    resolved: boolean;
  };
  "cell:commentDeleted": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    address: CellAddress;
    commentId: string;
  };
  "selection:changed": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    range: CellRange;
  };
  "split-pane:changed": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    splitPane?: SheetSplitPane;
  };
  "image:created": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    imageId: string;
    image: WorksheetImageObject;
  };
  "image:updated": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    imageId: string;
    image: WorksheetImageObject;
  };
  "image:deleted": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    imageId: string;
  };
  "image:moved": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    imageId: string;
    position: WorksheetObjectPosition;
  };
  "image:resized": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    imageId: string;
    position: WorksheetObjectPosition;
  };
  "image:selected": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    imageId: string;
  };
  "image:unselected": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    imageId: string;
  };
  "widget:created": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    widgetId: string;
    widget: WorksheetWidgetObject;
  };
  "widget:updated": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    widgetId: string;
    widget: WorksheetWidgetObject;
  };
  "widget:deleted": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    widgetId: string;
  };
  "widget:moved": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    widgetId: string;
    position: WorksheetObjectPosition;
  };
  "widget:resized": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    widgetId: string;
    position: WorksheetObjectPosition;
  };
  "row-model:changed": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
  };
  "client-side-query:applied": {
    timestamp: number;
    workbookId: string;
    sheetId: string;
    state: ClientSideQueryState;
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