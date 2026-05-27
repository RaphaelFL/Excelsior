import type {
  CellAddress,
  CellBatchUpdate,
  CellTransactionChange,
  CellRange,
  CellStyle,
  CellValidationConfig,
  CellValidationResult,
  ColumnSchema,
  ConditionalFormattingRule,
  RowSchema,
  SafeCellValidator,
  SheetMerge,
  SpreadsheetEventMap,
  SpreadsheetOperation,
  WorkbookDataInput,
  WorkbookModel
} from "../domain/types";

export type PluginDisposer = () => void;

export interface PluginState {
  [key: string]: unknown;
}

export interface PluginCommandApi {
  setCellValue(input: {
    sheetId: string;
    row: number;
    col: number;
    value: string | number | boolean | null;
  }): SpreadsheetOperation[];
  updateCells(input: { sheetId: string; updates: CellBatchUpdate[]; affectedRanges?: CellRange[] }): SpreadsheetOperation[];
  applyCellTransaction(input: {
    sheetId: string;
    changes: CellTransactionChange[];
    affectedRanges?: CellRange[];
  }): SpreadsheetOperation[];
  setRemoteGroupExpanded(sheetId: string, path: string[], expanded: boolean): void;
  setCellStyle(input: {
    sheetId: string;
    row: number;
    col: number;
    style: Partial<CellStyle>;
    mode?: "merge" | "replace";
  }): SpreadsheetOperation[];
  setCellValidation(input: {
    sheetId: string;
    row: number;
    col: number;
    validation?: CellValidationConfig;
  }): SpreadsheetOperation[];
  setConditionalFormattingRules(sheetId: string, rules: ConditionalFormattingRule[]): SpreadsheetOperation[];
  freezeRows(sheetId: string, count: number): SpreadsheetOperation[];
  freezeColumns(sheetId: string, count: number): SpreadsheetOperation[];
  setRowsHidden(sheetId: string, start: number, end?: number, hidden?: boolean): SpreadsheetOperation[];
  setColumnsHidden(sheetId: string, start: number, end?: number, hidden?: boolean): SpreadsheetOperation[];
  resizeColumn(sheetId: string, col: number, width: number): SpreadsheetOperation[];
  resizeRow(sheetId: string, row: number, height: number): SpreadsheetOperation[];
  mergeCells(input: { sheetId: string; start: CellAddress; end: CellAddress }): SpreadsheetOperation[];
  unmergeCells(input: { sheetId: string; row: number; col: number }): SpreadsheetOperation[];
  selectRange(input: {
    sheetId: string;
    rowStart: number;
    rowEnd: number;
    colStart: number;
    colEnd: number;
  }): void;
  applyOperations(operations: SpreadsheetOperation[]): void;
  getSelection(sheetId: string): CellRange;
  getMerge(sheetId: string, row: number, col: number): SheetMerge | undefined;
  getColumnSchema(sheetId: string, col: number): ColumnSchema | undefined;
  getRowSchema(sheetId: string, row: number): RowSchema | undefined;
  addSheet(input?: WorkbookDataInput): string;
  deleteSheet(sheetId: string): void;
  insertRows(sheetId: string, index: number, count?: number): void;
  deleteRows(sheetId: string, start: number, end?: number): void;
  insertColumns(sheetId: string, index: number, count?: number): void;
  deleteColumns(sheetId: string, start: number, end?: number): void;
  undo(): boolean;
  redo(): boolean;
  reportSecurityEvent(reason: string, details?: Record<string, unknown>): void;
  validateCellValue(input: {
    sheetId: string;
    row: number;
    col: number;
    value: string | number | boolean | null;
  }): CellValidationResult;
  getCellValidation(sheetId: string, row: number, col: number): CellValidationConfig | undefined;
  getConditionalFormattingRules(sheetId: string): ConditionalFormattingRule[];
  getConditionalStyle(sheetId: string, row: number, col: number): CellStyle | undefined;
  getFrozenPane(sheetId: string): { rows: number; columns: number };
  registerValidator(id: string, validator: SafeCellValidator): void;
  unregisterValidator(id: string): void;
}

export interface PluginContext {
  pluginId: string;
  workbookId: string;
  commands: PluginCommandApi;
  getSnapshot(): WorkbookModel;
  getState<TState extends PluginState = PluginState>(): TState | undefined;
  setState<TState extends PluginState = PluginState>(
    nextState: TState | ((previousState: TState | undefined) => TState)
  ): TState;
  on<TKey extends keyof SpreadsheetEventMap>(
    eventName: TKey,
    listener: (payload: SpreadsheetEventMap[TKey]) => void
  ): PluginDisposer;
}

export interface GridPlugin {
  id: string;
  setup(context: PluginContext): void | PluginDisposer;
}

export interface RegisteredGridPlugin {
  id: string;
  enabled: boolean;
}