import type { PivotSheetInput, WorkbookDataInput } from "./domain/types";

export type DataSortDirection = "asc" | "desc";

export type DataRequestKind = "rows" | "pivotSheet";

export interface DataSortModelItem {
  field: string;
  direction: DataSortDirection;
}

export type DataSortModel = DataSortModelItem[];

export interface DataFilterDescriptor {
  operator: string;
  value?: unknown;
  values?: unknown[];
}

export type DataFilterModel = Record<string, DataFilterDescriptor>;

export interface DataPivotModelItem {
  field: string;
}

export type DataPivotModel = DataPivotModelItem[];

export interface DataAggregateModelItem {
  field: string;
  function: string;
  as?: string;
}

export type DataAggregateModel = DataAggregateModelItem[];

export interface DataGroupInfo {
  key: string;
  path?: string[];
  level?: number;
  childCount?: number;
  expanded?: boolean;
}

export interface DataSourceRow {
  index: number;
  hidden?: boolean;
  height?: number;
}

export interface DataRequest {
  sheetId: string;
  startRow: number;
  endRow: number;
  requestKind?: DataRequestKind;
  sortModel?: DataSortModel;
  filterModel?: DataFilterModel;
  groupKeys?: string[];
  expandedGroupPaths?: string[][];
  pivotModel?: DataPivotModel;
  aggregateModel?: DataAggregateModel;
  pivotInput?: PivotSheetInput;
  visibleColumns?: string[];
  requestId: string;
}

export interface DataResponse {
  rows: DataSourceRow[];
  totalRows?: number;
  nextCursor?: string;
  groupInfo?: DataGroupInfo[];
  pivotSheet?: WorkbookDataInput;
  warnings?: string[];
}

export interface DataSourceRequestContext {
  signal?: AbortSignal;
}

export interface DataSource {
  getRows(request: DataRequest, context?: DataSourceRequestContext): Promise<DataResponse>;
}