import type {
  CellModel,
  CellPrimitive,
  PivotAggregateFunction,
  PivotBuildAsyncOptions,
  PivotInferenceInput,
  PivotSheetInput,
  PivotValueDefinition,
  SheetModel,
  WorkbookDataInput,
  WorkbookModel
} from "../domain/types";
import { createCoreOperationError } from "../errors/spreadsheet-operation-error";
import { getCellKey } from "../utils/cell-key";

interface AggregateState {
  sum: number;
  count: number;
  min?: number;
  max?: number;
}

interface PivotRecord {
  [field: string]: CellPrimitive;
}

interface PivotTreeNode {
  label: string;
  path: string[];
  children: PivotTreeNode[];
}

interface PivotContext {
  input: PivotSheetInput;
  sourceSheet: SheetModel;
  headers: string[];
  rowFields: string[];
  columnFields: string[];
  totalSourceRows: number;
}

interface PivotAggregation {
  rowOrder: string[][];
  columnOrder: string[][];
  rowSeen: Set<string>;
  columnSeen: Set<string>;
  leafAggregates: Map<string, AggregateState>;
  rowTotals: Map<string, AggregateState>;
  columnTotals: Map<string, AggregateState>;
  grandTotals: Map<string, AggregateState>;
  subtotalAggregates: Map<string, AggregateState>;
  subtotalRowTotals: Map<string, AggregateState>;
}

const BLANK_PIVOT_VALUE = "(blank)";
const TUPLE_SEPARATOR = "\u0001";

const tupleKey = (parts: string[]): string => parts.join(TUPLE_SEPARATOR);

const normalizeRange = <T extends PivotInferenceInput | PivotSheetInput>(input: T): T => ({
  ...input,
  sourceRange: {
    start: {
      row: Math.min(input.sourceRange.start.row, input.sourceRange.end.row),
      col: Math.min(input.sourceRange.start.col, input.sourceRange.end.col)
    },
    end: {
      row: Math.max(input.sourceRange.start.row, input.sourceRange.end.row),
      col: Math.max(input.sourceRange.start.col, input.sourceRange.end.col)
    }
  }
});

const getCellEffectiveValue = (sheet: SheetModel, row: number, col: number): CellPrimitive => {
  const cell = sheet.cells[getCellKey(row, col)];
  if (!cell) {
    return null;
  }

  return cell.formula ? cell.computedValue ?? null : cell.value;
};

const toPivotLabel = (value: CellPrimitive): string => {
  if (value == null || value === "") {
    return BLANK_PIVOT_VALUE;
  }

  return String(value);
};

const toNumericValue = (value: CellPrimitive): number | undefined => {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : undefined;
  }

  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }

  if (typeof value === "string") {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }

  return undefined;
};

const updateAggregate = (state: AggregateState | undefined, value: CellPrimitive): AggregateState => {
  const next: AggregateState = state
    ? { ...state }
    : {
        sum: 0,
        count: 0,
        min: undefined,
        max: undefined
      };
  const numeric = toNumericValue(value);

  if (numeric !== undefined) {
    next.sum += numeric;
    next.min = next.min === undefined ? numeric : Math.min(next.min, numeric);
    next.max = next.max === undefined ? numeric : Math.max(next.max, numeric);
  }

  if (value != null && value !== "") {
    next.count += 1;
  }

  return next;
};

const finalizeAggregate = (state: AggregateState | undefined, aggregate: PivotAggregateFunction): CellPrimitive => {
  if (!state) {
    return aggregate === "count" ? 0 : null;
  }

  switch (aggregate) {
    case "sum":
      return state.sum;
    case "avg":
      return state.count > 0 ? state.sum / state.count : null;
    case "min":
      return state.min ?? null;
    case "max":
      return state.max ?? null;
    case "count":
      return state.count;
  }
};

const aggregateLabel = (value: PivotValueDefinition): string => value.as ?? `${value.aggregate.toUpperCase()} ${value.field}`;

const buildHeaders = (sheet: SheetModel, input: PivotInferenceInput | PivotSheetInput): string[] => {
  const headers: string[] = [];
  for (let col = input.sourceRange.start.col; col <= input.sourceRange.end.col; col += 1) {
    const header = getCellEffectiveValue(sheet, input.sourceRange.start.row, col);
    const label = header == null || header === "" ? "" : String(header);
    if (!label) {
      throw new Error(`Pivot source header cannot be blank at column ${col}.`);
    }
    headers.push(label);
  }

  return headers;
};

const ensureFields = (headers: string[], fields: string[], role: string): void => {
  for (const field of fields) {
    if (!headers.includes(field)) {
      throw new Error(`Pivot ${role} field not found in source headers: ${field}`);
    }
  }
};

const buildRowTree = (rowTuples: string[][]): PivotTreeNode => {
  const root: PivotTreeNode = {
    label: "",
    path: [],
    children: []
  };

  for (const tuple of rowTuples) {
    let current = root;

    for (const label of tuple) {
      let next = current.children.find((child) => child.label === label);
      if (!next) {
        next = {
          label,
          path: [...current.path, label],
          children: []
        };
        current.children.push(next);
      }
      current = next;
    }
  }

  return root;
};

const createPivotCell = (value: CellPrimitive): CellModel => ({
  value,
  computedValue: value
});

const createPivotAggregation = (): PivotAggregation => ({
  rowOrder: [],
  columnOrder: [],
  rowSeen: new Set<string>(),
  columnSeen: new Set<string>(),
  leafAggregates: new Map<string, AggregateState>(),
  rowTotals: new Map<string, AggregateState>(),
  columnTotals: new Map<string, AggregateState>(),
  grandTotals: new Map<string, AggregateState>(),
  subtotalAggregates: new Map<string, AggregateState>(),
  subtotalRowTotals: new Map<string, AggregateState>()
});

const resolveSourceSheet = (workbook: Readonly<WorkbookModel>, sourceSheetId: string): SheetModel => {
  const sourceSheet = workbook.sheets.find((sheet) => sheet.id === sourceSheetId);
  if (!sourceSheet) {
    throw new Error(`Pivot source sheet not found: ${sourceSheetId}`);
  }

  return sourceSheet;
};

const validatePivotSourceRange = (input: PivotInferenceInput | PivotSheetInput): void => {
  if (input.sourceRange.end.row <= input.sourceRange.start.row) {
    throw new Error("Pivot source range must include at least one header row and one data row.");
  }
};

const validatePivotSourceLimit = (
  workbook: Readonly<WorkbookModel>,
  input: PivotSheetInput,
  totalSourceRows: number
): void => {
  const maxPivotSourceRows = workbook.settings.maxPivotSourceRows;

  if (maxPivotSourceRows !== undefined && totalSourceRows > maxPivotSourceRows) {
    throw createCoreOperationError(
      "CORE_PIVOT_CLIENT_ROW_LIMIT_EXCEEDED",
      `Pivot source range exceeds the configured limit of ${maxPivotSourceRows} rows for client-side pivoting.`,
      {
        sourceSheetId: input.sourceSheetId,
        totalSourceRows,
        maxPivotSourceRows,
        executionMode: input.executionMode ?? "client"
      }
    );
  }
};

const createPivotContext = (workbook: Readonly<WorkbookModel>, rawInput: PivotSheetInput): PivotContext => {
  const input = normalizeRange(rawInput);
  const sourceSheet = resolveSourceSheet(workbook, input.sourceSheetId);

  validatePivotSourceRange(input);

  if (!input.values.length) {
    throw new Error("Pivot requires at least one value definition.");
  }

  const rowFields = input.rows ?? [];
  const columnFields = input.columns ?? [];
  const headers = buildHeaders(sourceSheet, input);
  ensureFields(headers, rowFields, "row");
  ensureFields(headers, columnFields, "column");
  ensureFields(
    headers,
    input.values.map((value) => value.field),
    "value"
  );

  const totalSourceRows = input.sourceRange.end.row - input.sourceRange.start.row;
  validatePivotSourceLimit(workbook, input, totalSourceRows);

  return {
    input,
    sourceSheet,
    headers,
    rowFields,
    columnFields,
    totalSourceRows
  };
};

const createPivotRecord = (
  sheet: SheetModel,
  input: PivotSheetInput,
  headers: string[],
  row: number
): PivotRecord | undefined => {
  const record: PivotRecord = {};
  let hasValue = false;

  for (let col = input.sourceRange.start.col; col <= input.sourceRange.end.col; col += 1) {
    const header = headers[col - input.sourceRange.start.col];
    const value = getCellEffectiveValue(sheet, row, col);
    record[header] = value;
    hasValue ||= value != null && value !== "";
  }

  return hasValue ? record : undefined;
};

const applyPivotRecord = (
  aggregation: PivotAggregation,
  input: PivotSheetInput,
  rowFields: string[],
  columnFields: string[],
  record: PivotRecord
): void => {
  const rowTuple = rowFields.map((field) => toPivotLabel(record[field]));
  const columnTuple = columnFields.map((field) => toPivotLabel(record[field]));
  const rowKey = tupleKey(rowTuple);
  const columnKey = tupleKey(columnTuple);

  if (!aggregation.rowSeen.has(rowKey)) {
    aggregation.rowSeen.add(rowKey);
    aggregation.rowOrder.push(rowTuple);
  }

  if (!aggregation.columnSeen.has(columnKey)) {
    aggregation.columnSeen.add(columnKey);
    aggregation.columnOrder.push(columnTuple);
  }

  input.values.forEach((valueDefinition, valueIndex) => {
    const value = record[valueDefinition.field];

    aggregation.leafAggregates.set(
      `${rowKey}|${columnKey}|${valueIndex}`,
      updateAggregate(aggregation.leafAggregates.get(`${rowKey}|${columnKey}|${valueIndex}`), value)
    );
    aggregation.rowTotals.set(
      `${rowKey}|${valueIndex}`,
      updateAggregate(aggregation.rowTotals.get(`${rowKey}|${valueIndex}`), value)
    );
    aggregation.columnTotals.set(
      `${columnKey}|${valueIndex}`,
      updateAggregate(aggregation.columnTotals.get(`${columnKey}|${valueIndex}`), value)
    );
    aggregation.grandTotals.set(
      `${valueIndex}`,
      updateAggregate(aggregation.grandTotals.get(`${valueIndex}`), value)
    );

    for (let prefixLength = 1; prefixLength < rowTuple.length; prefixLength += 1) {
      const prefixKey = tupleKey(rowTuple.slice(0, prefixLength));
      aggregation.subtotalAggregates.set(
        `${prefixKey}|${columnKey}|${valueIndex}`,
        updateAggregate(aggregation.subtotalAggregates.get(`${prefixKey}|${columnKey}|${valueIndex}`), value)
      );
      aggregation.subtotalRowTotals.set(
        `${prefixKey}|${valueIndex}`,
        updateAggregate(aggregation.subtotalRowTotals.get(`${prefixKey}|${valueIndex}`), value)
      );
    }
  });
};

const materializePivotSheet = (context: PivotContext, aggregation: PivotAggregation): WorkbookDataInput => {
  const { input, sourceSheet, rowFields, columnFields } = context;

  if (!aggregation.columnOrder.length) {
    aggregation.columnOrder.push([]);
  }

  const showRowTotals = columnFields.length > 0 && input.includeRowTotals !== false;
  const headerLabels = [
    ...rowFields,
    ...aggregation.columnOrder.flatMap((columnTuple) => {
      if (columnFields.length === 0) {
        return input.values.map((valueDefinition) => aggregateLabel(valueDefinition));
      }

      return input.values.map((valueDefinition) =>
        input.values.length > 1
          ? `${columnTuple.join(" / ")} • ${aggregateLabel(valueDefinition)}`
          : columnTuple.join(" / ")
      );
    }),
    ...(showRowTotals
      ? input.values.map((valueDefinition) =>
          input.values.length > 1 ? `Total • ${aggregateLabel(valueDefinition)}` : "Total"
        )
      : [])
  ];

  const outputRows: Array<{ labels: string[]; values: CellPrimitive[] }> = [];
  const buildValueColumns = (rowKey: string, aggregateMap: Map<string, AggregateState>, totalsMap?: Map<string, AggregateState>) => [
    ...aggregation.columnOrder.flatMap((columnTuple) =>
      input.values.map((valueDefinition, valueIndex) =>
        finalizeAggregate(aggregateMap.get(`${rowKey}|${tupleKey(columnTuple)}|${valueIndex}`), valueDefinition.aggregate)
      )
    ),
    ...(showRowTotals && totalsMap
      ? input.values.map((valueDefinition, valueIndex) =>
          finalizeAggregate(totalsMap.get(`${rowKey}|${valueIndex}`), valueDefinition.aggregate)
        )
      : [])
  ];

  if (rowFields.length === 0) {
    outputRows.push({
      labels: [],
      values: [
        ...aggregation.columnOrder.flatMap((columnTuple) =>
          input.values.map((valueDefinition, valueIndex) =>
            finalizeAggregate(
              aggregation.leafAggregates.get(`${tupleKey([])}|${tupleKey(columnTuple)}|${valueIndex}`),
              valueDefinition.aggregate
            )
          )
        ),
        ...(showRowTotals
          ? input.values.map((valueDefinition, valueIndex) =>
              finalizeAggregate(aggregation.grandTotals.get(`${valueIndex}`), valueDefinition.aggregate)
            )
          : [])
      ]
    });
  } else {
    const rowTree = buildRowTree(aggregation.rowOrder);

    const appendNodeRows = (node: PivotTreeNode): void => {
      if (node.children.length === 0) {
        const rowKey = tupleKey(node.path);
        outputRows.push({
          labels: node.path,
          values: buildValueColumns(rowKey, aggregation.leafAggregates, aggregation.rowTotals)
        });
        return;
      }

      for (const child of node.children) {
        appendNodeRows(child);
      }

      if (input.includeSubtotals !== false && node.path.length > 0 && node.path.length < rowFields.length) {
        outputRows.push({
          labels: Array.from({ length: rowFields.length }, (_value, index) => {
            if (index < node.path.length - 1) {
              return node.path[index] ?? "";
            }
            if (index === node.path.length - 1) {
              return `${node.label} Total`;
            }
            return "";
          }),
          values: buildValueColumns(tupleKey(node.path), aggregation.subtotalAggregates, aggregation.subtotalRowTotals)
        });
      }
    };

    for (const child of rowTree.children) {
      appendNodeRows(child);
    }
  }

  if (input.includeColumnTotals !== false) {
    outputRows.push({
      labels: Array.from({ length: rowFields.length }, (_value, index) => (index === 0 ? "Grand Total" : "")),
      values: [
        ...aggregation.columnOrder.flatMap((columnTuple) =>
          input.values.map((valueDefinition, valueIndex) =>
            finalizeAggregate(aggregation.columnTotals.get(`${tupleKey(columnTuple)}|${valueIndex}`), valueDefinition.aggregate)
          )
        ),
        ...(showRowTotals
          ? input.values.map((valueDefinition, valueIndex) =>
              finalizeAggregate(aggregation.grandTotals.get(`${valueIndex}`), valueDefinition.aggregate)
            )
          : [])
      ]
    });
  }

  const cells: Record<string, CellModel> = {};
  headerLabels.forEach((label, index) => {
    cells[getCellKey(0, index)] = createPivotCell(label);
  });

  outputRows.forEach((row, rowIndex) => {
    row.labels.forEach((label, columnIndex) => {
      cells[getCellKey(rowIndex + 1, columnIndex)] = createPivotCell(label);
    });
    row.values.forEach((value, valueIndex) => {
      cells[getCellKey(rowIndex + 1, rowFields.length + valueIndex)] = createPivotCell(value);
    });
  });

  return {
    name: input.sheetName ?? `${sourceSheet.name} Pivot`,
    rowCount: Math.max(1, outputRows.length + 1),
    columnCount: Math.max(1, headerLabels.length),
    cells,
    merges: [],
    columns: {},
    rows: {}
  };
};

const defaultYieldControl = async (): Promise<void> => {
  await Promise.resolve();
};

const createAbortError = (): Error => {
  if (typeof DOMException === "function") {
    return new DOMException("Pivot build aborted.", "AbortError");
  }

  const error = new Error("Pivot build aborted.");
  error.name = "AbortError";
  return error;
};

const throwIfAborted = (signal?: AbortSignal): void => {
  if (signal?.aborted) {
    throw createAbortError();
  }
};

const hasNumericValuesInColumn = (sheet: SheetModel, normalizedInput: PivotInferenceInput, columnIndex: number): boolean => {
  const sheetColumn = normalizedInput.sourceRange.start.col + columnIndex;

  for (let row = normalizedInput.sourceRange.start.row + 1; row <= normalizedInput.sourceRange.end.row; row += 1) {
    if (toNumericValue(getCellEffectiveValue(sheet, row, sheetColumn)) !== undefined) {
      return true;
    }
  }

  return false;
};

export const buildPivotSheet = (workbook: Readonly<WorkbookModel>, rawInput: PivotSheetInput): WorkbookDataInput => {
  const context = createPivotContext(workbook, rawInput);
  const aggregation = createPivotAggregation();

  for (let row = context.input.sourceRange.start.row + 1; row <= context.input.sourceRange.end.row; row += 1) {
    const record = createPivotRecord(context.sourceSheet, context.input, context.headers, row);
    if (record) {
      applyPivotRecord(aggregation, context.input, context.rowFields, context.columnFields, record);
    }
  }

  return materializePivotSheet(context, aggregation);
};

export const buildPivotSheetAsync = async (
  workbook: Readonly<WorkbookModel>,
  rawInput: PivotSheetInput,
  options?: PivotBuildAsyncOptions
): Promise<WorkbookDataInput> => {
  const context = createPivotContext(workbook, rawInput);
  const aggregation = createPivotAggregation();
  const chunkSize = Math.max(1, options?.chunkSize ?? 250);
  const yieldControl = options?.yieldControl ?? defaultYieldControl;
  let processed = 0;

  for (let row = context.input.sourceRange.start.row + 1; row <= context.input.sourceRange.end.row; row += 1) {
    throwIfAborted(options?.signal);

    const record = createPivotRecord(context.sourceSheet, context.input, context.headers, row);
    if (record) {
      applyPivotRecord(aggregation, context.input, context.rowFields, context.columnFields, record);
    }

    processed += 1;

    if (processed % chunkSize === 0 || processed === context.totalSourceRows) {
      options?.onProgress?.({
        phase: "aggregate",
        completed: processed,
        total: context.totalSourceRows
      });
      throwIfAborted(options?.signal);

      if (processed < context.totalSourceRows) {
        await yieldControl();
      }
    }
  }

  throwIfAborted(options?.signal);
  const nextSheet = materializePivotSheet(context, aggregation);
  options?.onProgress?.({ phase: "materialize", completed: 1, total: 1 });
  return nextSheet;
};

export const inferPivotSheetInput = (
  workbook: Readonly<WorkbookModel>,
  rawInput: PivotInferenceInput
): PivotSheetInput => {
  const input = normalizeRange(rawInput);
  const sourceSheet = resolveSourceSheet(workbook, input.sourceSheetId);

  validatePivotSourceRange(input);

  const headers = buildHeaders(sourceSheet, input);
  const numericHeaderIndexes = headers
    .map((_header, index) => index)
    .filter((index) => hasNumericValuesInColumn(sourceSheet, input, index));
  const valueColumnIndex = numericHeaderIndexes.at(-1) ?? Math.max(headers.length - 1, 0);
  const valueField = headers[valueColumnIndex] ?? headers[0];
  const aggregate: PivotAggregateFunction = numericHeaderIndexes.includes(valueColumnIndex) ? "sum" : "count";
  const dimensionFields = headers.filter((_header, index) => index !== valueColumnIndex);

  return {
    sourceSheetId: input.sourceSheetId,
    sourceRange: input.sourceRange,
    rows: dimensionFields.length > 0 ? [dimensionFields[0]] : [],
    columns: dimensionFields.slice(1),
    values: [{ field: valueField, aggregate }],
    includeSubtotals: true,
    sheetName: input.sheetName,
    executionMode: input.executionMode
  };
};