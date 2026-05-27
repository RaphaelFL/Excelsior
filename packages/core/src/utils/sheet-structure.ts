import type { CellAddress, CellModel, CellRange, SheetMerge, SheetModel } from "../domain/types";
import { getCellKey } from "./cell-key";

export type Axis = "row" | "column";

const clamp = (value: number, min: number, max: number): number =>
  Math.min(Math.max(value, min), max);

const normalizeSelection = (selection: CellRange, sheet: SheetModel): CellRange => ({
  start: {
    row: clamp(selection.start.row, 0, Math.max(0, sheet.rowCount - 1)),
    col: clamp(selection.start.col, 0, Math.max(0, sheet.columnCount - 1))
  },
  end: {
    row: clamp(selection.end.row, 0, Math.max(0, sheet.rowCount - 1)),
    col: clamp(selection.end.col, 0, Math.max(0, sheet.columnCount - 1))
  }
});

const parseCellKey = (key: string): { row: number; col: number } => {
  const [rowText, colText] = key.split(":");
  return {
    row: Number(rowText),
    col: Number(colText)
  };
};

const shiftIndexedSchemaForInsert = <T>(
  schema: Record<number, T>,
  index: number,
  count: number
): Record<number, T> => {
  const nextSchema: Record<number, T> = {};

  for (const [key, value] of Object.entries(schema)) {
    const numericKey = Number(key);
    const nextKey = numericKey >= index ? numericKey + count : numericKey;
    nextSchema[nextKey] = value;
  }

  return nextSchema;
};

const shiftIndexedSchemaForDelete = <T>(
  schema: Record<number, T>,
  start: number,
  count: number
): Record<number, T> => {
  const end = start + count - 1;
  const nextSchema: Record<number, T> = {};

  for (const [key, value] of Object.entries(schema)) {
    const numericKey = Number(key);
    if (numericKey >= start && numericKey <= end) {
      continue;
    }

    const nextKey = numericKey > end ? numericKey - count : numericKey;
    nextSchema[nextKey] = value;
  }

  return nextSchema;
};

const shiftAddressForInsert = (address: CellAddress, axis: Axis, index: number, count: number): CellAddress => ({
  row: axis === "row" && address.row >= index ? address.row + count : address.row,
  col: axis === "column" && address.col >= index ? address.col + count : address.col
});

const shiftAddressForDelete = (address: CellAddress, axis: Axis, start: number, count: number): CellAddress | undefined => {
  const end = start + count - 1;
  const value = axis === "row" ? address.row : address.col;
  if (value >= start && value <= end) {
    return undefined;
  }

  return {
    row: axis === "row" && address.row > end ? address.row - count : address.row,
    col: axis === "column" && address.col > end ? address.col - count : address.col
  };
};

const shiftMergesForInsert = (merges: SheetMerge[], axis: Axis, index: number, count: number): SheetMerge[] =>
  merges.map((merge) => ({
    start: shiftAddressForInsert(merge.start, axis, index, count),
    end: shiftAddressForInsert(merge.end, axis, index, count)
  }));

const shiftMergesForDelete = (merges: SheetMerge[], axis: Axis, start: number, count: number): SheetMerge[] =>
  merges.flatMap((merge) => {
    const nextStart = shiftAddressForDelete(merge.start, axis, start, count);
    const nextEnd = shiftAddressForDelete(merge.end, axis, start, count);
    if (!nextStart || !nextEnd) {
      return [];
    }

    if (nextStart.row > nextEnd.row || nextStart.col > nextEnd.col) {
      return [];
    }

    return [{ start: nextStart, end: nextEnd }];
  });

const shiftSelectionForInsert = (
  selection: CellRange,
  axis: Axis,
  index: number,
  count: number,
  sheet: SheetModel
): CellRange => {
  const applyShift = (value: number): number => (value >= index ? value + count : value);
  const nextSelection: CellRange = {
    start: { ...selection.start },
    end: { ...selection.end }
  };

  if (axis === "row") {
    nextSelection.start.row = applyShift(nextSelection.start.row);
    nextSelection.end.row = applyShift(nextSelection.end.row);
  } else {
    nextSelection.start.col = applyShift(nextSelection.start.col);
    nextSelection.end.col = applyShift(nextSelection.end.col);
  }

  return normalizeSelection(nextSelection, sheet);
};

const shiftSelectionForDelete = (
  selection: CellRange,
  axis: Axis,
  start: number,
  count: number,
  sheet: SheetModel
): CellRange => {
  const end = start + count - 1;
  const applyShift = (value: number): number => {
    if (value < start) {
      return value;
    }

    if (value > end) {
      return value - count;
    }

    return start;
  };

  const nextSelection: CellRange = {
    start: { ...selection.start },
    end: { ...selection.end }
  };

  if (axis === "row") {
    nextSelection.start.row = applyShift(nextSelection.start.row);
    nextSelection.end.row = applyShift(nextSelection.end.row);
  } else {
    nextSelection.start.col = applyShift(nextSelection.start.col);
    nextSelection.end.col = applyShift(nextSelection.end.col);
  }

  return normalizeSelection(nextSelection, sheet);
};

export const insertAxis = (sheet: SheetModel, axis: Axis, index: number, count: number): void => {
  const nextCells: Record<string, CellModel> = {};

  for (const [key, cell] of Object.entries(sheet.cells)) {
    const address = parseCellKey(key);
    const nextRow = axis === "row" && address.row >= index ? address.row + count : address.row;
    const nextCol = axis === "column" && address.col >= index ? address.col + count : address.col;
    nextCells[getCellKey(nextRow, nextCol)] = cell;
  }

  sheet.cells = nextCells;
  sheet.merges = shiftMergesForInsert(sheet.merges, axis, index, count);
  if (axis === "row") {
    sheet.rowCount += count;
    sheet.rows = shiftIndexedSchemaForInsert(sheet.rows, index, count);
  } else {
    sheet.columnCount += count;
    sheet.columns = shiftIndexedSchemaForInsert(sheet.columns, index, count);
  }
  sheet.selection = shiftSelectionForInsert(sheet.selection, axis, index, count, sheet);
};

export const deleteAxis = (sheet: SheetModel, axis: Axis, start: number, count: number): void => {
  const end = start + count - 1;
  const nextCells: Record<string, CellModel> = {};

  for (const [key, cell] of Object.entries(sheet.cells)) {
    const address = parseCellKey(key);
    const value = axis === "row" ? address.row : address.col;

    if (value >= start && value <= end) {
      continue;
    }

    const nextRow = axis === "row" && address.row > end ? address.row - count : address.row;
    const nextCol = axis === "column" && address.col > end ? address.col - count : address.col;
    nextCells[getCellKey(nextRow, nextCol)] = cell;
  }

  sheet.cells = nextCells;
  sheet.merges = shiftMergesForDelete(sheet.merges, axis, start, count);
  if (axis === "row") {
    sheet.rowCount = Math.max(1, sheet.rowCount - count);
    sheet.rows = shiftIndexedSchemaForDelete(sheet.rows, start, count);
  } else {
    sheet.columnCount = Math.max(1, sheet.columnCount - count);
    sheet.columns = shiftIndexedSchemaForDelete(sheet.columns, start, count);
  }
  sheet.selection = shiftSelectionForDelete(sheet.selection, axis, start, count, sheet);
};