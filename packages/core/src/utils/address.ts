import type { CellAddress } from "../domain/types";
import { createCoreOperationError } from "../errors/spreadsheet-operation-error";

export const columnIndexToLabel = (columnIndex: number): string => {
  let current = columnIndex + 1;
  let label = "";

  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }

  return label;
};

export const cellAddressToLabel = ({ row, col }: CellAddress): string =>
  `${columnIndexToLabel(col)}${row + 1}`;

export const cellLabelToAddress = (label: string): CellAddress => {
  const match = /^(\$?[A-Z]+)(\$?\d+)$/i.exec(label.trim());

  if (!match) {
    throw createCoreOperationError("CORE_ADDRESS_INVALID_LABEL", `Invalid cell label: ${label}`, { label });
  }

  const [, rawColumnPart, rawRowPart] = match;
  const columnPart = rawColumnPart.replace(/\$/g, "");
  const rowPart = rawRowPart.replace(/\$/g, "");
  let columnIndex = 0;

  for (const char of columnPart.toUpperCase()) {
    columnIndex = columnIndex * 26 + (char.charCodeAt(0) - 64);
  }

  return {
    row: Number(rowPart) - 1,
    col: columnIndex - 1
  };
};