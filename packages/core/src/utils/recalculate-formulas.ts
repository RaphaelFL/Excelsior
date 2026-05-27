import type {
  CellModel,
  FormulaEngine,
  SpreadsheetError,
  SpreadsheetOperation,
  WorkbookModel
} from "../domain/types";
import { getCellKey } from "./cell-key";

const createFormulaError = (message: string): SpreadsheetError => ({
  code: "FORMULA_INVALID",
  message,
  area: "formula",
  recoverable: true
});

const createRecalcLimitError = (limit: number, count: number): SpreadsheetError => ({
  code: "FORMULA_RECALC_LIMIT_EXCEEDED",
  message: `Formula recalculation limit exceeded: ${count} cells requested, limit is ${limit}.`,
  area: "formula",
  recoverable: true
});

const getFormulaTrailKey = (sheetId: string, row: number, col: number): string =>
  `${sheetId}:${row}:${col}`;

const getRecalcLimit = (workbook: WorkbookModel): number => workbook.settings.maxRecalcCells ?? Number.POSITIVE_INFINITY;

const applyFormulaErrors = (workbook: WorkbookModel, formulaKeys: Iterable<string>, error: SpreadsheetError): WorkbookModel => {
  for (const key of formulaKeys) {
    const [sheetId, rowText, colText] = key.split(":");
    const sheet = workbook.sheets.find((item) => item.id === sheetId);
    if (!sheet) {
      continue;
    }

    const cell = sheet.cells[getCellKey(Number(rowText), Number(colText))];
    if (!cell?.formula) {
      continue;
    }

    cell.computedValue = null;
    cell.error = error;
  }

  return workbook;
};

export interface FormulaRecalculationTarget {
  sheetId: string;
  row: number;
  col: number;
}

const parseCellOperationKey = (value: string): { row: number; col: number } | undefined => {
  const match = /^(\d+):(\d+)$/.exec(value);
  if (!match) {
    return undefined;
  }

  return {
    row: Number(match[1]),
    col: Number(match[2])
  };
};

export const deriveFormulaRecalculationTargets = (
  operations: SpreadsheetOperation[]
): FormulaRecalculationTarget[] | undefined => {
  const targets = new Map<string, FormulaRecalculationTarget>();

  for (const operation of operations) {
    if (!["add", "replace", "remove"].includes(operation.op)) {
      return undefined;
    }

    if (operation.path[0] !== "cells" || typeof operation.path[1] !== "string") {
      return undefined;
    }

    const address = parseCellOperationKey(operation.path[1]);
    if (!address) {
      return undefined;
    }

    targets.set(`${operation.id}:${address.row}:${address.col}`, {
      sheetId: operation.id,
      row: address.row,
      col: address.col
    });
  }

  return Array.from(targets.values());
};

const resolveSheetRef = (workbook: WorkbookModel, currentSheetId: string, sheetRef?: string) => {
  if (!sheetRef) {
    return workbook.sheets.find((item) => item.id === currentSheetId);
  }

  return (
    workbook.sheets.find((item) => item.id === sheetRef) ??
    workbook.sheets.find((item) => item.name === sheetRef) ??
    workbook.sheets.find((item) => item.name.toLowerCase() === sheetRef.toLowerCase())
  );
};

const createFormulaEvaluator = (workbook: WorkbookModel, formulaEngine: FormulaEngine) => {
  const evaluated = new Set<string>();

  const evaluateAt = (sheetId: string, row: number, col: number, trail: string[]): CellModel | undefined => {
    const targetSheet = resolveSheetRef(workbook, sheetId, sheetId);
    if (!targetSheet) {
      return undefined;
    }

    const addressKey = getCellKey(row, col);
    const cell = targetSheet.cells[addressKey];
    if (!cell) {
      return undefined;
    }

    if (!cell.formula) {
      cell.computedValue = cell.value;
      cell.error = undefined;
      return cell;
    }

    const formulaKey = getFormulaTrailKey(targetSheet.id, row, col);
    if (evaluated.has(formulaKey)) {
      return cell;
    }

    const result = formulaEngine.evaluate(cell.formula, {
      currentCell: { row, col },
      currentSheetId: targetSheet.id,
      currentSheetName: targetSheet.name,
      getCell: (targetRow, targetCol, nextSheetRef) => {
        const resolvedSheet = resolveSheetRef(workbook, targetSheet.id, nextSheetRef ?? targetSheet.id);
        return resolvedSheet?.cells[getCellKey(targetRow, targetCol)];
      },
      evaluateCell: (targetRow, targetCol, nextTrail, nextSheetRef) => {
        const resolvedSheet = resolveSheetRef(workbook, targetSheet.id, nextSheetRef ?? targetSheet.id);
        if (!resolvedSheet) {
          return {
            value: null,
            error: createFormulaError("Referenced sheet was not found.")
          };
        }

        const currentTrail = nextTrail ?? trail;
        const targetKey = getFormulaTrailKey(resolvedSheet.id, targetRow, targetCol);
        if (currentTrail.includes(targetKey)) {
          return {
            value: null,
            error: createFormulaError("Circular reference detected.")
          };
        }

        const targetCell = evaluateAt(resolvedSheet.id, targetRow, targetCol, [...currentTrail, targetKey]);
        return {
          value: targetCell?.formula ? targetCell.computedValue ?? null : targetCell?.value ?? null,
          error: targetCell?.error
        };
      }
    });

    cell.computedValue = result.value;
    cell.error = result.error;
    evaluated.add(formulaKey);
    return cell;
  };

  return { evaluateAt };
};

const recalculateAllWorkbookFormulas = (workbook: WorkbookModel, formulaEngine: FormulaEngine): WorkbookModel => {
  const formulaKeys: string[] = [];

  for (const sheet of workbook.sheets) {
    for (const [cellKey, cell] of Object.entries(sheet.cells)) {
      if (!cell.formula) {
        cell.computedValue = cell.value;
        cell.error = undefined;
        continue;
      }

      const [rowText, colText] = cellKey.split(":");
      formulaKeys.push(getFormulaTrailKey(sheet.id, Number(rowText), Number(colText)));
    }
  }

  const recalcLimit = getRecalcLimit(workbook);
  if (formulaKeys.length > recalcLimit) {
    return applyFormulaErrors(workbook, formulaKeys, createRecalcLimitError(recalcLimit, formulaKeys.length));
  }

  const { evaluateAt } = createFormulaEvaluator(workbook, formulaEngine);

  for (const sheet of workbook.sheets) {
    for (const [cellKey, cell] of Object.entries(sheet.cells)) {
      if (!cell.formula) {
        cell.computedValue = cell.value;
        cell.error = undefined;
        continue;
      }

      const [rowText, colText] = cellKey.split(":");
      const row = Number(rowText);
      const col = Number(colText);
      evaluateAt(sheet.id, row, col, [getFormulaTrailKey(sheet.id, row, col)]);
    }
  }

  return workbook;
};

const buildDependentsGraph = (workbook: WorkbookModel, formulaEngine: FormulaEngine) => {
  if (!formulaEngine.collectReferences) {
    return undefined;
  }

  const formulaCells = new Set<string>();
  const dependents = new Map<string, Set<string>>();

  for (const sheet of workbook.sheets) {
    for (const [cellKey, cell] of Object.entries(sheet.cells)) {
      if (!cell.formula) {
        continue;
      }

      const [rowText, colText] = cellKey.split(":");
      const row = Number(rowText);
      const col = Number(colText);
      const formulaKey = getFormulaTrailKey(sheet.id, row, col);
      formulaCells.add(formulaKey);

      for (const reference of formulaEngine.collectReferences(cell.formula)) {
        const resolvedSheet = resolveSheetRef(workbook, sheet.id, reference.sheetRef);
        const dependencyKey = getFormulaTrailKey(resolvedSheet?.id ?? sheet.id, reference.row, reference.col);
        const current = dependents.get(dependencyKey) ?? new Set<string>();
        current.add(formulaKey);
        dependents.set(dependencyKey, current);
      }
    }
  }

  return { dependents, formulaCells };
};

const recalculateImpactedFormulas = (
  workbook: WorkbookModel,
  formulaEngine: FormulaEngine,
  dirtyCells: FormulaRecalculationTarget[]
): WorkbookModel => {
  const graph = buildDependentsGraph(workbook, formulaEngine);
  if (!graph) {
    return recalculateAllWorkbookFormulas(workbook, formulaEngine);
  }

  const impactedFormulaKeys = new Set<string>();
  const queue = dirtyCells.map((cell) => getFormulaTrailKey(cell.sheetId, cell.row, cell.col));

  while (queue.length > 0) {
    const current = queue.shift();
    if (!current) {
      continue;
    }

    if (graph.formulaCells.has(current)) {
      impactedFormulaKeys.add(current);
    }

    for (const dependent of graph.dependents.get(current) ?? []) {
      if (impactedFormulaKeys.has(dependent)) {
        continue;
      }

      impactedFormulaKeys.add(dependent);
      queue.push(dependent);
    }
  }

  if (impactedFormulaKeys.size === 0) {
    return workbook;
  }

  const recalcLimit = getRecalcLimit(workbook);
  if (impactedFormulaKeys.size > recalcLimit) {
    return applyFormulaErrors(
      workbook,
      impactedFormulaKeys,
      createRecalcLimitError(recalcLimit, impactedFormulaKeys.size)
    );
  }

  const { evaluateAt } = createFormulaEvaluator(workbook, formulaEngine);
  for (const key of impactedFormulaKeys) {
    const [sheetId, rowText, colText] = key.split(":");
    evaluateAt(sheetId, Number(rowText), Number(colText), [key]);
  }

  return workbook;
};

export const recalculateWorkbookFormulas = (
  workbook: WorkbookModel,
  formulaEngine?: FormulaEngine,
  dirtyCells?: FormulaRecalculationTarget[]
): WorkbookModel => {
  if (!formulaEngine) {
    return workbook;
  }

  if (!dirtyCells?.length) {
    return recalculateAllWorkbookFormulas(workbook, formulaEngine);
  }

  return recalculateImpactedFormulas(workbook, formulaEngine, dirtyCells);
};