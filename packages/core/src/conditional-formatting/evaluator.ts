import type {
  CellPrimitive,
  CellRange,
  CellStyle,
  ConditionalFormattingColorScaleRule,
  ConditionalFormattingComparisonRule,
  ConditionalFormattingContainsTextRule,
  ConditionalFormattingDateRule,
  ConditionalFormattingDuplicatesRule,
  ConditionalFormattingEqualityRule,
  ConditionalFormattingFormulaRule,
  ConditionalFormattingRule,
  FormulaEngine,
  SheetModel,
  WorkbookModel
} from "../domain/types";
import { getCellKey } from "../utils/cell-key";

const isWithinRange = (row: number, col: number, range: CellRange): boolean =>
  row >= range.start.row &&
  row <= range.end.row &&
  col >= range.start.col &&
  col <= range.end.col;

const normalizeStyle = (style?: Partial<CellStyle>): CellStyle | undefined => {
  if (!style) {
    return undefined;
  }

  const nextStyle: CellStyle = {
    ...style,
    border: style.border
      ? {
          ...style.border
        }
      : undefined
  };

  if (!nextStyle.border?.top && !nextStyle.border?.right && !nextStyle.border?.bottom && !nextStyle.border?.left) {
    delete nextStyle.border;
  }

  return Object.values(nextStyle).some((value) => value !== undefined) ? nextStyle : undefined;
};

const mergeStyles = (...styles: Array<Partial<CellStyle> | undefined>): CellStyle | undefined => {
  const merged = styles.reduce<Partial<CellStyle>>(
    (current, style) =>
      style
        ? {
            ...current,
            ...style,
            border: {
              ...current.border,
              ...style.border
            }
          }
        : current,
    {}
  );

  return normalizeStyle(merged);
};

const getSheet = (workbook: WorkbookModel, sheetId: string): SheetModel | undefined =>
  workbook.sheets.find((sheet) => sheet.id === sheetId);

const resolveSheet = (workbook: WorkbookModel, sheet: SheetModel, sheetRef?: string): SheetModel | undefined =>
  workbook.sheets.find((item) => item.id === (sheetRef ?? sheet.id) || item.name === sheetRef);

const getCellEffectiveValue = (sheet: SheetModel, row: number, col: number): CellPrimitive => {
  const cell = sheet.cells[getCellKey(row, col)];
  if (!cell) {
    return null;
  }

  return cell.formula ? cell.computedValue ?? null : cell.value;
};

const valuesEqual = (left: CellPrimitive, right: CellPrimitive): boolean => {
  if (left === right) {
    return true;
  }

  return String(left ?? "") === String(right ?? "");
};

const toNumber = (value: CellPrimitive): number | undefined => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }

  if (typeof value === "string") {
    const parsed = Number(value.trim());
    return Number.isFinite(parsed) ? parsed : undefined;
  }

  return undefined;
};

const toTimestamp = (value: CellPrimitive): number | undefined => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string") {
    const parsed = Date.parse(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }

  return undefined;
};

const toComparableText = (value: CellPrimitive): string => String(value ?? "").toLowerCase();

const getRangeValues = (sheet: SheetModel, range: CellRange): CellPrimitive[] => {
  const values: CellPrimitive[] = [];
  for (let row = range.start.row; row <= range.end.row; row += 1) {
    for (let col = range.start.col; col <= range.end.col; col += 1) {
      values.push(getCellEffectiveValue(sheet, row, col));
    }
  }
  return values;
};

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

const parseHexColor = (value: string): [number, number, number] | undefined => {
  const normalized = value.trim();
  if (!normalized.startsWith("#")) {
    return undefined;
  }

  const hex = normalized.slice(1);
  if (hex.length === 3) {
    const expanded = hex
      .split("")
      .map((char) => `${char}${char}`)
      .join("");
    const parsed = Number.parseInt(expanded, 16);
    if (Number.isNaN(parsed)) {
      return undefined;
    }
    return [(parsed >> 16) & 255, (parsed >> 8) & 255, parsed & 255];
  }

  if (hex.length === 6) {
    const parsed = Number.parseInt(hex, 16);
    if (Number.isNaN(parsed)) {
      return undefined;
    }
    return [(parsed >> 16) & 255, (parsed >> 8) & 255, parsed & 255];
  }

  return undefined;
};

const toHexColor = ([red, green, blue]: [number, number, number]): string =>
  `#${[red, green, blue]
    .map((channel) => clamp(Math.round(channel), 0, 255).toString(16).padStart(2, "0"))
    .join("")}`;

const interpolateColor = (start: string, end: string, ratio: number): string | undefined => {
  const startColor = parseHexColor(start);
  const endColor = parseHexColor(end);
  if (!startColor || !endColor) {
    return undefined;
  }

  return toHexColor([
    startColor[0] + (endColor[0] - startColor[0]) * ratio,
    startColor[1] + (endColor[1] - startColor[1]) * ratio,
    startColor[2] + (endColor[2] - startColor[2]) * ratio
  ] as [number, number, number]);
};

const evaluateFormulaRule = (
  workbook: WorkbookModel,
  sheet: SheetModel,
  row: number,
  col: number,
  formula: string,
  formulaEngine?: FormulaEngine
): boolean => {
  if (!formulaEngine) {
    return false;
  }

  const result = formulaEngine.evaluate(formula, {
    currentCell: { row, col },
    currentSheetId: sheet.id,
    currentSheetName: sheet.name,
    getCell: (targetRow, targetCol, nextSheetRef) => {
      const targetSheet = resolveSheet(workbook, sheet, nextSheetRef);
      return targetSheet?.cells[getCellKey(targetRow, targetCol)];
    },
    evaluateCell: (targetRow, targetCol, _trail, nextSheetRef) => {
      const targetSheet = resolveSheet(workbook, sheet, nextSheetRef);
      const targetCell = targetSheet?.cells[getCellKey(targetRow, targetCol)];
      return {
        value: targetCell?.formula ? targetCell.computedValue ?? null : targetCell?.value ?? null,
        error: targetCell?.error
      };
    }
  });

  return !result.error && (toNumber(result.value) ?? 0) !== 0;
};

const evaluateColorScaleRule = (
  sheet: SheetModel,
  row: number,
  col: number,
  rule: ConditionalFormattingColorScaleRule
): CellStyle | undefined => {
  const numericValues = getRangeValues(sheet, rule.range)
    .map((value) => toNumber(value))
    .filter((value): value is number => value != null);

  const currentValue = toNumber(getCellEffectiveValue(sheet, row, col));
  if (currentValue == null || numericValues.length === 0) {
    return undefined;
  }

  const min = Math.min(...numericValues);
  const max = Math.max(...numericValues);
  const ratio = min === max ? 1 : clamp((currentValue - min) / (max - min), 0, 1);
  const backgroundColor = interpolateColor(rule.minColor, rule.maxColor, ratio);

  return normalizeStyle({
    backgroundColor,
    textColor: rule.textColor
  });
};

const evaluateComparisonRule = (value: CellPrimitive, rule: ConditionalFormattingComparisonRule): CellStyle | undefined => {
  const numericValue = toNumber(value);
  if (numericValue == null) {
    return undefined;
  }

  const matches = rule.type === "greaterThan" ? numericValue > rule.value : numericValue < rule.value;
  return matches ? normalizeStyle(rule.style) : undefined;
};

const evaluateEqualityRule = (value: CellPrimitive, rule: ConditionalFormattingEqualityRule): CellStyle | undefined => {
  const matches = valuesEqual(value, rule.value);
  if (rule.type === "equal") {
    return matches ? normalizeStyle(rule.style) : undefined;
  }

  if (matches) {
    return undefined;
  }

  return normalizeStyle(rule.style);
};

const evaluateContainsTextRule = (
  value: CellPrimitive,
  rule: ConditionalFormattingContainsTextRule
): CellStyle | undefined =>
  toComparableText(value).includes(rule.text.toLowerCase()) ? normalizeStyle(rule.style) : undefined;

const evaluateDateRule = (value: CellPrimitive, rule: ConditionalFormattingDateRule): CellStyle | undefined => {
  const timestamp = toTimestamp(value);
  const limit = Date.parse(rule.date);
  if (timestamp == null || !Number.isFinite(limit)) {
    return undefined;
  }

  const matches = rule.type === "dateBefore" ? timestamp < limit : timestamp > limit;
  return matches ? normalizeStyle(rule.style) : undefined;
};

const evaluateDuplicatesRule = (
  sheet: SheetModel,
  value: CellPrimitive,
  rule: ConditionalFormattingDuplicatesRule
): CellStyle | undefined => {
  const normalized = String(value ?? "");
  if (!normalized) {
    return undefined;
  }

  const matches = getRangeValues(sheet, rule.range).filter((candidate) => String(candidate ?? "") === normalized);
  return matches.length > 1 ? normalizeStyle(rule.style) : undefined;
};

const evaluateFormulaFormattingRule = (
  workbook: WorkbookModel,
  sheet: SheetModel,
  row: number,
  col: number,
  rule: ConditionalFormattingFormulaRule,
  formulaEngine?: FormulaEngine
): CellStyle | undefined =>
  evaluateFormulaRule(workbook, sheet, row, col, rule.formula, formulaEngine) ? normalizeStyle(rule.style) : undefined;

const evaluateRuleStyle = (
  workbook: WorkbookModel,
  sheet: SheetModel,
  row: number,
  col: number,
  rule: ConditionalFormattingRule,
  formulaEngine?: FormulaEngine
): CellStyle | undefined => {
  if (!isWithinRange(row, col, rule.range)) {
    return undefined;
  }

  const value = getCellEffectiveValue(sheet, row, col);

  switch (rule.type) {
    case "greaterThan":
    case "lessThan":
      return evaluateComparisonRule(value, rule);
    case "equal":
    case "notEqual":
      return evaluateEqualityRule(value, rule);
    case "between": {
      const numericValue = toNumber(value);
      return numericValue != null && numericValue >= rule.min && numericValue <= rule.max
        ? normalizeStyle(rule.style)
        : undefined;
    }
    case "containsText":
      return evaluateContainsTextRule(value, rule);
    case "dateBefore":
    case "dateAfter":
      return evaluateDateRule(value, rule);
    case "duplicates":
      return evaluateDuplicatesRule(sheet, value, rule);
    case "colorScale":
      return evaluateColorScaleRule(sheet, row, col, rule);
    case "formula":
      return evaluateFormulaFormattingRule(workbook, sheet, row, col, rule, formulaEngine);
  }
};

export const getConditionalFormattingStyle = (
  workbook: WorkbookModel,
  sheetId: string,
  row: number,
  col: number,
  formulaEngine?: FormulaEngine
): CellStyle | undefined => {
  const sheet = getSheet(workbook, sheetId);
  if (!sheet?.conditionalFormats?.length) {
    return undefined;
  }

  const orderedRules = sheet.conditionalFormats
    .map((rule, index) => ({ rule, index }))
    .sort((left, right) => {
      const leftPriority = left.rule.priority ?? 100;
      const rightPriority = right.rule.priority ?? 100;
      if (leftPriority === rightPriority) {
        return left.index - right.index;
      }
      return rightPriority - leftPriority;
    })
    .map((entry) => entry.rule);

  return mergeStyles(
    ...orderedRules.map((rule) => evaluateRuleStyle(workbook, sheet, row, col, rule, formulaEngine))
  );
};