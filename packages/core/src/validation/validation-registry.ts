import type {
  CellModel,
  CellPrimitive,
  CellValidationConfig,
  CellValidationCustomRule,
  CellValidationIssue,
  CellValidationRegexRule,
  CellValidationResult,
  SafeCellValidator,
  SheetModel,
  SpreadsheetError,
  WorkbookModel
} from "../domain/types";
import { createCoreOperationError } from "../errors/spreadsheet-operation-error";

const MAX_REGEX_PATTERN_LENGTH = 120;
const ALLOWED_REGEX_FLAGS = /^[imu]*$/;

const createValidationError = (
  code: string,
  message: string,
  details?: Record<string, unknown>
): SpreadsheetError => ({
  code,
  message,
  area: "core",
  recoverable: true,
  details
});

const isEmptyValue = (value: CellPrimitive): boolean => value == null || (typeof value === "string" && value.trim() === "");

const isFormulaValue = (value: CellPrimitive): boolean => typeof value === "string" && value.startsWith("=");

const validateWorkbookLimits = (workbook: WorkbookModel, value: CellPrimitive): SpreadsheetError | undefined => {
  if (typeof value !== "string") {
    return undefined;
  }

  if (isFormulaValue(value)) {
    if (value.length > workbook.settings.maxFormulaLength) {
      return createValidationError(
        "CORE_VALIDATION_FORMULA_MAX_LENGTH",
        `A fórmula excede o limite de ${workbook.settings.maxFormulaLength} caracteres.`,
        { maxFormulaLength: workbook.settings.maxFormulaLength }
      );
    }

    return undefined;
  }

  if (value.length > workbook.settings.maxCellLength) {
    return createValidationError(
      "CORE_VALIDATION_CELL_MAX_LENGTH",
      `O valor excede o limite de ${workbook.settings.maxCellLength} caracteres.`,
      { maxCellLength: workbook.settings.maxCellLength }
    );
  }

  return undefined;
};

const parseNumber = (value: CellPrimitive): number | undefined => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string") {
    const normalized = value.trim();
    if (!normalized) {
      return undefined;
    }

    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : undefined;
  }

  return undefined;
};

const parseBoolean = (value: CellPrimitive): boolean | undefined => {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (["true", "1", "yes", "y"].includes(normalized)) {
      return true;
    }
    if (["false", "0", "no", "n"].includes(normalized)) {
      return false;
    }
  }

  return undefined;
};

const parseDate = (value: CellPrimitive): number | undefined => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string") {
    const parsed = Date.parse(value.trim());
    return Number.isFinite(parsed) ? parsed : undefined;
  }

  return undefined;
};

const buildIssue = (
  ruleType: CellValidationIssue["ruleType"],
  code: string,
  message: string,
  validator?: string
): CellValidationIssue => ({
  code,
  message,
  ruleType,
  validator
});

const shouldSkipRuleForFormula = (
  value: CellPrimitive,
  allowFormula: boolean | undefined,
  ruleType: CellValidationIssue["ruleType"]
): boolean => isFormulaValue(value) && ruleType !== "required" && allowFormula !== true;

const validateRegexRule = (value: CellPrimitive, rule: CellValidationRegexRule): CellValidationIssue | undefined => {
  if (typeof value !== "string") {
    return buildIssue("regex", "CORE_VALIDATION_REGEX", rule.message ?? "O valor deve ser um texto compatível com o padrão.");
  }

  if (rule.pattern.length > MAX_REGEX_PATTERN_LENGTH || !ALLOWED_REGEX_FLAGS.test(rule.flags ?? "")) {
    return buildIssue("regex", "CORE_VALIDATION_REGEX_UNSAFE", "A regra de regex configurada para a célula não é segura.");
  }

  const expression = new RegExp(rule.pattern, rule.flags);
  if (!expression.test(value)) {
    return buildIssue("regex", "CORE_VALIDATION_REGEX", rule.message ?? "O valor não corresponde ao padrão permitido.");
  }

  return undefined;
};

const validateCustomRule = (
  value: CellPrimitive,
  rule: CellValidationCustomRule,
  workbook: WorkbookModel,
  sheet: SheetModel,
  cell: CellModel | undefined,
  customValidators: Map<string, SafeCellValidator>,
  address: { row: number; col: number }
): CellValidationIssue | undefined => {
  const validator = customValidators.get(rule.validator);
  if (!validator) {
    return buildIssue(
      "custom",
      "CORE_VALIDATION_VALIDATOR_MISSING",
      `O validador customizado '${rule.validator}' nao foi registrado.`,
      rule.validator
    );
  }

  return (
    validator({
      workbook,
      sheet,
      cell,
      address,
      value,
      params: rule.params
    }) ?? undefined
  );
};

export class ValidationRegistry {
  private readonly customValidators = new Map<string, SafeCellValidator>();

  registerValidator(id: string, validator: SafeCellValidator): void {
    if (this.customValidators.has(id)) {
      throw createCoreOperationError("CORE_VALIDATOR_ALREADY_REGISTERED", `Validator already registered: ${id}`, {
        validatorId: id
      });
    }

    this.customValidators.set(id, validator);
  }

  unregisterValidator(id: string): void {
    this.customValidators.delete(id);
  }

  listValidators(): Array<{ id: string }> {
    return Array.from(this.customValidators.keys()).map((id) => ({ id }));
  }

  validateCellValue(input: {
    workbook: WorkbookModel;
    sheet: SheetModel;
    cell?: CellModel;
    row: number;
    col: number;
    value: CellPrimitive;
    validation?: CellValidationConfig;
  }): CellValidationResult {
    const workbookLimitError = validateWorkbookLimits(input.workbook, input.value);
    if (workbookLimitError) {
      return {
        valid: false,
        error: workbookLimitError
      };
    }

    const validation = input.validation ?? input.cell?.validation;
    if (!validation?.rules.length) {
      return { valid: true };
    }

    for (const rule of validation.rules) {
      if (shouldSkipRuleForFormula(input.value, rule.allowFormula, rule.type)) {
        continue;
      }

      let issue: CellValidationIssue | undefined;

      switch (rule.type) {
        case "required":
          if (isEmptyValue(input.value)) {
            issue = buildIssue("required", "CORE_VALIDATION_REQUIRED", rule.message ?? "Esta célula é obrigatória.");
          }
          break;
        case "text": {
          if (typeof input.value !== "string") {
            issue = buildIssue("text", "CORE_VALIDATION_TEXT", rule.message ?? "O valor deve ser um texto.");
            break;
          }

          const length = input.value.length;
          if (rule.minLength != null && length < rule.minLength) {
            issue = buildIssue("text", "CORE_VALIDATION_TEXT_MIN", rule.message ?? `O texto deve ter ao menos ${rule.minLength} caracteres.`);
          } else if (rule.maxLength != null && length > rule.maxLength) {
            issue = buildIssue("text", "CORE_VALIDATION_TEXT_MAX", rule.message ?? `O texto deve ter no máximo ${rule.maxLength} caracteres.`);
          }
          break;
        }
        case "number": {
          const numericValue = parseNumber(input.value);
          if (numericValue == null) {
            issue = buildIssue("number", "CORE_VALIDATION_NUMBER", rule.message ?? "O valor deve ser numérico.");
            break;
          }

          if (rule.min != null && numericValue < rule.min) {
            issue = buildIssue("number", "CORE_VALIDATION_NUMBER_MIN", rule.message ?? `O valor deve ser maior ou igual a ${rule.min}.`);
          } else if (rule.max != null && numericValue > rule.max) {
            issue = buildIssue("number", "CORE_VALIDATION_NUMBER_MAX", rule.message ?? `O valor deve ser menor ou igual a ${rule.max}.`);
          }
          break;
        }
        case "date": {
          const timestamp = parseDate(input.value);
          if (timestamp == null) {
            issue = buildIssue("date", "CORE_VALIDATION_DATE", rule.message ?? "O valor deve ser uma data válida.");
            break;
          }

          const min = rule.min ? Date.parse(rule.min) : undefined;
          const max = rule.max ? Date.parse(rule.max) : undefined;
          if (Number.isFinite(min) && timestamp < (min as number)) {
            issue = buildIssue("date", "CORE_VALIDATION_DATE_MIN", rule.message ?? `A data deve ser em ou após ${rule.min}.`);
          } else if (Number.isFinite(max) && timestamp > (max as number)) {
            issue = buildIssue("date", "CORE_VALIDATION_DATE_MAX", rule.message ?? `A data deve ser em ou antes de ${rule.max}.`);
          }
          break;
        }
        case "boolean":
        case "checkbox": {
          if (parseBoolean(input.value) == null) {
            issue = buildIssue(rule.type, "CORE_VALIDATION_BOOLEAN", rule.message ?? "O valor deve ser booleano.");
          }
          break;
        }
        case "list":
        case "dropdown": {
          const matches = rule.values.some((candidate) => candidate === input.value || String(candidate) === String(input.value));
          if (!matches) {
            issue = buildIssue(rule.type, "CORE_VALIDATION_LIST", rule.message ?? "O valor deve estar na lista permitida.");
          }
          break;
        }
        case "range": {
          const numericValue = parseNumber(input.value);
          if (numericValue == null || numericValue < rule.min || numericValue > rule.max) {
            issue = buildIssue("range", "CORE_VALIDATION_RANGE", rule.message ?? `O valor deve estar entre ${rule.min} e ${rule.max}.`);
          }
          break;
        }
        case "length": {
          const length = String(input.value ?? "").length;
          if (rule.min != null && length < rule.min) {
            issue = buildIssue("length", "CORE_VALIDATION_LENGTH_MIN", rule.message ?? `O valor deve ter ao menos ${rule.min} caracteres.`);
          } else if (rule.max != null && length > rule.max) {
            issue = buildIssue("length", "CORE_VALIDATION_LENGTH_MAX", rule.message ?? `O valor deve ter no máximo ${rule.max} caracteres.`);
          }
          break;
        }
        case "regex":
          issue = validateRegexRule(input.value, rule);
          break;
        case "custom":
          issue = validateCustomRule(
            input.value,
            rule,
            input.workbook,
            input.sheet,
            input.cell,
            this.customValidators,
            { row: input.row, col: input.col }
          );
          break;
      }

      if (issue) {
        return {
          valid: false,
          issue,
          error: createValidationError(issue.code, issue.message, {
            row: input.row,
            col: input.col,
            ruleType: issue.ruleType,
            validator: issue.validator
          })
        };
      }
    }

    return { valid: true };
  }
}