import { cellLabelToAddress } from "@excelsior/core";
import type {
  CellPrimitive,
  FormulaEngine,
  FormulaEvaluationContext,
  FormulaEvaluationResult,
  FormulaReference,
  SpreadsheetError
} from "@excelsior/core";

type Token =
  | { type: "number"; value: number }
  | { type: "cell"; value: string }
  | { type: "identifier"; value: string }
  | { type: "string"; value: string }
  | { type: "operator"; value: string }
  | { type: "paren"; value: "(" | ")" }
  | { type: "comma" }
  | { type: "colon" }
  | { type: "bang" }
  | { type: "eof" };

type AstNode =
  | { type: "number"; value: number }
  | { type: "cell"; value: string; sheet?: string }
  | { type: "range"; start: string; end: string; sheet?: string }
  | { type: "binary"; operator: string; left: AstNode; right: AstNode }
  | { type: "unary"; operator: string; operand: AstNode }
  | { type: "function"; name: string; args: AstNode[] };

const createFormulaError = (code: string, message: string): SpreadsheetError => ({
  code,
  message,
  area: "formula",
  recoverable: true
});

const isCellReference = (value: string): boolean => /^\$?[A-Za-z]+\$?\d+$/.test(value);

const normalizeCellReference = (value: string): string => value.replace(/\$/g, "").toUpperCase();

const tokenize = (expression: string): Token[] => {
  const source = expression.startsWith("=") ? expression.slice(1) : expression;
  const tokens: Token[] = [];
  let index = 0;

  while (index < source.length) {
    const char = source[index];

    if (!char) {
      break;
    }

    if (/\s/.test(char)) {
      index += 1;
      continue;
    }

    if (/[0-9.]/.test(char)) {
      let value = char;
      index += 1;
      while (index < source.length && /[0-9.]/.test(source[index] ?? "")) {
        value += source[index];
        index += 1;
      }
      tokens.push({ type: "number", value: Number(value) });
      continue;
    }

    if (char === "$" || /[A-Za-z_]/.test(char)) {
      let value = char;
      index += 1;
      while (index < source.length && /[$A-Za-z0-9_]/.test(source[index] ?? "")) {
        value += source[index];
        index += 1;
      }

      if (isCellReference(value)) {
        tokens.push({ type: "cell", value: value.toUpperCase() });
      } else {
        tokens.push({ type: "identifier", value });
      }
      continue;
    }

    if (char === "'") {
      let value = "";
      index += 1;
      while (index < source.length && source[index] !== "'") {
        value += source[index];
        index += 1;
      }
      if (source[index] !== "'") {
        throw createFormulaError("FORMULA_TOKEN_INVALID", "Unterminated sheet name string.");
      }
      index += 1;
      tokens.push({ type: "string", value });
      continue;
    }

    if ("+-*/^".includes(char)) {
      tokens.push({ type: "operator", value: char });
      index += 1;
      continue;
    }

    if (char === "(") {
      tokens.push({ type: "paren", value: "(" });
      index += 1;
      continue;
    }

    if (char === ")") {
      tokens.push({ type: "paren", value: ")" });
      index += 1;
      continue;
    }

    if (char === ",") {
      tokens.push({ type: "comma" });
      index += 1;
      continue;
    }

    if (char === ":") {
      tokens.push({ type: "colon" });
      index += 1;
      continue;
    }

    if (char === "!") {
      tokens.push({ type: "bang" });
      index += 1;
      continue;
    }

    throw createFormulaError("FORMULA_TOKEN_INVALID", `Unsupported token: ${char}`);
  }

  tokens.push({ type: "eof" });
  return tokens;
};

class Parser {
  constructor(private readonly tokens: Token[], private position = 0) {}

  private isOperator(value?: string): boolean {
    const token = this.current();
    return token.type === "operator" && (value ? token.value === value : true);
  }

  private isClosingParen(): boolean {
    const token = this.current();
    return token.type === "paren" && token.value === ")";
  }

  parse(): AstNode {
    const node = this.parseExpression();
    this.expect("eof");
    return node;
  }

  private current(): Token {
    return this.tokens[this.position] ?? { type: "eof" };
  }

  private advance(): Token {
    const token = this.current();
    this.position += 1;
    return token;
  }

  private expect(type: Token["type"]): Token {
    const token = this.current();
    if (token.type !== type) {
      throw createFormulaError("FORMULA_PARSE_INVALID", `Expected ${type} but received ${token.type}.`);
    }
    return this.advance();
  }

  private parseExpression(): AstNode {
    let node = this.parseTerm();

    while (this.isOperator("+") || this.isOperator("-")) {
      const operator = this.advance();
      if (operator.type !== "operator") {
        throw createFormulaError("FORMULA_PARSE_INVALID", "Expected expression operator.");
      }
      node = {
        type: "binary",
        operator: operator.value,
        left: node,
        right: this.parseTerm()
      };
    }

    return node;
  }

  private parseTerm(): AstNode {
    let node = this.parseFactor();

    while (this.isOperator("*") || this.isOperator("/")) {
      const operator = this.advance();
      if (operator.type !== "operator") {
        throw createFormulaError("FORMULA_PARSE_INVALID", "Expected term operator.");
      }
      node = {
        type: "binary",
        operator: operator.value,
        left: node,
        right: this.parseFactor()
      };
    }

    return node;
  }

  private parseFactor(): AstNode {
    if (this.isOperator("+") || this.isOperator("-")) {
      const operator = this.advance();
      if (operator.type !== "operator") {
        throw createFormulaError("FORMULA_PARSE_INVALID", "Expected unary operator.");
      }
      return {
        type: "unary",
        operator: operator.value,
        operand: this.parseFactor()
      };
    }

    let node = this.parsePrimary();

    while (this.isOperator("^")) {
      const operator = this.advance();
      if (operator.type !== "operator") {
        throw createFormulaError("FORMULA_PARSE_INVALID", "Expected power operator.");
      }
      node = {
        type: "binary",
        operator: operator.value,
        left: node,
        right: this.parsePrimary()
      };
    }

    return node;
  }

  private parsePrimary(): AstNode {
    const token = this.current();

    if (
      (token.type === "identifier" || token.type === "string") &&
      this.tokens[this.position + 1]?.type === "bang"
    ) {
      const sheet = token.value;
      this.advance();
      this.expect("bang");
      const cellToken = this.expect("cell");
      if (cellToken.type !== "cell") {
        throw createFormulaError("FORMULA_PARSE_INVALID", "Expected cell after sheet reference.");
      }
      if (this.current().type === "colon") {
        this.advance();
        const next = this.expect("cell");
        if (next.type !== "cell") {
          throw createFormulaError("FORMULA_PARSE_INVALID", "Expected end of range after sheet reference.");
        }
        return { type: "range", start: cellToken.value, end: next.value, sheet };
      }
      return { type: "cell", value: cellToken.value, sheet };
    }

    if (token.type === "number") {
      this.advance();
      return { type: "number", value: token.value };
    }

    if (token.type === "cell") {
      this.advance();
      if (this.current().type === "colon") {
        this.advance();
        const next = this.expect("cell");
        if (next.type !== "cell") {
          throw createFormulaError("FORMULA_PARSE_INVALID", "Expected cell reference after range separator.");
        }
        return { type: "range", start: token.value, end: next.value };
      }
      return { type: "cell", value: token.value };
    }

    if (token.type === "identifier") {
      this.advance();
      this.expect("paren");
      const args: AstNode[] = [];

      if (!this.isClosingParen()) {
        do {
          args.push(this.parseExpression());
          if (this.current().type !== "comma") {
            break;
          }
          this.advance();
        } while (true);
      }

      const closingParen = this.expect("paren");
      if (closingParen.type !== "paren" || closingParen.value !== ")") {
        throw createFormulaError("FORMULA_PARSE_INVALID", "Function call was not closed.");
      }

      return {
        type: "function",
        name: token.value.toUpperCase(),
        args
      };
    }

    if (token.type === "paren" && token.value === "(") {
      this.advance();
      const node = this.parseExpression();
      const closingParen = this.expect("paren");
      if (closingParen.type !== "paren" || closingParen.value !== ")") {
        throw createFormulaError("FORMULA_PARSE_INVALID", "Expression group was not closed.");
      }
      return node;
    }

    throw createFormulaError("FORMULA_PARSE_INVALID", `Unexpected token ${token.type}.`);
  }
}

const expandRange = (start: string, end: string): string[] => {
  const startAddress = cellLabelToAddress(start);
  const endAddress = cellLabelToAddress(end);
  const rowStart = Math.min(startAddress.row, endAddress.row);
  const rowEnd = Math.max(startAddress.row, endAddress.row);
  const colStart = Math.min(startAddress.col, endAddress.col);
  const colEnd = Math.max(startAddress.col, endAddress.col);
  const cells: string[] = [];

  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      const columnLabel = (() => {
        let current = col + 1;
        let label = "";
        while (current > 0) {
          const remainder = (current - 1) % 26;
          label = String.fromCharCode(65 + remainder) + label;
          current = Math.floor((current - 1) / 26);
        }
        return label;
      })();
      cells.push(`${columnLabel}${row + 1}`);
    }
  }

  return cells;
};

const toNumber = (value: CellPrimitive): number => {
  if (typeof value === "number") {
    return value;
  }

  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }

  if (typeof value === "string") {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  return 0;
};

const evaluateNode = (
  node: AstNode,
  context: FormulaEvaluationContext
): FormulaEvaluationResult & { values?: number[] } => {
  switch (node.type) {
    case "number":
      return { value: node.value, values: [node.value] };
    case "cell": {
      const address = cellLabelToAddress(normalizeCellReference(node.value));
      const result = context.evaluateCell(address.row, address.col, undefined, node.sheet);
      return {
        value: result.value,
        error: result.error,
        values: [toNumber(result.value)]
      };
    }
    case "range": {
      const values = expandRange(normalizeCellReference(node.start), normalizeCellReference(node.end)).map((label) => {
        const address = cellLabelToAddress(label);
        const result = context.evaluateCell(address.row, address.col, undefined, node.sheet);
        return toNumber(result.value);
      });
      return { value: values[0] ?? 0, values };
    }
    case "unary": {
      const result = evaluateNode(node.operand, context);
      const operandValue = toNumber(result.value);
      return {
        value: node.operator === "-" ? -operandValue : operandValue,
        error: result.error,
        values: [node.operator === "-" ? -operandValue : operandValue]
      };
    }
    case "binary": {
      const left = evaluateNode(node.left, context);
      const right = evaluateNode(node.right, context);
      if (left.error) {
        return left;
      }
      if (right.error) {
        return right;
      }
      const leftValue = toNumber(left.value);
      const rightValue = toNumber(right.value);
      const operations: Record<string, number> = {
        "+": leftValue + rightValue,
        "-": leftValue - rightValue,
        "*": leftValue * rightValue,
        "/": rightValue === 0 ? Number.NaN : leftValue / rightValue,
        "^": leftValue ** rightValue
      };
      return {
        value: Number.isNaN(operations[node.operator])
          ? null
          : operations[node.operator],
        error: Number.isNaN(operations[node.operator])
          ? createFormulaError("FORMULA_DIV_ZERO", "Division by zero.")
          : undefined,
        values: [operations[node.operator]]
      };
    }
    case "function": {
      const values = node.args.flatMap((argument) => {
        const result = evaluateNode(argument, context);
        if (argument.type === "range") {
          return result.values ?? [toNumber(result.value)];
        }
        return [toNumber(result.value)];
      });
      const reducerMap: Record<string, () => number> = {
        SUM: () => values.reduce((total, value) => total + value, 0),
        MIN: () => Math.min(...values),
        MAX: () => Math.max(...values),
        AVERAGE: () => (values.length ? values.reduce((total, value) => total + value, 0) / values.length : 0),
        COUNT: () => values.filter((value) => !Number.isNaN(value)).length,
        COUNTA: () => values.filter((value) => value !== 0).length,
        ABS: () => Math.abs(values[0] ?? 0),
        ROUND: () => {
          const base = values[0] ?? 0;
          const precision = values[1] ?? 0;
          const factor = 10 ** precision;
          return Math.round(base * factor) / factor;
        },
        IF: () => ((values[0] ?? 0) !== 0 ? values[1] ?? 0 : values[2] ?? 0),
        AND: () => (values.every((value) => value !== 0) ? 1 : 0),
        OR: () => (values.some((value) => value !== 0) ? 1 : 0),
        NOT: () => ((values[0] ?? 0) === 0 ? 1 : 0)
      };
      const reducer = reducerMap[node.name];
      if (!reducer) {
        return {
          value: null,
          error: createFormulaError("FORMULA_FUNCTION_UNSUPPORTED", `Unsupported function ${node.name}.`)
        };
      }
      return {
        value: reducer(),
        values
      };
    }
  }
};

const collectReferencesFromNode = (node: AstNode, references: FormulaReference[]): void => {
  switch (node.type) {
    case "number":
      return;
    case "cell": {
      const address = cellLabelToAddress(normalizeCellReference(node.value));
      references.push({
        row: address.row,
        col: address.col,
        sheetRef: node.sheet
      });
      return;
    }
    case "range": {
      for (const label of expandRange(normalizeCellReference(node.start), normalizeCellReference(node.end))) {
        const address = cellLabelToAddress(label);
        references.push({
          row: address.row,
          col: address.col,
          sheetRef: node.sheet
        });
      }
      return;
    }
    case "unary":
      collectReferencesFromNode(node.operand, references);
      return;
    case "binary":
      collectReferencesFromNode(node.left, references);
      collectReferencesFromNode(node.right, references);
      return;
    case "function":
      for (const argument of node.args) {
        collectReferencesFromNode(argument, references);
      }
      return;
  }
};

export class BasicFormulaEngine implements FormulaEngine {
  evaluate(expression: string, context: FormulaEvaluationContext): FormulaEvaluationResult {
    try {
      const tokens = tokenize(expression);
      const ast = new Parser(tokens).parse();
      const result = evaluateNode(ast, context);
      return {
        value: result.value,
        error: result.error
      };
    } catch (error) {
      if (typeof error === "object" && error && "code" in error) {
        return {
          value: null,
          error: error as SpreadsheetError
        };
      }
      return {
        value: null,
        error: createFormulaError("FORMULA_RUNTIME_INVALID", "Formula evaluation failed.")
      };
    }
  }

  collectReferences(expression: string): FormulaReference[] {
    try {
      const tokens = tokenize(expression);
      const ast = new Parser(tokens).parse();
      const references: FormulaReference[] = [];
      collectReferencesFromNode(ast, references);
      const unique = new Map<string, FormulaReference>();
      for (const reference of references) {
        unique.set(`${reference.sheetRef ?? ""}:${reference.row}:${reference.col}`, reference);
      }
      return Array.from(unique.values());
    } catch {
      return [];
    }
  }
}