import type { CellModel, CellPrimitive, WorkbookModel } from "@excelsior/core";

export type FindReplaceScope = "sheet" | "workbook";

export interface FindReplaceOptions {
  query: string;
  replaceText: string;
  scope: FindReplaceScope;
  caseSensitive: boolean;
  wholeCell: boolean;
  regex: boolean;
}

export interface FindReplacePrepared {
  matches: (text: string) => boolean;
  replace: (text: string) => string;
}

export interface FindReplaceMatch {
  sheetId: string;
  row: number;
  col: number;
  text: string;
}

export interface FindReplaceEntry {
  sheetId: string;
  row: number;
  col: number;
  text: string;
}

const REGEX_PATTERN_LIMIT = 120;
const REGEX_TEXT_LIMIT = 2048;

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const normalizeCellValue = (value: CellPrimitive): string => (value == null ? "" : String(value));

export const getSearchableCellText = (cell?: CellModel): string => {
  if (!cell) {
    return "";
  }

  return cell.formula ?? normalizeCellValue(cell.value);
};

export const collectFindReplaceEntries = (
  workbook: WorkbookModel,
  activeSheetId: string,
  scope: FindReplaceScope
): FindReplaceEntry[] => {
  const sheets = scope === "workbook" ? workbook.sheets : workbook.sheets.filter((sheet) => sheet.id === activeSheetId);
  const entries: FindReplaceEntry[] = [];

  for (const sheet of sheets) {
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const [rowPart, colPart] = key.split(":");
      const row = Number(rowPart);
      const col = Number(colPart);
      const text = getSearchableCellText(cell);

      if (!text) {
        continue;
      }

      entries.push({
        sheetId: sheet.id,
        row,
        col,
        text
      });
    }
  }

  return entries;
};

export const collectFindReplaceMatches = (
  entries: FindReplaceEntry[],
  prepared: FindReplacePrepared,
  start = 0,
  end = entries.length
): FindReplaceMatch[] =>
  entries.slice(start, end).filter((entry) => prepared.matches(entry.text));

export const createFindReplacePrepared = (
  options: FindReplaceOptions
): { prepared?: FindReplacePrepared; error?: string } => {
  if (!options.query) {
    return {};
  }

  if (!options.regex) {
    return createPlainPrepared(options);
  }

  return createRegexPrepared(options);
};

const createPlainPrepared = (options: FindReplaceOptions): { prepared: FindReplacePrepared } => {
  const source = options.caseSensitive ? options.query : options.query.toLocaleLowerCase();
  const replacePattern = new RegExp(escapeRegExp(options.query), options.caseSensitive ? "g" : "gi");

  return {
    prepared: {
      matches(text) {
        const candidate = options.caseSensitive ? text : text.toLocaleLowerCase();
        return options.wholeCell ? candidate === source : candidate.includes(source);
      },
      replace(text) {
        if (options.wholeCell) {
          return this.matches(text) ? options.replaceText : text;
        }
        return text.replace(replacePattern, options.replaceText);
      }
    }
  };
};

const createRegexPrepared = (options: FindReplaceOptions): { prepared?: FindReplacePrepared; error?: string } => {
  if (options.query.length > REGEX_PATTERN_LIMIT) {
    return {
      error: `Regex limitada a ${REGEX_PATTERN_LIMIT} caracteres por seguranca.`
    };
  }

  try {
    const baseSource = options.wholeCell ? `^(?:${options.query})$` : options.query;
    const flags = options.caseSensitive ? "" : "i";
    const testPattern = new RegExp(baseSource, flags);
    const replacePattern = new RegExp(baseSource, `${flags}g`);

    return {
      prepared: {
        matches(text) {
          if (text.length > REGEX_TEXT_LIMIT) {
            return false;
          }
          return testPattern.test(text);
        },
        replace(text) {
          if (text.length > REGEX_TEXT_LIMIT) {
            return text;
          }
          if (options.wholeCell) {
            return testPattern.test(text) ? text.replace(replacePattern, options.replaceText) : text;
          }
          return text.replace(replacePattern, options.replaceText);
        }
      }
    };
  } catch {
    return {
      error: "Regex invalida para a busca atual."
    };
  }
};
