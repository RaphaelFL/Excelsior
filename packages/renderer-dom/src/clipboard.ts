import type { ClipboardPolicy } from "@excelsior/core";

export interface ClipboardAnalysis {
  text: string;
  blocked: boolean;
  hardBlocked: boolean;
  reasons: string[];
}

const ALLOWED_TAGS = new Set([
  "HTML",
  "HEAD",
  "BODY",
  "TABLE",
  "THEAD",
  "TBODY",
  "TFOOT",
  "TR",
  "TD",
  "TH",
  "COLGROUP",
  "COL",
  "DIV",
  "P",
  "SPAN",
  "BR",
  "B",
  "STRONG",
  "I",
  "EM",
  "U",
  "S",
  "UL",
  "OL",
  "LI"
]);
const ALLOWED_ATTRIBUTES = new Set(["colspan", "rowspan"]);
const DANGEROUS_URL_PREFIXES = ["javascript:", "vbscript:", "data:"];

const normalizeClipboardText = (text: string): string => text.replaceAll("\r", "").trim();

const isAllowedClipboardAttribute = (name: string): boolean => ALLOWED_ATTRIBUTES.has(name);

const isDangerousClipboardUrl = (value: string): boolean =>
  DANGEROUS_URL_PREFIXES.some((prefix) => value.startsWith(prefix));

const collectBlockedAttributeReasons = (element: Element, reasons: Set<string>): void => {
  for (const attribute of Array.from(element.attributes)) {
    const lowerName = attribute.name.toLowerCase();
    const lowerValue = attribute.value.trim().toLowerCase();

    if (!isAllowedClipboardAttribute(lowerName)) {
      reasons.add(`blocked-attribute:${lowerName}`);
    }

    if (isDangerousClipboardUrl(lowerValue)) {
      reasons.add(`blocked-url:${lowerValue.split(":", 1)[0]}`);
    }
  }
};

const countClipboardCells = (text: string): number =>
  parseTabularText(text).reduce((total, row) => total + row.length, 0);

const withCellLimit = (analysis: ClipboardAnalysis, maxPasteCells?: number): ClipboardAnalysis => {
  if (!maxPasteCells || maxPasteCells < 1) {
    return analysis;
  }

  const cellCount = countClipboardCells(analysis.text);
  if (cellCount <= maxPasteCells) {
    return analysis;
  }

  return {
    ...analysis,
    blocked: true,
    hardBlocked: true,
    reasons: [...analysis.reasons, `max-paste-cells-exceeded:${cellCount}:${maxPasteCells}`]
  };
};

const collectTableText = (table: HTMLTableElement): string =>
  Array.from(table.rows)
    .map((row) =>
      Array.from(row.cells)
        .map((cell) => cell.textContent?.trim() ?? "")
        .join("\t")
    )
    .join("\n");

export const sanitizeClipboardHtml = (html: string): ClipboardAnalysis => {
  const parser = new DOMParser();
  const documentNode = parser.parseFromString(html, "text/html");
  const reasons = new Set<string>();

  for (const element of Array.from(documentNode.querySelectorAll("*"))) {
    if (!ALLOWED_TAGS.has(element.tagName)) {
      reasons.add(`blocked-tag:${element.tagName.toLowerCase()}`);
    }

    collectBlockedAttributeReasons(element, reasons);
  }

  const firstTable = documentNode.querySelector("table");
  const text = firstTable instanceof HTMLTableElement
    ? collectTableText(firstTable)
    : normalizeClipboardText(documentNode.body.textContent ?? "");

  return {
    text,
    blocked: reasons.size > 0,
    hardBlocked: false,
    reasons: [...reasons]
  };
};

export const resolveClipboardText = (
  text: string,
  html: string,
  policy: ClipboardPolicy,
  maxPasteCells?: number
): ClipboardAnalysis => {
  if (!html) {
    return withCellLimit({
      text,
      blocked: false,
      hardBlocked: false,
      reasons: []
    }, maxPasteCells);
  }

  if (policy === "text-only") {
    return withCellLimit({
      text,
      blocked: false,
      hardBlocked: false,
      reasons: []
    }, maxPasteCells);
  }

  const analysis = withCellLimit(sanitizeClipboardHtml(html), maxPasteCells);

  if (policy === "blocked-html" && analysis.blocked) {
    return {
      ...analysis,
      hardBlocked: true
    };
  }

  return {
    text: analysis.text || text,
    blocked: analysis.blocked,
    hardBlocked: analysis.hardBlocked,
    reasons: analysis.reasons
  };
};

export const parseTabularText = (text: string): string[][] => {
  if (!text.trim()) {
    return [];
  }

  return text
    .replaceAll("\r", "")
    .split("\n")
    .map((line) => line.split("\t"));
};