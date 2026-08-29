import fs from "node:fs";
import ts from "typescript";

const filePath = "packages/renderer-dom/src/dom-spreadsheet-renderer.ts";
let sourceText = fs.readFileSync(filePath, "utf8");

const sharedRenames = [
  ["ChartSurfaceMetrics", "VisualObjectSurfaceMetrics"],
  ["ChartRect", "VisualObjectRect"],
  ["ChartInteractionState", "VisualObjectInteractionState"],
  ["ChartPosition", "WorksheetObjectPosition"],
  ["chartSurfaceMetrics", "visualObjectSurfaceMetrics"],
  ["chartInteraction", "visualObjectInteraction"],
  ["chartLastInteractionMoveTs", "visualObjectLastInteractionMoveTs"],
  ["chartsLayer", "visualObjectsLayer"],
  ["handleChartLayer", "handleVisualObjectLayer"],
  ["handleChartInteraction", "handleVisualObjectInteraction"],
  ["resolveChartRect", "resolveVisualObjectRect"],
  ["resolveChartPositionFromRect", "resolveVisualObjectPositionFromRect"],
  ["renderChartObjects", "renderVisualObjects"],
  ["CHART_MIN_WIDTH", "VISUAL_OBJECT_MIN_WIDTH"],
  ["CHART_MIN_HEIGHT", "VISUAL_OBJECT_MIN_HEIGHT"]
];
for (const [from, to] of sharedRenames) {
  sourceText = sourceText.replaceAll(from, to);
}

sourceText = sourceText
  .replace(/import \{[\s\S]*?\} from "@excelsior\/charts";\r?\n/, "")
  .replace(/^\s*type WorksheetChartObject,\r?\n/m, "")
  .replace(/^\s*type WorksheetChartType,\r?\n/m, "")
  .replace(/^\s*(?:toolbarCharts|chart)\w+\??:[^\n]*\r?\n/gm, "")
  .replace(/^\s*(?:toolbarCharts|chart)\w+:\s*[^\n]*\r?\n/gm, "")
  .replace("// Fallback to JSON clone for plain chart payloads.", "// Fallback to JSON clone for plain serializable payloads.");

const sourceFile = ts.createSourceFile(filePath, sourceText, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
const edits = [];
const declarationName = (node) => {
  const name = node.name;
  return name && ts.isIdentifier(name) ? name.text : "";
};
const memberName = (node) => {
  const name = node.name;
  if (!name) return "";
  if (ts.isIdentifier(name) || ts.isStringLiteral(name) || ts.isNumericLiteral(name)) return name.text;
  return "";
};
const isChartName = (name) => /chart/i.test(name);

for (const statement of sourceFile.statements) {
  if (ts.isImportDeclaration(statement) || ts.isClassDeclaration(statement)) continue;
  if (
    (ts.isInterfaceDeclaration(statement) || ts.isTypeAliasDeclaration(statement) || ts.isFunctionDeclaration(statement)) &&
    isChartName(declarationName(statement))
  ) {
    edits.push([statement.getFullStart(), statement.end]);
    continue;
  }
  if (ts.isVariableStatement(statement) && statement.declarationList.declarations.some((item) => isChartName(declarationName(item)))) {
    edits.push([statement.getFullStart(), statement.end]);
  }
}

for (const statement of sourceFile.statements) {
  if (!ts.isClassDeclaration(statement)) continue;
  for (const member of statement.members) {
    if (isChartName(memberName(member))) {
      edits.push([member.getFullStart(), member.end]);
    }
  }
}

edits.sort((left, right) => right[0] - left[0]);
for (const [start, end] of edits) {
  sourceText = sourceText.slice(0, start) + sourceText.slice(end);
}

sourceText = sourceText
  .replace(
    /\s*const chartLimits = this\.getChartLimits\(\);[\s\S]*?\n\s*const imageIds = new Set<string>\(\);/,
    "\n    const imageIds = new Set<string>();"
  )
  .replaceAll("chartId", "objectId")
  .replaceAll("excelsior-charts-layer", "excelsior-visual-objects-layer")
  .replaceAll("excelsior-chart-object", "excelsior-visual-object")
  .replaceAll("excelsior-chart-preview-placeholder", "excelsior-widget-placeholder")
  .replaceAll("[data-chart-id], ", "")
  .replaceAll(", [data-chart-action]", "")
  .replaceAll("[data-chart-resize='true'], ", "")
  .replaceAll("[data-chart-move='true'], ", "");

fs.writeFileSync(filePath, sourceText);
