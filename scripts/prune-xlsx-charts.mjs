import fs from "node:fs";

const filePath = "packages/xlsx/src/xlsx-adapter.ts";
const source = fs.readFileSync(filePath, "utf8");
const startMarker = "  const parseChartTypeTag =";
const endMarker = "  const addWorkbookFeatures =";
const start = source.indexOf(startMarker);
const end = source.indexOf(endMarker, start + startMarker.length);
if (start < 0 || end < 0) {
  throw new Error("Could not find XLSX chart interop boundaries");
}
fs.writeFileSync(filePath, source.slice(0, start) + source.slice(end));