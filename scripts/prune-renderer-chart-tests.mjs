import fs from "node:fs";

const filePath = "packages/renderer-dom/tests/dom-spreadsheet-renderer.spec.ts";
const marker = "  it(\"applies text color and border color to selected cells from toolbar color controls\"";
const source = fs.readFileSync(filePath, "utf8");
const first = source.indexOf(marker);
const second = source.indexOf(marker, first + marker.length);
if (first < 0 || second < 0) {
  throw new Error("Could not find duplicated renderer test boundary");
}
fs.writeFileSync(filePath, source.slice(0, first) + source.slice(second));