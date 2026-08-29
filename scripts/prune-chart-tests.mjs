import fs from "node:fs";

const removeRange = (filePath, startMarker, endMarker) => {
  const source = fs.readFileSync(filePath, "utf8");
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  if (start < 0 || end < 0) {
    throw new Error(`Could not find chart test boundaries in ${filePath}`);
  }
  fs.writeFileSync(filePath, source.slice(0, start) + source.slice(end));
};

removeRange(
  "packages/renderer-dom/tests/dom-spreadsheet-renderer.spec.ts",
  "  /* chart-only tests removed for the Excel-only build. */",
  "  it(\"applies text color and border color to selected cells from toolbar color controls\""
);

removeRange(
  "packages/xlsx/tests/xlsx.spec.ts",
  "  it(\"roundtrips worksheet charts with source range and anchor metadata\"",
  "  it(\"returns a typed wrapper error when the requested sheet is missing\""
);