import { fileURLToPath } from "node:url";
import { createXlsxAdapter } from "./xlsx-adapter";

const { exportWorkbookToXlsx, importWorkbookFromXlsx, exportTableToXlsx, importTableFromXlsx } = createXlsxAdapter(
  fileURLToPath(import.meta.url)
);

export { exportWorkbookToXlsx, importWorkbookFromXlsx, exportTableToXlsx, importTableFromXlsx };
export type { ParseSheetDataResult, ReadSheet, Schema, WriterColumn, XlsxTableExportOptions, XlsxTableImportOptions } from "./xlsx-adapter";