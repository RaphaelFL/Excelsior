import type { WorkbookModel } from "@excelsior/core";
import { createXlsxAdapter } from "./xlsx-adapter";

const { exportWorkbookToXlsx, importWorkbookFromXlsx, exportTableToXlsx, importTableFromXlsx } = createXlsxAdapter();

export interface WorkbookInteropEngine {
  toJSON(): WorkbookModel;
  loadFromJSON(snapshot: WorkbookModel): void;
}

export const exportWorkbookEngineToXlsx = async (
  engine: WorkbookInteropEngine
): Promise<Uint8Array> => {
  return exportWorkbookToXlsx(engine.toJSON());
};

export const importWorkbookIntoEngineFromXlsx = async (
  engine: WorkbookInteropEngine,
  input: Uint8Array | ArrayBuffer
): Promise<WorkbookModel> => {
  const snapshot = await importWorkbookFromXlsx(input);
  engine.loadFromJSON(snapshot);
  return snapshot;
};

export { exportWorkbookToXlsx, importWorkbookFromXlsx, exportTableToXlsx, importTableFromXlsx };
export type {
  ParseSheetDataResult,
  ReadSheet,
  Schema,
  WriterColumn,
  XlsxTableExportOptions,
  XlsxTableImportOptions
} from "./xlsx-adapter";