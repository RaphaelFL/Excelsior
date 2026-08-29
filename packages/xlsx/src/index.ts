import type { WorkbookModel } from "@excelsior/core";
import { createXlsxAdapter } from "./xlsx-adapter";

const { exportWorkbookToXlsx, importWorkbookFromXlsx, exportTableToXlsx, importTableFromXlsx } = createXlsxAdapter();

export interface WorkbookChartInteropEngine {
  toJSON(options?: { emitChartExportEvents?: boolean }): WorkbookModel;
  loadFromJSON(snapshot: WorkbookModel): void;
  reportChartExported?: (sheetId: string, chartId: string) => void;
  reportChartUnsupportedFeature?: (sheetId: string, chartId: string, feature: string) => void;
  reportChartError?: (payload: { sheetId?: string; chartId?: string; errorCode: string; message: string }) => void;
}

export const exportWorkbookEngineToXlsx = async (
  engine: WorkbookChartInteropEngine
): Promise<Uint8Array> => {
  const snapshot = engine.toJSON();
  return exportWorkbookToXlsx(snapshot, {
    onChartExported: ({ sheetId, chartId }) => {
      engine.reportChartExported?.(sheetId, chartId);
    },
    onChartUnsupportedFeature: ({ sheetId, chartId, feature }) => {
      engine.reportChartUnsupportedFeature?.(sheetId, chartId, feature);
    },
    onChartError: (payload) => {
      engine.reportChartError?.(payload);
    }
  });
};

export const importWorkbookIntoEngineFromXlsx = async (
  engine: WorkbookChartInteropEngine,
  input: Uint8Array | ArrayBuffer
): Promise<WorkbookModel> => {
  const snapshot = await importWorkbookFromXlsx(input, {
    onChartError: (payload) => {
      engine.reportChartError?.(payload);
    }
  });
  engine.loadFromJSON(snapshot);
  return snapshot;
};

export { exportWorkbookToXlsx, importWorkbookFromXlsx, exportTableToXlsx, importTableFromXlsx };
export type {
  ParseSheetDataResult,
  ReadSheet,
  Schema,
  WriterColumn,
  XlsxChartInteropOptions,
  XlsxTableExportOptions,
  XlsxTableImportOptions
} from "./xlsx-adapter";