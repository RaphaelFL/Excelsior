import { WorkbookEngine, type RowModel, type SpreadsheetOperation, type WorkbookConfig } from "@excelsior/core";
import { BasicFormulaEngine } from "@excelsior/formulas";
import {
  DomSpreadsheetRenderer,
  type AutofillOptions,
  type CustomCellEditor,
  type CustomCellRenderer,
  type DomSpreadsheetRendererOptions,
  type RendererLocalizationOptions
} from "@excelsior/renderer-dom";

export interface CreateSpreadsheetOptions extends WorkbookConfig {
  onChange?: (operations: SpreadsheetOperation[]) => void;
  cellRenderers?: CustomCellRenderer[];
  cellEditors?: CustomCellEditor[];
  widgetRenderers?: DomSpreadsheetRendererOptions["widgetRenderers"];
  includeHiddenCellsInClipboard?: boolean;
  autofill?: AutofillOptions;
  localization?: RendererLocalizationOptions;
  renderDebounceMs?: number;
  chartLimits?: DomSpreadsheetRendererOptions["chartLimits"];
  chartPerformance?: DomSpreadsheetRendererOptions["chartPerformance"];
  chartInsertPreview?: DomSpreadsheetRendererOptions["chartInsertPreview"];
  rowModel?: RowModel;
}

export interface SpreadsheetInstance {
  engine: WorkbookEngine;
  renderer: DomSpreadsheetRenderer;
  destroy: () => void;
}

export const createSpreadsheet = (
  container: HTMLElement,
  options: CreateSpreadsheetOptions = {}
): SpreadsheetInstance => {
  const engine = new WorkbookEngine(options, new BasicFormulaEngine());
  if (options.rowModel) {
    engine.setRowModel(engine.getActiveSheet().id, options.rowModel);
  }
  const renderer = new DomSpreadsheetRenderer(container, engine, {
    onChange: options.onChange,
    cellRenderers: options.cellRenderers,
    cellEditors: options.cellEditors,
    widgetRenderers: options.widgetRenderers,
    includeHiddenCellsInClipboard: options.includeHiddenCellsInClipboard,
    autofill: options.autofill,
    localization: options.localization,
    renderDebounceMs: options.renderDebounceMs,
    chartLimits: options.chartLimits,
    chartPerformance: options.chartPerformance,
    chartInsertPreview: options.chartInsertPreview
  });

  return {
    engine,
    renderer,
    destroy: () => {
      renderer.dispose();
      engine.dispose();
    }
  };
};