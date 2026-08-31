import type { ChartFigure } from "./Figure";
import type { TraceType } from "./Trace";

export type WorksheetChartType = TraceType | "column" | "unknown";

export interface ChartRangeBinding {
  chartId: string;
  sheetId: string;
  sourceSheetId?: string;
  rangeAddress: string;
  orientation: "rows" | "columns";
  firstRowAsHeader: boolean;
  firstColumnAsLabel: boolean;
  autoRefresh: boolean;
}

export interface ChartPosition {
  fromCell: string;
  toCell?: string;
  offsetX: number;
  offsetY: number;
  width: number;
  height: number;
  zIndex: number;
}

export interface WorksheetChartObject {
  id: string;
  sheetId: string;
  type: WorksheetChartType;
  title?: string;
  sourceRange?: ChartRangeBinding;
  figure: ChartFigure;
  position: ChartPosition;
  style?: {
    theme?: string;
    background?: string;
    borderColor?: string;
    borderWidth?: number;
    fontFamily?: string;
    professionalPreset?: "spreadsheet" | "report" | "dashboard";
    displayMode?: "embedded" | "sheet";
  };
  state: {
    selected: boolean;
    visible: boolean;
    locked: boolean;
    lastRenderedAt?: number;
  };
  excelInterop: {
    originalChartType?: string;
    originalChartId?: string;
    originalAnchor?: unknown;
    unsupportedFeatures?: string[];
    fallbackImage?: boolean;
    preservedRawMetadata?: unknown;
  };
}
