import type { ChartConfig } from "./Config";
import type { ChartLayout } from "./Layout";
import type { ChartTrace } from "./Trace";

export interface ChartFigureMetadata {
  [key: string]: string | number | boolean | null | undefined;
}

export interface ChartFrame {
  name?: string;
  data?: ChartTrace[];
  layout?: Partial<ChartLayout>;
  metadata?: ChartFigureMetadata;
}

export interface ChartSelectedPoint {
  traceIndex: number;
  pointIndex: number;
  traceName: string;
  x: string;
  y: string;
}

export interface ChartSelectionState {
  mode: "click" | "rect" | "lasso";
  points: ChartSelectedPoint[];
  updatedAt: string;
}

export interface ChartFigure {
  data: ChartTrace[];
  layout: ChartLayout;
  config: ChartConfig;
  frames: ChartFrame[];
  selection: ChartSelectionState | null;
  metadata: ChartFigureMetadata;
  schemaVersion: string;
}

export interface ChartFigureInput {
  data: ChartTrace[];
  layout?: Partial<ChartLayout>;
  config?: Partial<ChartConfig>;
  frames?: ChartFrame[];
  selection?: ChartSelectionState | null;
  metadata?: ChartFigureMetadata;
  schemaVersion?: string;
}
