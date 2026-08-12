export interface ChartMargin {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export type ChartAxisType = "linear" | "log" | "date" | "category" | "multicategory";

export interface ChartAxis {
  title?: string;
  min?: number;
  max?: number;
  type?: ChartAxisType;
  reverse?: boolean;
  tickFormat?: string;
  tickCount?: number;
  zeroLine?: boolean;
  autoRange?: boolean;
}

export interface ChartSubplotLayout {
  rows: number;
  cols: number;
  gapX?: number;
  gapY?: number;
  syncZoom?: boolean;
}

export interface ChartShape {
  type: "line" | "rect" | "circle" | "path" | "region";
  x0: number;
  y0: number;
  x1?: number;
  y1?: number;
  radius?: number;
  path?: string;
  points?: Array<{ x: number; y: number }>;
  xRef?: "paper" | "data";
  yRef?: "paper" | "data";
  stroke?: string;
  strokeWidth?: number;
  fill?: string;
  opacity?: number;
}

export interface ChartAnnotation {
  text: string;
  x: number;
  y: number;
  xRef?: "paper" | "data";
  yRef?: "paper" | "data";
  color?: string;
  fontSize?: number;
  align?: "left" | "center" | "right";
  rotate?: number;
  showArrow?: boolean;
  arrowToX?: number;
  arrowToY?: number;
}

export interface ChartImageLayer {
  source: string;
  x: number;
  y: number;
  width: number;
  height: number;
  xRef?: "paper" | "data";
  yRef?: "paper" | "data";
  opacity?: number;
}

export interface ChartLegend {
  visible: boolean;
  position: "top" | "right" | "bottom" | "left";
}

export interface ChartLayout {
  title?: string;
  width?: number;
  height?: number;
  margin: ChartMargin;
  backgroundColor: string;
  xAxis: ChartAxis;
  xAxis2: ChartAxis;
  yAxis: ChartAxis;
  yAxis2: ChartAxis;
  legend: ChartLegend;
  subplots?: ChartSubplotLayout;
  shapes: ChartShape[];
  annotations: ChartAnnotation[];
  images: ChartImageLayer[];
  theme?: string;
  template?: string;
}

export const DEFAULT_CHART_LAYOUT: ChartLayout = {
  title: "",
  margin: {
    top: 56,
    right: 24,
    bottom: 48,
    left: 56
  },
  backgroundColor: "#ffffff",
  xAxis: {},
  xAxis2: {},
  yAxis: {},
  yAxis2: {},
  legend: {
    visible: true,
    position: "top"
  },
  shapes: [],
  annotations: [],
  images: []
};
