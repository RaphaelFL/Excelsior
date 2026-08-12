export type TraceType =
  | "scatter"
  | "line"
  | "bar"
  | "area"
  | "violin"
  | "density"
  | "distribution"
  | "pie"
  | "donut"
  | "sunburst"
  | "treemap"
  | "sankey"
  | "parallel-categories"
  | "histogram"
  | "box"
  | "heatmap"
  | "contour"
  | "quiver"
  | "candlestick"
  | "ohlc"
  | "waterfall"
  | "funnel"
  | "polar"
  | "ternary"
  | "geo"
  | "geo-scatter"
  | "geo-line"
  | "scatter3d"
  | "surface"
  | "mesh3d"
  | (string & {});
export type TracePoint = string | number | Date;
export type TraceAxisRef = "x" | "x2" | "y" | "y2";

export type TraceTransform =
  | {
      type: "sort";
      by: "x" | "y";
      direction?: "asc" | "desc";
    }
  | {
      type: "filter";
      field: "x" | "y";
      min?: number;
      max?: number;
    }
  | {
      type: "normalize";
      field: "y";
      strategy?: "max" | "sum";
    }
  | {
      type: "group";
      by: "x" | "y";
      aggregate?: "sum" | "avg" | "min" | "max" | "count";
    }
  | {
      type: "aggregate";
      field: "y";
      op: "sum" | "avg" | "min" | "max" | "count";
    }
  | {
      type: "bin";
      field: "x" | "y";
      bins?: number;
    }
  | {
      type: "stack";
      field: "y";
      strategy?: "running" | "offset";
    }
  | {
      type: "cumulative";
      field: "y";
    }
  | {
      type: "percent";
      field: "y";
    };

export interface BaseTrace {
  type: TraceType;
  name?: string;
  visible?: boolean;
  subplot?: number;
  xAxisRef?: TraceAxisRef;
  yAxisRef?: TraceAxisRef;
  transforms?: TraceTransform[];
}

export interface CartesianTrace extends BaseTrace {
  type: "scatter" | "line" | "bar" | "area";
  x: TracePoint[];
  y: number[];
  mode?: "lines" | "markers" | "lines+markers";
  orientation?: "vertical" | "horizontal";
  stackGroup?: string;
  normalizeStack?: boolean;
  fill?: "tozero" | "tonext";
  errorY?: {
    values: number[];
  };
  marker?: {
    color?: string | string[];
    size?: number | number[];
    opacity?: number;
  };
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface ViolinTrace extends BaseTrace {
  type: "violin";
  values: number[];
  categories?: TracePoint[];
  bandwidth?: number;
  showBox?: boolean;
  showPoints?: boolean;
  marker?: {
    color?: string;
    opacity?: number;
  };
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface DensityTrace extends BaseTrace {
  type: "density" | "distribution";
  values: number[];
  cumulative?: boolean;
  sampleLimit?: number;
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface PieTrace extends BaseTrace {
  type: "pie";
  values: number[];
  labels?: string[];
  hole?: number;
  pull?: number | number[];
  marker?: {
    colors?: string[];
    opacity?: number;
  };
}

export interface DonutTrace extends BaseTrace {
  type: "donut";
  values: number[];
  labels?: string[];
  hole?: number;
  pull?: number | number[];
  marker?: {
    colors?: string[];
    opacity?: number;
  };
}

export interface SunburstTrace extends BaseTrace {
  type: "sunburst";
  ids?: Array<string | number>;
  labels: string[];
  parents: string[];
  values: number[];
  rootId?: string | number;
  drilldownDepth?: number;
  marker?: {
    colors?: string[];
    opacity?: number;
  };
}

export interface TreemapTrace extends BaseTrace {
  type: "treemap";
  ids?: Array<string | number>;
  labels: string[];
  parents: string[];
  values: number[];
  rootId?: string | number;
  drilldownDepth?: number;
  marker?: {
    colors?: string[];
    opacity?: number;
  };
}

export interface SankeyTrace extends BaseTrace {
  type: "sankey";
  nodes: {
    ids: Array<string | number>;
    labels?: string[];
    colors?: string[];
  };
  links: {
    source: number[];
    target: number[];
    value: number[];
    colors?: string[];
  };
}

export interface ParallelCategoriesTrace extends BaseTrace {
  type: "parallel-categories";
  dimensions: Array<{
    name: string;
    values: Array<string | number | boolean | null>;
  }>;
  line?: {
    color?: string;
    opacity?: number;
  };
}

export interface HistogramTrace extends BaseTrace {
  type: "histogram";
  values: number[];
  bins?: number;
  orientation?: "vertical" | "horizontal";
  normalize?: boolean;
  cumulative?: boolean;
  marker?: {
    color?: string;
    opacity?: number;
  };
}

export interface BoxTrace extends BaseTrace {
  type: "box";
  values: number[];
  boxpoints?: boolean;
  showOutliers?: boolean;
  outlierIqrFactor?: number;
  marker?: {
    color?: string;
    opacity?: number;
  };
  line?: {
    color?: string;
    width?: number;
  };
}

export interface HeatmapTrace extends BaseTrace {
  type: "heatmap";
  z: number[][];
  x?: TracePoint[];
  y?: TracePoint[];
  colorscale?: string[];
  zMin?: number;
  zMax?: number;
}

export interface ContourTrace extends BaseTrace {
  type: "contour";
  z: number[][];
  x?: TracePoint[];
  y?: TracePoint[];
  levels?: number;
  fillContours?: boolean;
  labelLevels?: boolean;
  maxSegments?: number;
  colorscale?: string[];
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface QuiverTrace extends BaseTrace {
  type: "quiver";
  x: number[];
  y: number[];
  u: number[];
  v: number[];
  scale?: number;
  colorscale?: string[];
  colorByMagnitude?: boolean;
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface FinancialTrace extends BaseTrace {
  type: "candlestick" | "ohlc";
  x: TracePoint[];
  open: number[];
  high: number[];
  low: number[];
  close: number[];
  increasing?: {
    color?: string;
  };
  decreasing?: {
    color?: string;
  };
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface WaterfallTrace extends BaseTrace {
  type: "waterfall";
  x: TracePoint[];
  y: number[];
  measure?: Array<"relative" | "total" | "absolute">;
  increasing?: {
    color?: string;
  };
  decreasing?: {
    color?: string;
  };
  totals?: {
    color?: string;
  };
}

export interface FunnelTrace extends BaseTrace {
  type: "funnel";
  labels: string[];
  values: number[];
  sort?: "none" | "asc" | "desc";
  marker?: {
    color?: string[];
    opacity?: number;
  };
}

export interface PolarTrace extends BaseTrace {
  type: "polar";
  theta: TracePoint[];
  r: number[];
  variant?: "scatter" | "line" | "bar" | "area";
  barWidth?: number;
  mode?: "lines" | "markers" | "lines+markers";
  marker?: {
    color?: string;
    size?: number;
    opacity?: number;
  };
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface TernaryTrace extends BaseTrace {
  type: "ternary";
  a: number[];
  b: number[];
  c: number[];
  marker?: {
    color?: string;
    size?: number;
    opacity?: number;
  };
}

export interface GeoTrace extends BaseTrace {
  type: "geo";
  geojson: {
    type: "FeatureCollection";
    features: Array<{
      type: "Feature";
      geometry: {
        type: "Polygon" | "MultiPolygon";
        coordinates: number[][][] | number[][][][];
      };
      properties?: Record<string, string | number | boolean | null>;
    }>;
  };
  locations?: Array<string | number>;
  values?: number[];
  featureIdField?: string;
  valueField?: string;
  colorscale?: string[];
  reverseScale?: boolean;
  missingColor?: string;
  showColorLegend?: boolean;
  simplifyTolerance?: number;
}

export interface GeoScatterTrace extends BaseTrace {
  type: "geo-scatter";
  lat: number[];
  lon: number[];
  marker?: {
    color?: string;
    size?: number;
    opacity?: number;
  };
  mode?: "markers" | "lines" | "lines+markers";
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface GeoLineTrace extends BaseTrace {
  type: "geo-line";
  paths: Array<Array<{ lat: number; lon: number }>>;
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface Scatter3dTrace extends BaseTrace {
  type: "scatter3d";
  x: number[];
  y: number[];
  z: number[];
  mode?: "markers" | "lines" | "lines+markers";
  marker?: {
    color?: string;
    size?: number;
    opacity?: number;
  };
  line?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
}

export interface SurfaceTrace extends BaseTrace {
  type: "surface";
  z: number[][];
  x?: number[];
  y?: number[];
  colorscale?: string[];
  reverseScale?: boolean;
}

export interface Mesh3dTrace extends BaseTrace {
  type: "mesh3d";
  x: number[];
  y: number[];
  z: number[];
  i: number[];
  j: number[];
  k: number[];
  marker?: {
    color?: string;
    opacity?: number;
  };
}

export type ChartTrace =
  | CartesianTrace
  | ViolinTrace
  | DensityTrace
  | PieTrace
  | DonutTrace
  | SunburstTrace
  | TreemapTrace
  | SankeyTrace
  | ParallelCategoriesTrace
  | HistogramTrace
  | BoxTrace
  | HeatmapTrace
  | ContourTrace
  | QuiverTrace
  | FinancialTrace
  | WaterfallTrace
  | FunnelTrace
  | PolarTrace
  | TernaryTrace
  | GeoTrace
  | GeoScatterTrace
  | GeoLineTrace
  | Scatter3dTrace
  | SurfaceTrace
  | Mesh3dTrace
  | (BaseTrace & Record<string, unknown>);
