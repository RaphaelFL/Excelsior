import type { ChartFigure, ChartSelectedPoint } from "./Figure";

export interface ChartEventMap {
  "figure:created": { figure: ChartFigure };
  "figure:updated": { figure: ChartFigure };
  "figure:destroyed": { container: HTMLElement };
  "trace:hover": {
    traceIndex: number;
    pointIndex: number;
    traceName: string;
    x: string;
    y: string;
  };
  "trace:unhover": {
    traceIndex: number;
    pointIndex: number;
  };
  "trace:click": {
    traceIndex: number;
    pointIndex: number;
    traceName: string;
    x: string;
    y: string;
  };
  "trace:selected": {
    traceIndex: number;
    pointIndex: number;
    traceName: string;
    x: string;
    y: string;
  };
  "selection:changed": {
    mode: "click" | "rect" | "lasso";
    points: ChartSelectedPoint[];
  };
  "legend:toggled": {
    traceIndex: number;
    visible: boolean;
  };
  "axis:zoomed": {
    xMin: number;
    xMax: number;
    yMin: number;
    yMax: number;
  };
  "axis:panned": {
    xMin: number;
    xMax: number;
    yMin: number;
    yMax: number;
  };
  "layout:resized": { width: number; height: number };
  "animation:started": { frameCount: number; intervalMs: number };
  "animation:frame": { frameIndex: number; frameName?: string };
  "animation:stopped": { frameIndex: number };
  "export:started": { format: "svg" | "png" };
  "export:finished": { format: "svg" | "png" };
  "error:raised": { error: Error };
}
