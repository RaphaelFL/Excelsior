export type ChartRendererKind = "svg" | "canvas" | "webgl" | "hybrid";
export type ChartHoverMode = "point" | "x";

export interface ChartConfig {
  responsive: boolean;
  modebar: boolean;
  tooltip: boolean;
  hoverMode: ChartHoverMode;
  renderer: ChartRendererKind;
  maxRenderPoints: number;
  maxInteractivePoints: number;
  spatialHover: boolean;
  fullscreen: boolean;
  accessibleTable: boolean;
  ariaDescription: string;
  highContrast: boolean;
  frameDurationMs: number;
  webglFallback: boolean;
  syncSubplotZoom: boolean;
  maxDensitySamples: number;
}

export const DEFAULT_CHART_CONFIG: ChartConfig = {
  responsive: true,
  modebar: true,
  tooltip: true,
  hoverMode: "point",
  renderer: "svg",
  maxRenderPoints: 12_000,
  maxInteractivePoints: 25_000,
  spatialHover: true,
  fullscreen: true,
  accessibleTable: false,
  ariaDescription: "Chart visualization",
  highContrast: false,
  frameDurationMs: 450,
  webglFallback: true,
  syncSubplotZoom: true,
  maxDensitySamples: 20_000
};
