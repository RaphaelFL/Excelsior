import { ChartEngine } from "./core/ChartEngine";
import { ChartConfigurationError } from "./core/chart-errors";
import { FigureValidator } from "./core/FigureValidator";
import { TraceRegistry } from "./core/TraceRegistry";
import type { ChartFigureInput } from "./model/Figure";
import type { ChartConfig } from "./model/Config";
import type { ChartHandle, CreateFigureOptions } from "./types/PublicApi";

const resolveContainer = (target: string | HTMLElement): HTMLElement => {
  if (typeof target === "string") {
    const found = document.querySelector<HTMLElement>(target);
    if (!found) {
      throw new ChartConfigurationError("CHART_CONTAINER_NOT_FOUND", `Container '${target}' was not found.`);
    }
    return found;
  }

  return target;
};

export const createFigure = (
  target: string | HTMLElement,
  figure: ChartFigureInput,
  options?: CreateFigureOptions
): ChartHandle => {
  const container = resolveContainer(target);
  if (options?.containerClassName) {
    container.classList.add(options.containerClassName);
  }

  const engine = new ChartEngine(container, figure);
  return {
    update: (nextFigure) => engine.update(nextFigure),
    updateData: (nextData) => engine.updateData(nextData),
    updateLayout: (nextLayout) => engine.updateLayout(nextLayout),
    resize: () => engine.resize(),
    destroy: () => engine.destroy(),
    exportSvg: () => engine.exportSvg(),
    exportPng: (exportOptions) => engine.exportPng(exportOptions),
    toJson: () => engine.toJson(),
    exportDataTable: () => engine.exportDataTable(),
    getSelection: () => engine.getSelection(),
    clearSelection: () => engine.clearSelection(),
    playFrames: (playbackOptions) => engine.playFrames(playbackOptions),
    stopFrames: () => engine.stopFrames(),
    isAnimating: () => engine.isAnimating(),
    on: (event, handler) => engine.on(event, handler),
    off: (event, handler) => engine.off(event, handler)
  };
};

export const fromJson = (
  target: string | HTMLElement,
  serializedFigure: string,
  options?: CreateFigureOptions
): ChartHandle => {
  const figure = parseFigureFromJson(serializedFigure);
  return createFigure(target, figure, options);
};

export const parseFigureFromJson = (serializedFigure: string): ChartFigureInput => {
  let parsed: unknown;
  try {
    parsed = JSON.parse(serializedFigure);
  } catch (error) {
    throw new ChartConfigurationError("CHART_JSON_INVALID", `Failed to parse chart JSON: ${toErrorMessage(error)}`);
  }

  if (!parsed || typeof parsed !== "object") {
    throw new ChartConfigurationError("CHART_JSON_INVALID", "Chart JSON must represent an object figure.");
  }

  const candidate = parsed as { data?: unknown; schemaVersion?: unknown; config?: unknown };
  assertKnownKeys(candidate, ALLOWED_FIGURE_KEYS, "figure");

  if (!Array.isArray(candidate.data)) {
    throw new ChartConfigurationError("CHART_JSON_INVALID", "Chart JSON must include a data array.");
  }

  if (candidate.schemaVersion !== undefined && typeof candidate.schemaVersion !== "string") {
    throw new ChartConfigurationError("CHART_JSON_INVALID", "Chart JSON schemaVersion must be a string when provided.");
  }

  if (candidate.schemaVersion && !/^\d+\.\d+\.\d+/.test(candidate.schemaVersion)) {
    throw new ChartConfigurationError("CHART_JSON_INVALID", "Chart JSON schemaVersion must follow semantic version format.");
  }

  if (candidate.config !== undefined) {
    if (!candidate.config || typeof candidate.config !== "object" || Array.isArray(candidate.config)) {
      throw new ChartConfigurationError("CHART_JSON_INVALID", "Chart JSON config must be an object when provided.");
    }
    assertKnownKeys(candidate.config as Record<string, unknown>, ALLOWED_CONFIG_KEYS, "config");
  }

  const validator = new FigureValidator(new TraceRegistry());
  return validator.normalize(parsed as ChartFigureInput);
};

export { ChartEngine } from "./core/ChartEngine";
export { TraceRegistry } from "./core/TraceRegistry";
export { FigureValidator } from "./core/FigureValidator";
export { FigureManager } from "./core/FigureManager";
export { LayoutEngine } from "./core/LayoutEngine";
export { RenderScheduler } from "./core/RenderScheduler";
export { DashboardComposer } from "./core/DashboardComposer";
export { SvgRenderer } from "./renderers/SvgRenderer";
export { CanvasRenderer } from "./renderers/CanvasRenderer";
export { WebglRenderer } from "./renderers/WebglRenderer";
export { HybridRenderer } from "./renderers/HybridRenderer";
export { LinearScale } from "./scales/LinearScale";
export { CategoryScale } from "./scales/CategoryScale";
export { ChartConfigurationError } from "./core/chart-errors";
export { applyThemeAndTemplate, DEFAULT_THEMES, DEFAULT_TEMPLATES } from "./core/themes";
export { applyTraceTransforms } from "./core/TraceTransforms";
export { resolveColorFromScale, resolvePalette, interpolatePalette, normalizeRatio } from "./core/color-scales";
export { simplifyLine, clipLineToRect } from "./core/line-optimization";
export { formatAxisTick, getAxisTickCount, normalizeAxisType, toAxisScalar } from "./core/axis-utils";
export { TooltipController } from "./interactions/TooltipController";
export { HoverController } from "./interactions/HoverController";
export { LegendController } from "./interactions/LegendController";
export { ModebarController } from "./interactions/ModebarController";
export { ZoomPanController } from "./interactions/ZoomPanController";
export { SelectionController } from "./interactions/SelectionController";
export { createFigureFromSpreadsheetRange } from "./adapters/spreadsheet-adapter";
export type { ChartRenderer } from "./renderers/ChartRenderer";
export type { ChartFigure, ChartFigureInput, ChartFigureMetadata, ChartFrame, ChartSelectionState, ChartSelectedPoint } from "./model/Figure";
export type {
  ChartTrace,
  CartesianTrace,
  ViolinTrace,
  DensityTrace,
  PieTrace,
  DonutTrace,
  SunburstTrace,
  TreemapTrace,
  SankeyTrace,
  ParallelCategoriesTrace,
  HistogramTrace,
  BoxTrace,
  HeatmapTrace,
  ContourTrace,
  QuiverTrace,
  FinancialTrace,
  WaterfallTrace,
  FunnelTrace,
  PolarTrace,
  TernaryTrace,
  GeoTrace,
  GeoScatterTrace,
  GeoLineTrace,
  Scatter3dTrace,
  SurfaceTrace,
  Mesh3dTrace,
  TraceType,
  TraceTransform,
  TraceAxisRef
} from "./model/Trace";
export type {
  ChartLayout,
  ChartAxis,
  ChartLegend,
  ChartMargin,
  ChartAxisType,
  ChartSubplotLayout,
  ChartShape,
  ChartAnnotation,
  ChartImageLayer
} from "./model/Layout";
export type {
  ChartPosition,
  ChartRangeBinding,
  WorksheetChartObject,
  WorksheetChartType
} from "./model/WorksheetChartObject";
export type { ChartConfig, ChartRendererKind, ChartHoverMode } from "./model/Config";
export type { ChartEventMap } from "./model/Events";
export type { ChartHandle, CreateFigureOptions, FramePlaybackOptions } from "./types/PublicApi";
export type { InteractionMode } from "./interactions/ZoomPanController";
export type { SelectionShape } from "./interactions/SelectionController";
export type { SpreadsheetCellValue, SpreadsheetRangeInput, SpreadsheetAdapterOptions } from "./adapters/spreadsheet-adapter";
export type { DashboardComposerOptions, DashboardWidget } from "./core/DashboardComposer";

const toErrorMessage = (error: unknown): string => (error instanceof Error ? error.message : String(error));

const ALLOWED_FIGURE_KEYS = new Set(["data", "layout", "config", "frames", "selection", "metadata", "schemaVersion"]);
const ALLOWED_CONFIG_KEYS = new Set<keyof ChartConfig>([
  "responsive",
  "modebar",
  "tooltip",
  "hoverMode",
  "renderer",
  "maxRenderPoints",
  "maxInteractivePoints",
  "spatialHover",
  "fullscreen",
  "accessibleTable",
  "ariaDescription",
  "highContrast",
  "frameDurationMs",
  "webglFallback",
  "syncSubplotZoom",
  "maxDensitySamples"
]);

const assertKnownKeys = (value: Record<string, unknown>, allowedKeys: Set<string>, context: string): void => {
  for (const key of Object.keys(value)) {
    if (!allowedKeys.has(key)) {
      throw new ChartConfigurationError("CHART_JSON_UNKNOWN_FIELD", `Unknown ${context} field '${key}' in chart JSON.`);
    }
  }
};
