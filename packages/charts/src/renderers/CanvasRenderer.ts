import type {
  BoxTrace,
  CartesianTrace,
  ChartTrace,
  ContourTrace,
  DensityTrace,
  DonutTrace,
  FinancialTrace,
  FunnelTrace,
  GeoLineTrace,
  GeoScatterTrace,
  GeoTrace,
  HeatmapTrace,
  HistogramTrace,
  Mesh3dTrace,
  ParallelCategoriesTrace,
  PieTrace,
  PolarTrace,
  QuiverTrace,
  SankeyTrace,
  Scatter3dTrace,
  SunburstTrace,
  SurfaceTrace,
  TernaryTrace,
  TreemapTrace,
  ViolinTrace,
  WaterfallTrace
} from "../model/Trace";
import type { ChartFigure } from "../model/Figure";
import type { ComputedLayout, ComputedPlotArea } from "../core/LayoutEngine";
import type { ChartRenderer } from "./ChartRenderer";
import { ChartConfigurationError } from "../core/chart-errors";
import { buildCartesianDomains, isCartesianTrace } from "../core/cartesian-domain";
import { LinearScale } from "../scales/LinearScale";
import { CategoryScale } from "../scales/CategoryScale";
import {
  computeBoxOutliers,
  computeBoxStats,
  computeDensityCurve,
  computeHeatmapMatrix,
  computeHistogramBins,
  computeViolinProfile,
  formatNumeric
} from "../core/advanced-trace-utils";
import { SvgRenderer } from "./SvgRenderer";
import { clipLineToRect, simplifyLine } from "../core/line-optimization";
import { resolveColorFromScale, resolvePalette } from "../core/color-scales";
import { formatAxisTick, getAxisTickCount, normalizeAxisType, toAxisScalar } from "../core/axis-utils";
import { buildCartesianStackContext } from "../core/stacking-utils";

const DEFAULT_SERIES_COLORS = ["#2563eb", "#f97316", "#059669", "#dc2626", "#7c3aed", "#0891b2", "#ca8a04"];
const DEFAULT_HEATMAP_COLORS = ["#0ea5e9", "#f8fafc", "#ef4444"];

const isPieTrace = (trace: ChartTrace): trace is PieTrace => trace.type === "pie";
const isDonutTrace = (trace: ChartTrace): trace is DonutTrace => trace.type === "donut";
const isSunburstTrace = (trace: ChartTrace): trace is SunburstTrace => trace.type === "sunburst";
const isTreemapTrace = (trace: ChartTrace): trace is TreemapTrace => trace.type === "treemap";
const isHistogramTrace = (trace: ChartTrace): trace is HistogramTrace => trace.type === "histogram";
const isBoxTrace = (trace: ChartTrace): trace is BoxTrace => trace.type === "box";
const isHeatmapTrace = (trace: ChartTrace): trace is HeatmapTrace => trace.type === "heatmap";
const isContourTrace = (trace: ChartTrace): trace is ContourTrace => trace.type === "contour";
const isViolinTrace = (trace: ChartTrace): trace is ViolinTrace => trace.type === "violin";
const isDensityTrace = (trace: ChartTrace): trace is DensityTrace => trace.type === "density" || trace.type === "distribution";
const isFinancialTrace = (trace: ChartTrace): trace is FinancialTrace => trace.type === "candlestick" || trace.type === "ohlc";
const isWaterfallTrace = (trace: ChartTrace): trace is WaterfallTrace => trace.type === "waterfall";
const isFunnelTrace = (trace: ChartTrace): trace is FunnelTrace => trace.type === "funnel";
const isSankeyTrace = (trace: ChartTrace): trace is SankeyTrace => trace.type === "sankey";
const isParallelCategoriesTrace = (trace: ChartTrace): trace is ParallelCategoriesTrace => trace.type === "parallel-categories";
const isQuiverTrace = (trace: ChartTrace): trace is QuiverTrace => trace.type === "quiver";
const isPolarTrace = (trace: ChartTrace): trace is PolarTrace => trace.type === "polar";
const isTernaryTrace = (trace: ChartTrace): trace is TernaryTrace => trace.type === "ternary";
const isGeoTrace = (trace: ChartTrace): trace is GeoTrace => trace.type === "geo";
const isGeoScatterTrace = (trace: ChartTrace): trace is GeoScatterTrace => trace.type === "geo-scatter";
const isGeoLineTrace = (trace: ChartTrace): trace is GeoLineTrace => trace.type === "geo-line";
const isScatter3dTrace = (trace: ChartTrace): trace is Scatter3dTrace => trace.type === "scatter3d";
const isSurfaceTrace = (trace: ChartTrace): trace is SurfaceTrace => trace.type === "surface";
const isMesh3dTrace = (trace: ChartTrace): trace is Mesh3dTrace => trace.type === "mesh3d";

type IndexedTrace<TTrace extends ChartTrace = ChartTrace> = {
  trace: TTrace;
  index: number;
};

export class CanvasRenderer implements ChartRenderer {
  private container: HTMLElement | null = null;
  private rootElement: HTMLDivElement | null = null;
  private canvasElement: HTMLCanvasElement | null = null;
  private interactionLayer: HTMLDivElement | null = null;
  private context: CanvasRenderingContext2D | null = null;
  private lastFigure: ChartFigure | null = null;
  private lastLayout: ComputedLayout | null = null;
  private readonly imageCache = new Map<string, HTMLImageElement>();

  mount(container: HTMLElement): void {
    this.destroy();
    this.container = container;
    if (container.style.position === "" || container.style.position === "static") {
      container.style.position = "relative";
    }

    const root = document.createElement("div");
    root.className = "excelsior-chart-canvas-root";
    Object.assign(root.style, {
      position: "absolute",
      inset: "0",
      overflow: "hidden",
      touchAction: "none"
    });

    const canvas = document.createElement("canvas");
    Object.assign(canvas.style, {
      position: "absolute",
      inset: "0",
      width: "100%",
      height: "100%"
    });

    const interactionLayer = document.createElement("div");
    Object.assign(interactionLayer.style, {
      position: "absolute",
      inset: "0",
      pointerEvents: "auto"
    });

    root.append(canvas, interactionLayer);
    container.append(root);
    this.rootElement = root;
    this.canvasElement = canvas;
    this.interactionLayer = interactionLayer;
    this.context = canvas.getContext("2d");
    if (!this.context) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_MOUNTED", "Canvas 2D context could not be created.");
    }
  }

  render(figure: ChartFigure, layout: ComputedLayout): void {
    const context = this.ensureContext();
    const canvas = this.ensureCanvas();
    const layer = this.ensureInteractionLayer();
    this.lastFigure = structuredClone(figure);
    this.lastLayout = layout;

    canvas.width = Math.max(1, Math.round(layout.width));
    canvas.height = Math.max(1, Math.round(layout.height));
    context.clearRect(0, 0, canvas.width, canvas.height);
    layer.replaceChildren();

    context.fillStyle = figure.layout.backgroundColor;
    context.fillRect(0, 0, layout.width, layout.height);

    if (figure.layout.title) {
      context.fillStyle = "#0f172a";
      context.font = "600 16px Arial, sans-serif";
      context.textBaseline = "alphabetic";
      context.fillText(figure.layout.title, layout.margin.left, layout.titleY);
    }

    const allTraces = figure.data.map((trace, index) => ({ trace, index }));
    const visibleTraces = allTraces.filter(({ trace }) => trace.visible !== false);
    const cartesianTraces = visibleTraces.filter((entry): entry is IndexedTrace<CartesianTrace> => isCartesianTrace(entry.trace));
    const financialTraces = visibleTraces.filter((entry): entry is IndexedTrace<FinancialTrace> => isFinancialTrace(entry.trace));
    const histogramTraces = visibleTraces.filter((entry): entry is IndexedTrace<HistogramTrace> => isHistogramTrace(entry.trace));
    const boxTraces = visibleTraces.filter((entry): entry is IndexedTrace<BoxTrace> => isBoxTrace(entry.trace));
    const violinTraces = visibleTraces.filter((entry): entry is IndexedTrace<ViolinTrace> => isViolinTrace(entry.trace));
    const densityTraces = visibleTraces.filter((entry): entry is IndexedTrace<DensityTrace> => isDensityTrace(entry.trace));
    const heatmapTraces = visibleTraces.filter((entry): entry is IndexedTrace<HeatmapTrace> => isHeatmapTrace(entry.trace));
    const contourTraces = visibleTraces.filter((entry): entry is IndexedTrace<ContourTrace> => isContourTrace(entry.trace));
    const quiverTraces = visibleTraces.filter((entry): entry is IndexedTrace<QuiverTrace> => isQuiverTrace(entry.trace));
    const waterfallTraces = visibleTraces.filter((entry): entry is IndexedTrace<WaterfallTrace> => isWaterfallTrace(entry.trace));
    const funnelTraces = visibleTraces.filter((entry): entry is IndexedTrace<FunnelTrace> => isFunnelTrace(entry.trace));
    const pieTraces = visibleTraces.filter((entry): entry is IndexedTrace<PieTrace> => isPieTrace(entry.trace));
    const donutTraces = visibleTraces.filter((entry): entry is IndexedTrace<DonutTrace> => isDonutTrace(entry.trace));
    const sunburstTraces = visibleTraces.filter((entry): entry is IndexedTrace<SunburstTrace> => isSunburstTrace(entry.trace));
    const treemapTraces = visibleTraces.filter((entry): entry is IndexedTrace<TreemapTrace> => isTreemapTrace(entry.trace));
    const sankeyTraces = visibleTraces.filter((entry): entry is IndexedTrace<SankeyTrace> => isSankeyTrace(entry.trace));
    const parallelCategoryTraces = visibleTraces.filter(
      (entry): entry is IndexedTrace<ParallelCategoriesTrace> => isParallelCategoriesTrace(entry.trace)
    );
    const polarTraces = visibleTraces.filter((entry): entry is IndexedTrace<PolarTrace> => isPolarTrace(entry.trace));
    const ternaryTraces = visibleTraces.filter((entry): entry is IndexedTrace<TernaryTrace> => isTernaryTrace(entry.trace));
    const geoTraces = visibleTraces.filter((entry): entry is IndexedTrace<GeoTrace> => isGeoTrace(entry.trace));
    const geoScatterTraces = visibleTraces.filter((entry): entry is IndexedTrace<GeoScatterTrace> => isGeoScatterTrace(entry.trace));
    const geoLineTraces = visibleTraces.filter((entry): entry is IndexedTrace<GeoLineTrace> => isGeoLineTrace(entry.trace));
    const scatter3dTraces = visibleTraces.filter((entry): entry is IndexedTrace<Scatter3dTrace> => isScatter3dTrace(entry.trace));
    const surfaceTraces = visibleTraces.filter((entry): entry is IndexedTrace<SurfaceTrace> => isSurfaceTrace(entry.trace));
    const mesh3dTraces = visibleTraces.filter((entry): entry is IndexedTrace<Mesh3dTrace> => isMesh3dTrace(entry.trace));
    const maxRenderPoints = Math.max(500, Number(figure.config.maxRenderPoints ?? 12_000));
    const subplotAreas = layout.subplotAreas.length > 0 ? layout.subplotAreas : [layout.plotArea];

    if (cartesianTraces.length > 0) {
      this.renderSubplotGroups(cartesianTraces, subplotAreas, (group, plotArea, subplotIndex) =>
        this.renderCartesianScene(context, figure, group, plotArea, maxRenderPoints, subplotIndex)
      );
    } else if (financialTraces.length > 0) {
      this.renderSubplotGroups(financialTraces, subplotAreas, (group, plotArea, subplotIndex) =>
        this.renderFinancialScene(context, figure, group, plotArea, subplotIndex)
      );
    } else if (histogramTraces.length > 0) {
      this.renderSubplotGroups(histogramTraces, subplotAreas, (group, plotArea) => this.renderHistogramScene(context, figure, group, plotArea));
    } else if (boxTraces.length > 0) {
      this.renderSubplotGroups(boxTraces, subplotAreas, (group, plotArea) => this.renderBoxScene(context, figure, group, plotArea));
    } else if (violinTraces.length > 0) {
      this.renderSubplotGroups(violinTraces, subplotAreas, (group, plotArea) => this.renderViolinScene(context, figure, group, plotArea));
    } else if (densityTraces.length > 0) {
      this.renderSubplotGroups(densityTraces, subplotAreas, (group, plotArea) => this.renderDensityScene(context, figure, group, plotArea));
    } else if (heatmapTraces.length > 0) {
      this.renderSubplotGroups(heatmapTraces, subplotAreas, (group, plotArea) => this.renderHeatmapScene(context, group[0], plotArea));
    } else if (contourTraces.length > 0) {
      this.renderSubplotGroups(contourTraces, subplotAreas, (group, plotArea) => this.renderContourScene(context, figure, group[0], plotArea));
    } else if (quiverTraces.length > 0) {
      this.renderSubplotGroups(quiverTraces, subplotAreas, (group, plotArea) => this.renderQuiverScene(context, figure, group, plotArea));
    } else if (waterfallTraces.length > 0) {
      this.renderSubplotGroups(waterfallTraces, subplotAreas, (group, plotArea) => this.renderWaterfallScene(context, figure, group, plotArea));
    } else if (funnelTraces.length > 0) {
      this.renderSubplotGroups(funnelTraces, subplotAreas, (group, plotArea) => this.renderFunnelScene(context, group[0], plotArea));
    } else if (pieTraces.length > 0) {
      this.renderSubplotGroups(pieTraces, subplotAreas, (group, plotArea) => this.renderPieScene(context, group[0], plotArea));
    } else if (donutTraces.length > 0) {
      this.renderSubplotGroups(donutTraces, subplotAreas, (group, plotArea) => this.renderDonutScene(context, group[0], plotArea));
    } else if (sunburstTraces.length > 0) {
      this.renderSubplotGroups(sunburstTraces, subplotAreas, (group, plotArea) => this.renderSunburstScene(context, group[0], plotArea));
    } else if (treemapTraces.length > 0) {
      this.renderSubplotGroups(treemapTraces, subplotAreas, (group, plotArea) => this.renderTreemapScene(context, group[0], plotArea));
    } else if (sankeyTraces.length > 0) {
      this.renderSubplotGroups(sankeyTraces, subplotAreas, (group, plotArea) => this.renderSankeyScene(context, group[0], plotArea));
    } else if (parallelCategoryTraces.length > 0) {
      this.renderSubplotGroups(parallelCategoryTraces, subplotAreas, (group, plotArea) =>
        this.renderParallelCategoriesScene(context, figure, group[0], plotArea)
      );
    } else if (polarTraces.length > 0) {
      this.renderSubplotGroups(polarTraces, subplotAreas, (group, plotArea) => this.renderPolarScene(context, group, plotArea));
    } else if (ternaryTraces.length > 0) {
      this.renderSubplotGroups(ternaryTraces, subplotAreas, (group, plotArea) => this.renderTernaryScene(context, group, plotArea));
    } else if (geoTraces.length > 0 || geoScatterTraces.length > 0 || geoLineTraces.length > 0) {
      this.renderGeoScene(context, geoTraces, geoScatterTraces, geoLineTraces, layout.plotArea);
    } else if (scatter3dTraces.length > 0 || surfaceTraces.length > 0 || mesh3dTraces.length > 0) {
      this.renderProjected3dScene(context, scatter3dTraces, surfaceTraces, mesh3dTraces, layout.plotArea);
    }

    this.renderLayoutOverlays(context, figure, layout);

    if (figure.layout.legend.visible) {
      this.renderLegendControls(allTraces, layout);
    }
  }

  resize(layout: ComputedLayout): void {
    const canvas = this.ensureCanvas();
    canvas.width = Math.max(1, Math.round(layout.width));
    canvas.height = Math.max(1, Math.round(layout.height));
  }

  getRootElement(): Element | null {
    return this.interactionLayer;
  }

  exportSvg(): string {
    if (!this.lastFigure || !this.lastLayout) {
      throw new ChartConfigurationError("CHART_EXPORT_UNAVAILABLE", "No rendered figure available for SVG export.");
    }
    const surrogate = document.createElement("div");
    Object.assign(surrogate.style, {
      width: `${this.lastLayout.width}px`,
      height: `${this.lastLayout.height}px`,
      position: "absolute",
      left: "-99999px",
      top: "-99999px"
    });
    document.body.append(surrogate);

    const renderer = new SvgRenderer();
    renderer.mount(surrogate);
    renderer.render(this.lastFigure, this.lastLayout);
    const serialized = renderer.exportSvg();
    renderer.destroy();
    surrogate.remove();
    return serialized;
  }

  async exportPng(options?: { scale?: number; backgroundColor?: string }): Promise<Blob> {
    const source = this.ensureCanvas();
    const scale = Math.max(1, options?.scale ?? 1);
    const outputCanvas = document.createElement("canvas");
    outputCanvas.width = Math.max(1, Math.round(source.width * scale));
    outputCanvas.height = Math.max(1, Math.round(source.height * scale));
    const context = outputCanvas.getContext("2d");
    if (!context) {
      throw new ChartConfigurationError("CHART_EXPORT_UNAVAILABLE", "Canvas 2D context is unavailable for PNG export.");
    }

    if (options?.backgroundColor) {
      context.fillStyle = options.backgroundColor;
      context.fillRect(0, 0, outputCanvas.width, outputCanvas.height);
    }
    context.drawImage(source, 0, 0, outputCanvas.width, outputCanvas.height);

    return await new Promise<Blob>((resolve, reject) => {
      outputCanvas.toBlob((blob) => {
        if (!blob) {
          reject(new ChartConfigurationError("CHART_EXPORT_FAILED", "Failed to export canvas as PNG."));
          return;
        }
        resolve(blob);
      }, "image/png");
    });
  }

  destroy(): void {
    if (this.rootElement) {
      this.rootElement.remove();
    }
    this.rootElement = null;
    this.canvasElement = null;
    this.interactionLayer = null;
    this.context = null;
    this.lastFigure = null;
    this.lastLayout = null;
    this.imageCache.clear();
    this.container = null;
  }

  private renderCartesianScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traces: IndexedTrace<CartesianTrace>[],
    plotArea: ComputedPlotArea,
    maxRenderPoints: number,
    subplotIndex?: number
  ): void {
    const primaryYTraces = traces.filter((entry) => entry.trace.yAxisRef !== "y2");
    const secondaryYTraces = traces.filter((entry) => entry.trace.yAxisRef === "y2");
    const primaryXTraces = traces.filter((entry) => entry.trace.xAxisRef !== "x2");
    const secondaryXTraces = traces.filter((entry) => entry.trace.xAxisRef === "x2");
    const domains = this.resolveCartesianDomains(figure, traces, {
      xAxisKey: "xAxis",
      yAxisKey: "yAxis",
      subplotIndex
    });
    if (!domains) {
      return;
    }
    const primaryXDomains =
      this.resolveCartesianDomains(figure, primaryXTraces.length > 0 ? primaryXTraces : traces, {
        xAxisKey: "xAxis",
        yAxisKey: "yAxis",
        subplotIndex
      }) ?? domains;
    const secondaryXDomains =
      secondaryXTraces.length > 0
        ? this.resolveCartesianDomains(figure, secondaryXTraces, {
            xAxisKey: "xAxis2",
            yAxisKey: "yAxis",
            subplotIndex
          })
        : null;
    const primaryYDomains =
      this.resolveCartesianDomains(figure, primaryYTraces.length > 0 ? primaryYTraces : traces, {
        xAxisKey: "xAxis",
        yAxisKey: "yAxis",
        subplotIndex
      }) ?? domains;
    const secondaryYDomains =
      secondaryYTraces.length > 0
        ? this.resolveCartesianDomains(figure, secondaryYTraces, {
            xAxisKey: "xAxis",
            yAxisKey: "yAxis2",
            subplotIndex
          })
        : null;
    const xAxisType = normalizeAxisType(figure.layout.xAxis, "category");
    const xAxis2Type = normalizeAxisType(figure.layout.xAxis2, xAxisType);
    const yAxisType = normalizeAxisType(figure.layout.yAxis, "linear");
    const yAxis2Type = normalizeAxisType(figure.layout.yAxis2, yAxisType);
    const xScale = new LinearScale(primaryXDomains.x, [plotArea.x, plotArea.x + plotArea.width], {
      type: xAxisType === "log" ? "log" : "linear",
      reverse: figure.layout.xAxis.reverse === true
    });
    const xScaleSecondary = secondaryXDomains
      ? new LinearScale(secondaryXDomains.x, [plotArea.x, plotArea.x + plotArea.width], {
          type: xAxis2Type === "log" ? "log" : "linear",
          reverse: figure.layout.xAxis2.reverse === true
        })
      : null;
    const yScale = new LinearScale(primaryYDomains.y, [plotArea.y + plotArea.height, plotArea.y], {
      type: yAxisType === "log" ? "log" : "linear",
      reverse: figure.layout.yAxis.reverse === true
    });
    const yScaleSecondary = secondaryYDomains
      ? new LinearScale(secondaryYDomains.y, [plotArea.y + plotArea.height, plotArea.y], {
          type: yAxis2Type === "log" ? "log" : "linear",
          reverse: figure.layout.yAxis2.reverse === true
        })
      : null;

    this.drawPlotArea(context, plotArea);
    this.drawGrid(context, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });

    const bars = traces.filter((entry) => entry.trace.type === "bar");
    const unstackedBars = bars.filter((entry) => !entry.trace.stackGroup);
    const barCount = Math.max(1, unstackedBars.length);
    const stackContext = buildCartesianStackContext(traces);
    const previousAreaByGroup = new Map<string, Array<{ x: number; y: number }>>();
    const domainSpan = Math.max(1, Math.abs(xScale.invert(plotArea.x + plotArea.width) - xScale.invert(plotArea.x)));
    const slotWidth = plotArea.width / (domainSpan + 1);
    const barSlotWidth = Math.max(4, slotWidth * 0.72);
    const barWidth = Math.max(2, barSlotWidth / barCount);

    traces.forEach((entry) => {
      const color = this.resolveTraceColor(entry.trace, entry.index);
      const stackSeries = stackContext.get(entry.index);
      const targetXScale = entry.trace.xAxisRef === "x2" && xScaleSecondary ? xScaleSecondary : xScale;
      const targetXAxisType = entry.trace.xAxisRef === "x2" && xScaleSecondary ? xAxis2Type : xAxisType;
      if (entry.trace.type === "bar") {
        const barIndex = unstackedBars.findIndex((candidate) => candidate.index === entry.index);
        const targetYScale = entry.trace.yAxisRef === "y2" && yScaleSecondary ? yScaleSecondary : yScale;
        const orientation = entry.trace.orientation ?? "vertical";
        const errorValues = entry.trace.errorY?.values;
        const isStacked = stackSeries?.stacked === true && !!entry.trace.stackGroup;
        entry.trace.y.forEach((value, pointIndex) => {
          if (!Number.isFinite(value)) {
            return;
          }
          const xValue = toAxisScalar(entry.trace.x[pointIndex], targetXAxisType, pointIndex);
          const centerX = targetXScale.map(xValue);
          const activeBarWidth = isStacked ? barSlotWidth : barWidth;
          const x = centerX - barSlotWidth / 2 + (isStacked ? 0 : activeBarWidth * Math.max(0, barIndex));
          const stackedPoint = stackSeries?.points[pointIndex];
          const topValue = stackedPoint ? stackedPoint.top : value;
          const baseValue = stackedPoint ? stackedPoint.base : 0;
          const y = targetYScale.map(topValue);
          const baseY = targetYScale.map(baseValue);
          const top = Math.min(y, baseY);
          const height = Math.max(1, Math.abs(baseY - y));

          context.fillStyle = resolveMarkerColor(entry.trace.marker?.color, pointIndex, color);
          context.globalAlpha = entry.trace.marker?.opacity ?? 0.9;
          if (orientation === "horizontal") {
            const baseX = targetXScale.map(0);
            const width = Math.max(1, Math.abs(centerX - baseX));
            context.fillRect(Math.min(centerX, baseX), y - activeBarWidth / 2, width, activeBarWidth);
          } else {
            context.fillRect(x, top, activeBarWidth, height);
          }
          context.globalAlpha = 1;
          this.addInteractiveTarget(
            {
              traceIndex: entry.index,
              pointIndex,
              traceName: entry.trace.name,
              x: String(entry.trace.x[pointIndex] ?? pointIndex),
              y: formatNumeric(value)
            },
            orientation === "horizontal" ? Math.min(centerX, targetXScale.map(0)) : x,
            orientation === "horizontal" ? y - activeBarWidth / 2 : top,
            orientation === "horizontal" ? Math.max(1, Math.abs(centerX - targetXScale.map(0))) : activeBarWidth,
            orientation === "horizontal" ? activeBarWidth : height
          );

          const error = Number(errorValues?.[pointIndex]);
          if (Number.isFinite(error) && error > 0 && orientation !== "horizontal") {
            const yTop = targetYScale.map(topValue + error);
            const yBottom = targetYScale.map(topValue - error);
            context.strokeStyle = entry.trace.line?.color ?? "#334155";
            context.lineWidth = entry.trace.line?.width ?? 1.2;
            context.globalAlpha = entry.trace.line?.opacity ?? 0.9;
            context.beginPath();
            context.moveTo(centerX, yTop);
            context.lineTo(centerX, yBottom);
            context.stroke();
            context.globalAlpha = 1;
          }
        });
      } else {
        const mode = this.resolveTraceMode(entry.trace);
        const targetYScale = entry.trace.yAxisRef === "y2" && yScaleSecondary ? yScaleSecondary : yScale;
        const ySeries = stackSeries?.stacked ? stackSeries.points.map((point) => point.top) : entry.trace.y;
        const sampleStep = ySeries.length > maxRenderPoints ? Math.ceil(ySeries.length / maxRenderPoints) : 1;
        const points: Array<{ value: number; pointIndex: number; x: number; y: number }> = [];
        for (let pointIndex = 0; pointIndex < ySeries.length; pointIndex += sampleStep) {
          const value = ySeries[pointIndex];
          if (!Number.isFinite(value)) {
            continue;
          }
          points.push({
            value,
            pointIndex,
            x: targetXScale.map(toAxisScalar(entry.trace.x[pointIndex], targetXAxisType, pointIndex)),
            y: targetYScale.map(value)
          });
        }

        if (entry.trace.type === "area" && points.length >= 2) {
          const areaKey = entry.trace.stackGroup?.trim() || "default";
          const previousArea = previousAreaByGroup.get(areaKey);
          const baseline = stackSeries?.stacked
            ? stackSeries.points.map((point) => targetYScale.map(point.base))
            : previousArea && entry.trace.fill === "tonext" && previousArea.length === points.length
              ? previousArea.map((point) => point.y)
              : points.map(() => targetYScale.map(0));
          context.beginPath();
          points.forEach((point, index) => {
            if (index === 0) {
              context.moveTo(point.x, point.y);
            } else {
              context.lineTo(point.x, point.y);
            }
          });
          for (let reverse = points.length - 1; reverse >= 0; reverse -= 1) {
            context.lineTo(points[reverse].x, baseline[reverse] ?? baseline[baseline.length - 1] ?? targetYScale.map(0));
          }
          context.closePath();
          context.fillStyle = resolveColor(color, 0.24);
          context.fill();
          previousAreaByGroup.set(areaKey, points.map((point) => ({ x: point.x, y: point.y })));
        }

        if (mode.includes("lines") && points.length > 1) {
          const optimized = simplifyLine(
            clipLineToRect(
              points.map((point) => ({ x: point.x, y: point.y })),
              plotArea
            ),
            0.9
          );
          if (optimized.length >= 2) {
            context.strokeStyle = entry.trace.line?.color ?? color;
            context.lineWidth = entry.trace.line?.width ?? 2;
            context.globalAlpha = entry.trace.line?.opacity ?? 1;
            context.beginPath();
            optimized.forEach((point, index) => {
              if (index === 0) {
                context.moveTo(point.x, point.y);
              } else {
                context.lineTo(point.x, point.y);
              }
            });
            context.stroke();
            context.globalAlpha = 1;
          }
        }

        points.forEach((point) => {
          if (mode.includes("markers")) {
            context.beginPath();
            context.fillStyle = resolveMarkerColor(entry.trace.marker?.color, point.pointIndex, color);
            context.globalAlpha = entry.trace.marker?.opacity ?? 1;
            context.arc(point.x, point.y, resolveMarkerSize(entry.trace.marker?.size, point.pointIndex, 4), 0, Math.PI * 2);
            context.fill();
            context.globalAlpha = 1;
          }
          this.addInteractiveTarget(
            {
              traceIndex: entry.index,
              pointIndex: point.pointIndex,
              traceName: entry.trace.name,
              x: String(entry.trace.x[point.pointIndex] ?? point.pointIndex),
              y: formatNumeric(point.value)
            },
            point.x - 6,
            point.y - 6,
            12,
            12
          );
        });
      }
    });

    this.drawAxes(
      context,
      figure,
      plotArea,
      primaryXDomains.labels,
      xScale,
      yScale,
      new CategoryScale(primaryXDomains.labels, [plotArea.x, plotArea.x + plotArea.width]),
      yScaleSecondary,
      xScaleSecondary,
      secondaryXDomains?.labels,
      secondaryXDomains ? new CategoryScale(secondaryXDomains.labels, [plotArea.x, plotArea.x + plotArea.width]) : undefined
    );
  }

  private renderHistogramScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traces: IndexedTrace<HistogramTrace>[],
    plotArea: ComputedPlotArea
  ): void {
    const targetBins = Math.max(...traces.map((entry) => clampInt(entry.trace.bins ?? Math.ceil(Math.sqrt(entry.trace.values.length)), 3, 64)));
    const entries = traces.map((entry) => ({
      entry,
      bins: computeHistogramBins({
        ...entry.trace,
        bins: targetBins
      })
    }));
    entries.forEach((entry) => {
      const total = Math.max(1, entry.bins.reduce((sum, bin) => sum + bin.count, 0));
      if (entry.entry.trace.normalize) {
        entry.bins.forEach((bin) => {
          bin.count = bin.count / total;
        });
      }
      if (entry.entry.trace.cumulative) {
        let carry = 0;
        entry.bins.forEach((bin) => {
          carry += bin.count;
          bin.count = carry;
        });
      }
    });
    const labels = entries[0]?.bins.map((bin) => bin.label) ?? [];
    const horizontal = entries.some((entry) => entry.entry.trace.orientation === "horizontal");
    const maxCount = Math.max(1, ...entries.flatMap((entry) => entry.bins.map((bin) => bin.count)));
    const xScale = horizontal
      ? new LinearScale([0, maxCount * 1.1], [plotArea.x, plotArea.x + plotArea.width])
      : new LinearScale([0, Math.max(0, targetBins - 1)], [plotArea.x, plotArea.x + plotArea.width]);
    const yScale = horizontal
      ? new LinearScale([0, Math.max(0, targetBins - 1)], [plotArea.y, plotArea.y + plotArea.height])
      : new LinearScale([0, maxCount * 1.1], [plotArea.y + plotArea.height, plotArea.y]);

    this.drawPlotArea(context, plotArea);
    this.drawGrid(context, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });

    const barCount = Math.max(1, entries.length);
    const slotWidth = plotArea.width / Math.max(1, targetBins);
    const barSlotWidth = Math.max(6, slotWidth * 0.82);
    const barWidth = Math.max(2, barSlotWidth / barCount);

    entries.forEach((entryData, tracePosition) => {
      const color = this.resolveTraceColor(entryData.entry.trace, entryData.entry.index);
      entryData.bins.forEach((bin) => {
        context.fillStyle = color;
        context.globalAlpha = entryData.entry.trace.marker?.opacity ?? 0.9;
        if (horizontal) {
          const centerY = yScale.map(bin.index);
          const y = centerY - barSlotWidth / 2 + barWidth * tracePosition;
          const x0 = xScale.map(0);
          const x1 = xScale.map(bin.count);
          context.fillRect(Math.min(x0, x1), y, Math.max(1, Math.abs(x1 - x0)), barWidth);
          this.addInteractiveTarget(
            {
              traceIndex: entryData.entry.index,
              pointIndex: bin.index,
              traceName: entryData.entry.trace.name,
              x: bin.label,
              y: formatNumeric(bin.count)
            },
            Math.min(x0, x1),
            y,
            Math.max(1, Math.abs(x1 - x0)),
            barWidth
          );
        } else {
          const centerX = xScale.map(bin.index);
          const x = centerX - barSlotWidth / 2 + barWidth * tracePosition;
          const y = yScale.map(bin.count);
          const baseY = yScale.map(0);
          const top = Math.min(y, baseY);
          const height = Math.max(1, Math.abs(baseY - y));
          context.fillRect(x, top, barWidth, height);
          this.addInteractiveTarget(
            {
              traceIndex: entryData.entry.index,
              pointIndex: bin.index,
              traceName: entryData.entry.trace.name,
              x: bin.label,
              y: formatNumeric(bin.count)
            },
            x,
            top,
            barWidth,
            height
          );
        }
        context.globalAlpha = 1;
      });
    });

    this.drawAxes(context, figure, plotArea, labels, xScale, yScale, new CategoryScale(labels, [plotArea.x, plotArea.x + plotArea.width]));
  }

  private renderViolinScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traces: IndexedTrace<ViolinTrace>[],
    plotArea: ComputedPlotArea
  ): void {
    if (traces.length === 0) {
      return;
    }
    const labels = traces.map((entry, index) => entry.trace.name?.trim() || `Violin ${index + 1}`);
    const xScale = new LinearScale([0, Math.max(0, traces.length - 1)], [plotArea.x, plotArea.x + plotArea.width]);
    const allValues = traces.flatMap((entry) => entry.trace.values.map((value) => Number(value)).filter((value) => Number.isFinite(value)));
    if (allValues.length === 0) {
      return;
    }
    const yMin = Math.min(...allValues);
    const yMax = Math.max(...allValues);
    const yScale = new LinearScale([yMin, yMin === yMax ? yMax + 1 : yMax], [plotArea.y + plotArea.height, plotArea.y]);
    this.drawPlotArea(context, plotArea);
    this.drawGrid(context, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const slotWidth = plotArea.width / Math.max(1, traces.length);

    traces.forEach((entry, traceIndex) => {
      const profile = computeViolinProfile(entry.trace.values, { bandwidth: entry.trace.bandwidth });
      if (!profile || profile.points.length < 2) {
        return;
      }
      const centerX = xScale.map(traceIndex);
      const halfWidth = Math.max(6, slotWidth * 0.34);
      const color = this.resolveTraceColor(entry.trace, entry.index);
      context.beginPath();
      profile.points.forEach((point, index) => {
        const y = yScale.map(point.x);
        const spread = point.y * halfWidth;
        if (index === 0) {
          context.moveTo(centerX - spread, y);
        } else {
          context.lineTo(centerX - spread, y);
        }
      });
      for (let index = profile.points.length - 1; index >= 0; index -= 1) {
        const point = profile.points[index];
        const y = yScale.map(point.x);
        const spread = point.y * halfWidth;
        context.lineTo(centerX + spread, y);
      }
      context.closePath();
      context.fillStyle = resolveColor(color, 0.26);
      context.fill();
      context.strokeStyle = entry.trace.line?.color ?? color;
      context.lineWidth = entry.trace.line?.width ?? 1.6;
      context.globalAlpha = entry.trace.marker?.opacity ?? entry.trace.line?.opacity ?? 1;
      context.stroke();
      context.globalAlpha = 1;

      if (entry.trace.showBox) {
        const stats = computeBoxStats({ ...entry.trace, type: "box" });
        if (stats) {
          const boxWidth = Math.max(6, halfWidth * 0.45);
          const y0 = yScale.map(stats.q1);
          const y1 = yScale.map(stats.q3);
          context.fillStyle = "#ffffff";
          context.strokeStyle = entry.trace.line?.color ?? color;
          context.fillRect(centerX - boxWidth / 2, Math.min(y0, y1), boxWidth, Math.max(1, Math.abs(y1 - y0)));
          context.strokeRect(centerX - boxWidth / 2, Math.min(y0, y1), boxWidth, Math.max(1, Math.abs(y1 - y0)));
          const medianY = yScale.map(stats.median);
          context.beginPath();
          context.moveTo(centerX - boxWidth / 2, medianY);
          context.lineTo(centerX + boxWidth / 2, medianY);
          context.strokeStyle = "#0f172a";
          context.stroke();
        }
      }
    });

    this.drawAxes(context, figure, plotArea, labels, xScale, yScale, new CategoryScale(labels, [plotArea.x, plotArea.x + plotArea.width]));
  }

  private renderDensityScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traces: IndexedTrace<DensityTrace>[],
    plotArea: ComputedPlotArea
  ): void {
    const curves = traces
      .map((entry) => {
        const cumulative = entry.trace.cumulative === true || entry.trace.type === "distribution";
        return {
          entry,
          points: computeDensityCurve(entry.trace.values, {
            cumulative,
            sampleLimit: entry.trace.sampleLimit ?? figure.config.maxDensitySamples
          })
        };
      })
      .filter((curve) => curve.points.length >= 2);
    if (curves.length === 0) {
      return;
    }
    const xMin = Math.min(...curves.flatMap((curve) => curve.points.map((point) => point.x)));
    const xMax = Math.max(...curves.flatMap((curve) => curve.points.map((point) => point.x)));
    const yMax = Math.max(...curves.flatMap((curve) => curve.points.map((point) => point.y)), 1);
    const xScale = new LinearScale([xMin, xMin === xMax ? xMax + 1 : xMax], [plotArea.x, plotArea.x + plotArea.width]);
    const yScale = new LinearScale([0, yMax * 1.05], [plotArea.y + plotArea.height, plotArea.y]);
    this.drawPlotArea(context, plotArea);
    this.drawGrid(context, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });

    curves.forEach((curve) => {
      const color = this.resolveTraceColor(curve.entry.trace, curve.entry.index);
      context.strokeStyle = curve.entry.trace.line?.color ?? color;
      context.lineWidth = curve.entry.trace.line?.width ?? 2;
      context.globalAlpha = curve.entry.trace.line?.opacity ?? 0.95;
      context.beginPath();
      curve.points.forEach((point, index) => {
        const x = xScale.map(point.x);
        const y = yScale.map(point.y);
        if (index === 0) {
          context.moveTo(x, y);
        } else {
          context.lineTo(x, y);
        }
      });
      context.stroke();
      context.globalAlpha = 1;
    });

    this.drawAxes(context, figure, plotArea, [], xScale, yScale, new CategoryScale([], [plotArea.x, plotArea.x + plotArea.width]));
  }

  private renderBoxScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traces: IndexedTrace<BoxTrace>[],
    plotArea: ComputedPlotArea
  ): void {
    const statsEntries = traces
      .map((entry) => ({ entry, stats: computeBoxStats(entry.trace) }))
      .filter((entry): entry is { entry: IndexedTrace<BoxTrace>; stats: NonNullable<ReturnType<typeof computeBoxStats>> } => entry.stats !== null);

    if (statsEntries.length === 0) {
      return;
    }

    const labels = statsEntries.map((entry, index) => entry.entry.trace.name?.trim() || `Box ${index + 1}`);
    const xScale = new LinearScale([0, Math.max(0, statsEntries.length - 1)], [plotArea.x, plotArea.x + plotArea.width]);
    const values = statsEntries.flatMap((entry) => [entry.stats.min, entry.stats.q1, entry.stats.median, entry.stats.q3, entry.stats.max]);
    const minValue = Math.min(...values);
    const maxValue = Math.max(...values);
    const yScale = new LinearScale(
      [Math.min(0, minValue), maxValue === minValue ? maxValue + 1 : maxValue],
      [plotArea.y + plotArea.height, plotArea.y]
    );

    this.drawPlotArea(context, plotArea);
    this.drawGrid(context, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });

    const slotWidth = plotArea.width / Math.max(1, statsEntries.length);
    const boxWidth = Math.max(12, Math.min(48, slotWidth * 0.58));

    statsEntries.forEach(({ entry, stats }, tracePosition) => {
      const color = this.resolveTraceColor(entry.trace, entry.index);
      const centerX = xScale.map(tracePosition);
      const yMin = yScale.map(stats.min);
      const yQ1 = yScale.map(stats.q1);
      const yMedian = yScale.map(stats.median);
      const yQ3 = yScale.map(stats.q3);
      const yMax = yScale.map(stats.max);

      context.strokeStyle = entry.trace.line?.color ?? color;
      context.lineWidth = entry.trace.line?.width ?? 2;
      context.beginPath();
      context.moveTo(centerX, yMin);
      context.lineTo(centerX, yMax);
      context.stroke();

      const boxTop = Math.min(yQ1, yQ3);
      const boxHeight = Math.max(1, Math.abs(yQ3 - yQ1));
      context.fillStyle = color;
      context.globalAlpha = entry.trace.marker?.opacity ?? 0.48;
      context.fillRect(centerX - boxWidth / 2, boxTop, boxWidth, boxHeight);
      context.globalAlpha = 1;
      context.strokeStyle = entry.trace.line?.color ?? color;
      context.strokeRect(centerX - boxWidth / 2, boxTop, boxWidth, boxHeight);

      context.strokeStyle = entry.trace.line?.color ?? "#0f172a";
      context.beginPath();
      context.moveTo(centerX - boxWidth / 2, yMedian);
      context.lineTo(centerX + boxWidth / 2, yMedian);
      context.stroke();

      this.addInteractiveTarget(
        {
          traceIndex: entry.index,
          pointIndex: 0,
          traceName: entry.trace.name,
          x: labels[tracePosition],
          y: `mediana ${formatNumeric(stats.median)}`
        },
        centerX - boxWidth / 2,
        boxTop,
        boxWidth,
        boxHeight
      );

      if (entry.trace.showOutliers) {
        const factor = Number(entry.trace.outlierIqrFactor ?? 1.5);
        const outliers = computeBoxOutliers(entry.trace, Number.isFinite(factor) && factor > 0 ? factor : 1.5);
        outliers.forEach((outlier, pointIndex) => {
          const y = yScale.map(outlier);
          const jitter = ((pointIndex % 5) - 2) * 2.4;
          context.strokeStyle = entry.trace.line?.color ?? color;
          context.lineWidth = 1.2;
          context.beginPath();
          context.arc(centerX + jitter, y, 2.8, 0, Math.PI * 2);
          context.stroke();
          this.addInteractiveTarget(
            {
              traceIndex: entry.index,
              pointIndex,
              traceName: entry.trace.name,
              x: labels[tracePosition],
              y: `outlier ${formatNumeric(outlier)}`
            },
            centerX + jitter - 4,
            y - 4,
            8,
            8
          );
        });
      }
    });

    this.drawAxes(context, figure, plotArea, labels, xScale, yScale, new CategoryScale(labels, [plotArea.x, plotArea.x + plotArea.width]));
  }

  private renderFinancialScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traces: IndexedTrace<FinancialTrace>[],
    plotArea: ComputedPlotArea,
    subplotIndex?: number
  ): void {
    const primaryYTraces = traces.filter((entry) => entry.trace.yAxisRef !== "y2");
    const secondaryYTraces = traces.filter((entry) => entry.trace.yAxisRef === "y2");
    const primaryXTraces = traces.filter((entry) => entry.trace.xAxisRef !== "x2");
    const secondaryXTraces = traces.filter((entry) => entry.trace.xAxisRef === "x2");
    const domains = this.resolveFinancialDomains(figure, traces, {
      xAxisKey: "xAxis",
      yAxisKey: "yAxis",
      subplotIndex
    });
    if (!domains) {
      return;
    }
    const primaryXDomains =
      this.resolveFinancialDomains(figure, primaryXTraces.length > 0 ? primaryXTraces : traces, {
        xAxisKey: "xAxis",
        yAxisKey: "yAxis",
        subplotIndex
      }) ?? domains;
    const secondaryXDomains =
      secondaryXTraces.length > 0
        ? this.resolveFinancialDomains(figure, secondaryXTraces, {
            xAxisKey: "xAxis2",
            yAxisKey: "yAxis",
            subplotIndex
          })
        : null;
    const primaryYDomains =
      this.resolveFinancialDomains(figure, primaryYTraces.length > 0 ? primaryYTraces : traces, {
        xAxisKey: "xAxis",
        yAxisKey: "yAxis",
        subplotIndex
      }) ?? domains;
    const secondaryYDomains =
      secondaryYTraces.length > 0
        ? this.resolveFinancialDomains(figure, secondaryYTraces, {
            xAxisKey: "xAxis",
            yAxisKey: "yAxis2",
            subplotIndex
          })
        : null;

    const xAxisType = normalizeAxisType(figure.layout.xAxis, "category");
    const xAxis2Type = normalizeAxisType(figure.layout.xAxis2, xAxisType);
    const yAxisType = normalizeAxisType(figure.layout.yAxis, "linear");
    const yAxis2Type = normalizeAxisType(figure.layout.yAxis2, yAxisType);
    const xScale = new LinearScale(primaryXDomains.x, [plotArea.x, plotArea.x + plotArea.width], {
      type: xAxisType === "log" ? "log" : "linear",
      reverse: figure.layout.xAxis.reverse === true
    });
    const xScaleSecondary = secondaryXDomains
      ? new LinearScale(secondaryXDomains.x, [plotArea.x, plotArea.x + plotArea.width], {
          type: xAxis2Type === "log" ? "log" : "linear",
          reverse: figure.layout.xAxis2.reverse === true
        })
      : null;
    const yScale = new LinearScale(primaryYDomains.y, [plotArea.y + plotArea.height, plotArea.y], {
      type: yAxisType === "log" ? "log" : "linear",
      reverse: figure.layout.yAxis.reverse === true
    });
    const yScaleSecondary = secondaryYDomains
      ? new LinearScale(secondaryYDomains.y, [plotArea.y + plotArea.height, plotArea.y], {
          type: yAxis2Type === "log" ? "log" : "linear",
          reverse: figure.layout.yAxis2.reverse === true
        })
      : null;
    const categoryScale = new CategoryScale(primaryXDomains.labels, [plotArea.x, plotArea.x + plotArea.width]);
    const categoryScaleSecondary = secondaryXDomains
      ? new CategoryScale(secondaryXDomains.labels, [plotArea.x, plotArea.x + plotArea.width])
      : null;

    this.drawPlotArea(context, plotArea);
    this.drawGrid(context, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });

    const maxPoints = Math.max(...traces.map((entry) => entry.trace.x.length));
    const slotWidth = plotArea.width / Math.max(1, maxPoints);
    const bodyWidth = Math.max(2, Math.min(22, (slotWidth * 0.66) / Math.max(1, traces.length)));

    traces.forEach((entry, tracePosition) => {
      const offset = (tracePosition - (traces.length - 1) / 2) * bodyWidth;
      const lineWidth = entry.trace.line?.width ?? 1.5;
      const lineOpacity = entry.trace.line?.opacity ?? 1;

      for (let pointIndex = 0; pointIndex < entry.trace.x.length; pointIndex += 1) {
        const open = Number(entry.trace.open[pointIndex]);
        const high = Number(entry.trace.high[pointIndex]);
        const low = Number(entry.trace.low[pointIndex]);
        const close = Number(entry.trace.close[pointIndex]);
        if (!Number.isFinite(open) || !Number.isFinite(high) || !Number.isFinite(low) || !Number.isFinite(close)) {
          continue;
        }

        const increasing = close >= open;
        const upColor = entry.trace.increasing?.color ?? "#16a34a";
        const downColor = entry.trace.decreasing?.color ?? "#dc2626";
        const strokeColor = entry.trace.line?.color ?? (increasing ? upColor : downColor);
        const targetXScale = entry.trace.xAxisRef === "x2" && xScaleSecondary ? xScaleSecondary : xScale;
        const targetXAxisType = entry.trace.xAxisRef === "x2" && xScaleSecondary ? xAxis2Type : xAxisType;
        const xValue = toAxisScalar(entry.trace.x[pointIndex], targetXAxisType, pointIndex);
        const centerX = targetXScale.map(xValue) + offset;
        const targetYScale = entry.trace.yAxisRef === "y2" && yScaleSecondary ? yScaleSecondary : yScale;
        const yOpen = targetYScale.map(open);
        const yClose = targetYScale.map(close);
        const yHigh = targetYScale.map(high);
        const yLow = targetYScale.map(low);

        context.strokeStyle = strokeColor;
        context.globalAlpha = lineOpacity;
        context.lineWidth = lineWidth;
        context.beginPath();
        context.moveTo(centerX, yHigh);
        context.lineTo(centerX, yLow);
        context.stroke();

        if (entry.trace.type === "candlestick") {
          const top = Math.min(yOpen, yClose);
          const height = Math.max(1, Math.abs(yClose - yOpen));
          context.fillStyle = increasing ? upColor : downColor;
          context.fillRect(centerX - bodyWidth / 2, top, bodyWidth, height);
          context.strokeRect(centerX - bodyWidth / 2, top, bodyWidth, height);
          this.addInteractiveTarget(
            {
              traceIndex: entry.index,
              pointIndex,
              traceName: entry.trace.name,
              x: String(entry.trace.x[pointIndex] ?? pointIndex),
              y: `O ${formatNumeric(open)} H ${formatNumeric(high)} L ${formatNumeric(low)} C ${formatNumeric(close)}`
            },
            centerX - bodyWidth / 2,
            top,
            bodyWidth,
            height
          );
        } else {
          context.beginPath();
          context.moveTo(centerX - bodyWidth / 2, yOpen);
          context.lineTo(centerX, yOpen);
          context.moveTo(centerX, yClose);
          context.lineTo(centerX + bodyWidth / 2, yClose);
          context.stroke();
          this.addInteractiveTarget(
            {
              traceIndex: entry.index,
              pointIndex,
              traceName: entry.trace.name,
              x: String(entry.trace.x[pointIndex] ?? pointIndex),
              y: `O ${formatNumeric(open)} H ${formatNumeric(high)} L ${formatNumeric(low)} C ${formatNumeric(close)}`
            },
            centerX - bodyWidth / 2,
            Math.min(yHigh, yLow),
            bodyWidth,
            Math.max(2, Math.abs(yLow - yHigh))
          );
        }
      }
      context.globalAlpha = 1;
    });

    this.drawAxes(
      context,
      figure,
      plotArea,
      primaryXDomains.labels,
      xScale,
      yScale,
      categoryScale,
      yScaleSecondary,
      xScaleSecondary,
      secondaryXDomains?.labels,
      categoryScaleSecondary ?? undefined
    );
  }

  private renderHeatmapScene(context: CanvasRenderingContext2D, traceEntry: IndexedTrace<HeatmapTrace>, plotArea: ComputedPlotArea): void {
    const matrix = computeHeatmapMatrix(traceEntry.trace);
    if (!matrix) {
      return;
    }

    this.drawPlotArea(context, plotArea);
    const xLabels = Array.from({ length: matrix.cols }, (_, col) => String(traceEntry.trace.x?.[col] ?? `C${col + 1}`));
    const yLabels = Array.from({ length: matrix.rows }, (_, row) => String(traceEntry.trace.y?.[row] ?? `R${row + 1}`));
    const cellWidth = plotArea.width / matrix.cols;
    const cellHeight = plotArea.height / matrix.rows;
    const colors = resolvePalette("continuous", traceEntry.trace.colorscale, false);

    matrix.values.forEach((cell) => {
      const color = resolveColorFromScale(cell.value, {
        mode: "continuous",
        colors,
        min: matrix.min,
        max: matrix.max
      });
      const x = plotArea.x + cell.col * cellWidth;
      const y = plotArea.y + cell.row * cellHeight;
      context.fillStyle = color;
      context.fillRect(x, y, Math.max(1, cellWidth), Math.max(1, cellHeight));
      this.addInteractiveTarget(
        {
          traceIndex: traceEntry.index,
          pointIndex: cell.row * matrix.cols + cell.col,
          traceName: traceEntry.trace.name,
          x: `${xLabels[cell.col]} / ${yLabels[cell.row]}`,
          y: formatNumeric(cell.value)
        },
        x,
        y,
        Math.max(1, cellWidth),
        Math.max(1, cellHeight)
      );
    });
  }

  private renderContourScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traceEntry: IndexedTrace<ContourTrace>,
    plotArea: ComputedPlotArea
  ): void {
    const matrix = traceEntry.trace.z;
    if (matrix.length === 0 || matrix[0]?.length === 0) {
      return;
    }
    this.drawPlotArea(context, plotArea);
    const rows = matrix.length;
    const cols = matrix[0].length;
    const xLabels = Array.from({ length: cols }, (_, col) => String(traceEntry.trace.x?.[col] ?? `C${col + 1}`));
    const xScale = new LinearScale([0, Math.max(0, cols - 1)], [plotArea.x, plotArea.x + plotArea.width]);
    const yScale = new LinearScale([Math.max(0, rows - 1), 0], [plotArea.y + plotArea.height, plotArea.y]);
    const categoryScale = new CategoryScale(xLabels, [plotArea.x, plotArea.x + plotArea.width]);
    const values = matrix.flat().map((value) => Number(value)).filter((value) => Number.isFinite(value));
    if (values.length === 0) {
      return;
    }
    const min = Math.min(...values);
    const max = Math.max(...values);
    const levels = clampInt(traceEntry.trace.levels ?? 7, 3, 24);
    const palette = resolvePalette("continuous", traceEntry.trace.colorscale, false);
    if (traceEntry.trace.fillContours) {
      const cellWidth = plotArea.width / Math.max(1, cols);
      const cellHeight = plotArea.height / Math.max(1, rows);
      for (let row = 0; row < rows; row += 1) {
        for (let col = 0; col < cols; col += 1) {
          const raw = Number(matrix[row][col]);
          if (!Number.isFinite(raw)) {
            continue;
          }
          context.fillStyle = resolveColorFromScale(raw, {
            mode: "continuous",
            colors: palette,
            min,
            max
          });
          context.globalAlpha = 0.16;
          context.fillRect(plotArea.x + col * cellWidth, plotArea.y + row * cellHeight, Math.max(1, cellWidth), Math.max(1, cellHeight));
          context.globalAlpha = 1;
        }
      }
    }

    for (let levelIndex = 0; levelIndex < levels; levelIndex += 1) {
      const ratio = levels <= 1 ? 0.5 : levelIndex / (levels - 1);
      const level = min + ratio * (max - min);
      const levelColor = resolveColorFromScale(level, {
        mode: "continuous",
        colors: palette,
        min,
        max
      });
      const allSegments = buildContourSegments(matrix, level);
      const maxSegments = clampInt(traceEntry.trace.maxSegments ?? allSegments.length, 10, 100_000);
      const segments = allSegments.slice(0, maxSegments);
      context.strokeStyle = traceEntry.trace.line?.color ?? levelColor;
      context.lineWidth = traceEntry.trace.line?.width ?? 1.4;
      context.globalAlpha = traceEntry.trace.line?.opacity ?? 0.94;
      segments.forEach((segment, segmentIndex) => {
        const startX = plotArea.x + segment.a.x * (plotArea.width / Math.max(1, cols - 1));
        const startY = plotArea.y + segment.a.y * (plotArea.height / Math.max(1, rows - 1));
        const endX = plotArea.x + segment.b.x * (plotArea.width / Math.max(1, cols - 1));
        const endY = plotArea.y + segment.b.y * (plotArea.height / Math.max(1, rows - 1));
        context.beginPath();
        context.moveTo(startX, startY);
        context.lineTo(endX, endY);
        context.stroke();
        this.addInteractiveTarget(
          {
            traceIndex: traceEntry.index,
            pointIndex: levelIndex * 10_000 + segmentIndex,
            traceName: traceEntry.trace.name,
            x: `level ${formatNumeric(level)}`,
            y: traceEntry.trace.line?.color ?? levelColor
          },
          Math.min(startX, endX) - 4,
          Math.min(startY, endY) - 4,
          Math.max(8, Math.abs(endX - startX) + 8),
          Math.max(8, Math.abs(endY - startY) + 8)
        );
      });
      if (traceEntry.trace.labelLevels && segments.length > 0) {
        const sample = segments[Math.floor(segments.length / 2)];
        const x = plotArea.x + ((sample.a.x + sample.b.x) * 0.5) * (plotArea.width / Math.max(1, cols - 1));
        const y = plotArea.y + ((sample.a.y + sample.b.y) * 0.5) * (plotArea.height / Math.max(1, rows - 1));
        context.fillStyle = traceEntry.trace.line?.color ?? levelColor;
        context.font = "10px Arial, sans-serif";
        context.textAlign = "left";
        context.fillText(formatNumeric(level), x + 3, y - 3);
      }
      context.globalAlpha = 1;
    }

    this.drawAxes(context, figure, plotArea, xLabels, xScale, yScale, categoryScale);
  }

  private renderQuiverScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traces: IndexedTrace<QuiverTrace>[],
    plotArea: ComputedPlotArea
  ): void {
    if (traces.length === 0) {
      return;
    }
    const xValues = traces.flatMap((entry) => entry.trace.x.flatMap((value, index) => [Number(value), Number(value) + Number(entry.trace.u[index] ?? 0)]));
    const yValues = traces.flatMap((entry) => entry.trace.y.flatMap((value, index) => [Number(value), Number(value) + Number(entry.trace.v[index] ?? 0)]));
    const finiteX = xValues.filter((value) => Number.isFinite(value));
    const finiteY = yValues.filter((value) => Number.isFinite(value));
    if (finiteX.length === 0 || finiteY.length === 0) {
      return;
    }
    const xScale = new LinearScale([Math.min(...finiteX), Math.max(...finiteX)], [plotArea.x, plotArea.x + plotArea.width]);
    const yScale = new LinearScale([Math.min(...finiteY), Math.max(...finiteY)], [plotArea.y + plotArea.height, plotArea.y]);
    this.drawPlotArea(context, plotArea);
    this.drawGrid(context, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const magnitudes = traces.flatMap((entry) =>
      entry.trace.u.map((uValue, index) => Math.hypot(Number(uValue), Number(entry.trace.v[index])))
    );
    const finiteMagnitudes = magnitudes.filter((value) => Number.isFinite(value));
    const minMagnitude = finiteMagnitudes.length > 0 ? Math.min(...finiteMagnitudes) : 0;
    const maxMagnitude = finiteMagnitudes.length > 0 ? Math.max(...finiteMagnitudes) : 1;
    traces.forEach((entry, traceIndex) => {
      const color = entry.trace.line?.color ?? DEFAULT_SERIES_COLORS[traceIndex % DEFAULT_SERIES_COLORS.length];
      const scale = Number.isFinite(entry.trace.scale) ? Number(entry.trace.scale) : 1;
      context.lineWidth = entry.trace.line?.width ?? 1.2;
      context.globalAlpha = entry.trace.line?.opacity ?? 0.9;
      const magnitudePalette = resolvePalette("continuous", entry.trace.colorscale, false);
      for (let pointIndex = 0; pointIndex < entry.trace.x.length; pointIndex += 1) {
        const x = Number(entry.trace.x[pointIndex]);
        const y = Number(entry.trace.y[pointIndex]);
        const u = Number(entry.trace.u[pointIndex]);
        const v = Number(entry.trace.v[pointIndex]);
        if (!Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(u) || !Number.isFinite(v)) {
          continue;
        }
        const x0 = xScale.map(x);
        const y0 = yScale.map(y);
        const x1 = xScale.map(x + u * scale);
        const y1 = yScale.map(y + v * scale);
        const magnitude = Math.hypot(u, v);
        const pointColor =
          entry.trace.colorByMagnitude === true
            ? resolveColorFromScale(magnitude, {
                mode: "continuous",
                colors: magnitudePalette,
                min: minMagnitude,
                max: maxMagnitude
              })
            : color;
        context.strokeStyle = pointColor;
        context.fillStyle = pointColor;
        context.beginPath();
        context.moveTo(x0, y0);
        context.lineTo(x1, y1);
        context.stroke();

        const angle = Math.atan2(y1 - y0, x1 - x0);
        const len = 7;
        const wing = 0.6;
        const hx1 = x1 - Math.cos(angle - wing) * len;
        const hy1 = y1 - Math.sin(angle - wing) * len;
        const hx2 = x1 - Math.cos(angle + wing) * len;
        const hy2 = y1 - Math.sin(angle + wing) * len;
        context.beginPath();
        context.moveTo(x1, y1);
        context.lineTo(hx1, hy1);
        context.lineTo(hx2, hy2);
        context.closePath();
        context.fill();
        this.addInteractiveTarget(
          {
            traceIndex: entry.index,
            pointIndex,
            traceName: entry.trace.name,
            x: `x=${formatNumeric(x)} y=${formatNumeric(y)}`,
            y: `u=${formatNumeric(u)} v=${formatNumeric(v)} |m|=${formatNumeric(magnitude)}`
          },
          Math.min(x0, x1) - 6,
          Math.min(y0, y1) - 6,
          Math.max(12, Math.abs(x1 - x0) + 12),
          Math.max(12, Math.abs(y1 - y0) + 12)
        );
      }
      context.globalAlpha = 1;
    });
    this.drawAxes(context, figure, plotArea, [], xScale, yScale, new CategoryScale([], [plotArea.x, plotArea.x + plotArea.width]));
  }

  private renderWaterfallScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traces: IndexedTrace<WaterfallTrace>[],
    plotArea: ComputedPlotArea
  ): void {
    const allLabels = traces[0]?.trace.x.map((value) => String(value)) ?? [];
    const axisType = normalizeAxisType(figure.layout.xAxis, "category");
    const xValues = traces
      .flatMap((entry) => entry.trace.x.map((value, index) => toAxisScalar(value, axisType, index)))
      .filter((value) => Number.isFinite(value));
    const xMin = xValues.length > 0 ? Math.min(...xValues) : 0;
    const xMax = xValues.length > 0 ? Math.max(...xValues) : Math.max(1, allLabels.length - 1);
    const yValues: number[] = [];
    traces.forEach((entry) => {
      let running = 0;
      entry.trace.y.forEach((rawValue, index) => {
        const value = Number(rawValue);
        if (!Number.isFinite(value)) {
          return;
        }
        const measure = entry.trace.measure?.[index] ?? "relative";
        if (measure === "absolute" || measure === "total") {
          running = value;
        } else {
          running += value;
        }
        yValues.push(running);
      });
    });
    const yMin = Math.min(0, ...(yValues.length > 0 ? yValues : [0]));
    const yMax = Math.max(0, ...(yValues.length > 0 ? yValues : [1]));
    const xScale = new LinearScale([xMin, xMax], [plotArea.x, plotArea.x + plotArea.width], {
      type: axisType === "log" ? "log" : "linear",
      reverse: figure.layout.xAxis.reverse === true
    });
    const yScale = new LinearScale([yMin, yMax === yMin ? yMin + 1 : yMax], [plotArea.y + plotArea.height, plotArea.y], {
      type: normalizeAxisType(figure.layout.yAxis, "linear") === "log" ? "log" : "linear",
      reverse: figure.layout.yAxis.reverse === true
    });
    const categoryScale = new CategoryScale(allLabels, [plotArea.x, plotArea.x + plotArea.width]);

    this.drawPlotArea(context, plotArea);
    this.drawGrid(context, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });

    traces.forEach((entry, tracePosition) => {
      let running = 0;
      const slotWidth = Math.max(8, plotArea.width / Math.max(1, entry.trace.x.length) * 0.64);
      const offset = (tracePosition - (traces.length - 1) / 2) * Math.min(20, slotWidth / Math.max(1, traces.length));
      entry.trace.y.forEach((rawValue, pointIndex) => {
        const value = Number(rawValue);
        if (!Number.isFinite(value)) {
          return;
        }
        const measure = entry.trace.measure?.[pointIndex] ?? "relative";
        const start = running;
        if (measure === "absolute" || measure === "total") {
          running = value;
        } else {
          running += value;
        }
        const end = running;
        const xValue = toAxisScalar(entry.trace.x[pointIndex], axisType, pointIndex);
        const centerX = xScale.map(xValue) + offset;
        const y0 = yScale.map(start);
        const y1 = yScale.map(end);
        const top = Math.min(y0, y1);
        const height = Math.max(1, Math.abs(y1 - y0));
        const color =
          measure === "total"
            ? entry.trace.totals?.color ?? "#475569"
            : end >= start
              ? entry.trace.increasing?.color ?? "#16a34a"
              : entry.trace.decreasing?.color ?? "#dc2626";
        context.fillStyle = color;
        context.globalAlpha = 0.9;
        context.fillRect(centerX - slotWidth / 2, top, slotWidth, height);
        context.globalAlpha = 1;
        this.addInteractiveTarget(
          {
            traceIndex: entry.index,
            pointIndex,
            traceName: entry.trace.name,
            x: String(entry.trace.x[pointIndex] ?? pointIndex),
            y: formatNumeric(value)
          },
          centerX - slotWidth / 2,
          top,
          slotWidth,
          height
        );
      });
    });

    this.drawAxes(context, figure, plotArea, allLabels, xScale, yScale, categoryScale);
  }

  private renderFunnelScene(context: CanvasRenderingContext2D, traceEntry: IndexedTrace<FunnelTrace>, plotArea: ComputedPlotArea): void {
    this.drawPlotArea(context, plotArea);
    const steps = traceEntry.trace.values
      .map((value, index) => ({
        value: Number(value),
        index,
        label: traceEntry.trace.labels[index] ?? `Step ${index + 1}`
      }))
      .filter((entry) => Number.isFinite(entry.value) && entry.value >= 0);
    if (steps.length === 0) {
      return;
    }
    const sortMode = traceEntry.trace.sort ?? "none";
    if (sortMode === "asc") {
      steps.sort((left, right) => left.value - right.value);
    } else if (sortMode === "desc") {
      steps.sort((left, right) => right.value - left.value);
    }
    const values = steps.map((step) => step.value);
    const maxValue = Math.max(...values, 1);
    const initialValue = Math.max(values[0] ?? 0, Number.EPSILON);
    const count = values.length;
    const stepHeight = plotArea.height / count;
    const colors = traceEntry.trace.marker?.color ?? DEFAULT_SERIES_COLORS;
    const opacity = traceEntry.trace.marker?.opacity ?? 0.9;

    for (let index = 0; index < count; index += 1) {
      const value = values[index];
      const nextValue = values[index + 1] ?? value * 0.75;
      const topWidth = (value / maxValue) * plotArea.width * 0.94;
      const bottomWidth = (nextValue / maxValue) * plotArea.width * 0.94;
      const y0 = plotArea.y + index * stepHeight;
      const y1 = y0 + stepHeight;
      const cx = plotArea.x + plotArea.width / 2;
      context.beginPath();
      context.moveTo(cx - topWidth / 2, y0);
      context.lineTo(cx + topWidth / 2, y0);
      context.lineTo(cx + bottomWidth / 2, y1);
      context.lineTo(cx - bottomWidth / 2, y1);
      context.closePath();
      context.fillStyle = colors[index % colors.length];
      context.globalAlpha = opacity;
      context.fill();
      context.globalAlpha = 1;
      context.strokeStyle = "#ffffff";
      context.lineWidth = 1;
      context.stroke();

      const label = steps[index]?.label ?? `Step ${index + 1}`;
      context.fillStyle = "#0f172a";
      context.textAlign = "center";
      context.font = "11px Arial, sans-serif";
      context.fillText(label, cx, y0 + stepHeight / 2 + 3);
      const previousValue = index > 0 ? values[index - 1] : value;
      const conversionFromStart = (value / initialValue) * 100;
      const conversionFromPrevious = previousValue > 0 ? (value / previousValue) * 100 : 0;
      this.addInteractiveTarget(
        {
          traceIndex: traceEntry.index,
          pointIndex: steps[index]?.index ?? index,
          traceName: traceEntry.trace.name,
          x: label,
          y: `${formatNumeric(value)} (${conversionFromStart.toFixed(1)}% total, ${conversionFromPrevious.toFixed(1)}% etapa anterior)`
        },
        cx - topWidth / 2,
        y0,
        topWidth,
        stepHeight
      );
    }
  }

  private renderPieScene(context: CanvasRenderingContext2D, traceEntry: IndexedTrace<PieTrace>, plotArea: ComputedPlotArea): void {
    const trace = traceEntry.trace;
    const slices = trace.values
      .map((value, index) => ({
        value: Number(value),
        index,
        label: trace.labels?.[index] ?? `Slice ${index + 1}`
      }))
      .filter((entry) => Number.isFinite(entry.value) && entry.value > 0);
    if (slices.length === 0) {
      return;
    }

    const sum = slices.reduce((total, slice) => total + slice.value, 0);
    const radius = Math.max(12, Math.min(plotArea.width, plotArea.height) * 0.4);
    const cx = plotArea.x + plotArea.width / 2;
    const cy = plotArea.y + plotArea.height / 2;
    const hole = Math.max(0, Math.min(0.95, trace.hole ?? 0));
    const pullValues = Array.isArray(trace.pull) ? trace.pull : undefined;
    const defaultPull = pullValues ? 0 : Math.max(0, Number(trace.pull ?? 0));
    const labelCandidates: Array<{
      side: "left" | "right";
      x: number;
      y: number;
      anchorX: number;
      anchorY: number;
      text: string;
    }> = [];
    let start = -Math.PI / 2;

    slices.forEach((slice, index) => {
      const angle = (slice.value / sum) * Math.PI * 2;
      const end = start + angle;
      const mid = start + angle / 2;
      const pull = Math.max(0, Number(pullValues ? pullValues[index] ?? 0 : defaultPull) || 0);
      const pullOffset = radius * Math.min(1, pull) * 0.34;
      const sliceCx = cx + Math.cos(mid) * pullOffset;
      const sliceCy = cy + Math.sin(mid) * pullOffset;
      context.beginPath();
      if (hole > 0) {
        context.arc(sliceCx, sliceCy, radius, start, end);
        context.arc(sliceCx, sliceCy, radius * hole, end, start, true);
      } else {
        context.moveTo(sliceCx, sliceCy);
        context.arc(sliceCx, sliceCy, radius, start, end);
      }
      context.closePath();
      context.fillStyle = trace.marker?.colors?.[index] ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
      context.globalAlpha = trace.marker?.opacity ?? 1;
      context.fill();
      context.globalAlpha = 1;
      const targetX = sliceCx + Math.cos(mid) * radius * 0.7;
      const targetY = sliceCy + Math.sin(mid) * radius * 0.7;
      this.addInteractiveTarget(
        {
          traceIndex: traceEntry.index,
          pointIndex: slice.index,
          traceName: trace.name,
          x: slice.label,
          y: formatNumeric(slice.value)
        },
        targetX - 12,
        targetY - 12,
        24,
        24
      );
      const percent = (slice.value / Math.max(Number.EPSILON, sum)) * 100;
      const anchorRadius = radius + pullOffset;
      const anchorX = sliceCx + Math.cos(mid) * anchorRadius;
      const anchorY = sliceCy + Math.sin(mid) * anchorRadius;
      const textX = sliceCx + Math.cos(mid) * (anchorRadius + 20);
      labelCandidates.push({
        side: textX >= sliceCx ? "right" : "left",
        x: textX,
        y: anchorY,
        anchorX,
        anchorY,
        text: `${slice.label} (${percent.toFixed(1)}%)`
      });
      start = end;
    });

    if (labelCandidates.length > 1 && labelCandidates.length <= 28) {
      const minY = plotArea.y + 10;
      const maxY = plotArea.y + plotArea.height - 10;
      const relax = (items: Array<{ y: number }>): void => {
        if (items.length === 0) {
          return;
        }
        items.sort((left, right) => left.y - right.y);
        let nextY = minY;
        items.forEach((entry) => {
          entry.y = Math.max(entry.y, nextY);
          nextY = entry.y + 12;
        });
        nextY = maxY;
        for (let index = items.length - 1; index >= 0; index -= 1) {
          items[index].y = Math.min(items[index].y, nextY);
          nextY = items[index].y - 12;
        }
      };
      const left = labelCandidates.filter((candidate) => candidate.side === "left");
      const right = labelCandidates.filter((candidate) => candidate.side === "right");
      relax(left);
      relax(right);
      context.strokeStyle = "#94a3b8";
      context.lineWidth = 1;
      context.fillStyle = "#334155";
      context.font = "10px Arial, sans-serif";
      labelCandidates.forEach((candidate) => {
        const elbowX = candidate.side === "right" ? candidate.x - 8 : candidate.x + 8;
        context.beginPath();
        context.moveTo(candidate.anchorX, candidate.anchorY);
        context.lineTo(elbowX, candidate.y);
        context.lineTo(candidate.x, candidate.y);
        context.stroke();
        context.textAlign = candidate.side === "right" ? "left" : "right";
        context.fillText(candidate.text, candidate.x, candidate.y + 3);
      });
    }
  }

  private renderDonutScene(context: CanvasRenderingContext2D, traceEntry: IndexedTrace<DonutTrace>, plotArea: ComputedPlotArea): void {
    this.renderPieScene(
      context,
      {
        ...traceEntry,
        trace: {
          ...traceEntry.trace,
          type: "pie",
          hole: traceEntry.trace.hole ?? 0.48
        }
      },
      plotArea
    );
  }

  private renderSunburstScene(context: CanvasRenderingContext2D, traceEntry: IndexedTrace<SunburstTrace>, plotArea: ComputedPlotArea): void {
    this.drawPlotArea(context, plotArea);
    const trace = traceEntry.trace;
    const ids = trace.labels.map((label, index) => String(trace.ids?.[index] ?? label ?? index));
    const labelsById = new Map(ids.map((id, index) => [id, trace.labels[index] ?? id]));
    const parentById = new Map(ids.map((id, index) => [id, String(trace.parents[index] ?? "")]));
    const rootId = trace.rootId !== undefined ? String(trace.rootId) : "";
    const drillDepth = clampInt(trace.drilldownDepth ?? 99, 1, 99);
    const roots = trace.parents
      .map((parent, index) => ({ parent: String(parent ?? ""), index }))
      .filter((entry) => (rootId ? entry.parent === rootId : !entry.parent));
    const total = roots.reduce((sum, entry) => sum + Math.max(0, Number(trace.values[entry.index])), 0) || 1;
    const radius = Math.max(14, Math.min(plotArea.width, plotArea.height) * 0.42);
    const cx = plotArea.x + plotArea.width / 2;
    const cy = plotArea.y + plotArea.height / 2;
    const colors = trace.marker?.colors ?? DEFAULT_SERIES_COLORS;
    let start = -Math.PI / 2;

    roots.forEach((entry, rootPosition) => {
      const value = Math.max(0, Number(trace.values[entry.index]));
      const angle = (value / total) * Math.PI * 2;
      const end = start + angle;
      const pathLabels: string[] = [];
      let cursor = ids[entry.index];
      let depth = 0;
      while (cursor && depth < 24) {
        pathLabels.unshift(labelsById.get(cursor) ?? cursor);
        cursor = parentById.get(cursor) ?? "";
        depth += 1;
      }
      if (pathLabels.length > drillDepth + 1) {
        start = end;
        return;
      }
      context.beginPath();
      context.moveTo(cx, cy);
      context.arc(cx, cy, radius, start, end);
      context.closePath();
      context.fillStyle = colors[rootPosition % colors.length];
      context.globalAlpha = trace.marker?.opacity ?? 0.94;
      context.fill();
      context.globalAlpha = 1;
      this.addInteractiveTarget(
        {
          traceIndex: traceEntry.index,
          pointIndex: entry.index,
          traceName: trace.name,
          x: pathLabels.join(" > "),
          y: formatNumeric(value)
        },
        cx - radius,
        cy - radius,
        radius * 2,
        radius * 2
      );
      start = end;
    });
    if (rootId) {
      const backRadius = radius * 0.48;
      context.beginPath();
      context.fillStyle = "#ffffff";
      context.strokeStyle = "#94a3b8";
      context.lineWidth = 1;
      context.arc(cx, cy, backRadius, 0, Math.PI * 2);
      context.fill();
      context.stroke();
      context.fillStyle = "#334155";
      context.font = "11px Arial, sans-serif";
      context.textAlign = "center";
      context.fillText("Voltar", cx, cy + 4);
      this.addInteractiveTarget(
        {
          traceIndex: traceEntry.index,
          pointIndex: -1,
          traceName: trace.name,
          x: "voltar",
          y: rootId
        },
        cx - backRadius,
        cy - backRadius,
        backRadius * 2,
        backRadius * 2
      );
    }
  }

  private renderTreemapScene(context: CanvasRenderingContext2D, traceEntry: IndexedTrace<TreemapTrace>, plotArea: ComputedPlotArea): void {
    this.drawPlotArea(context, plotArea);
    const trace = traceEntry.trace;
    const ids = trace.labels.map((label, index) => String(trace.ids?.[index] ?? label ?? index));
    const labelsById = new Map(ids.map((id, index) => [id, trace.labels[index] ?? id]));
    const parentById = new Map(ids.map((id, index) => [id, String(trace.parents[index] ?? "")]));
    const rootId = trace.rootId !== undefined ? String(trace.rootId) : "";
    const drillDepth = clampInt(trace.drilldownDepth ?? 99, 1, 99);
    const headerHeight = rootId ? Math.min(28, Math.max(20, plotArea.height * 0.14)) : 0;
    const contentArea: ComputedPlotArea = {
      x: plotArea.x,
      y: plotArea.y + headerHeight,
      width: plotArea.width,
      height: Math.max(1, plotArea.height - headerHeight)
    };
    if (rootId) {
      context.fillStyle = "#f8fafc";
      context.fillRect(plotArea.x, plotArea.y, plotArea.width, headerHeight);
      context.strokeStyle = "#cbd5e1";
      context.lineWidth = 1;
      context.strokeRect(plotArea.x, plotArea.y, plotArea.width, headerHeight);
      context.fillStyle = "#334155";
      context.font = "11px Arial, sans-serif";
      context.textAlign = "left";
      context.fillText(`Voltar (${rootId})`, plotArea.x + 8, plotArea.y + headerHeight / 2 + 4);
      this.addInteractiveTarget(
        {
          traceIndex: traceEntry.index,
          pointIndex: -1,
          traceName: trace.name,
          x: "voltar",
          y: rootId
        },
        plotArea.x,
        plotArea.y,
        plotArea.width,
        headerHeight
      );
    }
    const visibleNodes = trace.values
      .map((value, index) => ({ value: Math.max(0, Number(value)), index }))
      .filter((entry) => {
        const path: string[] = [];
        let cursor = ids[entry.index];
        let depth = 0;
        while (cursor && depth < 24) {
          path.unshift(cursor);
          cursor = parentById.get(cursor) ?? "";
          depth += 1;
        }
        if (path.length > drillDepth + 1) {
          return false;
        }
        return rootId ? path.includes(rootId) : true;
      });
    const total = visibleNodes.reduce((sum, node) => sum + node.value, 0) || 1;
    const colors = trace.marker?.colors ?? DEFAULT_SERIES_COLORS;
    let cursorX = contentArea.x;

    visibleNodes.forEach(({ value, index }) => {
      const width = Math.max(2, (value / total) * contentArea.width);
      const pathLabels: string[] = [];
      let cursor = ids[index];
      let depth = 0;
      while (cursor && depth < 24) {
        pathLabels.unshift(labelsById.get(cursor) ?? cursor);
        cursor = parentById.get(cursor) ?? "";
        depth += 1;
      }
      context.fillStyle = colors[index % colors.length];
      context.globalAlpha = trace.marker?.opacity ?? 0.9;
      context.fillRect(cursorX, contentArea.y, width, contentArea.height);
      context.globalAlpha = 1;
      context.strokeStyle = "#ffffff";
      context.lineWidth = 1;
      context.strokeRect(cursorX, contentArea.y, width, contentArea.height);
      context.fillStyle = "#0f172a";
      context.font = "11px Arial, sans-serif";
      context.textAlign = "left";
      context.fillText(trace.labels[index], cursorX + 4, contentArea.y + 16);
      this.addInteractiveTarget(
        {
          traceIndex: traceEntry.index,
          pointIndex: index,
          traceName: trace.name,
          x: pathLabels.join(" > "),
          y: formatNumeric(value)
        },
        cursorX,
        contentArea.y,
        width,
        contentArea.height
      );
      cursorX += width;
    });
  }

  private renderSankeyScene(context: CanvasRenderingContext2D, traceEntry: IndexedTrace<SankeyTrace>, plotArea: ComputedPlotArea): void {
    this.drawPlotArea(context, plotArea);
    const trace = traceEntry.trace;
    const nodeCount = trace.nodes.ids.length;
    const labels = trace.nodes.labels ?? trace.nodes.ids.map((value) => String(value));
    const incoming = new Array<number>(nodeCount).fill(0);
    const outgoing = new Array<number>(nodeCount).fill(0);
    trace.links.value.forEach((raw, index) => {
      const value = Math.max(0, Number(raw));
      const source = clampInt(Number(trace.links.source[index]), 0, nodeCount - 1);
      const target = clampInt(Number(trace.links.target[index]), 0, nodeCount - 1);
      outgoing[source] += value;
      incoming[target] += value;
    });

    const levels = new Array<number>(nodeCount).fill(0);
    for (let pass = 0; pass < nodeCount; pass += 1) {
      let changed = false;
      trace.links.source.forEach((rawSource, index) => {
        const source = clampInt(Number(rawSource), 0, nodeCount - 1);
        const target = clampInt(Number(trace.links.target[index]), 0, nodeCount - 1);
        if (levels[target] <= levels[source]) {
          levels[target] = levels[source] + 1;
          changed = true;
        }
      });
      if (!changed) {
        break;
      }
    }
    const maxLevel = Math.max(1, ...levels);
    const nodesByLevel = new Map<number, number[]>();
    levels.forEach((level, nodeIndex) => {
      const bucket = nodesByLevel.get(level);
      if (bucket) {
        bucket.push(nodeIndex);
      } else {
        nodesByLevel.set(level, [nodeIndex]);
      }
    });
    const nodeWidth = Math.max(10, Math.min(22, plotArea.width * 0.045));
    const positions = new Array<{ x: number; y: number; height: number }>(nodeCount);
    for (const [level, nodes] of nodesByLevel.entries()) {
      const x = plotArea.x + (level / maxLevel) * Math.max(1, plotArea.width - nodeWidth);
      const totals = nodes.map((nodeIndex) => Math.max(1, Math.max(incoming[nodeIndex], outgoing[nodeIndex], 1)));
      const sum = totals.reduce((acc, value) => acc + value, 0);
      let yCursor = plotArea.y;
      nodes.forEach((nodeIndex, localIndex) => {
        const ratio = totals[localIndex] / Math.max(1, sum);
        const height = Math.max(12, ratio * plotArea.height - 4);
        positions[nodeIndex] = { x, y: yCursor, height };
        yCursor += height + 4;
      });
    }

    const outOffsets = new Array<number>(nodeCount).fill(0);
    const inOffsets = new Array<number>(nodeCount).fill(0);
    trace.links.value.forEach((rawValue, linkIndex) => {
      const value = Math.max(0, Number(rawValue));
      if (!Number.isFinite(value) || value <= 0) {
        return;
      }
      const source = clampInt(Number(trace.links.source[linkIndex]), 0, nodeCount - 1);
      const target = clampInt(Number(trace.links.target[linkIndex]), 0, nodeCount - 1);
      const from = positions[source];
      const to = positions[target];
      if (!from || !to) {
        return;
      }
      const sourceScale = Math.max(incoming[source], outgoing[source], 1);
      const targetScale = Math.max(incoming[target], outgoing[target], 1);
      const strokeWidth = Math.max(1.5, (value / sourceScale) * from.height);
      const startY = from.y + outOffsets[source] + strokeWidth / 2;
      const endY = to.y + inOffsets[target] + Math.max(1.5, (value / targetScale) * to.height) / 2;
      outOffsets[source] += strokeWidth;
      inOffsets[target] += Math.max(1.5, (value / targetScale) * to.height);
      const startX = from.x + nodeWidth;
      const endX = to.x;
      context.strokeStyle = trace.links.colors?.[linkIndex] ?? "rgba(37, 99, 235, 0.32)";
      context.lineWidth = strokeWidth;
      context.globalAlpha = 0.86;
      context.beginPath();
      context.moveTo(startX, startY);
      context.bezierCurveTo(startX + (endX - startX) * 0.45, startY, startX + (endX - startX) * 0.55, endY, endX, endY);
      context.stroke();
      context.globalAlpha = 1;
      this.addInteractiveTarget(
        {
          traceIndex: traceEntry.index,
          pointIndex: linkIndex,
          traceName: trace.name,
          x: `${labels[source]} -> ${labels[target]}`,
          y: formatNumeric(value)
        },
        Math.min(startX, endX),
        Math.min(startY, endY) - 6,
        Math.abs(endX - startX),
        Math.abs(endY - startY) + 12
      );
    });

    positions.forEach((position, index) => {
      if (!position) {
        return;
      }
      context.fillStyle = trace.nodes.colors?.[index] ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
      context.fillRect(position.x, position.y, nodeWidth, position.height);
      context.strokeStyle = "#334155";
      context.lineWidth = 0.8;
      context.strokeRect(position.x, position.y, nodeWidth, position.height);
      context.fillStyle = "#0f172a";
      context.font = "11px Arial, sans-serif";
      context.textAlign = "left";
      context.fillText(labels[index], position.x + nodeWidth + 4, position.y + 11);
    });
  }

  private renderParallelCategoriesScene(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    traceEntry: IndexedTrace<ParallelCategoriesTrace>,
    plotArea: ComputedPlotArea
  ): void {
    this.drawPlotArea(context, plotArea);
    const dimensions = traceEntry.trace.dimensions;
    if (dimensions.length < 2) {
      return;
    }
    const recordCount = dimensions[0]?.values.length ?? 0;
    if (recordCount === 0) {
      return;
    }
    const axisX = dimensions.map((_, index) => plotArea.x + (index / Math.max(1, dimensions.length - 1)) * plotArea.width);
    const categoryMaps = dimensions.map((dimension) => {
      const unique = Array.from(new Set(dimension.values.map((value) => String(value ?? ""))));
      return new Map(unique.map((value, index) => [value, index]));
    });
    const selectedRows = (figure.selection?.points ?? [])
      .filter((point) => point.traceIndex === traceEntry.index && point.pointIndex >= 0)
      .map((point) => clampInt(point.pointIndex, 0, recordCount - 1));
    const seedRow = selectedRows.length > 0 ? selectedRows[0] : null;
    const filterValues = seedRow === null ? null : dimensions.map((dimension) => String(dimension.values[seedRow] ?? ""));
    if (filterValues) {
      context.fillStyle = "#2563eb";
      context.font = "10px Arial, sans-serif";
      context.textAlign = "right";
      context.fillText("Filtro ativo", plotArea.x + plotArea.width - 4, plotArea.y - 8);
    }

    context.strokeStyle = "#cbd5e1";
    context.lineWidth = 1;
    axisX.forEach((x) => {
      context.beginPath();
      context.moveTo(x, plotArea.y);
      context.lineTo(x, plotArea.y + plotArea.height);
      context.stroke();
    });
    context.fillStyle = "#334155";
    context.font = "11px Arial, sans-serif";
    context.textAlign = "center";
    dimensions.forEach((dimension, index) => {
      context.fillText(dimension.name, axisX[index], plotArea.y - 8);
    });

    context.strokeStyle = traceEntry.trace.line?.color ?? "rgba(37, 99, 235, 0.42)";
    context.lineWidth = 1.1;
    context.globalAlpha = traceEntry.trace.line?.opacity ?? 0.7;
    for (let rowIndex = 0; rowIndex < recordCount; rowIndex += 1) {
      if (
        filterValues &&
        !dimensions.every((dimension, dimensionIndex) => String(dimension.values[rowIndex] ?? "") === filterValues[dimensionIndex])
      ) {
        continue;
      }
      const points: Array<{ x: number; y: number; label: string }> = [];
      dimensions.forEach((dimension, dimensionIndex) => {
        const raw = String(dimension.values[rowIndex] ?? "");
        const categoryIndex = categoryMaps[dimensionIndex].get(raw) ?? 0;
        const count = Math.max(1, categoryMaps[dimensionIndex].size - 1);
        const y = plotArea.y + (categoryIndex / Math.max(1, count)) * plotArea.height;
        points.push({ x: axisX[dimensionIndex], y, label: raw });
      });
      context.beginPath();
      points.forEach((point, index) => {
        if (index === 0) {
          context.moveTo(point.x, point.y);
        } else {
          context.lineTo(point.x, point.y);
        }
      });
      context.globalAlpha = selectedRows.includes(rowIndex) ? 0.95 : traceEntry.trace.line?.opacity ?? 0.7;
      context.lineWidth = selectedRows.includes(rowIndex) ? 1.8 : 1.1;
      context.stroke();
      this.addInteractiveTarget(
        {
          traceIndex: traceEntry.index,
          pointIndex: rowIndex,
          traceName: traceEntry.trace.name,
          x: `row ${rowIndex + 1}`,
          y: points.map((point) => point.label).join(" | ")
        },
        plotArea.x,
        Math.min(...points.map((point) => point.y)) - 4,
        plotArea.width,
        Math.max(8, Math.max(...points.map((point) => point.y)) - Math.min(...points.map((point) => point.y)) + 8)
      );
    }
    context.globalAlpha = 1;
  }

  private renderPolarScene(context: CanvasRenderingContext2D, traces: IndexedTrace<PolarTrace>[], plotArea: ComputedPlotArea): void {
    this.drawPlotArea(context, plotArea);
    const cx = plotArea.x + plotArea.width / 2;
    const cy = plotArea.y + plotArea.height / 2;
    const radius = Math.max(10, Math.min(plotArea.width, plotArea.height) * 0.42);
    context.strokeStyle = "#e2e8f0";
    context.lineWidth = 1;
    for (let level = 1; level <= 4; level += 1) {
      context.beginPath();
      context.arc(cx, cy, (radius * level) / 4, 0, Math.PI * 2);
      context.stroke();
    }

    traces.forEach((entry, traceIndex) => {
      const color = entry.trace.line?.color ?? entry.trace.marker?.color ?? DEFAULT_SERIES_COLORS[traceIndex % DEFAULT_SERIES_COLORS.length];
      const maxR = Math.max(...entry.trace.r.map((value) => Math.abs(Number(value))), 1);
      const points: Array<{ x: number; y: number; index: number; theta: string; r: number }> = [];
      entry.trace.r.forEach((rawR, pointIndex) => {
        const rValue = Number(rawR);
        if (!Number.isFinite(rValue)) {
          return;
        }
        const angle = toRadians(entry.trace.theta[pointIndex], pointIndex, entry.trace.theta.length);
        const distance = (rValue / maxR) * radius;
        points.push({
          x: cx + Math.cos(angle) * distance,
          y: cy + Math.sin(angle) * distance,
          index: pointIndex,
          theta: String(entry.trace.theta[pointIndex] ?? pointIndex),
          r: rValue
        });
      });
      const mode = entry.trace.mode ?? "lines+markers";
      const variant = entry.trace.variant ?? "scatter";
      if (variant === "bar") {
        const barWidth = Math.max(0.08, Math.min(Math.PI * 0.45, entry.trace.barWidth ?? (Math.PI * 2) / Math.max(6, points.length)));
        points.forEach((point) => {
          const theta = toRadians(entry.trace.theta[point.index], point.index, entry.trace.theta.length);
          const start = theta - barWidth / 2;
          const end = theta + barWidth / 2;
          context.beginPath();
          context.moveTo(cx, cy);
          context.arc(cx, cy, distance(cx, cy, point.x, point.y), start, end);
          context.closePath();
          context.fillStyle = resolveColor(color, 0.74);
          context.globalAlpha = entry.trace.marker?.opacity ?? 0.9;
          context.fill();
          context.globalAlpha = 1;
          context.strokeStyle = color;
          context.lineWidth = entry.trace.line?.width ?? 1;
          context.stroke();
        });
      }
      if ((variant === "line" || variant === "scatter" || variant === "area") && mode.includes("lines") && points.length >= 2) {
        context.strokeStyle = color;
        context.lineWidth = entry.trace.line?.width ?? 2;
        context.globalAlpha = entry.trace.line?.opacity ?? 1;
        context.beginPath();
        points.forEach((point, index) => {
          if (index === 0) {
            context.moveTo(point.x, point.y);
          } else {
            context.lineTo(point.x, point.y);
          }
        });
        if (variant === "area") {
          context.lineTo(cx, cy);
          context.closePath();
          context.fillStyle = resolveColor(color, 0.22);
          context.fill();
        }
        context.stroke();
        context.globalAlpha = 1;
      }
      if ((variant === "scatter" || variant === "line") && mode.includes("markers")) {
        points.forEach((point) => {
          context.beginPath();
          context.fillStyle = color;
          context.globalAlpha = entry.trace.marker?.opacity ?? 1;
          context.arc(point.x, point.y, entry.trace.marker?.size ?? 4, 0, Math.PI * 2);
          context.fill();
          context.globalAlpha = 1;
          this.addInteractiveTarget(
            {
              traceIndex: entry.index,
              pointIndex: point.index,
              traceName: entry.trace.name,
              x: point.theta,
              y: formatNumeric(point.r)
            },
            point.x - 7,
            point.y - 7,
            14,
            14
          );
        });
      }
    });
  }

  private renderTernaryScene(context: CanvasRenderingContext2D, traces: IndexedTrace<TernaryTrace>[], plotArea: ComputedPlotArea): void {
    this.drawPlotArea(context, plotArea);
    const side = Math.min(plotArea.width, plotArea.height) * 0.82;
    const centerX = plotArea.x + plotArea.width / 2;
    const centerY = plotArea.y + plotArea.height / 2;
    const half = side / 2;
    const height = (Math.sqrt(3) / 2) * side;
    const top = { x: centerX, y: centerY - height / 2 };
    const left = { x: centerX - half, y: centerY + height / 2 };
    const right = { x: centerX + half, y: centerY + height / 2 };

    context.strokeStyle = "#94a3b8";
    context.lineWidth = 1.5;
    context.beginPath();
    context.moveTo(top.x, top.y);
    context.lineTo(right.x, right.y);
    context.lineTo(left.x, left.y);
    context.closePath();
    context.stroke();

    traces.forEach((entry, traceIndex) => {
      const color = entry.trace.marker?.color ?? DEFAULT_SERIES_COLORS[traceIndex % DEFAULT_SERIES_COLORS.length];
      for (let index = 0; index < entry.trace.a.length; index += 1) {
        const a = Number(entry.trace.a[index]);
        const b = Number(entry.trace.b[index]);
        const c = Number(entry.trace.c[index]);
        if (!Number.isFinite(a) || !Number.isFinite(b) || !Number.isFinite(c)) {
          continue;
        }
        const total = a + b + c;
        if (total === 0) {
          continue;
        }
        const alpha = a / total;
        const beta = b / total;
        const gamma = c / total;
        const x = top.x * alpha + right.x * beta + left.x * gamma;
        const y = top.y * alpha + right.y * beta + left.y * gamma;
        context.beginPath();
        context.fillStyle = color;
        context.globalAlpha = entry.trace.marker?.opacity ?? 0.9;
        context.arc(x, y, entry.trace.marker?.size ?? 4, 0, Math.PI * 2);
        context.fill();
        context.globalAlpha = 1;
        this.addInteractiveTarget(
          {
            traceIndex: entry.index,
            pointIndex: index,
            traceName: entry.trace.name,
            x: `a=${formatNumeric(a)} b=${formatNumeric(b)}`,
            y: `c=${formatNumeric(c)}`
          },
          x - 6,
          y - 6,
          12,
          12
        );
      }
    });
  }

  private renderGeoScene(
    context: CanvasRenderingContext2D,
    geoTraces: IndexedTrace<GeoTrace>[],
    scatterTraces: IndexedTrace<GeoScatterTrace>[],
    lineTraces: IndexedTrace<GeoLineTrace>[],
    plotArea: ComputedPlotArea
  ): void {
    this.drawPlotArea(context, plotArea);
    const geoFeatures = geoTraces.flatMap((entry) => entry.trace.geojson.features.map((feature) => ({ entry, feature })));
    const bounds = computeGeoBounds(
      geoFeatures.map((item) => item.feature.geometry.coordinates),
      scatterTraces.flatMap((entry) => entry.trace.lat.map((lat, index) => [entry.trace.lon[index], lat] as [number, number])),
      lineTraces.flatMap((entry) => entry.trace.paths.flatMap((path) => path.map((point) => [point.lon, point.lat] as [number, number])))
    );
    if (!bounds) {
      return;
    }

    let minValue = Number.POSITIVE_INFINITY;
    let maxValue = Number.NEGATIVE_INFINITY;
    const valueResolvers = new Map<number, (feature: GeoTrace["geojson"]["features"][number]) => number | null>();
    geoTraces.forEach((traceEntry) => {
      const trace = traceEntry.trace;
      const featureIdField = trace.featureIdField ?? "id";
      const locationMap =
        Array.isArray(trace.locations) && Array.isArray(trace.values)
          ? new Map(trace.locations.map((location, index) => [String(location), Number(trace.values?.[index])]))
          : null;
      valueResolvers.set(traceEntry.index, (feature) => {
        if (locationMap) {
          const featureId = String(feature.properties?.[featureIdField] ?? feature.properties?.id ?? "");
          const mapped = locationMap.get(featureId);
          return Number.isFinite(mapped) ? Number(mapped) : null;
        }
        if (!trace.valueField) {
          return null;
        }
        const raw = feature.properties?.[trace.valueField];
        const numeric = typeof raw === "number" ? raw : Number(raw);
        return Number.isFinite(numeric) ? numeric : null;
      });
    });
    geoFeatures.forEach((item) => {
      const value = valueResolvers.get(item.entry.index)?.(item.feature);
      if (typeof value !== "number" || !Number.isFinite(value)) {
        return;
      }
      minValue = Math.min(minValue, value);
      maxValue = Math.max(maxValue, value);
    });
    if (!Number.isFinite(minValue) || !Number.isFinite(maxValue)) {
      minValue = 0;
      maxValue = 1;
    }

    geoFeatures.forEach((item, featureIndex) => {
      const value = valueResolvers.get(item.entry.index)?.(item.feature);
      const palette = resolvePalette("continuous", item.entry.trace.colorscale, item.entry.trace.reverseScale === true);
      const numericValue = typeof value === "number" && Number.isFinite(value) ? value : null;
      const fill = resolveColorFromScale(numericValue, {
        mode: "continuous",
        colors: palette,
        min: minValue,
        max: maxValue,
        missingColor: item.entry.trace.missingColor ?? "#cbd5e1"
      });
      const paths = expandGeoPaths(item.feature.geometry.coordinates).map((path) =>
        simplifyGeoPath(path, Number(item.entry.trace.simplifyTolerance ?? 0))
      );
      paths.forEach((pathPoints, pathIndex) => {
        const mapped = pathPoints.map((coordinate) => mapGeoToPlot(coordinate, bounds, plotArea));
        if (mapped.length < 3) {
          return;
        }
        context.beginPath();
        mapped.forEach((point, pointIndex) => {
          if (pointIndex === 0) {
            context.moveTo(point.x, point.y);
          } else {
            context.lineTo(point.x, point.y);
          }
        });
        context.closePath();
        context.fillStyle = fill;
        context.globalAlpha = 0.78;
        context.fill();
        context.globalAlpha = 1;
        context.strokeStyle = "#334155";
        context.lineWidth = 0.8;
        context.stroke();
        this.addInteractiveTarget(
          {
            traceIndex: item.entry.index,
            pointIndex: pathIndex,
            traceName: item.entry.trace.name,
            x: String(item.feature.properties?.[item.entry.trace.featureIdField ?? "id"] ?? `feature ${featureIndex + 1}`),
            y: numericValue !== null ? formatNumeric(numericValue) : item.entry.trace.missingColor ?? "n/a"
          },
          Math.min(...mapped.map((point) => point.x)),
          Math.min(...mapped.map((point) => point.y)),
          Math.max(6, Math.max(...mapped.map((point) => point.x)) - Math.min(...mapped.map((point) => point.x))),
          Math.max(6, Math.max(...mapped.map((point) => point.y)) - Math.min(...mapped.map((point) => point.y)))
        );
      });
    });

    const legendTrace = geoTraces.find((entry) => entry.trace.showColorLegend);
    if (legendTrace) {
      this.drawGeoColorLegend(context, legendTrace.trace, plotArea, minValue, maxValue);
    }

    lineTraces.forEach((entry, traceIndex) => {
      const color = entry.trace.line?.color ?? DEFAULT_SERIES_COLORS[traceIndex % DEFAULT_SERIES_COLORS.length];
      context.strokeStyle = color;
      context.lineWidth = entry.trace.line?.width ?? 1.8;
      context.globalAlpha = entry.trace.line?.opacity ?? 0.92;
      entry.trace.paths.forEach((pathPoints, pathIndex) => {
        const mapped = pathPoints.map((point) => mapGeoToPlot([point.lon, point.lat], bounds, plotArea));
        if (mapped.length < 2) {
          return;
        }
        context.beginPath();
        mapped.forEach((point, pointIndex) => {
          if (pointIndex === 0) {
            context.moveTo(point.x, point.y);
          } else {
            context.lineTo(point.x, point.y);
          }
        });
        context.stroke();
        this.addInteractiveTarget(
          {
            traceIndex: entry.index,
            pointIndex: pathIndex,
            traceName: entry.trace.name,
            x: "geo-line",
            y: color
          },
          Math.min(...mapped.map((point) => point.x)),
          Math.min(...mapped.map((point) => point.y)),
          Math.max(6, Math.max(...mapped.map((point) => point.x)) - Math.min(...mapped.map((point) => point.x))),
          Math.max(6, Math.max(...mapped.map((point) => point.y)) - Math.min(...mapped.map((point) => point.y)))
        );
      });
      context.globalAlpha = 1;
    });

    scatterTraces.forEach((entry, traceIndex) => {
      const points = entry.trace.lat
        .map((lat, index) => ({ lat: Number(lat), lon: Number(entry.trace.lon[index]), index }))
        .filter((point) => Number.isFinite(point.lat) && Number.isFinite(point.lon))
        .map((point) => ({ ...point, ...mapGeoToPlot([point.lon, point.lat], bounds, plotArea) }));
      const mode = entry.trace.mode ?? "markers";
      const color = entry.trace.marker?.color ?? entry.trace.line?.color ?? DEFAULT_SERIES_COLORS[traceIndex % DEFAULT_SERIES_COLORS.length];
      if (mode.includes("lines") && points.length >= 2) {
        context.strokeStyle = entry.trace.line?.color ?? color;
        context.lineWidth = entry.trace.line?.width ?? 1.4;
        context.globalAlpha = entry.trace.line?.opacity ?? entry.trace.marker?.opacity ?? 0.94;
        context.beginPath();
        points.forEach((point, pointIndex) => {
          if (pointIndex === 0) {
            context.moveTo(point.x, point.y);
          } else {
            context.lineTo(point.x, point.y);
          }
        });
        context.stroke();
        context.globalAlpha = 1;
      }
      if (mode.includes("markers")) {
        points.forEach((point) => {
          context.beginPath();
          context.fillStyle = color;
          context.globalAlpha = entry.trace.marker?.opacity ?? 0.94;
          context.arc(point.x, point.y, entry.trace.marker?.size ?? 4, 0, Math.PI * 2);
          context.fill();
          context.globalAlpha = 1;
          this.addInteractiveTarget(
            {
              traceIndex: entry.index,
              pointIndex: point.index,
              traceName: entry.trace.name,
              x: `${point.lat.toFixed(4)}, ${point.lon.toFixed(4)}`,
              y: color
            },
            point.x - 7,
            point.y - 7,
            14,
            14
          );
        });
      }
    });
  }

  private drawGeoColorLegend(
    context: CanvasRenderingContext2D,
    trace: GeoTrace,
    plotArea: ComputedPlotArea,
    minValue: number,
    maxValue: number
  ): void {
    const legendWidth = Math.max(88, Math.min(140, plotArea.width * 0.22));
    const legendHeight = 10;
    const x = plotArea.x + plotArea.width - legendWidth - 12;
    const y = plotArea.y + plotArea.height - 28;
    const palette = resolvePalette("continuous", trace.colorscale, trace.reverseScale === true);
    const steps = 22;
    for (let index = 0; index < steps; index += 1) {
      const ratio = steps <= 1 ? 0 : index / (steps - 1);
      const value = minValue + ratio * (maxValue - minValue);
      context.fillStyle = resolveColorFromScale(value, {
        mode: "continuous",
        colors: palette,
        min: minValue,
        max: maxValue
      });
      context.fillRect(x + (legendWidth / steps) * index, y, Math.ceil(legendWidth / steps), legendHeight);
    }
    context.strokeStyle = "#94a3b8";
    context.lineWidth = 0.7;
    context.strokeRect(x, y, legendWidth, legendHeight);
    context.fillStyle = "#334155";
    context.font = "10px Arial, sans-serif";
    context.textAlign = "left";
    context.fillText(formatNumeric(minValue), x, y - 2);
    context.textAlign = "right";
    context.fillText(formatNumeric(maxValue), x + legendWidth, y - 2);
  }

  private renderLayoutOverlays(context: CanvasRenderingContext2D, figure: ChartFigure, layout: ComputedLayout): void {
    const domains = buildCartesianDomains(figure);
    const plotArea = layout.plotArea;
    const mapX = (value: number, ref: "paper" | "data" | undefined): number => {
      if (ref === "paper") {
        return clamp(value, 0, 1) * layout.width;
      }
      if (!domains) {
        return plotArea.x + clamp(value, 0, 1) * plotArea.width;
      }
      return plotArea.x + ((value - domains.x[0]) / Math.max(Number.EPSILON, domains.x[1] - domains.x[0])) * plotArea.width;
    };
    const mapY = (value: number, ref: "paper" | "data" | undefined): number => {
      if (ref === "paper") {
        return clamp(value, 0, 1) * layout.height;
      }
      if (!domains) {
        return plotArea.y + (1 - clamp(value, 0, 1)) * plotArea.height;
      }
      return plotArea.y + (1 - (value - domains.y[0]) / Math.max(Number.EPSILON, domains.y[1] - domains.y[0])) * plotArea.height;
    };

    figure.layout.shapes.forEach((shape) => {
      context.save();
      context.strokeStyle = shape.stroke ?? "#2563eb";
      context.lineWidth = shape.strokeWidth ?? 1.2;
      context.fillStyle = shape.fill ?? "rgba(37, 99, 235, 0.12)";
      context.globalAlpha = shape.opacity ?? 1;
      if (shape.type === "line") {
        context.beginPath();
        context.moveTo(mapX(shape.x0, shape.xRef), mapY(shape.y0, shape.yRef));
        context.lineTo(mapX(shape.x1 ?? shape.x0, shape.xRef), mapY(shape.y1 ?? shape.y0, shape.yRef));
        context.stroke();
      } else if (shape.type === "rect") {
        const x0 = mapX(shape.x0, shape.xRef);
        const y0 = mapY(shape.y0, shape.yRef);
        const x1 = mapX(shape.x1 ?? shape.x0, shape.xRef);
        const y1 = mapY(shape.y1 ?? shape.y0, shape.yRef);
        context.fillRect(Math.min(x0, x1), Math.min(y0, y1), Math.max(1, Math.abs(x1 - x0)), Math.max(1, Math.abs(y1 - y0)));
        context.strokeRect(Math.min(x0, x1), Math.min(y0, y1), Math.max(1, Math.abs(x1 - x0)), Math.max(1, Math.abs(y1 - y0)));
      } else if (shape.type === "circle") {
        context.beginPath();
        context.arc(mapX(shape.x0, shape.xRef), mapY(shape.y0, shape.yRef), Math.max(1, shape.radius ?? 8), 0, Math.PI * 2);
        context.fill();
        context.stroke();
      } else if (shape.type === "path" && shape.path) {
        const path = new Path2D(shape.path);
        context.fill(path);
        context.stroke(path);
      } else if (shape.type === "region" && Array.isArray(shape.points) && shape.points.length >= 3) {
        context.beginPath();
        shape.points.forEach((point, index) => {
          const x = mapX(point.x, shape.xRef);
          const y = mapY(point.y, shape.yRef);
          if (index === 0) {
            context.moveTo(x, y);
          } else {
            context.lineTo(x, y);
          }
        });
        context.closePath();
        context.fill();
        context.stroke();
      }
      context.restore();
    });

    figure.layout.annotations.forEach((annotation) => {
      const x = mapX(annotation.x, annotation.xRef);
      const y = mapY(annotation.y, annotation.yRef);
      if (annotation.showArrow) {
        const ax = mapX(annotation.arrowToX ?? annotation.x, annotation.xRef);
        const ay = mapY(annotation.arrowToY ?? annotation.y, annotation.yRef);
        context.strokeStyle = annotation.color ?? "#334155";
        context.lineWidth = 1.2;
        context.beginPath();
        context.moveTo(ax, ay);
        context.lineTo(x, y);
        context.stroke();
        const angle = Math.atan2(y - ay, x - ax);
        const leftX = x - Math.cos(angle - 0.55) * 7;
        const leftY = y - Math.sin(angle - 0.55) * 7;
        const rightX = x - Math.cos(angle + 0.55) * 7;
        const rightY = y - Math.sin(angle + 0.55) * 7;
        context.fillStyle = annotation.color ?? "#334155";
        context.beginPath();
        context.moveTo(x, y);
        context.lineTo(leftX, leftY);
        context.lineTo(rightX, rightY);
        context.closePath();
        context.fill();
      }
      context.save();
      context.fillStyle = annotation.color ?? "#0f172a";
      context.font = `${annotation.fontSize ?? 12}px Arial, sans-serif`;
      context.textAlign = annotation.align === "left" ? "left" : annotation.align === "right" ? "right" : "center";
      if (typeof annotation.rotate === "number" && Number.isFinite(annotation.rotate)) {
        context.translate(x, y);
        context.rotate((annotation.rotate * Math.PI) / 180);
        context.fillText(annotation.text, 0, 0);
      } else {
        context.fillText(annotation.text, x, y);
      }
      context.restore();
    });

    figure.layout.images.forEach((imageLayer) => {
      if (!imageLayer.source) {
        return;
      }
      const image = this.resolveImageElement(imageLayer.source);
      if (!image || !image.complete || image.naturalWidth <= 0 || image.naturalHeight <= 0) {
        return;
      }
      const x = mapX(imageLayer.x, imageLayer.xRef);
      const y = mapY(imageLayer.y, imageLayer.yRef);
      context.save();
      context.globalAlpha = imageLayer.opacity ?? 1;
      context.drawImage(image, x, y, Math.max(1, imageLayer.width), Math.max(1, imageLayer.height));
      context.restore();
    });
  }

  private renderProjected3dScene(
    context: CanvasRenderingContext2D,
    scatterTraces: IndexedTrace<Scatter3dTrace>[],
    surfaceTraces: IndexedTrace<SurfaceTrace>[],
    meshTraces: IndexedTrace<Mesh3dTrace>[],
    plotArea: ComputedPlotArea
  ): void {
    this.drawPlotArea(context, plotArea);
    const projection = createDefaultProjection3d(plotArea);

    scatterTraces.forEach((entry, traceIndex) => {
      const color = entry.trace.marker?.color ?? entry.trace.line?.color ?? DEFAULT_SERIES_COLORS[traceIndex % DEFAULT_SERIES_COLORS.length];
      const mode = entry.trace.mode ?? "markers";
      const points = entry.trace.x
        .map((x, index) => ({
          x: Number(x),
          y: Number(entry.trace.y[index]),
          z: Number(entry.trace.z[index]),
          index
        }))
        .filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y) && Number.isFinite(point.z))
        .map((point) => ({ ...point, p: projection(point.x, point.y, point.z) }));

      if (mode.includes("lines") && points.length >= 2) {
        context.strokeStyle = entry.trace.line?.color ?? color;
        context.lineWidth = entry.trace.line?.width ?? 1.6;
        context.globalAlpha = entry.trace.line?.opacity ?? 0.9;
        context.beginPath();
        points.forEach((point, pointIndex) => {
          if (pointIndex === 0) {
            context.moveTo(point.p.x, point.p.y);
          } else {
            context.lineTo(point.p.x, point.p.y);
          }
        });
        context.stroke();
        context.globalAlpha = 1;
      }
      if (mode.includes("markers")) {
        points.forEach((point) => {
          context.beginPath();
          context.fillStyle = color;
          context.globalAlpha = entry.trace.marker?.opacity ?? 0.9;
          context.arc(point.p.x, point.p.y, entry.trace.marker?.size ?? 3.5, 0, Math.PI * 2);
          context.fill();
          context.globalAlpha = 1;
          this.addInteractiveTarget(
            {
              traceIndex: entry.index,
              pointIndex: point.index,
              traceName: entry.trace.name,
              x: `x:${formatNumeric(point.x)} y:${formatNumeric(point.y)}`,
              y: `z:${formatNumeric(point.z)}`
            },
            point.p.x - 7,
            point.p.y - 7,
            14,
            14
          );
        });
      }
    });

    surfaceTraces.forEach((entry) => {
      const rows = entry.trace.z.length;
      const cols = entry.trace.z[0]?.length ?? 0;
      if (rows < 2 || cols < 2) {
        return;
      }
      const palette = resolvePalette("continuous", entry.trace.colorscale, entry.trace.reverseScale === true);
      const values = entry.trace.z.flat().map((value) => Number(value)).filter((value) => Number.isFinite(value));
      const min = values.length > 0 ? Math.min(...values) : 0;
      const max = values.length > 0 ? Math.max(...values) : 1;
      for (let row = 0; row < rows - 1; row += 1) {
        for (let col = 0; col < cols - 1; col += 1) {
          const z00 = Number(entry.trace.z[row][col]);
          const z10 = Number(entry.trace.z[row][col + 1]);
          const z11 = Number(entry.trace.z[row + 1][col + 1]);
          const z01 = Number(entry.trace.z[row + 1][col]);
          if (!Number.isFinite(z00) || !Number.isFinite(z10) || !Number.isFinite(z11) || !Number.isFinite(z01)) {
            continue;
          }
          const p0 = projection(col, row, z00);
          const p1 = projection(col + 1, row, z10);
          const p2 = projection(col + 1, row + 1, z11);
          const p3 = projection(col, row + 1, z01);
          context.beginPath();
          context.moveTo(p0.x, p0.y);
          context.lineTo(p1.x, p1.y);
          context.lineTo(p2.x, p2.y);
          context.lineTo(p3.x, p3.y);
          context.closePath();
          context.fillStyle = resolveColorFromScale((z00 + z10 + z11 + z01) / 4, {
            mode: "continuous",
            colors: palette,
            min,
            max
          });
          context.globalAlpha = 0.92;
          context.fill();
          context.globalAlpha = 1;
          context.strokeStyle = "rgba(15, 23, 42, 0.2)";
          context.lineWidth = 0.5;
          context.stroke();
        }
      }
    });

    meshTraces.forEach((entry, traceIndex) => {
      context.fillStyle = entry.trace.marker?.color ?? DEFAULT_SERIES_COLORS[(traceIndex + 2) % DEFAULT_SERIES_COLORS.length];
      context.globalAlpha = entry.trace.marker?.opacity ?? 0.56;
      for (let index = 0; index < entry.trace.i.length; index += 1) {
        const ia = Number(entry.trace.i[index]);
        const jb = Number(entry.trace.j[index]);
        const kc = Number(entry.trace.k[index]);
        const a = projection(Number(entry.trace.x[ia]), Number(entry.trace.y[ia]), Number(entry.trace.z[ia]));
        const b = projection(Number(entry.trace.x[jb]), Number(entry.trace.y[jb]), Number(entry.trace.z[jb]));
        const c = projection(Number(entry.trace.x[kc]), Number(entry.trace.y[kc]), Number(entry.trace.z[kc]));
        context.beginPath();
        context.moveTo(a.x, a.y);
        context.lineTo(b.x, b.y);
        context.lineTo(c.x, c.y);
        context.closePath();
        context.fill();
        context.strokeStyle = "rgba(15, 23, 42, 0.25)";
        context.lineWidth = 0.4;
        context.stroke();
      }
      context.globalAlpha = 1;
    });
  }

  private renderSubplotGroups<TTrace extends ChartTrace>(
    traces: IndexedTrace<TTrace>[],
    subplotAreas: ComputedPlotArea[],
    renderGroup: (group: IndexedTrace<TTrace>[], plotArea: ComputedPlotArea, subplotIndex: number) => void
  ): void {
    const groups = new Map<number, IndexedTrace<TTrace>[]>();
    traces.forEach((entry) => {
      const subplotIndex = clampInt((entry.trace.subplot as number | undefined) ?? 0, 0, subplotAreas.length - 1);
      const bucket = groups.get(subplotIndex);
      if (bucket) {
        bucket.push(entry);
      } else {
        groups.set(subplotIndex, [entry]);
      }
    });

    for (const [subplotIndex, group] of groups.entries()) {
      renderGroup(group, subplotAreas[subplotIndex] ?? subplotAreas[0], subplotIndex);
    }
  }

  private resolveCartesianDomains(
    figure: ChartFigure,
    traces: IndexedTrace<CartesianTrace>[],
    options?: {
      xAxisKey?: "xAxis" | "xAxis2";
      yAxisKey?: "yAxis" | "yAxis2";
      subplotIndex?: number;
    }
  ): ReturnType<typeof buildCartesianDomains> {
    if (traces.length === 0) {
      return null;
    }
    const xAxis = options?.xAxisKey === "xAxis2" ? figure.layout.xAxis2 : figure.layout.xAxis;
    const yAxis = options?.yAxisKey === "yAxis2" ? figure.layout.yAxis2 : figure.layout.yAxis;
    const subplotOverride = this.resolveSubplotRangeOverride(figure, options?.subplotIndex);
    const subsetFigure: ChartFigure = {
      ...figure,
      layout: {
        ...figure.layout,
        xAxis: {
          ...xAxis,
          min: subplotOverride?.x?.[0] ?? xAxis.min,
          max: subplotOverride?.x?.[1] ?? xAxis.max
        },
        yAxis: {
          ...yAxis,
          min: subplotOverride?.y?.[0] ?? yAxis.min,
          max: subplotOverride?.y?.[1] ?? yAxis.max
        }
      },
      data: traces.map((entry) => entry.trace)
    };
    return buildCartesianDomains(subsetFigure);
  }

  private resolveFinancialDomains(
    figure: ChartFigure,
    traces: IndexedTrace<FinancialTrace>[],
    options?: {
      xAxisKey?: "xAxis" | "xAxis2";
      yAxisKey?: "yAxis" | "yAxis2";
      subplotIndex?: number;
    }
  ): ReturnType<typeof buildCartesianDomains> {
    if (traces.length === 0) {
      return null;
    }
    const xAxis = options?.xAxisKey === "xAxis2" ? figure.layout.xAxis2 : figure.layout.xAxis;
    const yAxis = options?.yAxisKey === "yAxis2" ? figure.layout.yAxis2 : figure.layout.yAxis;
    const subplotOverride = this.resolveSubplotRangeOverride(figure, options?.subplotIndex);
    const proxyTraces = traces.map((entry) => ({
      type: "line" as const,
      x: entry.trace.x,
      y: entry.trace.close
    }));
    const subsetFigure: ChartFigure = {
      ...figure,
      layout: {
        ...figure.layout,
        xAxis: {
          ...xAxis,
          min: subplotOverride?.x?.[0] ?? xAxis.min,
          max: subplotOverride?.x?.[1] ?? xAxis.max
        },
        yAxis: {
          ...yAxis,
          min: subplotOverride?.y?.[0] ?? yAxis.min,
          max: subplotOverride?.y?.[1] ?? yAxis.max
        }
      },
      data: proxyTraces
    };
    const domains = buildCartesianDomains(subsetFigure);
    if (!domains) {
      return null;
    }
    const yValues = traces.flatMap((entry) => [...entry.trace.low, ...entry.trace.high, ...entry.trace.open, ...entry.trace.close]);
    const finite = yValues.filter((value) => Number.isFinite(value));
    if (finite.length === 0) {
      return domains;
    }
    let min = Math.min(...finite);
    let max = Math.max(...finite);
    if (min === max) {
      min -= 1;
      max += 1;
    }
    return {
      ...domains,
      y: [min, max]
    };
  }

  private resolveSubplotRangeOverride(
    figure: ChartFigure,
    subplotIndex: number | undefined
  ): { x: [number, number]; y: [number, number] } | null {
    if (subplotIndex === undefined) {
      return null;
    }
    const runtimeRanges = (figure as ChartFigure & { __subplotAxisRanges?: Record<string, { x: [number, number]; y: [number, number] }> })
      .__subplotAxisRanges;
    if (!runtimeRanges) {
      return null;
    }
    const override = runtimeRanges[String(subplotIndex)];
    if (!override) {
      return null;
    }
    return {
      x: [Number(override.x[0]), Number(override.x[1])],
      y: [Number(override.y[0]), Number(override.y[1])]
    };
  }

  private resolveImageElement(source: string): HTMLImageElement | null {
    if (typeof Image === "undefined") {
      return null;
    }
    const cached = this.imageCache.get(source);
    if (cached) {
      return cached;
    }
    const image = new Image();
    image.decoding = "async";
    image.onload = () => {
      if (!this.lastFigure || !this.lastLayout) {
        return;
      }
      this.render(this.lastFigure, this.lastLayout);
    };
    image.onerror = () => {
      this.imageCache.delete(source);
    };
    image.src = source;
    this.imageCache.set(source, image);
    return image;
  }

  private renderLegendControls(traces: IndexedTrace[], layout: ComputedLayout): void {
    const layer = this.ensureInteractionLayer();
    const legendArea =
      layout.legendArea ??
      ({
        x: layout.margin.left,
        y: layout.margin.top,
        width: Math.max(1, layout.width - layout.margin.left - layout.margin.right),
        height: 44
      } as const);

    const wrapper = document.createElement("div");
    wrapper.className = "excelsior-chart-legend";
    Object.assign(wrapper.style, {
      position: "absolute",
      left: `${legendArea.x}px`,
      top: `${legendArea.y}px`,
      width: `${legendArea.width}px`,
      minHeight: `${legendArea.height}px`,
      display: "flex",
      flexWrap: "wrap",
      alignItems: "center",
      gap: "6px"
    });

    traces.forEach((entry, index) => {
      const button = document.createElement("button");
      button.type = "button";
      button.dataset.chartLegend = "item";
      button.dataset.legendTraceIndex = String(entry.index);
      Object.assign(button.style, {
        display: "inline-flex",
        alignItems: "center",
        gap: "4px",
        border: "none",
        background: "transparent",
        color: "#334155",
        fontSize: "11px",
        fontFamily: "Arial, sans-serif",
        cursor: "pointer",
        opacity: entry.trace.visible === false ? "0.45" : "1",
        padding: "2px 3px"
      });

      const marker = document.createElement("span");
      Object.assign(marker.style, {
        display: "inline-block",
        width: "10px",
        height: "10px",
        border: "1px solid #64748b",
        background: this.resolveTraceColor(entry.trace, entry.index)
      });
      const label = document.createElement("span");
      label.textContent = entry.trace.name?.trim() || `Trace ${index + 1}`;
      button.append(marker, label);
      wrapper.append(button);
    });

    layer.append(wrapper);
  }

  private drawPlotArea(context: CanvasRenderingContext2D, plotArea: ComputedPlotArea): void {
    context.fillStyle = "#ffffff";
    context.fillRect(plotArea.x, plotArea.y, plotArea.width, plotArea.height);
    context.strokeStyle = "#dbe2ea";
    context.lineWidth = 1;
    context.strokeRect(plotArea.x, plotArea.y, plotArea.width, plotArea.height);
  }

  private drawGrid(
    context: CanvasRenderingContext2D,
    plotArea: ComputedPlotArea,
    xScale: LinearScale,
    yScale: LinearScale,
    options?: { xZeroLine?: boolean; yZeroLine?: boolean }
  ): void {
    context.strokeStyle = "#e2e8f0";
    context.lineWidth = 1;
    yScale.ticks(6).forEach((tick) => {
      const y = yScale.map(tick);
      context.beginPath();
      context.moveTo(plotArea.x, y);
      context.lineTo(plotArea.x + plotArea.width, y);
      context.stroke();
    });
    xScale.ticks(8).forEach((tick) => {
      const x = xScale.map(tick);
      context.beginPath();
      context.moveTo(x, plotArea.y);
      context.lineTo(x, plotArea.y + plotArea.height);
      context.stroke();
    });
    const zeroY = yScale.map(0);
    if ((options?.yZeroLine ?? true) && Number.isFinite(zeroY) && zeroY >= plotArea.y && zeroY <= plotArea.y + plotArea.height) {
      context.strokeStyle = "#94a3b8";
      context.lineWidth = 1.5;
      context.beginPath();
      context.moveTo(plotArea.x, zeroY);
      context.lineTo(plotArea.x + plotArea.width, zeroY);
      context.stroke();
    }
    const zeroX = xScale.map(0);
    if ((options?.xZeroLine ?? true) && Number.isFinite(zeroX) && zeroX >= plotArea.x && zeroX <= plotArea.x + plotArea.width) {
      context.strokeStyle = "#94a3b8";
      context.lineWidth = 1.5;
      context.beginPath();
      context.moveTo(zeroX, plotArea.y);
      context.lineTo(zeroX, plotArea.y + plotArea.height);
      context.stroke();
    }
  }

  private drawAxes(
    context: CanvasRenderingContext2D,
    figure: ChartFigure,
    plotArea: ComputedPlotArea,
    labels: string[],
    xScale: LinearScale,
    yScale: LinearScale,
    categoryScale: CategoryScale,
    yScaleSecondary?: LinearScale | null,
    xScaleSecondary?: LinearScale | null,
    labelsSecondary?: string[],
    categoryScaleSecondary?: CategoryScale
  ): void {
    const xAxisType = normalizeAxisType(figure.layout.xAxis, "category");
    const xAxis2Type = normalizeAxisType(figure.layout.xAxis2, xAxisType);
    const axisBottomY = plotArea.y + plotArea.height;
    const axisLeftX = plotArea.x;

    context.strokeStyle = "#94a3b8";
    context.lineWidth = 1;
    context.beginPath();
    context.moveTo(plotArea.x, axisBottomY);
    context.lineTo(plotArea.x + plotArea.width, axisBottomY);
    context.moveTo(axisLeftX, plotArea.y);
    context.lineTo(axisLeftX, plotArea.y + plotArea.height);
    context.stroke();

    context.fillStyle = "#334155";
    context.font = "11px Arial, sans-serif";

    const xTickCount = getAxisTickCount(figure.layout.xAxis, Math.max(2, Math.min(10, Math.round(plotArea.width / 90))), 2, 20);
    const xTicks = xScale
      .ticks(xTickCount)
      .map((tick) => {
        const x = xScale.map(tick);
        const categoryIndex = Math.round(tick);
        const categoryLabel = labels[categoryIndex] ?? categoryScale.label(categoryIndex);
        return {
          tick,
          x,
          label:
            xAxisType === "category" || xAxisType === "multicategory"
              ? categoryLabel
              : formatAxisTick(tick, figure.layout.xAxis, categoryLabel)
        };
      })
      .filter((entry) => entry.x >= plotArea.x - 1 && entry.x <= plotArea.x + plotArea.width + 1);

    context.strokeStyle = "#94a3b8";
    context.lineWidth = 1;
    xTicks.forEach((entry) => {
      context.beginPath();
      context.moveTo(entry.x, axisBottomY);
      context.lineTo(entry.x, axisBottomY + 5);
      context.stroke();
    });

    if (xAxisType === "multicategory") {
      const parsed = xTicks.map((entry) => ({ ...entry, ...splitMulticategoryLabel(entry.label) }));
      context.fillStyle = "#334155";
      context.font = "11px Arial, sans-serif";
      context.textAlign = "center";
      parsed.forEach((entry) => {
        context.fillText(entry.leaf, entry.x, axisBottomY + 16);
      });
      context.fillStyle = "#64748b";
      context.font = "10px Arial, sans-serif";
      let groupStart = 0;
      while (groupStart < parsed.length) {
        const groupName = parsed[groupStart].group;
        let groupEnd = groupStart;
        while (groupEnd + 1 < parsed.length && parsed[groupEnd + 1].group === groupName) {
          groupEnd += 1;
        }
        if (groupName) {
          context.fillText(groupName, (parsed[groupStart].x + parsed[groupEnd].x) / 2, axisBottomY + 30);
        }
        groupStart = groupEnd + 1;
      }
    } else {
      context.fillStyle = "#334155";
      context.font = "11px Arial, sans-serif";
      context.textAlign = "center";
      xTicks.forEach((entry) => {
        context.fillText(entry.label, entry.x, axisBottomY + 16);
      });
    }

    const yTickCount = getAxisTickCount(figure.layout.yAxis, Math.max(2, Math.min(10, Math.round(plotArea.height / 56))), 2, 20);
    yScale.ticks(yTickCount).forEach((tick) => {
      const y = yScale.map(tick);
      context.textAlign = "right";
      context.fillText(formatAxisTick(tick, figure.layout.yAxis), axisLeftX - 8, y + 3);
    });

    if (figure.layout.xAxis.title) {
      context.textAlign = "center";
      context.fillStyle = "#0f172a";
      context.font = "12px Arial, sans-serif";
      context.fillText(figure.layout.xAxis.title, plotArea.x + plotArea.width / 2, axisBottomY + (xAxisType === "multicategory" ? 50 : 34));
    }

    if (figure.layout.yAxis.title) {
      context.save();
      context.translate(plotArea.x - 42, plotArea.y + plotArea.height / 2);
      context.rotate(-Math.PI / 2);
      context.textAlign = "center";
      context.fillStyle = "#0f172a";
      context.font = "12px Arial, sans-serif";
      context.fillText(figure.layout.yAxis.title, 0, 0);
      context.restore();
    }

    if (yScaleSecondary) {
      const axisRightX = plotArea.x + plotArea.width;
      context.strokeStyle = "#94a3b8";
      context.lineWidth = 1;
      context.beginPath();
      context.moveTo(axisRightX, plotArea.y);
      context.lineTo(axisRightX, plotArea.y + plotArea.height);
      context.stroke();

      context.fillStyle = "#334155";
      context.font = "11px Arial, sans-serif";
      const yTickCount = getAxisTickCount(figure.layout.yAxis2, Math.max(2, Math.min(10, Math.round(plotArea.height / 56))), 2, 20);
      yScaleSecondary.ticks(yTickCount).forEach((tick) => {
        const y = yScaleSecondary.map(tick);
        context.textAlign = "left";
        context.fillText(formatAxisTick(tick, figure.layout.yAxis2), axisRightX + 8, y + 3);
      });

      if (figure.layout.yAxis2.title) {
        context.save();
        context.translate(plotArea.x + plotArea.width + 42, plotArea.y + plotArea.height / 2);
        context.rotate(Math.PI / 2);
        context.textAlign = "center";
        context.fillStyle = "#0f172a";
        context.font = "12px Arial, sans-serif";
        context.fillText(figure.layout.yAxis2.title, 0, 0);
        context.restore();
      }
    }

    if (xScaleSecondary) {
      context.strokeStyle = "#94a3b8";
      context.lineWidth = 1;
      context.beginPath();
      context.moveTo(plotArea.x, plotArea.y);
      context.lineTo(plotArea.x + plotArea.width, plotArea.y);
      context.stroke();

      const secondaryLabels = labelsSecondary ?? [];
      const secondaryCategories = categoryScaleSecondary ?? new CategoryScale(secondaryLabels, [plotArea.x, plotArea.x + plotArea.width]);
      const x2TickCount = getAxisTickCount(figure.layout.xAxis2, Math.max(2, Math.min(10, Math.round(plotArea.width / 90))), 2, 20);
      const topTicks = xScaleSecondary
        .ticks(x2TickCount)
        .map((tick) => {
          const x = xScaleSecondary.map(tick);
          const categoryIndex = Math.round(tick);
          const categoryLabel = secondaryLabels[categoryIndex] ?? secondaryCategories.label(categoryIndex);
          return {
            tick,
            x,
            label:
              xAxis2Type === "category" || xAxis2Type === "multicategory"
                ? categoryLabel
                : formatAxisTick(tick, figure.layout.xAxis2, categoryLabel)
          };
        })
        .filter((entry) => entry.x >= plotArea.x - 1 && entry.x <= plotArea.x + plotArea.width + 1);

      topTicks.forEach((entry) => {
        context.beginPath();
        context.moveTo(entry.x, plotArea.y);
        context.lineTo(entry.x, plotArea.y - 5);
        context.stroke();
      });

      if (xAxis2Type === "multicategory") {
        const parsed = topTicks.map((entry) => ({ ...entry, ...splitMulticategoryLabel(entry.label) }));
        context.fillStyle = "#334155";
        context.font = "10px Arial, sans-serif";
        context.textAlign = "center";
        parsed.forEach((entry) => {
          context.fillText(entry.leaf, entry.x, plotArea.y - 8);
        });
        context.fillStyle = "#64748b";
        let groupStart = 0;
        while (groupStart < parsed.length) {
          const groupName = parsed[groupStart].group;
          let groupEnd = groupStart;
          while (groupEnd + 1 < parsed.length && parsed[groupEnd + 1].group === groupName) {
            groupEnd += 1;
          }
          if (groupName) {
            context.fillText(groupName, (parsed[groupStart].x + parsed[groupEnd].x) / 2, plotArea.y - 22);
          }
          groupStart = groupEnd + 1;
        }
      } else {
        context.fillStyle = "#334155";
        context.font = "10px Arial, sans-serif";
        context.textAlign = "center";
        topTicks.forEach((entry) => {
          context.fillText(entry.label, entry.x, plotArea.y - 8);
        });
      }

      if (figure.layout.xAxis2.title) {
        context.textAlign = "center";
        context.fillStyle = "#0f172a";
        context.font = "12px Arial, sans-serif";
        context.fillText(figure.layout.xAxis2.title, plotArea.x + plotArea.width / 2, plotArea.y - (xAxis2Type === "multicategory" ? 36 : 20));
      }
    }
  }

  private addInteractiveTarget(
    data: {
      traceIndex: number;
      pointIndex: number;
      traceName?: string;
      x: string;
      y: string;
    },
    x: number,
    y: number,
    width: number,
    height: number
  ): void {
    const layer = this.ensureInteractionLayer();
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.chartInteractive = "point";
    button.dataset.traceIndex = String(data.traceIndex);
    button.dataset.pointIndex = String(data.pointIndex);
    button.dataset.traceName = data.traceName?.trim() || `Trace ${data.traceIndex + 1}`;
    button.dataset.pointX = data.x;
    button.dataset.pointY = data.y;
    button.setAttribute("aria-label", `${button.dataset.traceName} ${data.x} ${data.y}`);
    Object.assign(button.style, {
      position: "absolute",
      left: `${Math.max(0, x)}px`,
      top: `${Math.max(0, y)}px`,
      width: `${Math.max(6, width)}px`,
      height: `${Math.max(6, height)}px`,
      border: "none",
      background: "transparent",
      margin: "0",
      padding: "0",
      cursor: "pointer"
    });
    layer.append(button);
  }

  private resolveTraceMode(trace: CartesianTrace): string {
    if (trace.mode) {
      return trace.mode;
    }
    if (trace.type === "line" || trace.type === "area") {
      return "lines";
    }
    if (trace.type === "scatter") {
      return "markers";
    }
    return "lines+markers";
  }

  private resolveTraceColor(trace: ChartTrace, index: number): string {
    if (isCartesianTrace(trace)) {
      return resolveMarkerColor(trace.marker?.color, index, trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length]);
    }
    if (isHistogramTrace(trace)) {
      return trace.marker?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isBoxTrace(trace)) {
      return trace.marker?.color ?? trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isViolinTrace(trace)) {
      return trace.marker?.color ?? trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isDensityTrace(trace)) {
      return trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isHeatmapTrace(trace)) {
      const colors = trace.colorscale && trace.colorscale.length > 0 ? trace.colorscale : DEFAULT_HEATMAP_COLORS;
      return colors[Math.floor(colors.length / 2)];
    }
    if (isContourTrace(trace)) {
      const colors = trace.colorscale && trace.colorscale.length > 0 ? trace.colorscale : DEFAULT_HEATMAP_COLORS;
      return colors[Math.floor(colors.length / 2)];
    }
    if (isFinancialTrace(trace)) {
      return trace.line?.color ?? trace.increasing?.color ?? trace.decreasing?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isDonutTrace(trace)) {
      return trace.marker?.colors?.[0] ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isSunburstTrace(trace) || isTreemapTrace(trace)) {
      return trace.marker?.colors?.[0] ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isPolarTrace(trace)) {
      return trace.marker?.color ?? trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isTernaryTrace(trace)) {
      return trace.marker?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isGeoTrace(trace)) {
      const colors = trace.colorscale && trace.colorscale.length > 0 ? trace.colorscale : DEFAULT_SERIES_COLORS;
      return colors[index % colors.length];
    }
    if (isGeoScatterTrace(trace)) {
      return trace.marker?.color ?? trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isGeoLineTrace(trace)) {
      return trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isScatter3dTrace(trace)) {
      return trace.marker?.color ?? trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isSurfaceTrace(trace)) {
      const palette = resolvePalette("continuous", trace.colorscale, trace.reverseScale === true);
      return palette[Math.floor(palette.length / 2)] ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isMesh3dTrace(trace)) {
      return trace.marker?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isWaterfallTrace(trace)) {
      return trace.increasing?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isFunnelTrace(trace)) {
      return trace.marker?.color?.[0] ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isQuiverTrace(trace)) {
      return trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isSankeyTrace(trace)) {
      return trace.nodes.colors?.[0] ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isParallelCategoriesTrace(trace)) {
      return trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    return DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
  }

  private ensureCanvas(): HTMLCanvasElement {
    if (!this.canvasElement) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_MOUNTED", "Canvas renderer must be mounted before rendering.");
    }
    return this.canvasElement;
  }

  private ensureContext(): CanvasRenderingContext2D {
    if (!this.context) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_MOUNTED", "Canvas renderer has no drawing context.");
    }
    return this.context;
  }

  private ensureInteractionLayer(): HTMLDivElement {
    if (!this.interactionLayer) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_MOUNTED", "Canvas renderer interaction layer is unavailable.");
    }
    return this.interactionLayer;
  }
}

const clampInt = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? Math.round(value) : min;
  return Math.min(max, Math.max(min, numeric));
};

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

const splitMulticategoryLabel = (label: string): { group: string; leaf: string } => {
  const normalized = String(label ?? "").trim();
  if (!normalized) {
    return { group: "", leaf: "" };
  }
  const delimiters = ["|", " / ", " > ", "::", ","];
  for (const delimiter of delimiters) {
    if (!normalized.includes(delimiter)) {
      continue;
    }
    const parts = normalized
      .split(delimiter)
      .map((part) => part.trim())
      .filter(Boolean);
    if (parts.length >= 2) {
      return {
        group: parts.slice(0, -1).join(" / "),
        leaf: parts[parts.length - 1]
      };
    }
  }
  return { group: "", leaf: normalized };
};

const resolveMarkerColor = (value: string | string[] | undefined, index: number, fallback: string): string => {
  if (Array.isArray(value)) {
    return value[index] ?? fallback;
  }
  return value ?? fallback;
};

const resolveMarkerSize = (value: number | number[] | undefined, index: number, fallback: number): number => {
  if (Array.isArray(value)) {
    const size = Number(value[index]);
    return Number.isFinite(size) && size > 0 ? size : fallback;
  }
  const size = Number(value);
  return Number.isFinite(size) && size > 0 ? size : fallback;
};

const resolveColor = (color: string, opacity: number): string => {
  if (/^#[0-9a-f]{6}$/i.test(color)) {
    const r = parseInt(color.slice(1, 3), 16);
    const g = parseInt(color.slice(3, 5), 16);
    const b = parseInt(color.slice(5, 7), 16);
    return `rgba(${r}, ${g}, ${b}, ${clamp(opacity, 0, 1)})`;
  }
  return color;
};

const buildContourSegments = (
  matrix: number[][],
  level: number
): Array<{ a: { x: number; y: number }; b: { x: number; y: number } }> => {
  const segments: Array<{ a: { x: number; y: number }; b: { x: number; y: number } }> = [];
  const rows = matrix.length;
  const cols = matrix[0]?.length ?? 0;
  for (let row = 0; row < rows - 1; row += 1) {
    for (let col = 0; col < cols - 1; col += 1) {
      const a = Number(matrix[row][col]);
      const b = Number(matrix[row][col + 1]);
      const c = Number(matrix[row + 1][col + 1]);
      const d = Number(matrix[row + 1][col]);
      if (![a, b, c, d].every((value) => Number.isFinite(value))) {
        continue;
      }
      const intersections: Array<{ x: number; y: number }> = [];
      addEdgeIntersection(intersections, col, row, col + 1, row, a, b, level);
      addEdgeIntersection(intersections, col + 1, row, col + 1, row + 1, b, c, level);
      addEdgeIntersection(intersections, col + 1, row + 1, col, row + 1, c, d, level);
      addEdgeIntersection(intersections, col, row + 1, col, row, d, a, level);
      if (intersections.length === 2) {
        segments.push({ a: intersections[0], b: intersections[1] });
      } else if (intersections.length === 4) {
        segments.push({ a: intersections[0], b: intersections[1] });
        segments.push({ a: intersections[2], b: intersections[3] });
      }
    }
  }
  return segments;
};

const addEdgeIntersection = (
  bucket: Array<{ x: number; y: number }>,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  v0: number,
  v1: number,
  level: number
): void => {
  const d0 = v0 - level;
  const d1 = v1 - level;
  if (d0 === 0 && d1 === 0) {
    return;
  }
  if ((d0 < 0 && d1 < 0) || (d0 > 0 && d1 > 0)) {
    return;
  }
  const ratio = Math.abs(v1 - v0) < Number.EPSILON ? 0.5 : (level - v0) / (v1 - v0);
  bucket.push({
    x: x0 + (x1 - x0) * ratio,
    y: y0 + (y1 - y0) * ratio
  });
};

const toRadians = (theta: unknown, index: number, total: number): number => {
  if (typeof theta === "number" && Number.isFinite(theta)) {
    return (theta / 180) * Math.PI;
  }
  const parsed = Number(theta);
  if (Number.isFinite(parsed)) {
    return (parsed / 180) * Math.PI;
  }
  return (index / Math.max(1, total)) * Math.PI * 2 - Math.PI / 2;
};

const expandGeoPaths = (coordinates: number[][][] | number[][][][]): number[][][] => {
  if (!Array.isArray(coordinates) || coordinates.length === 0) {
    return [];
  }
  if (Array.isArray(coordinates[0][0][0])) {
    const asMulti = coordinates as number[][][][];
    return asMulti.flatMap((polygon) => polygon.filter((ring) => ring.length >= 3));
  }
  const asPolygon = coordinates as number[][][];
  return asPolygon.filter((ring) => ring.length >= 3);
};

const computeGeoBounds = (
  polygons: Array<number[][][] | number[][][][]>,
  scatterPoints: Array<[number, number]>,
  linePoints: Array<[number, number]>
): { minLon: number; maxLon: number; minLat: number; maxLat: number } | null => {
  const allPoints: Array<[number, number]> = [...scatterPoints, ...linePoints];
  polygons.forEach((coordinates) => {
    expandGeoPaths(coordinates).forEach((path) => {
      path.forEach((coordinate) => {
        allPoints.push([Number(coordinate[0]), Number(coordinate[1])]);
      });
    });
  });
  if (allPoints.length === 0) {
    return null;
  }
  let minLon = Number.POSITIVE_INFINITY;
  let maxLon = Number.NEGATIVE_INFINITY;
  let minLat = Number.POSITIVE_INFINITY;
  let maxLat = Number.NEGATIVE_INFINITY;
  allPoints.forEach(([lon, lat]) => {
    if (!Number.isFinite(lon) || !Number.isFinite(lat)) {
      return;
    }
    minLon = Math.min(minLon, lon);
    maxLon = Math.max(maxLon, lon);
    minLat = Math.min(minLat, lat);
    maxLat = Math.max(maxLat, lat);
  });
  if (!Number.isFinite(minLon) || !Number.isFinite(maxLon) || !Number.isFinite(minLat) || !Number.isFinite(maxLat)) {
    return null;
  }
  if (minLon === maxLon) {
    minLon -= 0.5;
    maxLon += 0.5;
  }
  if (minLat === maxLat) {
    minLat -= 0.5;
    maxLat += 0.5;
  }
  return { minLon, maxLon, minLat, maxLat };
};

const mapGeoToPlot = (
  coordinate: number[],
  bounds: { minLon: number; maxLon: number; minLat: number; maxLat: number },
  plotArea: ComputedPlotArea
): { x: number; y: number } => {
  const lon = Number(coordinate[0]);
  const lat = Number(coordinate[1]);
  const xRatio = (lon - bounds.minLon) / Math.max(Number.EPSILON, bounds.maxLon - bounds.minLon);
  const yRatio = (lat - bounds.minLat) / Math.max(Number.EPSILON, bounds.maxLat - bounds.minLat);
  return {
    x: plotArea.x + clamp(xRatio, 0, 1) * plotArea.width,
    y: plotArea.y + (1 - clamp(yRatio, 0, 1)) * plotArea.height
  };
};

const simplifyGeoPath = (path: number[][], tolerance: number): number[][] => {
  if (!Number.isFinite(tolerance) || tolerance <= 0 || path.length <= 4) {
    return path;
  }
  const step = Math.max(2, Math.floor(tolerance));
  const simplified = path.filter((_, index) => index === 0 || index === path.length - 1 || index % step === 0);
  return simplified.length >= 3 ? simplified : path;
};

const distance = (x0: number, y0: number, x1: number, y1: number): number => Math.hypot(x1 - x0, y1 - y0);

const createDefaultProjection3d = (
  plotArea: ComputedPlotArea
): ((x: number, y: number, z: number) => { x: number; y: number }) => {
  const angleY = (38 * Math.PI) / 180;
  const angleX = (-30 * Math.PI) / 180;
  const scale = Math.min(plotArea.width, plotArea.height) * 0.34;
  const cx = plotArea.x + plotArea.width / 2;
  const cy = plotArea.y + plotArea.height / 2;
  const cosY = Math.cos(angleY);
  const sinY = Math.sin(angleY);
  const cosX = Math.cos(angleX);
  const sinX = Math.sin(angleX);

  return (x: number, y: number, z: number) => {
    const x1 = x * cosY - z * sinY;
    const z1 = x * sinY + z * cosY;
    const y2 = y * cosX - z1 * sinX;
    const z2 = y * sinX + z1 * cosX;
    const perspective = 1 / Math.max(0.2, 1 + z2 * 0.045);
    return {
      x: cx + x1 * scale * perspective,
      y: cy - y2 * scale * perspective
    };
  };
};
