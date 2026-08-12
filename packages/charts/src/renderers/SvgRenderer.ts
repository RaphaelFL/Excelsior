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
  TreemapTrace,
  TernaryTrace,
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
import { clipLineToRect, simplifyLine } from "../core/line-optimization";
import { resolveColorFromScale, resolvePalette } from "../core/color-scales";
import { formatAxisTick, getAxisTickCount, normalizeAxisType, toAxisScalar } from "../core/axis-utils";
import { buildCartesianStackContext } from "../core/stacking-utils";

const SVG_NS = "http://www.w3.org/2000/svg";
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

type CartesianPoint = {
  pointIndex: number;
  xPixel: number;
  yPixel: number;
  xLabel: string;
  yLabel: string;
};

export class SvgRenderer implements ChartRenderer {
  private container: HTMLElement | null = null;
  private svgElement: SVGSVGElement | null = null;
  private clipCounter = 0;
  private maxRenderPoints = 12_000;

  mount(container: HTMLElement): void {
    this.container = container;
    if (this.svgElement) {
      this.svgElement.remove();
    }

    this.svgElement = document.createElementNS(SVG_NS, "svg");
    this.svgElement.classList.add("excelsior-chart-svg");
    this.svgElement.setAttribute("role", "img");
    this.svgElement.setAttribute("aria-label", "Chart visualization");
    this.svgElement.style.touchAction = "none";
    container.append(this.svgElement);
  }

  render(figure: ChartFigure, layout: ComputedLayout): void {
    const svg = this.ensureSvg();
    this.maxRenderPoints = Math.max(500, Number(figure.config.maxRenderPoints ?? 12_000));
    svg.replaceChildren();
    svg.setAttribute("width", String(layout.width));
    svg.setAttribute("height", String(layout.height));
    svg.setAttribute("viewBox", `0 0 ${layout.width} ${layout.height}`);

    const background = createSvgElement("rect");
    background.setAttribute("x", "0");
    background.setAttribute("y", "0");
    background.setAttribute("width", String(layout.width));
    background.setAttribute("height", String(layout.height));
    background.setAttribute("fill", figure.layout.backgroundColor);
    svg.append(background);

    if (figure.layout.title) {
      const title = createSvgElement("text");
      title.textContent = figure.layout.title;
      title.setAttribute("x", String(layout.margin.left));
      title.setAttribute("y", String(layout.titleY));
      title.setAttribute("font-size", "16");
      title.setAttribute("font-family", "Arial, sans-serif");
      title.setAttribute("font-weight", "600");
      title.setAttribute("fill", "#0f172a");
      svg.append(title);
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
    const subplotAreas = layout.subplotAreas.length > 0 ? layout.subplotAreas : [layout.plotArea];

    if (cartesianTraces.length > 0) {
      this.renderSubplotGroups(cartesianTraces, subplotAreas, (group, plotArea, subplotIndex) =>
        this.renderCartesianScene(svg, figure, group, layout, plotArea, subplotIndex)
      );
    } else if (financialTraces.length > 0) {
      this.renderSubplotGroups(financialTraces, subplotAreas, (group, plotArea, subplotIndex) =>
        this.renderFinancialScene(svg, figure, group, layout, plotArea, subplotIndex)
      );
    } else if (histogramTraces.length > 0) {
      this.renderSubplotGroups(histogramTraces, subplotAreas, (group, plotArea) =>
        this.renderHistogramScene(svg, figure, group, layout, plotArea)
      );
    } else if (boxTraces.length > 0) {
      this.renderSubplotGroups(boxTraces, subplotAreas, (group, plotArea) => this.renderBoxScene(svg, figure, group, layout, plotArea));
    } else if (violinTraces.length > 0) {
      this.renderSubplotGroups(violinTraces, subplotAreas, (group, plotArea) => this.renderViolinScene(svg, figure, group, layout, plotArea));
    } else if (densityTraces.length > 0) {
      this.renderSubplotGroups(densityTraces, subplotAreas, (group, plotArea) => this.renderDensityScene(svg, figure, group, layout, plotArea));
    } else if (heatmapTraces.length > 0) {
      this.renderSubplotGroups(heatmapTraces, subplotAreas, (group, plotArea) => {
        this.renderHeatmapScene(svg, figure, group[0], layout, plotArea);
      });
    } else if (contourTraces.length > 0) {
      this.renderSubplotGroups(contourTraces, subplotAreas, (group, plotArea) => {
        this.renderContourScene(svg, figure, group[0], layout, plotArea);
      });
    } else if (quiverTraces.length > 0) {
      this.renderSubplotGroups(quiverTraces, subplotAreas, (group, plotArea, subplotIndex) => {
        this.renderQuiverScene(svg, figure, group, layout, plotArea, subplotIndex);
      });
    } else if (waterfallTraces.length > 0) {
      this.renderSubplotGroups(waterfallTraces, subplotAreas, (group, plotArea) => {
        this.renderWaterfallScene(svg, figure, group, layout, plotArea);
      });
    } else if (funnelTraces.length > 0) {
      this.renderSubplotGroups(funnelTraces, subplotAreas, (group, plotArea) => {
        this.renderFunnelScene(svg, group[0], plotArea);
      });
    } else if (pieTraces.length > 0) {
      this.renderSubplotGroups(pieTraces, subplotAreas, (group, plotArea, subplotIndex) => {
        this.renderPieTrace(svg, group[0], plotArea, subplotIndex);
      });
    } else if (donutTraces.length > 0) {
      this.renderSubplotGroups(donutTraces, subplotAreas, (group, plotArea, subplotIndex) => {
        this.renderDonutTrace(svg, group[0], plotArea, subplotIndex);
      });
    } else if (sunburstTraces.length > 0) {
      this.renderSubplotGroups(sunburstTraces, subplotAreas, (group, plotArea, subplotIndex) => {
        this.renderSunburstTrace(svg, group[0], plotArea, subplotIndex);
      });
    } else if (treemapTraces.length > 0) {
      this.renderSubplotGroups(treemapTraces, subplotAreas, (group, plotArea, subplotIndex) => {
        this.renderTreemapTrace(svg, group[0], plotArea, subplotIndex);
      });
    } else if (sankeyTraces.length > 0) {
      this.renderSubplotGroups(sankeyTraces, subplotAreas, (group, plotArea) => {
        this.renderSankeyScene(svg, group[0], plotArea);
      });
    } else if (parallelCategoryTraces.length > 0) {
      this.renderSubplotGroups(parallelCategoryTraces, subplotAreas, (group, plotArea) => {
        this.renderParallelCategoriesScene(svg, figure, group[0], plotArea);
      });
    } else if (polarTraces.length > 0) {
      this.renderSubplotGroups(polarTraces, subplotAreas, (group, plotArea, subplotIndex) =>
        this.renderPolarScene(svg, group, plotArea, subplotIndex)
      );
    } else if (ternaryTraces.length > 0) {
      this.renderSubplotGroups(ternaryTraces, subplotAreas, (group, plotArea, subplotIndex) =>
        this.renderTernaryScene(svg, group, plotArea, subplotIndex)
      );
    } else if (geoTraces.length > 0) {
      this.renderSubplotGroups(geoTraces, subplotAreas, (group, plotArea, subplotIndex) => {
        const scatter = geoScatterTraces.filter((entry) => ((entry.trace.subplot as number | undefined) ?? 0) === subplotIndex);
        const lines = geoLineTraces.filter((entry) => ((entry.trace.subplot as number | undefined) ?? 0) === subplotIndex);
        this.renderGeoScene(svg, group, plotArea, subplotIndex, scatter, lines);
      });
    } else if (geoScatterTraces.length > 0 || geoLineTraces.length > 0) {
      this.renderSubplotGroups(geoScatterTraces, subplotAreas, (group, plotArea, subplotIndex) => {
        const lines = geoLineTraces.filter((entry) => ((entry.trace.subplot as number | undefined) ?? 0) === subplotIndex);
        this.renderGeoScene(svg, [], plotArea, subplotIndex, group, lines);
      });
    } else if (scatter3dTraces.length > 0 || surfaceTraces.length > 0 || mesh3dTraces.length > 0) {
      this.renderProjected3dScene(svg, scatter3dTraces, surfaceTraces, mesh3dTraces, layout.plotArea);
    }

    if (figure.layout.legend.visible) {
      this.renderLegend(svg, allTraces, layout);
    }

    this.renderLayoutOverlays(svg, figure, layout);
  }

  resize(layout: ComputedLayout): void {
    const svg = this.ensureSvg();
    svg.setAttribute("width", String(layout.width));
    svg.setAttribute("height", String(layout.height));
    svg.setAttribute("viewBox", `0 0 ${layout.width} ${layout.height}`);
  }

  getRootElement(): Element | null {
    return this.svgElement;
  }

  exportSvg(): string {
    return this.ensureSvg().outerHTML;
  }

  async exportPng(options?: { scale?: number; backgroundColor?: string }): Promise<Blob> {
    const svg = this.ensureSvg();
    if (typeof Image === "undefined") {
      throw new ChartConfigurationError("CHART_EXPORT_UNAVAILABLE", "Image export is unavailable in this environment.");
    }

    const width = Number(svg.getAttribute("width") ?? "0");
    const height = Number(svg.getAttribute("height") ?? "0");
    const scale = Math.max(1, options?.scale ?? 1);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(width * scale));
    canvas.height = Math.max(1, Math.round(height * scale));
    const context = canvas.getContext("2d");
    if (!context) {
      throw new ChartConfigurationError("CHART_EXPORT_UNAVAILABLE", "Canvas 2D context is unavailable for PNG export.");
    }

    if (options?.backgroundColor) {
      context.fillStyle = options.backgroundColor;
      context.fillRect(0, 0, canvas.width, canvas.height);
    }

    const svgMarkup = this.exportSvg();
    const encodedSvg = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgMarkup)}`;
    const image = new Image();
    const blob = await new Promise<Blob>((resolve, reject) => {
      image.onload = () => {
        context.drawImage(image, 0, 0, canvas.width, canvas.height);
        canvas.toBlob((output) => {
          if (!output) {
            reject(new ChartConfigurationError("CHART_EXPORT_FAILED", "Failed to convert chart canvas to PNG."));
            return;
          }
          resolve(output);
        }, "image/png");
      };
      image.onerror = () => reject(new ChartConfigurationError("CHART_EXPORT_FAILED", "Failed to load SVG image for PNG export."));
      image.src = encodedSvg;
    });

    return blob;
  }

  destroy(): void {
    if (this.svgElement) {
      this.svgElement.remove();
    }
    this.svgElement = null;
    this.container = null;
  }

  private renderCartesianScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traces: IndexedTrace<CartesianTrace>[],
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea,
    subplotIndex?: number
  ): void {
    const plotArea = plotAreaOverride ?? layout.plotArea;
    const primaryYTraces = traces.filter((entry) => entry.trace.yAxisRef !== "y2");
    const secondaryYTraces = traces.filter((entry) => entry.trace.yAxisRef === "y2");
    const primaryXTraces = traces.filter((entry) => entry.trace.xAxisRef !== "x2");
    const secondaryXTraces = traces.filter((entry) => entry.trace.xAxisRef === "x2");

    const primaryDomains = this.resolveCartesianDomains(figure, traces, {
      xAxisKey: "xAxis",
      yAxisKey: "yAxis",
      subplotIndex
    });
    if (!primaryDomains) {
      return;
    }
    const primaryXDomains =
      this.resolveCartesianDomains(figure, primaryXTraces.length > 0 ? primaryXTraces : traces, {
        xAxisKey: "xAxis",
        yAxisKey: "yAxis",
        subplotIndex
      }) ?? primaryDomains;
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
      }) ?? primaryDomains;
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
    const categoryScale = new CategoryScale(primaryXDomains.labels, [plotArea.x, plotArea.x + plotArea.width]);
    const categoryScaleSecondary = secondaryXDomains
      ? new CategoryScale(secondaryXDomains.labels, [plotArea.x, plotArea.x + plotArea.width])
      : null;

    this.renderPlotBackground(svg, plotArea);
    this.renderCartesianGrid(svg, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const tracesLayer = this.createClippedLayer(svg, plotArea);

    const bars = traces.filter((entry) => entry.trace.type === "bar");
    const unstackedBars = bars.filter((entry) => !entry.trace.stackGroup);
    const barCount = Math.max(1, unstackedBars.length);
    const stackContext = buildCartesianStackContext(traces);
    const previousAreaByGroup = new Map<string, CartesianPoint[]>();

    traces.forEach((entry) => {
      const traceGroup = createSvgElement("g");
      traceGroup.setAttribute("data-trace", entry.trace.type);
      traceGroup.setAttribute("data-trace-index", String(entry.index));
      const color = this.resolveTraceColor(entry.trace, entry.index);

      const stackSeries = stackContext.get(entry.index);
      const targetXScale = entry.trace.xAxisRef === "x2" && xScaleSecondary ? xScaleSecondary : xScale;
      const targetXAxisType = entry.trace.xAxisRef === "x2" && xScaleSecondary ? xAxis2Type : xAxisType;
      if (entry.trace.type === "bar") {
        const barIndex = unstackedBars.findIndex((candidate) => candidate.index === entry.index);
        const targetYScale = entry.trace.yAxisRef === "y2" && yScaleSecondary ? yScaleSecondary : yScale;
        this.renderBarTrace(
          traceGroup,
          entry,
          plotArea,
          targetXScale,
          targetYScale,
          color,
          Math.max(0, barIndex),
          barCount,
          targetXAxisType,
          stackSeries
        );
      } else {
        const targetYScale = entry.trace.yAxisRef === "y2" && yScaleSecondary ? yScaleSecondary : yScale;
        const areaKey = entry.trace.stackGroup?.trim() || "default";
        const previousArea = previousAreaByGroup.get(areaKey);
        const points = this.renderLineLikeTrace(
          traceGroup,
          entry,
          targetXScale,
          targetYScale,
          color,
          plotArea,
          targetXAxisType,
          stackSeries,
          previousArea
        );
        if (entry.trace.type === "area" && points.length > 0) {
          previousAreaByGroup.set(areaKey, points);
        }
      }

      tracesLayer.append(traceGroup);
    });

    svg.append(tracesLayer);
    this.renderCartesianAxes(
      svg,
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

  private renderHistogramScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traces: IndexedTrace<HistogramTrace>[],
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea
  ): void {
    const plotArea = plotAreaOverride ?? layout.plotArea;
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
    const categoryScale = new CategoryScale(labels, [plotArea.x, plotArea.x + plotArea.width]);

    this.renderPlotBackground(svg, plotArea);
    this.renderCartesianGrid(svg, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const tracesLayer = this.createClippedLayer(svg, plotArea);
    const barCount = Math.max(1, entries.length);
    const slotWidth = plotArea.width / Math.max(1, targetBins);
    const barSlotWidth = Math.max(6, slotWidth * 0.82);
    const barWidth = Math.max(2, barSlotWidth / barCount);

    entries.forEach((data, tracePosition) => {
      const traceGroup = createSvgElement("g");
      traceGroup.setAttribute("data-trace", "histogram");
      traceGroup.setAttribute("data-trace-index", String(data.entry.index));
      const color = this.resolveTraceColor(data.entry.trace, data.entry.index);

      data.bins.forEach((bin) => {
        const rect = createSvgElement("rect");
        if (horizontal) {
          const centerY = yScale.map(bin.index);
          const y = centerY - barSlotWidth / 2 + barWidth * tracePosition;
          const x0 = xScale.map(0);
          const x1 = xScale.map(bin.count);
          rect.setAttribute("x", String(Math.min(x0, x1)));
          rect.setAttribute("y", String(y));
          rect.setAttribute("width", String(Math.max(1, Math.abs(x1 - x0))));
          rect.setAttribute("height", String(barWidth));
        } else {
          const centerX = xScale.map(bin.index);
          const x = centerX - barSlotWidth / 2 + barWidth * tracePosition;
          const y = yScale.map(bin.count);
          const baseY = yScale.map(0);
          const top = Math.min(y, baseY);
          const height = Math.max(1, Math.abs(baseY - y));
          rect.setAttribute("x", String(x));
          rect.setAttribute("y", String(top));
          rect.setAttribute("width", String(barWidth));
          rect.setAttribute("height", String(height));
        }
        rect.setAttribute("fill", color);
        rect.setAttribute("opacity", String(data.entry.trace.marker?.opacity ?? 0.9));
        this.attachInteractiveMetadata(rect, data.entry.index, bin.index, data.entry.trace.name, bin.label, formatNumeric(bin.count));
        traceGroup.append(rect);
      });

      tracesLayer.append(traceGroup);
    });

    svg.append(tracesLayer);
    this.renderCartesianAxes(svg, figure, plotArea, labels, xScale, yScale, categoryScale);
  }

  private renderFinancialScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traces: IndexedTrace<FinancialTrace>[],
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea,
    subplotIndex?: number
  ): void {
    const plotArea = plotAreaOverride ?? layout.plotArea;
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

    this.renderPlotBackground(svg, plotArea);
    this.renderCartesianGrid(svg, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const tracesLayer = this.createClippedLayer(svg, plotArea);

    const maxPoints = Math.max(...traces.map((entry) => entry.trace.x.length));
    const slotWidth = plotArea.width / Math.max(1, maxPoints);
    const bodyWidth = Math.max(2, Math.min(22, (slotWidth * 0.66) / Math.max(1, traces.length)));

    traces.forEach((entry, tracePosition) => {
      const group = createSvgElement("g");
      group.setAttribute("data-trace", entry.trace.type);
      group.setAttribute("data-trace-index", String(entry.index));
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

        const wick = createSvgElement("line");
        wick.setAttribute("x1", String(centerX));
        wick.setAttribute("x2", String(centerX));
        wick.setAttribute("y1", String(yHigh));
        wick.setAttribute("y2", String(yLow));
        wick.setAttribute("stroke", strokeColor);
        wick.setAttribute("stroke-width", String(lineWidth));
        wick.setAttribute("opacity", String(lineOpacity));
        group.append(wick);

        if (entry.trace.type === "candlestick") {
          const top = Math.min(yOpen, yClose);
          const height = Math.max(1, Math.abs(yClose - yOpen));
          const body = createSvgElement("rect");
          body.setAttribute("x", String(centerX - bodyWidth / 2));
          body.setAttribute("y", String(top));
          body.setAttribute("width", String(bodyWidth));
          body.setAttribute("height", String(height));
          body.setAttribute("fill", increasing ? upColor : downColor);
          body.setAttribute("opacity", String(lineOpacity));
          body.setAttribute("stroke", strokeColor);
          body.setAttribute("stroke-width", String(Math.max(1, lineWidth - 0.2)));
          this.attachInteractiveMetadata(
            body,
            entry.index,
            pointIndex,
            entry.trace.name,
            String(entry.trace.x[pointIndex] ?? pointIndex),
            `O ${formatNumeric(open)} H ${formatNumeric(high)} L ${formatNumeric(low)} C ${formatNumeric(close)}`
          );
          group.append(body);
        } else {
          const openTick = createSvgElement("line");
          openTick.setAttribute("x1", String(centerX - bodyWidth / 2));
          openTick.setAttribute("x2", String(centerX));
          openTick.setAttribute("y1", String(yOpen));
          openTick.setAttribute("y2", String(yOpen));
          openTick.setAttribute("stroke", strokeColor);
          openTick.setAttribute("stroke-width", String(lineWidth));
          group.append(openTick);

          const closeTick = createSvgElement("line");
          closeTick.setAttribute("x1", String(centerX));
          closeTick.setAttribute("x2", String(centerX + bodyWidth / 2));
          closeTick.setAttribute("y1", String(yClose));
          closeTick.setAttribute("y2", String(yClose));
          closeTick.setAttribute("stroke", strokeColor);
          closeTick.setAttribute("stroke-width", String(lineWidth));
          this.attachInteractiveMetadata(
            closeTick,
            entry.index,
            pointIndex,
            entry.trace.name,
            String(entry.trace.x[pointIndex] ?? pointIndex),
            `O ${formatNumeric(open)} H ${formatNumeric(high)} L ${formatNumeric(low)} C ${formatNumeric(close)}`
          );
          group.append(closeTick);
        }
      }

      tracesLayer.append(group);
    });

    svg.append(tracesLayer);
    this.renderCartesianAxes(
      svg,
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

  private renderBoxScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traces: IndexedTrace<BoxTrace>[],
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea
  ): void {
    const statsEntries = traces
      .map((entry) => ({ entry, stats: computeBoxStats(entry.trace) }))
      .filter((entry): entry is { entry: IndexedTrace<BoxTrace>; stats: NonNullable<ReturnType<typeof computeBoxStats>> } => entry.stats !== null);

    if (statsEntries.length === 0) {
      return;
    }

    const plotArea = plotAreaOverride ?? layout.plotArea;
    const labels = statsEntries.map((entry, index) => entry.entry.trace.name?.trim() || `Box ${index + 1}`);
    const xScale = new LinearScale([0, Math.max(0, statsEntries.length - 1)], [plotArea.x, plotArea.x + plotArea.width]);
    const categoryScale = new CategoryScale(labels, [plotArea.x, plotArea.x + plotArea.width]);
    const values = statsEntries.flatMap((entry) => [entry.stats.min, entry.stats.q1, entry.stats.median, entry.stats.q3, entry.stats.max]);
    const minValue = Math.min(...values);
    const maxValue = Math.max(...values);
    const yScale = new LinearScale(
      [Math.min(0, minValue), maxValue === minValue ? maxValue + 1 : maxValue],
      [plotArea.y + plotArea.height, plotArea.y]
    );

    this.renderPlotBackground(svg, plotArea);
    this.renderCartesianGrid(svg, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const tracesLayer = this.createClippedLayer(svg, plotArea);
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

      const group = createSvgElement("g");
      group.setAttribute("data-trace", "box");
      group.setAttribute("data-trace-index", String(entry.index));

      const whisker = createSvgElement("line");
      whisker.setAttribute("x1", String(centerX));
      whisker.setAttribute("x2", String(centerX));
      whisker.setAttribute("y1", String(yMin));
      whisker.setAttribute("y2", String(yMax));
      whisker.setAttribute("stroke", entry.trace.line?.color ?? color);
      whisker.setAttribute("stroke-width", String(entry.trace.line?.width ?? 2));
      group.append(whisker);

      const boxTop = Math.min(yQ1, yQ3);
      const boxHeight = Math.max(1, Math.abs(yQ3 - yQ1));
      const boxRect = createSvgElement("rect");
      boxRect.setAttribute("x", String(centerX - boxWidth / 2));
      boxRect.setAttribute("y", String(boxTop));
      boxRect.setAttribute("width", String(boxWidth));
      boxRect.setAttribute("height", String(boxHeight));
      boxRect.setAttribute("fill", color);
      boxRect.setAttribute("opacity", String(entry.trace.marker?.opacity ?? 0.48));
      boxRect.setAttribute("stroke", entry.trace.line?.color ?? color);
      boxRect.setAttribute("stroke-width", String(entry.trace.line?.width ?? 2));
      this.attachInteractiveMetadata(
        boxRect,
        entry.index,
        0,
        entry.trace.name,
        labels[tracePosition],
        `mediana ${formatNumeric(stats.median)}`
      );
      group.append(boxRect);

      const medianLine = createSvgElement("line");
      medianLine.setAttribute("x1", String(centerX - boxWidth / 2));
      medianLine.setAttribute("x2", String(centerX + boxWidth / 2));
      medianLine.setAttribute("y1", String(yMedian));
      medianLine.setAttribute("y2", String(yMedian));
      medianLine.setAttribute("stroke", entry.trace.line?.color ?? "#0f172a");
      medianLine.setAttribute("stroke-width", "2");
      group.append(medianLine);

      if (entry.trace.boxpoints) {
        const numericValues = entry.trace.values.map((value) => Number(value)).filter((value) => Number.isFinite(value));
        numericValues.forEach((value, pointIndex) => {
          const jitter = ((pointIndex % 5) - 2) * 2;
          const marker = createSvgElement("circle");
          marker.setAttribute("cx", String(centerX + jitter));
          marker.setAttribute("cy", String(yScale.map(value)));
          marker.setAttribute("r", "2.5");
          marker.setAttribute("fill", entry.trace.marker?.color ?? color);
          marker.setAttribute("opacity", "0.7");
          this.attachInteractiveMetadata(
            marker,
            entry.index,
            pointIndex,
            entry.trace.name,
            labels[tracePosition],
            formatNumeric(value)
          );
          group.append(marker);
        });
      }

      if (entry.trace.showOutliers) {
        const factor = Number(entry.trace.outlierIqrFactor ?? 1.5);
        const outliers = computeBoxOutliers(entry.trace, Number.isFinite(factor) && factor > 0 ? factor : 1.5);
        outliers.forEach((outlier, pointIndex) => {
          const marker = createSvgElement("circle");
          const jitter = ((pointIndex % 5) - 2) * 2.4;
          marker.setAttribute("cx", String(centerX + jitter));
          marker.setAttribute("cy", String(yScale.map(outlier)));
          marker.setAttribute("r", "2.8");
          marker.setAttribute("fill", "none");
          marker.setAttribute("stroke", entry.trace.line?.color ?? color);
          marker.setAttribute("stroke-width", "1.2");
          marker.setAttribute("opacity", "0.9");
          this.attachInteractiveMetadata(marker, entry.index, pointIndex, entry.trace.name, labels[tracePosition], `outlier ${formatNumeric(outlier)}`);
          group.append(marker);
        });
      }

      tracesLayer.append(group);
    });

    svg.append(tracesLayer);
    this.renderCartesianAxes(svg, figure, plotArea, labels, xScale, yScale, categoryScale);
  }

  private renderViolinScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traces: IndexedTrace<ViolinTrace>[],
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea
  ): void {
    if (traces.length === 0) {
      return;
    }
    const plotArea = plotAreaOverride ?? layout.plotArea;
    const labels = traces.map((entry, index) => entry.trace.name?.trim() || `Violin ${index + 1}`);
    const xScale = new LinearScale([0, Math.max(0, traces.length - 1)], [plotArea.x, plotArea.x + plotArea.width]);
    const categoryScale = new CategoryScale(labels, [plotArea.x, plotArea.x + plotArea.width]);
    const allValues = traces.flatMap((entry) => entry.trace.values.map((value) => Number(value)).filter((value) => Number.isFinite(value)));
    if (allValues.length === 0) {
      return;
    }
    const yMin = Math.min(...allValues);
    const yMax = Math.max(...allValues);
    const yScale = new LinearScale([yMin, yMin === yMax ? yMax + 1 : yMax], [plotArea.y + plotArea.height, plotArea.y]);
    this.renderPlotBackground(svg, plotArea);
    this.renderCartesianGrid(svg, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const tracesLayer = this.createClippedLayer(svg, plotArea);
    const slotWidth = plotArea.width / Math.max(1, traces.length);

    traces.forEach((entry, traceIndex) => {
      const profile = computeViolinProfile(entry.trace.values, {
        bandwidth: entry.trace.bandwidth,
        sampleLimit: this.maxRenderPoints
      });
      if (!profile || profile.points.length < 2) {
        return;
      }
      const centerX = xScale.map(traceIndex);
      const halfWidth = Math.max(6, slotWidth * 0.34);
      const upper: Array<{ x: number; y: number }> = [];
      const lower: Array<{ x: number; y: number }> = [];
      profile.points.forEach((point) => {
        const y = yScale.map(point.x);
        const spread = point.y * halfWidth;
        upper.push({ x: centerX + spread, y });
        lower.push({ x: centerX - spread, y });
      });
      const shape = createSvgElement("path");
      const pathData =
        `M ${lower[0].x.toFixed(2)} ${lower[0].y.toFixed(2)} ` +
        lower.slice(1).map((point) => `L ${point.x.toFixed(2)} ${point.y.toFixed(2)}`).join(" ") +
        " " +
        upper
          .slice()
          .reverse()
          .map((point) => `L ${point.x.toFixed(2)} ${point.y.toFixed(2)}`)
          .join(" ") +
        " Z";
      const color = this.resolveTraceColor(entry.trace, entry.index);
      shape.setAttribute("d", pathData);
      shape.setAttribute("fill", resolveColor(color, 0.28));
      shape.setAttribute("stroke", entry.trace.line?.color ?? color);
      shape.setAttribute("stroke-width", String(entry.trace.line?.width ?? 1.6));
      shape.setAttribute("opacity", String(entry.trace.marker?.opacity ?? entry.trace.line?.opacity ?? 1));
      this.attachInteractiveMetadata(shape, entry.index, 0, entry.trace.name, labels[traceIndex], `${profile.points.length} samples`);
      tracesLayer.append(shape);

      if (entry.trace.showBox) {
        const stats = computeBoxStats({ ...entry.trace, type: "box" });
        if (stats) {
          const box = createSvgElement("rect");
          const boxWidth = Math.max(6, halfWidth * 0.45);
          const y0 = yScale.map(stats.q1);
          const y1 = yScale.map(stats.q3);
          box.setAttribute("x", String(centerX - boxWidth / 2));
          box.setAttribute("y", String(Math.min(y0, y1)));
          box.setAttribute("width", String(boxWidth));
          box.setAttribute("height", String(Math.max(1, Math.abs(y1 - y0))));
          box.setAttribute("fill", "#ffffff");
          box.setAttribute("stroke", entry.trace.line?.color ?? color);
          box.setAttribute("stroke-width", "1.2");
          tracesLayer.append(box);

          const median = createSvgElement("line");
          const medianY = yScale.map(stats.median);
          median.setAttribute("x1", String(centerX - boxWidth / 2));
          median.setAttribute("x2", String(centerX + boxWidth / 2));
          median.setAttribute("y1", String(medianY));
          median.setAttribute("y2", String(medianY));
          median.setAttribute("stroke", "#0f172a");
          median.setAttribute("stroke-width", "1.4");
          tracesLayer.append(median);
        }
      }

      if (entry.trace.showPoints) {
        const points = entry.trace.values.map((value) => Number(value)).filter((value) => Number.isFinite(value));
        points.forEach((value, pointIndex) => {
          const marker = createSvgElement("circle");
          const jitter = ((pointIndex % 7) - 3) * 1.8;
          marker.setAttribute("cx", String(centerX + jitter));
          marker.setAttribute("cy", String(yScale.map(value)));
          marker.setAttribute("r", "1.8");
          marker.setAttribute("fill", entry.trace.marker?.color ?? color);
          marker.setAttribute("opacity", "0.65");
          tracesLayer.append(marker);
        });
      }
    });

    svg.append(tracesLayer);
    this.renderCartesianAxes(svg, figure, plotArea, labels, xScale, yScale, categoryScale);
  }

  private renderDensityScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traces: IndexedTrace<DensityTrace>[],
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea
  ): void {
    if (traces.length === 0) {
      return;
    }
    const plotArea = plotAreaOverride ?? layout.plotArea;
    const curves = traces
      .map((entry) => {
        const cumulative = entry.trace.cumulative === true || entry.trace.type === "distribution";
        const points = computeDensityCurve(entry.trace.values, {
          cumulative,
          sampleLimit: entry.trace.sampleLimit ?? figure.config.maxDensitySamples ?? this.maxRenderPoints
        });
        return { entry, points };
      })
      .filter((entry) => entry.points.length >= 2);
    if (curves.length === 0) {
      return;
    }
    const xMin = Math.min(...curves.flatMap((curve) => curve.points.map((point) => point.x)));
    const xMax = Math.max(...curves.flatMap((curve) => curve.points.map((point) => point.x)));
    const yMax = Math.max(...curves.flatMap((curve) => curve.points.map((point) => point.y)), 1);
    const xScale = new LinearScale([xMin, xMin === xMax ? xMax + 1 : xMax], [plotArea.x, plotArea.x + plotArea.width]);
    const yScale = new LinearScale([0, yMax * 1.05], [plotArea.y + plotArea.height, plotArea.y]);
    const labels = curves[0].points.map((point) => formatNumeric(point.x));
    const categoryScale = new CategoryScale(labels, [plotArea.x, plotArea.x + plotArea.width]);
    this.renderPlotBackground(svg, plotArea);
    this.renderCartesianGrid(svg, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const tracesLayer = this.createClippedLayer(svg, plotArea);

    curves.forEach((curve, index) => {
      const color = this.resolveTraceColor(curve.entry.trace, curve.entry.index);
      const path = createSvgElement("path");
      const d = curve.points
        .map((point, pointIndex) => `${pointIndex === 0 ? "M" : "L"} ${xScale.map(point.x).toFixed(2)} ${yScale.map(point.y).toFixed(2)}`)
        .join(" ");
      path.setAttribute("d", d);
      path.setAttribute("fill", "none");
      path.setAttribute("stroke", curve.entry.trace.line?.color ?? color);
      path.setAttribute("stroke-width", String(curve.entry.trace.line?.width ?? 2));
      path.setAttribute("opacity", String(curve.entry.trace.line?.opacity ?? 0.95));
      tracesLayer.append(path);
      if (index === 0) {
        curve.points.forEach((point, pointIndex) => {
          if (pointIndex % 8 !== 0) {
            return;
          }
          const marker = createSvgElement("circle");
          marker.setAttribute("cx", String(xScale.map(point.x)));
          marker.setAttribute("cy", String(yScale.map(point.y)));
          marker.setAttribute("r", "2");
          marker.setAttribute("fill", color);
          this.attachInteractiveMetadata(marker, curve.entry.index, pointIndex, curve.entry.trace.name, formatNumeric(point.x), formatNumeric(point.y));
          tracesLayer.append(marker);
        });
      }
    });

    svg.append(tracesLayer);
    this.renderCartesianAxes(svg, figure, plotArea, labels, xScale, yScale, categoryScale);
  }

  private renderQuiverScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traces: IndexedTrace<QuiverTrace>[],
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea,
    colorOffset = 0
  ): void {
    if (traces.length === 0) {
      return;
    }
    const plotArea = plotAreaOverride ?? layout.plotArea;
    const xValues = traces.flatMap((entry) => entry.trace.x.flatMap((value, index) => [Number(value), Number(value) + Number(entry.trace.u[index] ?? 0)]));
    const yValues = traces.flatMap((entry) => entry.trace.y.flatMap((value, index) => [Number(value), Number(value) + Number(entry.trace.v[index] ?? 0)]));
    const finiteX = xValues.filter((value) => Number.isFinite(value));
    const finiteY = yValues.filter((value) => Number.isFinite(value));
    if (finiteX.length === 0 || finiteY.length === 0) {
      return;
    }
    const xMin = Math.min(...finiteX);
    const xMax = Math.max(...finiteX);
    const yMin = Math.min(...finiteY);
    const yMax = Math.max(...finiteY);
    const xScale = new LinearScale([xMin, xMin === xMax ? xMax + 1 : xMax], [plotArea.x, plotArea.x + plotArea.width]);
    const yScale = new LinearScale([yMin, yMin === yMax ? yMax + 1 : yMax], [plotArea.y + plotArea.height, plotArea.y]);
    const categoryScale = new CategoryScale([], [plotArea.x, plotArea.x + plotArea.width]);
    this.renderPlotBackground(svg, plotArea);
    this.renderCartesianGrid(svg, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const tracesLayer = this.createClippedLayer(svg, plotArea);
    const magnitudes = traces.flatMap((entry) =>
      entry.trace.u.map((uValue, index) => Math.hypot(Number(uValue), Number(entry.trace.v[index])))
    );
    const finiteMagnitudes = magnitudes.filter((value) => Number.isFinite(value));
    const minMagnitude = finiteMagnitudes.length > 0 ? Math.min(...finiteMagnitudes) : 0;
    const maxMagnitude = finiteMagnitudes.length > 0 ? Math.max(...finiteMagnitudes) : 1;

    traces.forEach((entry, traceIndex) => {
      const color = entry.trace.line?.color ?? DEFAULT_SERIES_COLORS[(traceIndex + colorOffset) % DEFAULT_SERIES_COLORS.length];
      const strokeWidth = entry.trace.line?.width ?? 1.2;
      const opacity = entry.trace.line?.opacity ?? 0.9;
      const scale = Number.isFinite(entry.trace.scale) ? Number(entry.trace.scale) : 1;
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
        const arrow = createSvgElement("line");
        arrow.setAttribute("x1", String(x0));
        arrow.setAttribute("y1", String(y0));
        arrow.setAttribute("x2", String(x1));
        arrow.setAttribute("y2", String(y1));
        arrow.setAttribute("stroke", pointColor);
        arrow.setAttribute("stroke-width", String(strokeWidth));
        arrow.setAttribute("opacity", String(opacity));
        tracesLayer.append(arrow);

        const head = createSvgElement("path");
        const angle = Math.atan2(y1 - y0, x1 - x0);
        const len = 7;
        const wing = 0.6;
        const hx1 = x1 - Math.cos(angle - wing) * len;
        const hy1 = y1 - Math.sin(angle - wing) * len;
        const hx2 = x1 - Math.cos(angle + wing) * len;
        const hy2 = y1 - Math.sin(angle + wing) * len;
        head.setAttribute("d", `M ${x1.toFixed(2)} ${y1.toFixed(2)} L ${hx1.toFixed(2)} ${hy1.toFixed(2)} L ${hx2.toFixed(2)} ${hy2.toFixed(2)} Z`);
        head.setAttribute("fill", pointColor);
        head.setAttribute("opacity", String(opacity));
        this.attachInteractiveMetadata(
          head,
          entry.index,
          pointIndex,
          entry.trace.name,
          `x=${formatNumeric(x)} y=${formatNumeric(y)}`,
          `u=${formatNumeric(u)} v=${formatNumeric(v)} |m|=${formatNumeric(magnitude)}`
        );
        tracesLayer.append(head);
      }
    });

    svg.append(tracesLayer);
    this.renderCartesianAxes(svg, figure, plotArea, [], xScale, yScale, categoryScale);
  }

  private renderHeatmapScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traceEntry: IndexedTrace<HeatmapTrace>,
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea
  ): void {
    const matrix = computeHeatmapMatrix(traceEntry.trace);
    if (!matrix) {
      return;
    }

    const plotArea = plotAreaOverride ?? layout.plotArea;
    const xLabels = Array.from({ length: matrix.cols }, (_, col) => String(traceEntry.trace.x?.[col] ?? `C${col + 1}`));
    const yLabels = Array.from({ length: matrix.rows }, (_, row) => String(traceEntry.trace.y?.[row] ?? `R${row + 1}`));
    const xScale = new LinearScale([0, Math.max(0, matrix.cols - 1)], [plotArea.x, plotArea.x + plotArea.width]);
    const yScale = new LinearScale([Math.max(0, matrix.rows - 1), 0], [plotArea.y + plotArea.height, plotArea.y]);
    const categoryScale = new CategoryScale(xLabels, [plotArea.x, plotArea.x + plotArea.width]);

    this.renderPlotBackground(svg, plotArea);
    const tracesLayer = this.createClippedLayer(svg, plotArea);
    const colors = resolvePalette("continuous", traceEntry.trace.colorscale, false);
    const cellWidth = plotArea.width / matrix.cols;
    const cellHeight = plotArea.height / matrix.rows;

    matrix.values.forEach((cell) => {
      const color = resolveColorFromScale(cell.value, {
        mode: "continuous",
        colors,
        min: matrix.min,
        max: matrix.max
      });
      const x = plotArea.x + cell.col * cellWidth;
      const y = plotArea.y + cell.row * cellHeight;
      const rect = createSvgElement("rect");
      rect.setAttribute("x", String(x));
      rect.setAttribute("y", String(y));
      rect.setAttribute("width", String(Math.max(1, cellWidth)));
      rect.setAttribute("height", String(Math.max(1, cellHeight)));
      rect.setAttribute("fill", color);
      rect.setAttribute("data-trace", "heatmap");
      this.attachInteractiveMetadata(
        rect,
        traceEntry.index,
        cell.row * matrix.cols + cell.col,
        traceEntry.trace.name,
        `${xLabels[cell.col]} / ${yLabels[cell.row]}`,
        formatNumeric(cell.value)
      );
      tracesLayer.append(rect);
    });

    svg.append(tracesLayer);

    const axisFigure: ChartFigure = {
      ...figure,
      layout: {
        ...figure.layout,
        yAxis: {
          ...figure.layout.yAxis,
          title: figure.layout.yAxis.title || "Rows"
        }
      }
    };
    this.renderCartesianAxes(svg, axisFigure, plotArea, xLabels, xScale, yScale, categoryScale);
  }

  private renderContourScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traceEntry: IndexedTrace<ContourTrace>,
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea
  ): void {
    const matrix = traceEntry.trace.z;
    if (matrix.length === 0 || matrix[0]?.length === 0) {
      return;
    }
    const plotArea = plotAreaOverride ?? layout.plotArea;
    const rows = matrix.length;
    const cols = matrix[0].length;
    const xLabels = Array.from({ length: cols }, (_, col) => String(traceEntry.trace.x?.[col] ?? `C${col + 1}`));
    const xScale = new LinearScale([0, Math.max(0, cols - 1)], [plotArea.x, plotArea.x + plotArea.width]);
    const yScale = new LinearScale([Math.max(0, rows - 1), 0], [plotArea.y + plotArea.height, plotArea.y]);
    const categoryScale = new CategoryScale(xLabels, [plotArea.x, plotArea.x + plotArea.width]);
    const allValues = matrix.flat().map((value) => Number(value)).filter((value) => Number.isFinite(value));
    if (allValues.length === 0) {
      return;
    }
    const min = Math.min(...allValues);
    const max = Math.max(...allValues);
    const levels = clampInt(traceEntry.trace.levels ?? 7, 3, 24);
    const palette = resolvePalette("continuous", traceEntry.trace.colorscale, false);

    this.renderPlotBackground(svg, plotArea);
    const tracesLayer = this.createClippedLayer(svg, plotArea);
    if (traceEntry.trace.fillContours) {
      const cellWidth = plotArea.width / Math.max(1, cols);
      const cellHeight = plotArea.height / Math.max(1, rows);
      for (let row = 0; row < rows; row += 1) {
        for (let col = 0; col < cols; col += 1) {
          const raw = Number(matrix[row][col]);
          if (!Number.isFinite(raw)) {
            continue;
          }
          const fill = resolveColorFromScale(raw, {
            mode: "continuous",
            colors: palette,
            min,
            max
          });
          const rect = createSvgElement("rect");
          rect.setAttribute("x", String(plotArea.x + col * cellWidth));
          rect.setAttribute("y", String(plotArea.y + row * cellHeight));
          rect.setAttribute("width", String(Math.max(1, cellWidth)));
          rect.setAttribute("height", String(Math.max(1, cellHeight)));
          rect.setAttribute("fill", fill);
          rect.setAttribute("opacity", "0.16");
          tracesLayer.append(rect);
        }
      }
    }

    for (let levelIndex = 0; levelIndex < levels; levelIndex += 1) {
      const ratio = levels <= 1 ? 0.5 : levelIndex / (levels - 1);
      const level = min + ratio * (max - min);
      const allSegments = buildContourSegments(matrix, level);
      const maxSegments = clampInt(traceEntry.trace.maxSegments ?? allSegments.length, 10, 100_000);
      const segments = allSegments.slice(0, maxSegments);
      const levelColor = resolveColorFromScale(level, {
        mode: "continuous",
        colors: palette,
        min,
        max
      });
      const lineWidth = traceEntry.trace.line?.width ?? 1.4;
      const lineOpacity = traceEntry.trace.line?.opacity ?? 0.94;
      segments.forEach((segment, segmentIndex) => {
        const startX = plotArea.x + segment.a.x * (plotArea.width / Math.max(1, cols - 1));
        const startY = plotArea.y + segment.a.y * (plotArea.height / Math.max(1, rows - 1));
        const endX = plotArea.x + segment.b.x * (plotArea.width / Math.max(1, cols - 1));
        const endY = plotArea.y + segment.b.y * (plotArea.height / Math.max(1, rows - 1));
        const path = createSvgElement("line");
        path.setAttribute("x1", String(startX));
        path.setAttribute("y1", String(startY));
        path.setAttribute("x2", String(endX));
        path.setAttribute("y2", String(endY));
        path.setAttribute("stroke", traceEntry.trace.line?.color ?? levelColor);
        path.setAttribute("stroke-width", String(lineWidth));
        path.setAttribute("opacity", String(lineOpacity));
        this.attachInteractiveMetadata(
          path,
          traceEntry.index,
          levelIndex * 10_000 + segmentIndex,
          traceEntry.trace.name,
          `level ${formatNumeric(level)}`,
          traceEntry.trace.line?.color ?? levelColor
        );
        tracesLayer.append(path);
      });
      if (traceEntry.trace.labelLevels && segments.length > 0) {
        const sample = segments[Math.floor(segments.length / 2)];
        const x = plotArea.x + ((sample.a.x + sample.b.x) * 0.5) * (plotArea.width / Math.max(1, cols - 1));
        const y = plotArea.y + ((sample.a.y + sample.b.y) * 0.5) * (plotArea.height / Math.max(1, rows - 1));
        const label = createSvgElement("text");
        label.textContent = formatNumeric(level);
        label.setAttribute("x", String(x));
        label.setAttribute("y", String(y));
        label.setAttribute("font-size", "10");
        label.setAttribute("font-family", "Arial, sans-serif");
        label.setAttribute("fill", traceEntry.trace.line?.color ?? levelColor);
        label.setAttribute("stroke", "none");
        tracesLayer.append(label);
      }
    }

    svg.append(tracesLayer);
    this.renderCartesianAxes(svg, figure, plotArea, xLabels, xScale, yScale, categoryScale);
  }

  private renderWaterfallScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traces: IndexedTrace<WaterfallTrace>[],
    layout: ComputedLayout,
    plotAreaOverride?: ComputedPlotArea
  ): void {
    const plotArea = plotAreaOverride ?? layout.plotArea;
    const allLabels = traces[0]?.trace.x.map((value) => String(value)) ?? [];
    const axisType = normalizeAxisType(figure.layout.xAxis, "category");
    const xDomainValues = traces
      .flatMap((entry) => entry.trace.x.map((value, index) => toAxisScalar(value, axisType, index)))
      .filter((value) => Number.isFinite(value));
    const xMin = xDomainValues.length > 0 ? Math.min(...xDomainValues) : 0;
    const xMax = xDomainValues.length > 0 ? Math.max(...xDomainValues) : Math.max(1, allLabels.length - 1);

    const runningValues: number[] = [];
    traces.forEach((entry) => {
      let running = 0;
      entry.trace.y.forEach((rawValue, index) => {
        const value = Number(rawValue);
        if (!Number.isFinite(value)) {
          return;
        }
        const measure = entry.trace.measure?.[index] ?? "relative";
        if (measure === "absolute") {
          running = value;
        } else if (measure === "total") {
          running = value;
        } else {
          running += value;
        }
        runningValues.push(running);
      });
    });
    const yMin = Math.min(0, ...(runningValues.length > 0 ? runningValues : [0]));
    const yMax = Math.max(0, ...(runningValues.length > 0 ? runningValues : [1]));
    const xScale = new LinearScale([xMin, xMax], [plotArea.x, plotArea.x + plotArea.width], {
      type: axisType === "log" ? "log" : "linear",
      reverse: figure.layout.xAxis.reverse === true
    });
    const yScale = new LinearScale([yMin, yMax === yMin ? yMin + 1 : yMax], [plotArea.y + plotArea.height, plotArea.y], {
      type: normalizeAxisType(figure.layout.yAxis, "linear") === "log" ? "log" : "linear",
      reverse: figure.layout.yAxis.reverse === true
    });
    const categoryScale = new CategoryScale(allLabels, [plotArea.x, plotArea.x + plotArea.width]);
    this.renderPlotBackground(svg, plotArea);
    this.renderCartesianGrid(svg, plotArea, xScale, yScale, {
      xZeroLine: figure.layout.xAxis.zeroLine,
      yZeroLine: figure.layout.yAxis.zeroLine
    });
    const tracesLayer = this.createClippedLayer(svg, plotArea);

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
        if (measure === "absolute") {
          running = value;
        } else if (measure === "total") {
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
        const rect = createSvgElement("rect");
        rect.setAttribute("x", String(centerX - slotWidth / 2));
        rect.setAttribute("y", String(top));
        rect.setAttribute("width", String(slotWidth));
        rect.setAttribute("height", String(height));
        const color =
          measure === "total"
            ? entry.trace.totals?.color ?? "#475569"
            : end >= start
              ? entry.trace.increasing?.color ?? "#16a34a"
              : entry.trace.decreasing?.color ?? "#dc2626";
        rect.setAttribute("fill", color);
        rect.setAttribute("opacity", "0.9");
        this.attachInteractiveMetadata(rect, entry.index, pointIndex, entry.trace.name, String(entry.trace.x[pointIndex] ?? pointIndex), formatNumeric(value));
        tracesLayer.append(rect);

        if (pointIndex > 0) {
          const previousX = toAxisScalar(entry.trace.x[pointIndex - 1], axisType, pointIndex - 1);
          const connector = createSvgElement("line");
          connector.setAttribute("x1", String(xScale.map(previousX) + offset + slotWidth / 2));
          connector.setAttribute("x2", String(centerX - slotWidth / 2));
          connector.setAttribute("y1", String(yScale.map(start)));
          connector.setAttribute("y2", String(yScale.map(start)));
          connector.setAttribute("stroke", "#64748b");
          connector.setAttribute("stroke-width", "1");
          connector.setAttribute("stroke-dasharray", "3 2");
          tracesLayer.append(connector);
        }
      });
    });

    svg.append(tracesLayer);
    this.renderCartesianAxes(svg, figure, plotArea, allLabels, xScale, yScale, categoryScale);
  }

  private renderFunnelScene(svg: SVGSVGElement, traceEntry: IndexedTrace<FunnelTrace>, plotArea: ComputedPlotArea): void {
    this.renderPlotBackground(svg, plotArea);
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
    const palette = traceEntry.trace.marker?.color ?? DEFAULT_SERIES_COLORS;
    const opacity = traceEntry.trace.marker?.opacity ?? 0.9;

    for (let index = 0; index < count; index += 1) {
      const value = values[index];
      const nextValue = values[index + 1] ?? value * 0.75;
      const topWidth = (value / maxValue) * plotArea.width * 0.94;
      const bottomWidth = (nextValue / maxValue) * plotArea.width * 0.94;
      const y0 = plotArea.y + index * stepHeight;
      const y1 = y0 + stepHeight;
      const cx = plotArea.x + plotArea.width / 2;
      const path = createSvgElement("path");
      path.setAttribute(
        "d",
        `M ${(cx - topWidth / 2).toFixed(2)} ${y0.toFixed(2)} L ${(cx + topWidth / 2).toFixed(2)} ${y0.toFixed(2)} L ${(cx + bottomWidth / 2).toFixed(2)} ${y1.toFixed(2)} L ${(cx - bottomWidth / 2).toFixed(2)} ${y1.toFixed(2)} Z`
      );
      path.setAttribute("fill", palette[index % palette.length]);
      path.setAttribute("opacity", String(opacity));
      path.setAttribute("stroke", "#ffffff");
      path.setAttribute("stroke-width", "1");
      const label = steps[index]?.label ?? `Step ${index + 1}`;
      const previousValue = index > 0 ? values[index - 1] : value;
      const conversionFromStart = (value / initialValue) * 100;
      const conversionFromPrevious = previousValue > 0 ? (value / previousValue) * 100 : 0;
      this.attachInteractiveMetadata(
        path,
        traceEntry.index,
        steps[index]?.index ?? index,
        traceEntry.trace.name,
        label,
        `${formatNumeric(value)} (${conversionFromStart.toFixed(1)}% total, ${conversionFromPrevious.toFixed(1)}% etapa anterior)`
      );
      svg.append(path);

      const text = createSvgElement("text");
      text.textContent = label;
      text.setAttribute("x", String(cx));
      text.setAttribute("y", String(y0 + stepHeight / 2 + 3));
      text.setAttribute("text-anchor", "middle");
      text.setAttribute("font-size", "11");
      text.setAttribute("font-family", "Arial, sans-serif");
      text.setAttribute("fill", "#0f172a");
      text.setAttribute("stroke", "none");
      svg.append(text);
    }
  }

  private renderPolarScene(
    svg: SVGSVGElement,
    traces: IndexedTrace<PolarTrace>[],
    plotArea: ComputedPlotArea,
    colorOffset: number
  ): void {
    this.renderPlotBackground(svg, plotArea);
    const centerX = plotArea.x + plotArea.width / 2;
    const centerY = plotArea.y + plotArea.height / 2;
    const radius = Math.max(10, Math.min(plotArea.width, plotArea.height) * 0.42);

    for (let grid = 1; grid <= 4; grid += 1) {
      const circle = createSvgElement("circle");
      circle.setAttribute("cx", String(centerX));
      circle.setAttribute("cy", String(centerY));
      circle.setAttribute("r", String((radius * grid) / 4));
      circle.setAttribute("fill", "none");
      circle.setAttribute("stroke", "#e2e8f0");
      circle.setAttribute("stroke-width", "1");
      svg.append(circle);
    }

    traces.forEach((entry, tracePosition) => {
      const trace = entry.trace;
      const color = trace.line?.color ?? trace.marker?.color ?? DEFAULT_SERIES_COLORS[(colorOffset + tracePosition) % DEFAULT_SERIES_COLORS.length];
      const maxR = Math.max(...trace.r.map((value) => Math.abs(Number(value))), 1);
      const points: Array<{ x: number; y: number; theta: string; r: number; index: number }> = [];

      trace.r.forEach((rawR, index) => {
        const rValue = Number(rawR);
        if (!Number.isFinite(rValue)) {
          return;
        }
        const thetaValue = trace.theta[index];
        const angle = toRadians(thetaValue, index, trace.theta.length);
        const normalized = clamp(rValue / maxR, -1, 1);
        const distance = normalized * radius;
        points.push({
          x: centerX + Math.cos(angle) * distance,
          y: centerY + Math.sin(angle) * distance,
          theta: String(thetaValue ?? index),
          r: rValue,
          index
        });
      });

      if (points.length === 0) {
        return;
      }

      const mode = trace.mode ?? "lines+markers";
      const variant = trace.variant ?? "scatter";
      if (variant === "bar") {
        const barWidth = Math.max(0.08, Math.min(Math.PI * 0.45, trace.barWidth ?? (Math.PI * 2) / Math.max(6, points.length)));
        points.forEach((point, pointIndex) => {
          const theta = toRadians(trace.theta[point.index], point.index, trace.theta.length);
          const start = theta - barWidth / 2;
          const end = theta + barWidth / 2;
          const path = createSvgElement("path");
          path.setAttribute("d", describeArcSlice(centerX, centerY, distanceBetween(centerX, centerY, point.x, point.y), start, end, 0));
          path.setAttribute("fill", resolveColor(color, 0.74));
          path.setAttribute("stroke", color);
          path.setAttribute("stroke-width", String(trace.line?.width ?? 1));
          path.setAttribute("opacity", String(trace.marker?.opacity ?? 0.9));
          this.attachInteractiveMetadata(path, entry.index, pointIndex, trace.name, point.theta, formatNumeric(point.r));
          svg.append(path);
        });
      }
      if ((variant === "line" || variant === "scatter" || variant === "area") && mode.includes("lines") && points.length > 1) {
        const path = createSvgElement("path");
        path.setAttribute(
          "d",
          points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x.toFixed(2)} ${point.y.toFixed(2)}`).join(" ")
        );
        if (variant === "area") {
          path.setAttribute("d", `${path.getAttribute("d")} L ${centerX.toFixed(2)} ${centerY.toFixed(2)} Z`);
          path.setAttribute("fill", resolveColor(color, 0.22));
        } else {
          path.setAttribute("fill", "none");
        }
        path.setAttribute("stroke", color);
        path.setAttribute("stroke-width", String(trace.line?.width ?? 2));
        path.setAttribute("opacity", String(trace.line?.opacity ?? 1));
        svg.append(path);
      }

      if ((variant === "scatter" || variant === "line") && mode.includes("markers")) {
        points.forEach((point) => {
          const marker = createSvgElement("circle");
          marker.setAttribute("cx", String(point.x));
          marker.setAttribute("cy", String(point.y));
          marker.setAttribute("r", String(trace.marker?.size ?? 4));
          marker.setAttribute("fill", color);
          marker.setAttribute("opacity", String(trace.marker?.opacity ?? 1));
          this.attachInteractiveMetadata(marker, entry.index, point.index, trace.name, point.theta, formatNumeric(point.r));
          svg.append(marker);
        });
      }
    });
  }

  private renderTernaryScene(
    svg: SVGSVGElement,
    traces: IndexedTrace<TernaryTrace>[],
    plotArea: ComputedPlotArea,
    colorOffset: number
  ): void {
    this.renderPlotBackground(svg, plotArea);
    const side = Math.min(plotArea.width, plotArea.height) * 0.82;
    const centerX = plotArea.x + plotArea.width / 2;
    const centerY = plotArea.y + plotArea.height / 2;
    const half = side / 2;
    const triangleHeight = (Math.sqrt(3) / 2) * side;
    const top = { x: centerX, y: centerY - triangleHeight / 2 };
    const left = { x: centerX - half, y: centerY + triangleHeight / 2 };
    const right = { x: centerX + half, y: centerY + triangleHeight / 2 };

    const border = createSvgElement("path");
    border.setAttribute("d", `M ${top.x} ${top.y} L ${right.x} ${right.y} L ${left.x} ${left.y} Z`);
    border.setAttribute("fill", "none");
    border.setAttribute("stroke", "#94a3b8");
    border.setAttribute("stroke-width", "1.5");
    svg.append(border);

    traces.forEach((entry, tracePosition) => {
      const color = entry.trace.marker?.color ?? DEFAULT_SERIES_COLORS[(colorOffset + tracePosition) % DEFAULT_SERIES_COLORS.length];
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
        const marker = createSvgElement("circle");
        marker.setAttribute("cx", String(x));
        marker.setAttribute("cy", String(y));
        marker.setAttribute("r", String(entry.trace.marker?.size ?? 4));
        marker.setAttribute("fill", color);
        marker.setAttribute("opacity", String(entry.trace.marker?.opacity ?? 0.9));
        this.attachInteractiveMetadata(marker, entry.index, index, entry.trace.name, `a=${formatNumeric(a)} b=${formatNumeric(b)}`, `c=${formatNumeric(c)}`);
        svg.append(marker);
      }
    });
  }

  private renderGeoScene(
    svg: SVGSVGElement,
    traces: IndexedTrace<GeoTrace>[],
    plotArea: ComputedPlotArea,
    colorOffset: number,
    geoScatterTraces: IndexedTrace<GeoScatterTrace>[] = [],
    geoLineTraces: IndexedTrace<GeoLineTrace>[] = []
  ): void {
    this.renderPlotBackground(svg, plotArea);
    const geoFeatures = traces.flatMap((entry) => entry.trace.geojson.features.map((feature) => ({ entry, feature })));
    const boundsFromGeo = computeGeoBounds(geoFeatures.map((item) => item.feature.geometry.coordinates));
    const scatterCoordinates = geoScatterTraces.flatMap((entry) =>
      entry.trace.lat.map((latitude, index) => [entry.trace.lon[index], latitude] as [number, number])
    );
    const linesCoordinates = geoLineTraces.flatMap((entry) => entry.trace.paths.flatMap((path) => path.map((point) => [point.lon, point.lat] as [number, number])));
    const boundsFromPoints = computeGeoBoundsFromPoints([...scatterCoordinates, ...linesCoordinates]);
    const bounds = mergeGeoBounds(boundsFromGeo, boundsFromPoints);
    if (!bounds) {
      return;
    }

    let minValue = Number.POSITIVE_INFINITY;
    let maxValue = Number.NEGATIVE_INFINITY;
    const valueResolvers = new Map<number, (feature: GeoTrace["geojson"]["features"][number]) => number | null>();
    traces.forEach((traceEntry) => {
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
      const featureValue = valueResolvers.get(item.entry.index)?.(item.feature);
      const palette = resolvePalette("continuous", item.entry.trace.colorscale, item.entry.trace.reverseScale === true);
      const numericFeatureValue = typeof featureValue === "number" && Number.isFinite(featureValue) ? featureValue : null;
      const color = resolveColorFromScale(numericFeatureValue, {
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
        const path = createSvgElement("path");
        path.setAttribute(
          "d",
          pathPoints
            .map((coordinate, pointIndex) => {
              const mapped = mapGeoToPlot(coordinate, bounds, plotArea);
              return `${pointIndex === 0 ? "M" : "L"} ${mapped.x.toFixed(2)} ${mapped.y.toFixed(2)}`;
            })
            .join(" ") + " Z"
        );
        path.setAttribute("fill", color);
        path.setAttribute("stroke", "#334155");
        path.setAttribute("stroke-width", "0.8");
        path.setAttribute("opacity", "0.78");
        this.attachInteractiveMetadata(
          path,
          item.entry.index,
          pathIndex,
          item.entry.trace.name,
          String(item.feature.properties?.[item.entry.trace.featureIdField ?? "id"] ?? `feature ${featureIndex + 1}`),
          numericFeatureValue !== null ? formatNumeric(numericFeatureValue) : item.entry.trace.missingColor ?? "n/a"
        );
        svg.append(path);
      });
    });

    const legendTrace = traces.find((entry) => entry.trace.showColorLegend);
    if (legendTrace) {
      this.renderGeoColorLegend(svg, legendTrace.trace, plotArea, minValue, maxValue);
    }

    geoLineTraces.forEach((traceEntry, tracePosition) => {
      const color = traceEntry.trace.line?.color ?? DEFAULT_SERIES_COLORS[(colorOffset + tracePosition) % DEFAULT_SERIES_COLORS.length];
      const width = traceEntry.trace.line?.width ?? 1.8;
      const opacity = traceEntry.trace.line?.opacity ?? 0.92;
      traceEntry.trace.paths.forEach((pathPoints, pointGroupIndex) => {
        const mapped = pathPoints
          .map((point) => mapGeoToPlot([point.lon, point.lat], bounds, plotArea))
          .filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
        if (mapped.length < 2) {
          return;
        }
        const path = createSvgElement("path");
        path.setAttribute("d", mapped.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x.toFixed(2)} ${point.y.toFixed(2)}`).join(" "));
        path.setAttribute("fill", "none");
        path.setAttribute("stroke", color);
        path.setAttribute("stroke-width", String(width));
        path.setAttribute("opacity", String(opacity));
        this.attachInteractiveMetadata(path, traceEntry.index, pointGroupIndex, traceEntry.trace.name, "geo-line", color);
        svg.append(path);
      });
    });

    geoScatterTraces.forEach((traceEntry, tracePosition) => {
      const markerColor = traceEntry.trace.marker?.color ?? DEFAULT_SERIES_COLORS[(colorOffset + tracePosition) % DEFAULT_SERIES_COLORS.length];
      const markerSize = traceEntry.trace.marker?.size ?? 4;
      const markerOpacity = traceEntry.trace.marker?.opacity ?? 0.94;
      const lineColor = traceEntry.trace.line?.color ?? markerColor;
      const lineWidth = traceEntry.trace.line?.width ?? 1.4;
      const lineOpacity = traceEntry.trace.line?.opacity ?? markerOpacity;
      const points = traceEntry.trace.lat
        .map((latitude, index) => ({
          lat: Number(latitude),
          lon: Number(traceEntry.trace.lon[index]),
          index
        }))
        .filter((point) => Number.isFinite(point.lat) && Number.isFinite(point.lon))
        .map((point) => ({
          ...point,
          ...mapGeoToPlot([point.lon, point.lat], bounds, plotArea)
        }));

      if ((traceEntry.trace.mode ?? "markers").includes("lines") && points.length >= 2) {
        const linePath = createSvgElement("path");
        linePath.setAttribute("d", points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x.toFixed(2)} ${point.y.toFixed(2)}`).join(" "));
        linePath.setAttribute("fill", "none");
        linePath.setAttribute("stroke", lineColor);
        linePath.setAttribute("stroke-width", String(lineWidth));
        linePath.setAttribute("opacity", String(lineOpacity));
        svg.append(linePath);
      }

      if ((traceEntry.trace.mode ?? "markers").includes("markers")) {
        points.forEach((point) => {
          const marker = createSvgElement("circle");
          marker.setAttribute("cx", String(point.x));
          marker.setAttribute("cy", String(point.y));
          marker.setAttribute("r", String(markerSize));
          marker.setAttribute("fill", markerColor);
          marker.setAttribute("opacity", String(markerOpacity));
          this.attachInteractiveMetadata(
            marker,
            traceEntry.index,
            point.index,
            traceEntry.trace.name,
            `${point.lat.toFixed(4)}, ${point.lon.toFixed(4)}`,
            markerColor
          );
          svg.append(marker);
        });
      }
    });
  }

  private renderGeoColorLegend(svg: SVGSVGElement, trace: GeoTrace, plotArea: ComputedPlotArea, minValue: number, maxValue: number): void {
    const legendWidth = Math.max(88, Math.min(140, plotArea.width * 0.22));
    const legendHeight = 10;
    const x = plotArea.x + plotArea.width - legendWidth - 12;
    const y = plotArea.y + plotArea.height - 28;
    const palette = resolvePalette("continuous", trace.colorscale, trace.reverseScale === true);
    const steps = 22;
    for (let index = 0; index < steps; index += 1) {
      const ratio = steps <= 1 ? 0 : index / (steps - 1);
      const value = minValue + ratio * (maxValue - minValue);
      const color = resolveColorFromScale(value, {
        mode: "continuous",
        colors: palette,
        min: minValue,
        max: maxValue
      });
      const segment = createSvgElement("rect");
      segment.setAttribute("x", String(x + (legendWidth / steps) * index));
      segment.setAttribute("y", String(y));
      segment.setAttribute("width", String(Math.ceil(legendWidth / steps)));
      segment.setAttribute("height", String(legendHeight));
      segment.setAttribute("fill", color);
      segment.setAttribute("stroke", "none");
      svg.append(segment);
    }
    const border = createSvgElement("rect");
    border.setAttribute("x", String(x));
    border.setAttribute("y", String(y));
    border.setAttribute("width", String(legendWidth));
    border.setAttribute("height", String(legendHeight));
    border.setAttribute("fill", "none");
    border.setAttribute("stroke", "#94a3b8");
    border.setAttribute("stroke-width", "0.7");
    svg.append(border);

    const minText = createSvgElement("text");
    minText.textContent = formatNumeric(minValue);
    minText.setAttribute("x", String(x));
    minText.setAttribute("y", String(y - 2));
    minText.setAttribute("font-size", "10");
    minText.setAttribute("font-family", "Arial, sans-serif");
    minText.setAttribute("fill", "#334155");
    minText.setAttribute("stroke", "none");
    svg.append(minText);

    const maxText = createSvgElement("text");
    maxText.textContent = formatNumeric(maxValue);
    maxText.setAttribute("x", String(x + legendWidth));
    maxText.setAttribute("y", String(y - 2));
    maxText.setAttribute("text-anchor", "end");
    maxText.setAttribute("font-size", "10");
    maxText.setAttribute("font-family", "Arial, sans-serif");
    maxText.setAttribute("fill", "#334155");
    maxText.setAttribute("stroke", "none");
    svg.append(maxText);
  }

  private renderProjected3dScene(
    svg: SVGSVGElement,
    scatterTraces: IndexedTrace<Scatter3dTrace>[],
    surfaceTraces: IndexedTrace<SurfaceTrace>[],
    meshTraces: IndexedTrace<Mesh3dTrace>[],
    plotArea: ComputedPlotArea
  ): void {
    this.renderPlotBackground(svg, plotArea);
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
        const line = createSvgElement("path");
        line.setAttribute("d", points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.p.x.toFixed(2)} ${point.p.y.toFixed(2)}`).join(" "));
        line.setAttribute("fill", "none");
        line.setAttribute("stroke", entry.trace.line?.color ?? color);
        line.setAttribute("stroke-width", String(entry.trace.line?.width ?? 1.6));
        line.setAttribute("opacity", String(entry.trace.line?.opacity ?? 0.9));
        svg.append(line);
      }

      if (mode.includes("markers")) {
        points.forEach((point) => {
          const marker = createSvgElement("circle");
          marker.setAttribute("cx", String(point.p.x));
          marker.setAttribute("cy", String(point.p.y));
          marker.setAttribute("r", String(entry.trace.marker?.size ?? 3.5));
          marker.setAttribute("fill", color);
          marker.setAttribute("opacity", String(entry.trace.marker?.opacity ?? 0.9));
          this.attachInteractiveMetadata(
            marker,
            entry.index,
            point.index,
            entry.trace.name,
            `x:${formatNumeric(point.x)} y:${formatNumeric(point.y)}`,
            `z:${formatNumeric(point.z)}`
          );
          svg.append(marker);
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
          const polygon = createSvgElement("path");
          polygon.setAttribute("d", `M ${p0.x.toFixed(2)} ${p0.y.toFixed(2)} L ${p1.x.toFixed(2)} ${p1.y.toFixed(2)} L ${p2.x.toFixed(2)} ${p2.y.toFixed(2)} L ${p3.x.toFixed(2)} ${p3.y.toFixed(2)} Z`);
          polygon.setAttribute(
            "fill",
            resolveColorFromScale((z00 + z10 + z11 + z01) / 4, {
              mode: "continuous",
              colors: palette,
              min,
              max
            })
          );
          polygon.setAttribute("stroke", "#0f172a");
          polygon.setAttribute("stroke-opacity", "0.15");
          polygon.setAttribute("stroke-width", "0.4");
          polygon.setAttribute("opacity", "0.92");
          this.attachInteractiveMetadata(polygon, entry.index, row * cols + col, entry.trace.name, `r${row + 1} c${col + 1}`, formatNumeric(z00));
          svg.append(polygon);
        }
      }
    });

    meshTraces.forEach((entry, traceIndex) => {
      const fill = entry.trace.marker?.color ?? DEFAULT_SERIES_COLORS[(traceIndex + 2) % DEFAULT_SERIES_COLORS.length];
      const opacity = entry.trace.marker?.opacity ?? 0.56;
      for (let index = 0; index < entry.trace.i.length; index += 1) {
        const ia = Number(entry.trace.i[index]);
        const jb = Number(entry.trace.j[index]);
        const kc = Number(entry.trace.k[index]);
        const a = projection(Number(entry.trace.x[ia]), Number(entry.trace.y[ia]), Number(entry.trace.z[ia]));
        const b = projection(Number(entry.trace.x[jb]), Number(entry.trace.y[jb]), Number(entry.trace.z[jb]));
        const c = projection(Number(entry.trace.x[kc]), Number(entry.trace.y[kc]), Number(entry.trace.z[kc]));
        const triangle = createSvgElement("path");
        triangle.setAttribute("d", `M ${a.x.toFixed(2)} ${a.y.toFixed(2)} L ${b.x.toFixed(2)} ${b.y.toFixed(2)} L ${c.x.toFixed(2)} ${c.y.toFixed(2)} Z`);
        triangle.setAttribute("fill", fill);
        triangle.setAttribute("opacity", String(opacity));
        triangle.setAttribute("stroke", "#0f172a");
        triangle.setAttribute("stroke-width", "0.4");
        svg.append(triangle);
      }
    });
  }

  private renderLayoutOverlays(svg: SVGSVGElement, figure: ChartFigure, layout: ComputedLayout): void {
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

    for (const shape of figure.layout.shapes) {
      if (shape.type === "line") {
        const line = createSvgElement("line");
        line.setAttribute("x1", String(mapX(shape.x0, shape.xRef)));
        line.setAttribute("y1", String(mapY(shape.y0, shape.yRef)));
        line.setAttribute("x2", String(mapX(shape.x1 ?? shape.x0, shape.xRef)));
        line.setAttribute("y2", String(mapY(shape.y1 ?? shape.y0, shape.yRef)));
        line.setAttribute("stroke", shape.stroke ?? "#2563eb");
        line.setAttribute("stroke-width", String(shape.strokeWidth ?? 1.5));
        line.setAttribute("opacity", String(shape.opacity ?? 1));
        svg.append(line);
        continue;
      }

      if (shape.type === "rect") {
        const x0 = mapX(shape.x0, shape.xRef);
        const y0 = mapY(shape.y0, shape.yRef);
        const x1 = mapX(shape.x1 ?? shape.x0, shape.xRef);
        const y1 = mapY(shape.y1 ?? shape.y0, shape.yRef);
        const rect = createSvgElement("rect");
        rect.setAttribute("x", String(Math.min(x0, x1)));
        rect.setAttribute("y", String(Math.min(y0, y1)));
        rect.setAttribute("width", String(Math.max(1, Math.abs(x1 - x0))));
        rect.setAttribute("height", String(Math.max(1, Math.abs(y1 - y0))));
        rect.setAttribute("stroke", shape.stroke ?? "#2563eb");
        rect.setAttribute("stroke-width", String(shape.strokeWidth ?? 1.2));
        rect.setAttribute("fill", shape.fill ?? "rgba(37, 99, 235, 0.12)");
        rect.setAttribute("opacity", String(shape.opacity ?? 1));
        svg.append(rect);
        continue;
      }

      if (shape.type === "circle") {
        const circle = createSvgElement("circle");
        circle.setAttribute("cx", String(mapX(shape.x0, shape.xRef)));
        circle.setAttribute("cy", String(mapY(shape.y0, shape.yRef)));
        circle.setAttribute("r", String(Math.max(1, shape.radius ?? 8)));
        circle.setAttribute("stroke", shape.stroke ?? "#2563eb");
        circle.setAttribute("stroke-width", String(shape.strokeWidth ?? 1.2));
        circle.setAttribute("fill", shape.fill ?? "rgba(37, 99, 235, 0.12)");
        circle.setAttribute("opacity", String(shape.opacity ?? 1));
        svg.append(circle);
        continue;
      }

      if (shape.type === "path" && shape.path) {
        const path = createSvgElement("path");
        path.setAttribute("d", shape.path);
        path.setAttribute("stroke", shape.stroke ?? "#2563eb");
        path.setAttribute("stroke-width", String(shape.strokeWidth ?? 1.2));
        path.setAttribute("fill", shape.fill ?? "none");
        path.setAttribute("opacity", String(shape.opacity ?? 1));
        svg.append(path);
        continue;
      }

      if (shape.type === "region" && Array.isArray(shape.points) && shape.points.length >= 3) {
        const region = createSvgElement("path");
        const d = shape.points
          .map((point, index) => `${index === 0 ? "M" : "L"} ${mapX(point.x, shape.xRef).toFixed(2)} ${mapY(point.y, shape.yRef).toFixed(2)}`)
          .join(" ");
        region.setAttribute("d", `${d} Z`);
        region.setAttribute("stroke", shape.stroke ?? "#2563eb");
        region.setAttribute("stroke-width", String(shape.strokeWidth ?? 1.2));
        region.setAttribute("fill", shape.fill ?? "rgba(37, 99, 235, 0.12)");
        region.setAttribute("opacity", String(shape.opacity ?? 1));
        svg.append(region);
      }
    }

    for (const annotation of figure.layout.annotations) {
      if (annotation.showArrow) {
        const arrowToX = mapX(annotation.arrowToX ?? annotation.x, annotation.xRef);
        const arrowToY = mapY(annotation.arrowToY ?? annotation.y, annotation.yRef);
        const arrow = createSvgElement("line");
        arrow.setAttribute("x1", String(arrowToX));
        arrow.setAttribute("y1", String(arrowToY));
        arrow.setAttribute("x2", String(mapX(annotation.x, annotation.xRef)));
        arrow.setAttribute("y2", String(mapY(annotation.y, annotation.yRef)));
        arrow.setAttribute("stroke", annotation.color ?? "#334155");
        arrow.setAttribute("stroke-width", "1.2");
        svg.append(arrow);
        const ax = mapX(annotation.x, annotation.xRef);
        const ay = mapY(annotation.y, annotation.yRef);
        const angle = Math.atan2(ay - arrowToY, ax - arrowToX);
        const leftX = ax - Math.cos(angle - 0.55) * 7;
        const leftY = ay - Math.sin(angle - 0.55) * 7;
        const rightX = ax - Math.cos(angle + 0.55) * 7;
        const rightY = ay - Math.sin(angle + 0.55) * 7;
        const head = createSvgElement("path");
        head.setAttribute("d", `M ${ax.toFixed(2)} ${ay.toFixed(2)} L ${leftX.toFixed(2)} ${leftY.toFixed(2)} L ${rightX.toFixed(2)} ${rightY.toFixed(2)} Z`);
        head.setAttribute("fill", annotation.color ?? "#334155");
        svg.append(head);
      }
      const text = createSvgElement("text");
      const x = mapX(annotation.x, annotation.xRef);
      const y = mapY(annotation.y, annotation.yRef);
      text.textContent = annotation.text;
      text.setAttribute("x", String(x));
      text.setAttribute("y", String(y));
      text.setAttribute("font-size", String(annotation.fontSize ?? 12));
      text.setAttribute("font-family", "Arial, sans-serif");
      text.setAttribute("fill", annotation.color ?? "#0f172a");
      text.setAttribute(
        "text-anchor",
        annotation.align === "left" ? "start" : annotation.align === "right" ? "end" : "middle"
      );
      if (typeof annotation.rotate === "number" && Number.isFinite(annotation.rotate)) {
        text.setAttribute("transform", `rotate(${annotation.rotate} ${x} ${y})`);
      }
      svg.append(text);
    }

    for (const imageLayer of figure.layout.images) {
      if (!imageLayer.source) {
        continue;
      }
      if (!isSafeImageSource(imageLayer.source)) {
        continue;
      }
      const x = mapX(imageLayer.x, imageLayer.xRef);
      const y = mapY(imageLayer.y, imageLayer.yRef);
      const image = createSvgElement("image");
      image.setAttribute("x", String(x));
      image.setAttribute("y", String(y));
      image.setAttribute("width", String(Math.max(1, imageLayer.width)));
      image.setAttribute("height", String(Math.max(1, imageLayer.height)));
      image.setAttribute("opacity", String(imageLayer.opacity ?? 1));
      image.setAttributeNS("http://www.w3.org/1999/xlink", "href", imageLayer.source);
      svg.append(image);
    }
  }

  private renderSubplotGroups<TTrace extends ChartTrace>(
    traces: IndexedTrace<TTrace>[],
    subplotAreas: ComputedPlotArea[],
    renderGroup: (group: IndexedTrace<TTrace>[], plotArea: ComputedPlotArea, subplotIndex: number) => void
  ): void {
    const groups = new Map<number, IndexedTrace<TTrace>[]>();
    traces.forEach((entry) => {
      const requestedSubplot = clampInt((entry.trace.subplot as number | undefined) ?? 0, 0, subplotAreas.length - 1);
      const bucket = groups.get(requestedSubplot);
      if (bucket) {
        bucket.push(entry);
      } else {
        groups.set(requestedSubplot, [entry]);
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

  private renderCartesianGrid(
    svg: SVGSVGElement,
    plotArea: ComputedPlotArea,
    xScale: LinearScale,
    yScale: LinearScale,
    options?: { xZeroLine?: boolean; yZeroLine?: boolean }
  ): void {
    const gridGroup = createSvgElement("g");
    gridGroup.setAttribute("stroke", "#e2e8f0");
    gridGroup.setAttribute("stroke-width", "1");

    for (const tick of yScale.ticks(6)) {
      const y = yScale.map(tick);
      const line = createSvgElement("line");
      line.setAttribute("x1", String(plotArea.x));
      line.setAttribute("x2", String(plotArea.x + plotArea.width));
      line.setAttribute("y1", String(y));
      line.setAttribute("y2", String(y));
      gridGroup.append(line);
    }

    for (const tick of xScale.ticks(8)) {
      const x = xScale.map(tick);
      const line = createSvgElement("line");
      line.setAttribute("x1", String(x));
      line.setAttribute("x2", String(x));
      line.setAttribute("y1", String(plotArea.y));
      line.setAttribute("y2", String(plotArea.y + plotArea.height));
      gridGroup.append(line);
    }

    const zeroY = yScale.map(0);
    if ((options?.yZeroLine ?? true) && Number.isFinite(zeroY) && zeroY >= plotArea.y && zeroY <= plotArea.y + plotArea.height) {
      const zeroLine = createSvgElement("line");
      zeroLine.setAttribute("x1", String(plotArea.x));
      zeroLine.setAttribute("x2", String(plotArea.x + plotArea.width));
      zeroLine.setAttribute("y1", String(zeroY));
      zeroLine.setAttribute("y2", String(zeroY));
      zeroLine.setAttribute("stroke", "#94a3b8");
      zeroLine.setAttribute("stroke-width", "1.5");
      gridGroup.append(zeroLine);
    }
    const zeroX = xScale.map(0);
    if ((options?.xZeroLine ?? true) && Number.isFinite(zeroX) && zeroX >= plotArea.x && zeroX <= plotArea.x + plotArea.width) {
      const zeroLine = createSvgElement("line");
      zeroLine.setAttribute("x1", String(zeroX));
      zeroLine.setAttribute("x2", String(zeroX));
      zeroLine.setAttribute("y1", String(plotArea.y));
      zeroLine.setAttribute("y2", String(plotArea.y + plotArea.height));
      zeroLine.setAttribute("stroke", "#94a3b8");
      zeroLine.setAttribute("stroke-width", "1.5");
      gridGroup.append(zeroLine);
    }

    svg.append(gridGroup);
  }

  private renderCartesianAxes(
    svg: SVGSVGElement,
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
    const axesGroup = createSvgElement("g");
    axesGroup.setAttribute("stroke", "#94a3b8");
    axesGroup.setAttribute("stroke-width", "1");

    const axisBottomY = plotArea.y + plotArea.height;
    const axisLeftX = plotArea.x;

    const xAxisLine = createSvgElement("line");
    xAxisLine.setAttribute("x1", String(plotArea.x));
    xAxisLine.setAttribute("x2", String(plotArea.x + plotArea.width));
    xAxisLine.setAttribute("y1", String(axisBottomY));
    xAxisLine.setAttribute("y2", String(axisBottomY));
    axesGroup.append(xAxisLine);

    const yAxisLine = createSvgElement("line");
    yAxisLine.setAttribute("x1", String(axisLeftX));
    yAxisLine.setAttribute("x2", String(axisLeftX));
    yAxisLine.setAttribute("y1", String(plotArea.y));
    yAxisLine.setAttribute("y2", String(plotArea.y + plotArea.height));
    axesGroup.append(yAxisLine);

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

    xTicks.forEach((entry) => {
      const tickLine = createSvgElement("line");
      tickLine.setAttribute("x1", String(entry.x));
      tickLine.setAttribute("x2", String(entry.x));
      tickLine.setAttribute("y1", String(axisBottomY));
      tickLine.setAttribute("y2", String(axisBottomY + 5));
      axesGroup.append(tickLine);
    });

    if (xAxisType === "multicategory") {
      const parsed = xTicks.map((entry) => ({ ...entry, ...splitMulticategoryLabel(entry.label) }));
      parsed.forEach((entry) => {
        const tickText = createSvgElement("text");
        tickText.textContent = entry.leaf;
        tickText.setAttribute("x", String(entry.x));
        tickText.setAttribute("y", String(axisBottomY + 16));
        tickText.setAttribute("text-anchor", "middle");
        tickText.setAttribute("font-size", "11");
        tickText.setAttribute("font-family", "Arial, sans-serif");
        tickText.setAttribute("fill", "#334155");
        tickText.setAttribute("stroke", "none");
        axesGroup.append(tickText);
      });
      let groupStart = 0;
      while (groupStart < parsed.length) {
        const groupName = parsed[groupStart].group;
        let groupEnd = groupStart;
        while (groupEnd + 1 < parsed.length && parsed[groupEnd + 1].group === groupName) {
          groupEnd += 1;
        }
        if (groupName) {
          const groupText = createSvgElement("text");
          groupText.textContent = groupName;
          groupText.setAttribute("x", String((parsed[groupStart].x + parsed[groupEnd].x) / 2));
          groupText.setAttribute("y", String(axisBottomY + 30));
          groupText.setAttribute("text-anchor", "middle");
          groupText.setAttribute("font-size", "10");
          groupText.setAttribute("font-family", "Arial, sans-serif");
          groupText.setAttribute("fill", "#64748b");
          groupText.setAttribute("stroke", "none");
          axesGroup.append(groupText);
        }
        groupStart = groupEnd + 1;
      }
    } else {
      xTicks.forEach((entry) => {
        const tickText = createSvgElement("text");
        tickText.textContent = entry.label;
        tickText.setAttribute("x", String(entry.x));
        tickText.setAttribute("y", String(axisBottomY + 16));
        tickText.setAttribute("text-anchor", "middle");
        tickText.setAttribute("font-size", "11");
        tickText.setAttribute("font-family", "Arial, sans-serif");
        tickText.setAttribute("fill", "#334155");
        tickText.setAttribute("stroke", "none");
        axesGroup.append(tickText);
      });
    }

    const yTickCount = getAxisTickCount(figure.layout.yAxis, Math.max(2, Math.min(10, Math.round(plotArea.height / 56))), 2, 20);
    for (const tick of yScale.ticks(yTickCount)) {
      const y = yScale.map(tick);
      const tickLine = createSvgElement("line");
      tickLine.setAttribute("x1", String(axisLeftX - 5));
      tickLine.setAttribute("x2", String(axisLeftX));
      tickLine.setAttribute("y1", String(y));
      tickLine.setAttribute("y2", String(y));
      axesGroup.append(tickLine);

      const tickText = createSvgElement("text");
      tickText.textContent = formatAxisTick(tick, figure.layout.yAxis);
      tickText.setAttribute("x", String(axisLeftX - 8));
      tickText.setAttribute("y", String(y + 3));
      tickText.setAttribute("text-anchor", "end");
      tickText.setAttribute("font-size", "11");
      tickText.setAttribute("font-family", "Arial, sans-serif");
      tickText.setAttribute("fill", "#334155");
      tickText.setAttribute("stroke", "none");
      axesGroup.append(tickText);
    }

    if (figure.layout.xAxis.title) {
      const xTitle = createSvgElement("text");
      xTitle.textContent = figure.layout.xAxis.title;
      xTitle.setAttribute("x", String(plotArea.x + plotArea.width / 2));
      xTitle.setAttribute("y", String(axisBottomY + (xAxisType === "multicategory" ? 50 : 34)));
      xTitle.setAttribute("text-anchor", "middle");
      xTitle.setAttribute("font-size", "12");
      xTitle.setAttribute("font-family", "Arial, sans-serif");
      xTitle.setAttribute("fill", "#0f172a");
      xTitle.setAttribute("stroke", "none");
      axesGroup.append(xTitle);
    }

    if (figure.layout.yAxis.title) {
      const yTitle = createSvgElement("text");
      yTitle.textContent = figure.layout.yAxis.title;
      yTitle.setAttribute("x", String(plotArea.x - 42));
      yTitle.setAttribute("y", String(plotArea.y + plotArea.height / 2));
      yTitle.setAttribute("text-anchor", "middle");
      yTitle.setAttribute("font-size", "12");
      yTitle.setAttribute("font-family", "Arial, sans-serif");
      yTitle.setAttribute("fill", "#0f172a");
      yTitle.setAttribute("stroke", "none");
      yTitle.setAttribute("transform", `rotate(-90 ${plotArea.x - 42} ${plotArea.y + plotArea.height / 2})`);
      axesGroup.append(yTitle);
    }

    if (yScaleSecondary) {
      const axisRightX = plotArea.x + plotArea.width;
      const secondaryAxisLine = createSvgElement("line");
      secondaryAxisLine.setAttribute("x1", String(axisRightX));
      secondaryAxisLine.setAttribute("x2", String(axisRightX));
      secondaryAxisLine.setAttribute("y1", String(plotArea.y));
      secondaryAxisLine.setAttribute("y2", String(plotArea.y + plotArea.height));
      axesGroup.append(secondaryAxisLine);

      const secondaryTickCount = getAxisTickCount(figure.layout.yAxis2, Math.max(2, Math.min(10, Math.round(plotArea.height / 56))), 2, 20);
      for (const tick of yScaleSecondary.ticks(secondaryTickCount)) {
        const y = yScaleSecondary.map(tick);
        const tickLine = createSvgElement("line");
        tickLine.setAttribute("x1", String(axisRightX));
        tickLine.setAttribute("x2", String(axisRightX + 5));
        tickLine.setAttribute("y1", String(y));
        tickLine.setAttribute("y2", String(y));
        axesGroup.append(tickLine);

        const tickText = createSvgElement("text");
        tickText.textContent = formatAxisTick(tick, figure.layout.yAxis2);
        tickText.setAttribute("x", String(axisRightX + 8));
        tickText.setAttribute("y", String(y + 3));
        tickText.setAttribute("text-anchor", "start");
        tickText.setAttribute("font-size", "11");
        tickText.setAttribute("font-family", "Arial, sans-serif");
        tickText.setAttribute("fill", "#334155");
        tickText.setAttribute("stroke", "none");
        axesGroup.append(tickText);
      }

      if (figure.layout.yAxis2.title) {
        const y2Title = createSvgElement("text");
        y2Title.textContent = figure.layout.yAxis2.title;
        y2Title.setAttribute("x", String(plotArea.x + plotArea.width + 42));
        y2Title.setAttribute("y", String(plotArea.y + plotArea.height / 2));
        y2Title.setAttribute("text-anchor", "middle");
        y2Title.setAttribute("font-size", "12");
        y2Title.setAttribute("font-family", "Arial, sans-serif");
        y2Title.setAttribute("fill", "#0f172a");
        y2Title.setAttribute("stroke", "none");
        y2Title.setAttribute("transform", `rotate(90 ${plotArea.x + plotArea.width + 42} ${plotArea.y + plotArea.height / 2})`);
        axesGroup.append(y2Title);
      }
    }

    if (xScaleSecondary) {
      const topLine = createSvgElement("line");
      topLine.setAttribute("x1", String(plotArea.x));
      topLine.setAttribute("x2", String(plotArea.x + plotArea.width));
      topLine.setAttribute("y1", String(plotArea.y));
      topLine.setAttribute("y2", String(plotArea.y));
      axesGroup.append(topLine);
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
        const tickLine = createSvgElement("line");
        tickLine.setAttribute("x1", String(entry.x));
        tickLine.setAttribute("x2", String(entry.x));
        tickLine.setAttribute("y1", String(plotArea.y));
        tickLine.setAttribute("y2", String(plotArea.y - 5));
        axesGroup.append(tickLine);
      });
      if (xAxis2Type === "multicategory") {
        const parsed = topTicks.map((entry) => ({ ...entry, ...splitMulticategoryLabel(entry.label) }));
        parsed.forEach((entry) => {
          const tickText = createSvgElement("text");
          tickText.textContent = entry.leaf;
          tickText.setAttribute("x", String(entry.x));
          tickText.setAttribute("y", String(plotArea.y - 8));
          tickText.setAttribute("text-anchor", "middle");
          tickText.setAttribute("font-size", "10");
          tickText.setAttribute("font-family", "Arial, sans-serif");
          tickText.setAttribute("fill", "#334155");
          tickText.setAttribute("stroke", "none");
          axesGroup.append(tickText);
        });
        let groupStart = 0;
        while (groupStart < parsed.length) {
          const groupName = parsed[groupStart].group;
          let groupEnd = groupStart;
          while (groupEnd + 1 < parsed.length && parsed[groupEnd + 1].group === groupName) {
            groupEnd += 1;
          }
          if (groupName) {
            const groupText = createSvgElement("text");
            groupText.textContent = groupName;
            groupText.setAttribute("x", String((parsed[groupStart].x + parsed[groupEnd].x) / 2));
            groupText.setAttribute("y", String(plotArea.y - 22));
            groupText.setAttribute("text-anchor", "middle");
            groupText.setAttribute("font-size", "10");
            groupText.setAttribute("font-family", "Arial, sans-serif");
            groupText.setAttribute("fill", "#64748b");
            groupText.setAttribute("stroke", "none");
            axesGroup.append(groupText);
          }
          groupStart = groupEnd + 1;
        }
      } else {
        topTicks.forEach((entry) => {
          const tickText = createSvgElement("text");
          tickText.textContent = entry.label;
          tickText.setAttribute("x", String(entry.x));
          tickText.setAttribute("y", String(plotArea.y - 8));
          tickText.setAttribute("text-anchor", "middle");
          tickText.setAttribute("font-size", "10");
          tickText.setAttribute("font-family", "Arial, sans-serif");
          tickText.setAttribute("fill", "#334155");
          tickText.setAttribute("stroke", "none");
          axesGroup.append(tickText);
        });
      }
      if (figure.layout.xAxis2.title) {
        const title = createSvgElement("text");
        title.textContent = figure.layout.xAxis2.title;
        title.setAttribute("x", String(plotArea.x + plotArea.width / 2));
        title.setAttribute("y", String(plotArea.y - (xAxis2Type === "multicategory" ? 36 : 20)));
        title.setAttribute("text-anchor", "middle");
        title.setAttribute("font-size", "12");
        title.setAttribute("font-family", "Arial, sans-serif");
        title.setAttribute("fill", "#0f172a");
        title.setAttribute("stroke", "none");
        axesGroup.append(title);
      }
    }

    svg.append(axesGroup);
  }

  private renderLineLikeTrace(
    group: SVGGElement,
    traceEntry: IndexedTrace<CartesianTrace>,
    xScale: LinearScale,
    yScale: LinearScale,
    color: string,
    plotArea: ComputedPlotArea,
    xAxisType: "linear" | "log" | "date" | "category" | "multicategory",
    stackSeries?: { stacked: boolean; points: Array<{ base: number; top: number; value: number }> },
    previousAreaPoints?: CartesianPoint[]
  ): CartesianPoint[] {
    const trace = traceEntry.trace;
    const mode = this.resolveTraceMode(trace);
    const yOverride = stackSeries?.stacked ? stackSeries.points.map((point) => point.top) : undefined;
    const points = this.buildCartesianPoints(trace, xScale, yScale, xAxisType, yOverride);

    if (trace.type === "area" && points.length > 1) {
      const baselinePixels = stackSeries?.stacked
        ? stackSeries.points.map((point) => yScale.map(point.base))
        : previousAreaPoints && trace.fill === "tonext" && previousAreaPoints.length === points.length
          ? previousAreaPoints.map((point) => point.yPixel)
          : points.map(() => yScale.map(0));
      const areaPath = createSvgElement("path");
      let d = points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.xPixel.toFixed(2)} ${point.yPixel.toFixed(2)}`).join(" ");
      for (let index = points.length - 1; index >= 0; index -= 1) {
        d += ` L ${points[index].xPixel.toFixed(2)} ${baselinePixels[index].toFixed(2)}`;
      }
      d += " Z";
      areaPath.setAttribute("d", d);
      areaPath.setAttribute("fill", resolveColor(color, 0.24));
      areaPath.setAttribute("stroke", "none");
      group.append(areaPath);
    }

    if (mode.includes("lines") && points.length > 1) {
      const optimized = simplifyLine(
        clipLineToRect(
          points.map((point) => ({ x: point.xPixel, y: point.yPixel })),
          plotArea
        ),
        0.9
      );
      if (optimized.length >= 2) {
        const line = createSvgElement("path");
        line.setAttribute(
          "d",
          optimized.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x.toFixed(2)} ${point.y.toFixed(2)}`).join(" ")
        );
        line.setAttribute("fill", "none");
        line.setAttribute("stroke", trace.line?.color ?? color);
        line.setAttribute("stroke-width", String(trace.line?.width ?? 2));
        line.setAttribute("opacity", String(trace.line?.opacity ?? 1));
        group.append(line);
      }
    }

    const renderMarkers = mode.includes("markers");
    for (const point of points) {
      const marker = createSvgElement("circle");
      marker.setAttribute("cx", String(point.xPixel));
      marker.setAttribute("cy", String(point.yPixel));
      marker.setAttribute("r", String(resolveMarkerSize(trace.marker?.size, point.pointIndex, renderMarkers ? 4 : 6)));
      marker.setAttribute("fill", renderMarkers ? resolveMarkerColor(trace.marker?.color, point.pointIndex, color) : "transparent");
      marker.setAttribute("opacity", String(renderMarkers ? trace.marker?.opacity ?? 1 : 0));
      this.attachInteractiveMetadata(marker, traceEntry.index, point.pointIndex, trace.name, point.xLabel, point.yLabel);
      group.append(marker);
    }
    return points;
  }

  private renderBarTrace(
    group: SVGGElement,
    traceEntry: IndexedTrace<CartesianTrace>,
    plotArea: ComputedPlotArea,
    xScale: LinearScale,
    yScale: LinearScale,
    color: string,
    barIndex: number,
    barCount: number,
    xAxisType: "linear" | "log" | "date" | "category" | "multicategory",
    stackSeries?: { stacked: boolean; points: Array<{ base: number; top: number; value: number }> }
  ): void {
    const trace = traceEntry.trace;
    const orientation = trace.orientation ?? "vertical";
    const domainSpan = Math.max(1, Math.abs(xScale.invert(plotArea.x + plotArea.width) - xScale.invert(plotArea.x)));
    const slotWidth = plotArea.width / (domainSpan + 1);
    const barSlotWidth = Math.max(4, slotWidth * 0.72);
    const barWidth = Math.max(2, barSlotWidth / barCount);
    const errorValues = trace.errorY?.values;
    const isStacked = stackSeries?.stacked === true && !!trace.stackGroup;

    trace.y.forEach((value, pointIndex) => {
      if (!Number.isFinite(value)) {
        return;
      }
      const xValue = toAxisScalar(trace.x[pointIndex], xAxisType, pointIndex);
      const centerX = xScale.map(xValue);
      const activeBarWidth = isStacked ? barSlotWidth : barWidth;
      const x = centerX - barSlotWidth / 2 + (isStacked ? 0 : activeBarWidth * barIndex);
      const stackedPoint = stackSeries?.points[pointIndex];
      const topValue = stackedPoint ? stackedPoint.top : value;
      const baseValue = stackedPoint ? stackedPoint.base : 0;
      const y = yScale.map(topValue);
      const baseY = yScale.map(baseValue);
      const top = Math.min(y, baseY);
      const height = Math.max(1, Math.abs(baseY - y));

      const rect = createSvgElement("rect");
      if (orientation === "horizontal") {
        const baseX = xScale.map(0);
        const width = Math.max(1, Math.abs(centerX - baseX));
        rect.setAttribute("x", String(Math.min(centerX, baseX)));
        rect.setAttribute("y", String(y - activeBarWidth / 2));
        rect.setAttribute("width", String(width));
        rect.setAttribute("height", String(activeBarWidth));
      } else {
        rect.setAttribute("x", String(x));
        rect.setAttribute("y", String(top));
        rect.setAttribute("width", String(activeBarWidth));
        rect.setAttribute("height", String(height));
      }
      rect.setAttribute("fill", resolveMarkerColor(trace.marker?.color, pointIndex, color));
      rect.setAttribute("opacity", String(trace.marker?.opacity ?? 0.9));
      this.attachInteractiveMetadata(rect, traceEntry.index, pointIndex, trace.name, String(trace.x[pointIndex] ?? pointIndex), formatNumeric(value));
      group.append(rect);

      const error = Number(errorValues?.[pointIndex]);
      if (Number.isFinite(error) && error > 0 && orientation !== "horizontal") {
        const yTop = yScale.map(topValue + error);
        const yBottom = yScale.map(topValue - error);
        const errorLine = createSvgElement("line");
        errorLine.setAttribute("x1", String(centerX));
        errorLine.setAttribute("x2", String(centerX));
        errorLine.setAttribute("y1", String(yTop));
        errorLine.setAttribute("y2", String(yBottom));
        errorLine.setAttribute("stroke", trace.line?.color ?? "#334155");
        errorLine.setAttribute("stroke-width", String(trace.line?.width ?? 1.2));
        errorLine.setAttribute("opacity", String(trace.line?.opacity ?? 0.9));
        group.append(errorLine);
      }
    });
  }

  private buildCartesianPoints(
    trace: CartesianTrace,
    xScale: LinearScale,
    yScale: LinearScale,
    xAxisType: "linear" | "log" | "date" | "category" | "multicategory",
    yOverride?: number[]
  ): CartesianPoint[] {
    const points: CartesianPoint[] = [];
    const sourceY = yOverride ?? trace.y;
    const totalPoints = sourceY.length;
    const sampleStep = totalPoints > this.maxRenderPoints ? Math.ceil(totalPoints / this.maxRenderPoints) : 1;
    for (let index = 0; index < totalPoints; index += sampleStep) {
      const value = sourceY[index];
      if (!Number.isFinite(value)) {
        continue;
      }
      const xValue = toAxisScalar(trace.x[index], xAxisType, index);
      points.push({
        pointIndex: index,
        xPixel: xScale.map(xValue),
        yPixel: yScale.map(value),
        xLabel: String(trace.x[index] ?? index),
        yLabel: formatNumeric(value)
      });
    }
    if (sampleStep > 1 && totalPoints > 0) {
      const lastIndex = totalPoints - 1;
      const lastValue = sourceY[lastIndex];
      if (Number.isFinite(lastValue) && points.at(-1)?.pointIndex !== lastIndex) {
        const xValue = toAxisScalar(trace.x[lastIndex], xAxisType, lastIndex);
        points.push({
          pointIndex: lastIndex,
          xPixel: xScale.map(xValue),
          yPixel: yScale.map(lastValue),
          xLabel: String(trace.x[lastIndex] ?? lastIndex),
          yLabel: formatNumeric(lastValue)
        });
      }
    }
    return points;
  }

  private renderPieTrace(svg: SVGSVGElement, traceEntry: IndexedTrace<PieTrace>, plotArea: ComputedPlotArea, colorOffset: number): void {
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

    const sum = slices.reduce((accumulator, slice) => accumulator + slice.value, 0);
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
    let startAngle = -Math.PI / 2;

    slices.forEach((slice, index) => {
      const arcAngle = (slice.value / sum) * Math.PI * 2;
      const endAngle = startAngle + arcAngle;
      const midAngle = startAngle + arcAngle / 2;
      const pull = Math.max(0, Number(pullValues ? pullValues[index] ?? 0 : defaultPull) || 0);
      const pullOffset = radius * Math.min(1, pull) * 0.34;
      const sliceCx = cx + Math.cos(midAngle) * pullOffset;
      const sliceCy = cy + Math.sin(midAngle) * pullOffset;
      const path = createSvgElement("path");
      path.setAttribute("d", describeArcSlice(sliceCx, sliceCy, radius, startAngle, endAngle, radius * hole));
      path.setAttribute("fill", trace.marker?.colors?.[index] ?? DEFAULT_SERIES_COLORS[(index + colorOffset) % DEFAULT_SERIES_COLORS.length]);
      path.setAttribute("opacity", String(trace.marker?.opacity ?? 1));
      path.setAttribute("data-trace", "pie");
      this.attachInteractiveMetadata(path, traceEntry.index, slice.index, trace.name, slice.label, formatNumeric(slice.value));
      svg.append(path);

      const percent = (slice.value / Math.max(Number.EPSILON, sum)) * 100;
      const anchorRadius = radius + pullOffset;
      const anchorX = sliceCx + Math.cos(midAngle) * anchorRadius;
      const anchorY = sliceCy + Math.sin(midAngle) * anchorRadius;
      const textX = sliceCx + Math.cos(midAngle) * (anchorRadius + 22);
      labelCandidates.push({
        side: textX >= sliceCx ? "right" : "left",
        x: textX,
        y: anchorY,
        anchorX,
        anchorY,
        text: `${slice.label} (${percent.toFixed(1)}%)`
      });
      startAngle = endAngle;
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
      labelCandidates.forEach((candidate) => {
        const elbowX = candidate.side === "right" ? candidate.x - 8 : candidate.x + 8;
        const line = createSvgElement("polyline");
        line.setAttribute(
          "points",
          `${candidate.anchorX.toFixed(2)},${candidate.anchorY.toFixed(2)} ${elbowX.toFixed(2)},${candidate.y.toFixed(2)} ${candidate.x.toFixed(2)},${candidate.y.toFixed(2)}`
        );
        line.setAttribute("fill", "none");
        line.setAttribute("stroke", "#94a3b8");
        line.setAttribute("stroke-width", "1");
        svg.append(line);
        const label = createSvgElement("text");
        label.textContent = candidate.text;
        label.setAttribute("x", String(candidate.x));
        label.setAttribute("y", String(candidate.y + 3));
        label.setAttribute("text-anchor", candidate.side === "right" ? "start" : "end");
        label.setAttribute("font-size", "10");
        label.setAttribute("font-family", "Arial, sans-serif");
        label.setAttribute("fill", "#334155");
        label.setAttribute("stroke", "none");
        svg.append(label);
      });
    }
  }

  private renderDonutTrace(svg: SVGSVGElement, traceEntry: IndexedTrace<DonutTrace>, plotArea: ComputedPlotArea, colorOffset: number): void {
    const asPie: IndexedTrace<PieTrace> = {
      ...traceEntry,
      trace: {
        ...traceEntry.trace,
        type: "pie",
        hole: traceEntry.trace.hole ?? 0.48
      }
    };
    this.renderPieTrace(svg, asPie, plotArea, colorOffset);
  }

  private renderSunburstTrace(svg: SVGSVGElement, traceEntry: IndexedTrace<SunburstTrace>, plotArea: ComputedPlotArea, colorOffset: number): void {
    this.renderPlotBackground(svg, plotArea);
    const trace = traceEntry.trace;
    const ids = trace.labels.map((label, index) => String(trace.ids?.[index] ?? label ?? index));
    const labelsById = new Map(ids.map((id, index) => [id, trace.labels[index] ?? id]));
    const parentById = new Map(ids.map((id, index) => [id, String(trace.parents[index] ?? "")]));
    const rootId = trace.rootId !== undefined ? String(trace.rootId) : "";
    const drillDepth = clampInt(trace.drilldownDepth ?? 99, 1, 99);
    const visibleRoots = trace.parents
      .map((parent, index) => ({ parent: String(parent ?? ""), index }))
      .filter((entry) => (rootId ? entry.parent === rootId : !entry.parent));
    const radius = Math.max(14, Math.min(plotArea.width, plotArea.height) * 0.42);
    const cx = plotArea.x + plotArea.width / 2;
    const cy = plotArea.y + plotArea.height / 2;
    const palette = trace.marker?.colors ?? DEFAULT_SERIES_COLORS;

    const total = visibleRoots.reduce((sum, entry) => sum + Math.max(0, Number(trace.values[entry.index])), 0) || 1;
    let start = -Math.PI / 2;
    visibleRoots.forEach((entry, rootPosition) => {
      const value = Math.max(0, Number(trace.values[entry.index]));
      const angle = (value / total) * Math.PI * 2;
      const end = start + angle;
      const nodeId = ids[entry.index];
      const pathParts: string[] = [];
      let cursor = nodeId;
      let depth = 0;
      while (cursor && depth < 24) {
        pathParts.unshift(labelsById.get(cursor) ?? cursor);
        cursor = parentById.get(cursor) ?? "";
        depth += 1;
      }
      if (pathParts.length > drillDepth + 1) {
        start = end;
        return;
      }
      const slice = createSvgElement("path");
      slice.setAttribute("d", describeArcSlice(cx, cy, radius, start, end, radius * 0.55));
      slice.setAttribute("fill", palette[(colorOffset + rootPosition) % palette.length]);
      slice.setAttribute("opacity", String(trace.marker?.opacity ?? 0.95));
      this.attachInteractiveMetadata(slice, traceEntry.index, entry.index, trace.name, pathParts.join(" > "), formatNumeric(value));
      svg.append(slice);
      start = end;
    });

    if (rootId) {
      const backRadius = radius * 0.48;
      const back = createSvgElement("circle");
      back.setAttribute("cx", String(cx));
      back.setAttribute("cy", String(cy));
      back.setAttribute("r", String(backRadius));
      back.setAttribute("fill", "#ffffff");
      back.setAttribute("stroke", "#94a3b8");
      back.setAttribute("stroke-width", "1");
      this.attachInteractiveMetadata(back, traceEntry.index, -1, trace.name, "voltar", rootId);
      svg.append(back);
      const text = createSvgElement("text");
      text.textContent = "Voltar";
      text.setAttribute("x", String(cx));
      text.setAttribute("y", String(cy + 4));
      text.setAttribute("text-anchor", "middle");
      text.setAttribute("font-size", "11");
      text.setAttribute("font-family", "Arial, sans-serif");
      text.setAttribute("fill", "#334155");
      text.setAttribute("stroke", "none");
      svg.append(text);
    }
  }

  private renderTreemapTrace(svg: SVGSVGElement, traceEntry: IndexedTrace<TreemapTrace>, plotArea: ComputedPlotArea, colorOffset: number): void {
    this.renderPlotBackground(svg, plotArea);
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
      const backRect = createSvgElement("rect");
      backRect.setAttribute("x", String(plotArea.x));
      backRect.setAttribute("y", String(plotArea.y));
      backRect.setAttribute("width", String(plotArea.width));
      backRect.setAttribute("height", String(headerHeight));
      backRect.setAttribute("fill", "#f8fafc");
      backRect.setAttribute("stroke", "#cbd5e1");
      backRect.setAttribute("stroke-width", "1");
      this.attachInteractiveMetadata(backRect, traceEntry.index, -1, trace.name, "voltar", rootId);
      svg.append(backRect);
      const backText = createSvgElement("text");
      backText.textContent = `Voltar (${rootId})`;
      backText.setAttribute("x", String(plotArea.x + 8));
      backText.setAttribute("y", String(plotArea.y + headerHeight / 2 + 4));
      backText.setAttribute("font-size", "11");
      backText.setAttribute("font-family", "Arial, sans-serif");
      backText.setAttribute("fill", "#334155");
      backText.setAttribute("stroke", "none");
      svg.append(backText);
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
    const palette = trace.marker?.colors ?? DEFAULT_SERIES_COLORS;
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
      const rect = createSvgElement("rect");
      rect.setAttribute("x", String(cursorX));
      rect.setAttribute("y", String(contentArea.y));
      rect.setAttribute("width", String(width));
      rect.setAttribute("height", String(contentArea.height));
      rect.setAttribute("fill", palette[(colorOffset + index) % palette.length]);
      rect.setAttribute("stroke", "#ffffff");
      rect.setAttribute("stroke-width", "1");
      rect.setAttribute("opacity", String(trace.marker?.opacity ?? 0.9));
      this.attachInteractiveMetadata(rect, traceEntry.index, index, trace.name, pathLabels.join(" > "), formatNumeric(value));
      svg.append(rect);

      const label = createSvgElement("text");
      label.textContent = trace.labels[index];
      label.setAttribute("x", String(cursorX + 4));
      label.setAttribute("y", String(contentArea.y + 16));
      label.setAttribute("font-size", "11");
      label.setAttribute("font-family", "Arial, sans-serif");
      label.setAttribute("fill", "#0f172a");
      label.setAttribute("stroke", "none");
      svg.append(label);
      cursorX += width;
    });
  }

  private renderSankeyScene(svg: SVGSVGElement, traceEntry: IndexedTrace<SankeyTrace>, plotArea: ComputedPlotArea): void {
    this.renderPlotBackground(svg, plotArea);
    const trace = traceEntry.trace;
    const nodeCount = trace.nodes.ids.length;
    const nodeLabels = trace.nodes.labels ?? trace.nodes.ids.map((value) => String(value));
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
      const controlX1 = startX + (endX - startX) * 0.45;
      const controlX2 = startX + (endX - startX) * 0.55;
      const path = createSvgElement("path");
      path.setAttribute("d", `M ${startX.toFixed(2)} ${startY.toFixed(2)} C ${controlX1.toFixed(2)} ${startY.toFixed(2)}, ${controlX2.toFixed(2)} ${endY.toFixed(2)}, ${endX.toFixed(2)} ${endY.toFixed(2)}`);
      path.setAttribute("fill", "none");
      path.setAttribute("stroke", trace.links.colors?.[linkIndex] ?? "rgba(37, 99, 235, 0.32)");
      path.setAttribute("stroke-width", String(strokeWidth));
      path.setAttribute("opacity", "0.85");
      this.attachInteractiveMetadata(path, traceEntry.index, linkIndex, trace.name, `${nodeLabels[source]} -> ${nodeLabels[target]}`, formatNumeric(value));
      svg.append(path);
    });

    positions.forEach((position, index) => {
      if (!position) {
        return;
      }
      const node = createSvgElement("rect");
      node.setAttribute("x", String(position.x));
      node.setAttribute("y", String(position.y));
      node.setAttribute("width", String(nodeWidth));
      node.setAttribute("height", String(position.height));
      node.setAttribute("rx", "2");
      node.setAttribute("fill", trace.nodes.colors?.[index] ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length]);
      node.setAttribute("opacity", "0.95");
      this.attachInteractiveMetadata(node, traceEntry.index, index, trace.name, nodeLabels[index], `in ${formatNumeric(incoming[index])} / out ${formatNumeric(outgoing[index])}`);
      svg.append(node);

      const label = createSvgElement("text");
      label.textContent = nodeLabels[index];
      label.setAttribute("x", String(position.x + nodeWidth + 4));
      label.setAttribute("y", String(position.y + 11));
      label.setAttribute("font-size", "11");
      label.setAttribute("font-family", "Arial, sans-serif");
      label.setAttribute("fill", "#0f172a");
      label.setAttribute("stroke", "none");
      svg.append(label);
    });
  }

  private renderParallelCategoriesScene(
    svg: SVGSVGElement,
    figure: ChartFigure,
    traceEntry: IndexedTrace<ParallelCategoriesTrace>,
    plotArea: ComputedPlotArea
  ): void {
    this.renderPlotBackground(svg, plotArea);
    const trace = traceEntry.trace;
    const dimensions = trace.dimensions;
    if (dimensions.length < 2) {
      return;
    }
    const recordCount = dimensions[0]?.values.length ?? 0;
    if (recordCount === 0) {
      return;
    }
    const axisX = dimensions.map((_, index) =>
      plotArea.x + (index / Math.max(1, dimensions.length - 1)) * plotArea.width
    );
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
      const filterLabel = createSvgElement("text");
      filterLabel.textContent = "Filtro ativo";
      filterLabel.setAttribute("x", String(plotArea.x + plotArea.width - 8));
      filterLabel.setAttribute("y", String(plotArea.y - 8));
      filterLabel.setAttribute("text-anchor", "end");
      filterLabel.setAttribute("font-size", "10");
      filterLabel.setAttribute("font-family", "Arial, sans-serif");
      filterLabel.setAttribute("fill", "#2563eb");
      filterLabel.setAttribute("stroke", "none");
      svg.append(filterLabel);
    }

    dimensions.forEach((dimension, dimensionIndex) => {
      const axis = createSvgElement("line");
      axis.setAttribute("x1", String(axisX[dimensionIndex]));
      axis.setAttribute("x2", String(axisX[dimensionIndex]));
      axis.setAttribute("y1", String(plotArea.y));
      axis.setAttribute("y2", String(plotArea.y + plotArea.height));
      axis.setAttribute("stroke", "#cbd5e1");
      axis.setAttribute("stroke-width", "1");
      svg.append(axis);

      const title = createSvgElement("text");
      title.textContent = dimension.name;
      title.setAttribute("x", String(axisX[dimensionIndex]));
      title.setAttribute("y", String(plotArea.y - 8));
      title.setAttribute("text-anchor", "middle");
      title.setAttribute("font-size", "11");
      title.setAttribute("font-family", "Arial, sans-serif");
      title.setAttribute("fill", "#334155");
      title.setAttribute("stroke", "none");
      svg.append(title);
    });

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
      const path = createSvgElement("path");
      path.setAttribute(
        "d",
        points.map((point, pointIndex) => `${pointIndex === 0 ? "M" : "L"} ${point.x.toFixed(2)} ${point.y.toFixed(2)}`).join(" ")
      );
      path.setAttribute("fill", "none");
      path.setAttribute("stroke", trace.line?.color ?? "rgba(37, 99, 235, 0.42)");
      path.setAttribute("stroke-width", selectedRows.includes(rowIndex) ? "1.8" : "1.1");
      path.setAttribute("opacity", String(selectedRows.includes(rowIndex) ? 0.95 : trace.line?.opacity ?? 0.7));
      this.attachInteractiveMetadata(path, traceEntry.index, rowIndex, trace.name, `row ${rowIndex + 1}`, points.map((point) => point.label).join(" | "));
      svg.append(path);
    }
  }

  private renderLegend(svg: SVGSVGElement, traces: IndexedTrace[], layout: ComputedLayout): void {
    const legendArea =
      layout.legendArea ??
      ({
        x: layout.margin.left,
        y: layout.margin.top,
        width: Math.max(1, layout.width - layout.margin.left - layout.margin.right),
        height: 44
      } as const);

    const legendGroup = createSvgElement("g");
    legendGroup.setAttribute("class", "excelsior-chart-legend");

    const horizontal = legendArea.width >= legendArea.height * 2;
    const rowHeight = 18;
    let cursorX = legendArea.x + 8;
    let cursorY = legendArea.y + 14;

    traces.forEach((entry, index) => {
      const labelText = entry.trace.name?.trim() || `Trace ${index + 1}`;
      const itemWidth = Math.min(legendArea.width - 12, Math.max(68, 18 + labelText.length * 7));
      if (horizontal && cursorX + itemWidth > legendArea.x + legendArea.width - 4) {
        cursorX = legendArea.x + 8;
        cursorY += rowHeight;
      }

      const item = createSvgElement("g");
      item.setAttribute("data-chart-legend", "item");
      item.setAttribute("data-legend-trace-index", String(entry.index));
      item.setAttribute("transform", `translate(${cursorX}, ${cursorY})`);
      item.setAttribute("role", "button");
      item.setAttribute("tabindex", "0");
      item.setAttribute("focusable", "true");
      item.setAttribute("aria-label", `Alternar visibilidade de ${labelText}`);
      item.style.cursor = "pointer";
      item.style.opacity = entry.trace.visible === false ? "0.45" : "1";

      const marker = createSvgElement("rect");
      marker.setAttribute("x", "0");
      marker.setAttribute("y", "-9");
      marker.setAttribute("width", "10");
      marker.setAttribute("height", "10");
      marker.setAttribute("fill", this.resolveTraceColor(entry.trace, entry.index));
      marker.setAttribute("stroke", "#64748b");
      marker.setAttribute("stroke-width", "0.5");
      item.append(marker);

      const label = createSvgElement("text");
      label.textContent = labelText;
      label.setAttribute("x", "14");
      label.setAttribute("y", "0");
      label.setAttribute("font-size", "11");
      label.setAttribute("font-family", "Arial, sans-serif");
      label.setAttribute("fill", "#334155");
      item.append(label);

      legendGroup.append(item);
      if (horizontal) {
        cursorX += itemWidth;
      } else {
        cursorY += rowHeight;
      }
    });

    svg.append(legendGroup);
  }

  private renderPlotBackground(svg: SVGSVGElement, plotArea: ComputedPlotArea): void {
    const chartSurface = createSvgElement("rect");
    chartSurface.setAttribute("x", String(plotArea.x));
    chartSurface.setAttribute("y", String(plotArea.y));
    chartSurface.setAttribute("width", String(plotArea.width));
    chartSurface.setAttribute("height", String(plotArea.height));
    chartSurface.setAttribute("fill", "#ffffff");
    chartSurface.setAttribute("stroke", "#dbe2ea");
    svg.append(chartSurface);
  }

  private createClippedLayer(svg: SVGSVGElement, plotArea: ComputedPlotArea): SVGGElement {
    const clipId = `excelsior-chart-clip-${this.clipCounter++}`;
    const defs = createSvgElement("defs");
    const clipPath = createSvgElement("clipPath");
    clipPath.setAttribute("id", clipId);
    const clipRect = createSvgElement("rect");
    clipRect.setAttribute("x", String(plotArea.x));
    clipRect.setAttribute("y", String(plotArea.y));
    clipRect.setAttribute("width", String(plotArea.width));
    clipRect.setAttribute("height", String(plotArea.height));
    clipPath.append(clipRect);
    defs.append(clipPath);
    svg.append(defs);

    const layer = createSvgElement("g");
    layer.setAttribute("clip-path", `url(#${clipId})`);
    return layer;
  }

  private attachInteractiveMetadata(
    element: Element,
    traceIndex: number,
    pointIndex: number,
    traceName: string | undefined,
    xLabel: string,
    yLabel: string
  ): void {
    element.setAttribute("data-chart-interactive", "point");
    element.setAttribute("data-trace-index", String(traceIndex));
    element.setAttribute("data-point-index", String(pointIndex));
    element.setAttribute("data-trace-name", traceName?.trim() || `Trace ${traceIndex + 1}`);
    element.setAttribute("data-point-x", xLabel);
    element.setAttribute("data-point-y", yLabel);
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
    if (isFinancialTrace(trace)) {
      return trace.line?.color ?? trace.increasing?.color ?? trace.decreasing?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isHeatmapTrace(trace)) {
      const colors = trace.colorscale && trace.colorscale.length > 0 ? trace.colorscale : DEFAULT_HEATMAP_COLORS;
      return colors[Math.floor(colors.length / 2)];
    }
    if (isContourTrace(trace)) {
      const colors = trace.colorscale && trace.colorscale.length > 0 ? trace.colorscale : DEFAULT_HEATMAP_COLORS;
      return colors[Math.floor(colors.length / 2)];
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
    if (isGeoScatterTrace(trace) || isGeoLineTrace(trace)) {
      if (isGeoScatterTrace(trace)) {
        return trace.marker?.color ?? trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
      }
      return trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isScatter3dTrace(trace) || isMesh3dTrace(trace)) {
      if (isScatter3dTrace(trace)) {
        return trace.marker?.color ?? trace.line?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
      }
      return trace.marker?.color ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
    }
    if (isSurfaceTrace(trace)) {
      const palette = resolvePalette("continuous", trace.colorscale, trace.reverseScale === true);
      return palette[Math.floor(palette.length / 2)] ?? DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
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

  private ensureSvg(): SVGSVGElement {
    if (!this.svgElement || !this.container) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_MOUNTED", "Renderer must be mounted before rendering.");
    }
    return this.svgElement;
  }
}

const createSvgElement = <TName extends keyof SVGElementTagNameMap>(name: TName): SVGElementTagNameMap[TName] =>
  document.createElementNS(SVG_NS, name);

const polarToCartesian = (cx: number, cy: number, radius: number, angle: number): { x: number; y: number } => ({
  x: cx + radius * Math.cos(angle),
  y: cy + radius * Math.sin(angle)
});

const describeArcSlice = (
  cx: number,
  cy: number,
  outerRadius: number,
  startAngle: number,
  endAngle: number,
  innerRadius: number
): string => {
  const startOuter = polarToCartesian(cx, cy, outerRadius, startAngle);
  const endOuter = polarToCartesian(cx, cy, outerRadius, endAngle);
  const largeArcFlag = endAngle - startAngle > Math.PI ? "1" : "0";

  if (innerRadius <= 0) {
    return [
      `M ${cx} ${cy}`,
      `L ${startOuter.x} ${startOuter.y}`,
      `A ${outerRadius} ${outerRadius} 0 ${largeArcFlag} 1 ${endOuter.x} ${endOuter.y}`,
      "Z"
    ].join(" ");
  }

  const startInner = polarToCartesian(cx, cy, innerRadius, endAngle);
  const endInner = polarToCartesian(cx, cy, innerRadius, startAngle);
  return [
    `M ${startOuter.x} ${startOuter.y}`,
    `A ${outerRadius} ${outerRadius} 0 ${largeArcFlag} 1 ${endOuter.x} ${endOuter.y}`,
    `L ${startInner.x} ${startInner.y}`,
    `A ${innerRadius} ${innerRadius} 0 ${largeArcFlag} 0 ${endInner.x} ${endInner.y}`,
    "Z"
  ].join(" ");
};

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

const clampInt = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? Math.round(value) : min;
  return Math.min(max, Math.max(min, numeric));
};

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

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
  const ratio = total <= 1 ? 0 : index / total;
  return ratio * Math.PI * 2 - Math.PI / 2;
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
  allCoordinates: Array<number[][][] | number[][][][]>
): { minLon: number; maxLon: number; minLat: number; maxLat: number } | null => {
  let minLon = Number.POSITIVE_INFINITY;
  let maxLon = Number.NEGATIVE_INFINITY;
  let minLat = Number.POSITIVE_INFINITY;
  let maxLat = Number.NEGATIVE_INFINITY;

  for (const coordinates of allCoordinates) {
    const paths = expandGeoPaths(coordinates);
    paths.forEach((path) => {
      path.forEach((coordinate) => {
        const lon = Number(coordinate[0]);
        const lat = Number(coordinate[1]);
        if (!Number.isFinite(lon) || !Number.isFinite(lat)) {
          return;
        }
        minLon = Math.min(minLon, lon);
        maxLon = Math.max(maxLon, lon);
        minLat = Math.min(minLat, lat);
        maxLat = Math.max(maxLat, lat);
      });
    });
  }

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
  return {
    minLon,
    maxLon,
    minLat,
    maxLat
  };
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

const distanceBetween = (x0: number, y0: number, x1: number, y1: number): number => Math.hypot(x1 - x0, y1 - y0);

const isSafeImageSource = (source: string): boolean => {
  const trimmed = source.trim();
  return trimmed.startsWith("data:image/") || trimmed.startsWith("./") || trimmed.startsWith("../") || trimmed.startsWith("/");
};

const computeGeoBoundsFromPoints = (
  points: Array<[number, number]>
): { minLon: number; maxLon: number; minLat: number; maxLat: number } | null => {
  if (points.length === 0) {
    return null;
  }
  let minLon = Number.POSITIVE_INFINITY;
  let maxLon = Number.NEGATIVE_INFINITY;
  let minLat = Number.POSITIVE_INFINITY;
  let maxLat = Number.NEGATIVE_INFINITY;
  points.forEach(([lon, lat]) => {
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

const mergeGeoBounds = (
  first: { minLon: number; maxLon: number; minLat: number; maxLat: number } | null,
  second: { minLon: number; maxLon: number; minLat: number; maxLat: number } | null
): { minLon: number; maxLon: number; minLat: number; maxLat: number } | null => {
  if (!first && !second) {
    return null;
  }
  if (!first) {
    return second;
  }
  if (!second) {
    return first;
  }
  return {
    minLon: Math.min(first.minLon, second.minLon),
    maxLon: Math.max(first.maxLon, second.maxLon),
    minLat: Math.min(first.minLat, second.minLat),
    maxLat: Math.max(first.maxLat, second.maxLat)
  };
};

const createDefaultProjection3d = (
  plotArea: ComputedPlotArea
): ((x: number, y: number, z: number) => { x: number; y: number }) => {
  const angleY = (38 * Math.PI) / 180;
  const angleX = (-30 * Math.PI) / 180;
  const scale = Math.min(plotArea.width, plotArea.height) * 0.34;
  const cx = plotArea.x + plotArea.width / 2;
  const cy = plotArea.y + plotArea.height / 2;

  return (x: number, y: number, z: number) => {
    const nx = x;
    const ny = y;
    const nz = z;
    const cosY = Math.cos(angleY);
    const sinY = Math.sin(angleY);
    const cosX = Math.cos(angleX);
    const sinX = Math.sin(angleX);
    const x1 = nx * cosY - nz * sinY;
    const z1 = nx * sinY + nz * cosY;
    const y2 = ny * cosX - z1 * sinX;
    const z2 = ny * sinX + z1 * cosX;
    const perspective = 1 / Math.max(0.2, 1 + z2 * 0.045);
    return {
      x: cx + x1 * scale * perspective,
      y: cy - y2 * scale * perspective
    };
  };
};
