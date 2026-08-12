import type { ChartFigure, ChartFigureInput, ChartSelectedPoint, ChartSelectionState } from "../model/Figure";
import type { ChartLayout } from "../model/Layout";
import type {
  CartesianTrace,
  ChartTrace,
  DensityTrace,
  DonutTrace,
  FunnelTrace,
  PieTrace,
  SunburstTrace,
  TreemapTrace,
  ViolinTrace
} from "../model/Trace";
import type { ChartEventMap } from "../model/Events";
import { FigureManager } from "./FigureManager";
import { FigureValidator } from "./FigureValidator";
import { LayoutEngine, type ComputedLayout } from "./LayoutEngine";
import { TraceRegistry } from "./TraceRegistry";
import { TypedEventEmitter } from "./TypedEventEmitter";
import { SvgRenderer } from "../renderers/SvgRenderer";
import { CanvasRenderer } from "../renderers/CanvasRenderer";
import { WebglRenderer } from "../renderers/WebglRenderer";
import { HybridRenderer } from "../renderers/HybridRenderer";
import type { ChartRenderer } from "../renderers/ChartRenderer";
import { ChartConfigurationError } from "./chart-errors";
import { TooltipController } from "../interactions/TooltipController";
import { HoverController } from "../interactions/HoverController";
import { LegendController } from "../interactions/LegendController";
import { ModebarController } from "../interactions/ModebarController";
import { ZoomPanController, type InteractionMode } from "../interactions/ZoomPanController";
import { SelectionController, type SelectionShape } from "../interactions/SelectionController";
import { RenderScheduler } from "./RenderScheduler";
import { buildCartesianDomains } from "./cartesian-domain";
import type { FramePlaybackOptions } from "../types/PublicApi";

export class ChartEngine {
  private readonly events = new TypedEventEmitter<ChartEventMap>();
  private readonly traceRegistry: TraceRegistry;
  private readonly validator: FigureValidator;
  private readonly figureManager: FigureManager;
  private readonly layoutEngine = new LayoutEngine();
  private renderer: ChartRenderer;
  private readonly isRendererManaged: boolean;
  private readonly renderScheduler = new RenderScheduler();
  private readonly tooltipController = new TooltipController();
  private readonly hoverController: HoverController;
  private readonly legendController: LegendController;
  private readonly zoomPanController: ZoomPanController;
  private readonly selectionController: SelectionController;
  private readonly modebarController: ModebarController;
  private readonly resizeObserver?: ResizeObserver;
  private lastComputedLayout: ComputedLayout | null = null;
  private interactionMode: InteractionMode = "zoom";
  private animationTimer: number | null = null;
  private animationFrameIndex = -1;
  private animationLoop = true;
  private accessibilityTable: HTMLTableElement | null = null;
  private isDestroyed = false;
  private readonly subplotAxisOverrides = new Map<number, { x: [number, number]; y: [number, number] }>();
  private lastInteractionSubplotIndex = 0;

  constructor(
    private readonly container: HTMLElement,
    initialFigure: ChartFigureInput,
    renderer?: ChartRenderer
  ) {
    this.traceRegistry = new TraceRegistry();
    this.validator = new FigureValidator(this.traceRegistry);
    this.figureManager = new FigureManager(initialFigure, this.validator);
    this.isRendererManaged = !renderer;
    const initialConfig = this.figureManager.getFigure().config;
    this.renderer = renderer ?? this.createRenderer(initialConfig.renderer, initialConfig.webglFallback);
    this.hoverController = new HoverController({
      tooltip: this.tooltipController,
      onHover: (payload) => this.emit("trace:hover", payload),
      onUnhover: ({ traceIndex, pointIndex }) => this.emit("trace:unhover", { traceIndex, pointIndex }),
      onClick: (payload) => {
        this.handleTraceClick(payload.traceIndex, payload.pointIndex);
        this.emit("trace:click", payload);
        this.applyPointSelection("click", [payload], true);
      }
    });
    this.legendController = new LegendController((traceIndex) => this.toggleTraceVisibility(traceIndex));
    this.zoomPanController = new ZoomPanController({
      onZoom: (factor, anchorX, anchorY) => this.applyZoom(factor, anchorX, anchorY),
      onPan: (deltaX, deltaY, anchorX, anchorY) => this.applyPan(deltaX, deltaY, anchorX, anchorY),
      onZoomAxis: (axis, factor, anchorX, anchorY) => this.applyAxisZoom(axis, factor, anchorX, anchorY),
      onZoomRect: (bounds) => this.applyRectZoom(bounds)
    });
    this.selectionController = new SelectionController({
      onSelection: (points, shape) => this.applyPointSelection(shape, points, false),
      onCleared: () => this.applyPointSelection("rect", [], false)
    });
    this.modebarController = new ModebarController({
      onZoomIn: () => this.applyZoom(0.84),
      onZoomOut: () => this.applyZoom(1.18),
      onResetAxis: () => this.resetAxes(),
      onSetMode: (mode) => this.setInteractionMode(mode),
      onToggleLegend: () => this.toggleLegendVisibility(),
      onToggleFullscreen: () => void this.toggleFullscreen(),
      onExportSvg: () => this.downloadSvg(),
      onExportPng: () => void this.downloadPng(),
      allowFullscreen: this.figureManager.getFigure().config.fullscreen
    });

    this.renderer.mount(container);
    this.renderCurrentFigure(true);
    this.syncControllers();

    const figure = this.figureManager.getFigure();
    if (figure.config.responsive && typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => {
        this.resize();
      });
      this.resizeObserver.observe(this.container);
    }

    this.emit("figure:created", { figure });
  }

  update(nextFigure: ChartFigureInput): void {
    this.assertActive();
    this.subplotAxisOverrides.clear();
    const currentRendererKind = this.figureManager.getFigure().config.renderer;
    const figure = this.figureManager.update(nextFigure);
    if (this.isRendererManaged && figure.config.renderer !== currentRendererKind) {
      this.swapRenderer(figure.config.renderer, figure.config.webglFallback);
    }
    this.renderCurrentFigure(true);
    this.syncControllers();
    this.emit("figure:updated", { figure });
  }

  updateData(nextData: ChartTrace[]): void {
    this.assertActive();
    this.subplotAxisOverrides.clear();
    const figure = this.figureManager.updateData(nextData);
    this.renderCurrentFigure(true);
    this.syncControllers();
    this.emit("figure:updated", { figure });
  }

  updateLayout(nextLayout: Partial<ChartLayout>): void {
    this.assertActive();
    this.subplotAxisOverrides.clear();
    const figure = this.figureManager.updateLayout(nextLayout);
    this.renderCurrentFigure(true);
    this.syncControllers();
    this.emit("figure:updated", { figure });
  }

  resize(): void {
    this.assertActive();
    this.renderCurrentFigure(false, true);
  }

  destroy(): void {
    if (this.isDestroyed) {
      return;
    }

    this.stopFrames();
    this.resizeObserver?.disconnect();
    this.renderScheduler.destroy();
    this.hoverController.destroy();
    this.tooltipController.destroy();
    this.legendController.destroy();
    this.zoomPanController.destroy();
    this.selectionController.destroy();
    this.modebarController.destroy();
    this.renderer.destroy();
    this.accessibilityTable?.remove();
    this.accessibilityTable = null;
    this.isDestroyed = true;
    this.emit("figure:destroyed", { container: this.container });
    this.events.clear();
  }

  exportSvg(): string {
    this.assertActive();
    this.emit("export:started", { format: "svg" });
    const svg = this.renderer.exportSvg();
    this.emit("export:finished", { format: "svg" });
    return svg;
  }

  async exportPng(options?: { scale?: number; backgroundColor?: string }): Promise<Blob> {
    this.assertActive();
    this.emit("export:started", { format: "png" });
    const blob = await this.renderer.exportPng(options);
    this.emit("export:finished", { format: "png" });
    return blob;
  }

  toJson(): string {
    this.assertActive();
    return JSON.stringify(this.figureManager.getFigure());
  }

  getSelection(): ChartSelectionState | null {
    this.assertActive();
    const selection = this.figureManager.getFigure().selection;
    return selection ? structuredClone(selection) : null;
  }

  clearSelection(): void {
    this.assertActive();
    this.applyPointSelection("rect", [], false);
  }

  exportDataTable(): string {
    this.assertActive();
    const figure = this.figureManager.getFigure();
    const rows: string[] = ["trace,x,y"];
    figure.data.forEach((trace, traceIndex) => {
      if (trace.visible === false) {
        return;
      }
      const traceName = escapeCsv(trace.name?.trim() || `Trace ${traceIndex + 1}`);
      if (isCartesianDataTrace(trace)) {
        trace.y.forEach((value, pointIndex) => {
          const x = escapeCsv(String(trace.x[pointIndex] ?? pointIndex));
          const y = escapeCsv(String(value));
          rows.push(`${traceName},${x},${y}`);
        });
      } else if (isPieDataTrace(trace)) {
        trace.values.forEach((value, sliceIndex) => {
          const x = escapeCsv(String(trace.labels?.[sliceIndex] ?? `Slice ${sliceIndex + 1}`));
          const y = escapeCsv(String(value));
          rows.push(`${traceName},${x},${y}`);
        });
      } else if (isValueDataTrace(trace)) {
        trace.values.forEach((value, pointIndex) => {
          rows.push(`${traceName},${escapeCsv(String(pointIndex))},${escapeCsv(String(value))}`);
        });
      } else if (isFunnelDataTrace(trace)) {
        trace.values.forEach((value, pointIndex) => {
          rows.push(`${traceName},${escapeCsv(String(trace.labels?.[pointIndex] ?? pointIndex))},${escapeCsv(String(value))}`);
        });
      }
    });
    return rows.join("\n");
  }

  playFrames(options?: FramePlaybackOptions): void {
    this.assertActive();
    const figure = this.figureManager.getFigure();
    if (!figure.frames || figure.frames.length === 0) {
      return;
    }
    this.stopFrames();
    this.animationLoop = options?.loop ?? true;
    const intervalMs = clamp(options?.intervalMs ?? figure.config.frameDurationMs, 16, 60_000);
    this.emit("animation:started", { frameCount: figure.frames.length, intervalMs });

    const tick = (): void => {
      if (this.isDestroyed) {
        return;
      }
      const current = this.figureManager.getFigure();
      const frames = current.frames;
      if (!frames || frames.length === 0) {
        this.stopFrames();
        return;
      }
      this.animationFrameIndex += 1;
      if (this.animationFrameIndex >= frames.length) {
        if (!this.animationLoop) {
          this.stopFrames();
          return;
        }
        this.animationFrameIndex = 0;
      }

      const frame = frames[this.animationFrameIndex];
      const figureUpdate: ChartFigureInput = {
        ...current,
        data: frame.data ?? current.data,
        layout: frame.layout ? { ...current.layout, ...frame.layout } : current.layout
      };
      const normalized = this.figureManager.update(figureUpdate);
      this.renderCurrentFigure(false);
      this.syncControllers();
      this.emit("animation:frame", { frameIndex: this.animationFrameIndex, frameName: frame.name });
      this.emit("figure:updated", { figure: normalized });
      this.animationTimer = window.setTimeout(tick, intervalMs);
    };

    this.animationFrameIndex = -1;
    this.animationTimer = window.setTimeout(tick, intervalMs);
  }

  stopFrames(): void {
    const hadAnimation = this.animationTimer !== null || this.animationFrameIndex >= 0;
    if (this.animationTimer !== null) {
      clearTimeout(this.animationTimer);
      this.animationTimer = null;
    }
    const lastFrame = Math.max(0, this.animationFrameIndex);
    this.animationFrameIndex = -1;
    if (hadAnimation) {
      this.emit("animation:stopped", { frameIndex: lastFrame });
    }
  }

  isAnimating(): boolean {
    return this.animationTimer !== null;
  }

  on<TKey extends keyof ChartEventMap>(event: TKey, handler: (payload: ChartEventMap[TKey]) => void): void {
    this.events.on(event, handler);
  }

  off<TKey extends keyof ChartEventMap>(event: TKey, handler: (payload: ChartEventMap[TKey]) => void): void {
    this.events.off(event, handler);
  }

  getFigureSnapshot(): ChartFigure {
    this.assertActive();
    return structuredClone(this.figureManager.getFigure());
  }

  getRegisteredTraceTypes(): string[] {
    return this.traceRegistry.listTypes();
  }

  private renderCurrentFigure(immediate: boolean, emitResizeEvent = false): void {
    const renderTask = () => {
      const figure = this.figureManager.getFigure();
      const computedLayout = this.layoutEngine.compute(this.container, figure);
      this.lastComputedLayout = computedLayout;
      this.renderer.resize(computedLayout);
      this.renderer.render(this.createRuntimeFigure(figure), computedLayout);
      if (emitResizeEvent) {
        this.emit("layout:resized", { width: computedLayout.width, height: computedLayout.height });
      }
    };

    if (immediate) {
      this.renderScheduler.flushNow(renderTask);
      return;
    }
    this.renderScheduler.schedule(renderTask);
  }

  private syncControllers(): void {
    const root = this.renderer.getRootElement();
    if (!root) {
      return;
    }

    const figure = this.figureManager.getFigure();
    if (figure.config.tooltip) {
      this.tooltipController.mount(this.container);
      this.hoverController.mount(root);
      this.hoverController.setHoverMode(figure.config.hoverMode);
      this.hoverController.setSpatialHoverEnabled(figure.config.spatialHover);
    } else {
      this.hoverController.unmount();
      this.tooltipController.hide();
    }

    this.legendController.mount(root);
    this.zoomPanController.mount(root);
    this.zoomPanController.setMode(this.interactionMode);
    this.selectionController.mount(root, this.container);
    this.selectionController.setEnabled(this.interactionMode === "select");
    if (root instanceof HTMLElement || root instanceof SVGElement) {
      this.applyInteractionCursor(root);
    }

    if (figure.config.modebar) {
      this.modebarController.mount(this.container);
      this.modebarController.setMode(this.interactionMode);
    } else {
      this.modebarController.destroy();
    }

    this.syncAccessibility(figure);
  }

  private toggleTraceVisibility(traceIndex: number): void {
    this.assertActive();
    const currentFigure = this.figureManager.getFigure();
    const target = currentFigure.data[traceIndex];
    if (!target) {
      return;
    }

    const nextVisible = target.visible === false;
    const updatedData = currentFigure.data.map((trace, index) => {
      if (index !== traceIndex) {
        return trace;
      }
      return {
        ...trace,
        visible: nextVisible
      };
    });

    const figure = this.figureManager.updateData(updatedData);
    this.renderCurrentFigure(true);
    this.emit("legend:toggled", { traceIndex, visible: nextVisible });
    this.emit("figure:updated", { figure });
  }

  private applyZoom(factor: number, anchorClientX?: number, anchorClientY?: number): void {
    this.assertActive();
    const figure = this.figureManager.getFigure();
    const interaction = this.resolveInteractionContext(anchorClientX, anchorClientY);
    const domains = this.resolveInteractionDomains(figure, interaction.subplotIndex);
    if (!domains) {
      return;
    }

    const xRange = domains.x[1] - domains.x[0];
    const yRange = domains.y[1] - domains.y[0];
    if (xRange <= 0 || yRange <= 0) {
      return;
    }

    const rect = this.container.getBoundingClientRect();
    const chartX = anchorClientX === undefined ? interaction.plotArea.x + interaction.plotArea.width / 2 : anchorClientX - rect.left;
    const chartY = anchorClientY === undefined ? interaction.plotArea.y + interaction.plotArea.height / 2 : anchorClientY - rect.top;
    const xRatio = clamp((chartX - interaction.plotArea.x) / Math.max(1, interaction.plotArea.width), 0, 1);
    const yRatio = clamp((chartY - interaction.plotArea.y) / Math.max(1, interaction.plotArea.height), 0, 1);
    const xAnchor = domains.x[0] + xRatio * xRange;
    const yAnchor = domains.y[1] - yRatio * yRange;

    const nextXMin = xAnchor - (xAnchor - domains.x[0]) * factor;
    const nextXMax = xAnchor + (domains.x[1] - xAnchor) * factor;
    const nextYMin = yAnchor - (yAnchor - domains.y[0]) * factor;
    const nextYMax = yAnchor + (domains.y[1] - yAnchor) * factor;
    this.applyAxisRanges(nextXMin, nextXMax, nextYMin, nextYMax, "axis:zoomed", interaction.subplotIndex);
  }

  private applyAxisZoom(axis: "x" | "y", factor: number, anchorClientX?: number, anchorClientY?: number): void {
    this.assertActive();
    const figure = this.figureManager.getFigure();
    const interaction = this.resolveInteractionContext(anchorClientX, anchorClientY);
    const domains = this.resolveInteractionDomains(figure, interaction.subplotIndex);
    if (!domains) {
      return;
    }
    const xRange = domains.x[1] - domains.x[0];
    const yRange = domains.y[1] - domains.y[0];
    if (xRange <= 0 || yRange <= 0) {
      return;
    }
    const rect = this.container.getBoundingClientRect();
    const chartX = anchorClientX === undefined ? interaction.plotArea.x + interaction.plotArea.width / 2 : anchorClientX - rect.left;
    const chartY = anchorClientY === undefined ? interaction.plotArea.y + interaction.plotArea.height / 2 : anchorClientY - rect.top;
    const xRatio = clamp((chartX - interaction.plotArea.x) / Math.max(1, interaction.plotArea.width), 0, 1);
    const yRatio = clamp((chartY - interaction.plotArea.y) / Math.max(1, interaction.plotArea.height), 0, 1);
    const xAnchor = domains.x[0] + xRatio * xRange;
    const yAnchor = domains.y[1] - yRatio * yRange;
    if (axis === "x") {
      const nextXMin = xAnchor - (xAnchor - domains.x[0]) * factor;
      const nextXMax = xAnchor + (domains.x[1] - xAnchor) * factor;
      this.applyAxisRanges(nextXMin, nextXMax, domains.y[0], domains.y[1], "axis:zoomed", interaction.subplotIndex);
      return;
    }
    const nextYMin = yAnchor - (yAnchor - domains.y[0]) * factor;
    const nextYMax = yAnchor + (domains.y[1] - yAnchor) * factor;
    this.applyAxisRanges(domains.x[0], domains.x[1], nextYMin, nextYMax, "axis:zoomed", interaction.subplotIndex);
  }

  private applyRectZoom(bounds: { x0: number; y0: number; x1: number; y1: number }): void {
    this.assertActive();
    const figure = this.figureManager.getFigure();
    const centerX = (bounds.x0 + bounds.x1) / 2;
    const centerY = (bounds.y0 + bounds.y1) / 2;
    const interaction = this.resolveInteractionContext(centerX, centerY);
    const domains = this.resolveInteractionDomains(figure, interaction.subplotIndex);
    if (!domains) {
      return;
    }
    const rect = this.container.getBoundingClientRect();
    const left = Math.min(bounds.x0, bounds.x1) - rect.left;
    const right = Math.max(bounds.x0, bounds.x1) - rect.left;
    const top = Math.min(bounds.y0, bounds.y1) - rect.top;
    const bottom = Math.max(bounds.y0, bounds.y1) - rect.top;

    const xStart = clamp((left - interaction.plotArea.x) / Math.max(1, interaction.plotArea.width), 0, 1);
    const xEnd = clamp((right - interaction.plotArea.x) / Math.max(1, interaction.plotArea.width), 0, 1);
    const yStart = clamp((top - interaction.plotArea.y) / Math.max(1, interaction.plotArea.height), 0, 1);
    const yEnd = clamp((bottom - interaction.plotArea.y) / Math.max(1, interaction.plotArea.height), 0, 1);
    const nextXMin = domains.x[0] + xStart * (domains.x[1] - domains.x[0]);
    const nextXMax = domains.x[0] + xEnd * (domains.x[1] - domains.x[0]);
    const nextYMax = domains.y[1] - yStart * (domains.y[1] - domains.y[0]);
    const nextYMin = domains.y[1] - yEnd * (domains.y[1] - domains.y[0]);
    this.applyAxisRanges(nextXMin, nextXMax, nextYMin, nextYMax, "axis:zoomed", interaction.subplotIndex);
  }

  private applyPan(deltaX: number, deltaY: number, anchorClientX?: number, anchorClientY?: number): void {
    this.assertActive();
    const figure = this.figureManager.getFigure();
    const interaction = this.resolveInteractionContext(anchorClientX, anchorClientY);
    const domains = this.resolveInteractionDomains(figure, interaction.subplotIndex);
    if (!domains) {
      return;
    }

    const xRange = domains.x[1] - domains.x[0];
    const yRange = domains.y[1] - domains.y[0];
    if (xRange <= 0 || yRange <= 0) {
      return;
    }

    const xShift = (deltaX / Math.max(1, interaction.plotArea.width)) * xRange;
    const yShift = (deltaY / Math.max(1, interaction.plotArea.height)) * yRange;
    const nextXMin = domains.x[0] - xShift;
    const nextXMax = domains.x[1] - xShift;
    const nextYMin = domains.y[0] + yShift;
    const nextYMax = domains.y[1] + yShift;
    this.applyAxisRanges(nextXMin, nextXMax, nextYMin, nextYMax, "axis:panned", interaction.subplotIndex);
  }

  private resetAxes(): void {
    this.assertActive();
    const current = this.figureManager.getFigure();
    if (!this.shouldSyncSubplotZoom(current)) {
      if (this.subplotAxisOverrides.size > 0) {
        this.subplotAxisOverrides.clear();
        this.renderCurrentFigure(false);
        this.emit("figure:updated", { figure: current });
      }
      return;
    }
    this.subplotAxisOverrides.clear();
    const figure = this.figureManager.updateLayout({
      xAxis: {
        ...current.layout.xAxis,
        min: undefined,
        max: undefined
      },
      yAxis: {
        ...current.layout.yAxis,
        min: undefined,
        max: undefined
      }
    });
    this.renderCurrentFigure(false);
    this.emit("figure:updated", { figure });
  }

  private applyAxisRanges(
    xMin: number,
    xMax: number,
    yMin: number,
    yMax: number,
    eventName: "axis:zoomed" | "axis:panned",
    subplotIndex?: number
  ): void {
    const safeX = normalizeRange(xMin, xMax, 0.5);
    const safeY = normalizeRange(yMin, yMax, 1e-6);
    const current = this.figureManager.getFigure();
    const subplotCount = this.lastComputedLayout?.subplotAreas.length ?? 1;
    if (!this.shouldSyncSubplotZoom(current) && subplotCount > 1) {
      const targetSubplot = clampInt(subplotIndex ?? this.lastInteractionSubplotIndex, 0, subplotCount - 1);
      this.subplotAxisOverrides.set(targetSubplot, { x: safeX, y: safeY });
      this.renderCurrentFigure(false);
      this.emit(eventName, {
        xMin: safeX[0],
        xMax: safeX[1],
        yMin: safeY[0],
        yMax: safeY[1]
      });
      this.emit("figure:updated", { figure: current });
      return;
    }
    this.subplotAxisOverrides.clear();
    const figure = this.figureManager.updateLayout({
      xAxis: {
        ...current.layout.xAxis,
        min: safeX[0],
        max: safeX[1]
      },
      yAxis: {
        ...current.layout.yAxis,
        min: safeY[0],
        max: safeY[1]
      }
    });

    this.renderCurrentFigure(false);
    this.emit(eventName, {
      xMin: safeX[0],
      xMax: safeX[1],
      yMin: safeY[0],
      yMax: safeY[1]
    });
    this.emit("figure:updated", { figure });
  }

  private resolveInteractionContext(anchorClientX?: number, anchorClientY?: number): { subplotIndex: number; plotArea: ComputedLayout["plotArea"] } {
    const layout = this.lastComputedLayout;
    if (!layout) {
      return {
        subplotIndex: 0,
        plotArea: {
          x: 0,
          y: 0,
          width: Math.max(1, this.container.clientWidth || this.container.getBoundingClientRect().width || 1),
          height: Math.max(1, this.container.clientHeight || this.container.getBoundingClientRect().height || 1)
        }
      };
    }
    const subplotAreas = layout.subplotAreas.length > 0 ? layout.subplotAreas : [layout.plotArea];
    if (anchorClientX === undefined || anchorClientY === undefined) {
      const subplotIndex = clampInt(this.lastInteractionSubplotIndex, 0, subplotAreas.length - 1);
      this.lastInteractionSubplotIndex = subplotIndex;
      return {
        subplotIndex,
        plotArea: subplotAreas[subplotIndex] ?? layout.plotArea
      };
    }

    const rect = this.container.getBoundingClientRect();
    const localX = anchorClientX - rect.left;
    const localY = anchorClientY - rect.top;
    const matchedIndex = subplotAreas.findIndex(
      (area) => localX >= area.x && localX <= area.x + area.width && localY >= area.y && localY <= area.y + area.height
    );
    const subplotIndex = matchedIndex >= 0 ? matchedIndex : clampInt(this.lastInteractionSubplotIndex, 0, subplotAreas.length - 1);
    this.lastInteractionSubplotIndex = subplotIndex;
    return {
      subplotIndex,
      plotArea: subplotAreas[subplotIndex] ?? layout.plotArea
    };
  }

  private resolveInteractionDomains(figure: ChartFigure, subplotIndex: number): ReturnType<typeof buildCartesianDomains> {
    const syncSubplotZoom = this.shouldSyncSubplotZoom(figure);
    if (syncSubplotZoom) {
      return buildCartesianDomains(figure);
    }
    const subplotTraces = figure.data.filter((trace) => {
      if (trace.visible === false) {
        return false;
      }
      const traceSubplot = clampInt((trace.subplot as number | undefined) ?? 0, 0, Number.MAX_SAFE_INTEGER);
      if (traceSubplot !== subplotIndex) {
        return false;
      }
      return isCartesianDataTrace(trace) || isFinancialDataTrace(trace);
    });
    if (subplotTraces.length === 0) {
      return buildCartesianDomains(figure);
    }
    const subsetFigure: ChartFigure = {
      ...figure,
      data: subplotTraces
    };
    const domains = buildCartesianDomains(subsetFigure);
    if (!domains) {
      return null;
    }
    const override = this.subplotAxisOverrides.get(subplotIndex);
    if (!override) {
      return domains;
    }
    return {
      ...domains,
      x: [...override.x] as [number, number],
      y: [...override.y] as [number, number]
    };
  }

  private shouldSyncSubplotZoom(figure: ChartFigure): boolean {
    return figure.config.syncSubplotZoom !== false && figure.layout.subplots?.syncZoom !== false;
  }

  private createRuntimeFigure(figure: ChartFigure): ChartFigure {
    if (this.subplotAxisOverrides.size === 0) {
      return figure;
    }
    const ranges: Record<string, { x: [number, number]; y: [number, number] }> = {};
    this.subplotAxisOverrides.forEach((value, subplotIndex) => {
      ranges[String(subplotIndex)] = {
        x: [...value.x] as [number, number],
        y: [...value.y] as [number, number]
      };
    });
    return {
      ...figure,
      __subplotAxisRanges: ranges
    } as ChartFigure;
  }

  private handleTraceClick(traceIndex: number, pointIndex: number): void {
    const figure = this.figureManager.getFigure();
    const trace = figure.data[traceIndex];
    if (!trace || trace.visible === false) {
      return;
    }
    if (isSunburstDataTrace(trace) || isTreemapDataTrace(trace)) {
      this.applyHierarchyDrilldown(traceIndex, trace, pointIndex);
    }
  }

  private applyHierarchyDrilldown(traceIndex: number, trace: SunburstTrace | TreemapTrace, pointIndex: number): void {
    const ids = trace.labels.map((label, index) => String(trace.ids?.[index] ?? label ?? index));
    const parentById = new Map<string, string>();
    trace.parents.forEach((parent, index) => {
      parentById.set(ids[index], String(parent ?? ""));
    });
    const currentRoot = trace.rootId !== undefined ? String(trace.rootId) : "";
    const nextRoot = this.resolveHierarchyRoot(trace.parents, ids, parentById, currentRoot, pointIndex);
    if (nextRoot === null || nextRoot === currentRoot) {
      return;
    }
    const figure = this.figureManager.getFigure();
    const nextTrace: SunburstTrace | TreemapTrace = {
      ...trace,
      rootId: nextRoot || undefined
    };
    const updatedData = figure.data.map((candidate, index) => (index === traceIndex ? nextTrace : candidate));
    const nextFigure = this.figureManager.updateData(updatedData);
    this.renderCurrentFigure(true);
    this.syncControllers();
    this.emit("figure:updated", { figure: nextFigure });
  }

  private resolveHierarchyRoot(
    parents: Array<string | number>,
    ids: string[],
    parentById: Map<string, string>,
    currentRoot: string,
    pointIndex: number
  ): string | null {
    if (pointIndex < 0) {
      if (!currentRoot) {
        return null;
      }
      return parentById.get(currentRoot) ?? "";
    }
    const clickedId = ids[pointIndex];
    if (!clickedId) {
      return null;
    }
    const hasChildren = parents.some((parent) => String(parent ?? "") === clickedId);
    if (hasChildren) {
      if (clickedId === currentRoot) {
        return parentById.get(clickedId) ?? "";
      }
      return clickedId;
    }
    return null;
  }

  private setInteractionMode(mode: InteractionMode): void {
    this.interactionMode = mode;
    this.zoomPanController.setMode(mode);
    this.selectionController.setEnabled(mode === "select");
    const root = this.renderer.getRootElement();
    if (root instanceof HTMLElement || root instanceof SVGElement) {
      this.applyInteractionCursor(root);
    }
  }

  private toggleLegendVisibility(): void {
    this.assertActive();
    const current = this.figureManager.getFigure();
    const figure = this.figureManager.updateLayout({
      legend: {
        ...current.layout.legend,
        visible: !current.layout.legend.visible
      }
    });
    this.renderCurrentFigure(true);
    this.emit("figure:updated", { figure });
  }

  private downloadSvg(): void {
    const content = this.exportSvg();
    this.downloadFile("chart.svg", content, "image/svg+xml;charset=utf-8");
  }

  private async downloadPng(): Promise<void> {
    const blob = await this.exportPng();
    this.downloadBlob("chart.png", blob);
  }

  private downloadFile(filename: string, content: string, mimeType: string): void {
    const blob = new Blob([content], { type: mimeType });
    this.downloadBlob(filename, blob);
  }

  private downloadBlob(filename: string, blob: Blob): void {
    if (typeof URL === "undefined" || typeof document === "undefined") {
      return;
    }
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.style.display = "none";
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }

  private async toggleFullscreen(): Promise<void> {
    if (typeof document === "undefined") {
      return;
    }
    if (!this.figureManager.getFigure().config.fullscreen) {
      return;
    }

    const element = this.container;
    if (document.fullscreenElement === element) {
      await document.exitFullscreen?.();
      return;
    }
    await element.requestFullscreen?.();
  }

  private syncAccessibility(figure: ChartFigure): void {
    const description = figure.config.ariaDescription?.trim() || figure.layout.title?.trim() || "Chart visualization";
    this.container.setAttribute("role", "img");
    this.container.setAttribute("aria-label", description);

    if (!figure.config.accessibleTable) {
      this.accessibilityTable?.remove();
      this.accessibilityTable = null;
    } else {
      const table = this.accessibilityTable ?? document.createElement("table");
      table.className = "excelsior-chart-accessible-table";
      Object.assign(table.style, {
        marginTop: "8px",
        width: "100%",
        borderCollapse: "collapse",
        fontSize: "12px",
        fontFamily: "Arial, sans-serif"
      });
      table.setAttribute("aria-label", `${description} data table`);
      table.replaceChildren();
      const headerRow = document.createElement("tr");
      ["Trace", "X", "Y"].forEach((title) => {
        const th = document.createElement("th");
        th.textContent = title;
        Object.assign(th.style, {
          textAlign: "left",
          borderBottom: "1px solid #cbd5e1",
          padding: "4px 6px"
        });
        headerRow.append(th);
      });
      const thead = document.createElement("thead");
      thead.append(headerRow);
      table.append(thead);

      const tbody = document.createElement("tbody");
      const maxRows = 600;
      let rowCount = 0;
      figure.data.forEach((trace, traceIndex) => {
        if (trace.visible === false || rowCount >= maxRows) {
          return;
        }
        const traceName = trace.name?.trim() || `Trace ${traceIndex + 1}`;
        if (isCartesianDataTrace(trace)) {
          trace.y.forEach((value, pointIndex) => {
            if (rowCount >= maxRows) {
              return;
            }
            const row = document.createElement("tr");
            row.append(createTableCell(traceName), createTableCell(String(trace.x[pointIndex] ?? pointIndex)), createTableCell(String(value)));
            tbody.append(row);
            rowCount += 1;
          });
        } else if (isPieDataTrace(trace)) {
          trace.values.forEach((value, sliceIndex) => {
            if (rowCount >= maxRows) {
              return;
            }
            const label = trace.labels?.[sliceIndex] ?? `Slice ${sliceIndex + 1}`;
            const row = document.createElement("tr");
            row.append(createTableCell(traceName), createTableCell(String(label)), createTableCell(String(value)));
            tbody.append(row);
            rowCount += 1;
          });
        } else if (isValueDataTrace(trace)) {
          trace.values.forEach((value, pointIndex) => {
            if (rowCount >= maxRows) {
              return;
            }
            const row = document.createElement("tr");
            row.append(createTableCell(traceName), createTableCell(String(pointIndex)), createTableCell(String(value)));
            tbody.append(row);
            rowCount += 1;
          });
        } else if (isFunnelDataTrace(trace)) {
          trace.values.forEach((value, pointIndex) => {
            if (rowCount >= maxRows) {
              return;
            }
            const label = trace.labels?.[pointIndex] ?? `Step ${pointIndex + 1}`;
            const row = document.createElement("tr");
            row.append(createTableCell(traceName), createTableCell(String(label)), createTableCell(String(value)));
            tbody.append(row);
            rowCount += 1;
          });
        }
      });
      table.append(tbody);

      if (!this.accessibilityTable) {
        this.container.append(table);
      }
      this.accessibilityTable = table;
    }

    if (figure.config.highContrast) {
      const contrast = computeContrastRatio(figure.layout.backgroundColor, "#0f172a");
      if (contrast < 4.5) {
        this.emit("error:raised", {
          error: new ChartConfigurationError(
            "CHART_CONTRAST_LOW",
            `Contrast ratio is ${contrast.toFixed(2)}:1, below recommended minimum 4.5:1.`
          )
        });
      }
    }
  }

  private emit<TKey extends keyof ChartEventMap>(event: TKey, payload: ChartEventMap[TKey]): void {
    try {
      this.events.emit(event, payload);
    } catch (error) {
      if (event !== "error:raised") {
        this.events.emit("error:raised", { error: this.toError(error) });
      }
    }
  }

  private assertActive(): void {
    if (this.isDestroyed) {
      throw new ChartConfigurationError("CHART_INSTANCE_DESTROYED", "Chart instance is destroyed.");
    }
  }

  private toError(error: unknown): Error {
    return error instanceof Error ? error : new Error(String(error));
  }

  private applyInteractionCursor(target: HTMLElement | SVGElement): void {
    const cursorByMode: Record<InteractionMode, string> = {
      zoom: "crosshair",
      pan: "grab",
      select: "pointer"
    };
    target.style.cursor = cursorByMode[this.interactionMode];
  }

  private createRenderer(kind: "svg" | "canvas" | "webgl" | "hybrid", webglFallback = true): ChartRenderer {
    if (kind === "hybrid") {
      return new HybridRenderer("hybrid");
    }
    if (kind === "webgl" && webglFallback) {
      return new HybridRenderer("webgl");
    }
    if (kind === "canvas") {
      return new CanvasRenderer();
    }
    if (kind === "webgl") {
      return new WebglRenderer();
    }
    return new SvgRenderer();
  }

  private swapRenderer(kind: "svg" | "canvas" | "webgl" | "hybrid", webglFallback = true): void {
    this.renderer.destroy();
    this.renderer = this.createRenderer(kind, webglFallback);
    this.renderer.mount(this.container);
  }

  private applyPointSelection(mode: SelectionShape | "click", points: ChartSelectedPoint[], emitClickSelection: boolean): void {
    const normalizedPoints = deduplicateSelectionPoints(points);
    const selection: ChartSelectionState | null =
      normalizedPoints.length === 0
        ? null
        : {
            mode,
            points: normalizedPoints,
            updatedAt: new Date().toISOString()
          };

    const figure = this.figureManager.updateSelection(selection);
    this.renderCurrentFigure(true);
    this.syncControllers();
    if (emitClickSelection || normalizedPoints.length > 0) {
      normalizedPoints.forEach((point) => this.emit("trace:selected", point));
    }
    this.emit("selection:changed", {
      mode,
      points: normalizedPoints
    });
    this.emit("figure:updated", { figure });
  }
}

const normalizeRange = (rawMin: number, rawMax: number, minSpan: number): [number, number] => {
  const min = Math.min(rawMin, rawMax);
  let max = Math.max(rawMin, rawMax);
  if (max - min < minSpan) {
    max = min + minSpan;
  }
  return [min, max];
};

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

const deduplicateSelectionPoints = (points: ChartSelectedPoint[]): ChartSelectedPoint[] => {
  const seen = new Set<string>();
  const deduplicated: ChartSelectedPoint[] = [];
  for (const point of points) {
    const traceIndex = Number(point.traceIndex);
    const pointIndex = Number(point.pointIndex);
    if (!Number.isFinite(traceIndex) || !Number.isFinite(pointIndex)) {
      continue;
    }
    const key = `${traceIndex}:${pointIndex}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    deduplicated.push({
      traceIndex,
      pointIndex,
      traceName: String(point.traceName ?? `Trace ${traceIndex + 1}`),
      x: String(point.x ?? ""),
      y: String(point.y ?? "")
    });
  }
  return deduplicated;
};

const isCartesianDataTrace = (trace: ChartTrace): trace is CartesianTrace =>
  trace.type === "scatter" || trace.type === "line" || trace.type === "bar" || trace.type === "area";

const isFinancialDataTrace = (trace: ChartTrace): trace is Extract<ChartTrace, { type: "candlestick" | "ohlc" }> =>
  trace.type === "candlestick" || trace.type === "ohlc";

const isPieDataTrace = (trace: ChartTrace): trace is PieTrace | DonutTrace => trace.type === "pie" || trace.type === "donut";

const isValueDataTrace = (trace: ChartTrace): trace is ViolinTrace | DensityTrace =>
  trace.type === "violin" || trace.type === "density" || trace.type === "distribution";

const isFunnelDataTrace = (trace: ChartTrace): trace is FunnelTrace => trace.type === "funnel";

const isSunburstDataTrace = (trace: ChartTrace): trace is SunburstTrace => trace.type === "sunburst";

const isTreemapDataTrace = (trace: ChartTrace): trace is TreemapTrace => trace.type === "treemap";

const createTableCell = (text: string): HTMLTableCellElement => {
  const cell = document.createElement("td");
  cell.textContent = text;
  Object.assign(cell.style, {
    borderBottom: "1px solid #e2e8f0",
    padding: "3px 6px",
    color: "#0f172a"
  });
  return cell;
};

const escapeCsv = (value: string): string => {
  if (/[",\n]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
};

const clampInt = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? Math.round(value) : min;
  return Math.min(max, Math.max(min, numeric));
};

const computeContrastRatio = (background: string, foreground: string): number => {
  const left = toRelativeLuminance(parseHexColor(background));
  const right = toRelativeLuminance(parseHexColor(foreground));
  const lighter = Math.max(left, right);
  const darker = Math.min(left, right);
  return (lighter + 0.05) / (darker + 0.05);
};

const parseHexColor = (input: string): { r: number; g: number; b: number } => {
  const normalized = input.trim().replace("#", "");
  if (/^[0-9a-f]{3}$/i.test(normalized)) {
    const [r, g, b] = normalized.split("").map((char) => parseInt(char + char, 16));
    return { r, g, b };
  }
  if (/^[0-9a-f]{6}$/i.test(normalized)) {
    return {
      r: parseInt(normalized.slice(0, 2), 16),
      g: parseInt(normalized.slice(2, 4), 16),
      b: parseInt(normalized.slice(4, 6), 16)
    };
  }
  return { r: 255, g: 255, b: 255 };
};

const toRelativeLuminance = (color: { r: number; g: number; b: number }): number => {
  const convert = (channel: number): number => {
    const normalized = channel / 255;
    return normalized <= 0.03928 ? normalized / 12.92 : ((normalized + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * convert(color.r) + 0.7152 * convert(color.g) + 0.0722 * convert(color.b);
};
