import type { ChartRenderer } from "./ChartRenderer";
import type { ChartFigure } from "../model/Figure";
import type { ComputedLayout } from "../core/LayoutEngine";
import type { ChartRendererKind } from "../model/Config";
import { SvgRenderer } from "./SvgRenderer";
import { CanvasRenderer } from "./CanvasRenderer";
import { WebglRenderer } from "./WebglRenderer";
import { ChartConfigurationError } from "../core/chart-errors";

type RuntimeRendererKind = Exclude<ChartRendererKind, "hybrid">;

export class HybridRenderer implements ChartRenderer {
  private container: HTMLElement | null = null;
  private activeRenderer: ChartRenderer | null = null;
  private activeIndex = -1;
  private lastFigure: ChartFigure | null = null;
  private lastLayout: ComputedLayout | null = null;

  constructor(private readonly preferred: ChartRendererKind = "hybrid") {}

  mount(container: HTMLElement): void {
    this.destroyActiveRenderer();
    this.container = container;
    this.activateRenderer(0);
  }

  render(figure: ChartFigure, layout: ComputedLayout): void {
    this.lastFigure = structuredClone(figure);
    this.lastLayout = layout;
    const renderer = this.ensureActiveRenderer();
    try {
      renderer.render(figure, layout);
    } catch (error) {
      this.fallback(error);
      this.ensureActiveRenderer().render(figure, layout);
    }
  }

  resize(layout: ComputedLayout): void {
    this.lastLayout = layout;
    const renderer = this.ensureActiveRenderer();
    try {
      renderer.resize(layout);
    } catch (error) {
      this.fallback(error);
      this.ensureActiveRenderer().resize(layout);
    }
  }

  getRootElement(): Element | null {
    return this.activeRenderer?.getRootElement() ?? null;
  }

  exportSvg(): string {
    try {
      return this.ensureActiveRenderer().exportSvg();
    } catch (error) {
      this.fallback(error);
      return this.ensureActiveRenderer().exportSvg();
    }
  }

  async exportPng(options?: { scale?: number; backgroundColor?: string }): Promise<Blob> {
    try {
      return await this.ensureActiveRenderer().exportPng(options);
    } catch (error) {
      this.fallback(error);
      return await this.ensureActiveRenderer().exportPng(options);
    }
  }

  destroy(): void {
    this.destroyActiveRenderer();
    this.container = null;
    this.lastFigure = null;
    this.lastLayout = null;
  }

  private activateRenderer(startIndex: number): void {
    const container = this.container;
    if (!container) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_MOUNTED", "Hybrid renderer requires mount() before use.");
    }
    const kinds = this.resolveRendererChain();
    let lastError: unknown = null;
    for (let index = startIndex; index < kinds.length; index += 1) {
      const kind = kinds[index];
      const candidate = this.createRenderer(kind);
      try {
        candidate.mount(container);
        this.activeRenderer = candidate;
        this.activeIndex = index;
        if (this.lastFigure && this.lastLayout) {
          candidate.resize(this.lastLayout);
          candidate.render(this.lastFigure, this.lastLayout);
        }
        return;
      } catch (error) {
        candidate.destroy();
        lastError = error;
      }
    }
    throw new ChartConfigurationError(
      "CHART_RENDERER_NOT_AVAILABLE",
      `Unable to initialize a renderer from chain ${kinds.join(" -> ")}. Last error: ${toMessage(lastError)}`
    );
  }

  private ensureActiveRenderer(): ChartRenderer {
    if (!this.activeRenderer) {
      this.activateRenderer(0);
    }
    if (!this.activeRenderer) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_AVAILABLE", "No renderer available.");
    }
    return this.activeRenderer;
  }

  private fallback(error: unknown): void {
    const nextIndex = this.activeIndex + 1;
    const chain = this.resolveRendererChain();
    if (nextIndex >= chain.length) {
      throw new ChartConfigurationError("CHART_RENDER_FAILED", `Renderer failed and no fallback remains: ${toMessage(error)}`);
    }
    this.destroyActiveRenderer();
    this.activateRenderer(nextIndex);
  }

  private destroyActiveRenderer(): void {
    this.activeRenderer?.destroy();
    this.activeRenderer = null;
    this.activeIndex = -1;
  }

  private resolveRendererChain(): RuntimeRendererKind[] {
    if (this.preferred === "svg") {
      return ["svg"];
    }
    if (this.preferred === "canvas") {
      return ["canvas", "svg"];
    }
    if (this.preferred === "webgl") {
      return ["webgl", "canvas", "svg"];
    }
    return ["webgl", "canvas", "svg"];
  }

  private createRenderer(kind: RuntimeRendererKind): ChartRenderer {
    if (kind === "webgl") {
      return new WebglRenderer();
    }
    if (kind === "canvas") {
      return new CanvasRenderer();
    }
    return new SvgRenderer();
  }
}

const toMessage = (error: unknown): string => (error instanceof Error ? error.message : String(error));
