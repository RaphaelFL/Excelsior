import type { ChartFigure } from "../model/Figure";
import type { ComputedLayout } from "../core/LayoutEngine";

export interface ChartRenderer {
  mount(container: HTMLElement): void;
  render(figure: ChartFigure, layout: ComputedLayout): void;
  resize(layout: ComputedLayout): void;
  getRootElement(): Element | null;
  exportSvg(): string;
  exportPng(options?: { scale?: number; backgroundColor?: string }): Promise<Blob>;
  destroy(): void;
}
