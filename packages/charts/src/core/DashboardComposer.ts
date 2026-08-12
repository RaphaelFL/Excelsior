import type { ChartFigureInput } from "../model/Figure";
import { ChartEngine } from "./ChartEngine";

export interface DashboardWidget {
  id: string;
  figure: ChartFigureInput;
  row?: number;
  col?: number;
  rowSpan?: number;
  colSpan?: number;
  className?: string;
}

export interface DashboardComposerOptions {
  columns?: number;
  gap?: number;
  autoRowHeight?: number;
}

export class DashboardComposer {
  private readonly engines = new Map<string, ChartEngine>();
  private readonly cells = new Map<string, HTMLDivElement>();
  private readonly options: Required<DashboardComposerOptions>;

  constructor(private readonly container: HTMLElement, options?: DashboardComposerOptions) {
    this.options = {
      columns: clampInt(options?.columns ?? 12, 1, 48),
      gap: clampNumber(options?.gap ?? 10, 0, 40),
      autoRowHeight: clampInt(options?.autoRowHeight ?? 220, 80, 1400)
    };
    this.configureContainer();
  }

  setWidgets(widgets: DashboardWidget[]): void {
    const nextIds = new Set(widgets.map((widget) => widget.id));
    for (const [id, engine] of this.engines.entries()) {
      if (nextIds.has(id)) {
        continue;
      }
      engine.destroy();
      this.engines.delete(id);
      this.cells.get(id)?.remove();
      this.cells.delete(id);
    }

    widgets.forEach((widget) => {
      const existing = this.engines.get(widget.id);
      if (existing) {
        existing.update(widget.figure);
        this.applyCellLayout(widget.id, widget);
      } else {
        const cell = this.createCell(widget);
        const engine = new ChartEngine(cell, widget.figure);
        this.engines.set(widget.id, engine);
        this.cells.set(widget.id, cell);
      }
    });
  }

  updateWidget(id: string, figure: ChartFigureInput): void {
    const engine = this.engines.get(id);
    if (!engine) {
      throw new Error(`Dashboard widget '${id}' not found.`);
    }
    engine.update(figure);
  }

  removeWidget(id: string): void {
    this.engines.get(id)?.destroy();
    this.engines.delete(id);
    this.cells.get(id)?.remove();
    this.cells.delete(id);
  }

  destroy(): void {
    for (const engine of this.engines.values()) {
      engine.destroy();
    }
    this.engines.clear();
    this.cells.forEach((cell) => cell.remove());
    this.cells.clear();
  }

  private configureContainer(): void {
    Object.assign(this.container.style, {
      display: "grid",
      gridTemplateColumns: `repeat(${this.options.columns}, minmax(0, 1fr))`,
      gridAutoRows: `${this.options.autoRowHeight}px`,
      gap: `${this.options.gap}px`
    });
  }

  private createCell(widget: DashboardWidget): HTMLDivElement {
    const cell = document.createElement("div");
    cell.dataset.dashboardWidgetId = widget.id;
    cell.className = widget.className ?? "excelsior-chart-dashboard-widget";
    this.container.append(cell);
    this.applyCellLayout(widget.id, widget, cell);
    return cell;
  }

  private applyCellLayout(widgetId: string, widget: DashboardWidget, providedCell?: HTMLDivElement): void {
    const cell = providedCell ?? this.cells.get(widgetId);
    if (!cell) {
      return;
    }
    const row = clampInt(widget.row ?? 1, 1, 500);
    const col = clampInt(widget.col ?? 1, 1, this.options.columns);
    const rowSpan = clampInt(widget.rowSpan ?? 1, 1, 500);
    const colSpan = clampInt(widget.colSpan ?? 3, 1, this.options.columns);
    Object.assign(cell.style, {
      position: "relative",
      minHeight: "0",
      minWidth: "0",
      gridRow: `${row} / span ${rowSpan}`,
      gridColumn: `${col} / span ${colSpan}`
    });
  }
}

const clampInt = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? Math.round(value) : min;
  return Math.min(max, Math.max(min, numeric));
};

const clampNumber = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? value : min;
  return Math.min(max, Math.max(min, numeric));
};
