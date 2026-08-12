import type { TooltipController } from "./TooltipController";

type PointInfo = {
  traceIndex: number;
  pointIndex: number;
  traceName: string;
  x: string;
  y: string;
};

interface HoverControllerOptions {
  tooltip: TooltipController;
  onHover?: (info: PointInfo) => void;
  onUnhover?: (info: PointInfo) => void;
  onClick?: (info: PointInfo) => void;
}

interface SpatialPoint {
  info: PointInfo;
  centerX: number;
  centerY: number;
  element: HTMLElement;
}

export class HoverController {
  private target: HTMLElement | SVGElement | null = null;
  private hoverState: PointInfo | null = null;
  private hoverMode: "point" | "x" = "point";
  private spatialHoverEnabled = true;
  private spatialPoints: SpatialPoint[] = [];
  private spatialGrid = new Map<string, number[]>();
  private spatialSignature = "";
  private spatialCellSize = 28;
  private boundPointerMove = (event: PointerEvent): void => this.handlePointerMove(event);
  private boundPointerDown = (event: PointerEvent): void => this.handlePointerDown(event);
  private boundPointerUp = (event: PointerEvent): void => this.handlePointerUp(event);
  private boundPointerLeave = (): void => this.handlePointerLeave();
  private boundClick = (event: MouseEvent): void => this.handleClick(event);

  constructor(private readonly options: HoverControllerOptions) {}

  mount(target: Element): void {
    this.unmount();
    this.target = target as HTMLElement | SVGElement;
    this.target.addEventListener("pointermove", this.boundPointerMove as EventListener);
    this.target.addEventListener("pointerdown", this.boundPointerDown as EventListener);
    this.target.addEventListener("pointerup", this.boundPointerUp as EventListener);
    this.target.addEventListener("pointerleave", this.boundPointerLeave as EventListener);
    this.target.addEventListener("click", this.boundClick as EventListener);
  }

  setHoverMode(mode: "point" | "x"): void {
    this.hoverMode = mode;
  }

  setSpatialHoverEnabled(enabled: boolean): void {
    this.spatialHoverEnabled = enabled;
    this.invalidateSpatialIndex();
  }

  unmount(): void {
    if (!this.target) {
      return;
    }
    this.target.removeEventListener("pointermove", this.boundPointerMove as EventListener);
    this.target.removeEventListener("pointerdown", this.boundPointerDown as EventListener);
    this.target.removeEventListener("pointerup", this.boundPointerUp as EventListener);
    this.target.removeEventListener("pointerleave", this.boundPointerLeave as EventListener);
    this.target.removeEventListener("click", this.boundClick as EventListener);
    this.target = null;
    this.invalidateSpatialIndex();
  }

  destroy(): void {
    this.unmount();
    this.options.tooltip.hide();
    this.hoverState = null;
    this.invalidateSpatialIndex();
  }

  private handlePointerMove(event: PointerEvent): void {
    const target = event.target;
    if (!(target instanceof Element)) {
      this.clearHover();
      return;
    }

    const interactive = target.closest<HTMLElement>("[data-chart-interactive='point']") ?? this.findNearestInteractive(event.clientX, event.clientY);
    if (!interactive) {
      this.clearHover();
      return;
    }

    const pointInfo = this.readPointInfo(interactive);
    if (!pointInfo) {
      this.clearHover();
      return;
    }

    const hoverChanged =
      !this.hoverState ||
      this.hoverState.traceIndex !== pointInfo.traceIndex ||
      this.hoverState.pointIndex !== pointInfo.pointIndex;

    if (hoverChanged) {
      if (this.hoverState) {
        this.options.onUnhover?.(this.hoverState);
      }
      this.options.onHover?.(pointInfo);
      this.hoverState = pointInfo;
    }

    const tooltipLines = this.hoverMode === "x" ? this.buildComparativeLines(pointInfo.pointIndex) : [`x: ${pointInfo.x}`, `y: ${pointInfo.y}`];
    this.options.tooltip.show({
      title: this.hoverMode === "x" ? `x: ${pointInfo.x}` : pointInfo.traceName,
      lines: tooltipLines,
      clientX: event.clientX,
      clientY: event.clientY
    });
  }

  private handlePointerDown(event: PointerEvent): void {
    if (event.pointerType !== "touch") {
      return;
    }
    this.handlePointerMove(event);
  }

  private handlePointerUp(event: PointerEvent): void {
    if (event.pointerType === "touch") {
      window.setTimeout(() => this.clearHover(), 1200);
    }
  }

  private handlePointerLeave(): void {
    this.clearHover();
  }

  private handleClick(event: MouseEvent): void {
    const target = event.target;
    if (!(target instanceof Element)) {
      return;
    }
    const interactive = target.closest<HTMLElement>("[data-chart-interactive='point']");
    if (!interactive) {
      return;
    }
    const pointInfo = this.readPointInfo(interactive);
    if (!pointInfo) {
      return;
    }
    this.options.onClick?.(pointInfo);
  }

  private buildComparativeLines(pointIndex: number): string[] {
    if (!this.target) {
      return [];
    }
    const points = this.target.querySelectorAll<HTMLElement>(`[data-chart-interactive='point'][data-point-index='${pointIndex}']`);
    if (points.length === 0) {
      return [];
    }

    const byTrace = new Map<number, string>();
    points.forEach((point) => {
      const traceIndex = Number(point.dataset.traceIndex ?? "");
      if (!Number.isFinite(traceIndex) || byTrace.has(traceIndex)) {
        return;
      }
      const traceName = String(point.dataset.traceName ?? `Trace ${traceIndex + 1}`);
      const y = String(point.dataset.pointY ?? "");
      byTrace.set(traceIndex, `${traceName}: ${y}`);
    });

    return Array.from(byTrace.entries())
      .sort((a, b) => a[0] - b[0])
      .map((entry) => entry[1]);
  }

  private findNearestInteractive(clientX: number, clientY: number): HTMLElement | null {
    if (!this.target || !this.spatialHoverEnabled) {
      return null;
    }

    this.ensureSpatialIndex();
    if (this.spatialPoints.length === 0) {
      return null;
    }

    const column = Math.floor(clientX / this.spatialCellSize);
    const row = Math.floor(clientY / this.spatialCellSize);
    let nearest: SpatialPoint | null = null;
    let nearestDistance = Number.POSITIVE_INFINITY;
    const radiusCells = 2;
    const maxDistance = this.spatialCellSize * 1.8;

    for (let dy = -radiusCells; dy <= radiusCells; dy += 1) {
      for (let dx = -radiusCells; dx <= radiusCells; dx += 1) {
        const bucketKey = `${column + dx}:${row + dy}`;
        const indices = this.spatialGrid.get(bucketKey);
        if (!indices) {
          continue;
        }
        for (const pointIndex of indices) {
          const candidate = this.spatialPoints[pointIndex];
          const distance = Math.hypot(candidate.centerX - clientX, candidate.centerY - clientY);
          if (distance > maxDistance || distance >= nearestDistance) {
            continue;
          }
          nearest = candidate;
          nearestDistance = distance;
        }
      }
    }

    return nearest?.element ?? null;
  }

  private ensureSpatialIndex(): void {
    if (!this.target) {
      this.invalidateSpatialIndex();
      return;
    }

    const elements = this.target.querySelectorAll<HTMLElement>("[data-chart-interactive='point']");
    const signature = `${elements.length}:${this.target.getBoundingClientRect().width.toFixed(2)}:${this.target.getBoundingClientRect().height.toFixed(2)}`;
    if (signature === this.spatialSignature) {
      return;
    }
    this.spatialSignature = signature;
    this.spatialPoints = [];
    this.spatialGrid.clear();

    elements.forEach((element) => {
      const info = this.readPointInfo(element);
      if (!info) {
        return;
      }
      const rect = element.getBoundingClientRect();
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      const point: SpatialPoint = { info, centerX, centerY, element };
      this.spatialPoints.push(point);
      const pointIndex = this.spatialPoints.length - 1;
      const col = Math.floor(centerX / this.spatialCellSize);
      const row = Math.floor(centerY / this.spatialCellSize);
      const key = `${col}:${row}`;
      const bucket = this.spatialGrid.get(key);
      if (bucket) {
        bucket.push(pointIndex);
      } else {
        this.spatialGrid.set(key, [pointIndex]);
      }
    });
  }

  private invalidateSpatialIndex(): void {
    this.spatialSignature = "";
    this.spatialPoints = [];
    this.spatialGrid.clear();
  }

  private clearHover(): void {
    if (this.hoverState) {
      this.options.onUnhover?.(this.hoverState);
      this.hoverState = null;
    }
    this.options.tooltip.hide();
  }

  private readPointInfo(element: HTMLElement): PointInfo | null {
    const traceIndex = Number(element.dataset.traceIndex ?? "");
    const pointIndex = Number(element.dataset.pointIndex ?? "");
    const traceName = String(element.dataset.traceName ?? "Trace");
    const x = String(element.dataset.pointX ?? "");
    const y = String(element.dataset.pointY ?? "");

    if (!Number.isFinite(traceIndex) || !Number.isFinite(pointIndex)) {
      return null;
    }

    return {
      traceIndex,
      pointIndex,
      traceName,
      x,
      y
    };
  }
}
