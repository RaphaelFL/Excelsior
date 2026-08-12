import type { ChartSelectedPoint } from "../model/Figure";

export type SelectionShape = "rect" | "lasso";

interface SelectionControllerOptions {
  onSelection: (points: ChartSelectedPoint[], shape: SelectionShape) => void;
  onCleared?: () => void;
}

type SelectionState = {
  shape: SelectionShape;
  points: Array<{ x: number; y: number }>;
  pointerId: number;
  moved: boolean;
};

export class SelectionController {
  private container: HTMLElement | null = null;
  private target: HTMLElement | SVGElement | null = null;
  private overlaySvg: SVGSVGElement | null = null;
  private overlayRect: SVGRectElement | null = null;
  private overlayPath: SVGPathElement | null = null;
  private selectionState: SelectionState | null = null;
  private enabled = false;
  private readonly boundPointerDown = (event: PointerEvent): void => this.handlePointerDown(event);
  private readonly boundPointerMove = (event: PointerEvent): void => this.handlePointerMove(event);
  private readonly boundPointerUp = (event: PointerEvent): void => this.handlePointerUp(event);
  private readonly boundWindowKeyDown = (event: KeyboardEvent): void => this.handleWindowKeyDown(event);

  constructor(private readonly options: SelectionControllerOptions) {}

  mount(target: Element, container: HTMLElement): void {
    this.unmount();
    this.target = target as HTMLElement | SVGElement;
    this.container = container;
    this.target.addEventListener("pointerdown", this.boundPointerDown as EventListener);
    this.target.addEventListener("pointermove", this.boundPointerMove as EventListener);
    this.target.addEventListener("pointerup", this.boundPointerUp as EventListener);
    this.target.addEventListener("pointercancel", this.boundPointerUp as EventListener);
    window.addEventListener("keydown", this.boundWindowKeyDown);
  }

  setEnabled(enabled: boolean): void {
    this.enabled = enabled;
    if (!enabled) {
      this.clearOverlay();
    }
  }

  clearSelection(): void {
    this.options.onSelection([], "rect");
    this.options.onCleared?.();
  }

  unmount(): void {
    if (this.target) {
      this.target.removeEventListener("pointerdown", this.boundPointerDown as EventListener);
      this.target.removeEventListener("pointermove", this.boundPointerMove as EventListener);
      this.target.removeEventListener("pointerup", this.boundPointerUp as EventListener);
      this.target.removeEventListener("pointercancel", this.boundPointerUp as EventListener);
    }
    window.removeEventListener("keydown", this.boundWindowKeyDown);
    this.target = null;
    this.container = null;
    this.clearOverlay();
  }

  destroy(): void {
    this.unmount();
  }

  private handlePointerDown(event: PointerEvent): void {
    if (!this.enabled || event.button !== 0) {
      return;
    }
    if (!this.container || !this.target) {
      return;
    }

    const containerRect = this.container.getBoundingClientRect();
    const startPoint = {
      x: event.clientX - containerRect.left,
      y: event.clientY - containerRect.top
    };
    const shape: SelectionShape = event.shiftKey ? "lasso" : "rect";
    this.selectionState = {
      shape,
      points: [startPoint],
      pointerId: event.pointerId,
      moved: false
    };
    this.ensureOverlay();
    if (this.target instanceof SVGElement) {
      this.target.setPointerCapture(event.pointerId);
    }
    event.preventDefault();
  }

  private handlePointerMove(event: PointerEvent): void {
    const selectionState = this.selectionState;
    if (!selectionState || !this.container || !this.target || event.pointerId !== selectionState.pointerId) {
      return;
    }
    const containerRect = this.container.getBoundingClientRect();
    const point = {
      x: clamp(event.clientX - containerRect.left, 0, containerRect.width),
      y: clamp(event.clientY - containerRect.top, 0, containerRect.height)
    };

    const lastPoint = selectionState.points[selectionState.points.length - 1];
    const distance = Math.hypot(point.x - lastPoint.x, point.y - lastPoint.y);
    if (distance > 1) {
      selectionState.moved = true;
    }

    if (selectionState.shape === "lasso") {
      if (distance >= 2) {
        selectionState.points.push(point);
      }
    } else {
      selectionState.points[1] = point;
    }
    this.paintOverlay(selectionState);
  }

  private handlePointerUp(event: PointerEvent): void {
    const selectionState = this.selectionState;
    if (!selectionState || !this.target || event.pointerId !== selectionState.pointerId) {
      return;
    }

    const points = selectionState.moved ? this.collectSelection(selectionState) : [];
    this.options.onSelection(points, selectionState.shape);
    if (points.length === 0) {
      this.options.onCleared?.();
    }
    this.clearOverlay();
  }

  private handleWindowKeyDown(event: KeyboardEvent): void {
    if (!this.enabled) {
      return;
    }
    if (event.key === "Escape") {
      this.clearSelection();
      this.clearOverlay();
    }
  }

  private collectSelection(selectionState: SelectionState): ChartSelectedPoint[] {
    if (!this.target) {
      return [];
    }

    const interactiveElements = this.target.querySelectorAll<HTMLElement>("[data-chart-interactive='point']");
    const selected: ChartSelectedPoint[] = [];
    const seen = new Set<string>();

    interactiveElements.forEach((element) => {
      const rect = element.getBoundingClientRect();
      const point = {
        x: rect.left + rect.width / 2,
        y: rect.top + rect.height / 2
      };
      const inside =
        selectionState.shape === "lasso"
          ? pointInsidePolygon(point, selectionState.points, this.container)
          : pointInsideRect(point, selectionState.points, this.container);
      if (!inside) {
        return;
      }

      const traceIndex = Number(element.dataset.traceIndex ?? "");
      const pointIndex = Number(element.dataset.pointIndex ?? "");
      if (!Number.isFinite(traceIndex) || !Number.isFinite(pointIndex)) {
        return;
      }
      const key = `${traceIndex}:${pointIndex}`;
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      selected.push({
        traceIndex,
        pointIndex,
        traceName: String(element.dataset.traceName ?? `Trace ${traceIndex + 1}`),
        x: String(element.dataset.pointX ?? ""),
        y: String(element.dataset.pointY ?? "")
      });
    });

    return selected;
  }

  private ensureOverlay(): void {
    if (!this.container || this.overlaySvg) {
      return;
    }

    const overlay = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    overlay.setAttribute("class", "excelsior-chart-selection-overlay");
    overlay.setAttribute("aria-hidden", "true");
    Object.assign(overlay.style, {
      position: "absolute",
      inset: "0",
      width: "100%",
      height: "100%",
      pointerEvents: "none",
      zIndex: "4"
    });

    const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    rect.setAttribute("fill", "rgba(37, 99, 235, 0.12)");
    rect.setAttribute("stroke", "#2563eb");
    rect.setAttribute("stroke-dasharray", "4 3");
    rect.setAttribute("stroke-width", "1.5");
    rect.setAttribute("visibility", "hidden");
    overlay.append(rect);

    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("fill", "rgba(37, 99, 235, 0.12)");
    path.setAttribute("stroke", "#2563eb");
    path.setAttribute("stroke-dasharray", "4 3");
    path.setAttribute("stroke-width", "1.5");
    path.setAttribute("visibility", "hidden");
    overlay.append(path);

    this.container.append(overlay);
    this.overlaySvg = overlay;
    this.overlayRect = rect;
    this.overlayPath = path;
  }

  private paintOverlay(selectionState: SelectionState): void {
    if (!this.overlayRect || !this.overlayPath) {
      return;
    }

    if (selectionState.shape === "rect") {
      this.overlayPath.setAttribute("visibility", "hidden");
      const [start, end] = [selectionState.points[0], selectionState.points[1] ?? selectionState.points[0]];
      const x = Math.min(start.x, end.x);
      const y = Math.min(start.y, end.y);
      const width = Math.max(1, Math.abs(end.x - start.x));
      const height = Math.max(1, Math.abs(end.y - start.y));
      this.overlayRect.setAttribute("x", String(x));
      this.overlayRect.setAttribute("y", String(y));
      this.overlayRect.setAttribute("width", String(width));
      this.overlayRect.setAttribute("height", String(height));
      this.overlayRect.setAttribute("visibility", "visible");
      return;
    }

    this.overlayRect.setAttribute("visibility", "hidden");
    const pathData = selectionState.points
      .map((point, index) => `${index === 0 ? "M" : "L"} ${point.x.toFixed(1)} ${point.y.toFixed(1)}`)
      .join(" ");
    this.overlayPath.setAttribute("d", `${pathData} Z`);
    this.overlayPath.setAttribute("visibility", "visible");
  }

  private clearOverlay(): void {
    if (this.overlaySvg) {
      this.overlaySvg.remove();
    }
    this.overlaySvg = null;
    this.overlayRect = null;
    this.overlayPath = null;
    this.selectionState = null;
  }
}

const pointInsideRect = (
  viewportPoint: { x: number; y: number },
  selectionPoints: Array<{ x: number; y: number }>,
  container: HTMLElement | null
): boolean => {
  if (!container) {
    return false;
  }
  const [start, end] = [selectionPoints[0], selectionPoints[1] ?? selectionPoints[0]];
  const containerRect = container.getBoundingClientRect();
  const left = containerRect.left + Math.min(start.x, end.x);
  const right = containerRect.left + Math.max(start.x, end.x);
  const top = containerRect.top + Math.min(start.y, end.y);
  const bottom = containerRect.top + Math.max(start.y, end.y);
  return viewportPoint.x >= left && viewportPoint.x <= right && viewportPoint.y >= top && viewportPoint.y <= bottom;
};

const pointInsidePolygon = (
  viewportPoint: { x: number; y: number },
  selectionPoints: Array<{ x: number; y: number }>,
  container: HTMLElement | null
): boolean => {
  if (!container || selectionPoints.length < 3) {
    return false;
  }

  const containerRect = container.getBoundingClientRect();
  const point = {
    x: viewportPoint.x - containerRect.left,
    y: viewportPoint.y - containerRect.top
  };

  let inside = false;
  for (let i = 0, j = selectionPoints.length - 1; i < selectionPoints.length; j = i++) {
    const xi = selectionPoints[i].x;
    const yi = selectionPoints[i].y;
    const xj = selectionPoints[j].x;
    const yj = selectionPoints[j].y;
    const intersects = yi > point.y !== yj > point.y && point.x < ((xj - xi) * (point.y - yi)) / (yj - yi + Number.EPSILON) + xi;
    if (intersects) {
      inside = !inside;
    }
  }
  return inside;
};

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));
