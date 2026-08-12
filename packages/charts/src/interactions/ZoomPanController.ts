export type InteractionMode = "zoom" | "pan" | "select";

interface ZoomPanControllerOptions {
  onZoom: (factor: number, anchorX: number, anchorY: number) => void;
  onPan: (deltaX: number, deltaY: number, anchorX: number, anchorY: number) => void;
  onZoomAxis?: (axis: "x" | "y", factor: number, anchorX: number, anchorY: number) => void;
  onZoomRect?: (bounds: { x0: number; y0: number; x1: number; y1: number }) => void;
}

export class ZoomPanController {
  private target: HTMLElement | SVGElement | null = null;
  private mode: InteractionMode = "zoom";
  private isDragging = false;
  private dragLastX = 0;
  private dragLastY = 0;
  private zoomRectStartX = 0;
  private zoomRectStartY = 0;
  private zoomRectActive = false;

  private readonly boundWheel = (event: WheelEvent): void => this.handleWheel(event);
  private readonly boundPointerDown = (event: PointerEvent): void => this.handlePointerDown(event);
  private readonly boundPointerMove = (event: PointerEvent): void => this.handlePointerMove(event);
  private readonly boundPointerUp = (): void => this.handlePointerUp();

  constructor(private readonly options: ZoomPanControllerOptions) {}

  setMode(mode: InteractionMode): void {
    this.mode = mode;
  }

  mount(target: Element): void {
    this.unmount();
    this.target = target as HTMLElement | SVGElement;
    this.target.addEventListener("wheel", this.boundWheel as EventListener, { passive: false });
    this.target.addEventListener("pointerdown", this.boundPointerDown as EventListener);
    this.target.addEventListener("pointermove", this.boundPointerMove as EventListener);
    this.target.addEventListener("pointerup", this.boundPointerUp as EventListener);
    this.target.addEventListener("pointercancel", this.boundPointerUp as EventListener);
    this.target.addEventListener("pointerleave", this.boundPointerUp as EventListener);
  }

  unmount(): void {
    if (!this.target) {
      return;
    }
    this.target.removeEventListener("wheel", this.boundWheel as EventListener);
    this.target.removeEventListener("pointerdown", this.boundPointerDown as EventListener);
    this.target.removeEventListener("pointermove", this.boundPointerMove as EventListener);
    this.target.removeEventListener("pointerup", this.boundPointerUp as EventListener);
    this.target.removeEventListener("pointercancel", this.boundPointerUp as EventListener);
    this.target.removeEventListener("pointerleave", this.boundPointerUp as EventListener);
    this.isDragging = false;
    this.zoomRectActive = false;
    this.target = null;
  }

  destroy(): void {
    this.unmount();
  }

  private handleWheel(event: WheelEvent): void {
    event.preventDefault();
    const factor = event.deltaY > 0 ? 1.18 : 0.84;
    if (event.shiftKey) {
      this.options.onZoomAxis?.("x", factor, event.clientX, event.clientY);
      return;
    }
    if (event.altKey) {
      this.options.onZoomAxis?.("y", factor, event.clientX, event.clientY);
      return;
    }
    this.options.onZoom(factor, event.clientX, event.clientY);
  }

  private handlePointerDown(event: PointerEvent): void {
    if (this.mode === "zoom") {
      this.zoomRectStartX = event.clientX;
      this.zoomRectStartY = event.clientY;
      this.zoomRectActive = true;
      if (this.target instanceof SVGElement && "setPointerCapture" in this.target) {
        (this.target as SVGElement & { setPointerCapture?: (pointerId: number) => void }).setPointerCapture?.(event.pointerId);
      }
      return;
    }
    if (this.mode === "pan") {
      this.isDragging = true;
      this.dragLastX = event.clientX;
      this.dragLastY = event.clientY;
      if (this.target instanceof SVGElement && "setPointerCapture" in this.target) {
        (this.target as SVGElement & { setPointerCapture?: (pointerId: number) => void }).setPointerCapture?.(event.pointerId);
      }
    }
  }

  private handlePointerMove(event: PointerEvent): void {
    if (this.zoomRectActive && this.mode === "zoom") {
      this.dragLastX = event.clientX;
      this.dragLastY = event.clientY;
      return;
    }
    if (!this.isDragging || this.mode !== "pan") {
      return;
    }
    const deltaX = event.clientX - this.dragLastX;
    const deltaY = event.clientY - this.dragLastY;
    this.dragLastX = event.clientX;
    this.dragLastY = event.clientY;
    this.options.onPan(deltaX, deltaY, event.clientX, event.clientY);
  }

  private handlePointerUp(): void {
    if (this.zoomRectActive && this.mode === "zoom") {
      const x0 = this.zoomRectStartX;
      const y0 = this.zoomRectStartY;
      const x1 = this.dragLastX || this.zoomRectStartX;
      const y1 = this.dragLastY || this.zoomRectStartY;
      const width = Math.abs(x1 - x0);
      const height = Math.abs(y1 - y0);
      if (width >= 6 && height >= 6) {
        this.options.onZoomRect?.({ x0, y0, x1, y1 });
      }
      this.zoomRectActive = false;
    }
    this.isDragging = false;
  }
}
