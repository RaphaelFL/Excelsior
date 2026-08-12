export class LegendController {
  private target: HTMLElement | SVGElement | null = null;
  private readonly boundClick = (event: MouseEvent): void => this.handleClick(event);
  private readonly boundKeyDown = (event: KeyboardEvent): void => this.handleKeyDown(event);

  constructor(private readonly onToggle: (traceIndex: number) => void) {}

  mount(target: Element): void {
    this.unmount();
    this.target = target as HTMLElement | SVGElement;
    this.target.addEventListener("click", this.boundClick as EventListener);
    this.target.addEventListener("keydown", this.boundKeyDown as EventListener);
  }

  unmount(): void {
    if (!this.target) {
      return;
    }
    this.target.removeEventListener("click", this.boundClick as EventListener);
    this.target.removeEventListener("keydown", this.boundKeyDown as EventListener);
    this.target = null;
  }

  destroy(): void {
    this.unmount();
  }

  private handleClick(event: MouseEvent): void {
    const target = event.target;
    if (!(target instanceof Element)) {
      return;
    }

    const legendItem = target.closest<HTMLElement>("[data-chart-legend='item']");
    if (!legendItem) {
      return;
    }

    const traceIndex = Number(legendItem.dataset.legendTraceIndex ?? "");
    if (!Number.isFinite(traceIndex)) {
      return;
    }

    event.preventDefault();
    this.onToggle(traceIndex);
  }

  private handleKeyDown(event: KeyboardEvent): void {
    if (event.key !== "Enter" && event.key !== " ") {
      return;
    }

    const target = event.target;
    if (!(target instanceof Element)) {
      return;
    }
    const legendItem = target.closest<HTMLElement>("[data-chart-legend='item']");
    if (!legendItem) {
      return;
    }
    const traceIndex = Number(legendItem.dataset.legendTraceIndex ?? "");
    if (!Number.isFinite(traceIndex)) {
      return;
    }

    event.preventDefault();
    this.onToggle(traceIndex);
  }
}
