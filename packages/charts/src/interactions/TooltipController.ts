export interface TooltipPayload {
  title: string;
  lines: string[];
  clientX: number;
  clientY: number;
}

export class TooltipController {
  private tooltipElement: HTMLDivElement | null = null;

  mount(container: HTMLElement): void {
    if (this.tooltipElement) {
      this.tooltipElement.remove();
    }

    const tooltip = document.createElement("div");
    tooltip.className = "excelsior-chart-tooltip";
    Object.assign(tooltip.style, {
      position: "fixed",
      pointerEvents: "none",
      zIndex: "999999",
      background: "rgba(15, 23, 42, 0.95)",
      color: "#f8fafc",
      borderRadius: "6px",
      padding: "6px 8px",
      fontSize: "12px",
      lineHeight: "1.3",
      fontFamily: "Arial, sans-serif",
      boxShadow: "0 6px 16px rgba(15, 23, 42, 0.28)",
      maxWidth: "280px",
      opacity: "0",
      transform: "translate(-9999px, -9999px)",
      transition: "opacity 80ms ease-out"
    });
    container.append(tooltip);
    this.tooltipElement = tooltip;
  }

  show(payload: TooltipPayload): void {
    const tooltip = this.tooltipElement;
    if (!tooltip) {
      return;
    }

    tooltip.replaceChildren();
    const title = document.createElement("div");
    title.textContent = payload.title;
    title.style.fontWeight = "700";
    tooltip.append(title);

    for (const lineText of payload.lines) {
      const line = document.createElement("div");
      line.textContent = lineText;
      tooltip.append(line);
    }

    const offsetX = 12;
    const offsetY = 14;
    const viewportWidth = typeof window !== "undefined" ? window.innerWidth : payload.clientX + offsetX + 200;
    const viewportHeight = typeof window !== "undefined" ? window.innerHeight : payload.clientY + offsetY + 120;
    const tooltipWidth = Math.max(32, tooltip.offsetWidth || 160);
    const tooltipHeight = Math.max(24, tooltip.offsetHeight || 70);
    const left = clamp(payload.clientX + offsetX, 8, Math.max(8, viewportWidth - tooltipWidth - 8));
    const top = clamp(payload.clientY + offsetY, 8, Math.max(8, viewportHeight - tooltipHeight - 8));
    tooltip.style.transform = `translate(${left}px, ${top}px)`;
    tooltip.style.opacity = "1";
  }

  hide(): void {
    const tooltip = this.tooltipElement;
    if (!tooltip) {
      return;
    }
    tooltip.style.opacity = "0";
    tooltip.style.transform = "translate(-9999px, -9999px)";
  }

  destroy(): void {
    if (this.tooltipElement) {
      this.tooltipElement.remove();
    }
    this.tooltipElement = null;
  }
}

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));
