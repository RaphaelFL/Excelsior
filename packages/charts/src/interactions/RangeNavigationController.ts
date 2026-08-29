import type { ChartAxis } from "../model/Layout";

interface RangeNavigationControllerOptions {
  onRangeChanged: (startFraction: number, endFraction: number) => void;
}

export class RangeNavigationController {
  private element: HTMLDivElement | null = null;

  constructor(private readonly options: RangeNavigationControllerOptions) {}

  mount(container: HTMLElement, axis: ChartAxis): void {
    this.destroy();
    if (!axis.rangeSlider?.visible && !axis.rangeSelector?.visible) {
      return;
    }

    const element = document.createElement("div");
    element.className = "excelsior-chart-range-navigation";
    Object.assign(element.style, {
      position: "absolute",
      left: "56px",
      right: "24px",
      bottom: "6px",
      display: "flex",
      alignItems: "center",
      gap: "6px",
      zIndex: "4"
    });

    if (axis.rangeSelector?.visible) {
      const buttons = axis.rangeSelector.buttons?.length
        ? axis.rangeSelector.buttons.slice(0, 8)
        : [{ label: "25%", fraction: 0.25 }, { label: "50%", fraction: 0.5 }, { label: "Tudo", fraction: 1 }];
      for (const definition of buttons) {
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = definition.label.slice(0, 24);
        button.dataset.rangeFraction = String(definition.fraction ?? 1);
        button.setAttribute("aria-label", `Exibir ${definition.label}`);
        Object.assign(button.style, { minHeight: "24px", padding: "2px 7px" });
        button.addEventListener("click", () => {
          const fraction = Math.min(1, Math.max(0.01, definition.fraction ?? 1));
          this.options.onRangeChanged(1 - fraction, 1);
        });
        element.append(button);
      }
    }

    if (axis.rangeSlider?.visible) {
      const start = this.createSlider("Início da faixa", axis.rangeSlider.start ?? 0);
      const end = this.createSlider("Fim da faixa", axis.rangeSlider.end ?? 1);
      const apply = (): void => {
        const startFraction = Math.min(start.valueAsNumber, end.valueAsNumber - 0.01);
        const endFraction = Math.max(end.valueAsNumber, start.valueAsNumber + 0.01);
        start.value = String(Math.max(0, startFraction));
        end.value = String(Math.min(1, endFraction));
        this.options.onRangeChanged(Number(start.value), Number(end.value));
      };
      start.addEventListener("change", apply);
      end.addEventListener("change", apply);
      element.append(start, end);
    }

    container.style.position = container.style.position || "relative";
    container.append(element);
    this.element = element;
  }

  destroy(): void {
    this.element?.remove();
    this.element = null;
  }

  private createSlider(label: string, value: number): HTMLInputElement {
    const input = document.createElement("input");
    input.type = "range";
    input.min = "0";
    input.max = "1";
    input.step = "0.01";
    input.value = String(Math.min(1, Math.max(0, value)));
    input.setAttribute("aria-label", label);
    input.style.minWidth = "80px";
    input.style.flex = "1";
    return input;
  }
}