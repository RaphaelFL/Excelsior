import type { InteractionMode } from "./ZoomPanController";

interface ModebarControllerOptions {
  onZoomIn: () => void;
  onZoomOut: () => void;
  onResetAxis: () => void;
  onSetMode: (mode: InteractionMode) => void;
  onToggleLegend: () => void;
  onToggleFullscreen: () => void;
  onExportSvg: () => void;
  onExportPng: () => void;
  allowFullscreen: boolean;
}

export class ModebarController {
  private element: HTMLDivElement | null = null;
  private modeButtons = new Map<InteractionMode, HTMLButtonElement>();
  private mode: InteractionMode = "zoom";

  constructor(private readonly options: ModebarControllerOptions) {}

  mount(container: HTMLElement): void {
    this.destroy();
    container.style.position = container.style.position || "relative";

    const element = document.createElement("div");
    element.className = "excelsior-chart-modebar";
    Object.assign(element.style, {
      position: "absolute",
      top: "8px",
      right: "8px",
      display: "flex",
      gap: "4px",
      padding: "4px",
      borderRadius: "8px",
      background: "rgba(255, 255, 255, 0.94)",
      border: "1px solid rgba(148, 163, 184, 0.45)",
      boxShadow: "0 2px 8px rgba(15, 23, 42, 0.12)",
      zIndex: "5"
    });

    element.append(
      this.createButton("zoom-in", "Aproximar", () => this.options.onZoomIn()),
      this.createButton("zoom-out", "Afastar", () => this.options.onZoomOut()),
      this.createModeButton("zoom", "Modo zoom", "zoom"),
      this.createModeButton("pan", "Modo arrastar", "pan"),
      this.createModeButton("select", "Modo selecionar", "select"),
      this.createButton("reset", "Resetar eixos", () => this.options.onResetAxis()),
      this.createButton("legend", "Exibir/ocultar legenda", () => this.options.onToggleLegend()),
      ...(this.options.allowFullscreen ? [this.createButton("fullscreen", "Tela cheia", () => this.options.onToggleFullscreen())] : []),
      this.createButton("export-svg", "Exportar SVG", () => this.options.onExportSvg()),
      this.createButton("export-png", "Exportar PNG", () => this.options.onExportPng())
    );

    container.append(element);
    this.element = element;
    this.setMode(this.mode);
  }

  setMode(mode: InteractionMode): void {
    this.mode = mode;
    for (const [buttonMode, button] of this.modeButtons.entries()) {
      const active = buttonMode === mode;
      button.style.background = active ? "#1d4ed8" : "transparent";
      button.style.color = active ? "#ffffff" : "#0f172a";
      button.style.borderColor = active ? "#1d4ed8" : "#94a3b8";
    }
    this.options.onSetMode(mode);
  }

  destroy(): void {
    this.modeButtons.clear();
    if (this.element) {
      this.element.remove();
    }
    this.element = null;
  }

  private createModeButton(icon: ModebarIconName, title: string, mode: InteractionMode): HTMLButtonElement {
    const button = this.createButton(icon, title, () => this.setMode(mode));
    this.modeButtons.set(mode, button);
    return button;
  }

  private createButton(icon: ModebarIconName, title: string, onClick: () => void): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("aria-label", title);
    button.title = title;
    Object.assign(button.style, {
      height: "26px",
      minWidth: "26px",
      padding: "0 6px",
      borderRadius: "6px",
      border: "1px solid #94a3b8",
      background: "transparent",
      color: "#0f172a",
      fontSize: "11px",
      fontWeight: "600",
      cursor: "pointer",
      display: "inline-flex",
      alignItems: "center",
      justifyContent: "center"
    });
    button.append(createIcon(icon));
    button.addEventListener("click", (event) => {
      event.preventDefault();
      onClick();
    });
    return button;
  }
}

type ModebarIconName =
  | "zoom-in"
  | "zoom-out"
  | "zoom"
  | "pan"
  | "select"
  | "reset"
  | "legend"
  | "fullscreen"
  | "export-svg"
  | "export-png";

const createIcon = (name: ModebarIconName): SVGSVGElement => {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("width", "14");
  svg.setAttribute("height", "14");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.9");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");
  svg.setAttribute("aria-hidden", "true");
  svg.style.pointerEvents = "none";

  const commands = ICON_PATHS[name] ?? [];
  for (const command of commands) {
    const element = document.createElementNS("http://www.w3.org/2000/svg", command.tag);
    for (const [key, value] of Object.entries(command.attrs)) {
      element.setAttribute(key, value);
    }
    svg.append(element);
  }
  return svg;
};

const ICON_PATHS: Record<ModebarIconName, Array<{ tag: "path" | "line" | "rect" | "circle"; attrs: Record<string, string> }>> = {
  "zoom-in": [
    { tag: "circle", attrs: { cx: "10", cy: "10", r: "6" } },
    { tag: "line", attrs: { x1: "10", y1: "7", x2: "10", y2: "13" } },
    { tag: "line", attrs: { x1: "7", y1: "10", x2: "13", y2: "10" } },
    { tag: "line", attrs: { x1: "15", y1: "15", x2: "21", y2: "21" } }
  ],
  "zoom-out": [
    { tag: "circle", attrs: { cx: "10", cy: "10", r: "6" } },
    { tag: "line", attrs: { x1: "7", y1: "10", x2: "13", y2: "10" } },
    { tag: "line", attrs: { x1: "15", y1: "15", x2: "21", y2: "21" } }
  ],
  zoom: [
    { tag: "rect", attrs: { x: "4", y: "4", width: "16", height: "16", rx: "2" } },
    { tag: "line", attrs: { x1: "12", y1: "7", x2: "12", y2: "17" } },
    { tag: "line", attrs: { x1: "7", y1: "12", x2: "17", y2: "12" } }
  ],
  pan: [
    { tag: "path", attrs: { d: "M12 3l2.5 2.5L12 8 9.5 5.5 12 3z" } },
    { tag: "path", attrs: { d: "M21 12l-2.5 2.5L16 12l2.5-2.5L21 12z" } },
    { tag: "path", attrs: { d: "M12 21l-2.5-2.5L12 16l2.5 2.5L12 21z" } },
    { tag: "path", attrs: { d: "M3 12l2.5-2.5L8 12l-2.5 2.5L3 12z" } }
  ],
  select: [
    { tag: "rect", attrs: { x: "4", y: "4", width: "16", height: "16", rx: "2", "stroke-dasharray": "2 2" } },
    { tag: "path", attrs: { d: "M9 9l6 3-3 1-1 3-2-7z" } }
  ],
  reset: [
    { tag: "path", attrs: { d: "M4 12a8 8 0 1 0 2.3-5.7" } },
    { tag: "line", attrs: { x1: "4", y1: "5", x2: "4", y2: "10" } },
    { tag: "line", attrs: { x1: "4", y1: "5", x2: "9", y2: "5" } }
  ],
  legend: [
    { tag: "rect", attrs: { x: "4", y: "6", width: "5", height: "4", rx: "1" } },
    { tag: "line", attrs: { x1: "11", y1: "8", x2: "20", y2: "8" } },
    { tag: "rect", attrs: { x: "4", y: "14", width: "5", height: "4", rx: "1" } },
    { tag: "line", attrs: { x1: "11", y1: "16", x2: "20", y2: "16" } }
  ],
  fullscreen: [
    { tag: "path", attrs: { d: "M8 3H3v5M16 3h5v5M21 16v5h-5M3 16v5h5" } }
  ],
  "export-svg": [
    { tag: "rect", attrs: { x: "4", y: "4", width: "16", height: "16", rx: "2" } },
    { tag: "path", attrs: { d: "M8 15l3-4 2 3 3-4" } }
  ],
  "export-png": [
    { tag: "rect", attrs: { x: "4", y: "4", width: "16", height: "16", rx: "2" } },
    { tag: "line", attrs: { x1: "8", y1: "15", x2: "16", y2: "15" } },
    { tag: "line", attrs: { x1: "8", y1: "11", x2: "16", y2: "11" } }
  ]
};
