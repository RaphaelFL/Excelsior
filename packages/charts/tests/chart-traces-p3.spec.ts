import { describe, expect, it } from "vitest";
import { createFigure } from "../src";

const setContainerSize = (container: HTMLElement, width: number, height: number): void => {
  Object.defineProperty(container, "clientWidth", { configurable: true, value: width });
  Object.defineProperty(container, "clientHeight", { configurable: true, value: height });
  container.getBoundingClientRect = () =>
    ({
      width,
      height,
      x: 0,
      y: 0,
      top: 0,
      left: 0,
      right: width,
      bottom: height,
      toJSON: () => ({})
    }) as DOMRect;
};

const createPointerLikeEvent = (
  type: string,
  init: {
    pointerId?: number;
    clientX?: number;
    clientY?: number;
    shiftKey?: boolean;
    button?: number;
  }
): Event => {
  if (typeof PointerEvent !== "undefined") {
    return new PointerEvent(type, {
      bubbles: true,
      cancelable: true,
      pointerId: init.pointerId ?? 1,
      clientX: init.clientX ?? 0,
      clientY: init.clientY ?? 0,
      shiftKey: init.shiftKey ?? false,
      button: init.button ?? 0
    });
  }
  const fallback = new Event(type, { bubbles: true, cancelable: true });
  Object.defineProperty(fallback, "pointerId", { configurable: true, value: init.pointerId ?? 1 });
  Object.defineProperty(fallback, "clientX", { configurable: true, value: init.clientX ?? 0 });
  Object.defineProperty(fallback, "clientY", { configurable: true, value: init.clientY ?? 0 });
  Object.defineProperty(fallback, "shiftKey", { configurable: true, value: init.shiftKey ?? false });
  Object.defineProperty(fallback, "button", { configurable: true, value: init.button ?? 0 });
  return fallback;
};

const createWheelLikeEvent = (deltaY: number): Event => {
  if (typeof WheelEvent !== "undefined") {
    return new WheelEvent("wheel", {
      bubbles: true,
      cancelable: true,
      deltaY
    });
  }
  const fallback = new Event("wheel", { bubbles: true, cancelable: true });
  Object.defineProperty(fallback, "deltaY", { configurable: true, value: deltaY });
  return fallback;
};

describe("p3 traces and advanced axes", () => {
  it("renders area, waterfall, funnel and contour traces", () => {
    const container = document.createElement("div");
    setContainerSize(container, 860, 480);
    document.body.append(container);

    const areaChart = createFigure(container, {
      data: [
        {
          type: "area",
          name: "Acumulado",
          x: ["Q1", "Q2", "Q3", "Q4"],
          y: [12, 18, 16, 24]
        }
      ]
    });
    expect(container.querySelectorAll("[data-trace='area'] path").length).toBeGreaterThan(0);
    areaChart.destroy();

    const waterfallChart = createFigure(container, {
      data: [
        {
          type: "waterfall",
          x: ["Receita", "Custo", "Imposto", "Total"],
          y: [100, -40, -10, 50],
          measure: ["relative", "relative", "relative", "total"]
        }
      ]
    });
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    waterfallChart.destroy();

    const funnelChart = createFigure(container, {
      data: [
        {
          type: "funnel",
          labels: ["Visitas", "Leads", "Propostas", "Fechados"],
          values: [1200, 580, 240, 90]
        }
      ]
    });
    expect(container.querySelectorAll("path").length).toBeGreaterThan(0);
    funnelChart.destroy();

    const contourChart = createFigure(container, {
      data: [
        {
          type: "contour",
          levels: 8,
          z: [
            [10, 12, 16, 20],
            [9, 14, 21, 24],
            [7, 13, 18, 23],
            [5, 10, 14, 19]
          ]
        }
      ]
    });
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    contourChart.destroy();

    container.remove();
  });

  it("formats date and log axes with configured ticks", () => {
    const container = document.createElement("div");
    setContainerSize(container, 760, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: ["2026-01-01", "2026-02-01", "2026-03-01", "2026-04-01"],
          y: [1, 10, 100, 1000]
        }
      ],
      layout: {
        xAxis: {
          type: "date",
          tickFormat: "date",
          tickCount: 4
        },
        yAxis: {
          type: "log",
          tickFormat: "fixed:1",
          tickCount: 4
        }
      }
    });

    const labels = Array.from(container.querySelectorAll("text")).map((node) => node.textContent ?? "");
    expect(labels.some((label) => /^\d{4}-\d{2}-\d{2}$/.test(label))).toBe(true);
    expect(labels.some((label) => /^-?\d+\.\d$/.test(label))).toBe(true);

    chart.destroy();
    container.remove();
  });

  it("renders xAxis2 title and multicategory grouped labels", () => {
    const container = document.createElement("div");
    setContainerSize(container, 780, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: ["2026|Q1", "2026|Q2", "2027|Q1"],
          y: [12, 18, 16]
        },
        {
          type: "scatter",
          xAxisRef: "x2",
          mode: "markers",
          x: ["S1", "S2", "S3"],
          y: [2, 3, 2.4]
        }
      ],
      layout: {
        xAxis: {
          type: "multicategory",
          title: "Periodo"
        },
        xAxis2: {
          type: "category",
          title: "Topo"
        }
      }
    });

    const labels = Array.from(container.querySelectorAll("text")).map((node) => node.textContent ?? "");
    expect(labels).toContain("Topo");
    expect(labels).toContain("2026");
    expect(labels).toContain("Q1");

    chart.destroy();
    container.remove();
  });

  it("supports interactive WebGL camera controls for 3d traces", () => {
    const container = document.createElement("div");
    setContainerSize(container, 700, 420);
    document.body.append(container);
    const originalGetContext = HTMLCanvasElement.prototype.getContext;
    let drawCalls = 0;

    const fakeShader = {};
    const fakeProgram = {};
    const fakeBuffer = {};
    const fakeWebgl = {
      VERTEX_SHADER: 0x8b31,
      FRAGMENT_SHADER: 0x8b30,
      COMPILE_STATUS: 0x8b81,
      LINK_STATUS: 0x8b82,
      ARRAY_BUFFER: 0x8892,
      STATIC_DRAW: 0x88e4,
      COLOR_BUFFER_BIT: 0x4000,
      POINTS: 0x0000,
      TRIANGLES: 0x0004,
      viewport: () => {},
      clearColor: () => {},
      clear: () => {},
      createShader: () => fakeShader,
      shaderSource: () => {},
      compileShader: () => {},
      getShaderParameter: () => true,
      getShaderInfoLog: () => "",
      createProgram: () => fakeProgram,
      attachShader: () => {},
      linkProgram: () => {},
      getProgramParameter: () => true,
      getProgramInfoLog: () => "",
      deleteShader: () => {},
      createBuffer: () => fakeBuffer,
      useProgram: () => {},
      bindBuffer: () => {},
      getAttribLocation: () => 0,
      enableVertexAttribArray: () => {},
      vertexAttribPointer: () => {},
      getUniformLocation: () => ({}),
      uniform4f: () => {},
      uniform1f: () => {},
      bufferData: () => {},
      drawArrays: () => {
        drawCalls += 1;
      },
      deleteBuffer: () => {},
      deleteProgram: () => {}
    } as unknown;

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: (contextId: string) => {
        if (contextId === "webgl") {
          return fakeWebgl;
        }
        return null;
      }
    });

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter3d",
          x: [0, 1, 2, 3],
          y: [0, 1, 0.5, 2],
          z: [1, 1.8, 2.1, 3.2],
          mode: "markers"
        }
      ],
      config: {
        renderer: "webgl"
      }
    });

    const interactionLayer = container.querySelector(".excelsior-chart-webgl-root")?.lastElementChild as HTMLElement | null;
    expect(interactionLayer).not.toBeNull();
    if (interactionLayer) {
      interactionLayer.dispatchEvent(createPointerLikeEvent("pointerdown", { pointerId: 1, clientX: 120, clientY: 120 }));
      interactionLayer.dispatchEvent(createPointerLikeEvent("pointermove", { pointerId: 1, clientX: 165, clientY: 150 }));
      interactionLayer.dispatchEvent(createPointerLikeEvent("pointerup", { pointerId: 1, clientX: 165, clientY: 150 }));
      interactionLayer.dispatchEvent(createWheelLikeEvent(-120));
    }

    expect(drawCalls).toBeGreaterThan(1);

    chart.destroy();
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: originalGetContext
    });
    container.remove();
  });

  it("renders violin, density, quiver, sankey and parallel-categories traces", () => {
    const container = document.createElement("div");
    setContainerSize(container, 860, 480);
    document.body.append(container);

    const violin = createFigure(container, {
      data: [
        {
          type: "violin",
          name: "Distribuicao A",
          values: [10, 11, 12, 12, 13, 15, 18],
          showBox: true
        }
      ]
    });
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    violin.destroy();

    const density = createFigure(container, {
      data: [
        {
          type: "density",
          values: [1, 2, 2.3, 2.8, 3.1, 3.2, 4, 4.4, 5.1]
        }
      ]
    });
    expect(container.querySelectorAll("path").length).toBeGreaterThan(0);
    density.destroy();

    const quiver = createFigure(container, {
      data: [
        {
          type: "quiver",
          x: [0, 1, 2],
          y: [0, 1, 1],
          u: [1, 0.8, 0.4],
          v: [0.4, 0.2, 1]
        }
      ]
    });
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    quiver.destroy();

    const sankey = createFigure(container, {
      data: [
        {
          type: "sankey",
          nodes: {
            ids: ["A", "B", "C", "D"],
            labels: ["Entrada", "Processo", "Saida 1", "Saida 2"]
          },
          links: {
            source: [0, 1, 1],
            target: [1, 2, 3],
            value: [10, 6, 4]
          }
        }
      ]
    });
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    sankey.destroy();

    const parcat = createFigure(container, {
      data: [
        {
          type: "parallel-categories",
          dimensions: [
            { name: "Canal", values: ["Web", "Loja", "Web", "App"] },
            { name: "Regiao", values: ["Sul", "Sudeste", "Sul", "Nordeste"] },
            { name: "Status", values: ["Novo", "Recorrente", "Novo", "Recorrente"] }
          ]
        }
      ]
    });
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    parcat.destroy();

    container.remove();
  });

  it("includes conversion ratios on funnel and colors quiver by magnitude", () => {
    const container = document.createElement("div");
    setContainerSize(container, 860, 480);
    document.body.append(container);

    const funnel = createFigure(container, {
      data: [
        {
          type: "funnel",
          labels: ["Visitas", "Leads", "Propostas", "Fechados"],
          values: [1200, 600, 240, 96]
        }
      ]
    });

    const firstStage = container.querySelector<HTMLElement>("[data-chart-interactive='point'][data-point-index='0']");
    expect(firstStage).not.toBeNull();
    expect(firstStage?.dataset.pointY).toContain("% total");
    expect(firstStage?.dataset.pointY).toContain("% etapa anterior");
    funnel.destroy();

    const quiver = createFigure(container, {
      data: [
        {
          type: "quiver",
          x: [0, 1, 2, 3],
          y: [0, 1, 1, 0],
          u: [0.2, 1.2, 0.4, 1.8],
          v: [0.1, 0.6, 1.6, 0.2],
          colorByMagnitude: true,
          colorscale: ["#0ea5e9", "#f59e0b", "#ef4444"]
        }
      ]
    });

    const quiverPoint = container.querySelector<HTMLElement>("[data-chart-interactive='point'][data-point-index='0']");
    expect(quiverPoint).not.toBeNull();
    expect(quiverPoint?.dataset.pointY).toContain("|m|=");
    quiver.destroy();

    container.remove();
  });

  it("falls back from WebGL to 2D renderer when unavailable", () => {
    const container = document.createElement("div");
    setContainerSize(container, 720, 420);
    document.body.append(container);

    const originalGetContext = HTMLCanvasElement.prototype.getContext;
    const fake2d = {
      clearRect: () => {},
      fillRect: () => {},
      strokeRect: () => {},
      beginPath: () => {},
      moveTo: () => {},
      lineTo: () => {},
      stroke: () => {},
      fill: () => {},
      arc: () => {},
      closePath: () => {},
      save: () => {},
      restore: () => {},
      translate: () => {},
      rotate: () => {},
      fillText: () => {},
      strokeText: () => {},
      bezierCurveTo: () => {},
      createLinearGradient: () => ({ addColorStop: () => {} }),
      measureText: () => ({ width: 30 }),
      set globalAlpha(_: number) {},
      get globalAlpha() {
        return 1;
      },
      set strokeStyle(_: string) {},
      set fillStyle(_: string) {},
      set lineWidth(_: number) {},
      set font(_: string) {},
      set textAlign(_: CanvasTextAlign) {}
    } as unknown;
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: function patchedGetContext(this: HTMLCanvasElement, contextId: string, options?: unknown) {
        if (contextId === "webgl") {
          return null;
        }
        if (contextId === "2d") {
          return fake2d;
        }
        return originalGetContext.call(this, contextId as never, options as never);
      }
    });

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter3d",
          x: [0, 1, 2],
          y: [0, 1, 0.5],
          z: [1, 2, 3]
        }
      ],
      config: {
        renderer: "webgl",
        webglFallback: true
      }
    });

    expect(container.querySelector(".excelsior-chart-canvas-root") || container.querySelector("svg")).not.toBeNull();

    chart.destroy();
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: originalGetContext
    });
    container.remove();
  });
});
