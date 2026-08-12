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
    altKey?: boolean;
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
      altKey: init.altKey ?? false,
      button: init.button ?? 0
    });
  }
  const fallback = new Event(type, { bubbles: true, cancelable: true });
  Object.defineProperty(fallback, "pointerId", { configurable: true, value: init.pointerId ?? 1 });
  Object.defineProperty(fallback, "clientX", { configurable: true, value: init.clientX ?? 0 });
  Object.defineProperty(fallback, "clientY", { configurable: true, value: init.clientY ?? 0 });
  Object.defineProperty(fallback, "shiftKey", { configurable: true, value: init.shiftKey ?? false });
  Object.defineProperty(fallback, "altKey", { configurable: true, value: init.altKey ?? false });
  Object.defineProperty(fallback, "button", { configurable: true, value: init.button ?? 0 });
  return fallback;
};

describe("chart interactions", () => {
  it("toggles traces from legend clicks", () => {
    const container = document.createElement("div");
    setContainerSize(container, 740, 400);
    document.body.append(container);

    const legendEvents: Array<{ traceIndex: number; visible: boolean }> = [];
    const chart = createFigure(container, {
      data: [
        {
          type: "scatter",
          name: "Receita",
          x: ["Jan", "Fev", "Mar"],
          y: [120, 150, 180],
          mode: "markers"
        },
        {
          type: "line",
          name: "Custos",
          x: ["Jan", "Fev", "Mar"],
          y: [80, 110, 130]
        }
      ]
    });

    chart.on("legend:toggled", (payload) => legendEvents.push(payload));

    const legendItem = container.querySelector<SVGGElement>("[data-chart-legend='item'][data-legend-trace-index='0']");
    expect(legendItem).not.toBeNull();
    legendItem?.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(legendEvents).toHaveLength(1);
    expect(legendEvents[0]).toEqual({ traceIndex: 0, visible: false });
    expect(container.querySelector("[data-trace-index='0']")).toBeNull();

    chart.destroy();
    container.remove();
  });

  it("emits click payload for interactive points", () => {
    const container = document.createElement("div");
    setContainerSize(container, 640, 360);
    document.body.append(container);

    const clickedPayloads: Array<{ traceIndex: number; pointIndex: number; x: string; y: string }> = [];
    const selectedPayloads: Array<{ traceIndex: number; pointIndex: number }> = [];
    const selectionChanges: Array<{ mode: string; size: number }> = [];
    const chart = createFigure(container, {
      data: [
        {
          type: "bar",
          name: "Pedidos",
          x: ["A", "B"],
          y: [12, 18]
        }
      ]
    });

    chart.on("trace:click", (payload) =>
      clickedPayloads.push({
        traceIndex: payload.traceIndex,
        pointIndex: payload.pointIndex,
        x: payload.x,
        y: payload.y
      })
    );
    chart.on("trace:selected", (payload) =>
      selectedPayloads.push({
        traceIndex: payload.traceIndex,
        pointIndex: payload.pointIndex
      })
    );
    chart.on("selection:changed", (payload) => {
      selectionChanges.push({ mode: payload.mode, size: payload.points.length });
    });

    const firstBar = container.querySelector("[data-chart-interactive='point'][data-point-index='0']");
    expect(firstBar).not.toBeNull();
    firstBar?.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(clickedPayloads).toHaveLength(1);
    expect(selectedPayloads).toHaveLength(1);
    expect(clickedPayloads[0]).toEqual({
      traceIndex: 0,
      pointIndex: 0,
      x: "A",
      y: "12.00"
    });
    expect(selectedPayloads[0]).toEqual({
      traceIndex: 0,
      pointIndex: 0
    });
    expect(selectionChanges.at(-1)).toEqual({
      mode: "click",
      size: 1
    });
    expect(chart.getSelection()?.points).toHaveLength(1);
    chart.clearSelection();
    expect(chart.getSelection()).toBeNull();

    chart.destroy();
    container.remove();
  });

  it("applies wheel zoom and exposes modebar/tooltip by config", () => {
    const container = document.createElement("div");
    setContainerSize(container, 720, 420);
    document.body.append(container);

    const zoomEvents: Array<{ xMin: number; xMax: number; yMin: number; yMax: number }> = [];
    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: ["Jan", "Fev", "Mar", "Abr"],
          y: [100, 140, 160, 170]
        }
      ],
      config: {
        modebar: true,
        tooltip: true
      }
    });

    chart.on("axis:zoomed", (payload) => zoomEvents.push(payload));
    const svg = container.querySelector("svg");
    expect(svg).not.toBeNull();
    svg?.dispatchEvent(
      new WheelEvent("wheel", {
        bubbles: true,
        deltaY: -60,
        clientX: 280,
        clientY: 200
      })
    );

    expect(zoomEvents).toHaveLength(1);
    const figure = JSON.parse(chart.toJson()) as { layout: { xAxis: { min?: number; max?: number } } };
    expect(typeof figure.layout.xAxis.min).toBe("number");
    expect(typeof figure.layout.xAxis.max).toBe("number");
    expect(container.querySelector(".excelsior-chart-modebar")).not.toBeNull();
    expect(container.querySelector(".excelsior-chart-tooltip")).not.toBeNull();

    chart.destroy();
    container.remove();
  });

  it("supports comparative hover mode on shared x index", () => {
    const container = document.createElement("div");
    setContainerSize(container, 760, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter",
          name: "Serie A",
          x: ["Jan", "Fev", "Mar"],
          y: [10, 20, 25],
          mode: "markers"
        },
        {
          type: "scatter",
          name: "Serie B",
          x: ["Jan", "Fev", "Mar"],
          y: [8, 18, 29],
          mode: "markers"
        }
      ],
      config: {
        hoverMode: "x",
        tooltip: true
      }
    });

    const point = container.querySelector<HTMLElement>("[data-chart-interactive='point'][data-point-index='1']");
    expect(point).not.toBeNull();
    const moveEvent = new Event("pointermove", { bubbles: true }) as Event & { clientX?: number; clientY?: number };
    Object.defineProperty(moveEvent, "clientX", { configurable: true, value: 220 });
    Object.defineProperty(moveEvent, "clientY", { configurable: true, value: 180 });
    point?.dispatchEvent(moveEvent);

    const tooltipText = container.querySelector(".excelsior-chart-tooltip")?.textContent ?? "";
    expect(tooltipText).toContain("x: Fev");
    expect(tooltipText).toContain("Serie A");
    expect(tooltipText).toContain("Serie B");

    chart.destroy();
    container.remove();
  });

  it("supports axis wheel zoom and rectangular zoom", () => {
    const container = document.createElement("div");
    setContainerSize(container, 780, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: [0, 1, 2, 3, 4, 5],
          y: [10, 12, 15, 21, 34, 55]
        }
      ]
    });

    const root = container.querySelector("svg");
    expect(root).not.toBeNull();
    root?.dispatchEvent(
      new WheelEvent("wheel", {
        bubbles: true,
        cancelable: true,
        deltaY: -80,
        clientX: 280,
        clientY: 220,
        shiftKey: true
      })
    );
    root?.dispatchEvent(
      new WheelEvent("wheel", {
        bubbles: true,
        cancelable: true,
        deltaY: -80,
        clientX: 360,
        clientY: 190,
        altKey: true
      })
    );

    root?.dispatchEvent(createPointerLikeEvent("pointerdown", { pointerId: 7, clientX: 180, clientY: 130 }));
    root?.dispatchEvent(createPointerLikeEvent("pointermove", { pointerId: 7, clientX: 390, clientY: 280 }));
    root?.dispatchEvent(createPointerLikeEvent("pointerup", { pointerId: 7, clientX: 390, clientY: 280 }));

    const figure = JSON.parse(chart.toJson()) as {
      layout: {
        xAxis: { min?: number; max?: number };
        yAxis: { min?: number; max?: number };
      };
    };
    expect(typeof figure.layout.xAxis.min).toBe("number");
    expect(typeof figure.layout.xAxis.max).toBe("number");
    expect(typeof figure.layout.yAxis.min).toBe("number");
    expect(typeof figure.layout.yAxis.max).toBe("number");

    chart.destroy();
    container.remove();
  });

  it("keeps global axis layout unchanged when subplot sync is disabled", () => {
    const container = document.createElement("div");
    setContainerSize(container, 860, 460);
    document.body.append(container);

    const zoomEvents: Array<{ xMin: number; xMax: number; yMin: number; yMax: number }> = [];
    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          subplot: 0,
          x: [0, 1, 2, 3, 4],
          y: [1, 3, 2, 4, 5]
        },
        {
          type: "line",
          subplot: 1,
          x: [0, 1, 2, 3, 4],
          y: [10, 9, 8, 7, 6]
        }
      ],
      layout: {
        subplots: {
          rows: 1,
          cols: 2,
          syncZoom: false
        }
      },
      config: {
        syncSubplotZoom: false
      }
    });

    chart.on("axis:zoomed", (payload) => zoomEvents.push(payload));
    const root = container.querySelector("svg");
    expect(root).not.toBeNull();
    root?.dispatchEvent(
      new WheelEvent("wheel", {
        bubbles: true,
        cancelable: true,
        deltaY: -100,
        clientX: 220,
        clientY: 220
      })
    );

    const figure = JSON.parse(chart.toJson()) as {
      layout: {
        xAxis: { min?: number; max?: number };
        yAxis: { min?: number; max?: number };
      };
    };
    expect(zoomEvents).toHaveLength(1);
    expect(figure.layout.xAxis.min).toBeUndefined();
    expect(figure.layout.xAxis.max).toBeUndefined();
    expect(figure.layout.yAxis.min).toBeUndefined();
    expect(figure.layout.yAxis.max).toBeUndefined();

    chart.destroy();
    container.remove();
  });

  it("filters parallel-categories rows by clicked selection", () => {
    const container = document.createElement("div");
    setContainerSize(container, 800, 420);
    document.body.append(container);

    const chart = createFigure(container, {
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

    const before = container.querySelectorAll("[data-chart-interactive='point']").length;
    expect(before).toBe(4);

    const firstPath = container.querySelector<SVGElement>("[data-chart-interactive='point'][data-point-index='0']");
    expect(firstPath).not.toBeNull();
    firstPath?.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const afterIndices = Array.from(
      container.querySelectorAll<HTMLElement>("[data-chart-interactive='point']")
    ).map((node) => Number(node.dataset.pointIndex));

    expect(afterIndices).toContain(0);
    expect(afterIndices).toContain(2);
    expect(afterIndices).not.toContain(1);
    expect(afterIndices).not.toContain(3);
    expect(container.textContent ?? "").toContain("Filtro ativo");

    chart.destroy();
    container.remove();
  });
});
