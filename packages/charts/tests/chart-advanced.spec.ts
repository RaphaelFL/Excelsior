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

describe("advanced chart features", () => {
  it("navigates the x domain with range selector and range slider controls", () => {
    const container = document.createElement("div");
    setContainerSize(container, 760, 420);
    document.body.append(container);
    const chart = createFigure(container, {
      data: [{ type: "line", x: [0, 25, 50, 75, 100], y: [1, 2, 3, 4, 5] }],
      layout: {
        xAxis: {
          rangeSelector: { visible: true, buttons: [{ label: "Últimos 25%", fraction: 0.25 }] },
          rangeSlider: { visible: true, start: 0.1, end: 0.9 }
        }
      }
    });
    let zoomed: { xMin: number; xMax: number } | undefined;
    chart.on("axis:zoomed", (event) => {
      zoomed = event;
    });

    expect(container.querySelectorAll(".excelsior-chart-range-navigation input[type='range']")).toHaveLength(2);
    container.querySelector<HTMLButtonElement>("[data-range-fraction='0.25']")!.click();
    expect(zoomed?.xMin).toBeCloseTo(3);
    expect(zoomed?.xMax).toBeCloseTo(4);

    chart.destroy();
    container.remove();
  });

  it("applies safe transforms on cartesian traces", () => {
    const container = document.createElement("div");
    setContainerSize(container, 700, 380);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: [3, 1, 2],
          y: [30, 10, 20],
          transforms: [
            { type: "sort", by: "x", direction: "asc" },
            { type: "normalize", field: "y", strategy: "max" }
          ]
        }
      ]
    });

    const serialized = JSON.parse(chart.toJson()) as { data: Array<{ x: number[]; y: number[] }> };
    expect(serialized.data[0].x).toEqual([1, 2, 3]);
    expect(serialized.data[0].y[0]).toBeCloseTo(10 / 30, 5);
    expect(serialized.data[0].y[2]).toBeCloseTo(1, 5);

    chart.destroy();
    container.remove();
  });

  it("supports subplot layout with separated plot areas", () => {
    const container = document.createElement("div");
    setContainerSize(container, 840, 520);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter",
          subplot: 0,
          x: ["A", "B"],
          y: [1, 2],
          mode: "markers"
        },
        {
          type: "scatter",
          subplot: 3,
          x: ["A", "B"],
          y: [2, 3],
          mode: "markers"
        }
      ],
      layout: {
        subplots: {
          rows: 2,
          cols: 2
        }
      }
    });

    const plotBackgrounds = container.querySelectorAll("rect[stroke='#dbe2ea']");
    expect(plotBackgrounds.length).toBeGreaterThanOrEqual(2);

    chart.destroy();
    container.remove();
  });

  it("renders shapes, annotations and local image layers", () => {
    const container = document.createElement("div");
    setContainerSize(container, 760, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter",
          x: ["A", "B", "C"],
          y: [1, 3, 2]
        }
      ],
      layout: {
        shapes: [
          {
            type: "line",
            x0: 0,
            y0: 0,
            x1: 2,
            y1: 3,
            xRef: "data",
            yRef: "data",
            stroke: "#2563eb"
          }
        ],
        annotations: [
          {
            text: "Pico",
            x: 1,
            y: 3,
            xRef: "data",
            yRef: "data"
          }
        ],
        images: [
          {
            source: "data:image/gif;base64,R0lGODlhAQABAAAAACwAAAAAAQABAAA=",
            x: 0.75,
            y: 0.1,
            width: 18,
            height: 18,
            xRef: "paper",
            yRef: "paper"
          }
        ]
      }
    });

    const svg = container.querySelector("svg");
    expect(svg?.querySelectorAll("line").length).toBeGreaterThan(0);
    expect(svg?.textContent).toContain("Pico");
    expect(svg?.querySelector("image")).not.toBeNull();

    chart.destroy();
    container.remove();
  });

  it("respects axis zeroLine settings", () => {
    const container = document.createElement("div");
    setContainerSize(container, 760, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: [-2, -1, 0, 1, 2],
          y: [-3, 2, -1, 1, 4]
        }
      ],
      layout: {
        xAxis: {
          type: "linear",
          zeroLine: false
        },
        yAxis: {
          type: "linear",
          zeroLine: false
        }
      }
    });

    const zeroLines = container.querySelectorAll("line[stroke='#94a3b8'][stroke-width='1.5']");
    expect(zeroLines.length).toBe(0);

    chart.destroy();
    container.remove();
  });

  it("supports image overlays on canvas renderer", () => {
    const container = document.createElement("div");
    setContainerSize(container, 760, 420);
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
      drawImage: () => {},
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
        if (contextId === "2d") {
          return fake2d;
        }
        return originalGetContext.call(this, contextId as never, options as never);
      }
    });

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter",
          x: [0, 1, 2],
          y: [2, 3, 1]
        }
      ],
      layout: {
        images: [
          {
            source: "data:image/gif;base64,R0lGODlhAQABAAAAACwAAAAAAQABAAA=",
            x: 0.2,
            y: 0.2,
            width: 24,
            height: 24,
            xRef: "paper",
            yRef: "paper"
          }
        ]
      },
      config: {
        renderer: "canvas"
      }
    });

    expect(container.querySelector("canvas")).not.toBeNull();

    chart.destroy();
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: originalGetContext
    });
    container.remove();
  });
});
