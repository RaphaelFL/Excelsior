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

describe("p1 traces", () => {
  it("renders histogram trace on svg renderer", () => {
    const container = document.createElement("div");
    setContainerSize(container, 700, 360);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "histogram",
          name: "Distribuicao",
          values: [10, 12, 13, 12, 18, 16, 15, 12, 11, 9, 8, 17],
          bins: 5
        }
      ]
    });

    expect(container.querySelectorAll("[data-trace='histogram'] rect").length).toBeGreaterThan(0);
    chart.destroy();
    container.remove();
  });

  it("renders box trace on svg renderer", () => {
    const container = document.createElement("div");
    setContainerSize(container, 700, 360);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "box",
          name: "Latencia",
          values: [10, 11, 12, 12, 13, 18, 23, 14, 13, 12]
        }
      ]
    });

    expect(container.querySelector("[data-trace='box'] rect")).not.toBeNull();
    chart.destroy();
    container.remove();
  });

  it("renders heatmap trace and point metadata", () => {
    const container = document.createElement("div");
    setContainerSize(container, 720, 400);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "heatmap",
          name: "Matriz",
          x: ["A", "B", "C"],
          y: ["L1", "L2"],
          z: [
            [1, 3, 5],
            [2, 4, 6]
          ]
        }
      ]
    });

    expect(container.querySelectorAll("[data-trace='heatmap']").length).toBe(6);
    expect(container.querySelector("[data-chart-interactive='point'][data-point-index='0']")).not.toBeNull();
    chart.destroy();
    container.remove();
  });

  it("renders candlestick and ohlc traces", () => {
    const container = document.createElement("div");
    setContainerSize(container, 740, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "candlestick",
          name: "PETR4",
          x: ["D1", "D2", "D3"],
          open: [10, 11, 10.5],
          high: [12, 12.3, 11.2],
          low: [9.6, 10.8, 10.1],
          close: [11.5, 10.9, 11]
        },
        {
          type: "ohlc",
          name: "VALE3",
          x: ["D1", "D2", "D3"],
          open: [5, 5.4, 5.1],
          high: [5.6, 5.8, 5.5],
          low: [4.8, 5.1, 4.9],
          close: [5.2, 5.3, 5.4]
        }
      ]
    });

    expect(container.querySelector("[data-trace='candlestick']")).not.toBeNull();
    expect(container.querySelector("[data-trace='ohlc']")).not.toBeNull();
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);

    chart.destroy();
    container.remove();
  });

  it("renders polar trace", () => {
    const container = document.createElement("div");
    setContainerSize(container, 680, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "polar",
          name: "Radar",
          theta: [0, 72, 144, 216, 288],
          r: [4, 7, 5, 8, 6],
          mode: "lines+markers"
        }
      ]
    });

    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    chart.destroy();
    container.remove();
  });

  it("renders ternary trace", () => {
    const container = document.createElement("div");
    setContainerSize(container, 680, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "ternary",
          name: "Composicao",
          a: [30, 20, 50],
          b: [30, 50, 20],
          c: [40, 30, 30]
        }
      ]
    });

    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    chart.destroy();
    container.remove();
  });

  it("renders geo trace", () => {
    const container = document.createElement("div");
    setContainerSize(container, 760, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "geo",
          name: "Regioes",
          geojson: {
            type: "FeatureCollection",
            features: [
              {
                type: "Feature",
                geometry: {
                  type: "Polygon",
                  coordinates: [
                    [
                      [-47.0, -23.0],
                      [-46.0, -23.0],
                      [-46.0, -22.0],
                      [-47.0, -22.0],
                      [-47.0, -23.0]
                    ]
                  ]
                },
                properties: {
                  value: 10
                }
              }
            ]
          }
        }
      ]
    });

    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    chart.destroy();
    container.remove();
  });
});
