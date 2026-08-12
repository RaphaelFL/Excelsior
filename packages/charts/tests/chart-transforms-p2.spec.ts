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

describe("p2 transforms", () => {
  it("supports group and aggregate transforms", () => {
    const container = document.createElement("div");
    setContainerSize(container, 680, 320);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: ["A", "A", "B", "B"],
          y: [1, 2, 3, 5],
          transforms: [
            { type: "group", by: "x", aggregate: "sum" },
            { type: "aggregate", field: "y", op: "sum" }
          ]
        }
      ]
    });

    const serialized = JSON.parse(chart.toJson()) as { data: Array<{ y: number[] }> };
    expect(serialized.data[0].y).toEqual([11]);

    chart.destroy();
    container.remove();
  });

  it("supports bin and stack transforms", () => {
    const container = document.createElement("div");
    setContainerSize(container, 680, 320);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: [1, 2, 3, 4, 5, 6],
          y: [1, 2, 3, 4, 5, 6],
          transforms: [
            { type: "bin", field: "y", bins: 3 },
            { type: "stack", field: "y", strategy: "running" }
          ]
        }
      ]
    });

    const serialized = JSON.parse(chart.toJson()) as { data: Array<{ y: number[] }> };
    expect(serialized.data[0].y.length).toBe(3);
    expect(serialized.data[0].y[2]).toBeGreaterThanOrEqual(serialized.data[0].y[1]);

    chart.destroy();
    container.remove();
  });
});
