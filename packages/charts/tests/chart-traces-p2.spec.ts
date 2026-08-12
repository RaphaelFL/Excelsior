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

describe("p2 traces", () => {
  it("renders domain traces donut, sunburst and treemap", () => {
    const container = document.createElement("div");
    setContainerSize(container, 780, 420);
    document.body.append(container);

    const donut = createFigure(container, {
      data: [
        {
          type: "donut",
          values: [20, 30, 50],
          labels: ["A", "B", "C"],
          pull: [0.2, 0, 0]
        }
      ]
    });
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    donut.destroy();

    const sunburst = createFigure(container, {
      data: [
        {
          type: "sunburst",
          labels: ["Total", "A", "B"],
          parents: ["", "Total", "Total"],
          values: [100, 40, 60]
        }
      ]
    });
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    sunburst.destroy();

    const treemap = createFigure(container, {
      data: [
        {
          type: "treemap",
          labels: ["A", "B", "C"],
          parents: ["", "", ""],
          values: [20, 30, 50]
        }
      ]
    });
    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    treemap.destroy();
    container.remove();
  });

  it("renders geo-scatter and geo-line with local coordinates", () => {
    const container = document.createElement("div");
    setContainerSize(container, 760, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "geo-scatter",
          name: "Pontos",
          lat: [-23.55, -22.9, -22.4],
          lon: [-46.63, -43.2, -42.9],
          mode: "lines+markers"
        },
        {
          type: "geo-line",
          name: "Rota",
          paths: [
            [
              { lat: -23.55, lon: -46.63 },
              { lat: -22.9, lon: -43.2 }
            ]
          ]
        }
      ]
    });

    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    chart.destroy();
    container.remove();
  });

  it("supports drilldown and drill-up in hierarchy traces", () => {
    const container = document.createElement("div");
    setContainerSize(container, 780, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "sunburst",
          labels: ["A", "A1", "A2", "B", "B1"],
          parents: ["", "A", "A", "", "B"],
          ids: ["A", "A1", "A2", "B", "B1"],
          values: [60, 35, 25, 40, 40]
        }
      ]
    });

    const branch = container.querySelector<SVGElement>("[data-chart-interactive='point'][data-point-index='0']");
    expect(branch).not.toBeNull();
    branch?.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const afterDrilldown = JSON.parse(chart.toJson()) as {
      data: Array<{ rootId?: string }>;
    };
    expect(afterDrilldown.data[0].rootId).toBe("A");

    const backControl = container.querySelector<SVGElement>("[data-chart-interactive='point'][data-point-index='-1']");
    expect(backControl).not.toBeNull();
    backControl?.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const afterDrillup = JSON.parse(chart.toJson()) as {
      data: Array<{ rootId?: string }>;
    };
    expect(afterDrillup.data[0].rootId).toBeUndefined();

    chart.destroy();
    container.remove();
  });

  it("renders projected 3d traces in svg renderer", () => {
    const container = document.createElement("div");
    setContainerSize(container, 780, 460);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter3d",
          x: [0, 1, 2, 3],
          y: [1, 2, 1.5, 2.4],
          z: [1, 1.8, 2.2, 3.1],
          mode: "lines+markers"
        },
        {
          type: "surface",
          z: [
            [1, 2, 3],
            [2, 3, 4],
            [1, 2, 2]
          ]
        },
        {
          type: "mesh3d",
          x: [0, 2, 1, 1.5],
          y: [0, 0, 2, 1],
          z: [0, 1, 0.5, 2],
          i: [0, 0],
          j: [1, 1],
          k: [2, 3]
        }
      ]
    });

    expect(container.querySelectorAll("[data-chart-interactive='point']").length).toBeGreaterThan(0);
    expect(container.querySelectorAll("path").length).toBeGreaterThan(0);
    chart.destroy();
    container.remove();
  });
});
