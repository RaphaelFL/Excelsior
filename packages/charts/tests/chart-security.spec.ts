import { describe, expect, it } from "vitest";
import { ChartConfigurationError, createFigure } from "../src";

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

describe("chart security", () => {
  it("sanitizes annotation text and blocks unsafe image URLs", () => {
    const container = document.createElement("div");
    setContainerSize(container, 720, 380);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter",
          x: ["A", "B"],
          y: [1, 2]
        }
      ],
      layout: {
        annotations: [
          {
            text: "<script>alert(1)</script>",
            x: 0,
            y: 1
          }
        ],
        images: [
          {
            source: "http://evil.local/x.png",
            x: 0.2,
            y: 0.2,
            width: 20,
            height: 20
          }
        ]
      }
    });

    const svg = container.querySelector("svg");
    expect(svg?.textContent).toContain("‹script›alert(1)‹/script›");
    expect(svg?.querySelector("image")).toBeNull();

    chart.destroy();
    container.remove();
  });

  it("rejects invalid geo traces", () => {
    const container = document.createElement("div");
    setContainerSize(container, 680, 320);
    document.body.append(container);

    expect(() =>
      createFigure(container, {
        data: [
          {
            type: "geo",
            geojson: {
              type: "FeatureCollection",
              features: [
                {
                  type: "Feature",
                  geometry: {
                    type: "Point",
                    coordinates: [0, 0]
                  }
                }
              ]
            } as unknown as {
              type: "FeatureCollection";
              features: Array<{
                type: "Feature";
                geometry: {
                  type: "Polygon" | "MultiPolygon";
                  coordinates: number[][][] | number[][][][];
                };
              }>;
            }
          }
        ]
      })
    ).toThrow(ChartConfigurationError);

    container.remove();
  });
});
