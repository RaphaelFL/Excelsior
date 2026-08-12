import { describe, expect, it, vi } from "vitest";
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

describe("animation and accessibility", () => {
  it("plays frames and emits lifecycle events", () => {
    vi.useFakeTimers();
    const container = document.createElement("div");
    setContainerSize(container, 740, 420);
    document.body.append(container);

    const frameEvents: Array<{ index: number; name?: string }> = [];
    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: ["Q1", "Q2"],
          y: [1, 2]
        }
      ],
      frames: [
        {
          name: "f1",
          data: [
            {
              type: "line",
              x: ["Q1", "Q2"],
              y: [2, 4]
            }
          ]
        },
        {
          name: "f2",
          data: [
            {
              type: "line",
              x: ["Q1", "Q2"],
              y: [3, 6]
            }
          ]
        }
      ],
      config: {
        frameDurationMs: 10
      }
    });

    chart.on("animation:frame", (payload) => frameEvents.push({ index: payload.frameIndex, name: payload.frameName }));
    chart.playFrames({ loop: false, intervalMs: 10 });
    expect(chart.isAnimating()).toBe(true);
    vi.advanceTimersByTime(40);
    expect(frameEvents.length).toBeGreaterThanOrEqual(2);
    chart.stopFrames();
    expect(chart.isAnimating()).toBe(false);

    chart.destroy();
    container.remove();
    vi.useRealTimers();
  });

  it("renders accessible table and exports csv", () => {
    const container = document.createElement("div");
    setContainerSize(container, 740, 420);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter",
          x: ["Jan", "Fev"],
          y: [10, 12],
          mode: "markers"
        }
      ],
      config: {
        accessibleTable: true,
        ariaDescription: "Relatorio mensal"
      }
    });

    expect(container.getAttribute("role")).toBe("img");
    expect(container.getAttribute("aria-label")).toBe("Relatorio mensal");
    const table = container.querySelector(".excelsior-chart-accessible-table");
    expect(table).not.toBeNull();
    expect(table?.querySelectorAll("tbody tr").length).toBe(2);
    const csv = chart.exportDataTable();
    expect(csv).toContain("trace,x,y");
    expect(csv).toContain("Jan");

    chart.destroy();
    container.remove();
  });
});
