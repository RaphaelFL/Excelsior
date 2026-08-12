import { describe, expect, it } from "vitest";
import { DashboardComposer } from "../src";

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

describe("DashboardComposer", () => {
  it("creates and updates dashboard widgets", () => {
    const container = document.createElement("div");
    setContainerSize(container, 1200, 720);
    document.body.append(container);

    const composer = new DashboardComposer(container, {
      columns: 12,
      gap: 8
    });

    composer.setWidgets([
      {
        id: "kpi-sales",
        row: 1,
        col: 1,
        colSpan: 6,
        figure: {
          data: [{ type: "line", x: ["Jan", "Fev", "Mar"], y: [12, 19, 27] }]
        }
      },
      {
        id: "kpi-costs",
        row: 1,
        col: 7,
        colSpan: 6,
        figure: {
          data: [{ type: "bar", x: ["Jan", "Fev", "Mar"], y: [9, 14, 16] }]
        }
      }
    ]);

    expect(container.querySelectorAll("[data-dashboard-widget-id]").length).toBe(2);
    expect(container.querySelectorAll("svg").length).toBeGreaterThan(0);

    composer.updateWidget("kpi-sales", {
      data: [{ type: "line", x: ["Abr", "Mai"], y: [21, 30] }]
    });
    composer.removeWidget("kpi-costs");
    expect(container.querySelectorAll("[data-dashboard-widget-id]").length).toBe(1);

    composer.destroy();
    container.remove();
  });
});
