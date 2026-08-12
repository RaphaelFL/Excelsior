import { describe, expect, it } from "vitest";
import { ChartConfigurationError, TraceRegistry, type ChartTrace } from "../src/index";

describe("TraceRegistry", () => {
  it("registers default trace types", () => {
    const registry = new TraceRegistry();

    expect(registry.has("scatter")).toBe(true);
    expect(registry.has("line")).toBe(true);
    expect(registry.has("bar")).toBe(true);
    expect(registry.has("area")).toBe(true);
    expect(registry.has("violin")).toBe(true);
    expect(registry.has("density")).toBe(true);
    expect(registry.has("distribution")).toBe(true);
    expect(registry.has("pie")).toBe(true);
    expect(registry.has("histogram")).toBe(true);
    expect(registry.has("box")).toBe(true);
    expect(registry.has("heatmap")).toBe(true);
    expect(registry.has("contour")).toBe(true);
    expect(registry.has("quiver")).toBe(true);
    expect(registry.has("candlestick")).toBe(true);
    expect(registry.has("ohlc")).toBe(true);
    expect(registry.has("waterfall")).toBe(true);
    expect(registry.has("funnel")).toBe(true);
    expect(registry.has("polar")).toBe(true);
    expect(registry.has("ternary")).toBe(true);
    expect(registry.has("geo")).toBe(true);
    expect(registry.has("donut")).toBe(true);
    expect(registry.has("sunburst")).toBe(true);
    expect(registry.has("treemap")).toBe(true);
    expect(registry.has("sankey")).toBe(true);
    expect(registry.has("parallel-categories")).toBe(true);
    expect(registry.has("geo-scatter")).toBe(true);
    expect(registry.has("geo-line")).toBe(true);
    expect(registry.has("scatter3d")).toBe(true);
    expect(registry.has("surface")).toBe(true);
    expect(registry.has("mesh3d")).toBe(true);
  });

  it("throws controlled error for unknown traces", () => {
    const registry = new TraceRegistry();

    expect(() => registry.resolve("unknown")).toThrow(ChartConfigurationError);
  });

  it("allows custom trace registration", () => {
    const registry = new TraceRegistry();
    registry.register({
      type: "custom-points",
      renderer: "svg",
      validate: (trace: ChartTrace) => {
        if (!Array.isArray((trace as { x?: unknown }).x)) {
          throw new ChartConfigurationError("CUSTOM_TRACE_INVALID", "x is required.");
        }
      }
    });

    const definition = registry.resolve("custom-points");
    expect(definition.type).toBe("custom-points");
  });

  it("validates pie pull configuration", () => {
    const registry = new TraceRegistry();
    const pieDefinition = registry.resolve("pie");

    expect(() =>
      pieDefinition.validate({
        type: "pie",
        values: [10, 20, 30],
        pull: [0.1, 0.2]
      } as ChartTrace)
    ).toThrow(ChartConfigurationError);

    expect(() =>
      pieDefinition.validate({
        type: "pie",
        values: [10, 20, 30],
        pull: -0.1
      } as ChartTrace)
    ).toThrow(ChartConfigurationError);

    expect(() =>
      pieDefinition.validate({
        type: "pie",
        values: [10, 20, 30],
        pull: [0.1, 0.2, 0.3]
      } as ChartTrace)
    ).not.toThrow();
  });
});
