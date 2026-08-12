import { describe, expect, it } from "vitest";
import { ChartConfigurationError, createFigure, fromJson, parseFigureFromJson } from "../src/index";

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

describe("createFigure", () => {
  it("renders an svg figure and supports update lifecycle", () => {
    const container = document.createElement("div");
    setContainerSize(container, 720, 400);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter",
          name: "Receita",
          x: ["Jan", "Fev"],
          y: [120, 150]
        }
      ],
      layout: {
        title: "Receita mensal"
      }
    });

    expect(container.querySelector("svg.excelsior-chart-svg")).not.toBeNull();
    expect(container.querySelector("text")?.textContent).toContain("Receita mensal");
    expect(container.querySelectorAll("[data-trace='scatter'] circle")).toHaveLength(2);

    chart.updateData([
      {
        type: "scatter",
        x: ["Jan", "Fev", "Mar"],
        y: [120, 150, 180],
        mode: "markers"
      }
    ]);

    expect(container.querySelectorAll("[data-trace='scatter'] circle")).toHaveLength(3);

    const serialized = JSON.parse(chart.toJson()) as { data: unknown[] };
    expect(Array.isArray(serialized.data)).toBe(true);
    expect(serialized.data).toHaveLength(1);

    chart.destroy();
    expect(container.querySelector("svg.excelsior-chart-svg")).toBeNull();
    container.remove();
  });

  it("renders category axis ticks without duplicating labels", () => {
    const container = document.createElement("div");
    setContainerSize(container, 760, 320);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "bar",
          name: "Valor",
          x: ["Consultoria", "Licença"],
          y: [1200, 800]
        },
        {
          type: "bar",
          name: "Imposto",
          x: ["Consultoria", "Licença"],
          y: [120, 96]
        },
        {
          type: "bar",
          name: "Total",
          x: ["Consultoria", "Licença"],
          y: [1320, 896]
        }
      ],
      layout: {
        xAxis: {
          type: "category"
        }
      }
    });

    const axisTextValues = Array.from(container.querySelectorAll("text"))
      .map((node) => node.textContent?.trim() ?? "")
      .filter((value) => value.length > 0);
    expect(axisTextValues.filter((value) => value === "Consultoria")).toHaveLength(1);
    expect(axisTextValues.filter((value) => value === "Licença")).toHaveLength(1);

    chart.destroy();
    container.remove();
  });

  it("avoids overlapping y-axis labels for close values", () => {
    const container = document.createElement("div");
    setContainerSize(container, 780, 260);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "bar",
          name: "Valor",
          x: ["Consultoria", "Licença"],
          y: [1200, 800]
        },
        {
          type: "bar",
          name: "Imposto",
          x: ["Consultoria", "Licença"],
          y: [120, 96]
        },
        {
          type: "bar",
          name: "Total",
          x: ["Consultoria", "Licença"],
          y: [1320, 896]
        }
      ],
      layout: {
        xAxis: {
          type: "category"
        }
      }
    });

    const yLabels = Array.from(container.querySelectorAll<SVGTextElement>("text"))
      .filter((node) => (node.getAttribute("text-anchor") ?? "") === "end")
      .map((node) => ({
        text: node.textContent?.trim() ?? "",
        x: Number(node.getAttribute("x")),
        y: Number(node.getAttribute("y"))
      }))
      .filter(
        (entry) =>
          Number.isFinite(entry.x) &&
          Number.isFinite(entry.y) &&
          entry.x <= 80 &&
          /^-?\d+(?:\.\d+)?$/.test(entry.text)
      );

    const numericValues = yLabels.map((entry) => Number(entry.text));
    expect(numericValues.length).toBeGreaterThanOrEqual(2);
    expect(numericValues.some((value) => Math.abs(value - 1320) < 0.01)).toBe(true);
    const has96 = numericValues.some((value) => Math.abs(value - 96) < 0.01);
    const has120 = numericValues.some((value) => Math.abs(value - 120) < 0.01);
    expect(has96 && has120).toBe(false);

    for (let leftIndex = 0; leftIndex < yLabels.length; leftIndex += 1) {
      for (let rightIndex = leftIndex + 1; rightIndex < yLabels.length; rightIndex += 1) {
        const yDistance = Math.abs(yLabels[leftIndex].y - yLabels[rightIndex].y);
        expect(yDistance).toBeGreaterThanOrEqual(11);
      }
    }

    chart.destroy();
    container.remove();
  });

  it("throws controlled validation error for mismatched cartesian arrays", () => {
    const container = document.createElement("div");
    setContainerSize(container, 640, 360);
    document.body.append(container);

    expect(() =>
      createFigure(container, {
        data: [
          {
            type: "scatter",
            x: ["Jan", "Fev"],
            y: [120]
          }
        ]
      })
    ).toThrow(ChartConfigurationError);

    container.remove();
  });

  it("throws controlled error for a missing selector", () => {
    expect(() =>
      createFigure("#missing-chart-target", {
        data: [
          {
            type: "bar",
            x: ["A"],
            y: [1]
          }
        ]
      })
    ).toThrow(ChartConfigurationError);
  });

  it("preserves serializable frames on figure snapshots", () => {
    const container = document.createElement("div");
    setContainerSize(container, 600, 320);
    document.body.append(container);

    const chart = createFigure(container, {
      data: [
        {
          type: "line",
          x: ["Q1", "Q2"],
          y: [10, 20]
        }
      ],
      frames: [
        {
          name: "step-1",
          data: [
            {
              type: "line",
              x: ["Q1", "Q2"],
              y: [15, 30]
            }
          ]
        }
      ]
    });

    const snapshot = JSON.parse(chart.toJson()) as { frames: Array<{ name?: string }> };
    expect(snapshot.frames).toHaveLength(1);
    expect(snapshot.frames[0].name).toBe("step-1");

    chart.destroy();
    container.remove();
  });

  it("supports creating a chart from serialized JSON", () => {
    const container = document.createElement("div");
    setContainerSize(container, 640, 320);
    document.body.append(container);

    const chart = fromJson(
      container,
      JSON.stringify({
        data: [
          {
            type: "line",
            x: ["A", "B", "C"],
            y: [1, 3, 2]
          }
        ],
        layout: {
          title: "From JSON"
        },
        schemaVersion: "1.0.0"
      })
    );

    expect(container.querySelector("svg.excelsior-chart-svg")).not.toBeNull();
    expect(container.querySelector("text")?.textContent).toContain("From JSON");
    chart.destroy();
    container.remove();
  });

  it("rejects invalid serialized JSON", () => {
    expect(() => parseFigureFromJson("{invalid")).toThrow(ChartConfigurationError);
    expect(() => parseFigureFromJson(JSON.stringify({ layout: {} }))).toThrow(ChartConfigurationError);
    expect(() =>
      parseFigureFromJson(
        JSON.stringify({
          data: [{ type: "line", x: ["A"], y: [1] }],
          config: { unknownFlag: true }
        })
      )
    ).toThrow(ChartConfigurationError);
    expect(() =>
      parseFigureFromJson(
        JSON.stringify({
          data: [{ type: "line", x: ["A"], y: [1] }],
          unsupportedRootField: {}
        })
      )
    ).toThrow(ChartConfigurationError);
  });

  it("mounts canvas renderer when configured", () => {
    const container = document.createElement("div");
    setContainerSize(container, 640, 320);
    document.body.append(container);
    const originalGetContext = HTMLCanvasElement.prototype.getContext;
    const mockContext = {
      canvas: document.createElement("canvas"),
      clearRect: () => {},
      fillRect: () => {},
      strokeRect: () => {},
      fillText: () => {},
      beginPath: () => {},
      moveTo: () => {},
      lineTo: () => {},
      stroke: () => {},
      arc: () => {},
      closePath: () => {},
      fill: () => {},
      save: () => {},
      restore: () => {},
      translate: () => {},
      rotate: () => {},
      drawImage: () => {}
    } as unknown as CanvasRenderingContext2D;

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => mockContext
    });

    const chart = createFigure(container, {
      data: [
        {
          type: "scatter",
          x: ["A", "B"],
          y: [1, 2]
        }
      ],
      config: {
        renderer: "canvas"
      }
    });

    expect(container.querySelector("canvas")).not.toBeNull();
    expect(container.querySelector(".excelsior-chart-canvas-root")).not.toBeNull();
    chart.destroy();
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: originalGetContext
    });
    container.remove();
  });

  it("mounts webgl renderer when configured", () => {
    const container = document.createElement("div");
    setContainerSize(container, 640, 320);
    document.body.append(container);
    const originalGetContext = HTMLCanvasElement.prototype.getContext;

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
      drawArrays: () => {},
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
          type: "scatter",
          x: ["A", "B"],
          y: [1, 2]
        }
      ],
      config: {
        renderer: "webgl"
      }
    });

    expect(container.querySelector(".excelsior-chart-webgl-root")).not.toBeNull();
    chart.destroy();
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: originalGetContext
    });
    container.remove();
  });

  it("emits layout resized event with updated container dimensions", () => {
    const container = document.createElement("div");
    setContainerSize(container, 640, 320);
    document.body.append(container);

    const originalRaf = globalThis.requestAnimationFrame;
    const originalCancelRaf = globalThis.cancelAnimationFrame;
    const scheduledCallbacks: FrameRequestCallback[] = [];
    let frameId = 0;
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (callback: FrameRequestCallback) => {
        scheduledCallbacks.push(callback);
        frameId += 1;
        return frameId;
      }
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", {
      configurable: true,
      value: () => {}
    });

    try {
      const chart = createFigure(container, {
        data: [
          {
            type: "line",
            x: ["A", "B", "C"],
            y: [1, 3, 2]
          }
        ]
      });

      const resizedEvents: Array<{ width: number; height: number }> = [];
      chart.on("layout:resized", (payload) => resizedEvents.push(payload));

      setContainerSize(container, 880, 520);
      chart.resize();

      expect(scheduledCallbacks.length).toBeGreaterThan(0);
      while (scheduledCallbacks.length > 0) {
        const callback = scheduledCallbacks.shift();
        callback?.(0);
      }

      expect(resizedEvents.at(-1)).toEqual({ width: 880, height: 520 });
      chart.destroy();
    } finally {
      Object.defineProperty(globalThis, "requestAnimationFrame", {
        configurable: true,
        value: originalRaf
      });
      Object.defineProperty(globalThis, "cancelAnimationFrame", {
        configurable: true,
        value: originalCancelRaf
      });
      container.remove();
    }
  });

  it("disconnects resize observer and cancels pending frame on destroy", () => {
    const container = document.createElement("div");
    setContainerSize(container, 680, 360);
    document.body.append(container);

    const originalRaf = globalThis.requestAnimationFrame;
    const originalCancelRaf = globalThis.cancelAnimationFrame;
    const originalResizeObserver = globalThis.ResizeObserver;
    const cancelledIds: number[] = [];
    let scheduledFrameId = 0;
    let observeCalls = 0;
    let disconnectCalls = 0;

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: () => {
        scheduledFrameId += 1;
        return scheduledFrameId;
      }
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", {
      configurable: true,
      value: (id: number) => {
        cancelledIds.push(id);
      }
    });

    class MockResizeObserver {
      constructor(_callback: ResizeObserverCallback) {}

      observe(): void {
        observeCalls += 1;
      }

      disconnect(): void {
        disconnectCalls += 1;
      }
    }

    Object.defineProperty(globalThis, "ResizeObserver", {
      configurable: true,
      value: MockResizeObserver
    });

    try {
      const chart = createFigure(container, {
        data: [
          {
            type: "scatter",
            x: ["A", "B"],
            y: [2, 4]
          }
        ]
      });

      chart.resize();
      chart.destroy();

      expect(observeCalls).toBeGreaterThan(0);
      expect(disconnectCalls).toBe(1);
      expect(cancelledIds).toContain(scheduledFrameId);
      expect(() => chart.resize()).toThrow(ChartConfigurationError);
    } finally {
      Object.defineProperty(globalThis, "requestAnimationFrame", {
        configurable: true,
        value: originalRaf
      });
      Object.defineProperty(globalThis, "cancelAnimationFrame", {
        configurable: true,
        value: originalCancelRaf
      });
      Object.defineProperty(globalThis, "ResizeObserver", {
        configurable: true,
        value: originalResizeObserver
      });
      container.remove();
    }
  });
});
