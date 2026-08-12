import { FigureValidator, TraceRegistry } from "../dist/index.js";
import { performance } from "node:perf_hooks";

const validator = new FigureValidator(new TraceRegistry());

const scenarios = [10_000, 100_000, 1_000_000];

console.log("=== @excelsior/charts benchmark ===");
console.log("Case | Normalize(ms) | JSON parse(ms)");

for (const points of scenarios) {
  const figure = createLargeFigure(points);

  const startNormalize = performance.now();
  validator.normalize(figure);
  const normalizeMs = performance.now() - startNormalize;

  const json = JSON.stringify(figure);
  const startParse = performance.now();
  JSON.parse(json);
  const parseMs = performance.now() - startParse;

  console.log(`${points.toLocaleString("en-US")} | ${normalizeMs.toFixed(2)} | ${parseMs.toFixed(2)}`);
}

function createLargeFigure(points) {
  const maxPointsPerTrace = 200_000;
  const traces = [];
  let offset = 0;
  let traceIndex = 0;

  while (offset < points) {
    const size = Math.min(maxPointsPerTrace, points - offset);
    const x = new Array(size);
    const y = new Array(size);
    for (let index = 0; index < size; index += 1) {
      const absoluteIndex = offset + index;
      x[index] = absoluteIndex;
      y[index] = Math.sin(absoluteIndex / 48) * 100 + Math.cos(absoluteIndex / 17) * 42;
    }

    traces.push({
        type: "line",
        name: `Series ${traceIndex + 1}`,
        x,
        y
      });
    offset += size;
    traceIndex += 1;
  }

  return {
    data: traces,
    config: {
      renderer: "svg",
      maxRenderPoints: 12000,
      maxInteractivePoints: 25000
    }
  };
}
