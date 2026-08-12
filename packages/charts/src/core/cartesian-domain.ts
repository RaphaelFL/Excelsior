import type { ChartFigure } from "../model/Figure";
import type { CartesianTrace, ChartTrace, FinancialTrace } from "../model/Trace";
import { normalizeAxisType, normalizeDomainForAxis, toAxisScalar } from "./axis-utils";

export interface CartesianDomains {
  x: [number, number];
  y: [number, number];
  labels: string[];
}

export const isCartesianTrace = (trace: ChartTrace): trace is CartesianTrace =>
  trace.type === "scatter" || trace.type === "line" || trace.type === "bar" || trace.type === "area";

export const isFinancialTrace = (trace: ChartTrace): trace is FinancialTrace => trace.type === "candlestick" || trace.type === "ohlc";

export const buildCartesianDomains = (figure: ChartFigure): CartesianDomains | null => {
  const visible = figure.data.filter((trace) => trace.visible !== false);
  const cartesianTraces = visible.filter(isCartesianTrace);
  const financialTraces = visible.filter(isFinancialTrace);
  if (cartesianTraces.length === 0 && financialTraces.length === 0) {
    return null;
  }

  const xAxisType = normalizeAxisType(figure.layout.xAxis, "category");
  const yAxisType = normalizeAxisType(figure.layout.yAxis, "linear");
  const labels = resolveDomainLabels(cartesianTraces, financialTraces, xAxisType);
  const pointCount = Math.max(1, labels.length);

  const xValues =
    xAxisType === "category" || xAxisType === "multicategory"
      ? Array.from({ length: pointCount }, (_, index) => index)
      : [
          ...cartesianTraces.flatMap((trace) => trace.x.map((value, index) => toAxisScalar(value, xAxisType, index))),
          ...financialTraces.flatMap((trace) => trace.x.map((value, index) => toAxisScalar(value, xAxisType, index)))
        ].filter((value) => Number.isFinite(value));

  let defaultXMin = xValues.length > 0 ? Math.min(...xValues) : 0;
  let defaultXMax = xValues.length > 0 ? Math.max(...xValues) : Math.max(1, pointCount - 1);
  if (defaultXMin === defaultXMax) {
    defaultXMin -= 0.5;
    defaultXMax += 0.5;
  }

  let yValues = [
    ...cartesianTraces.flatMap((trace) => trace.y),
    ...financialTraces.flatMap((trace) => [...trace.low, ...trace.high, ...trace.open, ...trace.close])
  ].filter((value) => Number.isFinite(value));
  if (yAxisType === "log") {
    yValues = yValues.filter((value) => value > 0);
  }
  let defaultYMin = yValues.length > 0 ? Math.min(...yValues) : 0;
  let defaultYMax = yValues.length > 0 ? Math.max(...yValues) : 1;

  if (defaultYMin === defaultYMax) {
    defaultYMin -= 1;
    defaultYMax += 1;
  }
  if (yAxisType !== "log") {
    if (defaultYMin > 0) {
      defaultYMin = 0;
    }
    if (defaultYMax < 0) {
      defaultYMax = 0;
    }
  } else {
    defaultYMin = defaultYMin <= 0 ? 1 : defaultYMin;
    defaultYMax = defaultYMax <= defaultYMin ? defaultYMin * 10 : defaultYMax;
  }

  const xMin = Number.isFinite(figure.layout.xAxis.min) ? Number(figure.layout.xAxis.min) : defaultXMin;
  const xMax = Number.isFinite(figure.layout.xAxis.max) ? Number(figure.layout.xAxis.max) : defaultXMax;
  const yMin = Number.isFinite(figure.layout.yAxis.min) ? Number(figure.layout.yAxis.min) : defaultYMin;
  const yMax = Number.isFinite(figure.layout.yAxis.max) ? Number(figure.layout.yAxis.max) : defaultYMax;

  return {
    x: normalizeDomainForAxis(normalizeDomain(xMin, xMax, defaultXMin, defaultXMax), xAxisType),
    y: normalizeDomainForAxis(normalizeDomain(yMin, yMax, defaultYMin, defaultYMax), yAxisType),
    labels
  };
};

const resolveDomainLabels = (
  cartesianTraces: CartesianTrace[],
  financialTraces: FinancialTrace[],
  xAxisType: "linear" | "log" | "date" | "category" | "multicategory"
): string[] => {
  const allXSets: Array<Array<unknown>> = [...cartesianTraces.map((trace) => trace.x), ...financialTraces.map((trace) => trace.x)];
  const labelSource = allXSets.reduce<Array<unknown>>((selected, candidate) => {
    if (candidate.length > selected.length) {
      return candidate;
    }
    return selected;
  }, allXSets[0] ?? []);
  if (xAxisType === "date") {
    return labelSource.map((value) => normalizeDateLabel(value));
  }
  return labelSource.map((value) => normalizeTracePointLabel(value));
};

const normalizeTracePointLabel = (value: unknown): string => {
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value ?? "");
};

const normalizeDateLabel = (value: unknown): string => {
  if (value instanceof Date) {
    return formatDateLabel(value.getTime());
  }
  const timestamp = Date.parse(String(value));
  if (!Number.isFinite(timestamp)) {
    return String(value ?? "");
  }
  return formatDateLabel(timestamp);
};

const formatDateLabel = (timestamp: number): string => {
  const date = new Date(timestamp);
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
};

const normalizeDomain = (min: number, max: number, fallbackMin: number, fallbackMax: number): [number, number] => {
  if (!Number.isFinite(min) || !Number.isFinite(max) || min === max) {
    return [fallbackMin, fallbackMax];
  }
  return min < max ? [min, max] : [max, min];
};
