import type { ChartAxis, ChartAxisType } from "../model/Layout";
import type { TracePoint } from "../model/Trace";

export const normalizeAxisType = (axis: ChartAxis | undefined, fallback: ChartAxisType): ChartAxisType => {
  const type = axis?.type;
  if (type === "linear" || type === "log" || type === "date" || type === "category" || type === "multicategory") {
    return type;
  }
  return fallback;
};

export const toAxisScalar = (value: TracePoint | unknown, axisType: ChartAxisType, fallback: number): number => {
  if (axisType === "category" || axisType === "multicategory") {
    return fallback;
  }
  if (axisType === "date") {
    if (value instanceof Date) {
      const timestamp = value.getTime();
      return Number.isFinite(timestamp) ? timestamp : fallback;
    }
    const parsed = Date.parse(String(value));
    return Number.isFinite(parsed) ? parsed : fallback;
  }
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return fallback;
  }
  if (axisType === "log" && numeric <= 0) {
    return fallback;
  }
  return numeric;
};

export const normalizeDomainForAxis = (domain: [number, number], axisType: ChartAxisType): [number, number] => {
  let [min, max] = domain;
  if (!Number.isFinite(min) || !Number.isFinite(max) || min === max) {
    min = axisType === "log" ? 1 : 0;
    max = axisType === "log" ? 10 : 1;
  }
  if (axisType === "log") {
    min = min <= 0 ? 1 : min;
    max = max <= min ? min * 10 : max;
  }
  return min < max ? [min, max] : [max, min];
};

export const getAxisTickCount = (axis: ChartAxis | undefined, fallback: number, min = 2, max = 16): number => {
  const requested = Number(axis?.tickCount);
  if (!Number.isFinite(requested)) {
    return clampInt(fallback, min, max);
  }
  return clampInt(requested, min, max);
};

export const formatAxisTick = (value: number, axis: ChartAxis | undefined, fallbackLabel?: string): string => {
  const type = normalizeAxisType(axis, "linear");
  const tickFormat = axis?.tickFormat?.trim().toLowerCase();
  if (fallbackLabel && (type === "category" || type === "multicategory" || tickFormat === "label")) {
    return fallbackLabel;
  }

  if (type === "date" || tickFormat === "date") {
    const date = new Date(value);
    if (!Number.isNaN(date.getTime())) {
      return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`;
    }
  }

  if (tickFormat?.startsWith("fixed:")) {
    const digits = clampInt(Number(tickFormat.split(":")[1]), 0, 10);
    return Number.isFinite(value) ? value.toFixed(digits) : "0";
  }
  if (tickFormat === "percent") {
    return `${(value * 100).toFixed(1)}%`;
  }
  if (tickFormat === "compact") {
    return compactNumber(value);
  }
  if (tickFormat === "currency") {
    return `R$ ${value.toFixed(2)}`;
  }

  if (!Number.isFinite(value)) {
    return "0";
  }
  const absolute = Math.abs(value);
  if (absolute >= 1000) {
    return value.toFixed(0);
  }
  if (absolute >= 100) {
    return value.toFixed(1);
  }
  return value.toFixed(2);
};

const compactNumber = (value: number): string => {
  if (!Number.isFinite(value)) {
    return "0";
  }
  const absolute = Math.abs(value);
  if (absolute >= 1_000_000_000) {
    return `${(value / 1_000_000_000).toFixed(1)}B`;
  }
  if (absolute >= 1_000_000) {
    return `${(value / 1_000_000).toFixed(1)}M`;
  }
  if (absolute >= 1_000) {
    return `${(value / 1_000).toFixed(1)}K`;
  }
  return value.toFixed(0);
};

const clampInt = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, Math.round(value)));

const pad = (value: number): string => String(value).padStart(2, "0");
