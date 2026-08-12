import type { BoxTrace, HeatmapTrace, HistogramTrace } from "../model/Trace";

export interface HistogramBin {
  index: number;
  min: number;
  max: number;
  count: number;
  label: string;
}

export interface BoxStats {
  min: number;
  max: number;
  q1: number;
  median: number;
  q3: number;
}

export interface HeatmapCell {
  row: number;
  col: number;
  value: number;
}

export interface HeatmapMatrix {
  rows: number;
  cols: number;
  values: HeatmapCell[];
  min: number;
  max: number;
}

export interface DensityPoint {
  x: number;
  y: number;
}

export interface ViolinProfile {
  min: number;
  max: number;
  points: DensityPoint[];
}

export const computeHistogramBins = (trace: HistogramTrace): HistogramBin[] => {
  const values = trace.values.map((value) => Number(value)).filter((value) => Number.isFinite(value));
  if (values.length === 0) {
    return [];
  }

  let min = Math.min(...values);
  let max = Math.max(...values);
  if (min === max) {
    min -= 0.5;
    max += 0.5;
  }

  const requestedBins = Number.isFinite(trace.bins) ? Number(trace.bins) : undefined;
  const binsCount = clampInt(requestedBins ?? Math.ceil(Math.sqrt(values.length)), 3, 64);
  const width = (max - min) / binsCount;
  const bins = Array.from({ length: binsCount }, (_, index) => ({
    index,
    min: min + index * width,
    max: index === binsCount - 1 ? max : min + (index + 1) * width,
    count: 0,
    label: ""
  }));

  for (const value of values) {
    const ratio = (value - min) / (max - min);
    const target = Math.min(binsCount - 1, Math.max(0, Math.floor(ratio * binsCount)));
    bins[target].count += 1;
  }

  for (const bin of bins) {
    bin.label = `${formatNumeric(bin.min)}-${formatNumeric(bin.max)}`;
  }

  return bins;
};

export const computeBoxStats = (trace: BoxTrace): BoxStats | null => {
  const sorted = trace.values
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value))
    .sort((a, b) => a - b);
  if (sorted.length === 0) {
    return null;
  }

  return {
    min: sorted[0],
    max: sorted[sorted.length - 1],
    q1: percentile(sorted, 0.25),
    median: percentile(sorted, 0.5),
    q3: percentile(sorted, 0.75)
  };
};

export const computeBoxOutliers = (trace: BoxTrace, factor = 1.5): number[] => {
  const stats = computeBoxStats(trace);
  if (!stats) {
    return [];
  }
  const lowerFence = stats.q1 - (stats.q3 - stats.q1) * factor;
  const upperFence = stats.q3 + (stats.q3 - stats.q1) * factor;
  return trace.values
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value) && (value < lowerFence || value > upperFence));
};

export const computeDensityCurve = (
  rawValues: number[],
  options?: { samples?: number; bandwidth?: number; cumulative?: boolean; sampleLimit?: number }
): DensityPoint[] => {
  const maxInput = clampInt(options?.sampleLimit ?? rawValues.length, 50, 500_000);
  const values = rawValues
    .slice(0, maxInput)
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value));
  if (values.length === 0) {
    return [];
  }
  const sorted = [...values].sort((a, b) => a - b);
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  if (min === max) {
    return [{ x: min, y: 1 }];
  }
  const sampleCount = clampInt(options?.samples ?? 64, 12, 256);
  const sigma = standardDeviation(sorted);
  const iqr = percentile(sorted, 0.75) - percentile(sorted, 0.25);
  const baseBandwidth = options?.bandwidth ?? Math.max(1e-6, 0.9 * Math.min(sigma, iqr / 1.34 || sigma) * sorted.length ** -0.2);
  const bandwidth = Math.max(1e-6, baseBandwidth);
  const points: DensityPoint[] = [];
  for (let index = 0; index < sampleCount; index += 1) {
    const ratio = sampleCount <= 1 ? 0 : index / (sampleCount - 1);
    const x = min + ratio * (max - min);
    let y = 0;
    for (const value of sorted) {
      const z = (x - value) / bandwidth;
      y += gaussianKernel(z);
    }
    y = y / (sorted.length * bandwidth);
    points.push({ x, y });
  }
  if (options?.cumulative) {
    let carry = 0;
    for (const point of points) {
      carry += point.y;
      point.y = carry;
    }
    if (carry > 0) {
      points.forEach((point) => {
        point.y = point.y / carry;
      });
    }
  }
  return points;
};

export const computeViolinProfile = (
  rawValues: number[],
  options?: { samples?: number; bandwidth?: number; sampleLimit?: number }
): ViolinProfile | null => {
  const numericValues = rawValues.map((value) => Number(value)).filter((value) => Number.isFinite(value));
  if (numericValues.length === 0) {
    return null;
  }
  const curve = computeDensityCurve(numericValues, {
    samples: options?.samples ?? 48,
    bandwidth: options?.bandwidth,
    sampleLimit: options?.sampleLimit
  });
  if (curve.length === 0) {
    return null;
  }
  const min = Math.min(...numericValues);
  const max = Math.max(...numericValues);
  const peak = Math.max(...curve.map((point) => point.y), Number.EPSILON);
  return {
    min,
    max,
    points: curve.map((point) => ({
      x: point.x,
      y: point.y / peak
    }))
  };
};

export const computeHeatmapMatrix = (trace: HeatmapTrace): HeatmapMatrix | null => {
  const rowCount = trace.z.length;
  if (rowCount === 0) {
    return null;
  }
  const colCount = trace.z[0]?.length ?? 0;
  if (colCount === 0) {
    return null;
  }

  const values: HeatmapCell[] = [];
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;

  for (let row = 0; row < rowCount; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      const raw = Number(trace.z[row][col]);
      const value = Number.isFinite(raw) ? raw : 0;
      values.push({ row, col, value });
      min = Math.min(min, value);
      max = Math.max(max, value);
    }
  }

  const forcedMin = Number.isFinite(trace.zMin) ? Number(trace.zMin) : min;
  const forcedMax = Number.isFinite(trace.zMax) ? Number(trace.zMax) : max;
  if (forcedMin === forcedMax) {
    return {
      rows: rowCount,
      cols: colCount,
      values,
      min: forcedMin - 0.5,
      max: forcedMax + 0.5
    };
  }

  return {
    rows: rowCount,
    cols: colCount,
    values,
    min: Math.min(forcedMin, forcedMax),
    max: Math.max(forcedMin, forcedMax)
  };
};

const clampInt = (value: number, min: number, max: number): number => {
  const asInt = Math.round(value);
  return Math.min(max, Math.max(min, asInt));
};

const percentile = (sortedValues: number[], ratio: number): number => {
  if (sortedValues.length === 1) {
    return sortedValues[0];
  }

  const position = (sortedValues.length - 1) * ratio;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) {
    return sortedValues[lower];
  }

  const weight = position - lower;
  return sortedValues[lower] * (1 - weight) + sortedValues[upper] * weight;
};

const standardDeviation = (values: number[]): number => {
  if (values.length <= 1) {
    return 0;
  }
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  const variance = values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, values.length - 1);
  return Math.sqrt(variance);
};

const gaussianKernel = (value: number): number => Math.exp((-0.5 * value * value) / Math.max(Number.EPSILON, 1)) / Math.sqrt(Math.PI * 2);

export const formatNumeric = (value: number): string => {
  if (!Number.isFinite(value)) {
    return "0";
  }
  const absolute = Math.abs(value);
  if (absolute >= 1000 || absolute === 0) {
    return value.toFixed(0);
  }
  if (absolute >= 100) {
    return value.toFixed(1);
  }
  return value.toFixed(2);
};
