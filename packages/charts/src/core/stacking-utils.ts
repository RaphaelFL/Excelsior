import type { CartesianTrace } from "../model/Trace";

export interface StackedPoint {
  base: number;
  top: number;
  value: number;
}

export interface StackedSeries {
  stacked: boolean;
  points: StackedPoint[];
}

export interface CartesianStackEntry {
  index: number;
  trace: CartesianTrace;
}

export const buildCartesianStackContext = (entries: CartesianStackEntry[]): Map<number, StackedSeries> => {
  const result = new Map<number, StackedSeries>();
  const stackGroups = new Map<string, CartesianStackEntry[]>();

  entries.forEach((entry) => {
    if ((entry.trace.type !== "bar" && entry.trace.type !== "area") || !entry.trace.stackGroup?.trim()) {
      result.set(entry.index, {
        stacked: false,
        points: entry.trace.y.map((value) => ({ base: 0, top: Number(value), value: Number(value) }))
      });
      return;
    }
    const key = `${entry.trace.type}:${entry.trace.stackGroup.trim()}`;
    const bucket = stackGroups.get(key);
    if (bucket) {
      bucket.push(entry);
    } else {
      stackGroups.set(key, [entry]);
    }
  });

  for (const groupEntries of stackGroups.values()) {
    const maxPoints = Math.max(0, ...groupEntries.map((entry) => entry.trace.y.length));
    const positiveCarry = new Array<number>(maxPoints).fill(0);
    const negativeCarry = new Array<number>(maxPoints).fill(0);
    const totalsAbs = new Array<number>(maxPoints).fill(0);

    groupEntries.forEach((entry) => {
      entry.trace.y.forEach((rawValue, pointIndex) => {
        const numeric = Number(rawValue);
        if (!Number.isFinite(numeric)) {
          return;
        }
        totalsAbs[pointIndex] += Math.abs(numeric);
      });
    });

    groupEntries.forEach((entry) => {
      const points: StackedPoint[] = [];
      for (let pointIndex = 0; pointIndex < maxPoints; pointIndex += 1) {
        const rawValue = Number(entry.trace.y[pointIndex] ?? 0);
        const value = Number.isFinite(rawValue) ? rawValue : 0;
        const denominator = totalsAbs[pointIndex] <= 0 ? 1 : totalsAbs[pointIndex];
        const normalizedValue = entry.trace.normalizeStack ? value / denominator : value;
        if (normalizedValue >= 0) {
          const base = positiveCarry[pointIndex];
          const top = base + normalizedValue;
          positiveCarry[pointIndex] = top;
          points.push({ base, top, value: normalizedValue });
        } else {
          const base = negativeCarry[pointIndex];
          const top = base + normalizedValue;
          negativeCarry[pointIndex] = top;
          points.push({ base, top, value: normalizedValue });
        }
      }
      result.set(entry.index, {
        stacked: true,
        points
      });
    });
  }

  return result;
};
