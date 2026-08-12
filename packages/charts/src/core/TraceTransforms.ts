import type {
  BoxTrace,
  CartesianTrace,
  ChartTrace,
  DensityTrace,
  FinancialTrace,
  HistogramTrace,
  PolarTrace,
  TernaryTrace,
  TracePoint,
  TraceTransform,
  ViolinTrace
} from "../model/Trace";

export const applyTraceTransforms = (trace: ChartTrace): ChartTrace => {
  if (!Array.isArray((trace as { transforms?: unknown }).transforms) || (trace as { transforms?: unknown[] }).transforms?.length === 0) {
    return trace;
  }

  if (isCartesianTrace(trace)) {
    return applyCartesianTransforms(trace);
  }
  if (isFinancialTrace(trace)) {
    return applyFinancialTransforms(trace);
  }
  if (isHistogramTrace(trace)) {
    return applyValueTransforms(trace);
  }
  if (isBoxTrace(trace)) {
    return applyValueTransforms(trace);
  }
  if (isViolinTrace(trace)) {
    return applyValueTransforms(trace);
  }
  if (isDensityTrace(trace)) {
    return applyValueTransforms(trace);
  }
  if (isPolarTrace(trace)) {
    return applyPolarTransforms(trace);
  }
  if (isTernaryTrace(trace)) {
    return applyTernaryTransforms(trace);
  }
  return trace;
};

const applyCartesianTransforms = (trace: CartesianTrace): CartesianTrace => {
  let points = trace.x.map((x, index) => ({ x, y: Number(trace.y[index]) }));
  for (const transform of trace.transforms ?? []) {
    points = runCartesianTransform(points, transform);
  }
  return {
    ...trace,
    x: points.map((point) => point.x),
    y: points.map((point) => point.y)
  };
};

const applyFinancialTransforms = (trace: FinancialTrace): FinancialTrace => {
  let points = trace.x.map((x, index) => ({
    x,
    open: Number(trace.open[index]),
    high: Number(trace.high[index]),
    low: Number(trace.low[index]),
    close: Number(trace.close[index])
  }));

  for (const transform of trace.transforms ?? []) {
    if (transform.type === "sort") {
      const direction = transform.direction === "desc" ? -1 : 1;
      points = [...points].sort((left, right) => compareScalar(left.x, right.x) * direction);
      continue;
    }
    if (transform.type === "filter") {
      points = points.filter((point) => {
        const value = transform.field === "x" ? toScalar(point.x) : point.close;
        if (transform.min !== undefined && value < transform.min) {
          return false;
        }
        if (transform.max !== undefined && value > transform.max) {
          return false;
        }
        return true;
      });
      continue;
    }
    if (transform.type === "normalize") {
      const base = transform.strategy === "sum" ? points.reduce((sum, point) => sum + Math.abs(point.close), 0) : Math.max(...points.map((point) => Math.abs(point.close)), 1);
      const safeBase = base === 0 ? 1 : base;
      points = points.map((point) => ({
        ...point,
        open: point.open / safeBase,
        high: point.high / safeBase,
        low: point.low / safeBase,
        close: point.close / safeBase
      }));
      continue;
    }
    if (transform.type === "cumulative") {
      let carry = 0;
      points = points.map((point) => {
        carry += point.close;
        return {
          ...point,
          close: carry
        };
      });
      continue;
    }
    if (transform.type === "percent") {
      const total = points.reduce((sum, point) => sum + point.close, 0);
      const safeTotal = total === 0 ? 1 : total;
      points = points.map((point) => ({
        ...point,
        close: (point.close / safeTotal) * 100
      }));
    }
  }

  return {
    ...trace,
    x: points.map((point) => point.x),
    open: points.map((point) => point.open),
    high: points.map((point) => point.high),
    low: points.map((point) => point.low),
    close: points.map((point) => point.close)
  };
};

const applyValueTransforms = <TTrace extends HistogramTrace | BoxTrace | ViolinTrace | DensityTrace>(trace: TTrace): TTrace => {
  let values = trace.values.map((value) => Number(value));
  for (const transform of trace.transforms ?? []) {
    values = runValueTransform(values, transform);
  }
  return {
    ...trace,
    values
  };
};

const applyPolarTransforms = (trace: PolarTrace): PolarTrace => {
  let points = trace.theta.map((theta, index) => ({ theta, r: Number(trace.r[index]) }));
  for (const transform of trace.transforms ?? []) {
    if (transform.type === "sort") {
      const direction = transform.direction === "desc" ? -1 : 1;
      points = [...points].sort((left, right) => {
        const leftValue = transform.by === "x" ? toScalar(left.theta) : left.r;
        const rightValue = transform.by === "x" ? toScalar(right.theta) : right.r;
        return (leftValue - rightValue) * direction;
      });
      continue;
    }
    if (transform.type === "filter") {
      points = points.filter((point) => {
        const value = transform.field === "x" ? toScalar(point.theta) : point.r;
        if (transform.min !== undefined && value < transform.min) {
          return false;
        }
        if (transform.max !== undefined && value > transform.max) {
          return false;
        }
        return true;
      });
      continue;
    }
    if (transform.type === "normalize") {
      const denominator = transform.strategy === "sum" ? points.reduce((sum, point) => sum + Math.abs(point.r), 0) : Math.max(...points.map((point) => Math.abs(point.r)), 1);
      const safe = denominator === 0 ? 1 : denominator;
      points = points.map((point) => ({ ...point, r: point.r / safe }));
      continue;
    }
    if (transform.type === "cumulative") {
      let carry = 0;
      points = points.map((point) => {
        carry += point.r;
        return { ...point, r: carry };
      });
      continue;
    }
    if (transform.type === "percent") {
      const total = points.reduce((sum, point) => sum + point.r, 0);
      const safe = total === 0 ? 1 : total;
      points = points.map((point) => ({ ...point, r: (point.r / safe) * 100 }));
    }
  }

  return {
    ...trace,
    theta: points.map((point) => point.theta),
    r: points.map((point) => point.r)
  };
};

const applyTernaryTransforms = (trace: TernaryTrace): TernaryTrace => {
  let points = trace.a.map((a, index) => ({ a: Number(a), b: Number(trace.b[index]), c: Number(trace.c[index]) }));
  for (const transform of trace.transforms ?? []) {
    if (transform.type === "sort") {
      const direction = transform.direction === "desc" ? -1 : 1;
      points = [...points].sort((left, right) => {
        const leftValue = transform.by === "x" ? left.a : left.c;
        const rightValue = transform.by === "x" ? right.a : right.c;
        return (leftValue - rightValue) * direction;
      });
      continue;
    }
    if (transform.type === "filter") {
      points = points.filter((point) => {
        const value = transform.field === "x" ? point.a : point.c;
        if (transform.min !== undefined && value < transform.min) {
          return false;
        }
        if (transform.max !== undefined && value > transform.max) {
          return false;
        }
        return true;
      });
    }
  }

  return {
    ...trace,
    a: points.map((point) => point.a),
    b: points.map((point) => point.b),
    c: points.map((point) => point.c)
  };
};

const runCartesianTransform = (
  points: Array<{ x: TracePoint; y: number }>,
  transform: TraceTransform
): Array<{ x: TracePoint; y: number }> => {
  if (transform.type === "sort") {
    const direction = transform.direction === "desc" ? -1 : 1;
    return [...points].sort((left, right) => {
      const leftValue = transform.by === "x" ? toScalar(left.x) : left.y;
      const rightValue = transform.by === "x" ? toScalar(right.x) : right.y;
      return (leftValue - rightValue) * direction;
    });
  }

  if (transform.type === "filter") {
    return points.filter((point) => {
      const value = transform.field === "x" ? toScalar(point.x) : point.y;
      if (transform.min !== undefined && value < transform.min) {
        return false;
      }
      if (transform.max !== undefined && value > transform.max) {
        return false;
      }
      return true;
    });
  }

  if (transform.type === "normalize") {
    const denominator =
      transform.strategy === "sum"
        ? points.reduce((sum, point) => sum + Math.abs(point.y), 0)
        : Math.max(...points.map((point) => Math.abs(point.y)), 1);
    const safe = denominator === 0 ? 1 : denominator;
    return points.map((point) => ({ ...point, y: point.y / safe }));
  }

  if (transform.type === "group") {
    const aggregate = transform.aggregate ?? "sum";
    if (transform.by === "x") {
      const grouped = new Map<string, { x: TracePoint; values: number[] }>();
      points.forEach((point) => {
        const key = serializeKey(point.x);
        const current = grouped.get(key);
        if (current) {
          current.values.push(point.y);
        } else {
          grouped.set(key, { x: point.x, values: [point.y] });
        }
      });
      return Array.from(grouped.values()).map((group) => ({
        x: group.x,
        y: aggregateNumbers(group.values, aggregate)
      }));
    }

    const grouped = new Map<string, { y: number; values: number[] }>();
    points.forEach((point) => {
      const key = serializeKey(point.y);
      const current = grouped.get(key);
      if (current) {
        current.values.push(toScalar(point.x));
      } else {
        grouped.set(key, { y: point.y, values: [toScalar(point.x)] });
      }
    });
    return Array.from(grouped.values()).map((group, index) => ({
      x: index,
      y: group.y
    }));
  }

  if (transform.type === "aggregate") {
    if (points.length === 0) {
      return points;
    }
    return [
      {
        x: points[0].x,
        y: aggregateNumbers(points.map((point) => point.y), transform.op)
      }
    ];
  }

  if (transform.type === "bin") {
    const bins = clampInt(transform.bins ?? Math.ceil(Math.sqrt(points.length)), 2, 200);
    if (transform.field === "y") {
      return computeBins(points.map((point) => point.y), bins).map((bin) => ({
        x: bin.center,
        y: bin.count
      }));
    }
    return computeBins(points.map((point) => toScalar(point.x)), bins).map((bin) => ({
      x: bin.center,
      y: bin.count
    }));
  }

  if (transform.type === "stack") {
    let carry = 0;
    return points.map((point) => {
      const nextY = transform.strategy === "offset" ? carry + point.y : point.y + carry;
      carry = nextY;
      return {
        ...point,
        y: nextY
      };
    });
  }

  if (transform.type === "cumulative") {
    let carry = 0;
    return points.map((point) => {
      carry += point.y;
      return { ...point, y: carry };
    });
  }

  if (transform.type === "percent") {
    const total = points.reduce((sum, point) => sum + point.y, 0);
    const safe = total === 0 ? 1 : total;
    return points.map((point) => ({ ...point, y: (point.y / safe) * 100 }));
  }

  return points;
};

const runValueTransform = (values: number[], transform: TraceTransform): number[] => {
  if (transform.type === "sort") {
    const direction = transform.direction === "desc" ? -1 : 1;
    return [...values].sort((left, right) => (left - right) * direction);
  }
  if (transform.type === "filter") {
    return values.filter((value) => {
      if (transform.min !== undefined && value < transform.min) {
        return false;
      }
      if (transform.max !== undefined && value > transform.max) {
        return false;
      }
      return true;
    });
  }
  if (transform.type === "normalize") {
    const denominator =
      transform.strategy === "sum" ? values.reduce((sum, value) => sum + Math.abs(value), 0) : Math.max(...values.map((value) => Math.abs(value)), 1);
    const safe = denominator === 0 ? 1 : denominator;
    return values.map((value) => value / safe);
  }
  if (transform.type === "group") {
    const grouped = new Map<string, number[]>();
    values.forEach((value) => {
      const key = serializeKey(value);
      const bucket = grouped.get(key);
      if (bucket) {
        bucket.push(value);
      } else {
        grouped.set(key, [value]);
      }
    });
    const aggregate = transform.aggregate ?? "sum";
    return Array.from(grouped.values()).map((bucket) => aggregateNumbers(bucket, aggregate));
  }
  if (transform.type === "aggregate") {
    return values.length > 0 ? [aggregateNumbers(values, transform.op)] : values;
  }
  if (transform.type === "bin") {
    const bins = clampInt(transform.bins ?? Math.ceil(Math.sqrt(values.length)), 2, 200);
    return computeBins(values, bins).map((bin) => bin.count);
  }
  if (transform.type === "stack") {
    let carry = 0;
    return values.map((value) => {
      const next = transform.strategy === "offset" ? carry + value : value + carry;
      carry = next;
      return next;
    });
  }
  if (transform.type === "cumulative") {
    let carry = 0;
    return values.map((value) => {
      carry += value;
      return carry;
    });
  }
  if (transform.type === "percent") {
    const total = values.reduce((sum, value) => sum + value, 0);
    const safe = total === 0 ? 1 : total;
    return values.map((value) => (value / safe) * 100);
  }
  return values;
};

const toScalar = (value: unknown): number => {
  if (value instanceof Date) {
    return value.getTime();
  }
  if (typeof value === "number") {
    return value;
  }
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : 0;
};

const compareScalar = (left: unknown, right: unknown): number => {
  const leftValue = toScalar(left);
  const rightValue = toScalar(right);
  if (leftValue === rightValue) {
    return 0;
  }
  return leftValue > rightValue ? 1 : -1;
};

const serializeKey = (value: unknown): string => {
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
};

const aggregateNumbers = (values: number[], op: "sum" | "avg" | "min" | "max" | "count"): number => {
  if (op === "count") {
    return values.length;
  }
  if (values.length === 0) {
    return 0;
  }
  if (op === "min") {
    return Math.min(...values);
  }
  if (op === "max") {
    return Math.max(...values);
  }
  const sum = values.reduce((total, value) => total + value, 0);
  if (op === "avg") {
    return sum / Math.max(1, values.length);
  }
  return sum;
};

const computeBins = (values: number[], bins: number): Array<{ center: number; count: number }> => {
  const numeric = values.filter((value) => Number.isFinite(value));
  if (numeric.length === 0) {
    return [];
  }
  let min = Math.min(...numeric);
  let max = Math.max(...numeric);
  if (min === max) {
    min -= 0.5;
    max += 0.5;
  }
  const width = (max - min) / bins;
  const result = Array.from({ length: bins }, (_, index) => ({
    center: min + width * (index + 0.5),
    count: 0
  }));
  numeric.forEach((value) => {
    const ratio = (value - min) / (max - min);
    const target = Math.min(bins - 1, Math.max(0, Math.floor(ratio * bins)));
    result[target].count += 1;
  });
  return result;
};

const clampInt = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? Math.round(value) : min;
  return Math.min(max, Math.max(min, numeric));
};

const isCartesianTrace = (trace: ChartTrace): trace is CartesianTrace =>
  trace.type === "scatter" || trace.type === "line" || trace.type === "bar" || trace.type === "area";

const isFinancialTrace = (trace: ChartTrace): trace is FinancialTrace => trace.type === "candlestick" || trace.type === "ohlc";

const isHistogramTrace = (trace: ChartTrace): trace is HistogramTrace => trace.type === "histogram";

const isBoxTrace = (trace: ChartTrace): trace is BoxTrace => trace.type === "box";

const isViolinTrace = (trace: ChartTrace): trace is ViolinTrace => trace.type === "violin";

const isDensityTrace = (trace: ChartTrace): trace is DensityTrace => trace.type === "density" || trace.type === "distribution";

const isPolarTrace = (trace: ChartTrace): trace is PolarTrace => trace.type === "polar";

const isTernaryTrace = (trace: ChartTrace): trace is TernaryTrace => trace.type === "ternary";
