import { DEFAULT_CHART_CONFIG, type ChartConfig } from "../model/Config";
import type {
  ChartFigure,
  ChartFigureInput,
  ChartFigureMetadata,
  ChartFrame,
  ChartSelectionState
} from "../model/Figure";
import { DEFAULT_CHART_LAYOUT, type ChartLayout } from "../model/Layout";
import type { ChartTrace } from "../model/Trace";
import { TraceRegistry } from "./TraceRegistry";
import { ChartConfigurationError } from "./chart-errors";
import { applyTraceTransforms } from "./TraceTransforms";
import { applyThemeAndTemplate } from "./themes";

const MAX_VALIDATION_DEPTH = 24;
const MAX_ARRAY_ITEMS = 200_000;
const MAX_SELECTION_POINTS = 5_000;

export class FigureValidator {
  constructor(private readonly traceRegistry: TraceRegistry) {}

  normalize(input: ChartFigureInput): ChartFigure {
    this.assertSerializableValue(input, "figure", 0);

    if (!Array.isArray(input.data)) {
      throw new ChartConfigurationError("CHART_DATA_INVALID", "Figure data must be an array of traces.");
    }

    const data = input.data.map((trace, index) => this.normalizeTrace(trace, index));
    const layout = this.mergeLayout(input.layout);
    const config = this.mergeConfig(input.config);
    const frames = this.normalizeFrames(input.frames);
    const selection = this.normalizeSelection(input.selection);
    const metadata = this.sanitizeMetadata(input.metadata);

    return {
      data,
      layout,
      config,
      frames,
      selection,
      metadata,
      schemaVersion: input.schemaVersion?.trim() || "1.0.0"
    };
  }

  private normalizeTrace(trace: ChartTrace, index: number): ChartTrace {
    if (!trace || typeof trace !== "object") {
      throw new ChartConfigurationError("CHART_TRACE_INVALID", `Trace at index ${index} must be an object.`);
    }

    const type = String((trace as { type?: unknown }).type ?? "").trim();
    if (!type) {
      throw new ChartConfigurationError("CHART_TRACE_TYPE_REQUIRED", `Trace at index ${index} requires a type.`);
    }

    const nextTrace = applyTraceTransforms(structuredClone(trace) as ChartTrace);
    const definition = this.traceRegistry.resolve(type);
    definition.validate(nextTrace);
    if ("name" in nextTrace && typeof nextTrace.name === "string") {
      nextTrace.name = sanitizeText(nextTrace.name);
    }
    if (nextTrace.visible === undefined) {
      nextTrace.visible = true;
    }

    return nextTrace;
  }

  private mergeLayout(partial?: Partial<ChartLayout>): ChartLayout {
    const sanitizeShape = (shape: ChartLayout["shapes"][number]): ChartLayout["shapes"][number] => ({
      ...shape,
      type: shape.type,
      stroke: shape.stroke ? sanitizeColor(shape.stroke) : shape.stroke,
      fill: shape.fill ? sanitizeColor(shape.fill) : shape.fill
    });
    const sanitizeAnnotation = (annotation: ChartLayout["annotations"][number]): ChartLayout["annotations"][number] => ({
      ...annotation,
      text: sanitizeText(annotation.text),
      color: annotation.color ? sanitizeColor(annotation.color) : annotation.color
    });
    const sanitizeImage = (image: ChartLayout["images"][number]): ChartLayout["images"][number] => ({
      ...image,
      source: sanitizeImageSource(image.source)
    });

    const mergedLayout: ChartLayout = {
      ...DEFAULT_CHART_LAYOUT,
      ...partial,
      margin: {
        ...DEFAULT_CHART_LAYOUT.margin,
        ...(partial?.margin ?? {})
      },
      xAxis: {
        ...DEFAULT_CHART_LAYOUT.xAxis,
        ...(partial?.xAxis ?? {})
      },
      xAxis2: {
        ...DEFAULT_CHART_LAYOUT.xAxis2,
        ...(partial?.xAxis2 ?? {})
      },
      yAxis: {
        ...DEFAULT_CHART_LAYOUT.yAxis,
        ...(partial?.yAxis ?? {})
      },
      yAxis2: {
        ...DEFAULT_CHART_LAYOUT.yAxis2,
        ...(partial?.yAxis2 ?? {})
      },
      legend: {
        ...DEFAULT_CHART_LAYOUT.legend,
        ...(partial?.legend ?? {})
      },
      subplots: partial?.subplots
        ? {
            rows: clampInt(partial.subplots.rows ?? 1, 1, 16),
            cols: clampInt(partial.subplots.cols ?? 1, 1, 16),
            gapX: clampNumber(partial.subplots.gapX ?? 14, 0, 120),
            gapY: clampNumber(partial.subplots.gapY ?? 14, 0, 120),
            syncZoom: partial.subplots.syncZoom ?? true
          }
        : undefined,
      shapes: (partial?.shapes ?? DEFAULT_CHART_LAYOUT.shapes).map((shape) => sanitizeShape(shape)),
      annotations: (partial?.annotations ?? DEFAULT_CHART_LAYOUT.annotations).map((annotation) => sanitizeAnnotation(annotation)),
      images: (partial?.images ?? DEFAULT_CHART_LAYOUT.images).map((image) => sanitizeImage(image)),
      theme: partial?.theme ? sanitizeText(partial.theme) : undefined,
      template: partial?.template ? sanitizeText(partial.template) : undefined,
      title: partial?.title ? sanitizeText(partial.title) : DEFAULT_CHART_LAYOUT.title
    };
    return applyThemeAndTemplate(mergedLayout);
  }

  private mergeConfig(partial?: Partial<ChartConfig>): ChartConfig {
    const hoverMode = partial?.hoverMode === "x" ? "x" : "point";
    const maxRenderPoints = clampInt(Number(partial?.maxRenderPoints ?? DEFAULT_CHART_CONFIG.maxRenderPoints), 500, 500_000);
    const maxInteractivePoints = clampInt(
      Number(partial?.maxInteractivePoints ?? DEFAULT_CHART_CONFIG.maxInteractivePoints),
      100,
      1_000_000
    );
    const frameDurationMs = clampInt(Number(partial?.frameDurationMs ?? DEFAULT_CHART_CONFIG.frameDurationMs), 16, 60_000);
    const maxDensitySamples = clampInt(Number(partial?.maxDensitySamples ?? DEFAULT_CHART_CONFIG.maxDensitySamples), 100, 500_000);
    return {
      ...DEFAULT_CHART_CONFIG,
      ...(partial ?? {}),
      hoverMode,
      maxRenderPoints,
      maxInteractivePoints,
      spatialHover: partial?.spatialHover ?? DEFAULT_CHART_CONFIG.spatialHover,
      fullscreen: partial?.fullscreen ?? DEFAULT_CHART_CONFIG.fullscreen,
      accessibleTable: partial?.accessibleTable ?? DEFAULT_CHART_CONFIG.accessibleTable,
      ariaDescription: partial?.ariaDescription ? sanitizeText(partial.ariaDescription) : DEFAULT_CHART_CONFIG.ariaDescription,
      highContrast: partial?.highContrast ?? DEFAULT_CHART_CONFIG.highContrast,
      frameDurationMs,
      webglFallback: partial?.webglFallback ?? DEFAULT_CHART_CONFIG.webglFallback,
      syncSubplotZoom: partial?.syncSubplotZoom ?? DEFAULT_CHART_CONFIG.syncSubplotZoom,
      maxDensitySamples
    };
  }

  private sanitizeMetadata(metadata: ChartFigureMetadata | undefined): ChartFigureMetadata {
    if (!metadata) {
      return {};
    }

    const next: ChartFigureMetadata = {};
    for (const [rawKey, rawValue] of Object.entries(metadata)) {
      const key = sanitizeText(rawKey);
      if (!key) {
        continue;
      }
      next[key] = typeof rawValue === "string" ? sanitizeText(rawValue) : rawValue;
    }

    return next;
  }

  private normalizeFrames(frames: ChartFrame[] | undefined): ChartFrame[] {
    if (!frames || frames.length === 0) {
      return [];
    }
    return frames.map((frame) => {
      const normalizedData = Array.isArray(frame.data)
        ? frame.data.map((trace, index) => this.normalizeTrace(trace, index))
        : undefined;
      return {
        name: typeof frame.name === "string" ? sanitizeText(frame.name) : undefined,
        data: normalizedData,
        layout: frame.layout ? structuredClone(frame.layout) : undefined,
        metadata: this.sanitizeMetadata(frame.metadata)
      };
    });
  }

  private normalizeSelection(selection: ChartSelectionState | null | undefined): ChartSelectionState | null {
    if (!selection) {
      return null;
    }

    const mode: ChartSelectionState["mode"] =
      selection.mode === "rect" || selection.mode === "lasso" || selection.mode === "click" ? selection.mode : "click";
    const rawPoints = Array.isArray(selection.points) ? selection.points : [];
    if (rawPoints.length > MAX_SELECTION_POINTS) {
      throw new ChartConfigurationError(
        "CHART_SELECTION_TOO_LARGE",
        `Selection exceeds the maximum supported size of ${MAX_SELECTION_POINTS} points.`
      );
    }

    const points = rawPoints
      .map((point) => {
        const traceIndex = Number(point.traceIndex);
        const pointIndex = Number(point.pointIndex);
        if (!Number.isFinite(traceIndex) || !Number.isFinite(pointIndex)) {
          return null;
        }
        return {
          traceIndex,
          pointIndex,
          traceName: sanitizeText(String(point.traceName ?? "Trace")),
          x: sanitizeText(String(point.x ?? "")),
          y: sanitizeText(String(point.y ?? ""))
        };
      })
      .filter((point): point is NonNullable<typeof point> => point !== null);

    const updatedAtSource = selection.updatedAt && !Number.isNaN(Date.parse(selection.updatedAt)) ? selection.updatedAt : undefined;
    return {
      mode,
      points,
      updatedAt: updatedAtSource ?? new Date().toISOString()
    };
  }

  private assertSerializableValue(value: unknown, path: string, depth: number): void {
    if (depth > MAX_VALIDATION_DEPTH) {
      throw new ChartConfigurationError("CHART_INPUT_TOO_DEEP", `Figure input is too deep near '${path}'.`);
    }

    if (typeof value === "function" || typeof value === "symbol") {
      throw new ChartConfigurationError("CHART_INPUT_UNSAFE", `Unsupported value '${typeof value}' found at '${path}'.`);
    }

    if (!value || typeof value !== "object") {
      return;
    }

    if (Array.isArray(value)) {
      if (value.length > MAX_ARRAY_ITEMS) {
        throw new ChartConfigurationError(
          "CHART_INPUT_TOO_LARGE",
          `Array at '${path}' exceeds the supported limit of ${MAX_ARRAY_ITEMS} items.`
        );
      }
      value.forEach((item, index) => this.assertSerializableValue(item, `${path}[${index}]`, depth + 1));
      return;
    }

    for (const [key, nested] of Object.entries(value)) {
      this.assertSerializableValue(nested, `${path}.${key}`, depth + 1);
    }
  }
}

const sanitizeText = (value: string): string => {
  return value
    .replace(/\u0000/g, "")
    .replace(/</g, "\u2039")
    .replace(/>/g, "\u203A")
    .trim();
};

const sanitizeColor = (value: string): string => {
  const trimmed = value.trim();
  if (/^#[0-9a-f]{3,8}$/i.test(trimmed)) {
    return trimmed;
  }
  if (/^(rgb|rgba|hsl|hsla)\(/i.test(trimmed)) {
    return trimmed;
  }
  if (/^[a-z]+$/i.test(trimmed)) {
    return trimmed;
  }
  return "#334155";
};

const sanitizeImageSource = (value: string): string => {
  const trimmed = value.trim();
  if (trimmed.startsWith("data:image/")) {
    return trimmed;
  }
  if (trimmed.startsWith("./") || trimmed.startsWith("../") || trimmed.startsWith("/")) {
    return trimmed;
  }
  return "";
};

const clampInt = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? Math.round(value) : min;
  return Math.min(max, Math.max(min, numeric));
};

const clampNumber = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? value : min;
  return Math.min(max, Math.max(min, numeric));
};
