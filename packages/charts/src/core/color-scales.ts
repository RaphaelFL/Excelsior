export type ColorScaleMode = "continuous" | "discrete" | "diverging";

export interface ColorScaleOptions {
  mode?: ColorScaleMode;
  colors?: string[];
  reverse?: boolean;
  missingColor?: string;
  min?: number;
  max?: number;
}

const DEFAULT_CONTINUOUS = ["#0ea5e9", "#f8fafc", "#ef4444"];
const DEFAULT_DIVERGING = ["#2563eb", "#f8fafc", "#dc2626"];
const DEFAULT_DISCRETE = ["#2563eb", "#f97316", "#059669", "#dc2626", "#7c3aed", "#0891b2", "#ca8a04"];

export const resolveColorFromScale = (value: number | null | undefined, options: ColorScaleOptions): string => {
  if (value === null || value === undefined || !Number.isFinite(value)) {
    return sanitizeColor(options.missingColor ?? "#cbd5e1");
  }

  const mode = options.mode ?? "continuous";
  const min = Number.isFinite(options.min) ? Number(options.min) : 0;
  const max = Number.isFinite(options.max) ? Number(options.max) : 1;
  const normalized = normalizeRatio(value, min, max);
  const palette = resolvePalette(mode, options.colors, options.reverse === true);

  if (mode === "discrete") {
    if (palette.length === 0) {
      return "#64748b";
    }
    const index = clampInt(Math.floor(normalized * palette.length), 0, palette.length - 1);
    return palette[index];
  }

  return interpolatePalette(palette, normalized);
};

export const normalizeRatio = (value: number, min: number, max: number): number => {
  if (!Number.isFinite(value)) {
    return 0;
  }
  if (!Number.isFinite(min) || !Number.isFinite(max)) {
    return 0;
  }
  if (min === max) {
    return 0.5;
  }
  const low = Math.min(min, max);
  const high = Math.max(min, max);
  return clamp((value - low) / (high - low), 0, 1);
};

export const interpolatePalette = (colors: string[], ratio: number): string => {
  const palette = colors.length >= 2 ? colors.map((color) => sanitizeColor(color)) : DEFAULT_CONTINUOUS;
  const bounded = clamp(ratio, 0, 1);
  const scaled = bounded * (palette.length - 1);
  const startIndex = Math.floor(scaled);
  const endIndex = Math.min(palette.length - 1, startIndex + 1);
  const local = scaled - startIndex;
  const start = toRgb(palette[startIndex]);
  const end = toRgb(palette[endIndex]);
  return `rgb(${Math.round(start.r + (end.r - start.r) * local)}, ${Math.round(start.g + (end.g - start.g) * local)}, ${Math.round(
    start.b + (end.b - start.b) * local
  )})`;
};

export const resolvePalette = (mode: ColorScaleMode, colors: string[] | undefined, reverse: boolean): string[] => {
  const fallback = mode === "discrete" ? DEFAULT_DISCRETE : mode === "diverging" ? DEFAULT_DIVERGING : DEFAULT_CONTINUOUS;
  const selected = (colors?.filter((color) => typeof color === "string").map((color) => sanitizeColor(color)) ?? fallback).slice();
  if (selected.length < 2 && mode !== "discrete") {
    return fallback.slice();
  }
  if (selected.length === 0) {
    return fallback.slice();
  }
  return reverse ? selected.reverse() : selected;
};

export const sanitizeColor = (color: string): string => {
  const value = color.trim();
  if (/^#[0-9a-f]{3}$/i.test(value) || /^#[0-9a-f]{6}$/i.test(value) || /^#[0-9a-f]{8}$/i.test(value)) {
    return value;
  }
  if (/^(rgb|rgba|hsl|hsla)\(/i.test(value)) {
    return value;
  }
  if (/^[a-z]+$/i.test(value)) {
    return value;
  }
  return "#64748b";
};

const toRgb = (value: string): { r: number; g: number; b: number } => {
  const normalized = value.replace("#", "");
  if (normalized.length === 3) {
    const [r, g, b] = normalized.split("").map((char) => parseInt(char + char, 16));
    return { r, g, b };
  }
  if (/^[0-9a-f]{6}$/i.test(normalized)) {
    return {
      r: parseInt(normalized.slice(0, 2), 16),
      g: parseInt(normalized.slice(2, 4), 16),
      b: parseInt(normalized.slice(4, 6), 16)
    };
  }
  return { r: 100, g: 116, b: 139 };
};

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

const clampInt = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, Math.round(value)));
