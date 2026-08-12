import type { ChartEventMap } from "../model/Events";
import type { ChartFigureInput, ChartSelectionState } from "../model/Figure";
import type { ChartLayout } from "../model/Layout";
import type { ChartTrace } from "../model/Trace";

export interface ChartHandle {
  update(nextFigure: ChartFigureInput): void;
  updateData(nextData: ChartTrace[]): void;
  updateLayout(nextLayout: Partial<ChartLayout>): void;
  resize(): void;
  destroy(): void;
  exportSvg(): string;
  exportPng(options?: { scale?: number; backgroundColor?: string }): Promise<Blob>;
  toJson(): string;
  exportDataTable(): string;
  getSelection(): ChartSelectionState | null;
  clearSelection(): void;
  playFrames(options?: FramePlaybackOptions): void;
  stopFrames(): void;
  isAnimating(): boolean;
  on<TKey extends keyof ChartEventMap>(event: TKey, handler: (payload: ChartEventMap[TKey]) => void): void;
  off<TKey extends keyof ChartEventMap>(event: TKey, handler: (payload: ChartEventMap[TKey]) => void): void;
}

export interface CreateFigureOptions {
  containerClassName?: string;
}

export interface FramePlaybackOptions {
  loop?: boolean;
  intervalMs?: number;
}
