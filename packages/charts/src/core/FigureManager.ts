import type { ChartFigure, ChartFigureInput, ChartSelectionState } from "../model/Figure";
import type { ChartLayout } from "../model/Layout";
import type { ChartTrace } from "../model/Trace";
import { FigureValidator } from "./FigureValidator";

export class FigureManager {
  private figure: ChartFigure;

  constructor(initialFigure: ChartFigureInput, private readonly validator: FigureValidator) {
    this.figure = this.validator.normalize(initialFigure);
  }

  getFigure(): ChartFigure {
    return this.figure;
  }

  update(nextFigure: ChartFigureInput): ChartFigure {
    this.figure = this.validator.normalize(nextFigure);
    return this.figure;
  }

  updateData(data: ChartTrace[]): ChartFigure {
    this.figure = this.validator.normalize({
      ...this.figure,
      data
    });
    return this.figure;
  }

  updateLayout(layout: Partial<ChartLayout>): ChartFigure {
    const current = this.figure.layout;
    this.figure = this.validator.normalize({
      ...this.figure,
      layout: {
        ...current,
        ...layout,
        margin: {
          ...current.margin,
          ...(layout.margin ?? {})
        },
        xAxis: {
          ...current.xAxis,
          ...(layout.xAxis ?? {})
        },
        xAxis2: {
          ...current.xAxis2,
          ...(layout.xAxis2 ?? {})
        },
        yAxis: {
          ...current.yAxis,
          ...(layout.yAxis ?? {})
        },
        yAxis2: {
          ...current.yAxis2,
          ...(layout.yAxis2 ?? {})
        },
        legend: {
          ...current.legend,
          ...(layout.legend ?? {})
        },
        shapes: layout.shapes ?? current.shapes,
        annotations: layout.annotations ?? current.annotations,
        images: layout.images ?? current.images
      }
    });
    return this.figure;
  }

  updateSelection(selection: ChartSelectionState | null): ChartFigure {
    this.figure = this.validator.normalize({
      ...this.figure,
      selection
    });
    return this.figure;
  }
}
