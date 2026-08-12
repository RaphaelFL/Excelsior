import type { ChartFigure } from "../model/Figure";
import type { ChartMargin } from "../model/Layout";

const MIN_CHART_WIDTH = 160;
const MIN_CHART_HEIGHT = 120;

export interface ComputedPlotArea {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ComputedLegendArea {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ComputedLayout {
  width: number;
  height: number;
  margin: ChartMargin;
  plotArea: ComputedPlotArea;
  subplotAreas: ComputedPlotArea[];
  legendArea: ComputedLegendArea | null;
  titleY: number;
}

export class LayoutEngine {
  compute(container: HTMLElement, figure: ChartFigure): ComputedLayout {
    const width = Math.max(
      MIN_CHART_WIDTH,
      figure.layout.width ?? container.clientWidth ?? container.getBoundingClientRect().width ?? 640
    );
    const height = Math.max(
      MIN_CHART_HEIGHT,
      figure.layout.height ?? container.clientHeight ?? container.getBoundingClientRect().height ?? 360
    );

    const margin = figure.layout.margin;
    const titleOffset = figure.layout.title ? 22 : 0;
    let x = margin.left;
    let y = margin.top + titleOffset;
    let plotWidth = Math.max(1, width - margin.left - margin.right);
    let plotHeight = Math.max(1, height - margin.top - margin.bottom - titleOffset);
    let legendArea: ComputedLegendArea | null = null;

    if (figure.layout.legend.visible) {
      const position = figure.layout.legend.position;
      if (position === "top" || position === "bottom") {
        const legendHeight = Math.max(22, Math.min(66, Math.round(height * 0.12)));
        if (position === "top") {
          legendArea = {
            x: margin.left,
            y: margin.top,
            width: plotWidth,
            height: legendHeight
          };
          y += legendHeight;
        } else {
          legendArea = {
            x: margin.left,
            y: height - margin.bottom - legendHeight,
            width: plotWidth,
            height: legendHeight
          };
        }
        plotHeight = Math.max(1, plotHeight - legendHeight);
      } else {
        const legendWidth = Math.max(96, Math.min(180, Math.round(width * 0.2)));
        if (position === "left") {
          legendArea = {
            x: margin.left,
            y: y,
            width: legendWidth,
            height: plotHeight
          };
          x += legendWidth;
        } else {
          legendArea = {
            x: width - margin.right - legendWidth,
            y: y,
            width: legendWidth,
            height: plotHeight
          };
        }
        plotWidth = Math.max(1, plotWidth - legendWidth);
      }
    }

    const plotArea: ComputedPlotArea = {
      x,
      y,
      width: plotWidth,
      height: plotHeight
    };

    const subplotAreas = this.computeSubplotAreas(plotArea, figure.layout.subplots?.rows, figure.layout.subplots?.cols, figure.layout.subplots?.gapX, figure.layout.subplots?.gapY);

    return {
      width,
      height,
      margin,
      plotArea,
      subplotAreas,
      legendArea,
      titleY: Math.max(24, margin.top - 18)
    };
  }

  private computeSubplotAreas(
    plotArea: ComputedPlotArea,
    requestedRows: number | undefined,
    requestedCols: number | undefined,
    requestedGapX: number | undefined,
    requestedGapY: number | undefined
  ): ComputedPlotArea[] {
    const rows = clampInt(requestedRows ?? 1, 1, 16);
    const cols = clampInt(requestedCols ?? 1, 1, 16);
    const gapX = clampNumber(requestedGapX ?? 14, 0, 160);
    const gapY = clampNumber(requestedGapY ?? 14, 0, 160);
    if (rows === 1 && cols === 1) {
      return [plotArea];
    }

    const totalGapX = gapX * (cols - 1);
    const totalGapY = gapY * (rows - 1);
    const cellWidth = Math.max(1, (plotArea.width - totalGapX) / cols);
    const cellHeight = Math.max(1, (plotArea.height - totalGapY) / rows);
    const areas: ComputedPlotArea[] = [];

    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        areas.push({
          x: plotArea.x + col * (cellWidth + gapX),
          y: plotArea.y + row * (cellHeight + gapY),
          width: cellWidth,
          height: cellHeight
        });
      }
    }
    return areas;
  }
}

const clampInt = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? Math.round(value) : min;
  return Math.min(max, Math.max(min, numeric));
};

const clampNumber = (value: number, min: number, max: number): number => {
  const numeric = Number.isFinite(value) ? value : min;
  return Math.min(max, Math.max(min, numeric));
};
