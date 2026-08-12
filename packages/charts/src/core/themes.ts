import type { ChartLayout } from "../model/Layout";

export interface ChartThemeTokens {
  name: string;
  backgroundColor: string;
  axisColor: string;
  gridColor: string;
  textColor: string;
}

export const DEFAULT_THEMES: Record<string, ChartThemeTokens> = {
  light: {
    name: "light",
    backgroundColor: "#ffffff",
    axisColor: "#94a3b8",
    gridColor: "#e2e8f0",
    textColor: "#0f172a"
  },
  dark: {
    name: "dark",
    backgroundColor: "#0f172a",
    axisColor: "#475569",
    gridColor: "#1e293b",
    textColor: "#f8fafc"
  }
};

export const DEFAULT_TEMPLATES: Record<string, Partial<ChartLayout>> = {
  standard: {},
  dashboard: {
    margin: {
      top: 44,
      right: 16,
      bottom: 44,
      left: 44
    },
    legend: {
      visible: true,
      position: "right"
    }
  },
  report: {
    margin: {
      top: 64,
      right: 30,
      bottom: 58,
      left: 58
    },
    legend: {
      visible: true,
      position: "top"
    }
  },
  minimal: {
    legend: {
      visible: false,
      position: "top"
    }
  },
  dark: {
    backgroundColor: "#0f172a",
    legend: {
      visible: true,
      position: "right"
    }
  }
};

export const applyThemeAndTemplate = (layout: ChartLayout): ChartLayout => {
  const template = layout.template ? DEFAULT_TEMPLATES[layout.template] : undefined;
  const theme = layout.theme ? DEFAULT_THEMES[layout.theme] : undefined;

  const merged: ChartLayout = {
    ...layout,
    ...(template ?? {}),
    margin: {
      ...layout.margin,
      ...(template?.margin ?? {})
    },
    legend: {
      ...layout.legend,
      ...(template?.legend ?? {})
    },
    xAxis: {
      ...layout.xAxis,
      ...(template?.xAxis ?? {})
    },
    xAxis2: {
      ...layout.xAxis2,
      ...(template?.xAxis2 ?? {})
    },
    yAxis: {
      ...layout.yAxis,
      ...(template?.yAxis ?? {})
    },
    yAxis2: {
      ...layout.yAxis2,
      ...(template?.yAxis2 ?? {})
    },
    shapes: template?.shapes ?? layout.shapes,
    annotations: template?.annotations ?? layout.annotations,
    images: template?.images ?? layout.images
  };

  if (!theme) {
    return merged;
  }

  return {
    ...merged,
    backgroundColor: theme.backgroundColor
  };
};
