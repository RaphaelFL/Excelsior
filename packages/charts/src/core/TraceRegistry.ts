import type { ChartRendererKind } from "../model/Config";
import type { ChartTrace } from "../model/Trace";
import { ChartConfigurationError } from "./chart-errors";

export interface TraceDefinition {
  type: string;
  renderer: ChartRendererKind;
  validate: (trace: ChartTrace) => void;
}

export class TraceRegistry {
  private readonly definitions = new Map<string, TraceDefinition>();

  constructor(seed: readonly TraceDefinition[] = DEFAULT_TRACE_DEFINITIONS) {
    for (const definition of seed) {
      this.register(definition, { allowOverride: true });
    }
  }

  register(definition: TraceDefinition, options?: { allowOverride?: boolean }): void {
    const normalizedType = definition.type.trim();
    if (!normalizedType) {
      throw new ChartConfigurationError("CHART_TRACE_TYPE_EMPTY", "Trace type must be a non-empty string.");
    }

    if (!options?.allowOverride && this.definitions.has(normalizedType)) {
      throw new ChartConfigurationError(
        "CHART_TRACE_ALREADY_REGISTERED",
        `Trace type '${normalizedType}' is already registered.`
      );
    }

    this.definitions.set(normalizedType, { ...definition, type: normalizedType });
  }

  has(type: string): boolean {
    return this.definitions.has(type);
  }

  resolve(type: string): TraceDefinition {
    const definition = this.definitions.get(type);
    if (!definition) {
      throw new ChartConfigurationError("CHART_TRACE_UNKNOWN", `Unknown trace type '${type}'.`);
    }
    return definition;
  }

  listTypes(): string[] {
    return Array.from(this.definitions.keys());
  }
}

const assertCartesianTrace = (trace: ChartTrace): void => {
  const x = (trace as { x?: unknown }).x;
  const y = (trace as { y?: unknown }).y;
  if (!Array.isArray(x) || !Array.isArray(y)) {
    throw new ChartConfigurationError("CHART_TRACE_CARTESIAN_INVALID", "Cartesian traces require array fields x and y.");
  }
  if (x.length !== y.length) {
    throw new ChartConfigurationError("CHART_TRACE_CARTESIAN_LENGTH", "Trace fields x and y must have the same length.");
  }
  if (x.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_EMPTY", "Trace x/y arrays cannot be empty.");
  }
  const errorY = (trace as { errorY?: unknown }).errorY as { values?: unknown } | undefined;
  if (errorY?.values !== undefined) {
    if (!Array.isArray(errorY.values) || errorY.values.length !== x.length) {
      throw new ChartConfigurationError("CHART_TRACE_CARTESIAN_INVALID", "Cartesian errorY.values must match x/y length.");
    }
  }
  const orientation = (trace as { orientation?: unknown }).orientation;
  if (orientation !== undefined && orientation !== "vertical" && orientation !== "horizontal") {
    throw new ChartConfigurationError("CHART_TRACE_CARTESIAN_INVALID", "Cartesian orientation must be 'vertical' or 'horizontal'.");
  }
};

const assertPieTrace = (trace: ChartTrace): void => {
  const values = (trace as { values?: unknown }).values;
  if (!Array.isArray(values)) {
    throw new ChartConfigurationError("CHART_TRACE_PIE_INVALID", "Pie trace requires an array field values.");
  }
  if (values.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_EMPTY", "Pie trace values cannot be empty.");
  }
  for (let index = 0; index < values.length; index += 1) {
    const numeric = Number(values[index]);
    if (!Number.isFinite(numeric)) {
      throw new ChartConfigurationError("CHART_TRACE_PIE_INVALID", `Pie trace contains non-numeric value at index ${index}.`);
    }
    if (numeric < 0) {
      throw new ChartConfigurationError("CHART_TRACE_PIE_INVALID", `Pie trace contains negative value at index ${index}.`);
    }
  }
  const labels = (trace as { labels?: unknown }).labels;
  if (labels !== undefined && (!Array.isArray(labels) || labels.length !== values.length)) {
    throw new ChartConfigurationError("CHART_TRACE_PIE_INVALID", "Pie trace labels must match values length when provided.");
  }
  const pull = (trace as { pull?: unknown }).pull;
  if (pull !== undefined) {
    if (Array.isArray(pull)) {
      if (pull.length !== values.length) {
        throw new ChartConfigurationError("CHART_TRACE_PIE_INVALID", "Pie trace pull array must match values length when provided.");
      }
      pull.forEach((value, index) => {
        const numeric = Number(value);
        if (!Number.isFinite(numeric) || numeric < 0) {
          throw new ChartConfigurationError("CHART_TRACE_PIE_INVALID", `Pie trace pull contains invalid value at index ${index}.`);
        }
      });
    } else {
      const numeric = Number(pull);
      if (!Number.isFinite(numeric) || numeric < 0) {
        throw new ChartConfigurationError("CHART_TRACE_PIE_INVALID", "Pie trace pull must be a non-negative number or an array.");
      }
    }
  }
};

const assertSunburstLikeTrace = (trace: ChartTrace, kind: "sunburst" | "treemap"): void => {
  const ids = (trace as { ids?: unknown }).ids;
  const labels = (trace as { labels?: unknown }).labels;
  const parents = (trace as { parents?: unknown }).parents;
  const values = (trace as { values?: unknown }).values;
  if (!Array.isArray(labels) || !Array.isArray(parents) || !Array.isArray(values)) {
    throw new ChartConfigurationError(`CHART_TRACE_${kind.toUpperCase()}_INVALID`, `${kind} trace requires labels, parents and values arrays.`);
  }
  if (labels.length === 0 || labels.length !== parents.length || labels.length !== values.length) {
    throw new ChartConfigurationError(
      `CHART_TRACE_${kind.toUpperCase()}_LENGTH`,
      `${kind} trace labels/parents/values arrays must have same non-empty length.`
    );
  }
  if (ids !== undefined && (!Array.isArray(ids) || ids.length !== labels.length)) {
    throw new ChartConfigurationError(`CHART_TRACE_${kind.toUpperCase()}_INVALID`, `${kind} ids must match labels length when provided.`);
  }

  const nodeIds = labels.map((label, index) => String((Array.isArray(ids) ? ids[index] : label) ?? index));
  const nodeSet = new Set(nodeIds);
  const parentById = new Map<string, string>();
  nodeIds.forEach((id, index) => {
    const parentRaw = parents[index];
    const parentId = String(parentRaw ?? "");
    if (parentId && !nodeSet.has(parentId)) {
      throw new ChartConfigurationError(
        `CHART_TRACE_${kind.toUpperCase()}_INVALID`,
        `${kind} parent '${parentId}' at index ${index} does not exist in ids/labels.`
      );
    }
    if (parentId && parentId === id) {
      throw new ChartConfigurationError(`CHART_TRACE_${kind.toUpperCase()}_INVALID`, `${kind} node '${id}' cannot reference itself as parent.`);
    }
    parentById.set(id, parentId);
  });
  for (const id of nodeIds) {
    const seen = new Set<string>();
    let current = id;
    while (true) {
      const parent = parentById.get(current);
      if (!parent) {
        break;
      }
      if (seen.has(parent)) {
        throw new ChartConfigurationError(`CHART_TRACE_${kind.toUpperCase()}_INVALID`, `${kind} hierarchy contains cycle at node '${id}'.`);
      }
      seen.add(parent);
      current = parent;
    }
  }
};

const assertViolinTrace = (trace: ChartTrace): void => {
  const values = (trace as { values?: unknown }).values;
  if (!Array.isArray(values) || values.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_VIOLIN_INVALID", "Violin trace requires non-empty numeric values.");
  }
};

const assertDensityTrace = (trace: ChartTrace): void => {
  const values = (trace as { values?: unknown }).values;
  if (!Array.isArray(values) || values.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_DENSITY_INVALID", "Density/distribution trace requires non-empty numeric values.");
  }
  const sampleLimit = Number((trace as { sampleLimit?: unknown }).sampleLimit);
  if ((trace as { sampleLimit?: unknown }).sampleLimit !== undefined && (!Number.isFinite(sampleLimit) || sampleLimit <= 0)) {
    throw new ChartConfigurationError("CHART_TRACE_DENSITY_INVALID", "Density/distribution sampleLimit must be positive.");
  }
};

const assertSankeyTrace = (trace: ChartTrace): void => {
  const nodes = (trace as { nodes?: unknown }).nodes as { ids?: unknown; labels?: unknown; colors?: unknown } | undefined;
  const links = (trace as { links?: unknown }).links as { source?: unknown; target?: unknown; value?: unknown } | undefined;
  if (!nodes || !links || !Array.isArray(nodes.ids) || !Array.isArray(links.source) || !Array.isArray(links.target) || !Array.isArray(links.value)) {
    throw new ChartConfigurationError("CHART_TRACE_SANKEY_INVALID", "Sankey trace requires nodes.ids and links source/target/value arrays.");
  }
  const nodeCount = nodes.ids.length;
  if (nodeCount === 0) {
    throw new ChartConfigurationError("CHART_TRACE_SANKEY_INVALID", "Sankey nodes.ids cannot be empty.");
  }
  if (
    links.source.length === 0 ||
    links.source.length !== links.target.length ||
    links.source.length !== links.value.length
  ) {
    throw new ChartConfigurationError("CHART_TRACE_SANKEY_INVALID", "Sankey links source/target/value arrays must share same non-empty length.");
  }
  for (let index = 0; index < links.source.length; index += 1) {
    const source = Number(links.source[index]);
    const target = Number(links.target[index]);
    const value = Number(links.value[index]);
    if (!Number.isFinite(source) || !Number.isFinite(target) || !Number.isFinite(value)) {
      throw new ChartConfigurationError("CHART_TRACE_SANKEY_INVALID", `Sankey link at index ${index} contains invalid values.`);
    }
    if (source < 0 || source >= nodeCount || target < 0 || target >= nodeCount) {
      throw new ChartConfigurationError("CHART_TRACE_SANKEY_INVALID", `Sankey link at index ${index} references out-of-bounds node.`);
    }
  }
};

const assertParallelCategoriesTrace = (trace: ChartTrace): void => {
  const dimensions = (trace as { dimensions?: unknown }).dimensions;
  if (!Array.isArray(dimensions) || dimensions.length < 2) {
    throw new ChartConfigurationError("CHART_TRACE_PARCAT_INVALID", "Parallel categories trace requires at least two dimensions.");
  }
  const firstLength = Array.isArray((dimensions[0] as { values?: unknown }).values) ? (dimensions[0] as { values: unknown[] }).values.length : -1;
  if (firstLength <= 0) {
    throw new ChartConfigurationError("CHART_TRACE_PARCAT_INVALID", "Parallel categories dimensions must include non-empty values.");
  }
  dimensions.forEach((dimension, index) => {
    const values = (dimension as { values?: unknown }).values;
    const name = (dimension as { name?: unknown }).name;
    if (typeof name !== "string" || !name.trim() || !Array.isArray(values) || values.length !== firstLength) {
      throw new ChartConfigurationError("CHART_TRACE_PARCAT_INVALID", `Parallel categories dimension at index ${index} is invalid.`);
    }
  });
};

const assertHistogramTrace = (trace: ChartTrace): void => {
  const values = (trace as { values?: unknown }).values;
  if (!Array.isArray(values)) {
    throw new ChartConfigurationError("CHART_TRACE_HISTOGRAM_INVALID", "Histogram trace requires an array field values.");
  }
  if (values.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_EMPTY", "Histogram values cannot be empty.");
  }
  const orientation = (trace as { orientation?: unknown }).orientation;
  if (orientation !== undefined && orientation !== "vertical" && orientation !== "horizontal") {
    throw new ChartConfigurationError("CHART_TRACE_HISTOGRAM_INVALID", "Histogram orientation must be 'vertical' or 'horizontal'.");
  }
};

const assertBoxTrace = (trace: ChartTrace): void => {
  const values = (trace as { values?: unknown }).values;
  if (!Array.isArray(values)) {
    throw new ChartConfigurationError("CHART_TRACE_BOX_INVALID", "Box trace requires an array field values.");
  }
  if (values.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_EMPTY", "Box values cannot be empty.");
  }
  const factor = Number((trace as { outlierIqrFactor?: unknown }).outlierIqrFactor);
  if ((trace as { outlierIqrFactor?: unknown }).outlierIqrFactor !== undefined && (!Number.isFinite(factor) || factor <= 0)) {
    throw new ChartConfigurationError("CHART_TRACE_BOX_INVALID", "Box outlierIqrFactor must be positive when provided.");
  }
};

const assertHeatmapTrace = (trace: ChartTrace): void => {
  const z = (trace as { z?: unknown }).z;
  if (!Array.isArray(z) || z.length === 0 || !Array.isArray(z[0])) {
    throw new ChartConfigurationError("CHART_TRACE_HEATMAP_INVALID", "Heatmap trace requires a 2D numeric matrix field z.");
  }

  const width = (z[0] as unknown[]).length;
  if (width === 0) {
    throw new ChartConfigurationError("CHART_TRACE_HEATMAP_INVALID", "Heatmap trace z matrix cannot contain empty rows.");
  }

  for (const row of z as unknown[]) {
    if (!Array.isArray(row) || row.length !== width) {
      throw new ChartConfigurationError(
        "CHART_TRACE_HEATMAP_INVALID",
        "Heatmap trace z matrix must have rows with a consistent length."
      );
    }
  }
};

const assertContourTrace = (trace: ChartTrace): void => {
  const z = (trace as { z?: unknown }).z;
  if (!Array.isArray(z) || z.length === 0 || !Array.isArray(z[0])) {
    throw new ChartConfigurationError("CHART_TRACE_CONTOUR_INVALID", "Contour trace requires a 2D numeric matrix field z.");
  }
  const width = (z[0] as unknown[]).length;
  if (width === 0) {
    throw new ChartConfigurationError("CHART_TRACE_CONTOUR_INVALID", "Contour trace z matrix cannot contain empty rows.");
  }
  for (const row of z as unknown[]) {
    if (!Array.isArray(row) || row.length !== width) {
      throw new ChartConfigurationError("CHART_TRACE_CONTOUR_INVALID", "Contour trace z matrix rows must have consistent length.");
    }
  }
  const maxSegments = Number((trace as { maxSegments?: unknown }).maxSegments);
  if ((trace as { maxSegments?: unknown }).maxSegments !== undefined && (!Number.isFinite(maxSegments) || maxSegments < 10)) {
    throw new ChartConfigurationError("CHART_TRACE_CONTOUR_INVALID", "Contour maxSegments must be >= 10 when provided.");
  }
};

const assertQuiverTrace = (trace: ChartTrace): void => {
  const x = (trace as { x?: unknown }).x;
  const y = (trace as { y?: unknown }).y;
  const u = (trace as { u?: unknown }).u;
  const v = (trace as { v?: unknown }).v;
  if (!Array.isArray(x) || !Array.isArray(y) || !Array.isArray(u) || !Array.isArray(v)) {
    throw new ChartConfigurationError("CHART_TRACE_QUIVER_INVALID", "Quiver trace requires x, y, u and v arrays.");
  }
  if (x.length === 0 || x.length !== y.length || x.length !== u.length || x.length !== v.length) {
    throw new ChartConfigurationError("CHART_TRACE_QUIVER_INVALID", "Quiver trace x/y/u/v arrays must have same non-empty length.");
  }
};

const assertFinancialTrace = (trace: ChartTrace): void => {
  const x = (trace as { x?: unknown }).x;
  const open = (trace as { open?: unknown }).open;
  const high = (trace as { high?: unknown }).high;
  const low = (trace as { low?: unknown }).low;
  const close = (trace as { close?: unknown }).close;

  if (!Array.isArray(x) || !Array.isArray(open) || !Array.isArray(high) || !Array.isArray(low) || !Array.isArray(close)) {
    throw new ChartConfigurationError(
      "CHART_TRACE_FINANCIAL_INVALID",
      "Financial traces require x, open, high, low and close arrays."
    );
  }

  const expectedLength = x.length;
  if (expectedLength === 0) {
    throw new ChartConfigurationError("CHART_TRACE_EMPTY", "Financial trace arrays cannot be empty.");
  }
  if (open.length !== expectedLength || high.length !== expectedLength || low.length !== expectedLength || close.length !== expectedLength) {
    throw new ChartConfigurationError(
      "CHART_TRACE_FINANCIAL_LENGTH",
      "Financial trace arrays x/open/high/low/close must have the same length."
    );
  }

  for (let index = 0; index < expectedLength; index += 1) {
    const openValue = Number(open[index]);
    const highValue = Number(high[index]);
    const lowValue = Number(low[index]);
    const closeValue = Number(close[index]);
    if (!Number.isFinite(openValue) || !Number.isFinite(highValue) || !Number.isFinite(lowValue) || !Number.isFinite(closeValue)) {
      throw new ChartConfigurationError(
        "CHART_TRACE_FINANCIAL_INVALID",
        `Financial trace contains a non-numeric OHLC value at index ${index}.`
      );
    }
    if (lowValue > highValue) {
      throw new ChartConfigurationError(
        "CHART_TRACE_FINANCIAL_RANGE",
        `Financial trace low value cannot exceed high value at index ${index}.`
      );
    }
  }
};

const assertWaterfallTrace = (trace: ChartTrace): void => {
  const x = (trace as { x?: unknown }).x;
  const y = (trace as { y?: unknown }).y;
  if (!Array.isArray(x) || !Array.isArray(y)) {
    throw new ChartConfigurationError("CHART_TRACE_WATERFALL_INVALID", "Waterfall trace requires x and y arrays.");
  }
  if (x.length === 0 || x.length !== y.length) {
    throw new ChartConfigurationError("CHART_TRACE_WATERFALL_LENGTH", "Waterfall trace x and y arrays must have same non-empty length.");
  }
  const measure = (trace as { measure?: unknown }).measure;
  if (measure !== undefined) {
    if (!Array.isArray(measure) || measure.length !== x.length) {
      throw new ChartConfigurationError("CHART_TRACE_WATERFALL_INVALID", "Waterfall measure array must match x/y length.");
    }
    for (let index = 0; index < measure.length; index += 1) {
      const value = measure[index];
      if (value !== "relative" && value !== "total" && value !== "absolute") {
        throw new ChartConfigurationError("CHART_TRACE_WATERFALL_INVALID", `Waterfall measure at index ${index} is invalid.`);
      }
    }
  }
};

const assertFunnelTrace = (trace: ChartTrace): void => {
  const labels = (trace as { labels?: unknown }).labels;
  const values = (trace as { values?: unknown }).values;
  if (!Array.isArray(labels) || !Array.isArray(values)) {
    throw new ChartConfigurationError("CHART_TRACE_FUNNEL_INVALID", "Funnel trace requires labels and values arrays.");
  }
  if (labels.length === 0 || labels.length !== values.length) {
    throw new ChartConfigurationError("CHART_TRACE_FUNNEL_LENGTH", "Funnel trace labels/values arrays must have same non-empty length.");
  }
  const sort = (trace as { sort?: unknown }).sort;
  if (sort !== undefined && sort !== "none" && sort !== "asc" && sort !== "desc") {
    throw new ChartConfigurationError("CHART_TRACE_FUNNEL_INVALID", "Funnel sort must be 'none', 'asc', or 'desc'.");
  }
};

const assertPolarTrace = (trace: ChartTrace): void => {
  const theta = (trace as { theta?: unknown }).theta;
  const r = (trace as { r?: unknown }).r;
  if (!Array.isArray(theta) || !Array.isArray(r)) {
    throw new ChartConfigurationError("CHART_TRACE_POLAR_INVALID", "Polar trace requires theta and r arrays.");
  }
  if (theta.length !== r.length || theta.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_POLAR_LENGTH", "Polar trace theta and r arrays must have same non-empty length.");
  }
  const variant = (trace as { variant?: unknown }).variant;
  if (variant !== undefined && variant !== "scatter" && variant !== "line" && variant !== "bar" && variant !== "area") {
    throw new ChartConfigurationError("CHART_TRACE_POLAR_INVALID", "Polar variant must be 'scatter', 'line', 'bar', or 'area'.");
  }
};

const assertTernaryTrace = (trace: ChartTrace): void => {
  const a = (trace as { a?: unknown }).a;
  const b = (trace as { b?: unknown }).b;
  const c = (trace as { c?: unknown }).c;
  if (!Array.isArray(a) || !Array.isArray(b) || !Array.isArray(c)) {
    throw new ChartConfigurationError("CHART_TRACE_TERNARY_INVALID", "Ternary trace requires a, b and c arrays.");
  }
  if (a.length !== b.length || b.length !== c.length || a.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_TERNARY_LENGTH", "Ternary trace arrays a/b/c must have same non-empty length.");
  }
};

const assertGeoTrace = (trace: ChartTrace): void => {
  const geojson = (trace as { geojson?: unknown }).geojson;
  if (!geojson || typeof geojson !== "object") {
    throw new ChartConfigurationError("CHART_TRACE_GEO_INVALID", "Geo trace requires a GeoJSON object.");
  }

  const candidate = geojson as { type?: unknown; features?: unknown };
  if (candidate.type !== "FeatureCollection" || !Array.isArray(candidate.features)) {
    throw new ChartConfigurationError("CHART_TRACE_GEO_INVALID", "Geo trace geojson must be a FeatureCollection with features array.");
  }
  if (candidate.features.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_EMPTY", "Geo trace requires at least one feature.");
  }

  candidate.features.forEach((feature, index) => {
    const geometry = (feature as { geometry?: unknown }).geometry as
      | { type?: unknown; coordinates?: unknown }
      | undefined;
    if (!geometry || (geometry.type !== "Polygon" && geometry.type !== "MultiPolygon")) {
      throw new ChartConfigurationError(
        "CHART_TRACE_GEO_INVALID",
        `Geo trace feature at index ${index} must use Polygon or MultiPolygon geometry.`
      );
    }
    if (!Array.isArray(geometry.coordinates)) {
      throw new ChartConfigurationError(
        "CHART_TRACE_GEO_INVALID",
        `Geo trace feature at index ${index} has invalid coordinates.`
      );
    }
  });
  const locations = (trace as { locations?: unknown }).locations;
  const values = (trace as { values?: unknown }).values;
  if (locations !== undefined || values !== undefined) {
    if (!Array.isArray(locations) || !Array.isArray(values) || locations.length !== values.length) {
      throw new ChartConfigurationError("CHART_TRACE_GEO_INVALID", "Geo trace locations and values must be arrays with same length.");
    }
  }
};

const assertGeoScatterTrace = (trace: ChartTrace): void => {
  const lat = (trace as { lat?: unknown }).lat;
  const lon = (trace as { lon?: unknown }).lon;
  if (!Array.isArray(lat) || !Array.isArray(lon)) {
    throw new ChartConfigurationError("CHART_TRACE_GEO_SCATTER_INVALID", "Geo scatter trace requires lat and lon arrays.");
  }
  if (lat.length === 0 || lat.length !== lon.length) {
    throw new ChartConfigurationError("CHART_TRACE_GEO_SCATTER_LENGTH", "Geo scatter lat/lon arrays must have same non-empty length.");
  }
  for (let index = 0; index < lat.length; index += 1) {
    const latitude = Number(lat[index]);
    const longitude = Number(lon[index]);
    if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) {
      throw new ChartConfigurationError("CHART_TRACE_GEO_SCATTER_INVALID", `Geo scatter contains invalid coordinates at index ${index}.`);
    }
    if (latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) {
      throw new ChartConfigurationError("CHART_TRACE_GEO_SCATTER_INVALID", `Geo scatter coordinates out of bounds at index ${index}.`);
    }
  }
};

const assertGeoLineTrace = (trace: ChartTrace): void => {
  const paths = (trace as { paths?: unknown }).paths;
  if (!Array.isArray(paths) || paths.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_GEO_LINE_INVALID", "Geo line trace requires a non-empty paths array.");
  }
  paths.forEach((path, pathIndex) => {
    if (!Array.isArray(path) || path.length < 2) {
      throw new ChartConfigurationError("CHART_TRACE_GEO_LINE_INVALID", `Geo line path at index ${pathIndex} must contain at least two points.`);
    }
    path.forEach((point, pointIndex) => {
      const latitude = Number((point as { lat?: unknown }).lat);
      const longitude = Number((point as { lon?: unknown }).lon);
      if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) {
        throw new ChartConfigurationError(
          "CHART_TRACE_GEO_LINE_INVALID",
          `Geo line path ${pathIndex} contains invalid coordinates at point ${pointIndex}.`
        );
      }
    });
  });
};

const assertScatter3dTrace = (trace: ChartTrace): void => {
  const x = (trace as { x?: unknown }).x;
  const y = (trace as { y?: unknown }).y;
  const z = (trace as { z?: unknown }).z;
  if (!Array.isArray(x) || !Array.isArray(y) || !Array.isArray(z)) {
    throw new ChartConfigurationError("CHART_TRACE_SCATTER3D_INVALID", "Scatter3D trace requires x, y and z arrays.");
  }
  if (x.length === 0 || x.length !== y.length || x.length !== z.length) {
    throw new ChartConfigurationError("CHART_TRACE_SCATTER3D_LENGTH", "Scatter3D trace x/y/z arrays must have same non-empty length.");
  }
};

const assertSurfaceTrace = (trace: ChartTrace): void => {
  const z = (trace as { z?: unknown }).z;
  if (!Array.isArray(z) || z.length === 0 || !Array.isArray(z[0])) {
    throw new ChartConfigurationError("CHART_TRACE_SURFACE_INVALID", "Surface trace requires a non-empty 2D z matrix.");
  }
  const width = (z[0] as unknown[]).length;
  if (width === 0) {
    throw new ChartConfigurationError("CHART_TRACE_SURFACE_INVALID", "Surface z matrix rows cannot be empty.");
  }
  for (const row of z as unknown[]) {
    if (!Array.isArray(row) || row.length !== width) {
      throw new ChartConfigurationError("CHART_TRACE_SURFACE_INVALID", "Surface z matrix rows must have consistent width.");
    }
  }
};

const assertMesh3dTrace = (trace: ChartTrace): void => {
  const x = (trace as { x?: unknown }).x;
  const y = (trace as { y?: unknown }).y;
  const z = (trace as { z?: unknown }).z;
  const i = (trace as { i?: unknown }).i;
  const j = (trace as { j?: unknown }).j;
  const k = (trace as { k?: unknown }).k;
  if (!Array.isArray(x) || !Array.isArray(y) || !Array.isArray(z) || !Array.isArray(i) || !Array.isArray(j) || !Array.isArray(k)) {
    throw new ChartConfigurationError("CHART_TRACE_MESH3D_INVALID", "Mesh3D trace requires x, y, z, i, j and k arrays.");
  }
  if (x.length === 0 || x.length !== y.length || x.length !== z.length) {
    throw new ChartConfigurationError("CHART_TRACE_MESH3D_LENGTH", "Mesh3D x/y/z arrays must have same non-empty length.");
  }
  if (i.length !== j.length || i.length !== k.length || i.length === 0) {
    throw new ChartConfigurationError("CHART_TRACE_MESH3D_LENGTH", "Mesh3D index arrays i/j/k must have same non-empty length.");
  }
  const maxVertexIndex = x.length - 1;
  for (let index = 0; index < i.length; index += 1) {
    const ia = Number(i[index]);
    const jb = Number(j[index]);
    const kc = Number(k[index]);
    if (!Number.isFinite(ia) || !Number.isFinite(jb) || !Number.isFinite(kc)) {
      throw new ChartConfigurationError("CHART_TRACE_MESH3D_INVALID", `Mesh3D contains invalid triangle index at position ${index}.`);
    }
    if (ia < 0 || jb < 0 || kc < 0 || ia > maxVertexIndex || jb > maxVertexIndex || kc > maxVertexIndex) {
      throw new ChartConfigurationError("CHART_TRACE_MESH3D_INVALID", `Mesh3D triangle index out of bounds at position ${index}.`);
    }
  }
};

export const DEFAULT_TRACE_DEFINITIONS: readonly TraceDefinition[] = [
  {
    type: "scatter",
    renderer: "svg",
    validate: assertCartesianTrace
  },
  {
    type: "line",
    renderer: "svg",
    validate: assertCartesianTrace
  },
  {
    type: "bar",
    renderer: "svg",
    validate: assertCartesianTrace
  },
  {
    type: "area",
    renderer: "svg",
    validate: assertCartesianTrace
  },
  {
    type: "violin",
    renderer: "svg",
    validate: assertViolinTrace
  },
  {
    type: "density",
    renderer: "svg",
    validate: assertDensityTrace
  },
  {
    type: "distribution",
    renderer: "svg",
    validate: assertDensityTrace
  },
  {
    type: "pie",
    renderer: "svg",
    validate: assertPieTrace
  },
  {
    type: "donut",
    renderer: "svg",
    validate: assertPieTrace
  },
  {
    type: "sunburst",
    renderer: "svg",
    validate: (trace) => assertSunburstLikeTrace(trace, "sunburst")
  },
  {
    type: "treemap",
    renderer: "svg",
    validate: (trace) => assertSunburstLikeTrace(trace, "treemap")
  },
  {
    type: "sankey",
    renderer: "svg",
    validate: assertSankeyTrace
  },
  {
    type: "parallel-categories",
    renderer: "svg",
    validate: assertParallelCategoriesTrace
  },
  {
    type: "histogram",
    renderer: "svg",
    validate: assertHistogramTrace
  },
  {
    type: "box",
    renderer: "svg",
    validate: assertBoxTrace
  },
  {
    type: "heatmap",
    renderer: "svg",
    validate: assertHeatmapTrace
  },
  {
    type: "contour",
    renderer: "svg",
    validate: assertContourTrace
  },
  {
    type: "quiver",
    renderer: "svg",
    validate: assertQuiverTrace
  },
  {
    type: "candlestick",
    renderer: "svg",
    validate: assertFinancialTrace
  },
  {
    type: "ohlc",
    renderer: "svg",
    validate: assertFinancialTrace
  },
  {
    type: "waterfall",
    renderer: "svg",
    validate: assertWaterfallTrace
  },
  {
    type: "funnel",
    renderer: "svg",
    validate: assertFunnelTrace
  },
  {
    type: "polar",
    renderer: "svg",
    validate: assertPolarTrace
  },
  {
    type: "ternary",
    renderer: "svg",
    validate: assertTernaryTrace
  },
  {
    type: "geo",
    renderer: "svg",
    validate: assertGeoTrace
  },
  {
    type: "geo-scatter",
    renderer: "svg",
    validate: assertGeoScatterTrace
  },
  {
    type: "geo-line",
    renderer: "svg",
    validate: assertGeoLineTrace
  },
  {
    type: "scatter3d",
    renderer: "webgl",
    validate: assertScatter3dTrace
  },
  {
    type: "surface",
    renderer: "webgl",
    validate: assertSurfaceTrace
  },
  {
    type: "mesh3d",
    renderer: "webgl",
    validate: assertMesh3dTrace
  }
];
