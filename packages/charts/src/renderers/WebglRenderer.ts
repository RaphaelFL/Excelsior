import type { CartesianTrace, ChartTrace, Mesh3dTrace, Scatter3dTrace, SurfaceTrace } from "../model/Trace";
import type { ChartFigure } from "../model/Figure";
import type { ComputedLayout } from "../core/LayoutEngine";
import type { ChartRenderer } from "./ChartRenderer";
import { ChartConfigurationError } from "../core/chart-errors";
import { buildCartesianDomains, isCartesianTrace } from "../core/cartesian-domain";
import { SvgRenderer } from "./SvgRenderer";
import { normalizeAxisType, toAxisScalar } from "../core/axis-utils";

const VERTEX_SHADER_SOURCE = `
attribute vec2 a_position;
uniform float u_pointSize;
void main() {
  gl_Position = vec4(a_position, 0.0, 1.0);
  gl_PointSize = u_pointSize;
}
`;

const FRAGMENT_SHADER_SOURCE = `
precision mediump float;
uniform vec4 u_color;
void main() {
  gl_FragColor = u_color;
}
`;

const DEFAULT_SERIES_COLORS = ["#2563eb", "#f97316", "#059669", "#dc2626", "#7c3aed", "#0891b2", "#ca8a04"];
const isScatter3dTrace = (trace: ChartTrace): trace is Scatter3dTrace => trace.type === "scatter3d";
const isSurfaceTrace = (trace: ChartTrace): trace is SurfaceTrace => trace.type === "surface";
const isMesh3dTrace = (trace: ChartTrace): trace is Mesh3dTrace => trace.type === "mesh3d";

interface CameraState3d {
  rotationX: number;
  rotationY: number;
  zoom: number;
  panX: number;
  panY: number;
}

const DEFAULT_CAMERA_STATE: CameraState3d = {
  rotationX: (-30 * Math.PI) / 180,
  rotationY: (38 * Math.PI) / 180,
  zoom: 1,
  panX: 0,
  panY: 0
};

export class WebglRenderer implements ChartRenderer {
  private container: HTMLElement | null = null;
  private root: HTMLDivElement | null = null;
  private canvas: HTMLCanvasElement | null = null;
  private interactionLayer: HTMLDivElement | null = null;
  private gl: WebGLRenderingContext | null = null;
  private program: WebGLProgram | null = null;
  private positionBuffer: WebGLBuffer | null = null;
  private lastFigure: ChartFigure | null = null;
  private lastLayout: ComputedLayout | null = null;
  private camera: CameraState3d = { ...DEFAULT_CAMERA_STATE };
  private dragPointerId: number | null = null;
  private dragMode: "rotate" | "pan" = "rotate";
  private lastPointerX = 0;
  private lastPointerY = 0;

  mount(container: HTMLElement): void {
    this.destroy();
    this.container = container;
    if (container.style.position === "" || container.style.position === "static") {
      container.style.position = "relative";
    }

    const root = document.createElement("div");
    root.className = "excelsior-chart-webgl-root";
    Object.assign(root.style, {
      position: "absolute",
      inset: "0",
      overflow: "hidden",
      touchAction: "none"
    });

    const canvas = document.createElement("canvas");
    Object.assign(canvas.style, {
      position: "absolute",
      inset: "0",
      width: "100%",
      height: "100%"
    });

    const interactionLayer = document.createElement("div");
    Object.assign(interactionLayer.style, {
      position: "absolute",
      inset: "0",
      pointerEvents: "auto"
    });

    root.append(canvas, interactionLayer);
    container.append(root);
    this.root = root;
    this.canvas = canvas;
    this.interactionLayer = interactionLayer;

    const gl = canvas.getContext("webgl");
    if (!gl) {
      throw new ChartConfigurationError("CHART_WEBGL_UNAVAILABLE", "WebGL context is unavailable in this environment.");
    }
    this.gl = gl;
    this.initializeProgram();
    this.attachCameraListeners(interactionLayer);
  }

  render(figure: ChartFigure, layout: ComputedLayout): void {
    const gl = this.ensureGl();
    const canvas = this.ensureCanvas();
    const layer = this.ensureInteractionLayer();
    layer.replaceChildren();
    this.lastFigure = structuredClone(figure);
    this.lastLayout = layout;

    canvas.width = Math.max(1, Math.round(layout.width));
    canvas.height = Math.max(1, Math.round(layout.height));
    gl.viewport(0, 0, canvas.width, canvas.height);
    const background = hexToRgb(figure.layout.backgroundColor);
    gl.clearColor(background.r / 255, background.g / 255, background.b / 255, 1);
    gl.clear(gl.COLOR_BUFFER_BIT);

    const visibleTraces = figure.data.filter((trace) => trace.visible !== false);
    const traces2d = visibleTraces.filter(isCartesianTrace);
    const scatter3dTraces = visibleTraces.filter(isScatter3dTrace);
    const surfaceTraces = visibleTraces.filter(isSurfaceTrace);
    const mesh3dTraces = visibleTraces.filter(isMesh3dTrace);
    const has3d = scatter3dTraces.length > 0 || surfaceTraces.length > 0 || mesh3dTraces.length > 0;
    layer.style.cursor = has3d ? (this.dragPointerId !== null ? "grabbing" : "grab") : "default";
    if (traces2d.length === 0 && scatter3dTraces.length === 0 && surfaceTraces.length === 0 && mesh3dTraces.length === 0) {
      return;
    }

    gl.useProgram(this.program);
    gl.bindBuffer(gl.ARRAY_BUFFER, this.positionBuffer);
    const positionLocation = gl.getAttribLocation(this.program as WebGLProgram, "a_position");
    gl.enableVertexAttribArray(positionLocation);
    gl.vertexAttribPointer(positionLocation, 2, gl.FLOAT, false, 0, 0);

    const pointSizeLocation = gl.getUniformLocation(this.program as WebGLProgram, "u_pointSize");
    const colorLocation = gl.getUniformLocation(this.program as WebGLProgram, "u_color");
    const maxInteractivePoints = Math.max(100, Number(figure.config.maxInteractivePoints ?? 25_000));
    const maxRenderPoints = Math.max(500, Number(figure.config.maxRenderPoints ?? 12_000));
    let interactiveCounter = 0;

    if (has3d) {
      const projection = createProjection3d(layout.plotArea, canvas.width, canvas.height, this.camera);
      scatter3dTraces.forEach((trace, traceIndex) => {
        const color = trace.marker?.color ?? trace.line?.color ?? DEFAULT_SERIES_COLORS[traceIndex % DEFAULT_SERIES_COLORS.length];
        const rgb = hexToRgb(color);
        gl.uniform4f(colorLocation, rgb.r / 255, rgb.g / 255, rgb.b / 255, trace.marker?.opacity ?? trace.line?.opacity ?? 0.95);
        gl.uniform1f(pointSizeLocation, trace.marker?.size ?? 5);
        const vertices: number[] = [];
        const sampleStep = trace.z.length > maxRenderPoints ? Math.ceil(trace.z.length / maxRenderPoints) : 1;
        for (let pointIndex = 0; pointIndex < trace.z.length; pointIndex += sampleStep) {
          const x = Number(trace.x[pointIndex]);
          const y = Number(trace.y[pointIndex]);
          const z = Number(trace.z[pointIndex]);
          if (!Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(z)) {
            continue;
          }
          const projected = projection(x, y, z);
          vertices.push(projected.clipX, projected.clipY);
          if (interactiveCounter < maxInteractivePoints) {
            this.addInteractiveTarget(
              layer,
              {
                traceIndex,
                pointIndex,
                traceName: trace.name,
                x: `x:${formatNumeric(x)} y:${formatNumeric(y)}`,
                y: `z:${formatNumeric(z)}`
              },
              projected.px - 7,
              projected.py - 7,
              14,
              14
            );
            interactiveCounter += 1;
          }
        }
        this.drawVertexBuffer(gl, vertices, gl.POINTS);
      });

      surfaceTraces.forEach((trace, traceIndex) => {
        const color = trace.colorscale?.[0] ?? DEFAULT_SERIES_COLORS[(traceIndex + 2) % DEFAULT_SERIES_COLORS.length];
        const rgb = hexToRgb(color);
        gl.uniform4f(colorLocation, rgb.r / 255, rgb.g / 255, rgb.b / 255, 0.58);
        gl.uniform1f(pointSizeLocation, 1);
        const vertices: number[] = [];
        const rows = trace.z.length;
        const cols = trace.z[0]?.length ?? 0;
        for (let row = 0; row < rows - 1; row += 1) {
          for (let col = 0; col < cols - 1; col += 1) {
            const p0 = projection(col, row, Number(trace.z[row][col]));
            const p1 = projection(col + 1, row, Number(trace.z[row][col + 1]));
            const p2 = projection(col + 1, row + 1, Number(trace.z[row + 1][col + 1]));
            const p3 = projection(col, row + 1, Number(trace.z[row + 1][col]));
            if (![p0, p1, p2, p3].every((point) => Number.isFinite(point.clipX) && Number.isFinite(point.clipY))) {
              continue;
            }
            vertices.push(
              p0.clipX,
              p0.clipY,
              p1.clipX,
              p1.clipY,
              p2.clipX,
              p2.clipY,
              p0.clipX,
              p0.clipY,
              p2.clipX,
              p2.clipY,
              p3.clipX,
              p3.clipY
            );
          }
        }
        this.drawVertexBuffer(gl, vertices, gl.TRIANGLES);
      });

      mesh3dTraces.forEach((trace, traceIndex) => {
        const color = trace.marker?.color ?? DEFAULT_SERIES_COLORS[(traceIndex + 4) % DEFAULT_SERIES_COLORS.length];
        const rgb = hexToRgb(color);
        gl.uniform4f(colorLocation, rgb.r / 255, rgb.g / 255, rgb.b / 255, trace.marker?.opacity ?? 0.52);
        gl.uniform1f(pointSizeLocation, 1);
        const vertices: number[] = [];
        for (let index = 0; index < trace.i.length; index += 1) {
          const ia = Number(trace.i[index]);
          const jb = Number(trace.j[index]);
          const kc = Number(trace.k[index]);
          const pa = projection(Number(trace.x[ia]), Number(trace.y[ia]), Number(trace.z[ia]));
          const pb = projection(Number(trace.x[jb]), Number(trace.y[jb]), Number(trace.z[jb]));
          const pc = projection(Number(trace.x[kc]), Number(trace.y[kc]), Number(trace.z[kc]));
          if (![pa, pb, pc].every((point) => Number.isFinite(point.clipX) && Number.isFinite(point.clipY))) {
            continue;
          }
          vertices.push(pa.clipX, pa.clipY, pb.clipX, pb.clipY, pc.clipX, pc.clipY);
        }
        this.drawVertexBuffer(gl, vertices, gl.TRIANGLES);
      });
      return;
    }

    const domains = buildCartesianDomains(figure);
    if (!domains) {
      return;
    }
    const xAxisType = normalizeAxisType(figure.layout.xAxis, "category");
    traces2d.forEach((trace, traceIndex) => {
      const color = resolveTraceColor(trace, traceIndex);
      const rgb = hexToRgb(color);
      gl.uniform4f(colorLocation, rgb.r / 255, rgb.g / 255, rgb.b / 255, trace.marker?.opacity ?? trace.line?.opacity ?? 1);
      const markerSizeRaw = Array.isArray(trace.marker?.size) ? trace.marker?.size[0] : trace.marker?.size;
      const markerSize = Number(markerSizeRaw);
      gl.uniform1f(pointSizeLocation, Number.isFinite(markerSize) && markerSize > 0 ? markerSize : 5);

      const vertices: number[] = [];
      const sampleStep = trace.y.length > maxRenderPoints ? Math.ceil(trace.y.length / maxRenderPoints) : 1;
      for (let pointIndex = 0; pointIndex < trace.y.length; pointIndex += sampleStep) {
        const value = trace.y[pointIndex];
        if (!Number.isFinite(value)) {
          continue;
        }
        const xValue = toAxisScalar(trace.x[pointIndex], xAxisType, pointIndex);
        const xRatio = domains.x[1] === domains.x[0] ? 0 : (xValue - domains.x[0]) / (domains.x[1] - domains.x[0]);
        const yRatio = domains.y[1] === domains.y[0] ? 0 : (value - domains.y[0]) / (domains.y[1] - domains.y[0]);
        const clampedXRatio = clamp(xRatio, 0, 1);
        const clampedYRatio = clamp(yRatio, 0, 1);
        const px = layout.plotArea.x + clampedXRatio * layout.plotArea.width;
        const py = layout.plotArea.y + layout.plotArea.height - clampedYRatio * layout.plotArea.height;
        const clipX = (px / canvas.width) * 2 - 1;
        const clipY = -((py / canvas.height) * 2 - 1);
        vertices.push(clipX, clipY);

        if (interactiveCounter < maxInteractivePoints) {
          this.addInteractiveTarget(
            layer,
            {
              traceIndex,
              pointIndex,
              traceName: trace.name,
              x: String(trace.x[pointIndex] ?? pointIndex),
              y: formatNumeric(value)
            },
            px - 6,
            py - 6,
            12,
            12
          );
          interactiveCounter += 1;
        }
      }

      this.drawVertexBuffer(gl, vertices, gl.POINTS);
    });
  }

  resize(layout: ComputedLayout): void {
    const canvas = this.ensureCanvas();
    canvas.width = Math.max(1, Math.round(layout.width));
    canvas.height = Math.max(1, Math.round(layout.height));
  }

  getRootElement(): Element | null {
    return this.interactionLayer;
  }

  exportSvg(): string {
    if (!this.lastFigure || !this.lastLayout) {
      throw new ChartConfigurationError("CHART_EXPORT_UNAVAILABLE", "No rendered figure available for SVG export.");
    }
    const surrogate = document.createElement("div");
    Object.assign(surrogate.style, {
      width: `${this.lastLayout.width}px`,
      height: `${this.lastLayout.height}px`,
      position: "absolute",
      left: "-99999px",
      top: "-99999px"
    });
    document.body.append(surrogate);

    const renderer = new SvgRenderer();
    renderer.mount(surrogate);
    renderer.render(this.lastFigure, this.lastLayout);
    const serialized = renderer.exportSvg();
    renderer.destroy();
    surrogate.remove();
    return serialized;
  }

  async exportPng(): Promise<Blob> {
    const canvas = this.ensureCanvas();
    return await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (!blob) {
          reject(new ChartConfigurationError("CHART_EXPORT_FAILED", "Failed to export WebGL canvas as PNG."));
          return;
        }
        resolve(blob);
      }, "image/png");
    });
  }

  destroy(): void {
    if (this.interactionLayer) {
      this.interactionLayer.removeEventListener("pointerdown", this.handlePointerDown);
      this.interactionLayer.removeEventListener("pointermove", this.handlePointerMove);
      this.interactionLayer.removeEventListener("pointerup", this.handlePointerUp);
      this.interactionLayer.removeEventListener("pointercancel", this.handlePointerUp);
      this.interactionLayer.removeEventListener("wheel", this.handleWheel);
    }
    if (this.gl) {
      if (this.positionBuffer) {
        this.gl.deleteBuffer(this.positionBuffer);
      }
      if (this.program) {
        this.gl.deleteProgram(this.program);
      }
    }
    if (this.root) {
      this.root.remove();
    }

    this.container = null;
    this.root = null;
    this.canvas = null;
    this.interactionLayer = null;
    this.gl = null;
    this.program = null;
    this.positionBuffer = null;
    this.lastFigure = null;
    this.lastLayout = null;
    this.dragPointerId = null;
    this.camera = { ...DEFAULT_CAMERA_STATE };
  }

  private initializeProgram(): void {
    const gl = this.ensureGl();
    const vertexShader = compileShader(gl, gl.VERTEX_SHADER, VERTEX_SHADER_SOURCE);
    const fragmentShader = compileShader(gl, gl.FRAGMENT_SHADER, FRAGMENT_SHADER_SOURCE);
    const program = gl.createProgram();
    if (!program) {
      throw new ChartConfigurationError("CHART_WEBGL_UNAVAILABLE", "Unable to create WebGL program.");
    }
    gl.attachShader(program, vertexShader);
    gl.attachShader(program, fragmentShader);
    gl.linkProgram(program);
    if (!gl.getProgramParameter(program, gl.LINK_STATUS)) {
      const info = gl.getProgramInfoLog(program) ?? "Unknown WebGL link error.";
      throw new ChartConfigurationError("CHART_WEBGL_UNAVAILABLE", `Failed to link WebGL program: ${info}`);
    }
    gl.deleteShader(vertexShader);
    gl.deleteShader(fragmentShader);
    this.program = program;

    const positionBuffer = gl.createBuffer();
    if (!positionBuffer) {
      throw new ChartConfigurationError("CHART_WEBGL_UNAVAILABLE", "Unable to allocate WebGL position buffer.");
    }
    this.positionBuffer = positionBuffer;
  }

  private attachCameraListeners(layer: HTMLDivElement): void {
    layer.addEventListener("pointerdown", this.handlePointerDown);
    layer.addEventListener("pointermove", this.handlePointerMove);
    layer.addEventListener("pointerup", this.handlePointerUp);
    layer.addEventListener("pointercancel", this.handlePointerUp);
    layer.addEventListener("wheel", this.handleWheel, { passive: false });
  }

  private readonly handlePointerDown = (event: PointerEvent): void => {
    if (!this.has3dFigure()) {
      return;
    }
    this.dragPointerId = event.pointerId;
    this.lastPointerX = event.clientX;
    this.lastPointerY = event.clientY;
    this.dragMode = event.shiftKey || event.button === 1 ? "pan" : "rotate";
    this.ensureInteractionLayer().setPointerCapture?.(event.pointerId);
    event.preventDefault();
  };

  private readonly handlePointerMove = (event: PointerEvent): void => {
    if (!this.has3dFigure() || this.dragPointerId === null || event.pointerId !== this.dragPointerId) {
      return;
    }
    const dx = event.clientX - this.lastPointerX;
    const dy = event.clientY - this.lastPointerY;
    this.lastPointerX = event.clientX;
    this.lastPointerY = event.clientY;
    if (this.dragMode === "pan") {
      this.camera.panX += dx;
      this.camera.panY += dy;
    } else {
      this.camera.rotationY += dx * 0.009;
      this.camera.rotationX = clamp(this.camera.rotationX + dy * 0.009, -Math.PI / 2 + 0.05, Math.PI / 2 - 0.05);
    }
    this.rerenderLastFigure();
    event.preventDefault();
  };

  private readonly handlePointerUp = (event: PointerEvent): void => {
    if (this.dragPointerId === null || event.pointerId !== this.dragPointerId) {
      return;
    }
    this.ensureInteractionLayer().releasePointerCapture?.(event.pointerId);
    this.dragPointerId = null;
    event.preventDefault();
  };

  private readonly handleWheel = (event: WheelEvent): void => {
    if (!this.has3dFigure()) {
      return;
    }
    const zoomFactor = event.deltaY > 0 ? 0.92 : 1.08;
    this.camera.zoom = clamp(this.camera.zoom * zoomFactor, 0.35, 6);
    this.rerenderLastFigure();
    event.preventDefault();
  };

  private rerenderLastFigure(): void {
    if (!this.lastFigure || !this.lastLayout) {
      return;
    }
    this.render(this.lastFigure, this.lastLayout);
  }

  private has3dFigure(): boolean {
    if (!this.lastFigure) {
      return false;
    }
    return this.lastFigure.data.some((trace) => trace.visible !== false && (isScatter3dTrace(trace) || isSurfaceTrace(trace) || isMesh3dTrace(trace)));
  }

  private addInteractiveTarget(
    layer: HTMLElement,
    data: {
      traceIndex: number;
      pointIndex: number;
      traceName?: string;
      x: string;
      y: string;
    },
    x: number,
    y: number,
    width: number,
    height: number
  ): void {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.chartInteractive = "point";
    button.dataset.traceIndex = String(data.traceIndex);
    button.dataset.pointIndex = String(data.pointIndex);
    button.dataset.traceName = data.traceName?.trim() || `Trace ${data.traceIndex + 1}`;
    button.dataset.pointX = data.x;
    button.dataset.pointY = data.y;
    button.setAttribute("aria-label", `${button.dataset.traceName} ${data.x} ${data.y}`);
    Object.assign(button.style, {
      position: "absolute",
      left: `${Math.max(0, x)}px`,
      top: `${Math.max(0, y)}px`,
      width: `${Math.max(6, width)}px`,
      height: `${Math.max(6, height)}px`,
      border: "none",
      background: "transparent",
      margin: "0",
      padding: "0",
      cursor: "pointer"
    });
    layer.append(button);
  }

  private drawVertexBuffer(gl: WebGLRenderingContext, vertices: number[], primitive: number): void {
    const pointsCount = vertices.length / 2;
    if (pointsCount === 0) {
      return;
    }
    gl.bufferData(gl.ARRAY_BUFFER, new Float32Array(vertices), gl.STATIC_DRAW);
    gl.drawArrays(primitive, 0, pointsCount);
  }

  private ensureCanvas(): HTMLCanvasElement {
    if (!this.canvas) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_MOUNTED", "WebGL renderer must be mounted before rendering.");
    }
    return this.canvas;
  }

  private ensureInteractionLayer(): HTMLDivElement {
    if (!this.interactionLayer) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_MOUNTED", "WebGL interaction layer is unavailable.");
    }
    return this.interactionLayer;
  }

  private ensureGl(): WebGLRenderingContext {
    if (!this.gl) {
      throw new ChartConfigurationError("CHART_RENDERER_NOT_MOUNTED", "WebGL context is unavailable.");
    }
    return this.gl;
  }
}

const compileShader = (gl: WebGLRenderingContext, type: number, source: string): WebGLShader => {
  const shader = gl.createShader(type);
  if (!shader) {
    throw new ChartConfigurationError("CHART_WEBGL_UNAVAILABLE", "Unable to create WebGL shader.");
  }
  gl.shaderSource(shader, source);
  gl.compileShader(shader);
  if (!gl.getShaderParameter(shader, gl.COMPILE_STATUS)) {
    const info = gl.getShaderInfoLog(shader) ?? "Unknown WebGL shader compile error.";
    gl.deleteShader(shader);
    throw new ChartConfigurationError("CHART_WEBGL_UNAVAILABLE", `Failed to compile WebGL shader: ${info}`);
  }
  return shader;
};

const resolveTraceColor = (trace: CartesianTrace, index: number): string =>
  (Array.isArray(trace.marker?.color) ? trace.marker?.color[0] : trace.marker?.color) ??
  trace.line?.color ??
  DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];

const hexToRgb = (hex: string): { r: number; g: number; b: number } => {
  const normalized = hex.replace("#", "");
  if (normalized.length === 3) {
    const [r, g, b] = normalized.split("").map((char) => parseInt(char + char, 16));
    return { r, g, b };
  }
  if (!/^[0-9a-f]{6}$/i.test(normalized)) {
    return { r: 37, g: 99, b: 235 };
  }
  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16)
  };
};

const formatNumeric = (value: number): string => {
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

const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

const createProjection3d = (
  plotArea: ComputedLayout["plotArea"],
  canvasWidth: number,
  canvasHeight: number,
  camera: CameraState3d
): ((x: number, y: number, z: number) => { clipX: number; clipY: number; px: number; py: number }) => {
  const angleY = camera.rotationY;
  const angleX = camera.rotationX;
  const scale = Math.min(plotArea.width, plotArea.height) * 0.34 * camera.zoom;
  const centerX = plotArea.x + plotArea.width / 2;
  const centerY = plotArea.y + plotArea.height / 2;
  const cosY = Math.cos(angleY);
  const sinY = Math.sin(angleY);
  const cosX = Math.cos(angleX);
  const sinX = Math.sin(angleX);

  return (x: number, y: number, z: number) => {
    const x1 = x * cosY - z * sinY;
    const z1 = x * sinY + z * cosY;
    const y2 = y * cosX - z1 * sinX;
    const z2 = y * sinX + z1 * cosX;
    const perspective = 1 / Math.max(0.2, 1 + z2 * 0.045);
    const px = centerX + camera.panX + x1 * scale * perspective;
    const py = centerY + camera.panY - y2 * scale * perspective;
    return {
      clipX: clamp((px / Math.max(1, canvasWidth)) * 2 - 1, -1.2, 1.2),
      clipY: clamp(-((py / Math.max(1, canvasHeight)) * 2 - 1), -1.2, 1.2),
      px,
      py
    };
  };
};
