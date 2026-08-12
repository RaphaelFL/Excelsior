export interface LinePoint {
  x: number;
  y: number;
}

export interface PlotRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export const simplifyLine = (points: LinePoint[], tolerance = 1.1): LinePoint[] => {
  if (points.length <= 2) {
    return points;
  }
  const squaredTolerance = tolerance * tolerance;
  const simplified = [points[0]];
  simplifyDfs(points, 0, points.length - 1, squaredTolerance, simplified);
  simplified.push(points[points.length - 1]);
  return dedupeSequential(simplified);
};

export const clipLineToRect = (points: LinePoint[], rect: PlotRect): LinePoint[] => {
  if (points.length < 2) {
    return points.filter((point) => insideRect(point.x, point.y, rect));
  }
  const clipped: LinePoint[] = [];
  for (let index = 1; index < points.length; index += 1) {
    const start = points[index - 1];
    const end = points[index];
    const segment = clipSegment(start, end, rect);
    if (!segment) {
      continue;
    }
    if (clipped.length === 0 || !isSamePoint(clipped[clipped.length - 1], segment[0])) {
      clipped.push(segment[0]);
    }
    clipped.push(segment[1]);
  }
  return dedupeSequential(clipped);
};

const simplifyDfs = (
  points: LinePoint[],
  start: number,
  end: number,
  squaredTolerance: number,
  output: LinePoint[]
): void => {
  let maxDistance = squaredTolerance;
  let nextIndex = -1;
  const a = points[start];
  const b = points[end];

  for (let index = start + 1; index < end; index += 1) {
    const distance = squaredDistanceToSegment(points[index], a, b);
    if (distance > maxDistance) {
      nextIndex = index;
      maxDistance = distance;
    }
  }

  if (nextIndex === -1) {
    return;
  }

  if (nextIndex - start > 1) {
    simplifyDfs(points, start, nextIndex, squaredTolerance, output);
  }
  output.push(points[nextIndex]);
  if (end - nextIndex > 1) {
    simplifyDfs(points, nextIndex, end, squaredTolerance, output);
  }
};

const squaredDistanceToSegment = (point: LinePoint, start: LinePoint, end: LinePoint): number => {
  let x = start.x;
  let y = start.y;
  let dx = end.x - x;
  let dy = end.y - y;

  if (dx !== 0 || dy !== 0) {
    const t = ((point.x - x) * dx + (point.y - y) * dy) / (dx * dx + dy * dy);
    if (t > 1) {
      x = end.x;
      y = end.y;
    } else if (t > 0) {
      x += dx * t;
      y += dy * t;
    }
  }

  dx = point.x - x;
  dy = point.y - y;
  return dx * dx + dy * dy;
};

const clipSegment = (a: LinePoint, b: LinePoint, rect: PlotRect): [LinePoint, LinePoint] | null => {
  const minX = rect.x;
  const maxX = rect.x + rect.width;
  const minY = rect.y;
  const maxY = rect.y + rect.height;
  let x0 = a.x;
  let y0 = a.y;
  let x1 = b.x;
  let y1 = b.y;

  let outCode0 = computeOutCode(x0, y0, minX, maxX, minY, maxY);
  let outCode1 = computeOutCode(x1, y1, minX, maxX, minY, maxY);

  while (true) {
    if (!(outCode0 | outCode1)) {
      return [
        { x: x0, y: y0 },
        { x: x1, y: y1 }
      ];
    }
    if (outCode0 & outCode1) {
      return null;
    }

    const outCodeOut = outCode0 ? outCode0 : outCode1;
    let x = 0;
    let y = 0;

    if (outCodeOut & 8) {
      x = x0 + ((x1 - x0) * (maxY - y0)) / (y1 - y0 || Number.EPSILON);
      y = maxY;
    } else if (outCodeOut & 4) {
      x = x0 + ((x1 - x0) * (minY - y0)) / (y1 - y0 || Number.EPSILON);
      y = minY;
    } else if (outCodeOut & 2) {
      y = y0 + ((y1 - y0) * (maxX - x0)) / (x1 - x0 || Number.EPSILON);
      x = maxX;
    } else {
      y = y0 + ((y1 - y0) * (minX - x0)) / (x1 - x0 || Number.EPSILON);
      x = minX;
    }

    if (outCodeOut === outCode0) {
      x0 = x;
      y0 = y;
      outCode0 = computeOutCode(x0, y0, minX, maxX, minY, maxY);
    } else {
      x1 = x;
      y1 = y;
      outCode1 = computeOutCode(x1, y1, minX, maxX, minY, maxY);
    }
  }
};

const computeOutCode = (x: number, y: number, minX: number, maxX: number, minY: number, maxY: number): number => {
  let code = 0;
  if (x < minX) {
    code |= 1;
  } else if (x > maxX) {
    code |= 2;
  }
  if (y < minY) {
    code |= 4;
  } else if (y > maxY) {
    code |= 8;
  }
  return code;
};

const insideRect = (x: number, y: number, rect: PlotRect): boolean =>
  x >= rect.x && x <= rect.x + rect.width && y >= rect.y && y <= rect.y + rect.height;

const dedupeSequential = (points: LinePoint[]): LinePoint[] => {
  if (points.length <= 1) {
    return points;
  }
  const output: LinePoint[] = [points[0]];
  for (let index = 1; index < points.length; index += 1) {
    if (!isSamePoint(points[index], output[output.length - 1])) {
      output.push(points[index]);
    }
  }
  return output;
};

const isSamePoint = (left: LinePoint, right: LinePoint): boolean => Math.abs(left.x - right.x) < 0.001 && Math.abs(left.y - right.y) < 0.001;
