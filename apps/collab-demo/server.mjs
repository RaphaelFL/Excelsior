import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { WorkbookEngine } from "@excelsior/core";
import { BasicFormulaEngine } from "@excelsior/formulas";

const modulePath = fileURLToPath(import.meta.url);
const __dirname = path.dirname(modulePath);
const distDir = path.join(__dirname, "dist");
const port = Number(process.env.PORT || 4177);
const MAX_JSON_BODY_BYTES = 64 * 1024;
const MAX_OPERATION_BATCH = 5_000;
const COMMON_SECURITY_HEADERS = {
  "x-content-type-options": "nosniff",
  "x-frame-options": "DENY",
  "referrer-policy": "no-referrer",
  "cross-origin-opener-policy": "same-origin",
  "cross-origin-resource-policy": "same-origin"
};
const HTML_SECURITY_HEADERS = {
  ...COMMON_SECURITY_HEADERS,
  "content-security-policy": "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; connect-src 'self'; img-src 'self' data:; font-src 'self' data:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'"
};
const STATIC_CONTENT_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8"
};

class HttpRequestError extends Error {
  constructor(statusCode, code, message) {
    super(message);
    this.name = code;
    this.statusCode = statusCode;
  }
}

export const createInitialWorkbook = () => ({
  id: "collab-workbook",
  activeSheetId: "sheet-revenue",
  metadata: { demo: true },
  settings: {
    maxRows: 1000,
    maxColumns: 100,
    rowHeight: 28,
    columnWidth: 120,
    viewportBuffer: 6,
    maxHistorySize: 100,
    enableFormulas: true,
    clipboardPolicy: "safe-html"
  },
  sheets: [
    {
      id: "sheet-revenue",
      name: "Revenue",
      rowCount: 200,
      columnCount: 24,
      selection: {
        start: { row: 0, col: 0 },
        end: { row: 0, col: 0 }
      },
      cells: {
        "0:0": { value: "Region", computedValue: "Region" },
        "0:1": { value: "Q1", computedValue: "Q1" },
        "0:2": { value: "Q2", computedValue: "Q2" },
        "1:0": { value: "LATAM", computedValue: "LATAM" },
        "1:1": { value: 120000, computedValue: 120000 },
        "1:2": { value: 140000, computedValue: 140000 }
      }
    },
    {
      id: "sheet-summary",
      name: "Summary",
      rowCount: 120,
      columnCount: 24,
      selection: {
        start: { row: 0, col: 0 },
        end: { row: 0, col: 0 }
      },
      cells: {
        "0:0": { value: "='Revenue'!B2+'Revenue'!C2", formula: "='Revenue'!B2+'Revenue'!C2", computedValue: 260000 }
      }
    }
  ]
});

const isPathInsideRoot = (rootPath, targetPath) => {
  const relative = path.relative(rootPath, targetPath);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
};

const getRequestUrl = (request) => new URL(request.url, `http://${request.headers.host ?? "localhost"}`);

const isAssetRequest = (relativePath) => path.extname(relativePath) !== "";

const isAllowedOrigin = (request) => {
  const origin = request.headers.origin;
  if (!origin) {
    return true;
  }

  const host = request.headers.host;
  if (!host) {
    return false;
  }

  try {
    return new URL(origin).host === host;
  } catch {
    return false;
  }
};

export const resolveStaticAssetPath = (request) => {
  const requestUrl = getRequestUrl(request);
  let pathname;

  try {
    pathname = decodeURIComponent(requestUrl.pathname);
  } catch {
    throw new HttpRequestError(400, "INVALID_PATH", "Invalid request path.");
  }

  const normalizedPath = path.posix.normalize(pathname);
  const relativePath = normalizedPath === "/" ? "index.html" : normalizedPath.replace(/^\/+/, "");
  const filePath = path.resolve(distDir, relativePath);

  if (!isPathInsideRoot(distDir, filePath)) {
    throw new HttpRequestError(403, "STATIC_PATH_FORBIDDEN", "Static asset path is not allowed.");
  }

  return { filePath, relativePath };
};

const isValidOperation = (operation) =>
  Boolean(
    operation &&
      typeof operation === "object" &&
      typeof operation.op === "string" &&
      typeof operation.id === "string" &&
      Array.isArray(operation.path)
  );

export const isValidOpsPayload = (payload) =>
  Boolean(
    payload &&
      typeof payload === "object" &&
      typeof payload.clientId === "string" &&
      payload.clientId.length > 0 &&
      payload.clientId.length <= 128 &&
      Array.isArray(payload.operations) &&
      payload.operations.length <= MAX_OPERATION_BATCH &&
      payload.operations.every(isValidOperation)
  );

const sendJson = (response, statusCode, body) => {
  response.writeHead(statusCode, {
    ...COMMON_SECURITY_HEADERS,
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store"
  });
  response.end(JSON.stringify(body));
};

const sendError = (response, error) => {
  if (error instanceof HttpRequestError) {
    sendJson(response, error.statusCode, {
      error: error.name,
      message: error.message
    });
    return;
  }

  console.error(error);
  sendJson(response, 500, {
    error: "INTERNAL_ERROR",
    message: "Internal server error."
  });
};

const broadcast = (clients, payload) => {
  const serialized = `data: ${JSON.stringify(payload)}\n\n`;
  for (const client of clients) {
    client.write(serialized);
  }
};

const readJsonBody = async (request) => {
  const contentType = request.headers["content-type"];
  if (!contentType || !contentType.toLowerCase().startsWith("application/json")) {
    throw new HttpRequestError(415, "UNSUPPORTED_MEDIA_TYPE", "Request body must be JSON.");
  }

  const chunks = [];
  let totalBytes = 0;

  for await (const chunk of request) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    totalBytes += buffer.length;

    if (totalBytes > MAX_JSON_BODY_BYTES) {
      throw new HttpRequestError(413, "PAYLOAD_TOO_LARGE", `Request body exceeds ${MAX_JSON_BODY_BYTES} bytes.`);
    }

    chunks.push(buffer);
  }

  if (totalBytes === 0) {
    throw new HttpRequestError(400, "EMPTY_BODY", "Request body cannot be empty.");
  }

  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw new HttpRequestError(400, "INVALID_JSON", "Request body is not valid JSON.");
  }
};

const serveStatic = async (request, response) => {
  const { filePath, relativePath } = resolveStaticAssetPath(request);

  try {
    const fileStats = await stat(filePath);
    if (!fileStats.isFile()) {
      throw new Error("not file");
    }

    const content = await readFile(filePath);
    const extension = path.extname(filePath);

    response.writeHead(200, {
      ...(extension === ".html" ? HTML_SECURITY_HEADERS : COMMON_SECURITY_HEADERS),
      "content-type": STATIC_CONTENT_TYPES[extension] ?? "application/octet-stream"
    });
    response.end(content);
  } catch {
    if (isAssetRequest(relativePath)) {
      throw new HttpRequestError(404, "STATIC_ASSET_NOT_FOUND", "Static asset was not found.");
    }

    const fallback = await readFile(path.join(distDir, "index.html"));
    response.writeHead(200, {
      ...HTML_SECURITY_HEADERS,
      "content-type": "text/html; charset=utf-8"
    });
    response.end(fallback);
  }
};

export const createCollabDemoServer = () => {
  const engine = WorkbookEngine.fromJSON(createInitialWorkbook(), new BasicFormulaEngine());
  const clients = new Set();

  return createServer(async (request, response) => {
    try {
      if (!request.url) {
        throw new HttpRequestError(400, "REQUEST_URL_REQUIRED", "Request URL is required.");
      }

      if (request.method === "GET" && request.url === "/api/workbook") {
        sendJson(response, 200, engine.toJSON());
        return;
      }

      if (request.method === "GET" && request.url.startsWith("/api/events")) {
        response.writeHead(200, {
          ...COMMON_SECURITY_HEADERS,
          "content-type": "text/event-stream",
          "cache-control": "no-cache, no-transform",
          connection: "keep-alive"
        });
        response.write("retry: 1000\n\n");
        clients.add(response);
        request.on("close", () => clients.delete(response));
        return;
      }

      if (request.method === "POST" && request.url === "/api/ops") {
        if (!isAllowedOrigin(request)) {
          throw new HttpRequestError(403, "ORIGIN_FORBIDDEN", "Origin is not allowed for this endpoint.");
        }

        const payload = await readJsonBody(request);
        if (!isValidOpsPayload(payload)) {
          throw new HttpRequestError(400, "INVALID_OPS_PAYLOAD", "Operations payload is invalid.");
        }

        engine.applyOperations(payload.operations);
        broadcast(clients, { type: "ops", clientId: payload.clientId, operations: payload.operations });
        sendJson(response, 200, { ok: true });
        return;
      }

      if (request.method === "POST" && request.url === "/api/reset") {
        if (!isAllowedOrigin(request)) {
          throw new HttpRequestError(403, "ORIGIN_FORBIDDEN", "Origin is not allowed for this endpoint.");
        }

        engine.loadFromJSON(createInitialWorkbook());
        broadcast(clients, { type: "reset", clientId: "server", operations: [] });
        sendJson(response, 200, { ok: true });
        return;
      }

      await serveStatic(request, response);
    } catch (error) {
      sendError(response, error);
    }
  });
};

const isDirectExecution = process.argv[1] ? path.resolve(process.argv[1]) === modulePath : false;

if (isDirectExecution) {
  const server = createCollabDemoServer();

  server.listen(port, () => {
    console.log(`Collaboration demo running at http://localhost:${port}`);
  });
}