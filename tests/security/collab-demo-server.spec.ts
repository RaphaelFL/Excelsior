import type { AddressInfo } from "node:net";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
// @ts-expect-error Runtime security test imports a JavaScript server module.
import { createCollabDemoServer, isValidOpsPayload, resolveStaticAssetPath } from "../../apps/collab-demo/server.mjs";

describe("collab demo server security", () => {
  let server: ReturnType<typeof createCollabDemoServer>;
  let baseUrl: string;

  beforeEach(async () => {
    server = createCollabDemoServer();
    await new Promise<void>((resolve, reject) => {
      server.listen(0, "127.0.0.1", () => resolve());
      server.on("error", reject);
    });
    const address = server.address() as AddressInfo;
    baseUrl = `http://127.0.0.1:${address.port}`;
  });

  afterEach(async () => {
    await new Promise<void>((resolve, reject) => {
      server.close((error: Error | undefined) => {
        if (error) {
          reject(error);
          return;
        }
        resolve();
      });
    });
  });

  it("rejects encoded static path traversal attempts", async () => {
    const response = await fetch(`${baseUrl}/%2e%2e/server.mjs`);
    const body = await response.json();

    expect(response.status).toBe(404);
    expect(body).toMatchObject({
      error: "STATIC_ASSET_NOT_FOUND"
    });
  });

  it("rejects oversized JSON bodies for operations", async () => {
    const oversized = JSON.stringify({
      clientId: "attacker",
      operations: [],
      padding: "x".repeat(70_000)
    });

    const response = await fetch(`${baseUrl}/api/ops`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        origin: baseUrl
      },
      body: oversized
    });

    expect(response.status).toBe(413);
    await expect(response.json()).resolves.toMatchObject({
      error: "PAYLOAD_TOO_LARGE"
    });
  });

  it("rejects malformed operation payloads", async () => {
    const response = await fetch(`${baseUrl}/api/ops`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        origin: baseUrl
      },
      body: JSON.stringify({ clientId: "demo", operations: [{ id: "1" }] })
    });

    expect(response.status).toBe(400);
    await expect(response.json()).resolves.toMatchObject({
      error: "INVALID_OPS_PAYLOAD"
    });
  });

  it("adds no-sniff headers to JSON API responses", async () => {
    const response = await fetch(`${baseUrl}/api/workbook`);

    expect(response.status).toBe(200);
    expect(response.headers.get("x-content-type-options")).toBe("nosniff");
    expect(response.headers.get("cross-origin-opener-policy")).toBe("same-origin");
  });

  it("validates operation payload shape locally", () => {
    expect(
      isValidOpsPayload({
        clientId: "demo",
        operations: [{ op: "replace", id: "op-1", path: ["sheets", 0, "cells"] }]
      })
    ).toBe(true);
    expect(isValidOpsPayload({ clientId: "demo", operations: [{ id: "op-1" }] })).toBe(false);
  });

  it("normalizes encoded static paths inside dist", () => {
    const resolved = resolveStaticAssetPath({
      url: "/%2e%2e/server.mjs",
      headers: { host: "localhost:4177" }
    });

    expect(resolved.relativePath).toBe("server.mjs");
  });
});