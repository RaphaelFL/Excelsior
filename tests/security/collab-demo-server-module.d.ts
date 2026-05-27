declare module "../../apps/collab-demo/server.mjs" {
  import type { Server } from "node:http";

  export function createCollabDemoServer(): Server;

  export function isValidOpsPayload(payload: unknown): boolean;

  export function resolveStaticAssetPath(request: {
    url: string;
    headers: {
      host?: string;
    };
  }): {
    filePath: string;
    relativePath: string;
  };
}