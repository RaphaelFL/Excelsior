import path from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const rootDir = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  test: {
    environment: "jsdom",
    include: ["packages/**/*.spec.ts", "tests/**/*.spec.ts"],
    coverage: {
      reporter: ["text", "html"]
    }
  },
  resolve: {
    alias: {
      "@excelsior/core": path.resolve(rootDir, "packages/core/src/index.ts"),
      "@excelsior/devtools": path.resolve(rootDir, "packages/devtools/src/index.ts"),
      "@excelsior/formulas": path.resolve(rootDir, "packages/formulas/src/index.ts"),
      "@excelsior/renderer-dom": path.resolve(rootDir, "packages/renderer-dom/src/index.ts"),
      "@excelsior/xlsx": path.resolve(rootDir, "packages/xlsx/src/index.ts"),
      "@excelsior/vanilla": path.resolve(rootDir, "packages/vanilla/src/index.ts")
    }
  }
});