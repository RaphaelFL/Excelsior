import { readFile, readdir } from "node:fs/promises";
import { join, relative } from "node:path";
import { fileURLToPath } from "node:url";
import process from "node:process";

const root = new URL("../", import.meta.url);
const rootPath = fileURLToPath(root);
const runtimePackages = ["charts", "core", "devtools", "formulas", "renderer-dom", "styles", "vanilla", "xlsx"];
const sourceExtensions = new Set([".js", ".jsx", ".ts", ".tsx", ".mjs", ".cjs"]);
const externalImports = [];
const externalDependencies = [];

const walk = async (directory) => {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) files.push(...await walk(path));
    else files.push(path);
  }
  return files;
};

for (const packageName of runtimePackages) {
  const packageDirectory = new URL(`packages/${packageName}/`, root);
  const manifest = JSON.parse(await readFile(new URL("package.json", packageDirectory), "utf8"));
  for (const [dependency] of Object.entries(manifest.dependencies ?? {})) {
    if (!dependency.startsWith("@excelsior/")) externalDependencies.push(`${manifest.name}: ${dependency}`);
  }

  const sourceDirectory = new URL("src/", packageDirectory);
  for (const file of await walk(fileURLToPath(sourceDirectory))) {
    const extension = file.slice(file.lastIndexOf("."));
    if (!sourceExtensions.has(extension)) continue;
    const source = await readFile(file, "utf8");
    const imports = source.matchAll(/(?:from\s+|import\s*\(|require\s*\()\s*["']([^"']+)["']/g);
    for (const match of imports) {
      const specifier = match[1];
      if (!specifier.startsWith(".") && !specifier.startsWith("@excelsior/") && !specifier.startsWith("node:")) {
        externalImports.push(`${relative(rootPath, file)}: ${specifier}`);
      }
    }
  }
}

if (externalDependencies.length || externalImports.length) {
  console.error("Runtime externo detectado.");
  for (const finding of [...externalDependencies, ...externalImports]) console.error(`- ${finding}`);
  process.exit(1);
}

console.log(`Runtime auditado: ${runtimePackages.length} pacotes, sem dependências ou imports externos.`);
