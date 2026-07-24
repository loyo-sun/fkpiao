import { cp, mkdir, rm } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const source = resolve(root, "node_modules/pdfjs-dist");
const target = resolve(root, "public/pdfjs");

await rm(target, { recursive: true, force: true });
await mkdir(target, { recursive: true });

for (const directory of ["cmaps", "standard_fonts", "wasm"]) {
  await cp(resolve(source, directory), resolve(target, directory), { recursive: true });
}
