import { mkdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import { dirname } from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";

const entry = fileURLToPath(new URL("../src/index.ts", import.meta.url));
const outfile = fileURLToPath(
  new URL("../site/assets/demo.js", import.meta.url),
);

async function main() {
  if (!existsSync(entry)) {
    process.stdout.write(
      `skipping build:site: ${entry} not found (expected this repo's src/index.ts)\n`,
    );
    return;
  }

  await mkdir(dirname(outfile), { recursive: true });

  await esbuild.build({
    entryPoints: [entry],
    outfile,
    bundle: true,
    format: "iife",
    globalName: "OfficeLens",
    platform: "browser",
    target: "es2020",
    minify: true,
    logLevel: "info",
  });

  process.stdout.write("built site/assets/demo.js\n");
}

main().catch((error) => {
  process.stderr.write(`${error?.stack ?? error}\n`);
  process.exit(1);
});
