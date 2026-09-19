#!/usr/bin/env bun
import { audit } from "./audit";
import { formatJson, formatText } from "./report";

function usage(): string {
  return [
    "Usage: ooxml-a11y <file...> [--json] [--quiet]",
    "",
    "Audit DOCX and PPTX files for accessibility issues.",
    "",
    "Options:",
    "  --json     Print machine-readable JSON for each file",
    "  --quiet    Print only a one-line summary for each file",
    "  -h, --help Show this message",
  ].join("\n");
}

function splitArgs(argv: string[]): {
  files: string[];
  json: boolean;
  quiet: boolean;
  help: boolean;
} {
  const files: string[] = [];
  let json = false;
  let quiet = false;
  let help = false;

  for (const arg of argv) {
    if (arg === "--json") json = true;
    else if (arg === "--quiet" || arg === "-q") quiet = true;
    else if (arg === "--help" || arg === "-h") help = true;
    else if (arg.startsWith("-")) {
      console.error(`ooxml-a11y: unknown option ${arg}`);
      help = true;
    } else files.push(arg);
  }

  return { files, json, quiet, help };
}

async function main(argv: string[]): Promise<number> {
  const { files, json, quiet, help } = splitArgs(argv);

  if (help || files.length === 0) {
    if (help) console.log(usage());
    else console.error(usage());
    return help ? 0 : 2;
  }

  let errorCount = 0;

  for (const file of files) {
    let data: Uint8Array;
    try {
      data = new Uint8Array(await Bun.file(file).arrayBuffer());
    } catch (err) {
      console.error(
        `ooxml-a11y: cannot read ${file}: ${(err as Error).message}`,
      );
      errorCount += 1;
      continue;
    }

    const result = audit(data, file);
    errorCount += result.counts.error;

    if (quiet) {
      console.log(
        `${file}: ${result.counts.error} error(s), ${result.counts.warning} warning(s), ${result.counts.info} info`,
      );
    } else if (json) {
      console.log(formatJson(result));
    } else {
      console.log(formatText(result));
    }
  }

  return errorCount > 0 ? 1 : 0;
}

if (import.meta.main) {
  main(process.argv.slice(2))
    .then((code) => process.exit(code))
    .catch((err: unknown) => {
      console.error(err);
      process.exit(1);
    });
}
