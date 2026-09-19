#!/usr/bin/env bun
import { audit } from "./audit";
import { formatJson, formatText } from "./report";
import { writeSarif } from "./sarif";
import type { AuditResult } from "./types";

type FailOn = "error" | "warning" | "info" | "none";

const FAIL_ON_LEVELS: FailOn[] = ["error", "warning", "info", "none"];
const TOOL_NAME = "officelens";

function usage(): string {
  return [
    "Usage: officelens <file...> [options]",
    "",
    "Audit DOCX and PPTX files for accessibility issues.",
    "",
    "Options:",
    "  --json               Print machine-readable JSON for each file",
    "  --quiet, -q          Print only a one-line summary for each file",
    "  --sarif <path>       Write a SARIF 2.1.0 report to <path>",
    "  --fail-on <level>    Exit 1 on: error (default), warning, info, none",
    "  --version, -v        Print the version and exit",
    "  -h, --help           Show this message",
  ].join("\n");
}

interface Options {
  files: string[];
  json: boolean;
  quiet: boolean;
  help: boolean;
  version: boolean;
  sarif?: string;
  failOn: FailOn;
  unknown: boolean;
}

function splitArgs(argv: string[]): Options {
  const files: string[] = [];
  let json = false;
  let quiet = false;
  let help = false;
  let version = false;
  let sarif: string | undefined;
  let failOn: FailOn = "error";
  let unknown = false;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]!;
    if (arg === "--json") json = true;
    else if (arg === "--quiet" || arg === "-q") quiet = true;
    else if (arg === "--help" || arg === "-h") help = true;
    else if (arg === "--version" || arg === "-v") version = true;
    else if (arg === "--sarif") {
      const value = argv[i + 1];
      if (value === undefined || value.startsWith("-")) {
        console.error("officelens: --sarif requires a path");
        unknown = true;
      } else {
        sarif = value;
        i += 1;
      }
    } else if (arg === "--fail-on") {
      const value = argv[i + 1];
      if (value === undefined || !FAIL_ON_LEVELS.includes(value as FailOn)) {
        console.error("officelens: --fail-on must be one of error, warning, info, none");
        unknown = true;
      } else {
        failOn = value as FailOn;
        i += 1;
      }
    } else if (arg.startsWith("-")) {
      console.error(`officelens: unknown option ${arg}`);
      unknown = true;
    } else files.push(arg);
  }

  return { files, json, quiet, help, version, sarif, failOn, unknown };
}

async function readVersion(): Promise<string> {
  try {
    const pkg = (await Bun.file(new URL("../package.json", import.meta.url)).json()) as {
      version?: string;
    };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
}

function exceedsThreshold(results: AuditResult[], failOn: FailOn): boolean {
  if (failOn === "none") return false;
  let error = 0;
  let warning = 0;
  let info = 0;
  for (const result of results) {
    error += result.counts.error;
    warning += result.counts.warning;
    info += result.counts.info;
  }
  if (failOn === "error") return error > 0;
  if (failOn === "warning") return error + warning > 0;
  return error + warning + info > 0;
}

async function main(argv: string[]): Promise<number> {
  const opts = splitArgs(argv);

  if (opts.help) {
    console.log(usage());
    return 0;
  }
  if (opts.unknown) {
    console.error(usage());
    return 2;
  }

  const version = await readVersion();

  if (opts.version) {
    console.log(version);
    return 0;
  }

  if (opts.files.length === 0) {
    console.error(usage());
    return 2;
  }

  const results: AuditResult[] = [];
  let readErrors = 0;

  for (const file of opts.files) {
    let data: Uint8Array;
    try {
      data = new Uint8Array(await Bun.file(file).arrayBuffer());
    } catch (err) {
      console.error(
        `officelens: cannot read ${file}: ${(err as Error).message}`,
      );
      readErrors += 1;
      continue;
    }

    const result = audit(data, file);
    results.push(result);

    if (opts.quiet) {
      console.log(
        `${file}: ${result.counts.error} error(s), ${result.counts.warning} warning(s), ${result.counts.info} info`,
      );
    } else if (opts.json) {
      console.log(formatJson(result));
    } else {
      console.log(formatText(result));
    }
  }

  if (opts.sarif) {
    await writeSarif(opts.sarif, results, TOOL_NAME, version);
  }

  if (readErrors > 0) return 1;
  return exceedsThreshold(results, opts.failOn) ? 1 : 0;
}

if (import.meta.main) {
  main(process.argv.slice(2))
    .then((code) => process.exit(code))
    .catch((err: unknown) => {
      console.error(err);
      process.exit(1);
    });
}
