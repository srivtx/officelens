#!/usr/bin/env bun
import { readdirSync } from "node:fs";
import { join } from "node:path";
import { audit } from "./audit";
import { formatJson, formatText } from "./report";
import { writeSarif } from "./sarif";
import type { AuditResult } from "./types";

export type FailOn = "error" | "warning" | "info" | "none";

export const FAIL_ON_LEVELS: FailOn[] = ["error", "warning", "info", "none"];
const TOOL_NAME = "officelens";
const DOC_EXTENSIONS = [".docx", ".pptx"];

export const USAGE = `officelens - offline accessibility auditor for DOCX and PPTX

Usage:
  officelens <file...> [options]
  officelens --dir <path> [options]

Options:
  --dir <path>         Audit every .docx/.pptx file in <path> (non-recursive, sorted)
  --json               Print machine-readable JSON; one object for a single file,
                       a JSON array when auditing more than one
  --quiet, -q          Print only a one-line summary for each file
  --sarif <path>       Write a SARIF 2.1.0 report to <path>
  --fail-on <level>    Exit 1 on: error (default), warning, info, none
  --version, -v        Print the version and exit
  -h, --help           Show this message

Value flags accept either "--flag value" or "--flag=value". Use "--" to stop
option parsing; every following argument is treated as a file path.

Exit codes:
  0  no findings at or above --fail-on
  1  findings at or above --fail-on
  2  invalid usage, or a package that cannot be parsed as OOXML
  3  an input file or directory could not be read, or the report could not be written
`;

export interface Options {
  files: string[];
  json: boolean;
  quiet: boolean;
  dir: string | null;
  help: boolean;
  version: boolean;
  sarif: string | null;
  failOn: FailOn;
}

export function parseArgs(argv: string[]): Options {
  const files: string[] = [];
  let json = false;
  let quiet = false;
  let help = false;
  let version = false;
  let sarif: string | null = null;
  let dir: string | null = null;
  let failOn: FailOn = "error";
  let endOfOptions = false;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === undefined) continue;

    if (endOfOptions) {
      files.push(arg);
      continue;
    }
    if (arg === "--") {
      endOfOptions = true;
      continue;
    }
    if (arg === "--json") {
      json = true;
      continue;
    }
    if (arg === "--quiet" || arg === "-q") {
      quiet = true;
      continue;
    }
    if (arg === "--help" || arg === "-h") {
      help = true;
      continue;
    }
    if (arg === "--version" || arg === "-v") {
      version = true;
      continue;
    }

    const readValue = (name: string): string | undefined => {
      if (arg === name) {
        const next = argv[i + 1];
        if (next === undefined || next.startsWith("-")) {
          throw new Error(`${name} requires a value`);
        }
        i += 1;
        return next;
      }
      if (arg.startsWith(`${name}=`)) {
        const inline = arg.slice(name.length + 1);
        if (inline === "") throw new Error(`${name} requires a value`);
        return inline;
      }
      return undefined;
    };

    const dirValue = readValue("--dir");
    if (dirValue !== undefined) {
      dir = dirValue;
      continue;
    }

    const sarifValue = readValue("--sarif");
    if (sarifValue !== undefined) {
      sarif = sarifValue;
      continue;
    }

    const failValue = readValue("--fail-on");
    if (failValue !== undefined) {
      if (!FAIL_ON_LEVELS.includes(failValue as FailOn)) {
        throw new Error(
          `--fail-on must be one of error, warning, info, none (got ${failValue})`,
        );
      }
      failOn = failValue as FailOn;
      continue;
    }

    if (arg.startsWith("-")) {
      throw new Error(`unknown option ${arg}`);
    }
    files.push(arg);
  }

  return { files, json, quiet, dir, help, version, sarif, failOn };
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

function collectFiles(opts: Options): string[] {
  const files = [...opts.files];
  if (opts.dir !== null) {
    let entries: string[];
    try {
      entries = readdirSync(opts.dir);
    } catch (err) {
      throw new Error(
        `cannot read --dir ${opts.dir}: ${(err as Error).message}`,
      );
    }
    const matches = entries
      .filter((name) =>
        DOC_EXTENSIONS.some((ext) => name.toLowerCase().endsWith(ext)),
      )
      .sort();
    for (const name of matches) files.push(join(opts.dir, name));
  }
  return files;
}

export async function run(argv: string[]): Promise<number> {
  let opts: Options;
  try {
    opts = parseArgs(argv);
  } catch (err) {
    console.error(`officelens: ${(err as Error).message}`);
    console.error(USAGE);
    return 2;
  }

  if (opts.help) {
    console.log(USAGE);
    return 0;
  }

  const version = await readVersion();

  if (opts.version) {
    console.log(version);
    return 0;
  }

  let files: string[];
  try {
    files = collectFiles(opts);
  } catch (err) {
    console.error(`officelens: ${(err as Error).message}`);
    return 3;
  }

  if (files.length === 0) {
    if (opts.dir !== null) return 0;
    console.error(USAGE);
    return 2;
  }

  const results: AuditResult[] = [];
  let readErrors = 0;
  let parseErrors = 0;

  for (const file of files) {
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

    if (result.parseError) {
      parseErrors += 1;
      console.error(`officelens: cannot parse ${file}: ${result.parseError}`);
    }

    if (opts.quiet) {
      if (!result.parseError) {
        console.log(
          `${file}: ${result.counts.error} error(s), ${result.counts.warning} warning(s), ${result.counts.info} info`,
        );
      }
    } else if (!opts.json && !result.parseError) {
      console.log(formatText(result));
    }
  }

  if (opts.json) {
    const asArray = files.length !== 1 || results.length !== 1;
    const payload = asArray ? results : results[0];
    console.log(JSON.stringify(payload, null, 2));
  }

  if (opts.sarif) {
    try {
      await writeSarif(opts.sarif, results, TOOL_NAME, version);
    } catch (err) {
      console.error(
        `officelens: cannot write SARIF report to ${opts.sarif}: ${(err as Error).message}`,
      );
      return 3;
    }
  }

  if (readErrors > 0) return 3;
  if (parseErrors > 0) return 2;
  return exceedsThreshold(results, opts.failOn) ? 1 : 0;
}

if (import.meta.main) {
  run(process.argv.slice(2))
    .then((code) => process.exit(code))
    .catch((err: unknown) => {
      console.error(err);
      process.exit(1);
    });
}
