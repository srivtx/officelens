/// <reference types="bun" />
import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { strToU8, zipSync } from "fflate";
import { audit } from "../src/audit";
import { formatText } from "../src/report";
import { parseArgs } from "../src/cli";
import {
  makeBadDocx,
  makeBadPptx,
  makeGoodDocx,
  makeGoodPptx,
} from "../scripts/make-fixtures";

const ROOT = join(import.meta.dir, "..");
const CLI = join(ROOT, "src", "cli.ts");

const CORRUPT_DOCX = zipSync({
  "[Content_Types].xml": strToU8(
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  ),
  "_rels/.rels": strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  ),
  "word/document.xml": strToU8("<w:document><w:body><w:p>corrupt"),
});

interface CliRun {
  code: number;
  stdout: string;
  stderr: string;
}

async function runCli(args: string[]): Promise<CliRun> {
  const proc = Bun.spawn(["bun", "run", CLI, ...args], {
    cwd: ROOT,
    stdout: "pipe",
    stderr: "pipe",
  });
  const [stdout, stderr] = await Promise.all([
    new Response(proc.stdout).text(),
    new Response(proc.stderr).text(),
  ]);
  return { code: await proc.exited, stdout, stderr };
}

async function tempFile(name: string, data: Uint8Array): Promise<string> {
  const dir = mkdtempSync(join(tmpdir(), "officelens-cli-"));
  const file = join(dir, name);
  await Bun.write(file, data);
  return file;
}

describe("docx fixtures", () => {
  test("bad docx reports errors", () => {
    expect(audit(makeBadDocx(), "bad.docx").counts.error).toBeGreaterThan(0);
  });

  test("good docx reports no errors", () => {
    expect(audit(makeGoodDocx(), "good.docx").counts.error).toBe(0);
  });

  test("formatText includes DOCX-ALT-001 for the bad docx", () => {
    const text = formatText(audit(makeBadDocx(), "bad.docx"));
    expect(text).toContain("DOCX-ALT-001");
  });
});

describe("pptx fixtures", () => {
  test("bad pptx reports errors", () => {
    expect(audit(makeBadPptx(), "bad.pptx").counts.error).toBeGreaterThan(0);
  });

  test("good pptx reports no errors", () => {
    expect(audit(makeGoodPptx(), "good.pptx").counts.error).toBe(0);
  });
});

describe("cli", () => {
  test("prints usage and exits 2 with no arguments", async () => {
    const proc = Bun.spawn(["bun", "run", CLI], {
      cwd: ROOT,
      stdout: "pipe",
      stderr: "pipe",
    });
    const code = await proc.exited;
    expect(code).toBe(2);
  });

  test("exits 1 for a file with errors", async () => {
    const dir = mkdtempSync(join(tmpdir(), "officelens-"));
    const file = join(dir, "bad.docx");
    await Bun.write(file, makeBadDocx());

    const proc = Bun.spawn(["bun", "run", CLI, "--quiet", file], {
      cwd: ROOT,
      stdout: "pipe",
      stderr: "pipe",
    });
    const code = await proc.exited;
    expect(code).toBe(1);
  });

  test("exits 0 for a clean file", async () => {
    const dir = mkdtempSync(join(tmpdir(), "officelens-"));
    const file = join(dir, "good.docx");
    await Bun.write(file, makeGoodDocx());

    const proc = Bun.spawn(["bun", "run", CLI, "--json", file], {
      cwd: ROOT,
      stdout: "pipe",
      stderr: "pipe",
    });
    const code = await proc.exited;
    expect(code).toBe(0);
  });
});

describe("cli option parsing", () => {
  test("accepts --flag=value for every value flag", () => {
    const opts = parseArgs([
      "--dir=docs",
      "--sarif=out.sarif",
      "--fail-on=warning",
    ]);
    expect(opts.dir).toBe("docs");
    expect(opts.sarif).toBe("out.sarif");
    expect(opts.failOn).toBe("warning");
  });

  test("treats everything after -- as a file path", () => {
    const opts = parseArgs(["--", "--not-an-option", "-x"]);
    expect(opts.files).toEqual(["--not-an-option", "-x"]);
  });

  test("rejects an empty inline value", () => {
    expect(() => parseArgs(["--sarif="])).toThrow(/requires a value/);
  });

  test("rejects an unknown option", () => {
    expect(() => parseArgs(["--bogus"])).toThrow(/unknown option --bogus/);
  });

  test("rejects an invalid --fail-on value", () => {
    expect(() => parseArgs(["--fail-on=loud"])).toThrow(/--fail-on/);
  });
});

describe("cli json contract", () => {
  test("a single input emits a JSON object", async () => {
    const file = await tempFile("good.docx", makeGoodDocx());
    const run = await runCli(["--json", file]);
    const parsed = JSON.parse(run.stdout) as unknown;
    expect(Array.isArray(parsed)).toBe(false);
    expect((parsed as { kind: string }).kind).toBe("docx");
    expect(run.code).toBe(0);
  });

  test("multiple inputs emit a single valid JSON array", async () => {
    const good = await tempFile("good.docx", makeGoodDocx());
    const bad = await tempFile("bad.pptx", makeBadPptx());
    const run = await runCli(["--json", good, bad]);
    const parsed = JSON.parse(run.stdout) as Array<{ file: string }>;
    expect(Array.isArray(parsed)).toBe(true);
    expect(parsed).toHaveLength(2);
    expect(run.code).toBe(1);
  });
});

describe("cli exit codes", () => {
  test("a corrupt package logs to stderr, keeps stdout free of the finding, and exits 2", async () => {
    const file = await tempFile("corrupt.docx", CORRUPT_DOCX);
    const run = await runCli([file]);
    expect(run.code).toBe(2);
    expect(run.stderr).toContain("cannot parse");
    expect(run.stderr).toContain("word/document.xml");
    expect(run.stdout).not.toContain("OOXML-000");
  });

  test("an unreadable input exits 3", async () => {
    const run = await runCli([join(tmpdir(), "officelens-missing-xyz.docx")]);
    expect(run.code).toBe(3);
    expect(run.stderr).toContain("cannot read");
  });

  test("an unreadable --dir exits 3", async () => {
    const run = await runCli(["--dir", join(tmpdir(), "officelens-no-dir-xyz")]);
    expect(run.code).toBe(3);
    expect(run.stderr).toContain("cannot read --dir");
  });

  test("an unwritable SARIF output exits 3", async () => {
    const good = await tempFile("good.docx", makeGoodDocx());
    const run = await runCli(["--quiet", "--sarif", "/no-such-dir/x.sarif", good]);
    expect(run.code).toBe(3);
    expect(run.stderr).toContain("cannot write SARIF");
  });

  test("an unknown option exits 2 with usage", async () => {
    const run = await runCli(["--bogus"]);
    expect(run.code).toBe(2);
    expect(run.stderr).toContain("unknown option --bogus");
    expect(run.stderr).toContain("Usage:");
  });

  test("--help documents the exit codes", async () => {
    const run = await runCli(["--help"]);
    expect(run.code).toBe(0);
    expect(run.stdout).toContain("Exit codes:");
    expect(run.stdout).toContain("0  no findings");
    expect(run.stdout).toContain("1  findings");
    expect(run.stdout).toContain("2  invalid usage");
    expect(run.stdout).toContain("3  an input file");
  });

  test("--version and -v print exactly the version string", async () => {
    const expected = (await Bun.file(join(ROOT, "package.json")).json()) as {
      version: string;
    };
    for (const flag of ["--version", "-v"]) {
      const run = await runCli([flag]);
      expect(run.code).toBe(0);
      expect(run.stdout.trim()).toBe(expected.version);
    }
  });
});

describe("cli --dir", () => {
  test("audits every document in the directory as a JSON array", async () => {
    const dir = mkdtempSync(join(tmpdir(), "officelens-dir-"));
    await Bun.write(join(dir, "good.docx"), makeGoodDocx());
    await Bun.write(join(dir, "good.pptx"), makeGoodPptx());
    await Bun.write(join(dir, "notes.txt"), strToU8("ignore me"));

    const run = await runCli(["--json", "--dir", dir]);
    const parsed = JSON.parse(run.stdout) as Array<{ file: string }>;
    expect(parsed).toHaveLength(2);
    expect(parsed.map((entry) => entry.file).sort()).toEqual([
      join(dir, "good.docx"),
      join(dir, "good.pptx"),
    ]);
    expect(run.code).toBe(0);
  });
});
