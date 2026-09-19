/// <reference types="bun" />
import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { audit } from "../src/audit";
import { formatText } from "../src/report";
import {
  makeBadDocx,
  makeBadPptx,
  makeGoodDocx,
  makeGoodPptx,
} from "../scripts/make-fixtures";

const ROOT = join(import.meta.dir, "..");
const CLI = join(ROOT, "src", "cli.ts");

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
    const dir = mkdtempSync(join(tmpdir(), "ooxml-a11y-"));
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
    const dir = mkdtempSync(join(tmpdir(), "ooxml-a11y-"));
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
