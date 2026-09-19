/// <reference types="bun" />
import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { strToU8, zipSync } from "fflate";
import { audit } from "../src/audit";
import { openOoxml } from "../src/package";

const ROOT = join(import.meta.dir, "..");
const CLI = join(ROOT, "src", "cli.ts");

const MALFORMED = new Uint8Array([
  0x50, 0x4b, 0x03, 0x04, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
]);

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;

function cliExit(file: string): Promise<number> {
  const proc = Bun.spawn(["bun", "run", CLI, "--quiet", file], {
    cwd: ROOT,
    stdout: "pipe",
    stderr: "pipe",
  });
  return proc.exited;
}

async function writeTemp(name: string, data: Uint8Array): Promise<string> {
  const dir = mkdtempSync(join(tmpdir(), "officelens-bad-"));
  const file = join(dir, name);
  await Bun.write(file, data);
  return file;
}

describe("unreadable input is a hard failure", () => {
  test("a malformed zip does not pass", () => {
    expect(() => openOoxml(MALFORMED)).toThrow();

    const result = audit(MALFORMED, "malformed.zip");
    expect(result.parseError).toBeDefined();
    expect(result.kind).toBe("unknown");
    expect(result.counts.error).toBeGreaterThan(0);
  });

  test("a zip missing [Content_Types].xml does not pass", () => {
    const data = zipSync({ "word/document.xml": strToU8("<w:document/>") });
    expect(() => openOoxml(data)).toThrow();

    const result = audit(data, "no-content-types.docx");
    expect(result.parseError).toBeDefined();
    expect(result.counts.error).toBeGreaterThan(0);
  });

  test("a valid non-OOXML zip does not pass", () => {
    const data = zipSync({ "hello.txt": strToU8("not a document") });
    expect(() => openOoxml(data)).toThrow();

    const result = audit(data, "plain.zip");
    expect(result.parseError).toBeDefined();
    expect(result.counts.error).toBeGreaterThan(0);
  });

  test("an OPC package with no main document part does not pass", () => {
    const data = zipSync({
      "[Content_Types].xml": strToU8(CONTENT_TYPES),
      "_rels/.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
      ),
      "docProps/app.xml": strToU8("<Properties/>"),
    });
    const result = audit(data, "opc-only.zip");
    expect(result.parseError).toBeDefined();
    expect(result.kind).toBe("unknown");
    expect(result.counts.error).toBeGreaterThan(0);
  });

  test("a malformed zip exits the CLI with code 2", async () => {
    const file = await writeTemp("malformed.zip", MALFORMED);
    expect(await cliExit(file)).toBe(2);
  });

  test("a non-OOXML zip exits the CLI with code 2", async () => {
    const file = await writeTemp(
      "plain.zip",
      zipSync({ "hello.txt": strToU8("not a document") }),
    );
    expect(await cliExit(file)).toBe(2);
  });
});
