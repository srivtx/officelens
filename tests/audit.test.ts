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

const DOCX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const PPTX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`;

const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

function rootRels(main: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="${main}"/>
</Relationships>`;
}

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

describe("present but unparseable parts", () => {
  test("a corrupt DOCX main document is a parse error, not a clean audit", () => {
    const data = zipSync({
      "[Content_Types].xml": strToU8(DOCX_CONTENT_TYPES),
      "_rels/.rels": strToU8(rootRels("word/document.xml")),
      "word/document.xml": strToU8("<w:document><w:body><w:p>corrupt"),
    });

    const result = audit(data, "corrupt.docx");
    expect(result.kind).toBe("docx");
    expect(result.parseError).toBeDefined();
    expect(result.counts.error).toBeGreaterThan(0);
    const fatal = result.issues.find((issue) => issue.code === "OOXML-000");
    expect(fatal?.severity).toBe("error");
  });

  test("a corrupt PPTX slide is a parse error, not a clean audit", () => {
    const data = zipSync({
      "[Content_Types].xml": strToU8(PPTX_CONTENT_TYPES),
      "_rels/.rels": strToU8(rootRels("ppt/presentation.xml")),
      "ppt/presentation.xml": strToU8(
        `<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OFFICE_REL}/slide" Target="slides/slide1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8("<p:sld><p:cSld><p:spTree>corrupt"),
    });

    const result = audit(data, "corrupt.pptx");
    expect(result.kind).toBe("pptx");
    expect(result.parseError).toBeDefined();
    expect(result.counts.error).toBeGreaterThan(0);
    expect(result.issues.map((issue) => issue.code)).toContain("OOXML-000");
  });

  test("a corrupt DOCX main document exits the CLI with code 2", async () => {
    const data = zipSync({
      "[Content_Types].xml": strToU8(DOCX_CONTENT_TYPES),
      "_rels/.rels": strToU8(rootRels("word/document.xml")),
      "word/document.xml": strToU8("<w:document><w:body><w:p>corrupt"),
    });
    const file = await writeTemp("corrupt.docx", data);
    expect(await cliExit(file)).toBe(2);
  });
});

describe("zip limits", () => {
  const DOCX_WITH_MEMBER = (size: number): Uint8Array =>
    zipSync({
      "[Content_Types].xml": strToU8(DOCX_CONTENT_TYPES),
      "_rels/.rels": strToU8(rootRels("word/document.xml")),
      "word/document.xml": strToU8(
        `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`,
      ),
      "word/media/big.bin": strToU8("A".repeat(size)),
    });

  test("rejects a member whose declared size exceeds maxEntryBytes", () => {
    const data = DOCX_WITH_MEMBER(4096);
    const limits = {
      maxMembers: 100,
      maxEntryBytes: 1024,
      maxTotalBytes: 1024 * 1024,
    };
    expect(() => openOoxml(data, limits)).toThrow(/expands beyond/);

    const result = audit(data, "bomb.docx", limits);
    expect(result.parseError).toBeDefined();
    expect(result.counts.error).toBeGreaterThan(0);
  });

  test("rejects a package with too many members", () => {
    const data = DOCX_WITH_MEMBER(1);
    const limits = {
      maxMembers: 1,
      maxEntryBytes: 1024 * 1024,
      maxTotalBytes: 1024 * 1024,
    };
    expect(() => openOoxml(data, limits)).toThrow(/more than 1 members/);
  });

  test("rejects a package whose total uncompressed size exceeds maxTotalBytes", () => {
    const data = DOCX_WITH_MEMBER(2048);
    const limits = {
      maxMembers: 100,
      maxEntryBytes: 1024 * 1024,
      maxTotalBytes: 1024,
    };
    expect(() => openOoxml(data, limits)).toThrow(/expands beyond/);
  });

  test("accepts the same package under the default limits", () => {
    expect(() => openOoxml(DOCX_WITH_MEMBER(4096))).not.toThrow();
  });
});
