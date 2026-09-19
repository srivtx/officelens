/// <reference types="bun" />
import { describe, expect, test } from "bun:test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { audit } from "../src/audit";
import { toSarif, writeSarif } from "../src/sarif";
import { makeBadDocx } from "../scripts/make-fixtures";

describe("sarif", () => {
  test("emits SARIF 2.1.0 for a single result", () => {
    const result = audit(makeBadDocx(), "bad.docx");
    const sarif = toSarif(result, "officelens", "0.1.0");

    expect(sarif.version).toBe("2.1.0");
    expect(sarif.$schema).toBe("https://json.schemastore.org/sarif-2.1.0.json");
    expect(sarif.runs[0]!.tool.driver.name).toBe("officelens");
    expect(sarif.runs[0]!.tool.driver.version).toBe("0.1.0");

    const alt = sarif.runs[0]!.results.find((r) => r.ruleId === "DOCX-ALT-001");
    expect(alt).toBeDefined();
    expect(alt!.level).toBe("error");
    expect(alt!.locations[0]!.physicalLocation.region.startLine).toBe(1);
  });

  test("deduplicates rules across many results", () => {
    const result = audit(makeBadDocx(), "bad.docx");
    const sarif = toSarif([result, result], "officelens", "0.1.0");
    const ids = sarif.runs[0]!.tool.driver.rules.map((rule) => rule.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  test("writeSarif writes parseable JSON", async () => {
    const dir = mkdtempSync(join(tmpdir(), "officelens-sarif-"));
    const out = join(dir, "officelens.sarif");
    const result = audit(makeBadDocx(), "bad.docx");

    await writeSarif(out, result, "officelens", "0.1.0");

    const parsed = JSON.parse(await Bun.file(out).text());
    expect(parsed.version).toBe("2.1.0");
    expect(parsed.runs[0].tool.driver.name).toBe("officelens");
    expect(Array.isArray(parsed.runs[0].results)).toBe(true);
  });
});
