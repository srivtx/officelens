import { XMLParser } from "fast-xml-parser";
import type { Issue, OoxmlPackage } from "./types.ts";
import { findMainDocument } from "./package.ts";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  trimValues: false,
});

export function findAll(node: unknown, names: string[]): any[] {
  const out: any[] = [];
  const wanted = new Set(names);

  const visit = (current: unknown): void => {
    if (Array.isArray(current)) {
      for (const item of current) visit(item);
      return;
    }
    if (!current || typeof current !== "object") return;

    const obj = current as Record<string, unknown>;
    for (const key of Object.keys(obj)) {
      if (!wanted.has(key)) continue;
      const value = obj[key];
      if (Array.isArray(value)) out.push(...value);
      else if (value !== undefined) out.push(value);
    }
    for (const key of Object.keys(obj)) visit(obj[key]);
  };

  visit(node);
  return out;
}

function toArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

function textContent(value: unknown): string {
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (Array.isArray(value)) return value.map(textContent).join("");
  if (value && typeof value === "object") {
    const rec = value as Record<string, unknown>;
    if (rec["#text"] !== undefined) return textContent(rec["#text"]);
  }
  return "";
}

function nodeText(node: unknown): string {
  return findAll(node, ["w:t"]).map(textContent).join("");
}

function headingLevel(paragraph: unknown): number | undefined {
  if (!paragraph || typeof paragraph !== "object") return undefined;
  const rec = paragraph as Record<string, unknown>;
  const pPr = rec["w:pPr"];
  if (pPr && typeof pPr === "object") {
    const pr = pPr as Record<string, unknown>;
    const style = pr["w:pStyle"];
    if (style && typeof style === "object") {
      const val = (style as Record<string, unknown>)["@_w:val"];
      if (typeof val === "string") {
        const match = /^heading\s*(\d+)$/i.exec(val);
        if (match && match[1]) return Number(match[1]);
        if (/^heading/i.test(val)) return 1;
      }
    }
    const outline = pr["w:outlineLvl"];
    if (outline && typeof outline === "object") {
      const val = (outline as Record<string, unknown>)["@_w:val"];
      const num = Number(val);
      if (val !== undefined && Number.isFinite(num)) return num + 1;
    }
  }
  return undefined;
}

function paragraphs(doc: unknown): unknown[] {
  return findAll(doc, ["w:p"]);
}

function ruleAlt(doc: unknown): Issue[] {
  const issues: Issue[] = [];
  for (const pr of findAll(doc, ["docPr", "wp:docPr"])) {
    if (!pr || typeof pr !== "object") continue;
    const rec = pr as Record<string, unknown>;
    const descr = rec["@_descr"];
    const title = rec["@_title"];
    const hasDescr = typeof descr === "string" && descr.trim() !== "";
    const hasTitle = typeof title === "string" && title.trim() !== "";
    if (!hasDescr && !hasTitle) {
      const id = rec["@_id"] ?? "unknown";
      issues.push({
        code: "DOCX-ALT-001",
        severity: "error",
        message: `Drawing (docPr id=${String(id)}) has no alt text (descr or title).`,
        location: "word/document.xml",
        wcag: "1.1.1",
      });
    }
  }
  return issues;
}

function ruleHeadingsPresent(doc: unknown): Issue[] {
  const bodyText = nodeText(doc).replace(/\s+/g, "");
  if (!bodyText) return [];
  const hasHeading = paragraphs(doc).some(
    (p) => headingLevel(p) !== undefined,
  );
  if (hasHeading) return [];
  return [
    {
      code: "DOCX-HEAD-002",
      severity: "warning",
      message: "Document contains body text but no headings.",
      location: "word/document.xml",
      wcag: "1.3.1",
    },
  ];
}

function ruleHeadingSkip(doc: unknown): Issue[] {
  const levels: number[] = [];
  for (const p of paragraphs(doc)) {
    const level = headingLevel(p);
    if (level !== undefined) levels.push(level);
  }

  let previous: number | undefined;
  for (const level of levels) {
    if (previous !== undefined && level - previous > 1) {
      return [
        {
          code: "DOCX-HEAD-003",
          severity: "warning",
          message: `Heading level skips from ${previous} to ${level}.`,
          location: "word/document.xml",
          wcag: "1.3.1",
        },
      ];
    }
    previous = level;
  }
  return [];
}

function ruleLanguage(doc: unknown, styles: unknown): Issue[] {
  const inDocument = findAll(doc, ["w:lang"]).length > 0;
  const inStyles = styles ? findAll(styles, ["w:lang"]).length > 0 : false;
  if (inDocument || inStyles) return [];
  return [
    {
      code: "DOCX-LANG-004",
      severity: "warning",
      message: "No document language (w:lang) is specified.",
      location: "word/document.xml",
      wcag: "3.1.1",
    },
  ];
}

function ruleTableHeader(doc: unknown): Issue[] {
  const issues: Issue[] = [];
  for (const tbl of findAll(doc, ["w:tbl"])) {
    if (!tbl || typeof tbl !== "object") continue;
    const rows = toArray((tbl as Record<string, unknown>)["w:tr"]);
    const first = rows[0];
    let hasHeader = false;
    if (first && typeof first === "object") {
      const trPr = (first as Record<string, unknown>)["w:trPr"];
      if (trPr && typeof trPr === "object") {
        hasHeader = "w:tblHeader" in (trPr as Record<string, unknown>);
      }
    }
    if (!hasHeader) {
      issues.push({
        code: "DOCX-TBL-005",
        severity: "error",
        message: "Table's first row is not marked as a header (w:tblHeader).",
        location: "word/document.xml",
        wcag: "1.3.1",
      });
    }
  }
  return issues;
}

function ruleLinkText(doc: unknown): Issue[] {
  const issues: Issue[] = [];
  for (const link of findAll(doc, ["w:hyperlink"])) {
    const text = nodeText(link).trim();
    if (/^https?:\/\//.test(text)) {
      issues.push({
        code: "DOCX-LINK-006",
        severity: "warning",
        message: `Hyperlink uses a raw URL as its text: "${text}".`,
        location: "word/document.xml",
        wcag: "2.4.4",
      });
    }
  }
  return issues;
}

export function auditDocx(pkg: OoxmlPackage): Issue[] {
  let doc: unknown;
  try {
    const main = findMainDocument(pkg, "docx") ?? "word/document.xml";
    const xml = pkg.text(main);
    if (!xml) return [];
    doc = parser.parse(xml);
  } catch {
    return [];
  }

  let styles: unknown;
  try {
    const xml = pkg.text("word/styles.xml");
    if (xml) styles = parser.parse(xml);
  } catch {
    styles = undefined;
  }

  const rules: Array<() => Issue[]> = [
    () => ruleAlt(doc),
    () => ruleHeadingsPresent(doc),
    () => ruleHeadingSkip(doc),
    () => ruleLanguage(doc, styles),
    () => ruleTableHeader(doc),
    () => ruleLinkText(doc),
  ];

  const issues: Issue[] = [];
  for (const rule of rules) {
    try {
      issues.push(...rule());
    } catch {
      // a single malformed rule must not abort the audit
    }
  }

  return issues.sort((a, b) => a.code.localeCompare(b.code));
}
