import { XMLParser } from "fast-xml-parser";
import type { Issue, OoxmlPackage } from "./types.ts";
import {
  findMainDocument,
  findOfficeDocument,
  readRels,
  type Relationship,
} from "./package.ts";

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

function hasNonEmptyText(node: unknown): boolean {
  const bodies = findAll(node, ["w:txbxContent", "v:textbox"]);
  return bodies.some((body) => nodeText(body).replace(/\s+/g, "") !== "");
}

function hasLangValue(node: unknown): boolean {
  if (!node || typeof node !== "object") return false;
  const rec = node as Record<string, unknown>;
  const val = rec["@_w:val"];
  const eastAsia = rec["@_w:eastAsia"];
  return (
    (typeof val === "string" && val.trim() !== "") ||
    (typeof eastAsia === "string" && eastAsia.trim() !== "")
  );
}

function hasAltText(rec: Record<string, unknown>): boolean {
  const descr = rec["@_descr"] ?? rec["@_alt"];
  if (typeof descr === "string" && descr.trim() !== "") return true;
  const title = rec["@_title"] ?? rec["@_o:title"];
  return typeof title === "string" && title.trim() !== "";
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

function ruleAlt(doc: unknown, location: string): Issue[] {
  const issues: Issue[] = [];

  for (const drawing of findAll(doc, ["w:drawing"])) {
    if (hasNonEmptyText(drawing)) continue;
    for (const pr of findAll(drawing, ["wp:docPr", "docPr"])) {
      if (!pr || typeof pr !== "object") continue;
      const rec = pr as Record<string, unknown>;
      if (hasAltText(rec)) continue;
      const id = rec["@_id"] ?? "unknown";
      issues.push({
        code: "DOCX-ALT-001",
        severity: "error",
        message: `Drawing (docPr id=${String(id)}) has no alt text (descr or title).`,
        location,
        wcag: "1.1.1",
      });
    }
  }

  for (const pict of findAll(doc, ["w:pict"])) {
    if (hasNonEmptyText(pict)) continue;
    for (const shape of findAll(pict, ["v:shape"])) {
      if (!shape || typeof shape !== "object") continue;
      if (findAll(shape, ["v:imagedata"]).length === 0) continue;
      const rec = shape as Record<string, unknown>;
      if (hasAltText(rec)) continue;
      const id = rec["@_id"] ?? "unknown";
      issues.push({
        code: "DOCX-ALT-001",
        severity: "error",
        message: `VML image (v:shape id=${String(id)}) has no alt text (alt or title).`,
        location,
        wcag: "1.1.1",
      });
    }
  }

  return issues;
}

function ruleHeadingsPresent(doc: unknown, location: string): Issue[] {
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
      location,
      wcag: "1.3.1",
    },
  ];
}

function ruleHeadingSkip(doc: unknown, location: string): Issue[] {
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
          location,
          wcag: "1.3.1",
        },
      ];
    }
    previous = level;
  }
  return [];
}

function hasDocumentLanguage(
  styles: unknown,
  core: unknown,
  settings: unknown,
): boolean {
  for (const lang of findAll(core, ["dc:language"])) {
    if (textContent(lang).trim() !== "") return true;
  }

  for (const lang of findAll(settings, ["w:themeFontLang"])) {
    if (hasLangValue(lang)) return true;
  }

  for (const defaults of findAll(styles, ["w:docDefaults"])) {
    if (findAll(defaults, ["w:lang"]).some(hasLangValue)) return true;
  }

  for (const style of findAll(styles, ["w:style"])) {
    if (!style || typeof style !== "object") continue;
    const rec = style as Record<string, unknown>;
    const isDefault =
      rec["@_w:default"] === "1" || rec["@_w:default"] === 1;
    if (isDefault && findAll(style, ["w:lang"]).some(hasLangValue)) return true;
  }

  return false;
}

function ruleLanguage(
  styles: unknown,
  core: unknown,
  settings: unknown,
  location: string,
): Issue[] {
  if (hasDocumentLanguage(styles, core, settings)) return [];
  return [
    {
      code: "DOCX-LANG-004",
      severity: "warning",
      message: "No document language (w:lang) is specified.",
      location,
      wcag: "3.1.1",
    },
  ];
}

function ruleTableHeader(doc: unknown, location: string): Issue[] {
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
        location,
        wcag: "1.3.1",
      });
    }
  }
  return issues;
}

const URL_TOKEN =
  /(?:[a-z][a-z0-9+.-]*:\/\/|mailto:|tel:|www\.)[^\s<>"']+|\b[a-z0-9-]+(?:\.[a-z0-9-]+)*\.(?:com|org|net|edu|gov|io|co|dev|app|ai)\b/i;

function ruleLinkText(doc: unknown, location: string): Issue[] {
  const issues: Issue[] = [];
  for (const link of findAll(doc, ["w:hyperlink"])) {
    const text = nodeText(link).trim();
    const match = URL_TOKEN.exec(text);
    if (match) {
      issues.push({
        code: "DOCX-LINK-006",
        severity: "warning",
        message: `Hyperlink text contains a raw address: "${match[0]}".`,
        location,
        wcag: "2.4.4",
      });
    }
  }
  return issues;
}

const CONTENT_REL_SUFFIXES = [
  "/header",
  "/footer",
  "/footnotes",
  "/endnotes",
  "/comments",
];

function relationshipTarget(
  rels: Record<string, Relationship>,
  suffix: string,
): string | undefined {
  for (const rel of Object.values(rels)) {
    if (rel.type.endsWith(suffix)) return rel.target;
  }
  return undefined;
}

function parsePart(pkg: OoxmlPackage, path: string): unknown {
  try {
    const xml = pkg.text(path);
    if (!xml) return undefined;
    return parser.parse(xml);
  } catch {
    return undefined;
  }
}

function resolveMainPart(pkg: OoxmlPackage): string | undefined {
  const office = findOfficeDocument(pkg);
  if (office && pkg.get(office.path)) return office.path;
  const conventional = findMainDocument(pkg, "docx") ?? "word/document.xml";
  return pkg.get(conventional) ? conventional : undefined;
}

export function auditDocx(pkg: OoxmlPackage): Issue[] {
  const main = resolveMainPart(pkg);
  if (!main) return [];

  const doc = parsePart(pkg, main);
  if (doc === undefined) return [];

  const rels = readRels(pkg, main);
  const rootRels = readRels(pkg, "");
  const stylesPath =
    relationshipTarget(rels, "/styles") ??
    (pkg.get("word/styles.xml") ? "word/styles.xml" : undefined);
  const settingsPath =
    relationshipTarget(rels, "/settings") ??
    (pkg.get("word/settings.xml") ? "word/settings.xml" : undefined);
  const corePath =
    relationshipTarget(rootRels, "/core-properties") ??
    (pkg.get("docProps/core.xml") ? "docProps/core.xml" : undefined);

  const styles = stylesPath ? parsePart(pkg, stylesPath) : undefined;
  const settings = settingsPath ? parsePart(pkg, settingsPath) : undefined;
  const core = corePath ? parsePart(pkg, corePath) : undefined;

  const contentParts: Array<{ path: string; root: unknown }> = [
    { path: main, root: doc },
  ];
  for (const rel of Object.values(rels)) {
    if (!CONTENT_REL_SUFFIXES.some((suffix) => rel.type.endsWith(suffix))) {
      continue;
    }
    if (rel.mode.toLowerCase() === "external") continue;
    const root = parsePart(pkg, rel.target);
    if (root !== undefined) contentParts.push({ path: rel.target, root });
  }

  const issues: Issue[] = [];
  const run = (fn: () => Issue[]): void => {
    try {
      issues.push(...fn());
    } catch {
      // a single malformed rule must not abort the audit
    }
  };

  run(() => ruleAlt(doc, main));
  run(() => ruleHeadingsPresent(doc, main));
  run(() => ruleHeadingSkip(doc, main));
  run(() => ruleLanguage(styles, core, settings, main));
  run(() => ruleTableHeader(doc, main));
  run(() => ruleLinkText(doc, main));

  for (const part of contentParts) {
    if (part.path === main) continue;
    run(() => ruleAlt(part.root, part.path));
    run(() => ruleHeadingSkip(part.root, part.path));
    run(() => ruleTableHeader(part.root, part.path));
    run(() => ruleLinkText(part.root, part.path));
  }

  return issues.sort((a, b) => a.code.localeCompare(b.code));
}
