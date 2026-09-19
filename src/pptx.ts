import { XMLParser } from "fast-xml-parser";
import type { Issue, OoxmlPackage } from "./types";
import { readRels } from "./package";

const SHAPE_TAGS = ["p:sp", "p:pic", "p:cxnSp"];
const CNVPR_TAGS = ["p:cNvPr", "a:cNvPr", "cNvPr"];

function makeParser(): XMLParser {
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    parseTagValue: false,
    trimValues: false,
  });
}

function collect(node: any, names: string[], out: any[] = []): any[] {
  if (node === null || node === undefined) return out;
  if (typeof node !== "object") return out;
  if (Array.isArray(node)) {
    for (const item of node) collect(item, names, out);
    return out;
  }
  for (const [key, value] of Object.entries(node)) {
    if (names.includes(key)) {
      if (Array.isArray(value)) {
        for (const v of value) out.push(v);
      } else {
        out.push(value);
      }
    }
    collect(value, names, out);
  }
  return out;
}

function readPart(pkg: OoxmlPackage, path: string): string | undefined {
  const p: any = pkg as any;

  if (p instanceof Map) {
    const v = p.get(path);
    if (typeof v === "string") return v;
    if (v instanceof Uint8Array) return new TextDecoder().decode(v);
  }

  for (const name of [
    "read",
    "get",
    "getFile",
    "readFile",
    "readText",
    "getText",
    "text",
    "open",
  ]) {
    const fn = p?.[name];
    if (typeof fn === "function") {
      try {
        const v = fn.call(p, path);
        if (typeof v === "string") return v;
        if (v instanceof Uint8Array) return new TextDecoder().decode(v);
      } catch {
        void 0;
      }
    }
  }

  for (const bag of [p?.files, p?.parts, p?.entries, p?.map, p?.contents]) {
    if (!bag) continue;
    try {
      if (bag instanceof Map) {
        if (bag.has(path)) {
          const v = bag.get(path);
          if (typeof v === "string") return v;
          if (v instanceof Uint8Array) return new TextDecoder().decode(v);
        }
      } else if (typeof bag === "object") {
        const v = (bag as any)[path];
        if (typeof v === "string") return v;
        if (v instanceof Uint8Array) return new TextDecoder().decode(v);
      }
    } catch {
      void 0;
    }
  }

  return undefined;
}

function toPairs(raw: any): Array<[string, string]> {
  const pairs: Array<[string, string]> = [];
  const add = (id: any, target: any) => {
    if (typeof id === "string" && typeof target === "string") pairs.push([id, target]);
  };

  if (!raw) return pairs;

  if (raw instanceof Map) {
    for (const [key, value] of raw) {
      if (typeof value === "string") add(key, value);
      else if (value && typeof value === "object") {
        add(
          (value as any).id ?? (value as any)["@_Id"] ?? (value as any).Id,
          (value as any).target ?? (value as any)["@_Target"] ?? (value as any).Target,
        );
      }
    }
    return pairs;
  }

  if (Array.isArray(raw)) {
    for (const r of raw) {
      add(
        r?.id ?? r?.["@_Id"] ?? r?.Id,
        r?.target ?? r?.["@_Target"] ?? r?.Target,
      );
    }
    return pairs;
  }

  if (typeof raw === "object") {
    const nested =
      (raw as any).Relationships?.Relationship ??
      (raw as any).Relationship ??
      (raw as any).relationship;
    if (nested) return toPairs(nested);

    if ((raw as any)["@_Id"] || (raw as any).Id) {
      add((raw as any)["@_Id"] ?? (raw as any).Id, (raw as any)["@_Target"] ?? (raw as any).Target);
      return pairs;
    }

    for (const [key, value] of Object.entries(raw)) {
      if (typeof value === "string") add(key, value);
      else if (value && typeof value === "object") {
        add(
          (value as any).id ?? (value as any)["@_Id"] ?? (value as any).Id,
          (value as any).target ?? (value as any)["@_Target"] ?? (value as any).Target,
        );
      }
    }
  }

  return pairs;
}

function resolvePath(base: string, target: string): string {
  const t = target.replace(/\\/g, "/");
  if (t.startsWith("/")) return t.replace(/^\/+/, "");
  if (t.startsWith(base + "/")) return t;
  const segments = (base + "/" + t).split("/");
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") out.pop();
    else out.push(segment);
  }
  return out.join("/");
}

function getRels(pkg: OoxmlPackage): Map<string, string> {
  const parser = makeParser();
  const map = new Map<string, string>();
  const relsPath = "ppt/_rels/presentation.xml.rels";

  const direct = readPart(pkg, relsPath);
  if (direct) {
    try {
      for (const [id, target] of toPairs(parser.parse(direct))) map.set(id, target);
    } catch {
      void 0;
    }
  }

  if (map.size === 0) {
    for (const arg of ["ppt/presentation.xml", relsPath]) {
      try {
        const raw: any = readRels(pkg, arg);
        for (const [id, target] of toPairs(raw)) map.set(id, target);
      } catch {
        void 0;
      }
      if (map.size > 0) break;
    }
  }

  return map;
}

interface SlidePart {
  path: string;
  root: any;
  layoutPath?: string;
}

function parseXml(xml: string): any {
  try {
    return makeParser().parse(xml);
  } catch {
    return undefined;
  }
}

function findLayoutPath(pkg: OoxmlPackage, slidePath: string): string | undefined {
  let rels: Record<string, { target: string; type: string }> = {};
  try {
    rels = readRels(pkg, slidePath);
  } catch {
    return undefined;
  }
  for (const rel of Object.values(rels)) {
    if (rel.type.endsWith("/slideLayout")) return rel.target;
  }
  return undefined;
}

function getSlides(pkg: OoxmlPackage): SlidePart[] {
  const parser = makeParser();
  const presXml = readPart(pkg, "ppt/presentation.xml");
  if (!presXml) return [];

  let pres: any;
  try {
    pres = parser.parse(presXml);
  } catch {
    return [];
  }

  const rels = getRels(pkg);
  const sldIds = collect(pres, ["p:sldId"]);
  const slides: SlidePart[] = [];

  for (const sldId of sldIds) {
    const rid = sldId?.["@_r:id"] ?? sldId?.["@_id"] ?? sldId?.id;
    if (!rid) continue;
    const target = rels.get(rid);
    if (!target) continue;
    const path = resolvePath("ppt", target);
    const xml = readPart(pkg, path);
    if (!xml) continue;
    let root: any;
    try {
      root = parser.parse(xml);
    } catch {
      continue;
    }
    slides.push({ path, root, layoutPath: findLayoutPath(pkg, path) });
  }

  return slides;
}

export function auditPptx(pkg: OoxmlPackage): Issue[] {
  const issues: Issue[] = [];
  const slides = getSlides(pkg);

  try {
    for (const slide of slides) {
      for (const shape of collect(slide.root, SHAPE_TAGS)) {
        if (collect(shape, ["p:ph"]).length > 0) continue;
        for (const cnv of collect(shape, CNVPR_TAGS)) {
          const descr = cnv?.["@_descr"];
          if (typeof descr !== "string" || descr.trim() === "") {
            const name = cnv?.["@_name"] ?? cnv?.["@_id"] ?? "";
            issues.push({
              code: "PPTX-ALT-001",
              severity: "error",
              wcag: "1.1.1",
              message: `Missing alternative text (descr)${name ? ` for shape "${name}"` : ""}`,
              location: slide.path,
            });
          }
        }
      }
    }
  } catch {
    void 0;
  }

  const isTitlePlaceholder = (ph: any): boolean => {
    const type = ph?.["@_type"];
    return type === "title" || type === "ctrTitle";
  };

  try {
    for (const slide of slides) {
      const hasTitlePlaceholder = collect(slide.root, ["p:ph"]).some(
        isTitlePlaceholder,
      );

      let hasTitleName = false;
      if (!hasTitlePlaceholder) {
        hasTitleName = collect(slide.root, ["p:sp"]).some((sp) =>
          collect(sp, ["p:cNvPr"]).some(
            (cnv) => typeof cnv?.["@_name"] === "string" && /title/i.test(cnv["@_name"]),
          ),
        );
      }

      let layoutHasTitle = false;
      if (!hasTitlePlaceholder && !hasTitleName && slide.layoutPath) {
        const layoutXml = readPart(pkg, slide.layoutPath);
        const layoutRoot = layoutXml ? parseXml(layoutXml) : undefined;
        if (layoutRoot) {
          layoutHasTitle = collect(layoutRoot, ["p:ph"]).some(isTitlePlaceholder);
        }
      }

      if (!hasTitlePlaceholder && !hasTitleName && !layoutHasTitle) {
        issues.push({
          code: "PPTX-TITLE-002",
          severity: "warning",
          wcag: "1.3.1",
          message: "Slide has no title placeholder",
          location: slide.path,
        });
      }
    }
  } catch {
    void 0;
  }

  try {
    for (const slide of slides) {
      const tables = collect(slide.root, ["a:tbl"]);
      const hasUnlabeledTable = tables.some((tbl) => {
        const props = collect(tbl, ["a:tblPr"]);
        if (props.length === 0) return true;
        return !props.some(
          (pr) => pr?.["@_firstRow"] === "1" || pr?.["@_firstRow"] === 1,
        );
      });
      if (hasUnlabeledTable) {
        issues.push({
          code: "PPTX-TBL-003",
          severity: "warning",
          wcag: "1.3.1",
          message: "Table does not have a header row (firstRow) enabled",
          location: slide.path,
        });
      }
    }
  } catch {
    void 0;
  }

  try {
    const runLanguage = (run: any): boolean => {
      const lang = run?.["@_lang"];
      return typeof lang === "string" && lang.trim() !== "";
    };
    const hasLanguage = (node: any): boolean =>
      collect(node, ["a:rPr", "a:defRPr", "a:endParaRPr"]).some(runLanguage);

    let hasDefaultLanguage = false;
    const presXml = readPart(pkg, "ppt/presentation.xml");
    const pres = presXml ? parseXml(presXml) : undefined;
    if (pres) {
      for (const style of collect(pres, ["p:defaultTextStyle"])) {
        if (hasLanguage(style)) hasDefaultLanguage = true;
      }
    }

    const layoutsWithLanguage = new Set<string>();
    for (const path of pkg.list()) {
      const isMaster = /^ppt\/slideMasters\/[^/]+\.xml$/i.test(path);
      const isLayout = /^ppt\/slideLayouts\/[^/]+\.xml$/i.test(path);
      if (!isMaster && !isLayout) continue;
      const xml = readPart(pkg, path);
      const root = xml ? parseXml(xml) : undefined;
      if (!root) continue;
      const hasStylesLanguage = collect(root, ["p:txStyles"]).some(hasLanguage);
      if (!hasStylesLanguage) continue;
      if (isMaster) hasDefaultLanguage = true;
      else layoutsWithLanguage.add(path);
    }

    for (const slide of slides) {
      const hasTextRuns = collect(slide.root, ["a:t"]).length > 0;
      if (!hasTextRuns) continue;
      const inheritsFromLayout =
        slide.layoutPath !== undefined && layoutsWithLanguage.has(slide.layoutPath);
      if (hasLanguage(slide.root) || hasDefaultLanguage || inheritsFromLayout) {
        continue;
      }
      issues.push({
        code: "PPTX-LANG-004",
        severity: "warning",
        wcag: "3.1.1",
        message: "No language specified on this slide's text runs",
        location: slide.path,
      });
    }
  } catch {
    void 0;
  }

  issues.sort((a, b) => (a.code < b.code ? -1 : a.code > b.code ? 1 : 0));
  return issues;
}
