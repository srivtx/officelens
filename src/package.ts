import { strFromU8, unzipSync } from "fflate";
import { XMLParser } from "fast-xml-parser";
import type { OoxmlPackage, PackagePart } from "./types.ts";

const relsParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  trimValues: false,
});

export interface Relationship {
  target: string;
  type: string;
  mode: string;
}

export function openOoxml(data: Uint8Array): OoxmlPackage {
  const files = unzipSync(data);
  const map = new Map<string, Uint8Array>();
  for (const [path, bytes] of Object.entries(files)) {
    if (path.endsWith("/")) continue;
    map.set(path, bytes);
  }
  const parts: PackagePart[] = [...map.entries()].map(([path, bytes]) => ({
    path,
    data: bytes,
  }));
  return {
    parts,
    list: () => [...map.keys()].sort(),
    get: (path) => map.get(path),
    text: (path) => {
      const bytes = map.get(path);
      return bytes ? strFromU8(bytes) : undefined;
    },
  };
}

export function partName(relTarget: string, baseDir: string): string {
  const isAbsolute = relTarget.startsWith("/");
  const combined = isAbsolute
    ? relTarget
    : baseDir && baseDir !== "." && baseDir !== ""
      ? `${baseDir}/${relTarget}`
      : relTarget;
  const out: string[] = [];
  for (const seg of combined.split("/")) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") {
      out.pop();
      continue;
    }
    out.push(seg);
  }
  return out.join("/");
}

function relsPathFor(partPath: string): string {
  if (partPath === "" || partPath === "/") return "_rels/.rels";
  const slash = partPath.lastIndexOf("/");
  const dir = slash >= 0 ? partPath.slice(0, slash) : "";
  const base = slash >= 0 ? partPath.slice(slash + 1) : partPath;
  return dir ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`;
}

export function readRels(
  pkg: OoxmlPackage,
  partPath: string,
): Record<string, Relationship> {
  const result: Record<string, Relationship> = {};
  const xml = pkg.text(relsPathFor(partPath));
  if (!xml) return result;

  const parsed = relsParser.parse(xml) as Record<string, unknown> | undefined;
  const root = parsed?.["Relationships"] as Record<string, unknown> | undefined;
  if (!root) return result;

  const raw = root["Relationship"];
  const rels: unknown[] = Array.isArray(raw) ? raw : raw ? [raw] : [];
  const slash = partPath.lastIndexOf("/");
  const dir = slash >= 0 ? partPath.slice(0, slash) : "";

  for (const rel of rels) {
    if (!rel || typeof rel !== "object") continue;
    const rec = rel as Record<string, unknown>;
    const id = rec["@_Id"];
    if (id === undefined) continue;
    result[String(id)] = {
      target: partName(String(rec["@_Target"] ?? ""), dir),
      type: String(rec["@_Type"] ?? ""),
      mode: String(rec["@_TargetMode"] ?? "Internal"),
    };
  }
  return result;
}

export function findMainDocument(
  pkg: OoxmlPackage,
  kind: "docx" | "pptx",
): string | undefined {
  try {
    const rels = readRels(pkg, "");
    for (const rel of Object.values(rels)) {
      if (rel.type.endsWith("/officeDocument")) return rel.target;
    }
  } catch {
    // fall through to the conventional path
  }
  return kind === "docx" ? "word/document.xml" : "ppt/presentation.xml";
}
