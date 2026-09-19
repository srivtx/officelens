import { strFromU8, unzipSync, type UnzipFileInfo } from "fflate";
import { XMLParser } from "fast-xml-parser";
import type { OoxmlPackage, PackagePart } from "./types.ts";

/**
 * Caps applied while reading the ZIP container, before any member is
 * decompressed. A hostile package must be rejected by its declared sizes
 * rather than allowed to exhaust memory.
 */
export interface UnzipLimits {
  maxMembers: number;
  maxEntryBytes: number;
  maxTotalBytes: number;
}

export const DEFAULT_UNZIP_LIMITS: UnzipLimits = {
  maxMembers: 65535,
  maxEntryBytes: 256 * 1024 * 1024,
  maxTotalBytes: 1024 * 1024 * 1024,
};

const relsParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  trimValues: false,
});

const contentTypeParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  trimValues: true,
});

export const CONTENT_TYPES_PART = "[Content_Types].xml";

export interface Relationship {
  target: string;
  type: string;
  mode: string;
}

interface ContentTypes {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

function readContentTypes(xml: string): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();

  let parsed: Record<string, unknown> | undefined;
  try {
    parsed = contentTypeParser.parse(xml) as Record<string, unknown> | undefined;
  } catch {
    return { defaults, overrides };
  }

  const root = parsed?.["Types"] as Record<string, unknown> | undefined;
  if (!root) return { defaults, overrides };

  const asList = (value: unknown): unknown[] =>
    Array.isArray(value) ? value : value ? [value] : [];

  for (const entry of asList(root["Default"])) {
    if (!entry || typeof entry !== "object") continue;
    const rec = entry as Record<string, unknown>;
    const extension = String(rec["@_Extension"] ?? "").toLowerCase();
    if (extension) defaults.set(extension, String(rec["@_ContentType"] ?? ""));
  }
  for (const entry of asList(root["Override"])) {
    if (!entry || typeof entry !== "object") continue;
    const rec = entry as Record<string, unknown>;
    const part = String(rec["@_PartName"] ?? "").replace(/^\/+/, "");
    if (part) overrides.set(part, String(rec["@_ContentType"] ?? ""));
  }

  return { defaults, overrides };
}

export function openOoxml(
  data: Uint8Array,
  limits: UnzipLimits = DEFAULT_UNZIP_LIMITS,
): OoxmlPackage {
  if (!data || data.length === 0) {
    throw new Error("empty input is not an OOXML package");
  }

  let files: Record<string, Uint8Array>;
  try {
    let memberCount = 0;
    let totalBytes = 0;
    files = unzipSync(data, {
      filter: (file: UnzipFileInfo): boolean => {
        if (file.name.endsWith("/")) return false;
        memberCount += 1;
        if (memberCount > limits.maxMembers) {
          throw new Error(`archive has more than ${limits.maxMembers} members`);
        }
        if (file.originalSize > limits.maxEntryBytes) {
          throw new Error(
            `archive member ${JSON.stringify(file.name)} expands beyond ${limits.maxEntryBytes} bytes`,
          );
        }
        totalBytes += file.originalSize;
        if (totalBytes > limits.maxTotalBytes) {
          throw new Error(
            `archive expands beyond ${limits.maxTotalBytes} bytes`,
          );
        }
        return true;
      },
    });
  } catch (err) {
    throw new Error(`not a readable zip archive: ${(err as Error).message}`);
  }

  const map = new Map<string, Uint8Array>();
  for (const [path, bytes] of Object.entries(files)) {
    if (path.endsWith("/")) continue;
    map.set(path, bytes);
  }

  const contentTypes = map.get(CONTENT_TYPES_PART);
  if (!contentTypes) {
    throw new Error(`missing ${CONTENT_TYPES_PART} (not an OOXML package)`);
  }
  const contentTypesXml = strFromU8(contentTypes);
  let hasTypesRoot = false;
  try {
    const parsed = contentTypeParser.parse(contentTypesXml) as
      | Record<string, unknown>
      | undefined;
    hasTypesRoot = Boolean(parsed?.["Types"]);
  } catch {
    hasTypesRoot = false;
  }
  if (!hasTypesRoot) {
    throw new Error(`invalid ${CONTENT_TYPES_PART} (no Types root)`);
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

export function findOfficeDocument(
  pkg: OoxmlPackage,
): { path: string; type: string } | undefined {
  try {
    const rels = readRels(pkg, "");
    for (const rel of Object.values(rels)) {
      if (rel.type.endsWith("/officeDocument")) {
        return { path: rel.target, type: rel.type };
      }
    }
  } catch {
    // fall through to no main part
  }
  return undefined;
}

export function contentTypeFor(
  pkg: OoxmlPackage,
  partPath: string,
): string | undefined {
  const xml = pkg.text(CONTENT_TYPES_PART);
  if (!xml) return undefined;
  const { defaults, overrides } = readContentTypes(xml);
  const override = overrides.get(partPath);
  if (override) return override;
  const dot = partPath.lastIndexOf(".");
  const extension = dot >= 0 ? partPath.slice(dot + 1).toLowerCase() : "";
  return extension ? defaults.get(extension) : undefined;
}

const DOCX_MAIN_TYPE =
  /application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document(?:\.macroEnabled)?\.main/i;
const PPTX_MAIN_TYPE =
  /application\/vnd\.openxmlformats-officedocument\.presentationml\.presentation(?:\.macroEnabled)?\.main/i;

export function detectDocumentKind(
  pkg: OoxmlPackage,
): "docx" | "pptx" | undefined {
  const candidates: Array<{ path: string; contentType: string }> = [];

  const office = findOfficeDocument(pkg);
  if (office) {
    candidates.push({
      path: office.path,
      contentType: contentTypeFor(pkg, office.path) ?? "",
    });
  }

  const xml = pkg.text(CONTENT_TYPES_PART);
  if (xml) {
    const { overrides } = readContentTypes(xml);
    for (const [path, contentType] of overrides) {
      if (DOCX_MAIN_TYPE.test(contentType) || PPTX_MAIN_TYPE.test(contentType)) {
        candidates.push({ path, contentType });
      }
    }
  }

  for (const candidate of candidates) {
    if (
      PPTX_MAIN_TYPE.test(candidate.contentType) ||
      /presentation\.xml$/i.test(candidate.path) ||
      candidate.path.startsWith("ppt/")
    ) {
      return "pptx";
    }
    if (
      DOCX_MAIN_TYPE.test(candidate.contentType) ||
      /document\.xml$/i.test(candidate.path) ||
      candidate.path.startsWith("word/")
    ) {
      return "docx";
    }
  }
  return undefined;
}
