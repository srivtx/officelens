import type { AuditResult, DocumentKind, Issue, OoxmlPackage } from "./types";
import {
  DEFAULT_UNZIP_LIMITS,
  detectDocumentKind,
  openOoxml,
  type UnzipLimits,
} from "./package";
import { auditDocx } from "./docx";
import { auditPptx } from "./pptx";
import { PartParseError } from "./errors";

function countIssues(issues: Issue[]): { error: number; warning: number; info: number } {
  const counts = { error: 0, warning: 0, info: 0 };
  for (const issue of issues) {
    if (issue.severity === "error") counts.error += 1;
    else if (issue.severity === "warning") counts.warning += 1;
    else counts.info += 1;
  }
  return counts;
}

function unreadable(file: string, reason: string): AuditResult {
  const issues: Issue[] = [
    {
      code: "OOXML-000",
      severity: "error",
      message: reason,
      location: file,
    },
  ];
  return {
    file,
    kind: "unknown",
    issues,
    counts: countIssues(issues),
    parseError: reason,
  };
}

function parseFailed(
  file: string,
  kind: DocumentKind,
  reason: string,
): AuditResult {
  const issues: Issue[] = [
    {
      code: "OOXML-000",
      severity: "error",
      message: reason,
      location: file,
    },
  ];
  return {
    file,
    kind,
    issues,
    counts: countIssues(issues),
    parseError: reason,
  };
}

export function audit(
  data: Uint8Array,
  file = "document",
  limits: UnzipLimits = DEFAULT_UNZIP_LIMITS,
): AuditResult {
  let pkg: OoxmlPackage;
  try {
    pkg = openOoxml(data, limits);
  } catch (err) {
    return unreadable(file, `Not a readable OOXML package: ${(err as Error).message}`);
  }

  const kind = detectDocumentKind(pkg);
  if (!kind) {
    return unreadable(
      file,
      "Not a DOCX or PPTX package: no main document part found",
    );
  }

  let issues: Issue[] = [];
  try {
    issues = kind === "docx" ? auditDocx(pkg) : auditPptx(pkg);
  } catch (err) {
    if (err instanceof PartParseError) {
      return parseFailed(file, kind, `Unparseable package part: ${err.message}`);
    }
    return unreadable(file, `Failed to audit package: ${(err as Error).message}`);
  }

  return { file, kind, issues, counts: countIssues(issues) };
}
