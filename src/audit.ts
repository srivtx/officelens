import type { AuditResult, Issue, OoxmlPackage } from "./types";
import { findMainDocument, openOoxml } from "./package";
import { auditDocx } from "./docx";
import { auditPptx } from "./pptx";

function countIssues(issues: Issue[]): { error: number; warning: number; info: number } {
  const counts = { error: 0, warning: 0, info: 0 };
  for (const issue of issues) {
    if (issue.severity === "error") counts.error += 1;
    else if (issue.severity === "warning") counts.warning += 1;
    else counts.info += 1;
  }
  return counts;
}

export function audit(data: Uint8Array, file = "document"): AuditResult {
  let pkg: OoxmlPackage;
  try {
    pkg = openOoxml(data);
  } catch {
    const issues: Issue[] = [
      {
        code: "OOXML-000",
        severity: "info",
        message: "Not a readable OOXML package",
        location: file,
      },
    ];
    return { file, kind: "docx", issues, counts: countIssues(issues) };
  }

  let kind: "docx" | "pptx" = "pptx";
  try {
    const main = findMainDocument(pkg, "docx");
    const isPresentation = typeof main === "string" && /presentation\.xml$/i.test(main);
    const hasMain = typeof main === "string" && pkg.text(main) !== undefined;
    if (main && !isPresentation && hasMain) kind = "docx";
  } catch {
    void 0;
  }

  let issues: Issue[] = [];
  try {
    issues = kind === "docx" ? auditDocx(pkg) : auditPptx(pkg);
  } catch {
    issues = [];
  }

  return { file, kind, issues, counts: countIssues(issues) };
}
