import type { AuditResult } from "./types";

export function formatText(result: AuditResult): string {
  const counts = (result as any).counts ?? { error: 0, warning: 0, info: 0 };
  const header = `${result.file} (${result.kind})  errors:${counts.error ?? 0}  warnings:${
    counts.warning ?? 0
  }  info:${counts.info ?? 0}`;

  const lines: string[] = [header];
  for (const issue of result.issues ?? []) {
    const location = issue.location && issue.location.length > 0 ? issue.location : "-";
    lines.push(
      `${String(issue.severity).toUpperCase()}  ${issue.code}  ${location}  ${issue.message}`,
    );
  }
  return lines.join("\n");
}

export function formatJson(result: AuditResult): string {
  return JSON.stringify(result, null, 2);
}
