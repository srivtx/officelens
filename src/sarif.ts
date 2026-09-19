import type { AuditResult, Issue, Severity } from "./types";

const SARIF_VERSION = "2.1.0";
const SARIF_SCHEMA = "https://json.schemastore.org/sarif-2.1.0.json";
const INFORMATION_URI = "https://github.com/srivtx/officelens";

export interface SarifRule {
  id: string;
  shortDescription: { text: string };
  properties: { tags: string[] };
}

export interface SarifResult {
  ruleId: string;
  level: "error" | "warning" | "note";
  message: { text: string };
  locations: Array<{
    physicalLocation: {
      artifactLocation: { uri: string };
      region: { startLine: number };
    };
  }>;
  properties: { tags: string[] };
}

export interface SarifLog {
  version: "2.1.0";
  $schema: string;
  runs: Array<{
    tool: {
      driver: {
        name: string;
        version: string;
        informationUri: string;
        rules: SarifRule[];
      };
    };
    results: SarifResult[];
  }>;
}

function toLevel(severity: Severity): "error" | "warning" | "note" {
  if (severity === "error") return "error";
  if (severity === "warning") return "warning";
  return "note";
}

function resultTags(issue: Issue): string[] {
  const tags = ["accessibility"];
  if (issue.wcag) tags.push(`wcag:${issue.wcag}`);
  return tags;
}

export function toSarif(
  results: AuditResult | AuditResult[],
  toolName: string,
  toolVersion: string,
): SarifLog {
  const list = Array.isArray(results) ? results : [results];
  const issues = list.flatMap((result) => result.issues ?? []);

  const descriptions = new Map<string, string>();
  for (const issue of issues) {
    if (!descriptions.has(issue.code)) descriptions.set(issue.code, issue.message);
  }

  const rules: SarifRule[] = [...descriptions.keys()].sort().map((code) => ({
    id: code,
    shortDescription: { text: descriptions.get(code) ?? code },
    properties: { tags: ["accessibility"] },
  }));

  const sarifResults: SarifResult[] = issues.map((issue) => ({
    ruleId: issue.code,
    level: toLevel(issue.severity),
    message: { text: issue.message },
    locations: [
      {
        physicalLocation: {
          artifactLocation: { uri: issue.location },
          region: { startLine: 1 },
        },
      },
    ],
    properties: { tags: resultTags(issue) },
  }));

  return {
    version: SARIF_VERSION,
    $schema: SARIF_SCHEMA,
    runs: [
      {
        tool: {
          driver: {
            name: toolName,
            version: toolVersion,
            informationUri: INFORMATION_URI,
            rules,
          },
        },
        results: sarifResults,
      },
    ],
  };
}

export async function writeSarif(
  path: string,
  results: AuditResult | AuditResult[],
  toolName: string,
  toolVersion: string,
): Promise<void> {
  const sarif = toSarif(results, toolName, toolVersion);
  await Bun.write(path, JSON.stringify(sarif, null, 2));
}
