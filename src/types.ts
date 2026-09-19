export type Severity = "error" | "warning" | "info";

export type DocumentKind = "docx" | "pptx" | "unknown";

export interface Issue {
  code: string;
  severity: Severity;
  message: string;
  location: string;
  wcag?: string;
}

export interface AuditResult {
  file: string;
  kind: DocumentKind;
  issues: Issue[];
  counts: Record<Severity, number>;
  parseError?: string;
}

export interface PackagePart {
  path: string;
  data: Uint8Array;
}

export interface OoxmlPackage {
  parts: PackagePart[];
  list(): string[];
  get(path: string): Uint8Array | undefined;
  text(path: string): string | undefined;
}
