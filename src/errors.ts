/**
 * Thrown when a part that is present in the package cannot be parsed as
 * well-formed XML. Distinct from a missing part: an absent part is simply
 * skipped, while an unparseable part must fail the audit instead of being
 * silently treated as clean.
 */
export class PartParseError extends Error {
  readonly part: string;
  readonly detail: string;

  constructor(part: string, detail: string) {
    super(`could not parse ${part}: ${detail}`);
    this.name = "PartParseError";
    this.part = part;
    this.detail = detail;
  }
}
