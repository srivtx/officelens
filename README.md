# ooxml-a11y

An offline, dependency-light accessibility auditor for **DOCX** and **PPTX**
files, designed to run in scripts and CI.

## Why

Microsoft Office ships an accessibility checker, but it is GUI-only: there is no
supported way to run it headlessly, so accessibility regressions in Word and
PowerPoint documents cannot be caught automatically. Other document formats do
not have this problem — PDF, EPUB, and HTML all have mature offline auditing
tools. Office Open XML (OOXML) does not. `ooxml-a11y` fills that gap with a
small, auditable, zero-network CLI that parses the OOXML package directly.

- **Offline** — nothing is uploaded; the tool never touches the network.
- **Dependency-light** — only `fflate` (zip) and `fast-xml-parser`.
- **Scriptable** — stable exit codes and machine-readable JSON output.
- **CI-ready** — one command, no Office installation required.

## Install

```sh
bun install
```

Run the CLI directly with Bun:

```sh
bun run src/cli.ts report.docx slides.pptx
```

Or link the `ooxml-a11y` binary defined in `package.json`:

```sh
bun link
ooxml-a11y report.docx slides.pptx
```

## Usage

```
ooxml-a11y <file...> [--json] [--quiet]
```

| Flag | Description |
| --- | --- |
| `--json` | Print one JSON object per file. |
| `--quiet` | Print only a one-line summary per file. |
| `-h`, `--help` | Show usage. |

Human-readable output (default):

```sh
ooxml-a11y report.docx
```

```text
report.docx (docx)  errors:2  warnings:1  info:0
ERROR  DOCX-ALT-001  word/document.xml  Drawing (docPr id=1) has no alt text (descr or title).
ERROR  DOCX-TBL-005  word/document.xml  Table's first row is not marked as a header (w:tblHeader).
WARNING  DOCX-LANG-004  word/document.xml  No document language (w:lang) is specified.
```

JSON output:

```sh
ooxml-a11y --json report.docx
```

```json
{
  "file": "report.docx",
  "kind": "docx",
  "issues": [
    {
      "code": "DOCX-ALT-001",
      "severity": "error",
      "message": "Drawing (docPr id=1) has no alt text (descr or title).",
      "location": "word/document.xml",
      "wcag": "1.1.1"
    }
  ],
  "counts": { "error": 1, "warning": 0, "info": 0 }
}
```

Exit codes:

- `0` — no errors found (warnings and info may still be present).
- `1` — at least one error was found, or a file could not be read.
- `2` — no input files were provided.

## Rules

Each rule maps to a WCAG 2.1 success criterion.

### DOCX

| Code | Severity | Description | WCAG |
| --- | --- | --- | --- |
| `DOCX-ALT-001` | error | Images and drawings must have alternative text (`descr` or `title`). | 1.1.1 (A) |
| `DOCX-TBL-005` | error | A data table's first row must be marked as a header row (`w:tblHeader`). | 1.3.1 (A) |
| `DOCX-HEAD-002` | warning | Documents with body text should use at least one heading. | 1.3.1 (A) |
| `DOCX-HEAD-003` | warning | Heading levels must not be skipped (e.g. `Heading1` to `Heading3`). | 1.3.1 (A) |
| `DOCX-LINK-006` | warning | Link text must be descriptive, not a raw URL. | 2.4.4 (A) |
| `DOCX-LANG-004` | warning | The document must declare a language (`w:lang`). | 3.1.1 (A) |

### PPTX

| Code | Severity | Description | WCAG |
| --- | --- | --- | --- |
| `PPTX-ALT-001` | error | Pictures and other non-placeholder shapes must have alternative text (`descr`). | 1.1.1 (A) |
| `PPTX-TITLE-002` | error | Every slide must have a title placeholder. | 1.3.1 (A), 2.4.2 (A) |
| `PPTX-TBL-003` | warning | A table's first row must be marked with `firstRow="1"`. | 1.3.1 (A) |
| `PPTX-LANG-004` | warning | Slide text runs should declare a language (`lang`). | 3.1.1 (A) |

## CI integration

Fail a build when a document regresses:

```yaml
- name: Audit office documents
  run: bunx ooxml-a11y --quiet docs/*.docx docs/*.pptx
```

Because the exit code is `1` whenever any error is reported, no extra parsing of
the output is needed. Use `--json` if you want to post results as a review
comment or feed them into another tool.

## Library API

```ts
import { audit, formatText } from "./src/index";
import { readFileSync } from "node:fs";

const data = new Uint8Array(readFileSync("report.docx"));
const result = audit(data, "report.docx");

console.log(formatText(result));
```

The package's `src/index.ts` entry point also exports `formatJson`, `openOoxml`,
`auditDocx`, `auditPptx`, and the `AuditResult`, `Issue`, and `Severity` types.

## Development

```sh
bun run typecheck     # bunx tsc --noEmit
bun test              # run the test suite
bun run make-fixtures # regenerate fixtures/*.docx and fixtures/*.pptx
```

## License

MIT — see [LICENSE](./LICENSE).
